//! M0 import gate — fails the build when a new hand-rolled coordinate
//! parser or formatter appears outside `refs/refs.zig`.
//!
//! Before M0 the tree had six independent A1 parsers and seven
//! column-letter formatters that disagreed on the column base, on case
//! handling, and on whether `A01` was a valid row. They were not
//! written by careless people — each was added locally, in a file that
//! could not see the others, because there was nowhere shared to put
//! it. `refs/refs.zig` is that place, and this gate is what stops the
//! sixth-and-seventh from growing back.
//!
//! Detection is deliberately narrow and syntactic: bijective base-26
//! arithmetic is the fingerprint every one of those implementations
//! shared, and it has essentially no other use in this codebase. A
//! file is flagged when a `26` multiply-or-modulo appears within a few
//! lines of an ASCII-letter offset (`'A'`) — a window, not a single
//! line, because real implementations split the two halves across
//! lines and a line-local test would be defeated by pressing Enter.
//!
//! Run as `zig build test` (wired into the test step), so CI enforces
//! it. Paths come from `build.zig` via argv — no directory walking, no
//! assumption about the working directory.

const std = @import("std");

/// Files allowed to contain base-26 coordinate arithmetic. Every entry
/// needs a reason; an entry without one is a bug in review, not a
/// license to duplicate.
const Allowed = struct {
    /// Suffix of the path, matched against the tail so the build's
    /// absolute paths work.
    suffix: []const u8,
    why: []const u8,
};

const allowlist = [_]Allowed{
    .{
        .suffix = "refs/refs.zig",
        .why = "the one owner — this is the implementation",
    },
    .{
        .suffix = "refs/import_gate.zig",
        .why = "this file; the patterns appear in its own source",
    },
    .{
        .suffix = "src/xlsx.zig",
        .why = "test-only: a fuzz generator builds synthetic refs with 'A' + (i % 26). " ++
            "Not a parser — it emits input, never interprets it.",
    },
    .{
        .suffix = "src/writer.zig",
        .why = "test-only: a comment-fixture loop walks A..Z the same way.",
    },
    .{
        .suffix = "src/dbx.zig",
        .why = "false positive: a percent-encoding test string contains the literal '%26'.",
    },
};

pub fn main(init: std.process.Init) !u8 {
    // 0.16 supplies the allocator, Io and argv through process.Init.
    const arena = init.arena.allocator();
    const io = init.io;

    const args = try init.minimal.args.toSlice(arena);
    if (args.len < 2) {
        std.debug.print("import_gate: no files to scan (build.zig wiring bug)\n", .{});
        return error.NoInputs;
    }

    var violations: usize = 0;
    var scanned: usize = 0;

    for (args[1..]) |path| {
        if (allowedFor(path) != null) continue;

        const source = std.Io.Dir.cwd().readFileAlloc(
            io,
            path,
            arena,
            .limited(8 * 1024 * 1024),
        ) catch |err| {
            std.debug.print("import_gate: cannot read {s}: {s}\n", .{ path, @errorName(err) });
            return err;
        };
        scanned += 1;

        if (try findViolation(arena, source)) |line_no| {
            violations += 1;
            std.debug.print(
                \\
                \\import gate: {s}:{d}
                \\  bijective base-26 coordinate arithmetic outside `refs/refs.zig`.
                \\
                \\  That is how a column letter is parsed or formatted, and there is
                \\  exactly one place for it now. Use `zlsx_refs`:
                \\    parse  a cell   -> refs.parseCell(s, .{{ .case = ..., .leading_zero_row = ... }})
                \\    parse  a column -> refs.parseColNumber(s, .{{ ... }}) / refs.scanColPrefix(...)
                \\    format a cell   -> refs.formatCell(buf, cell)
                \\    format a column -> refs.writeColLetters(buf, col)      // in-grid, typed
                \\                       refs.writeColNumberLetters(buf, n)  // unchecked, legacy
                \\
                \\  If this file genuinely needs its own (a test fixture generator, say),
                \\  add it to `allowlist` in refs/import_gate.zig WITH A REASON.
                \\
            , .{ path, line_no });
        }
    }

    if (violations > 0) {
        std.debug.print("import gate: {d} file(s) violated; {d} scanned\n", .{ violations, scanned });
        return 1;
    }
    return 0;
}

fn allowedFor(path: []const u8) ?[]const u8 {
    for (allowlist) |entry| {
        if (std.mem.endsWith(u8, path, entry.suffix)) return entry.why;
    }
    return null;
}

/// How many lines either side of the base-26 arithmetic to search for
/// the ASCII letter origin. Real implementations split the two halves
/// across lines all the time — `src/writer.zig` does — so a
/// single-line test would be bypassed by pressing Enter.
const window_lines: usize = 3;

/// Returns the 1-based line number of the first base-26 coordinate
/// arithmetic, or null. Arithmetic qualifies when a `26`
/// multiply-or-modulo appears within `window_lines` of an ASCII letter
/// origin (`'A'`) — the two halves together are what make it a column
/// codec rather than an incidental 26.
fn findViolation(gpa: std.mem.Allocator, source: []const u8) !?usize {
    // The allocation failure propagates rather than being swallowed:
    // a gate that fails open is worse than no gate at all.
    var lines: std.ArrayList([]const u8) = .empty;
    defer lines.deinit(gpa);

    var it = std.mem.splitScalar(u8, source, '\n');
    while (it.next()) |line| {
        try lines.append(gpa, line);
    }

    for (lines.items, 0..) |line, i| {
        if (!hasBase26(line)) continue;
        const lo = i -| window_lines;
        const hi = @min(i + window_lines + 1, lines.items.len);
        for (lines.items[lo..hi]) |near| {
            if (std.mem.indexOf(u8, near, "'A'") != null) return i + 1;
        }
    }
    return null;
}

fn hasBase26(line: []const u8) bool {
    return std.mem.indexOf(u8, line, "* 26") != null or
        std.mem.indexOf(u8, line, "*26") != null or
        std.mem.indexOf(u8, line, "% 26") != null or
        std.mem.indexOf(u8, line, "%26") != null;
}

test "findViolation catches a hand-rolled parser" {
    const parser =
        \\fn parseCol(s: []const u8) u32 {
        \\    var v: u32 = 0;
        \\    for (s) |c| v = v * 26 + (c - 'A' + 1);
        \\    return v;
        \\}
    ;
    try std.testing.expectEqual(@as(?usize, 3), try findViolation(std.testing.allocator, parser));
}

test "findViolation catches a hand-rolled formatter" {
    const formatter =
        \\while (n > 0) {
        \\    buf[i] = 'A' + @as(u8, @intCast((n - 1) % 26));
        \\    n = (n - 1) / 26;
        \\}
    ;
    try std.testing.expectEqual(@as(?usize, 2), try findViolation(std.testing.allocator, formatter));
}

test "findViolation catches a codec split across lines" {
    const split =
        \\const upper = std.ascii.toUpper(c);
        \\const offset: u32 = upper - 'A' + 1;
        \\
        \\v = v * 26 + offset;
    ;
    try std.testing.expectEqual(@as(?usize, 4), try findViolation(std.testing.allocator, split));
}

test "findViolation ignores an incidental 26" {
    const innocent =
        \\const limit = count * 26; // rows per page
        \\const letter_count = 26;
        \\if (std.mem.eql(u8, s, "..%2Fx%3Fy%3D1%26z")) return;
    ;
    try std.testing.expectEqual(@as(?usize, null), try findViolation(std.testing.allocator, innocent));
}

test "allowlist entries all carry a reason" {
    for (allowlist) |entry| {
        try std.testing.expect(entry.suffix.len > 0);
        try std.testing.expect(entry.why.len > 10);
    }
}
