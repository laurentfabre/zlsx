//! §9.1e's TEXT-MASS workload: unique multi-kilobyte strings.
//!
//! §9.1d measured three shapes whose retained text never exceeded
//! **9.11 B per cell** against a slope of ~473 B/cell, and said so: a
//! design in which the quantity under test is 1.9 % of the figure has no
//! power to price it, however far that quantity is spread. This fixture
//! exists to give the text term power. Every data row carries a **unique
//! 2 048-byte string** and a formula that produces a second one, so the
//! text the run retains is `2 × 2 048` bytes over two cells — **2 048 B
//! per cell**, four times the whole per-cell slope rather than two
//! percent of it.
//!
//! Topology:
//!
//! | Col | Content | Rows |
//! |---|---|---|
//! | A | a unique 2 048-byte lowercase-ASCII string | data |
//! | B | `UPPER(A…)` — a second 2 048-byte string, per row | data |
//!
//! `UPPER` and not `TRIM` or `LEFT`: the result must be a *new*
//! allocation the size of its input, so the run holds two multi-kilobyte
//! strings per row and not one plus a slice of it. The pool is
//! lowercase letters and digits only, so the uppercase result is
//! byte-for-byte the same length as its input (`ß` and the Greek final
//! sigma are length-changing, and this fixture is measuring bytes, not
//! casing edge cases — `synth_text_mix.zig` already owns those).
//!
//! Uniqueness is load-bearing. The shared-string table deduplicates, so
//! a repeated pool would make the retained text a constant rather than a
//! per-row cost, and the fixture would measure nothing. Each string
//! opens with its row index and continues with an LCG-driven body, so no
//! two rows share bytes beyond coincidence.
//!
//! No leading or trailing whitespace and no XML-significant characters:
//! an SST entry with edge whitespace needs `xml:space="preserve"` to
//! round-trip, and a fixture that loses bytes on the way back in
//! measures a different workbook than the one it emitted.

const std = @import("std");
const xlsx = @import("zlsx");

const Allocator = std.mem.Allocator;

/// Data columns per row: the string.
pub const data_cols: u32 = 1;

/// Formula columns per row: `UPPER` of it.
pub const formula_cols: u32 = 1;

/// Bytes in every generated string at the identity size. Multi-kilobyte
/// by the goal's wording, and two orders of magnitude above the 82-byte
/// rows the M8c text fixture carries.
pub const default_string_bytes: u32 = 2_048;

/// The generator's scratch buffer, and so the largest string a sweep can
/// ask for. Excel's own cell limit is 32 767 characters; this is well
/// under it and covers the 512 → 8 192 sweep §9.1e runs.
pub const max_string_bytes: u32 = 16_384;

/// The alphabet the body is drawn from: lowercase ASCII and digits, so
/// `UPPER` preserves length exactly and nothing needs XML escaping.
const alphabet = "abcdefghijklmnopqrstuvwxyz0123456789";

pub const Geometry = struct {
    /// Data rows, excluding the one header row.
    data_rows: u32,

    /// Bytes per generated string. A *knob*, not a constant, and that is
    /// the point: holding `data_rows` fixed while moving this moves the
    /// retained text with the cell count, the row count, the formula
    /// count and the graph all literally unchanged. It is the only way
    /// to price a text term by measurement rather than by attributing a
    /// residual to it — which is the error §9.1d's Codex round killed
    /// three separate times.
    string_bytes: u32 = default_string_bytes,

    pub fn cells(self: Geometry) u64 {
        return @as(u64, self.data_rows) * (data_cols + formula_cols);
    }

    pub fn formulaCells(self: Geometry) u64 {
        return @as(u64, self.data_rows) * formula_cols;
    }

    /// Bytes of text the run retains: the input string plus the string
    /// its formula produces, per row. The quantity §9.1d could not
    /// price.
    pub fn retainedTextBytes(self: Geometry) u64 {
        return @as(u64, self.data_rows) * 2 * self.string_bytes;
    }
};

/// The small size: 2 000 cells, 4 MB of text.
pub const tiny: Geometry = .{ .data_rows = 1_000 };

/// The identity size the recorded §9.1e numbers bind to: 5 000 rows,
/// 10 000 cells, 20 MB of retained text.
pub const small: Geometry = .{ .data_rows = 5_000 };

/// SHA-256 of `bytes(gpa, io, small)`. Same contract as every other
/// fixture's identity digest.
pub const small_digest_sha256 =
    "cc36c810b0025caf834c011230d8f4e393e5a9ba59b59b114497e5c0e5e34b79";

pub const digest_len = std.crypto.hash.sha2.Sha256.digest_length;

/// Serialize the fixture into memory — `Writer.saveToOwnedBuffer`, so
/// the digest is a property of the archive and not of a filesystem
/// round-trip.
pub fn bytes(gpa: Allocator, io: std.Io, g: Geometry) ![]u8 {
    var w = xlsx.Writer.init(gpa);
    defer w.deinit();
    try build(&w, g);
    return w.saveToOwnedBuffer(gpa, io);
}

/// Lowercase hex SHA-256 of `data`, into a caller-provided buffer.
pub fn digestHex(data: []const u8, out: *[digest_len * 2]u8) []const u8 {
    var raw: [digest_len]u8 = undefined;
    std.crypto.hash.sha2.Sha256.hash(data, &raw, .{});
    return std.fmt.bufPrint(out, "{x}", .{&raw}) catch unreachable;
}

/// One row's string: an index prefix that guarantees uniqueness, then an
/// LCG body so the rest of the bytes differ too. Both halves are pure
/// functions of `i`, which is what makes the archive reproducible.
pub fn bigString(buf: []u8, i: u32) []const u8 {
    const head = std.fmt.bufPrint(buf, "row-{d:0>9}-", .{i}) catch unreachable;
    var st: u32 = 0x9E37_79B9 ^ i;
    for (buf[head.len..]) |*b| {
        st = st *% 1_664_525 +% 1_013_904_223;
        b.* = alphabet[(st >> 24) % alphabet.len];
    }
    return buf;
}

fn build(w: *xlsx.Writer, g: Geometry) !void {
    std.debug.assert(g.string_bytes >= 16);
    std.debug.assert(g.string_bytes <= max_string_bytes);

    var s = try w.addSheet("BIGTEXT");
    try s.writeRow(&.{ .{ .string = "blob" }, .{ .string = "upper" } });

    var str_storage: [max_string_bytes]u8 = undefined;
    const str_buf = str_storage[0..g.string_bytes];
    var fbuf: [32]u8 = undefined;
    var cells: [data_cols + formula_cols]xlsx.Cell = undefined;
    var formulas: [data_cols + formula_cols]?[]const u8 = undefined;

    var i: u32 = 0;
    while (i < g.data_rows) : (i += 1) {
        const r = i + 2; // one header row above
        cells[0] = .{ .string = bigString(str_buf, i) };
        formulas[0] = null;
        // A deliberately wrong cache, and a string one — the result is
        // text, and a poison value the formula could itself produce is a
        // cell the recalc would be free to leave alone.
        cells[1] = .{ .string = "#STALE" };
        formulas[1] = try std.fmt.bufPrint(&fbuf, "UPPER(A{d})", .{r});
        try s.writeRowWithFormulas(&cells, &formulas);
    }
}

// ─── tests ───────────────────────────────────────────────────────────

const testing = std.testing;

test "bigtext-mix: the generator is deterministic" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    const g: Geometry = .{ .data_rows = 32 };
    const x = try bytes(a, io, g);
    defer a.free(x);
    const y = try bytes(a, io, g);
    defer a.free(y);
    try testing.expectEqualSlices(u8, x, y);

    var hx: [digest_len * 2]u8 = undefined;
    var hy: [digest_len * 2]u8 = undefined;
    try testing.expectEqualStrings(digestHex(x, &hx), digestHex(y, &hy));
}

test "bigtext-mix: the strings are the size claimed, unique, and uppercase-safe" {
    var a_buf: [default_string_bytes]u8 = undefined;
    var b_buf: [default_string_bytes]u8 = undefined;

    const a_str = bigString(&a_buf, 7);
    try testing.expectEqual(@as(usize, default_string_bytes), a_str.len);
    // Uniqueness is what keeps the SST from folding the column into one
    // entry, so it is asserted rather than assumed.
    const b_str = bigString(&b_buf, 8);
    try testing.expect(!std.mem.eql(u8, a_str, b_str));
    try testing.expect(std.mem.startsWith(u8, a_str, "row-000000007-"));

    // Every byte is in the alphabet or the fixed prefix: no XML escape,
    // no edge whitespace, and `UPPER` cannot change the length.
    for (a_str) |c| {
        const ok = (c >= 'a' and c <= 'z') or (c >= '0' and c <= '9') or c == '-';
        try testing.expect(ok);
    }
}

test "bigtext-mix: the topology is the one the table describes" {
    const a = testing.allocator;

    const g: Geometry = .{ .data_rows = 50 };
    try testing.expectEqual(@as(u64, 100), g.cells());
    try testing.expectEqual(@as(u64, 50), g.formulaCells());
    try testing.expectEqual(@as(u64, 50 * 2 * default_string_bytes), g.retainedTextBytes());

    var w = xlsx.Writer.init(a);
    defer w.deinit();
    try build(&w, g);
    const body = w.sheets.items[0].body.items;

    try testing.expectEqual(@as(usize, 50), std.mem.count(u8, body, "<f>"));
    try testing.expectEqual(@as(usize, 50), std.mem.count(u8, body, "<f>UPPER(A"));
    try testing.expect(std.mem.indexOf(u8, body, "<f>UPPER(A2)</f>") != null);

    // 50 distinct SST entries for the blobs plus the two header labels,
    // and **not** a 53rd for the poison cache: a formula cell's cached
    // text ships as `t="str"` inline, never through the shared table.
    // If the SST folded the blob column the fixture would retain a
    // constant rather than a per-row cost, which is the thing to catch.
    try testing.expectEqual(@as(usize, 52), w.sst_plan.new_strings.items.len);
}

test "bigtext-mix: the density knob moves text and nothing else" {
    const a = testing.allocator;

    // §9.1e's text-term measurement rests on this: at a fixed row count,
    // `string_bytes` must move the retained text and leave the cell
    // count, the formula count and the formulas themselves alone. If it
    // moved anything else, the sweep's ΔRSS would have two causes and
    // the per-byte figure would be the residual-attribution error again.
    const lo: Geometry = .{ .data_rows = 20, .string_bytes = 512 };
    const hi: Geometry = .{ .data_rows = 20, .string_bytes = 4_096 };

    try testing.expectEqual(lo.cells(), hi.cells());
    try testing.expectEqual(lo.formulaCells(), hi.formulaCells());
    try testing.expectEqual(@as(u64, 20 * 2 * 512), lo.retainedTextBytes());
    try testing.expectEqual(@as(u64, 20 * 2 * 4_096), hi.retainedTextBytes());

    var w_lo = xlsx.Writer.init(a);
    defer w_lo.deinit();
    try build(&w_lo, lo);
    var w_hi = xlsx.Writer.init(a);
    defer w_hi.deinit();
    try build(&w_hi, hi);

    const body_lo = w_lo.sheets.items[0].body.items;
    const body_hi = w_hi.sheets.items[0].body.items;
    // The sheet body holds the formulas and the cell skeleton — the
    // strings live in the SST — so it is *byte-identical* across the
    // sweep. That is the invariance the measurement needs.
    try testing.expectEqualSlices(u8, body_lo, body_hi);
    try testing.expectEqual(
        w_lo.sst_plan.new_strings.items.len,
        w_hi.sst_plan.new_strings.items.len,
    );
}
