//! §9's TEXT-heavy bench workload: the M8c fixture.
//!
//! The F1-mix fixture prices the graph and the criteria fixture prices
//! the aligned cursor; this one prices **text evaluation in volume**.
//! Every data row carries six text formulas over its own three cells,
//! so the formula count scales WITH the data — the opposite of the
//! criteria fixture's fixed report — and what is being timed is the
//! per-formula cost of the text stack: `casing_v1`'s segmentation,
//! `criteria.fold`'s positional map, §5.3b's three-way parse, and
//! `numfmt_v1`'s renderer, each reached through the registry the way a
//! workbook reaches it.
//!
//! A bench over one function would baseline the wrong mix (M8a
//! decision 14 deferred exactly this until F3 landed); the mix here
//! spans three rows of the ladder on purpose: M4f's TRIM, M8b's
//! PROPER, M8c's TEXTBEFORE/TEXTAFTER/NUMBERVALUE/FIXED/DOLLAR, with
//! M8a's renderer under the last two.
//!
//! Topology (fixed, same discipline and same reasons as
//! `synth_f1_mix.zig`):
//!
//! | Col | Content | Rows |
//! |---|---|---|
//! | A | dirty display name — padding, mangled case, `'`/`-`/`ß`/CJK | data |
//! | B | comma-joined 4-segment path, mixed-case pool | data |
//! | C | number spelling — grouped, plain, percent, padded | data |
//! | D | `PROPER(TRIM(A…))` | data |
//! | E | `TEXTBEFORE(B…,",")` — raw match | data |
//! | F | `TEXTAFTER(B…,",",-1,1)` — FOLDED match | data |
//! | G | `NUMBERVALUE(C…)` | data |
//! | H | `FIXED(G…,1)` — a real graph edge into numfmt | data |
//! | I | `DOLLAR(G…)` | data |
//!
//! `E` and `F` split one delimiter across the two match modes, so a
//! regression in the fold and one in the scan are visible as such.
//! Every formula is row-local — no whole-column scans — so the two
//! sizes separate the **marginal per-row cost** from everything paid
//! once, which is what §9 asks this baseline to state.

const std = @import("std");
const xlsx = @import("zlsx");

const Allocator = std.mem.Allocator;

/// Data columns per row: raw name, path, amount spelling.
pub const data_cols: u32 = 3;

/// Text formulas per row: D…I.
pub const formula_cols: u32 = 6;

/// Distinct dirty first names column `A` cycles through. The mangled
/// case is the point: PROPER has to move something in every cell.
const first_pool = [_][]const u8{
    "aLiCe", "bOb", "cHaRlIe", "dAnA", "eLiF", "fRaNk", "gRaCe", "hAnK",
};

/// Distinct last names — the apostrophe class, the hyphen boundary,
/// SpecialCasing's `ß`, precomposed accents and an uncased-letter tail,
/// so the bench walks the same paths M8b's fixtures pinned.
const last_pool = [_][]const u8{
    "jOhNsOn", "o'BRIEN",
    "müLLER",
    "sTRAße",
    "dU-PONT", "iVanOv",
    "五十嵐a",
    "ΣΟΦΟΣ",
};

/// Path segments column `B` joins with commas.
const seg_pool = [_][]const u8{
    "alpha", "BETA", "gamma", "DELTA", "epsilon", "ZETA", "eta", "THETA",
};

pub const Geometry = struct {
    /// Data rows, excluding the one header row.
    data_rows: u32,

    pub fn cells(self: Geometry) u64 {
        return @as(u64, self.data_rows) * (data_cols + formula_cols);
    }

    pub fn formulaCells(self: Geometry) u64 {
        return @as(u64, self.data_rows) * formula_cols;
    }
};

/// The smaller size: the same mix, a tenth of the rows.
pub const tiny: Geometry = .{ .data_rows = 1_000 };

/// The identity size the recorded TEXT baseline binds to: 60 000 text
/// formulas over 10 000 rows.
pub const small: Geometry = .{ .data_rows = 10_000 };

/// SHA-256 of `bytes(gpa, io, small)` — the same contract as
/// `synth_criteria_mix.small_digest_sha256`: this identifies the
/// workload the recorded baseline was measured on, and a mismatch
/// means re-measure, not "fix the writer".
pub const small_digest_sha256 =
    "63eddc8fc5472ba9a3d8747ced82464ff7fdbfc1a96cec434aba0e0f50f39899";

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

fn rawName(buf: []u8, i: u32) []const u8 {
    return std.fmt.bufPrint(buf, "  {s}  {s}  ", .{
        first_pool[i % first_pool.len],
        last_pool[(i / first_pool.len) % last_pool.len],
    }) catch unreachable;
}

fn path(buf: []u8, i: u32) []const u8 {
    return std.fmt.bufPrint(buf, "{s},{s},{s},{s}", .{
        seg_pool[i % seg_pool.len],
        seg_pool[(i + 3) % seg_pool.len],
        seg_pool[(i + 5) % seg_pool.len],
        seg_pool[(i + 6) % seg_pool.len],
    }) catch unreachable;
}

/// Deterministic value spread over 0…999 — the criteria fixture's
/// coprime multiplier, reused.
fn valueOf(i: u32) u32 {
    return (i * 37) % 1_000;
}

/// The amount spelling cycles four shapes, so NUMBERVALUE's grouped,
/// plain, percent and whitespace paths all run in volume.
fn amount(buf: []u8, i: u32) []const u8 {
    return switch (i % 4) {
        0 => std.fmt.bufPrint(buf, "{d},{d:0>3}", .{ (i % 9) + 1, valueOf(i) }),
        1 => std.fmt.bufPrint(buf, "{d}.5", .{valueOf(i)}),
        2 => std.fmt.bufPrint(buf, "{d}.25%", .{valueOf(i)}),
        else => std.fmt.bufPrint(buf, " {d} ", .{valueOf(i)}),
    } catch unreachable;
}

fn build(w: *xlsx.Writer, g: Geometry) !void {
    var s = try w.addSheet("TEXT");
    try s.writeRow(&.{
        .{ .string = "raw_name" }, .{ .string = "path" }, .{ .string = "amount" },
    });

    var name_buf: [64]u8 = undefined;
    var path_buf: [64]u8 = undefined;
    var amt_buf: [24]u8 = undefined;
    var fbuf: [formula_cols][48]u8 = undefined;
    var cells: [data_cols + formula_cols]xlsx.Cell = undefined;
    var formulas: [data_cols + formula_cols]?[]const u8 = undefined;

    var i: u32 = 0;
    while (i < g.data_rows) : (i += 1) {
        const row = i + 2; // one header row above
        cells[0] = .{ .string = rawName(&name_buf, i) };
        cells[1] = .{ .string = path(&path_buf, i) };
        cells[2] = .{ .string = amount(&amt_buf, i) };
        formulas[0] = null;
        formulas[1] = null;
        formulas[2] = null;

        // A deliberately wrong cache in every formula cell, for the
        // same reason the other fixtures carry one: the recalc then
        // writes every formula cell and the stage/patch phase is part
        // of the measurement. A poison STRING, not −1: five of the six
        // results are text, and a sentinel a formula could produce is
        // a cell the recalc silently leaves out (the sixth, G, parses
        // spellings that are all non-negative, so the string is wrong
        // there too).
        for (cells[data_cols..]) |*c| c.* = .{ .string = "#STALE" };
        formulas[3] = try std.fmt.bufPrint(&fbuf[0], "PROPER(TRIM(A{d}))", .{row});
        formulas[4] = try std.fmt.bufPrint(&fbuf[1], "TEXTBEFORE(B{d},\",\")", .{row});
        formulas[5] = try std.fmt.bufPrint(&fbuf[2], "TEXTAFTER(B{d},\",\",-1,1)", .{row});
        formulas[6] = try std.fmt.bufPrint(&fbuf[3], "NUMBERVALUE(C{d})", .{row});
        formulas[7] = try std.fmt.bufPrint(&fbuf[4], "FIXED(G{d},1)", .{row});
        formulas[8] = try std.fmt.bufPrint(&fbuf[5], "DOLLAR(G{d})", .{row});
        try s.writeRowWithFormulas(&cells, &formulas);
    }
}

// ─── tests ───────────────────────────────────────────────────────────
//
// Same split as the other fixtures: determinism and topology run on
// the default test path at a size they can afford; the identity digest
// is verified by `zlsx-bench-recalc emit --workload text` in the
// ReleaseFast lane, where the baseline is measured anyway.

const testing = std.testing;

test "text-mix: the generator is deterministic" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    const sixty_four: Geometry = .{ .data_rows = 64 };
    const x = try bytes(a, io, sixty_four);
    defer a.free(x);
    const y = try bytes(a, io, sixty_four);
    defer a.free(y);
    try testing.expectEqualSlices(u8, x, y);

    var hx: [digest_len * 2]u8 = undefined;
    var hy: [digest_len * 2]u8 = undefined;
    try testing.expectEqualStrings(digestHex(x, &hx), digestHex(y, &hy));
}

test "text-mix: the topology is the one the table describes" {
    const a = testing.allocator;

    const g: Geometry = .{ .data_rows = 100 };
    try testing.expectEqual(@as(u64, 900), g.cells());
    try testing.expectEqual(@as(u64, 600), g.formulaCells());

    var w = xlsx.Writer.init(a);
    defer w.deinit();
    try build(&w, g);
    const body = w.sheets.items[0].body.items;

    // Six formulas per row, and not one more — a stray formula on a
    // data column would put the inputs inside the graph the bench
    // claims they are not in.
    try testing.expectEqual(@as(usize, 600), std.mem.count(u8, body, "<f>"));
    try testing.expectEqual(@as(usize, 100), std.mem.count(u8, body, "<f>PROPER(TRIM(A"));
    try testing.expectEqual(@as(usize, 100), std.mem.count(u8, body, "<f>TEXTBEFORE(B"));
    try testing.expectEqual(@as(usize, 100), std.mem.count(u8, body, "<f>TEXTAFTER(B"));
    try testing.expectEqual(@as(usize, 100), std.mem.count(u8, body, "<f>NUMBERVALUE(C"));
    try testing.expectEqual(@as(usize, 100), std.mem.count(u8, body, "<f>FIXED(G"));
    try testing.expectEqual(@as(usize, 100), std.mem.count(u8, body, "<f>DOLLAR(G"));
    // The two match modes really split one delimiter. The writer
    // escapes `"` as `&quot;` in element content, so the probes match
    // the emitted bytes, not the formula's spelling.
    try testing.expect(std.mem.indexOf(
        u8,
        body,
        "<f>TEXTBEFORE(B2,&quot;,&quot;)</f>",
    ) != null);
    try testing.expect(std.mem.indexOf(
        u8,
        body,
        "<f>TEXTAFTER(B2,&quot;,&quot;,-1,1)</f>",
    ) != null);
}
