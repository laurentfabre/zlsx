//! §9.1e's AST-MASS workload: very long formulas.
//!
//! Every fixture in the §9.1d matrix carries formulas of 12 to 45
//! characters — a handful of AST nodes each — so "per cell" and "per AST
//! node" are the same quantity there and nothing separates them. This
//! one holds the cell count and the edge count fixed and multiplies the
//! **program** behind each formula cell by two orders of magnitude:
//! `terms` products summed, ~500 AST nodes and ~1 KB of formula text per
//! formula cell.
//!
//! Topology:
//!
//! | Col | Content | Rows |
//! |---|---|---|
//! | A | an integer | data |
//! | B | an integer | data |
//! | C | `A…*1+B…*2+A…*3+…` — `terms` products, summed | data |
//!
//! **Two precedents, whatever `terms` is.** The graph deduplicates
//! edges, so C's 128 references collapse to A and B and the fixture's
//! edge density stays at 0.667 per cell while its AST mass runs 128×
//! ahead of every other fixture's. That is the isolation: if §9.1d's
//! per-cell figure is really per-cell, this fixture prices a cell like
//! the others; if a formula's parsed program is retained per cell, this
//! is where it becomes visible.
//!
//! The chain is flat — a sum of products over two literals-bearing
//! cells, no nesting — so `max_fn_nesting` (64) and `max_parse_depth`
//! (256) are untouched, and at 128 terms the ~520 nodes and ~1 100
//! characters sit an order of magnitude below `max_ast_nodes` (16 384)
//! and well inside `max_formula_chars` (8 192). A fixture that refused
//! would measure nothing.

const std = @import("std");
const xlsx = @import("zlsx");

const Allocator = std.mem.Allocator;

/// Data columns per row: two integers.
pub const data_cols: u32 = 2;

/// Formula columns per row: the long one.
pub const formula_cols: u32 = 1;

/// Products summed in each formula at the identity size. 128 puts ~511
/// AST nodes behind every formula cell against the 4–14 the other
/// fixtures carry.
pub const default_terms: u32 = 128;

/// The generator's scratch buffer, and so the ceiling on a sweep: 512
/// terms at a six-digit row index is ~6 100 bytes, still inside the
/// parser's 8 192-character limit.
const formula_buf_bytes: usize = 16_384;

pub const Geometry = struct {
    /// Data rows, excluding the one header row.
    data_rows: u32,

    /// Products summed in each formula. A *knob*, for the same reason
    /// `bigtext`'s `string_bytes` is one: holding `data_rows` fixed
    /// while moving this moves the AST mass with the cell count, the
    /// row count, the formula count and the edge set unchanged, so the
    /// per-node cost is measured rather than left over.
    terms: u32 = default_terms,

    pub fn cells(self: Geometry) u64 {
        return @as(u64, self.data_rows) * (data_cols + formula_cols);
    }

    pub fn formulaCells(self: Geometry) u64 {
        return @as(u64, self.data_rows) * formula_cols;
    }

    /// AST nodes per formula, by the grammar: `terms` refs, `terms`
    /// coefficients, `terms` products and `terms − 1` sums.
    pub fn astNodesPerFormula(self: Geometry) u64 {
        return 4 * @as(u64, self.terms) - 1;
    }

    /// Total `<f>` payload bytes across the fixture. Summed by running
    /// the generator rather than by a digit-count formula, so it is the
    /// bytes actually emitted; §9.1e needs it as a *named term* to
    /// separate the source text a run retains from everything the parse
    /// produces out of it.
    pub fn formulaTextBytes(self: Geometry) u64 {
        var buf: [formula_buf_bytes]u8 = undefined;
        var total: u64 = 0;
        var i: u32 = 0;
        while (i < self.data_rows) : (i += 1) {
            total += longFormula(&buf, i + 2, self.terms).len;
        }
        return total;
    }
};

/// The small size: 3 000 cells.
pub const tiny: Geometry = .{ .data_rows = 1_000 };

/// The identity size the recorded §9.1e numbers bind to: 10 000 rows,
/// 30 000 cells, 10 000 long formulas.
pub const small: Geometry = .{ .data_rows = 10_000 };

/// SHA-256 of `bytes(gpa, io, small)`. Same contract as every other
/// fixture's identity digest.
pub const small_digest_sha256 =
    "719628e8ee8589267ee087013f1796afaa3f9fadab29a397b938db344143c02f";

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

/// Deterministic value spread over 0…999 — the same coprime multiplier
/// the criteria and text fixtures use.
fn valueOf(i: u32) u32 {
    return (i * 37) % 1_000;
}

/// `A{r}*1+B{r}*2+A{r}*3+…` — `terms` products of one of the row's two
/// inputs by its 1-based position, joined by `+`. Alternating the column
/// keeps both inputs in the program; the coefficients keep the terms
/// textually distinct, so nothing downstream can fold them.
pub fn longFormula(buf: *[formula_buf_bytes]u8, r: u32, terms: u32) []const u8 {
    var used: usize = 0;
    var k: u32 = 1;
    while (k <= terms) : (k += 1) {
        const col: u8 = if (k % 2 == 1) 'A' else 'B';
        const part = (if (k == 1)
            std.fmt.bufPrint(buf[used..], "{c}{d}*{d}", .{ col, r, k })
        else
            std.fmt.bufPrint(buf[used..], "+{c}{d}*{d}", .{ col, r, k })) catch unreachable;
        used += part.len;
    }
    return buf[0..used];
}

fn build(w: *xlsx.Writer, g: Geometry) !void {
    std.debug.assert(g.terms >= 1);

    var s = try w.addSheet("LONGFORM");
    try s.writeRow(&.{
        .{ .string = "a" }, .{ .string = "b" }, .{ .string = "sum" },
    });

    var fbuf: [formula_buf_bytes]u8 = undefined;
    var cells: [data_cols + formula_cols]xlsx.Cell = undefined;
    var formulas: [data_cols + formula_cols]?[]const u8 = undefined;

    var i: u32 = 0;
    while (i < g.data_rows) : (i += 1) {
        const r = i + 2; // one header row above
        cells[0] = .{ .integer = @intCast(valueOf(i)) };
        cells[1] = .{ .integer = @intCast(valueOf(i + 1)) };
        // The usual deliberately wrong cache: the recalc writes every
        // formula cell, so staging is measured and not skipped.
        cells[2] = .{ .integer = 0 };
        formulas[0] = null;
        formulas[1] = null;
        formulas[2] = longFormula(&fbuf, r, g.terms);
        try s.writeRowWithFormulas(&cells, &formulas);
    }
}

// ─── tests ───────────────────────────────────────────────────────────

const testing = std.testing;

test "longform-mix: the generator is deterministic" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    const g: Geometry = .{ .data_rows = 40 };
    const x = try bytes(a, io, g);
    defer a.free(x);
    const y = try bytes(a, io, g);
    defer a.free(y);
    try testing.expectEqualSlices(u8, x, y);

    var hx: [digest_len * 2]u8 = undefined;
    var hy: [digest_len * 2]u8 = undefined;
    try testing.expectEqualStrings(digestHex(x, &hx), digestHex(y, &hy));
}

test "longform-mix: the formula is long, flat, and over two precedents" {
    var fbuf: [formula_buf_bytes]u8 = undefined;
    const f = longFormula(&fbuf, 2, default_terms);

    // `terms` products means `terms - 1` joining `+` signs and `terms`
    // `*` signs; the count IS the AST mass claim, so it is asserted.
    try testing.expectEqual(@as(usize, default_terms), std.mem.count(u8, f, "*"));
    try testing.expectEqual(@as(usize, default_terms - 1), std.mem.count(u8, f, "+"));
    try testing.expect(std.mem.startsWith(u8, f, "A2*1+B2*2+A2*3"));
    try testing.expect(std.mem.endsWith(u8, f, "+B2*128"));
    // Flat: no parentheses and no call, so no nesting limit is in play.
    try testing.expectEqual(@as(usize, 0), std.mem.count(u8, f, "("));
    // Two distinct precedents however many terms there are.
    try testing.expectEqual(@as(usize, default_terms / 2), std.mem.count(u8, f, "A2*"));
    try testing.expectEqual(@as(usize, default_terms / 2), std.mem.count(u8, f, "B2*"));
    // Long enough to be the point, short enough to parse: the parser's
    // char limit is 8 192. The exact length is arithmetic, not a range —
    // `A2*1` is 4 bytes and every later term is `+`, a column, `2`, `*`
    // and the coefficient's digits, so 4 + 8×5 + 90×6 + 29×7 = 787 at
    // row 2. The fixture's *average* is 1 157 B because its row indices
    // run to five digits.
    try testing.expectEqual(@as(usize, 787), f.len);

    // The widest point of §9.1e's sweep — 512 terms at a six-digit row —
    // still fits the scratch buffer and the parser's limit. A fixture
    // that refused would measure nothing.
    const widest = longFormula(&fbuf, 999_999, 512);
    try testing.expect(widest.len < formula_buf_bytes);
    try testing.expect(widest.len < 8_192);
    try testing.expectEqual(@as(usize, 512), std.mem.count(u8, widest, "*"));
}

test "longform-mix: the AST knob moves the program and nothing else" {
    const a = testing.allocator;

    // The mirror of `bigtext`'s density test, and load-bearing for the
    // same reason: at a fixed row count, `terms` must move the AST mass
    // and leave the cells, the formula count and the precedent set
    // alone, or the sweep's ΔRSS would have more than one cause.
    const lo: Geometry = .{ .data_rows = 20, .terms = 32 };
    const hi: Geometry = .{ .data_rows = 20, .terms = 512 };
    try testing.expectEqual(lo.cells(), hi.cells());
    try testing.expectEqual(lo.formulaCells(), hi.formulaCells());
    try testing.expectEqual(@as(u64, 127), lo.astNodesPerFormula());
    try testing.expectEqual(@as(u64, 2_047), hi.astNodesPerFormula());

    var w = xlsx.Writer.init(a);
    defer w.deinit();
    try build(&w, hi);
    const body = w.sheets.items[0].body.items;
    // Still one formula per row, still two precedents in it.
    try testing.expectEqual(@as(usize, 20), std.mem.count(u8, body, "<f>"));
    try testing.expect(std.mem.indexOf(u8, body, "+B2*512</f>") != null);
}

test "longform-mix: the topology is the one the table describes" {
    const a = testing.allocator;

    const g: Geometry = .{ .data_rows = 20 };
    try testing.expectEqual(@as(u64, 60), g.cells());
    try testing.expectEqual(@as(u64, 20), g.formulaCells());

    var w = xlsx.Writer.init(a);
    defer w.deinit();
    try build(&w, g);
    const body = w.sheets.items[0].body.items;

    // One formula per row, in column C and nowhere else.
    try testing.expectEqual(@as(usize, 20), std.mem.count(u8, body, "<f>"));
    try testing.expectEqual(@as(usize, 20), std.mem.count(u8, body, "<f>A"));
    try testing.expect(std.mem.indexOf(u8, body, "r=\"C2\"") != null);
    try testing.expect(std.mem.indexOf(u8, body, "<f>A2*1+B2*2+") != null);
}
