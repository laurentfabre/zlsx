//! §9's named bench workload: the F1-mix fixture (M5d3).
//!
//! §9 makes the workloads **committed artifacts** — "a checked-in
//! generator script producing fixed-topology workbooks (dependency
//! fan-in/out recorded, formula mix per the F1 profile, shared-string
//! density, archive size)". This is that generator, and the reason it is
//! a generator rather than a checked-in `.xlsx` is the same reason
//! `synth_100k_x_10.zig` is: a hundred thousand cells of committed
//! binary is a repository cost paid on every clone, and the bytes are a
//! pure function of the code that emits them.
//!
//! Which makes the digest the actual artifact. §9's absolute ceilings
//! (evaluate ≤ 500 ms, end-to-end ≤ 1 s, peak RSS ≤ 3× model bytes) bind
//! to **one named workload**, so "the fixture" has to mean specific
//! bytes rather than "whatever the generator produces today".
//! `named_digest_sha256` is those bytes. It is not a correctness
//! assertion about the writer: if a legitimate writer change moves it,
//! the workload moved, and the recorded baseline must be re-measured
//! with it — the digest and the numbers in §9 are one artifact, not two.
//!
//! Topology (fixed, and fixed on purpose — §5.6a's cost model is a
//! function of the graph, so a bench whose graph drifts is measuring a
//! different question each time):
//!
//! | Col | Content | F1 batch | Fan-in |
//! |---|---|---|---|
//! | A | integer input | — | 0 |
//! | B | shared string from a 64-entry pool | — | 0 |
//! | C | `A{r}*2+1` | operators (M4c) | 1 |
//! | D | `ROUND(SQRT(ABS(C{r})),3)` | F1a-2 (M4d) | 1 |
//! | E | `IF(D{r}>10,C{r},A{r})` | F1a-1 (M4c) | 3 |
//! | F | `SUM(A{r-4}:A{r})` | F1b (M4e) | 5 (cross-row) |
//! | G | `MAX(MIN(F{r},1000),0)[+G{r-1}]` | F1b (M4e) | 1–2 (chained) |
//! | H | `LEN(UPPER(TRIM(B{r})))` | F1c-text (M4f) | 1 |
//! | I | `YEAR(EOMONTH(DATE(2020,MOD(A{r},12)+1,1),0))` | F1c-date (M4g) | 1 |
//! | J | `E{r}+G{r}+H{r}+I{r}` | operators | 4 |
//!
//! Fan-out: `A` is read by C, E, F and I (4); `F` by G alone; `G` by J
//! and by the next row's `G`. The `G` chain resets every
//! `chain_block_rows` rows, so the longest cell→cell path is bounded at
//! 32 + 3 — deliberately well under §9's `max_eval_depth` of 512, since
//! a fixture that refused would measure nothing.
//!
//! Every formula returns a **number**, so every cached `<v>` is numeric
//! and the fixture never raises §5.7.2's `t="str"` cache-state question.
//! The text and date batches are still exercised — `LEN` and `YEAR`
//! consume their results — which is the point: the mix is about the
//! work, not about the result type.
//!
//! No leading or trailing whitespace in the string pool: an SST entry
//! with edge whitespace needs `xml:space="preserve"` to round-trip, and
//! a fixture that loses bytes on the way back in is a fixture measuring
//! a different workbook than the one it emitted. `TRIM` still runs; it
//! just has nothing to strip.

const std = @import("std");
const xlsx = @import("zlsx");

const Allocator = std.mem.Allocator;

/// Columns per data row. Ten, and the table above is the reason for each.
pub const cols: u32 = 10;

/// Formula columns per data row: C…J.
pub const formula_cols: u32 = 8;

/// `G`'s chain length before it restarts. Bounded so the deepest
/// cell→cell path stays far below §9's `max_eval_depth`.
pub const chain_block_rows: u32 = 32;

/// Distinct strings column `B` draws from. 64 over `data_rows` entries
/// is the shared-string density §9 asks to be recorded: 0.64 % unique at
/// the named size.
pub const string_pool_len: u32 = 64;

/// `SUM(A{r-4}:A{r})` — five-wide backward window, so the graph carries
/// cross-row edges rather than being a forest of independent rows.
pub const window_rows: u32 = 5;

pub const Geometry = struct {
    /// Data rows, excluding the one header row.
    data_rows: u32,

    pub fn cells(self: Geometry) u64 {
        return @as(u64, self.data_rows) * cols;
    }

    pub fn formulaCells(self: Geometry) u64 {
        return @as(u64, self.data_rows) * formula_cols;
    }
};

/// The second size. Its only job is to make the fixed costs visible:
/// one measurement cannot separate "opens the archive" from "evaluates
/// the graph", and two an order of magnitude apart can.
pub const small: Geometry = .{ .data_rows = 1_000 };

/// The cheapest of the three, and the one CI can afford to pair with
/// `small` on every pull request. Two points a decade apart is the
/// minimum that separates "opens the archive" from "evaluates the
/// graph"; one point cannot.
pub const tiny: Geometry = .{ .data_rows = 100 };

/// §9's named workload: 10 000 × 10 = **100 000 cells**, 80 000 of them
/// formulas, plus a ten-cell header row.
pub const named: Geometry = .{ .data_rows = 10_000 };

/// SHA-256 of `bytes(gpa, io, named)`. See the module docstring: this
/// identifies the workload the recorded baseline was measured on, and a
/// mismatch means re-measure, not "fix the writer".
pub const named_digest_sha256 =
    "b2b42c0b67703a8a4935ddc525213ebeabf0568507bedded120ac23c32778ad0";

pub const digest_len = std.crypto.hash.sha2.Sha256.digest_length;

/// Serialize the fixture into memory. Uses `Writer.saveToOwnedBuffer`
/// (M5c) rather than a temp file so the digest below is a property of
/// the archive and not of a filesystem round-trip.
pub fn bytes(gpa: Allocator, io: std.Io, g: Geometry) ![]u8 {
    var w = xlsx.Writer.init(gpa);
    defer w.deinit();
    try build(&w, g);
    return w.saveToOwnedBuffer(gpa, io);
}

// There is deliberately no `emit(path)` alongside `bytes`. Every caller
// needs the digest of what it wrote, so every caller needs the bytes in
// hand — a path-only variant would be a second serialization the digest
// gate could not see, and Zig only analyses what something references,
// so an unused one would rot unnoticed.

/// Lowercase hex SHA-256 of `data`, into a caller-provided buffer.
pub fn digestHex(data: []const u8, out: *[digest_len * 2]u8) []const u8 {
    var raw: [digest_len]u8 = undefined;
    std.crypto.hash.sha2.Sha256.hash(data, &raw, .{});
    return std.fmt.bufPrint(out, "{x}", .{&raw}) catch unreachable;
}

/// One entry of the string pool. Fixed width and fixed alphabet, so
/// `LEN(UPPER(TRIM(B)))` is constant across rows and the text batch's
/// cost does not vary with the row index.
fn poolEntry(buf: *[24]u8, i: u32) []const u8 {
    return std.fmt.bufPrint(buf, "item-{d:0>2}-abcdefghij", .{i}) catch unreachable;
}

fn build(w: *xlsx.Writer, g: Geometry) !void {
    std.debug.assert(g.data_rows >= 1);

    var s = try w.addSheet("F1");
    try s.writeRow(&.{
        .{ .string = "n" }, .{ .string = "label" }, .{ .string = "c" },
        .{ .string = "d" }, .{ .string = "e" },     .{ .string = "f" },
        .{ .string = "g" }, .{ .string = "h" },     .{ .string = "i" },
        .{ .string = "j" },
    });

    // Reused scratch: the hot loop allocates nothing beyond what the
    // Writer's own buffers grow by.
    var pool_buf: [24]u8 = undefined;
    var fbuf: [formula_cols][96]u8 = undefined;
    var cells: [cols]xlsx.Cell = undefined;
    var formulas: [cols]?[]const u8 = undefined;

    var i: u32 = 0;
    while (i < g.data_rows) : (i += 1) {
        // 1-based sheet row: row 1 is the header, so data starts at 2.
        const r = i + 2;
        const win_start = if (r >= 2 + window_rows - 1) r - (window_rows - 1) else 2;
        const chain_head = i % chain_block_rows == 0;

        cells[0] = .{ .integer = @intCast(i) };
        cells[1] = .{ .string = poolEntry(&pool_buf, i % string_pool_len) };
        // A deliberately wrong cache in every formula cell. The recalc
        // therefore writes all 80 000 of them at the named size, which
        // is what makes the stage/patch phase part of the measurement
        // rather than a no-op the fixture happened to allow.
        for (cells[2..]) |*c| c.* = .{ .integer = 0 };

        formulas[0] = null;
        formulas[1] = null;
        formulas[2] = try std.fmt.bufPrint(&fbuf[0], "A{d}*2+1", .{r});
        formulas[3] = try std.fmt.bufPrint(&fbuf[1], "ROUND(SQRT(ABS(C{d})),3)", .{r});
        formulas[4] = try std.fmt.bufPrint(&fbuf[2], "IF(D{d}>10,C{d},A{d})", .{ r, r, r });
        formulas[5] = try std.fmt.bufPrint(&fbuf[3], "SUM(A{d}:A{d})", .{ win_start, r });
        formulas[6] = if (chain_head)
            try std.fmt.bufPrint(&fbuf[4], "MAX(MIN(F{d},1000),0)", .{r})
        else
            try std.fmt.bufPrint(&fbuf[4], "MAX(MIN(F{d},1000),0)+G{d}", .{ r, r - 1 });
        formulas[7] = try std.fmt.bufPrint(&fbuf[5], "LEN(UPPER(TRIM(B{d})))", .{r});
        formulas[8] = try std.fmt.bufPrint(&fbuf[6], "YEAR(EOMONTH(DATE(2020,MOD(A{d},12)+1,1),0))", .{r});
        formulas[9] = try std.fmt.bufPrint(&fbuf[7], "E{d}+G{d}+H{d}+I{d}", .{ r, r, r, r });

        try s.writeRowWithFormulas(&cells, &formulas);
    }
}

// ─── tests ───────────────────────────────────────────────────────────
//
// The named geometry's digest is NOT checked here: emitting 100 000
// cells and deflating a few megabytes in a Debug build would put a
// multi-second cost on every `zig build test` for a check that only the
// bench lane acts on. `zlsx-bench-recalc --emit` verifies it in
// ReleaseFast, which is where the baseline is measured anyway. What runs
// on the default path is what a code change can silently break: that the
// generator is deterministic, and that the topology is the one the
// docstring claims.

const testing = std.testing;

test "f1-mix: the generator is deterministic" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    const forty: Geometry = .{ .data_rows = 40 };
    const x = try bytes(a, io, forty);
    defer a.free(x);
    const y = try bytes(a, io, forty);
    defer a.free(y);
    try testing.expectEqualSlices(u8, x, y);

    // Same input, same digest — the property the named constant relies
    // on, checked at a size the default test path can afford.
    var hx: [digest_len * 2]u8 = undefined;
    var hy: [digest_len * 2]u8 = undefined;
    try testing.expectEqualStrings(digestHex(x, &hx), digestHex(y, &hy));
}

test "f1-mix: the topology is the one the table describes" {
    const a = testing.allocator;

    // Reads the sheet body straight off the Writer: the topology is a
    // property of what `build` emitted, and serializing it first would
    // put a deflate between the assertion and the thing asserted.
    const forty: Geometry = .{ .data_rows = 40 };
    try testing.expectEqual(@as(u64, 400), forty.cells());
    try testing.expectEqual(@as(u64, 320), forty.formulaCells());

    var w = xlsx.Writer.init(a);
    defer w.deinit();
    try build(&w, forty);
    const body = w.sheets.items[0].body.items;

    // One `<f>` per formula cell, and not one more: a stray formula on
    // an input column would change the graph the baseline describes.
    try testing.expectEqual(
        @as(usize, @intCast(forty.formulaCells())),
        std.mem.count(u8, body, "<f>"),
    );
    // The chain restarts every `chain_block_rows`, so a 40-row fixture
    // carries two heads — G2 and G34 — and 38 links. A head is the form
    // that ends at `,0)`; every other G continues with `+G{r-1}`.
    try testing.expectEqual(@as(usize, 40), std.mem.count(u8, body, "<f>MAX(MIN(F"));
    try testing.expectEqual(@as(usize, 2), std.mem.count(u8, body, ",1000),0)</f>"));
    try testing.expectEqual(@as(usize, 38), std.mem.count(u8, body, ",1000),0)+G"));
    try testing.expect(std.mem.indexOf(u8, body, "<f>MAX(MIN(F34,1000),0)</f>") != null);
    try testing.expect(std.mem.indexOf(u8, body, "<f>MAX(MIN(F35,1000),0)+G34</f>") != null);
    // The window clamps at the top of the sheet rather than running off
    // it: row 2 sums a single cell.
    try testing.expect(std.mem.indexOf(u8, body, "<f>SUM(A2:A2)</f>") != null);
    try testing.expect(std.mem.indexOf(u8, body, "<f>SUM(A2:A6)</f>") != null);
}
