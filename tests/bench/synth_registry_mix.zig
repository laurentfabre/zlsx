//! §9's mixed full-registry bench workload: the M9d fixture, and the
//! ladder's last lane.
//!
//! The F1-mix fixture prices the graph, the criteria fixture prices the
//! aligned cursor, the TEXT fixture prices the text stack; this one
//! prices **registry breadth** — every data row carries twelve
//! formulas that together span the ladder's batches: F1a arithmetic and
//! logic, F1b's small-collection statistics shape (via M7b3's MEDIAN),
//! F1c text and dates, M8a's numfmt renderer under TEXT, F4a's TVM
//! closed forms and depreciation, and F4b's four engineering families
//! (CONVERT's unit table, a base conversion, a BIT* mask, and a
//! parse+format complex round trip). A regression in any of those
//! stacks — a unit resolver that starts allocating, a complex format
//! that re-parses, a TVM spelling that loses its `log1p` arm — shows up
//! here as a slope change before any correctness test notices.
//!
//! Every formula is row-local — no whole-column scans, no cross-row
//! chains — so the two sizes separate the **marginal per-row cost** of
//! the mixed registry from everything paid once, which is what §9 asks
//! this baseline to state.
//!
//! Topology (fixed, same discipline and same reasons as
//! `synth_f1_mix.zig`):
//!
//! | Col | Content | Rows |
//! |---|---|---|
//! | A | annual rate in basis points, 300…1000 | data |
//! | B | principal, 1 000…9 999 | data |
//! | C | integer 0…999 | data |
//! | D | date serial, 45 000…45 364 | data |
//! | E | mixed-case word | data |
//! | F | `ROUND(A…/10000*B…,2)` — F1a arithmetic | data |
//! | G | `IF(C…>500,"hi","lo")` — F1a logic | data |
//! | H | `PMT(A…/10000/12,120,-B…)` — F4a TVM | data |
//! | I | `SLN(B…,C…,10)` — F4a depreciation | data |
//! | J | `EOMONTH(D…,1)` — F1c serial dates | data |
//! | K | `TEXT(C…,"0.00")` — M8a numfmt | data |
//! | L | `CONVERT(C…,"mi","km")` — F4b unit table | data |
//! | M | `DEC2HEX(C…)` — F4b base conversion | data |
//! | N | `BITXOR(C…,255)` — F4b bit field | data |
//! | O | `IMABS(COMPLEX(A…/100,C…))` — F4b complex, nested | data |
//! | P | `MEDIAN(A…,B…,C…)` — F2 statistics | data |
//! | Q | `UPPER(E…)` — F1c text | data |

const std = @import("std");
const xlsx = @import("zlsx");

const Allocator = std.mem.Allocator;

/// Data columns per row: rate, principal, integer, serial, word.
pub const data_cols: u32 = 5;

/// Registry formulas per row: F…Q.
pub const formula_cols: u32 = 12;

/// Mixed-case words for column `E` — UPPER has to move something in
/// every cell.
const word_pool = [_][]const u8{
    "aLpHa", "bRaVo", "cHaRlIe", "dElTa", "eChO", "fOxTrOt", "gOlF", "hOtEl",
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

/// The identity size the recorded registry baseline binds to: 120 000
/// mixed formulas over 10 000 rows.
pub const small: Geometry = .{ .data_rows = 10_000 };

/// SHA-256 of `bytes(gpa, io, small)` — the same contract as
/// `synth_text_mix.small_digest_sha256`: this identifies the workload
/// the recorded baseline was measured on, and a mismatch means
/// re-measure, not "fix the writer".
pub const small_digest_sha256 =
    "957fb4d0b933aa46968a4682c05c9237438c9c6d176767a52ed16ae6dbfb834f";

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

/// Deterministic value spread over 0…999 — the criteria fixture's
/// coprime multiplier, reused.
fn valueOf(i: u32) u32 {
    return (i * 37) % 1_000;
}

fn build(w: *xlsx.Writer, g: Geometry) !void {
    var s = try w.addSheet("REGISTRY");
    try s.writeRow(&.{
        .{ .string = "rate_bp" }, .{ .string = "principal" }, .{ .string = "n" },
        .{ .string = "serial" },  .{ .string = "word" },      .{ .string = "f" },
        .{ .string = "g" },       .{ .string = "h" },         .{ .string = "i" },
        .{ .string = "j" },       .{ .string = "k" },         .{ .string = "l" },
        .{ .string = "m" },       .{ .string = "n2" },        .{ .string = "o" },
        .{ .string = "p" },       .{ .string = "q" },
    });

    var fbuf: [formula_cols][64]u8 = undefined;
    var cells: [data_cols + formula_cols]xlsx.Cell = undefined;
    var formulas: [data_cols + formula_cols]?[]const u8 = undefined;

    var i: u32 = 0;
    while (i < g.data_rows) : (i += 1) {
        const r = i + 2; // one header row above
        cells[0] = .{ .integer = @intCast(300 + (i % 8) * 100) };
        cells[1] = .{ .integer = @intCast(1_000 + (i * 37) % 9_000) };
        cells[2] = .{ .integer = @intCast(valueOf(i)) };
        cells[3] = .{ .integer = @intCast(45_000 + i % 365) };
        cells[4] = .{ .string = word_pool[i % word_pool.len] };
        for (formulas[0..data_cols]) |*f| f.* = null;

        // A deliberately wrong cache in every formula cell, for the
        // same reason the other fixtures carry one: the recalc then
        // writes every formula cell and the stage/patch phase is part
        // of the measurement. A poison STRING no formula in the mix
        // can produce, so a silently-skipped cell cannot hide.
        for (cells[data_cols..]) |*c| c.* = .{ .string = "#STALE" };
        formulas[5] = try std.fmt.bufPrint(&fbuf[0], "ROUND(A{d}/10000*B{d},2)", .{ r, r });
        formulas[6] = try std.fmt.bufPrint(&fbuf[1], "IF(C{d}>500,\"hi\",\"lo\")", .{r});
        formulas[7] = try std.fmt.bufPrint(&fbuf[2], "PMT(A{d}/10000/12,120,-B{d})", .{ r, r });
        formulas[8] = try std.fmt.bufPrint(&fbuf[3], "SLN(B{d},C{d},10)", .{ r, r });
        formulas[9] = try std.fmt.bufPrint(&fbuf[4], "EOMONTH(D{d},1)", .{r});
        formulas[10] = try std.fmt.bufPrint(&fbuf[5], "TEXT(C{d},\"0.00\")", .{r});
        formulas[11] = try std.fmt.bufPrint(&fbuf[6], "CONVERT(C{d},\"mi\",\"km\")", .{r});
        formulas[12] = try std.fmt.bufPrint(&fbuf[7], "DEC2HEX(C{d})", .{r});
        formulas[13] = try std.fmt.bufPrint(&fbuf[8], "BITXOR(C{d},255)", .{r});
        formulas[14] = try std.fmt.bufPrint(&fbuf[9], "IMABS(COMPLEX(A{d}/100,C{d}))", .{ r, r });
        formulas[15] = try std.fmt.bufPrint(&fbuf[10], "MEDIAN(A{d},B{d},C{d})", .{ r, r, r });
        formulas[16] = try std.fmt.bufPrint(&fbuf[11], "UPPER(E{d})", .{r});
        try s.writeRowWithFormulas(&cells, &formulas);
    }
}

// ─── tests ───────────────────────────────────────────────────────────
//
// Same split as the other fixtures: determinism and topology run on
// the default test path at a size they can afford; the identity digest
// is verified by `zlsx-bench-recalc emit --workload registry` in the
// ReleaseFast lane, where the baseline is measured anyway.

const testing = std.testing;

test "registry-mix: the generator is deterministic" {
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

test "registry-mix: the topology is the one the table describes" {
    const a = testing.allocator;

    const g: Geometry = .{ .data_rows = 100 };
    try testing.expectEqual(@as(u64, 1_700), g.cells());
    try testing.expectEqual(@as(u64, 1_200), g.formulaCells());

    var w = xlsx.Writer.init(a);
    defer w.deinit();
    try build(&w, g);
    const body = w.sheets.items[0].body.items;

    // Twelve formulas per row, and not one more — a stray formula on a
    // data column would put the inputs inside the graph the bench
    // claims they are not in.
    try testing.expectEqual(@as(usize, 1_200), std.mem.count(u8, body, "<f>"));
    try testing.expectEqual(@as(usize, 100), std.mem.count(u8, body, "<f>PMT(A"));
    try testing.expectEqual(@as(usize, 100), std.mem.count(u8, body, "<f>CONVERT(C"));
    try testing.expectEqual(@as(usize, 100), std.mem.count(u8, body, "<f>IMABS(COMPLEX(A"));
    try testing.expectEqual(@as(usize, 100), std.mem.count(u8, body, "<f>BITXOR(C"));
    // The writer escapes `"` as `&quot;` in element content, so the
    // probes match the emitted bytes, not the formula's spelling.
    try testing.expect(std.mem.indexOf(
        u8,
        body,
        "<f>CONVERT(C2,&quot;mi&quot;,&quot;km&quot;)</f>",
    ) != null);
    try testing.expect(std.mem.indexOf(
        u8,
        body,
        "<f>TEXT(C2,&quot;0.00&quot;)</f>",
    ) != null);
}
