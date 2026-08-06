//! §9's criteria bench workload: the M7b2 fixture.
//!
//! The F1-mix fixture prices the graph; this one prices the **aligned
//! cursor**. Every report formula scans whole columns (`A:A`, `B:B`,
//! `C:C`) of a sheet whose stored rows stop far above row 1 048 576, so
//! what is being timed is exactly §5.6a's run-based iteration: stored
//! cells cost per cell, and the million-row tail costs one blank run.
//! A cursor that quietly degraded to a per-coordinate walk would show
//! up here as a ~100× cliff before any correctness test noticed.
//!
//! Topology (fixed, same discipline and same reasons as
//! `synth_f1_mix.zig` — a bench whose graph drifts measures a different
//! question each time):
//!
//! | Col | Content | Rows |
//! |---|---|---|
//! | A | category string, 16-entry pool | data |
//! | B | region string, 4-entry pool | data |
//! | C | integer value, 0…999 | data |
//! | D | `SUMIFS(C:C,A:A,"cat-…")` | report |
//! | E | `COUNTIFS(A:A,"cat-…")` | report |
//! | F | `SUMIFS(C:C,A:A,"cat-…",B:B,"r-…")` | report |
//! | G | `COUNTIFS(A:A,"cat-…",B:B,"r-…",C:C,">500")` | report |
//! | H | `AVERAGEIFS(C:C,A:A,"cat-…")` | report |
//! | I | `MINIFS(C:C,B:B,"r-…")` | report |
//! | J | `MAXIFS(C:C,B:B,"r-…")` | report |
//! | K | `SUMIF(A:A,"cat-…",C:C)` | report |
//!
//! `K` is deliberate: the single-criterion door (M4e) over the same
//! column as the family (M7b2), so a regression in one shape and not
//! the other is visible as such.
//!
//! The report block is a FIXED 64 rows whatever the data size. That is
//! what makes two sizes informative: 512 formula cells scanning 1 000
//! rows against the same 512 scanning 10 000 separates the per-stored-
//! cell cost from everything paid per formula, which one size cannot.

const std = @import("std");
const xlsx = @import("zlsx");

const Allocator = std.mem.Allocator;

/// Data columns per data row: category, region, value.
pub const data_cols: u32 = 3;

/// Report formulas per report row: D…K.
pub const report_cols: u32 = 8;

/// Report rows — fixed across sizes on purpose (see the docstring).
pub const report_rows: u32 = 64;

/// Distinct categories column `A` cycles through. 16 over the report's
/// 64 rows means every criterion resolves to a nonempty match set, so
/// `AVERAGEIFS` never divides by an empty one.
pub const category_pool_len: u32 = 16;

/// Distinct regions column `B` cycles through.
pub const region_pool_len: u32 = 4;

pub const Geometry = struct {
    /// Data rows, excluding the one header row.
    data_rows: u32,

    pub fn cells(self: Geometry) u64 {
        return @as(u64, self.data_rows) * data_cols +
            @as(u64, report_rows) * report_cols;
    }

    pub fn formulaCells(self: Geometry) u64 {
        _ = self;
        return @as(u64, report_rows) * report_cols;
    }
};

/// The smaller size: same 512 formulas, a tenth of the stored rows.
pub const tiny: Geometry = .{ .data_rows = 1_000 };

/// The identity size the recorded criteria baseline binds to: 512
/// whole-column formulas over 10 000 stored rows.
pub const small: Geometry = .{ .data_rows = 10_000 };

/// SHA-256 of `bytes(gpa, io, small)` — the same contract as
/// `synth_f1_mix.named_digest_sha256`: this identifies the workload the
/// recorded criteria baseline was measured on, and a mismatch means
/// re-measure, not "fix the writer".
pub const small_digest_sha256 =
    "31dbe2bd812f8426023217c11855b83b3e1ae6dffedbc8fe3bb74d0a294e74e2";

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

fn category(buf: *[8]u8, i: u32) []const u8 {
    return std.fmt.bufPrint(buf, "cat-{d:0>2}", .{i % category_pool_len}) catch unreachable;
}

fn region(buf: *[4]u8, i: u32) []const u8 {
    return std.fmt.bufPrint(buf, "r-{d}", .{i % region_pool_len}) catch unreachable;
}

/// Deterministic value spread over 0…999 — coprime multiplier, so every
/// category sees the full range and `>500` splits each roughly in half.
fn valueOf(i: u32) u32 {
    return (i * 37) % 1_000;
}

fn build(w: *xlsx.Writer, g: Geometry) !void {
    // The report needs a full cycle of criteria under it.
    std.debug.assert(g.data_rows >= report_rows);

    var s = try w.addSheet("CRIT");
    try s.writeRow(&.{
        .{ .string = "category" }, .{ .string = "region" }, .{ .string = "value" },
    });

    var cat_buf: [8]u8 = undefined;
    var reg_buf: [4]u8 = undefined;
    var fbuf: [report_cols][96]u8 = undefined;
    var cells: [data_cols + report_cols]xlsx.Cell = undefined;
    var formulas: [data_cols + report_cols]?[]const u8 = undefined;

    var i: u32 = 0;
    while (i < g.data_rows) : (i += 1) {
        const cat = category(&cat_buf, i);
        const reg = region(&reg_buf, i);
        cells[0] = .{ .string = cat };
        cells[1] = .{ .string = reg };
        cells[2] = .{ .integer = valueOf(i) };
        formulas[0] = null;
        formulas[1] = null;
        formulas[2] = null;

        if (i >= report_rows) {
            try s.writeRowWithFormulas(cells[0..data_cols], formulas[0..data_cols]);
            continue;
        }

        // A deliberately wrong cache in every report cell, for the same
        // reason the F1 mix carries one: the recalc then writes all 512,
        // and the stage/patch phase is part of the measurement. −1 and
        // not 0, because every fold here is over non-negative values —
        // `MINIFS` over the region that holds the value 0 really is 0,
        // and a cache that is accidentally right is a cell the recalc
        // silently leaves out of the measurement.
        for (cells[data_cols..]) |*c| c.* = .{ .integer = -1 };
        formulas[3] = try std.fmt.bufPrint(&fbuf[0], "SUMIFS(C:C,A:A,\"{s}\")", .{cat});
        formulas[4] = try std.fmt.bufPrint(&fbuf[1], "COUNTIFS(A:A,\"{s}\")", .{cat});
        formulas[5] = try std.fmt.bufPrint(&fbuf[2], "SUMIFS(C:C,A:A,\"{s}\",B:B,\"{s}\")", .{ cat, reg });
        formulas[6] = try std.fmt.bufPrint(&fbuf[3], "COUNTIFS(A:A,\"{s}\",B:B,\"{s}\",C:C,\">500\")", .{ cat, reg });
        formulas[7] = try std.fmt.bufPrint(&fbuf[4], "AVERAGEIFS(C:C,A:A,\"{s}\")", .{cat});
        formulas[8] = try std.fmt.bufPrint(&fbuf[5], "MINIFS(C:C,B:B,\"{s}\")", .{reg});
        formulas[9] = try std.fmt.bufPrint(&fbuf[6], "MAXIFS(C:C,B:B,\"{s}\")", .{reg});
        formulas[10] = try std.fmt.bufPrint(&fbuf[7], "SUMIF(A:A,\"{s}\",C:C)", .{cat});
        try s.writeRowWithFormulas(&cells, &formulas);
    }
}

// ─── tests ───────────────────────────────────────────────────────────
//
// Same split as the F1 mix: determinism and topology run on the default
// test path at a size it can afford; the identity digest is verified by
// `zlsx-bench-recalc emit --workload criteria` in the ReleaseFast lane,
// where the baseline is measured anyway.

const testing = std.testing;

test "criteria-mix: the generator is deterministic" {
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

test "criteria-mix: the topology is the one the table describes" {
    const a = testing.allocator;

    const g: Geometry = .{ .data_rows = 100 };
    try testing.expectEqual(@as(u64, 300 + 512), g.cells());
    try testing.expectEqual(@as(u64, 512), g.formulaCells());

    var w = xlsx.Writer.init(a);
    defer w.deinit();
    try build(&w, g);
    const body = w.sheets.items[0].body.items;

    // One `<f>` per report cell, and not one more: a stray formula on a
    // data column would put the scanned columns inside the graph the
    // bench claims they are not in.
    try testing.expectEqual(@as(usize, 512), std.mem.count(u8, body, "<f>"));
    // Every formula scans whole columns — the fixture's entire point.
    // The two SUMIFS columns (one criterion and two) share a prefix, as
    // do the two COUNTIFS columns, so those counts are 128.
    try testing.expectEqual(@as(usize, 128), std.mem.count(u8, body, "<f>SUMIFS(C:C,A:A,"));
    try testing.expectEqual(@as(usize, 128), std.mem.count(u8, body, "<f>COUNTIFS(A:A,"));
    try testing.expectEqual(@as(usize, 64), std.mem.count(u8, body, "<f>AVERAGEIFS(C:C,"));
    try testing.expectEqual(@as(usize, 64), std.mem.count(u8, body, "<f>MINIFS(C:C,"));
    try testing.expectEqual(@as(usize, 64), std.mem.count(u8, body, "<f>MAXIFS(C:C,"));
    try testing.expectEqual(@as(usize, 64), std.mem.count(u8, body, "<f>SUMIF(A:A,"));
    // The three-criteria row really carries three pairs. The writer
    // escapes `"` as `&quot;` and `>` as `&gt;` in element content, so
    // the probe matches the emitted bytes, not the formula's spelling.
    try testing.expect(std.mem.indexOf(
        u8,
        body,
        "<f>COUNTIFS(A:A,&quot;cat-00&quot;,B:B,&quot;r-0&quot;,C:C,&quot;&gt;500&quot;)</f>",
    ) != null);
    // The report stops at its fixed 64 rows; the data continues alone.
    try testing.expectEqual(@as(usize, 64), std.mem.count(u8, body, "<f>SUMIF(A:A,&quot;cat-"));
}
