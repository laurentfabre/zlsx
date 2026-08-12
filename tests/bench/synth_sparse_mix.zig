//! §9.1e's ROW-OVERHEAD workload: one cell per row.
//!
//! Every other fixture in the §9.1d matrix packs 3–10 cells into each
//! row, so a per-row cost — the row record, the row's index entry, the
//! `<row>` element and its attributes — is divided by 3 to 10 before it
//! reaches the B/cell figure. This one divides it by **one**. If the
//! per-cell claim §9.1d made is really about cells, this fixture prices
//! a cell the same as the others; if it is partly about rows wearing a
//! cell's name, this fixture is where that shows.
//!
//! Topology:
//!
//! | Col | Content | Rows |
//! |---|---|---|
//! | A | on even data rows, an integer | data/2 |
//! | A | on odd data rows, `A{r-1}*2+1` | data/2 |
//!
//! One column, and nothing else on the row. The formula rows carry a
//! deliberately wrong cache for the same reason every other fixture
//! does: the recalc then writes every formula cell, so the stage/patch
//! phase is measured rather than skipped.
//!
//! **The sizes are chosen to land on the f1 matrix's cell counts, not on
//! its row counts.** 10 000 / 100 000 / 400 000 rows here are 10 000 /
//! 100 000 / 400 000 cells, which is exactly f1/1k, f1/10k and f1/40k —
//! so the comparison is at equal cells with ten times the rows, and the
//! difference between the two columns is the row term with everything
//! else held as close as two different fixtures can hold it.
//!
//! The chain is two deep by construction (a literal feeds one formula),
//! so nothing here approaches an evaluation-depth limit however many
//! rows are asked for.

const std = @import("std");
const xlsx = @import("zlsx");

const Allocator = std.mem.Allocator;

/// Cells per data row at the identity size. One, which is the whole
/// point of the fixture.
pub const default_cols: u32 = 1;

/// The widest row a sweep may ask for: A…J, the F1 mix's own width.
pub const max_cols: u32 = 10;

pub const Geometry = struct {
    /// Data rows, excluding the one header row.
    data_rows: u32,

    /// Cells per data row. A *knob*, and the instrument this fixture
    /// exists for: at a fixed **cell** count, moving `cols` moves the
    /// row count and nothing else — same formula fraction, same formula
    /// text, same precedent structure, same literal values. 400 000
    /// cells as 400 000 × 1 against 40 000 × 10 differ in rows alone, so
    /// the row term is measured rather than left over.
    cols: u32 = default_cols,

    pub fn cells(self: Geometry) u64 {
        return @as(u64, self.data_rows) * self.cols;
    }

    /// Odd rows carry the formulas, so half the cells — rounded down,
    /// because an odd `data_rows` ends on a literal row.
    pub fn formulaCells(self: Geometry) u64 {
        return @as(u64, self.data_rows) / 2 * self.cols;
    }
};

/// The small size: 10 000 cells, matching f1/1k's cell count.
pub const tiny: Geometry = .{ .data_rows = 10_000 };

/// The identity size the recorded §9.1e numbers bind to: 100 000 cells
/// in 100 000 rows, against f1_named's 100 000 cells in 10 000.
pub const small: Geometry = .{ .data_rows = 100_000 };

/// SHA-256 of `bytes(gpa, io, small)` — same contract as every other
/// fixture's identity digest: it names the workload the recorded numbers
/// were measured on, and a mismatch means re-measure, not "fix the
/// writer".
pub const small_digest_sha256 =
    "31d0342e6bb4d8a9f63efba8606c01febedd291cb31843b4310813a0838737f2";

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

/// Deterministic value spread over 0…999 — the coprime multiplier the
/// criteria and text fixtures already use, reused so the input column
/// varies the same way theirs do.
fn valueOf(i: u32) u32 {
    return (i * 37) % 1_000;
}

fn build(w: *xlsx.Writer, g: Geometry) !void {
    std.debug.assert(g.cols >= 1);
    std.debug.assert(g.cols <= max_cols);

    var s = try w.addSheet("SPARSE");
    var header: [max_cols]xlsx.Cell = @splat(.{ .string = "n" });
    try s.writeRow(header[0..g.cols]);

    var fbuf: [max_cols][24]u8 = undefined;
    var cells: [max_cols]xlsx.Cell = undefined;
    var formulas: [max_cols]?[]const u8 = undefined;

    var i: u32 = 0;
    while (i < g.data_rows) : (i += 1) {
        const r = i + 2; // one header row above
        // Whole rows alternate rather than columns within a row: the
        // formula fraction is then exactly one half at every `cols`,
        // and every formula reads the cell directly above it in its own
        // column — the same one-precedent shape whatever the width.
        var j: u32 = 0;
        while (j < g.cols) : (j += 1) {
            if (i % 2 == 0) {
                cells[j] = .{ .integer = @intCast(valueOf(i +% j)) };
                formulas[j] = null;
            } else {
                cells[j] = .{ .integer = 0 };
                formulas[j] = try std.fmt.bufPrint(
                    &fbuf[j],
                    "{c}{d}*2+1",
                    .{ @as(u8, 'A') + @as(u8, @intCast(j)), r - 1 },
                );
            }
        }
        try s.writeRowWithFormulas(cells[0..g.cols], formulas[0..g.cols]);
    }
}

// ─── tests ───────────────────────────────────────────────────────────
//
// Same split as the other fixtures: determinism and topology on the
// default test path at a size it can afford; the identity digest is
// gated by `zlsx-bench-recalc emit --workload sparse` in the release
// lane, where the numbers are measured anyway.

const testing = std.testing;

test "sparse-mix: the generator is deterministic" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    const g: Geometry = .{ .data_rows = 64 };
    const x = try bytes(a, io, g);
    defer a.free(x);
    const y = try bytes(a, io, g);
    defer a.free(y);
    try testing.expectEqualSlices(u8, x, y);

    var hx: [digest_len * 2]u8 = undefined;
    var hy: [digest_len * 2]u8 = undefined;
    try testing.expectEqualStrings(digestHex(x, &hx), digestHex(y, &hy));
}

test "sparse-mix: one cell per row, half of them formulas" {
    const a = testing.allocator;

    const g: Geometry = .{ .data_rows = 100 };
    try testing.expectEqual(@as(u64, 100), g.cells());
    try testing.expectEqual(@as(u64, 50), g.formulaCells());

    var w = xlsx.Writer.init(a);
    defer w.deinit();
    try build(&w, g);
    const body = w.sheets.items[0].body.items;

    // 101 rows (header + data), one `<c` each and not one more: a second
    // cell anywhere would make this a two-column fixture measuring a
    // different thing than its name.
    try testing.expectEqual(@as(usize, 101), std.mem.count(u8, body, "<row "));
    try testing.expectEqual(@as(usize, 101), std.mem.count(u8, body, "<c "));
    try testing.expectEqual(@as(usize, 50), std.mem.count(u8, body, "<f>"));
    // Every formula reads the literal one row above it.
    try testing.expect(std.mem.indexOf(u8, body, "<f>A2*2+1</f>") != null);
    try testing.expect(std.mem.indexOf(u8, body, "<f>A4*2+1</f>") != null);
    // …and the last data row (i = 99, odd) is a formula over row 100.
    try testing.expect(std.mem.indexOf(u8, body, "<f>A100*2+1</f>") != null);
    // No cell lands outside column A.
    try testing.expectEqual(@as(usize, 0), std.mem.count(u8, body, "r=\"B"));
}

test "sparse-mix: the width knob moves rows and nothing else" {
    const a = testing.allocator;

    // §9.1e's row-term measurement rests on this: 200 cells as 200 × 1
    // and as 20 × 10 must agree on cells, on formula cells and on the
    // shape of every formula, and differ only in how many rows carry
    // them. Anything else moving would give ΔRSS a second cause.
    const tall: Geometry = .{ .data_rows = 200, .cols = 1 };
    const wide: Geometry = .{ .data_rows = 20, .cols = 10 };
    try testing.expectEqual(tall.cells(), wide.cells());
    try testing.expectEqual(tall.formulaCells(), wide.formulaCells());

    var w_tall = xlsx.Writer.init(a);
    defer w_tall.deinit();
    try build(&w_tall, tall);
    var w_wide = xlsx.Writer.init(a);
    defer w_wide.deinit();
    try build(&w_wide, wide);

    const body_tall = w_tall.sheets.items[0].body.items;
    const body_wide = w_wide.sheets.items[0].body.items;
    try testing.expectEqual(
        std.mem.count(u8, body_tall, "<f>"),
        std.mem.count(u8, body_wide, "<f>"),
    );
    try testing.expectEqual(@as(usize, 201), std.mem.count(u8, body_tall, "<row "));
    try testing.expectEqual(@as(usize, 21), std.mem.count(u8, body_wide, "<row "));
    // Same one-precedent formula shape at both widths, in each column.
    try testing.expect(std.mem.indexOf(u8, body_wide, "<f>A2*2+1</f>") != null);
    try testing.expect(std.mem.indexOf(u8, body_wide, "<f>J2*2+1</f>") != null);
}
