//! `zlsx_recalc` — the third public module (§5.10).
//!
//! M5c of the tier-D1 ladder (`goal_formula.md`). This is the **shell**:
//! the module exists, imports `zlsx` and `zlsx_pkg`, and proves that a
//! producer's bytes can reach a consumer without a filesystem and without
//! a cycle. `writerSaveWithRecalc` — the composition that threads a
//! recalculation between those two halves — lands at M5d3 with the
//! `tests/consumer` dependency test.
//!
//! Why a third module rather than a function on one of the two
//! -------------------------------------------------------------
//! The composition needs BOTH halves at once: `zlsx.Writer` produces an
//! archive, `zlsx_pkg.Workbook` consumes one. Putting it in `zlsx` would
//! make the reader/writer depend on the package layer; putting it in
//! `zlsx_pkg` would make the package layer depend on the writer. Either
//! direction closes a loop that today runs one way — `zlsx_pkg → zlsx` —
//! and a loop is not a style preference here: `pkg/zip.zig` and
//! `pkg/fresh_emit.zig` are deliberately stdlib-only and take deflate as
//! a **function pointer** precisely so the graph stays a DAG.
//!
//! So the orchestrator sits *above* both, importing each as a named
//! module. Named modules are what makes that safe: a file belongs to
//! exactly one module, and `zlsx_pkg` already imports `zlsx` as the same
//! module object this file does — so `zlsx.Cell` here and `zlsx.Cell`
//! inside `Editor` are one type, not two structurally-identical ones.
//! The `comptime` block below is that claim, asserted rather than
//! assumed; the build's own acyclicity gate is the other half.
//!
//! Rooted at top-level `recalc/` rather than under `src/` or `pkg/` for
//! the same reason `unicode/` and `refs/` are: a module root inside
//! either tree would put this file in that tree's module as well, and a
//! file claimed by two modules is two distinct types.

const std = @import("std");
const assert = std.debug.assert;

pub const zlsx = @import("zlsx");
pub const pkg = @import("zlsx_pkg");

comptime {
    // The identity that makes composition possible at all. If the build
    // ever handed this module a different `zlsx` instance than the one
    // `zlsx_pkg` was built against, these would be two distinct types
    // with identical fields — and every error message about it would be
    // a wall of "expected xlsx.Cell, found xlsx.Cell".
    assert(@FieldType(pkg.Edit, "cell") == zlsx.Cell);
}

// ─── the two halves, under the names the composition will use ────

/// The producer. `saveToOwnedBuffer(allocator, io)` is what an
/// orchestrator hands onward (§5.10).
pub const Writer = zlsx.Writer;

/// The eager, whole-file reader.
pub const Book = zlsx.Book;

/// The typed package overlay. `openBuffer(allocator, io, bytes)` is the
/// other end of the handoff; the borrow ends when it returns.
pub const Workbook = pkg.Workbook;
pub const Worksheet = pkg.Worksheet;
pub const Editor = pkg.Editor;
pub const PartStore = pkg.PartStore;

/// M5b2's prepare/swap transaction, which M5d2's `recalculate()` and
/// M5d3's `writerSaveWithRecalc` both drive.
pub const recalc_txn = pkg.recalc_txn;

/// §9's cap on a serialized output archive, re-exported so a caller that
/// only imports the orchestrator can name the bound it is subject to.
pub const max_output_archive_bytes = zlsx.max_output_archive_bytes;

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

test {
    testing.refAllDecls(@This());
}

/// Populate `w` with a workbook that exercises the emitter's conditional
/// branches — SST with a dedup hit, a style, and a second sheet — so the
/// handoff below carries more than a one-cell archive.
fn buildWorkbook(w: *Writer) !void {
    const bold = try w.addStyle(.{ .font_bold = true });

    var s1 = try w.addSheet("Summary");
    try s1.writeRowStyled(&.{ .{ .string = "Region" }, .{ .string = "Units" } }, &.{ bold, bold });
    try s1.writeRow(&.{ .{ .string = "North" }, .{ .integer = 120 } });
    try s1.writeRow(&.{ .{ .string = "North" }, .{ .number = 7.5 } });

    var s2 = try w.addSheet("Notes");
    try s2.writeRow(&.{.{ .string = "second sheet" }});
}

test "composition: a Writer's bytes open as a Workbook, with no file at all" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var w = Writer.init(a);
    defer w.deinit();
    try buildWorkbook(&w);

    const bytes = try w.saveToOwnedBuffer(a, io);
    defer a.free(bytes);

    var wb = try Workbook.openBuffer(a, io, bytes);
    defer wb.deinit();

    // The whole point of the shell: this line reaches a type from
    // `zlsx_pkg` holding bytes produced by a type from `zlsx`, in one
    // compilation, with the import graph still a DAG.
    try testing.expectEqual(@as(u32, 2), wb.sheetCount());
    const ws = try wb.sheet(0);
    try testing.expectEqualStrings("Summary", ws.name());
    const view = try ws.ensureParsed();
    try testing.expect(view.rows.len == 3);
}

test "composition: the buffer handoff and the path handoff agree" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmp.dir.realPathFileAlloc(io, ".", a);
    defer a.free(dir);
    const path = try std.fs.path.join(a, &.{ dir, "handoff.xlsx" });
    defer a.free(path);

    var w = Writer.init(a);
    defer w.deinit();
    try buildWorkbook(&w);

    try w.save(io, path);
    const bytes = try w.saveToOwnedBuffer(a, io);
    defer a.free(bytes);

    var from_path = try Workbook.open(a, io, path);
    defer from_path.deinit();
    var from_buffer = try Workbook.openBuffer(a, io, bytes);
    defer from_buffer.deinit();

    try testing.expectEqual(from_path.sheetCount(), from_buffer.sheetCount());
    for (0..from_path.sheetCount()) |i| {
        const idx: u32 = @intCast(i);
        const p = try from_path.sheet(idx);
        const b = try from_buffer.sheet(idx);
        try testing.expectEqualStrings(p.name(), b.name());
        const pv = try p.ensureParsed();
        const bv = try b.ensureParsed();
        try testing.expectEqual(pv.rows.len, bv.rows.len);
    }
}

test "composition: the transaction reaches a buffer-opened workbook" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var w = Writer.init(a);
    defer w.deinit();
    try buildWorkbook(&w);
    const bytes = try w.saveToOwnedBuffer(a, io);
    defer a.free(bytes);

    var wb = try Workbook.openBuffer(a, io, bytes);
    defer wb.deinit();

    // M5b2's transaction over a workbook that never had a path. Nothing
    // in the recalc path knows which arm of the backing it is reading,
    // and `markRecalcOnLoad` is the smallest thing that proves it: it
    // runs a full prepare/swap and retains a generation.
    try wb.markRecalcOnLoad();
    try testing.expectEqual(@as(usize, 1), wb.retained.items.len);

    const part = (try wb.store.part("xl/workbook.xml")) orelse return error.MissingPart;
    try testing.expect(std.mem.indexOf(u8, part.bytes, "fullCalcOnLoad=\"1\"") != null);
}
