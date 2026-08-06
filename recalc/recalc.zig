//! `zlsx_recalc` — the third public module (§5.10).
//!
//! M5c built the **shell**: the module exists, imports `zlsx` and
//! `zlsx_pkg`, and proves that a producer's bytes can reach a consumer
//! without a filesystem and without a cycle. M5d3 puts the composition
//! on top of it — `writerSaveWithRecalc`, which threads a recalculation
//! between those two halves and a `Control` through all three stages —
//! and gates the whole graph on `tests/consumer`, a downstream package
//! that imports all three public modules in one compilation.
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

/// §5.10's `Control`, which M5d3's `writerSaveWithRecalc` threads into
/// BOTH pre-recalc stages. Nameable from the orchestrator module because
/// that is where a caller composing the two halves lives; the identity
/// assertion below is the same argument as the `Edit.cell` one above —
/// two `Control` types with identical fields would make the threading
/// uncompilable for reasons no error message would explain.
pub const Control = zlsx.Control;
pub const CancelToken = zlsx.CancelToken;

comptime {
    assert(Control == pkg.Control);
    assert(@FieldType(Control, "cancel") == ?CancelToken);
}

/// §5.5's reproducible inputs and §5.7's policy, under the names the
/// composition's own signature uses. Re-exported for the same reason
/// `zlsx_pkg` re-exports them: a caller that imports only the
/// orchestrator cannot otherwise spell what it must pass.
pub const RunInputs = pkg.RunInputs;
pub const Options = pkg.RecalcOptions;
pub const Report = pkg.RecalcReport;
/// §5.8c (M7c): the authoring dialect, **Zig-only** — the versioned C
/// export and the Python binding land at M9a2.
pub const FormulaWrite = pkg.FormulaWrite;

// ─── the composition (§5.10) ─────────────────────────────────────

/// §5.10's orchestrator: a `Writer`'s workbook, recalculated, committed
/// to `path` — with no temporary file and no second archive on disk.
///
/// Three stages, and the middle one is the reason this module exists:
///
///   1. `Writer.saveToOwnedBufferControlled` serializes the authored
///      workbook into memory (`zlsx`),
///   2. `Workbook.openBufferControlled` reads those bytes back as a
///      package (`zlsx_pkg`) — the borrow ends when it returns, so the
///      buffer is freed before this function does,
///   3. `Workbook.saveWithRecalc` runs M5d2's pipeline and §5.7.9's file
///      transaction.
///
/// **The `Control` threads into BOTH pre-recalc stages (§5.10,
/// normative).** Stage 1 and stage 2 can each process gigabytes before
/// a single formula is evaluated, so a `run` carrying a deadline that
/// only reached stage 3 would leave the §5.5 polling bound true of a
/// third of the work. `run.cancel` and `run.deadline` are the whole of
/// `Control` by construction — both are outside `EffectiveRunInputs`, so
/// threading them here cannot change what a run fingerprints as.
///
/// What a caller gets from the failure modes is what `saveWithRecalc`
/// promises, extended backwards: a cancellation in stage 1 or stage 2
/// happens strictly before §5.7.9's commit point, so the destination
/// still holds its prior bytes (or is still absent), no temp file is
/// left beside it, and the `Writer` is exactly as it was — a save never
/// consumed it.
///
/// The returned `Report` is the caller's and outlives the `Workbook`
/// this builds and tears down: its census is allocated from `gpa`, which
/// is the allocator stage 2 hands the workbook, so `Report.deinit(gpa)`
/// is the right call and there is no borrowed sheet name in it to
/// dangle.
pub fn writerSaveWithRecalc(
    gpa: std.mem.Allocator,
    io: std.Io,
    writer: *Writer,
    path: []const u8,
    run: RunInputs,
    opts: Options,
) !Report {
    assert(path.len > 0);

    const ctl: Control = .{ .cancel = run.cancel, .deadline = run.deadline };

    const bytes = try writer.saveToOwnedBufferControlled(gpa, io, ctl);
    defer gpa.free(bytes);

    var wb = try Workbook.openBufferControlled(gpa, io, bytes, ctl);
    defer wb.deinit();

    return wb.saveWithRecalc(gpa, io, path, run, opts);
}

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

// ─── M5d3: the composition ───────────────────────────────────────

const Allocator = std.mem.Allocator;
const control = pkg.control;

/// The run every composition test uses. Fixed clock and fixed seed
/// because §5.5 makes a run reproducible from its inputs, and the
/// byte-equality below is a statement about the pipeline, not about
/// what time it happened to be.
const fixed_run: RunInputs = .{
    .now_utc_ms = 1_700_000_000_000,
    .rng_seed = 0x5EED_5D3,
    .limits = .{},
};

/// A workbook whose caches are all wrong. `999` is not what any of the
/// three formulas says, so a composition that serialized, opened and
/// saved without recalculating would produce a file this test can tell
/// apart from one that did.
fn buildFormulaWorkbook(w: *Writer) !void {
    var s = try w.addSheet("Calc");
    try s.writeRow(&.{ .{ .string = "n" }, .{ .string = "next" } });
    try s.writeRowWithFormulas(
        &.{ .{ .integer = 1 }, .{ .integer = 999 } },
        &.{ null, "A2+1" },
    );
    try s.writeRowWithFormulas(
        &.{ .{ .integer = 2 }, .{ .integer = 999 } },
        &.{ null, "A3+1" },
    );
    try s.writeRowWithFormulas(
        &.{ .empty, .{ .integer = 999 } },
        &.{ null, "SUM(B2:B3)" },
    );
}

fn readAll(gpa: Allocator, io: std.Io, path: []const u8) ![]u8 {
    var f = try std.Io.Dir.cwd().openFile(io, path, .{});
    defer f.close(io);
    const size = (try f.stat(io)).size;
    const buf = try gpa.alloc(u8, @intCast(size));
    errdefer gpa.free(buf);
    var r = f.reader(io, &.{});
    try r.interface.readSliceAll(buf);
    return buf;
}

fn cellCache(ws: *Worksheet, ref: []const u8) ![]const u8 {
    const view = try ws.ensureParsed();
    for (view.rows) |row| {
        for (row.cells) |c| {
            if (std.mem.eql(u8, c.ref, ref)) return c.raw_value orelse "";
        }
    }
    return error.CellNotFound;
}

test "composition: writerSaveWithRecalc is the three steps, byte for byte" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmp.dir.realPathFileAlloc(io, ".", a);
    defer a.free(dir);
    const composed = try std.fs.path.join(a, &.{ dir, "composed.xlsx" });
    defer a.free(composed);
    const by_hand = try std.fs.path.join(a, &.{ dir, "by_hand.xlsx" });
    defer a.free(by_hand);

    // One Writer for both: `saveToOwnedBuffer` does not consume it, and
    // using two would leave "the fixtures were equal" as an alternative
    // explanation for equal outputs.
    var w = Writer.init(a);
    defer w.deinit();
    try buildFormulaWorkbook(&w);

    {
        var report = try writerSaveWithRecalc(a, io, &w, composed, fixed_run, .{});
        defer report.deinit(a);
        try testing.expectEqual(@as(u32, 3), report.cells_written);
        try testing.expectEqual(@as(u32, 1), report.sheets_patched);
    }

    // The same three stages, spelled out by a caller who does not have
    // the orchestrator.
    {
        const bytes = try w.saveToOwnedBuffer(a, io);
        defer a.free(bytes);
        var wb = try Workbook.openBuffer(a, io, bytes);
        defer wb.deinit();
        var report = try wb.saveWithRecalc(a, io, by_hand, fixed_run, .{});
        defer report.deinit(a);
    }

    const x = try readAll(a, io, composed);
    defer a.free(x);
    const y = try readAll(a, io, by_hand);
    defer a.free(y);
    try testing.expectEqualSlices(u8, x, y);

    // And both hold what the formulas say rather than what the caches
    // did — without this the equality above would also hold for two
    // runs that recalculated nothing.
    var out = try Workbook.open(a, io, composed);
    defer out.deinit();
    const ws = try out.sheet(0);
    try testing.expectEqualStrings("2", try cellCache(ws, "B2"));
    try testing.expectEqualStrings("3", try cellCache(ws, "B3"));
    try testing.expectEqualStrings("5", try cellCache(ws, "B4"));
}

// ─── the Control reaches BOTH pre-recalc stages (§5.10) ──────────

/// A workbook big enough that each pre-recalc stage spans many chunks:
/// high-entropy strings, so deflate cannot collapse the sheet body back
/// into one poll's worth of input. The formula rows ride along so the
/// fixture is the same *kind* of workbook the byte-equality test uses.
fn buildLargeFormulaWorkbook(w: *Writer, rows: usize) !void {
    var s = try w.addSheet("Big");
    var rng: std.Random.DefaultPrng = .init(0x5d3_5d3);
    var cell: [64]u8 = undefined;
    for (0..rows) |_| {
        rng.random().bytes(&cell);
        for (&cell) |*c| c.* = 'a' + (c.* % 26);
        try s.writeRow(&.{ .{ .string = &cell }, .{ .integer = 1 } });
    }
    try s.writeRowWithFormulas(
        &.{ .empty, .{ .integer = 999 } },
        &.{ null, "1+1" },
    );
}

/// Where the poll counter stands at the end of each pre-recalc stage.
///
/// Measured rather than hardcoded: a chunk-size change or an extra part
/// would silently move a constant into the wrong stage, and the whole
/// point of the tests below is *which* stage refused.
const StagePolls = struct { serialize: u64, open_end: u64 };

fn countStagePolls(gpa: Allocator, base: std.Io, w: *Writer) !StagePolls {
    // A clock that never advances, against a deadline that therefore
    // never fires: every poll is one counted read and nothing ends the
    // run early. The same measurement M5d1's polling-bound test makes.
    const ctl: Control = .{ .deadline = .{ .nanoseconds = std.math.maxInt(i64) } };
    const io = control.inject.wrap(base, .{});

    const bytes = try w.saveToOwnedBufferControlled(gpa, io, ctl);
    defer gpa.free(bytes);
    const serialize = control.inject.state.now_calls;

    var wb = try Workbook.openBufferControlled(gpa, io, bytes, ctl);
    defer wb.deinit();
    return .{ .serialize = serialize, .open_end = control.inject.state.now_calls };
}

const Stage = enum { serialize, buffer_open };

/// Cancel inside `stage` and assert what §5.7.9 promises about a failure
/// before the commit point: the destination is exactly as it was, and
/// nothing is left beside it.
///
/// `prior` picks which half of that promise is under test — known bytes
/// that must survive, or a destination that never existed and must stay
/// absent.
fn expectStageCancelled(stage: Stage, prior: ?[]const u8) !void {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const base = threaded.io();

    var tmp = testing.tmpDir(.{ .iterate = true });
    defer tmp.cleanup();
    const dir = try tmp.dir.realPathFileAlloc(base, ".", a);
    defer a.free(dir);
    const dest = try std.fs.path.join(a, &.{ dir, "dest.xlsx" });
    defer a.free(dest);
    if (prior) |bytes| try tmp.dir.writeFile(base, .{ .sub_path = "dest.xlsx", .data = bytes });

    var w = Writer.init(a);
    defer w.deinit();
    try buildLargeFormulaWorkbook(&w, 6000);

    const polls = try countStagePolls(a, base, &w);
    // "Inside this stage" has to be a statement about placement, so each
    // stage needs room for a trip that is unambiguously within it.
    try testing.expect(polls.serialize >= 4);
    try testing.expect(polls.open_end - polls.serialize >= 4);

    // `Control.check` reads the token BEFORE the clock, so the read that
    // arms the flag is not the poll that refuses — tripping at N means
    // poll N+1 returns `Cancelled`. Both offsets below put that poll
    // strictly inside the named stage.
    const trip_at: u64 = switch (stage) {
        .serialize => polls.serialize - 2,
        .buffer_open => polls.serialize + 1,
    };

    var flag: u8 = 0;
    const io = control.inject.wrap(base, .{ .trip_at = trip_at, .trip_flag = &flag });
    var run = fixed_run;
    run.deadline = .{ .nanoseconds = std.math.maxInt(i64) };
    run.cancel = .{ .flag = &flag };

    try testing.expectError(error.Cancelled, writerSaveWithRecalc(a, io, &w, dest, run, .{}));

    // Which stage refused, not merely that one did. The refusing poll
    // never reaches the clock, so the counter stops where the flag was
    // armed.
    switch (stage) {
        .serialize => try testing.expect(control.inject.state.now_calls < polls.serialize),
        .buffer_open => {
            try testing.expect(control.inject.state.now_calls > polls.serialize);
            try testing.expect(control.inject.state.now_calls <= polls.open_end);
        },
    }

    if (prior) |want| {
        const after = try readAll(a, base, dest);
        defer a.free(after);
        try testing.expectEqualSlices(u8, want, after);
    } else {
        try testing.expectError(error.FileNotFound, std.Io.Dir.cwd().access(base, dest, .{}));
    }

    var it = tmp.dir.iterate();
    while (try it.next(base)) |entry| {
        if (std.mem.startsWith(u8, entry.name, ".ztmp")) return error.TempFileLeftBehind;
    }

    // The Writer is untouched: a save never consumed it, and a cancelled
    // one is no different.
    const again = try w.saveToOwnedBuffer(a, base);
    defer a.free(again);
    try testing.expect(again.len > 22);
}

test "composition: a cancel mid-serialization leaves the prior bytes" {
    try expectStageCancelled(.serialize, "PRIOR WORKBOOK BYTES");
}

test "composition: a cancel mid-serialization leaves an absent file absent" {
    try expectStageCancelled(.serialize, null);
}

test "composition: a cancel mid-buffer-open leaves the prior bytes" {
    try expectStageCancelled(.buffer_open, "PRIOR WORKBOOK BYTES");
}

test "composition: a cancel mid-buffer-open leaves an absent file absent" {
    try expectStageCancelled(.buffer_open, null);
}

test "composition: a deadline expires in the first stage, before any file exists" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const base = threaded.io();

    var tmp = testing.tmpDir(.{ .iterate = true });
    defer tmp.cleanup();
    const dir = try tmp.dir.realPathFileAlloc(base, ".", a);
    defer a.free(dir);
    const dest = try std.fs.path.join(a, &.{ dir, "dest.xlsx" });
    defer a.free(dest);

    var w = Writer.init(a);
    defer w.deinit();
    try buildLargeFormulaWorkbook(&w, 6000);

    // 1 ms per poll against a 5 ms budget: the sixth read is past it.
    // Nothing but the `Control` carries this — `RunInputs.deadline` is
    // outside `EffectiveRunInputs`, so the run it describes is the same
    // run either way.
    const io = control.inject.wrap(base, .{ .step_ns = std.time.ns_per_ms });
    var run = fixed_run;
    run.deadline = .{ .nanoseconds = 5 * std.time.ns_per_ms };

    try testing.expectError(error.Cancelled, writerSaveWithRecalc(a, io, &w, dest, run, .{}));
    try testing.expectError(error.FileNotFound, std.Io.Dir.cwd().access(base, dest, .{}));
}

test "composition: a disarmed run reaches the same bytes as an armed one that never fires" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmp.dir.realPathFileAlloc(io, ".", a);
    defer a.free(dir);
    const plain_path = try std.fs.path.join(a, &.{ dir, "plain.xlsx" });
    defer a.free(plain_path);
    const armed_path = try std.fs.path.join(a, &.{ dir, "armed.xlsx" });
    defer a.free(armed_path);

    var w = Writer.init(a);
    defer w.deinit();
    try buildFormulaWorkbook(&w);

    {
        var r = try writerSaveWithRecalc(a, io, &w, plain_path, fixed_run, .{});
        r.deinit(a);
    }
    {
        // A token that exists and never fires. The polling seam runs,
        // and produces the same archive — cancellation is a control
        // channel, not an input to the run.
        var never = std.atomic.Value(bool).init(false);
        var run = fixed_run;
        run.cancel = .{ .atomic = &never };
        var r = try writerSaveWithRecalc(a, io, &w, armed_path, run, .{});
        r.deinit(a);
    }

    const x = try readAll(a, io, plain_path);
    defer a.free(x);
    const y = try readAll(a, io, armed_path);
    defer a.free(y);
    try testing.expectEqualSlices(u8, x, y);
}
