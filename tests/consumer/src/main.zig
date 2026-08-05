//! Z2 gate: round-trips a workbook through ALL THREE public zlsx
//! modules in a single compilation:
//!
//!   const zlsx   = @import("zlsx");         // Writer + Book.open + Rows
//!   const pkg    = @import("zlsx_pkg");     // Editor.open + setCells + save
//!   const recalc = @import("zlsx_recalc");  // writerSaveWithRecalc (§5.10)
//!
//! write -> read a cell -> mutate it via Editor -> save -> re-read and
//! verify. That is the exact shape nemonym needs (reader for scanning,
//! Editor for byte-preserving mask writes), so if this builds and
//! passes, the co-import question is settled for downstream users.
//!
//! Step 5 is M5d3's dependency test: a Writer with a deliberately wrong
//! cached value goes through `zlsx_recalc.writerSaveWithRecalc`, and the
//! re-read has to show what the formula says. It exercises the one thing
//! a unit test inside the repo cannot — that the three modules resolve
//! to one graph when a *downstream* `build.zig` wires them.
//!
//! The fixture is generated rather than taken from `tests/corpus/`, for
//! two reasons: corpus fetching is deliberately fail-soft so the files
//! may be absent, and several corpus archives use ZIP data descriptors
//! which `Editor.open` refuses by design. Generating also means this
//! one binary exercises Writer, Book and Editor together.
//!
//! Exits non-zero with a diagnostic on any mismatch, so it works as a
//! CI gate with no test harness around it.

const std = @import("std");
const zlsx = @import("zlsx");
const pkg = @import("zlsx_pkg");
const recalc = @import("zlsx_recalc");

const SENTINEL = "Z2_CO_IMPORT_OK";
const ORIGINAL = "before";

/// The composition's inputs, fixed. §5.5 makes a run reproducible from
/// them, and a gate that read the wall clock would be asserting
/// something it cannot repeat.
const RUN: recalc.RunInputs = .{
    .now_utc_ms = 1_700_000_000_000,
    .rng_seed = 0x5EED_5D3,
    .limits = .{},
};

pub fn main(init: std.process.Init) !u8 {
    const gpa = init.gpa;
    const io = init.io;
    const args = try init.minimal.args.toSlice(init.arena.allocator());

    var out: std.Io.File.Writer = .init(std.Io.File.stdout(), io, &.{});
    const w = &out.interface;
    defer w.flush() catch {};

    if (args.len < 3) {
        try w.print("usage: {s} <scratch-in.xlsx> <scratch-out.xlsx>\n", .{args[0]});
        return 2;
    }
    const in_path = args[1];
    const out_path = args[2];

    // ── 1. Write a fixture, via the `zlsx` module ──────────────────────
    {
        var writer = zlsx.Writer.init(gpa);
        defer writer.deinit();
        var sheet = try writer.addSheet("Data");
        try sheet.writeRow(&.{.{ .string = "header" }});
        try sheet.writeRow(&.{.{ .string = ORIGINAL }});
        try writer.save(io, in_path);
    }
    try w.print("wrote  {s}\n", .{in_path});

    // ── 2. Read side, via the `zlsx` module ────────────────────────────
    const before = try readA2(gpa, io, in_path);
    defer gpa.free(before);
    try w.print("read   A2 = \"{s}\"\n", .{before});
    if (!std.mem.eql(u8, before, ORIGINAL)) {
        try w.print("FAIL: fixture A2 should be \"{s}\"\n", .{ORIGINAL});
        return 1;
    }

    // ── 3. Mutate, via the `zlsx_pkg` module ───────────────────────────
    {
        var ed = try pkg.Editor.open(gpa, io, in_path);
        defer ed.deinit();
        // Spell the batch as `pkg.Edit` deliberately: an inferred anonymous
        // literal compiles even without the re-export, so only a named type
        // pins it. This is the consumer shape that motivated exporting it —
        // building the slice somewhere other than the call site.
        const edits = [_]pkg.Edit{
            .{ .row = 2, .col = 0, .cell = .{ .string = SENTINEL } },
        };
        try ed.setCells(0, &edits);
        try ed.save(io, out_path);
    }
    try w.print("edited {s}\n", .{out_path});

    // ── 4. Re-read, via the `zlsx` module again ────────────────────────
    const after = try readA2(gpa, io, out_path);
    defer gpa.free(after);
    try w.print("reread A2 = \"{s}\"\n", .{after});

    if (!std.mem.eql(u8, after, SENTINEL)) {
        try w.print("FAIL: expected \"{s}\", got \"{s}\"\n", .{ SENTINEL, after });
        return 1;
    }

    // ── 5. Compose all three, via `zlsx_recalc` (§5.10) ────────────────
    //
    // Derived from `out_path` rather than taken as a third argument so
    // the two-argument invocation documented in AGENTS.md and run by CI
    // keeps working.
    const recalc_path = try std.fmt.allocPrint(
        init.arena.allocator(),
        "{s}.recalc.xlsx",
        .{out_path},
    );
    {
        var writer = zlsx.Writer.init(gpa);
        defer writer.deinit();
        var sheet = try writer.addSheet("Calc");
        // A2 = 41, B2 = A2+1 with a cache that says 999. A composition
        // that serialized and saved without recalculating would leave
        // the 999 behind, so the assertion below is about the pipeline
        // and not about the writer.
        try sheet.writeRow(&.{.{ .string = "n" }});
        try sheet.writeRowWithFormulas(
            &.{ .{ .integer = 41 }, .{ .integer = 999 } },
            &.{ null, "A2+1" },
        );

        var report = try recalc.writerSaveWithRecalc(gpa, io, &writer, recalc_path, RUN, .{});
        defer report.deinit(gpa);
        if (report.cells_written != 1) {
            try w.print("FAIL: expected 1 recalculated cell, got {d}\n", .{report.cells_written});
            return 1;
        }
    }
    try w.print("recalc {s}\n", .{recalc_path});

    const computed = try readB2(gpa, io, recalc_path);
    defer gpa.free(computed);
    try w.print("reread B2 = \"{s}\"\n", .{computed});
    if (!std.mem.eql(u8, computed, "42")) {
        try w.print("FAIL: expected B2 = \"42\" after recalc, got \"{s}\"\n", .{computed});
        return 1;
    }

    try w.writeAll("round-trip OK\n");
    return 0;
}

/// Read the B2 cell of sheet 0 as text, through the *package* module —
/// so the check reads a cached formula result rather than the value
/// `Book`'s row iterator would coerce it to. Caller owns the result.
fn readB2(gpa: std.mem.Allocator, io: std.Io, path: []const u8) ![]u8 {
    var wb = try pkg.Workbook.open(gpa, io, path);
    defer wb.deinit();

    const ws = try wb.sheet(0);
    const view = try ws.ensureParsed();
    for (view.rows) |row| {
        for (row.cells) |c| {
            if (std.mem.eql(u8, c.ref, "B2")) return gpa.dupe(u8, c.raw_value orelse "");
        }
    }
    return error.CellNotFound;
}

/// Read the A2 cell of sheet 0 as text. Caller owns the result.
fn readA2(gpa: std.mem.Allocator, io: std.Io, path: []const u8) ![]u8 {
    var book = try zlsx.Book.open(gpa, io, path);
    defer book.deinit();

    if (book.sheets.len == 0) return error.NoSheets;

    var rows = try book.rows(book.sheets[0], gpa);
    defer rows.deinit();

    // Cells come back positionally (index 0 == column A); the 1-based
    // OOXML row number is read off the iterator after each `next()`.
    while (try rows.next()) |cells| {
        if (rows.currentRowNumber() != 2) continue;
        if (cells.len == 0) break;
        return switch (cells[0]) {
            .string => |s| try gpa.dupe(u8, s),
            .integer => |n| try std.fmt.allocPrint(gpa, "{d}", .{n}),
            .number => |n| try std.fmt.allocPrint(gpa, "{d}", .{n}),
            .boolean => |b| try gpa.dupe(u8, if (b) "true" else "false"),
            .empty => try gpa.dupe(u8, ""),
        };
    }
    return error.CellNotFound;
}
