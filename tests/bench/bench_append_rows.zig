// B2 iter-er-3 baseline bench — Editor.appendRows hot-path timing.
//
// Establishes the perf baseline for iter-er-3's walk-away gate: the
// rebased `Editor.appendRows` (route through `Worksheet.appendRows`
// on the Workbook overlay) MUST stay within 1.10× of the legacy
// substring-and-substitute fast-path on this fixture.
// Above 1.10× → keep the legacy path indefinitely (per
// docs/plans/archive/editor-rebase.md).
//
// Fixture: 100k rows × 5 cells per row, written into a synthetic
// base xlsx (zero pre-existing rows), then read back via `Editor.open`
// + `Editor.appendRows` + `Editor.save`. Reports total wall-clock,
// per-row µs, and per-cell ns.
//
// Uses `std.heap.smp_allocator` — same rationale as the read/write
// benches: matches what a production caller plugs in (DebugAllocator
// adds ~ms of allocation-tracking overhead that isn't representative
// of downstream behaviour).
//
// Usage:
//   zlsx-bench-append-rows <tmpdir> [rows]
//
// Default `rows = 100_000`. Output is one CSV-like line:
//   total_ms=NN.N rows=R cells=C avg_us_per_row=U avg_ns_per_cell=N
const std = @import("std");
const xlsx = @import("zlsx");
const zlsx_pkg = @import("zlsx_pkg");

const default_rows: usize = 100_000;
const cells_per_row: usize = 5;

pub fn main(init: std.process.Init) !void {
    const io = init.io;
    const alloc = std.heap.smp_allocator;

    const args = try init.minimal.args.toSlice(init.arena.allocator());
    if (args.len < 2) {
        std.debug.print("usage: {s} <tmpdir> [rows]\n", .{args[0]});
        return;
    }
    const tmpdir = args[1];
    const rows = if (args.len >= 3) try std.fmt.parseInt(usize, args[2], 10) else default_rows;

    var src_path_buf: [512]u8 = undefined;
    const src_path = try std.fmt.bufPrint(&src_path_buf, "{s}/append_src.xlsx", .{tmpdir});
    var dst_path_buf: [512]u8 = undefined;
    const dst_path = try std.fmt.bufPrint(&dst_path_buf, "{s}/append_dst.xlsx", .{tmpdir});

    // ── Step 1: synthesise the base xlsx (single empty sheet). ─────
    {
        var w = xlsx.Writer.init(alloc);
        defer w.deinit();
        var s = try w.addSheet("Bench");
        // Touch one row so the sheet has a valid <sheetData> body
        // shape — Editor.appendRows won't trigger fresh-SST creation
        // for our numeric appends.
        try s.writeRow(&.{.{ .integer = 0 }});
        try w.save(io, src_path);
    }

    // Pre-allocate the row payload — 5 numeric cells per row, same
    // values reused N times. Keeps the bench measuring append-only
    // overhead, not allocation noise.
    const row_template: [cells_per_row]xlsx.Cell = .{
        .{ .integer = 1 },
        .{ .number = 2.5 },
        .{ .integer = 3 },
        .{ .number = 4.75 },
        .{ .integer = 5 },
    };
    const all_rows = try alloc.alloc([]xlsx.Cell, rows);
    defer {
        for (all_rows) |r| alloc.free(r);
        alloc.free(all_rows);
    }
    for (all_rows) |*r| {
        const buf = try alloc.alloc(xlsx.Cell, cells_per_row);
        @memcpy(buf, &row_template);
        r.* = buf;
    }

    // ── Step 2: open + append + save with wall-clock timing. ──────
    // 0.16 removed std.time.Timer along with the rest of std.time's
    // functions; monotonic timing comes from the Io clock now.
    const t_start = std.Io.Clock.now(.awake, io).nanoseconds;
    {
        var ed = try zlsx_pkg.Editor.open(alloc, io, src_path);
        defer ed.deinit();
        try ed.appendRows(0, all_rows);
        try ed.save(io, dst_path);
    }
    const elapsed_ns: u64 = @intCast(std.Io.Clock.now(.awake, io).nanoseconds - t_start);

    const total_cells = rows * cells_per_row;
    const total_ms = @as(f64, @floatFromInt(elapsed_ns)) / 1_000_000.0;
    const avg_us_per_row = @as(f64, @floatFromInt(elapsed_ns)) / @as(f64, @floatFromInt(rows)) / 1_000.0;
    const avg_ns_per_cell = @as(f64, @floatFromInt(elapsed_ns)) / @as(f64, @floatFromInt(total_cells));

    std.debug.print(
        "total_ms={d:.2} rows={d} cells={d} avg_us_per_row={d:.2} avg_ns_per_cell={d:.1}\n",
        .{ total_ms, rows, total_cells, avg_us_per_row, avg_ns_per_cell },
    );

    // Cleanup so repeated bench runs in the same tmpdir don't OOM.
    std.Io.Dir.cwd().deleteFile(io, src_path) catch {};
    std.Io.Dir.cwd().deleteFile(io, dst_path) catch {};
}
