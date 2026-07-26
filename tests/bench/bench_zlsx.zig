// Benchmark: open an xlsx, iterate every row of the first sheet,
// tally cells by type. Prints one summary line at the end.
//
// Uses `std.heap.smp_allocator` for the same reason the writer bench
// does (see tests/bench/bench_write_zlsx.zig) — it matches what a
// production caller would actually plug in. `DebugAllocator` would
// add several ms of per-alloc tracking overhead that isn't
// representative of downstream behaviour; see docs/benchmarks.md's
// "Methodology" section for the full rationale.
//
// Mode: defaults to the streaming `Book.rows` iterator over
// `Book.open` (which eagerly preloads every sheet's side-indices —
// merged ranges, hyperlinks, validations, comments — so the
// metadata getters are populated synchronously). Pass:
//   --materialise   drive `Book.materialiseSheet` instead of the
//                   streaming iterator (row-materialise delta).
//   --lazy          use `Book.openLazy` — defers per-sheet
//                   side-index parsing until the caller asks for
//                   it. The right comparison against calamine,
//                   which is also lazy by default; the win is
//                   noticeable on multi-sheet workbooks
//                   (`ons_cpi_detailed.xlsx`: 42 ms eager → 4 ms
//                   lazy).
const std = @import("std");
const xlsx = @import("zlsx");

pub fn main(init: std.process.Init) !void {
    const io = init.io;
    const alloc = std.heap.smp_allocator;

    const args = try init.minimal.args.toSlice(init.arena.allocator());
    if (args.len < 2) {
        std.debug.print("usage: {s} [--materialise] <xlsx>\n", .{args[0]});
        return;
    }

    var path: ?[]const u8 = null;
    var materialise = false;
    var lazy = false;
    for (args[1..]) |a| {
        if (std.mem.eql(u8, a, "--materialise")) {
            materialise = true;
        } else if (std.mem.eql(u8, a, "--lazy")) {
            lazy = true;
        } else {
            path = a;
        }
    }
    const xlsx_path = path orelse {
        std.debug.print("usage: {s} [--materialise] [--lazy] <xlsx>\n", .{args[0]});
        return;
    };

    var book = if (lazy)
        try xlsx.Book.openLazy(alloc, io, xlsx_path)
    else
        try xlsx.Book.open(alloc, io, xlsx_path);
    defer book.deinit();

    if (book.sheets.len == 0) return;

    var n_rows: usize = 0;
    var n_str: usize = 0;
    var n_int: usize = 0;
    var n_num: usize = 0;
    var n_bool: usize = 0;
    var n_empty: usize = 0;

    if (materialise) {
        var matrix = try book.materialiseSheet(book.sheets[0], alloc);
        defer matrix.deinit();
        for (matrix.rows) |row| {
            n_rows += 1;
            for (row) |c| switch (c) {
                .string => n_str += 1,
                .integer => n_int += 1,
                .number => n_num += 1,
                .boolean => n_bool += 1,
                .empty => n_empty += 1,
            };
        }
    } else {
        var rows = try book.rows(book.sheets[0], alloc);
        defer rows.deinit();
        while (try rows.next()) |cells| {
            n_rows += 1;
            for (cells) |c| switch (c) {
                .string => n_str += 1,
                .integer => n_int += 1,
                .number => n_num += 1,
                .boolean => n_bool += 1,
                .empty => n_empty += 1,
            };
        }
    }

    var buf: [256]u8 = undefined;
    const tag: []const u8 = if (materialise)
        (if (lazy) "matrix-lazy" else "matrix")
    else
        (if (lazy) "stream-lazy" else "stream");
    const msg = try std.fmt.bufPrint(&buf, "mode={s} rows={d} str={d} int={d} num={d} bool={d} empty={d}\n", .{ tag, n_rows, n_str, n_int, n_num, n_bool, n_empty });
    // 0.16 removed File.write; stdout goes through a writer interface.
    var out_buf: [512]u8 = undefined;
    var ow = std.Io.File.stdout().writer(io, &out_buf);
    ow.interface.writeAll(msg) catch {};
    ow.interface.flush() catch {};
}
