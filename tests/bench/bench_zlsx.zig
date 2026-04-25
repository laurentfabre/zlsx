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
// Mode: defaults to the streaming `Book.rows` iterator. Pass
// `--materialise` to drive the new `Book.materialiseSheet` API
// instead — useful for measuring the row-materialise delta on the
// same fixture without forking a second binary.
const std = @import("std");
const xlsx = @import("zlsx");

pub fn main() !void {
    const alloc = std.heap.smp_allocator;

    const args = try std.process.argsAlloc(alloc);
    defer std.process.argsFree(alloc, args);
    if (args.len < 2) {
        std.debug.print("usage: {s} [--materialise] <xlsx>\n", .{args[0]});
        return;
    }

    var path: ?[]const u8 = null;
    var materialise = false;
    for (args[1..]) |a| {
        if (std.mem.eql(u8, a, "--materialise")) {
            materialise = true;
        } else {
            path = a;
        }
    }
    const xlsx_path = path orelse {
        std.debug.print("usage: {s} [--materialise] <xlsx>\n", .{args[0]});
        return;
    };

    var book = try xlsx.Book.open(alloc, xlsx_path);
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
    const tag: []const u8 = if (materialise) "matrix" else "stream";
    const msg = try std.fmt.bufPrint(&buf, "mode={s} rows={d} str={d} int={d} num={d} bool={d} empty={d}\n", .{ tag, n_rows, n_str, n_int, n_num, n_bool, n_empty });
    _ = std.fs.File.stdout().write(msg) catch {};
}
