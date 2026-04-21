// Benchmark: open an xlsx, iterate every row of the first sheet,
// tally cells by type. Prints one summary line at the end.
const std = @import("std");
const xlsx = @import("zlsx");

pub fn main() !void {
    var gpa: std.heap.DebugAllocator(.{}) = .init;
    defer _ = gpa.deinit();
    const alloc = gpa.allocator();

    const args = try std.process.argsAlloc(alloc);
    defer std.process.argsFree(alloc, args);
    if (args.len < 2) {
        std.debug.print("usage: {s} <xlsx>\n", .{args[0]});
        return;
    }

    var book = try xlsx.Book.open(alloc, args[1]);
    defer book.deinit();

    if (book.sheets.len == 0) return;
    var rows = try book.rows(book.sheets[0], alloc);
    defer rows.deinit();

    var n_rows: usize = 0;
    var n_str: usize = 0;
    var n_int: usize = 0;
    var n_num: usize = 0;
    var n_bool: usize = 0;
    var n_empty: usize = 0;
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
    var buf: [256]u8 = undefined;
    const msg = try std.fmt.bufPrint(&buf, "rows={d} str={d} int={d} num={d} bool={d} empty={d}\n", .{ n_rows, n_str, n_int, n_num, n_bool, n_empty });
    _ = std.fs.File.stdout().write(msg) catch {};
}
