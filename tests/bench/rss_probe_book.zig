//! Bench helper: open the fixture as `xlsx.Book.openLazy` and emit
//! the post-open RSS delta on stdout (decimal bytes, single line).
//!
//! Companion to `rss_probe_workbook.zig` — same shape, different
//! backend. Lives in its own binary because `cli_mod` + `zlsx_pkg`
//! + `writer` can't coexist in a single Zig 0.15.2 compilation
//! (see `build.zig` + `AGENTS.md`).
//!
//! Argv contract: `<exe> <fixture-path>`.

const std = @import("std");
const xlsx = @import("zlsx");
const rss = @import("rss");

pub fn main(init: std.process.Init) !void {
    const io = init.io;
    var gpa: std.heap.DebugAllocator(.{}) = .init;
    defer _ = gpa.deinit();
    const allocator = gpa.allocator();

    const args = try init.minimal.args.toSlice(init.arena.allocator());
    if (args.len < 2) {
        std.debug.print("usage: {s} <fixture.xlsx>\n", .{args[0]});
        std.process.exit(2);
    }
    const fixture = args[1];

    // Settle the allocator before the baseline read so post-warmup
    // pages are part of the baseline rather than the workload delta.
    {
        const warmup = try allocator.alloc(u8, 1 * 1024 * 1024);
        defer allocator.free(warmup);
        @memset(warmup, 0);
    }

    const baseline = try rss.rssBytes(io);

    var book = try xlsx.Book.openLazy(allocator, io, fixture);
    defer book.deinit();
    // Touch nothing — the gate is "before any sheet is touched".

    const after = try rss.rssBytes(io);
    const delta: u64 = if (after > baseline) after - baseline else 0;

    var stdout_buf: [64]u8 = undefined;
    const out = try std.fmt.bufPrint(&stdout_buf, "{d}\n", .{delta});
    {
        var obuf: [4096]u8 = undefined;
        var ow = std.Io.File.stdout().writer(io, &obuf);
        try ow.interface.writeAll(out);
        try ow.interface.flush();
    }
}
