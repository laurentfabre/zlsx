//! Bench helper: open the fixture as `pkg.Workbook.openLazy` and
//! emit the post-open RSS delta on stdout.
//!
//! Companion to `rss_probe_book.zig`. See that file's preamble.

const std = @import("std");
const pkg = @import("zlsx_pkg");
const rss = @import("rss");

pub fn main() !void {
    var gpa: std.heap.DebugAllocator(.{}) = .init;
    defer _ = gpa.deinit();
    const allocator = gpa.allocator();

    const args = try std.process.argsAlloc(allocator);
    defer std.process.argsFree(allocator, args);
    if (args.len < 2) {
        std.debug.print("usage: {s} <fixture.xlsx>\n", .{args[0]});
        std.process.exit(2);
    }
    const fixture = args[1];

    {
        const warmup = try allocator.alloc(u8, 1 * 1024 * 1024);
        defer allocator.free(warmup);
        @memset(warmup, 0);
    }

    const baseline = try rss.rssBytes();

    var wb = try pkg.Workbook.openLazy(allocator, fixture);
    defer wb.deinit();
    // Touch nothing — same contract as the Book probe.

    const after = try rss.rssBytes();
    const delta: u64 = if (after > baseline) after - baseline else 0;

    var stdout_buf: [64]u8 = undefined;
    const out = try std.fmt.bufPrint(&stdout_buf, "{d}\n", .{delta});
    _ = try std.Io.File.stdout().writeAll(out);
}
