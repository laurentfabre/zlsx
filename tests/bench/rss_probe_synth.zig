//! Bench helper: synthesise the iter-wb-6 RSS gate fixture if absent.
//!
//! Pulled out of the orchestration test process so the orchestrator
//! can stay independent of `zlsx`'s module (which it would otherwise
//! pull in just to drive the writer). Argv: `<exe> <out-path>`.

const std = @import("std");
const synth = @import("synth");

pub fn main(init: std.process.Init) !void {
    const io = init.io;
    var gpa: std.heap.DebugAllocator(.{}) = .init;
    defer _ = gpa.deinit();
    const allocator = gpa.allocator();

    const args = try init.minimal.args.toSlice(init.arena.allocator());
    if (args.len < 2) {
        std.debug.print("usage: {s} <out.xlsx>\n", .{args[0]});
        std.process.exit(2);
    }
    try synth.synthesize(allocator, io, args[1]);
}
