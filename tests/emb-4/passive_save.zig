//! emb-4 passive-save control.
//!
//! Opens an .xlsx and saves it back unedited via zlsx's own
//! delta-on-bytes writer. This is the "control" leg of the compat
//! matrix: it confirms zlsx's *own* load → save round-trip preserves
//! the `xl/zlsxEmbeddings/*` arc (workbook→index rel, content-type
//! overrides, per-coverage vec.bin/hashes.bin). A third-party tool
//! that strips on save is then a property of that tool, not of the
//! part format — this leg pins the zlsx baseline at exit 0.
//!
//! Usage:
//!   zlsx-emb4-passive-save <in.xlsx> <out.xlsx>

const std = @import("std");
const pkg = @import("zlsx_pkg");

pub fn main(init: std.process.Init) !u8 {
    const allocator = init.gpa;
    const io = init.io;

    const args = try init.minimal.args.toSlice(init.arena.allocator());

    var err_buf: [256]u8 = undefined;
    var err_w = std.Io.File.stderr().writer(io, &err_buf);
    const err = &err_w.interface;
    defer err.flush() catch {};

    if (args.len < 3) {
        try err.print("usage: {s} <in.xlsx> <out.xlsx>\n", .{args[0]});
        return 2;
    }

    var wb = try pkg.Workbook.open(allocator, io, args[1]);
    defer wb.deinit();
    try wb.save(io, args[2]);

    var out_buf: [512]u8 = undefined;
    var out_w = std.Io.File.stdout().writer(io, &out_buf);
    const out = &out_w.interface;
    defer out.flush() catch {};
    try out.print("passive save: {s} -> {s}\n", .{ args[1], args[2] });
    return 0;
}
