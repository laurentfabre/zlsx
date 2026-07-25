//! C2a: standalone CLI for extracting embedded images from an
//! .xlsx workbook to a target directory. Uses `zlsx_pkg.PartStore`
//! directly so it works on workbooks the full reader / Editor
//! would refuse (corrupt sheet XML, unsupported features) — only
//! the ZIP layer + content-type lookup matter.
//!
//! Usage:
//!     zlsx-extract-images <in.xlsx> <out-dir>
//!
//! Output filename per image is the basename of the image part
//! (e.g. `image1.jpeg`). Mirrors `unzip xl/media/* -d out-dir`.
//!
//! Exit codes:
//!     0  — success (zero or more images written)
//!     1  — bad usage / missing args
//!     2  — cannot open the workbook
//!     5  — cannot create or write to out-dir
//!
//! Lives as a separate binary instead of a `zlsx` subcommand
//! because Zig 0.15.2's module-graph computation rejects
//! cli_mod + zlsx_pkg + writer in one compilation. See
//! docs/plans/post-0.2.9-roadmap.md for the constraint write-up.

const std = @import("std");
const pkg = @import("zlsx_pkg");

pub fn main(init: std.process.Init) !u8 {
    // 0.16 supplies the allocator, Io and argv through process.Init;
    // std.heap.GeneralPurposeAllocator and std.process.argsAlloc are
    // both gone.
    const alloc = init.gpa;
    const io = init.io;

    const args = try init.minimal.args.toSlice(init.arena.allocator());

    var stderr_buf: [256]u8 = undefined;
    var stderr_w = std.Io.File.stderr().writer(io, &stderr_buf);
    const err = &stderr_w.interface;
    defer err.flush() catch {};

    if (args.len < 3) {
        try err.print(
            "usage: {s} <in.xlsx> <out-dir>\n",
            .{args[0]},
        );
        return 1;
    }
    const in_path = args[1];
    const out_dir = args[2];

    // Reject empty path strings up front. Without this, an unset
    // shell variable (`zlsx-extract-images "$tmp" "$out"`) would
    // bottom out in std.fs with a cryptic `BadPathName`. Better to
    // tell the caller what's actually wrong.
    if (in_path.len == 0) {
        try err.writeAll("error: <in.xlsx> path is empty (unset shell variable?)\n");
        return 1;
    }
    if (out_dir.len == 0) {
        try err.writeAll("error: <out-dir> path is empty (unset shell variable?)\n");
        return 1;
    }

    var store = pkg.PartStore.open(alloc, io, in_path) catch |e| {
        try err.print("cannot open '{s}': {s}\n", .{ in_path, @errorName(e) });
        return 2;
    };
    defer store.deinit();

    std.Io.Dir.cwd().createDirPath(io, out_dir) catch |e| {
        try err.print("cannot create '{s}': {s}\n", .{ out_dir, @errorName(e) });
        return 5;
    };
    var dir = std.Io.Dir.cwd().openDir(io, out_dir, .{}) catch |e| {
        try err.print("cannot open '{s}': {s}\n", .{ out_dir, @errorName(e) });
        return 5;
    };
    defer dir.close(io);

    var stdout_buf: [256]u8 = undefined;
    var stdout_w = std.Io.File.stdout().writer(io, &stdout_buf);
    const out = &stdout_w.interface;
    defer out.flush() catch {};

    var written: usize = 0;
    const images = store.imageParts() catch |e| {
        try err.print("imageParts: {s}\n", .{@errorName(e)});
        return 5;
    };
    for (images) |p| {
        // Filter to xl/media/ — that's the canonical home for
        // sheet-embedded images. Other image-content-type parts
        // exist (docProps/thumbnail.jpeg, custom workbook resources)
        // but they aren't anchored to a worksheet and the CLI's
        // documented contract is "mirror unzip xl/media/* -d out".
        if (!std.mem.startsWith(u8, p.name, "xl/media/")) continue;
        const basename = std.fs.path.basename(p.name);
        dir.writeFile(io, .{ .sub_path = basename, .data = p.bytes }) catch |e| {
            try err.print("write '{s}': {s}\n", .{ basename, @errorName(e) });
            return 5;
        };
        try out.print("{s}\n", .{basename});
        written += 1;
    }
    try err.print("zlsx-extract-images: wrote {d} image(s) to {s}\n", .{ written, out_dir });
    return 0;
}
