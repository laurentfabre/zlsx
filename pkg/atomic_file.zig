//! Crash-safe whole-file replacement.
//!
//! Zig 0.15 shipped `std.fs.Dir.atomicFile`; 0.16 removed it along with
//! the rest of the `std.fs` file surface (everything moved under
//! `std.Io`). Both `PartStore.save` and `Editor.save` depend on the
//! write-to-temp-then-rename contract to keep the "a failed save never
//! leaves a half-written workbook" guarantee in `docs/plans/`, so the
//! primitive is reimplemented here on 0.16 `std.Io` calls.
//!
//! The shape deliberately mirrors 0.15's `std.fs.AtomicFile` so the
//! call sites read the same:
//!
//! ```zig
//! var af = try AtomicFile.init(io, path, &write_buf);
//! defer af.deinit();
//! const w = &af.file_writer.interface;
//! try w.writeAll(bytes);
//! try af.finish();
//! ```
//!
//! Move contract (same as 0.15's): `file_writer`'s buffer pointer aims
//! at the caller-owned `write_buffer`, never into this struct, so
//! returning by value is safe. Take `&af.file_writer.interface` only
//! after the value has landed in its final location.

const std = @import("std");

pub const AtomicFile = struct {
    io: std.Io,
    /// Directory the rename happens within. Both temp and destination
    /// are resolved against it, which is what makes the rename atomic:
    /// POSIX only guarantees atomicity within one filesystem.
    dir: std.Io.Dir,
    dest_sub_path: []const u8,
    /// Temp name storage. Held inline (not a slice into elsewhere) so
    /// the struct stays movable; reconstruct the slice via `tmpPath()`
    /// rather than caching a pointer into this array.
    tmp_buf: [tmp_name_len]u8,
    tmp_len: u8,
    file: std.Io.File,
    file_writer: std.Io.File.Writer,
    /// Set by `finish`. Gates `deinit` so a successful rename is not
    /// followed by deleting the file we just published.
    finished: bool,
    /// Set the moment `finish` closes the fd, BEFORE the rename can
    /// fail. Without it a rename failure leaves `finished == false` and
    /// `deinit` closes the same fd twice — EBADF, which Debug builds
    /// escalate to a panic instead of the caller's error path.
    file_closed: bool,

    /// 0.16 removed `std.crypto.random`, and randomness is not actually
    /// needed here: an exclusive create is itself atomic, so walking a
    /// counter and taking the first name the filesystem grants us gives
    /// collision-freedom without any entropy source. `next_suffix` only
    /// biases the starting point so concurrent savers in one process
    /// rarely probe the same name twice.
    const prefix = ".ztmp-";
    const max_digits = 10;
    const tmp_name_len = prefix.len + max_digits;
    const max_probes = 1024;

    var next_suffix: std.atomic.Value(u32) = .init(0);

    pub const InitError = std.Io.File.OpenError || error{TempNameExhausted};
    pub const FinishError = std.Io.File.Writer.Error || std.Io.Dir.RenameError;

    pub fn init(
        io: std.Io,
        dest_path: []const u8,
        write_buffer: []u8,
    ) InitError!AtomicFile {
        // Resolve the temp file into the *destination's* directory.
        // Putting it in /tmp would make the rename cross-device and
        // therefore non-atomic (and often EXDEV).
        const dir_path = std.fs.path.dirname(dest_path) orelse ".";
        const base = std.fs.path.basename(dest_path);

        var dir = try std.Io.Dir.cwd().openDir(io, dir_path, .{});
        errdefer dir.close(io);

        var self: AtomicFile = .{
            .io = io,
            .dir = dir,
            .dest_sub_path = base,
            .tmp_buf = undefined,
            .tmp_len = 0,
            .file = undefined,
            .file_writer = undefined,
            .finished = false,
            .file_closed = false,
        };

        @memcpy(self.tmp_buf[0..prefix.len], prefix);
        var probe: u32 = 0;
        self.file = while (probe < max_probes) : (probe += 1) {
            const n = next_suffix.fetchAdd(1, .monotonic);
            const name = std.fmt.bufPrint(
                self.tmp_buf[prefix.len..],
                "{d}",
                .{n % 1_000_000_000},
            ) catch unreachable; // max_digits sizes the buffer
            self.tmp_len = @intCast(prefix.len + name.len);

            break dir.createFile(io, self.tmpPath(), .{ .exclusive = true }) catch |err| switch (err) {
                // Someone else holds this name — take the next one.
                error.PathAlreadyExists => continue,
                else => return err,
            };
        } else return error.TempNameExhausted;

        self.file_writer = self.file.writer(io, write_buffer);
        return self;
    }

    fn tmpPath(self: *const AtomicFile) []const u8 {
        return self.tmp_buf[0..self.tmp_len];
    }

    /// Flush buffered bytes, then atomically publish the temp file over
    /// the destination. After this returns, `deinit` is a no-op beyond
    /// releasing handles.
    pub fn finish(self: *AtomicFile) FinishError!void {
        try self.file_writer.flush();
        self.file.close(self.io);
        self.file_closed = true;
        try self.dir.rename(self.tmpPath(), self.dir, self.dest_sub_path, self.io);
        self.finished = true;
    }

    /// Safe to call whether or not `finish` succeeded. On the failure
    /// path it removes the temp file so a crashed/errored save leaves
    /// no debris next to the user's workbook.
    pub fn deinit(self: *AtomicFile) void {
        if (!self.finished) {
            if (!self.file_closed) self.file.close(self.io);
            // Best-effort: if unlink fails there is nothing useful to
            // report from a deinit, and the original file is untouched
            // either way.
            self.dir.deleteFile(self.io, self.tmpPath()) catch {};
        }
        self.dir.close(self.io);
    }
};

test "atomic file publishes on finish" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();

    const path = try tmp.dir.realPathFileAlloc(io, ".", std.testing.allocator);
    defer std.testing.allocator.free(path);
    const dest = try std.fs.path.join(std.testing.allocator, &.{ path, "out.bin" });
    defer std.testing.allocator.free(dest);

    var write_buf: [64]u8 = undefined;
    var af = try AtomicFile.init(io, dest, &write_buf);
    defer af.deinit();
    try af.file_writer.interface.writeAll("hello");
    try af.finish();

    var read_buf: [16]u8 = undefined;
    const n = try std.Io.Dir.cwd().readFile(io, dest, &read_buf);
    try std.testing.expectEqualStrings("hello", n);
}

test "atomic file leaves no debris when abandoned" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();

    // `.iterate = true` is required to call `iterate()` below. macOS
    // tolerates a non-iterable handle; Linux returns BADF and Windows
    // ACCESS_DENIED, so omitting it passes locally and fails in CI.
    var tmp = std.testing.tmpDir(.{ .iterate = true });
    defer tmp.cleanup();

    const path = try tmp.dir.realPathFileAlloc(io, ".", std.testing.allocator);
    defer std.testing.allocator.free(path);
    const dest = try std.fs.path.join(std.testing.allocator, &.{ path, "never.bin" });
    defer std.testing.allocator.free(dest);

    {
        var write_buf: [64]u8 = undefined;
        var af = try AtomicFile.init(io, dest, &write_buf);
        defer af.deinit(); // abandoned without finish()
        try af.file_writer.interface.writeAll("partial");
    }

    // Destination was never created...
    try std.testing.expectError(
        error.FileNotFound,
        std.Io.Dir.cwd().access(io, dest, .{}),
    );
    // ...and the temp file was cleaned up, so the directory is empty.
    var it = tmp.dir.iterate();
    try std.testing.expectEqual(@as(?std.Io.Dir.Entry, null), try it.next(io));
}
