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
// M5d1: only for `control.inject`, the shared `std.Io` fault double. The
// commit region itself takes no `Poller` — that is the point of it.
const control = @import("zlsx_control");

/// What `syncDir` reports. A success value, always — §5.7.9 forbids the
/// post-commit step from producing an error, so the only thing it can say
/// is "committed, but the directory entry is not yet durable".
///
/// Shaped to be copied straight into `recalc_txn.Report.durability`, whose
/// slot is preallocated with the report precisely because this outcome is
/// discovered after all allocation must be done.
pub const Commit = struct {
    durability_warning: bool = false,
    /// POSIX errno of the failed directory fsync, 0 when clean. An errno
    /// rather than a Zig error so the C ABI and the Python binding can
    /// forward it without a translation table each.
    durability_errno: i32 = 0,
};

/// The report's slot is an errno, so the Zig error has to become one.
/// Only the codes `File.SyncError` can actually carry are mapped; the
/// catch-all is `EIO`, which is what a caller acts on anyway ("the
/// filesystem did not confirm the write").
fn errnoFor(err: std.Io.File.SyncError) i32 {
    return @intFromEnum(switch (err) {
        error.InputOutput => std.posix.E.IO,
        error.NoSpaceLeft => std.posix.E.NOSPC,
        error.DiskQuota => std.posix.E.DQUOT,
        error.AccessDenied => std.posix.E.ACCES,
        error.Canceled => std.posix.E.CANCELED,
        else => std.posix.E.IO,
    });
}

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
    pub const FinishError = std.Io.File.Writer.Error ||
        std.Io.File.SyncError ||
        std.Io.Dir.RenameError;

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

    /// §5.7.9's commit region: flush → **sync** → close → rename.
    ///
    /// The sync is what M5d1 added. Flushing moves the bytes out of the
    /// user-space buffer and into the kernel's; only `File.sync` gets them
    /// onto the medium. Without it the rename could commit a directory
    /// entry pointing at a file whose contents a power loss has not yet
    /// reached — the classic "atomic rename over unsynced data" hole,
    /// which on a crash yields a *zero-length or torn* workbook where the
    /// user's previous one used to be. Ordering it before the rename is
    /// the whole point: after the rename there is no failure left to
    /// report.
    ///
    /// **The rename is the commit point.** Every failure above it leaves
    /// the destination's prior bytes intact (or the destination still
    /// absent when it never existed) and the temp file deleted by
    /// `deinit`. Nothing below it may report failure as an error — see
    /// `syncDir`.
    ///
    /// **SLA exception 1 (§5.5)**: the sync is an uncancellable,
    /// untimed, blocking wait. It takes no `Poller` and has no deadline.
    /// A cancel token that fires while the kernel is flushing a
    /// half-gigabyte of dirty pages is observed at the *next* poll site,
    /// which is in the caller's next operation — there is no safe point
    /// inside `fsync(2)` to return from, and abandoning the wait would
    /// not un-issue the writeback anyway.
    pub fn finish(self: *AtomicFile) FinishError!void {
        try self.file_writer.flush();
        // Before the close: `sync` needs the descriptor. A failure here
        // leaves `file_closed == false`, so `deinit` closes it and
        // removes the temp file — no debris, destination untouched.
        try self.file.sync(self.io);
        self.file.close(self.io);
        self.file_closed = true;
        try self.dir.rename(self.tmpPath(), self.dir, self.dest_sub_path, self.io);
        self.finished = true;
    }

    /// §5.7.9's post-commit durability step, and the reason it cannot be
    /// an error.
    ///
    /// The rename has already committed: the destination now names the
    /// new file, and (for `saveWithRecalc`) the in-memory swap has
    /// already happened. Syncing the *directory* is what makes the new
    /// name itself survive a crash — on POSIX, `rename(2)` orders the
    /// entry but does not durably persist it. If that fsync fails, the
    /// file and memory are still consistent with each other and with what
    /// the caller asked for; only the guarantee that the *name* outlives
    /// a power cut is weakened. Reporting that as an error would tell a
    /// caller their save failed when it did not, so it comes back as a
    /// flag the report copies into its preallocated dormant slot
    /// (`recalc_txn.Durability`) without allocating.
    ///
    /// **SLA exception 2 (§5.5)**: like the file sync, uncancellable and
    /// untimed. There is additionally nothing a cancellation could mean
    /// here — the operation whose cancellation is being requested has
    /// already succeeded.
    ///
    /// Windows has no directory-fsync equivalent; there the rename's own
    /// metadata ordering is the guarantee, and this returns clean.
    pub fn syncDir(self: *AtomicFile) Commit {
        std.debug.assert(self.finished);
        if (@import("builtin").os.tag == .windows) return .{};
        // Routed through `std.Io.File.sync` rather than `std.posix.fsync`
        // so it crosses the same vtable seam as every other I/O this file
        // performs: one injection point covers both SLA exceptions, and
        // nothing test-only has to exist in the production path.
        const as_file: std.Io.File = .{
            .handle = self.dir.handle,
            .flags = .{ .nonblocking = false },
        };
        as_file.sync(self.io) catch |err| return .{
            .durability_warning = true,
            .durability_errno = errnoFor(err),
        };
        return .{};
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
    const commit = af.syncDir();
    try std.testing.expect(!commit.durability_warning);
    try std.testing.expectEqual(@as(i32, 0), commit.durability_errno);

    var read_buf: [16]u8 = undefined;
    const n = try std.Io.Dir.cwd().readFile(io, dest, &read_buf);
    try std.testing.expectEqualStrings("hello", n);
}

// ─── M5d1: the commit region under injected failure ──────────────────
//
// One helper for all four, because the interesting variable is *which*
// step fails and the rest of the fixture is identical: a destination that
// either exists with known prior bytes or does not exist at all, a save
// of different bytes over it, and an assertion about what survived.

const Fixture = struct {
    tmp: std.testing.TmpDir,
    /// Sentinel-terminated because `realPathFileAlloc` returns `[:0]u8`
    /// and it allocated `len + 1`. Storing it as a plain `[]u8` frees one
    /// byte short — which the debug allocator reports as a size mismatch,
    /// not as a leak.
    dir_path: [:0]u8,
    dest: []u8,

    /// Takes the caller's `io` rather than owning a `Threaded`: an
    /// `Io.Threaded` hands out an `Io` whose `userdata` points back at
    /// it, so returning one by value from a constructor leaves that
    /// pointer aimed at a dead frame.
    fn init(io: std.Io, prior: ?[]const u8) !Fixture {
        var self: Fixture = .{
            // `.iterate` so the debris assertions can walk the directory.
            .tmp = std.testing.tmpDir(.{ .iterate = true }),
            .dir_path = undefined,
            .dest = undefined,
        };
        self.dir_path = try self.tmp.dir.realPathFileAlloc(io, ".", std.testing.allocator);
        self.dest = try std.fs.path.join(std.testing.allocator, &.{ self.dir_path, "book.bin" });
        if (prior) |bytes| {
            try self.tmp.dir.writeFile(io, .{ .sub_path = "book.bin", .data = bytes });
        }
        return self;
    }

    fn deinit(self: *Fixture) void {
        std.testing.allocator.free(self.dest);
        std.testing.allocator.free(self.dir_path);
        self.tmp.cleanup();
    }

    /// Assert the destination holds exactly `expected`, or is absent when
    /// `expected` is null. §5.7.9 splits those two cases deliberately:
    /// "no output file" is only the promise for a destination that never
    /// existed.
    fn expectDest(self: *Fixture, io: std.Io, expected: ?[]const u8) !void {
        if (expected) |want| {
            var buf: [64]u8 = undefined;
            const got = try std.Io.Dir.cwd().readFile(io, self.dest, &buf);
            try std.testing.expectEqualStrings(want, got);
        } else {
            try std.testing.expectError(
                error.FileNotFound,
                std.Io.Dir.cwd().access(io, self.dest, .{}),
            );
        }
    }

    /// No `.ztmp-N` left next to the workbook, whatever failed.
    fn expectNoDebris(self: *Fixture, io: std.Io) !void {
        var it = self.tmp.dir.iterate();
        while (try it.next(io)) |entry| {
            try std.testing.expect(!std.mem.startsWith(u8, entry.name, ".ztmp-"));
        }
    }
};

test "M5d1: finish syncs before the rename" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const base = threaded.io();
    var fx = try Fixture.init(base, "OLD");
    defer fx.deinit();
    // Fail the first (and only) file sync — the temp file's.
    const io = control.inject.wrap(base, .{ .fail_file_sync_at = 1 });

    {
        var write_buf: [64]u8 = undefined;
        var af = try AtomicFile.init(io, fx.dest, &write_buf);
        defer af.deinit();
        try af.file_writer.interface.writeAll("NEW");
        try std.testing.expectError(error.InputOutput, af.finish());
    }

    // The sync failed, so the rename never ran: prior bytes intact. That
    // this test can fail the sync at all is the proof the sync happens —
    // and its position before the rename is what `expectDest` reads.
    try fx.expectDest(base, "OLD");
    try fx.expectNoDebris(base);
    try std.testing.expectEqual(@as(u64, 0), control.inject.state.rename_calls);
}

test "M5d1: a failed sync over an absent destination leaves it absent" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const base = threaded.io();
    var fx = try Fixture.init(base, null);
    defer fx.deinit();
    const io = control.inject.wrap(base, .{ .fail_file_sync_at = 1 });

    {
        var write_buf: [64]u8 = undefined;
        var af = try AtomicFile.init(io, fx.dest, &write_buf);
        defer af.deinit();
        try af.file_writer.interface.writeAll("NEW");
        try std.testing.expectError(error.InputOutput, af.finish());
    }

    try fx.expectDest(base, null);
    try fx.expectNoDebris(base);
}

test "M5d1: a failed rename is an error, and the destination is untouched" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const base = threaded.io();
    var fx = try Fixture.init(base, "OLD");
    defer fx.deinit();
    const io = control.inject.wrap(base, .{ .fail_rename = true });

    {
        var write_buf: [64]u8 = undefined;
        var af = try AtomicFile.init(io, fx.dest, &write_buf);
        defer af.deinit();
        try af.file_writer.interface.writeAll("NEW");
        try std.testing.expectError(error.AccessDenied, af.finish());
    }

    try fx.expectDest(base, "OLD");
    try fx.expectNoDebris(base);
    // The sync ran (once, on the temp file) and the rename was reached
    // and refused — so this is genuinely the rename failing, not an
    // earlier step short-circuiting.
    try std.testing.expectEqual(@as(u64, 1), control.inject.state.file_sync_calls);
    try std.testing.expectEqual(@as(u64, 1), control.inject.state.rename_calls);
}

test "M5d1: a failed post-commit dir fsync is a warning, not an error" {
    if (@import("builtin").os.tag == .windows) return error.SkipZigTest;

    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const base = threaded.io();
    var fx = try Fixture.init(base, "OLD");
    defer fx.deinit();
    // Sync #1 is the temp file's and must succeed; #2 is the directory's,
    // which happens only after the rename has committed.
    const io = control.inject.wrap(base, .{ .fail_file_sync_at = 2 });

    var commit: Commit = .{};
    {
        var write_buf: [64]u8 = undefined;
        var af = try AtomicFile.init(io, fx.dest, &write_buf);
        defer af.deinit();
        try af.file_writer.interface.writeAll("NEW");
        try af.finish();
        commit = af.syncDir();
    }

    // Committed: the new bytes are the destination's, and memory (here,
    // the caller's view of the outcome) agrees with the file.
    try fx.expectDest(base, "NEW");
    try std.testing.expect(commit.durability_warning);
    try std.testing.expectEqual(@intFromEnum(std.posix.E.IO), commit.durability_errno);

    // …and the outcome reaches M5b2's dormant slot without allocating:
    // two scalar stores into a struct that already exists.
    var report_slot: struct {
        warning: bool = false,
        err_code: i32 = 0,
        fn warn(self: *@This(), code: i32) void {
            self.warning = true;
            self.err_code = code;
        }
    } = .{};
    if (commit.durability_warning) report_slot.warn(commit.durability_errno);
    try std.testing.expect(report_slot.warning);
    try std.testing.expectEqual(@intFromEnum(std.posix.E.IO), report_slot.err_code);
}

test "M5d1: the two SLA exceptions take no poller and cannot fail the commit" {
    // §5.5 promises a poll at bounded work intervals inside long
    // operations. These two are the documented exceptions, and this test
    // is where the exception is written down in a form that breaks if
    // someone later "fixes" it.
    //
    // 1. `AtomicFile.finish`'s `File.sync` — an uncancellable, untimed
    //    blocking wait. There is no safe point inside `fsync(2)` to
    //    return from, and abandoning the wait would not un-issue the
    //    writeback.
    // 2. `AtomicFile.syncDir`'s directory fsync — same, plus there is
    //    nothing a cancellation could mean: per §5.7.9 the rename has
    //    already committed and the status is already success.
    //
    // The machine-checkable form of "cannot be polled" is that neither
    // signature has anywhere to put a poller, and that the second cannot
    // return an error at all.
    const finish_params = @typeInfo(@TypeOf(AtomicFile.finish)).@"fn".params;
    try std.testing.expectEqual(@as(usize, 1), finish_params.len);
    try std.testing.expectEqual(*AtomicFile, finish_params[0].type.?);

    const sync_dir = @typeInfo(@TypeOf(AtomicFile.syncDir)).@"fn";
    try std.testing.expectEqual(@as(usize, 1), sync_dir.params.len);
    try std.testing.expectEqual(*AtomicFile, sync_dir.params[0].type.?);
    // Not an error union: post-commit, failure is a flag on a success.
    try std.testing.expectEqual(Commit, sync_dir.return_type.?);
    try std.testing.expect(@typeInfo(Commit) == .@"struct");
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
