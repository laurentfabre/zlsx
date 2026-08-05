//! Cooperative cancellation, deadlines, and the 64 KiB seam every long
//! operation polls (§5.5, §5.7.9, §5.10).
//!
//! M5d1 of the tier-D1 ladder (`goal_formula.md`). §5.5 promises that a
//! cancel token or a deadline is observed "at bounded work intervals
//! inside long ops". Before this module there was no such interval: a
//! `DeflateFn` took `(allocator, input, out)` and compressed a 200 MiB
//! sheet in one call, so a token set the instant after the call began was
//! read for the first time when it ended. Same for decompression at
//! materialization, for the raw-entry copies `PartStore.save` streams, and
//! for the whole-archive serialization behind `Writer.saveToOwnedBuffer`.
//!
//! Three types, in dependency order:
//!
//!   - `Control` is §5.10's `{ cancel, deadline }` — what a *caller*
//!     supplies. It is the only one of the three that appears in a public
//!     signature.
//!   - `Watch` binds a `Control` to the `std.Io` its deadline is read
//!     against. A deadline is a timestamp on the `.awake` clock and
//!     reading that clock needs an `Io`; a cancel flag does not. Pairing
//!     them here is why `Control` itself can stay `io`-free.
//!   - `Poller` is the erased seam — a context pointer and the function
//!     that inspects it. Every long operation takes one of these, not a
//!     `Control`, for two reasons. `pkg/zip.zig` is stdlib-only and sits
//!     beneath the formula engine in the module graph, so it cannot name
//!     a `CancelToken`; and a test that wants to prove *where* a poll
//!     lands supplies its own context instead of trying to observe a
//!     volatile read it has no hook on.
//!
//! ## Why the token lives here and not in the evaluator
//!
//! `CancelToken` was `src/formula/run_inputs.zig`'s until M5d1. It moved
//! because `src/writer.zig` needs it — `saveToOwnedBufferControlled` is a
//! Writer method — and `writer_mod` does not import the formula engine.
//! Copying the union into the archive layer would have left two
//! structurally identical types that a conversion function has to keep in
//! sync forever, so instead the definition moved down to a std-only leaf
//! both trees already sit above. `run_inputs.zig` re-exports the name, so
//! nothing in the engine changed.
//!
//! ## What is NOT pollable
//!
//! Two waits in the save transaction are uncancellable by construction and
//! §5.7.9 says so: the blocking `File.sync` that makes the temp file
//! durable, and the POSIX directory fsync that runs *after* the rename has
//! already committed. Neither takes a `Poller`, neither has a timeout, and
//! `pkg/atomic_file.zig` names both. Returning `error.Cancelled` from the
//! second one would contradict a status that is already success.

const std = @import("std");

/// §5.5's bound, made a number. Every long operation polls at least once
/// per this many bytes of input.
///
/// 64 KiB is large enough that the poll is noise next to the work it
/// brackets (a deflate of the same chunk is four orders of magnitude more
/// expensive) and small enough that the worst-case latency between a
/// caller setting a flag and the operation noticing stays sub-millisecond
/// on any machine that can run a spreadsheet engine.
pub const chunk_bytes: usize = 64 * 1024;

/// The one error a poll produces. Deliberately spelled with two `l`s:
/// `std.Io.Cancelable` uses `Canceled`, and these are different outcomes.
/// The std one means an `Io` operation was cancelled by the runtime; this
/// one means the *caller's* token fired or their deadline passed, which is
/// what §12.2 maps to C `-5` and CLI 130/143.
pub const Error = error{Cancelled};

/// Cooperative cancellation, as a no-alloc union over the two storage
/// kinds a caller can actually have: an atomic for multi-threaded hosts,
/// and a plain volatile flag for a signal handler in a single-threaded
/// CLI (writing an atomic from a signal handler is not async-signal-safe
/// in general). One `isTriggered` seam, so every poll site is identical.
pub const CancelToken = union(enum) {
    atomic: *const std.atomic.Value(bool),
    /// `sig_atomic_t`-shaped: written by a signal handler, read here.
    flag: *const volatile u8,

    pub fn isTriggered(self: CancelToken) bool {
        return switch (self) {
            .atomic => |a| a.load(.acquire),
            .flag => |f| f.* != 0,
        };
    }
};

/// §5.10's caller-facing control. Both fields are excluded from
/// `EffectiveRunInputs` by construction — they are not in that struct, so
/// they cannot leak into a fingerprint or a cache key.
pub const Control = struct {
    cancel: ?CancelToken = null,
    /// Absolute, on the `.awake` clock of the same `std.Io` the operation
    /// was given. Zig 0.16 has no `Instant`; `std.Io.Timestamp` is the
    /// monotonic reading.
    deadline: ?std.Io.Timestamp = null,

    /// What every plain (non-`Controlled`) signature forwards.
    pub const none: Control = .{};

    /// True when at least one of the two can fire. A control that cannot
    /// fire is skipped entirely — see `Watch.poller`.
    pub fn armed(self: Control) bool {
        return self.cancel != null or self.deadline != null;
    }

    /// First to fire wins (§5.5). Cancel is checked first because it is a
    /// load and the deadline is a clock read; on the path where both are
    /// set and the flag is already up, the syscall is skipped.
    pub fn check(self: Control, io: std.Io) Error!void {
        if (self.cancel) |token| {
            if (token.isTriggered()) return Error.Cancelled;
        }
        if (self.deadline) |limit| {
            const now = std.Io.Timestamp.now(io, .awake);
            if (now.nanoseconds >= limit.nanoseconds) return Error.Cancelled;
        }
    }
};

/// A `Control` bound to the `Io` its deadline reads. Produces the
/// `Poller` the long operations actually take.
///
/// **Lifetime**: `poller()` hands out a pointer to this value, so the
/// `Watch` must outlive every operation that holds the result. Keep it in
/// the frame that owns the call.
pub const Watch = struct {
    io: std.Io,
    control: Control = .{},

    pub fn init(io: std.Io, control: Control) Watch {
        return .{ .io = io, .control = control };
    }

    /// A disarmed control produces `Poller.none` rather than a poller
    /// that calls back to discover there is nothing to check. That is
    /// what makes "the plain form is byte-identical to the controlled
    /// form under a null control" true by construction rather than by
    /// test: with no cancel and no deadline the two paths run the same
    /// instructions.
    pub fn poller(self: *const Watch) Poller {
        if (!self.control.armed()) return .none;
        return .{ .ctx = self, .check_fn = checkErased };
    }

    fn checkErased(ctx: ?*const anyopaque) Error!void {
        const self: *const Watch = @ptrCast(@alignCast(ctx.?));
        return self.control.check(self.io);
    }
};

/// The erased poll seam threaded through every long operation.
///
/// Context plus function rather than a `Control` value so that stdlib-only
/// modules beneath the engine can carry it, and so a test can supply a
/// context that trips on the Nth poll instead of trying to observe a
/// volatile load.
pub const Poller = struct {
    ctx: ?*const anyopaque = null,
    check_fn: ?*const fn (?*const anyopaque) Error!void = null,

    /// No cancel, no deadline, no call. The value every plain signature
    /// forwards and the value `Watch.poller` returns for a disarmed
    /// control.
    pub const none: Poller = .{};

    pub inline fn check(self: Poller) Error!void {
        const f = self.check_fn orelse return;
        return f(self.ctx);
    }

    /// Split `input` into `chunk_bytes`-sized pieces, polling before each.
    ///
    /// The poll happens *before* the chunk is handed out, so an operation
    /// that is already cancelled when it starts does no work at all — and
    /// a zero-length input polls zero times, which is what keeps the
    /// per-entry cost of a package full of tiny XML parts at zero.
    pub fn chunks(self: Poller, input: []const u8) Chunks {
        return .{ .poller = self, .rest = input };
    }
};

/// `Poller.chunks`' iterator. `while (try it.next()) |chunk|`.
pub const Chunks = struct {
    poller: Poller,
    rest: []const u8,

    pub fn next(self: *Chunks) Error!?[]const u8 {
        if (self.rest.len == 0) return null;
        try self.poller.check();
        const n = @min(self.rest.len, chunk_bytes);
        const out = self.rest[0..n];
        self.rest = self.rest[n..];
        return out;
    }
};

/// How many polls a correctly chunked operation owes for `len` bytes.
/// The measurement in M5d1's polling-bound test compares against this
/// rather than against a hand-computed constant.
pub fn chunkCount(len: usize) usize {
    return (len + chunk_bytes - 1) / chunk_bytes;
}

// ─── injection substrate (tests) ─────────────────────────────────────

/// The `std.Io` doubles M5d1's tests share.
///
/// Lives beside the seam rather than in any one test file because four
/// layers need the same two doubles — `pkg/atomic_file.zig`,
/// `pkg/store.zig`, `pkg/zip.zig` and `src/writer.zig` — and a copy per
/// layer is four copies to keep in sync with a stdlib vtable that moves.
///
/// The mechanism is `std.Io`'s own vtable: a wrapper keeps the base
/// `userdata` and replaces exactly the three entries M5d1 cares about, so
/// every other operation still reaches the real implementation. That is
/// also why there is no fault-injection hook anywhere in the production
/// code — the seam the tests drive is the one the process already has.
pub const inject = struct {
    pub const State = struct {
        base_vtable: *const std.Io.VTable = undefined,

        /// `now` reads. With a deadline armed this is exactly the number
        /// of polls, which is how the §5.5 bound is measured rather than
        /// asserted.
        now_calls: u64 = 0,
        /// Reported clock: `origin + step_ns * now_calls`. A step of zero
        /// is a clock that never advances, so a far-future deadline never
        /// fires however many times it is polled.
        origin_ns: i96 = 0,
        step_ns: i96 = 0,

        /// Once `now_calls` reaches this, set `trip_flag`. Turns "cancel
        /// arrives mid-operation" into something deterministic: the poll
        /// that reads the clock for the Nth time arms the flag, and the
        /// next poll sees it.
        trip_at: ?u64 = null,
        trip_flag: ?*volatile u8 = null,

        /// 1-based: fail the Nth `File.sync`. The temp-file sync is the
        /// first one a save performs and the post-commit directory fsync
        /// is the second, so `1` and `2` select the two SLA exceptions.
        fail_file_sync_at: ?u64 = null,
        file_sync_calls: u64 = 0,
        file_sync_error: std.Io.File.SyncError = error.InputOutput,

        fail_rename: bool = false,
        rename_calls: u64 = 0,
    };

    /// Thread-local so a parallel test runner cannot cross wires; the
    /// vtable copy is thread-local for the same reason.
    pub threadlocal var state: State = .{};
    threadlocal var vtable: std.Io.VTable = undefined;

    /// Install `initial` and return an `Io` that behaves like `base`
    /// except for the three injected entries.
    pub fn wrap(base: std.Io, initial: State) std.Io {
        state = initial;
        state.base_vtable = base.vtable;
        vtable = base.vtable.*;
        vtable.now = injectedNow;
        vtable.fileSync = injectedFileSync;
        vtable.dirRename = injectedRename;
        return .{ .userdata = base.userdata, .vtable = &vtable };
    }

    fn injectedNow(_: ?*anyopaque, _: std.Io.Clock) std.Io.Timestamp {
        state.now_calls += 1;
        if (state.trip_at) |n| {
            if (state.now_calls >= n) {
                if (state.trip_flag) |f| f.* = 1;
            }
        }
        return .{ .nanoseconds = state.origin_ns + state.step_ns * @as(i96, @intCast(state.now_calls)) };
    }

    fn injectedFileSync(userdata: ?*anyopaque, file: std.Io.File) std.Io.File.SyncError!void {
        state.file_sync_calls += 1;
        if (state.fail_file_sync_at) |n| {
            if (state.file_sync_calls == n) return state.file_sync_error;
        }
        return state.base_vtable.fileSync(userdata, file);
    }

    fn injectedRename(
        userdata: ?*anyopaque,
        old_dir: std.Io.Dir,
        old_sub_path: []const u8,
        new_dir: std.Io.Dir,
        new_sub_path: []const u8,
    ) std.Io.Dir.RenameError!void {
        state.rename_calls += 1;
        if (state.fail_rename) return error.AccessDenied;
        return state.base_vtable.dirRename(userdata, old_dir, old_sub_path, new_dir, new_sub_path);
    }
};

// ─── Tests ───────────────────────────────────────────────────────────

const testing = std.testing;

test "Poller.none never calls back" {
    var p: Poller = .none;
    try p.check();
    var it = p.chunks("abcdef");
    var seen: usize = 0;
    while (try it.next()) |c| seen += c.len;
    try testing.expectEqual(@as(usize, 6), seen);
}

test "chunks yields 64 KiB pieces and polls once per piece" {
    const Counter = struct {
        n: usize = 0,
        fn check(ctx: ?*const anyopaque) Error!void {
            const self: *const @This() = @ptrCast(@alignCast(ctx.?));
            @constCast(self).n += 1;
        }
    };
    var counter: Counter = .{};
    const p: Poller = .{ .ctx = &counter, .check_fn = Counter.check };

    const len = chunk_bytes * 3 + 7;
    const buf = try testing.allocator.alloc(u8, len);
    defer testing.allocator.free(buf);
    @memset(buf, 'x');

    var it = p.chunks(buf);
    var total: usize = 0;
    var pieces: usize = 0;
    while (try it.next()) |c| {
        try testing.expect(c.len <= chunk_bytes);
        total += c.len;
        pieces += 1;
    }
    try testing.expectEqual(len, total);
    try testing.expectEqual(@as(usize, 4), pieces);
    try testing.expectEqual(chunkCount(len), pieces);
    try testing.expectEqual(pieces, counter.n);
}

test "chunks stops at the piece the poller refuses" {
    const TripAt = struct {
        n: usize = 0,
        at: usize,
        fn check(ctx: ?*const anyopaque) Error!void {
            const self: *const @This() = @ptrCast(@alignCast(ctx.?));
            @constCast(self).n += 1;
            if (self.n >= self.at) return Error.Cancelled;
        }
    };
    var trip: TripAt = .{ .at = 3 };
    const p: Poller = .{ .ctx = &trip, .check_fn = TripAt.check };

    const buf = try testing.allocator.alloc(u8, chunk_bytes * 10);
    defer testing.allocator.free(buf);
    @memset(buf, 0);

    var it = p.chunks(buf);
    var consumed: usize = 0;
    while (true) {
        const c = it.next() catch |err| {
            try testing.expectEqual(Error.Cancelled, err);
            break;
        } orelse return error.PollerNeverRefused;
        consumed += c.len;
    }
    // Two chunks got through; the third poll refused before handing out
    // any bytes, so nothing past the boundary was touched.
    try testing.expectEqual(chunk_bytes * 2, consumed);
}

test "Control: cancel fires, and a disarmed control produces no poller" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var flag: u8 = 0;
    var armed: Watch = .init(io, .{ .cancel = .{ .flag = &flag } });
    try armed.poller().check();
    flag = 1;
    try testing.expectError(Error.Cancelled, armed.poller().check());

    const idle: Watch = .init(io, .{});
    try testing.expect(idle.poller().check_fn == null);
    try testing.expect(!Control.none.armed());
}

test "Control: atomic token and deadline both fire" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const base = threaded.io();

    var atomic: std.atomic.Value(bool) = .init(false);
    const w: Watch = .init(base, .{ .cancel = .{ .atomic = &atomic } });
    try w.poller().check();
    atomic.store(true, .release);
    try testing.expectError(Error.Cancelled, w.poller().check());

    // A clock that advances 1 ms per read against a 5 ms deadline: the
    // fifth poll is the one that trips.
    const io = inject.wrap(base, .{ .step_ns = std.time.ns_per_ms });
    const d: Watch = .init(io, .{ .deadline = .{ .nanoseconds = 5 * std.time.ns_per_ms } });
    var polls: usize = 0;
    while (polls < 10) : (polls += 1) {
        d.poller().check() catch break;
    }
    try testing.expectEqual(@as(usize, 4), polls);
    try testing.expectEqual(@as(u64, 5), inject.state.now_calls);
}

test "inject: the counting clock arms a cancel flag mid-run" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();

    var flag: u8 = 0;
    const io = inject.wrap(threaded.io(), .{
        .trip_at = 3,
        .trip_flag = &flag,
    });
    // Deadline armed so every poll reads the clock; the clock never
    // advances, so only the injected trip can end the loop.
    const w: Watch = .init(io, .{
        .cancel = .{ .flag = &flag },
        .deadline = .{ .nanoseconds = std.math.maxInt(i64) },
    });
    var polls: usize = 0;
    while (polls < 100) : (polls += 1) {
        w.poller().check() catch break;
    }
    try testing.expectEqual(@as(usize, 3), polls);
}
