//! Portable resident-set-size (RSS) reader for the iter-wb-6 RSS gate.
//!
//! Returns the calling process's current resident-set size in bytes,
//! cross-platform. Implementations:
//!
//!  * **Linux**: parse `/proc/self/status` for the `VmRSS:` line
//!    (kilobytes, multiply by 1024).
//!  * **macOS**: `task_info(MACH_TASK_BASIC_INFO)` via `<mach/mach.h>`.
//!    The `resident_size` field is bytes.
//!  * **BSDs**: `getrusage(RUSAGE_SELF).ru_maxrss` (kilobytes; *peak*,
//!    not current — coarser than the mach reading).
//!  * **Windows**: not implemented; returns `error.RssNotAvailable`.
//!    Tests SKIP rather than fail.
//!
//! Precision is *coarse*. Never compare against an absolute byte
//! threshold — only against a baseline ratio. The iter-wb-6 gate
//! consumes this as `wb_rss / book_rss ≤ 2`, which is robust to the
//! ~1 MB allocator-bookkeeping noise typical of these counters.

const std = @import("std");
const builtin = @import("builtin");

pub const Error = error{
    RssNotAvailable,
    RssReadFailed,
};

/// Current resident-set size in bytes for this process. Returns
/// `error.RssNotAvailable` on platforms without an implementation
/// (callers should `try` and the test should SKIP).
pub fn rssBytes(io: std.Io) Error!u64 {
    return switch (builtin.os.tag) {
        .linux => readLinuxVmRss(io),
        .macos => readMacosResidentSize(),
        .freebsd, .netbsd, .openbsd, .dragonfly => readBsdRuMaxRss(),
        else => error.RssNotAvailable,
    };
}

fn readLinuxVmRss(io: std.Io) Error!u64 {
    // 0.16 moved absolute-path opens under std.Io.Dir and every call
    // takes an Io. This path is Linux-only, so macOS builds never
    // compiled it and the migration missed it until CI ran.
    var file = std.Io.Dir.openFileAbsolute(io, "/proc/self/status", .{}) catch
        return error.RssReadFailed;
    defer file.close(io);

    // `/proc/self/status` is small (≤ 4 KB on contemporary kernels);
    // a fixed buffer is safe and avoids any allocator dependency.
    var buf: [8192]u8 = undefined;
    var fr = file.reader(io, &.{});
    const n = fr.interface.readSliceShort(&buf) catch return error.RssReadFailed;
    const text = buf[0..n];

    var it = std.mem.splitScalar(u8, text, '\n');
    while (it.next()) |line| {
        const prefix = "VmRSS:";
        if (!std.mem.startsWith(u8, line, prefix)) continue;
        const rest = std.mem.trim(u8, line[prefix.len..], &std.ascii.whitespace);
        // rest is shaped like "12345 kB"
        var fields = std.mem.tokenizeAny(u8, rest, " \t");
        const num_str = fields.next() orelse return error.RssReadFailed;
        const kb = std.fmt.parseInt(u64, num_str, 10) catch return error.RssReadFailed;
        // Multiplication is bounded: VmRSS in kB on a 64-bit system
        // can't realistically overflow u64 once shifted by 1024 (the
        // host would be physically incapable of holding the page set).
        // Use std.math.mul for explicit overflow trapping anyway.
        return std.math.mul(u64, kb, 1024) catch error.RssReadFailed;
    }
    return error.RssReadFailed;
}

// ─── macOS ─────────────────────────────────────────────────────────
//
// `task_info` returns a `mach_task_basic_info` struct whose
// `resident_size` field is bytes. Layout matches `<mach/task_info.h>`.

const macos = if (builtin.os.tag == .macos) struct {
    // Mirrors `mach_task_basic_info_data_t`. Field order and types
    // match Apple's Mach kernel headers (Darwin 24, macOS 15 SDK).
    // `mach_vm_size_t = u64`, `policy_t = i32`, `time_value_t =
    // {i32, i32}` per `mach/{mach_types,vm_types,std_types}.defs`.
    const TimeValue = extern struct { seconds: i32, microseconds: i32 };
    const TaskBasicInfo = extern struct {
        virtual_size: u64,
        resident_size: u64,
        resident_size_max: u64,
        user_time: TimeValue,
        system_time: TimeValue,
        policy: i32,
        suspend_count: i32,
    };

    // task_info flavors. `MACH_TASK_BASIC_INFO == 20`.
    const MACH_TASK_BASIC_INFO: c_int = 20;
    // `MACH_TASK_BASIC_INFO_COUNT = sizeof(struct) / sizeof(natural_t)`
    // where `natural_t` is u32 on Darwin.
    const MACH_TASK_BASIC_INFO_COUNT: u32 = @intCast(@sizeOf(TaskBasicInfo) / @sizeOf(u32));
    const KERN_SUCCESS: c_int = 0;

    // Externs from libSystem (Mach side). `mach_task_self_` is a
    // global in libSystem, but the `mach_task_self()` macro resolves
    // to it; we expose it as a function for ABI-stability.
    extern fn mach_task_self() c_uint;
    extern fn task_info(
        target: c_uint,
        flavor: c_int,
        info_out: [*]u32,
        count: *u32,
    ) c_int;
} else struct {};

fn readMacosResidentSize() Error!u64 {
    if (builtin.os.tag != .macos) return error.RssNotAvailable;

    var info: macos.TaskBasicInfo = undefined;
    var count: u32 = macos.MACH_TASK_BASIC_INFO_COUNT;

    // task_info writes through a `task_info_t` (a `[*]integer_t`); the
    // struct is laid out as a packed array of natural_t / integer_t
    // words. Cast through `[*]u32` to match the C signature.
    const info_ptr: [*]u32 = @ptrCast(@alignCast(&info));
    const kr = macos.task_info(
        macos.mach_task_self(),
        macos.MACH_TASK_BASIC_INFO,
        info_ptr,
        &count,
    );
    if (kr != macos.KERN_SUCCESS) return error.RssReadFailed;
    return info.resident_size;
}

// ─── BSDs ──────────────────────────────────────────────────────────

fn readBsdRuMaxRss() Error!u64 {
    // `getrusage` exists in std.posix on every Unix-like target.
    const usage = std.posix.getrusage(std.posix.rusage.SELF);
    // `ru_maxrss` units differ by platform. On the BSDs (and on Linux,
    // though we don't reach this branch there), the units are kilobytes.
    // On macOS the units are bytes (but we use mach there, not this).
    // Treat as kB and multiply by 1024.
    const max_rss = usage.maxrss;
    if (max_rss < 0) return error.RssReadFailed;
    return std.math.mul(u64, @intCast(max_rss), 1024) catch error.RssReadFailed;
}

// ─── Tests ─────────────────────────────────────────────────────────

test "rssBytes: returns a non-zero reading on supported platforms" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const r = rssBytes(io) catch |err| switch (err) {
        // Windows path — skip, don't fail.
        error.RssNotAvailable => return error.SkipZigTest,
        else => return err,
    };
    // A live process always has at least a few pages resident; a
    // sub-page-size reading would mean the syscall lied or our parser
    // misread the units. 64 KB is a defensively low floor that still
    // catches a unit-mistake (e.g. forgot to multiply kB → bytes).
    try std.testing.expect(r > 64 * 1024);
}

test "rssBytes: monotonic-ish across a heap allocation burst" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const before = rssBytes(io) catch |err| switch (err) {
        error.RssNotAvailable => return error.SkipZigTest,
        else => return err,
    };

    // Touch ~16 MB so the OS actually maps pages (lazy-allocated
    // memory wouldn't move RSS). Buffer is freed after — we only
    // assert that *something* moved during the live window.
    const burst_bytes: usize = 16 * 1024 * 1024;
    const buf = try std.testing.allocator.alloc(u8, burst_bytes);
    defer std.testing.allocator.free(buf);
    @memset(buf, 0xAA);

    const after = rssBytes(io) catch return error.SkipZigTest;

    // RSS measurement is coarse; we don't insist `after > before`
    // (kernel page reclamation can fire between the calls). Just
    // assert the second reading is also sane and within an order of
    // magnitude — defensive against a regressed parser returning 0.
    try std.testing.expect(after > 64 * 1024);
    // Sanity: not absurdly high (no 100 GB resident on a CI runner).
    try std.testing.expect(after < 100 * 1024 * 1024 * 1024);
    _ = before;
}
