//! B1 iter-wb-6 RSS gate — orchestration test.
//!
//! Walk-away gate from `docs/plans/workbook-overlay.md`:
//!
//! > "100k-row × 10-sheet workbook needs ≤ 2× current `Book.openLazy`
//! >  RSS before any sheet is touched."
//!
//! Why two child processes? `cli_mod` + `zlsx_pkg` + `writer` cannot
//! coexist in a single Zig 0.15.2 compilation (documented in
//! `build.zig` and `AGENTS.md`: the same `src/writer.zig` file ends
//! up claimed by both `zlsx`'s path-import tree and `zlsx_pkg`'s
//! `addImport("writer", ...)` tree). The bench step therefore spawns
//! two purpose-built helper executables — one linked only against
//! `zlsx`, one only against `zlsx_pkg` — and each prints its own
//! post-open RSS delta. Comparing the two satisfies the gate without
//! ever needing both modules in one binary.
//!
//! This test is **off the default `zig build test` path**. It runs
//! only via `zig build bench-workbook-rss`. Two reasons:
//!   1. Synthesising the 100k × 10 fixture takes 1-3 minutes the
//!      first time (cached afterwards under `.zig-cache/`).
//!   2. RSS measurement is order-sensitive — running it alongside
//!      other tests in the same binary risks contamination.

const std = @import("std");
const builtin = @import("builtin");

const rss = @import("rss");

const FIXTURE_PATH = ".zig-cache/bench-100k-x-10.xlsx";

/// Workbook-vs-Book RSS ratio ceiling. Roadmap target was ≤ 2.0×.
/// Reality after the file-streaming PartStore: the Workbook overlay
/// uses LESS RSS than `Book.openLazy` because PartStore no longer
/// slurps the file (only EOCD + CD + structural parts come into
/// memory at open; everything else streams from disk on demand via
/// seek + readAll). Locally the ratio is ~0.78× on the 10k×10
/// fixture.
///
/// 1.5× locks in the win as a regression detector — any future
/// change that re-introduces a file slurp or eager decompress will
/// trip this immediately. Tighten further if headroom shrinks.
const RATIO_CEILING: f64 = 1.5;

fn currentRss() !u64 {
    return rss.rssBytes() catch |err| switch (err) {
        error.RssNotAvailable => return error.SkipZigTest,
        else => return err,
    };
}

/// Locate one of the three probe executables produced by the build
/// step. They install at `zig-out/bin/zlsx-bench-rss-{book,workbook,synth}`.
fn findProbe(allocator: std.mem.Allocator, io: std.Io, name: []const u8) ![]u8 {
    var path_buf: [256]u8 = undefined;
    const path = try std.fmt.bufPrint(&path_buf, "zig-out/bin/{s}", .{name});
    std.Io.Dir.cwd().access(io, path, .{}) catch return error.ProbeBinaryNotFound;
    return try allocator.dupe(u8, path);
}

/// Spawn the RSS probe (book or workbook backend) and parse the
/// single decimal line it prints to stdout.
fn runRssProbe(io: std.Io, exe: []const u8, fixture: []const u8) !u64 {
    var child = try std.process.spawn(io, .{
        .argv = &.{ exe, fixture },
        .stdout = .pipe,
        .stderr = .inherit,
    });

    const stdout = child.stdout.?;
    var buf: [256]u8 = undefined;
    var rbuf: [256]u8 = undefined;
    var reader = stdout.reader(io, &rbuf);
    const n = reader.interface.readSliceShort(&buf) catch |err| {
        child.kill(io);
        return err;
    };

    const term = try child.wait(io);
    switch (term) {
        .exited => |code| if (code != 0) return error.ProbeFailed,
        else => return error.ProbeFailed,
    }

    const text = std.mem.trim(u8, buf[0..n], &std.ascii.whitespace);
    return std.fmt.parseInt(u64, text, 10) catch error.ProbeOutputUnparseable;
}

/// Spawn the synth probe to (re)generate the cached fixture. Inherits
/// stdout/stderr so a long synth run isn't silent.
fn runSynthProbe(io: std.Io, exe: []const u8, out_path: []const u8) !void {
    // 0.16 replaced Child.init + child.spawn() with process.spawn(io,
    // options); stdio inheritance moved into SpawnOptions and the
    // allocator is no longer part of the call.
    var child = try std.process.spawn(io, .{
        .argv = &.{ exe, out_path },
        .stdout = .inherit,
        .stderr = .inherit,
    });
    const term = try child.wait(io);
    switch (term) {
        .exited => |code| if (code != 0) return error.SynthFailed,
        else => return error.SynthFailed,
    }
}

test "RSS gate: Workbook.openLazy ≤ 2× Book.openLazy on 100k × 10 fixture" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    // Probe RSS support up front so an unsupported platform skips
    // before the synth runs.
    _ = currentRss() catch return error.SkipZigTest;

    const allocator = std.testing.allocator;

    // 1. Locate the three probe binaries the bench step installs.
    const synth_exe = findProbe(allocator, io, "zlsx-bench-rss-synth") catch return error.SkipZigTest;
    defer allocator.free(synth_exe);
    const book_exe = findProbe(allocator, io, "zlsx-bench-rss-book") catch return error.SkipZigTest;
    defer allocator.free(book_exe);
    const wb_exe = findProbe(allocator, io, "zlsx-bench-rss-workbook") catch return error.SkipZigTest;
    defer allocator.free(wb_exe);

    // 2. Generate fixture (cached). On a fresh checkout this is the
    //    longest step; the probe self-skips when the file is present.
    try runSynthProbe(io, synth_exe, FIXTURE_PATH);
    const fixture_size = (try std.Io.Dir.cwd().statFile(io, FIXTURE_PATH, .{})).size;
    std.debug.print(
        "[rss-gate] fixture: {s} ({d:.2} MB)\n",
        .{ FIXTURE_PATH, @as(f64, @floatFromInt(fixture_size)) / (1024.0 * 1024.0) },
    );

    // 3. Two child invocations — one per backend. Each probe
    //    independently measures its own RSS delta.
    const book_rss = try runRssProbe(io, book_exe, FIXTURE_PATH);
    const wb_rss = try runRssProbe(io, wb_exe, FIXTURE_PATH);

    const book_mb = @as(f64, @floatFromInt(book_rss)) / (1024.0 * 1024.0);
    const wb_mb = @as(f64, @floatFromInt(wb_rss)) / (1024.0 * 1024.0);
    const ratio: f64 = if (book_rss == 0)
        std.math.inf(f64)
    else
        @as(f64, @floatFromInt(wb_rss)) / @as(f64, @floatFromInt(book_rss));

    std.debug.print(
        "[rss-gate] Book.openLazy delta: {d:.2} MB | Workbook.openLazy delta: {d:.2} MB | ratio: {d:.3}× (ceiling {d:.1}×)\n",
        .{ book_mb, wb_mb, ratio, RATIO_CEILING },
    );

    // Defensive: a zero baseline means the measurement was unreliable
    // (allocator returned scratch faster than we could re-poll, or
    // the kernel reclaimed pages aggressively between syscalls).
    // Skip rather than emit a false pass.
    if (book_rss == 0) {
        std.debug.print("[rss-gate] book_rss == 0; measurement unreliable, skipping\n", .{});
        return error.SkipZigTest;
    }

    try std.testing.expect(ratio <= RATIO_CEILING);
}

test "rss.rssBytes: returns a non-zero, sane reading" {
    const r = currentRss() catch return error.SkipZigTest;
    try std.testing.expect(r > 64 * 1024);
    // A test process pulling in zlsx + pkg shouldn't be using more
    // than 32 GB resident at unit-test time. Catches a units bug
    // that would otherwise hide behind a passing absolute check.
    try std.testing.expect(r < 32 * 1024 * 1024 * 1024);
}

comptime {
    _ = builtin;
}
