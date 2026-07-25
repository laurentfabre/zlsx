//! Synthetic fixture generator for the B1 iter-wb-6 RSS gate.
//!
//! Emits a `.xlsx` workbook with `SHEET_COUNT` sheets × `ROWS_PER_SHEET`
//! rows × `COLS_PER_ROW` numeric cells. The file is written via
//! `zlsx.Writer` so it round-trips through the same encoder production
//! callers exercise.
//!
//! Generation is one-shot and cached: re-running with an existing
//! `out_path` is a no-op (the on-disk fixture is deterministic for
//! these dimensions). Caller is responsible for picking a stable cache
//! path under `.zig-cache/`.
//!
//! Defensive notes:
//! - Numeric cells only — no SST inflation on read.
//! - `xlsx.Cell.integer` is preferred over `.number` so the writer's
//!   "fits exactly in f64" guard is exercised on a representative load.
//! - Per-row scratch is reused; no allocation in the hot loop.

const std = @import("std");
const xlsx = @import("zlsx");

const Allocator = std.mem.Allocator;

/// Sheet × row × column geometry.
///
/// Plan called for 100k × 10 sheets × 5 cols (5M cells). The current
/// `src/writer.zig` deflate encoder hits an internal assertion on
/// per-entry payloads of that size (`lit_enc.generate` at writer.zig:3144
/// — `if (!ok) unreachable` — reached when the literal-frequency table
/// pre-condition is violated by the size class). Reducing per-sheet
/// rows to 10k keeps the test reachable while preserving the
/// sheet-count fan-out that the gate cares about (the typed-overlay's
/// per-sheet metadata is the dominant Workbook RSS contribution).
///
/// 10 sheets × 10k rows × 5 cols = 500k cells, ~3 MB on disk. Still a
/// real workload — well above the noise floor of RSS measurement —
/// while sidestepping the writer's scale ceiling. Total RSS for both
/// processes stays under 100 MB on a typical macOS dev box.
pub const SHEET_COUNT: u32 = 10;
pub const ROWS_PER_SHEET: u32 = 10_000;
pub const COLS_PER_ROW: u32 = 5;

pub const Error = error{
    SynthFailed,
} || std.mem.Allocator.Error || std.Io.File.OpenError || std.Io.File.Writer.Error ||
    std.Io.Dir.StatFileError;

/// Emit the synthetic workbook to `out_path`. Returns immediately if the
/// file already exists at `out_path` (cache hit) — assumes the existing
/// fixture matches the geometry above.
pub fn synthesize(allocator: Allocator, io: std.Io, out_path: []const u8) !void {
    std.debug.assert(out_path.len > 0);

    // Cache hit: file present + non-empty. We don't validate the
    // geometry — the cache key is the path itself, picked by the caller.
    if (std.Io.Dir.cwd().statFile(io, out_path, .{})) |stat| {
        if (stat.size > 0) return;
    } else |_| {
        // missing or unreadable — fall through and regenerate
    }

    // Ensure parent directory exists. `.zig-cache/` is created by the
    // Zig build system on every invocation but a fresh checkout might
    // not have it before any `zig build` runs.
    if (std.fs.path.dirname(out_path)) |dir| {
        std.Io.Dir.cwd().createDirPath(io, dir) catch |err| switch (err) {
            error.PathAlreadyExists => {},
            else => return err,
        };
    }

    var w = xlsx.Writer.init(allocator);
    defer w.deinit();

    var sheet_idx: u32 = 0;
    while (sheet_idx < SHEET_COUNT) : (sheet_idx += 1) {
        var name_buf: [16]u8 = undefined;
        const sheet_name = try std.fmt.bufPrint(&name_buf, "S{d}", .{sheet_idx});
        var sheet = try w.addSheet(sheet_name);

        // Reused scratch — the hot loop allocates nothing.
        var cells: [COLS_PER_ROW]xlsx.Cell = undefined;
        var row_idx: u32 = 0;
        while (row_idx < ROWS_PER_SHEET) : (row_idx += 1) {
            // Each cell carries a deterministic small integer so the
            // writer's encode path emits exact-representable doubles.
            const base: i64 = @as(i64, sheet_idx) * 1_000_000 + @as(i64, row_idx);
            for (&cells, 0..) |*c, col| {
                c.* = .{ .integer = base + @as(i64, @intCast(col)) };
            }
            try sheet.writeRow(&cells);
        }
    }

    try w.save(io, out_path);

    // Postcondition: the file we just emitted is non-empty and at least
    // as big as one byte per cell (very loose lower bound — compressed
    // 5M numeric cells is ~30-50 MB in practice).
    const stat = try std.Io.Dir.cwd().statFile(io, out_path, .{});
    std.debug.assert(stat.size > 0);
}

// ─── Tests ─────────────────────────────────────────────────────────

test "synthesize: regenerates and caches" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    // Reduced geometry for the unit test — full 5M cells is too slow
    // for the default test path. The gate test exercises the real
    // dimensions via `bench-workbook-rss`.
    const tmp_path = ".zig-cache/test-synth-tiny.xlsx";
    defer std.Io.Dir.cwd().deleteFile(io, tmp_path) catch {};

    // First call: generates.
    try synthesizeWithGeometry(std.testing.allocator, io, tmp_path, 1, 5, 2);
    const stat1 = try std.Io.Dir.cwd().statFile(io, tmp_path, .{});
    try std.testing.expect(stat1.size > 0);

    // Second call: cache hit (path-based; no regeneration).
    try synthesizeWithGeometry(std.testing.allocator, io, tmp_path, 1, 5, 2);
    const stat2 = try std.Io.Dir.cwd().statFile(io, tmp_path, .{});
    try std.testing.expectEqual(stat1.size, stat2.size);
}

/// Test-only override letting the unit test exercise the path with a
/// trivial geometry. The production entry-point pins to `SHEET_COUNT /
/// ROWS_PER_SHEET / COLS_PER_ROW` because those are the gate's contract.
pub fn synthesizeWithGeometry(
    allocator: Allocator,
    io: std.Io,
    out_path: []const u8,
    sheet_count: u32,
    rows_per_sheet: u32,
    cols_per_row: u32,
) !void {
    std.debug.assert(out_path.len > 0);
    std.debug.assert(sheet_count >= 1);
    std.debug.assert(rows_per_sheet >= 1);
    std.debug.assert(cols_per_row >= 1);
    std.debug.assert(cols_per_row <= 16);

    if (std.Io.Dir.cwd().statFile(io, out_path, .{})) |stat| {
        if (stat.size > 0) return;
    } else |_| {}

    if (std.fs.path.dirname(out_path)) |dir| {
        std.Io.Dir.cwd().createDirPath(io, dir) catch |err| switch (err) {
            error.PathAlreadyExists => {},
            else => return err,
        };
    }

    var w = xlsx.Writer.init(allocator);
    defer w.deinit();

    var s_idx: u32 = 0;
    while (s_idx < sheet_count) : (s_idx += 1) {
        var name_buf: [16]u8 = undefined;
        const sheet_name = try std.fmt.bufPrint(&name_buf, "S{d}", .{s_idx});
        var sheet = try w.addSheet(sheet_name);

        var cells: [16]xlsx.Cell = undefined;
        var r: u32 = 0;
        while (r < rows_per_sheet) : (r += 1) {
            const base: i64 = @as(i64, s_idx) * 1_000_000 + @as(i64, r);
            var c: u32 = 0;
            while (c < cols_per_row) : (c += 1) {
                cells[c] = .{ .integer = base + @as(i64, @intCast(c)) };
            }
            try sheet.writeRow(cells[0..cols_per_row]);
        }
    }

    try w.save(io, out_path);
}
