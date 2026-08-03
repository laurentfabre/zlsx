//! `zlsx-oracle-record` — turn a recalculated workbook into a manifest
//! (M1b, `goal_formula.md` §8.2).
//!
//! This is the step between "an application saved a file" and "we have
//! a golden". It extracts, checks the sentinels, and only then writes a
//! manifest — so a run that never recalculated cannot become evidence.
//! The check is not advisory: a rejected run exits non-zero and writes
//! nothing at all.
//!
//! Usage:
//!   zlsx-oracle-record <recalculated.xlsx> <provenance.json> <case> <fidelity> <out.json>

const std = @import("std");
const extractor = @import("extractor.zig");
const manifest = @import("manifest.zig");
const provenance = @import("provenance.zig");
const sentinel = @import("sentinel.zig");
const sentinel_set = @import("sentinel_set.zig");

const usage =
    \\usage: zlsx-oracle-record <recalculated.xlsx> <provenance.json> <case> <fidelity> <out.json>
    \\
    \\  fidelity: excel | ieee
    \\
    \\Exit codes:
    \\  0  manifest written
    \\  1  usage or I/O error
    \\  2  run REJECTED by the sentinel check — nothing written
    \\
;

pub fn main(init: std.process.Init) !u8 {
    var gpa: std.heap.DebugAllocator(.{}) = .init;
    defer _ = gpa.deinit();
    const allocator = gpa.allocator();

    // 0.16 hands main its Io and argv through `std.process.Init`;
    // `std.process.argsAlloc` is gone, and argv's lifetime is the
    // process arena's.
    const io = init.io;
    const args = try init.minimal.args.toSlice(init.arena.allocator());

    var stderr_buf: [4096]u8 = undefined;
    var stderr_file = std.Io.File.stderr().writer(io, &stderr_buf);
    const err = &stderr_file.interface;
    defer err.flush() catch {};

    if (args.len != 6) {
        try err.writeAll(usage);
        return 1;
    }

    const workbook_path = args[1];
    const provenance_path = args[2];
    const case_name = args[3];
    const fidelity = manifest.Fidelity.parse(args[4]) catch {
        try err.print("unknown fidelity '{s}' (expected excel | ieee)\n", .{args[4]});
        return 1;
    };
    const out_path = args[5];

    const workbook_bytes = try std.Io.Dir.cwd().readFileAlloc(
        io,
        workbook_path,
        allocator,
        .limited(64 << 20),
    );
    defer allocator.free(workbook_bytes);

    const provenance_bytes = try std.Io.Dir.cwd().readFileAlloc(
        io,
        provenance_path,
        allocator,
        .limited(1 << 20),
    );
    defer allocator.free(provenance_bytes);

    const parsed_provenance = try std.json.parseFromSlice(
        provenance.Record,
        allocator,
        provenance_bytes,
        // `alloc_always` so the record owns its strings rather than
        // borrowing from `provenance_bytes` — see `manifest.parse`.
        .{ .ignore_unknown_fields = true, .allocate = .alloc_always },
    );
    defer parsed_provenance.deinit();
    const prov = parsed_provenance.value;
    try prov.validate();

    var wb = try extractor.extract(allocator, workbook_bytes);
    defer wb.deinit();

    // The provenance claims a digest; the file in front of us has one.
    // A mismatch means the recorded evidence and the recorded identity
    // are of different files, which makes the manifest untraceable.
    if (!std.mem.eql(u8, prov.workbook_digest, &wb.digest)) {
        try err.print(
            "provenance digest does not match the workbook:\n  claimed:  {s}\n  actual:   {s}\n",
            .{ prov.workbook_digest, wb.digest },
        );
        return 1;
    }

    const adapter = try prov.adapterEnum();
    const sentinels = sentinel_set.forAdapter(adapter);
    if (sentinels.len > 0) {
        const verdict = try sentinel.check(allocator, sentinels, wb);
        defer sentinel.freeVerdict(allocator, verdict);
        switch (verdict) {
            .accepted => {},
            .rejected => |failures| {
                const text = try sentinel.explain(allocator, failures);
                defer allocator.free(text);
                try err.writeAll(text);
                return 2;
            },
        }
    }

    var arena: std.heap.ArenaAllocator = .init(allocator);
    defer arena.deinit();
    const caps = @import("adapters.zig").get(adapter);
    const m = try manifest.build(arena.allocator(), wb, .{
        .case = case_name,
        .fidelity = fidelity,
        .prov = prov,
        .exclude_volatiles = caps.excludes_volatiles,
        .exclude_char_code_high = caps.excludes_char_code_high,
    });
    try m.validate();

    var out: std.Io.Writer.Allocating = .init(allocator);
    defer out.deinit();
    // Pretty-printed on purpose: a manifest is reviewed by a human
    // before it becomes a golden, and diffed by one when it changes.
    // Null optionals are omitted for the same reason — a cell entry
    // padded with six `null`s buries the one field that matters.
    try std.json.Stringify.value(m, .{
        .whitespace = .indent_2,
        .emit_null_optional_fields = false,
    }, &out.writer);
    try out.writer.writeByte('\n');

    try std.Io.Dir.cwd().writeFile(io, .{ .sub_path = out_path, .data = out.written() });

    var stdout_buf: [1024]u8 = undefined;
    var stdout_file = std.Io.File.stdout().writer(io, &stdout_buf);
    const w = &stdout_file.interface;
    try w.print("wrote {s}: {d} cells ({d} asserted) from {s}\n", .{
        out_path,
        m.cells.len,
        m.assertedCount(),
        prov.adapter,
    });
    try w.flush();
    return 0;
}
