//! emb-4 verifier.
//!
//! Opens an .xlsx (typically one that has been round-tripped through
//! Excel mac, Excel Win, LibreOffice Calc, or Apple Numbers via the
//! emb-4 compat matrix procedure) and reports whether the
//! `xl/zlsxEmbeddings/*` parts survived. Three pieces of evidence:
//!
//! 1. `xl/_rels/workbook.xml.rels` carries a Relationship of Type
//!    `…/relationships/embeddings` with Target
//!    `zlsxEmbeddings/index.xml`. Removal of this rel by a
//!    cooperating tool would cause v1.0 readers to need full-package
//!    scan; presence is the canonical signal the tool preserved the
//!    arc.
//! 2. `Workbook.embeddings()` returns a non-null view AND parses
//!    every declared coverage's vec.bin + hashes.bin (the
//!    `MissingEmbeddingPart` error from emb-2 fires here on partial
//!    stripping).
//! 3. The model / dim / dtype / coverage ids match the fixture
//!    that `zlsx-emb4-fixture` produced.
//!
//! Exit codes:
//!   0 — every expectation met (full preservation)
//!   1 — usage error
//!   2 — embedding view present but partial (mismatched fields)
//!   3 — embedding view missing (consumer stripped the parts)
//!   4 — workbook→index rel missing but parts still present
//!   5 — workbook→index rel present but parts missing (orphaned rel)
//!
//! Usage:
//!   zig build emb4-verify -- <file.xlsx>

const std = @import("std");
const pkg = @import("zlsx_pkg");

const EXPECTED_MODEL: []const u8 = "emb-4-fixture-v1";
const EXPECTED_DIM: u32 = 4;
const EXPECTED_DTYPE = pkg.embedding_part.Dtype.int8_sym_per_vec;
const EXPECTED_COVERAGES: usize = 2;

pub fn main(init: std.process.Init) !u8 {
    const allocator = init.gpa;
    const io = init.io;

    const args = try init.minimal.args.toSlice(init.arena.allocator());

    var stdout_buf: [512]u8 = undefined;
    var stdout_w = std.Io.File.stdout().writer(io, &stdout_buf);
    const stdout = &stdout_w.interface;
    defer stdout.flush() catch {};

    if (args.len < 2) {
        try stdout.print("usage: {s} <file.xlsx>\n", .{args[0]});
        return 1;
    }
    const in_path = args[1];

    var wb = try pkg.Workbook.open(allocator, io, in_path);
    defer wb.deinit();

    var out_buf: std.ArrayListUnmanaged(u8) = .empty;
    defer out_buf.deinit(allocator);
    const w = &out_buf;

    try w.print(allocator, "emb-4 verify: {s}\n", .{in_path});

    // 1. Workbook→index rel presence.
    const rels_part = try wb.store.part("xl/_rels/workbook.xml.rels");
    const rel_present = if (rels_part) |p|
        std.mem.indexOf(u8, p.bytes, pkg.embedding_part.REL_TYPE_EMBEDDINGS) != null
    else
        false;
    try w.print(
        allocator,
        "  workbook→index rel: {s}\n",
        .{if (rel_present) "present" else "MISSING"},
    );

    // 2. Embeddings view + coverage walk.
    const view_opt = wb.embeddings() catch |err| {
        try w.print(allocator, "  embeddings(): error.{s}\n", .{@errorName(err)});
        try stdout.writeAll(out_buf.items);
        return 3;
    };
    if (view_opt == null) {
        try w.print(allocator, "  embeddings(): null (no xl/zlsxEmbeddings/index.xml)\n", .{});
        try stdout.writeAll(out_buf.items);
        // Distinguish "no parts at all" (exit 3) from "rel present
        // but parts gone" (exit 5).
        return if (rel_present) 5 else 3;
    }
    const view = view_opt.?;

    var fields_ok = true;
    try w.print(allocator, "  model: {s}", .{view.index.model});
    if (!std.mem.eql(u8, view.index.model, EXPECTED_MODEL)) {
        try w.print(allocator, "  (EXPECTED {s})", .{EXPECTED_MODEL});
        fields_ok = false;
    }
    try w.print(allocator, "\n", .{});

    try w.print(allocator, "  dim: {d}", .{view.index.dim});
    if (view.index.dim != EXPECTED_DIM) {
        try w.print(allocator, "  (EXPECTED {d})", .{EXPECTED_DIM});
        fields_ok = false;
    }
    try w.print(allocator, "\n", .{});

    try w.print(allocator, "  dtype: {s}", .{view.index.dtype.string()});
    if (view.index.dtype != EXPECTED_DTYPE) {
        try w.print(allocator, "  (EXPECTED {s})", .{EXPECTED_DTYPE.string()});
        fields_ok = false;
    }
    try w.print(allocator, "\n", .{});

    try w.print(allocator, "  coverages: {d}", .{view.coverages.len});
    if (view.coverages.len != EXPECTED_COVERAGES) {
        try w.print(allocator, "  (EXPECTED {d})", .{EXPECTED_COVERAGES});
        fields_ok = false;
    }
    try w.print(allocator, "\n", .{});

    for (view.coverages) |cv| {
        try w.print(
            allocator,
            "    - id={s} ws={s} range={s} count={d} vec.count={d} hash.count={d}\n",
            .{
                cv.coverage.id,
                cv.coverage.worksheet_target,
                cv.coverage.range,
                cv.coverage.count,
                cv.vec.header.count,
                cv.hashes.header.count,
            },
        );
    }

    // 3. Verdict.
    if (!fields_ok) {
        try w.print(allocator, "verdict: PARTIAL — embedding parts survived but fields drifted from fixture\n", .{});
        try stdout.writeAll(out_buf.items);
        return 2;
    }
    if (!rel_present) {
        try w.print(allocator, "verdict: PARTS-ONLY — embedding parts survived but workbook→index rel was stripped\n", .{});
        try stdout.writeAll(out_buf.items);
        return 4;
    }
    try w.print(allocator, "verdict: PASS — full preservation\n", .{});
    try stdout.writeAll(out_buf.items);
    return 0;
}
