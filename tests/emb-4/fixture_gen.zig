//! emb-4 fixture generator.
//!
//! Produces a small but realistic .xlsx that exercises the
//! `setEmbeddings` writer surface end-to-end (workbook → index rel,
//! Content_Types overrides, two-coverage index, per-coverage vec.bin
//! + hashes.bin, idempotent re-write). The file is the input to the
//! emb-4 compat matrix: open in Excel mac, Excel Win, LibreOffice
//! Calc, and Apple Numbers, do a passive save in each, then run
//! `zlsx-emb4-verify <out.xlsx>` to confirm the embedding parts
//! survived.
//!
//! Usage:
//!   zig build emb4-fixture -- <out.xlsx>
//!
//! Workbook layout:
//!   sheet "Items"
//!     A1: "Title"    B1: "Body"
//!     A2: "Alpha"    B2: "First entry, body text alpha."
//!     A3: "Beta"     B3: "Second entry, body text beta."
//!     A4: "Gamma"    B4: "Third entry, body text gamma."
//!     A5: "Delta"    B5: "Fourth entry, body text delta."
//!
//! Embeddings:
//!   model "emb-4-fixture-v1", dim 4, dtype int8-sym-per-vec.
//!   Two coverages:
//!     title — worksheets/sheet1.xml @ A2:A5 (column A)
//!     body  — worksheets/sheet1.xml @ B2:B5 (column B)
//!   Vectors are deterministic synthetic values so the produced
//!   file is byte-stable across runs.

const std = @import("std");
const pkg = @import("zlsx_pkg");
const zlsx = @import("zlsx");

const Cell = zlsx.Cell;
const Dtype = pkg.embedding_part.Dtype;

pub fn main(init: std.process.Init) !u8 {
    const allocator = init.gpa;
    const io = init.io;

    const args = try init.minimal.args.toSlice(init.arena.allocator());

    var err_buf: [256]u8 = undefined;
    var err_w = std.Io.File.stderr().writer(io, &err_buf);
    const err = &err_w.interface;
    defer err.flush() catch {};

    if (args.len < 2) {
        try err.print("usage: {s} <out.xlsx>\n", .{args[0]});
        return 2;
    }
    const out_path = args[1];

    var wb = try pkg.Workbook.empty(allocator, io);
    defer wb.deinit();

    // Sheet with 1 header + 4 data rows. `addSheet` returns a *Worksheet;
    // the slot stays valid as long as no further `addSheet`/`deleteSheet`
    // calls run.
    const ws = try wb.addSheet("Items");

    const rows = [_][]const Cell{
        &.{ .{ .string = "Title" }, .{ .string = "Body" } },
        &.{ .{ .string = "Alpha" }, .{ .string = "First entry, body text alpha." } },
        &.{ .{ .string = "Beta" }, .{ .string = "Second entry, body text beta." } },
        &.{ .{ .string = "Gamma" }, .{ .string = "Third entry, body text gamma." } },
        &.{ .{ .string = "Delta" }, .{ .string = "Fourth entry, body text delta." } },
    };
    try ws.appendRows(&rows);

    // int8-sym-per-vec with dim=4 → 8 bytes per record (f32 scale + i8[4]).
    // 4 rows per coverage → 32 bytes per coverage body.
    const dim: u32 = 4;
    var title_body: [4 * (4 + 4)]u8 = undefined;
    var body_body: [4 * (4 + 4)]u8 = undefined;

    encodeInt8SymRows(&title_body, 0.5, [_]i8{ 10, -10, 5, -5 });
    encodeInt8SymRows(&body_body, 0.7, [_]i8{ 30, -30, 15, -15 });

    // Deterministic synthetic hashes — not real xxh3 of the cell text.
    // The compat matrix tests preservation, not hash validity; emb-6
    // CLI will produce real hashes via embedding_part.xxh3Canonical.
    const title_hashes = [_]u64{ 0xA000_0001, 0xA000_0002, 0xA000_0003, 0xA000_0004 };
    const body_hashes = [_]u64{ 0xB000_0001, 0xB000_0002, 0xB000_0003, 0xB000_0004 };

    const inputs = [_]pkg.EmbeddingCoverageInput{
        .{
            .id = "title",
            .worksheet_target = "worksheets/sheet1.xml",
            .range = "A2:A5",
            .column = "A",
            .include_formulas = false,
            .vec_body = &title_body,
            .hashes = &title_hashes,
        },
        .{
            .id = "body",
            .worksheet_target = "worksheets/sheet1.xml",
            .range = "B2:B5",
            .column = "B",
            .include_formulas = false,
            .vec_body = &body_body,
            .hashes = &body_hashes,
        },
    };
    try wb.setEmbeddings("emb-4-fixture-v1", dim, .int8_sym_per_vec, &inputs);

    try wb.save(io, out_path);

    // Mirror the path back so the build-step output is greppable in
    // CI logs / shell history.
    var stdout_buf: [512]u8 = undefined;
    var stdout_w = std.Io.File.stdout().writer(io, &stdout_buf);
    const out = &stdout_w.interface;
    defer out.flush() catch {};
    try out.print(
        "wrote emb-4 fixture: {s}\n  model=emb-4-fixture-v1 dim=4 dtype=int8-sym-per-vec coverages=title,body\n",
        .{out_path},
    );
    return 0;
}

/// Encode 4 rows of int8-sym-per-vec records into `out` (length 32):
/// each record is a little-endian f32 scale (4 bytes) followed by
/// `dim`-byte i8 array. All four rows share the same scale +
/// quantized values so the fixture is reproducible and human-
/// readable in a hex dump.
fn encodeInt8SymRows(out: *[4 * (4 + 4)]u8, scale: f32, q: [4]i8) void {
    var i: usize = 0;
    while (i < 4) : (i += 1) {
        const off = i * (4 + 4);
        std.mem.writeInt(u32, out[off..][0..4], @bitCast(scale), .little);
        // i8 is one byte; @bitCast to u8 to avoid sign-extension on copy.
        out[off + 4] = @bitCast(q[0]);
        out[off + 5] = @bitCast(q[1]);
        out[off + 6] = @bitCast(q[2]);
        out[off + 7] = @bitCast(q[3]);
    }
}
