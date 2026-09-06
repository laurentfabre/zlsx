//! The `embed --extract` NDJSON records — S3c's embeddable-rows read
//! (`docs/cli.md`, "embed --extract"), written once for every surface.
//!
//! `zlsx embed --extract` streams these; the C and Python legs of row
//! S3c (`zlsx_editor_embeddable_rows_ndjson`, `Editor.embeddable_rows`)
//! hand over the same bytes. One writer, so the record a Python caller
//! sees is byte-for-byte the one the CLI prints — the
//! `defined_name_ndjson.zig` precedent.
//!
//! The record is `Workbook.embeddableRows`' row: the 1-based sheet
//! row, the text a model should see, and the canonical xxh3-64
//! content hash the write stores beside the vector — `hash` as an
//! unsigned 64-bit decimal, the value `zlsx_emb_hashes` reads back.
//! Nothing here judges a row: the workbook read already refused a
//! text that is not UTF-8, so the JSON escaper's pass-through of every
//! byte above the C0 range is a pass-through of valid UTF-8.

const std = @import("std");
const json = @import("json_text.zig");
const workbook_mod = @import("workbook.zig");

pub const EmbeddableRow = workbook_mod.EmbeddableRow;

/// One `{"kind":"embed_row",…}` line. The field order is the
/// docs/cli.md contract; a change here is a wire-format change on
/// every surface at once.
pub fn writeRow(out: *std.Io.Writer, r: EmbeddableRow) !void {
    try out.print("{{\"kind\":\"embed_row\",\"row\":{d},\"text\":", .{r.row});
    try json.writeString(out, r.text);
    try out.print(",\"hash\":{d}}}\n", .{r.hash});
}

/// The whole stream — every row, range order. The CLI's and the C
/// leg's entry point (`c_abi.zig::embeddableRowsNdjsonOwned`).
pub fn writeAll(out: *std.Io.Writer, rows: []const EmbeddableRow) !void {
    for (rows) |r| try writeRow(out, r);
}

// ─── Tests ───────────────────────────────────────────────────────────

test "writeAll: the wire shape — row, text escaped by the shared escaper, hash as an unsigned decimal" {
    var out: std.Io.Writer.Allocating = .init(std.testing.allocator);
    defer out.deinit();
    try writeAll(&out.writer, &.{
        .{ .row = 2, .text = "alpha", .hash = 0x1111 },
        // The escaper's own rules ride along: the two metacharacters
        // and the tab are escaped, `& < >` and UTF-8 pass through.
        .{ .row = 1048576, .text = "a \"quoted\" tab\t& < > é", .hash = std.math.maxInt(u64) },
        .{ .row = 3, .text = "0", .hash = 0 },
    });
    try std.testing.expectEqualStrings(
        "{\"kind\":\"embed_row\",\"row\":2,\"text\":\"alpha\",\"hash\":4369}\n" ++
            "{\"kind\":\"embed_row\",\"row\":1048576,\"text\":\"a \\\"quoted\\\" tab\\t& < > é\",\"hash\":18446744073709551615}\n" ++
            "{\"kind\":\"embed_row\",\"row\":3,\"text\":\"0\",\"hash\":0}\n",
        out.written(),
    );
}

test "writeAll: no rows, no bytes" {
    var out: std.Io.Writer.Allocating = .init(std.testing.allocator);
    defer out.deinit();
    try writeAll(&out.writer, &.{});
    try std.testing.expectEqual(@as(usize, 0), out.written().len);
}
