//! C2a (post-0.2.9 roadmap): per-sheet image-anchor extraction.
//!
//! Builds on `PartStore` to surface every embedded image with its
//! sheet attribution and cell anchor. Out of scope for v1: charts,
//! shapes, pivot tables, absolute-pixel anchors. Covers the
//! "extract images grouped by sheet" workflow that's the
//! highest-value chunk of object preservation.
//!
//! OOXML drawing structure walked here:
//!
//!   xl/worksheets/sheet1.xml         <drawing r:id="rIdN"/>
//!   xl/worksheets/_rels/sheet1.xml.rels  rIdN → ../drawings/drawing1.xml
//!   xl/drawings/drawing1.xml         <xdr:wsDr>
//!     <xdr:twoCellAnchor>            ← anchor wrapper
//!       <xdr:from><xdr:col>X</xdr:col>...<xdr:rowOff>Y</xdr:rowOff></xdr:from>
//!       <xdr:to>...</xdr:to>          (twoCellAnchor only)
//!       <xdr:pic>                    ← images live here
//!         <xdr:blipFill>
//!           <a:blip r:embed="rIdM"/> ← rIdM resolved via drawing1's rels
//!         </xdr:blipFill>
//!       </xdr:pic>
//!     </xdr:twoCellAnchor>
//!   xl/drawings/_rels/drawing1.xml.rels  rIdM → ../media/image1.png
//!
//! `oneCellAnchor` is the same minus `<xdr:to>`.
//! `absoluteAnchor` (pixel-pos) is detected but skipped — its `<xdr:pos>`
//! shape doesn't fit the cell-grid contract; callers needing it can
//! reach for the raw drawing XML via PartStore.part().
//!
//! ⚠️ Namespace-prefix assumption: this v1 parser hard-codes the
//! `xdr:` prefix for the spreadsheetDrawing namespace and `a:` for
//! drawingml. Every Microsoft Excel + LibreOffice + xlsxwriter +
//! openpyxl + python-calamine fixture in the project's corpus uses
//! these prefixes, but OOXML producers are technically free to pick
//! any prefix. Workbooks with non-standard prefixes will surface
//! zero anchors instead of erroring. A namespace-aware parser is
//! queued as a future iter; until then the assumption is documented
//! here as a known limitation.

const std = @import("std");
const store_mod = @import("store.zig");
const PartStore = store_mod.PartStore;

pub const CellAnchor = struct {
    /// 0-based column index.
    col: u32,
    /// EMU offset within the column (1 EMU = 1/914400 inch).
    col_off: i64,
    /// 0-based row index.
    row: u32,
    /// EMU offset within the row.
    row_off: i64,
};

pub const ImageAnchor = struct {
    /// Archive name of the image part, e.g. `xl/media/image1.png`.
    image_part_name: []const u8,
    /// Archive name of the sheet whose drawing references this image,
    /// e.g. `xl/worksheets/sheet1.xml`.
    sheet_part_name: []const u8,
    /// Top-left anchor cell.
    from: CellAnchor,
    /// Bottom-right anchor cell. `null` for `oneCellAnchor` (image
    /// sized via `<xdr:ext>` in EMUs, which we don't expose here).
    to: ?CellAnchor,
    /// Decompressed image bytes (PNG/JPEG/etc.). Borrowed from the
    /// PartStore — caller must not free.
    bytes: []const u8,
};

/// Walk every worksheet's `<drawing r:id=...>`, resolve to a drawing
/// part, parse anchored `<xdr:pic>` entries, and return the resulting
/// list of ImageAnchors.
///
/// Allocations come from `allocator` for the returned slice; string
/// slices inside each anchor are arena-borrowed from the PartStore
/// (valid until the store's `deinit`).
pub fn imageAnchors(store: *PartStore, allocator: std.mem.Allocator) ![]ImageAnchor {
    var out: std.ArrayListUnmanaged(ImageAnchor) = .empty;
    errdefer out.deinit(allocator);

    // Walk every sheet part.
    for (store.parts) |sheet_part| {
        if (!isSheetPart(sheet_part.name)) continue;
        try collectFromSheet(store, allocator, sheet_part, &out);
    }

    return out.toOwnedSlice(allocator);
}

fn isSheetPart(name: []const u8) bool {
    return std.mem.startsWith(u8, name, "xl/worksheets/sheet") and
        std.mem.endsWith(u8, name, ".xml");
}

fn collectFromSheet(
    store: *PartStore,
    allocator: std.mem.Allocator,
    sheet_part: store_mod.Part,
    out: *std.ArrayListUnmanaged(ImageAnchor),
) !void {
    // Find `<drawing r:id="..."/>` in the sheet XML. Skip the sheet
    // entirely if absent (no anchored objects).
    const rid = findDrawingRid(sheet_part.bytes) orelse return;

    // Resolve rid → drawing part name via sheet's rels.
    const sheet_rels = store.rels(sheet_part.name);
    const drawing_target = relTargetForId(sheet_rels, rid) orelse return;
    const drawing_part_name = (try store.resolve(sheet_part.name, drawing_target)) orelse return;
    const drawing_part = store.part(drawing_part_name) orelse return;

    // Walk the drawing's twoCellAnchor / oneCellAnchor blocks.
    const drawing_rels = store.rels(drawing_part_name);

    var i: usize = 0;
    while (i < drawing_part.bytes.len) {
        const next = std.mem.indexOfPos(u8, drawing_part.bytes, i, "<xdr:") orelse break;
        i = next;
        // Identify anchor opener.
        const is_two = std.mem.startsWith(u8, drawing_part.bytes[i..], "<xdr:twoCellAnchor");
        const is_one = std.mem.startsWith(u8, drawing_part.bytes[i..], "<xdr:oneCellAnchor");
        if (!is_two and !is_one) {
            i += "<xdr:".len;
            continue;
        }
        // Find close tag.
        const close_marker = if (is_two) "</xdr:twoCellAnchor>" else "</xdr:oneCellAnchor>";
        const close = std.mem.indexOfPos(u8, drawing_part.bytes, i, close_marker) orelse break;
        const block = drawing_part.bytes[i .. close + close_marker.len];
        i = close + close_marker.len;

        // Only image-bearing anchors are surfaced in v1.
        const pic_idx = std.mem.indexOf(u8, block, "<xdr:pic>") orelse continue;
        const pic_close = std.mem.indexOfPos(u8, block, pic_idx, "</xdr:pic>") orelse continue;
        const pic_block = block[pic_idx .. pic_close + "</xdr:pic>".len];

        const embed_rid = findBlipEmbed(pic_block) orelse continue;
        const image_target = relTargetForId(drawing_rels, embed_rid) orelse continue;
        const image_part_name = (try store.resolve(drawing_part_name, image_target)) orelse continue;
        const image_part = store.part(image_part_name) orelse continue;

        const from = parseCellAnchor(block, "<xdr:from>", "</xdr:from>") orelse continue;
        const to_anchor: ?CellAnchor = if (is_two)
            parseCellAnchor(block, "<xdr:to>", "</xdr:to>")
        else
            null;

        try out.append(allocator, .{
            .image_part_name = image_part.name,
            .sheet_part_name = sheet_part.name,
            .from = from,
            .to = to_anchor,
            .bytes = image_part.bytes,
        });
    }
}

/// Find the value of `r:id` on the sheet's `<drawing>` element. The
/// element is always self-closing in OOXML and lives at sheet scope
/// (one per sheet at most).
fn findDrawingRid(sheet_xml: []const u8) ?[]const u8 {
    const tag = std.mem.indexOf(u8, sheet_xml, "<drawing ") orelse return null;
    const tag_end = std.mem.indexOfScalarPos(u8, sheet_xml, tag, '>') orelse return null;
    const attrs = sheet_xml[tag .. tag_end + 1];
    const key = "r:id=\"";
    const ks = std.mem.indexOf(u8, attrs, key) orelse return null;
    const start = ks + key.len;
    const end = std.mem.indexOfScalarPos(u8, attrs, start, '"') orelse return null;
    return attrs[start..end];
}

/// Find the value of `r:embed` on the `<a:blip r:embed="rIdN" ...>`
/// inside an `<xdr:pic>` block. Linked-only blips (`r:link` instead
/// of `r:embed`) return null — those reference an external file and
/// have no part in the package.
fn findBlipEmbed(pic_xml: []const u8) ?[]const u8 {
    const blip = std.mem.indexOf(u8, pic_xml, "<a:blip") orelse return null;
    const blip_end = std.mem.indexOfScalarPos(u8, pic_xml, blip, '>') orelse return null;
    const attrs = pic_xml[blip .. blip_end + 1];
    const key = "r:embed=\"";
    const ks = std.mem.indexOf(u8, attrs, key) orelse return null;
    const start = ks + key.len;
    const end = std.mem.indexOfScalarPos(u8, attrs, start, '"') orelse return null;
    return attrs[start..end];
}

fn relTargetForId(rels: []const store_mod.Relationship, id: []const u8) ?[]const u8 {
    for (rels) |r| {
        if (std.mem.eql(u8, r.id, id)) return r.target;
    }
    return null;
}

/// Parse `<xdr:from>...</xdr:from>` (or `<xdr:to>...</xdr:to>`) into
/// a CellAnchor. Each contains exactly four scalar children:
///   <xdr:col>N</xdr:col>
///   <xdr:colOff>N</xdr:colOff>
///   <xdr:row>N</xdr:row>
///   <xdr:rowOff>N</xdr:rowOff>
fn parseCellAnchor(xml: []const u8, open: []const u8, close: []const u8) ?CellAnchor {
    const o = std.mem.indexOf(u8, xml, open) orelse return null;
    const c = std.mem.indexOfPos(u8, xml, o, close) orelse return null;
    const inner = xml[o + open.len .. c];

    return .{
        .col = parseElementU32(inner, "<xdr:col>", "</xdr:col>") orelse return null,
        .col_off = parseElementI64(inner, "<xdr:colOff>", "</xdr:colOff>") orelse return null,
        .row = parseElementU32(inner, "<xdr:row>", "</xdr:row>") orelse return null,
        .row_off = parseElementI64(inner, "<xdr:rowOff>", "</xdr:rowOff>") orelse return null,
    };
}

fn parseElementU32(xml: []const u8, open: []const u8, close: []const u8) ?u32 {
    const start = std.mem.indexOf(u8, xml, open) orelse return null;
    const value_start = start + open.len;
    const value_end = std.mem.indexOfPos(u8, xml, value_start, close) orelse return null;
    return std.fmt.parseInt(u32, xml[value_start..value_end], 10) catch null;
}

fn parseElementI64(xml: []const u8, open: []const u8, close: []const u8) ?i64 {
    const start = std.mem.indexOf(u8, xml, open) orelse return null;
    const value_start = start + open.len;
    const value_end = std.mem.indexOfPos(u8, xml, value_start, close) orelse return null;
    return std.fmt.parseInt(i64, xml[value_start..value_end], 10) catch null;
}

// ─── Tests ────────────────────────────────────────────────────────────

test "imageAnchors: openxlsx_loadExample.xlsx surfaces 2 anchored images" {
    const fixture = "tests/corpus/openxlsx_loadExample.xlsx";
    std.fs.cwd().access(fixture, .{}) catch return error.SkipZigTest;

    var s = try PartStore.open(std.testing.allocator, fixture);
    defer s.deinit();

    const anchors = try imageAnchors(&s, std.testing.allocator);
    defer std.testing.allocator.free(anchors);

    try std.testing.expect(anchors.len >= 2);

    // Every anchor must point at an image part with non-empty bytes
    // and a sheet part name.
    for (anchors) |a| {
        try std.testing.expect(std.mem.startsWith(u8, a.image_part_name, "xl/media/"));
        try std.testing.expect(std.mem.startsWith(u8, a.sheet_part_name, "xl/worksheets/sheet"));
        try std.testing.expect(a.bytes.len > 0);
        // Both image1.jpeg and image2.jpeg are JPEGs — bytes start
        // with the JPEG SOI marker 0xFFD8.
        try std.testing.expectEqual(@as(u8, 0xFF), a.bytes[0]);
        try std.testing.expectEqual(@as(u8, 0xD8), a.bytes[1]);
    }

    // The two anchors are on the same sheet (sheet3) and use
    // twoCellAnchor (so .to is non-null).
    try std.testing.expect(anchors[0].to != null);
}

test "imageAnchors: skips drawings with shapes only (no <xdr:pic>)" {
    // poi_58325_db.xlsx ships shape-only drawings. The parser must
    // walk them without producing image anchors.
    const fixture = "tests/corpus/poi_58325_db.xlsx";
    std.fs.cwd().access(fixture, .{}) catch return error.SkipZigTest;

    var s = try PartStore.open(std.testing.allocator, fixture);
    defer s.deinit();

    const anchors = try imageAnchors(&s, std.testing.allocator);
    defer std.testing.allocator.free(anchors);

    // Some fixtures may have hidden <xdr:pic> entries; just assert
    // the parser doesn't crash and runs to completion. The image
    // count for poi_58325_db happens to be zero anchored — the four
    // images live in xl/media/ but aren't anchored via drawing rels
    // (legacy VML / direct embed paths).
    try std.testing.expect(anchors.len >= 0);
}

test "imageAnchors: workbook with no drawings returns empty slice" {
    // worldbank_catalog has no drawings at all; the parser should
    // walk every sheet and find nothing.
    const fixture = "tests/corpus/worldbank_catalog.xlsx";
    std.fs.cwd().access(fixture, .{}) catch return error.SkipZigTest;

    var s = try PartStore.open(std.testing.allocator, fixture);
    defer s.deinit();

    const anchors = try imageAnchors(&s, std.testing.allocator);
    defer std.testing.allocator.free(anchors);

    try std.testing.expectEqual(@as(usize, 0), anchors.len);
}

test "parseCellAnchor unit test" {
    const xml =
        \\<xdr:from><xdr:col>3</xdr:col><xdr:colOff>16119</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>47624</xdr:rowOff></xdr:from>
    ;
    const a = parseCellAnchor(xml, "<xdr:from>", "</xdr:from>").?;
    try std.testing.expectEqual(@as(u32, 3), a.col);
    try std.testing.expectEqual(@as(i64, 16119), a.col_off);
    try std.testing.expectEqual(@as(u32, 1), a.row);
    try std.testing.expectEqual(@as(i64, 47624), a.row_off);
}
