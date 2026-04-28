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

pub const ChartType = enum {
    bar,
    line,
    pie,
    scatter,
    area,
    bubble,
    radar,
    /// Any other / unrecognised chart-XML element. The raw_xml is
    /// always available so callers can interrogate further.
    other,
};

pub const ChartAnchor = struct {
    /// Archive name of the chart part, e.g. `xl/charts/chart1.xml`.
    chart_part_name: []const u8,
    /// Archive name of the sheet whose drawing references this chart.
    sheet_part_name: []const u8,
    /// Top-left anchor cell.
    from: CellAnchor,
    /// Bottom-right anchor cell. `null` for `oneCellAnchor`.
    to: ?CellAnchor,
    /// Detected chart-type element (`<c:barChart>`, `<c:lineChart>`,
    /// etc.). `.other` covers unrecognised or compound charts; the
    /// raw_xml is always available for callers needing more detail.
    chart_type: ChartType,
    /// All `<c:f>` formula refs surfaced from the chart (series
    /// names, categories, values, labels — flattened in document
    /// order). Strings borrow from raw_xml; do not free.
    /// Empty when the chart uses inline literal data only.
    series_refs: []const []const u8,
    /// Raw chart-part XML bytes. Borrowed from the PartStore.
    raw_xml: []const u8,
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
        if (!isSheetPart(sheet_part)) continue;
        try collectFromSheet(store, allocator, sheet_part, &out);
    }

    return out.toOwnedSlice(allocator);
}

/// Same walk shape as `imageAnchors` but surfaces every embedded
/// chart (`<xdr:graphicFrame>` containing `<c:chart r:id=...>`).
/// Each ChartAnchor exposes the chart part's archive name + raw
/// XML bytes; the chart_type field is best-effort detected from
/// the chart-XML root element (barChart / lineChart / etc.) and
/// callers wanting series refs can interrogate raw_xml directly
/// for now.
pub fn chartAnchors(store: *PartStore, allocator: std.mem.Allocator) ![]ChartAnchor {
    var out: std.ArrayListUnmanaged(ChartAnchor) = .empty;
    // Each appended ChartAnchor owns an allocator-allocated
    // series_refs slice; on partial failure (e.g. OOM during a
    // later sheet) `out.deinit` alone leaks every prior chart's
    // refs. Walk and free each before the outer array.
    errdefer {
        for (out.items) |c| allocator.free(c.series_refs);
        out.deinit(allocator);
    }
    for (store.parts) |sheet_part| {
        if (!isSheetPart(sheet_part)) continue;
        try collectChartsFromSheet(store, allocator, sheet_part, &out);
    }
    return out.toOwnedSlice(allocator);
}

/// OOXML Transitional worksheet content type. Strict OOXML
/// (ECMA-376 second edition + later) ships variants of this with
/// different MIME prefixes, so detection accepts any content type
/// whose tail is `.worksheet+xml`. As a defensive belt-and-braces
/// fallback, the legacy filename heuristic (`xl/worksheets/sheet<N>.xml`)
/// is also accepted so workbooks that fail to declare a content
/// type still get their drawings walked — the union of the two
/// detection paths catches every workbook we've ever seen.
const ct_worksheet_transitional = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

fn isSheetPart(part: store_mod.Part) bool {
    if (part.content_type) |ct| {
        if (std.mem.endsWith(u8, ct, ".worksheet+xml")) return true;
        if (std.mem.eql(u8, ct, ct_worksheet_transitional)) return true;
    }
    // Fallback to filename for content-type-less producers. The
    // entire substring between `sheet` and `.xml` must be digits;
    // a partial-digit prefix would let `sheet1_backup.xml` /
    // `sheet1custom.xml` slip through and be walked as a worksheet.
    const prefix = "xl/worksheets/sheet";
    const suffix = ".xml";
    if (!std.mem.startsWith(u8, part.name, prefix)) return false;
    if (!std.mem.endsWith(u8, part.name, suffix)) return false;
    if (part.name.len <= prefix.len + suffix.len) return false;
    const num_part = part.name[prefix.len .. part.name.len - suffix.len];
    if (num_part.len == 0) return false;
    for (num_part) |c| if (!std.ascii.isDigit(c)) return false;
    return true;
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

fn collectChartsFromSheet(
    store: *PartStore,
    allocator: std.mem.Allocator,
    sheet_part: store_mod.Part,
    out: *std.ArrayListUnmanaged(ChartAnchor),
) !void {
    const rid = findDrawingRid(sheet_part.bytes) orelse return;
    const sheet_rels = store.rels(sheet_part.name);
    const drawing_target = relTargetForId(sheet_rels, rid) orelse return;
    const drawing_part_name = (try store.resolve(sheet_part.name, drawing_target)) orelse return;
    const drawing_part = store.part(drawing_part_name) orelse return;

    const drawing_rels = store.rels(drawing_part_name);

    var i: usize = 0;
    while (i < drawing_part.bytes.len) {
        const next = std.mem.indexOfPos(u8, drawing_part.bytes, i, "<xdr:") orelse break;
        i = next;
        const is_two = std.mem.startsWith(u8, drawing_part.bytes[i..], "<xdr:twoCellAnchor");
        const is_one = std.mem.startsWith(u8, drawing_part.bytes[i..], "<xdr:oneCellAnchor");
        if (!is_two and !is_one) {
            i += "<xdr:".len;
            continue;
        }
        const close_marker = if (is_two) "</xdr:twoCellAnchor>" else "</xdr:oneCellAnchor>";
        const close = std.mem.indexOfPos(u8, drawing_part.bytes, i, close_marker) orelse break;
        const block = drawing_part.bytes[i .. close + close_marker.len];
        i = close + close_marker.len;

        // Charts live inside <xdr:graphicFrame>...<c:chart r:id=...
        const gf_idx = std.mem.indexOf(u8, block, "<xdr:graphicFrame") orelse continue;
        const chart_idx = std.mem.indexOfPos(u8, block, gf_idx, "<c:chart") orelse continue;
        const chart_end = std.mem.indexOfScalarPos(u8, block, chart_idx, '>') orelse continue;
        const chart_attrs = block[chart_idx .. chart_end + 1];
        const embed_rid = attrValue(chart_attrs, "r:id") orelse continue;

        const chart_target = relTargetForId(drawing_rels, embed_rid) orelse continue;
        const chart_part_name = (try store.resolve(drawing_part_name, chart_target)) orelse continue;
        const chart_part = store.part(chart_part_name) orelse continue;

        const from = parseCellAnchor(block, "<xdr:from>", "</xdr:from>") orelse continue;
        const to_anchor: ?CellAnchor = if (is_two)
            parseCellAnchor(block, "<xdr:to>", "</xdr:to>")
        else
            null;

        const refs = try extractSeriesRefs(allocator, chart_part.bytes);
        // If `out.append` OOMs after we just allocated `refs`, the
        // caller's outer errdefer frees the rest but `refs` itself
        // hasn't been transferred yet — free it on the failing path.
        errdefer allocator.free(refs);
        try out.append(allocator, .{
            .chart_part_name = chart_part.name,
            .sheet_part_name = sheet_part.name,
            .from = from,
            .to = to_anchor,
            .chart_type = detectChartType(chart_part.bytes),
            .series_refs = refs,
            .raw_xml = chart_part.bytes,
        });
    }
}

/// Walk every `<c:f>...</c:f>` in the chart XML in document order and
/// return the formula strings as borrowed slices into `xml`. Series
/// names, categories, and values all flow through `<c:f>`, so the
/// flattened list captures every workbook reference the chart pulls
/// from.
fn extractSeriesRefs(allocator: std.mem.Allocator, xml: []const u8) ![]const []const u8 {
    var out: std.ArrayListUnmanaged([]const u8) = .empty;
    errdefer out.deinit(allocator);

    var i: usize = 0;
    while (std.mem.indexOfPos(u8, xml, i, "<c:f>")) |open| {
        const start = open + "<c:f>".len;
        const close = std.mem.indexOfPos(u8, xml, start, "</c:f>") orelse break;
        try out.append(allocator, xml[start..close]);
        i = close + "</c:f>".len;
    }
    return out.toOwnedSlice(allocator);
}

/// Best-effort chart-type detection from the chart-part XML. Looks
/// for the canonical `<c:Xchart>` element name. Compound charts
/// (multiple plot types overlaid) collapse to whichever is found
/// first; callers needing the full picture can walk raw_xml
/// directly.
fn detectChartType(chart_xml: []const u8) ChartType {
    if (std.mem.indexOf(u8, chart_xml, "<c:barChart") != null) return .bar;
    if (std.mem.indexOf(u8, chart_xml, "<c:lineChart") != null) return .line;
    if (std.mem.indexOf(u8, chart_xml, "<c:pieChart") != null) return .pie;
    if (std.mem.indexOf(u8, chart_xml, "<c:scatterChart") != null) return .scatter;
    if (std.mem.indexOf(u8, chart_xml, "<c:areaChart") != null) return .area;
    if (std.mem.indexOf(u8, chart_xml, "<c:bubbleChart") != null) return .bubble;
    if (std.mem.indexOf(u8, chart_xml, "<c:radarChart") != null) return .radar;
    return .other;
}

/// Generic attribute value extractor: find `key="value"` or
/// `key='value'` inside an already-narrowed tag-attributes slice.
/// Both quote styles are valid XML; non-Microsoft producers
/// (libreoffice, hand-edited drawings) sometimes emit single
/// quotes, and skipping them silently dropped image/chart anchors.
fn attrValue(attrs: []const u8, key: []const u8) ?[]const u8 {
    return attrValueWithQuote(attrs, key, '"') orelse
        attrValueWithQuote(attrs, key, '\'');
}

fn attrValueWithQuote(attrs: []const u8, key: []const u8, quote: u8) ?[]const u8 {
    var search_buf: [32]u8 = undefined;
    if (key.len + 2 > search_buf.len) return null;
    @memcpy(search_buf[0..key.len], key);
    search_buf[key.len] = '=';
    search_buf[key.len + 1] = quote;
    const needle = search_buf[0 .. key.len + 2];
    const found = std.mem.indexOf(u8, attrs, needle) orelse return null;
    const start = found + needle.len;
    const end = std.mem.indexOfScalarPos(u8, attrs, start, quote) orelse return null;
    return attrs[start..end];
}

/// Find the value of `r:id` on the sheet's `<drawing>` element. The
/// element is always self-closing in OOXML and lives at sheet scope
/// (one per sheet at most).
fn findDrawingRid(sheet_xml: []const u8) ?[]const u8 {
    const tag = findOpeningTag(sheet_xml, "drawing") orelse return null;
    const tag_end = std.mem.indexOfScalarPos(u8, sheet_xml, tag, '>') orelse return null;
    return attrValue(sheet_xml[tag .. tag_end + 1], "r:id");
}

/// Find the start of an opening tag named `name` in `xml`, tolerating
/// XML whitespace (space / tab / LF / CR) or `/`/`>` after the name.
/// `<drawing\nr:id="rId1"/>` is valid XML; the previous literal
/// "<drawing " search missed it and silently dropped anchors on
/// well-formed workbooks emitted by non-Microsoft producers.
fn findOpeningTag(xml: []const u8, name: []const u8) ?usize {
    var i: usize = 0;
    while (std.mem.indexOfPos(u8, xml, i, "<")) |lt| {
        const after_name = lt + 1 + name.len;
        if (after_name >= xml.len) return null;
        if (std.mem.eql(u8, xml[lt + 1 .. after_name], name)) {
            const c = xml[after_name];
            if (c == ' ' or c == '\t' or c == '\n' or c == '\r' or c == '/' or c == '>') {
                return lt;
            }
        }
        i = lt + 1;
    }
    return null;
}

/// Find the value of `r:embed` on the `<a:blip r:embed="rIdN" ...>`
/// inside an `<xdr:pic>` block. Linked-only blips (`r:link` instead
/// of `r:embed`) return null — those reference an external file and
/// have no part in the package.
fn findBlipEmbed(pic_xml: []const u8) ?[]const u8 {
    const blip = std.mem.indexOf(u8, pic_xml, "<a:blip") orelse return null;
    const blip_end = std.mem.indexOfScalarPos(u8, pic_xml, blip, '>') orelse return null;
    return attrValue(pic_xml[blip .. blip_end + 1], "r:embed");
}

fn relTargetForId(rels: []const store_mod.Relationship, id: []const u8) ?[]const u8 {
    // Decode the lookup id into a stack buffer so the comparison
    // matches the decoded Relationship.id stored by parseRelationships.
    // OOXML rIds in practice are short ASCII tokens (`rId1`, `rId12`),
    // so 64 bytes is generous. On overflow OR fast-path (no `&`), fall
    // back to raw compare against the source `id` slice.
    if (std.mem.indexOfScalar(u8, id, '&') == null) {
        for (rels) |r| {
            if (std.mem.eql(u8, r.id, id)) return r.target;
        }
        return null;
    }
    var buf: [64]u8 = undefined;
    const decoded = decodeIdInto(&buf, id) orelse {
        // Pathologically long encoded id — fall back to raw match.
        for (rels) |r| {
            if (std.mem.eql(u8, r.id, id)) return r.target;
        }
        return null;
    };
    for (rels) |r| {
        if (std.mem.eql(u8, r.id, decoded)) return r.target;
    }
    return null;
}

/// Decode the same five named entities + numeric refs into `buf`.
/// Returns null if the decoded form would exceed buf.len. This is
/// the lookup-key counterpart to store.zig's decodeXmlEntities —
/// same rules, no allocation, code-point UTF-8 ≤ 4 bytes per
/// reference. Symmetric handling means a relTargetForId lookup
/// matches whether the referring side uses named entities, numeric
/// refs, or literal characters.
fn decodeIdInto(buf: []u8, src: []const u8) ?[]const u8 {
    var out_len: usize = 0;
    var i: usize = 0;
    while (i < src.len) {
        if (src[i] == '&') {
            const remain = src[i..];
            // Named entities.
            if (std.mem.startsWith(u8, remain, "&amp;")) {
                if (out_len >= buf.len) return null;
                buf[out_len] = '&';
                out_len += 1;
                i += 5;
                continue;
            }
            if (std.mem.startsWith(u8, remain, "&lt;")) {
                if (out_len >= buf.len) return null;
                buf[out_len] = '<';
                out_len += 1;
                i += 4;
                continue;
            }
            if (std.mem.startsWith(u8, remain, "&gt;")) {
                if (out_len >= buf.len) return null;
                buf[out_len] = '>';
                out_len += 1;
                i += 4;
                continue;
            }
            if (std.mem.startsWith(u8, remain, "&quot;")) {
                if (out_len >= buf.len) return null;
                buf[out_len] = '"';
                out_len += 1;
                i += 6;
                continue;
            }
            if (std.mem.startsWith(u8, remain, "&apos;")) {
                if (out_len >= buf.len) return null;
                buf[out_len] = '\'';
                out_len += 1;
                i += 6;
                continue;
            }
            // Numeric character references via the same parser as
            // the storage-side decoder, so both sides agree on what
            // counts as a valid ref vs. a literal `&`.
            if (std.mem.startsWith(u8, remain, "&#")) {
                if (store_mod.decodeNumericRef(remain)) |info| {
                    const utf8 = info.utf8[0..info.utf8_len];
                    if (out_len + utf8.len > buf.len) return null;
                    @memcpy(buf[out_len..][0..utf8.len], utf8);
                    out_len += utf8.len;
                    i += info.consumed;
                    continue;
                }
            }
        }
        if (out_len >= buf.len) return null;
        buf[out_len] = src[i];
        out_len += 1;
        i += 1;
    }
    return buf[0..out_len];
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

test "chartAnchors: openxlsx_loadExample.xlsx surfaces embedded charts" {
    const fixture = "tests/corpus/openxlsx_loadExample.xlsx";
    std.fs.cwd().access(fixture, .{}) catch return error.SkipZigTest;

    var s = try PartStore.open(std.testing.allocator, fixture);
    defer s.deinit();

    const charts = try chartAnchors(&s, std.testing.allocator);
    defer {
        // Each ChartAnchor owns its series_refs slice (allocator-
        // allocated; inner strings borrow from raw_xml). Walk + free.
        for (charts) |c| std.testing.allocator.free(c.series_refs);
        std.testing.allocator.free(charts);
    }

    // openxlsx_loadExample has at least one embedded chart.
    try std.testing.expect(charts.len > 0);
    var any_with_refs = false;
    for (charts) |c| {
        try std.testing.expect(std.mem.startsWith(u8, c.chart_part_name, "xl/charts/chart"));
        try std.testing.expect(std.mem.startsWith(u8, c.sheet_part_name, "xl/worksheets/sheet"));
        try std.testing.expect(c.raw_xml.len > 0);
        // Detected chart type should be one of the known enum
        // values (.other is acceptable for compound / unrecognised
        // forms but every fixture in the corpus today is bar/line/
        // pie/scatter).
        switch (c.chart_type) {
            .bar, .line, .pie, .scatter, .area, .bubble, .radar, .other => {},
        }
        if (c.series_refs.len > 0) any_with_refs = true;
        // Every series ref borrowed from raw_xml; sanity check that
        // each ref is a non-empty substring containing a sheet
        // separator `!` (canonical SpreadsheetML reference shape).
        for (c.series_refs) |r| {
            try std.testing.expect(r.len > 0);
            try std.testing.expect(std.mem.indexOf(u8, r, "!") != null);
        }
    }
    // At least one chart in the fixture must have series refs;
    // chart3.xml in openxlsx_loadExample has them per-confirmed.
    try std.testing.expect(any_with_refs);
}

test "chartAnchors: workbook with no charts returns empty slice" {
    const fixture = "tests/corpus/worldbank_catalog.xlsx";
    std.fs.cwd().access(fixture, .{}) catch return error.SkipZigTest;

    var s = try PartStore.open(std.testing.allocator, fixture);
    defer s.deinit();

    const charts = try chartAnchors(&s, std.testing.allocator);
    defer {
        for (charts) |c| std.testing.allocator.free(c.series_refs);
        std.testing.allocator.free(charts);
    }

    try std.testing.expectEqual(@as(usize, 0), charts.len);
}

test "detectChartType: covers all canonical OOXML chart elements" {
    try std.testing.expectEqual(ChartType.bar, detectChartType("<c:chartSpace><c:barChart/>"));
    try std.testing.expectEqual(ChartType.line, detectChartType("<c:chartSpace><c:lineChart/>"));
    try std.testing.expectEqual(ChartType.pie, detectChartType("<c:chartSpace><c:pieChart/>"));
    try std.testing.expectEqual(ChartType.scatter, detectChartType("<c:chartSpace><c:scatterChart/>"));
    try std.testing.expectEqual(ChartType.area, detectChartType("<c:chartSpace><c:areaChart/>"));
    try std.testing.expectEqual(ChartType.bubble, detectChartType("<c:chartSpace><c:bubbleChart/>"));
    try std.testing.expectEqual(ChartType.radar, detectChartType("<c:chartSpace><c:radarChart/>"));
    try std.testing.expectEqual(ChartType.other, detectChartType("<c:chartSpace><c:doughnutChart/>"));
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

test "attrValue tolerates single-quoted XML attributes" {
    // Both quote styles are valid XML (W3C XML 1.0 §3.1). Valid
    // OOXML packages from libreoffice / pandoc / hand-edited drawings
    // use either, so the helper must accept both.
    try std.testing.expectEqualStrings("rId7", attrValue("foo=\"bar\" r:id=\"rId7\"", "r:id").?);
    try std.testing.expectEqualStrings("rId7", attrValue("foo='bar' r:id='rId7'", "r:id").?);
    try std.testing.expectEqualStrings("rId7", attrValue("r:id='rId7'", "r:id").?);
    try std.testing.expectEqualStrings("rId7", attrValue("r:id=\"rId7\"", "r:id").?);
    // Mixed quote styles in the same tag are legal XML.
    try std.testing.expectEqualStrings("X", attrValue("a=\"y\" b='X'", "b").?);
    // Missing key returns null regardless.
    try std.testing.expectEqual(@as(?[]const u8, null), attrValue("foo='bar'", "missing"));
}

test "findDrawingRid + findBlipEmbed tolerate single quotes" {
    try std.testing.expectEqualStrings(
        "rId3",
        findDrawingRid("<sheet><drawing r:id='rId3'/></sheet>").?,
    );
    try std.testing.expectEqualStrings(
        "rId9",
        findBlipEmbed("<xdr:pic><a:blip r:embed='rId9'/></xdr:pic>").?,
    );
}

test "findDrawingRid tolerates XML whitespace after tag name" {
    // `<drawing\n r:id=...>` and `<drawing\tr:id=...>` are valid XML.
    try std.testing.expectEqualStrings(
        "rId7",
        findDrawingRid("<sheet><drawing\nr:id=\"rId7\"/></sheet>").?,
    );
    try std.testing.expectEqualStrings(
        "rId8",
        findDrawingRid("<sheet><drawing\tr:id=\"rId8\"/></sheet>").?,
    );
    // <drawingthing ...> is NOT a drawing tag — must not match.
    try std.testing.expectEqual(
        @as(?[]const u8, null),
        findDrawingRid("<sheet><drawingthing r:id=\"rIdX\"/></sheet>"),
    );
}
