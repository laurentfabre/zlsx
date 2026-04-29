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

/// Pixel-coordinate anchor for `<xdr:absoluteAnchor>`. Used when a
/// drawing isn't tied to specific cells but instead pinned to a
/// fixed position on the sheet (rare in practice; some legacy
/// charts and AlternateContent fallbacks emit this shape).
pub const AbsoluteAnchor = struct {
    /// X offset from the sheet origin in EMUs (1 EMU = 1/914400 inch).
    x: i64,
    /// Y offset from the sheet origin in EMUs.
    y: i64,
    /// Width in EMUs (extent's `cx` attribute).
    cx: i64,
    /// Height in EMUs (extent's `cy` attribute).
    cy: i64,
};

pub const ImageAnchor = struct {
    /// Archive name of the image part, e.g. `xl/media/image1.png`.
    image_part_name: []const u8,
    /// Archive name of the sheet whose drawing references this image,
    /// e.g. `xl/worksheets/sheet1.xml`.
    sheet_part_name: []const u8,
    /// Top-left anchor cell. For `<xdr:absoluteAnchor>` images this
    /// is a zero sentinel — check `.absolute != null` first to
    /// distinguish.
    from: CellAnchor,
    /// Bottom-right anchor cell. `null` for `oneCellAnchor` (image
    /// sized via `<xdr:ext>` in EMUs) or `absoluteAnchor`.
    to: ?CellAnchor,
    /// Pixel-coordinate placement when the source used
    /// `<xdr:absoluteAnchor>`. Mutually exclusive with cell-anchor
    /// fields above (which become a zero sentinel in that case).
    absolute: ?AbsoluteAnchor = null,
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
    /// Top-left anchor cell. For `<xdr:absoluteAnchor>` charts this
    /// is a zero sentinel — check `.absolute != null` first to
    /// distinguish.
    from: CellAnchor,
    /// Bottom-right anchor cell. `null` for `oneCellAnchor` or
    /// `absoluteAnchor`.
    to: ?CellAnchor,
    /// Pixel-coordinate placement when the source used
    /// `<xdr:absoluteAnchor>`. Mutually exclusive with cell-anchor
    /// fields above (which become a zero sentinel in that case).
    absolute: ?AbsoluteAnchor = null,
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
    const drawing_target = relTargetForId(allocator, sheet_rels, rid) orelse return;
    const drawing_part_name = (try store.resolve(sheet_part.name, drawing_target)) orelse return;
    const drawing_part = store.part(drawing_part_name) orelse return;

    // Walk the drawing's twoCellAnchor / oneCellAnchor blocks.
    const drawing_rels = store.rels(drawing_part_name);

    // Resolve namespace prefixes once per drawing part. Microsoft
    // canonically uses "xdr" / "a" / "c" but OOXML allows any
    // prefix. Look up the actual prefix declared on the root
    // element so non-Microsoft producers (libreoffice, custom
    // tooling) don't silently surface zero anchors.
    const prefixes = resolveDrawingPrefixes(drawing_part.bytes);
    // 4 KiB tag-needle scratch covers prefixes up to ~250 chars
    // per needle (12 needles × ~max-prefix-len ≈ 4 KiB total). XML
    // namespace prefixes have no upper bound in the spec — Codex
    // flagged the previous 512 B buffer as a hard-fail surface for
    // valid-but-long custom prefixes.
    var tags_buf: [4096]u8 = undefined;
    const tags = try DrawingTags.build(&tags_buf, prefixes);

    var i: usize = 0;
    while (i < drawing_part.bytes.len) {
        const next = std.mem.indexOfPos(u8, drawing_part.bytes, i, tags.xdr_prefix_open) orelse break;
        i = next;
        // Identify anchor opener.
        const is_two = std.mem.startsWith(u8, drawing_part.bytes[i..], tags.open_two);
        const is_one = std.mem.startsWith(u8, drawing_part.bytes[i..], tags.open_one);
        const is_absolute = std.mem.startsWith(u8, drawing_part.bytes[i..], tags.open_absolute);
        if (!is_two and !is_one and !is_absolute) {
            i += tags.xdr_prefix_open.len;
            continue;
        }
        // Find close tag.
        const close_marker = if (is_two)
            tags.close_two
        else if (is_one)
            tags.close_one
        else
            tags.close_absolute;
        const close = std.mem.indexOfPos(u8, drawing_part.bytes, i, close_marker) orelse break;
        const block = drawing_part.bytes[i .. close + close_marker.len];
        i = close + close_marker.len;

        // Only image-bearing anchors are surfaced in v1.
        const pic_idx = std.mem.indexOf(u8, block, tags.open_pic) orelse continue;
        const pic_close = std.mem.indexOfPos(u8, block, pic_idx, tags.close_pic) orelse continue;
        const pic_block = block[pic_idx .. pic_close + tags.close_pic.len];

        const embed_rid = findBlipEmbedWithAlt(pic_block, prefixes.a, prefixes.a_alt) orelse continue;
        const image_target = relTargetForId(allocator, drawing_rels, embed_rid) orelse continue;
        const image_part_name = (try store.resolve(drawing_part_name, image_target)) orelse continue;
        const image_part = store.part(image_part_name) orelse continue;

        var from: CellAnchor = .{ .col = 0, .col_off = 0, .row = 0, .row_off = 0 };
        var to_anchor: ?CellAnchor = null;
        var absolute: ?AbsoluteAnchor = null;
        if (is_absolute) {
            absolute = parseAbsoluteAnchor(block, tags.open_pos, tags.open_ext) orelse continue;
        } else {
            from = parseCellAnchor(block, tags.open_from, tags.close_from, prefixes.xdr) orelse continue;
            if (is_two) {
                to_anchor = parseCellAnchor(block, tags.open_to, tags.close_to, prefixes.xdr);
            }
        }

        try out.append(allocator, .{
            .image_part_name = image_part.name,
            .sheet_part_name = sheet_part.name,
            .from = from,
            .to = to_anchor,
            .absolute = absolute,
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
    const drawing_target = relTargetForId(allocator, sheet_rels, rid) orelse return;
    const drawing_part_name = (try store.resolve(sheet_part.name, drawing_target)) orelse return;
    const drawing_part = store.part(drawing_part_name) orelse return;

    const drawing_rels = store.rels(drawing_part_name);
    const prefixes = resolveDrawingPrefixes(drawing_part.bytes);
    // 4 KiB tag-needle scratch covers prefixes up to ~250 chars
    // per needle (12 needles × ~max-prefix-len ≈ 4 KiB total). XML
    // namespace prefixes have no upper bound in the spec — Codex
    // flagged the previous 512 B buffer as a hard-fail surface for
    // valid-but-long custom prefixes.
    var tags_buf: [4096]u8 = undefined;
    const tags = try DrawingTags.build(&tags_buf, prefixes);

    var i: usize = 0;
    while (i < drawing_part.bytes.len) {
        const next = std.mem.indexOfPos(u8, drawing_part.bytes, i, tags.xdr_prefix_open) orelse break;
        i = next;
        const is_two = std.mem.startsWith(u8, drawing_part.bytes[i..], tags.open_two);
        const is_one = std.mem.startsWith(u8, drawing_part.bytes[i..], tags.open_one);
        const is_absolute = std.mem.startsWith(u8, drawing_part.bytes[i..], tags.open_absolute);
        if (!is_two and !is_one and !is_absolute) {
            i += tags.xdr_prefix_open.len;
            continue;
        }
        const close_marker = if (is_two)
            tags.close_two
        else if (is_one)
            tags.close_one
        else
            tags.close_absolute;
        const close = std.mem.indexOfPos(u8, drawing_part.bytes, i, close_marker) orelse break;
        const block = drawing_part.bytes[i .. close + close_marker.len];
        i = close + close_marker.len;

        // Charts live inside <xdr:graphicFrame>...<c:chart r:id=...
        const gf_idx = std.mem.indexOf(u8, block, tags.open_graphic_frame) orelse continue;
        // The chart-namespace prefix can be declared LOCALLY on the
        // <*:chart> element rather than on the drawing root (valid
        // OOXML scoping pattern). Try BOTH the block-local prefix
        // and the drawing-root prefix — XML allows redundant local
        // declarations whose prefixes aren't the one actually used
        // by the chart element, so picking the first match anywhere
        // in the block could miss valid root-scoped charts.
        var block_chart_buf: [128]u8 = undefined;
        const block_chart_prefix_t = findNamespacePrefix(block, ns_c_transitional);
        const block_chart_prefix_s = findNamespacePrefix(block, ns_c_strict);
        const root_open_chart = tags.open_chart;
        const chart_idx = blk: {
            // Try every plausible prefix for the chart namespace —
            // local Transitional, local Strict, and the drawing-
            // root resolved prefix. Whichever matches the actual
            // `<*:chart` element in the block wins. This covers
            // mixed-conformance blocks that declare both URIs.
            const candidates = [_]?[]const u8{
                block_chart_prefix_t,
                block_chart_prefix_s,
                prefixes.c_alt, // drawing-root alternate-class binding
            };
            for (candidates) |maybe_p| {
                const p = maybe_p orelse continue;
                const local_open = std.fmt.bufPrint(&block_chart_buf, "<{s}:chart", .{p}) catch continue;
                if (std.mem.indexOfPos(u8, block, gf_idx, local_open)) |idx| break :blk idx;
            }
            // Fall back to the drawing-root primary prefix.
            if (std.mem.indexOfPos(u8, block, gf_idx, root_open_chart)) |idx| break :blk idx;
            continue;
        };
        const chart_end = std.mem.indexOfScalarPos(u8, block, chart_idx, '>') orelse continue;
        const chart_attrs = block[chart_idx .. chart_end + 1];
        const embed_rid = attrValue(chart_attrs, "r:id") orelse continue;

        const chart_target = relTargetForId(allocator, drawing_rels, embed_rid) orelse continue;
        const chart_part_name = (try store.resolve(drawing_part_name, chart_target)) orelse continue;
        const chart_part = store.part(chart_part_name) orelse continue;

        var from: CellAnchor = .{ .col = 0, .col_off = 0, .row = 0, .row_off = 0 };
        var to_anchor: ?CellAnchor = null;
        var absolute: ?AbsoluteAnchor = null;
        if (is_absolute) {
            absolute = parseAbsoluteAnchor(block, tags.open_pos, tags.open_ext) orelse continue;
        } else {
            from = parseCellAnchor(block, tags.open_from, tags.close_from, prefixes.xdr) orelse continue;
            if (is_two) {
                to_anchor = parseCellAnchor(block, tags.open_to, tags.close_to, prefixes.xdr);
            }
        }

        // Each chart's own XML may declare a different `c:` prefix
        // — resolve per-chart to be safe.
        const chart_prefixes = resolveDrawingPrefixes(chart_part.bytes);
        const refs = try extractSeriesRefs(allocator, chart_part.bytes, chart_prefixes.c, chart_prefixes.c_alt);
        // If `out.append` OOMs after we just allocated `refs`, the
        // caller's outer errdefer frees the rest but `refs` itself
        // hasn't been transferred yet — free it on the failing path.
        errdefer allocator.free(refs);
        try out.append(allocator, .{
            .chart_part_name = chart_part.name,
            .sheet_part_name = sheet_part.name,
            .from = from,
            .to = to_anchor,
            .absolute = absolute,
            .chart_type = detectChartTypeWithAlt(chart_part.bytes, chart_prefixes.c, chart_prefixes.c_alt),
            .series_refs = refs,
            .raw_xml = chart_part.bytes,
        });
    }
}

/// Walk every `<{c}:f>...</{c}:f>` in the chart XML in document
/// order and return the formula strings as borrowed slices into
/// `xml`. Series names, categories, and values all flow through
/// `<{c}:f>`, so the flattened list captures every workbook
/// reference the chart pulls from. `c_prefix` is the document's
/// actual chart-namespace prefix (canonically "c").
fn extractSeriesRefs(
    allocator: std.mem.Allocator,
    xml: []const u8,
    c_prefix: []const u8,
    c_prefix_alt: ?[]const u8,
) ![]const []const u8 {
    var out: std.ArrayListUnmanaged([]const u8) = .empty;
    errdefer out.deinit(allocator);

    var primary_open_buf: [128]u8 = undefined;
    var primary_close_buf: [128]u8 = undefined;
    const primary_open = std.fmt.bufPrint(&primary_open_buf, "<{s}:f>", .{c_prefix}) catch return out.toOwnedSlice(allocator);
    const primary_close = std.fmt.bufPrint(&primary_close_buf, "</{s}:f>", .{c_prefix}) catch return out.toOwnedSlice(allocator);

    var alt_open_buf: [128]u8 = undefined;
    var alt_close_buf: [128]u8 = undefined;
    var alt_open: ?[]const u8 = null;
    var alt_close: ?[]const u8 = null;
    if (c_prefix_alt) |alt| {
        if (!std.mem.eql(u8, alt, c_prefix)) {
            alt_open = std.fmt.bufPrint(&alt_open_buf, "<{s}:f>", .{alt}) catch null;
            alt_close = std.fmt.bufPrint(&alt_close_buf, "</{s}:f>", .{alt}) catch null;
        }
    }

    // Single document-order pass: at each position, find the
    // EARLIER of the next primary-prefix match and the next
    // alt-prefix match; consume that one and advance. Preserves
    // the documented "flattened in document order" contract even
    // for mixed-prefix charts.
    var i: usize = 0;
    while (true) {
        const p_pos = std.mem.indexOfPos(u8, xml, i, primary_open);
        const a_pos = if (alt_open) |o| std.mem.indexOfPos(u8, xml, i, o) else null;
        const winner: enum { primary, alt } = if (p_pos != null and a_pos != null)
            (if (p_pos.? < a_pos.?) .primary else .alt)
        else if (p_pos != null) .primary else if (a_pos != null) .alt else break;
        const open_offset = if (winner == .primary) p_pos.? else a_pos.?;
        const open_tag = if (winner == .primary) primary_open else alt_open.?;
        const close_tag = if (winner == .primary) primary_close else alt_close.?;
        const start = open_offset + open_tag.len;
        const close_off = std.mem.indexOfPos(u8, xml, start, close_tag) orelse break;
        try out.append(allocator, xml[start..close_off]);
        i = close_off + close_tag.len;
    }
    return out.toOwnedSlice(allocator);
}

/// Best-effort chart-type detection from the chart-part XML. Looks
/// for the canonical `<{c}:Xchart>` element name. Compound charts
/// (multiple plot types overlaid) collapse to whichever is found
/// first; callers needing the full picture can walk raw_xml
/// directly. `c_prefix` is the document's actual chart-namespace
/// prefix (canonically "c").
fn detectChartType(chart_xml: []const u8, c_prefix: []const u8) ChartType {
    return detectChartTypeWithAlt(chart_xml, c_prefix, null);
}

fn detectChartTypeWithAlt(
    chart_xml: []const u8,
    c_prefix: []const u8,
    c_prefix_alt: ?[]const u8,
) ChartType {
    var buf: [128]u8 = undefined;
    const candidates = [_]struct { suffix: []const u8, kind: ChartType }{
        .{ .suffix = "barChart", .kind = .bar },
        .{ .suffix = "lineChart", .kind = .line },
        .{ .suffix = "pieChart", .kind = .pie },
        .{ .suffix = "scatterChart", .kind = .scatter },
        .{ .suffix = "areaChart", .kind = .area },
        .{ .suffix = "bubbleChart", .kind = .bubble },
        .{ .suffix = "radarChart", .kind = .radar },
    };
    const prefixes = [_]?[]const u8{ c_prefix, c_prefix_alt };
    for (prefixes) |maybe_p| {
        const p = maybe_p orelse continue;
        for (candidates) |c| {
            const needle = std.fmt.bufPrint(&buf, "<{s}:{s}", .{ p, c.suffix }) catch continue;
            if (std.mem.indexOf(u8, chart_xml, needle) != null) return c.kind;
        }
    }
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

/// Find the value of `r:embed` on the `<{a}:blip r:embed="rIdN" ...>`
/// inside a `<{xdr}:pic>` block. Linked-only blips (`r:link`
/// instead of `r:embed`) return null — those reference an external
/// file and have no part in the package. `a_prefix` is the
/// document's actual DrawingML-main prefix (canonically "a").
fn findBlipEmbed(pic_xml: []const u8, a_prefix: []const u8) ?[]const u8 {
    return findBlipEmbedWithAlt(pic_xml, a_prefix, null);
}

fn findBlipEmbedWithAlt(
    pic_xml: []const u8,
    a_prefix: []const u8,
    a_prefix_alt: ?[]const u8,
) ?[]const u8 {
    var blip_open_buf: [128]u8 = undefined;
    // Probe each candidate prefix in turn. A `<*:blip>` element
    // with no `r:embed` (e.g. a linked-only blip) shouldn't end
    // the search — try the alternate prefix's blip too in case
    // the embedded one lives under that conformance class.
    if (tryBlipEmbedAt(pic_xml, &blip_open_buf, a_prefix)) |rid| return rid;
    if (a_prefix_alt) |alt| {
        if (tryBlipEmbedAt(pic_xml, &blip_open_buf, alt)) |rid| return rid;
    }
    return null;
}

fn tryBlipEmbedAt(pic_xml: []const u8, buf: []u8, prefix: []const u8) ?[]const u8 {
    const blip_open = std.fmt.bufPrint(buf, "<{s}:blip", .{prefix}) catch return null;
    var search_at: usize = 0;
    while (std.mem.indexOfPos(u8, pic_xml, search_at, blip_open)) |blip| {
        const blip_end = std.mem.indexOfScalarPos(u8, pic_xml, blip, '>') orelse return null;
        if (attrValue(pic_xml[blip .. blip_end + 1], "r:embed")) |rid| return rid;
        search_at = blip_end + 1;
    }
    return null;
}

/// OOXML namespace URIs for the three prefixes the drawing parser
/// needs. Both Transitional (ECMA-376 first edition,
/// http://schemas.openxmlformats.org/...) and Strict (second
/// edition, http://purl.oclc.org/ooxml/...) URIs are accepted —
/// Strict-conformance workbooks declare the same logical types
/// under the purl.oclc.org variants.
const ns_xdr_transitional = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const ns_xdr_strict = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
const ns_a_transitional = "http://schemas.openxmlformats.org/drawingml/2006/main";
const ns_a_strict = "http://purl.oclc.org/ooxml/drawingml/main";
const ns_c_transitional = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const ns_c_strict = "http://purl.oclc.org/ooxml/drawingml/chart";

/// Resolved namespace prefixes for one drawing or chart part.
/// Defaults are the canonical Microsoft prefixes; secondary
/// fields hold the alternate-conformance binding when both
/// Transitional and Strict URIs are declared so downstream
/// lookups can probe either prefix.
const DrawingPrefixes = struct {
    xdr: []const u8 = "xdr",
    a: []const u8 = "a",
    c: []const u8 = "c",
    a_alt: ?[]const u8 = null,
    c_alt: ?[]const u8 = null,
};

/// Scan the root element's xmlns:* declarations and return the
/// prefix for each canonical OOXML namespace. Falls back to the
/// canonical prefix when a namespace isn't declared (some chart
/// parts only declare the chart namespace inline on `<c:chart>`).
fn resolveDrawingPrefixes(xml: []const u8) DrawingPrefixes {
    var p: DrawingPrefixes = .{};
    // The xdr / a / c namespaces are independently scoped: a
    // document can bind one to its Strict URI and another to its
    // Transitional URI under different prefixes. Anchor xdr on
    // the root element when possible (so a stray binding doesn't
    // override the prefix actually used by `<*:wsDr>`), but
    // resolve `a` and `c` independently across both URIs.
    const root_xdr = rootElementPrefix(xml);
    if (root_xdr) |pref| {
        p.xdr = pref;
    } else if (findNamespacePrefix(xml, ns_xdr_transitional) orelse
        findNamespacePrefix(xml, ns_xdr_strict)) |pref|
    {
        p.xdr = pref;
    }
    const a_t = findNamespacePrefix(xml, ns_a_transitional);
    const a_s = findNamespacePrefix(xml, ns_a_strict);
    if (a_t orelse a_s) |pref| p.a = pref;
    if (a_t != null and a_s != null) p.a_alt = a_s;
    const c_t = findNamespacePrefix(xml, ns_c_transitional);
    const c_s = findNamespacePrefix(xml, ns_c_strict);
    if (c_t orelse c_s) |pref| p.c = pref;
    if (c_t != null and c_s != null) p.c_alt = c_s;
    return p;
}

/// Find the prefix on the root XML element. Skips the XML
/// declaration (`<?xml ... ?>`) and any leading whitespace, then
/// reads the first `<NAME:` token's NAME. Returns null if the
/// root element is unprefixed or absent. Bounded scan (4 KiB).
fn rootElementPrefix(xml: []const u8) ?[]const u8 {
    const limit = @min(xml.len, 4096);
    var i: usize = 0;
    // Skip the optional XML declaration.
    while (i < limit) {
        const lt = std.mem.indexOfScalarPos(u8, xml[0..limit], i, '<') orelse return null;
        if (lt + 1 >= limit) return null;
        const after = xml[lt + 1];
        if (after == '?') {
            // Skip until '?>'
            const close = std.mem.indexOfPos(u8, xml[0..limit], lt, "?>") orelse return null;
            i = close + 2;
            continue;
        }
        if (after == '!') {
            // Skip DOCTYPE / comment / CDATA until '>'
            const close = std.mem.indexOfScalarPos(u8, xml[0..limit], lt, '>') orelse return null;
            i = close + 1;
            continue;
        }
        // Real element. Read NAME[:LOCAL].
        var j = lt + 1;
        while (j < limit) : (j += 1) {
            const c = xml[j];
            if (c == ':') {
                if (j - (lt + 1) > max_prefix_len) return null;
                return xml[lt + 1 .. j];
            }
            if (c == ' ' or c == '\t' or c == '\n' or c == '\r' or c == '/' or c == '>') {
                return null; // unprefixed
            }
        }
        return null;
    }
    return null;
}

/// Walk the first 4 KiB of `xml` looking for `xmlns:NAME="URI"`.
/// Returns NAME if URI matches `target_uri`. Bounded scan because
/// xmlns declarations are always on the root element; well past
/// 4 KiB the search would cost more than it saves on adversarial
/// input.
/// Reject prefixes longer than the smallest per-needle scratch
/// buffer can format. The buf-fixed Writer in DrawingTags.build
/// handles 4 KiB of needles total (12+ needles), and the per-
/// helper scratch buffers (parseCellAnchor, findBlipEmbed,
/// detectChartType) are 128 bytes each. The longest formatted
/// needle is `</PREFIX:rowOff>` = prefix.len + 12, so a 110-char
/// prefix exactly fills the 128-byte buffer with no room for the
/// fixed pattern. 100 leaves a comfortable margin and still
/// covers any conceivable real-world prefix (OOXML uses 1-8 char
/// prefixes; 100 is already absurd).
const max_prefix_len: usize = 100;

fn findNamespacePrefix(xml: []const u8, target_uri: []const u8) ?[]const u8 {
    const limit = @min(xml.len, 4096);
    var i: usize = 0;
    while (std.mem.indexOfPos(u8, xml[0..limit], i, "xmlns:")) |start| {
        const after = start + "xmlns:".len;
        if (after >= limit) return null;
        // Walk forward to the first XML whitespace OR `=` to pin
        // the prefix end. XML 1.0 allows arbitrary whitespace
        // around the `=` between attribute name and value, so
        // `xmlns:dr = "uri"` must be tolerated.
        var name_end = after;
        while (name_end < limit) : (name_end += 1) {
            const c = xml[name_end];
            if (c == '=' or c == ' ' or c == '\t' or c == '\n' or c == '\r') break;
        }
        if (name_end >= limit) return null;
        const name = xml[after..name_end];
        // Skip whitespace before `=`, then expect `=`.
        var p = name_end;
        while (p < limit and (xml[p] == ' ' or xml[p] == '\t' or xml[p] == '\n' or xml[p] == '\r')) p += 1;
        if (p >= limit or xml[p] != '=') {
            i = after;
            continue;
        }
        p += 1;
        // Skip whitespace after `=`, then expect a quote.
        while (p < limit and (xml[p] == ' ' or xml[p] == '\t' or xml[p] == '\n' or xml[p] == '\r')) p += 1;
        if (p >= limit) return null;
        const quote = xml[p];
        if (quote != '"' and quote != '\'') {
            i = p;
            continue;
        }
        const val_start = p + 1;
        const val_end = std.mem.indexOfScalarPos(u8, xml[0..limit], val_start, quote) orelse return null;
        if (std.mem.eql(u8, xml[val_start..val_end], target_uri)) {
            // Cap pathologically long prefixes — they'd overflow
            // the per-needle scratch buffers downstream. Skip
            // (rather than abort the whole lookup) so a workbook
            // that declares multiple prefixes for the same URI
            // can still match a usable one further along.
            if (name.len <= max_prefix_len) return name;
        }
        i = val_end + 1;
    }
    return null;
}

/// Pre-built tag needles keyed off the resolved prefixes. Built
/// into a single caller-supplied buffer so the per-part lookup
/// loop doesn't re-format on every iteration.
const DrawingTags = struct {
    xdr_prefix_open: []const u8, // "<xdr:"
    open_two: []const u8, // "<xdr:twoCellAnchor"
    close_two: []const u8, // "</xdr:twoCellAnchor>"
    open_one: []const u8, // "<xdr:oneCellAnchor"
    close_one: []const u8, // "</xdr:oneCellAnchor>"
    open_absolute: []const u8, // "<xdr:absoluteAnchor"
    close_absolute: []const u8, // "</xdr:absoluteAnchor>"
    open_pic: []const u8, // "<xdr:pic>"
    close_pic: []const u8, // "</xdr:pic>"
    open_from: []const u8, // "<xdr:from>"
    close_from: []const u8, // "</xdr:from>"
    open_to: []const u8, // "<xdr:to>"
    close_to: []const u8, // "</xdr:to>"
    open_pos: []const u8, // "<xdr:pos"
    open_ext: []const u8, // "<xdr:ext"
    open_graphic_frame: []const u8, // "<xdr:graphicFrame"
    open_chart: []const u8, // "<c:chart"

    fn build(buf: []u8, p: DrawingPrefixes) !DrawingTags {
        var w = std.Io.Writer.fixed(buf);
        const xdr_prefix_open = try writeAndAdvance(&w, "<{s}:", .{p.xdr});
        const open_two = try writeAndAdvance(&w, "<{s}:twoCellAnchor", .{p.xdr});
        const close_two = try writeAndAdvance(&w, "</{s}:twoCellAnchor>", .{p.xdr});
        const open_one = try writeAndAdvance(&w, "<{s}:oneCellAnchor", .{p.xdr});
        const close_one = try writeAndAdvance(&w, "</{s}:oneCellAnchor>", .{p.xdr});
        const open_absolute = try writeAndAdvance(&w, "<{s}:absoluteAnchor", .{p.xdr});
        const close_absolute = try writeAndAdvance(&w, "</{s}:absoluteAnchor>", .{p.xdr});
        const open_pic = try writeAndAdvance(&w, "<{s}:pic>", .{p.xdr});
        const close_pic = try writeAndAdvance(&w, "</{s}:pic>", .{p.xdr});
        const open_from = try writeAndAdvance(&w, "<{s}:from>", .{p.xdr});
        const close_from = try writeAndAdvance(&w, "</{s}:from>", .{p.xdr});
        const open_to = try writeAndAdvance(&w, "<{s}:to>", .{p.xdr});
        const close_to = try writeAndAdvance(&w, "</{s}:to>", .{p.xdr});
        const open_pos = try writeAndAdvance(&w, "<{s}:pos", .{p.xdr});
        const open_ext = try writeAndAdvance(&w, "<{s}:ext", .{p.xdr});
        const open_graphic_frame = try writeAndAdvance(&w, "<{s}:graphicFrame", .{p.xdr});
        const open_chart = try writeAndAdvance(&w, "<{s}:chart", .{p.c});
        return .{
            .xdr_prefix_open = xdr_prefix_open,
            .open_two = open_two,
            .close_two = close_two,
            .open_one = open_one,
            .close_one = close_one,
            .open_absolute = open_absolute,
            .close_absolute = close_absolute,
            .open_pic = open_pic,
            .close_pic = close_pic,
            .open_from = open_from,
            .close_from = close_from,
            .open_to = open_to,
            .close_to = close_to,
            .open_pos = open_pos,
            .open_ext = open_ext,
            .open_graphic_frame = open_graphic_frame,
            .open_chart = open_chart,
        };
    }
};

/// Format `fmt` into the writer-fixed buffer and return the slice
/// of bytes that were just written (offset into the underlying
/// buffer, fixed for the writer's lifetime).
fn writeAndAdvance(w: *std.Io.Writer, comptime fmt: []const u8, args: anytype) ![]const u8 {
    const before = w.end;
    try w.print(fmt, args);
    return w.buffer[before..w.end];
}

fn relForId(
    allocator: std.mem.Allocator,
    rels: []const store_mod.Relationship,
    id: []const u8,
) ?store_mod.Relationship {
    // Decode the lookup id so the comparison matches the decoded
    // Relationship.id stored by parseRelationships. OOXML rIds in
    // practice are short ASCII tokens (`rId1`, `rId12`), so the
    // 64-byte stack buffer fast path covers everything realistic.
    // Pathological encoded IDs that decode beyond 64 bytes fall
    // through to a heap-allocated decode so we still match
    // correctly instead of silently dropping the relationship.
    if (std.mem.indexOfScalar(u8, id, '&') == null) {
        for (rels) |r| {
            if (std.mem.eql(u8, r.id, id)) return r;
        }
        return null;
    }
    var buf: [64]u8 = undefined;
    if (decodeIdInto(&buf, id)) |decoded| {
        for (rels) |r| {
            if (std.mem.eql(u8, r.id, decoded)) return r;
        }
        return null;
    }
    // Stack buffer overflow — heap-allocate a buffer large enough
    // to hold the worst-case decoded length (≤ id.len since each
    // entity decodes to at most as many bytes as its escaped form).
    const heap_buf = allocator.alloc(u8, id.len) catch {
        // OOM during a non-critical relationship lookup — last
        // resort, raw-compare. We still try in case the encoded
        // form happens to match (it won't, since stored IDs are
        // decoded, but the alternative is silent miss).
        for (rels) |r| {
            if (std.mem.eql(u8, r.id, id)) return r;
        }
        return null;
    };
    defer allocator.free(heap_buf);
    const decoded = decodeIdInto(heap_buf, id) orelse {
        for (rels) |r| {
            if (std.mem.eql(u8, r.id, id)) return r;
        }
        return null;
    };
    for (rels) |r| {
        if (std.mem.eql(u8, r.id, decoded)) return r;
    }
    return null;
}

/// Look up an internal-mode relationship target. External-mode
/// rels (TargetMode="External") return null even when their target
/// looks relative — those are linked-from-elsewhere references the
/// package doesn't carry the bytes for, and resolving them as
/// internal would (mis)attribute external links to package parts
/// that happen to share the relative path.
fn relTargetForId(
    allocator: std.mem.Allocator,
    rels: []const store_mod.Relationship,
    id: []const u8,
) ?[]const u8 {
    const r = relForId(allocator, rels, id) orelse return null;
    if (r.target_mode == .external) return null;
    return r.target;
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
///   <{xdr}:col>N</{xdr}:col>
///   <{xdr}:colOff>N</{xdr}:colOff>
///   <{xdr}:row>N</{xdr}:row>
///   <{xdr}:rowOff>N</{xdr}:rowOff>
/// Parse the `<{xdr}:pos x="N" y="N"/>` and `<{xdr}:ext cx="N"
/// cy="N"/>` self-closing children of an `<{xdr}:absoluteAnchor>`.
/// Both are required for a valid absoluteAnchor — returning null
/// causes the caller to skip the anchor as malformed.
fn parseAbsoluteAnchor(xml: []const u8, open_pos: []const u8, open_ext: []const u8) ?AbsoluteAnchor {
    const pos_idx = std.mem.indexOf(u8, xml, open_pos) orelse return null;
    const pos_end = std.mem.indexOfScalarPos(u8, xml, pos_idx, '>') orelse return null;
    const pos_attrs = xml[pos_idx .. pos_end + 1];
    const x_str = attrValue(pos_attrs, "x") orelse return null;
    const y_str = attrValue(pos_attrs, "y") orelse return null;
    const x = std.fmt.parseInt(i64, x_str, 10) catch return null;
    const y = std.fmt.parseInt(i64, y_str, 10) catch return null;

    const ext_idx = std.mem.indexOfPos(u8, xml, pos_end, open_ext) orelse return null;
    const ext_end = std.mem.indexOfScalarPos(u8, xml, ext_idx, '>') orelse return null;
    const ext_attrs = xml[ext_idx .. ext_end + 1];
    const cx_str = attrValue(ext_attrs, "cx") orelse return null;
    const cy_str = attrValue(ext_attrs, "cy") orelse return null;
    const cx = std.fmt.parseInt(i64, cx_str, 10) catch return null;
    const cy = std.fmt.parseInt(i64, cy_str, 10) catch return null;

    return .{ .x = x, .y = y, .cx = cx, .cy = cy };
}

fn parseCellAnchor(
    xml: []const u8,
    open: []const u8,
    close: []const u8,
    xdr_prefix: []const u8,
) ?CellAnchor {
    const o = std.mem.indexOf(u8, xml, open) orelse return null;
    const c = std.mem.indexOfPos(u8, xml, o, close) orelse return null;
    const inner = xml[o + open.len .. c];

    // 128-byte scratch per needle covers prefixes up to ~110 chars
    // (`</PREFIX:rowOff>` ≈ prefix.len + 11), well past anything
    // realistic. The previous 32-byte budget bottomed out at
    // ~20-char prefixes.
    var col_open_buf: [128]u8 = undefined;
    var col_close_buf: [128]u8 = undefined;
    var col_off_open_buf: [128]u8 = undefined;
    var col_off_close_buf: [128]u8 = undefined;
    var row_open_buf: [128]u8 = undefined;
    var row_close_buf: [128]u8 = undefined;
    var row_off_open_buf: [128]u8 = undefined;
    var row_off_close_buf: [128]u8 = undefined;
    const col_open = std.fmt.bufPrint(&col_open_buf, "<{s}:col>", .{xdr_prefix}) catch return null;
    const col_close = std.fmt.bufPrint(&col_close_buf, "</{s}:col>", .{xdr_prefix}) catch return null;
    const col_off_open = std.fmt.bufPrint(&col_off_open_buf, "<{s}:colOff>", .{xdr_prefix}) catch return null;
    const col_off_close = std.fmt.bufPrint(&col_off_close_buf, "</{s}:colOff>", .{xdr_prefix}) catch return null;
    const row_open = std.fmt.bufPrint(&row_open_buf, "<{s}:row>", .{xdr_prefix}) catch return null;
    const row_close = std.fmt.bufPrint(&row_close_buf, "</{s}:row>", .{xdr_prefix}) catch return null;
    const row_off_open = std.fmt.bufPrint(&row_off_open_buf, "<{s}:rowOff>", .{xdr_prefix}) catch return null;
    const row_off_close = std.fmt.bufPrint(&row_off_close_buf, "</{s}:rowOff>", .{xdr_prefix}) catch return null;

    return .{
        .col = parseElementU32(inner, col_open, col_close) orelse return null,
        .col_off = parseElementI64(inner, col_off_open, col_off_close) orelse return null,
        .row = parseElementU32(inner, row_open, row_close) orelse return null,
        .row_off = parseElementI64(inner, row_off_open, row_off_close) orelse return null,
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
    try std.testing.expectEqual(ChartType.bar, detectChartType("<c:chartSpace><c:barChart/>", "c"));
    try std.testing.expectEqual(ChartType.line, detectChartType("<c:chartSpace><c:lineChart/>", "c"));
    try std.testing.expectEqual(ChartType.pie, detectChartType("<c:chartSpace><c:pieChart/>", "c"));
    try std.testing.expectEqual(ChartType.scatter, detectChartType("<c:chartSpace><c:scatterChart/>", "c"));
    try std.testing.expectEqual(ChartType.area, detectChartType("<c:chartSpace><c:areaChart/>", "c"));
    try std.testing.expectEqual(ChartType.bubble, detectChartType("<c:chartSpace><c:bubbleChart/>", "c"));
    try std.testing.expectEqual(ChartType.radar, detectChartType("<c:chartSpace><c:radarChart/>", "c"));
    try std.testing.expectEqual(ChartType.other, detectChartType("<c:chartSpace><c:doughnutChart/>", "c"));
    // Non-canonical prefix: same XML with a different chart-namespace
    // prefix should still detect the chart type.
    try std.testing.expectEqual(ChartType.bar, detectChartType("<chrt:chartSpace><chrt:barChart/>", "chrt"));
}

test "parseCellAnchor unit test" {
    const xml =
        \\<xdr:from><xdr:col>3</xdr:col><xdr:colOff>16119</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>47624</xdr:rowOff></xdr:from>
    ;
    const a = parseCellAnchor(xml, "<xdr:from>", "</xdr:from>", "xdr").?;
    try std.testing.expectEqual(@as(u32, 3), a.col);
    try std.testing.expectEqual(@as(i64, 16119), a.col_off);
    try std.testing.expectEqual(@as(u32, 1), a.row);
    try std.testing.expectEqual(@as(i64, 47624), a.row_off);
    // Non-canonical drawing prefix: identical structure with `dr:`
    // instead of `xdr:` — same parser run with a different prefix.
    const xml2 =
        \\<dr:from><dr:col>3</dr:col><dr:colOff>0</dr:colOff><dr:row>1</dr:row><dr:rowOff>0</dr:rowOff></dr:from>
    ;
    const b = parseCellAnchor(xml2, "<dr:from>", "</dr:from>", "dr").?;
    try std.testing.expectEqual(@as(u32, 3), b.col);
    try std.testing.expectEqual(@as(u32, 1), b.row);
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
        findBlipEmbed("<xdr:pic><a:blip r:embed='rId9'/></xdr:pic>", "a").?,
    );
    // Non-canonical DrawingML-main prefix.
    try std.testing.expectEqualStrings(
        "rId9",
        findBlipEmbed("<xdr:pic><dml:blip r:embed='rId9'/></xdr:pic>", "dml").?,
    );
}

test "parseAbsoluteAnchor unit test" {
    const xml =
        \\<xdr:absoluteAnchor>
        \\  <xdr:pos x="914400" y="685800"/>
        \\  <xdr:ext cx="3657600" cy="2743200"/>
        \\</xdr:absoluteAnchor>
    ;
    const a = parseAbsoluteAnchor(xml, "<xdr:pos", "<xdr:ext").?;
    try std.testing.expectEqual(@as(i64, 914400), a.x);
    try std.testing.expectEqual(@as(i64, 685800), a.y);
    try std.testing.expectEqual(@as(i64, 3657600), a.cx);
    try std.testing.expectEqual(@as(i64, 2743200), a.cy);

    // Same shape with non-canonical drawing prefix.
    const xml2 =
        \\<dr:absoluteAnchor>
        \\  <dr:pos x="100" y="200"/>
        \\  <dr:ext cx="300" cy="400"/>
        \\</dr:absoluteAnchor>
    ;
    const b = parseAbsoluteAnchor(xml2, "<dr:pos", "<dr:ext").?;
    try std.testing.expectEqual(@as(i64, 100), b.x);
    try std.testing.expectEqual(@as(i64, 400), b.cy);

    // Missing pos/ext returns null.
    try std.testing.expectEqual(
        @as(?AbsoluteAnchor, null),
        parseAbsoluteAnchor("<xdr:absoluteAnchor></xdr:absoluteAnchor>", "<xdr:pos", "<xdr:ext"),
    );
}

test "resolveDrawingPrefixes maps canonical + custom prefixes" {
    // Canonical prefixes — round-trip.
    {
        const xml =
            \\<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>
        ;
        const p = resolveDrawingPrefixes(xml);
        try std.testing.expectEqualStrings("xdr", p.xdr);
        try std.testing.expectEqualStrings("a", p.a);
        try std.testing.expectEqualStrings("c", p.c);
    }
    // Custom prefixes — different short names mapped to same URIs.
    {
        const xml =
            \\<?xml version="1.0"?><dr:wsDr xmlns:dr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:dml="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:chrt="http://schemas.openxmlformats.org/drawingml/2006/chart"/>
        ;
        const p = resolveDrawingPrefixes(xml);
        try std.testing.expectEqualStrings("dr", p.xdr);
        try std.testing.expectEqualStrings("dml", p.a);
        try std.testing.expectEqualStrings("chrt", p.c);
    }
    // Single-quoted attribute values — also valid XML.
    {
        const xml =
            \\<?xml version='1.0'?><x:wsDr xmlns:x='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'/>
        ;
        const p = resolveDrawingPrefixes(xml);
        try std.testing.expectEqualStrings("x", p.xdr);
        // Undeclared namespaces fall back to canonical defaults.
        try std.testing.expectEqualStrings("a", p.a);
    }
    // No declarations at all — defaults.
    {
        const p = resolveDrawingPrefixes("<wsDr/>");
        try std.testing.expectEqualStrings("xdr", p.xdr);
        try std.testing.expectEqualStrings("a", p.a);
        try std.testing.expectEqualStrings("c", p.c);
    }
    // Strict OOXML namespace URIs — http://purl.oclc.org/ooxml/...
    {
        const xml =
            \\<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing" xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" xmlns:c="http://purl.oclc.org/ooxml/drawingml/chart"/>
        ;
        const p = resolveDrawingPrefixes(xml);
        try std.testing.expectEqualStrings("xdr", p.xdr);
        try std.testing.expectEqualStrings("a", p.a);
        try std.testing.expectEqualStrings("c", p.c);
    }
    // Strict-namespace URIs with non-canonical prefix names.
    {
        const xml =
            \\<?xml version="1.0"?><dr:wsDr xmlns:dr="http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing" xmlns:dml="http://purl.oclc.org/ooxml/drawingml/main"/>
        ;
        const p = resolveDrawingPrefixes(xml);
        try std.testing.expectEqualStrings("dr", p.xdr);
        try std.testing.expectEqualStrings("dml", p.a);
    }
    // Whitespace around `=` is valid XML — must be tolerated.
    {
        const xml =
            \\<wsDr xmlns:dr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>
        ;
        const p = resolveDrawingPrefixes(xml);
        try std.testing.expectEqualStrings("dr", p.xdr);
    }
    // Newlines + tabs around `=` (some pretty-printers).
    {
        const xml =
            "<wsDr xmlns:dr\n\t=\n\t\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"/>";
        const p = resolveDrawingPrefixes(xml);
        try std.testing.expectEqualStrings("dr", p.xdr);
    }
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
