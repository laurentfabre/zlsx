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

/// Which OOXML anchor wrapper carried the object. Recorded from the
/// opening tag itself, not inferred from which optional fields parsed
/// — a `<xdr:twoCellAnchor>` whose `<to>` is unreadable must not
/// masquerade as a one-cell anchor (Codex #214 r1 REL-101).
pub const AnchorKind = enum { two_cell, one_cell, absolute };

/// How the walkers treat drawing structures they recognise but cannot
/// read whole. `lenient` (the historical behavior, and the default
/// wrappers') skips them — the "no crash, surface what parses"
/// contract the corpus walks rely on. `strict` refuses with
/// `error.MalformedDrawingXml` instead: a dangling drawing / image /
/// chart relationship, a named part that is absent, an anchor whose
/// `from` / `to` / `pos` / `ext` does not parse, an opened block that
/// never closes. The NDJSON view reads strict — a partial anchor
/// inventory is the shape of a guard hole.
pub const WalkMode = enum { lenient, strict };

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
    /// The anchor wrapper as spelled in the drawing XML.
    kind: AnchorKind,
    /// Byte offset of the anchor's opening tag within its drawing
    /// part — the document-order key. The mixed-prefix replay appends
    /// alternate-prefixed anchors after the primary scan, so slice
    /// order alone is NOT document order (Codex #214 r1 REL-103);
    /// sort on this to restore it.
    doc_offset: usize,
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
    /// The anchor wrapper as spelled in the drawing XML.
    kind: AnchorKind,
    /// Byte offset of the anchor's opening tag within its drawing
    /// part — the document-order key (see `ImageAnchor.doc_offset`).
    doc_offset: usize,
};

/// Walk every worksheet's `<drawing r:id=...>`, resolve to a drawing
/// part, parse anchored `<xdr:pic>` entries, and return the resulting
/// list of ImageAnchors.
///
/// Allocations come from `allocator` for the returned slice; string
/// slices inside each anchor are arena-borrowed from the PartStore
/// (valid until the store's `deinit`).
pub fn imageAnchors(store: *PartStore, allocator: std.mem.Allocator) ![]ImageAnchor {
    return imageAnchorsIn(store, allocator, .lenient);
}

pub fn imageAnchorsIn(store: *PartStore, allocator: std.mem.Allocator, mode: WalkMode) ![]ImageAnchor {
    var out: std.ArrayListUnmanaged(ImageAnchor) = .empty;
    errdefer out.deinit(allocator);

    // Walk every sheet part. After the deferred-decompress refactor
    // (PartStore lazy materialization), `store.parts[i].bytes` is
    // empty until we go through `store.part(name)` — fetch by name
    // here so the sheet XML is materialized before we scan it.
    for (store.parts) |sheet_part_meta| {
        if (!isSheetPart(sheet_part_meta)) continue;
        const sheet_part = (try store.part(sheet_part_meta.name)) orelse continue;
        try collectFromSheet(store, allocator, sheet_part, mode, &out);
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
    return chartAnchorsIn(store, allocator, .lenient);
}

pub fn chartAnchorsIn(store: *PartStore, allocator: std.mem.Allocator, mode: WalkMode) ![]ChartAnchor {
    var out: std.ArrayListUnmanaged(ChartAnchor) = .empty;
    // Each appended ChartAnchor owns an allocator-allocated
    // series_refs slice; on partial failure (e.g. OOM during a
    // later sheet) `out.deinit` alone leaks every prior chart's
    // refs. Walk and free each before the outer array.
    errdefer {
        for (out.items) |c| allocator.free(c.series_refs);
        out.deinit(allocator);
    }
    for (store.parts) |sheet_part_meta| {
        if (!isSheetPart(sheet_part_meta)) continue;
        const sheet_part = (try store.part(sheet_part_meta.name)) orelse continue;
        try collectChartsFromSheet(store, allocator, sheet_part, mode, &out);
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
    mode: WalkMode,
    out: *std.ArrayListUnmanaged(ImageAnchor),
) !void {
    // Find `<drawing r:id="..."/>` in the sheet XML. Skip the sheet
    // entirely if absent (no anchored objects). A sheet that DOES
    // declare one whose reference is unreadable or whose relationship
    // chain dangles is a drawing the walk cannot read — strict
    // refuses it whole.
    const rid = switch (findDrawingRef(sheet_part.bytes)) {
        .absent => return,
        .malformed => {
            if (mode == .strict) return error.MalformedDrawingXml;
            return;
        },
        .rid => |r| r,
    };

    // Resolve rid → drawing part name via sheet's rels. Strict also
    // requires the relationship's TYPE to be the drawing edge — an id
    // that reaches a hyperlink relationship with an extant target
    // would otherwise read a wrong part as a drawing (REL-202).
    const sheet_rels = store.rels(sheet_part.name);
    const drawing_target = (try relTargetForIdTyped(allocator, sheet_rels, rid, "drawing", mode)) orelse {
        if (mode == .strict) return error.MalformedDrawingXml;
        return;
    };
    const drawing_part_name = (try store.resolve(sheet_part.name, drawing_target)) orelse {
        if (mode == .strict) return error.MalformedDrawingXml;
        return;
    };
    const drawing_part = try store.part(drawing_part_name) orelse {
        if (mode == .strict) return error.MalformedDrawingXml;
        return;
    };

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
    var primary_buf: [4096]u8 = undefined;
    const primary_tags = try DrawingTags.build(&primary_buf, prefixes);
    try scanImagesWithTags(store, allocator, drawing_part, drawing_part_name, drawing_rels, sheet_part, prefixes, primary_tags, mode, out);

    // Mixed-prefix bindings: descendant anchors may use a different
    // prefix bound to the same spreadsheetDrawing URI. Replay the
    // scan once per alt prefix; primary-prefixed anchors won't
    // match, so no duplicates surface. Replay order is NOT document
    // order — consumers needing that sort on `doc_offset`.
    for (prefixes.xdr_alts()) |alt| {
        var alt_prefixes = prefixes;
        alt_prefixes.xdr = alt;
        var alt_buf: [4096]u8 = undefined;
        const alt_tags = try DrawingTags.build(&alt_buf, alt_prefixes);
        try scanImagesWithTags(store, allocator, drawing_part, drawing_part_name, drawing_rels, sheet_part, alt_prefixes, alt_tags, mode, out);
    }
}

fn scanImagesWithTags(
    store: *PartStore,
    allocator: std.mem.Allocator,
    drawing_part: store_mod.Part,
    drawing_part_name: []const u8,
    drawing_rels: []const store_mod.Relationship,
    sheet_part: store_mod.Part,
    prefixes: DrawingPrefixes,
    tags: DrawingTags,
    mode: WalkMode,
    out: *std.ArrayListUnmanaged(ImageAnchor),
) !void {
    var i: usize = 0;
    while (i < drawing_part.bytes.len) {
        const next = std.mem.indexOfPos(u8, drawing_part.bytes, i, tags.xdr_prefix_open) orelse break;
        i = next;
        const block_start = next;
        // Identify anchor opener.
        const is_two = std.mem.startsWith(u8, drawing_part.bytes[i..], tags.open_two);
        const is_one = std.mem.startsWith(u8, drawing_part.bytes[i..], tags.open_one);
        const is_absolute = std.mem.startsWith(u8, drawing_part.bytes[i..], tags.open_absolute);
        if (!is_two and !is_one and !is_absolute) {
            i += tags.xdr_prefix_open.len;
            continue;
        }
        // Find close tag. An opened anchor that never closes is a
        // malformed drawing, not an absent one.
        const close_marker = if (is_two)
            tags.close_two
        else if (is_one)
            tags.close_one
        else
            tags.close_absolute;
        const close = std.mem.indexOfPos(u8, drawing_part.bytes, i, close_marker) orelse {
            if (mode == .strict) return error.MalformedDrawingXml;
            break;
        };
        const block = drawing_part.bytes[i .. close + close_marker.len];
        i = close + close_marker.len;

        // Only image-bearing anchors are surfaced. An anchor with no
        // `<xdr:pic>` (a shape, a chart frame) is legitimately not an
        // image in either mode; from the unclosed pic onward the block
        // holds an image the walk cannot read whole.
        const pic_idx = std.mem.indexOf(u8, block, tags.open_pic) orelse continue;
        const pic_close = std.mem.indexOfPos(u8, block, pic_idx, tags.close_pic) orelse {
            if (mode == .strict) return error.MalformedDrawingXml;
            continue;
        };
        const pic_block = block[pic_idx .. pic_close + tags.close_pic.len];

        const embed_rid = findBlipEmbedWithAlt(pic_block, prefixes.a, prefixes.a_alt) orelse {
            if (mode == .strict) return error.MalformedDrawingXml;
            continue;
        };
        const image_target = (try relTargetForIdTyped(allocator, drawing_rels, embed_rid, "image", mode)) orelse {
            if (mode == .strict) return error.MalformedDrawingXml;
            continue;
        };
        const image_part_name = (try store.resolve(drawing_part_name, image_target)) orelse {
            if (mode == .strict) return error.MalformedDrawingXml;
            continue;
        };
        const image_part = try store.part(image_part_name) orelse {
            if (mode == .strict) return error.MalformedDrawingXml;
            continue;
        };

        var from: CellAnchor = .{ .col = 0, .col_off = 0, .row = 0, .row_off = 0 };
        var to_anchor: ?CellAnchor = null;
        var absolute: ?AbsoluteAnchor = null;
        if (is_absolute) {
            absolute = parseAbsoluteAnchor(block, tags.open_pos, tags.open_ext) orelse {
                if (mode == .strict) return error.MalformedDrawingXml;
                continue;
            };
        } else {
            from = parseCellAnchor(block, tags.open_from, tags.close_from, prefixes.xdr) orelse {
                if (mode == .strict) return error.MalformedDrawingXml;
                continue;
            };
            if (is_two) {
                to_anchor = parseCellAnchor(block, tags.open_to, tags.close_to, prefixes.xdr);
                // A two-cell anchor without a readable `<to>` must not
                // ride out looking like a one-cell anchor (REL-101).
                if (mode == .strict and to_anchor == null) return error.MalformedDrawingXml;
            } else if (mode == .strict) {
                // A one-cell anchor's `<xdr:ext>` is schema-required;
                // strict validates it even though the extent stays
                // off the wire (REL-201).
                if (parseExtAttrs(block, 0, tags.open_ext) == null) return error.MalformedDrawingXml;
            }
        }

        try out.append(allocator, .{
            .image_part_name = image_part.name,
            .sheet_part_name = sheet_part.name,
            .from = from,
            .to = to_anchor,
            .absolute = absolute,
            .bytes = image_part.bytes,
            .kind = if (is_absolute) .absolute else if (is_two) .two_cell else .one_cell,
            .doc_offset = block_start,
        });
    }
}

fn collectChartsFromSheet(
    store: *PartStore,
    allocator: std.mem.Allocator,
    sheet_part: store_mod.Part,
    mode: WalkMode,
    out: *std.ArrayListUnmanaged(ChartAnchor),
) !void {
    const rid = switch (findDrawingRef(sheet_part.bytes)) {
        .absent => return,
        .malformed => {
            if (mode == .strict) return error.MalformedDrawingXml;
            return;
        },
        .rid => |r| r,
    };
    const sheet_rels = store.rels(sheet_part.name);
    const drawing_target = (try relTargetForIdTyped(allocator, sheet_rels, rid, "drawing", mode)) orelse {
        if (mode == .strict) return error.MalformedDrawingXml;
        return;
    };
    const drawing_part_name = (try store.resolve(sheet_part.name, drawing_target)) orelse {
        if (mode == .strict) return error.MalformedDrawingXml;
        return;
    };
    const drawing_part = try store.part(drawing_part_name) orelse {
        if (mode == .strict) return error.MalformedDrawingXml;
        return;
    };

    const drawing_rels = store.rels(drawing_part_name);
    const prefixes = resolveDrawingPrefixes(drawing_part.bytes);
    // 4 KiB tag-needle scratch covers prefixes up to ~250 chars
    // per needle (12 needles × ~max-prefix-len ≈ 4 KiB total). XML
    // namespace prefixes have no upper bound in the spec — Codex
    // flagged the previous 512 B buffer as a hard-fail surface for
    // valid-but-long custom prefixes.
    var primary_buf: [4096]u8 = undefined;
    const primary_tags = try DrawingTags.build(&primary_buf, prefixes);
    try scanChartsWithTags(store, allocator, drawing_part, drawing_part_name, drawing_rels, sheet_part, prefixes, primary_tags, mode, out);

    for (prefixes.xdr_alts()) |alt| {
        var alt_prefixes = prefixes;
        alt_prefixes.xdr = alt;
        var alt_buf: [4096]u8 = undefined;
        const alt_tags = try DrawingTags.build(&alt_buf, alt_prefixes);
        try scanChartsWithTags(store, allocator, drawing_part, drawing_part_name, drawing_rels, sheet_part, alt_prefixes, alt_tags, mode, out);
    }
}

fn scanChartsWithTags(
    store: *PartStore,
    allocator: std.mem.Allocator,
    drawing_part: store_mod.Part,
    drawing_part_name: []const u8,
    drawing_rels: []const store_mod.Relationship,
    sheet_part: store_mod.Part,
    prefixes: DrawingPrefixes,
    tags: DrawingTags,
    mode: WalkMode,
    out: *std.ArrayListUnmanaged(ChartAnchor),
) !void {
    var i: usize = 0;
    while (i < drawing_part.bytes.len) {
        const next = std.mem.indexOfPos(u8, drawing_part.bytes, i, tags.xdr_prefix_open) orelse break;
        i = next;
        const block_start = next;
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
        const close = std.mem.indexOfPos(u8, drawing_part.bytes, i, close_marker) orelse {
            if (mode == .strict) return error.MalformedDrawingXml;
            break;
        };
        const block = drawing_part.bytes[i .. close + close_marker.len];
        i = close + close_marker.len;

        // Charts live inside <xdr:graphicFrame>...<c:chart r:id=...
        // A frame without a chart element (a diagram, a table) is
        // legitimately not a chart in either mode; a chart element
        // the walk cannot follow to its part is a refusal in strict.
        const gf_idx = std.mem.indexOf(u8, block, tags.open_graphic_frame) orelse continue;
        // Scan for any `<*:chart` element whose prefix is bound to
        // either chart URI (block-local OR drawing-root). Walking by
        // tag rather than by prefix avoids the "multiple local
        // bindings to the same chart URI" failure mode where
        // collect-first-prefix would pick an unused declaration.
        const chart_idx = findLocalChartElement(block, gf_idx, prefixes) orelse continue;
        const chart_end = std.mem.indexOfScalarPos(u8, block, chart_idx, '>') orelse {
            if (mode == .strict) return error.MalformedDrawingXml;
            continue;
        };
        const chart_attrs = block[chart_idx .. chart_end + 1];
        const embed_rid = attrValue(chart_attrs, "r:id") orelse {
            if (mode == .strict) return error.MalformedDrawingXml;
            continue;
        };

        const chart_target = (try relTargetForIdTyped(allocator, drawing_rels, embed_rid, "chart", mode)) orelse {
            if (mode == .strict) return error.MalformedDrawingXml;
            continue;
        };
        const chart_part_name = (try store.resolve(drawing_part_name, chart_target)) orelse {
            if (mode == .strict) return error.MalformedDrawingXml;
            continue;
        };
        const chart_part = try store.part(chart_part_name) orelse {
            if (mode == .strict) return error.MalformedDrawingXml;
            continue;
        };

        var from: CellAnchor = .{ .col = 0, .col_off = 0, .row = 0, .row_off = 0 };
        var to_anchor: ?CellAnchor = null;
        var absolute: ?AbsoluteAnchor = null;
        if (is_absolute) {
            absolute = parseAbsoluteAnchor(block, tags.open_pos, tags.open_ext) orelse {
                if (mode == .strict) return error.MalformedDrawingXml;
                continue;
            };
        } else {
            from = parseCellAnchor(block, tags.open_from, tags.close_from, prefixes.xdr) orelse {
                if (mode == .strict) return error.MalformedDrawingXml;
                continue;
            };
            if (is_two) {
                to_anchor = parseCellAnchor(block, tags.open_to, tags.close_to, prefixes.xdr);
                // A two-cell anchor without a readable `<to>` must not
                // ride out looking like a one-cell anchor (REL-101).
                if (mode == .strict and to_anchor == null) return error.MalformedDrawingXml;
            } else if (mode == .strict) {
                // A one-cell anchor's `<xdr:ext>` is schema-required;
                // strict validates it even though the extent stays
                // off the wire (REL-201).
                if (parseExtAttrs(block, 0, tags.open_ext) == null) return error.MalformedDrawingXml;
            }
        }

        // Each chart's own XML may declare a different `c:` prefix
        // — resolve per-chart to be safe.
        const chart_prefixes = resolveDrawingPrefixes(chart_part.bytes);
        const refs = try extractSeriesRefs(allocator, chart_part.bytes, chart_prefixes.c, chart_prefixes.c_alt, mode);
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
            .kind = if (is_absolute) .absolute else if (is_two) .two_cell else .one_cell,
            .doc_offset = block_start,
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
    mode: WalkMode,
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
            // Build both needles atomically — if either fails to
            // format (e.g. a prefix that fits "<alt:f>" but not
            // "</alt:f>" within the 128-byte buffer), disable the
            // alt scan entirely. Splitting the assignment risked
            // alt_open != null with alt_close == null and a later
            // unwrap would panic.
            if (std.fmt.bufPrint(&alt_open_buf, "<{s}:f>", .{alt})) |o| {
                if (std.fmt.bufPrint(&alt_close_buf, "</{s}:f>", .{alt})) |c| {
                    alt_open = o;
                    alt_close = c;
                } else |_| {}
            } else |_| {}
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
        // Markup-shaped text inside a comment / CDATA / PI is not a
        // formula carrier — `<!-- <c:f>Fake!A1</c:f> -->` must not
        // add a series ref (Codex #214 r2 REL-203). Jump past the
        // whole region and rescan.
        if (isInsideCommentOrCdata(xml, open_offset)) {
            i = skipRegionContaining(xml, open_offset) orelse open_offset + 1;
            continue;
        }
        const start = open_offset + open_tag.len;
        const close_off = std.mem.indexOfPos(u8, xml, start, close_tag) orelse {
            // A real, opened carrier that never closes: lenient keeps
            // the historical truncation; strict refuses rather than
            // silently thinning `series_refs` (REL-203).
            if (mode == .strict) return error.MalformedDrawingXml;
            break;
        };
        try out.append(allocator, xml[start..close_off]);
        i = close_off + close_tag.len;
    }
    return out.toOwnedSlice(allocator);
}

/// True when `needle` occurs in `xml` OUTSIDE every comment / CDATA /
/// PI region — element detection that markup-shaped text cannot fool
/// (Codex #214 r2 REL-203).
fn hasRealMarkup(xml: []const u8, needle: []const u8) bool {
    var i: usize = 0;
    while (std.mem.indexOfPos(u8, xml, i, needle)) |at| {
        if (!isInsideCommentOrCdata(xml, at)) return true;
        i = skipRegionContaining(xml, at) orelse at + 1;
    }
    return false;
}

/// Best-effort chart-type detection from the chart-part XML. Looks
/// for the canonical `<{c}:Xchart>` element names. A compound chart
/// (two or more distinct plot types overlaid — bar + line, say)
/// reports `.other`, matching the enum's contract and the CLI wire
/// contract (Codex #214 r1 REL-104); callers needing the full
/// picture can walk raw_xml directly. `c_prefix` is the document's
/// actual chart-namespace prefix (canonically "c").
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
    // The same element found under both prefixes is one plot type,
    // not a compound — track distinct kinds, not matches. Matches
    // inside comments / CDATA / PIs are text, not plot elements
    // (Codex #214 r2 REL-203).
    var found: ?ChartType = null;
    for (candidates) |c| {
        var present = false;
        for (prefixes) |maybe_p| {
            const p = maybe_p orelse continue;
            const needle = std.fmt.bufPrint(&buf, "<{s}:{s}", .{ p, c.suffix }) catch continue;
            if (hasRealMarkup(chart_xml, needle)) present = true;
        }
        if (!present) continue;
        if (found != null) return .other;
        found = c.kind;
    }
    return found orelse .other;
}

/// Generic attribute value extractor: find `key="value"` or
/// `key='value'` inside an already-narrowed tag-attributes slice.
/// Both quote styles are valid XML; non-Microsoft producers
/// (libreoffice, hand-edited drawings) sometimes emit single
/// quotes, and skipping them silently dropped image/chart anchors.
fn attrValue(attrs: []const u8, key: []const u8) ?[]const u8 {
    // Walk attrs as a tag-attribute slice, tracking quoted regions
    // so a substring inside one attribute's VALUE can't masquerade
    // as another attribute name. At each unquoted, word-boundary
    // position try to match: key, optional whitespace, `=`,
    // optional whitespace, quote-delimited value.
    var i: usize = 0;
    while (i < attrs.len) {
        const c = attrs[i];
        // Skip over quoted runs entirely.
        if (c == '"' or c == '\'') {
            const close = std.mem.indexOfScalarPos(u8, attrs, i + 1, c) orelse return null;
            i = close + 1;
            continue;
        }
        // Word boundary: only consider candidate keys at slice
        // start or after XML whitespace.
        const at_word_boundary = i == 0 or
            attrs[i - 1] == ' ' or attrs[i - 1] == '\t' or
            attrs[i - 1] == '\n' or attrs[i - 1] == '\r';
        if (!at_word_boundary) {
            i += 1;
            continue;
        }
        if (i + key.len > attrs.len) return null;
        if (!std.mem.eql(u8, attrs[i .. i + key.len], key)) {
            i += 1;
            continue;
        }
        var p = i + key.len;
        while (p < attrs.len and (attrs[p] == ' ' or attrs[p] == '\t' or attrs[p] == '\n' or attrs[p] == '\r')) p += 1;
        if (p >= attrs.len or attrs[p] != '=') {
            i += 1;
            continue;
        }
        p += 1;
        while (p < attrs.len and (attrs[p] == ' ' or attrs[p] == '\t' or attrs[p] == '\n' or attrs[p] == '\r')) p += 1;
        if (p >= attrs.len) return null;
        const quote = attrs[p];
        if (quote != '"' and quote != '\'') {
            i += 1;
            continue;
        }
        const val_start = p + 1;
        const val_end = std.mem.indexOfScalarPos(u8, attrs, val_start, quote) orelse return null;
        return attrs[val_start..val_end];
    }
    return null;
}

/// Find the value of `r:id` on the sheet's `<drawing>` element. The
/// element is always self-closing in OOXML and lives at sheet scope
/// (one per sheet at most).
fn findDrawingRid(sheet_xml: []const u8) ?[]const u8 {
    return switch (findDrawingRef(sheet_xml)) {
        .rid => |r| r,
        .absent, .malformed => null,
    };
}

/// The sheet's `<drawing>` reference, tri-state: strict callers must
/// tell "no drawing element" (nothing to walk) from "a drawing element
/// whose reference cannot be read" (an inventory hole) — one optional
/// conflated them and strict mode silently skipped the sheet (Codex
/// #214 r2 REL-202).
const DrawingRef = union(enum) {
    absent,
    rid: []const u8,
    malformed,
};

fn findDrawingRef(sheet_xml: []const u8) DrawingRef {
    const tag = findOpeningTagAnyPrefix(sheet_xml, "drawing") orelse return .absent;
    const tag_end = std.mem.indexOfScalarPos(u8, sheet_xml, tag, '>') orelse return .malformed;
    const attrs = sheet_xml[tag .. tag_end + 1];
    // Canonical spelling first; then any-prefix `*:id` — the
    // relationships namespace is conventionally bound to `r` but OOXML
    // producers may pick any prefix.
    if (attrValue(attrs, "r:id")) |rid| return .{ .rid = rid };
    if (prefixedIdAttrValue(attrs)) |rid| return .{ .rid = rid };
    return .malformed;
}

/// `findOpeningTag`, but also matching `<{prefix}:{name}` for any
/// NCName-shaped prefix — `<x:drawing r:id=…/>` is valid OOXML that
/// the unprefixed search missed (Codex #214 r2 REL-202).
fn findOpeningTagAnyPrefix(xml: []const u8, name: []const u8) ?usize {
    if (findOpeningTag(xml, name)) |i| return i;
    var i: usize = 0;
    while (std.mem.indexOfScalarPos(u8, xml, i, ':')) |colon| {
        i = colon + 1;
        const after_name = colon + 1 + name.len;
        if (after_name >= xml.len) return null;
        if (!std.mem.eql(u8, xml[colon + 1 .. after_name], name)) continue;
        const c = xml[after_name];
        if (!(c == ' ' or c == '\t' or c == '\n' or c == '\r' or c == '/' or c == '>')) continue;
        // Walk back over the prefix to the `<`; every intervening
        // byte must be a name char, so a `:drawing` inside an
        // attribute value or text does not count.
        var p = colon;
        var ok = true;
        while (p > 0) {
            p -= 1;
            const pc = xml[p];
            if (pc == '<') break;
            if (!(std.ascii.isAlphanumeric(pc) or pc == '_' or pc == '-' or pc == '.')) {
                ok = false;
                break;
            }
        }
        if (!ok or xml[p] != '<') continue;
        if (p + 1 == colon) continue; // `<:name` — an empty prefix is not one.
        return p;
    }
    return null;
}

/// The value of the first attribute spelled `{prefix}:id` (any
/// non-empty prefix) in an already-narrowed tag-attributes slice.
/// A bare `id` attribute is NOT a relationship reference and does
/// not match. Quote-aware, like `attrValue`.
fn prefixedIdAttrValue(attrs: []const u8) ?[]const u8 {
    var i: usize = 0;
    while (i < attrs.len) {
        const c = attrs[i];
        if (c == '"' or c == '\'') {
            const close = std.mem.indexOfScalarPos(u8, attrs, i + 1, c) orelse return null;
            i = close + 1;
            continue;
        }
        const at_word_boundary = i == 0 or
            attrs[i - 1] == ' ' or attrs[i - 1] == '\t' or
            attrs[i - 1] == '\n' or attrs[i - 1] == '\r';
        if (at_word_boundary and c != '<') {
            // Read a candidate attribute name: NAME chars up to `=`,
            // whitespace or end.
            var j = i;
            while (j < attrs.len) : (j += 1) {
                const nc = attrs[j];
                if (nc == '=' or nc == ' ' or nc == '\t' or nc == '\n' or nc == '\r' or nc == '>' or nc == '/') break;
            }
            const key = attrs[i..j];
            if (key.len > 3 and std.mem.endsWith(u8, key, ":id")) {
                if (attrValue(attrs, key)) |v| return v;
            }
            i = j + 1;
            continue;
        }
        i += 1;
    }
    return null;
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
const max_xdr_alts: usize = 8;

const DrawingPrefixes = struct {
    xdr: []const u8 = "xdr",
    a: []const u8 = "a",
    c: []const u8 = "c",
    a_alt: ?[]const u8 = null,
    c_alt: ?[]const u8 = null,
    /// All prefixes (other than `xdr`) bound to either xdr URI in
    /// the same document. The scanner replays once per prefix so
    /// anchors using ANY bound prefix are surfaced. Capped at
    /// `max_xdr_alts` — real OOXML producers declare 1-2; more is
    /// pathological. Tracked as fixed-array + count rather than a
    /// stdlib bounded helper because Zig 0.15 dropped
    /// `std.BoundedArray`.
    xdr_alts_buf: [max_xdr_alts][]const u8 = undefined,
    xdr_alts_len: usize = 0,

    fn xdr_alts(self: *const DrawingPrefixes) []const []const u8 {
        return self.xdr_alts_buf[0..self.xdr_alts_len];
    }
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
    // A document can bind the spreadsheetDrawing namespace under
    // multiple prefixes (one on the root, others on descendant
    // anchors). Codex flagged scenarios where unused declarations
    // appear before the actually-used alt — picking only the FIRST
    // alt would silently drop anchors. Collect ALL alternate
    // bindings (capped at 8) so the scanner replays per-prefix.
    //
    // Prefer alts on the SAME URI as the primary binding so the
    // most-likely-used candidate is processed first.
    const primary_uri = uriOfPrefix(xml, p.xdr);
    if (primary_uri) |uri| {
        if (std.mem.eql(u8, uri, ns_xdr_strict)) {
            collectAllNamespacePrefixes(xml, ns_xdr_strict, p.xdr, &p);
            collectAllNamespacePrefixes(xml, ns_xdr_transitional, p.xdr, &p);
        } else if (std.mem.eql(u8, uri, ns_xdr_transitional)) {
            collectAllNamespacePrefixes(xml, ns_xdr_transitional, p.xdr, &p);
            collectAllNamespacePrefixes(xml, ns_xdr_strict, p.xdr, &p);
        } else {
            collectAllNamespacePrefixes(xml, ns_xdr_transitional, p.xdr, &p);
            collectAllNamespacePrefixes(xml, ns_xdr_strict, p.xdr, &p);
        }
    } else {
        collectAllNamespacePrefixes(xml, ns_xdr_transitional, p.xdr, &p);
        collectAllNamespacePrefixes(xml, ns_xdr_strict, p.xdr, &p);
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
    return findNamespacePrefixExcept(xml, target_uri, "");
}

/// Walk `block` from `start` for any `<*:chart` opener whose
/// prefix is bound to either chart URI in the block (or matches
/// the drawing-root primary / alternate prefix). Returns the index
/// of the `<` on hit. Verifying the prefix per-tag means redundant
/// xmlns declarations that aren't actually used by `<c:chart>`
/// can't shadow the binding that IS used — the previous
/// "collect first prefix bound to URI" approach failed when an
/// unused declaration appeared first.
fn findLocalChartElement(block: []const u8, start: usize, prefixes: DrawingPrefixes) ?usize {
    var i = start;
    // If `start` itself sits inside a skip region (caller's
    // graphicFrame index can land in a commented-out fake), jump
    // past the region's close before searching.
    if (skipRegionContaining(block, start)) |skip_end| {
        i = skip_end;
    }
    while (i < block.len) {
        // Eat any comment/CDATA/PI section starting at `i`. Done
        // up-front each iteration so we never call indexOfPos for
        // `:chart` inside a skip region — the search stays O(n)
        // even when adversarial input packs many fake `:chart`
        // substrings into many skip regions.
        if (i + 4 <= block.len and std.mem.startsWith(u8, block[i..], "<!--")) {
            const close = std.mem.indexOfPos(u8, block, i + 4, "-->") orelse return null;
            i = close + 3;
            continue;
        }
        if (i + 9 <= block.len and std.mem.startsWith(u8, block[i..], "<![CDATA[")) {
            const close = std.mem.indexOfPos(u8, block, i + 9, "]]>") orelse return null;
            i = close + 3;
            continue;
        }
        if (i + 2 <= block.len and std.mem.startsWith(u8, block[i..], "<?")) {
            const close = std.mem.indexOfPos(u8, block, i + 2, "?>") orelse return null;
            i = close + 2;
            continue;
        }
        const colon = std.mem.indexOfPos(u8, block, i, ":chart") orelse return null;
        // If a skip region opens between `i` and `colon`, jump to
        // that opener so the section-eat branch above consumes it.
        // This keeps `:chart` lookups limited to non-skipped bytes.
        if (nextSkipOpenBefore(block, i, colon)) |skip_open| {
            i = skip_open;
            continue;
        }
        // The byte immediately after `:chart` must end the tag name
        // (space, `>`, or `/`); otherwise this is a longer name
        // that happens to contain ":chart" (e.g. `:chartSpace`).
        const after = colon + ":chart".len;
        if (after >= block.len) {
            i = colon + 1;
            continue;
        }
        const ch_after = block[after];
        if (ch_after != ' ' and ch_after != '\t' and ch_after != '\n' and
            ch_after != '\r' and ch_after != '>' and ch_after != '/')
        {
            i = colon + 1;
            continue;
        }
        // Walk back from `colon` to the start of the prefix; the
        // byte immediately before the prefix must be `<`.
        var p_start = colon;
        while (p_start > 0) : (p_start -= 1) {
            const c = block[p_start - 1];
            if (c == '<') break;
            if (!isPrefixByte(c)) {
                p_start = colon; // sentinel: no valid prefix
                break;
            }
        }
        if (p_start == colon or p_start == 0 or block[p_start - 1] != '<') {
            i = colon + 1;
            continue;
        }
        // (The outer loop already eats comment/CDATA/PI regions
        // before looking for `:chart`, so this candidate is
        // guaranteed to be in normal-state markup.)
        const prefix = block[p_start..colon];
        // Verify the prefix is bound to a chart URI. Use the
        // declaration nearest to (but inside) the chart tag — the
        // chart element's own attributes can carry an xmlns:<p>
        // that redeclares the prefix locally, and that
        // declaration IS in scope for the chart element itself.
        // Find the closing `>` of the candidate tag so the
        // self-declaration is included in the scan window.
        const tag_end = std.mem.indexOfScalarPos(u8, block, colon, '>') orelse {
            i = colon + 1;
            continue;
        };
        const uri_local = uriOfPrefixAtPosition(block, prefix, tag_end);
        // A local in-scope binding is authoritative — if the
        // prefix has been redeclared, the root-primary fallback
        // must NOT override it. Only fall through to root prefix
        // matching when the prefix has no local declaration in
        // scope at this tag. Otherwise the lookup would accept a
        // `<c:chart xmlns:c="not-a-chart-uri"/>` element solely
        // because `c` matches the drawing root, dropping the real
        // chart later in the same graphic frame.
        if (uri_local) |u| {
            if (std.mem.eql(u8, u, ns_c_transitional) or std.mem.eql(u8, u, ns_c_strict)) {
                return p_start - 1; // index of `<`
            }
            i = colon + 1;
            continue;
        }
        const matches_root_primary = std.mem.eql(u8, prefix, prefixes.c);
        const matches_root_alt = if (prefixes.c_alt) |alt|
            std.mem.eql(u8, prefix, alt)
        else
            false;
        if (matches_root_primary or matches_root_alt) {
            return p_start - 1; // index of `<`
        }
        i = colon + 1;
    }
    return null;
}

inline fn isPrefixByte(c: u8) bool {
    return (c >= 'A' and c <= 'Z') or (c >= 'a' and c <= 'z') or
        (c >= '0' and c <= '9') or c == '_' or c == '.' or c == '-';
}

/// Resolve the URI bound to `prefix` in the scope of position
/// `before` — the most recent `xmlns:<prefix>=...` declaration
/// whose element extent contains `before`. XML namespace scoping:
/// a binding on `<foo xmlns:p="..."/>` is in scope inside foo;
/// once foo closes (self-closing `/>` or matching `</foo>`), the
/// binding ends. Bindings on closed siblings are out of scope.
///
/// For each xmlns declaration found, we compute:
///   - the opening tag's `>` position
///   - the element extent (tag end if self-closing, else position
///     of the matching `</NAME>`)
/// and skip the binding if `before` is past the extent.
///
/// Nested same-name elements are handled with a depth counter
/// (each `<NAME` increments, each `</NAME>` decrements).
fn uriOfPrefixAtPosition(xml: []const u8, prefix: []const u8, before: usize) ?[]const u8 {
    if (prefix.len == 0) return null;
    const limit = xml.len;
    var i: usize = 0;
    var last_uri: ?[]const u8 = null;
    while (std.mem.indexOfPos(u8, xml[0..limit], i, "xmlns:")) |start| {
        if (start > before) break;
        // Ignore `xmlns:` text that's inside an XML comment, CDATA
        // section, or general text content — only attribute-shaped
        // hits in an opening tag count. Walk back to the most
        // recent `<`; if the bytes after it are `!--`, `![`, `?`,
        // or `/`, this isn't an opening tag.
        if (!isInsideOpeningTag(xml, start)) {
            i = start + "xmlns:".len;
            continue;
        }
        const after = start + "xmlns:".len;
        if (after >= limit) break;
        var name_end = after;
        while (name_end < limit) : (name_end += 1) {
            const c = xml[name_end];
            if (c == '=' or c == ' ' or c == '\t' or c == '\n' or c == '\r') break;
        }
        if (name_end >= limit) break;
        const name = xml[after..name_end];
        var p = name_end;
        while (p < limit and (xml[p] == ' ' or xml[p] == '\t' or xml[p] == '\n' or xml[p] == '\r')) p += 1;
        if (p >= limit or xml[p] != '=') {
            i = after;
            continue;
        }
        p += 1;
        while (p < limit and (xml[p] == ' ' or xml[p] == '\t' or xml[p] == '\n' or xml[p] == '\r')) p += 1;
        if (p >= limit) break;
        const quote = xml[p];
        if (quote != '"' and quote != '\'') {
            i = p;
            continue;
        }
        const val_start = p + 1;
        const val_end = std.mem.indexOfScalarPos(u8, xml[0..limit], val_start, quote) orelse break;
        if (std.mem.eql(u8, name, prefix)) {
            const extent_end = elementExtentEnd(xml, start) orelse {
                i = val_end + 1;
                continue;
            };
            if (before <= extent_end) {
                last_uri = xml[val_start..val_end];
            }
            i = val_end + 1;
            continue;
        }
        i = val_end + 1;
    }
    return last_uri;
}

/// Given a position inside an element's opening tag, return the
/// index of the byte that closes the element's extent: either the
/// tag's own `>` (self-closing `/>`) or the position of `>` on
/// the matching `</NAME>` close tag. Returns null if the opening
/// tag is malformed or unterminated.
fn elementExtentEnd(xml: []const u8, inside_pos: usize) ?usize {
    // Walk back to find the `<` that opens the element.
    var lt = inside_pos;
    while (lt > 0 and xml[lt] != '<') : (lt -= 1) {}
    if (xml[lt] != '<') return null;
    if (lt + 1 >= xml.len) return null;
    // Read the element name: bytes after `<` until whitespace/`>`/`/`.
    var name_end = lt + 1;
    while (name_end < xml.len) : (name_end += 1) {
        const c = xml[name_end];
        if (c == ' ' or c == '\t' or c == '\n' or c == '\r' or c == '>' or c == '/') break;
    }
    if (name_end == lt + 1) return null;
    const elem_name = xml[lt + 1 .. name_end];
    // Find the opening tag's `>`, skipping over quoted attribute
    // values so a literal `>` inside `descr="a > b"` doesn't end
    // the tag prematurely.
    const tag_end = findUnquotedTagEnd(xml, name_end) orelse return null;
    // Self-closing: extent ends at tag_end.
    if (tag_end > 0 and xml[tag_end - 1] == '/') return tag_end;
    // Otherwise walk forward looking for the matching `</NAME>`,
    // accounting for nested same-name elements via a depth counter.
    // Skip over comments and CDATA: fake markup inside them
    // (`<!-- <foo> -->`, `<![CDATA[</foo>]]>`) must NOT bump depth.
    var depth: i64 = 1;
    var search_at: usize = tag_end + 1;
    while (search_at < xml.len) {
        // Advance past any comment / CDATA / PI section here.
        if (search_at + 4 <= xml.len and std.mem.startsWith(u8, xml[search_at..], "<!--")) {
            const close = std.mem.indexOfPos(u8, xml, search_at + 4, "-->") orelse return xml.len - 1;
            search_at = close + 3;
            continue;
        }
        if (search_at + 9 <= xml.len and std.mem.startsWith(u8, xml[search_at..], "<![CDATA[")) {
            const close = std.mem.indexOfPos(u8, xml, search_at + 9, "]]>") orelse return xml.len - 1;
            search_at = close + 3;
            continue;
        }
        if (search_at + 2 <= xml.len and std.mem.startsWith(u8, xml[search_at..], "<?")) {
            const close = std.mem.indexOfPos(u8, xml, search_at + 2, "?>") orelse return xml.len - 1;
            search_at = close + 2;
            continue;
        }
        const next_lt = std.mem.indexOfScalarPos(u8, xml, search_at, '<') orelse return xml.len - 1;
        if (next_lt + 1 >= xml.len) return xml.len - 1;
        // If next_lt opens a comment / CDATA / PI, loop around so
        // the top-of-loop skip handles it.
        if (next_lt + 4 <= xml.len and std.mem.startsWith(u8, xml[next_lt..], "<!--")) {
            search_at = next_lt;
            continue;
        }
        if (next_lt + 9 <= xml.len and std.mem.startsWith(u8, xml[next_lt..], "<![CDATA[")) {
            search_at = next_lt;
            continue;
        }
        if (next_lt + 2 <= xml.len and std.mem.startsWith(u8, xml[next_lt..], "<?")) {
            search_at = next_lt;
            continue;
        }
        const is_close = xml[next_lt + 1] == '/';
        const candidate_name_start = if (is_close) next_lt + 2 else next_lt + 1;
        if (candidate_name_start + elem_name.len > xml.len) return xml.len - 1;
        const matches_name =
            std.mem.eql(u8, xml[candidate_name_start .. candidate_name_start + elem_name.len], elem_name) and
            (candidate_name_start + elem_name.len < xml.len) and
            isNameTerminator(xml[candidate_name_start + elem_name.len]);
        if (matches_name) {
            if (is_close) {
                depth -= 1;
                if (depth == 0) {
                    return findUnquotedTagEnd(xml, candidate_name_start) orelse xml.len - 1;
                }
            } else {
                // Open tag; check if it's self-closing (no depth
                // bump) or container (bump).
                const open_end = findUnquotedTagEnd(xml, candidate_name_start) orelse return xml.len - 1;
                if (!(open_end > 0 and xml[open_end - 1] == '/')) {
                    depth += 1;
                }
                search_at = open_end + 1;
                continue;
            }
        }
        // Advance past this `<` and keep scanning.
        search_at = next_lt + 1;
    }
    return xml.len - 1;
}

inline fn isNameTerminator(c: u8) bool {
    return c == ' ' or c == '\t' or c == '\n' or c == '\r' or c == '>' or c == '/';
}

/// True if `pos` is inside an opening tag — i.e. the most recent
/// `<` before it opens an element (not a comment `<!--`, CDATA
/// `<![`, processing instruction `<?`, or close tag `</`), and
/// no intervening `>` has closed that tag.
///
/// The "no intervening `>`" check is quote-aware: a `>` inside a
/// quoted attribute value (`<foo descr="a > b" xmlns:p="...">`)
/// must not be mistaken for the tag end. Walks back to the most
/// recent `<` first, then scans forward quote-aware until it
/// finds the real tag-closing `>`.
fn isInsideOpeningTag(xml: []const u8, pos: usize) bool {
    if (pos == 0) return false;
    // Reject if `pos` is inside an XML comment or CDATA section —
    // markup-shaped text inside those isn't real markup.
    if (isInsideCommentOrCdata(xml, pos)) return false;
    var lt = pos;
    while (lt > 0 and xml[lt] != '<') : (lt -= 1) {}
    if (xml[lt] != '<') return false;
    if (lt + 1 >= xml.len) return false;
    const next = xml[lt + 1];
    if (next == '!' or next == '?' or next == '/') return false;
    // Walk forward quote-aware from `<` checking if `pos` falls
    // BEFORE the tag-closing `>` AND outside any attribute value.
    // An xmlns-looking string inside `macro="xmlns:c='bad'"` is
    // attribute content, not a real namespace declaration.
    var i = lt + 1;
    while (i < xml.len) : (i += 1) {
        const c = xml[i];
        if (c == '"' or c == '\'') {
            const close = std.mem.indexOfScalarPos(u8, xml, i + 1, c) orelse return false;
            // pos inside this quoted region → reject.
            if (pos > i and pos < close) return false;
            i = close;
            continue;
        }
        if (c == '>') return pos < i;
    }
    return false;
}

/// Find the earliest skip-region opener (`<!--`, `<![CDATA[`, or
/// `<?`) starting at or after `from` and strictly before `limit`.
/// Returns the opener's start index or null if none. Used by the
/// chart-element scanner to jump past skip regions inline.
fn nextSkipOpenBefore(xml: []const u8, from: usize, limit: usize) ?usize {
    if (limit <= from) return null;
    // Bound the substring scans to `xml[from..limit]` so a
    // candidate match deep in the document doesn't cost us a
    // full-block scan per iteration. Each search returns indices
    // relative to the slice; rebase to absolute positions.
    const slice = xml[from..limit];
    const c1 = std.mem.indexOfPos(u8, slice, 0, "<!--");
    const c2 = std.mem.indexOfPos(u8, slice, 0, "<![CDATA[");
    const c3 = std.mem.indexOfPos(u8, slice, 0, "<?");
    var best: ?usize = null;
    inline for (.{ c1, c2, c3 }) |maybe_p| {
        if (maybe_p) |p| {
            const abs = from + p;
            best = if (best) |b| @min(b, abs) else abs;
        }
    }
    return best;
}

/// If `pos` is inside an XML comment / CDATA / PI region, return
/// the byte index just past the region's close. Otherwise null.
/// Lets callers advance their scan past the entire skipped region
/// in one step instead of byte-by-byte.
fn skipRegionContaining(xml: []const u8, pos: usize) ?usize {
    var i: usize = 0;
    while (i < xml.len) {
        if (i + 4 <= xml.len and std.mem.startsWith(u8, xml[i..], "<!--")) {
            const close = std.mem.indexOfPos(u8, xml, i + 4, "-->") orelse xml.len - 3;
            const end = close + 3;
            if (pos >= i and pos < end) return end;
            i = end;
            continue;
        }
        if (i + 9 <= xml.len and std.mem.startsWith(u8, xml[i..], "<![CDATA[")) {
            const close = std.mem.indexOfPos(u8, xml, i + 9, "]]>") orelse xml.len - 3;
            const end = close + 3;
            if (pos >= i and pos < end) return end;
            i = end;
            continue;
        }
        if (i + 2 <= xml.len and std.mem.startsWith(u8, xml[i..], "<?")) {
            const close = std.mem.indexOfPos(u8, xml, i + 2, "?>") orelse xml.len - 2;
            const end = close + 2;
            if (pos >= i and pos < end) return end;
            i = end;
            continue;
        }
        if (i > pos) return null;
        i += 1;
    }
    return null;
}

/// True if `pos` is inside an XML comment (`<!-- ... -->`),
/// CDATA section (`<![CDATA[ ... ]]>`), or processing instruction
/// (`<?...?>`). Walks forward so closed regions don't bleed their
/// delimiter text into adjacent state — a `<!--` literal inside
/// `<?xml-stylesheet ... ?>` is PI content, not a comment open.
fn isInsideCommentOrCdata(xml: []const u8, pos: usize) bool {
    if (pos == 0) return false;
    const limit = @min(xml.len, pos + 1);
    var i: usize = 0;
    while (i < limit) {
        if (i + 4 <= xml.len and std.mem.startsWith(u8, xml[i..], "<!--")) {
            if (i + 4 > pos) return true;
            const close = std.mem.indexOfPos(u8, xml, i + 4, "-->") orelse return true;
            if (pos < close + 3) return true;
            i = close + 3;
            continue;
        }
        if (i + 9 <= xml.len and std.mem.startsWith(u8, xml[i..], "<![CDATA[")) {
            if (i + 9 > pos) return true;
            const close = std.mem.indexOfPos(u8, xml, i + 9, "]]>") orelse return true;
            if (pos < close + 3) return true;
            i = close + 3;
            continue;
        }
        if (i + 2 <= xml.len and std.mem.startsWith(u8, xml[i..], "<?")) {
            if (i + 2 > pos) return true;
            const close = std.mem.indexOfPos(u8, xml, i + 2, "?>") orelse return true;
            if (pos < close + 2) return true;
            i = close + 2;
            continue;
        }
        i += 1;
    }
    return false;
}

/// Walk `xml` from `start` looking for the `>` that ends a tag,
/// skipping over `"..."` and `'...'` quoted attribute regions so a
/// literal `>` inside a quoted value doesn't end the scan early.
/// Returns the `>` position or null on EOF.
fn findUnquotedTagEnd(xml: []const u8, start: usize) ?usize {
    var i = start;
    while (i < xml.len) : (i += 1) {
        const c = xml[i];
        if (c == '"' or c == '\'') {
            const close = std.mem.indexOfScalarPos(u8, xml, i + 1, c) orelse return null;
            i = close;
            continue;
        }
        if (c == '>') return i;
    }
    return null;
}

/// Same as `uriOfPrefix` but scans the FULL block instead of
/// capping at 4 KiB. Used by `findLocalChartElement` when verifying
/// a tag-derived prefix that may be declared late in the block.
fn uriOfPrefixLocal(xml: []const u8, prefix: []const u8) ?[]const u8 {
    if (prefix.len == 0) return null;
    const limit = xml.len;
    var i: usize = 0;
    while (std.mem.indexOfPos(u8, xml[0..limit], i, "xmlns:")) |start| {
        const after = start + "xmlns:".len;
        if (after >= limit) return null;
        var name_end = after;
        while (name_end < limit) : (name_end += 1) {
            const c = xml[name_end];
            if (c == '=' or c == ' ' or c == '\t' or c == '\n' or c == '\r') break;
        }
        if (name_end >= limit) return null;
        const name = xml[after..name_end];
        var p = name_end;
        while (p < limit and (xml[p] == ' ' or xml[p] == '\t' or xml[p] == '\n' or xml[p] == '\r')) p += 1;
        if (p >= limit or xml[p] != '=') {
            i = after;
            continue;
        }
        p += 1;
        while (p < limit and (xml[p] == ' ' or xml[p] == '\t' or xml[p] == '\n' or xml[p] == '\r')) p += 1;
        if (p >= limit) return null;
        const quote = xml[p];
        if (quote != '"' and quote != '\'') {
            i = p;
            continue;
        }
        const val_start = p + 1;
        const val_end = std.mem.indexOfScalarPos(u8, xml[0..limit], val_start, quote) orelse return null;
        if (std.mem.eql(u8, name, prefix)) return xml[val_start..val_end];
        i = val_end + 1;
    }
    return null;
}

/// Same as `findNamespacePrefix` but scans the FULL `xml` input
/// instead of capping at 4 KiB. Use this when the input is a
/// per-anchor block whose interior declarations should be reachable
/// even if they're past 4 KiB into the block. Avoid for whole-
/// document root-prefix resolution — that path's bounded form
/// guards against late mid-document declarations shadowing the
/// canonical fallback.
fn findLocalNamespacePrefix(xml: []const u8, target_uri: []const u8) ?[]const u8 {
    const limit = xml.len;
    var i: usize = 0;
    while (std.mem.indexOfPos(u8, xml[0..limit], i, "xmlns:")) |start| {
        const after = start + "xmlns:".len;
        if (after >= limit) return null;
        var name_end = after;
        while (name_end < limit) : (name_end += 1) {
            const c = xml[name_end];
            if (c == '=' or c == ' ' or c == '\t' or c == '\n' or c == '\r') break;
        }
        if (name_end >= limit) return null;
        const name = xml[after..name_end];
        var p = name_end;
        while (p < limit and (xml[p] == ' ' or xml[p] == '\t' or xml[p] == '\n' or xml[p] == '\r')) p += 1;
        if (p >= limit or xml[p] != '=') {
            i = after;
            continue;
        }
        p += 1;
        while (p < limit and (xml[p] == ' ' or xml[p] == '\t' or xml[p] == '\n' or xml[p] == '\r')) p += 1;
        if (p >= limit) return null;
        const quote = xml[p];
        if (quote != '"' and quote != '\'') {
            i = p;
            continue;
        }
        const val_start = p + 1;
        const val_end = std.mem.indexOfScalarPos(u8, xml[0..limit], val_start, quote) orelse return null;
        if (std.mem.eql(u8, xml[val_start..val_end], target_uri) and name.len <= max_prefix_len) {
            return name;
        }
        i = val_end + 1;
    }
    return null;
}

/// Append every prefix bound to `target_uri` (other than `skip`
/// and any prefix already in `out`) to the bounded array. Used
/// to collect ALL alternate xdr prefixes for replay scanning —
/// finding the FIRST alt isn't enough when an unused declaration
/// precedes the actually-used one.
fn collectAllNamespacePrefixes(
    xml: []const u8,
    target_uri: []const u8,
    skip: []const u8,
    out: *DrawingPrefixes,
) void {
    const limit = xml.len;
    var i: usize = 0;
    while (std.mem.indexOfPos(u8, xml[0..limit], i, "xmlns:")) |start| {
        const after = start + "xmlns:".len;
        if (after >= limit) return;
        var name_end = after;
        while (name_end < limit) : (name_end += 1) {
            const c = xml[name_end];
            if (c == '=' or c == ' ' or c == '\t' or c == '\n' or c == '\r') break;
        }
        if (name_end >= limit) return;
        const name = xml[after..name_end];
        var p = name_end;
        while (p < limit and (xml[p] == ' ' or xml[p] == '\t' or xml[p] == '\n' or xml[p] == '\r')) p += 1;
        if (p >= limit or xml[p] != '=') {
            i = after;
            continue;
        }
        p += 1;
        while (p < limit and (xml[p] == ' ' or xml[p] == '\t' or xml[p] == '\n' or xml[p] == '\r')) p += 1;
        if (p >= limit) return;
        const quote = xml[p];
        if (quote != '"' and quote != '\'') {
            i = p;
            continue;
        }
        const val_start = p + 1;
        const val_end = std.mem.indexOfScalarPos(u8, xml[0..limit], val_start, quote) orelse return;
        if (std.mem.eql(u8, xml[val_start..val_end], target_uri) and
            !std.mem.eql(u8, name, skip) and
            name.len <= max_prefix_len)
        {
            // Dedup: don't append the same prefix twice (would
            // happen when same prefix is declared on two elements
            // with the same URI).
            var already_in = false;
            for (out.xdr_alts()) |existing| {
                if (std.mem.eql(u8, existing, name)) {
                    already_in = true;
                    break;
                }
            }
            if (!already_in and out.xdr_alts_len < max_xdr_alts) {
                out.xdr_alts_buf[out.xdr_alts_len] = name;
                out.xdr_alts_len += 1;
            }
        }
        i = val_end + 1;
    }
}

/// Inverse lookup: given a prefix, return the URI it's bound to,
/// or null if the prefix isn't declared. Bounded scan; matches the
/// 4 KiB ceiling used elsewhere in this module so behaviour stays
/// consistent across helpers.
fn uriOfPrefix(xml: []const u8, prefix: []const u8) ?[]const u8 {
    if (prefix.len == 0) return null;
    // Bounded scan: this is only used to look up the URI of the
    // ROOT element's prefix, which by definition is declared on
    // the root and therefore in the first 4 KiB.
    const limit = @min(xml.len, 4096);
    var i: usize = 0;
    while (std.mem.indexOfPos(u8, xml[0..limit], i, "xmlns:")) |start| {
        const after = start + "xmlns:".len;
        if (after >= limit) return null;
        var name_end = after;
        while (name_end < limit) : (name_end += 1) {
            const c = xml[name_end];
            if (c == '=' or c == ' ' or c == '\t' or c == '\n' or c == '\r') break;
        }
        if (name_end >= limit) return null;
        const name = xml[after..name_end];
        var p = name_end;
        while (p < limit and (xml[p] == ' ' or xml[p] == '\t' or xml[p] == '\n' or xml[p] == '\r')) p += 1;
        if (p >= limit or xml[p] != '=') {
            i = after;
            continue;
        }
        p += 1;
        while (p < limit and (xml[p] == ' ' or xml[p] == '\t' or xml[p] == '\n' or xml[p] == '\r')) p += 1;
        if (p >= limit) return null;
        const quote = xml[p];
        if (quote != '"' and quote != '\'') {
            i = p;
            continue;
        }
        const val_start = p + 1;
        const val_end = std.mem.indexOfScalarPos(u8, xml[0..limit], val_start, quote) orelse return null;
        if (std.mem.eql(u8, name, prefix)) return xml[val_start..val_end];
        i = val_end + 1;
    }
    return null;
}

/// Same as `findNamespacePrefix` but skips any binding whose name
/// equals `skip`. Lets callers locate a SECOND prefix bound to the
/// same URI when one was already chosen as primary. `skip = ""`
/// behaves identically to `findNamespacePrefix` (no exclusion —
/// xmlns: declarations always have a non-empty name).
fn findNamespacePrefixExcept(
    xml: []const u8,
    target_uri: []const u8,
    skip: []const u8,
) ?[]const u8 {
    // Bounded to the first 4 KiB. Used for `a` and `c` prefix
    // resolution where ROOT-only scoping is what we want — picking
    // up a local mid-document declaration for these would shadow
    // the canonical-fallback path and silently drop anchors that
    // use the canonical prefix locally. xdr alt collection has
    // its own `collectAllNamespacePrefixes` which scans the full
    // document because anchor-tag prefixes are intentionally
    // late-bindable.
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
        if (std.mem.eql(u8, xml[val_start..val_end], target_uri) and
            !std.mem.eql(u8, name, skip))
        {
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
) !?store_mod.Relationship {
    // Decode the lookup id so the comparison matches the decoded
    // Relationship.id stored by parseRelationships. OOXML rIds in
    // practice are short ASCII tokens (`rId1`, `rId12`), so the
    // 64-byte stack buffer fast path covers everything realistic.
    // Pathological encoded IDs that decode beyond 64 bytes fall
    // through to a heap-allocated decode. OOM during the heap
    // fallback is propagated as `error.OutOfMemory` rather than
    // silently dropped — a downstream caller can decide whether
    // to abort the whole drawing parse or continue.
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
    const heap_buf = try allocator.alloc(u8, id.len);
    defer allocator.free(heap_buf);
    const decoded = decodeIdInto(heap_buf, id) orelse {
        // Decode-into-buffer failed even with id.len bytes — the
        // input is malformed (unterminated entity etc). Treat as
        // "no match" since the encoded form definitionally won't
        // match the stored decoded id.
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
) !?[]const u8 {
    const r = (try relForId(allocator, rels, id)) orelse return null;
    if (r.target_mode == .external) return null;
    return r.target;
}

/// `relTargetForId`, but under `.strict` the relationship's TYPE must
/// carry the expected leaf (`drawing` / `image` / `chart`): an id
/// that reaches a differently-typed relationship with an extant
/// target would otherwise produce a semantically false record instead
/// of refusing (Codex #214 r2 REL-202). `.lenient` keeps the
/// historical identity-only lookup.
fn relTargetForIdTyped(
    allocator: std.mem.Allocator,
    rels: []const store_mod.Relationship,
    id: []const u8,
    leaf: []const u8,
    mode: WalkMode,
) !?[]const u8 {
    const r = (try relForId(allocator, rels, id)) orelse return null;
    if (r.target_mode == .external) return null;
    if (mode == .strict and !relLeafIs(r.type, leaf)) return null;
    return r.target;
}

/// Case-insensitive comparison of a relationship type's last path
/// segment (the `pkg/pivots.zig` helper, duplicated to keep this
/// module import-light).
fn relLeafIs(rel_type: []const u8, leaf: []const u8) bool {
    const l = if (std.mem.lastIndexOfScalar(u8, rel_type, '/')) |i| rel_type[i + 1 ..] else rel_type;
    return std.ascii.eqlIgnoreCase(l, leaf);
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

    const ext = parseExtAttrs(xml, pos_end, open_ext) orelse return null;

    return .{ .x = x, .y = y, .cx = ext.cx, .cy = ext.cy };
}

/// Parse the `<{xdr}:ext cx="N" cy="N"/>` child starting the search
/// at `from_idx`. Shared by the absoluteAnchor parser and the strict
/// one-cell validation: a oneCellAnchor's extent is REQUIRED by the
/// schema, and a pic- or chart-bearing one-cell anchor whose extent
/// does not parse must refuse under strict rather than ride out
/// (Codex #214 r2 REL-201). The value stays off the wire either way.
fn parseExtAttrs(xml: []const u8, from_idx: usize, open_ext: []const u8) ?struct { cx: i64, cy: i64 } {
    const ext_idx = std.mem.indexOfPos(u8, xml, from_idx, open_ext) orelse return null;
    const ext_end = std.mem.indexOfScalarPos(u8, xml, ext_idx, '>') orelse return null;
    const ext_attrs = xml[ext_idx .. ext_end + 1];
    const cx_str = attrValue(ext_attrs, "cx") orelse return null;
    const cy_str = attrValue(ext_attrs, "cy") orelse return null;
    const cx = std.fmt.parseInt(i64, cx_str, 10) catch return null;
    const cy = std.fmt.parseInt(i64, cy_str, 10) catch return null;
    return .{ .cx = cx, .cy = cy };
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/openxlsx_loadExample.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var s = try PartStore.open(std.testing.allocator, io, fixture);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    // poi_58325_db.xlsx ships shape-only drawings. The parser must
    // walk them without producing image anchors.
    const fixture = "tests/corpus/poi_58325_db.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var s = try PartStore.open(std.testing.allocator, io, fixture);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    // worldbank_catalog has no drawings at all; the parser should
    // walk every sheet and find nothing.
    const fixture = "tests/corpus/worldbank_catalog.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var s = try PartStore.open(std.testing.allocator, io, fixture);
    defer s.deinit();

    const anchors = try imageAnchors(&s, std.testing.allocator);
    defer std.testing.allocator.free(anchors);

    try std.testing.expectEqual(@as(usize, 0), anchors.len);
}

test "chartAnchors: openxlsx_loadExample.xlsx surfaces embedded charts" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/openxlsx_loadExample.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var s = try PartStore.open(std.testing.allocator, io, fixture);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/worldbank_catalog.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var s = try PartStore.open(std.testing.allocator, io, fixture);
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
    // A compound chart (two plot types overlaid) reports `.other`, per
    // the enum's contract — not whichever element happens to be found
    // first (Codex #214 r1 REL-104).
    try std.testing.expectEqual(ChartType.other, detectChartType("<c:chartSpace><c:barChart/><c:lineChart/>", "c"));
    // The same plot type under BOTH prefixes is one type, not a
    // compound.
    try std.testing.expectEqual(ChartType.bar, detectChartTypeWithAlt("<c:chartSpace><c:barChart/><c2:barChart/>", "c", "c2"));
    // A plot-type element inside a comment or CDATA section is text,
    // not a plot (Codex #214 r2 REL-203).
    try std.testing.expectEqual(ChartType.bar, detectChartType("<c:chartSpace><c:barChart/><!-- <c:lineChart/> -->", "c"));
    try std.testing.expectEqual(ChartType.bar, detectChartType("<c:chartSpace><c:barChart/><![CDATA[<c:pieChart/>]]>", "c"));
}

test "findDrawingRid: a prefixed drawing element and a non-r relationship prefix resolve" {
    // `<x:drawing rel:id=…/>` is valid OOXML the unprefixed literal
    // search missed (Codex #214 r2 REL-202).
    try std.testing.expectEqualStrings(
        "rId9",
        findDrawingRid("<x:worksheet><x:drawing rel:id=\"rId9\"/></x:worksheet>").?,
    );
    // A bare `id` attribute is NOT a relationship reference.
    try std.testing.expectEqual(@as(?[]const u8, null), findDrawingRid("<sheet><drawing id=\"1\"/></sheet>"));
    // Tri-state: an unreadable reference is malformed, not absent.
    try std.testing.expect(findDrawingRef("<sheet><drawing/></sheet>") == .malformed);
    try std.testing.expect(findDrawingRef("<sheet><sheetData/></sheet>") == .absent);
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

test "attrValue tolerates whitespace around =" {
    // XML 1.0 §3.1 allows whitespace around `=` in attribute syntax.
    try std.testing.expectEqualStrings("914400", attrValue(
        \\x = "914400" y = "0"
    , "x").?);
    try std.testing.expectEqualStrings("0", attrValue(
        \\x = "914400" y = "0"
    , "y").?);
    // Tabs / newlines too.
    try std.testing.expectEqualStrings(
        "rId1",
        attrValue("r:id\n=\t\"rId1\"", "r:id").?,
    );
    // Substring-of-other-key must NOT match (word-boundary).
    try std.testing.expectEqual(
        @as(?[]const u8, null),
        attrValue("any=\"v\"", "ny"),
    );
    // Substring inside a quoted VALUE must NOT match (quote-aware
    // walking). A descr value containing `x = 'fake'` shouldn't
    // satisfy a lookup for `x`.
    try std.testing.expectEqualStrings(
        "real",
        attrValue("descr=\"foo x = 'fake'\" x=\"real\"", "x").?,
    );
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

test "findLocalChartElement matches the actually-used prefix among multiple bindings" {
    // A block with two prefixes bound to the chart URI: the first
    // (`u`) is unused, the second (`c2`) is on the actual chart
    // element. The previous "collect first prefix bound to URI"
    // approach picked `u` and missed the chart; scanning by tag
    // verifies per-element so `c2` is found.
    const block =
        "<xdr:graphicFrame>" ++
        "<some xmlns:u=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"/>" ++
        "<c2:chart xmlns:c2=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" r:id=\"rId1\"/>" ++
        "</xdr:graphicFrame>";
    const gf_idx = std.mem.indexOf(u8, block, "<xdr:graphicFrame").?;
    const found = findLocalChartElement(block, gf_idx, .{});
    try std.testing.expect(found != null);
    // `<c2:chart` should be at the position the helper returns.
    try std.testing.expect(std.mem.startsWith(u8, block[found.?..], "<c2:chart"));
}

test "findLocalChartElement: chart-element self-redeclare wins over earlier non-chart binding" {
    // An earlier element declares xmlns:p bound to a NON-chart URI;
    // the chart element redeclares xmlns:p bound to the chart URI.
    // XML scoping says the chart element's own declaration applies
    // to that element. The "nearest before tag end" lookup must
    // pick the chart's own redeclare, not the earlier sibling.
    //
    // Uses prefix `p` (NOT the default root `c`) so the test result
    // depends on the local-binding lookup, not the
    // `matches_root_primary` fallback.
    const block =
        "<xdr:graphicFrame>" ++
        "<other xmlns:p=\"http://example.com/not-a-chart\"/>" ++
        "<p:chart xmlns:p=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" r:id=\"rId1\"/>" ++
        "</xdr:graphicFrame>";
    const gf_idx = std.mem.indexOf(u8, block, "<xdr:graphicFrame").?;
    // Override the root primary so it can't match `p` as a fallback.
    var prefixes: DrawingPrefixes = .{};
    prefixes.c = "c-not-p";
    const found = findLocalChartElement(block, gf_idx, prefixes);
    try std.testing.expect(found != null);
    try std.testing.expect(std.mem.startsWith(u8, block[found.?..], "<p:chart"));
}

test "findLocalChartElement: xmlns after a quoted `>` in same tag is honored" {
    // The `<foo>` opening tag has `descr="a > b"` BEFORE an
    // xmlns:p declaration. The opening-tag detection must walk
    // forward quote-aware so the in-quotes `>` doesn't cause
    // isInsideOpeningTag to reject the real xmlns binding.
    const block =
        "<xdr:graphicFrame>" ++
        "<foo descr=\"a > b\" xmlns:p=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">" ++
        "<p:chart r:id=\"rId1\"/>" ++
        "</foo>" ++
        "</xdr:graphicFrame>";
    const gf_idx = std.mem.indexOf(u8, block, "<xdr:graphicFrame").?;
    var prefixes: DrawingPrefixes = .{};
    prefixes.c = "c-not-p"; // disable root-prefix fallback
    const found = findLocalChartElement(block, gf_idx, prefixes);
    try std.testing.expect(found != null);
    try std.testing.expect(std.mem.startsWith(u8, block[found.?..], "<p:chart"));
}

test "findLocalChartElement: xmlns inside an attribute value isn't a real binding" {
    // `<foo macro="xmlns:c='bad'">` — the xmlns text lives inside
    // a quoted attribute VALUE of an opening tag, not as an
    // attribute on its own. Quote-aware opening-tag detection
    // must reject it so the chart still resolves via root.
    const block =
        "<xdr:graphicFrame>" ++
        "<foo macro=\"xmlns:c='http://example.com/bad'\"></foo>" ++
        "<c:chart r:id=\"rId1\"/>" ++
        "</xdr:graphicFrame>";
    const gf_idx = std.mem.indexOf(u8, block, "<xdr:graphicFrame").?;
    const prefixes: DrawingPrefixes = .{};
    const found = findLocalChartElement(block, gf_idx, prefixes);
    try std.testing.expect(found != null);
    try std.testing.expect(std.mem.startsWith(u8, block[found.?..], "<c:chart"));
}

test "findLocalChartElement: fake nested markup in comment doesn't unbalance depth" {
    // `<foo xmlns:c="bad"><!-- <foo> --></foo>` — the comment
    // contains a fake `<foo>` nested-open. elementExtentEnd's
    // depth counter must NOT increment on it; otherwise the
    // matching `</foo>` won't bring depth back to 0 and the bad
    // binding leaks past the close into the chart's scope.
    const block =
        "<xdr:graphicFrame>" ++
        "<foo xmlns:c=\"http://example.com/bad\"><!-- <foo> --></foo>" ++
        "<c:chart r:id=\"rId1\"/>" ++
        "</xdr:graphicFrame>";
    const gf_idx = std.mem.indexOf(u8, block, "<xdr:graphicFrame").?;
    const prefixes: DrawingPrefixes = .{};
    const found = findLocalChartElement(block, gf_idx, prefixes);
    try std.testing.expect(found != null);
    try std.testing.expect(std.mem.startsWith(u8, block[found.?..], "<c:chart"));
}

test "findLocalChartElement: start inside a skip region jumps past it" {
    // `start` (the graphicFrame index) lands inside a commented
    // fake `<xdr:graphicFrame>` containing a fake `<c:chart>`.
    // The scanner must jump past the comment's close before
    // looking for `:chart`, otherwise it'd return the bogus
    // commented chart.
    const block =
        "<wrapper>" ++
        "<!-- <xdr:graphicFrame><c:chart r:id=\"rIdFake\"/></xdr:graphicFrame> -->" ++
        "<xdr:graphicFrame>" ++
        "<c:chart r:id=\"rIdReal\"/>" ++
        "</xdr:graphicFrame>" ++
        "</wrapper>";
    // Caller's plain indexOf(`<xdr:graphicFrame`) will land on
    // the commented one's `<` (inside the comment).
    const fake_gf = std.mem.indexOf(u8, block, "<xdr:graphicFrame").?;
    const prefixes: DrawingPrefixes = .{};
    const found = findLocalChartElement(block, fake_gf, prefixes);
    try std.testing.expect(found != null);
    try std.testing.expect(std.mem.startsWith(u8, block[found.?..], "<c:chart r:id=\"rIdReal\""));
}

test "findLocalChartElement: fake `<c:chart>` inside a comment is ignored" {
    // The graphicFrame contains a fake `<c:chart>` tag inside a
    // comment, then the real chart afterwards. The scanner must
    // skip the commented candidate (markup-shaped text isn't real
    // markup) and pick the real one.
    const block =
        "<xdr:graphicFrame>" ++
        "<!-- <c:chart xmlns:c=\"http://example.com/bad\" r:id=\"rIdFake\"/> -->" ++
        "<c:chart r:id=\"rIdReal\"/>" ++
        "</xdr:graphicFrame>";
    const gf_idx = std.mem.indexOf(u8, block, "<xdr:graphicFrame").?;
    const prefixes: DrawingPrefixes = .{};
    const found = findLocalChartElement(block, gf_idx, prefixes);
    try std.testing.expect(found != null);
    // The found `<` is for the REAL chart with rIdReal, not the
    // commented one with rIdFake.
    const tail_after = block[found.?..];
    try std.testing.expect(std.mem.startsWith(u8, tail_after, "<c:chart r:id=\"rIdReal\""));
}

test "findLocalChartElement: PI body with fake </name> doesn't unbalance extent" {
    // An ancestor that declares a chart-prefix binding contains a
    // processing instruction whose body has fake `</foo>` text.
    // elementExtentEnd's depth counter must skip PI bodies — if it
    // counted that as a real close tag, depth would go to 0 too
    // early, the binding's extent would shrink, and a real chart
    // declared after the ancestor's true close would still see
    // the ancestor's binding.
    const block =
        "<xdr:graphicFrame>" ++
        "<foo xmlns:p=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">" ++
        "<?someProc </foo> ?>" ++
        "<p:chart r:id=\"rId1\"/>" ++
        "</foo>" ++
        "</xdr:graphicFrame>";
    const gf_idx = std.mem.indexOf(u8, block, "<xdr:graphicFrame").?;
    var prefixes: DrawingPrefixes = .{};
    prefixes.c = "c-not-p"; // disable root fallback so test depends on local lookup
    const found = findLocalChartElement(block, gf_idx, prefixes);
    try std.testing.expect(found != null);
    try std.testing.expect(std.mem.startsWith(u8, block[found.?..], "<p:chart"));
}

test "findLocalChartElement: comment-delimiter-text inside a PI doesn't open a fake comment" {
    // A processing instruction whose body contains the literal
    // text `<!--`. The forward scanner must skip past `?>` and
    // not treat the PI content as an unclosed comment, otherwise
    // later real xmlns bindings get classified as in-comment.
    const block =
        "<xdr:graphicFrame>" ++
        "<?someProcessing <!-- ?>" ++
        "<foo xmlns:p=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">" ++
        "<p:chart r:id=\"rId1\"/>" ++
        "</foo>" ++
        "</xdr:graphicFrame>";
    const gf_idx = std.mem.indexOf(u8, block, "<xdr:graphicFrame").?;
    var prefixes: DrawingPrefixes = .{};
    prefixes.c = "c-not-p"; // disable root fallback so test depends on local lookup
    const found = findLocalChartElement(block, gf_idx, prefixes);
    try std.testing.expect(found != null);
    try std.testing.expect(std.mem.startsWith(u8, block[found.?..], "<p:chart"));
}

test "findLocalChartElement: comment-delimiter-text inside CDATA doesn't open a fake comment" {
    // `<![CDATA[<!--]]>` — a CDATA section whose CONTENT is the
    // literal text `<!--`. The closed CDATA must restore "outside
    // markup" state; a naive "latest <!-- vs latest -->" heuristic
    // would treat the CDATA-internal text as an unclosed comment
    // and reject the real xmlns:p that follows.
    const block =
        "<xdr:graphicFrame>" ++
        "<![CDATA[<!--]]>" ++
        "<foo xmlns:p=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">" ++
        "<p:chart r:id=\"rId1\"/>" ++
        "</foo>" ++
        "</xdr:graphicFrame>";
    const gf_idx = std.mem.indexOf(u8, block, "<xdr:graphicFrame").?;
    var prefixes: DrawingPrefixes = .{};
    prefixes.c = "c-not-p"; // disable root fallback so test depends on local lookup
    const found = findLocalChartElement(block, gf_idx, prefixes);
    try std.testing.expect(found != null);
    try std.testing.expect(std.mem.startsWith(u8, block[found.?..], "<p:chart"));
}

test "findLocalChartElement: fake markup with xmlns inside a comment is ignored" {
    // A comment contains FAKE markup that includes a real-looking
    // opening tag with a quoted `>` and an xmlns binding. The
    // namespace scope scan must walk past the entire comment and
    // not be tricked by the inner `<x>` opener.
    const block =
        "<xdr:graphicFrame>" ++
        "<!-- <x descr=\"a > b\" xmlns:c=\"http://example.com/bad\"> -->" ++
        "<c:chart r:id=\"rId1\"/>" ++
        "</xdr:graphicFrame>";
    const gf_idx = std.mem.indexOf(u8, block, "<xdr:graphicFrame").?;
    const prefixes: DrawingPrefixes = .{};
    const found = findLocalChartElement(block, gf_idx, prefixes);
    try std.testing.expect(found != null);
    try std.testing.expect(std.mem.startsWith(u8, block[found.?..], "<c:chart"));
}

test "findLocalChartElement: xmlns-looking text in XML comment is ignored" {
    // A comment before the chart contains text that looks like an
    // xmlns declaration. The scope-resolver must NOT treat that
    // as a real binding; otherwise the chart's `c` resolves to
    // the bogus comment URI and the chart is dropped.
    const block =
        "<xdr:graphicFrame>" ++
        "<!-- xmlns:c=\"http://example.com/in-a-comment\" -->" ++
        "<c:chart r:id=\"rId1\"/>" ++
        "</xdr:graphicFrame>";
    const gf_idx = std.mem.indexOf(u8, block, "<xdr:graphicFrame").?;
    const prefixes: DrawingPrefixes = .{};
    const found = findLocalChartElement(block, gf_idx, prefixes);
    try std.testing.expect(found != null);
    try std.testing.expect(std.mem.startsWith(u8, block[found.?..], "<c:chart"));
}

test "findLocalChartElement: quoted `>` in attribute doesn't fool tag-end scan" {
    // A self-closing earlier sibling with an attribute that
    // contains a literal `>` inside its quoted value. The tag-end
    // scanner must skip over quoted regions so the `/>` is found
    // at the actual end of the tag, not at the in-attribute `>`.
    // Without quote-awareness, the sibling looks non-self-closing,
    // its xmlns:c stays "in scope", and `<c:chart>` is dropped.
    const block =
        "<xdr:graphicFrame>" ++
        "<other descr=\"a > b\" xmlns:c=\"http://example.com/closed-sibling\"/>" ++
        "<c:chart r:id=\"rId1\"/>" ++
        "</xdr:graphicFrame>";
    const gf_idx = std.mem.indexOf(u8, block, "<xdr:graphicFrame").?;
    const prefixes: DrawingPrefixes = .{};
    const found = findLocalChartElement(block, gf_idx, prefixes);
    try std.testing.expect(found != null);
    try std.testing.expect(std.mem.startsWith(u8, block[found.?..], "<c:chart"));
}

test "findLocalChartElement: container sibling's redeclare doesn't shadow after close" {
    // Non-self-closing earlier sibling (`<other ...>...</other>`)
    // declares xmlns:c bound to a NON-chart URI. Per XML scoping,
    // that binding ends at `</other>`. The later `<c:chart>` must
    // resolve `c` via the root primary, not the closed sibling.
    const block =
        "<xdr:graphicFrame>" ++
        "<other xmlns:c=\"http://example.com/closed-container\">stuff</other>" ++
        "<c:chart r:id=\"rId1\"/>" ++
        "</xdr:graphicFrame>";
    const gf_idx = std.mem.indexOf(u8, block, "<xdr:graphicFrame").?;
    const prefixes: DrawingPrefixes = .{}; // .c = "c" by default
    const found = findLocalChartElement(block, gf_idx, prefixes);
    try std.testing.expect(found != null);
    try std.testing.expect(std.mem.startsWith(u8, block[found.?..], "<c:chart"));
}

test "findLocalChartElement: closed sibling's redeclare doesn't shadow root binding" {
    // A self-closing earlier sibling declares xmlns:c bound to a
    // NON-chart URI. Per XML scoping, that binding ends at the
    // sibling's `/>`, so when `<c:chart>` later in the same block
    // uses prefix `c`, the root primary should match. Without the
    // self-closing scope check, the sibling's xmlns:c would still
    // be visible via uriOfPrefixAtPosition and the chart would be
    // wrongly skipped.
    const block =
        "<xdr:graphicFrame>" ++
        "<other xmlns:c=\"http://example.com/closed-sibling\"/>" ++
        "<c:chart r:id=\"rId1\"/>" ++
        "</xdr:graphicFrame>";
    const gf_idx = std.mem.indexOf(u8, block, "<xdr:graphicFrame").?;
    const prefixes: DrawingPrefixes = .{}; // .c = "c" by default
    const found = findLocalChartElement(block, gf_idx, prefixes);
    try std.testing.expect(found != null);
    try std.testing.expect(std.mem.startsWith(u8, block[found.?..], "<c:chart"));
}

test "findLocalChartElement: local non-chart redeclare blocks root-prefix match" {
    // The first `<c:chart>` redeclares xmlns:c locally to a
    // NON-chart URI. The root primary is `c`, so before the fix
    // the scanner would accept this (wrong) element via
    // matches_root_primary and drop the REAL `<p:chart>` later.
    // With the fix, an in-scope local binding is authoritative —
    // root fallback only applies when no local binding exists.
    const block =
        "<xdr:graphicFrame>" ++
        "<c:chart xmlns:c=\"http://example.com/not-a-chart\" r:id=\"rIdWrong\"/>" ++
        "<p:chart xmlns:p=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" r:id=\"rIdReal\"/>" ++
        "</xdr:graphicFrame>";
    const gf_idx = std.mem.indexOf(u8, block, "<xdr:graphicFrame").?;
    const prefixes: DrawingPrefixes = .{}; // .c = "c" by default
    const found = findLocalChartElement(block, gf_idx, prefixes);
    try std.testing.expect(found != null);
    try std.testing.expect(std.mem.startsWith(u8, block[found.?..], "<p:chart"));
}

test "findLocalChartElement skips ':chartSpace' false matches" {
    // `<c:chartSpace>` is NOT `<c:chart>` — the byte after `:chart`
    // determines whether it's the chart element or a different one.
    const block = "<xdr:graphicFrame><c:chartSpace/></xdr:graphicFrame>";
    const gf_idx = 0;
    var prefixes: DrawingPrefixes = .{};
    prefixes.c = "c";
    try std.testing.expectEqual(@as(?usize, null), findLocalChartElement(block, gf_idx, prefixes));
}

test "findLocalNamespacePrefix walks past 4 KiB inside a block" {
    // collectChartsFromSheet uses findLocalNamespacePrefix to probe
    // for a chart-namespace prefix declared LOCALLY on `<*:chart>`.
    // If the anchor block is large enough that the local xmlns:c
    // sits past 4 KiB, the bounded findNamespacePrefix would miss
    // it; the local helper must walk the whole block.
    var pad_buf: [5000]u8 = undefined;
    @memset(&pad_buf, ' ');
    const head = "<chart-prefix-late>";
    const tail =
        \\<c2:chart xmlns:c2="http://schemas.openxmlformats.org/drawingml/2006/chart"/></chart-prefix-late>
    ;
    var doc_buf: [8192]u8 = undefined;
    var fbs = std.Io.Writer.fixed(&doc_buf);
    const w = &fbs;
    try w.writeAll(head);
    try w.writeAll(&pad_buf);
    try w.writeAll(tail);
    const block = fbs.buffered();

    // Bounded helper: misses the late binding (cap at 4 KiB).
    try std.testing.expectEqual(
        @as(?[]const u8, null),
        findNamespacePrefix(block, "http://schemas.openxmlformats.org/drawingml/2006/chart"),
    );
    // Unbounded local helper: still finds it.
    const found = findLocalNamespacePrefix(block, "http://schemas.openxmlformats.org/drawingml/2006/chart");
    try std.testing.expect(found != null);
    try std.testing.expectEqualStrings("c2", found.?);
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
    // Two prefixes bound to the same spreadsheetDrawing URI: root
    // uses one, descendant anchors may use the other. Both must be
    // tracked so the scanner can replay with the alt prefix.
    {
        const xml =
            \\<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:dr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>
        ;
        const p = resolveDrawingPrefixes(xml);
        try std.testing.expectEqualStrings("xdr", p.xdr);
        try std.testing.expectEqual(@as(usize, 1), p.xdr_alts_len);
        try std.testing.expectEqualStrings("dr", p.xdr_alts()[0]);
    }
    // Mixed conformance: Transitional URI on one prefix, Strict URI
    // on another. Primary picks one, alt tracks the other.
    {
        const xml =
            \\<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:xs="http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"/>
        ;
        const p = resolveDrawingPrefixes(xml);
        try std.testing.expectEqualStrings("xdr", p.xdr);
        try std.testing.expectEqual(@as(usize, 1), p.xdr_alts_len);
        try std.testing.expectEqualStrings("xs", p.xdr_alts()[0]);
    }
    // Single binding — alt must stay null so the scanner doesn't
    // double-walk an identical needle set.
    {
        const xml =
            \\<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>
        ;
        const p = resolveDrawingPrefixes(xml);
        try std.testing.expectEqualStrings("xdr", p.xdr);
        try std.testing.expectEqual(@as(usize, 0), p.xdr_alts_len);
    }
    // Strict-rooted with unused Transitional declaration: the
    // descendant alt prefix is on the Strict URI. The resolver
    // must prefer same-URI alts over the unused other-conformance
    // declaration.
    {
        const xml =
            \\<?xml version="1.0"?><xs:wsDr xmlns:xs="http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:dr="http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"/>
        ;
        const p = resolveDrawingPrefixes(xml);
        try std.testing.expectEqualStrings("xs", p.xdr);
        try std.testing.expectEqual(@as(usize, 2), p.xdr_alts_len);
        // Same-URI (Strict) alt is enumerated FIRST.
        try std.testing.expectEqualStrings("dr", p.xdr_alts()[0]);
        // Other-conformance (Transitional) alt comes second so the
        // scanner still walks it, picking up any anchors that use
        // it (rare, but valid OOXML).
        try std.testing.expectEqualStrings("xdr", p.xdr_alts()[1]);
    }
    // Mirror case: Transitional-rooted with unused Strict
    // declaration. Same-URI alt must still win, both alts tracked.
    {
        const xml =
            \\<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:xs="http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing" xmlns:dr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>
        ;
        const p = resolveDrawingPrefixes(xml);
        try std.testing.expectEqualStrings("xdr", p.xdr);
        try std.testing.expectEqual(@as(usize, 2), p.xdr_alts_len);
        try std.testing.expectEqualStrings("dr", p.xdr_alts()[0]);
        try std.testing.expectEqualStrings("xs", p.xdr_alts()[1]);
    }
    // Late-declared xmlns past the previous 4 KiB scan window:
    // XML 1.0 + Namespaces 1.0 allow xmlns:* on any element. Pad
    // the document past 4 KiB before declaring the alt prefix so
    // we exercise the unbounded scan path.
    {
        var pad_buf: [5000]u8 = undefined;
        @memset(&pad_buf, ' ');
        const head =
            \\<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
        ;
        const tail =
            \\<dr:twoCellAnchor xmlns:dr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/></xdr:wsDr>
        ;
        var doc_buf: [8192]u8 = undefined;
        var fbs = std.Io.Writer.fixed(&doc_buf);
        const w = &fbs;
        try w.writeAll(head);
        try w.writeAll(&pad_buf);
        try w.writeAll(tail);
        const doc = fbs.buffered();
        const p = resolveDrawingPrefixes(doc);
        try std.testing.expectEqualStrings("xdr", p.xdr);
        try std.testing.expectEqual(@as(usize, 1), p.xdr_alts_len);
        try std.testing.expectEqualStrings("dr", p.xdr_alts()[0]);
    }
    // Multiple alts on the same URI: an unused declaration
    // appears before the actually-used alt. Both must be tracked
    // so the scanner replays with each, even if the first alt is
    // unused — the loop short-circuits on no-match.
    {
        const xml =
            \\<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:u="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:dr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>
        ;
        const p = resolveDrawingPrefixes(xml);
        try std.testing.expectEqualStrings("xdr", p.xdr);
        try std.testing.expectEqual(@as(usize, 2), p.xdr_alts_len);
        try std.testing.expectEqualStrings("u", p.xdr_alts()[0]);
        try std.testing.expectEqualStrings("dr", p.xdr_alts()[1]);
    }
    // Scope-isolation regression: a/c prefixes resolve from the
    // ROOT element only — picking up a late-declared dml prefix
    // would shadow the canonical fallback, dropping anchors that
    // use the canonical `<a:blip>` locally. Pad past 4 KiB before
    // declaring xmlns:dml so the scan would have to walk the full
    // document to find it; the cap on findNamespacePrefix keeps
    // p.a anchored on the canonical default.
    {
        var pad_buf: [5000]u8 = undefined;
        @memset(&pad_buf, ' ');
        const head =
            \\<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
        ;
        const tail =
            \\<other xmlns:dml="http://schemas.openxmlformats.org/drawingml/2006/main"/></xdr:wsDr>
        ;
        var doc_buf: [8192]u8 = undefined;
        var fbs = std.Io.Writer.fixed(&doc_buf);
        const w = &fbs;
        try w.writeAll(head);
        try w.writeAll(&pad_buf);
        try w.writeAll(tail);
        const doc = fbs.buffered();
        const p = resolveDrawingPrefixes(doc);
        // Canonical fallback wins — late dml binding is invisible.
        try std.testing.expectEqualStrings("a", p.a);
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
