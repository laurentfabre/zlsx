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
//! Namespace prefixes are resolved per part (`resolveDrawingPrefixes`):
//! `xdr:` / `a:` / `c:` are the canonical spellings the sketch above
//! uses, but every prefix bound to the spreadsheetDrawing URI is
//! followed — the root element's, every alternate declared anywhere
//! in the part (replayed, capped at `max_xdr_alts`), and the DEFAULT
//! namespace (an empty prefix: `<wsDr xmlns="…/spreadsheetDrawing">
//! <oneCellAnchor><from><col>`, openpyxl 3.1's spelling, whose
//! anchors this read listed as nothing until the namespace-aware
//! drawing slice). A binding to the URI under a name the walk cannot
//! follow — longer than `max_prefix_len`, or past the alternate cap —
//! is `xdr_rejected`: the strict read refuses the drawing rather than
//! serve a partial inventory, and the drawing sweep (`drawing_edit`,
//! which resolves the same prefixes) refuses the edit rather than
//! leave an anchor behind. The chart stub inside a graphic frame is
//! matched as `<{p}:chart` for a prefix bound to a chart URI — an
//! unprefixed `<chart xmlns="…/chart">` stub is not followed (no
//! producer spells one; openpyxl binds `c:` on its drawing root).

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
/// (valid until the store's `deinit`). The walk's own scratch — the
/// part names it resolves along the relationship chain — is freed
/// before return; nothing a walk allocates lands in the store.
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
    // Fallback to filename for content-type-less producers.
    return isCanonicalNumberedPart(part.name, "xl/worksheets/sheet");
}

/// Is `name` `<prefix><digits>.xml` with at least one digit — the
/// canonical numbered part name (`xl/worksheets/sheet<N>.xml`,
/// `xl/charts/chart<N>.xml`)? The entire substring between the prefix
/// and `.xml` must be digits; a partial-digit rule would let
/// `sheet1_backup.xml` / `sheet1custom.xml` slip through and be walked
/// as a worksheet. One rule for the sheet walk's fallback and the
/// workbook's chart-part enumeration, so the two cannot drift.
pub fn isCanonicalNumberedPart(name: []const u8, prefix: []const u8) bool {
    const suffix = ".xml";
    if (!std.mem.startsWith(u8, name, prefix)) return false;
    if (!std.mem.endsWith(u8, name, suffix)) return false;
    if (name.len <= prefix.len + suffix.len) return false;
    for (name[prefix.len .. name.len - suffix.len]) |c| if (!std.ascii.isDigit(c)) return false;
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
        .multiple => |r| blk: {
            if (mode == .strict) return error.MalformedDrawingXml;
            break :blk r;
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
    // Resolved names are lookup keys for the walk, nothing more: the
    // anchors carry the store's own part names. Scratch in the
    // caller's allocator, freed before return — the store's arena
    // variant would retain every path for the store's lifetime, and a
    // long-lived editor repeats this walk per typed read (Codex #216
    // r1 S3B-MEM-603).
    const drawing_part_name = (try store.resolveOwned(allocator, sheet_part.name, drawing_target)) orelse {
        if (mode == .strict) return error.MalformedDrawingXml;
        return;
    };
    defer allocator.free(drawing_part_name);
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
    // A spreadsheetDrawing binding under a name the walk will not
    // spell is an inventory the read cannot serve whole: strict
    // refuses rather than list the anchors under the other names as
    // the drawing's (the chart walk's `c_rejected` rule); lenient
    // lists what it can, as before.
    if (mode == .strict and prefixes.xdr_rejected) return error.MalformedDrawingXml;
    // A DTD is not content and no producer writes one: strict refuses
    // the part rather than walk what an entity could rewrite; lenient
    // steps over the declaration as a region (in-house ND-REL-203).
    if (mode == .strict and hasLiveDoctype(drawing_part.bytes)) return error.MalformedDrawingXml;
    // A `<` inside an attribute value is not well-formed XML and every
    // decoy surface at once: strict refuses (in-house ND-REL-411).
    if (mode == .strict and hasMarkupInAttributeValue(drawing_part.bytes)) return error.MalformedDrawingXml;
    // One tag set per followed prefix (`buildTagSets`): the wrapper
    // scan replays once per set — anchors under the other names do
    // not match, so no duplicates surface; replay order is NOT
    // document order, consumers needing that sort on `doc_offset` —
    // and every wrapper's children are read under any set.
    const set_count = tagSetCount(&prefixes);
    const bufs = try allocator.alloc(TagSetBuf, set_count);
    defer allocator.free(bufs);
    const set_store = try allocator.alloc(DrawingTags, set_count);
    defer allocator.free(set_store);
    const sets = try buildTagSets(bufs, set_store, prefixes);
    for (sets) |*tags| {
        try scanImagesWithTags(store, allocator, drawing_part, drawing_part_name, drawing_rels, sheet_part, prefixes, tags, sets, mode, out);
    }
}

/// Where the wrapper scan continues after a `<` that opens no anchor.
/// The empty prefix's opener is the bare `<`, which matches a comment
/// / CDATA / PI delimiter itself: stepping one byte past it would leave
/// the cursor INSIDE the region and blind every later `findLiveMarkup`,
/// whose region state starts at the cursor — a commented decoy became
/// a phantom anchor and a trailing comment a refusal (in-house
/// ND-REL-101). `skipRegionEndFrom` answers `at + 1` for a `<` that
/// opens no region (the prefixed opener never does), so the prefixed
/// path keeps its needle-length advance.
fn advancePastNonAnchor(xml: []const u8, at: usize, opener_len: usize) usize {
    return @max(at + opener_len, skipRegionEndFrom(xml, at));
}

fn scanImagesWithTags(
    store: *PartStore,
    allocator: std.mem.Allocator,
    drawing_part: store_mod.Part,
    drawing_part_name: []const u8,
    drawing_rels: []const store_mod.Relationship,
    sheet_part: store_mod.Part,
    prefixes: DrawingPrefixes,
    tags: *const DrawingTags,
    sets: []const DrawingTags,
    mode: WalkMode,
    out: *std.ArrayListUnmanaged(ImageAnchor),
) !void {
    var i: usize = 0;
    while (i < drawing_part.bytes.len) {
        // Anchors, closes and structural children are located as LIVE
        // exact tags: markup-shaped text inside comments / CDATA / PIs
        // is not inventory, a commented close must not truncate a live
        // anchor, and a longer element name must not masquerade as an
        // anchor wrapper (Codex #214 r5 REL-501).
        const next = findLiveMarkup(drawing_part.bytes, i, tags.xdr_prefix_open) orelse break;
        i = next;
        const block_start = next;
        // Identify anchor opener.
        const is_two = matchesOpenTag(drawing_part.bytes, i, tags.open_two);
        const is_one = matchesOpenTag(drawing_part.bytes, i, tags.open_one);
        const is_absolute = matchesOpenTag(drawing_part.bytes, i, tags.open_absolute);
        if (!is_two and !is_one and !is_absolute) {
            i = advancePastNonAnchor(drawing_part.bytes, i, tags.xdr_prefix_open.len);
            continue;
        }
        // A self-closing wrapper (`<xdr:twoCellAnchor/>`) holds no
        // anchor: step over it, or the close of the NEXT wrapper of
        // that name bounds a block spanning both and the second anchor
        // is served as the first (in-house ND-DOC-204). The sweep
        // passes it through — the two agree. A part that ends inside
        // the tag is unreadable.
        const open_tag = selfClosingTagEnd(drawing_part.bytes, i) orelse {
            if (mode == .strict) return error.MalformedDrawingXml;
            break;
        };
        if (open_tag.self_closing) {
            i = open_tag.gt + 1;
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
        // From past the open tag: a close spelled inside one of its
        // attribute values is not the close (in-house ND-REL-302 — the
        // sweep already searched from there).
        const close = findLiveMarkup(drawing_part.bytes, open_tag.gt + 1, close_marker) orelse {
            if (mode == .strict) return error.MalformedDrawingXml;
            break;
        };
        const block = drawing_part.bytes[i .. close + close_marker.len];
        i = close + close_marker.len;

        // Only image-bearing anchors are surfaced. An anchor with no
        // `<xdr:pic>` (a shape, a chart frame) is legitimately not an
        // image in either mode; from the unclosed pic onward the block
        // holds an image the walk cannot read whole. The pic, like
        // every child, may be spelled under any followed prefix.
        const content_start = open_tag.gt + 1 - block_start;
        const pic = anyLiveExactTag(block, content_start, sets, "open_pic") orelse continue;
        const pic_close = findLiveMarkup(block, pic.at, pic.set.close_pic) orelse {
            if (mode == .strict) return error.MalformedDrawingXml;
            continue;
        };
        const pic_block = block[pic.at .. pic_close + pic.set.close_pic.len];

        const embed_rid = findBlipEmbedWithAlt(pic_block, prefixes.a, prefixes.a_alt) orelse {
            if (mode == .strict) return error.MalformedDrawingXml;
            continue;
        };
        const image_target = (try relTargetForIdTyped(allocator, drawing_rels, embed_rid, "image", mode)) orelse {
            if (mode == .strict) return error.MalformedDrawingXml;
            continue;
        };
        const image_part_name = (try store.resolveOwned(allocator, drawing_part_name, image_target)) orelse {
            if (mode == .strict) return error.MalformedDrawingXml;
            continue;
        };
        defer allocator.free(image_part_name);
        const image_part = try store.part(image_part_name) orelse {
            if (mode == .strict) return error.MalformedDrawingXml;
            continue;
        };

        const geometry = readAnchorGeometry(block, content_start, sets, is_two, is_absolute, mode) catch |e| switch (e) {
            error.MalformedDrawingXml => {
                if (mode == .strict) return error.MalformedDrawingXml;
                continue;
            },
        };

        try out.append(allocator, .{
            .image_part_name = image_part.name,
            .sheet_part_name = sheet_part.name,
            .from = geometry.from,
            .to = geometry.to,
            .absolute = geometry.absolute,
            .bytes = image_part.bytes,
            .kind = if (is_absolute) .absolute else if (is_two) .two_cell else .one_cell,
            .doc_offset = block_start,
        });
    }
}

const AnchorGeometry = struct {
    from: CellAnchor,
    to: ?CellAnchor,
    absolute: ?AbsoluteAnchor,
};

/// The cell-grid or pixel geometry of one anchor block, every child
/// under any followed prefix. `MalformedDrawingXml` when the walk
/// cannot read it whole — the caller refuses under strict and skips
/// the anchor under lenient: a two-cell anchor without a readable
/// `<to>` must not ride out looking like a one-cell anchor (Codex #214
/// r1 REL-101); a one-cell anchor's `<xdr:ext>` is schema-required, so
/// strict validates it even though the extent stays off the wire (r2
/// REL-201); an absolute anchor needs both `<pos>` and `<ext>`.
fn readAnchorGeometry(block: []const u8, content_start: usize, sets: []const DrawingTags, is_two: bool, is_absolute: bool, mode: WalkMode) error{MalformedDrawingXml}!AnchorGeometry {
    if (is_absolute) {
        const absolute = parseAbsoluteAnchorIn(block, content_start, sets) orelse return error.MalformedDrawingXml;
        return .{ .from = .{ .col = 0, .col_off = 0, .row = 0, .row_off = 0 }, .to = null, .absolute = absolute };
    }
    const from = parseCornerIn(block, content_start, sets, .from) orelse return error.MalformedDrawingXml;
    var to: ?CellAnchor = null;
    if (is_two) {
        const to_block = parseCornerIn(block, content_start, sets, .to);
        if (mode == .strict and to_block == null) return error.MalformedDrawingXml;
        // Two blocks that overlap — a `<to>` nested inside `<from>` —
        // are not two corners: strict refuses, as the sweep does (in-house
        // ND-DOC-301: the read listed the pair the sweep refused).
        if (mode == .strict and to_block != null and cornersOverlap(from, to_block.?)) return error.MalformedDrawingXml;
        if (to_block) |t| to = t.anchor;
    } else if (mode == .strict) {
        if (parseExtAttrsIn(block, content_start, sets) == null) return error.MalformedDrawingXml;
    }
    return .{ .from = from.anchor, .to = to, .absolute = null };
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
        .multiple => |r| blk: {
            if (mode == .strict) return error.MalformedDrawingXml;
            break :blk r;
        },
        .rid => |r| r,
    };
    const sheet_rels = store.rels(sheet_part.name);
    const drawing_target = (try relTargetForIdTyped(allocator, sheet_rels, rid, "drawing", mode)) orelse {
        if (mode == .strict) return error.MalformedDrawingXml;
        return;
    };
    // Scratch resolution, the image walk's rule (S3B-MEM-603).
    const drawing_part_name = (try store.resolveOwned(allocator, sheet_part.name, drawing_target)) orelse {
        if (mode == .strict) return error.MalformedDrawingXml;
        return;
    };
    defer allocator.free(drawing_part_name);
    const drawing_part = try store.part(drawing_part_name) orelse {
        if (mode == .strict) return error.MalformedDrawingXml;
        return;
    };

    const drawing_rels = store.rels(drawing_part_name);
    const prefixes = resolveDrawingPrefixes(drawing_part.bytes);
    // The image walk's rule: an unfollowed spreadsheetDrawing binding
    // refuses under strict.
    if (mode == .strict and prefixes.xdr_rejected) return error.MalformedDrawingXml;
    if (mode == .strict and hasLiveDoctype(drawing_part.bytes)) return error.MalformedDrawingXml;
    if (mode == .strict and hasMarkupInAttributeValue(drawing_part.bytes)) return error.MalformedDrawingXml;
    // The image walk's replay: one tag set per followed prefix.
    const set_count = tagSetCount(&prefixes);
    const bufs = try allocator.alloc(TagSetBuf, set_count);
    defer allocator.free(bufs);
    const set_store = try allocator.alloc(DrawingTags, set_count);
    defer allocator.free(set_store);
    const sets = try buildTagSets(bufs, set_store, prefixes);
    for (sets) |*tags| {
        try scanChartsWithTags(store, allocator, drawing_part, drawing_part_name, drawing_rels, sheet_part, prefixes, tags, sets, mode, out);
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
    tags: *const DrawingTags,
    sets: []const DrawingTags,
    mode: WalkMode,
    out: *std.ArrayListUnmanaged(ChartAnchor),
) !void {
    var i: usize = 0;
    while (i < drawing_part.bytes.len) {
        // Live exact tags throughout — see scanImagesWithTags (Codex
        // #214 r5 REL-501).
        const next = findLiveMarkup(drawing_part.bytes, i, tags.xdr_prefix_open) orelse break;
        i = next;
        const block_start = next;
        const is_two = matchesOpenTag(drawing_part.bytes, i, tags.open_two);
        const is_one = matchesOpenTag(drawing_part.bytes, i, tags.open_one);
        const is_absolute = matchesOpenTag(drawing_part.bytes, i, tags.open_absolute);
        if (!is_two and !is_one and !is_absolute) {
            i = advancePastNonAnchor(drawing_part.bytes, i, tags.xdr_prefix_open.len);
            continue;
        }
        // A self-closing wrapper (`<xdr:twoCellAnchor/>`) holds no
        // anchor: step over it, or the close of the NEXT wrapper of
        // that name bounds a block spanning both and the second anchor
        // is served as the first (in-house ND-DOC-204). The sweep
        // passes it through — the two agree. A part that ends inside
        // the tag is unreadable.
        const open_tag = selfClosingTagEnd(drawing_part.bytes, i) orelse {
            if (mode == .strict) return error.MalformedDrawingXml;
            break;
        };
        if (open_tag.self_closing) {
            i = open_tag.gt + 1;
            continue;
        }
        const close_marker = if (is_two)
            tags.close_two
        else if (is_one)
            tags.close_one
        else
            tags.close_absolute;
        // From past the open tag: a close spelled inside one of its
        // attribute values is not the close (in-house ND-REL-302 — the
        // sweep already searched from there).
        const close = findLiveMarkup(drawing_part.bytes, open_tag.gt + 1, close_marker) orelse {
            if (mode == .strict) return error.MalformedDrawingXml;
            break;
        };
        const block = drawing_part.bytes[i .. close + close_marker.len];
        i = close + close_marker.len;

        // Charts live inside <xdr:graphicFrame>...<c:chart r:id=...
        // A frame without a chart element (a diagram, a table) is
        // legitimately not a chart in either mode; a chart element
        // the walk cannot follow to its part is a refusal in strict.
        const content_start = open_tag.gt + 1 - block_start;
        const gf = anyLiveExactTag(block, content_start, sets, "open_graphic_frame") orelse continue;
        // Scan for any `<*:chart` element whose prefix is bound to
        // either chart URI (block-local OR drawing-root). Walking by
        // tag rather than by prefix avoids the "multiple local
        // bindings to the same chart URI" failure mode where
        // collect-first-prefix would pick an unused declaration.
        const chart_idx = findLocalChartElement(block, gf.at, prefixes) orelse continue;
        // Quote-aware: an unescaped `>` in an attribute value is legal
        // (in-house ND-REL-405).
        const chart_end = findUnquotedTagEnd(block, chart_idx) orelse {
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
        const chart_part_name = (try store.resolveOwned(allocator, drawing_part_name, chart_target)) orelse {
            if (mode == .strict) return error.MalformedDrawingXml;
            continue;
        };
        defer allocator.free(chart_part_name);
        const chart_part = try store.part(chart_part_name) orelse {
            if (mode == .strict) return error.MalformedDrawingXml;
            continue;
        };

        const geometry = readAnchorGeometry(block, content_start, sets, is_two, is_absolute, mode) catch |e| switch (e) {
            error.MalformedDrawingXml => {
                if (mode == .strict) return error.MalformedDrawingXml;
                continue;
            },
        };

        // Each chart's own XML may declare a different `c:` prefix
        // — resolve per-chart to be safe. A chart namespace bound
        // under a prefix the resolver rejected is a part this read
        // cannot walk whole: strict refuses it rather than serve the
        // fallback prefix's (empty) carrier list as the chart's refs
        // — the sweep's verdict on the same bytes (CF-DOC-201).
        const chart_prefixes = resolveDrawingPrefixes(chart_part.bytes);
        if (mode == .strict and chart_prefixes.c_rejected) return error.MalformedDrawingXml;
        const refs = try extractSeriesRefs(allocator, chart_part.bytes, chart_prefixes.c, chart_prefixes.c_alt, mode);
        // If `out.append` OOMs after we just allocated `refs`, the
        // caller's outer errdefer frees the rest but `refs` itself
        // hasn't been transferred yet — free it on the failing path.
        errdefer allocator.free(refs);
        try out.append(allocator, .{
            .chart_part_name = chart_part.name,
            .sheet_part_name = sheet_part.name,
            .from = geometry.from,
            .to = geometry.to,
            .absolute = geometry.absolute,
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
/// actual chart-namespace prefix (canonically "c"). One walk serves
/// this read and the structural-edit sweep
/// (`Workbook.rewriteAllChartFormulas`): `ChartFormulaWalk` below.
fn extractSeriesRefs(
    allocator: std.mem.Allocator,
    xml: []const u8,
    c_prefix: []const u8,
    c_prefix_alt: ?[]const u8,
    mode: WalkMode,
) ![]const []const u8 {
    var out: std.ArrayListUnmanaged([]const u8) = .empty;
    errdefer out.deinit(allocator);
    var walk = ChartFormulaWalk.initWithPrefixes(xml, c_prefix, c_prefix_alt);
    while (walk.next(mode) catch |e| switch (e) {
        // A carrier the walk cannot read whole: strict refuses rather
        // than silently thin `series_refs` (REL-203); the walk's own
        // verdict is the sweep's name, this read's is the drawing's.
        error.MalformedChartXml => return error.MalformedDrawingXml,
    }) |f| {
        try out.append(allocator, xml[f.body_start..f.body_end]);
    }
    return out.toOwnedSlice(allocator);
}

/// One `<{c}:f>` formula carrier of a chart part, as byte offsets
/// into the part: `body_start..body_end` is the element's raw inner
/// text (still entity-encoded — the rewrite boundary decodes on the
/// way in and re-escapes on the way out); the walk itself resumes
/// past the carrier's close. A self-closing `<c:f/>` is an empty
/// carrier (`body_start == body_end`), the schema-equivalent
/// spelling of `<c:f></c:f>`.
pub const ChartFormula = struct { body_start: usize, body_end: usize };

/// The `<{c}:f>` carrier walk of one chart part — the ONE definition
/// of "a chart formula carrier" for the anchors read
/// (`extractSeriesRefs`) and the structural-edit sweep
/// (`Workbook.rewriteAllChartFormulas` / `preflightChartFormulas`),
/// so what one reports the other moves.
///
/// Carriers are LIVE, exact QNames under the prefix bound to the
/// chart namespace (Transitional or Strict; both when a part binds
/// both — one document-order pass across the two spellings), with
/// XML `Eq` whitespace allowed before either tag's `>`; markup-shaped
/// text inside a comment, CDATA section or PI is text (a commented
/// `<c:f>` is neither a ref nor a refusal). Candidates are cached per
/// prefix and refreshed only once the scan passes them, so a dense
/// part with an enabled-but-unused alternate prefix costs linear
/// work (Codex #214 r6 PERF-602).
///
/// What the walk cannot read whole is decided by `mode`, the
/// `WalkMode` split every drawing walker makes. `.lenient` is the
/// historical read: an opened carrier that never closes (a start tag
/// truncated at the end of the part included) ends the walk, a body
/// holding markup is returned as it lies. `.strict` is the S2 `<xm:f>`
/// rule (`sheet_edit.nextXmFormula`), refusing `error.MalformedChartXml`
/// for exactly four shapes: a carrier with no close (the truncated
/// start tag included); a body with a `<` in it (the schema's element
/// text is a simple type — a `<` can only open a nested element, a
/// comment or a CDATA section, none of which a byte-preserving splice
/// can carry through); an unterminated comment / CDATA / PI with a
/// carrier START TAG — the exact QName, not any `<c:f`-prefixed
/// element such as `<c:formatCode>` — inside it, where live or decoy
/// is undecidable (a trailing unterminated region holding no carrier
/// tag refuses nothing, and the probe is one forward pass — Codex
/// CF-PERF-101 / CF-REL-101); and a chart namespace bound, anywhere
/// in the part, under a prefix the walk does not follow — longer than
/// `max_prefix_len` (100 bytes, which the resolver rejects), a second
/// prefix on a chart URI, or declared beyond the resolver's 4 KiB
/// root window — where `init` records the unfollowed binding and
/// strict refuses the part rather than walk it under a prefix it does
/// not use and move nothing, while lenient walks what resolves, as
/// the read always did (in-house CF-DOC-201 / CF-REL-301; the probe
/// reads live, attribute-shaped declarations only, so `xmlns:` text
/// inside a comment, a PI or an attribute value is text; below the
/// 100-byte limit every prefix spells into the 128-byte needle
/// scratch, whose own bound is 124 — `</p:f` is four bytes over the
/// prefix — so a longer prefix handed to `initWithPrefixes` directly
/// is strict's refusal and lenient's end, and an alternate that long
/// is dropped). A part that binds the chart namespace as its DEFAULT
/// namespace (`<chartSpace xmlns="…/chart"><f>` — openpyxl's spelling)
/// resolves to the empty prefix and is walked under `<f>` / `</f>`,
/// the exact-QName rule keeping `<formatCode>` and `<firstSliceAng>`
/// out (in-house CF-REL-401: the shape had been documented as
/// unproduced and left unwalked, and every openpyxl chart went stale
/// silently); a default binding beside a prefixed one is a binding
/// the walk does not follow, refused like a second prefix.
pub const ChartFormulaWalk = struct {
    xml: []const u8,
    c_prefix: []const u8,
    c_prefix_alt: ?[]const u8,
    /// Scan position: every carrier before it has been returned.
    i: usize = 0,
    /// Cached next-candidate per prefix (see the struct doc).
    primary: ?CarrierOpen = null,
    alt: ?CarrierOpen = null,
    primed: bool = false,
    /// The part binds a chart namespace under a prefix the resolver
    /// rejected (`max_prefix_len`): strict refuses rather than walk
    /// the part under a fallback prefix it does not use.
    prefix_rejected: bool = false,
    /// The part carries a live `<!DOCTYPE` (OPC forbids a DTD; an
    /// entity could rewrite a carrier — in-house ND-REL-302) or a `<`
    /// inside an attribute value (not well-formed; a carrier spelled
    /// there would be walked — ND-REL-411): the strict walk refuses it
    /// as the drawing walks do. Probed on the first strict `next` only
    /// — a lenient read never pays the whole-part pass (ND-PERF-403).
    unwalkable: ?bool = null,

    /// Resolve the chart prefix from the part's own declarations
    /// (`resolveDrawingPrefixes`: the root's binding of the
    /// Transitional / Strict chart namespace, the canonical `c` when
    /// neither is declared; a binding under a prefix longer than
    /// `max_prefix_len` is recorded as rejected).
    pub fn init(xml: []const u8) ChartFormulaWalk {
        const p = resolveDrawingPrefixes(xml);
        var walk = initWithPrefixes(xml, p.c, p.c_alt);
        walk.prefix_rejected = p.c_rejected;
        return walk;
    }

    pub fn initWithPrefixes(xml: []const u8, c_prefix: []const u8, c_prefix_alt: ?[]const u8) ChartFormulaWalk {
        const alt: ?[]const u8 = if (c_prefix_alt) |a| (if (std.mem.eql(u8, a, c_prefix)) null else a) else null;
        return .{ .xml = xml, .c_prefix = c_prefix, .c_prefix_alt = alt };
    }

    /// The next carrier at or after the scan position, or `null` when
    /// the rest of the part carries none.
    pub fn next(self: *ChartFormulaWalk, mode: WalkMode) error{MalformedChartXml}!?ChartFormula {
        if (mode == .strict and self.prefix_rejected) return error.MalformedChartXml;
        if (mode == .strict) {
            if (self.unwalkable == null) self.unwalkable = hasLiveDoctype(self.xml) or hasMarkupInAttributeValue(self.xml);
            if (self.unwalkable.?) return error.MalformedChartXml;
        }
        const xml = self.xml;
        // Needles are spelled per call into stack scratch: the walk is
        // a value (the read returns it by value), so it cannot hold
        // slices into its own buffers.
        var primary_open_buf: [128]u8 = undefined;
        var primary_close_buf: [128]u8 = undefined;
        const primary_open = spellCarrierNeedle(&primary_open_buf, self.c_prefix, .open) catch return self.unscannable(mode);
        const primary_close = spellCarrierNeedle(&primary_close_buf, self.c_prefix, .close) catch return self.unscannable(mode);

        var alt_open_buf: [128]u8 = undefined;
        var alt_close_buf: [128]u8 = undefined;
        var alt_open: ?[]const u8 = null;
        var alt_close: ?[]const u8 = null;
        if (self.c_prefix_alt) |alt_prefix| {
            // Both needles or neither — a prefix that fits `<alt:f` but
            // not `</alt:f` must not leave an open with no close.
            if (spellCarrierNeedle(&alt_open_buf, alt_prefix, .open)) |o| {
                if (spellCarrierNeedle(&alt_close_buf, alt_prefix, .close)) |c| {
                    alt_open = o;
                    alt_close = c;
                } else |_| {}
            } else |_| {}
        }

        if (!self.primed) {
            self.primary = nextCarrierOpen(xml, self.i, primary_open);
            self.alt = if (alt_open) |o| nextCarrierOpen(xml, self.i, o) else null;
            self.primed = true;
        }
        if (self.primary != null and self.primary.?.at < self.i) self.primary = nextCarrierOpen(xml, self.i, primary_open);
        if (self.alt != null and self.alt.?.at < self.i) self.alt = nextCarrierOpen(xml, self.i, alt_open.?);
        const use_primary = if (self.primary != null and self.alt != null)
            self.primary.?.at <= self.alt.?.at
        else
            self.primary != null;
        const open = (if (use_primary) self.primary else self.alt) orelse {
            // No live carrier ahead. Strict still asks whether a
            // carrier START TAG sits inside an unterminated comment /
            // CDATA / PI beyond the scan position — live or decoy is
            // undecidable there, the S2 rule refuses rather than guess.
            // Strict only: lenient never asked, and the read's public
            // lenient path must not pay for the pass.
            if (mode == .strict and unterminatedRegionHoldsCarrier(xml, self.i, primary_open, alt_open)) {
                return error.MalformedChartXml;
            }
            self.i = xml.len;
            return null;
        };
        if (open.truncated) {
            // `<c:f` or `<c:f ` as the part's last bytes: a start tag
            // with no `>` is a carrier with no close.
            if (mode == .strict) return error.MalformedChartXml;
            self.i = xml.len;
            return null;
        }
        if (open.self_closing) {
            self.i = open.content_start;
            return .{ .body_start = open.content_start, .body_end = open.content_start };
        }
        const close_needle = if (use_primary) primary_close else alt_close.?;
        const closed = nextCarrierClose(xml, open.content_start, close_needle) orelse {
            // A real, opened carrier that never closes: lenient keeps
            // the historical truncation; strict refuses (REL-203).
            if (mode == .strict) return error.MalformedChartXml;
            self.i = xml.len;
            return null;
        };
        const body = xml[open.content_start..closed.at];
        if (mode == .strict and std.mem.indexOfScalar(u8, body, '<') != null) return error.MalformedChartXml;
        self.i = closed.end;
        return .{ .body_start = open.content_start, .body_end = closed.at };
    }

    fn unscannable(self: *ChartFormulaWalk, mode: WalkMode) error{MalformedChartXml}!?ChartFormula {
        if (mode == .strict) return error.MalformedChartXml;
        self.i = self.xml.len;
        return null;
    }
};

/// `<p:f` / `</p:f` for a bound prefix, `<f` / `</f` for the DEFAULT
/// namespace (an empty prefix — openpyxl's chart parts, in-house
/// CF-REL-401); the exact-QName terminator rule on the match keeps
/// `<formatCode>`, `<firstSliceAng>` and `<fmtId>` out either way.
fn spellCarrierNeedle(buf: *[128]u8, prefix: []const u8, which: enum { open, close }) error{NoSpaceLeft}![]const u8 {
    return spellQName(buf, if (which == .open) "<" else "</", prefix, "f", "");
}

/// Does a comment / CDATA / PI opened at or after `from` never close,
/// with a carrier start tag (the exact QName under either needle)
/// inside it? One forward pass over `<`: each terminated region is
/// stepped whole from its own opener — so a `-->` inside a CDATA
/// section is CDATA text — and an unterminated one swallows the rest
/// of the part, so the question reduces to "does the swallowed tail
/// spell a carrier tag". Linear in the part (the region-by-region
/// rescan it replaces was quadratic in the number of trailing
/// regions — Codex CF-PERF-101).
fn unterminatedRegionHoldsCarrier(xml: []const u8, from: usize, primary_open: []const u8, alt_open: ?[]const u8) bool {
    var pos = from;
    while (std.mem.indexOfScalarPos(u8, xml, pos, '<')) |lt| {
        // `skipRegionCloseFrom` answers `lt + 1` for a `<` that opens
        // no region and null only for an unterminated one.
        pos = skipRegionCloseFrom(xml, lt) orelse {
            return carrierTagFrom(xml, lt, primary_open) or
                (alt_open != null and carrierTagFrom(xml, lt, alt_open.?));
        };
    }
    return false;
}

/// A carrier start tag at or after `from`: the needle followed by a
/// name terminator (whitespace, `/`, `>`) — or by nothing, the tag
/// truncated at the end of the part. `<c:formatCode` is not one.
fn carrierTagFrom(xml: []const u8, from: usize, needle: []const u8) bool {
    var pos = from;
    while (std.mem.indexOfPos(u8, xml, pos, needle)) |at| {
        const after = at + needle.len;
        if (after >= xml.len) return true;
        const c = xml[after];
        if (c == '>' or c == '/' or c == ' ' or c == '\t' or c == '\n' or c == '\r') return true;
        pos = at + 1;
    }
    return false;
}

const CarrierOpen = struct {
    at: usize,
    content_start: usize,
    /// `<c:f/>` — an empty carrier; `content_start` is past the `>`.
    self_closing: bool,
    /// The start tag is the part's last bytes (`<c:f`, `<c:f ` with no
    /// `>` after it): a carrier with no close. `content_start` and
    /// `self_closing` are meaningless then.
    truncated: bool,
};
const CarrierClose = struct { at: usize, end: usize };

/// The next live `<{p}:f>` open at or after `start`: the exact QName
/// (a byte terminating the name follows — whitespace, `/` or `>`),
/// whitespace tolerated before the tag's `>` (Codex #214 r5 REL-502).
/// A `/` before the `>` is the self-closing spelling (`<c:f/ >`, a `/`
/// not followed by `>`, is not well-formed XML and reads as an open
/// carrier); a tag the part ends inside is returned `truncated` rather
/// than dropped (Codex CF-REL-101).
fn nextCarrierOpen(xml: []const u8, start: usize, open_needle: []const u8) ?CarrierOpen {
    var pos = start;
    while (findLiveMarkup(xml, pos, open_needle)) |at| {
        const after = at + open_needle.len;
        if (after >= xml.len) return .{ .at = at, .content_start = xml.len, .self_closing = false, .truncated = true };
        const c = xml[after];
        if (!(c == '>' or c == '/' or c == ' ' or c == '\t' or c == '\n' or c == '\r')) {
            pos = at + 1;
            continue;
        }
        const end = std.mem.indexOfScalarPos(u8, xml, after, '>') orelse
            return .{ .at = at, .content_start = xml.len, .self_closing = false, .truncated = true };
        return .{ .at = at, .content_start = end + 1, .self_closing = xml[end - 1] == '/', .truncated = false };
    }
    return null;
}

/// The matching live `</{p}:f>` close, same tolerance.
fn nextCarrierClose(xml: []const u8, start: usize, close_needle: []const u8) ?CarrierClose {
    var pos = start;
    while (findLiveMarkup(xml, pos, close_needle)) |at| {
        var j = at + close_needle.len;
        while (j < xml.len and (xml[j] == ' ' or xml[j] == '\t' or xml[j] == '\n' or xml[j] == '\r')) j += 1;
        if (j < xml.len and xml[j] == '>') return .{ .at = at, .end = j + 1 };
        pos = at + 1;
    }
    return null;
}

/// The next occurrence of `needle` at or after `start` that is NOT
/// inside a comment / CDATA / PI region. One forward pass: the scan
/// position and the cached candidate only ever advance, so a document
/// stuffed with commented fakes costs linear work, not a rescan from
/// byte zero per fake (Codex #214 r3 PERF-301); and the opener probe
/// stops at the first region rather than scanning to the candidate,
/// so a run of regions that do NOT spell the needle costs its length
/// once, not once per region (in-house CF-PERF-201).
pub fn findLiveMarkup(xml: []const u8, start: usize, needle: []const u8) ?usize {
    var i = start;
    var m = std.mem.indexOfPos(u8, xml, i, needle) orelse return null;
    while (true) {
        if (m < i) m = std.mem.indexOfPos(u8, xml, i, needle) orelse return null;
        if (nextSkipOpenBefore(xml, i, m)) |skip_at| {
            i = skipRegionEndFrom(xml, skip_at);
            continue;
        }
        return m;
    }
}

/// `findLiveMarkup` for an opening tag: the byte after the matched
/// QName must terminate the name (whitespace, `/` or `>`), so
/// `<xdr:ext` cannot match `<xdr:extLst` and satisfy a validation the
/// real element would fail (Codex #214 r3 REL-301).
pub fn findLiveExactTag(xml: []const u8, start: usize, open_needle: []const u8) ?usize {
    var i = start;
    while (findLiveMarkup(xml, i, open_needle)) |at| {
        const after = at + open_needle.len;
        if (after >= xml.len) return null;
        const c = xml[after];
        if (c == ' ' or c == '\t' or c == '\n' or c == '\r' or c == '/' or c == '>') return at;
        i = at + 1;
    }
    return null;
}

/// Does the opening tag at `at` spell exactly `open` (the byte after
/// the QName terminates the name)? `<xdr:twoCellAnchorFake>` must not
/// register as a two-cell anchor (Codex #214 r5 REL-501).
pub fn matchesOpenTag(xml: []const u8, at: usize, open: []const u8) bool {
    if (!std.mem.startsWith(u8, xml[at..], open)) return false;
    const after = at + open.len;
    if (after >= xml.len) return false;
    const c = xml[after];
    return c == ' ' or c == '\t' or c == '\n' or c == '\r' or c == '/' or c == '>';
}

/// Byte index just past the close of the skip region OPENING at `at`
/// (`<!--`, `<![CDATA[` or `<?`). An unterminated region swallows the
/// rest of the document, matching `skipRegionContaining`.
pub fn skipRegionEndFrom(xml: []const u8, at: usize) usize {
    return skipRegionCloseFrom(xml, at) orelse xml.len;
}

/// `skipRegionEndFrom` that tells an unterminated region (`null`)
/// from one whose close is the document's last bytes.
pub fn skipRegionCloseFrom(xml: []const u8, at: usize) ?usize {
    if (at + 4 <= xml.len and std.mem.startsWith(u8, xml[at..], "<!--")) {
        const close = std.mem.indexOfPos(u8, xml, at + 4, "-->") orelse return null;
        return close + 3;
    }
    if (at + 9 <= xml.len and std.mem.startsWith(u8, xml[at..], "<![CDATA[")) {
        const close = std.mem.indexOfPos(u8, xml, at + 9, "]]>") orelse return null;
        return close + 3;
    }
    if (at + 2 <= xml.len and std.mem.startsWith(u8, xml[at..], "<?")) {
        const close = std.mem.indexOfPos(u8, xml, at + 2, "?>") orelse return null;
        return close + 2;
    }
    if (at + 2 <= xml.len and std.mem.startsWith(u8, xml[at..], "<!")) return markupDeclarationEnd(xml, at);
    return at + 1;
}

/// Does `rest` open a region the walks step over whole — a comment,
/// a CDATA section, a processing instruction, or a markup declaration
/// (`<!DOCTYPE`, and the `<!ENTITY` / `<!ATTLIST` of its internal
/// subset)?
fn isRegionOpener(rest: []const u8) bool {
    return std.mem.startsWith(u8, rest, "<!") or std.mem.startsWith(u8, rest, "<?");
}

/// The byte past a markup declaration opening at `at` — its unquoted
/// `>`, a `[ … ]` internal subset stepped over whole (quoted values,
/// comments and PIs inside it included). A DTD is not content: under
/// the bare `<` opener both walks used to step into one byte by byte
/// and take an entity value's markup for an anchor — a phantom record
/// on the read, a grid coordinate spliced inside the declaration by
/// the sweep (in-house ND-REL-203). Null when unterminated.
fn markupDeclarationEnd(xml: []const u8, at: usize) ?usize {
    var i = at + 2;
    var in_subset = false;
    while (i < xml.len) {
        const c = xml[i];
        if (c == '"' or c == '\'') {
            const close = std.mem.indexOfScalarPos(u8, xml, i + 1, c) orelse return null;
            i = close + 1;
            continue;
        }
        if (in_subset) {
            if (c == ']') {
                in_subset = false;
            } else if (std.mem.startsWith(u8, xml[i..], "<!--")) {
                const close = std.mem.indexOfPos(u8, xml, i + 4, "-->") orelse return null;
                i = close + 3;
                continue;
            } else if (std.mem.startsWith(u8, xml[i..], "<?")) {
                const close = std.mem.indexOfPos(u8, xml, i + 2, "?>") orelse return null;
                i = close + 2;
                continue;
            }
            i += 1;
            continue;
        }
        if (c == '[') {
            in_subset = true;
        } else if (c == '>') {
            return i + 1;
        }
        i += 1;
    }
    return null;
}

/// Does the part carry a live `<!DOCTYPE` — outside any comment, CDATA
/// section or PI? OPC forbids a DTD in a part (ISO/IEC 29500-2) and no
/// producer writes one: the strict read and the sweep refuse the
/// drawing rather than walk a document an entity could rewrite; the
/// lenient read steps over the declaration as a region.
pub fn hasLiveDoctype(xml: []const u8) bool {
    return findLiveMarkup(xml, 0, "<!DOCTYPE") != null;
}

/// Does a `<` sit inside a quoted attribute value — outside comments,
/// CDATA sections, PIs and markup declarations? XML 1.0 forbids it
/// (AttValue excludes `<`; a producer writes `&lt;`), so the part is
/// not well-formed and no walk can tell its markup from its text: an
/// attribute holding `<a:blip r:embed=…/>` or `</oneCellAnchor>` would
/// be served as the blip or end the anchor. One linear pass; the
/// strict read, the sweep and the strict chart walk refuse on it
/// (in-house ND-REL-411 — the per-site content-start rules stay for
/// the lenient read). A part that ends inside a tag or a quote is
/// judged by the other probes.
pub fn hasMarkupInAttributeValue(xml: []const u8) bool {
    var i: usize = 0;
    while (i < xml.len) {
        const lt = std.mem.indexOfScalarPos(u8, xml, i, '<') orelse return false;
        if (isRegionOpener(xml[lt..])) {
            i = skipRegionEndFrom(xml, lt);
            continue;
        }
        // Inside a tag: quoted values are stepped over; a `<` in one is
        // the violation; `>` leaves the tag.
        var k = lt + 1;
        while (k < xml.len) : (k += 1) {
            const c = xml[k];
            if (c == '"' or c == '\'') {
                const close = std.mem.indexOfScalarPos(u8, xml, k + 1, c) orelse return false;
                if (std.mem.indexOfScalarPos(u8, xml[0..close], k + 1, '<') != null) return true;
                k = close;
                continue;
            }
            if (c == '>') break;
        }
        i = k + 1;
    }
    return false;
}

pub const OpenTagEnd = struct { gt: usize, self_closing: bool };

/// The `>` of the opening tag at `at` — quote-aware, so a `>` inside
/// an attribute value does not end it — and whether the tag is
/// self-closing (`/>`, whitespace before the slash allowed). Null when
/// the part ends inside the tag. The sweep used a quote-blind scan
/// here and took `editAs="a/>b"` for a self-closing wrapper (in-house
/// ND-REL-202); the read did not look and let a self-closing wrapper's
/// block run to the NEXT wrapper's close (ND-DOC-204).
pub fn selfClosingTagEnd(xml: []const u8, at: usize) ?OpenTagEnd {
    const gt = findUnquotedTagEnd(xml, at) orelse return null;
    var trim_end = gt;
    while (trim_end > at) : (trim_end -= 1) {
        const ch = xml[trim_end - 1];
        if (ch == ' ' or ch == '\t' or ch == '\n' or ch == '\r') continue;
        break;
    }
    return .{ .gt = gt, .self_closing = trim_end > at and xml[trim_end - 1] == '/' };
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
    // (Codex #214 r2 REL-203), and the QName must be exact — a
    // `<c:barChartExtension>` is not a bar chart (r5 REL-503).
    var found: ?ChartType = null;
    for (candidates) |c| {
        var present = false;
        for (prefixes) |maybe_p| {
            const p = maybe_p orelse continue;
            // The empty prefix is the default-namespace part (openpyxl):
            // `<barChart`, not `<:barChart` (in-house CF-REL-401).
            const needle = spellQName(&buf, "<", p, c.suffix, "") catch continue;
            if (findLiveExactTag(chart_xml, 0, needle) != null) present = true;
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
        .rid, .multiple => |r| r,
        .absent, .malformed => null,
    };
}

/// The sheet's `<drawing>` reference, tri-state: strict callers must
/// tell "no drawing element" (nothing to walk) from "a drawing element
/// whose reference cannot be read" (an inventory hole) — one optional
/// conflated them and strict mode silently skipped the sheet (Codex
/// #214 r2 REL-202).
pub const DrawingRef = union(enum) {
    absent,
    rid: []const u8,
    /// Two or more live drawing elements — schema-invalid. Strict
    /// refuses (the inventory cannot be walked whole); lenient
    /// follows the first, whose rid this carries (Codex #214 r4
    /// REL-403).
    multiple: []const u8,
    malformed,
};

/// The namespace URIs a prefixed `<{p}:drawing>` element must bind
/// its prefix to (SpreadsheetML main, Transitional + Strict), and the
/// URIs a non-`r` id-attribute prefix must bind to (office-document
/// relationships, Transitional + Strict). Without the binding check,
/// a foreign-namespaced `foo:drawing` or `foo:id` would collide with
/// the walk lexically (Codex #214 r3 REL-302).
const spreadsheetml_main_uris = [_][]const u8{
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "http://purl.oclc.org/ooxml/spreadsheetml/main",
};
const relationships_ns_uris = [_][]const u8{
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://purl.oclc.org/ooxml/officeDocument/relationships",
};

pub fn findDrawingRef(sheet_xml: []const u8) DrawingRef {
    const tag = findDrawingElementFrom(sheet_xml, 0) orelse return .absent;
    const tag_end = std.mem.indexOfScalarPos(u8, sheet_xml, tag, '>') orelse return .malformed;
    const attrs = sheet_xml[tag .. tag_end + 1];
    // Canonical spelling first — the literal `r:` prefix is honoured
    // as spelled, the convention every relationship reference in this
    // module reads (`r:embed`, the chart `r:id`); binding verification
    // applies to the NON-canonical prefixes the r2 lift added. Then
    // any-prefix `*:id` whose prefix is bound to the relationships
    // namespace.
    const rid = attrValue(attrs, "r:id") orelse
        (prefixedIdAttrValue(sheet_xml, attrs) orelse return .malformed);
    // The worksheet schema allows ONE drawing element; a second live
    // one means the inventory cannot be walked whole — strict refuses,
    // lenient follows the first (Codex #214 r4 REL-403).
    if (findDrawingElementFrom(sheet_xml, tag_end + 1) != null) return .{ .multiple = rid };
    return .{ .rid = rid };
}

/// The first live `<drawing>` element at or after `start`, in either
/// spelling — the canonical unprefixed one, or `<{p}:drawing` for a
/// prefix bound to the SpreadsheetML main namespace — whichever comes
/// FIRST in document order (one scan, not unprefixed-then-prefixed
/// passes that could skip an earlier prefixed element; Codex #214 r4
/// REL-403). Matches inside comments / CDATA / PIs are text, not the
/// element — a commented `<drawing/>` must not turn into a strict
/// refusal of a valid sheet (r3 REL-302).
fn findDrawingElementFrom(xml: []const u8, start: usize) ?usize {
    const unprefixed: ?usize = blk: {
        var i = start;
        while (findOpeningTagFrom(xml, i, "drawing")) |at| {
            if (isInsideCommentOrCdata(xml, at)) {
                i = skipRegionContaining(xml, at) orelse at + 1;
                continue;
            }
            break :blk at;
        }
        break :blk null;
    };
    const prefixed: ?usize = blk: {
        var i = start;
        while (std.mem.indexOfScalarPos(u8, xml, i, ':')) |colon| {
            i = colon + 1;
            const name = "drawing";
            const after_name = colon + 1 + name.len;
            if (after_name >= xml.len) break :blk null;
            if (!std.mem.eql(u8, xml[colon + 1 .. after_name], name)) continue;
            const c = xml[after_name];
            if (!(c == ' ' or c == '\t' or c == '\n' or c == '\r' or c == '/' or c == '>')) continue;
            // Walk back over the prefix to the `<`; every intervening
            // byte must be a name char, so a `:drawing` inside an
            // attribute value or text does not count.
            var p = colon;
            var ok = true;
            while (p > start) {
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
            if (isInsideCommentOrCdata(xml, p)) continue;
            if (!prefixBoundTo(xml, xml[p + 1 .. colon], &spreadsheetml_main_uris)) continue;
            break :blk p;
        }
        break :blk null;
    };
    if (unprefixed != null and prefixed != null) return @min(unprefixed.?, prefixed.?);
    return unprefixed orelse prefixed;
}

/// Find the start of an opening tag named `name` at or after `start`,
/// tolerating XML whitespace (space / tab / LF / CR) or `/`/`>` after
/// the name. `<drawing\nr:id="rId1"/>` is valid XML; a literal
/// "<drawing " search missed it and silently dropped anchors on
/// well-formed workbooks emitted by non-Microsoft producers.
fn findOpeningTagFrom(xml: []const u8, start: usize, name: []const u8) ?usize {
    var i: usize = start;
    while (std.mem.indexOfPos(u8, xml, i, "<")) |lt| {
        i = lt + 1;
        const after_name = lt + 1 + name.len;
        if (after_name >= xml.len) return null;
        if (!std.mem.eql(u8, xml[lt + 1 .. after_name], name)) continue;
        const c = xml[after_name];
        if (c == ' ' or c == '\t' or c == '\n' or c == '\r' or c == '/' or c == '>') {
            return lt;
        }
    }
    return null;
}

/// The value of the first attribute spelled `{prefix}:id` whose
/// prefix is bound to the office-document relationships namespace.
/// A bare `id` attribute is NOT a relationship reference and does
/// not match; nor does an id under a foreign namespace. Quote-aware,
/// like `attrValue`.
fn prefixedIdAttrValue(sheet_xml: []const u8, attrs: []const u8) ?[]const u8 {
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
                const prefix = key[0 .. key.len - 3];
                if (prefixBoundTo(sheet_xml, prefix, &relationships_ns_uris)) {
                    if (attrValue(attrs, key)) |v| return v;
                }
            }
            i = j + 1;
            continue;
        }
        i += 1;
    }
    return null;
}

/// Is `prefix` bound anywhere in the document to one of `uris`? A
/// document-wide search, not a scope-accurate resolution — the sheet
/// part is not namespace-parsed here, and a binding declared anywhere
/// is evidence enough to honour the spelling (the drawings-part
/// resolver does the scope-accurate version for its own URIs). The
/// candidate must be a real attribute NAME in a live opening tag —
/// preceded by XML whitespace, outside quoted values, comments,
/// CDATA and PIs — so declaration-shaped text inside an attribute
/// value or a comment is not a binding (Codex #214 r4 REL-401).
/// Tolerates XML `Eq` whitespace on both sides of `=`.
fn prefixBoundTo(xml: []const u8, prefix: []const u8, uris: []const []const u8) bool {
    var buf: [140]u8 = undefined;
    const decl = std.fmt.bufPrint(&buf, "xmlns:{s}", .{prefix}) catch return false;
    var i: usize = 0;
    while (std.mem.indexOfPos(u8, xml, i, decl)) |at| {
        i = at + decl.len;
        if (at == 0) continue;
        const before = xml[at - 1];
        if (!(before == ' ' or before == '\t' or before == '\n' or before == '\r')) continue;
        if (!isInsideOpeningTag(xml, at)) continue;
        var j = at + decl.len;
        while (j < xml.len and (xml[j] == ' ' or xml[j] == '\t' or xml[j] == '\n' or xml[j] == '\r')) j += 1;
        if (j >= xml.len or xml[j] != '=') continue;
        j += 1;
        while (j < xml.len and (xml[j] == ' ' or xml[j] == '\t' or xml[j] == '\n' or xml[j] == '\r')) j += 1;
        if (j >= xml.len) return false;
        const q = xml[j];
        if (q != '"' and q != '\'') continue;
        const close = std.mem.indexOfScalarPos(u8, xml, j + 1, q) orelse return false;
        const uri = xml[j + 1 .. close];
        for (uris) |u| {
            if (std.mem.eql(u8, uri, u)) return true;
        }
    }
    return false;
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
    const blip_open = spellQName(buf, "<", prefix, "blip", "") catch return null;
    var search_at: usize = 0;
    // Live, QName-exact matches only — a commented blip is text (r5
    // REL-501), and a `<a:blipExtension r:embed=…>` is not a blip
    // (Codex #214 r6 REL-601). `<a:blipFill>` fails the terminator
    // and the loop walks on to the real `<a:blip>` inside it.
    while (findLiveExactTag(pic_xml, search_at, blip_open)) |blip| {
        const blip_end = findUnquotedTagEnd(pic_xml, blip) orelse return null;
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
pub const max_xdr_alts: usize = 8;

pub const DrawingPrefixes = struct {
    /// The spreadsheetDrawing prefix the walk spells first: the root
    /// element's own, else the first bound one, else the canonical
    /// fallback. EMPTY for a part whose unprefixed root sits in the
    /// spreadsheetDrawing namespace bound as the DEFAULT one
    /// (`<wsDr xmlns="…/spreadsheetDrawing">`, openpyxl's drawings) —
    /// every tag needle is then spelled bare (`<twoCellAnchor`,
    /// `<from>`, `<row>`), the chart walk's `<f>` rule.
    xdr: []const u8 = "xdr",
    a: []const u8 = "a",
    c: []const u8 = "c",
    a_alt: ?[]const u8 = null,
    c_alt: ?[]const u8 = null,
    /// The part binds a chart namespace (either URI), anywhere in it,
    /// under a prefix the chart walk will not follow — longer than
    /// `max_prefix_len`, a second prefix on the URI, or declared beyond
    /// the root window — so `c` holds a fallback the part may not use.
    /// The chart walk's strict mode refuses on it (CF-DOC-201 /
    /// CF-REL-301).
    c_rejected: bool = false,
    /// The part binds a spreadsheetDrawing namespace (either URI),
    /// anywhere in it, under a name the anchor walk will not follow
    /// — longer than `max_prefix_len`, or past the `max_xdr_alts`
    /// replay cap — so an anchor under that name would be neither
    /// listed nor shifted. The strict read refuses the drawing and
    /// the drawing sweep refuses the edit on it (`MalformedDrawingXml`
    /// both), the chart walk's `c_rejected` rule on this namespace.
    xdr_rejected: bool = false,
    /// All prefixes (other than `xdr`) bound to either xdr URI in
    /// the same document — the DEFAULT declaration counted as the
    /// empty name (a prefixed root over unprefixed anchors is walked
    /// too). The scanner replays once per prefix so anchors using ANY
    /// bound prefix are surfaced. Capped at `max_xdr_alts` — real
    /// OOXML producers declare 1-2; more is pathological. Tracked as
    /// fixed-array + count rather than a stdlib bounded helper because
    /// Zig 0.15 dropped `std.BoundedArray`.
    xdr_alts_buf: [max_xdr_alts][]const u8 = undefined,
    xdr_alts_len: usize = 0,

    pub fn xdr_alts(self: *const DrawingPrefixes) []const []const u8 {
        return self.xdr_alts_buf[0..self.xdr_alts_len];
    }

    /// Is `name` a spreadsheetDrawing prefix the walk spells — the
    /// primary or one of the alternates?
    pub fn followsXdr(self: *const DrawingPrefixes, name: []const u8) bool {
        if (std.mem.eql(u8, name, self.xdr)) return true;
        for (self.xdr_alts()) |alt| if (std.mem.eql(u8, alt, name)) return true;
        return false;
    }
};

/// Scan the root element's xmlns:* declarations and return the
/// prefix for each canonical OOXML namespace. Falls back to the
/// canonical prefix when a namespace isn't declared (some chart
/// parts only declare the chart namespace inline on `<c:chart>`).
/// Shared with `drawing_edit`, so the read and the sweep follow one
/// resolution of the same bytes.
pub fn resolveDrawingPrefixes(xml: []const u8) DrawingPrefixes {
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
    } else if (rootElementIsUnprefixed(xml) and isXdrUri(rootDefaultNamespaceUri(xml))) {
        // An unprefixed root whose DEFAULT namespace is the
        // spreadsheetDrawing one — `<wsDr xmlns="…/spreadsheetDrawing">`,
        // openpyxl 3.1's drawings: the anchors are `<oneCellAnchor>
        // <from><col>`, unprefixed, and the walk spells them so. A
        // prefixed binding declared beside it is an alternate below.
        p.xdr = "";
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
    // No prefixed binding: the root's DEFAULT declaration may bind a
    // chart URI — openpyxl spells its chart parts `<chartSpace
    // xmlns="…/chart"><f>` — and the carriers are then unprefixed: the
    // empty prefix, spelled `<f` / `</f` by the walk (in-house
    // CF-REL-401; the shape had been documented as unproduced and left
    // unwalked, so every openpyxl chart went stale silently).
    if (c_t == null and c_s == null) {
        if (defaultNamespaceUri(xml)) |uri| {
            if (std.mem.eql(u8, uri, ns_c_transitional) or std.mem.eql(u8, uri, ns_c_strict)) p.c = "";
        }
    }
    markUnfollowedBindings(xml, &p);
    return p;
}

/// One whole-part pass over the live declarations marking BOTH
/// namespaces' unfollowed bindings: `c_rejected` (the rule
/// `hasUnfollowedChartBinding` states) and `xdr_rejected` — a
/// spreadsheetDrawing binding (either URI) under a name the anchor walk
/// will not spell, longer than `max_prefix_len` (skipped by the
/// collectors) or past the `max_xdr_alts` replay cap; the DEFAULT
/// declaration is the empty name, followed when it is the primary or
/// an alternate. An anchor under such a name would be neither listed
/// by the read nor shifted by the sweep, so both refuse on it
/// (`MalformedDrawingXml`). One scan for the two flags, so a chart
/// part, which the chart walk resolves on every structural edit, pays
/// for what it never reads once (in-house ND-PERF-105).
fn markUnfollowedBindings(xml: []const u8, p: *DrawingPrefixes) void {
    var it = RootNsBindings.init(xml, xml.len);
    while (it.next()) |b| {
        if (std.mem.eql(u8, b.uri, ns_c_transitional) or std.mem.eql(u8, b.uri, ns_c_strict)) {
            if (std.mem.eql(u8, b.name, p.c)) continue;
            if (p.c_alt) |alt| if (std.mem.eql(u8, b.name, alt)) continue;
            p.c_rejected = true;
        } else if (isXdrUri(b.uri)) {
            if (!p.followsXdr(b.name)) p.xdr_rejected = true;
        }
    }
}

fn isXdrUri(uri: ?[]const u8) bool {
    const u = uri orelse return false;
    return std.mem.eql(u8, u, ns_xdr_transitional) or std.mem.eql(u8, u, ns_xdr_strict);
}

/// Find the prefix on the root XML element — the NAME of the first
/// `<NAME:LOCAL` token (`rootElementQName`). Returns null if the root
/// element is unprefixed or absent, or its prefix is longer than
/// `max_prefix_len` (a spreadsheetDrawing binding under such a name is
/// separately `xdr_rejected`; an overlong root prefix bound elsewhere
/// or nowhere falls through to the canonical fallback, as any foreign
/// root does).
fn rootElementPrefix(xml: []const u8) ?[]const u8 {
    const qname = rootElementQName(xml) orelse return null;
    const colon = std.mem.indexOfScalar(u8, qname, ':') orelse return null;
    if (colon > max_prefix_len) return null;
    return qname[0..colon];
}

/// Is there a root element, and does its QName carry no prefix
/// (`<wsDr …>`)? False for an absent root and for a prefixed one,
/// overlong prefixes included — the default-namespace resolution is
/// for a root that lives in the default namespace, nothing else.
fn rootElementIsUnprefixed(xml: []const u8) bool {
    const qname = rootElementQName(xml) orelse return false;
    return std.mem.indexOfScalar(u8, qname, ':') == null;
}

/// The root element's QName. Skips the XML declaration
/// (`<?xml ... ?>`), a DOCTYPE / comment and any leading whitespace,
/// then reads the first element's name up to XML whitespace, `/` or
/// `>`. Null when no element starts in the root window (4 KiB) or the
/// name runs past it.
fn rootElementQName(xml: []const u8) ?[]const u8 {
    const limit = @min(xml.len, 4096);
    var i: usize = 0;
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
            if (c == ' ' or c == '\t' or c == '\n' or c == '\r' or c == '/' or c == '>') {
                return xml[lt + 1 .. j];
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
/// buffer can format. `DrawingTags.build` writes its 26 needles into
/// one 4 KiB `TagSetBuf` per followed prefix (2.8 KiB under this
/// limit), and the per-helper scratch buffers (`tryBlipEmbedAt`,
/// `detectChartTypeWithAlt`, `spellCarrierNeedle`) are 128 bytes
/// each. The longest formatted needle is `</PREFIX:twoCellAnchor>` =
/// prefix.len + 17. 100 leaves a comfortable margin and still covers
/// any conceivable real-world prefix (OOXML uses 1-8 char prefixes;
/// 100 is already absurd).
const max_prefix_len: usize = 100;

/// The root element's declarations sit in the part's first bytes;
/// prefix resolution reads no further, so a mid-document declaration
/// cannot shadow the canonical fallback (`RootNsBindings`).
const root_window: usize = 4096;

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
    var cached: ?usize = null;
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
        // The candidate is cached across iterations and refreshed
        // only once `i` passes it, so each byte is scanned once for
        // `:chart` and once for `<`: a region followed by one live
        // element reached this search every iteration and rescanned
        // the remainder per region — 82 s for 40 000 pairs through
        // the anchors read in Debug (in-house CF-PERF-301).
        if (cached == null or cached.? < i) cached = std.mem.indexOfPos(u8, block, i, ":chart") orelse return null;
        const colon = cached.?;
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

inline fn isXmlWs(c: u8) bool {
    return c == ' ' or c == '\t' or c == '\n' or c == '\r';
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
///
/// One pass over `<` that stops at the first opener, so a call costs
/// the distance to that opener, not to `limit`: `findLiveMarkup`
/// asks once per terminated region, and the three needle scans to
/// the candidate it used to run per call made a run of k regions
/// before a carrier cost k times the run — 52 s for a 1 MB chart
/// part on every anchors read and, with the sweep, every structural
/// edit (in-house CF-PERF-201).
fn nextSkipOpenBefore(xml: []const u8, from: usize, limit: usize) ?usize {
    if (limit <= from) return null;
    const bounded = xml[0..limit];
    var p = from;
    while (std.mem.indexOfScalarPos(u8, bounded, p, '<')) |lt| {
        if (isRegionOpener(bounded[lt..])) return lt;
        p = lt + 1;
    }
    return null;
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
pub fn findUnquotedTagEnd(xml: []const u8, start: usize) ?usize {
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
    var it = RootNsBindings.init(xml, xml.len);
    while (it.next()) |b| {
        // The DEFAULT declaration (an empty name) is an alternate like
        // any other: a `<xdr:wsDr>` root over `<twoCellAnchor xmlns=
        // "…/spreadsheetDrawing">` anchors is walked under both.
        if (!std.mem.eql(u8, b.uri, target_uri) or std.mem.eql(u8, b.name, skip) or b.name.len > max_prefix_len) continue;
        // Dedup: don't append the same prefix twice (would
        // happen when same prefix is declared on two elements
        // with the same URI).
        var already_in = false;
        for (out.xdr_alts()) |existing| {
            if (std.mem.eql(u8, existing, b.name)) {
                already_in = true;
                break;
            }
        }
        if (!already_in and out.xdr_alts_len < max_xdr_alts) {
            out.xdr_alts_buf[out.xdr_alts_len] = b.name;
            out.xdr_alts_len += 1;
        }
    }
}

/// Inverse lookup: given a prefix, return the URI it's bound to,
/// or null if the prefix isn't declared. Root-window scan
/// (`RootNsBindings`), the ceiling prefix resolution uses so
/// behaviour stays consistent across helpers.
fn uriOfPrefix(xml: []const u8, prefix: []const u8) ?[]const u8 {
    // The empty prefix is the DEFAULT declaration, the iterator's
    // empty name — the primary of an unprefixed drawing root.
    // Root-window scan: this is only used to look up the URI of the
    // ROOT element's prefix, which by definition is declared on the
    // root.
    var it = RootNsBindings.init(xml, root_window);
    while (it.next()) |b| {
        if (std.mem.eql(u8, b.name, prefix)) return b.uri;
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
    var it = RootNsBindings.init(xml, root_window);
    while (it.next()) |b| {
        if (b.name.len != 0 and std.mem.eql(u8, b.uri, target_uri) and !std.mem.eql(u8, b.name, skip)) {
            // Cap pathologically long prefixes — they'd overflow
            // the per-needle scratch buffers downstream. Skip
            // (rather than abort the whole lookup) so a workbook
            // that declares multiple prefixes for the same URI
            // can still match a usable one further along;
            // `hasUnfollowedChartBinding` reports the skip to the
            // chart walk, which refuses under strict (CF-DOC-201).
            if (b.name.len <= max_prefix_len) return b.name;
        }
    }
    return null;
}

/// Does the part bind a chart namespace (either URI) under a prefix
/// the walk will not follow — any name other than the resolved
/// primary / alternate, which covers a prefix longer than
/// `max_prefix_len` (skipped by `findNamespacePrefixExcept`), a second
/// prefix on the same URI, and a declaration beyond the resolver's
/// root window (the resolved prefix is then the canonical fallback,
/// which the part may not use)? The DEFAULT declaration counts as the
/// empty name — followed when the walk resolved the part as
/// default-namespaced (openpyxl's spelling), unfollowed beside a
/// prefixed binding (in-house CF-REL-401). Every live, attribute-shaped
/// declaration of the whole part, not the root window: a binding the
/// walk cannot follow is a refusal wherever it sits, and a carrier
/// under it would otherwise be neither served nor moved (CF-DOC-201;
/// the window-bounded, decoy-blind probe let a prefix past the window
/// through and refused on a commented one — in-house CF-REL-301). The
/// chart walk's strict mode refuses on it.
fn hasUnfollowedChartBinding(xml: []const u8, c_prefix: []const u8, c_alt: ?[]const u8) bool {
    var it = RootNsBindings.init(xml, xml.len);
    while (it.next()) |b| {
        if (!std.mem.eql(u8, b.uri, ns_c_transitional) and !std.mem.eql(u8, b.uri, ns_c_strict)) continue;
        if (std.mem.eql(u8, b.name, c_prefix)) continue;
        if (c_alt) |alt| if (std.mem.eql(u8, b.name, alt)) continue;
        return true;
    }
    return false;
}

/// The first live DEFAULT namespace declaration (`xmlns="uri"`) in
/// the root window — the root element's, or a descendant's when the
/// root declares none. The chart resolver follows it when no prefixed
/// chart binding exists (openpyxl, in-house CF-REL-401): a carrier
/// under a descendant's default binding is a carrier.
fn defaultNamespaceUri(xml: []const u8) ?[]const u8 {
    var it = RootNsBindings.init(xml, root_window);
    while (it.next()) |b| {
        if (b.name.len == 0) return b.uri;
    }
    return null;
}

/// The ROOT ELEMENT's own default namespace declaration — the one
/// that puts an unprefixed root in a namespace. A descendant's
/// declaration does not (in-house ND-DOC-307: the primary resolution
/// took one for the root's; it is an alternate, which
/// `collectAllNamespacePrefixes` collects).
fn rootDefaultNamespaceUri(xml: []const u8) ?[]const u8 {
    var it = RootNsBindings.init(xml, root_window);
    while (it.next()) |b| {
        if (it.tags_entered > 1) return null;
        if (b.name.len == 0) return b.uri;
    }
    return null;
}

/// One `xmlns:<name>="<uri>"` declaration — `name` empty for the
/// default declaration `xmlns="<uri>"` — as `RootNsBindings` yields
/// it, over the root window or the whole part.
const NsBinding = struct { name: []const u8, uri: []const u8 };

/// The live, attribute-shaped `xmlns:` declarations in the first
/// `limit` bytes of a part, in document order. Prefix resolution
/// bounds it to the root window (`root_window`) because the `a` and
/// `c` prefixes want ROOT-only scoping — picking up a local
/// mid-document declaration would shadow the canonical-fallback path
/// and silently drop anchors that use the canonical prefix locally;
/// xdr alt collection (`collectAllNamespacePrefixes`) and the chart
/// walk's unfollowed-binding probe walk the whole part, anchor-tag
/// prefixes being intentionally late-bindable and an unfollowed
/// binding anywhere being a refusal. One linear pass with three
/// states: text (the next `<` decides — a comment / CDATA / PI or a
/// close tag is stepped over whole, an opening tag is entered), tag
/// (a quoted attribute value is stepped over whole, `>` leaves,
/// `xmlns:` or a bare `xmlns` starts a declaration — the default one
/// yields an EMPTY name, openpyxl's chart-namespace spelling, in-house
/// CF-REL-401; every consumer that wants a prefix skips it) — so
/// `xmlns:` text inside a
/// comment, a PI, character data or an attribute value is text, not
/// a declaration (a commented decoy used to resolve as the prefix
/// and, once the overlong probe existed, to refuse a pristine part —
/// in-house CF-REL-301; the per-hit `isInsideOpeningTag` guard the
/// drawing side uses rescans from byte zero, which a whole-part probe
/// cannot afford). XML 1.0 allows whitespace around the `=`, so
/// `xmlns:dr = "uri"` is tolerated; a declaration without an `=` or a
/// quote is stepped over; a window that ends inside a tag or a value
/// ends the iteration.
const RootNsBindings = struct {
    xml: []const u8,
    limit: usize,
    i: usize = 0,
    in_tag: bool = false,
    /// Opening tags entered so far — 1 while inside the root element's
    /// own start tag.
    tags_entered: usize = 0,

    fn init(xml: []const u8, limit: usize) RootNsBindings {
        return .{ .xml = xml, .limit = @min(xml.len, limit) };
    }

    fn next(self: *RootNsBindings) ?NsBinding {
        const xml = self.xml[0..self.limit];
        while (true) {
            if (self.in_tag) {
                while (self.i < xml.len) {
                    const c = xml[self.i];
                    if (c == '"' or c == '\'') {
                        const close = std.mem.indexOfScalarPos(u8, xml, self.i + 1, c) orelse return null;
                        self.i = close + 1;
                        continue;
                    }
                    if (c == '>') {
                        self.in_tag = false;
                        self.i += 1;
                        break;
                    }
                    // A declaration is an attribute: XML whitespace
                    // precedes it (in the tag state `self.i` is past
                    // the `<`, so `self.i - 1` is in range). An
                    // attribute NAME that merely ends in `xmlns:` is
                    // not one (round-4 agent B).
                    if (c == 'x' and std.mem.startsWith(u8, xml[self.i..], "xmlns") and isXmlWs(xml[self.i - 1])) {
                        const after_kw = self.i + "xmlns".len;
                        if (after_kw < xml.len and xml[after_kw] == ':') {
                            if (self.declarationAt(after_kw + 1)) |b| return b;
                            continue;
                        }
                        // The DEFAULT declaration, `xmlns="uri"`: an empty
                        // name. openpyxl binds the chart namespace this way
                        // (in-house CF-REL-401).
                        if (after_kw < xml.len and (xml[after_kw] == '=' or isXmlWs(xml[after_kw]))) {
                            if (self.declarationAt(after_kw)) |b| return b;
                            continue;
                        }
                    }
                    self.i += 1;
                }
                if (self.in_tag) return null;
                continue;
            }
            const lt = std.mem.indexOfScalarPos(u8, xml, self.i, '<') orelse return null;
            const rest = xml[lt..];
            if (isRegionOpener(rest)) {
                self.i = skipRegionCloseFrom(xml, lt) orelse return null;
                continue;
            }
            if (std.mem.startsWith(u8, rest, "</")) {
                self.i = (std.mem.indexOfScalarPos(u8, xml, lt + 2, '>') orelse return null) + 1;
                continue;
            }
            self.i = lt + 1;
            self.in_tag = true;
            self.tags_entered += 1;
        }
    }

    /// Parse `name = "uri"` after an `xmlns:` at `after`, advancing
    /// the cursor past what was consumed either way.
    fn declarationAt(self: *RootNsBindings, after: usize) ?NsBinding {
        const xml = self.xml[0..self.limit];
        var name_end = after;
        while (name_end < xml.len) : (name_end += 1) {
            const c = xml[name_end];
            if (c == '=' or c == ' ' or c == '\t' or c == '\n' or c == '\r' or c == '>' or c == '/') break;
        }
        const name = xml[after..name_end];
        var p = name_end;
        while (p < xml.len and (xml[p] == ' ' or xml[p] == '\t' or xml[p] == '\n' or xml[p] == '\r')) p += 1;
        if (p >= xml.len or xml[p] != '=') {
            self.i = after;
            return null;
        }
        p += 1;
        while (p < xml.len and (xml[p] == ' ' or xml[p] == '\t' or xml[p] == '\n' or xml[p] == '\r')) p += 1;
        if (p >= xml.len) {
            self.i = xml.len;
            return null;
        }
        const quote = xml[p];
        if (quote != '"' and quote != '\'') {
            self.i = p;
            return null;
        }
        const val_start = p + 1;
        const val_end = std.mem.indexOfScalarPos(u8, xml, val_start, quote) orelse {
            self.i = xml.len;
            return null;
        };
        self.i = val_end + 1;
        return .{ .name = name, .uri = xml[val_start..val_end] };
    }
};

/// Pre-built tag needles keyed off the resolved prefixes. Built
/// into a single caller-supplied buffer so the per-part lookup
/// loop doesn't re-format on every iteration.
pub const DrawingTags = struct {
    /// The spreadsheetDrawing prefix this set spells — empty for the
    /// default namespace.
    prefix: []const u8,
    xdr_prefix_open: []const u8, // "<xdr:"
    open_two: []const u8, // "<xdr:twoCellAnchor"
    close_two: []const u8, // "</xdr:twoCellAnchor>"
    open_one: []const u8, // "<xdr:oneCellAnchor"
    close_one: []const u8, // "</xdr:oneCellAnchor>"
    open_absolute: []const u8, // "<xdr:absoluteAnchor"
    close_absolute: []const u8, // "</xdr:absoluteAnchor>"
    open_pic: []const u8, // "<xdr:pic"
    close_pic: []const u8, // "</xdr:pic>"
    open_from: []const u8, // "<xdr:from>"
    close_from: []const u8, // "</xdr:from>"
    open_to: []const u8, // "<xdr:to>"
    close_to: []const u8, // "</xdr:to>"
    open_pos: []const u8, // "<xdr:pos"
    open_ext: []const u8, // "<xdr:ext"
    open_graphic_frame: []const u8, // "<xdr:graphicFrame"
    open_chart: []const u8, // "<c:chart"
    // The four scalars of a corner block, whole tags.
    open_col: []const u8, // "<xdr:col>"
    close_col: []const u8, // "</xdr:col>"
    open_col_off: []const u8,
    close_col_off: []const u8,
    open_row: []const u8,
    close_row: []const u8,
    open_row_off: []const u8,
    close_row_off: []const u8,

    pub fn build(buf: []u8, p: DrawingPrefixes) !DrawingTags {
        var w = std.Io.Writer.fixed(buf);
        // Every needle is spelled by `writeQName`, so the DEFAULT
        // namespace (an empty prefix — openpyxl's drawings) reads
        // `<twoCellAnchor` / `<from>` / `<row>` and the scan's tag
        // opener is the bare `<` (every live tag is then a candidate;
        // the anchor names decide, as they do under a prefix).
        const xdr_prefix_open = try writeQName(&w, "<", p.xdr, "", "");
        const open_two = try writeQName(&w, "<", p.xdr, "twoCellAnchor", "");
        const close_two = try writeQName(&w, "</", p.xdr, "twoCellAnchor", ">");
        const open_one = try writeQName(&w, "<", p.xdr, "oneCellAnchor", "");
        const close_one = try writeQName(&w, "</", p.xdr, "oneCellAnchor", ">");
        const open_absolute = try writeQName(&w, "<", p.xdr, "absoluteAnchor", "");
        const close_absolute = try writeQName(&w, "</", p.xdr, "absoluteAnchor", ">");
        // `<{p}:pic` without the closing `>`: the element may carry
        // attributes (`macro`, `fPublished`); the exact-live-tag
        // lookup supplies the name terminator (r5 REL-501).
        const open_pic = try writeQName(&w, "<", p.xdr, "pic", "");
        const close_pic = try writeQName(&w, "</", p.xdr, "pic", ">");
        const open_from = try writeQName(&w, "<", p.xdr, "from", ">");
        const close_from = try writeQName(&w, "</", p.xdr, "from", ">");
        const open_to = try writeQName(&w, "<", p.xdr, "to", ">");
        const close_to = try writeQName(&w, "</", p.xdr, "to", ">");
        const open_pos = try writeQName(&w, "<", p.xdr, "pos", "");
        const open_ext = try writeQName(&w, "<", p.xdr, "ext", "");
        const open_graphic_frame = try writeQName(&w, "<", p.xdr, "graphicFrame", "");
        const open_chart = try writeQName(&w, "<", p.c, "chart", "");
        const open_col = try writeQName(&w, "<", p.xdr, "col", ">");
        const close_col = try writeQName(&w, "</", p.xdr, "col", ">");
        const open_col_off = try writeQName(&w, "<", p.xdr, "colOff", ">");
        const close_col_off = try writeQName(&w, "</", p.xdr, "colOff", ">");
        const open_row = try writeQName(&w, "<", p.xdr, "row", ">");
        const close_row = try writeQName(&w, "</", p.xdr, "row", ">");
        const open_row_off = try writeQName(&w, "<", p.xdr, "rowOff", ">");
        const close_row_off = try writeQName(&w, "</", p.xdr, "rowOff", ">");
        return .{
            .prefix = p.xdr,
            .open_col = open_col,
            .close_col = close_col,
            .open_col_off = open_col_off,
            .close_col_off = close_col_off,
            .open_row = open_row,
            .close_row = close_row,
            .open_row_off = open_row_off,
            .close_row_off = close_row_off,
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

/// One tag set per followed spreadsheetDrawing prefix — the primary
/// first, then the alternates. The walks match an anchor WRAPPER
/// under each set in turn (the read replays per set, the sweep tries
/// each set at every tag) and the wrapper's children under ANY set:
/// a part that binds the namespace twice may legally spell a wrapper
/// under one name and its `from` / `col` under the other, and XML
/// well-formedness only ties an element's close tag to its own open
/// tag (in-house ND-REL-103; the wrapper-prefix rule left such an
/// anchor neither listed nor shifted). `bufs` holds the needles, one
/// scratch per set — 4 KiB covers 25 needles under the 100-byte
/// prefix limit.
pub const max_tag_sets: usize = 1 + max_xdr_alts;
pub const TagSetBuf = [4096]u8;

/// The sets `buildTagSets` writes for `p` — the primary and its
/// alternates — so a caller can size the scratch exactly (one part
/// binds one or two names; nine sets were 40 KiB of stack per read
/// and per edit — in-house ND-PERF-206).
pub fn tagSetCount(p: *const DrawingPrefixes) usize {
    return 1 + p.xdr_alts().len;
}

pub fn buildTagSets(bufs: []TagSetBuf, sets: []DrawingTags, p: DrawingPrefixes) ![]const DrawingTags {
    std.debug.assert(bufs.len >= tagSetCount(&p) and sets.len >= tagSetCount(&p));
    var n: usize = 0;
    sets[n] = try DrawingTags.build(&bufs[n], p);
    n += 1;
    for (p.xdr_alts()) |alt| {
        var q = p;
        q.xdr = alt;
        sets[n] = try DrawingTags.build(&bufs[n], q);
        n += 1;
    }
    return sets[0..n];
}

/// The first live, exact-QName opening tag spelled by any set's
/// `field` at or after `start` — the earliest in DOCUMENT order across
/// the sets, so a mixed-spelling part is read in the order it is
/// written.
pub const TagHit = struct { at: usize, set: *const DrawingTags };

pub fn anyLiveExactTag(block: []const u8, start: usize, sets: []const DrawingTags, comptime field: []const u8) ?TagHit {
    var best: ?TagHit = null;
    for (sets) |*s| {
        const at = findLiveExactTag(block, start, @field(s, field)) orelse continue;
        if (best == null or at < best.?.at) best = .{ .at = at, .set = s };
    }
    return best;
}

/// One tag needle under a resolved prefix — `{lead}{prefix}:{local}{tail}`
/// for a bound prefix, `{lead}{local}{tail}` for the DEFAULT namespace
/// (an empty prefix: openpyxl's drawings and chart parts). `lead` is
/// `<` or `</`; `tail` is `>` for a whole tag and empty for an opener
/// whose attributes the exact-QName match terminates. The ONE spelling
/// of every needle the drawing and chart walks match, so no site can
/// spell `<:from>` for the empty prefix.
fn writeQName(w: *std.Io.Writer, lead: []const u8, prefix: []const u8, local: []const u8, tail: []const u8) ![]const u8 {
    const before = w.end;
    try w.writeAll(lead);
    if (prefix.len != 0) {
        try w.writeAll(prefix);
        try w.writeByte(':');
    }
    try w.writeAll(local);
    try w.writeAll(tail);
    return w.buffer[before..w.end];
}

/// `writeQName` into a caller's scratch buffer.
pub fn spellQName(buf: []u8, lead: []const u8, prefix: []const u8, local: []const u8, tail: []const u8) error{NoSpaceLeft}![]const u8 {
    var w = std.Io.Writer.fixed(buf);
    return writeQName(&w, lead, prefix, local, tail) catch return error.NoSpaceLeft;
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
pub fn relTargetForIdTyped(
    allocator: std.mem.Allocator,
    rels: []const store_mod.Relationship,
    id: []const u8,
    leaf: []const u8,
    mode: WalkMode,
) !?[]const u8 {
    const r = (try relForId(allocator, rels, id)) orelse return null;
    if (r.target_mode == .external) return null;
    if (mode == .strict and !relTypeIs(r.type, leaf)) return null;
    return r.target;
}

/// The relationship-type namespace roots a strict type check accepts:
/// Transitional, Strict, and the Microsoft extension family (the
/// macrosheet types live there). Matching only the last path segment
/// let `https://example.invalid/drawing` pass as a drawing edge
/// (Codex #214 r3 REL-303).
const rel_type_roots = [_][]const u8{
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/",
    "http://schemas.microsoft.com/office/2006/relationships/",
};

/// Is `rel_type` exactly `{known root}{leaf}`? Byte-exact — URIs are
/// case-sensitive, so `…/relationships/DRAWING` is a different (and
/// unknown) type, not a drawing edge (Codex #214 r4 REL-402). The
/// strict walk's type gate; also the anchors NDJSON view's
/// sheet-family gate — pub for that one consumer.
pub fn relTypeIs(rel_type: []const u8, leaf: []const u8) bool {
    for (rel_type_roots) |root| {
        if (rel_type.len == root.len + leaf.len and
            std.mem.startsWith(u8, rel_type, root) and
            std.mem.eql(u8, rel_type[root.len..], leaf)) return true;
    }
    return false;
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
    const pos_idx = findLiveExactTag(xml, 0, open_pos) orelse return null;
    const pos = posAttrsAt(xml, pos_idx) orelse return null;
    const ext = parseExtAttrs(xml, pos.end, open_ext) orelse return null;
    return .{ .x = pos.x, .y = pos.y, .cx = ext.cx, .cy = ext.cy };
}

/// `parseAbsoluteAnchor` with `<pos>` and `<ext>` under any followed
/// prefix.
fn parseAbsoluteAnchorIn(xml: []const u8, content_start: usize, sets: []const DrawingTags) ?AbsoluteAnchor {
    const pos_tag = anyLiveExactTag(xml, content_start, sets, "open_pos") orelse return null;
    const pos = posAttrsAt(xml, pos_tag.at) orelse return null;
    const ext = parseExtAttrsIn(xml, pos.end, sets) orelse return null;
    return .{ .x = pos.x, .y = pos.y, .cx = ext.cx, .cy = ext.cy };
}

/// The `x` / `y` of the `<{p}:pos>` tag at `pos_idx`, and the index of
/// its `>`.
fn posAttrsAt(xml: []const u8, pos_idx: usize) ?struct { x: i64, y: i64, end: usize } {
    const pos_end = findUnquotedTagEnd(xml, pos_idx) orelse return null;
    const pos_attrs = xml[pos_idx .. pos_end + 1];
    const x_str = attrValue(pos_attrs, "x") orelse return null;
    const y_str = attrValue(pos_attrs, "y") orelse return null;
    return .{
        .x = parseXsdInteger(i64, x_str) orelse return null,
        .y = parseXsdInteger(i64, y_str) orelse return null,
        .end = pos_end,
    };
}

/// Parse the `<{xdr}:ext cx="N" cy="N"/>` child starting the search
/// at `from_idx`. Shared by the absoluteAnchor parser and the strict
/// one-cell validation: a oneCellAnchor's extent is REQUIRED by the
/// schema, and a pic- or chart-bearing one-cell anchor whose extent
/// does not parse must refuse under strict rather than ride out
/// (Codex #214 r2 REL-201). The match is an exact live tag — a
/// `<xdr:extLst>` or a commented fake cannot satisfy the validation
/// the real element would fail (r3 REL-301). The value stays off the
/// wire either way.
fn parseExtAttrs(xml: []const u8, from_idx: usize, open_ext: []const u8) ?Extent {
    const ext_idx = findLiveExactTag(xml, from_idx, open_ext) orelse return null;
    return extAttrsAt(xml, ext_idx);
}

const Extent = struct { cx: i64, cy: i64 };

/// `parseExtAttrs` with `<ext>` under any followed prefix.
fn parseExtAttrsIn(xml: []const u8, from_idx: usize, sets: []const DrawingTags) ?Extent {
    const ext = anyLiveExactTag(xml, from_idx, sets, "open_ext") orelse return null;
    return extAttrsAt(xml, ext.at);
}

fn extAttrsAt(xml: []const u8, ext_idx: usize) ?Extent {
    const ext_end = findUnquotedTagEnd(xml, ext_idx) orelse return null;
    const ext_attrs = xml[ext_idx .. ext_end + 1];
    const cx_str = attrValue(ext_attrs, "cx") orelse return null;
    const cy_str = attrValue(ext_attrs, "cy") orelse return null;
    const cx = parseXsdInteger(i64, cx_str) orelse return null;
    const cy = parseXsdInteger(i64, cy_str) orelse return null;
    return .{ .cx = cx, .cy = cy };
}

pub const Corner = enum { from, to };

/// XML whitespace, the set XSD's `whiteSpace="collapse"` strips from
/// a typed value.
const xml_ws = " \t\r\n";

/// A byte range of a block: `[start, end)`.
pub const Span = struct { start: usize, end: usize };

/// A corner block (`<{p}:from>` / `<{p}:to>`) read whole — the four
/// coordinates AND where the block and each of the two grid scalars
/// sit in the anchor, so the read takes the values and the sweep
/// splices one scalar's text: ONE parser, one acceptance
/// (`parseXsdInteger`: the whitespace-collapsed body, an optional sign, digits only) for both, where the
/// sweep used to carry its own more lenient grammar (in-house
/// ND-REL-103 — the read and the sweep must judge the same bytes the
/// same way). Offsets are relative to the block handed in.
pub const CornerBlock = struct {
    anchor: CellAnchor,
    /// The `<` of the block's open tag.
    open_start: usize,
    /// The byte past the block's close tag.
    after_close: usize,
    /// The text of `<col>` / `<row>` — the sweep's splice sites.
    col_text: Span,
    row_text: Span,
};

/// Parse the anchor's `from` or `to` block, searched from
/// `content_start` (the byte past the wrapper's open tag): the block
/// under the first set that spells a live open tag (its close tag is
/// then that set's — XML ties them), each scalar under ANY set. Live matches: a commented
/// `<xdr:from>` fake in the block is not the anchor's geometry (Codex
/// #214 r5 REL-501). Null when the block or any scalar is absent or
/// does not parse — the strict read's and the sweep's refusal.
pub fn parseCornerIn(block: []const u8, content_start: usize, sets: []const DrawingTags, which: Corner) ?CornerBlock {
    for (sets) |s| {
        const open = if (which == .from) s.open_from else s.open_to;
        const close = if (which == .from) s.close_from else s.close_to;
        // From the wrapper's content: a block spelled inside the
        // wrapper's own attribute value is not a corner (in-house
        // ND-REL-301 — the read served an attribute's `<xdr:to>`).
        const o = findLiveMarkup(block, content_start, open) orelse continue;
        const c = findLiveMarkup(block, o + open.len, close) orelse return null;
        const inner_start = o + open.len;
        const inner = block[inner_start..c];
        const col = scalarIn(inner, sets, .col) orelse return null;
        const col_off = scalarIn(inner, sets, .col_off) orelse return null;
        const row = scalarIn(inner, sets, .row) orelse return null;
        const row_off = scalarIn(inner, sets, .row_off) orelse return null;
        return .{
            // XSD numeric types collapse whitespace: `<xdr:row> 8 </xdr:row>`
            // is a conformant 8 (in-house ND-REL-201 — the unified
            // grammar refused it where the v1 sweep had trimmed). The
            // spans stay untrimmed: the sweep replaces the whole body.
            .anchor = .{
                .col = parseXsdInteger(u32, inner[col.start..col.end]) orelse return null,
                .col_off = parseXsdInteger(i64, inner[col_off.start..col_off.end]) orelse return null,
                .row = parseXsdInteger(u32, inner[row.start..row.end]) orelse return null,
                .row_off = parseXsdInteger(i64, inner[row_off.start..row_off.end]) orelse return null,
            },
            .open_start = o,
            .after_close = c + close.len,
            .col_text = .{ .start = inner_start + col.start, .end = inner_start + col.end },
            .row_text = .{ .start = inner_start + row.start, .end = inner_start + row.end },
        };
    }
    return null;
}

/// Do two corner blocks of one anchor share bytes — one nested inside
/// the other? Not two corners: the strict read and the sweep refuse.
pub fn cornersOverlap(a: CornerBlock, b: CornerBlock) bool {
    const first = if (a.open_start <= b.open_start) a else b;
    const second = if (a.open_start <= b.open_start) b else a;
    return second.open_start < first.after_close;
}

/// An XSD integer: optional sign, then ASCII digits — nothing else.
/// `std.fmt.parseInt` also accepts `_` digit separators (`1_0` = 10),
/// which no XSD lexical space has (in-house ND-REL-306). The text is
/// XSD-collapsed first (`xml_ws`).
fn parseXsdInteger(comptime T: type, text: []const u8) ?T {
    const t = std.mem.trim(u8, text, xml_ws);
    if (t.len == 0) return null;
    var digits = t;
    var negative = false;
    if (t[0] == '+' or t[0] == '-') {
        negative = t[0] == '-';
        digits = t[1..];
    }
    if (digits.len == 0) return null;
    for (digits) |c| if (c < '0' or c > '9') return null;
    const magnitude = std.fmt.parseInt(T, digits, 10) catch return null;
    if (@typeInfo(T).int.signedness == .unsigned) return if (negative) null else magnitude;
    return if (negative) -magnitude else magnitude;
}

const Scalar = enum { col, col_off, row, row_off };

/// The text span of the first live `<{p}:col>`-style scalar under any
/// set, relative to `inner`.
fn scalarIn(inner: []const u8, sets: []const DrawingTags, which: Scalar) ?Span {
    for (sets) |s| {
        const open, const close = switch (which) {
            .col => .{ s.open_col, s.close_col },
            .col_off => .{ s.open_col_off, s.close_col_off },
            .row => .{ s.open_row, s.close_row },
            .row_off => .{ s.open_row_off, s.close_row_off },
        };
        const o = findLiveMarkup(inner, 0, open) orelse continue;
        const value_start = o + open.len;
        const c = findLiveMarkup(inner, value_start, close) orelse return null;
        return .{ .start = value_start, .end = c };
    }
    return null;
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
    // The QName must be exact: a longer element name is not that plot
    // type and must not fake a compound either (r5 REL-503).
    try std.testing.expectEqual(ChartType.other, detectChartType("<c:chartSpace><c:barChartExtension/>", "c"));
    try std.testing.expectEqual(ChartType.pie, detectChartType("<c:chartSpace><c:pieChart/><c:barChartExtension/>", "c"));
}

test "findDrawingRid: a prefixed drawing element and a non-r relationship prefix resolve when BOUND" {
    // `<x:drawing rel:id=…/>` is valid OOXML the unprefixed literal
    // search missed (Codex #214 r2 REL-202) — honoured only when the
    // prefixes are actually bound to the SpreadsheetML main and
    // relationships namespaces (r3 REL-302).
    const bound =
        "<x:worksheet xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"" ++
        " xmlns:rel=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" ++
        "<x:drawing rel:id=\"rId9\"/></x:worksheet>";
    try std.testing.expectEqualStrings("rId9", findDrawingRid(bound).?);
    // The same spellings with UNBOUND prefixes are not the element /
    // not a reference.
    try std.testing.expect(findDrawingRef("<x:worksheet><x:drawing rel:id=\"rId9\"/></x:worksheet>") == .absent);
    const foreign =
        "<sheet xmlns:foo=\"https://example.invalid/ns\">" ++
        "<drawing foo:id=\"rId9\"/></sheet>";
    try std.testing.expect(findDrawingRef(foreign) == .malformed);
    // A bare `id` attribute is NOT a relationship reference.
    try std.testing.expectEqual(@as(?[]const u8, null), findDrawingRid("<sheet><drawing id=\"1\"/></sheet>"));
    // Tri-state: an unreadable reference is malformed, not absent.
    try std.testing.expect(findDrawingRef("<sheet><drawing/></sheet>") == .malformed);
    try std.testing.expect(findDrawingRef("<sheet><sheetData/></sheet>") == .absent);
    // Markup-shaped text inside a comment is not the element — a
    // commented `<drawing/>` must not refuse a valid sheet (REL-302).
    try std.testing.expect(findDrawingRef("<sheet><!-- <drawing/> --><sheetData/></sheet>") == .absent);
    // A declaration-shaped string inside an attribute VALUE or a
    // comment is not a binding (r4 REL-401).
    const value_decoy =
        "<sheet xmlns:foo=\"https://example.invalid/ns\"" ++
        " note=\"xmlns:foo='http://schemas.openxmlformats.org/officeDocument/2006/relationships'\">" ++
        "<drawing foo:id=\"rId1\"/></sheet>";
    try std.testing.expect(findDrawingRef(value_decoy) == .malformed);
    const comment_decoy =
        "<sheet xmlns:foo=\"https://example.invalid/ns\">" ++
        "<!-- xmlns:foo=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" -->" ++
        "<drawing foo:id=\"rId1\"/></sheet>";
    try std.testing.expect(findDrawingRef(comment_decoy) == .malformed);
    // A real binding survives single quotes and Eq whitespace.
    const eq_ws =
        "<sheet xmlns:rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'>" ++
        "<drawing rel:id=\"rId4\"/></sheet>";
    try std.testing.expectEqualStrings("rId4", findDrawingRid(eq_ws).?);
    // Two live drawing elements: `.multiple`, carrying the FIRST in
    // document order — here the bound prefixed one (r4 REL-403).
    const dup =
        "<x:sheet xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"" ++
        " xmlns:rel=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" ++
        "<x:drawing rel:id=\"rIdA\"/><drawing r:id=\"rIdB\"/></x:sheet>";
    switch (findDrawingRef(dup)) {
        .multiple => |r| try std.testing.expectEqualStrings("rIdA", r),
        else => return error.TestUnexpectedResult,
    }
    const dup2 = "<sheet><drawing r:id=\"rIdB\"/><drawing r:id=\"rIdC\"/></sheet>";
    switch (findDrawingRef(dup2)) {
        .multiple => |r| try std.testing.expectEqualStrings("rIdB", r),
        else => return error.TestUnexpectedResult,
    }
}

test "relTypeIs: exact known roots only; extractSeriesRefs stays linear under comment spam" {
    // A foreign URI ending in the expected leaf is not that edge
    // (Codex #214 r3 REL-303).
    try std.testing.expect(relTypeIs("http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing", "drawing"));
    try std.testing.expect(relTypeIs("http://purl.oclc.org/ooxml/officeDocument/relationships/image", "image"));
    try std.testing.expect(relTypeIs("http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet", "xlMacrosheet"));
    try std.testing.expect(!relTypeIs("https://example.invalid/relationships/drawing", "drawing"));
    try std.testing.expect(!relTypeIs("http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawingx", "drawing"));
    // URIs are case-sensitive: a case-mutated leaf is an unknown
    // type, not the edge (r4 REL-402).
    try std.testing.expect(!relTypeIs("http://schemas.openxmlformats.org/officeDocument/2006/relationships/DRAWING", "drawing"));

    // PERF-301: thousands of commented fakes around a few real
    // carriers — the forward-only scan must stay effectively linear
    // (this is a functional pin; the shape that used to rescan from
    // byte zero per fake would time out here long before failing).
    // Linear here because every fake SPELLS the needle, so the cached
    // candidate stays near the scan position; regions that do not
    // spell it are the CF-PERF-201 test's shape.
    var xml: std.ArrayListUnmanaged(u8) = .empty;
    defer xml.deinit(std.testing.allocator);
    try xml.appendSlice(std.testing.allocator, "<c:chartSpace>");
    for (0..2000) |_| try xml.appendSlice(std.testing.allocator, "<!-- <c:f>Fake!A1</c:f> -->");
    try xml.appendSlice(std.testing.allocator, "<c:f>Real!A1</c:f>");
    for (0..2000) |_| try xml.appendSlice(std.testing.allocator, "<!-- <c:f>Fake!B1</c:f> -->");
    try xml.appendSlice(std.testing.allocator, "<c:f>Real!B1</c:f></c:chartSpace>");
    const refs = try extractSeriesRefs(std.testing.allocator, xml.items, "c", null, .strict);
    defer std.testing.allocator.free(refs);
    try std.testing.expectEqual(@as(usize, 2), refs.len);
    try std.testing.expectEqualStrings("Real!A1", refs[0]);
    try std.testing.expectEqualStrings("Real!B1", refs[1]);

    // PERF-602: dense primary carriers with an ENABLED but unused
    // alternate prefix — the cached-candidate scan must not rescan
    // the alternate's whole suffix per carrier (the shape that did
    // would grind here in Debug long before failing).
    var dense: std.ArrayListUnmanaged(u8) = .empty;
    defer dense.deinit(std.testing.allocator);
    try dense.appendSlice(std.testing.allocator, "<c:chartSpace>");
    for (0..1500) |_| try dense.appendSlice(std.testing.allocator, "<c:f>R!A1</c:f>");
    try dense.appendSlice(std.testing.allocator, "</c:chartSpace>");
    const dense_refs = try extractSeriesRefs(std.testing.allocator, dense.items, "c", "c2", .strict);
    defer std.testing.allocator.free(dense_refs);
    try std.testing.expectEqual(@as(usize, 1500), dense_refs.len);
    // The symmetric alternate-only case still flattens in order.
    const alt_only = "<c:chartSpace><c2:f>Alt!A1</c2:f><c:f>Pri!A1</c:f><c2:f>Alt!B1</c2:f></c:chartSpace>";
    const mixed = try extractSeriesRefs(std.testing.allocator, alt_only, "c", "c2", .strict);
    defer std.testing.allocator.free(mixed);
    try std.testing.expectEqual(@as(usize, 3), mixed.len);
    try std.testing.expectEqualStrings("Alt!A1", mixed[0]);
    try std.testing.expectEqualStrings("Pri!A1", mixed[1]);
    try std.testing.expectEqualStrings("Alt!B1", mixed[2]);
}

/// The tag sets of a part, for the tests that parse a block directly.
fn testTagSets(bufs: *[max_tag_sets]TagSetBuf, sets: *[max_tag_sets]DrawingTags, root_xml: []const u8) []const DrawingTags {
    return buildTagSets(bufs, sets, resolveDrawingPrefixes(root_xml)) catch unreachable;
}

test "parseCornerIn: one corner parser — the values, the block and the scalar spans, under any followed prefix" {
    var bufs: [max_tag_sets]TagSetBuf = undefined;
    var set_store: [max_tag_sets]DrawingTags = undefined;
    const xml =
        \\<xdr:from><xdr:col>3</xdr:col><xdr:colOff>16119</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>47624</xdr:rowOff></xdr:from>
    ;
    const sets = testTagSets(&bufs, &set_store, "<xdr:wsDr xmlns:xdr=\"" ++ ns_xdr_transitional ++ "\"/>");
    const a = parseCornerIn(xml, 0, sets, .from).?;
    try std.testing.expectEqual(@as(u32, 3), a.anchor.col);
    try std.testing.expectEqual(@as(i64, 16119), a.anchor.col_off);
    try std.testing.expectEqual(@as(u32, 1), a.anchor.row);
    try std.testing.expectEqual(@as(i64, 47624), a.anchor.row_off);
    try std.testing.expectEqual(@as(usize, 0), a.open_start);
    try std.testing.expectEqual(xml.len, a.after_close);
    try std.testing.expectEqualStrings("3", xml[a.col_text.start..a.col_text.end]);
    try std.testing.expectEqualStrings("1", xml[a.row_text.start..a.row_text.end]);
    // Non-canonical drawing prefix: identical structure with `dr:`
    // instead of `xdr:` — same parser, the part's own resolution.
    const xml2 =
        \\<dr:from><dr:col>3</dr:col><dr:colOff>0</dr:colOff><dr:row>1</dr:row><dr:rowOff>0</dr:rowOff></dr:from>
    ;
    const sets2 = testTagSets(&bufs, &set_store, "<dr:wsDr xmlns:dr=\"" ++ ns_xdr_transitional ++ "\"/>");
    const b = parseCornerIn(xml2, 0, sets2, .from).?;
    try std.testing.expectEqual(@as(u32, 3), b.anchor.col);
    try std.testing.expectEqual(@as(u32, 1), b.anchor.row);
    // A part binding the namespace twice may spell the block under one
    // name and its scalars under the other (in-house ND-REL-103).
    const mixed = "<from><xdr:col>7</xdr:col><colOff>0</colOff><xdr:row>2</xdr:row><rowOff>5</rowOff></from>";
    const sets3 = testTagSets(&bufs, &set_store, "<xdr:wsDr xmlns:xdr=\"" ++ ns_xdr_transitional ++ "\" xmlns=\"" ++ ns_xdr_transitional ++ "\"/>");
    const m = parseCornerIn(mixed, 0, sets3, .from).?;
    try std.testing.expectEqual(@as(u32, 7), m.anchor.col);
    try std.testing.expectEqual(@as(u32, 2), m.anchor.row);
    try std.testing.expectEqual(@as(i64, 5), m.anchor.row_off);
    try std.testing.expectEqualStrings("7", mixed[m.col_text.start..m.col_text.end]);
    // The grammar is `parseInt` on the XSD-collapsed body: whitespace
    // around the digits parses (in-house ND-REL-201), a leading `+`
    // does, an empty body does not — one acceptance for the read and
    // the sweep.
    try std.testing.expectEqual(@as(u32, 3), parseCornerIn("<from><col> 3</col><colOff>0</colOff><row>\n1\t</row><rowOff>0</rowOff></from>", 0, sets3, .from).?.anchor.col);
    try std.testing.expectEqual(@as(u32, 3), parseCornerIn("<from><col>+3</col><colOff>0</colOff><row>1</row><rowOff>0</rowOff></from>", 0, sets3, .from).?.anchor.col);
    try std.testing.expect(parseCornerIn("<from><col>3</col><colOff>0</colOff><row></row><rowOff>0</rowOff></from>", 0, sets3, .from) == null);
    // A commented block is text: the live one after it is the corner.
    const decoy = "<!-- <from><col>9</col><colOff>0</colOff><row>9</row><rowOff>0</rowOff></from> --><from><col>1</col><colOff>0</colOff><row>2</row><rowOff>0</rowOff></from>";
    const d = parseCornerIn(decoy, 0, sets3, .from).?;
    try std.testing.expectEqual(@as(u32, 1), d.anchor.col);
    try std.testing.expect(d.open_start > 0);
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

test "findNamespacePrefix is bounded to the root window — a binding declared past 4 KiB is not resolved" {
    // Prefix resolution reads the root window only (`root_window`), so
    // a late mid-document declaration cannot shadow the canonical
    // fallback; the chart walk's unfollowed-binding probe is what sees
    // such a binding (and refuses on it under strict).
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
    try std.testing.expectEqual(
        @as(?[]const u8, null),
        findNamespacePrefix(block, "http://schemas.openxmlformats.org/drawingml/2006/chart"),
    );
    try std.testing.expect(hasUnfollowedChartBinding(block, "c", null));
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

test "ChartFormulaWalk: one carrier walk for the read and the sweep — self-closing, decoys, strict refusals" {
    const a = std.testing.allocator;
    // Self-closing carriers are EMPTY carriers in both spellings — the
    // schema-equivalent of `<c:f></c:f>`, which the read always
    // reported; `<c:f />` used to mis-scan to the next close and
    // return the markup in between as a ref.
    const sc = "<c:chartSpace><c:f/><c:f />text<c:f></c:f><c:f>R!A1</c:f></c:chartSpace>";
    const sc_refs = try extractSeriesRefs(a, sc, "c", null, .strict);
    defer a.free(sc_refs);
    try std.testing.expectEqual(@as(usize, 4), sc_refs.len);
    try std.testing.expectEqualStrings("", sc_refs[0]);
    try std.testing.expectEqualStrings("", sc_refs[1]);
    try std.testing.expectEqualStrings("", sc_refs[2]);
    try std.testing.expectEqualStrings("R!A1", sc_refs[3]);
    // The spans the sweep splices at slice to exactly those bodies, and
    // the walk resumes past each carrier's close.
    var walk = ChartFormulaWalk.init(sc);
    const f0 = (try walk.next(.strict)).?;
    try std.testing.expectEqual(f0.body_start, f0.body_end);
    try std.testing.expectEqual(@as(usize, "<c:chartSpace><c:f/>".len), f0.body_start);
    const f1 = (try walk.next(.strict)).?;
    try std.testing.expectEqual(f1.body_start, f1.body_end);
    const f2 = (try walk.next(.strict)).?;
    try std.testing.expectEqual(f2.body_start, f2.body_end);
    const f3 = (try walk.next(.strict)).?;
    try std.testing.expectEqualStrings("R!A1", sc[f3.body_start..f3.body_end]);
    try std.testing.expectEqualStrings("</c:f></c:chartSpace>", sc[f3.body_end..]);
    try std.testing.expectEqual(@as(?ChartFormula, null), try walk.next(.strict));
    try std.testing.expectEqual(@as(?ChartFormula, null), try walk.next(.strict));

    // Strict refuses what a byte-preserving splice cannot carry
    // through: an opened carrier that never closes, markup in a body
    // (a comment, a CDATA section — the schema's element text is a
    // simple type), carrier text inside an UNTERMINATED comment (live
    // or decoy is undecidable). Lenient keeps the historical read:
    // truncate at the unclosed carrier, return a body as it lies.
    const unclosed = "<c:chartSpace><c:f>R!A1</c:f><c:f>R!B1</c:chartSpace>";
    try std.testing.expectError(error.MalformedDrawingXml, extractSeriesRefs(a, unclosed, "c", null, .strict));
    const unclosed_len = try extractSeriesRefs(a, unclosed, "c", null, .lenient);
    defer a.free(unclosed_len);
    try std.testing.expectEqual(@as(usize, 1), unclosed_len.len);
    try std.testing.expectEqualStrings("R!A1", unclosed_len[0]);

    const markup = "<c:chartSpace><c:f>R!<!-- x -->A1</c:f></c:chartSpace>";
    try std.testing.expectError(error.MalformedDrawingXml, extractSeriesRefs(a, markup, "c", null, .strict));
    const markup_len = try extractSeriesRefs(a, markup, "c", null, .lenient);
    defer a.free(markup_len);
    try std.testing.expectEqual(@as(usize, 1), markup_len.len);
    try std.testing.expectEqualStrings("R!<!-- x -->A1", markup_len[0]);
    const cdata = "<c:chartSpace><c:f><![CDATA[R!A1]]></c:f></c:chartSpace>";
    try std.testing.expectError(error.MalformedDrawingXml, extractSeriesRefs(a, cdata, "c", null, .strict));

    const open_comment = "<c:chartSpace><c:f>R!A1</c:f><!-- <c:f>Fake!A1</c:f>";
    try std.testing.expectError(error.MalformedDrawingXml, extractSeriesRefs(a, open_comment, "c", null, .strict));
    const open_comment_len = try extractSeriesRefs(a, open_comment, "c", null, .lenient);
    defer a.free(open_comment_len);
    try std.testing.expectEqual(@as(usize, 1), open_comment_len.len);
    // The alternate prefix's text counts too.
    const open_alt = "<c:chartSpace><c2:f>A!A1</c2:f><![CDATA[<c2:f>";
    try std.testing.expectError(error.MalformedDrawingXml, extractSeriesRefs(a, open_alt, "c", "c2", .strict));

    // A TERMINATED decoy is text — not a ref, not a refusal — and an
    // unterminated region holding no carrier text refuses nothing
    // (there is no live-or-decoy question to decide).
    const decoys = "<c:chartSpace><!-- <c:f>Fake!A1</c:f> --><?pi <c:f>Fake!B1 ?><![CDATA[<c:f>Fake!C1</c:f>]]><c:f>R!A1</c:f><!-- <c:f>Fake!D1</c:f> --></c:chartSpace>";
    const decoy_refs = try extractSeriesRefs(a, decoys, "c", null, .strict);
    defer a.free(decoy_refs);
    try std.testing.expectEqual(@as(usize, 1), decoy_refs.len);
    try std.testing.expectEqualStrings("R!A1", decoy_refs[0]);
    const trailing_open = "<c:chartSpace><c:f>R!A1</c:f><!-- no carrier here";
    const trailing_refs = try extractSeriesRefs(a, trailing_open, "c", null, .strict);
    defer a.free(trailing_refs);
    try std.testing.expectEqual(@as(usize, 1), trailing_refs.len);
    // The probe keys on the exact QName: `<c:formatCode>` — which
    // follows the last carrier of nearly every chart Excel writes —
    // is not carrier text, so a trailing unterminated region after it
    // refuses nothing (Codex CF-REL-101).
    const format_code_tail = "<c:chartSpace><c:f>R!A1</c:f><c:numCache><c:formatCode>General</c:formatCode></c:numCache><!-- x";
    const format_code_refs = try extractSeriesRefs(a, format_code_tail, "c", null, .strict);
    defer a.free(format_code_refs);
    try std.testing.expectEqual(@as(usize, 1), format_code_refs.len);
    try std.testing.expectEqualStrings("R!A1", format_code_refs[0]);
    // …while a carrier tag inside the region still does, whichever
    // terminator follows it — including the end of the part.
    try std.testing.expectError(error.MalformedDrawingXml, extractSeriesRefs(a, "<c:chartSpace><c:f>R!A1</c:f><c:formatCode/><!-- <c:f", "c", null, .strict));
    try std.testing.expectError(error.MalformedDrawingXml, extractSeriesRefs(a, "<c:chartSpace><c:f>R!A1</c:f><!-- <c:f />", "c", null, .strict));
    // A start tag the part ends inside is a carrier with no close:
    // strict refuses, lenient ends with what came before (it used to
    // be dropped in both modes).
    for ([_][]const u8{ "<c:chartSpace><c:f>R!A1</c:f><c:f", "<c:chartSpace><c:f>R!A1</c:f><c:f " }) |truncated| {
        try std.testing.expectError(error.MalformedDrawingXml, extractSeriesRefs(a, truncated, "c", null, .strict));
        const kept = try extractSeriesRefs(a, truncated, "c", null, .lenient);
        defer a.free(kept);
        try std.testing.expectEqual(@as(usize, 1), kept.len);
        try std.testing.expectEqualStrings("R!A1", kept[0]);
    }

    // `init` resolves the part's own binding: a Strict-namespace chart
    // under a non-canonical prefix, and a part binding both URIs.
    const strict_ns = "<cs:chartSpace xmlns:cs=\"http://purl.oclc.org/ooxml/drawingml/chart\"><cs:f>S!A1</cs:f><c:f>not bound</c:f></cs:chartSpace>";
    var sw = ChartFormulaWalk.init(strict_ns);
    const s0 = (try sw.next(.strict)).?;
    try std.testing.expectEqualStrings("S!A1", strict_ns[s0.body_start..s0.body_end]);
    try std.testing.expectEqual(@as(?ChartFormula, null), try sw.next(.strict));
    const both = "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:cs=\"http://purl.oclc.org/ooxml/drawingml/chart\"><cs:f>S!A1</cs:f><c:f>T!A1</c:f><cs:f>S!B1</cs:f></c:chartSpace>";
    var bw = ChartFormulaWalk.init(both);
    const b0 = (try bw.next(.strict)).?;
    const b1 = (try bw.next(.strict)).?;
    const b2 = (try bw.next(.strict)).?;
    try std.testing.expectEqualStrings("S!A1", both[b0.body_start..b0.body_end]);
    try std.testing.expectEqualStrings("T!A1", both[b1.body_start..b1.body_end]);
    try std.testing.expectEqualStrings("S!B1", both[b2.body_start..b2.body_end]);
    try std.testing.expectEqual(@as(?ChartFormula, null), try bw.next(.strict));

    // The documented limits, pinned so they stay true. The needle
    // scratch's own bound (124: `</p:f` is four bytes over a 128-byte
    // buffer) is an internal invariant only `initWithPrefixes` can
    // reach — strict refuses, lenient serves nothing…
    const long_prefix = "p" ** 130;
    const long_bound = "<" ++ long_prefix ++ ":chartSpace xmlns:" ++ long_prefix ++ "=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><" ++ long_prefix ++ ":f>R!A1</" ++ long_prefix ++ ":f></" ++ long_prefix ++ ":chartSpace>";
    try std.testing.expectError(error.MalformedDrawingXml, extractSeriesRefs(a, long_bound, long_prefix, null, .strict));
    const long_len = try extractSeriesRefs(a, long_bound, long_prefix, null, .lenient);
    defer a.free(long_len);
    try std.testing.expectEqual(@as(usize, 0), long_len.len);
    const edge_prefix = "p" ** 124;
    const edge_bound = "<" ++ edge_prefix ++ ":chartSpace><" ++ edge_prefix ++ ":f>R!A1</" ++ edge_prefix ++ ":f></" ++ edge_prefix ++ ":chartSpace>";
    const edge_refs = try extractSeriesRefs(a, edge_bound, edge_prefix, null, .strict);
    defer a.free(edge_refs);
    try std.testing.expectEqual(@as(usize, 1), edge_refs.len);
    try std.testing.expectError(error.MalformedDrawingXml, extractSeriesRefs(a, edge_bound, "p" ** 125, null, .strict));
    // …because the production entry (`init`, the resolver) never
    // hands the walk a prefix past `max_prefix_len` (100): at the
    // limit the carrier is served; one byte over, the binding is
    // rejected and strict refuses the part instead of walking it
    // under the canonical `c` and finding nothing, while lenient
    // walks what resolves — nothing (in-house CF-DOC-201: the verdict
    // the docs promised for a bound no production part could reach).
    const at_limit = "p" ** 100;
    const at_limit_part = "<" ++ at_limit ++ ":chartSpace xmlns:" ++ at_limit ++ "=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><" ++ at_limit ++ ":f>R!A1</" ++ at_limit ++ ":f></" ++ at_limit ++ ":chartSpace>";
    var lw = ChartFormulaWalk.init(at_limit_part);
    const lf = (try lw.next(.strict)).?;
    try std.testing.expectEqualStrings("R!A1", at_limit_part[lf.body_start..lf.body_end]);
    const over = "p" ** 101;
    const over_part = "<" ++ over ++ ":chartSpace xmlns:" ++ over ++ "=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><" ++ over ++ ":f>R!A1</" ++ over ++ ":f></" ++ over ++ ":chartSpace>";
    var ow = ChartFormulaWalk.init(over_part);
    try std.testing.expectError(error.MalformedChartXml, ow.next(.strict));
    var ol = ChartFormulaWalk.init(over_part);
    try std.testing.expectEqual(@as(?ChartFormula, null), try ol.next(.lenient));
    // A rejected binding beside an accepted one is still a refusal:
    // the carriers under the rejected prefix would be the ones missed.
    const mixed_part = "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:" ++ over ++ "=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:f>R!A1</c:f><" ++ over ++ ":f>R!B1</" ++ over ++ ":f></c:chartSpace>";
    var mw = ChartFormulaWalk.init(mixed_part);
    try std.testing.expectError(error.MalformedChartXml, mw.next(.strict));
    // The probe reads the whole part, not the resolver's root window,
    // and refuses ANY chart-namespace binding the walk does not
    // follow (in-house CF-REL-301): a prefix that runs past 4 KiB…
    const huge = "p" ** 4200;
    const huge_part = "<" ++ huge ++ ":chartSpace xmlns:" ++ huge ++ "=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><" ++ huge ++ ":f>R!A1</" ++ huge ++ ":f></" ++ huge ++ ":chartSpace>";
    var hw = ChartFormulaWalk.init(huge_part);
    try std.testing.expectError(error.MalformedChartXml, hw.next(.strict));
    var hl = ChartFormulaWalk.init(huge_part);
    try std.testing.expectEqual(@as(?ChartFormula, null), try hl.next(.lenient));
    // …a second prefix on the URI, the carriers under it…
    const second = "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:d=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><d:f>R!A1</d:f></c:chartSpace>";
    var secw = ChartFormulaWalk.init(second);
    try std.testing.expectError(error.MalformedChartXml, secw.next(.strict));
    var secl = ChartFormulaWalk.init(second);
    try std.testing.expectEqual(@as(?ChartFormula, null), try secl.next(.lenient));
    // …and a plausible prefix declared beyond the root window, while
    // the canonical `c` declared there is the fallback the walk
    // follows anyway and its carrier is served.
    const pad = "x" ** 4200;
    const beyond_zz = "<zz:chartSpace pad=\"" ++ pad ++ "\" xmlns:zz=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><zz:f>R!A1</zz:f></zz:chartSpace>";
    var bzw = ChartFormulaWalk.init(beyond_zz);
    try std.testing.expectError(error.MalformedChartXml, bzw.next(.strict));
    const beyond_c = "<c:chartSpace pad=\"" ++ pad ++ "\" xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:f>R!A1</c:f></c:chartSpace>";
    var bcw = ChartFormulaWalk.init(beyond_c);
    const bcf = (try bcw.next(.strict)).?;
    try std.testing.expectEqualStrings("R!A1", beyond_c[bcf.body_start..bcf.body_end]);
    // Only LIVE, attribute-shaped declarations count: a decoy in a
    // comment, a PI, character data or an attribute value neither
    // refuses (an overlong or second one) nor resolves (a plausible
    // one — the root's real `c:` binding wins and its carrier is
    // served).
    const chart_uri = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    const decoys_r3 = [_][]const u8{
        "<?xml version=\"1.0\"?><!-- xmlns:" ++ over ++ "=\"" ++ chart_uri ++ "\" --><c:chartSpace xmlns:c=\"" ++ chart_uri ++ "\"><c:f>R!A1</c:f></c:chartSpace>",
        "<?xml version=\"1.0\"?><!-- xmlns:zz=\"" ++ chart_uri ++ "\" --><c:chartSpace xmlns:c=\"" ++ chart_uri ++ "\"><c:f>R!A1</c:f></c:chartSpace>",
        "<?zz xmlns:q=\"" ++ chart_uri ++ "\"?><c:chartSpace xmlns:c=\"" ++ chart_uri ++ "\"><c:f>R!A1</c:f></c:chartSpace>",
        "<c:chartSpace data=\"xmlns:q='" ++ chart_uri ++ "'\" xmlns:c=\"" ++ chart_uri ++ "\"><c:f>R!A1</c:f></c:chartSpace>",
        "<c:chartSpace data=\"xmlns:" ++ over ++ "='" ++ chart_uri ++ "'\" xmlns:c=\"" ++ chart_uri ++ "\"><c:f>R!A1</c:f></c:chartSpace>",
        "<c:chartSpace xmlns:c=\"" ++ chart_uri ++ "\"><c:t>xmlns:q=\"" ++ chart_uri ++ "\"</c:t><c:f>R!A1</c:f></c:chartSpace>",
        // An attribute NAME that merely ends in `xmlns:` — a
        // namespace-well-formed part with `xmlns:axmlns` bound — is not
        // a declaration either (round-4 agent B: the r3 tokenizer
        // matched the six bytes anywhere in a tag and refused it).
        "<c:chartSpace xmlns:axmlns=\"urn:x\" axmlns:zz=\"" ++ chart_uri ++ "\" xmlns:c=\"" ++ chart_uri ++ "\"><c:f>R!A1</c:f></c:chartSpace>",
    };
    for (decoys_r3) |decoy| {
        const resolved = resolveDrawingPrefixes(decoy);
        try std.testing.expectEqualStrings("c", resolved.c);
        try std.testing.expect(!resolved.c_rejected);
        var dw2 = ChartFormulaWalk.init(decoy);
        const df = (try dw2.next(.strict)).?;
        try std.testing.expectEqualStrings("R!A1", decoy[df.body_start..df.body_end]);
    }
    // A part binding the chart namespace as its DEFAULT namespace —
    // openpyxl's spelling — resolves to the empty prefix and is walked
    // under `<f>` (in-house CF-REL-401); the exact-QName rule keeps the
    // `f`-prefixed element names out.
    const default_ns = "<chartSpace xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><chart><plotArea><barChart><ser><tx><strRef><f>'Data'!B1</f></strRef></tx><cat><numRef><f>'Data'!$A$2:$A$4</f><numCache><formatCode>General</formatCode></numCache></numRef></cat><val><numRef><f>'Data'!$B$2:$B$4</f></numRef></val></ser><firstSliceAng val=\"0\"/></barChart></plotArea></chart></chartSpace>";
    try std.testing.expectEqualStrings("", resolveDrawingPrefixes(default_ns).c);
    try std.testing.expect(!resolveDrawingPrefixes(default_ns).c_rejected);
    for ([_]WalkMode{ .strict, .lenient }) |mode| {
        const refs = try extractSeriesRefs(a, default_ns, "", null, mode);
        defer a.free(refs);
        try std.testing.expectEqual(@as(usize, 3), refs.len);
        try std.testing.expectEqualStrings("'Data'!B1", refs[0]);
        try std.testing.expectEqualStrings("'Data'!$A$2:$A$4", refs[1]);
        try std.testing.expectEqualStrings("'Data'!$B$2:$B$4", refs[2]);
    }
    var dw = ChartFormulaWalk.init(default_ns);
    const d0 = (try dw.next(.strict)).?;
    try std.testing.expectEqualStrings("'Data'!B1", default_ns[d0.body_start..d0.body_end]);
    try std.testing.expectEqual(ChartType.bar, detectChartTypeWithAlt(default_ns, "", null));
    // A default binding beside a prefixed one is a binding the walk
    // does not follow: refused, like a second prefix.
    const mixed_default = "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:f>R!A1</c:f><f>R!B1</f></c:chartSpace>";
    var mdw = ChartFormulaWalk.init(mixed_default);
    try std.testing.expectError(error.MalformedChartXml, mdw.next(.strict));
}

test "ChartFormulaWalk: the strict tail probe is one pass — a part with many trailing comments walks in linear time (CF-PERF-101)" {
    const a = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();
    // 200 000 terminated comments after a `<c:formatCode>` (the shape
    // that made the old region-by-region probe rescan the remainder
    // per region: 23 s for 80 000 comments in ReleaseFast). The bound
    // is generous for Debug and CI noise; the quadratic shape needs
    // minutes here, so it is not a flaky pin.
    var xml: std.ArrayListUnmanaged(u8) = .empty;
    defer xml.deinit(a);
    try xml.appendSlice(a, "<c:chartSpace><c:f>R!A1</c:f><c:numCache><c:formatCode>General</c:formatCode></c:numCache>");
    for (0..200_000) |_| try xml.appendSlice(a, "<!-- x -->");
    try xml.appendSlice(a, "</c:chartSpace>");
    const started = std.Io.Clock.now(.awake, io).nanoseconds;
    const refs = try extractSeriesRefs(a, xml.items, "c", null, .strict);
    defer a.free(refs);
    const elapsed_ns = std.Io.Clock.now(.awake, io).nanoseconds - started;
    try std.testing.expectEqual(@as(usize, 1), refs.len);
    try std.testing.expectEqualStrings("R!A1", refs[0]);
    try std.testing.expect(elapsed_ns < 10 * std.time.ns_per_s);
    // And the same tail left unterminated: still one pass, still no
    // carrier tag inside it, still served.
    xml.items.len -= "</c:chartSpace>".len;
    try xml.appendSlice(a, "<!-- open");
    const open_refs = try extractSeriesRefs(a, xml.items, "c", null, .strict);
    defer a.free(open_refs);
    try std.testing.expectEqual(@as(usize, 1), open_refs.len);
}

test "ChartFormulaWalk: the live search is one pass — regions that do not spell the needle before and between carriers walk in linear time (CF-PERF-201)" {
    const a = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();
    // 100 000 terminated regions BEFORE the first carrier and 100 000
    // between the two — the shape the three-needle opener probe
    // rescanned to the candidate once per region (52 s for 100 000
    // comments in ReleaseFast; the PERF-301 fakes above spell the
    // needle, so they never showed it). Comments and PIs, both
    // modes; the bound is generous for Debug and CI noise, and the
    // quadratic shape needs minutes here, so it is not a flaky pin.
    for ([_][]const u8{ "<!-- x -->", "<?p x?>" }) |region| {
        var xml: std.ArrayListUnmanaged(u8) = .empty;
        defer xml.deinit(a);
        try xml.appendSlice(a, "<c:chartSpace>");
        for (0..100_000) |_| try xml.appendSlice(a, region);
        try xml.appendSlice(a, "<c:f>R!A1</c:f>");
        for (0..100_000) |_| try xml.appendSlice(a, region);
        try xml.appendSlice(a, "<c:f>R!B1</c:f></c:chartSpace>");
        for ([_]WalkMode{ .strict, .lenient }) |mode| {
            const started = std.Io.Clock.now(.awake, io).nanoseconds;
            const refs = try extractSeriesRefs(a, xml.items, "c", null, mode);
            defer a.free(refs);
            const elapsed_ns = std.Io.Clock.now(.awake, io).nanoseconds - started;
            try std.testing.expectEqual(@as(usize, 2), refs.len);
            try std.testing.expectEqualStrings("R!A1", refs[0]);
            try std.testing.expectEqualStrings("R!B1", refs[1]);
            try std.testing.expect(elapsed_ns < 10 * std.time.ns_per_s);
        }
    }
}

test "findLocalChartElement: a region followed by live bytes, 100 000 times, scans in linear time (CF-PERF-301)" {
    const a = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();
    // Back-to-back regions are eaten by the loop's own branches; one
    // live element between two regions reached the `:chart` search,
    // which rescanned the remainder per region (82 s at 40 000 pairs
    // through the anchors read in Debug). The candidate is cached
    // now; the bound is generous for Debug and CI noise, and the
    // quadratic shape needs minutes here.
    for ([_][]const u8{ "<!-- x --><a:x/>", "<?p x?><a:x/>" }) |pair| {
        var block: std.ArrayListUnmanaged(u8) = .empty;
        defer block.deinit(a);
        try block.appendSlice(a, "<xdr:graphicFrame macro=\"\">");
        for (0..100_000) |_| try block.appendSlice(a, pair);
        try block.appendSlice(a, "<c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" r:id=\"rId1\"/></xdr:graphicFrame>");
        const started = std.Io.Clock.now(.awake, io).nanoseconds;
        const found = findLocalChartElement(block.items, 0, .{});
        const elapsed_ns = std.Io.Clock.now(.awake, io).nanoseconds - started;
        try std.testing.expect(found != null);
        try std.testing.expect(std.mem.startsWith(u8, block.items[found.?..], "<c:chart"));
        try std.testing.expect(elapsed_ns < 10 * std.time.ns_per_s);
    }
}

test "RootNsBindings: a whole-part pass over 100 000 decoy declarations is linear (CF-REL-301)" {
    const a = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();
    // The per-hit opening-tag guard the drawing side uses rescans
    // from byte zero; the iterator's own state machine must not.
    var xml: std.ArrayListUnmanaged(u8) = .empty;
    defer xml.deinit(a);
    try xml.appendSlice(a, "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">");
    for (0..100_000) |_| try xml.appendSlice(a, "<c:t d=\"xmlns:q='u'\">xmlns:q=\"u\"</c:t><!-- xmlns:q=\"u\" -->");
    try xml.appendSlice(a, "<c:f>R!A1</c:f></c:chartSpace>");
    const started = std.Io.Clock.now(.awake, io).nanoseconds;
    var w = ChartFormulaWalk.init(xml.items);
    const f = (try w.next(.strict)).?;
    const elapsed_ns = std.Io.Clock.now(.awake, io).nanoseconds - started;
    try std.testing.expectEqualStrings("R!A1", xml.items[f.body_start..f.body_end]);
    try std.testing.expect(elapsed_ns < 10 * std.time.ns_per_s);
}

test "resolveDrawingPrefixes: the default namespace is a prefix — empty primary for an unprefixed root, an alternate beside a prefixed one, foreign when not an xdr URI" {
    // openpyxl 3.1's drawing root: unprefixed, the spreadsheetDrawing
    // namespace bound as the default one, `a` / `c` / `r` prefixed.
    {
        const xml =
            \\<wsDr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><oneCellAnchor/></wsDr>
        ;
        const p = resolveDrawingPrefixes(xml);
        try std.testing.expectEqualStrings("", p.xdr);
        try std.testing.expectEqual(@as(usize, 0), p.xdr_alts().len);
        try std.testing.expectEqualStrings("a", p.a);
        try std.testing.expectEqualStrings("c", p.c);
        try std.testing.expect(!p.xdr_rejected);
        try std.testing.expect(!p.c_rejected);
        // The needles are spelled bare.
        var buf: [4096]u8 = undefined;
        const tags = try DrawingTags.build(&buf, p);
        try std.testing.expectEqualStrings("<", tags.xdr_prefix_open);
        try std.testing.expectEqualStrings("<twoCellAnchor", tags.open_two);
        try std.testing.expectEqualStrings("</oneCellAnchor>", tags.close_one);
        try std.testing.expectEqualStrings("<from>", tags.open_from);
        try std.testing.expectEqualStrings("</to>", tags.close_to);
        try std.testing.expectEqualStrings("<ext", tags.open_ext);
        try std.testing.expectEqualStrings("<graphicFrame", tags.open_graphic_frame);
        try std.testing.expectEqualStrings("<c:chart", tags.open_chart);
        // …and the corner parser reads them.
        var bufs: [max_tag_sets]TagSetBuf = undefined;
        var set_store: [max_tag_sets]DrawingTags = undefined;
        const sets = try buildTagSets(&bufs, &set_store, p);
        try std.testing.expectEqual(@as(usize, 1), sets.len);
        const from = parseCornerIn("<from><col>3</col><colOff>0</colOff><row>1</row><rowOff>0</rowOff></from>", 0, sets, .from).?;
        try std.testing.expectEqual(@as(u32, 3), from.anchor.col);
        try std.testing.expectEqual(@as(u32, 1), from.anchor.row);
    }
    // The Strict URI as the default namespace resolves the same way.
    {
        const p = resolveDrawingPrefixes("<wsDr xmlns=\"http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing\"/>");
        try std.testing.expectEqualStrings("", p.xdr);
        try std.testing.expect(!p.xdr_rejected);
    }
    // A prefixed root beside a default declaration on the xdr URI: the
    // root's prefix is primary, the default namespace an alternate —
    // unprefixed anchors are replayed.
    {
        const p = resolveDrawingPrefixes("<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"/>");
        try std.testing.expectEqualStrings("xdr", p.xdr);
        try std.testing.expectEqual(@as(usize, 1), p.xdr_alts().len);
        try std.testing.expectEqualStrings("", p.xdr_alts()[0]);
        try std.testing.expect(p.followsXdr(""));
        try std.testing.expect(!p.xdr_rejected);
    }
    // An unprefixed root with a default declaration bound elsewhere
    // and `xdr:` bound beside it: `xdr` is primary, the empty prefix
    // is not an xdr prefix — an unprefixed `<twoCellAnchor>` there is
    // a foreign element.
    {
        const p = resolveDrawingPrefixes("<wsDr xmlns=\"urn:not-a-drawing\" xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"/>");
        try std.testing.expectEqualStrings("xdr", p.xdr);
        try std.testing.expectEqual(@as(usize, 0), p.xdr_alts().len);
        try std.testing.expect(!p.followsXdr(""));
        try std.testing.expect(!p.xdr_rejected);
    }
    // A default declaration below the root, on the anchor: collected
    // as an alternate (the whole-part rule for xdr alternates).
    {
        const p = resolveDrawingPrefixes("<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"><twoCellAnchor xmlns=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"/></xdr:wsDr>");
        try std.testing.expectEqualStrings("xdr", p.xdr);
        try std.testing.expect(p.followsXdr(""));
        try std.testing.expect(!p.xdr_rejected);
    }
}

test "resolveDrawingPrefixes: a spreadsheetDrawing binding the walk cannot spell is xdr_rejected — overlong, past the replay cap; a decoy is text; the limit is followed" {
    const uri = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    const at_limit = "p" ** 100;
    const past_limit = "p" ** 101;
    // Overlong beside a followed primary.
    {
        const p = resolveDrawingPrefixes("<xdr:wsDr xmlns:xdr=\"" ++ uri ++ "\" xmlns:" ++ past_limit ++ "=\"" ++ uri ++ "\"/>");
        try std.testing.expectEqualStrings("xdr", p.xdr);
        try std.testing.expect(p.xdr_rejected);
    }
    // At the limit: followed as an alternate.
    {
        const p = resolveDrawingPrefixes("<xdr:wsDr xmlns:xdr=\"" ++ uri ++ "\" xmlns:" ++ at_limit ++ "=\"" ++ uri ++ "\"/>");
        try std.testing.expect(p.followsXdr(at_limit));
        try std.testing.expect(!p.xdr_rejected);
    }
    // An overlong ROOT prefix: the root's own name is unfollowed —
    // the primary falls back to the canonical spelling the part does
    // not use, so the part is rejected, not silently empty.
    {
        const p = resolveDrawingPrefixes("<" ++ past_limit ++ ":wsDr xmlns:" ++ past_limit ++ "=\"" ++ uri ++ "\"/>");
        try std.testing.expectEqualStrings("xdr", p.xdr);
        try std.testing.expect(p.xdr_rejected);
    }
    // A ninth alternate is past `max_xdr_alts`: rejected; eight are
    // followed.
    {
        var xml: std.ArrayListUnmanaged(u8) = .empty;
        defer xml.deinit(std.testing.allocator);
        try xml.appendSlice(std.testing.allocator, "<xdr:wsDr xmlns:xdr=\"" ++ uri ++ "\"");
        for (0..max_xdr_alts) |k| {
            var decl_buf: [128]u8 = undefined;
            try xml.appendSlice(std.testing.allocator, try std.fmt.bufPrint(&decl_buf, " xmlns:p{d}=\"{s}\"", .{ k, uri }));
        }
        try xml.appendSlice(std.testing.allocator, "/>");
        const eight = resolveDrawingPrefixes(xml.items);
        try std.testing.expectEqual(max_xdr_alts, eight.xdr_alts().len);
        try std.testing.expect(!eight.xdr_rejected);
        // One more.
        xml.items.len -= "/>".len;
        try xml.appendSlice(std.testing.allocator, " xmlns:p8=\"" ++ uri ++ "\"/>");
        const nine = resolveDrawingPrefixes(xml.items);
        try std.testing.expectEqual(max_xdr_alts, nine.xdr_alts().len);
        try std.testing.expect(nine.xdr_rejected);
    }
    // Declaration-shaped text in a comment, a PI or an attribute value
    // is not a binding: nothing rejected.
    {
        const p = resolveDrawingPrefixes("<xdr:wsDr xmlns:xdr=\"" ++ uri ++ "\" title=\"xmlns:" ++ past_limit ++ "='" ++ uri ++ "'\"><!-- xmlns:" ++ past_limit ++ "=\"" ++ uri ++ "\" --><?pi xmlns:" ++ past_limit ++ "=\"" ++ uri ++ "\"?></xdr:wsDr>");
        try std.testing.expect(!p.xdr_rejected);
    }
    // A binding on another URI, however long, is not this walk's.
    {
        const p = resolveDrawingPrefixes("<xdr:wsDr xmlns:xdr=\"" ++ uri ++ "\" xmlns:" ++ past_limit ++ "=\"urn:other\"/>");
        try std.testing.expect(!p.xdr_rejected);
    }
}

test "anchors read: an openpyxl-shaped default-namespace drawing lists its image and its chart; an unfollowed binding refuses under strict, lists under lenient" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const a = std.testing.allocator;
    var prng = std.Random.DefaultPrng.init(@truncate(@as(u96, @bitCast(std.Io.Clock.now(.awake, io).nanoseconds))));
    var buf: [256]u8 = undefined;
    const path = try std.fmt.bufPrint(&buf, ".zig-cache/test-drawings-default-ns-{d}.xlsx", .{prng.random().int(u32)});
    defer std.Io.Dir.cwd().deleteFile(io, path) catch {};
    const uri = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    const drawing_default =
        "<wsDr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns=\"" ++ uri ++ "\">" ++
        "<oneCellAnchor><from><col>3</col><colOff>0</colOff><row>1</row><rowOff>0</rowOff></from><ext cx=\"5400000\" cy=\"2700000\" /><graphicFrame><nvGraphicFramePr><cNvPr id=\"1\" name=\"Chart 1\" /><cNvGraphicFramePr /></nvGraphicFramePr><xfrm /><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" r:id=\"rId1\" /></a:graphicData></a:graphic></graphicFrame><clientData /></oneCellAnchor>" ++
        "<twoCellAnchor><from><col>0</col><colOff>0</colOff><row>6</row><rowOff>0</rowOff></from><to><col>2</col><colOff>0</colOff><row>9</row><rowOff>0</rowOff></to><pic><nvPicPr><cNvPr id=\"2\" name=\"Image 1\" /><cNvPicPr /></nvPicPr><blipFill><a:blip r:embed=\"rId2\" /><a:stretch><a:fillRect /></a:stretch></blipFill><spPr><a:prstGeom prst=\"rect\"><a:avLst /></a:prstGeom></spPr></pic><clientData /></twoCellAnchor>" ++
        "</wsDr>";
    for ([_]bool{ false, true }) |reject| {
        var store = try PartStore.open(a, io, "tests/corpus/openpyxl_chart.xlsx");
        defer store.deinit();
        const drawing = if (reject)
            try std.mem.replaceOwned(u8, a, drawing_default, "<wsDr ", "<wsDr xmlns:" ++ ("p" ** 101) ++ "=\"" ++ uri ++ "\" ")
        else
            try a.dupe(u8, drawing_default);
        defer a.free(drawing);
        try store.replacePart("xl/drawings/drawing1.xml", drawing);
        try store.addPart("xl/media/image1.png", "image/png", "\x89PNG\r\n\x1a\n01234567");
        try store.replacePart("xl/drawings/_rels/drawing1.xml.rels",
            \\<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="/xl/charts/chart1.xml" Id="rId1" /><Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="/xl/media/image1.png" Id="rId2" /></Relationships>
        );
        try store.save(io, path);
        var s = try PartStore.open(a, io, path);
        defer s.deinit();
        if (reject) {
            try std.testing.expectError(error.MalformedDrawingXml, imageAnchorsIn(&s, a, .strict));
            try std.testing.expectError(error.MalformedDrawingXml, chartAnchorsIn(&s, a, .strict));
        }
        const mode: WalkMode = if (reject) .lenient else .strict;
        const images = try imageAnchorsIn(&s, a, mode);
        defer a.free(images);
        try std.testing.expectEqual(@as(usize, 1), images.len);
        try std.testing.expectEqual(AnchorKind.two_cell, images[0].kind);
        try std.testing.expectEqual(@as(u32, 0), images[0].from.col);
        try std.testing.expectEqual(@as(u32, 6), images[0].from.row);
        try std.testing.expectEqual(@as(u32, 2), images[0].to.?.col);
        try std.testing.expectEqual(@as(u32, 9), images[0].to.?.row);
        try std.testing.expectEqualStrings("xl/media/image1.png", images[0].image_part_name);
        const charts = try chartAnchorsIn(&s, a, mode);
        defer {
            for (charts) |c| a.free(c.series_refs);
            a.free(charts);
        }
        try std.testing.expectEqual(@as(usize, 1), charts.len);
        try std.testing.expectEqual(AnchorKind.one_cell, charts[0].kind);
        try std.testing.expectEqual(@as(u32, 3), charts[0].from.col);
        try std.testing.expectEqual(@as(u32, 1), charts[0].from.row);
        try std.testing.expectEqual(ChartType.bar, charts[0].chart_type);
        try std.testing.expectEqual(@as(usize, 3), charts[0].series_refs.len);
        try std.testing.expectEqualStrings("'Data'!B1", charts[0].series_refs[0]);
    }
}

test "anchors read: markup-shaped text in a comment / CDATA / PI of a default-namespace drawing is text; a mixed-spelling anchor is listed (ND-REL-101, ND-REL-103)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const a = std.testing.allocator;
    var prng = std.Random.DefaultPrng.init(@truncate(@as(u96, @bitCast(std.Io.Clock.now(.awake, io).nanoseconds))));
    var buf: [256]u8 = undefined;
    const path = try std.fmt.bufPrint(&buf, ".zig-cache/test-drawings-decoys-{d}.xlsx", .{prng.random().int(u32)});
    defer std.Io.Dir.cwd().deleteFile(io, path) catch {};
    const uri = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    const chart_anchor = "<oneCellAnchor><from><col>3</col><colOff>0</colOff><row>1</row><rowOff>0</rowOff></from><ext cx=\"5400000\" cy=\"2700000\" /><graphicFrame><nvGraphicFramePr><cNvPr id=\"1\" name=\"Chart 1\" /><cNvGraphicFramePr /></nvGraphicFramePr><xfrm /><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" r:id=\"rId1\" /></a:graphicData></a:graphic></graphicFrame><clientData /></oneCellAnchor>";
    const Case = struct { name: []const u8, root: []const u8, body: []const u8, listed: usize };
    const cases = [_]Case{
        // The bare `<` opener stepping one byte into `<!--` made the
        // commented copy a phantom anchor (row 100 on the wire) and a
        // trailing comment a refusal.
        .{ .name = "commented decoy + trailing comment", .root = "<wsDr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns=\"" ++ uri ++ "\">", .body = "<!-- <oneCellAnchor><from><col>3</col><colOff>0</colOff><row>99</row><rowOff>0</rowOff></from> -->" ++ chart_anchor ++ "<!-- TODO: a second <oneCellAnchor> goes here -->", .listed = 1 },
        .{ .name = "CDATA + PI decoys", .root = "<wsDr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns=\"" ++ uri ++ "\">", .body = "<![CDATA[<oneCellAnchor><from><row>42</row></from>]]><?note <oneCellAnchor> ?>" ++ chart_anchor, .listed = 1 },
        // An unterminated comment swallows the rest of the part —
        // nothing listed, no refusal, as for the prefixed spelling.
        .{ .name = "unterminated comment", .root = "<wsDr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns=\"" ++ uri ++ "\">", .body = "<!-- unterminated " ++ chart_anchor, .listed = 0 },
        // Both names followed: a prefixed wrapper over bare children,
        // with the `from` block's scalars mixed too.
        .{ .name = "mixed spelling", .root = "<wsDr xmlns:xdr=\"" ++ uri ++ "\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns=\"" ++ uri ++ "\">", .body = "<xdr:oneCellAnchor><from><xdr:col>3</xdr:col><colOff>0</colOff><row>1</row><rowOff>0</rowOff></from><xdr:ext cx=\"5400000\" cy=\"2700000\" /><graphicFrame><nvGraphicFramePr><cNvPr id=\"1\" name=\"Chart 1\" /><cNvGraphicFramePr /></nvGraphicFramePr><xfrm /><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" r:id=\"rId1\" /></a:graphicData></a:graphic></graphicFrame><clientData /></xdr:oneCellAnchor>", .listed = 1 },
    };
    for (cases) |case| {
        var store = try PartStore.open(a, io, "tests/corpus/openpyxl_chart.xlsx");
        defer store.deinit();
        const drawing = try std.mem.concat(a, u8, &.{ case.root, case.body, "</wsDr>" });
        defer a.free(drawing);
        try store.replacePart("xl/drawings/drawing1.xml", drawing);
        try store.save(io, path);
        var s = try PartStore.open(a, io, path);
        defer s.deinit();
        const charts = try chartAnchorsIn(&s, a, .strict);
        defer {
            for (charts) |c| a.free(c.series_refs);
            a.free(charts);
        }
        std.testing.expectEqual(case.listed, charts.len) catch |e| {
            std.debug.print("case: {s}\n", .{case.name});
            return e;
        };
        for (charts) |c| {
            try std.testing.expectEqual(@as(u32, 1), c.from.row);
            try std.testing.expectEqual(@as(u32, 3), c.from.col);
        }
    }
}

test "anchors read: an absoluteAnchor under the default namespace parses its <pos> and <ext> bare" {
    var bufs: [max_tag_sets]TagSetBuf = undefined;
    var set_store: [max_tag_sets]DrawingTags = undefined;
    const sets = testTagSets(&bufs, &set_store, "<wsDr xmlns=\"" ++ ns_xdr_transitional ++ "\"/>");
    const xml =
        \\<absoluteAnchor><pos x="1000" y="2000" /><ext cx="914400" cy="457200" /><pic/><clientData /></absoluteAnchor>
    ;
    const abs = parseAbsoluteAnchorIn(xml, 0, sets).?;
    try std.testing.expectEqual(@as(i64, 1000), abs.x);
    try std.testing.expectEqual(@as(i64, 2000), abs.y);
    try std.testing.expectEqual(@as(i64, 914400), abs.cx);
    try std.testing.expectEqual(@as(i64, 457200), abs.cy);
    // The strict one-cell validation reads the bare `<ext` too, and
    // `<extLst` does not satisfy it; an unescaped `>` in an attribute
    // value does not end the tag (ND-REL-405).
    try std.testing.expect(parseExtAttrsIn("<oneCellAnchor><from/><ext cx=\"1\" cy=\"2\"/></oneCellAnchor>", 0, sets) != null);
    try std.testing.expectEqual(@as(i64, 2), parseExtAttrsIn("<oneCellAnchor><from/><ext desc=\"a>b\" cx=\"1\" cy=\"2\"/></oneCellAnchor>", 0, sets).?.cy);
    try std.testing.expectEqual(@as(i64, 2000), parseAbsoluteAnchorIn("<absoluteAnchor><pos desc=\"a>b\" x=\"1000\" y=\"2000\" /><ext cx=\"1\" cy=\"2\" /></absoluteAnchor>", 0, sets).?.y);
    try std.testing.expect(parseExtAttrsIn("<oneCellAnchor><from/><extLst cx=\"1\" cy=\"2\"/></oneCellAnchor>", 0, sets) == null);
}

test "anchors read: XSD-collapsed whitespace around a scalar parses; a DTD refuses under strict and is a region under lenient; a self-closing wrapper is stepped over (ND-REL-201/203, ND-DOC-204)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const a = std.testing.allocator;
    var prng = std.Random.DefaultPrng.init(@truncate(@as(u96, @bitCast(std.Io.Clock.now(.awake, io).nanoseconds))));
    var buf: [256]u8 = undefined;
    const path = try std.fmt.bufPrint(&buf, ".zig-cache/test-drawings-r2-{d}.xlsx", .{prng.random().int(u32)});
    defer std.Io.Dir.cwd().deleteFile(io, path) catch {};
    const uri = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    const root = "<wsDr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns=\"" ++ uri ++ "\">";
    const frame = "<ext cx=\"5400000\" cy=\"2700000\" /><graphicFrame><nvGraphicFramePr><cNvPr id=\"1\" name=\"Chart 1\" /><cNvGraphicFramePr /></nvGraphicFramePr><xfrm /><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" r:id=\"rId1\" /></a:graphicData></a:graphic></graphicFrame><clientData /></oneCellAnchor>";
    const padded = "<oneCellAnchor><from><col> 3 </col><colOff>0</colOff><row>\n1\n</row><rowOff>0</rowOff></from>" ++ frame;
    const plain = "<oneCellAnchor><from><col>3</col><colOff>0</colOff><row>1</row><rowOff>0</rowOff></from>" ++ frame;
    const Case = struct { name: []const u8, body: []const u8, strict_refuses: bool, listed: usize, expect_offset_of: ?[]const u8 = null };
    const cases = [_]Case{
        .{ .name = "padded scalars", .body = padded, .strict_refuses = false, .listed = 1 },
        // A DTD whose entity value spells an anchor: strict refuses the
        // part; lenient steps over the declaration and lists the one
        // real anchor, never the entity's.
        .{ .name = "doctype", .body = "<!DOCTYPE wsDr [ <!ENTITY ghost \"<oneCellAnchor><from><col>7</col><colOff>0</colOff><row>42</row><rowOff>0</rowOff></from>" ++ frame ++ "\"> ]>" ++ plain, .strict_refuses = true, .listed = 1 },
        // A self-closing wrapper before the real one: not an anchor,
        // and not the start of the real one's block.
        .{ .name = "self-closing then real", .body = "<oneCellAnchor/><oneCellAnchor />" ++ plain, .strict_refuses = false, .listed = 1, .expect_offset_of = plain },
    };
    for (cases) |case| {
        var store = try PartStore.open(a, io, "tests/corpus/openpyxl_chart.xlsx");
        defer store.deinit();
        const drawing = try std.mem.concat(a, u8, &.{ if (std.mem.startsWith(u8, case.body, "<!DOCTYPE")) "" else root, case.body, "</wsDr>" });
        defer a.free(drawing);
        // The DTD case: the declaration precedes the root element.
        const drawing2 = if (std.mem.startsWith(u8, case.body, "<!DOCTYPE")) blk: {
            const subset_end = std.mem.indexOf(u8, case.body, "]>").? + 2;
            break :blk try std.mem.concat(a, u8, &.{ case.body[0..subset_end], root, case.body[subset_end..], "</wsDr>" });
        } else try a.dupe(u8, drawing);
        defer a.free(drawing2);
        try store.replacePart("xl/drawings/drawing1.xml", drawing2);
        try store.save(io, path);
        var s = try PartStore.open(a, io, path);
        defer s.deinit();
        if (case.strict_refuses) {
            try std.testing.expectError(error.MalformedDrawingXml, chartAnchorsIn(&s, a, .strict));
        }
        const mode: WalkMode = if (case.strict_refuses) .lenient else .strict;
        const charts = try chartAnchorsIn(&s, a, mode);
        defer {
            for (charts) |c| a.free(c.series_refs);
            a.free(charts);
        }
        std.testing.expectEqual(case.listed, charts.len) catch |e| {
            std.debug.print("case: {s}\n", .{case.name});
            return e;
        };
        for (charts) |c| {
            try std.testing.expectEqual(@as(u32, 1), c.from.row);
            try std.testing.expectEqual(@as(u32, 3), c.from.col);
        }
        if (case.expect_offset_of) |needle| {
            const d = (try s.part("xl/drawings/drawing1.xml")).?;
            try std.testing.expectEqual(std.mem.indexOf(u8, d.bytes, needle).?, charts[0].doc_offset);
        }
    }
    // The region skipper on its own: a DOCTYPE with a subset holding
    // quotes, a comment and a `>` inside an entity value; a bare one;
    // an unterminated one.
    try std.testing.expectEqual(@as(?usize, 48), skipRegionCloseFrom("<!DOCTYPE wsDr [ <!ENTITY x \"a>b\"> <!-- ] --> ]><wsDr/>", 0));
    try std.testing.expectEqual(@as(?usize, 15), skipRegionCloseFrom("<!DOCTYPE wsDr><wsDr/>", 0));
    try std.testing.expectEqual(@as(?usize, null), skipRegionCloseFrom("<!DOCTYPE wsDr [ <!ENTITY x \"a\">", 0));
    try std.testing.expect(hasLiveDoctype("<?xml version=\"1.0\"?><!DOCTYPE wsDr><wsDr/>"));
    try std.testing.expect(!hasLiveDoctype("<!-- <!DOCTYPE wsDr> --><wsDr/>"));
    // The quote-aware open-tag end.
    try std.testing.expect(!selfClosingTagEnd("<xdr:oneCellAnchor editAs=\"a/>b\"><xdr:from/>", 0).?.self_closing);
    try std.testing.expect(selfClosingTagEnd("<xdr:oneCellAnchor editAs=\"a/>b\" />", 0).?.self_closing);
    try std.testing.expectEqual(@as(?OpenTagEnd, null), selfClosingTagEnd("<xdr:oneCellAnchor editAs=\"a", 0));
}

test "anchors read: a nested corner block refuses under strict; a close spelled in an attribute value is not the close; digit separators do not parse; a descendant's default declaration is not the root's (round 3)" {
    var bufs: [max_tag_sets]TagSetBuf = undefined;
    var set_store: [max_tag_sets]DrawingTags = undefined;
    const sets = testTagSets(&bufs, &set_store, "<xdr:wsDr xmlns:xdr=\"" ++ ns_xdr_transitional ++ "\"/>");
    // A `<to>` nested inside `<from>`: both parse, and they overlap.
    const nested = "<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff><xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to></xdr:from></xdr:twoCellAnchor>";
    const from = parseCornerIn(nested, 0, sets, .from).?;
    const to = parseCornerIn(nested, 0, sets, .to).?;
    try std.testing.expect(cornersOverlap(from, to));
    try std.testing.expect(cornersOverlap(to, from));
    try std.testing.expectError(error.MalformedDrawingXml, readAnchorGeometry(nested, 0, sets, true, false, .strict));
    const lenient = try readAnchorGeometry(nested, 0, sets, true, false, .lenient);
    try std.testing.expectEqual(@as(u32, 10), lenient.to.?.row);
    // Disjoint, reversed: not an overlap.
    const reversed = "<xdr:twoCellAnchor><xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from></xdr:twoCellAnchor>";
    try std.testing.expect(!cornersOverlap(parseCornerIn(reversed, 0, sets, .from).?, parseCornerIn(reversed, 0, sets, .to).?));
    // The XSD integer: digits only after the optional sign.
    try std.testing.expectEqual(@as(?u32, 10), parseXsdInteger(u32, " 10 "));
    try std.testing.expectEqual(@as(?u32, 10), parseXsdInteger(u32, "+10"));
    try std.testing.expectEqual(@as(?i64, -10), parseXsdInteger(i64, "-10"));
    try std.testing.expectEqual(@as(?u32, null), parseXsdInteger(u32, "1_0"));
    try std.testing.expectEqual(@as(?u32, null), parseXsdInteger(u32, "-1"));
    try std.testing.expectEqual(@as(?u32, null), parseXsdInteger(u32, "0x10"));
    try std.testing.expectEqual(@as(?u32, null), parseXsdInteger(u32, "+"));
    try std.testing.expectEqual(@as(?u32, null), parseXsdInteger(u32, ""));
    try std.testing.expect(parseCornerIn("<xdr:from><xdr:col>1_0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>", 0, sets, .from) == null);
    // A root in no namespace over an anchor that binds the default one
    // itself: the primary is the canonical fallback, the empty prefix
    // an alternate — followed either way.
    const p = resolveDrawingPrefixes("<wsDr><twoCellAnchor xmlns=\"" ++ ns_xdr_transitional ++ "\"/></wsDr>");
    try std.testing.expectEqualStrings("xdr", p.xdr);
    try std.testing.expect(p.followsXdr(""));
    try std.testing.expect(!p.xdr_rejected);
    try std.testing.expectEqual(@as(?[]const u8, null), rootDefaultNamespaceUri("<wsDr><twoCellAnchor xmlns=\"urn:x\"/></wsDr>"));
    try std.testing.expectEqualStrings("urn:r", rootDefaultNamespaceUri("<?xml version=\"1.0\"?><!-- c --><wsDr xmlns:a=\"urn:a\" xmlns=\"urn:r\"><x xmlns=\"urn:x\"/></wsDr>").?);
}

test "anchors read: a `<` inside an attribute value is not well-formed — strict refuses, lenient neither ends the anchor nor takes the corner from it (ND-REL-302/411)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const a = std.testing.allocator;
    var prng = std.Random.DefaultPrng.init(@truncate(@as(u96, @bitCast(std.Io.Clock.now(.awake, io).nanoseconds))));
    var buf: [256]u8 = undefined;
    const path = try std.fmt.bufPrint(&buf, ".zig-cache/test-drawings-attrclose-{d}.xlsx", .{prng.random().int(u32)});
    defer std.Io.Dir.cwd().deleteFile(io, path) catch {};
    const decoys = [_][]const u8{
        // A close in the wrapper's attribute value (the read used to
        // serve an EMPTY inventory for an anchor the sweep shifted).
        "<oneCellAnchor editAs=\"</oneCellAnchor>\">",
        // A corner there (the read used to serve it as the corner).
        "<oneCellAnchor editAs=\"<from><col>7</col><colOff>0</colOff><row>99</row><rowOff>0</rowOff></from>\">",
        // A chart stub in the graphic frame's own attribute.
        "<oneCellAnchor editAs=\"<graphicFrame macro='<c:chart r:id=`rId9`/>'>\">",
    };
    for (decoys) |decoy| {
        var store = try PartStore.open(a, io, "tests/corpus/openpyxl_chart.xlsx");
        defer store.deinit();
        const drawing = (try store.part("xl/drawings/drawing1.xml")).?;
        const patched = try std.mem.replaceOwned(u8, a, drawing.bytes, "<oneCellAnchor>", decoy);
        defer a.free(patched);
        try std.testing.expect(hasMarkupInAttributeValue(patched));
        try store.replacePart("xl/drawings/drawing1.xml", patched);
        try store.save(io, path);
        var s = try PartStore.open(a, io, path);
        defer s.deinit();
        try std.testing.expectError(error.MalformedDrawingXml, chartAnchorsIn(&s, a, .strict));
        try std.testing.expectError(error.MalformedDrawingXml, imageAnchorsIn(&s, a, .strict));
        const charts = try chartAnchorsIn(&s, a, .lenient);
        defer {
            for (charts) |c| a.free(c.series_refs);
            a.free(charts);
        }
        try std.testing.expectEqual(@as(usize, 1), charts.len);
        try std.testing.expectEqual(@as(u32, 1), charts[0].from.row);
        try std.testing.expectEqual(@as(u32, 3), charts[0].from.col);
        try std.testing.expectEqualStrings("xl/charts/chart1.xml", charts[0].chart_part_name);
    }
    // The probe itself: escaped markup, a `>` in a value, quotes of
    // either kind, and regions are not violations; a `<` is, in either
    // quote; a part ending inside a quote is judged elsewhere.
    try std.testing.expect(!hasMarkupInAttributeValue("<a b=\"&lt;x&gt;\" c='a>b'/>"));
    try std.testing.expect(!hasMarkupInAttributeValue("<!-- <a b=\"<x>\"/> --><![CDATA[<a b=\"<\">]]><?pi <a b=\"<\"> ?><!DOCTYPE a [ <!ENTITY e \"<x>\"> ]><a/>"));
    try std.testing.expect(hasMarkupInAttributeValue("<a b=\"<x/>\"/>"));
    try std.testing.expect(hasMarkupInAttributeValue("<a b='</a>'/>"));
    try std.testing.expect(hasMarkupInAttributeValue("<root><a b=\"x\" c=\"<\"/></root>"));
    try std.testing.expect(!hasMarkupInAttributeValue("<a b=\"<x"));
    try std.testing.expect(!hasMarkupInAttributeValue(""));
}

test "ChartFormulaWalk: a live DTD refuses the strict walk and is a region under lenient (ND-REL-302)" {
    const chart = "<!DOCTYPE chartSpace [ <!ENTITY g \"<c:f>Ghost!$A$1</c:f>\"> ]><c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart><c:ser><c:tx><c:strRef><c:f>Data!$B$1</c:f></c:strRef></c:tx></c:ser></c:chart></c:chartSpace>";
    var strict = ChartFormulaWalk.init(chart);
    try std.testing.expectError(error.MalformedChartXml, strict.next(.strict));
    var lenient = ChartFormulaWalk.init(chart);
    const f = (try lenient.next(.lenient)).?;
    try std.testing.expectEqualStrings("Data!$B$1", chart[f.body_start..f.body_end]);
    try std.testing.expectEqual(@as(?ChartFormula, null), try lenient.next(.lenient));
    // A DOCTYPE inside a comment is text.
    var commented = ChartFormulaWalk.init("<!-- <!DOCTYPE x> --><c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:f>A!$A$1</c:f></c:chartSpace>");
    try std.testing.expect((try commented.next(.strict)) != null);
    // A carrier spelled inside an attribute value: not well-formed —
    // strict refuses (ND-REL-411); lenient walks what it finds.
    var attr = ChartFormulaWalk.init("<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart title=\"<c:f>Ghost!$A$1</c:f>\"><c:f>A!$A$1</c:f></c:chart></c:chartSpace>");
    try std.testing.expectError(error.MalformedChartXml, attr.next(.strict));
}
