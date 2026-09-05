//! DrawingML anchor rewriter for OOXML `xl/drawings/drawingN.xml`
//! parts. Pure-function; consumed by `pkg/editor.zig`'s row/col
//! edit path after the worksheet's own bytes have been rewritten
//! by `pkg/sheet_edit.zig`. Companion to (not extension of)
//! sheet_edit because drawing anchors live in a separate part,
//! not the worksheet body.
//!
//! Walks `<xdr:wsDr>` for `<xdr:twoCellAnchor>` and
//! `<xdr:oneCellAnchor>` blocks. Inside each anchor's `<xdr:from>`
//! and `<xdr:to>` sub-blocks, shifts the `<xdr:col>` value on
//! column edits and the `<xdr:row>` value on row edits — both are
//! 0-based unsigned ints (ECMA-376 §20.5.2.13 / §20.5.2.21).
//!
//! `<xdr:absoluteAnchor>` carries pixel coordinates in `<xdr:pos>`
//! and passes through unchanged — there is no row/col coordinate
//! to shift.
//!
//! Drop semantics (anchor + body removed entirely from the
//! drawing):
//!
//! - `twoCellAnchor`: when BOTH `<xdr:from>` and `<xdr:to>`
//!   coordinates on the edited axis equal the deleted row/column.
//!   The image's full extent is on the deleted axis line.
//! - `oneCellAnchor`: when the `<xdr:from>` coordinate on the
//!   edited axis equals the deleted row/column. The single anchor
//!   cell is gone.
//!
//! The walk is the anchors read's (`pkg/drawings.zig`), not a copy of
//! it — the namespace-aware drawing slice: `xdr:` above is the
//! canonical spelling, not a literal. `drawings.resolveDrawingPrefixes`
//! resolves the part ONCE for both — the root element's prefix, every
//! alternate bound to a spreadsheetDrawing URI anywhere in the part,
//! and the DEFAULT namespace as the empty prefix (`<wsDr xmlns="…/
//! spreadsheetDrawing"><oneCellAnchor><from><row>`, openpyxl 3.1's
//! spelling, whose anchors the `xdr:`-literal v1 of this sweep left
//! in place while the grid and the chart's series formulas moved:
//! in-house chart-sweep round 5 CF-DOC-501). `drawings.buildTagSets`
//! spells one tag set per followed prefix; an anchor WRAPPER is
//! matched under any set (one document-order pass, since a rewrite
//! must keep the order) and its children under any set too — a part
//! that binds the namespace twice may spell a wrapper under one name
//! and its `from` / `col` under the other (in-house ND-REL-103). The
//! lexical layer is the read's: comments / CDATA / PIs are copied
//! through whole and never matched (`drawings.skipRegionEndFrom`,
//! `findLiveMarkup` — the v1 walker took a commented `<xdr:twoCellAnchor>`
//! for an anchor and spliced the real ones after it, in-house
//! ND-REL-101), tag names are exact QNames, and a corner is read by
//! the read's own parser (`drawings.parseCornerIn` — the values AND the
//! scalar spans this sweep splices; one acceptance, `parseXsdInteger`:
//! XSD-collapsed, an optional sign, digits only).
//!
//! What the sweep cannot move it refuses (`MalformedDrawingXml`),
//! rather than leave an anchor behind: a spreadsheetDrawing binding
//! the resolver cannot spell (`DrawingPrefixes.xdr_rejected` — a
//! prefix past its 100-byte limit, or past its eight-alternate cap),
//! a `<!DOCTYPE` (no producer writes one; an entity could rewrite
//! the markup), a `<` inside an attribute value (not well-formed XML —
//! a close or a corner spelled there would be taken for the real one),
//! an anchor wrapper with no close or an unterminated
//! start tag, a corner block absent or with a scalar that does not
//! parse, two corner blocks that overlap. On the anchors both walk
//! these are the strict read's refusals on the same bytes, by the
//! same parser; `Workbook.applySheetEditTransform` runs the sweep
//! before the edit's first mutation, so the verdict lands with
//! nothing installed. The two walks differ where their jobs do, and
//! only there: the sweep must move EVERY anchor, so it reads the
//! corners of a shape (`<xdr:sp>`) or a frame without a chart, which
//! the read never lists and never validates; the strict read requires
//! a one-cell anchor's schema-mandated `<ext>`, which the sweep does
//! not move and does not check; a self-closing wrapper
//! (`<xdr:twoCellAnchor/>`) holds no anchor — the read steps over it,
//! the sweep passes it through; a reversed pair — `<to>` written
//! before `<from>` — is schema-invalid but readable: the read lists
//! it and the sweep moves both blocks, emitted in document order (the
//! v1 walker sliced backwards on it — in-house ND-REL-102). The
//! sheet's `<drawing>` element is followed by the read's own edge
//! resolution (`drawings.findDrawingRef` — any bound prefix, live
//! elements only — and the typed relationship lookup), and a
//! reference the strict read cannot follow refuses the edit as it
//! refuses the inventory (the v1 sweep's raw `<drawing r:id` scan
//! found nothing behind a prefixed or decoyed element and moved the
//! grid alone — in-house ND-REL-201). An unprefixed anchor under a
//! root whose default namespace is NOT a spreadsheetDrawing one is
//! not an anchor and passes through, as the read lists nothing for
//! it.

const std = @import("std");
const drawings = @import("drawings.zig");

const Allocator = std.mem.Allocator;

pub const Error = error{
    MalformedDrawingXml,
    DrawingCoordinateOverflow,
} || Allocator.Error;

pub const EditKind = enum { insert, delete };

pub const Axis = enum { row, col };

/// Apply one row OR column edit to a drawing-part body. `index`
/// is 0-based to match the wire format of `<xdr:col>` /
/// `<xdr:row>`. The Editor's row/col-edit surfaces use 1-based
/// indices and must subtract 1 before calling. Returns a freshly
/// allocated buffer. Refuses `MalformedDrawingXml` for a part the
/// walk cannot rewrite whole (the header's list) — before writing
/// anything, so a dry run is the refusal the apply would raise.
pub fn applyEditToDrawing(
    allocator: Allocator,
    src: []const u8,
    axis: Axis,
    index: u32,
    kind: EditKind,
) Error![]u8 {
    const prefixes = drawings.resolveDrawingPrefixes(src);
    if (prefixes.xdr_rejected) return Error.MalformedDrawingXml;
    // A DTD is not content (the strict read's rule): refuse rather
    // than splice a grid coordinate inside an entity declaration.
    if (drawings.hasLiveDoctype(src)) return Error.MalformedDrawingXml;
    // A `<` inside an attribute value is not well-formed XML: no walk
    // can tell its markup from its text (the strict read's rule).
    if (drawings.hasMarkupInAttributeValue(src)) return Error.MalformedDrawingXml;
    const set_count = drawings.tagSetCount(&prefixes);
    const bufs = try allocator.alloc(drawings.TagSetBuf, set_count);
    defer allocator.free(bufs);
    const set_store = try allocator.alloc(drawings.DrawingTags, set_count);
    defer allocator.free(set_store);
    // Every needle fits under the resolver's prefix limit; a set that
    // cannot be spelled is a part the walk cannot read.
    const sets = drawings.buildTagSets(bufs, set_store, prefixes) catch return Error.MalformedDrawingXml;

    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(allocator);

    var i: usize = 0;
    while (i < src.len) {
        const lt = std.mem.indexOfScalarPos(u8, src, i, '<') orelse {
            try out.appendSlice(allocator, src[i..]);
            return try out.toOwnedSlice(allocator);
        };
        try out.appendSlice(allocator, src[i..lt]);
        // A comment / CDATA / PI is copied whole: markup-shaped text
        // inside it is text (the read's rule); an unterminated region
        // swallows the rest of the part, as it does for the read.
        // `skipRegionEndFrom` answers `lt + 1` for a `<` that opens no
        // region — no region is one byte long.
        const region_end = drawings.skipRegionEndFrom(src, lt);
        if (region_end != lt + 1) {
            try out.appendSlice(allocator, src[lt..region_end]);
            i = region_end;
            continue;
        }
        if (try matchAnchorOpenAt(src, lt, sets)) |a| {
            try processAnchor(allocator, &out, src, a, sets, axis, index, kind);
            i = a.after_close;
        } else {
            try out.append(allocator, '<');
            i = lt + 1;
        }
    }
    return try out.toOwnedSlice(allocator);
}

const AnchorKind = enum { two_cell, one_cell };

const AnchorMatch = struct {
    kind: AnchorKind,
    /// Absolute byte offset of the opening `<xdr:NAME` token.
    open_start: usize,
    /// Absolute byte offset just past the opening tag's `>`.
    after_open: usize,
    /// Absolute byte offset just past the closing tag (`after_open`
    /// for a self-closing wrapper).
    after_close: usize,
    /// True when the opening tag was self-closing (`<xdr:foo/>`),
    /// which is malformed for an anchor element (per ECMA-376 the
    /// anchor MUST contain a `<xdr:from>` etc) — but we tolerate
    /// it by emitting and skipping the empty anchor verbatim.
    self_closing: bool,
};

/// The anchor wrapper opening at `at` under any of `sets` — its kind
/// and extent — or null when the tag at `at` is not one. A wrapper the
/// walk cannot bound is a refusal: an opening tag with no `>`, or an
/// open with no live close (the strict read's verdict on the same
/// bytes).
fn matchAnchorOpenAt(src: []const u8, at: usize, sets: []const drawings.DrawingTags) Error!?AnchorMatch {
    for (sets) |*s| {
        const kind: AnchorKind, const close_needle = if (drawings.matchesOpenTag(src, at, s.open_two))
            .{ .two_cell, s.close_two }
        else if (drawings.matchesOpenTag(src, at, s.open_one))
            .{ .one_cell, s.close_one }
        else
            continue;
        // The open tag's end is quote-aware — `editAs="a/>b"` is not a
        // self-closing wrapper (in-house ND-REL-202) — and shared with
        // the read.
        const open_tag = drawings.selfClosingTagEnd(src, at) orelse return Error.MalformedDrawingXml;
        const after_open = open_tag.gt + 1;
        if (open_tag.self_closing) {
            return .{ .kind = kind, .open_start = at, .after_open = after_open, .after_close = after_open, .self_closing = true };
        }
        // XML ties the close tag to the open tag's QName: the wrapper's
        // own set spells it. A commented close is text.
        const close = drawings.findLiveMarkup(src, after_open, close_needle) orelse return Error.MalformedDrawingXml;
        return .{ .kind = kind, .open_start = at, .after_open = after_open, .after_close = close + close_needle.len, .self_closing = false };
    }
    return null;
}

/// Shift a 0-based row/column index for an insert at `edit_index`.
fn shiftForInsert(value: u32, edit_index: u32) Error!u32 {
    if (value < edit_index) return value;
    if (value == std.math.maxInt(u32)) return Error.DrawingCoordinateOverflow;
    return value + 1;
}

/// Shift a 0-based row/column index for a delete at `edit_index`.
/// `is_br_corner` shrinks the BR corner by one on a delete-match.
/// On the TL corner the index stays put on a delete-match (the
/// data shifted into the deleted slot is at the same grid index).
fn shiftForDelete(value: u32, edit_index: u32, is_br_corner: bool) u32 {
    if (value > edit_index) return value - 1;
    if (value == edit_index and is_br_corner and edit_index > 0) return value - 1;
    return value;
}

fn shifted(value: u32, index: u32, kind: EditKind, is_br_corner: bool) Error!u32 {
    return switch (kind) {
        .insert => try shiftForInsert(value, index),
        .delete => shiftForDelete(value, index, is_br_corner),
    };
}

fn axisValue(anchor: drawings.CellAnchor, axis: Axis) u32 {
    return switch (axis) {
        .row => anchor.row,
        .col => anchor.col,
    };
}

fn axisSpan(corner: drawings.CornerBlock, axis: Axis) drawings.Span {
    return switch (axis) {
        .row => corner.row_text,
        .col => corner.col_text,
    };
}

/// One corner and the value its edited-axis scalar takes.
const Splice = struct {
    corner: drawings.CornerBlock,
    new_value: u32,
};

fn processAnchor(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    a: AnchorMatch,
    sets: []const drawings.DrawingTags,
    axis: Axis,
    index: u32,
    kind: EditKind,
) Error!void {
    if (a.self_closing) {
        try out.appendSlice(allocator, src[a.open_start..a.after_open]);
        return;
    }
    const block = src[a.open_start..a.after_close];
    // The corners are read by the read's parser — a block absent or a
    // scalar that does not parse is the strict read's refusal, and the
    // sweep's: it cannot move what it cannot read.
    const content_start = a.after_open - a.open_start;
    const from = drawings.parseCornerIn(block, content_start, sets, .from) orelse return Error.MalformedDrawingXml;
    const from_value = axisValue(from.anchor, axis);

    if (a.kind == .one_cell) {
        // Drop-on-collapse: delete-match on the from coord. The
        // single anchor cell is gone.
        if (kind == .delete and from_value == index) return;
        const one = [_]Splice{.{ .corner = from, .new_value = try shifted(from_value, index, kind, false) }};
        try emitAnchor(allocator, out, block, &one, axis);
        return;
    }

    const to = drawings.parseCornerIn(block, content_start, sets, .to) orelse return Error.MalformedDrawingXml;
    const to_value = axisValue(to.anchor, axis);
    // Two blocks that overlap are not two corners — judged BEFORE the
    // collapse rule, or a nested pair at the deleted index would be
    // dropped whole instead of refused (in-house ND-REL-401).
    if (drawings.cornersOverlap(from, to)) return Error.MalformedDrawingXml;
    // Drop-on-collapse: delete that wipes both corners' value on
    // the edited axis.
    if (kind == .delete and from_value == index and to_value == index) return;
    const from_splice: Splice = .{ .corner = from, .new_value = try shifted(from_value, index, kind, false) };
    const to_splice: Splice = .{ .corner = to, .new_value = try shifted(to_value, index, kind, true) };
    // Document order decides the emit order: the schema says `from`
    // then `to`, but the read is order-agnostic and lists a reversed
    // pair, so the sweep moves it (the v1 walker sliced backwards —
    // in-house ND-REL-102).
    const ordered = if (from.open_start <= to.open_start) [_]Splice{ from_splice, to_splice } else [_]Splice{ to_splice, from_splice };
    try emitAnchor(allocator, out, block, &ordered, axis);
}

/// Emit the anchor block with each corner's edited-axis scalar spliced
/// to its new value — every other byte as written.
fn emitAnchor(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    block: []const u8,
    splices: []const Splice,
    axis: Axis,
) Error!void {
    var at: usize = 0;
    for (splices) |s| {
        const c = s.corner;
        try out.appendSlice(allocator, block[at..c.open_start]);
        if (s.new_value == axisValue(c.anchor, axis)) {
            // No rewrite needed: the block passes through verbatim.
            try out.appendSlice(allocator, block[c.open_start..c.after_close]);
        } else {
            const span = axisSpan(c, axis);
            try out.appendSlice(allocator, block[c.open_start..span.start]);
            // Splice the new value as a decimal integer. u32 max is
            // 10 digits — bufPrint into a 16-byte buffer cannot exhaust.
            var num_buf: [16]u8 = undefined;
            const new_text = std.fmt.bufPrint(&num_buf, "{d}", .{s.new_value}) catch unreachable;
            try out.appendSlice(allocator, new_text);
            try out.appendSlice(allocator, block[span.end..c.after_close]);
        }
        at = c.after_close;
    }
    try out.appendSlice(allocator, block[at..]);
}

// ---------------------------------------------------------------------------
// Tests — pure-function coverage.
// ---------------------------------------------------------------------------

const testing = std.testing;

fn wrapDrawing(allocator: Allocator, body: []const u8) ![]u8 {
    const head = "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">";
    const tail = "</xdr:wsDr>";
    var buf = try allocator.alloc(u8, head.len + body.len + tail.len);
    @memcpy(buf[0..head.len], head);
    @memcpy(buf[head.len .. head.len + body.len], body);
    @memcpy(buf[head.len + body.len ..], tail);
    return buf;
}

const sample_two = "<xdr:twoCellAnchor>" ++
    "<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>" ++
    "<xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>" ++
    "<xdr:pic><xdr:nvPicPr/><xdr:blipFill/></xdr:pic>" ++
    "<xdr:clientData/>" ++
    "</xdr:twoCellAnchor>";

const sample_one = "<xdr:oneCellAnchor>" ++
    "<xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>" ++
    "<xdr:ext cx=\"914400\" cy=\"914400\"/>" ++
    "<xdr:pic><xdr:nvPicPr/><xdr:blipFill/></xdr:pic>" ++
    "<xdr:clientData/>" ++
    "</xdr:oneCellAnchor>";

test "twoCellAnchor: insert at col before from shifts both corners right" {
    const a = testing.allocator;
    const src = try wrapDrawing(a, sample_two);
    defer a.free(src);
    const out = try applyEditToDrawing(a, src, .col, 0, .insert);
    defer a.free(out);
    // Original col=1, col=3 → expect col=2, col=4.
    try testing.expect(std.mem.indexOf(u8, out, "<xdr:col>2</xdr:col>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<xdr:col>4</xdr:col>") != null);
    // Rows untouched.
    try testing.expect(std.mem.indexOf(u8, out, "<xdr:row>4</xdr:row>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<xdr:row>10</xdr:row>") != null);
}

test "twoCellAnchor: insert at col strictly inside range expands BR" {
    const a = testing.allocator;
    const src = try wrapDrawing(a, sample_two);
    defer a.free(src);
    // Insert at col 2 (between original from=1 and to=3).
    const out = try applyEditToDrawing(a, src, .col, 2, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<xdr:col>1</xdr:col>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<xdr:col>4</xdr:col>") != null);
}

test "twoCellAnchor: insert at col past to is no-op for the anchor" {
    const a = testing.allocator;
    const src = try wrapDrawing(a, sample_two);
    defer a.free(src);
    // Insert at col 10 — well past to=3.
    const out = try applyEditToDrawing(a, src, .col, 10, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<xdr:col>1</xdr:col>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<xdr:col>3</xdr:col>") != null);
}

test "twoCellAnchor: delete inside range shrinks BR" {
    const a = testing.allocator;
    const src = try wrapDrawing(a, sample_two);
    defer a.free(src);
    // Delete at col 2 (between from=1 and to=3).
    const out = try applyEditToDrawing(a, src, .col, 2, .delete);
    defer a.free(out);
    // from=1 unchanged (1 < 2); to=3 → 2.
    try testing.expect(std.mem.indexOf(u8, out, "<xdr:col>1</xdr:col>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<xdr:col>2</xdr:col>") != null);
}

test "twoCellAnchor: delete at TL corner leaves from in place; data underneath shifts" {
    const a = testing.allocator;
    const src = try wrapDrawing(a, sample_two);
    defer a.free(src);
    // Delete at col 1 (== from). from stays 1; to 3 → 2.
    const out = try applyEditToDrawing(a, src, .col, 1, .delete);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<xdr:col>1</xdr:col>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<xdr:col>2</xdr:col>") != null);
}

test "twoCellAnchor: delete that collapses entire range drops the anchor" {
    const a = testing.allocator;
    // Anchor spans only column 2 (from=2, to=2).
    const body = "<xdr:twoCellAnchor>" ++
        "<xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>" ++
        "<xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>" ++
        "<xdr:pic/><xdr:clientData/></xdr:twoCellAnchor>";
    const src = try wrapDrawing(a, body);
    defer a.free(src);
    const out = try applyEditToDrawing(a, src, .col, 2, .delete);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<xdr:twoCellAnchor") == null);
}

test "twoCellAnchor: delete at col past to shifts only to" {
    const a = testing.allocator;
    const src = try wrapDrawing(a, sample_two);
    defer a.free(src);
    // Delete at col 5 — past to=3, so no shift.
    const out = try applyEditToDrawing(a, src, .col, 5, .delete);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<xdr:col>1</xdr:col>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<xdr:col>3</xdr:col>") != null);
}

test "twoCellAnchor: row insert shifts both corners' rows" {
    const a = testing.allocator;
    const src = try wrapDrawing(a, sample_two);
    defer a.free(src);
    // Insert at row 5 (between from=4 and to=10).
    const out = try applyEditToDrawing(a, src, .row, 5, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<xdr:row>4</xdr:row>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<xdr:row>11</xdr:row>") != null);
}

test "oneCellAnchor: insert at col before from shifts col" {
    const a = testing.allocator;
    const src = try wrapDrawing(a, sample_one);
    defer a.free(src);
    // Insert at col 0 — before from=2.
    const out = try applyEditToDrawing(a, src, .col, 0, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<xdr:col>3</xdr:col>") != null);
}

test "oneCellAnchor: delete at from drops the anchor" {
    const a = testing.allocator;
    const src = try wrapDrawing(a, sample_one);
    defer a.free(src);
    // Delete at col 2 (== from). Anchor is gone.
    const out = try applyEditToDrawing(a, src, .col, 2, .delete);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<xdr:oneCellAnchor") == null);
}

test "absoluteAnchor passes through unchanged" {
    const a = testing.allocator;
    const body = "<xdr:absoluteAnchor>" ++
        "<xdr:pos x=\"0\" y=\"0\"/><xdr:ext cx=\"914400\" cy=\"914400\"/>" ++
        "<xdr:pic/><xdr:clientData/></xdr:absoluteAnchor>";
    const src = try wrapDrawing(a, body);
    defer a.free(src);
    const out = try applyEditToDrawing(a, src, .col, 0, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<xdr:absoluteAnchor>") != null);
}

test "twoCellAnchor: rewrite preserves <xdr:to> opener (regression: emit-block bug)" {
    const a = testing.allocator;
    const src = try wrapDrawing(a, sample_two);
    defer a.free(src);
    const out = try applyEditToDrawing(a, src, .col, 0, .insert);
    defer a.free(out);
    // Both opener tags AND closer tags must survive the splice —
    // the prior emit-block code dropped <xdr:to> in valid input.
    try testing.expect(std.mem.indexOf(u8, out, "<xdr:from>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "</xdr:from>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<xdr:to>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "</xdr:to>") != null);
    // Each appears exactly once for a single anchor.
    var i: usize = 0;
    var to_open_count: usize = 0;
    while (std.mem.indexOfPos(u8, out, i, "<xdr:to>")) |idx| {
        to_open_count += 1;
        i = idx + 1;
    }
    try testing.expectEqual(@as(usize, 1), to_open_count);
}

test "twoCellAnchor: no-op rewrite is byte-identical (SHA256 passthrough)" {
    const a = testing.allocator;
    const src = try wrapDrawing(a, sample_two);
    defer a.free(src);
    // Insert at col 100 — well past the anchor's to=3. No coord changes.
    const out = try applyEditToDrawing(a, src, .col, 100, .insert);
    defer a.free(out);
    try testing.expectEqualSlices(u8, src, out);
}

test "two adjacent twoCellAnchors both rewrite correctly" {
    const a = testing.allocator;
    const body = sample_two ++ sample_two;
    const src = try wrapDrawing(a, body);
    defer a.free(src);
    const out = try applyEditToDrawing(a, src, .col, 0, .insert);
    defer a.free(out);
    // Both anchors should have shifted col=1 → 2 and col=3 → 4.
    var count_col2: usize = 0;
    var count_col4: usize = 0;
    var i: usize = 0;
    while (std.mem.indexOfPos(u8, out, i, "<xdr:col>2</xdr:col>")) |idx| {
        count_col2 += 1;
        i = idx + 1;
    }
    i = 0;
    while (std.mem.indexOfPos(u8, out, i, "<xdr:col>4</xdr:col>")) |idx| {
        count_col4 += 1;
        i = idx + 1;
    }
    try testing.expectEqual(@as(usize, 2), count_col2);
    try testing.expectEqual(@as(usize, 2), count_col4);
}

// ─── namespace-aware: the default namespace, custom + mixed prefixes ──
//
// The anchors read and this sweep share `drawings.resolveDrawingPrefixes`;
// these pin the sweep's half — every prefix the read lists under, the
// sweep shifts under, in one document-order pass.

const ns_xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const ns_xdr_strict = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
const ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main";

/// openpyxl 3.1.5's drawing for `tests/corpus/openpyxl_chart.xlsx`,
/// byte for byte: the spreadsheetDrawing namespace bound as the
/// DEFAULT one, every anchor element unprefixed, `<c:chart>` bound
/// on the root.
const openpyxl_drawing =
    "<wsDr xmlns:a=\"" ++ ns_a ++ "\" xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns=\"" ++ ns_xdr ++ "\">" ++
    "<oneCellAnchor><from><col>3</col><colOff>0</colOff><row>1</row><rowOff>0</rowOff></from><ext cx=\"5400000\" cy=\"2700000\" />" ++
    "<graphicFrame><nvGraphicFramePr><cNvPr id=\"1\" name=\"Chart 1\" /><cNvGraphicFramePr /></nvGraphicFramePr><xfrm />" ++
    "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart r:id=\"rId1\" /></a:graphicData></a:graphic></graphicFrame><clientData /></oneCellAnchor></wsDr>";

/// `sample_two` with no prefix at all — the default-namespace
/// spelling of the same two-cell anchor.
const sample_two_default = "<twoCellAnchor>" ++
    "<from><col>1</col><colOff>0</colOff><row>4</row><rowOff>0</rowOff></from>" ++
    "<to><col>3</col><colOff>0</colOff><row>10</row><rowOff>0</rowOff></to>" ++
    "<pic><nvPicPr/><blipFill/></pic>" ++
    "<clientData/>" ++
    "</twoCellAnchor>";

fn wrapWithRoot(allocator: Allocator, root_open: []const u8, body: []const u8, root_close: []const u8) ![]u8 {
    return std.mem.concat(allocator, u8, &.{ root_open, body, root_close });
}

test "default namespace (openpyxl): a row insert above the anchor shifts <row> and nothing else — the CF-DOC-501 gap closed" {
    const a = testing.allocator;
    const out = try applyEditToDrawing(a, openpyxl_drawing, .row, 0, .insert);
    defer a.free(out);
    // The one splice: `<row>1</row>` → `<row>2</row>`; every other
    // byte, the `<c:chart r:id>` stub and the `/>`-spaced empties
    // included, is as openpyxl wrote it.
    const expected = try std.mem.replaceOwned(u8, a, openpyxl_drawing, "<row>1</row>", "<row>2</row>");
    defer a.free(expected);
    try testing.expectEqualStrings(expected, out);
    // A column insert past the anchor is byte-identical; one at the
    // anchor's column moves it right.
    const same = try applyEditToDrawing(a, openpyxl_drawing, .col, 4, .insert);
    defer a.free(same);
    try testing.expectEqualStrings(openpyxl_drawing, same);
    const col = try applyEditToDrawing(a, openpyxl_drawing, .col, 3, .insert);
    defer a.free(col);
    try testing.expect(std.mem.indexOf(u8, col, "<col>4</col>") != null);
    try testing.expect(std.mem.indexOf(u8, col, "<row>1</row>") != null);
    // A delete of the anchor's own row drops the one-cell anchor
    // whole — the root's open tag and close are all that remain.
    const dropped = try applyEditToDrawing(a, openpyxl_drawing, .row, 1, .delete);
    defer a.free(dropped);
    const anchor_at = std.mem.indexOf(u8, openpyxl_drawing, "<oneCellAnchor>").?;
    const expected_dropped = try std.mem.concat(a, u8, &.{ openpyxl_drawing[0..anchor_at], "</wsDr>" });
    defer a.free(expected_dropped);
    try testing.expectEqualStrings(expected_dropped, dropped);
}

test "default namespace: a two-cell anchor shifts, shrinks and collapses under both URIs" {
    const a = testing.allocator;
    for ([_][]const u8{ ns_xdr, ns_xdr_strict }) |uri| {
        const root = try std.mem.concat(a, u8, &.{ "<wsDr xmlns=\"", uri, "\" xmlns:a=\"", ns_a, "\">" });
        defer a.free(root);
        const src = try wrapWithRoot(a, root, sample_two_default, "</wsDr>");
        defer a.free(src);
        const ins = try applyEditToDrawing(a, src, .col, 0, .insert);
        defer a.free(ins);
        try testing.expect(std.mem.indexOf(u8, ins, "<col>2</col>") != null);
        try testing.expect(std.mem.indexOf(u8, ins, "<col>4</col>") != null);
        try testing.expect(std.mem.indexOf(u8, ins, "<row>4</row>") != null);
        const del = try applyEditToDrawing(a, src, .row, 6, .delete);
        defer a.free(del);
        try testing.expect(std.mem.indexOf(u8, del, "<row>4</row>") != null);
        try testing.expect(std.mem.indexOf(u8, del, "<row>9</row>") != null);
        // Collapse: an anchor spanning one column, that column deleted.
        const narrow = try std.mem.replaceOwned(u8, a, sample_two_default, "<to><col>3</col>", "<to><col>1</col>");
        defer a.free(narrow);
        const src_narrow = try wrapWithRoot(a, root, narrow, "</wsDr>");
        defer a.free(src_narrow);
        const gone = try applyEditToDrawing(a, src_narrow, .col, 1, .delete);
        defer a.free(gone);
        try testing.expect(std.mem.indexOf(u8, gone, "<twoCellAnchor") == null);
    }
}

test "a non-canonical prefix bound to the xdr URI shifts — `dr:` where v1 saw nothing" {
    const a = testing.allocator;
    const canonical = try wrapDrawing(a, sample_two);
    defer a.free(canonical);
    // `xmlns:xdr` → `xmlns:dr`, every tag with it.
    const src = try std.mem.replaceOwned(u8, a, canonical, "xdr", "dr");
    defer a.free(src);
    const out = try applyEditToDrawing(a, src, .row, 5, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<dr:row>4</dr:row>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<dr:row>11</dr:row>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "xdr") == null);
}

test "mixed prefixes in one part shift in one document-order pass — the root's, a declared alternate, the default namespace" {
    const a = testing.allocator;
    const root = "<xdr:wsDr xmlns:xdr=\"" ++ ns_xdr ++ "\" xmlns:dr=\"" ++ ns_xdr ++ "\" xmlns:a=\"" ++ ns_a ++ "\">";
    const alt = try std.mem.replaceOwned(u8, a, sample_two, "xdr:", "dr:");
    defer a.free(alt);
    // The default namespace declared mid-document, on the anchor.
    const default_anchor = try std.mem.replaceOwned(u8, a, sample_two_default, "<twoCellAnchor>", "<twoCellAnchor xmlns=\"" ++ ns_xdr ++ "\">");
    defer a.free(default_anchor);
    const src = try std.mem.concat(a, u8, &.{ root, sample_two, alt, default_anchor, "</xdr:wsDr>" });
    defer a.free(src);
    const out = try applyEditToDrawing(a, src, .col, 0, .insert);
    defer a.free(out);
    const x = std.mem.indexOf(u8, out, "<xdr:col>2</xdr:col>").?;
    const d = std.mem.indexOf(u8, out, "<dr:col>2</dr:col>").?;
    const n = std.mem.indexOf(u8, out, "<col>2</col>").?;
    try testing.expect(x < d and d < n);
    try testing.expectEqual(@as(usize, 1), std.mem.count(u8, out, "<xdr:col>4</xdr:col>"));
    try testing.expectEqual(@as(usize, 1), std.mem.count(u8, out, "<dr:col>4</dr:col>"));
    try testing.expectEqual(@as(usize, 1), std.mem.count(u8, out, "<col>4</col>"));
    try testing.expect(std.mem.indexOf(u8, out, "<col>1</col>") == null);
    // The same input with a no-op edit is byte-identical.
    const same = try applyEditToDrawing(a, src, .col, 50, .insert);
    defer a.free(same);
    try testing.expectEqualStrings(src, same);
}

test "a prefixed root over default-namespace anchors shifts them — the default declaration is an alternate" {
    const a = testing.allocator;
    const root = "<xdr:wsDr xmlns:xdr=\"" ++ ns_xdr ++ "\" xmlns=\"" ++ ns_xdr ++ "\">";
    const src = try wrapWithRoot(a, root, sample_two_default, "</xdr:wsDr>");
    defer a.free(src);
    const out = try applyEditToDrawing(a, src, .row, 0, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<row>5</row>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<row>11</row>") != null);
}

test "an unprefixed anchor under a foreign default namespace is not an anchor — byte-identical" {
    const a = testing.allocator;
    // The root's default namespace is not a spreadsheetDrawing one;
    // `xdr:` is bound but unused. `<twoCellAnchor>` here is a
    // foreign element the read lists nothing for, so the sweep
    // leaves it as written.
    const root = "<wsDr xmlns=\"urn:not-a-drawing\" xmlns:xdr=\"" ++ ns_xdr ++ "\">";
    const src = try wrapWithRoot(a, root, sample_two_default, "</wsDr>");
    defer a.free(src);
    const out = try applyEditToDrawing(a, src, .col, 0, .insert);
    defer a.free(out);
    try testing.expectEqualStrings(src, out);
}

test "a spreadsheetDrawing binding the walk cannot spell refuses MalformedDrawingXml; the 100-byte limit itself is followed" {
    const a = testing.allocator;
    const at_limit = "p" ** 100;
    const past_limit = "p" ** 101;
    // Under the limit: the anchor under the long prefix shifts.
    {
        const root = "<xdr:wsDr xmlns:xdr=\"" ++ ns_xdr ++ "\" xmlns:" ++ at_limit ++ "=\"" ++ ns_xdr ++ "\">";
        const body = try std.mem.replaceOwned(u8, a, sample_two, "xdr:", at_limit ++ ":");
        defer a.free(body);
        const src = try wrapWithRoot(a, root, body, "</xdr:wsDr>");
        defer a.free(src);
        const out = try applyEditToDrawing(a, src, .col, 0, .insert);
        defer a.free(out);
        try testing.expect(std.mem.indexOf(u8, out, "<" ++ at_limit ++ ":col>2</" ++ at_limit ++ ":col>") != null);
    }
    // Past it: the binding is unfollowed, so the edit refuses rather
    // than leave that anchor behind — whether or not any anchor uses
    // the name (the read refuses the same bytes under strict).
    {
        const root = "<xdr:wsDr xmlns:xdr=\"" ++ ns_xdr ++ "\" xmlns:" ++ past_limit ++ "=\"" ++ ns_xdr ++ "\">";
        const src = try wrapWithRoot(a, root, sample_two, "</xdr:wsDr>");
        defer a.free(src);
        try testing.expectError(error.MalformedDrawingXml, applyEditToDrawing(a, src, .col, 0, .insert));
        try testing.expectError(error.MalformedDrawingXml, applyEditToDrawing(a, src, .row, 100, .delete));
    }
    // A ninth alternate is past the replay cap: refused too. A
    // commented decoy past the limit is text, not a binding.
    {
        var root: std.ArrayListUnmanaged(u8) = .empty;
        defer root.deinit(a);
        try root.appendSlice(a, "<xdr:wsDr xmlns:xdr=\"" ++ ns_xdr ++ "\"");
        for (0..drawings.max_xdr_alts + 1) |k| {
            var decl_buf: [128]u8 = undefined;
            try root.appendSlice(a, try std.fmt.bufPrint(&decl_buf, " xmlns:p{d}=\"{s}\"", .{ k, ns_xdr }));
        }
        try root.appendSlice(a, ">");
        const src = try wrapWithRoot(a, root.items, sample_two, "</xdr:wsDr>");
        defer a.free(src);
        try testing.expectError(error.MalformedDrawingXml, applyEditToDrawing(a, src, .col, 0, .insert));
    }
    {
        const root = "<xdr:wsDr xmlns:xdr=\"" ++ ns_xdr ++ "\"><!-- xmlns:" ++ past_limit ++ "=\"" ++ ns_xdr ++ "\" -->";
        const src = try wrapWithRoot(a, root, sample_two, "</xdr:wsDr>");
        defer a.free(src);
        const out = try applyEditToDrawing(a, src, .col, 0, .insert);
        defer a.free(out);
        try testing.expect(std.mem.indexOf(u8, out, "<xdr:col>2</xdr:col>") != null);
    }
}

test "a <to> before <from> moves both corners in document order; nested blocks refuse (ND-REL-102)" {
    const a = testing.allocator;
    // The v1 walker sliced backwards on this shape (a panic in safe
    // builds); the read lists it, so the sweep moves it.
    const reversed = "<xdr:twoCellAnchor>" ++
        "<xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>" ++
        "<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>" ++
        "<xdr:pic/><xdr:clientData/></xdr:twoCellAnchor>";
    const src = try wrapDrawing(a, reversed);
    defer a.free(src);
    const out = try applyEditToDrawing(a, src, .row, 5, .insert);
    defer a.free(out);
    const expected = try std.mem.replaceOwned(u8, a, src, "<xdr:row>10</xdr:row>", "<xdr:row>11</xdr:row>");
    defer a.free(expected);
    try testing.expectEqualStrings(expected, out);
    // Both corners, a column insert before both.
    const both = try applyEditToDrawing(a, src, .col, 0, .insert);
    defer a.free(both);
    try testing.expect(std.mem.indexOf(u8, both, "<xdr:to><xdr:col>4</xdr:col>") != null);
    try testing.expect(std.mem.indexOf(u8, both, "<xdr:from><xdr:col>2</xdr:col>") != null);
    // The default-namespace twin.
    const bare = try std.mem.replaceOwned(u8, a, reversed, "xdr:", "");
    defer a.free(bare);
    const src_bare = try wrapWithRoot(a, "<wsDr xmlns=\"" ++ ns_xdr ++ "\">", bare, "</wsDr>");
    defer a.free(src_bare);
    const out_bare = try applyEditToDrawing(a, src_bare, .row, 5, .insert);
    defer a.free(out_bare);
    try testing.expect(std.mem.indexOf(u8, out_bare, "<row>11</row>") != null);
    try testing.expect(std.mem.indexOf(u8, out_bare, "<row>4</row>") != null);
    // A `<to>` nested inside `<from>` is not two corners.
    const nested = "<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff>" ++
        "<xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to></xdr:from></xdr:twoCellAnchor>";
    const src_nested = try wrapDrawing(a, nested);
    defer a.free(src_nested);
    try testing.expectError(error.MalformedDrawingXml, applyEditToDrawing(a, src_nested, .row, 5, .insert));
    // …under every edit, the delete that would collapse it included:
    // the overlap is judged before the drop rule (ND-REL-401).
    const collapsing = try std.mem.replaceOwned(u8, a, src_nested, "<xdr:row>10</xdr:row>", "<xdr:row>4</xdr:row>");
    defer a.free(collapsing);
    try testing.expectError(error.MalformedDrawingXml, applyEditToDrawing(a, collapsing, .row, 4, .delete));
    try testing.expectError(error.MalformedDrawingXml, applyEditToDrawing(a, src_nested, .col, 7, .delete));
}

test "a wrapper under one followed spelling with its children under the other shifts — both directions (ND-REL-103)" {
    const a = testing.allocator;
    const root = "<xdr:wsDr xmlns:xdr=\"" ++ ns_xdr ++ "\" xmlns=\"" ++ ns_xdr ++ "\">";
    // A prefixed wrapper over bare corners…
    {
        const body = "<xdr:twoCellAnchor editAs=\"oneCell\">" ++
            "<from><col>1</col><colOff>0</colOff><row>4</row><rowOff>0</rowOff></from>" ++
            "<to><col>3</col><colOff>0</colOff><row>10</row><rowOff>0</rowOff></to>" ++
            "<pic/><clientData/></xdr:twoCellAnchor>";
        const src = try wrapWithRoot(a, root, body, "</xdr:wsDr>");
        defer a.free(src);
        const out = try applyEditToDrawing(a, src, .col, 0, .insert);
        defer a.free(out);
        try testing.expect(std.mem.indexOf(u8, out, "<col>2</col>") != null);
        try testing.expect(std.mem.indexOf(u8, out, "<col>4</col>") != null);
        try testing.expect(std.mem.indexOf(u8, out, "<col>1</col>") == null);
    }
    // …and a bare wrapper over prefixed corners, the scalars mixed
    // inside one block too.
    {
        const body = "<oneCellAnchor>" ++
            "<xdr:from><col>2</col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><rowOff>0</rowOff></xdr:from>" ++
            "<ext cx=\"1\" cy=\"1\"/><pic/><clientData/></oneCellAnchor>";
        const src = try wrapWithRoot(a, root, body, "</xdr:wsDr>");
        defer a.free(src);
        const out = try applyEditToDrawing(a, src, .row, 0, .insert);
        defer a.free(out);
        try testing.expect(std.mem.indexOf(u8, out, "<xdr:row>6</xdr:row>") != null);
        const col = try applyEditToDrawing(a, src, .col, 0, .insert);
        defer a.free(col);
        try testing.expect(std.mem.indexOf(u8, col, "<col>3</col>") != null);
    }
}

test "comments, CDATA and PIs are copied whole and never matched — under the default namespace too (ND-REL-101)" {
    const a = testing.allocator;
    const root = "<wsDr xmlns=\"" ++ ns_xdr ++ "\" xmlns:a=\"" ++ ns_a ++ "\">";
    const one = "<oneCellAnchor><from><col>0</col><colOff>0</colOff><row>0</row><rowOff>0</rowOff></from><ext cx=\"1\" cy=\"1\"/><pic/><clientData/></oneCellAnchor>";
    const decoys = "<!-- a stray <twoCellAnchor><from><col>9</col></from> mention --><![CDATA[<oneCellAnchor><from><row>9</row></from>]]><?note <oneCellAnchor> ?>";
    const src = try std.mem.concat(a, u8, &.{ root, decoys, one, sample_two_default, "<!-- TODO: a second <oneCellAnchor> goes here -->", "</wsDr>" });
    defer a.free(src);
    // A row insert at 0: the one-cell anchor and BOTH corners of the
    // two-cell one move; the decoys are byte-identical (the v1 walker
    // took the commented wrapper for an anchor and spliced the real
    // one-cell's `from` onto the two-cell's `to`).
    const out = try applyEditToDrawing(a, src, .row, 0, .insert);
    defer a.free(out);
    try testing.expect(std.mem.startsWith(u8, out, root ++ decoys));
    try testing.expect(std.mem.endsWith(u8, out, "<!-- TODO: a second <oneCellAnchor> goes here --></wsDr>"));
    try testing.expect(std.mem.indexOf(u8, out, "<from><col>0</col><colOff>0</colOff><row>1</row>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<from><col>1</col><colOff>0</colOff><row>5</row>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<to><col>3</col><colOff>0</colOff><row>11</row>") != null);
    // A commented close does not end a live anchor.
    const commented_close = "<twoCellAnchor><!-- </twoCellAnchor> -->" ++ sample_two_default["<twoCellAnchor>".len..];
    const src2 = try wrapWithRoot(a, root, commented_close, "</wsDr>");
    defer a.free(src2);
    const out2 = try applyEditToDrawing(a, src2, .col, 0, .insert);
    defer a.free(out2);
    try testing.expect(std.mem.indexOf(u8, out2, "<!-- </twoCellAnchor> -->") != null);
    try testing.expect(std.mem.indexOf(u8, out2, "<col>2</col>") != null);
    try testing.expect(std.mem.indexOf(u8, out2, "<col>4</col>") != null);
    // An unterminated comment swallows the rest of the part: nothing
    // after it is an anchor, and the bytes pass through (the read lists
    // nothing for them either).
    const src3 = try wrapWithRoot(a, root, "<!-- unterminated " ++ sample_two_default, "</wsDr>");
    defer a.free(src3);
    const out3 = try applyEditToDrawing(a, src3, .col, 0, .insert);
    defer a.free(out3);
    try testing.expectEqualStrings(src3, out3);
}

test "what the sweep cannot move it refuses — MalformedDrawingXml, the strict read's verdicts; a self-closing wrapper passes through" {
    const a = testing.allocator;
    const cases = [_][]const u8{
        // A wrapper with no close.
        "<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>",
        "<xdr:oneCellAnchor editAs=\"oneCell\">",
        // A two-cell anchor without `<to>`.
        "<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from></xdr:twoCellAnchor>",
        // A one-cell anchor without `<from>`.
        "<xdr:oneCellAnchor><xdr:ext cx=\"1\" cy=\"1\"/><xdr:pic/></xdr:oneCellAnchor>",
        // A scalar that does not parse (the read's grammar: an empty
        // body, a non-digit).
        "<xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff></xdr:rowOff></xdr:from></xdr:oneCellAnchor>",
        "<xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4x</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from></xdr:oneCellAnchor>",
        // A DTD.
        "<!DOCTYPE wsDr [ <!ENTITY g \"<oneCellAnchor><from><col>0</col></from></oneCellAnchor>\"> ]>" ++ sample_one,
        // A scalar missing — the row under a column edit too: the
        // read requires all four.
        "<xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:rowOff>0</xdr:rowOff></xdr:from></xdr:oneCellAnchor>",
    };
    for (cases) |body| {
        const src = try wrapDrawing(a, body);
        defer a.free(src);
        try testing.expectError(error.MalformedDrawingXml, applyEditToDrawing(a, src, .col, 0, .insert));
        try testing.expectError(error.MalformedDrawingXml, applyEditToDrawing(a, src, .row, 100, .delete));
    }
    // Whitespace around a scalar is XSD-collapsed, not a refusal: the
    // value moves, the body is replaced whole (in-house ND-REL-201).
    {
        const padded = try wrapDrawing(a, "<xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row> 4 </xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx=\"1\" cy=\"1\"/><xdr:pic/></xdr:oneCellAnchor>");
        defer a.free(padded);
        const moved = try applyEditToDrawing(a, padded, .row, 0, .insert);
        defer a.free(moved);
        const expected = try std.mem.replaceOwned(u8, a, padded, "<xdr:row> 4 </xdr:row>", "<xdr:row>5</xdr:row>");
        defer a.free(expected);
        try testing.expectEqualStrings(expected, moved);
        // …and a column edit leaves the padded row as written.
        const col = try applyEditToDrawing(a, padded, .col, 0, .insert);
        defer a.free(col);
        try testing.expect(std.mem.indexOf(u8, col, "<xdr:row> 4 </xdr:row>") != null);
        try testing.expect(std.mem.indexOf(u8, col, "<xdr:col>2</xdr:col>") != null);
    }
    // A `<` inside an attribute value is not well-formed XML — a close
    // or a corner spelled there would be taken for the real one — so
    // the part refuses, as the strict read refuses it (ND-REL-302 /
    // ND-REL-411); a digit separator is not a digit (`1_0` is not 10 —
    // ND-REL-306).
    {
        const attr_close = try wrapDrawing(a, "<xdr:oneCellAnchor editAs=\"</xdr:oneCellAnchor>\"><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx=\"1\" cy=\"1\"/><xdr:pic/></xdr:oneCellAnchor>");
        defer a.free(attr_close);
        try testing.expectError(error.MalformedDrawingXml, applyEditToDrawing(a, attr_close, .row, 0, .insert));
        const attr_corner = try wrapDrawing(a, "<xdr:oneCellAnchor editAs=\"<xdr:from><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>99</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>\"><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:pic/></xdr:oneCellAnchor>");
        defer a.free(attr_corner);
        try testing.expectError(error.MalformedDrawingXml, applyEditToDrawing(a, attr_corner, .col, 0, .insert));
        const attr_pic = try wrapDrawing(a, "<xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:pic macro=\"<a:blip r:embed='rId9'/>\"/></xdr:oneCellAnchor>");
        defer a.free(attr_pic);
        try testing.expectError(error.MalformedDrawingXml, applyEditToDrawing(a, attr_pic, .row, 0, .insert));
        // An escaped one is text, and moves.
        const escaped = try wrapDrawing(a, "<xdr:oneCellAnchor editAs=\"&lt;xdr:from&gt;\"><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:pic/></xdr:oneCellAnchor>");
        defer a.free(escaped);
        const moved = try applyEditToDrawing(a, escaped, .row, 0, .insert);
        defer a.free(moved);
        try testing.expect(std.mem.indexOf(u8, moved, "editAs=\"&lt;xdr:from&gt;\"><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row>") != null);
        const sep = try wrapDrawing(a, "<xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1_0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:pic/></xdr:oneCellAnchor>");
        defer a.free(sep);
        try testing.expectError(error.MalformedDrawingXml, applyEditToDrawing(a, sep, .row, 0, .insert));
    }
    // A `/>` inside an attribute value does not make the wrapper
    // self-closing (in-house ND-REL-202).
    {
        const quoted = try wrapDrawing(a, "<xdr:oneCellAnchor editAs=\"a/>b\"><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx=\"1\" cy=\"1\"/><xdr:pic/></xdr:oneCellAnchor>");
        defer a.free(quoted);
        const moved = try applyEditToDrawing(a, quoted, .row, 0, .insert);
        defer a.free(moved);
        try testing.expect(std.mem.indexOf(u8, moved, "editAs=\"a/>b\"><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row>") != null);
    }
    // A self-closing wrapper followed by a real one of the same name:
    // the real one moves (the read steps over the empty wrapper and
    // lists the real one — ND-DOC-204).
    {
        const pair = try wrapDrawing(a, "<xdr:twoCellAnchor/>" ++ sample_two);
        defer a.free(pair);
        const moved = try applyEditToDrawing(a, pair, .col, 0, .insert);
        defer a.free(moved);
        try testing.expect(std.mem.startsWith(u8, moved, pair[0..std.mem.indexOf(u8, pair, sample_two).?]));
        try testing.expect(std.mem.indexOf(u8, moved, "<xdr:col>2</xdr:col>") != null);
        try testing.expect(std.mem.indexOf(u8, moved, "<xdr:col>4</xdr:col>") != null);
    }
    // An opening tag the part ends inside.
    try testing.expectError(error.MalformedDrawingXml, applyEditToDrawing(a, "<xdr:wsDr xmlns:xdr=\"" ++ ns_xdr ++ "\"><xdr:oneCellAnchor editAs=\"oneCell\"", .col, 0, .insert));
    // `<xdr:oneCellAnchor</xdr:wsDr>` — a `<` inside a tag — is not
    // well-formed: refused, as the strict read refuses it (round 5).
    const glued = try wrapDrawing(a, "<xdr:oneCellAnchor");
    defer a.free(glued);
    try testing.expectError(error.MalformedDrawingXml, applyEditToDrawing(a, glued, .col, 0, .insert));
    // A self-closing wrapper carries nothing to move.
    const src = try wrapDrawing(a, "<xdr:twoCellAnchor/><xdr:oneCellAnchor />" ++ sample_two);
    defer a.free(src);
    const out = try applyEditToDrawing(a, src, .col, 0, .insert);
    defer a.free(out);
    try testing.expect(std.mem.startsWith(u8, out, src[0..std.mem.indexOf(u8, src, sample_two).?]));
    try testing.expect(std.mem.indexOf(u8, out, "<xdr:col>2</xdr:col>") != null);
}

// ─── fuzz target ────────────────────────────────────────────────────
//
// See the note in sheet_edit.zig. This walker parses nested
// `<xdr:from>` / `<xdr:to>` blocks with 0-based coordinates, so its
// index arithmetic differs from the A1-based walkers.

fn fuzzDrawingEditTarget(_: void, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();
    var smith_buf: [4096]u8 = undefined;
    const input = smith_buf[0..smith.slice(&smith_buf)];

    var arena = std.heap.ArenaAllocator.init(std.testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    for ([_]u32{ 1, 2, 5 }) |idx| {
        inline for (.{ .row, .col }) |axis| {
            inline for (.{ .insert, .delete }) |kind| {
                if (applyEditToDrawing(a, input, axis, idx, kind)) |out| {
                    a.free(out);
                } else |_| {}
            }
        }
    }
}

test "fuzz: applyEditToDrawing never crashes on adversarial XML" {
    try std.testing.fuzz({}, fuzzDrawingEditTarget, .{
        .corpus = &[_][]const u8{
            "",
            "<xdr:wsDr/>",
            "<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:row>2</xdr:row></xdr:to></xdr:twoCellAnchor>",
            "<xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:row>1</xdr:row></xdr:from></xdr:oneCellAnchor>",
            "<xdr:absoluteAnchor><xdr:pos x=\"0\" y=\"0\"/></xdr:absoluteAnchor>",
            // Both corners present, reversed (the ND-REL-102 shape), and
            // nested.
            "<xdr:twoCellAnchor><xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from></xdr:twoCellAnchor>",
            "<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:to><xdr:col>2</xdr:col></xdr:to></xdr:from></xdr:twoCellAnchor>",
            // The default namespace: every spelling above, bare.
            "<wsDr xmlns=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"><oneCellAnchor><from><col>1</col><colOff>0</colOff><row>1</row><rowOff>0</rowOff></from></oneCellAnchor></wsDr>",
            "<wsDr xmlns=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"><twoCellAnchor><to><col>2</col></to><from><col>0</col></from></twoCellAnchor></wsDr>",
            "<wsDr xmlns=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"><!-- <twoCellAnchor> --><oneCellAnchor><from><col>1</col><colOff>0</colOff><row>1</row><rowOff>0</rowOff></from></oneCellAnchor><![CDATA[</oneCellAnchor>]]></wsDr>",
            "<wsDr xmlns=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"><!-- unterminated <oneCellAnchor>",
            "<wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"><xdr:oneCellAnchor><from><xdr:col>1</xdr:col><colOff>0</colOff><row>1</row><rowOff>0</rowOff></from></xdr:oneCellAnchor></wsDr>",
            // A DTD with an anchor-shaped entity value; a `/>` in a value.
            "<!DOCTYPE wsDr [ <!ENTITY x \"<oneCellAnchor><from><row>1</row></from>\"> ]><wsDr xmlns=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"/>",
            "<xdr:twoCellAnchor editAs=\"a/>b\"><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to></xdr:twoCellAnchor>",
            // from without to, and the reverse.
            "<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col></xdr:from></xdr:twoCellAnchor>",
            "<xdr:twoCellAnchor><xdr:to><xdr:col>2</xdr:col></xdr:to></xdr:twoCellAnchor>",
            // Non-numeric and extreme coordinates.
            "<xdr:from><xdr:col>abc</xdr:col><xdr:row>-1</xdr:row></xdr:from>",
            "<xdr:from><xdr:col>4294967296</xdr:col></xdr:from>",
            "<xdr:from><xdr:col></xdr:col></xdr:from>",
            "<xdr:from><xdr:col>16383</xdr:col><xdr:row>1048575</xdr:row></xdr:from>",
            // Truncations.
            "<xdr:twoCellAnchor><xdr:from><xdr:col>0",
            "<xdr:twoCellAnchor><xdr:from",
            "<xdr:col>",
            "<xdr:",
            "<xdr:twoCellAnchor",
        },
    });
}
