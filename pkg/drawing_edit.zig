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
//! Namespace prefixes: `xdr:` above is the canonical spelling, not a
//! literal. The anchors are matched under every prefix the anchors
//! read follows — `drawings.resolveDrawingPrefixes`, ONE resolution
//! of the same bytes for the read and this sweep: the root element's
//! prefix, every alternate bound to a spreadsheetDrawing URI anywhere
//! in the part, and the DEFAULT namespace as the empty prefix
//! (`<wsDr xmlns="…/spreadsheetDrawing"><oneCellAnchor><from><row>`,
//! openpyxl 3.1's spelling, whose anchors the `xdr:`-literal v1 of
//! this sweep left in place while the grid and — since the chart
//! `<c:f>` sweep — the chart's series formulas moved: in-house
//! chart-sweep round 5 CF-DOC-501, the edit-side half of the
//! namespace-aware drawing slice). One pass in document order; an
//! anchor's inner `from` / `to` / `col` / `row` elements are spelled
//! under the anchor wrapper's own prefix (the read's rule). A
//! binding the resolver cannot follow — a prefix longer than its
//! 100-byte limit, or past its replay cap — refuses
//! `MalformedDrawingXml`: an anchor under it would be left behind,
//! the silent corruption this sweep exists to prevent, and
//! `Workbook.preflightDrawingEditForSheet` raises it before the
//! edit's first mutation. An unprefixed anchor under a root whose
//! default namespace is NOT a spreadsheetDrawing one is not an
//! anchor and passes through, as the read lists nothing for it.

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
/// allocated buffer. Refuses `MalformedDrawingXml` for a part that
/// binds a spreadsheetDrawing namespace under a name the walk cannot
/// spell (`DrawingPrefixes.xdr_rejected`) — before reading an anchor,
/// so a dry run is the refusal the apply would raise.
pub fn applyEditToDrawing(
    allocator: Allocator,
    src: []const u8,
    axis: Axis,
    index: u32,
    kind: EditKind,
) Error![]u8 {
    const prefixes = drawings.resolveDrawingPrefixes(src);
    if (prefixes.xdr_rejected) return Error.MalformedDrawingXml;
    const set = PrefixSet.init(&prefixes);

    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(allocator);

    var i: usize = 0;
    while (i < src.len) {
        const next_lt = std.mem.indexOfScalarPos(u8, src, i, '<') orelse {
            try out.appendSlice(allocator, src[i..]);
            return try out.toOwnedSlice(allocator);
        };
        try out.appendSlice(allocator, src[i..next_lt]);
        i = next_lt;

        if (matchAnchorOpenAt(src, i, set.items(), "twoCellAnchor")) |a| {
            try processTwoCellAnchor(allocator, &out, src, a, axis, index, kind, &i);
        } else if (matchAnchorOpenAt(src, i, set.items(), "oneCellAnchor")) |a| {
            try processOneCellAnchor(allocator, &out, src, a, axis, index, kind, &i);
        } else {
            try out.append(allocator, '<');
            i += 1;
        }
    }
    return try out.toOwnedSlice(allocator);
}

/// The spreadsheetDrawing prefixes one walk spells: the resolver's
/// primary first, then its alternates — the read's replay set, in
/// one pass here because a rewrite must keep document order.
const PrefixSet = struct {
    buf: [1 + drawings.max_xdr_alts][]const u8 = undefined,
    len: usize = 0,

    fn init(p: *const drawings.DrawingPrefixes) PrefixSet {
        var s: PrefixSet = .{};
        s.buf[0] = p.xdr;
        s.len = 1;
        for (p.xdr_alts()) |alt| {
            s.buf[s.len] = alt;
            s.len += 1;
        }
        return s;
    }

    fn items(self: *const PrefixSet) []const []const u8 {
        return self.buf[0..self.len];
    }
};

/// Scratch for one spelled needle: `</PREFIX:twoCellAnchor>` under
/// the resolver's 100-byte prefix limit is 117 bytes.
const needle_buf_len: usize = 128;

const AnchorMatch = struct {
    /// Absolute byte offset of the opening `<xdr:NAME` token.
    open_start: usize,
    /// Absolute byte offset just past the opening tag's `>`.
    after_open: usize,
    /// Absolute byte offset of `</xdr:NAME>` (or end of src on
    /// malformed XML — caller handles).
    close_start: usize,
    /// Absolute byte offset just past the closing tag.
    after_close: usize,
    /// True when the opening tag was self-closing (`<xdr:foo/>`),
    /// which is malformed for an anchor element (per ECMA-376 the
    /// anchor MUST contain a `<xdr:from>` etc) — but we tolerate
    /// it by emitting and skipping the empty anchor verbatim.
    self_closing: bool,
    /// The prefix the wrapper was matched under — empty for the
    /// default namespace; its inner elements are spelled under it.
    prefix: []const u8,
};

/// Match `<{p}:NAME` at `i` for the first `p` in `prefixes` that
/// spells it and locate the matching closing tag. `name` is the bare
/// element name (e.g. `twoCellAnchor`).
fn matchAnchorOpenAt(src: []const u8, i: usize, prefixes: []const []const u8, name: []const u8) ?AnchorMatch {
    if (i >= src.len or src[i] != '<') return null;
    for (prefixes) |prefix| {
        if (matchAnchorOpenUnder(src, i, prefix, name)) |m| return m;
    }
    return null;
}

fn matchAnchorOpenUnder(src: []const u8, i: usize, prefix: []const u8, name: []const u8) ?AnchorMatch {
    var open_buf: [needle_buf_len]u8 = undefined;
    const open = drawings.spellQName(&open_buf, "<", prefix, name, "") catch return null;
    if (!std.mem.startsWith(u8, src[i..], open)) return null;
    const after_name = i + open.len;
    if (after_name >= src.len) return null;
    const c = src[after_name];
    // Disambiguate from longer element names sharing a prefix
    // (e.g. `twoCellAnchorThing` would otherwise match
    // `twoCellAnchor`). Anchors only legitimately end with one of
    // these byte classes.
    if (c != ' ' and c != '\t' and c != '\n' and c != '\r' and c != '/' and c != '>') return null;
    const gt = std.mem.indexOfScalarPos(u8, src, i, '>') orelse return null;
    const after_open = gt + 1;

    // Self-closing form: `<xdr:foo/>` — locate the slash before `>`.
    var trim_end = gt;
    while (trim_end > after_name) : (trim_end -= 1) {
        const ch = src[trim_end - 1];
        if (ch == ' ' or ch == '\t' or ch == '\n' or ch == '\r') continue;
        break;
    }
    const self_closing = trim_end > after_name and src[trim_end - 1] == '/';

    if (self_closing) {
        return .{
            .open_start = i,
            .after_open = after_open,
            .close_start = after_open,
            .after_close = after_open,
            .self_closing = true,
            .prefix = prefix,
        };
    }

    // Locate `</xdr:NAME>` close.
    var close_buf: [needle_buf_len]u8 = undefined;
    const close_needle = drawings.spellQName(&close_buf, "</", prefix, name, ">") catch return null;
    const close = std.mem.indexOfPos(u8, src, after_open, close_needle) orelse return null;
    return .{
        .open_start = i,
        .after_open = after_open,
        .close_start = close,
        .after_close = close + close_needle.len,
        .self_closing = false,
        .prefix = prefix,
    };
}

/// Locate `<xdr:NAME>...</xdr:NAME>` strictly inside `[lo, hi)`
/// and return (text-content-start, text-content-end, after-close).
/// Returns null if either tag is absent. The text content is the
/// span between `>` of the open and `<` of the close; we return
/// inner-only because the only callers parse a plain integer.
const InnerSpan = struct {
    text_start: usize,
    text_end: usize,
    /// Absolute offset just past the closing tag's `>`.
    after_close: usize,
};

fn findInnerInt(src: []const u8, lo: usize, hi: usize, prefix: []const u8, name: []const u8) ?InnerSpan {
    var open_buf: [needle_buf_len]u8 = undefined;
    var close_buf: [needle_buf_len]u8 = undefined;
    const open_needle = drawings.spellQName(&open_buf, "<", prefix, name, ">") catch return null;
    const close_needle = drawings.spellQName(&close_buf, "</", prefix, name, ">") catch return null;

    const open_pos = std.mem.indexOfPos(u8, src, lo, open_needle) orelse return null;
    if (open_pos >= hi) return null;
    const text_start = open_pos + open_needle.len;
    const close_pos = std.mem.indexOfPos(u8, src, text_start, close_needle) orelse return null;
    if (close_pos >= hi) return null;
    return .{
        .text_start = text_start,
        .text_end = close_pos,
        .after_close = close_pos + close_needle.len,
    };
}

/// Trim the standard XML whitespace set from both ends of `text`.
fn trimXml(text: []const u8) []const u8 {
    return std.mem.trim(u8, text, " \t\n\r");
}

/// Parse a non-negative decimal integer, ignoring leading +/0
/// (per OOXML schema xsd:unsignedInt). Returns null on any
/// non-digit char (after the optional `+`).
fn parseUInt(text: []const u8) ?u32 {
    const trimmed = trimXml(text);
    if (trimmed.len == 0) return null;
    var s = trimmed;
    if (s[0] == '+') s = s[1..];
    if (s.len == 0) return null;
    return std.fmt.parseInt(u32, s, 10) catch null;
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

/// Get the sub-element name we need to read inside `<xdr:from>` /
/// `<xdr:to>` for the given axis.
fn axisInnerName(axis: Axis) []const u8 {
    return switch (axis) {
        .row => "row",
        .col => "col",
    };
}

/// Result of pre-scanning a from/to corner for the axis under
/// edit.
const CornerCoord = struct {
    span: InnerSpan,
    value: u32,
};

/// Inside the corner block bounded by `[corner_lo, corner_hi)`,
/// find the axis sub-element and parse it. Returns null if the
/// element is missing or unparseable; callers fall back to
/// passing the corner through verbatim in that case.
fn readCorner(src: []const u8, corner_lo: usize, corner_hi: usize, prefix: []const u8, axis: Axis) ?CornerCoord {
    const span = findInnerInt(src, corner_lo, corner_hi, prefix, axisInnerName(axis)) orelse return null;
    const text = src[span.text_start..span.text_end];
    const value = parseUInt(text) orelse return null;
    return .{ .span = span, .value = value };
}

/// Locate `<xdr:from>` / `<xdr:to>` block bounds inside the
/// anchor body `[lo, hi)`. Returns null if the block isn't
/// present.
const BlockBounds = struct {
    /// Byte offset of the `<` of the opening tag.
    open_start: usize,
    /// Byte offset just past the `>` of the opening tag (start
    /// of the inner text content).
    lo: usize,
    /// Byte offset of the `<` of the closing tag.
    hi: usize,
    /// Byte offset just past the `>` of the closing tag.
    after_close: usize,
};

fn findCornerBlock(src: []const u8, lo: usize, hi: usize, prefix: []const u8, name: []const u8) ?BlockBounds {
    var open_buf: [needle_buf_len]u8 = undefined;
    var close_buf: [needle_buf_len]u8 = undefined;
    const open_needle = drawings.spellQName(&open_buf, "<", prefix, name, ">") catch return null;
    const close_needle = drawings.spellQName(&close_buf, "</", prefix, name, ">") catch return null;

    const open_pos = std.mem.indexOfPos(u8, src, lo, open_needle) orelse return null;
    if (open_pos >= hi) return null;
    const inner_lo = open_pos + open_needle.len;
    const close_pos = std.mem.indexOfPos(u8, src, inner_lo, close_needle) orelse return null;
    if (close_pos >= hi) return null;
    return .{ .open_start = open_pos, .lo = inner_lo, .hi = close_pos, .after_close = close_pos + close_needle.len };
}

fn processTwoCellAnchor(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    a: AnchorMatch,
    axis: Axis,
    index: u32,
    kind: EditKind,
    cursor: *usize,
) Error!void {
    if (a.self_closing) {
        try out.appendSlice(allocator, src[a.open_start..a.after_open]);
        cursor.* = a.after_open;
        return;
    }

    const body_lo = a.after_open;
    const body_hi = a.close_start;

    // Locate from + to corner blocks. If either is missing the
    // anchor is malformed; pass through verbatim and let downstream
    // readers tolerate it.
    const from_block = findCornerBlock(src, body_lo, body_hi, a.prefix, "from") orelse {
        try out.appendSlice(allocator, src[a.open_start..a.after_close]);
        cursor.* = a.after_close;
        return;
    };
    const to_block = findCornerBlock(src, body_lo, body_hi, a.prefix, "to") orelse {
        try out.appendSlice(allocator, src[a.open_start..a.after_close]);
        cursor.* = a.after_close;
        return;
    };

    const from_corner = readCorner(src, from_block.lo, from_block.hi, a.prefix, axis);
    const to_corner = readCorner(src, to_block.lo, to_block.hi, a.prefix, axis);

    // Drop-on-collapse: delete that wipes both corners' value on
    // the edited axis.
    if (kind == .delete and from_corner != null and to_corner != null) {
        if (from_corner.?.value == index and to_corner.?.value == index) {
            cursor.* = a.after_close;
            return;
        }
    }

    // Compute new values. Skip rewriting either corner whose
    // sub-element wasn't parseable; emit that corner verbatim.
    var new_from: ?u32 = null;
    var new_to: ?u32 = null;
    if (from_corner) |fc| {
        new_from = switch (kind) {
            .insert => try shiftForInsert(fc.value, index),
            .delete => shiftForDelete(fc.value, index, false),
        };
    }
    if (to_corner) |tc| {
        new_to = switch (kind) {
            .insert => try shiftForInsert(tc.value, index),
            .delete => shiftForDelete(tc.value, index, true),
        };
    }

    // Emit the anchor open tag verbatim.
    try out.appendSlice(allocator, src[a.open_start..a.after_open]);

    // Body layout (unchanged regions emit verbatim; corner
    // bodies emit spliced):
    //   src[body_lo .. from_block.open_start]      lead-in to <xdr:from>
    //   <xdr:from>...spliced...</xdr:from>          from-block (incl. tags)
    //   src[from_close_end .. to_block.open_start]  between from + to
    //   <xdr:to>...spliced...</xdr:to>              to-block (incl. tags)
    //   src[to_close_end .. a.after_close]          tail (e.g. <xdr:pic>...</xdr:twoCellAnchor>)
    try out.appendSlice(allocator, src[body_lo..from_block.open_start]);
    try emitCornerBlock(allocator, out, src, from_block, from_corner, new_from);
    try out.appendSlice(allocator, src[from_block.after_close..to_block.open_start]);
    try emitCornerBlock(allocator, out, src, to_block, to_corner, new_to);
    try out.appendSlice(allocator, src[to_block.after_close..a.after_close]);
    cursor.* = a.after_close;
}

fn processOneCellAnchor(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    a: AnchorMatch,
    axis: Axis,
    index: u32,
    kind: EditKind,
    cursor: *usize,
) Error!void {
    if (a.self_closing) {
        try out.appendSlice(allocator, src[a.open_start..a.after_open]);
        cursor.* = a.after_open;
        return;
    }

    const body_lo = a.after_open;
    const body_hi = a.close_start;

    const from_block = findCornerBlock(src, body_lo, body_hi, a.prefix, "from") orelse {
        try out.appendSlice(allocator, src[a.open_start..a.after_close]);
        cursor.* = a.after_close;
        return;
    };
    const from_corner = readCorner(src, from_block.lo, from_block.hi, a.prefix, axis);

    // Drop-on-collapse: delete-match on the from coord. The
    // single anchor cell is gone.
    if (kind == .delete and from_corner != null and from_corner.?.value == index) {
        cursor.* = a.after_close;
        return;
    }

    var new_from: ?u32 = null;
    if (from_corner) |fc| {
        new_from = switch (kind) {
            .insert => try shiftForInsert(fc.value, index),
            .delete => shiftForDelete(fc.value, index, false),
        };
    }

    try out.appendSlice(allocator, src[a.open_start..a.after_open]);
    try out.appendSlice(allocator, src[body_lo..from_block.open_start]);
    try emitCornerBlock(allocator, out, src, from_block, from_corner, new_from);
    try out.appendSlice(allocator, src[from_block.after_close..a.after_close]);
    cursor.* = a.after_close;
}

/// Emit `<xdr:NAME>…</xdr:NAME>` (open + body + close) with the
/// axis sub-element's text spliced to `new_value`. If the corner
/// sub-element couldn't be parsed (corner is null) the block
/// emits verbatim.
fn emitCornerBlock(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    block: BlockBounds,
    corner: ?CornerCoord,
    new_value: ?u32,
) Error!void {
    // Open tag: `src[open_start .. lo]` is `<xdr:NAME>`.
    try out.appendSlice(allocator, src[block.open_start..block.lo]);

    if (corner == null or new_value == null or corner.?.value == new_value.?) {
        // No rewrite needed: pass through the body verbatim.
        try out.appendSlice(allocator, src[block.lo..block.hi]);
    } else {
        const c = corner.?;
        // Emit body up to the axis sub-element's text.
        try out.appendSlice(allocator, src[block.lo..c.span.text_start]);
        // Splice the new value as a decimal integer. u32 max is
        // 10 digits — bufPrint into a 16-byte buffer cannot exhaust.
        var num_buf: [16]u8 = undefined;
        const new_text = std.fmt.bufPrint(&num_buf, "{d}", .{new_value.?}) catch unreachable;
        try out.appendSlice(allocator, new_text);
        // Emit the rest of the body.
        try out.appendSlice(allocator, src[c.span.text_end..block.hi]);
    }

    // Close tag: `src[hi .. after_close]` is `</xdr:NAME>` —
    // `findCornerBlock` located it, so the bounds are its own.
    try out.appendSlice(allocator, src[block.hi..block.after_close]);
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
