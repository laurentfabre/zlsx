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
//! v1 LIMITATION: hard-codes the `xdr:` namespace prefix. Every
//! Microsoft Excel / LibreOffice / xlsxwriter / openpyxl /
//! python-calamine fixture in the project's corpus uses this
//! prefix; non-Microsoft producers with non-canonical prefixes
//! will surface zero rewrites (and therefore silent corruption
//! when row/col edits cross drawing anchors). pkg/drawings.zig
//! grew namespace-aware support post-C2a; the same machinery can
//! be lifted here in a follow-up iter when needed.

const std = @import("std");

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
/// allocated buffer.
pub fn applyEditToDrawing(
    allocator: Allocator,
    src: []const u8,
    axis: Axis,
    index: u32,
    kind: EditKind,
) Error![]u8 {
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

        if (matchAnchorOpenAt(src, i, "twoCellAnchor")) |a| {
            try processTwoCellAnchor(allocator, &out, src, a, axis, index, kind, &i);
        } else if (matchAnchorOpenAt(src, i, "oneCellAnchor")) |a| {
            try processOneCellAnchor(allocator, &out, src, a, axis, index, kind, &i);
        } else {
            try out.append(allocator, '<');
            i += 1;
        }
    }
    return try out.toOwnedSlice(allocator);
}

const xdr_prefix = "xdr:";

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
};

/// Match `<xdr:NAME` at `i` and locate the matching closing tag.
/// `name` is the bare element name (e.g. `twoCellAnchor`).
fn matchAnchorOpenAt(src: []const u8, i: usize, name: []const u8) ?AnchorMatch {
    if (i >= src.len or src[i] != '<') return null;
    const open_token_len = 1 + xdr_prefix.len + name.len;
    if (i + open_token_len > src.len) return null;
    if (!std.mem.eql(u8, src[i + 1 .. i + 1 + xdr_prefix.len], xdr_prefix)) return null;
    if (!std.mem.eql(u8, src[i + 1 + xdr_prefix.len .. i + open_token_len], name)) return null;
    const after_name = i + open_token_len;
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
        };
    }

    // Locate `</xdr:NAME>` close.
    var close_buf: [64]u8 = undefined;
    if (xdr_prefix.len + name.len + 3 > close_buf.len) return null;
    var n: usize = 0;
    close_buf[n] = '<';
    n += 1;
    close_buf[n] = '/';
    n += 1;
    @memcpy(close_buf[n .. n + xdr_prefix.len], xdr_prefix);
    n += xdr_prefix.len;
    @memcpy(close_buf[n .. n + name.len], name);
    n += name.len;
    close_buf[n] = '>';
    n += 1;
    const close_needle = close_buf[0..n];

    const close = std.mem.indexOfPos(u8, src, after_open, close_needle) orelse return null;
    return .{
        .open_start = i,
        .after_open = after_open,
        .close_start = close,
        .after_close = close + close_needle.len,
        .self_closing = false,
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

fn findInnerInt(src: []const u8, lo: usize, hi: usize, name: []const u8) ?InnerSpan {
    var open_buf: [64]u8 = undefined;
    var close_buf: [64]u8 = undefined;
    if (1 + xdr_prefix.len + name.len + 1 > open_buf.len) return null;
    if (2 + xdr_prefix.len + name.len + 1 > close_buf.len) return null;

    var n: usize = 0;
    open_buf[n] = '<';
    n += 1;
    @memcpy(open_buf[n .. n + xdr_prefix.len], xdr_prefix);
    n += xdr_prefix.len;
    @memcpy(open_buf[n .. n + name.len], name);
    n += name.len;
    open_buf[n] = '>';
    n += 1;
    const open_needle = open_buf[0..n];

    var m: usize = 0;
    close_buf[m] = '<';
    m += 1;
    close_buf[m] = '/';
    m += 1;
    @memcpy(close_buf[m .. m + xdr_prefix.len], xdr_prefix);
    m += xdr_prefix.len;
    @memcpy(close_buf[m .. m + name.len], name);
    m += name.len;
    close_buf[m] = '>';
    m += 1;
    const close_needle = close_buf[0..m];

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
fn readCorner(src: []const u8, corner_lo: usize, corner_hi: usize, axis: Axis) ?CornerCoord {
    const span = findInnerInt(src, corner_lo, corner_hi, axisInnerName(axis)) orelse return null;
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
};

fn findCornerBlock(src: []const u8, lo: usize, hi: usize, name: []const u8) ?BlockBounds {
    var open_buf: [16]u8 = undefined;
    var close_buf: [16]u8 = undefined;
    var n: usize = 0;
    open_buf[n] = '<';
    n += 1;
    @memcpy(open_buf[n .. n + xdr_prefix.len], xdr_prefix);
    n += xdr_prefix.len;
    @memcpy(open_buf[n .. n + name.len], name);
    n += name.len;
    open_buf[n] = '>';
    n += 1;
    const open_needle = open_buf[0..n];

    var m: usize = 0;
    close_buf[m] = '<';
    m += 1;
    close_buf[m] = '/';
    m += 1;
    @memcpy(close_buf[m .. m + xdr_prefix.len], xdr_prefix);
    m += xdr_prefix.len;
    @memcpy(close_buf[m .. m + name.len], name);
    m += name.len;
    close_buf[m] = '>';
    m += 1;
    const close_needle = close_buf[0..m];

    const open_pos = std.mem.indexOfPos(u8, src, lo, open_needle) orelse return null;
    if (open_pos >= hi) return null;
    const inner_lo = open_pos + open_needle.len;
    const close_pos = std.mem.indexOfPos(u8, src, inner_lo, close_needle) orelse return null;
    if (close_pos >= hi) return null;
    return .{ .open_start = open_pos, .lo = inner_lo, .hi = close_pos };
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
    const from_block = findCornerBlock(src, body_lo, body_hi, "from") orelse {
        try out.appendSlice(allocator, src[a.open_start..a.after_close]);
        cursor.* = a.after_close;
        return;
    };
    const to_block = findCornerBlock(src, body_lo, body_hi, "to") orelse {
        try out.appendSlice(allocator, src[a.open_start..a.after_close]);
        cursor.* = a.after_close;
        return;
    };

    const from_corner = readCorner(src, from_block.lo, from_block.hi, axis);
    const to_corner = readCorner(src, to_block.lo, to_block.hi, axis);

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
    try emitCornerBlock(allocator, out, src, from_block, "from", from_corner, new_from);
    try out.appendSlice(allocator, src[closeAfter(from_block, "from")..to_block.open_start]);
    try emitCornerBlock(allocator, out, src, to_block, "to", to_corner, new_to);
    try out.appendSlice(allocator, src[closeAfter(to_block, "to")..a.after_close]);
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

    const from_block = findCornerBlock(src, body_lo, body_hi, "from") orelse {
        try out.appendSlice(allocator, src[a.open_start..a.after_close]);
        cursor.* = a.after_close;
        return;
    };
    const from_corner = readCorner(src, from_block.lo, from_block.hi, axis);

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
    try emitCornerBlock(allocator, out, src, from_block, "from", from_corner, new_from);
    try out.appendSlice(allocator, src[closeAfter(from_block, "from")..a.after_close]);
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
    name: []const u8,
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

    // Close tag: `src[hi .. closeAfter(block, name)]` is
    // `</xdr:NAME>`.
    try out.appendSlice(allocator, src[block.hi..closeAfter(block, name)]);
}

/// Byte offset just past the `>` of the corner block's closing
/// `</xdr:NAME>` tag. `findCornerBlock` already validated that
/// the closing tag is well-formed in the input, so we compute
/// the offset arithmetically rather than rescanning.
fn closeAfter(block: BlockBounds, name: []const u8) usize {
    return block.hi + "</".len + xdr_prefix.len + name.len + 1; // 1 for '>'
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
