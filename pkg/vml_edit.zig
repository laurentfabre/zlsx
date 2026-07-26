//! VML (Vector Markup Language) anchor rewriter for OOXML
//! `xl/drawings/vmlDrawingN.vml` parts. Pure-function; consumed
//! by `pkg/editor.zig`'s row/col edit path after the worksheet's
//! own bytes have been rewritten by `pkg/sheet_edit.zig` and the
//! modern xdr drawing part has been rewritten by
//! `pkg/drawing_edit.zig`. Companion to (not extension of) those
//! modules because the legacy VML format lives in a different
//! file with a different schema.
//!
//! VML is the legacy spreadsheet-drawing format, used by Excel
//! today only for cell comments / notes. Each `<v:shape>` block
//! contains one comment shape with three coordinate attachment
//! points the rewriter must keep in sync:
//!
//!   <v:shape ...>
//!     <x:ClientData ObjectType="Note">
//!       <x:Anchor>FC, FCO, FR, FRO, TC, TCO, TR, TRO</x:Anchor>
//!       <x:Row>R</x:Row>      ← anchor cell row    (0-based)
//!       <x:Column>C</x:Column> ← anchor cell col   (0-based)
//!     </x:ClientData>
//!   </v:shape>
//!
//! The 8-int `<x:Anchor>` payload encodes the comment's display
//! rectangle (FC..TC × FR..TR with pixel offsets); `<x:Row>` /
//! `<x:Column>` identify the cell the comment is attached to.
//!
//! Drop semantics: when the anchor cell is deleted (delete-axis
//! match against `<x:Row>` for row edits, `<x:Column>` for col
//! edits), the entire `<v:shape>` block is removed. The
//! `<o:idmap>` chunks at the top of the file are over-provisioned
//! by 1024 IDs each, so a missing shape ID is harmless (per the
//! emitter's documented invariant).
//!
//! On a non-drop edit the rewriter shifts both the anchor-cell
//! coordinate (`<x:Row>` or `<x:Column>`) AND the corresponding
//! pair from the `<x:Anchor>` 8-tuple (FC/TC for col edits,
//! FR/TR for row edits). The display rectangle's "from" corner
//! follows the data: TL stays put on a delete-match (the next
//! cell's data shifts into the slot); BR shrinks. Insert at any
//! index ≥ the corner shifts that corner.
//!
//! v1 LIMITATION: hard-codes the `v:` and `x:` namespace prefixes.
//! Every Microsoft Excel + LibreOffice + xlsxwriter + openpyxl
//! corpus fixture uses these prefixes; non-Microsoft producers
//! with non-canonical prefixes will surface zero rewrites (and
//! therefore silent corruption when row/col edits cross VML
//! anchors). Namespace-aware support can land in a follow-up
//! iter when needed.

const std = @import("std");

const Allocator = std.mem.Allocator;

pub const Error = error{
    MalformedVmlDrawing,
    VmlCoordinateOverflow,
    MalformedCommentsXml,
} || Allocator.Error;

pub const EditKind = enum { insert, delete };

pub const Axis = enum { row, col };

/// Apply one row OR column edit to a VML drawing-part body.
/// `index` is 0-based to match the wire format of `<x:Row>` /
/// `<x:Column>` and the comma-separated `<x:Anchor>` ints. The
/// Editor's row/col-edit surfaces use 1-based indices and must
/// subtract 1 before calling. Returns a freshly allocated buffer.
pub fn applyEditToVmlDrawing(
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

        if (matchVShapeAt(src, i)) |s| {
            try processVShape(allocator, &out, src, s, axis, index, kind, &i);
        } else {
            try out.append(allocator, '<');
            i += 1;
        }
    }
    return try out.toOwnedSlice(allocator);
}

const ShapeMatch = struct {
    open_start: usize,
    after_open: usize,
    close_start: usize,
    after_close: usize,
    self_closing: bool,
};

fn matchVShapeAt(src: []const u8, i: usize) ?ShapeMatch {
    const open_token = "<v:shape";
    if (i + open_token.len > src.len) return null;
    if (!std.mem.eql(u8, src[i .. i + open_token.len], open_token)) return null;
    const after_name = i + open_token.len;
    if (after_name >= src.len) return null;
    const c = src[after_name];
    // Disambiguate from `<v:shapetype>`, `<v:shadow>`, etc.
    if (c != ' ' and c != '\t' and c != '\n' and c != '\r' and c != '/' and c != '>') return null;
    // Find unquoted `>` ending the open tag. XML attribute values
    // can legitimately contain `>` inside quoted strings (e.g.
    // `style="...; foo:>bar"`); a naive `indexOfScalarPos` would
    // truncate the open tag too early. Codex review REL-704.
    const gt = findUnquotedGt(src, after_name) orelse return null;
    const after_open = gt + 1;

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

    const close_needle = "</v:shape>";
    const close = std.mem.indexOfPos(u8, src, after_open, close_needle) orelse return null;
    return .{
        .open_start = i,
        .after_open = after_open,
        .close_start = close,
        .after_close = close + close_needle.len,
        .self_closing = false,
    };
}

/// Find the byte offset of the next unquoted `>` starting at
/// `from`. Tracks single + double quotes so quoted attribute
/// values containing `>` don't truncate the open-tag scan.
/// Returns null if no unquoted `>` exists.
fn findUnquotedGt(src: []const u8, from: usize) ?usize {
    var j: usize = from;
    var quote: u8 = 0; // 0 = not inside quotes; '"' or '\'' otherwise
    while (j < src.len) : (j += 1) {
        const ch = src[j];
        if (quote == 0) {
            if (ch == '"' or ch == '\'') {
                quote = ch;
            } else if (ch == '>') {
                return j;
            }
        } else if (ch == quote) {
            quote = 0;
        }
    }
    return null;
}

const InnerSpan = struct {
    text_start: usize,
    text_end: usize,
};

/// Find `<x:NAME>...</x:NAME>` strictly inside `[lo, hi)` and
/// return the text-content span. Tolerates whitespace inside the
/// element text (callers trim before parsing).
fn findXmlIntElement(src: []const u8, lo: usize, hi: usize, name: []const u8) ?InnerSpan {
    var open_buf: [32]u8 = undefined;
    var close_buf: [32]u8 = undefined;

    var n: usize = 0;
    open_buf[n] = '<';
    n += 1;
    open_buf[n] = 'x';
    n += 1;
    open_buf[n] = ':';
    n += 1;
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
    close_buf[m] = 'x';
    m += 1;
    close_buf[m] = ':';
    m += 1;
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
    return .{ .text_start = text_start, .text_end = close_pos };
}

fn parseUInt(text: []const u8) ?u32 {
    const trimmed = std.mem.trim(u8, text, " \t\n\r");
    if (trimmed.len == 0) return null;
    var s = trimmed;
    if (s[0] == '+') s = s[1..];
    if (s.len == 0) return null;
    return std.fmt.parseInt(u32, s, 10) catch null;
}

fn shiftForInsert(value: u32, edit_index: u32) Error!u32 {
    if (value < edit_index) return value;
    if (value == std.math.maxInt(u32)) return Error.VmlCoordinateOverflow;
    return value + 1;
}

fn shiftForDelete(value: u32, edit_index: u32, is_br_corner: bool) u32 {
    if (value > edit_index) return value - 1;
    if (value == edit_index and is_br_corner and edit_index > 0) return value - 1;
    return value;
}

fn processVShape(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    s: ShapeMatch,
    axis: Axis,
    index: u32,
    kind: EditKind,
    cursor: *usize,
) Error!void {
    if (s.self_closing) {
        try out.appendSlice(allocator, src[s.open_start..s.after_open]);
        cursor.* = s.after_open;
        return;
    }

    const body_lo = s.after_open;
    const body_hi = s.close_start;

    // The anchor cell is identified by `<x:Row>` (rows) and
    // `<x:Column>` (cols). Shapes without these (e.g. floating
    // form controls) pass through verbatim.
    const cell_axis_name: []const u8 = switch (axis) {
        .row => "Row",
        .col => "Column",
    };
    const cell_span = findXmlIntElement(src, body_lo, body_hi, cell_axis_name);

    // Shapes without `<x:Row>` / `<x:Column>` (form controls,
    // standalone anchors) are not the comment-attachment shape
    // this rewriter targets — pass them through verbatim. Codex
    // review REL-703.
    if (cell_span == null) {
        try out.appendSlice(allocator, src[s.open_start..s.after_close]);
        cursor.* = s.after_close;
        return;
    }

    // Refuse on malformed present `<x:Row>` / `<x:Column>` —
    // silently shifting the anchor tuple while leaving the
    // anchor-cell coord stale / malformed is exactly the
    // silent-corruption case the refusal exists to prevent.
    // Codex review REL-706.
    const cell_value: u32 = parseUInt(src[cell_span.?.text_start..cell_span.?.text_end]) orelse
        return Error.MalformedVmlDrawing;

    // Drop on delete-match against the anchor cell.
    if (kind == .delete and cell_value == index) {
        cursor.* = s.after_close;
        return;
    }

    // Compute new anchor-cell value if it shifts.
    const new_cell_value: u32 = switch (kind) {
        .insert => try shiftForInsert(cell_value, index),
        .delete => shiftForDelete(cell_value, index, false),
    };

    // Locate `<x:Anchor>` and parse the 8-int payload. Compute
    // shifted from/to corners on the edited axis. A present-but-
    // malformed `<x:Anchor>` (wrong int count, non-integer token)
    // is a hard refusal — emitting a shifted anchor cell with a
    // stale display rect is the silent-corruption case the
    // refusal axis exists to prevent. Codex review REL-702.
    const anchor_span = findXmlIntElement(src, body_lo, body_hi, "Anchor");
    var anchor_pair: ?VmlAnchorPair = null;
    if (anchor_span) |span| {
        anchor_pair = parseAnchorAxisPair(src[span.text_start..span.text_end], axis) orelse
            return Error.MalformedVmlDrawing;
    }

    var new_from: ?u32 = null;
    var new_to: ?u32 = null;
    if (anchor_pair) |pair| {
        new_from = switch (kind) {
            .insert => try shiftForInsert(pair.from, index),
            .delete => shiftForDelete(pair.from, index, false),
        };
        new_to = switch (kind) {
            .insert => try shiftForInsert(pair.to, index),
            .delete => shiftForDelete(pair.to, index, true),
        };
        // Maintain non-inverted invariant. shiftForDelete on TL
        // can leave from > to in pathological cases (BR shrunk
        // to from-1); clamp to from.
        if (new_from.? > new_to.?) new_to = new_from;
    }

    // Emit shape open + walk body splicing the changed elements.
    try out.appendSlice(allocator, src[s.open_start..s.after_open]);

    var cur: usize = body_lo;

    // Process anchor + cell-axis splices in document order. The
    // emitter places `<x:Anchor>` BEFORE `<x:Row>`/`<x:Column>`,
    // but third-party files might differ — we sort by text_start
    // ascending to handle either order safely.
    var splices: [2]Splice = undefined;
    var splice_count: usize = 0;
    if (anchor_span != null and anchor_pair != null and new_from != null and new_to != null and
        (anchor_pair.?.from != new_from.? or anchor_pair.?.to != new_to.?))
    {
        splices[splice_count] = .{
            .span = anchor_span.?,
            .new_text_kind = .{ .anchor = .{
                .axis = axis,
                .from = new_from.?,
                .to = new_to.?,
            } },
        };
        splice_count += 1;
    }
    if (cell_value != new_cell_value) {
        splices[splice_count] = .{
            .span = cell_span.?,
            .new_text_kind = .{ .integer = new_cell_value },
        };
        splice_count += 1;
    }
    // Sort by text_start (insertion sort over up to 2 entries).
    if (splice_count == 2 and splices[0].span.text_start > splices[1].span.text_start) {
        std.mem.swap(Splice, &splices[0], &splices[1]);
    }

    for (splices[0..splice_count]) |sp| {
        try out.appendSlice(allocator, src[cur..sp.span.text_start]);
        switch (sp.new_text_kind) {
            .integer => |v| {
                var num_buf: [16]u8 = undefined;
                const txt = std.fmt.bufPrint(&num_buf, "{d}", .{v}) catch unreachable;
                try out.appendSlice(allocator, txt);
            },
            .anchor => |a| {
                try writeRewrittenAnchor(allocator, out, src[sp.span.text_start..sp.span.text_end], a);
            },
        }
        cur = sp.span.text_end;
    }
    try out.appendSlice(allocator, src[cur..s.after_close]);
    cursor.* = s.after_close;
}

const VmlAnchorPair = struct { from: u32, to: u32 };

const NewAnchor = struct {
    axis: Axis,
    from: u32,
    to: u32,
};

const Splice = struct {
    span: InnerSpan,
    new_text_kind: union(enum) {
        integer: u32,
        anchor: NewAnchor,
    },
};

/// Parse the from/to pair on the requested axis from the 8-int
/// `<x:Anchor>` payload. Returns null if the payload doesn't
/// have 8 comma-separated integers.
fn parseAnchorAxisPair(text: []const u8, axis: Axis) ?VmlAnchorPair {
    var parts: [8]u32 = undefined;
    var idx: usize = 0;
    var it = std.mem.splitScalar(u8, text, ',');
    while (it.next()) |raw| {
        if (idx >= 8) return null;
        parts[idx] = parseUInt(raw) orelse return null;
        idx += 1;
    }
    if (idx != 8) return null;
    return switch (axis) {
        .col => .{ .from = parts[0], .to = parts[4] },
        .row => .{ .from = parts[2], .to = parts[6] },
    };
}

/// Re-emit the 8-int `<x:Anchor>` payload with the from/to pair
/// on the requested axis replaced. Preserves the off-axis ints
/// + the original spacing convention (single space after each
/// comma, matching zlsx's writer and Office's emit).
fn writeRewrittenAnchor(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    original_text: []const u8,
    new_anchor: NewAnchor,
) Error!void {
    var parts: [8]u32 = undefined;
    var idx: usize = 0;
    var it = std.mem.splitScalar(u8, original_text, ',');
    while (it.next()) |raw| {
        if (idx >= 8) return Error.MalformedVmlDrawing;
        parts[idx] = parseUInt(raw) orelse return Error.MalformedVmlDrawing;
        idx += 1;
    }
    if (idx != 8) return Error.MalformedVmlDrawing;

    switch (new_anchor.axis) {
        .col => {
            parts[0] = new_anchor.from;
            parts[4] = new_anchor.to;
        },
        .row => {
            parts[2] = new_anchor.from;
            parts[6] = new_anchor.to;
        },
    }

    var buf: [128]u8 = undefined;
    const txt = std.fmt.bufPrint(
        &buf,
        "{d}, {d}, {d}, {d}, {d}, {d}, {d}, {d}",
        .{ parts[0], parts[1], parts[2], parts[3], parts[4], parts[5], parts[6], parts[7] },
    ) catch unreachable;
    try out.appendSlice(allocator, txt);
}

// ---------------------------------------------------------------------------
// Comments part rewriter (xl/commentsN.xml).
// ---------------------------------------------------------------------------
//
// VML comment shapes and their text content live in two parts: the VML
// drawing (handled above) carries the visual anchor, while
// `xl/commentsN.xml` carries the comment text + its anchor cell ref:
//
//   <comments xmlns="…spreadsheetml/2006/main">
//     <authors><author>…</author></authors>
//     <commentList>
//       <comment ref="C5" authorId="0" shapeId="0">
//         <text>…</text>
//       </comment>
//     </commentList>
//   </comments>
//
// The two parts MUST stay synchronized: a row/col edit that drops a
// VML shape MUST also drop the matching `<comment>` block (the VML
// shape's anchor cell == the comment's `ref`).
//
// On non-drop edits, `<comment ref="…">` shifts as a single A1 cell
// ref — same shift logic as everywhere else in the byte transform.

/// Apply one row OR column edit to a `xl/commentsN.xml` body.
/// `index` is 1-based to match the A1-ref `<comment ref>` encoding
/// (this is the OPPOSITE of the VML rewriter, which uses 0-based
/// `<x:Row>` / `<x:Column>` integers).
pub fn applyEditToCommentsXml(
    allocator: Allocator,
    src: []const u8,
    axis: Axis,
    index_1based: u32,
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

        if (matchCommentOpenAt(src, i)) |c| {
            try processCommentTag(allocator, &out, src, c, axis, index_1based, kind, &i);
        } else {
            try out.append(allocator, '<');
            i += 1;
        }
    }
    return try out.toOwnedSlice(allocator);
}

const CommentMatch = struct {
    open_start: usize,
    after_open: usize,
    /// Byte offset of `<` of the closing `</comment>` tag.
    close_start: usize,
    /// Byte offset just past the `>` of the closing tag.
    after_close: usize,
};

fn matchCommentOpenAt(src: []const u8, i: usize) ?CommentMatch {
    const open_token = "<comment";
    if (i + open_token.len > src.len) return null;
    if (!std.mem.eql(u8, src[i .. i + open_token.len], open_token)) return null;
    const after_name = i + open_token.len;
    if (after_name >= src.len) return null;
    const ch = src[after_name];
    // Disambiguate from `<commentList>` etc.
    if (ch != ' ' and ch != '\t' and ch != '\n' and ch != '\r' and ch != '/' and ch != '>') return null;
    const gt = findUnquotedGt(src, after_name) orelse return null;
    const after_open = gt + 1;

    // Self-closing form is unusual for `<comment>` (it always has
    // a `<text>` child) but handle it defensively.
    var trim_end = gt;
    while (trim_end > after_name) : (trim_end -= 1) {
        const c = src[trim_end - 1];
        if (c == ' ' or c == '\t' or c == '\n' or c == '\r') continue;
        break;
    }
    if (trim_end > after_name and src[trim_end - 1] == '/') {
        return .{
            .open_start = i,
            .after_open = after_open,
            .close_start = after_open,
            .after_close = after_open,
        };
    }

    const close_needle = "</comment>";
    const close = std.mem.indexOfPos(u8, src, after_open, close_needle) orelse return null;
    return .{
        .open_start = i,
        .after_open = after_open,
        .close_start = close,
        .after_close = close + close_needle.len,
    };
}

fn processCommentTag(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    m: CommentMatch,
    axis: Axis,
    index_1based: u32,
    kind: EditKind,
    cursor: *usize,
) Error!void {
    // Read the `ref` attribute from the open tag's attribute span.
    const open_attrs = src[m.open_start + "<comment".len .. m.after_open - 1];
    const ref = getCommentRef(open_attrs) orelse {
        // No ref attribute — pass through verbatim.
        try out.appendSlice(allocator, src[m.open_start..m.after_close]);
        cursor.* = m.after_close;
        return;
    };

    const a1 = parseA1Single(ref) orelse return Error.MalformedCommentsXml;

    // Drop on delete-match against the relevant axis.
    if (kind == .delete) switch (axis) {
        .row => if (a1.row == index_1based) {
            cursor.* = m.after_close;
            return;
        },
        .col => if (a1.col == index_1based) {
            cursor.* = m.after_close;
            return;
        },
    };

    // Compute the shifted ref. `shiftForInsert`/`shiftForDelete`
    // are the same 1-based-aware shifts used by the VML half of
    // this module — they're index-base-agnostic.
    var new_a1 = a1;
    switch (axis) {
        .row => new_a1.row = switch (kind) {
            .insert => try shiftForInsert(a1.row, index_1based),
            .delete => shiftForDelete(a1.row, index_1based, false),
        },
        .col => new_a1.col = switch (kind) {
            .insert => try shiftForInsert(a1.col, index_1based),
            .delete => shiftForDelete(a1.col, index_1based, false),
        },
    }

    if (new_a1.row == a1.row and new_a1.col == a1.col) {
        // No-op rewrite — pass through verbatim.
        try out.appendSlice(allocator, src[m.open_start..m.after_close]);
        cursor.* = m.after_close;
        return;
    }

    // Rewrite the open tag's `ref` attribute with the new A1 ref.
    var new_ref_buf: [16]u8 = undefined;
    const new_ref = formatA1(&new_ref_buf, new_a1) catch return Error.MalformedCommentsXml;
    try writeCommentOpenWithNewRef(allocator, out, src, m, ref, new_ref);
    // Body + close tag emit verbatim.
    try out.appendSlice(allocator, src[m.after_open..m.after_close]);
    cursor.* = m.after_close;
}

const A1Ref = struct { col: u32, row: u32 };

fn parseA1Single(ref: []const u8) ?A1Ref {
    if (ref.len == 0) return null;
    var i: usize = 0;
    var col: u32 = 0;
    while (i < ref.len and ref[i] >= 'A' and ref[i] <= 'Z') : (i += 1) {
        col = col * 26 + (ref[i] - 'A' + 1);
        if (col > 16384) return null;
    }
    if (i == 0 or i == ref.len) return null;
    var row: u32 = 0;
    while (i < ref.len) : (i += 1) {
        const ch = ref[i];
        if (ch < '0' or ch > '9') return null;
        row = row * 10 + (ch - '0');
        if (row > 1048576) return null;
    }
    if (row == 0) return null;
    return .{ .col = col, .row = row };
}

fn formatA1(buf: []u8, ref: A1Ref) ![]const u8 {
    if (ref.col == 0 or ref.col > 16384 or ref.row == 0 or ref.row > 1048576) return error.MalformedCommentsXml;
    var letters: [4]u8 = undefined;
    var n_letters: usize = 0;
    var c: u32 = ref.col;
    while (c > 0) {
        const digit = (c - 1) % 26;
        letters[n_letters] = 'A' + @as(u8, @intCast(digit));
        n_letters += 1;
        c = (c - 1) / 26;
    }
    std.mem.reverse(u8, letters[0..n_letters]);
    return std.fmt.bufPrint(buf, "{s}{d}", .{ letters[0..n_letters], ref.row }) catch unreachable;
}

/// Read the `ref="..."` attribute value from a `<comment>` open
/// tag's attribute span.
fn getCommentRef(attrs: []const u8) ?[]const u8 {
    // Walk attributes looking for `ref` followed by optional
    // whitespace, `=`, optional whitespace, then a quote.
    var j: usize = 0;
    while (j < attrs.len) {
        // Find a name boundary: start-of-string or whitespace.
        if (j > 0 and attrs[j - 1] != ' ' and attrs[j - 1] != '\t' and
            attrs[j - 1] != '\n' and attrs[j - 1] != '\r')
        {
            j += 1;
            continue;
        }
        if (j + 3 > attrs.len) return null;
        if (!std.mem.eql(u8, attrs[j .. j + 3], "ref")) {
            j += 1;
            continue;
        }
        // Confirm the byte after `ref` is `=` or whitespace.
        var k = j + 3;
        while (k < attrs.len and (attrs[k] == ' ' or attrs[k] == '\t' or attrs[k] == '\n' or attrs[k] == '\r')) k += 1;
        if (k >= attrs.len or attrs[k] != '=') {
            j += 1;
            continue;
        }
        k += 1;
        while (k < attrs.len and (attrs[k] == ' ' or attrs[k] == '\t' or attrs[k] == '\n' or attrs[k] == '\r')) k += 1;
        if (k >= attrs.len) return null;
        const quote = attrs[k];
        if (quote != '"' and quote != '\'') {
            j += 1;
            continue;
        }
        const start = k + 1;
        const end = std.mem.indexOfScalarPos(u8, attrs, start, quote) orelse return null;
        return attrs[start..end];
    }
    return null;
}

fn writeCommentOpenWithNewRef(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    m: CommentMatch,
    old_ref: []const u8,
    new_ref: []const u8,
) Error!void {
    // The old_ref is a slice into the original attrs span; locate
    // its absolute byte offset within `src` so we can splice.
    const old_ref_start = @intFromPtr(old_ref.ptr) - @intFromPtr(src.ptr);
    const old_ref_end = old_ref_start + old_ref.len;
    try out.appendSlice(allocator, src[m.open_start..old_ref_start]);
    try out.appendSlice(allocator, new_ref);
    try out.appendSlice(allocator, src[old_ref_end..m.after_open]);
}

// ---------------------------------------------------------------------------
// Tests — pure-function coverage.
// ---------------------------------------------------------------------------

const testing = std.testing;

fn wrapVml(allocator: Allocator, body: []const u8) ![]u8 {
    const head = "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">" ++
        "<o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/></o:shapelayout>" ++
        "<v:shapetype id=\"_x0000_t202\" path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/></v:shapetype>";
    const tail = "</xml>";
    var buf = try allocator.alloc(u8, head.len + body.len + tail.len);
    @memcpy(buf[0..head.len], head);
    @memcpy(buf[head.len .. head.len + body.len], body);
    @memcpy(buf[head.len + body.len ..], tail);
    return buf;
}

const sample_shape =
    "<v:shape id=\"_x0000_s1025\" type=\"#_x0000_t202\" style=\"position:absolute\">" ++
    "<v:fill color2=\"#ffffe1\"/>" ++
    "<v:textbox><div/></v:textbox>" ++
    "<x:ClientData ObjectType=\"Note\">" ++
    "<x:MoveWithCells/><x:SizeWithCells/>" ++
    "<x:Anchor>2, 15, 4, 2, 5, 31, 8, 3</x:Anchor>" ++
    "<x:AutoFill>False</x:AutoFill>" ++
    "<x:Row>4</x:Row><x:Column>3</x:Column>" ++
    "</x:ClientData>" ++
    "</v:shape>";

test "VML: insert col before anchor shifts column + anchor's from/to col" {
    const a = testing.allocator;
    const src = try wrapVml(a, sample_shape);
    defer a.free(src);
    // Insert at col 0. Anchor cell <x:Column>3</x:Column> shifts to 4.
    // <x:Anchor> from-col 2 → 3, to-col 5 → 6.
    const out = try applyEditToVmlDrawing(a, src, .col, 0, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<x:Column>4</x:Column>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<x:Anchor>3, 15, 4, 2, 6, 31, 8, 3</x:Anchor>") != null);
    // Row coords untouched.
    try testing.expect(std.mem.indexOf(u8, out, "<x:Row>4</x:Row>") != null);
}

test "VML: insert row before anchor shifts row + anchor's from/to row" {
    const a = testing.allocator;
    const src = try wrapVml(a, sample_shape);
    defer a.free(src);
    // Insert at row 0. <x:Row>4</x:Row> → 5. From-row 4 → 5, to-row 8 → 9.
    const out = try applyEditToVmlDrawing(a, src, .row, 0, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<x:Row>5</x:Row>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<x:Anchor>2, 15, 5, 2, 5, 31, 9, 3</x:Anchor>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<x:Column>3</x:Column>") != null);
}

test "VML: delete at anchor's column drops the entire shape" {
    const a = testing.allocator;
    const src = try wrapVml(a, sample_shape);
    defer a.free(src);
    // Delete at col 3 (== <x:Column>3</x:Column>). Drop.
    const out = try applyEditToVmlDrawing(a, src, .col, 3, .delete);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<v:shape ") == null);
    try testing.expect(std.mem.indexOf(u8, out, "<x:Anchor>") == null);
    try testing.expect(std.mem.indexOf(u8, out, "<x:Column>") == null);
}

test "VML: delete at anchor's row drops the entire shape" {
    const a = testing.allocator;
    const src = try wrapVml(a, sample_shape);
    defer a.free(src);
    const out = try applyEditToVmlDrawing(a, src, .row, 4, .delete);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<v:shape ") == null);
}

test "VML: delete past anchor shifts column + anchor's to" {
    const a = testing.allocator;
    const src = try wrapVml(a, sample_shape);
    defer a.free(src);
    // Delete at col 4 — between anchor's from (2) and to (5).
    // <x:Column>3</x:Column>: 3 < 4 → unchanged.
    // From-col 2 < 4: unchanged. To-col 5 > 4: 4.
    const out = try applyEditToVmlDrawing(a, src, .col, 4, .delete);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<x:Column>3</x:Column>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<x:Anchor>2, 15, 4, 2, 4, 31, 8, 3</x:Anchor>") != null);
}

test "VML: delete past everything is a no-op" {
    const a = testing.allocator;
    const src = try wrapVml(a, sample_shape);
    defer a.free(src);
    const out = try applyEditToVmlDrawing(a, src, .col, 100, .delete);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<x:Column>3</x:Column>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<x:Anchor>2, 15, 4, 2, 5, 31, 8, 3</x:Anchor>") != null);
}

test "VML: shapetype + shapelayout pass through unchanged" {
    const a = testing.allocator;
    const src = try wrapVml(a, sample_shape);
    defer a.free(src);
    const out = try applyEditToVmlDrawing(a, src, .col, 0, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<v:shapetype ") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<o:shapelayout") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<o:idmap v:ext=\"edit\" data=\"1\"/>") != null);
}

test "VML: two shapes both rewrite independently" {
    const a = testing.allocator;
    const body = sample_shape ++ sample_shape;
    const src = try wrapVml(a, body);
    defer a.free(src);
    const out = try applyEditToVmlDrawing(a, src, .col, 0, .insert);
    defer a.free(out);
    var count: usize = 0;
    var i: usize = 0;
    while (std.mem.indexOfPos(u8, out, i, "<x:Column>4</x:Column>")) |idx| {
        count += 1;
        i = idx + 1;
    }
    try testing.expectEqual(@as(usize, 2), count);
}

test "VML: deleting one of two shapes leaves the other intact" {
    const a = testing.allocator;
    const body = sample_shape ++
        "<v:shape id=\"_x0000_s1026\" type=\"#_x0000_t202\">" ++
        "<x:ClientData ObjectType=\"Note\">" ++
        "<x:Anchor>0, 0, 0, 0, 1, 0, 1, 0</x:Anchor>" ++
        "<x:Row>0</x:Row><x:Column>0</x:Column>" ++
        "</x:ClientData></v:shape>";
    const src = try wrapVml(a, body);
    defer a.free(src);
    // Delete col 3 (first shape's anchor) — first shape drops,
    // second's <x:Column>0</x:Column> < 3 stays put.
    const out = try applyEditToVmlDrawing(a, src, .col, 3, .delete);
    defer a.free(out);
    var shape_count: usize = 0;
    var i: usize = 0;
    while (std.mem.indexOfPos(u8, out, i, "<v:shape ")) |idx| {
        shape_count += 1;
        i = idx + 1;
    }
    try testing.expectEqual(@as(usize, 1), shape_count);
    try testing.expect(std.mem.indexOf(u8, out, "_x0000_s1026") != null);
    try testing.expect(std.mem.indexOf(u8, out, "_x0000_s1025") == null);
}

test "VML: shape without ClientData (form control) passes through verbatim" {
    const a = testing.allocator;
    const body = "<v:shape id=\"_x0000_s99\" type=\"#_x0000_t202\"><v:textbox><div/></v:textbox></v:shape>";
    const src = try wrapVml(a, body);
    defer a.free(src);
    const out = try applyEditToVmlDrawing(a, src, .col, 0, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "_x0000_s99") != null);
}

test "VML: malformed <x:Anchor> (7 ints) refuses with MalformedVmlDrawing (REL-702)" {
    const a = testing.allocator;
    const body =
        "<v:shape id=\"_x0000_s1025\" type=\"#_x0000_t202\">" ++
        "<x:ClientData ObjectType=\"Note\">" ++
        "<x:Anchor>2, 15, 4, 2, 5, 31, 8</x:Anchor>" ++ // only 7 ints
        "<x:Row>4</x:Row><x:Column>3</x:Column>" ++
        "</x:ClientData></v:shape>";
    const src = try wrapVml(a, body);
    defer a.free(src);
    const result = applyEditToVmlDrawing(a, src, .col, 0, .insert);
    try testing.expectError(error.MalformedVmlDrawing, result);
}

test "VML: malformed <x:Anchor> (non-integer token) refuses with MalformedVmlDrawing (REL-702)" {
    const a = testing.allocator;
    const body =
        "<v:shape id=\"_x0000_s1025\" type=\"#_x0000_t202\">" ++
        "<x:ClientData ObjectType=\"Note\">" ++
        "<x:Anchor>2, 15, 4, abc, 5, 31, 8, 3</x:Anchor>" ++
        "<x:Row>4</x:Row><x:Column>3</x:Column>" ++
        "</x:ClientData></v:shape>";
    const src = try wrapVml(a, body);
    defer a.free(src);
    const result = applyEditToVmlDrawing(a, src, .col, 0, .insert);
    try testing.expectError(error.MalformedVmlDrawing, result);
}

test "VML: anchor without x:Row/x:Column passes through unchanged (REL-703)" {
    const a = testing.allocator;
    // Shape has <x:ClientData> + <x:Anchor> but no <x:Row>/<x:Column>.
    // Should pass through verbatim — this isn't a comment-attachment shape.
    const body =
        "<v:shape id=\"_x0000_s1025\" type=\"#_x0000_t202\">" ++
        "<x:ClientData ObjectType=\"Drop\">" ++
        "<x:Anchor>2, 15, 4, 2, 5, 31, 8, 3</x:Anchor>" ++
        "</x:ClientData></v:shape>";
    const src = try wrapVml(a, body);
    defer a.free(src);
    const out = try applyEditToVmlDrawing(a, src, .col, 0, .insert);
    defer a.free(out);
    // Anchor's from-col (2) MUST stay 2 — no rewrite without an
    // anchor-cell coord to pin the shape to.
    try testing.expect(std.mem.indexOf(u8, out, "<x:Anchor>2, 15, 4, 2, 5, 31, 8, 3</x:Anchor>") != null);
}

test "VML: <v:shape> open tag with quoted '>' inside style attr (REL-704)" {
    const a = testing.allocator;
    // Include a quoted `>` and `/>` inside the style attribute. A
    // naive `>` scan would truncate the open tag mid-attribute.
    const body =
        "<v:shape id=\"_x0000_s1025\" type=\"#_x0000_t202\" style=\"width:>100pt; arrow:/>foo\">" ++
        "<x:ClientData ObjectType=\"Note\">" ++
        "<x:Anchor>2, 15, 4, 2, 5, 31, 8, 3</x:Anchor>" ++
        "<x:Row>4</x:Row><x:Column>3</x:Column>" ++
        "</x:ClientData></v:shape>";
    const src = try wrapVml(a, body);
    defer a.free(src);
    // Insert at col 0 — column 3 should shift to 4, anchor from
    // 2 → 3, to 5 → 6.
    const out = try applyEditToVmlDrawing(a, src, .col, 0, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<x:Column>4</x:Column>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<x:Anchor>3, 15, 4, 2, 6, 31, 8, 3</x:Anchor>") != null);
    // The quoted style attribute survives unchanged.
    try testing.expect(std.mem.indexOf(u8, out, "style=\"width:>100pt; arrow:/>foo\"") != null);
}

// ---------------------------------------------------------------------------
// Comments part rewriter tests (REL-705).
// ---------------------------------------------------------------------------

fn wrapComments(allocator: Allocator, body: []const u8) ![]u8 {
    const head = "<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><authors><author>Author</author></authors><commentList>";
    const tail = "</commentList></comments>";
    var buf = try allocator.alloc(u8, head.len + body.len + tail.len);
    @memcpy(buf[0..head.len], head);
    @memcpy(buf[head.len .. head.len + body.len], body);
    @memcpy(buf[head.len + body.len ..], tail);
    return buf;
}

const sample_comment =
    "<comment ref=\"C5\" authorId=\"0\" shapeId=\"0\">" ++
    "<text><r><t>hello</t></r></text>" ++
    "</comment>";

test "comments: insertColumn before anchor shifts ref's col" {
    const a = testing.allocator;
    const src = try wrapComments(a, sample_comment);
    defer a.free(src);
    // Insert before col B (1-based 2). C5: col=3 >= 2 → 4 (D5).
    const out = try applyEditToCommentsXml(a, src, .col, 2, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "ref=\"D5\"") != null);
}

test "comments: insertRow before anchor shifts ref's row" {
    const a = testing.allocator;
    const src = try wrapComments(a, sample_comment);
    defer a.free(src);
    // Insert before row 3. C5: row=5 >= 3 → 6.
    const out = try applyEditToCommentsXml(a, src, .row, 3, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "ref=\"C6\"") != null);
}

test "comments: deleteColumn at anchor's column drops the entire <comment>" {
    const a = testing.allocator;
    const src = try wrapComments(a, sample_comment);
    defer a.free(src);
    // Delete col C (1-based 3, == ref's col).
    const out = try applyEditToCommentsXml(a, src, .col, 3, .delete);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<comment ") == null);
    try testing.expect(std.mem.indexOf(u8, out, "<text>") == null);
}

test "comments: deleteRow at anchor's row drops the entire <comment>" {
    const a = testing.allocator;
    const src = try wrapComments(a, sample_comment);
    defer a.free(src);
    const out = try applyEditToCommentsXml(a, src, .row, 5, .delete);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<comment ") == null);
}

test "comments: edit past the comment's ref is a no-op" {
    const a = testing.allocator;
    const src = try wrapComments(a, sample_comment);
    defer a.free(src);
    const out = try applyEditToCommentsXml(a, src, .col, 100, .insert);
    defer a.free(out);
    try testing.expectEqualSlices(u8, src, out);
}

test "comments: malformed ref refuses with MalformedCommentsXml" {
    const a = testing.allocator;
    const body =
        "<comment ref=\"NOT_A_REF\" authorId=\"0\" shapeId=\"0\">" ++
        "<text><r><t>x</t></r></text></comment>";
    const src = try wrapComments(a, body);
    defer a.free(src);
    const result = applyEditToCommentsXml(a, src, .col, 2, .insert);
    try testing.expectError(error.MalformedCommentsXml, result);
}

test "comments: two comments rewrite + drop independently" {
    const a = testing.allocator;
    const body = sample_comment ++
        "<comment ref=\"E10\" authorId=\"0\" shapeId=\"1\">" ++
        "<text><r><t>two</t></r></text></comment>";
    const src = try wrapComments(a, body);
    defer a.free(src);
    // Delete col C (== first comment's col). First drops; second
    // (E10, col 5 > 3) shifts to D10.
    const out = try applyEditToCommentsXml(a, src, .col, 3, .delete);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "ref=\"D10\"") != null);
    try testing.expect(std.mem.indexOf(u8, out, "ref=\"C5\"") == null);
    var i: usize = 0;
    var count: usize = 0;
    while (std.mem.indexOfPos(u8, out, i, "<comment ")) |idx| {
        count += 1;
        i = idx + 1;
    }
    try testing.expectEqual(@as(usize, 1), count);
}
