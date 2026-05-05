//! Worksheet XML row/col edit primitives. Shared between
//! `pkg/editor.zig` (legacy save-time path) and `pkg/workbook.zig`
//! (B2 iter-er-4 (3/N) typed-overlay surfaces:
//! `Workbook.insertRow` / `deleteRow` / `insertColumn` /
//! `deleteColumn`). Lifted out of `pkg/editor.zig` verbatim — the
//! XML walker is unchanged; only its visibility moves.
//!
//! Both `applyRowEditToWorksheet` and `applyColEditToWorksheet`
//! take the source sheet XML bytes and an edit kind, and return a
//! freshly-allocated buffer with the row/col attributes shifted in
//! place. The implementation is byte-walk + tag-recognition; it
//! does NOT parse formulas, hyperlinks, data validations,
//! conditional formats, drawings, or tables. Callers are expected
//! to either refuse sheets that carry those constructs (legacy
//! Editor's recordRowEdit / recordColEdit) or run the typed-overlay
//! rewriters first (Workbook.rewriteAllFormulas etc.) and call
//! these helpers on the resulting bytes.

const std = @import("std");
const xlsx = @import("zlsx");

const Allocator = std.mem.Allocator;
const TagOpen = xlsx.TagOpen;
const max_row = xlsx.max_row;
const max_col_1based = xlsx.max_col_1based;
const getAttr = xlsx.getAttr;


/// Apply one column edit (insert or delete at `col_1based`) to a
/// worksheet XML buffer. iter-col-3/4 v1: rewrites `<c r="A1">`
/// column letter, `<col min=N max=M>` bounds, `<mergeCells>` rect
/// bounds, and `<dimension>`. Other elements pass through —
/// recordColEdit refuses sheets that contain formulas / hyperlinks
/// / validations / cond-formats / drawings / tables, so we don't
/// have to handle those here.
pub fn applyColEditToWorksheet(
    allocator: Allocator,
    src: []const u8,
    col_1based: u32,
    kind: RowEditKind,
) ![]u8 {
    var out: std.ArrayListUnmanaged(u8) = .{};
    errdefer out.deinit(allocator);

    var i: usize = 0;
    while (i < src.len) {
        const next_lt = std.mem.indexOfScalarPos(u8, src, i, '<') orelse {
            try out.appendSlice(allocator, src[i..]);
            return try out.toOwnedSlice(allocator);
        };
        try out.appendSlice(allocator, src[i..next_lt]);
        i = next_lt;

        if (matchTagAt(src, i, "c")) |t| {
            try processCellTagCol(allocator, &out, src, t, col_1based, kind, &i);
        } else if (matchTagAt(src, i, "col")) |t| {
            try processColTag(allocator, &out, src, t, col_1based, kind);
            i = t.after_open;
        } else if (matchTagAt(src, i, "mergeCell")) |t| {
            try processMergeCellTagCol(allocator, &out, src, t, col_1based, kind);
            i = t.after_open;
        } else if (matchTagAt(src, i, "dimension")) |t| {
            try processDimensionTagCol(allocator, &out, src, t, col_1based, kind);
            i = t.after_open;
        } else {
            try out.append(allocator, '<');
            i += 1;
        }
    }
    return try out.toOwnedSlice(allocator);
}

fn processCellTagCol(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    col_1based: u32,
    kind: RowEditKind,
    i: *usize,
) !void {
    const attrs = src[t.start + "<c".len .. t.after_open - 1];
    const trimmed = std.mem.trimRight(u8, attrs, " \t\r\n");
    const is_self_closing = trimmed.len > 0 and trimmed[trimmed.len - 1] == '/';
    const attrs_for_lookup = if (is_self_closing) trimmed[0 .. trimmed.len - 1] else trimmed;
    const r_attr = getAttr(attrs_for_lookup, "r");
    if (r_attr == null) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        i.* = t.after_open;
        return;
    }
    const ref = r_attr.?;
    var letters_end: usize = 0;
    while (letters_end < ref.len and ref[letters_end] >= 'A' and ref[letters_end] <= 'Z') letters_end += 1;
    if (letters_end == 0 or letters_end == ref.len) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        i.* = t.after_open;
        return;
    }
    const old_col = parseColLetters(ref[0..letters_end]) orelse {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        i.* = t.after_open;
        return;
    };

    // Delete-match: drop the entire <c> element (open + body + close).
    if (kind == .delete and old_col == col_1based) {
        if (is_self_closing) {
            i.* = t.after_open;
        } else {
            const close = std.mem.indexOfPos(u8, src, t.after_open, "</c>") orelse t.after_open;
            i.* = if (close + "</c>".len <= src.len) close + "</c>".len else t.after_open;
        }
        return;
    }

    var new_col: u32 = old_col;
    switch (kind) {
        .insert => if (old_col >= col_1based) {
            if (old_col >= max_col_1based) return error.ColEditExceedsMaxCol;
            new_col = old_col + 1;
        },
        .delete => if (old_col > col_1based) {
            new_col = old_col - 1;
        },
    }
    if (new_col == old_col) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        i.* = t.after_open;
        return;
    }
    var letters_buf: [8]u8 = undefined;
    const new_letters = formatColLetters(&letters_buf, new_col - 1) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        i.* = t.after_open;
        return;
    };
    var new_ref_buf: [16]u8 = undefined;
    const new_ref = try std.fmt.bufPrint(&new_ref_buf, "{s}{s}", .{ new_letters, ref[letters_end..] });
    try writeWithReplacedAttr(allocator, out, src, t, "<c".len, "r", new_ref);
    i.* = t.after_open;
}

fn processColTag(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    col_1based: u32,
    kind: RowEditKind,
) !void {
    // <col min="N" max="M" .../> covers a contiguous range of cols.
    // For inserts: shift min/max if >= col_1based; if min < col <= max,
    // split into two col entries (we just shift max in v1; the
    // inserted column gets default formatting).
    // For deletes: similar shift; if min == max == col, drop entry.
    //   Else if min == col, increment min.
    //   Else if max == col, decrement max.
    //   Else if min <= col <= max, shrink max by 1.
    const attrs = src[t.start + "<col".len .. t.after_open - 1];
    const min_str = getAttr(attrs, "min");
    const max_str = getAttr(attrs, "max");
    if (min_str == null or max_str == null) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    const old_min = std.fmt.parseInt(u32, min_str.?, 10) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    const old_max = std.fmt.parseInt(u32, max_str.?, 10) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    var new_min = old_min;
    var new_max = old_max;
    var drop = false;
    var split_emit_first_max: ?u32 = null;
    switch (kind) {
        .insert => {
            // Insert point strictly INSIDE the existing range:
            // split into two `<col>` entries, leaving the inserted
            // column with default formatting. e.g. <col min=2 max=4>
            // + insertColumn(3) → <col min=2 max=2/> +
            // <col min=4 max=5/>; col 3 (the inserted one) gets no
            // formatting entry.
            if (old_min < col_1based and col_1based <= old_max) {
                split_emit_first_max = col_1based - 1;
                new_min = col_1based + 1;
                new_max = old_max + 1;
            } else {
                if (old_min >= col_1based) new_min = old_min + 1;
                if (old_max >= col_1based) new_max = old_max + 1;
            }
            // Insert that pushes a <col> range past XFD (16384)
            // would emit max="16385", which Excel rejects as out
            // of range. Mirror the cell-reference path's error.
            if (new_max > max_col_1based or new_min > max_col_1based) {
                return error.ColEditExceedsMaxCol;
            }
        },
        .delete => {
            if (old_min == col_1based and old_max == col_1based) {
                drop = true;
            } else {
                if (old_min > col_1based) new_min = old_min - 1;
                if (old_max >= col_1based) new_max = old_max - 1;
                if (new_min > new_max) drop = true;
            }
        },
    }
    if (drop) return;
    if (split_emit_first_max == null and new_min == old_min and new_max == old_max) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    if (split_emit_first_max) |first_max| {
        // Emit two <col> entries: one for the unchanged head, one
        // for the shifted tail.
        try emitColEntry(allocator, out, attrs, src, t, old_min, first_max);
        try emitColEntry(allocator, out, attrs, src, t, new_min, new_max);
    } else {
        try emitColEntry(allocator, out, attrs, src, t, new_min, new_max);
    }
}

/// Emit a `<col>` tag with its attrs preserved verbatim except
/// `min` and `max`, which take the supplied values.
fn emitColEntry(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    attrs: []const u8,
    src: []const u8,
    t: TagOpen,
    new_min: u32,
    new_max: u32,
) !void {
    var min_buf: [12]u8 = undefined;
    var max_buf: [12]u8 = undefined;
    const min_s = try std.fmt.bufPrint(&min_buf, "{d}", .{new_min});
    const max_s = try std.fmt.bufPrint(&max_buf, "{d}", .{new_max});
    try out.appendSlice(allocator, "<col");
    var ai: usize = 0;
    while (ai < attrs.len) {
        const ws_start = ai;
        while (ai < attrs.len and (attrs[ai] == ' ' or attrs[ai] == '\t' or
            attrs[ai] == '\n' or attrs[ai] == '\r')) ai += 1;
        try out.appendSlice(allocator, attrs[ws_start..ai]);
        if (ai >= attrs.len) break;
        const name_start = ai;
        while (ai < attrs.len and attrs[ai] != '=' and attrs[ai] != ' ' and
            attrs[ai] != '\t' and attrs[ai] != '\n' and attrs[ai] != '\r') ai += 1;
        const aname = attrs[name_start..ai];
        while (ai < attrs.len and attrs[ai] != '=') ai += 1;
        if (ai >= attrs.len) break;
        ai += 1;
        while (ai < attrs.len and (attrs[ai] == ' ' or attrs[ai] == '\t' or
            attrs[ai] == '\n' or attrs[ai] == '\r')) ai += 1;
        if (ai >= attrs.len or (attrs[ai] != '"' and attrs[ai] != '\'')) break;
        const quote = attrs[ai];
        ai += 1;
        const val_start = ai;
        while (ai < attrs.len and attrs[ai] != quote) ai += 1;
        const val = attrs[val_start..ai];
        if (ai < attrs.len) ai += 1;
        try out.appendSlice(allocator, aname);
        try out.append(allocator, '=');
        try out.append(allocator, quote);
        if (std.mem.eql(u8, aname, "min")) {
            try out.appendSlice(allocator, min_s);
        } else if (std.mem.eql(u8, aname, "max")) {
            try out.appendSlice(allocator, max_s);
        } else {
            try out.appendSlice(allocator, val);
        }
        try out.append(allocator, quote);
    }
    // OOXML <col/> is always an empty element. Detect the original
    // self-closing form (last non-ws byte of attrs is `/`) and
    // emit `/>` accordingly, so we don't leave open `<col>` tags.
    const trimmed_attrs = std.mem.trimRight(u8, attrs, " \t\r\n");
    const was_self_closing = trimmed_attrs.len > 0 and trimmed_attrs[trimmed_attrs.len - 1] == '/';
    if (was_self_closing) {
        try out.appendSlice(allocator, "/>");
    } else {
        try out.appendSlice(allocator, src[t.after_open - 1 .. t.after_open]);
    }
}

fn processMergeCellTagCol(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    col_1based: u32,
    kind: RowEditKind,
) !void {
    const attrs = src[t.start + "<mergeCell".len .. t.after_open - 1];
    const r_attr = getAttr(attrs, "ref");
    if (r_attr == null) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    const ref = r_attr.?;
    const colon = std.mem.indexOfScalar(u8, ref, ':') orelse {
        // Single-cell merge.
        if (kind == .delete) {
            const c = parseColFromA1(ref) orelse 0;
            if (c == col_1based) return; // drop
        }
        var new_buf: [16]u8 = undefined;
        const new_ref = shiftSingleA1Col(ref, col_1based, kind, &new_buf, false) catch {
            try out.appendSlice(allocator, src[t.start..t.after_open]);
            return;
        };
        if (std.mem.eql(u8, ref, new_ref)) {
            try out.appendSlice(allocator, src[t.start..t.after_open]);
            return;
        }
        try writeWithReplacedAttr(allocator, out, src, t, "<mergeCell".len, "ref", new_ref);
        return;
    };
    if (kind == .delete) {
        const tl_col = parseColFromA1(ref[0..colon]) orelse 0;
        const br_col = parseColFromA1(ref[colon + 1 ..]) orelse 0;
        if (tl_col == col_1based and br_col == col_1based) return;
    }
    var tl_buf: [16]u8 = undefined;
    var br_buf: [16]u8 = undefined;
    const tl_new = shiftSingleA1Col(ref[0..colon], col_1based, kind, &tl_buf, false) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    const br_new = shiftSingleA1Col(ref[colon + 1 ..], col_1based, kind, &br_buf, true) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    var new_ref_buf: [40]u8 = undefined;
    const new_ref = try std.fmt.bufPrint(&new_ref_buf, "{s}:{s}", .{ tl_new, br_new });
    if (std.mem.eql(u8, ref, new_ref)) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    try writeWithReplacedAttr(allocator, out, src, t, "<mergeCell".len, "ref", new_ref);
}

fn processDimensionTagCol(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    col_1based: u32,
    kind: RowEditKind,
) !void {
    const attrs = src[t.start + "<dimension".len .. t.after_open - 1];
    const r_attr = getAttr(attrs, "ref");
    if (r_attr == null) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    const ref = r_attr.?;
    const colon = std.mem.indexOfScalar(u8, ref, ':') orelse {
        // Single-cell dimension. On a delete-match (the only used
        // column == col_1based) the cell is gone; fall back to "A1"
        // so the dimension stays valid (Excel recomputes on open).
        if (kind == .delete) {
            if (parseColFromA1(ref)) |c| {
                if (c == col_1based) {
                    try writeWithReplacedAttr(allocator, out, src, t, "<dimension".len, "ref", "A1");
                    return;
                }
            }
        }
        var b: [16]u8 = undefined;
        const new_ref = shiftSingleA1Col(ref, col_1based, kind, &b, false) catch {
            try out.appendSlice(allocator, src[t.start..t.after_open]);
            return;
        };
        if (std.mem.eql(u8, ref, new_ref)) {
            try out.appendSlice(allocator, src[t.start..t.after_open]);
            return;
        }
        try writeWithReplacedAttr(allocator, out, src, t, "<dimension".len, "ref", new_ref);
        return;
    };
    // Detect "entire range is the deleted col" up front: both
    // corners at col_1based on a delete. Fall back to a safe
    // sentinel "A1" so the dimension stays valid (Excel
    // recomputes on open).
    if (kind == .delete) {
        const tl_col = parseColFromA1(ref[0..colon]) orelse 0;
        const br_col = parseColFromA1(ref[colon + 1 ..]) orelse 0;
        if (tl_col == col_1based and br_col == col_1based) {
            try writeWithReplacedAttr(allocator, out, src, t, "<dimension".len, "ref", "A1");
            return;
        }
    }
    var tl_buf: [16]u8 = undefined;
    var br_buf: [16]u8 = undefined;
    const tl_new = shiftSingleA1Col(ref[0..colon], col_1based, kind, &tl_buf, false) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    const br_new = shiftSingleA1Col(ref[colon + 1 ..], col_1based, kind, &br_buf, true) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    var new_ref_buf: [40]u8 = undefined;
    const new_ref = try std.fmt.bufPrint(&new_ref_buf, "{s}:{s}", .{ tl_new, br_new });
    if (std.mem.eql(u8, ref, new_ref)) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    try writeWithReplacedAttr(allocator, out, src, t, "<dimension".len, "ref", new_ref);
}

/// Shift the column letters of an A1 ref. `is_br_corner` shrinks
/// the BR corner by one on delete-match.
fn shiftSingleA1Col(ref: []const u8, col_1based: u32, kind: RowEditKind, buf: *[16]u8, is_br_corner: bool) ![]const u8 {
    var letters_end: usize = 0;
    while (letters_end < ref.len and ref[letters_end] >= 'A' and ref[letters_end] <= 'Z') letters_end += 1;
    if (letters_end == 0) return error.MalformedXml;
    const old_col = parseColLetters(ref[0..letters_end]) orelse return error.MalformedXml;
    var new_col: u32 = old_col;
    switch (kind) {
        .insert => if (old_col >= col_1based) {
            if (old_col >= max_col_1based) return error.ColEditExceedsMaxCol;
            new_col = old_col + 1;
        },
        .delete => if (old_col > col_1based) {
            new_col = old_col - 1;
        } else if (old_col == col_1based and is_br_corner and old_col > 1) {
            new_col = old_col - 1;
        },
    }
    if (new_col == old_col) {
        @memcpy(buf[0..ref.len], ref);
        return buf[0..ref.len];
    }
    var letters_buf: [8]u8 = undefined;
    const new_letters = try formatColLetters(&letters_buf, new_col - 1);
    return try std.fmt.bufPrint(buf, "{s}{s}", .{ new_letters, ref[letters_end..] });
}

fn parseColFromA1(ref: []const u8) ?u32 {
    var letters_end: usize = 0;
    while (letters_end < ref.len and ref[letters_end] >= 'A' and ref[letters_end] <= 'Z') letters_end += 1;
    if (letters_end == 0) return null;
    return parseColLetters(ref[0..letters_end]);
}

/// Render a 0-based col_idx as A1 letters (A=0, Z=25, AA=26, ...).
/// Caller-provided buffer; result borrows from buf.
fn formatColLetters(buf: *[8]u8, col_idx: u32) ![]const u8 {
    var col_chars: [8]u8 = undefined;
    var pos: usize = col_chars.len;
    var c = col_idx + 1;
    while (c > 0) {
        c -= 1;
        pos -= 1;
        if (pos == std.math.maxInt(usize)) return error.ColumnIndexOutOfRange;
        col_chars[pos] = @intCast('A' + (c % 26));
        c /= 26;
        if (pos == 0 and c > 0) return error.ColumnIndexOutOfRange;
    }
    const len = col_chars.len - pos;
    @memcpy(buf[0..len], col_chars[pos..]);
    return buf[0..len];
}

/// Apply one row edit (insert or delete at `row`) to a worksheet
/// XML buffer. iter-row-2/3 v1: rewrites `<row r=>` (renumber or
/// drop), `<c r="A1">` row component, `<mergeCells>` rect bounds,
/// and `<dimension>`. Other elements pass through verbatim — the
/// caller's recordRowEdit guard refuses sheets that contain
/// formulas / hyperlinks / validations / cond-formats / drawings,
/// so we don't have to handle those here.
pub const RowEditKind = enum { insert, delete };

pub fn applyRowEditToWorksheet(
    allocator: Allocator,
    src: []const u8,
    row: u32,
    kind: RowEditKind,
) ![]u8 {
    var out: std.ArrayListUnmanaged(u8) = .{};
    errdefer out.deinit(allocator);

    var i: usize = 0;
    while (i < src.len) {
        // Find the next interesting tag opening.
        const next_lt = std.mem.indexOfScalarPos(u8, src, i, '<') orelse {
            try out.appendSlice(allocator, src[i..]);
            return try out.toOwnedSlice(allocator);
        };
        try out.appendSlice(allocator, src[i..next_lt]);
        i = next_lt;

        // Identify which tag we're at.
        if (matchTagAt(src, i, "row")) |t| {
            try processRowTag(allocator, &out, src, t, row, kind, &i);
        } else if (matchTagAt(src, i, "c")) |t| {
            try processCellTag(allocator, &out, src, t, row, kind);
            i = t.after_open;
        } else if (matchTagAt(src, i, "mergeCell")) |t| {
            try processMergeCellTag(allocator, &out, src, t, row, kind);
            i = t.after_open;
        } else if (matchTagAt(src, i, "dimension")) |t| {
            try processDimensionTag(allocator, &out, src, t, row, kind);
            i = t.after_open;
        } else {
            // Some other tag; emit `<` and continue past it.
            try out.append(allocator, '<');
            i += 1;
        }
    }
    return try out.toOwnedSlice(allocator);
}

fn matchTagAt(src: []const u8, i: usize, tag: []const u8) ?TagOpen {
    if (i >= src.len or src[i] != '<') return null;
    const after = i + 1 + tag.len;
    if (after > src.len) return null;
    if (!std.mem.eql(u8, src[i + 1 .. i + 1 + tag.len], tag)) return null;
    const c = src[after];
    if (c != ' ' and c != '\t' and c != '\n' and c != '\r' and c != '/' and c != '>') return null;
    const gt = std.mem.indexOfScalarPos(u8, src, i, '>') orelse return null;
    return .{ .start = i, .after_open = gt + 1 };
}

fn processRowTag(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    row: u32,
    kind: RowEditKind,
    i: *usize,
) !void {
    const attrs = src[t.start + "<row".len .. t.after_open - 1];
    const trimmed = std.mem.trimRight(u8, attrs, " \t\r\n");
    const is_self_closing = trimmed.len > 0 and trimmed[trimmed.len - 1] == '/';
    const attrs_for_lookup = if (is_self_closing) trimmed[0 .. trimmed.len - 1] else trimmed;
    const r_attr = getAttr(attrs_for_lookup, "r");
    const old_r: ?u32 = if (r_attr) |s| (std.fmt.parseInt(u32, s, 10) catch null) else null;

    // Decide: drop entirely (delete-match), shift, or pass-through.
    var drop = false;
    var new_r: ?u32 = null;
    if (old_r) |r_val| {
        switch (kind) {
            .insert => if (r_val >= row) {
                if (r_val >= max_row) return error.RowEditExceedsMaxRow;
                new_r = r_val + 1;
            },
            .delete => if (r_val == row) {
                drop = true;
            } else if (r_val > row) {
                new_r = r_val - 1;
            },
        }
    }

    if (drop) {
        // Skip the entire <row>...</row> block. For self-closing
        // form, just skip the opener.
        if (is_self_closing) {
            i.* = t.after_open;
        } else {
            const close = std.mem.indexOfPos(u8, src, t.after_open, "</row>") orelse t.after_open;
            i.* = if (close + "</row>".len <= src.len) close + "</row>".len else t.after_open;
        }
        return;
    }

    // Emit the row open with possibly-rewritten r=.
    if (new_r == null and old_r != null) {
        // Same row number — emit verbatim.
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        i.* = t.after_open;
        // If body form, also recursively process inner tags via
        // the outer loop (they include <c> tags whose row will
        // be rewritten on subsequent iterations).
        return;
    }
    if (new_r == null) {
        // No r= attribute at all — pass through.
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        i.* = t.after_open;
        return;
    }
    // Rewrite r="..." to the new value.
    try writeWithReplacedRowAttr(allocator, out, src, t, "<row".len, new_r.?);
    i.* = t.after_open;
}

fn processCellTag(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    row: u32,
    kind: RowEditKind,
) !void {
    const attrs = src[t.start + "<c".len .. t.after_open - 1];
    const trimmed = std.mem.trimRight(u8, attrs, " \t\r\n");
    const is_self_closing = trimmed.len > 0 and trimmed[trimmed.len - 1] == '/';
    const attrs_for_lookup = if (is_self_closing) trimmed[0 .. trimmed.len - 1] else trimmed;
    const r_attr = getAttr(attrs_for_lookup, "r");
    if (r_attr == null) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    const ref = r_attr.?;
    var letters_end: usize = 0;
    while (letters_end < ref.len and ref[letters_end] >= 'A' and ref[letters_end] <= 'Z') letters_end += 1;
    if (letters_end == 0 or letters_end == ref.len) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    const old_row = std.fmt.parseInt(u32, ref[letters_end..], 10) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    var new_row: u32 = old_row;
    switch (kind) {
        .insert => if (old_row >= row) {
            if (old_row >= max_row) return error.RowEditExceedsMaxRow;
            new_row = old_row + 1;
        },
        .delete => if (old_row > row) {
            new_row = old_row - 1;
        }, // (a cell at the deleted row is dropped along with the <row> block)
    }
    if (new_row == old_row) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    // Rewrite the r= attribute, preserving column letters.
    var new_ref_buf: [16]u8 = undefined;
    const new_ref = try std.fmt.bufPrint(&new_ref_buf, "{s}{d}", .{ ref[0..letters_end], new_row });
    try writeWithReplacedAttr(allocator, out, src, t, "<c".len, "r", new_ref);
}

fn processMergeCellTag(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    row: u32,
    kind: RowEditKind,
) !void {
    const attrs = src[t.start + "<mergeCell".len .. t.after_open - 1];
    const r_attr = getAttr(attrs, "ref");
    if (r_attr == null) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    const ref = r_attr.?;
    const colon = std.mem.indexOfScalar(u8, ref, ':') orelse {
        // Single-cell merge. On delete-match, drop the entire tag
        // so it doesn't transfer onto whatever shifts into the
        // deleted row.
        if (kind == .delete) {
            const tl_row = parseRowFromA1(ref) orelse 0;
            if (tl_row == row) return; // drop tag entirely
        }
        var new_ref_buf: [16]u8 = undefined;
        const new_ref = shiftSingleA1(ref, row, kind, &new_ref_buf, false) catch {
            try out.appendSlice(allocator, src[t.start..t.after_open]);
            return;
        };
        if (std.mem.eql(u8, ref, new_ref)) {
            try out.appendSlice(allocator, src[t.start..t.after_open]);
            return;
        }
        try writeWithReplacedAttr(allocator, out, src, t, "<mergeCell".len, "ref", new_ref);
        return;
    };
    // Rectangle merge. On delete, if BOTH the top-left and
    // bottom-right rows equal the deleted row, the entire merge
    // is on the dead row — drop the tag.
    if (kind == .delete) {
        const tl_row = parseRowFromA1(ref[0..colon]) orelse 0;
        const br_row = parseRowFromA1(ref[colon + 1 ..]) orelse 0;
        if (tl_row == row and br_row == row) return;
    }
    var tl_buf: [16]u8 = undefined;
    var br_buf: [16]u8 = undefined;
    const tl_new = shiftSingleA1(ref[0..colon], row, kind, &tl_buf, false) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    const br_new = shiftSingleA1(ref[colon + 1 ..], row, kind, &br_buf, true) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    var new_ref_buf: [40]u8 = undefined;
    const new_ref = try std.fmt.bufPrint(&new_ref_buf, "{s}:{s}", .{ tl_new, br_new });
    if (std.mem.eql(u8, ref, new_ref)) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    try writeWithReplacedAttr(allocator, out, src, t, "<mergeCell".len, "ref", new_ref);
}

/// Parse the row component (digits) from an A1-style ref. Returns
/// null on malformed input.
fn parseRowFromA1(ref: []const u8) ?u32 {
    var letters_end: usize = 0;
    while (letters_end < ref.len and ref[letters_end] >= 'A' and ref[letters_end] <= 'Z') letters_end += 1;
    if (letters_end == 0 or letters_end == ref.len) return null;
    return std.fmt.parseInt(u32, ref[letters_end..], 10) catch null;
}

fn processDimensionTag(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    row: u32,
    kind: RowEditKind,
) !void {
    // Reuse mergeCell logic — same `ref="A1:Z100"` / `ref="A1"` shape.
    const attrs = src[t.start + "<dimension".len .. t.after_open - 1];
    const r_attr = getAttr(attrs, "ref");
    if (r_attr == null) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    const ref = r_attr.?;
    const colon = std.mem.indexOfScalar(u8, ref, ':') orelse {
        // Single-cell dimension. On a delete-match (the only used
        // row == deleted row) the cell is gone; fall back to "A1"
        // so the dimension stays valid (Excel recomputes on open).
        if (kind == .delete) {
            if (parseRowFromA1(ref)) |r| {
                if (r == row) {
                    try writeWithReplacedAttr(allocator, out, src, t, "<dimension".len, "ref", "A1");
                    return;
                }
            }
        }
        var b: [16]u8 = undefined;
        const new_ref = shiftSingleA1(ref, row, kind, &b, false) catch {
            try out.appendSlice(allocator, src[t.start..t.after_open]);
            return;
        };
        if (std.mem.eql(u8, ref, new_ref)) {
            try out.appendSlice(allocator, src[t.start..t.after_open]);
            return;
        }
        try writeWithReplacedAttr(allocator, out, src, t, "<dimension".len, "ref", new_ref);
        return;
    };
    // Detect "entire range is the deleted row" up front: both
    // corners at the deleted row on a delete. Fall back to a safe
    // sentinel "A1" so the dimension stays valid (Excel recomputes
    // on open).
    if (kind == .delete) {
        const tl_row = parseRowFromA1(ref[0..colon]) orelse 0;
        const br_row = parseRowFromA1(ref[colon + 1 ..]) orelse 0;
        if (tl_row == row and br_row == row) {
            try writeWithReplacedAttr(allocator, out, src, t, "<dimension".len, "ref", "A1");
            return;
        }
    }
    var tl_buf: [16]u8 = undefined;
    var br_buf: [16]u8 = undefined;
    const tl_new = shiftSingleA1(ref[0..colon], row, kind, &tl_buf, false) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    const br_new = shiftSingleA1(ref[colon + 1 ..], row, kind, &br_buf, true) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    var new_ref_buf: [40]u8 = undefined;
    const new_ref = try std.fmt.bufPrint(&new_ref_buf, "{s}:{s}", .{ tl_new, br_new });
    if (std.mem.eql(u8, ref, new_ref)) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    try writeWithReplacedAttr(allocator, out, src, t, "<dimension".len, "ref", new_ref);
}

/// Shift the row component of an A1-style ref. `is_br_corner`
/// says "this is the bottom-right of a rectangular range" — on a
/// delete-match (old_row == row), BR corners shrink by 1 so the
/// range collapses to the row above the deleted one.
fn shiftSingleA1(ref: []const u8, row: u32, kind: RowEditKind, buf: *[16]u8, is_br_corner: bool) ![]const u8 {
    var letters_end: usize = 0;
    while (letters_end < ref.len and ref[letters_end] >= 'A' and ref[letters_end] <= 'Z') letters_end += 1;
    if (letters_end == 0 or letters_end == ref.len) return error.MalformedXml;
    const old_row = std.fmt.parseInt(u32, ref[letters_end..], 10) catch return error.MalformedXml;
    var new_row: u32 = old_row;
    switch (kind) {
        .insert => if (old_row >= row) {
            // Cap at Excel's max row (1_048_576). If a shift would
            // overflow, refuse the entire edit — saving an invalid
            // workbook is worse than a typed error at the call
            // site (recordRowEdit catches the propagation).
            if (old_row >= max_row) return error.RowEditExceedsMaxRow;
            new_row = old_row + 1;
        },
        .delete => if (old_row > row) {
            new_row = old_row - 1;
        } else if (old_row == row and is_br_corner and old_row > 1) {
            // BR shrink: don't go below row 1. If the entire range
            // collapses to row 0 we'd produce invalid XML; leaving
            // it at row 1 keeps the dimension valid (Excel will
            // recompute on next save anyway).
            new_row = old_row - 1;
        },
    }
    if (new_row == old_row) {
        @memcpy(buf[0..ref.len], ref);
        return buf[0..ref.len];
    }
    return try std.fmt.bufPrint(buf, "{s}{d}", .{ ref[0..letters_end], new_row });
}

/// Emit the original `<tag attrs>` with `attr_name="..."` value
/// replaced by `new_value`. `tag_name_len` is the length of the
/// tag name including the leading `<` (e.g. `"<c".len` = 2).
fn writeWithReplacedAttr(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    tag_name_len: usize,
    attr_name: []const u8,
    new_value: []const u8,
) !void {
    const attrs_full_start = t.start + tag_name_len;
    const attrs_full_end = t.after_open - 1;
    const attrs = src[attrs_full_start..attrs_full_end];
    // Find the `name="..."` occurrence inside attrs.
    var pat_buf: [32]u8 = undefined;
    const pat = try std.fmt.bufPrint(&pat_buf, "{s}=\"", .{attr_name});
    const pat_off = std.mem.indexOf(u8, attrs, pat) orelse {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    const val_start_in_src = attrs_full_start + pat_off + pat.len;
    const val_end_in_src = std.mem.indexOfScalarPos(u8, src, val_start_in_src, '"') orelse {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    try out.appendSlice(allocator, src[t.start..val_start_in_src]);
    try out.appendSlice(allocator, new_value);
    try out.appendSlice(allocator, src[val_end_in_src..t.after_open]);
}

/// Rewrite the row attribute on `<row r="…">`. Just delegates to
/// writeWithReplacedAttr formatted as a number.
fn writeWithReplacedRowAttr(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    tag_name_len: usize,
    new_row: u32,
) !void {
    var buf: [16]u8 = undefined;
    const s = try std.fmt.bufPrint(&buf, "{d}", .{new_row});
    try writeWithReplacedAttr(allocator, out, src, t, tag_name_len, "r", s);
}


/// Parse uppercase A-Z letters as a 1-based Excel column index
/// (A=1, B=2, ..., Z=26, AA=27, ..., XFD=16384). Returns null on
/// empty input or anything past `max_col_1based`.
pub fn parseColLetters(s: []const u8) ?u32 {
    if (s.len == 0) return null;
    var n: u32 = 0;
    for (s) |c| {
        if (c < 'A' or c > 'Z') return null;
        n = n * 26 + (c - 'A' + 1);
        if (n > max_col_1based) return null;
    }
    return n;
}

/// Render `col_idx` (0-based) as A, B, ..., Z, AA, AB, ... into `buf`.
/// Capacity 8 is more than enough (Excel max is XFD = 3 letters).
pub fn colLetterEditor(buf: []u8, col_idx: u32) []u8 {
    var n: u32 = col_idx + 1;
    var i: usize = 0;
    while (n > 0) {
        const r = (n - 1) % 26;
        buf[i] = 'A' + @as(u8, @intCast(r));
        i += 1;
        n = (n - 1) / 26;
    }
    std.mem.reverse(u8, buf[0..i]);
    return buf[0..i];
}
