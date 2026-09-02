//! Worksheet XML row/col edit primitives. Shared between
//! `pkg/editor.zig` (legacy save-time path) and `pkg/workbook.zig`
//! (B2 iter-er-4 (3/N) typed-overlay surfaces:
//! `Workbook.insertRow` / `deleteRow` / `insertColumn` /
//! `deleteColumn`). Lifted out of `pkg/editor.zig` verbatim — the
//! XML walker is unchanged; only its visibility moves.
//!
//! Coordinate-bearing elements handled here: `<row r>`, `<c r>`,
//! `<col min/max>`, `<mergeCell ref>`, `<dimension ref>`,
//! `<pane xSplit/ySplit/topLeftCell>`, `<autoFilter ref>` +
//! `<filterColumn colId>`, `<sheetView topLeftCell>`,
//! `<selection activeCell/sqref>`, `<sortState ref>`,
//! `<sortCondition ref>`, `<conditionalFormatting sqref>` and
//! `<dataValidation sqref>`.
//!
//! Both `applyRowEditToWorksheet` and `applyColEditToWorksheet`
//! take the source sheet XML bytes and an edit kind, and return a
//! freshly-allocated buffer with the row/col attributes shifted in
//! place. The implementation is byte-walk + tag-recognition; it
//! does NOT parse formulas, hyperlinks, drawings, or tables — and of
//! a DV/CF block it moves only the `sqref` envelope. Callers run the
//! typed-overlay rewriters (Workbook.rewriteAllFormulas, the DV/CF
//! formula sweep, the table/drawing walkers) around these helpers;
//! historically the legacy Editor refused such sheets instead
//! (recordRowEdit / recordColEdit), which is why the envelope shift
//! arrived late (Codex #216 r1 S3B-REL-301).

const std = @import("std");
const xlsx = @import("zlsx");
const coords = @import("zlsx_refs");
// Only for `skipNonElement` — the decoy-aware step the `<xm:f>` scan
// shares with the workbook-part walkers. std-only, so no cycle.
const workbook_xml = @import("typed_parts/workbook_xml.zig");
// Only for `decodeXmlEntities` — a DV/CF `sqref` is an entity-bearing
// attribute carrier, and the shift must move what the value MEANS.
const store_mod = @import("store.zig");

const Allocator = std.mem.Allocator;
const TagOpen = xlsx.TagOpen;
const max_row = xlsx.max_row;
const max_col_1based = xlsx.max_col_1based;
const getAttr = xlsx.getAttr;

/// Apply one column edit (insert or delete at `col_1based`) to a
/// worksheet XML buffer: `<c r="A1">` column letters, `<col min/max>`
/// bounds, merge rects, `<dimension>`, panes, autoFilter, view state,
/// sort state, DV/CF `sqref` envelopes and `<xm:sqref>`. Formula
/// bodies, hyperlinks, drawings and tables move in their own
/// workbook-layer sweeps around this transform.
pub fn applyColEditToWorksheet(
    allocator: Allocator,
    src: []const u8,
    col_1based: u32,
    kind: RowEditKind,
) ![]u8 {
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

        // A comment, CDATA section, PI or DOCTYPE passes through
        // VERBATIM: a `<conditionalFormatting` or `<mergeCell` spelled
        // inside one is prose, not a tag — dispatching on it rewrote
        // (or, under the strict sqref posture, refused on) bytes that
        // are not elements (Codex #216 r4 S3B-REL-804). An
        // unterminated construct falls through to the ordinary
        // emit-`<`-and-continue posture.
        if (i + 1 < src.len and (src[i + 1] == '!' or src[i + 1] == '?')) {
            if (workbook_xml.skipNonElement(src, i)) |end| {
                try out.appendSlice(allocator, src[i..end]);
                i = end;
                continue;
            } else |_| {}
        }

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
        } else if (matchTagAt(src, i, "pane")) |t| {
            try processPaneTagCol(allocator, &out, src, t, col_1based, kind);
            i = t.after_open;
        } else if (matchTagAt(src, i, "autoFilter")) |t| {
            try processAutoFilterTagCol(allocator, &out, src, t, col_1based, kind, &i);
        } else if (matchTagAt(src, i, "sheetView")) |t| {
            try processSheetViewTagCol(allocator, &out, src, t, col_1based, kind);
            i = t.after_open;
        } else if (matchTagAt(src, i, "selection")) |t| {
            try processSelectionTagCol(allocator, &out, src, t, col_1based, kind);
            i = t.after_open;
        } else if (matchTagAt(src, i, "pivotSelection")) |t| {
            try processPivotSelectionTag(allocator, &out, src, t, .col, col_1based, kind);
            i = t.after_open;
        } else if (matchTagAt(src, i, "sortState")) |t| {
            try processSortStateTagCol(allocator, &out, src, t, col_1based, kind, &i);
        } else if (matchTagAt(src, i, "sortCondition")) |t| {
            try processSortConditionTagCol(allocator, &out, src, t, col_1based, kind, &i);
        } else if (matchTagAt(src, i, "conditionalFormatting")) |t| {
            try processSqrefListTag(allocator, &out, src, t, "<conditionalFormatting".len, .col, col_1based, kind);
            i = t.after_open;
        } else if (matchTagAt(src, i, "dataValidation")) |t| {
            try processSqrefListTag(allocator, &out, src, t, "<dataValidation".len, .col, col_1based, kind);
            i = t.after_open;
        } else if (matchTagAt(src, i, "xm:sqref")) |t| {
            try processXmSqrefTag(allocator, &out, src, t, .col, col_1based, kind, &i);
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
    const trimmed = std.mem.trimEnd(u8, attrs, " \t\r\n");
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
    const trimmed_attrs = std.mem.trimEnd(u8, attrs, " \t\r\n");
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

/// Rewrite `<autoFilter ref="A1:E10">…</autoFilter>` for a column
/// edit. Shifts the col halves of `ref` on the open tag AND walks
/// `<filterColumn colId="N">` children: `colId` is a 0-based offset
/// from the autoFilter's left edge (ECMA-376 §18.3.2.7), so any col
/// edit that changes the range's starting column requires
/// recomputing every surviving filterColumn's `colId`. A
/// filterColumn whose absolute column was deleted is dropped
/// entirely (open + body + close).
///
/// On range collapse (single-column filter where that column is
/// deleted), the entire `<autoFilter>` is dropped.
///
/// Nested `<sortState ref="…">` and its `<sortCondition ref="…">`
/// children are rewritten too (iter-sv-1), closing the caveat this
/// comment used to carry. One implementation covers all three
/// contexts — sheet-bare, autoFilter-nested, and table-nested (which
/// `pkg/table_edit.zig` reaches by delegating here).
pub fn processAutoFilterTagCol(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    col_1based: u32,
    kind: RowEditKind,
    i: *usize,
) !void {
    const attrs_full = src[t.start + "<autoFilter".len .. t.after_open - 1];
    const trimmed = std.mem.trimEnd(u8, attrs_full, " \t\r\n");
    const is_self_closing = trimmed.len > 0 and trimmed[trimmed.len - 1] == '/';
    const attrs_for_lookup = if (is_self_closing) trimmed[0 .. trimmed.len - 1] else trimmed;
    const r_attr = getAttr(attrs_for_lookup, "ref");
    if (r_attr == null) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        i.* = t.after_open;
        return;
    }
    const ref = r_attr.?;
    const colon = std.mem.indexOfScalar(u8, ref, ':');

    // Capture old + new column bounds; needed for filterColumn
    // colId rebasing.
    var old_tl_col: u32 = 0;
    var old_br_col: u32 = 0;
    if (colon) |c| {
        old_tl_col = parseColFromA1(ref[0..c]) orelse 0;
        old_br_col = parseColFromA1(ref[c + 1 ..]) orelse 0;
    } else {
        old_tl_col = parseColFromA1(ref) orelse 0;
        old_br_col = old_tl_col;
    }

    // Detect collapse: delete-match where every column in the
    // range is exactly the deleted column.
    var drop_entire = false;
    if (kind == .delete and old_tl_col == col_1based and old_br_col == col_1based and old_tl_col != 0) {
        drop_entire = true;
    }

    if (drop_entire) {
        if (is_self_closing) {
            i.* = t.after_open;
        } else {
            const close = std.mem.indexOfPos(u8, src, t.after_open, "</autoFilter>") orelse t.after_open;
            i.* = if (close + "</autoFilter>".len <= src.len) close + "</autoFilter>".len else t.after_open;
        }
        return;
    }

    // Compute new ref. Pass through on any malformed half.
    var new_ref_buf: [40]u8 = undefined;
    var new_ref: []const u8 = ref;
    if (colon) |c| {
        var tl_buf: [16]u8 = undefined;
        var br_buf: [16]u8 = undefined;
        const tl_new = shiftSingleA1Col(ref[0..c], col_1based, kind, &tl_buf, false) catch {
            try out.appendSlice(allocator, src[t.start..t.after_open]);
            i.* = t.after_open;
            return;
        };
        const br_new = shiftSingleA1Col(ref[c + 1 ..], col_1based, kind, &br_buf, true) catch {
            try out.appendSlice(allocator, src[t.start..t.after_open]);
            i.* = t.after_open;
            return;
        };
        new_ref = try std.fmt.bufPrint(&new_ref_buf, "{s}:{s}", .{ tl_new, br_new });
    } else {
        var b: [16]u8 = undefined;
        const shifted = shiftSingleA1Col(ref, col_1based, kind, &b, false) catch {
            try out.appendSlice(allocator, src[t.start..t.after_open]);
            i.* = t.after_open;
            return;
        };
        @memcpy(new_ref_buf[0..shifted.len], shifted);
        new_ref = new_ref_buf[0..shifted.len];
    }

    // Emit open tag with possibly-rewritten ref.
    if (std.mem.eql(u8, ref, new_ref)) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
    } else {
        try writeWithReplacedAttr(allocator, out, src, t, "<autoFilter".len, "ref", new_ref);
    }

    if (is_self_closing) {
        i.* = t.after_open;
        return;
    }

    // Open form: walk children to `</autoFilter>`. Rewrite any
    // `<filterColumn colId>` to its post-edit offset; drop
    // filterColumns whose absolute column was deleted.
    const new_tl_col: u32 = blk: {
        const new_colon = std.mem.indexOfScalar(u8, new_ref, ':');
        if (new_colon) |nc| {
            break :blk parseColFromA1(new_ref[0..nc]) orelse old_tl_col;
        }
        break :blk parseColFromA1(new_ref) orelse old_tl_col;
    };

    const close_tag = "</autoFilter>";
    const close_pos = std.mem.indexOfPos(u8, src, t.after_open, close_tag) orelse {
        // Malformed: no close tag found. Stop after the open tag.
        i.* = t.after_open;
        return;
    };

    var j: usize = t.after_open;
    while (j < close_pos) {
        const next_lt = std.mem.indexOfScalarPos(u8, src, j, '<') orelse {
            try out.appendSlice(allocator, src[j..close_pos]);
            j = close_pos;
            break;
        };
        if (next_lt >= close_pos) {
            try out.appendSlice(allocator, src[j..close_pos]);
            j = close_pos;
            break;
        }
        try out.appendSlice(allocator, src[j..next_lt]);
        j = next_lt;

        if (matchTagAt(src, j, "filterColumn")) |ft| {
            try processFilterColumnTag(
                allocator,
                out,
                src,
                ft,
                col_1based,
                kind,
                old_tl_col,
                new_tl_col,
                close_pos,
                &j,
            );
        } else if (matchTagAt(src, j, "sortState")) |st| {
            // Only the column walker consumes autoFilter children, so
            // a nested sortState never reaches the top-level dispatch
            // on this axis and has to be handled here. (The row walker
            // stops at the open tag, so its children fall through.)
            try processSortStateTagCol(allocator, out, src, st, col_1based, kind, &j);
        } else if (matchTagAt(src, j, "sortCondition")) |sc| {
            try processSortConditionTagCol(allocator, out, src, sc, col_1based, kind, &j);
        } else {
            try out.append(allocator, '<');
            j += 1;
        }
    }

    // Emit `</autoFilter>` and advance past it.
    try out.appendSlice(allocator, src[close_pos .. close_pos + close_tag.len]);
    i.* = close_pos + close_tag.len;
}

/// Rewrite a `<filterColumn colId="N">` child of `<autoFilter>`.
/// `colId` is a 0-based offset from the autoFilter's left edge.
/// Drop the filterColumn (skipping any children up to
/// `</filterColumn>`) when its absolute column is the one being
/// deleted; otherwise rewrite `colId` to the post-edit offset.
fn processFilterColumnTag(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    col_1based: u32,
    kind: RowEditKind,
    old_tl_col: u32,
    new_tl_col: u32,
    autofilter_close_pos: usize,
    j: *usize,
) !void {
    const attrs_full = src[t.start + "<filterColumn".len .. t.after_open - 1];
    const trimmed = std.mem.trimEnd(u8, attrs_full, " \t\r\n");
    const is_self_closing = trimmed.len > 0 and trimmed[trimmed.len - 1] == '/';
    const attrs_for_lookup = if (is_self_closing) trimmed[0 .. trimmed.len - 1] else trimmed;
    const id_attr = getAttr(attrs_for_lookup, "colId");

    // No colId or unparseable / out-of-range autoFilter — emit
    // verbatim, advance past the open tag only.
    if (id_attr == null or old_tl_col == 0) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        j.* = t.after_open;
        return;
    }
    const old_id = std.fmt.parseInt(u32, id_attr.?, 10) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        j.* = t.after_open;
        return;
    };
    const old_abs = old_tl_col + old_id;

    var drop = false;
    var new_abs = old_abs;
    switch (kind) {
        .insert => if (old_abs >= col_1based) {
            if (old_abs >= max_col_1based) {
                try out.appendSlice(allocator, src[t.start..t.after_open]);
                j.* = t.after_open;
                return;
            }
            new_abs = old_abs + 1;
        },
        .delete => if (old_abs == col_1based) {
            drop = true;
        } else if (old_abs > col_1based) {
            new_abs = old_abs - 1;
        },
    }

    if (drop) {
        if (is_self_closing) {
            j.* = t.after_open;
        } else {
            const close_str = "</filterColumn>";
            const close = std.mem.indexOfPos(u8, src, t.after_open, close_str) orelse autofilter_close_pos;
            const after_close = if (close + close_str.len <= autofilter_close_pos)
                close + close_str.len
            else
                close;
            j.* = after_close;
        }
        return;
    }

    // Compute new colId. new_abs should never sit before
    // new_tl_col under a well-formed edit; if it does, fall back
    // to a verbatim emit.
    if (new_tl_col == 0 or new_abs < new_tl_col) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        j.* = t.after_open;
        return;
    }
    const new_id = new_abs - new_tl_col;
    if (new_id == old_id) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
    } else {
        var idbuf: [16]u8 = undefined;
        const new_id_str = try std.fmt.bufPrint(&idbuf, "{d}", .{new_id});
        try writeWithReplacedAttr(allocator, out, src, t, "<filterColumn".len, "colId", new_id_str);
    }
    j.* = t.after_open;
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
pub fn shiftSingleA1Col(ref: []const u8, col_1based: u32, kind: RowEditKind, buf: *[16]u8, is_br_corner: bool) ![]const u8 {
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
        // `buf` is a fixed 16 bytes; `ref` comes from an attribute and
        // is arbitrary length. A valid column letter followed by an
        // over-long row number — `A99999999999999999` — parses fine and
        // then overran this copy. Refuse rather than truncate: a
        // silently shortened reference is a wrong reference.
        // Found by fuzzing, 2026-07-27.
        if (ref.len > buf.len) return error.MalformedXml;
        @memcpy(buf[0..ref.len], ref);
        return buf[0..ref.len];
    }
    var letters_buf: [8]u8 = undefined;
    const new_letters = try formatColLetters(&letters_buf, new_col - 1);
    return try std.fmt.bufPrint(buf, "{s}{s}", .{ new_letters, ref[letters_end..] });
}

pub fn parseColFromA1(ref: []const u8) ?u32 {
    var letters_end: usize = 0;
    while (letters_end < ref.len and ref[letters_end] >= 'A' and ref[letters_end] <= 'Z') letters_end += 1;
    if (letters_end == 0) return null;
    return parseColLetters(ref[0..letters_end]);
}

/// Render a 0-based col_idx as A1 letters (A=0, Z=25, AA=26, ...).
/// Caller-provided buffer; result borrows from buf.
fn formatColLetters(buf: *[8]u8, col_idx: u32) ![]const u8 {
    // M0 adapter over `zlsx_refs`. Deliberately the UNCHECKED writer:
    // a shifted index can land past XFD here, and this path has always
    // formatted it rather than erroring. The only failure is not
    // fitting the buffer — plus a `col_idx` of `maxInt(u32)`, which the
    // old `col_idx + 1` would have panicked on in Debug.
    const one_based = std.math.add(u32, col_idx, 1) catch
        return error.ColumnIndexOutOfRange;
    const len = try coords.writeColNumberLetters(buf, one_based);
    return buf[0..len];
}

/// Apply one row edit (insert or delete at `row`) to a worksheet
/// XML buffer: `<row r=>` (renumber or drop), `<c r="A1">` row
/// components, merge rects, `<dimension>`, panes, autoFilter, view
/// state, sort state, DV/CF `sqref` envelopes and `<xm:sqref>`.
/// Formula bodies, hyperlinks, drawings and tables move in their own
/// workbook-layer sweeps around this transform.
pub const RowEditKind = enum { insert, delete };

pub fn applyRowEditToWorksheet(
    allocator: Allocator,
    src: []const u8,
    row: u32,
    kind: RowEditKind,
) ![]u8 {
    var out: std.ArrayListUnmanaged(u8) = .empty;
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

        // Non-elements pass through verbatim — the column walker's
        // rule (Codex #216 r4 S3B-REL-804).
        if (i + 1 < src.len and (src[i + 1] == '!' or src[i + 1] == '?')) {
            if (workbook_xml.skipNonElement(src, i)) |end| {
                try out.appendSlice(allocator, src[i..end]);
                i = end;
                continue;
            } else |_| {}
        }

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
        } else if (matchTagAt(src, i, "pane")) |t| {
            try processPaneTagRow(allocator, &out, src, t, row, kind);
            i = t.after_open;
        } else if (matchTagAt(src, i, "autoFilter")) |t| {
            try processAutoFilterTagRow(allocator, &out, src, t, row, kind, &i);
        } else if (matchTagAt(src, i, "sheetView")) |t| {
            try processSheetViewTagRow(allocator, &out, src, t, row, kind);
            i = t.after_open;
        } else if (matchTagAt(src, i, "selection")) |t| {
            try processSelectionTagRow(allocator, &out, src, t, row, kind);
            i = t.after_open;
        } else if (matchTagAt(src, i, "pivotSelection")) |t| {
            try processPivotSelectionTag(allocator, &out, src, t, .row, row, kind);
            i = t.after_open;
        } else if (matchTagAt(src, i, "sortState")) |t| {
            try processSortStateTagRow(allocator, &out, src, t, row, kind, &i);
        } else if (matchTagAt(src, i, "sortCondition")) |t| {
            try processSortConditionTagRow(allocator, &out, src, t, row, kind, &i);
        } else if (matchTagAt(src, i, "conditionalFormatting")) |t| {
            try processSqrefListTag(allocator, &out, src, t, "<conditionalFormatting".len, .row, row, kind);
            i = t.after_open;
        } else if (matchTagAt(src, i, "dataValidation")) |t| {
            try processSqrefListTag(allocator, &out, src, t, "<dataValidation".len, .row, row, kind);
            i = t.after_open;
        } else if (matchTagAt(src, i, "xm:sqref")) |t| {
            try processXmSqrefTag(allocator, &out, src, t, .row, row, kind, &i);
        } else {
            // Some other tag; emit `<` and continue past it.
            try out.append(allocator, '<');
            i += 1;
        }
    }
    return try out.toOwnedSlice(allocator);
}

pub fn matchTagAt(src: []const u8, i: usize, tag: []const u8) ?TagOpen {
    if (i >= src.len or src[i] != '<') return null;
    const after = i + 1 + tag.len;
    // `>=`, not `>`: `after == src.len` means the input ends exactly
    // where the delimiter would be, so there is no byte to read and no
    // way for the tag to be validly terminated. The `>` form read one
    // past the end on inputs like `"<row"` — an OOB read reachable from
    // any truncated or corrupt sheet part, which is attacker-controlled
    // input on the load-modify-save path. Found by fuzzing, 2026-07-27.
    if (after >= src.len) return null;
    if (!std.mem.eql(u8, src[i + 1 .. i + 1 + tag.len], tag)) return null;
    const c = src[after];
    if (c != ' ' and c != '\t' and c != '\n' and c != '\r' and c != '/' and c != '>') return null;
    const gt = tagEnd(src, i) orelse return null;
    return .{ .start = i, .after_open = gt + 1 };
}

/// The `>` that closes the start tag opened at `lt` — a `>` inside a
/// quoted attribute value is data, not the end (Codex #206 r13
/// REL-1302). Null when the tag never closes.
pub fn tagEnd(src: []const u8, lt: usize) ?usize {
    var i = lt + 1;
    var quote: ?u8 = null;
    while (i < src.len) : (i += 1) {
        const c = src[i];
        if (quote) |q| {
            if (c == q) quote = null;
        } else if (c == '"' or c == '\'') {
            quote = c;
        } else if (c == '>') {
            return i;
        }
    }
    return null;
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
    const trimmed = std.mem.trimEnd(u8, attrs, " \t\r\n");
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
    const trimmed = std.mem.trimEnd(u8, attrs, " \t\r\n");
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
        const new_ref = shiftSingleA1Row(ref, row, kind, &new_ref_buf, false) catch {
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
    const tl_new = shiftSingleA1Row(ref[0..colon], row, kind, &tl_buf, false) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    const br_new = shiftSingleA1Row(ref[colon + 1 ..], row, kind, &br_buf, true) catch {
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

/// Rewrite `<autoFilter ref="A1:E10">…</autoFilter>` for a row
/// edit. Shifts the row halves of `ref` on the open tag; children
/// (`<filterColumn>`, `<sortState>`, …) are unaffected by row
/// edits and pass through verbatim. If the entire range collapses
/// onto the deleted row, the whole element (open form: open + body
/// + close; self-closing form: just the open tag) is dropped.
///
/// Caveat: nested `<sortState ref="…">` is not rewritten when the
/// `<autoFilter>` is sheet-bare (open form). For table-inner
/// `<autoFilter>` (delegated from `pkg/table_edit.zig`), sortState
/// is handled by table_edit's outer walker. zlsx's writer never
/// emits open-form sheet-bare autoFilter, so the remaining caveat
/// only matters for third-party files
/// that mix autoFilter sortState with row edits.
pub fn processAutoFilterTagRow(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    row: u32,
    kind: RowEditKind,
    i: *usize,
) !void {
    const attrs_full = src[t.start + "<autoFilter".len .. t.after_open - 1];
    const trimmed = std.mem.trimEnd(u8, attrs_full, " \t\r\n");
    const is_self_closing = trimmed.len > 0 and trimmed[trimmed.len - 1] == '/';
    const attrs_for_lookup = if (is_self_closing) trimmed[0 .. trimmed.len - 1] else trimmed;
    const r_attr = getAttr(attrs_for_lookup, "ref");
    if (r_attr == null) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        i.* = t.after_open;
        return;
    }
    const ref = r_attr.?;
    const colon = std.mem.indexOfScalar(u8, ref, ':');

    // Detect collapse: delete-match where every row in the range is
    // exactly the deleted row.
    var drop_entire = false;
    if (kind == .delete) {
        if (colon) |c| {
            const tl_row = parseRowFromA1(ref[0..c]) orelse 0;
            const br_row = parseRowFromA1(ref[c + 1 ..]) orelse 0;
            if (tl_row == row and br_row == row) drop_entire = true;
        } else {
            const tl_row = parseRowFromA1(ref) orelse 0;
            if (tl_row == row) drop_entire = true;
        }
    }

    if (drop_entire) {
        if (is_self_closing) {
            i.* = t.after_open;
        } else {
            const close = std.mem.indexOfPos(u8, src, t.after_open, "</autoFilter>") orelse t.after_open;
            i.* = if (close + "</autoFilter>".len <= src.len) close + "</autoFilter>".len else t.after_open;
        }
        return;
    }

    // Compute new ref. On any malformed half, pass through verbatim.
    var new_ref_buf: [40]u8 = undefined;
    var new_ref: []const u8 = ref;
    if (colon) |c| {
        var tl_buf: [16]u8 = undefined;
        var br_buf: [16]u8 = undefined;
        const tl_new = shiftSingleA1Row(ref[0..c], row, kind, &tl_buf, false) catch {
            try out.appendSlice(allocator, src[t.start..t.after_open]);
            i.* = t.after_open;
            return;
        };
        const br_new = shiftSingleA1Row(ref[c + 1 ..], row, kind, &br_buf, true) catch {
            try out.appendSlice(allocator, src[t.start..t.after_open]);
            i.* = t.after_open;
            return;
        };
        new_ref = try std.fmt.bufPrint(&new_ref_buf, "{s}:{s}", .{ tl_new, br_new });
    } else {
        var b: [16]u8 = undefined;
        const shifted = shiftSingleA1Row(ref, row, kind, &b, false) catch {
            try out.appendSlice(allocator, src[t.start..t.after_open]);
            i.* = t.after_open;
            return;
        };
        @memcpy(new_ref_buf[0..shifted.len], shifted);
        new_ref = new_ref_buf[0..shifted.len];
    }

    if (std.mem.eql(u8, ref, new_ref)) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
    } else {
        try writeWithReplacedAttr(allocator, out, src, t, "<autoFilter".len, "ref", new_ref);
    }
    i.* = t.after_open;
}

/// Parse the row component (digits) from an A1-style ref. Returns
/// null on malformed input.
pub fn parseRowFromA1(ref: []const u8) ?u32 {
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
        const new_ref = shiftSingleA1Row(ref, row, kind, &b, false) catch {
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
    const tl_new = shiftSingleA1Row(ref[0..colon], row, kind, &tl_buf, false) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    const br_new = shiftSingleA1Row(ref[colon + 1 ..], row, kind, &br_buf, true) catch {
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
pub fn shiftSingleA1Row(ref: []const u8, row: u32, kind: RowEditKind, buf: *[16]u8, is_br_corner: bool) ![]const u8 {
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
        // `buf` is a fixed 16 bytes; `ref` comes from an attribute and
        // is arbitrary length. A valid column letter followed by an
        // over-long row number — `A99999999999999999` — parses fine and
        // then overran this copy. Refuse rather than truncate: a
        // silently shortened reference is a wrong reference.
        // Found by fuzzing, 2026-07-27.
        if (ref.len > buf.len) return error.MalformedXml;
        @memcpy(buf[0..ref.len], ref);
        return buf[0..ref.len];
    }
    return try std.fmt.bufPrint(buf, "{s}{d}", .{ ref[0..letters_end], new_row });
}

/// Emit the original `<tag attrs>` with `attr_name="..."` value
/// replaced by `new_value`. `tag_name_len` is the length of the
/// tag name including the leading `<` (e.g. `"<c".len` = 2).
pub fn writeWithReplacedAttr(
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
    // Bound the search to the tag's own extent. Searching all of `src`
    // meant an unterminated attribute value — `<autoFilter ref="A1:E10>`
    // — matched a quote belonging to some later element, putting
    // `val_end_in_src` past `t.after_open` and making the final
    // `src[val_end_in_src..t.after_open]` slice backwards: a panic, not
    // a typed error, on input any corrupt or hand-edited sheet part can
    // carry. Found by fuzzing, 2026-07-27.
    const val_end_in_src = blk: {
        const found = std.mem.indexOfScalarPos(u8, src, val_start_in_src, '"') orelse
            break :blk null;
        break :blk if (found < t.after_open) found else null;
    } orelse {
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

/// Rewrite `<pane>` for a row edit: shift `ySplit` if the edit
/// lands inside the frozen-row region; shift `topLeftCell` as a
/// regular A1 ref. Only `state="frozen"` and `state="frozenSplit"`
/// are rewritten — `state="split"` (or absent, which OOXML defaults
/// to split) carries pixel offsets in xSplit/ySplit and can't be
/// meaningfully shifted by integer row/col edits.
fn processPaneTagRow(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    row: u32,
    kind: RowEditKind,
) !void {
    const attrs = src[t.start + "<pane".len .. t.after_open - 1];
    if (!paneStateIsFrozen(attrs)) return error.SplitPaneNotSupported;

    var y_buf: [16]u8 = undefined;
    var tl_buf: [16]u8 = undefined;
    var subs: [2]AttrSub = undefined;
    var n_subs: usize = 0;

    if (getAttr(attrs, "ySplit")) |y_str| {
        const y_old = std.fmt.parseInt(u32, y_str, 10) catch return error.MalformedPaneSplit;
        var y_new: ?u32 = null;
        switch (kind) {
            .insert => if (y_old != 0 and row <= y_old) {
                if (y_old >= max_row) return error.RowEditExceedsMaxRow;
                y_new = y_old + 1;
            },
            .delete => if (y_old != 0 and row <= y_old) {
                y_new = y_old - 1;
            },
        }
        if (y_new) |yn| {
            const s = try std.fmt.bufPrint(&y_buf, "{d}", .{yn});
            subs[n_subs] = .{ .name = "ySplit", .new_value = s };
            n_subs += 1;
        }
    }

    if (getAttr(attrs, "topLeftCell")) |tl| {
        const shifted = shiftSingleA1Row(tl, row, kind, &tl_buf, false) catch return error.MalformedPaneSplit;
        if (!std.mem.eql(u8, shifted, tl)) {
            subs[n_subs] = .{ .name = "topLeftCell", .new_value = shifted };
            n_subs += 1;
        }
    }

    if (n_subs == 0) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    try writeWithReplacedAttrs(allocator, out, src, t, "<pane".len, subs[0..n_subs]);
}

/// Column counterpart of `processPaneTagRow`. Shifts `xSplit` and
/// the column component of `topLeftCell`. Same `state="frozen"` /
/// `state="frozenSplit"` precondition.
fn processPaneTagCol(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    col_1based: u32,
    kind: RowEditKind,
) !void {
    const attrs = src[t.start + "<pane".len .. t.after_open - 1];
    if (!paneStateIsFrozen(attrs)) return error.SplitPaneNotSupported;

    var x_buf: [16]u8 = undefined;
    var tl_buf: [16]u8 = undefined;
    var subs: [2]AttrSub = undefined;
    var n_subs: usize = 0;

    if (getAttr(attrs, "xSplit")) |x_str| {
        const x_old = std.fmt.parseInt(u32, x_str, 10) catch return error.MalformedPaneSplit;
        var x_new: ?u32 = null;
        switch (kind) {
            .insert => if (x_old != 0 and col_1based <= x_old) {
                if (x_old >= max_col_1based) return error.ColEditExceedsMaxCol;
                x_new = x_old + 1;
            },
            .delete => if (x_old != 0 and col_1based <= x_old) {
                x_new = x_old - 1;
            },
        }
        if (x_new) |xn| {
            const s = try std.fmt.bufPrint(&x_buf, "{d}", .{xn});
            subs[n_subs] = .{ .name = "xSplit", .new_value = s };
            n_subs += 1;
        }
    }

    if (getAttr(attrs, "topLeftCell")) |tl| {
        const shifted = shiftSingleA1Col(tl, col_1based, kind, &tl_buf, false) catch return error.MalformedPaneSplit;
        if (!std.mem.eql(u8, shifted, tl)) {
            subs[n_subs] = .{ .name = "topLeftCell", .new_value = shifted };
            n_subs += 1;
        }
    }

    if (n_subs == 0) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    try writeWithReplacedAttrs(allocator, out, src, t, "<pane".len, subs[0..n_subs]);
}

fn paneStateIsFrozen(attrs: []const u8) bool {
    const s = getAttr(attrs, "state") orelse return false;
    return std.mem.eql(u8, s, "frozen") or std.mem.eql(u8, s, "frozenSplit");
}

/// Axis selector for the handlers that share one implementation
/// across row and column edits.
///
/// The older handlers in this file predate it and come as explicit
/// `…TagRow` / `…TagCol` pairs. The newer ones below keep that public
/// shape — the two dispatch loops stay branch-free — but back it with
/// a single axis-generic body, because the row and column logic for
/// these elements differs only in which A1 component moves.
pub const Axis = enum { row, col };

fn shiftHalfOnAxis(
    ref: []const u8,
    axis: Axis,
    idx_1based: u32,
    kind: RowEditKind,
    buf: *[16]u8,
    is_br_corner: bool,
) ![]const u8 {
    return switch (axis) {
        .row => shiftSingleA1Row(ref, idx_1based, kind, buf, is_br_corner),
        .col => shiftSingleA1Col(ref, idx_1based, kind, buf, is_br_corner),
    };
}

fn parseOnAxis(ref: []const u8, axis: Axis) ?u32 {
    return switch (axis) {
        .row => parseRowFromA1(ref),
        .col => parseColFromA1(ref),
    };
}

/// Shift a single A1 ref (`B5`) or an `A1:B2` range along `axis`.
///
/// Returns `null` when the range **collapses** — a delete whose target
/// is the only row/column the range spans on that axis, leaving
/// nothing for it to point at. Callers decide what a collapse means
/// for their element: `<sortState>` drops entirely, an `sqref` entry
/// is dropped from the list.
///
/// Result borrows from `out_buf`.
/// Shift `<xm:sqref>A1:A10 C3</xm:sqref>` — the range an `x14:`/`x15:`
/// extension applies to.
///
/// `<extLst>` was the last coordinate-bearing surface passing through
/// **verbatim**: this file contained no reference to it, so an
/// `x14:conditionalFormatting`, `x14:dataValidation`,
/// `x14:sparklineGroup` or `x14:ignoredErrors` kept pointing at the
/// pre-shift grid after a row/col edit. `docs/plans/refusal-audit.md`
/// named it as the one surface its method could not reach.
///
/// Handled by element name rather than by walking into `<extLst>`:
/// `xm:sqref` means the same thing — "the range this extension covers"
/// — in every extension that carries it, and it appears nowhere else.
/// Matching the leaf directly avoids tracking extension nesting, and a
/// future `x14:` element carrying `xm:sqref` is then correct for free.
///
/// Unlike an attribute, this is element *text*, so the value lives
/// between the open and close tags. Malformed input (no close tag,
/// unshiftable ref) emits the original bytes unchanged — same failure
/// posture as every other handler here.
fn processXmSqrefTag(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    axis: Axis,
    idx_1based: u32,
    kind: RowEditKind,
    i: *usize,
) !void {
    const CLOSE = "</xm:sqref>";
    const text_start = t.after_open;
    const close_at = std.mem.indexOfPos(u8, src, text_start, CLOSE) orelse {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        i.* = t.after_open;
        return;
    };
    const body = src[text_start..close_at];

    var shifted: std.ArrayListUnmanaged(u8) = .empty;
    defer shifted.deinit(allocator);
    var it = std.mem.tokenizeAny(u8, body, " \t\r\n");
    while (it.next()) |entry| {
        var ent_buf: [40]u8 = undefined;
        const s = shiftRefOrRange(entry, axis, idx_1based, kind, &ent_buf) catch {
            // Unparseable ref: emit the whole element untouched rather
            // than a partially-shifted list.
            try out.appendSlice(allocator, src[t.start .. close_at + CLOSE.len]);
            i.* = close_at + CLOSE.len;
            return;
        };
        const keep = s orelse continue; // range collapsed by a delete
        if (shifted.items.len > 0) try shifted.append(allocator, ' ');
        try shifted.appendSlice(allocator, keep);
    }

    // Every range collapsed. An empty `<xm:sqref/>` is schema-invalid,
    // and dropping the element alone would orphan its sibling rule, so
    // emit the original bytes and let the edit stand — the same
    // trade-off `<selection sqref>` makes when its list empties.
    if (shifted.items.len == 0) {
        try out.appendSlice(allocator, src[t.start .. close_at + CLOSE.len]);
        i.* = close_at + CLOSE.len;
        return;
    }

    try out.appendSlice(allocator, src[t.start..text_start]);
    try out.appendSlice(allocator, shifted.items);
    try out.appendSlice(allocator, CLOSE);
    i.* = close_at + CLOSE.len;
}

/// One `<xm:f>` formula carrier inside a worksheet's `<extLst>`:
/// a sparkline's data range or date axis (`x14:sparkline`,
/// `x14:sparklineGroup`), an `x14:cfRule` / `x14:cfvo` expression, an
/// `x14:formula1` / `x14:formula2` validation body. Byte offsets into
/// the source; `body` is the element's raw inner text (still
/// entity-encoded — the rewrite boundary decodes on the way in and
/// re-escapes on the way out), `next` is where the scan resumes.
pub const XmFormula = struct { body_start: usize, body_end: usize, next: usize };

/// Locate the `<xm:f>` element at or after `from`, or `null` when the
/// rest of `src` carries none.
///
/// This is the carrier walk behind S2's lift: `<xm:f>` holds a
/// *formula*, not a plain range like `<xm:sqref>`, so the byte
/// transform above cannot shift it — it needs the formula rewriter
/// and a sheet-name context only the workbook layer has
/// (`Workbook.rewriteAllExtensionFormulas`). As with `xm:sqref`, the
/// leaf is matched by name rather than by walking into `<extLst>`:
/// `xm:f` means "a formula" in every extension that carries it and
/// appears nowhere else, so a future `x14:` element carrying one is
/// rewritten for free.
///
/// Unlike every other handler in this file, a malformed carrier is an
/// **error**, not a pass-through: an `<xm:f>` with no `</xm:f>`, or
/// one whose text holds markup, cannot be rewritten wholly, and the
/// edit's contract is all-or-nothing — `Workbook` preflights every
/// sheet with this scan before its first mutation and refuses the
/// whole edit on `MalformedExtensionXml`. A self-closing `<xm:f/>`
/// is an empty formula and is returned with an empty body.
///
/// Prefix-literal, like every scanner in this file: the element is
/// matched as `xm:f`, not resolved through its namespace binding, so
/// a producer that binds `…/excel/2006/main` to another prefix
/// (`<m:f xmlns:m=…>`) has no carrier here and its formula passes
/// through unmaintained — the same gap `xm:sqref` (#140) and every
/// `x14:` match carry, and the one #140's guard had too (it scanned
/// for the literal `<xm:f`). Excel and LibreOffice both write `xm`;
/// namespace-aware scanning is the tracked backlog item
/// (`goal_formula.md`, M10+), not an S2 concern.
pub fn nextXmFormula(src: []const u8, from: usize) error{MalformedExtensionXml}!?XmFormula {
    const OPEN = "<xm:f";
    const CLOSE = "</xm:f>";
    // No carrier text anywhere ahead: done, whatever else the XML
    // holds. Only a sheet that does spell `<xm:f` pays for the
    // decoy-aware walk below.
    if (std.mem.indexOfPos(u8, src, from, OPEN) == null) return null;
    var i = from;
    while (std.mem.indexOfScalarPos(u8, src, i, '<')) |lt| {
        // A comment, CDATA section, PI or DOCTYPE may spell `<xm:f>`
        // without carrying one (Codex, S2 r1). Step over the whole
        // construct so a decoy is neither spliced nor refused; an
        // unterminated construct leaves the live/decoy question
        // undecidable, and this sheet does hold `<xm:f` text, so
        // refuse rather than guess.
        const past = workbook_xml.skipNonElement(src, lt) catch return error.MalformedExtensionXml;
        if (past != lt) {
            i = past;
            continue;
        }
        // `<xm:f` is a prefix of nothing else in the `xm:` namespace
        // today, but `matchTagAt` insists on a delimiter after the
        // name so that stays true by construction.
        const t = matchTagAt(src, lt, "xm:f") orelse {
            i = lt + 1;
            continue;
        };
        std.debug.assert(t.after_open >= 2);
        if (src[t.after_open - 2] == '/') {
            return .{ .body_start = t.after_open, .body_end = t.after_open, .next = t.after_open };
        }
        const close_at = std.mem.indexOfPos(u8, src, t.after_open, CLOSE) orelse
            return error.MalformedExtensionXml;
        const body = src[t.after_open..close_at];
        // ST_Formula is a simple type: element text only. A `<` in
        // the body is a nested element, a comment or a CDATA section
        // — none of which the rewriter can carry through a splice.
        if (std.mem.indexOfScalar(u8, body, '<') != null) return error.MalformedExtensionXml;
        return .{ .body_start = t.after_open, .body_end = close_at, .next = close_at + CLOSE.len };
    }
    return null;
}

fn shiftRefOrRange(
    ref: []const u8,
    axis: Axis,
    idx_1based: u32,
    kind: RowEditKind,
    out_buf: *[40]u8,
) !?[]const u8 {
    var tl_buf: [16]u8 = undefined;
    var br_buf: [16]u8 = undefined;

    if (std.mem.indexOfScalar(u8, ref, ':')) |c| {
        const tl_old = parseOnAxis(ref[0..c], axis) orelse 0;
        const br_old = parseOnAxis(ref[c + 1 ..], axis) orelse 0;
        if (kind == .delete and tl_old != 0 and tl_old == idx_1based and br_old == idx_1based) return null;
        const tl_new = try shiftHalfOnAxis(ref[0..c], axis, idx_1based, kind, &tl_buf, false);
        const br_new = try shiftHalfOnAxis(ref[c + 1 ..], axis, idx_1based, kind, &br_buf, true);
        return try std.fmt.bufPrint(out_buf, "{s}:{s}", .{ tl_new, br_new });
    }

    const old = parseOnAxis(ref, axis) orelse 0;
    if (kind == .delete and old != 0 and old == idx_1based) return null;
    const shifted = try shiftHalfOnAxis(ref, axis, idx_1based, kind, &tl_buf, false);
    @memcpy(out_buf[0..shifted.len], shifted);
    return out_buf[0..shifted.len];
}

/// Rewrite `<sheetView topLeftCell="…">` — the scroll anchor of the
/// view, distinct from the `<pane topLeftCell>` handled above.
///
/// A scroll anchor never collapses. Deleting the very row/column it
/// names does not orphan it: whatever slides into that position
/// becomes the new top-left, which is exactly the behaviour a reader
/// expects, so the index is held rather than pulled back. That is why
/// this passes `is_br_corner = false` and ignores the collapse path.
fn processSheetViewTag(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    axis: Axis,
    idx_1based: u32,
    kind: RowEditKind,
) !void {
    const attrs = src[t.start + "<sheetView".len .. t.after_open - 1];
    const tl = getAttr(attrs, "topLeftCell") orelse {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    var buf: [16]u8 = undefined;
    const shifted = shiftHalfOnAxis(tl, axis, idx_1based, kind, &buf, false) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    if (std.mem.eql(u8, shifted, tl)) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    try writeWithReplacedAttr(allocator, out, src, t, "<sheetView".len, "topLeftCell", shifted);
}

pub fn processSheetViewTagCol(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    col_1based: u32,
    kind: RowEditKind,
) !void {
    return processSheetViewTag(allocator, out, src, t, .col, col_1based, kind);
}

pub fn processSheetViewTagRow(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    row: u32,
    kind: RowEditKind,
) !void {
    return processSheetViewTag(allocator, out, src, t, .row, row, kind);
}

/// Rewrite `<selection activeCell="B5" sqref="B5 D1:D9"/>`.
///
/// `sqref` is a space-separated list of refs and ranges (ECMA-376
/// §18.3.1.78), so each entry shifts independently and an entry whose
/// range collapses is dropped from the list.
///
/// `activeCell` is an anchor, not a data reference, and gets the same
/// hold-the-index treatment as `topLeftCell`. If every `sqref` entry
/// collapses the element is kept but its `sqref` falls back to the
/// active cell (or `A1`): selection is pure view state, so any
/// schema-valid value is correct, and rewriting an attribute is safer
/// than excising an element mid-walk.
fn processSelectionTag(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    axis: Axis,
    idx_1based: u32,
    kind: RowEditKind,
) !void {
    const attrs = src[t.start + "<selection".len .. t.after_open - 1];

    var active_buf: [16]u8 = undefined;
    var new_active: ?[]const u8 = null;
    if (getAttr(attrs, "activeCell")) |ac| {
        new_active = shiftHalfOnAxis(ac, axis, idx_1based, kind, &active_buf, false) catch {
            try out.appendSlice(allocator, src[t.start..t.after_open]);
            return;
        };
    }

    const sqref = getAttr(attrs, "sqref");
    var sq_out: std.ArrayListUnmanaged(u8) = .empty;
    defer sq_out.deinit(allocator);
    if (sqref) |sq| {
        var it = std.mem.tokenizeAny(u8, sq, " \t\r\n");
        while (it.next()) |entry| {
            var ent_buf: [40]u8 = undefined;
            const shifted = shiftRefOrRange(entry, axis, idx_1based, kind, &ent_buf) catch {
                try out.appendSlice(allocator, src[t.start..t.after_open]);
                return;
            };
            const keep = shifted orelse continue; // collapsed — drop this entry
            if (sq_out.items.len > 0) try sq_out.append(allocator, ' ');
            try sq_out.appendSlice(allocator, keep);
        }
        if (sq_out.items.len == 0) {
            try sq_out.appendSlice(allocator, new_active orelse "A1");
        }
    }

    var subs: [2]AttrSub = undefined;
    var n_subs: usize = 0;
    if (new_active) |na| {
        if (!std.mem.eql(u8, na, getAttr(attrs, "activeCell").?)) {
            subs[n_subs] = .{ .name = "activeCell", .new_value = na };
            n_subs += 1;
        }
    }
    if (sqref) |sq| {
        if (!std.mem.eql(u8, sq_out.items, sq)) {
            subs[n_subs] = .{ .name = "sqref", .new_value = sq_out.items };
            n_subs += 1;
        }
    }

    if (n_subs == 0) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    try writeWithReplacedAttrs(allocator, out, src, t, "<selection".len, subs[0..n_subs]);
}

pub fn processSelectionTagCol(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    col_1based: u32,
    kind: RowEditKind,
) !void {
    return processSelectionTag(allocator, out, src, t, .col, col_1based, kind);
}

pub fn processSelectionTagRow(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    row: u32,
    kind: RowEditKind,
) !void {
    return processSelectionTag(allocator, out, src, t, .row, row, kind);
}

/// One half of an sqref area, parsed by the strict ST_Ref grammar:
/// `[$]LETTERS[[$]DIGITS]` or `[$]DIGITS` — uppercase letters within
/// `XFD`, a row within the grid with no leading zero, `$` anchors
/// admitted per component (MS-XLSX defines ST_Ref through the A1
/// grammar, anchors included — Codex #216 r4 S3B-REL-803). Null when
/// the spelling is none of those.
const SqrefHalf = struct {
    col: ?u32 = null, // 1-based
    col_anchor: bool = false,
    row: ?u32 = null, // 1-based
    row_anchor: bool = false,
};

fn parseSqrefHalf(s: []const u8) ?SqrefHalf {
    var rest = s;
    var h: SqrefHalf = .{};
    if (rest.len > 0 and rest[0] == '$') {
        h.col_anchor = true;
        rest = rest[1..];
    }
    var letters_end: usize = 0;
    while (letters_end < rest.len and rest[letters_end] >= 'A' and rest[letters_end] <= 'Z') letters_end += 1;
    if (letters_end > 0) {
        h.col = parseColLetters(rest[0..letters_end]) orelse return null;
        rest = rest[letters_end..];
    } else if (h.col_anchor) {
        // `$3`: the anchor belongs to the row component.
        h.col_anchor = false;
        h.row_anchor = true;
    }
    if (rest.len > 0 and rest[0] == '$') {
        if (h.row_anchor) return null; // `$$3`
        h.row_anchor = true;
        rest = rest[1..];
    }
    if (rest.len > 0) {
        if (rest[0] == '0') return null; // row 0 / leading zero
        var row: u32 = 0;
        for (rest) |c| {
            if (c < '0' or c > '9') return null;
            row = std.math.mul(u32, row, 10) catch return null;
            row = std.math.add(u32, row, c - '0') catch return null;
            if (row > max_row) return null;
        }
        h.row = row;
    } else if (h.row_anchor) {
        return null; // a trailing `$` anchoring nothing
    }
    if (h.col == null and h.row == null) return null;
    return h;
}

fn appendSqrefHalf(out_buf: []u8, pos: *usize, h: SqrefHalf) !void {
    if (h.col) |c1| {
        if (h.col_anchor) {
            out_buf[pos.*] = '$';
            pos.* += 1;
        }
        var letters: [8]u8 = undefined;
        const s = try formatColLetters(&letters, c1 - 1);
        @memcpy(out_buf[pos.* .. pos.* + s.len], s);
        pos.* += s.len;
    }
    if (h.row) |r| {
        if (h.row_anchor) {
            out_buf[pos.*] = '$';
            pos.* += 1;
        }
        const s = try std.fmt.bufPrint(out_buf[pos.*..], "{d}", .{r});
        pos.* += s.len;
    }
}

/// Shift one sqref area. Unlike a merge ref, an sqref area legally
/// spells whole columns (`A:A`), whole rows (`1:1`) and `$` anchors:
/// each half is parsed by the strict ST_Ref grammar (both axes
/// validated, not just the edited one), an area with no component on
/// the edited axis ABSORBS the edit — Excel keeps `A:A` as written
/// when a row is inserted — and one with the component is an interval
/// on that axis, the merge-rect rules, anchors and the cross-axis
/// components preserved. Null = the area collapsed under a delete.
/// An area the grammar refuses is an error the caller turns into a
/// whole-edit refusal (Codex #216 r3 S3B-REL-802, r4 S3B-REL-803).
fn shiftSqrefArea(
    entry: []const u8,
    axis: Axis,
    idx_1based: u32,
    kind: RowEditKind,
    out_buf: *[40]u8,
) !?[]const u8 {
    var h1: SqrefHalf = undefined;
    var h2: SqrefHalf = undefined;
    var is_range = false;
    if (std.mem.indexOfScalar(u8, entry, ':')) |c| {
        is_range = true;
        h1 = parseSqrefHalf(entry[0..c]) orelse return error.MalformedXml;
        h2 = parseSqrefHalf(entry[c + 1 ..]) orelse return error.MalformedXml;
        // ST_Ref pairs like with like: cell:cell, col:col, row:row —
        // and Excel writes corners normalized.
        if ((h1.col == null) != (h2.col == null)) return error.MalformedXml;
        if ((h1.row == null) != (h2.row == null)) return error.MalformedXml;
        if (h1.col != null and h1.col.? > h2.col.?) return error.MalformedXml;
        if (h1.row != null and h1.row.? > h2.row.?) return error.MalformedXml;
    } else {
        h1 = parseSqrefHalf(entry) orelse return error.MalformedXml;
        // A lone `A` or `3` is not an area; only a full cell stands
        // alone.
        if (h1.col == null or h1.row == null) return error.MalformedXml;
        h2 = h1;
    }

    // No component on the edited axis: the area absorbs the edit.
    const absent = switch (axis) {
        .row => h1.row == null,
        .col => h1.col == null,
    };
    if (absent) {
        if (entry.len > out_buf.len) return error.MalformedXml;
        @memcpy(out_buf[0..entry.len], entry);
        return out_buf[0..entry.len];
    }

    const v1 = switch (axis) {
        .row => h1.row.?,
        .col => h1.col.?,
    };
    const v2 = switch (axis) {
        .row => h2.row.?,
        .col => h2.col.?,
    };
    const axis_max: u32 = switch (axis) {
        .row => max_row,
        .col => max_col_1based,
    };
    if (kind == .delete and v1 == idx_1based and v2 == idx_1based) return null;
    var n1 = v1;
    var n2 = v2;
    switch (kind) {
        .insert => {
            if (v1 >= idx_1based) {
                if (v1 >= axis_max) return switch (axis) {
                    .row => error.RowEditExceedsMaxRow,
                    .col => error.ColEditExceedsMaxCol,
                };
                n1 = v1 + 1;
            }
            if (v2 >= idx_1based) {
                if (v2 >= axis_max) return switch (axis) {
                    .row => error.RowEditExceedsMaxRow,
                    .col => error.ColEditExceedsMaxCol,
                };
                n2 = v2 + 1;
            }
        },
        .delete => {
            if (v1 > idx_1based) n1 = v1 - 1;
            // The BR corner shrinks on a delete at its own line; the
            // TL corner holds (what slides up takes its place). The
            // both-equal case collapsed above, so n2 stays >= n1 >= 1.
            if (v2 > idx_1based) {
                n2 = v2 - 1;
            } else if (v2 == idx_1based and v2 > 1) {
                n2 = v2 - 1;
            }
        },
    }
    switch (axis) {
        .row => {
            h1.row = n1;
            h2.row = n2;
        },
        .col => {
            h1.col = n1;
            h2.col = n2;
        },
    }
    var pos: usize = 0;
    try appendSqrefHalf(out_buf, &pos, h1);
    if (is_range) {
        out_buf[pos] = ':';
        pos += 1;
        try appendSqrefHalf(out_buf, &pos, h2);
    }
    return out_buf[0..pos];
}

/// Rewrite the `sqref` attribute of `<conditionalFormatting>`
/// (ECMA-376 §18.3.1.18) and `<dataValidation>` (§18.3.2.32) — the
/// DATA ranges a rule block or validation applies to, a
/// space-separated list of A1 areas. Each area shifts by the
/// merge-rect interval semantics (`shiftSqrefArea`: shift below an
/// insert, grow on an insert inside, shrink on a delete inside, the
/// whole-column / whole-row spellings absorbing the cross-axis edit)
/// and a collapsed area drops from the list — Excel moves these
/// ranges with the grid exactly like a merge. The FORMULA bodies of the same
/// rules move separately, in the workbook layer's DV/CF sweep
/// (`rewriteAllValidationsAndConditionalFormats`); this handler is
/// what keeps the envelope and those bodies on one grid — before it,
/// an insert rewrote `B1>3` to `B2>3` while `sqref="B1:B4"` stayed
/// (Codex #216 r1 S3B-REL-301).
///
/// An entity-spelled value is decoded first — the shift moves what
/// the value MEANS — and an area the strict ST_Ref grammar refuses is
/// an **error**, not a pass-through: the formula sweep is about to
/// move this rule's bodies, and a frozen envelope beside moved bodies
/// is the skew this handler exists to prevent, so the whole edit
/// refuses pre-mutation, the `<xm:f>` all-or-nothing posture (Codex
/// #216 r3 S3B-REL-802). A delete that collapses EVERY area refuses
/// too (`SqrefCollapseUnsafe`): Excel deletes the rule outright, this
/// walker cannot excise an element with children mid-walk, and kept
/// bytes would silently retarget the rule (r4 S3B-REL-805).
fn processSqrefListTag(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    open_name_len: usize,
    axis: Axis,
    idx_1based: u32,
    kind: RowEditKind,
) !void {
    const attrs = src[t.start + open_name_len .. t.after_open - 1];
    const sq_raw = getAttr(attrs, "sqref") orelse {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    const sq = try store_mod.decodeXmlEntities(allocator, sq_raw);
    defer allocator.free(sq);
    var sq_out: std.ArrayListUnmanaged(u8) = .empty;
    defer sq_out.deinit(allocator);
    var any_changed = false;
    var had_areas = false;
    var it = std.mem.tokenizeAny(u8, sq, " \t\r\n");
    while (it.next()) |entry| {
        had_areas = true;
        var ent_buf: [40]u8 = undefined;
        const shifted = try shiftSqrefArea(entry, axis, idx_1based, kind, &ent_buf);
        const keep = shifted orelse {
            any_changed = true; // collapsed — drop this area
            continue;
        };
        if (!std.mem.eql(u8, keep, entry)) any_changed = true;
        if (sq_out.items.len > 0) try sq_out.append(allocator, ' ');
        try sq_out.appendSlice(allocator, keep);
    }
    // Every area collapsed: Excel deletes the rule outright, and this
    // walker cannot excise an element with children mid-walk — keeping
    // the bytes would silently retarget the rule to whatever slides
    // into its old coordinates, so the edit refuses pre-mutation, the
    // `TableCollapseUnsafe` shape (Codex #216 r4 S3B-REL-805). An
    // sqref that never held an area is inert and keeps its bytes.
    if (had_areas and sq_out.items.len == 0) return error.SqrefCollapseUnsafe;
    if (!any_changed) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    // The multi-attribute writer, not `writeWithReplacedAttr`: the
    // substring writer matches `sqref="` anywhere in the tag, so a
    // prefixed `x:sqref="…"` decoy earlier in the tag would take the
    // splice while the real attribute stayed, and a single-quoted or
    // `sqref = "…"` spelling `getAttr` reads would go unspliced while
    // the formula sweep still moved the bodies (Codex #216 r2
    // S3B-REL-701). This one matches the exact name and preserves the
    // quote style and Eq spacing.
    const subs = [_]AttrSub{.{ .name = "sqref", .new_value = sq_out.items }};
    try writeWithReplacedAttrs(allocator, out, src, t, open_name_len, &subs);
}

/// Rewrite `<pivotSelection activeRow="11" activeCol="1" previousRow="11"
/// previousCol="1" …>` (ECMA-376 §18.3.1.62): the selection state of a
/// pivot on the host sheet. The four are ABSOLUTE 0-based grid
/// coordinates and move with the grid like `activeCell` — the
/// hold-the-index rule on a delete-match, since selection is view
/// state and any in-grid value is schema-valid. `start` / `min` /
/// `max` are pivot-area item indices and `<pivotArea offset>` is
/// relative to the pivot; neither is a grid coordinate. S7a (Codex
/// #200 r1 REL-034): the pivot's own rectangle moves with the same
/// edit, so a selection left behind would point outside it.
fn processPivotSelectionTag(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    axis: Axis,
    idx_1based: u32,
    kind: RowEditKind,
) !void {
    const attrs = src[t.start + "<pivotSelection".len .. t.after_open - 1];
    const names: [2][]const u8 = switch (axis) {
        .row => .{ "activeRow", "previousRow" },
        .col => .{ "activeCol", "previousCol" },
    };
    var bufs: [2][12]u8 = undefined;
    var subs: [2]AttrSub = undefined;
    var n_subs: usize = 0;
    for (names, 0..) |name, k| {
        const raw = getAttr(attrs, name) orelse continue;
        // Character references are the value too (`&#51;` is `3`, Codex
        // #200 r3 REL-042); a value that is still not a number is left
        // as written — the walker never invents view state. A shifted
        // value is re-emitted plain: the entity spelling is not
        // preserved, the number is.
        var dec_buf: [16]u8 = undefined;
        const decoded = workbook_xml.decodeScalarAttr(&dec_buf, raw) orelse continue;
        const v = std.fmt.parseInt(u32, decoded, 10) catch continue;
        const shifted = shiftIndex0(v, idx_1based, kind) orelse continue;
        if (shifted == v) continue;
        subs[n_subs] = .{ .name = name, .new_value = try std.fmt.bufPrint(&bufs[k], "{d}", .{shifted}) };
        n_subs += 1;
    }
    if (n_subs == 0) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    try writeWithReplacedAttrs(allocator, out, src, t, "<pivotSelection".len, subs[0..n_subs]);
}

/// Hold-the-index shift of a 0-based coordinate under a 1-based edit:
/// an insert at or before it pushes it by one, a delete before it pulls
/// it by one, a delete AT it leaves it (it now names the cell that
/// moved up). Null when the push would leave `u32`.
fn shiftIndex0(v: u32, idx_1based: u32, kind: RowEditKind) ?u32 {
    // `Workbook.applySheetEditTransform` refuses index 0 before any
    // walker runs; this is the arithmetic's own statement of it.
    std.debug.assert(idx_1based >= 1);
    const edit0 = idx_1based - 1;
    return switch (kind) {
        .insert => if (v >= edit0) (std.math.add(u32, v, 1) catch return null) else v,
        .delete => if (v > edit0) v - 1 else v,
    };
}

/// Rewrite `<sortState ref="…">` and drop it whole when its range
/// collapses (the body's `<sortCondition>` children go with it).
///
/// This is the single implementation for both contexts: sheet-bare
/// `<sortState>`, `<sortState>` nested inside an open-form
/// `<autoFilter>`, and `<sortState>` inside a `<table>` — which
/// `pkg/table_edit.zig` reaches by delegating here, the same way it
/// already delegates `<autoFilter>`.
fn processSortStateTag(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    axis: Axis,
    idx_1based: u32,
    kind: RowEditKind,
    i: *usize,
) !void {
    const attrs_full = src[t.start + "<sortState".len .. t.after_open - 1];
    const trimmed = std.mem.trimEnd(u8, attrs_full, " \t\r\n");
    const is_self_closing = trimmed.len > 0 and trimmed[trimmed.len - 1] == '/';
    const attrs = if (is_self_closing) trimmed[0 .. trimmed.len - 1] else trimmed;

    const ref = getAttr(attrs, "ref") orelse {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        i.* = t.after_open;
        return;
    };

    var buf: [40]u8 = undefined;
    const shifted = shiftRefOrRange(ref, axis, idx_1based, kind, &buf) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        i.* = t.after_open;
        return;
    };

    if (shifted == null) {
        // Collapsed: drop the open tag, the body, and the close tag.
        if (is_self_closing) {
            i.* = t.after_open;
        } else {
            const close = std.mem.indexOfPos(u8, src, t.after_open, "</sortState>") orelse {
                i.* = t.after_open;
                return;
            };
            i.* = close + "</sortState>".len;
        }
        return;
    }

    if (std.mem.eql(u8, ref, shifted.?)) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
    } else {
        try writeWithReplacedAttr(allocator, out, src, t, "<sortState".len, "ref", shifted.?);
    }
    // Children are left to the caller's walker, which dispatches
    // `<sortCondition>` as a sibling tag.
    i.* = t.after_open;
}

pub fn processSortStateTagCol(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    col_1based: u32,
    kind: RowEditKind,
    i: *usize,
) !void {
    return processSortStateTag(allocator, out, src, t, .col, col_1based, kind, i);
}

pub fn processSortStateTagRow(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    row: u32,
    kind: RowEditKind,
    i: *usize,
) !void {
    return processSortStateTag(allocator, out, src, t, .row, row, kind, i);
}

/// Rewrite `<sortCondition ref="…">`, the sort key range inside a
/// `<sortState>`. It must move in step with its parent — a parent
/// that shifts while the key stays behind sorts on the wrong cells,
/// with no error surfaced anywhere.
///
/// A collapsed key range drops the condition; the parent `<sortState>`
/// survives, since its own range may still span other columns.
fn processSortConditionTag(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    axis: Axis,
    idx_1based: u32,
    kind: RowEditKind,
    i: *usize,
) !void {
    const attrs_full = src[t.start + "<sortCondition".len .. t.after_open - 1];
    const trimmed = std.mem.trimEnd(u8, attrs_full, " \t\r\n");
    const is_self_closing = trimmed.len > 0 and trimmed[trimmed.len - 1] == '/';
    const attrs = if (is_self_closing) trimmed[0 .. trimmed.len - 1] else trimmed;

    const ref = getAttr(attrs, "ref") orelse {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        i.* = t.after_open;
        return;
    };

    var buf: [40]u8 = undefined;
    const shifted = shiftRefOrRange(ref, axis, idx_1based, kind, &buf) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        i.* = t.after_open;
        return;
    };

    if (shifted == null) {
        if (is_self_closing) {
            i.* = t.after_open;
        } else {
            const close = std.mem.indexOfPos(u8, src, t.after_open, "</sortCondition>") orelse {
                i.* = t.after_open;
                return;
            };
            i.* = close + "</sortCondition>".len;
        }
        return;
    }

    if (std.mem.eql(u8, ref, shifted.?)) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
    } else {
        try writeWithReplacedAttr(allocator, out, src, t, "<sortCondition".len, "ref", shifted.?);
    }
    i.* = t.after_open;
}

pub fn processSortConditionTagCol(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    col_1based: u32,
    kind: RowEditKind,
    i: *usize,
) !void {
    return processSortConditionTag(allocator, out, src, t, .col, col_1based, kind, i);
}

pub fn processSortConditionTagRow(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    row: u32,
    kind: RowEditKind,
    i: *usize,
) !void {
    return processSortConditionTag(allocator, out, src, t, .row, row, kind, i);
}

const AttrSub = struct {
    name: []const u8,
    new_value: []const u8,
};

/// Re-emit a tag with N attribute values substituted. Walks the
/// attribute list once; for each attr whose name appears in `subs`,
/// emit the substituted value (preserving the original quote char).
/// Unknown attrs, surrounding whitespace, and the closing `/>` / `>`
/// pass through verbatim. Stable across attr reordering between
/// inputs because each attr is matched by name, not position.
fn writeWithReplacedAttrs(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    tag_name_len: usize,
    subs: []const AttrSub,
) !void {
    try out.appendSlice(allocator, src[t.start .. t.start + tag_name_len]);
    var i: usize = t.start + tag_name_len;
    const end = t.after_open;
    while (i < end) {
        const ws_start = i;
        while (i < end and (src[i] == ' ' or src[i] == '\t' or src[i] == '\n' or src[i] == '\r')) i += 1;
        try out.appendSlice(allocator, src[ws_start..i]);
        if (i >= end) break;
        // Reached the closing `>` or self-close `/>` — emit verbatim
        // and stop.
        if (src[i] == '/' or src[i] == '>') {
            try out.appendSlice(allocator, src[i..end]);
            return;
        }
        const name_start = i;
        while (i < end and src[i] != '=' and !isXmlWs(src[i]) and src[i] != '/' and src[i] != '>') i += 1;
        const name = src[name_start..i];
        // XML §3.1 `Eq ::= S? '=' S?`: whitespace on either side of the
        // `=` is legal, and `getAttr` reads through it — so the writer
        // must too, or a substitution computed for `activeRow = "11"`
        // re-emits the old value (Codex #200 r2 REL-039). The spacing
        // is preserved verbatim; only the value changes.
        var j = i;
        while (j < end and isXmlWs(src[j])) j += 1;
        if (j >= end or src[j] != '=') {
            try out.appendSlice(allocator, name);
            continue;
        }
        j += 1;
        while (j < end and isXmlWs(src[j])) j += 1;
        if (j >= end or (src[j] != '"' and src[j] != '\'')) {
            try out.appendSlice(allocator, src[name_start..j]);
            i = j;
            continue;
        }
        const quote = src[j];
        const val_start = j + 1;
        i = val_start;
        while (i < end and src[i] != quote) i += 1;
        if (i < end) i += 1;
        var replacement: ?[]const u8 = null;
        for (subs) |sub| {
            if (std.mem.eql(u8, name, sub.name)) {
                replacement = sub.new_value;
                break;
            }
        }
        if (replacement) |v| {
            try out.appendSlice(allocator, src[name_start..val_start]);
            try out.appendSlice(allocator, v);
            try out.append(allocator, quote);
        } else {
            try out.appendSlice(allocator, src[name_start..i]);
        }
    }
}

fn isXmlWs(c: u8) bool {
    return c == ' ' or c == '\t' or c == '\n' or c == '\r';
}

/// Parse uppercase A-Z letters as a 1-based Excel column index
/// (A=1, B=2, ..., Z=26, AA=27, ..., XFD=16384). Returns null on
/// empty input or anything past `max_col_1based`.
pub fn parseColLetters(s: []const u8) ?u32 {
    // M0 adapter over `zlsx_refs`. Policy preserved exactly: uppercase
    // only (Excel-authored XML), grid-bounded, 1-based result.
    return coords.parseColNumber(s, .{ .case = .upper_only }) catch null;
}

/// Render `col_idx` (0-based) as A, B, ..., Z, AA, AB, ... into `buf`.
/// Capacity 8 is more than enough (Excel max is XFD = 3 letters).
pub fn colLetterEditor(buf: []u8, col_idx: u32) []u8 {
    // Unchecked writer: this path takes an already-shifted 0-based
    // index and has never bounds-checked it against the grid. The
    // assert — not any caller's discipline — is what makes the error
    // branch unreachable: 8 bytes covers the longest run any `u32` can
    // produce. (No caller exists today; it is public API surface.)
    std.debug.assert(buf.len >= 8);
    const n = coords.writeColNumberLetters(buf, col_idx + 1) catch unreachable;
    return buf[0..n];
}

// ---------------------------------------------------------------------------
// autoFilter rewriter tests (refusal lift; see docs/plans/refusal-audit.md).
// ---------------------------------------------------------------------------

const testing = std.testing;

fn wrapSheet(allocator: Allocator, body: []const u8) ![]u8 {
    const head = "<worksheet><sheetData><row r=\"1\"><c r=\"A1\"/></row></sheetData>";
    const tail = "</worksheet>";
    var buf = try allocator.alloc(u8, head.len + body.len + tail.len);
    @memcpy(buf[0..head.len], head);
    @memcpy(buf[head.len .. head.len + body.len], body);
    @memcpy(buf[head.len + body.len ..], tail);
    return buf;
}

test "autoFilter row insert: shifts ref row halves" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<autoFilter ref=\"B2:D5\"/>");
    defer a.free(src);
    const out = try applyRowEditToWorksheet(a, src, 3, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "ref=\"B2:D6\"") != null);
}

test "autoFilter row delete inside range: shifts BR row only" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<autoFilter ref=\"B2:D5\"/>");
    defer a.free(src);
    const out = try applyRowEditToWorksheet(a, src, 3, .delete);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "ref=\"B2:D4\"") != null);
}

test "autoFilter row delete that collapses entire range: drops self-closing tag" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<autoFilter ref=\"A5:D5\"/>");
    defer a.free(src);
    const out = try applyRowEditToWorksheet(a, src, 5, .delete);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<autoFilter") == null);
}

test "autoFilter row delete that collapses entire range: drops open form + body" {
    const a = testing.allocator;
    const src = try wrapSheet(
        a,
        "<autoFilter ref=\"A5:D5\"><filterColumn colId=\"0\"><filters><filter val=\"x\"/></filters></filterColumn></autoFilter>",
    );
    defer a.free(src);
    const out = try applyRowEditToWorksheet(a, src, 5, .delete);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<autoFilter") == null);
    try testing.expect(std.mem.indexOf(u8, out, "<filterColumn") == null);
    try testing.expect(std.mem.indexOf(u8, out, "</autoFilter>") == null);
}

test "autoFilter col insert before range: shifts both halves right" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<autoFilter ref=\"D2:G5\"/>");
    defer a.free(src);
    const out = try applyColEditToWorksheet(a, src, 2, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "ref=\"E2:H5\"") != null);
}

test "autoFilter col delete inside range: BR shrinks by one" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<autoFilter ref=\"B2:E5\"/>");
    defer a.free(src);
    const out = try applyColEditToWorksheet(a, src, 4, .delete);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "ref=\"B2:D5\"") != null);
}

test "autoFilter col delete on single-column range: drops entire tag" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<autoFilter ref=\"C1:C10\"/>");
    defer a.free(src);
    const out = try applyColEditToWorksheet(a, src, 3, .delete);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<autoFilter") == null);
}

test "autoFilter open form: filterColumn colId rebases on col delete inside range" {
    const a = testing.allocator;
    // Range B2:E5 (cols B,C,D,E). Delete col B → range B2:D5
    // (cols B,C,D — original C/D/E shifted left). filterColumn
    // colId=0 (was B) is dropped; colId=1 (was C) → 0; colId=2
    // (was D) → 1; colId=3 (was E) → 2.
    const body =
        "<autoFilter ref=\"B2:E5\">" ++
        "<filterColumn colId=\"0\"><filters><filter val=\"x\"/></filters></filterColumn>" ++
        "<filterColumn colId=\"1\"><filters><filter val=\"y\"/></filters></filterColumn>" ++
        "<filterColumn colId=\"2\"><filters><filter val=\"z\"/></filters></filterColumn>" ++
        "<filterColumn colId=\"3\"><filters><filter val=\"w\"/></filters></filterColumn>" ++
        "</autoFilter>";
    const src = try wrapSheet(a, body);
    defer a.free(src);
    const out = try applyColEditToWorksheet(a, src, 2, .delete);
    defer a.free(out);
    // The col-B filterColumn (val=x) was dropped entirely.
    try testing.expect(std.mem.indexOf(u8, out, "val=\"x\"") == null);
    // Range shrinks to B2:D5.
    try testing.expect(std.mem.indexOf(u8, out, "ref=\"B2:D5\"") != null);
    // Surviving filterColumns rebased: y→0, z→1, w→2.
    const y_at = std.mem.indexOf(u8, out, "val=\"y\"").?;
    const z_at = std.mem.indexOf(u8, out, "val=\"z\"").?;
    const w_at = std.mem.indexOf(u8, out, "val=\"w\"").?;
    try testing.expect(std.mem.indexOf(u8, out[0..y_at], "colId=\"0\"") != null);
    try testing.expect(std.mem.indexOf(u8, out[0..z_at], "colId=\"1\"") != null);
    try testing.expect(std.mem.indexOf(u8, out[0..w_at], "colId=\"2\"") != null);
}

test "autoFilter open form: filterColumn colId rebases on col insert inside range" {
    const a = testing.allocator;
    // Range B2:E5. Insert col C (=3) → range B2:F5; new col is at
    // colId=1 (no filterColumn there). Old colIds: 0 (B) stays 0;
    // 1 (C) → 2; 2 (D) → 3; 3 (E) → 4.
    const body =
        "<autoFilter ref=\"B2:E5\">" ++
        "<filterColumn colId=\"0\"><filters><filter val=\"x\"/></filters></filterColumn>" ++
        "<filterColumn colId=\"1\"><filters><filter val=\"y\"/></filters></filterColumn>" ++
        "<filterColumn colId=\"2\"><filters><filter val=\"z\"/></filters></filterColumn>" ++
        "<filterColumn colId=\"3\"><filters><filter val=\"w\"/></filters></filterColumn>" ++
        "</autoFilter>";
    const src = try wrapSheet(a, body);
    defer a.free(src);
    const out = try applyColEditToWorksheet(a, src, 3, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "ref=\"B2:F5\"") != null);
    const x_at = std.mem.indexOf(u8, out, "val=\"x\"").?;
    const y_at = std.mem.indexOf(u8, out, "val=\"y\"").?;
    const z_at = std.mem.indexOf(u8, out, "val=\"z\"").?;
    const w_at = std.mem.indexOf(u8, out, "val=\"w\"").?;
    try testing.expect(std.mem.indexOf(u8, out[0..x_at], "colId=\"0\"") != null);
    try testing.expect(std.mem.indexOf(u8, out[0..y_at], "colId=\"2\"") != null);
    try testing.expect(std.mem.indexOf(u8, out[0..z_at], "colId=\"3\"") != null);
    try testing.expect(std.mem.indexOf(u8, out[0..w_at], "colId=\"4\"") != null);
}

test "autoFilter col insert past range: ref unchanged" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<autoFilter ref=\"B2:D5\"/>");
    defer a.free(src);
    const out = try applyColEditToWorksheet(a, src, 10, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "ref=\"B2:D5\"") != null);
}

test "autoFilter passes through when ref attribute is missing" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<autoFilter/>");
    defer a.free(src);
    const out = try applyRowEditToWorksheet(a, src, 1, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<autoFilter/>") != null);
}

// ─── iter-sv-1: sheetView / selection / sortState coordinate attrs ──
//
// Each of these carries a coordinate that row/col edits used to leave
// behind. None of them surfaced an error when they went stale — the
// workbook just quietly pointed somewhere wrong, which is the failure
// mode the library exists to prevent.

test "sheetView topLeftCell shifts on row insert" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<sheetViews><sheetView topLeftCell=\"B5\" workbookViewId=\"0\"/></sheetViews>");
    defer a.free(src);
    const out = try applyRowEditToWorksheet(a, src, 2, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "topLeftCell=\"B6\"") != null);
}

test "sheetView topLeftCell shifts on col insert" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<sheetViews><sheetView topLeftCell=\"C1\" workbookViewId=\"0\"/></sheetViews>");
    defer a.free(src);
    const out = try applyColEditToWorksheet(a, src, 2, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "topLeftCell=\"D1\"") != null);
}

test "sheetView topLeftCell holds its index when its own row is deleted" {
    const a = testing.allocator;
    // A scroll anchor does not collapse: whatever slides into row 5
    // becomes the new top-left, so the index stays put.
    const src = try wrapSheet(a, "<sheetViews><sheetView topLeftCell=\"B5\"/></sheetViews>");
    defer a.free(src);
    const out = try applyRowEditToWorksheet(a, src, 5, .delete);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "topLeftCell=\"B5\"") != null);
}

test "sheetViews container is not mistaken for sheetView" {
    const a = testing.allocator;
    // `<sheetViews>` shares a prefix with `<sheetView>`; matchTagAt's
    // delimiter check is what keeps them apart. Guard it explicitly.
    const src = try wrapSheet(a, "<sheetViews><sheetView topLeftCell=\"A9\"/></sheetViews>");
    defer a.free(src);
    const out = try applyRowEditToWorksheet(a, src, 1, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<sheetViews>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "</sheetViews>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "topLeftCell=\"A10\"") != null);
}

test "pivotSelection: the four absolute coordinates move with the grid, the item indices do not" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<sheetView workbookViewId=\"0\"><selection activeCell=\"C12\" sqref=\"C12\"/><pivotSelection pane=\"bottomRight\" showHeader=\"1\" axis=\"axisRow\" dimension=\"0\" start=\"2\" min=\"2\" max=\"3\" activeRow=\"11\" activeCol=\"1\" previousRow=\"11\" previousCol=\"1\" click=\"1\" r:id=\"rIdPT\"><pivotArea dataOnly=\"0\" labelOnly=\"1\" outline=\"0\" fieldPosition=\"0\"><references count=\"1\"><reference field=\"0\" count=\"1\"><x v=\"1\"/></reference></references></pivotArea></pivotSelection></sheetView>");
    defer a.free(src);
    // Insert row 1: rows 11 → 12; columns untouched.
    {
        const out = try applyRowEditToWorksheet(a, src, 1, .insert);
        defer a.free(out);
        try testing.expect(std.mem.indexOf(u8, out, "activeRow=\"12\" activeCol=\"1\" previousRow=\"12\" previousCol=\"1\"") != null);
        try testing.expect(std.mem.indexOf(u8, out, "start=\"2\" min=\"2\" max=\"3\"") != null);
        try testing.expect(std.mem.indexOf(u8, out, "activeCell=\"C13\"") != null);
    }
    // Delete row 12 (the selected one): hold the index.
    {
        const out = try applyRowEditToWorksheet(a, src, 12, .delete);
        defer a.free(out);
        try testing.expect(std.mem.indexOf(u8, out, "activeRow=\"11\" activeCol=\"1\" previousRow=\"11\" previousCol=\"1\"") != null);
    }
    // Delete row 3 (above): 11 → 10.
    {
        const out = try applyRowEditToWorksheet(a, src, 3, .delete);
        defer a.free(out);
        try testing.expect(std.mem.indexOf(u8, out, "activeRow=\"10\" activeCol=\"1\" previousRow=\"10\" previousCol=\"1\"") != null);
    }
    // Insert column B (2): cols 1 → 2; rows untouched.
    {
        const out = try applyColEditToWorksheet(a, src, 2, .insert);
        defer a.free(out);
        try testing.expect(std.mem.indexOf(u8, out, "activeRow=\"11\" activeCol=\"2\" previousRow=\"11\" previousCol=\"2\"") != null);
    }
    // Insert row 20 (below): byte-identical.
    {
        const out = try applyRowEditToWorksheet(a, src, 20, .insert);
        defer a.free(out);
        try testing.expectEqualStrings(src, out);
    }
}

test "pivotSelection: whitespace around `=`, single quotes and an open tag are rewritten in place" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<sheetView><selection activeCell = \"C12\" sqref = \"C12\"/><pivotSelection activeRow = \"11\" previousRow =\t'11' activeCol= \"1\" r:id=\"rIdPT\"><pivotArea/></pivotSelection></sheetView>");
    defer a.free(src);
    const out = try applyRowEditToWorksheet(a, src, 2, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<pivotSelection activeRow = \"12\" previousRow =\t'12' activeCol= \"1\" r:id=\"rIdPT\">") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<selection activeCell = \"C13\" sqref = \"C13\"/>") != null);
}

test "pivotSelection: character references spell the coordinate too" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<sheetView><pivotSelection activeRow=\"&#51;\" activeCol=\"&#x31;\" previousRow=\"3\" previousCol=\"1\" r:id=\"rIdPT\"/></sheetView>");
    defer a.free(src);
    {
        const out = try applyRowEditToWorksheet(a, src, 2, .insert);
        defer a.free(out);
        try testing.expect(std.mem.indexOf(u8, out, "activeRow=\"4\" activeCol=\"&#x31;\" previousRow=\"4\" previousCol=\"1\"") != null);
    }
    {
        const out = try applyColEditToWorksheet(a, src, 1, .insert);
        defer a.free(out);
        try testing.expect(std.mem.indexOf(u8, out, "activeRow=\"&#51;\" activeCol=\"2\" previousRow=\"3\" previousCol=\"2\"") != null);
    }
}

test "pivotSelection: a coordinate that is not a number is left as written" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<sheetView><pivotSelection activeRow=\"x\" activeCol=\"1\" previousRow=\"30\" r:id=\"rIdPT\"/></sheetView>");
    defer a.free(src);
    // Row 20 is below the fixture's data, so only the selection moves:
    // the numeric coordinate shifts, the unreadable one stays.
    const out = try applyRowEditToWorksheet(a, src, 20, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "activeRow=\"x\" activeCol=\"1\" previousRow=\"31\"") != null);
}

test "selection activeCell and sqref both shift" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<selection activeCell=\"B5\" sqref=\"B5\"/>");
    defer a.free(src);
    const out = try applyRowEditToWorksheet(a, src, 2, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "activeCell=\"B6\"") != null);
    try testing.expect(std.mem.indexOf(u8, out, "sqref=\"B6\"") != null);
}

test "selection sqref shifts every entry of a multi-range list" {
    const a = testing.allocator;
    // sqref is a space-separated list (ECMA-376 18.3.1.78).
    const src = try wrapSheet(a, "<selection activeCell=\"A2\" sqref=\"A2 C4:D6 F8\"/>");
    defer a.free(src);
    const out = try applyRowEditToWorksheet(a, src, 1, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "sqref=\"A3 C5:D7 F9\"") != null);
}

test "conditionalFormatting sqref shifts with the grid, children untouched (S3B-REL-301)" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<conditionalFormatting sqref=\"B1:B4\"><cfRule type=\"expression\" priority=\"1\"><formula>B1&gt;3</formula></cfRule></conditionalFormatting>");
    defer a.free(src);
    // Insert above: the whole range slides down — the same move the
    // formula sweep gives its `<formula>` bodies, so envelope and body
    // stay on one grid.
    const out = try applyRowEditToWorksheet(a, src, 1, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<conditionalFormatting sqref=\"B2:B5\">") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<formula>B1&gt;3</formula>") != null);
}

test "conditionalFormatting sqref grows on an insert inside, shrinks on a delete inside" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<conditionalFormatting sqref=\"A2:A5\"><cfRule type=\"containsBlanks\" priority=\"1\"/></conditionalFormatting>");
    defer a.free(src);
    const grown = try applyRowEditToWorksheet(a, src, 3, .insert);
    defer a.free(grown);
    try testing.expect(std.mem.indexOf(u8, grown, "sqref=\"A2:A6\"") != null);
    const shrunk = try applyRowEditToWorksheet(a, src, 3, .delete);
    defer a.free(shrunk);
    try testing.expect(std.mem.indexOf(u8, shrunk, "sqref=\"A2:A4\"") != null);
}

test "conditionalFormatting sqref drops a collapsed area; ALL collapsed refuses the edit (S3B-REL-805)" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<conditionalFormatting sqref=\"A1 C3:D3 F8\"><cfRule type=\"containsBlanks\" priority=\"1\"/></conditionalFormatting>");
    defer a.free(src);
    const out = try applyRowEditToWorksheet(a, src, 3, .delete);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "sqref=\"A1 F7\"") != null);

    // Every area collapsed: Excel deletes the rule outright; keeping
    // the bytes would retarget it to whatever slides into row 3, so
    // the edit refuses pre-mutation — the TableCollapseUnsafe shape.
    const src2 = try wrapSheet(a, "<conditionalFormatting sqref=\"C3:D3\"><cfRule type=\"containsBlanks\" priority=\"1\"/></conditionalFormatting>");
    defer a.free(src2);
    try testing.expectError(error.SqrefCollapseUnsafe, applyRowEditToWorksheet(a, src2, 3, .delete));
    const src3 = try wrapSheet(a, "<dataValidation type=\"list\" sqref=\"D2\"><formula1>\"a\"</formula1></dataValidation>");
    defer a.free(src3);
    try testing.expectError(error.SqrefCollapseUnsafe, applyColEditToWorksheet(a, src3, 4, .delete));
    // An sqref that never held an area is inert, not a collapse.
    const src4 = try wrapSheet(a, "<conditionalFormatting sqref=\"\"><cfRule type=\"containsBlanks\" priority=\"1\"/></conditionalFormatting>");
    defer a.free(src4);
    const out4 = try applyRowEditToWorksheet(a, src4, 3, .delete);
    defer a.free(out4);
    try testing.expect(std.mem.indexOf(u8, out4, "sqref=\"\"") != null);
}

test "sqref $ anchors parse, shift and survive rendering (S3B-REL-803)" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<conditionalFormatting sqref=\"$B$1:B4 $A:$B 2:$3\"><cfRule type=\"containsBlanks\" priority=\"1\"/></conditionalFormatting>");
    defer a.free(src);
    // Row insert at 1: the cell range shifts rows with anchors kept,
    // the whole-column pair absorbs, the whole-row pair shifts.
    const out = try applyRowEditToWorksheet(a, src, 1, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "sqref=\"$B$2:B5 $A:$B 3:$4\"") != null);
    // Column insert at 1: the cell range and the column pair shift
    // columns, the whole-row pair absorbs.
    const cout = try applyColEditToWorksheet(a, src, 1, .insert);
    defer a.free(cout);
    try testing.expect(std.mem.indexOf(u8, cout, "sqref=\"$C$1:C4 $B:$C 2:$3\"") != null);
}

test "sqref validates BOTH axes strictly; the grammar refuses out-of-grid spellings (S3B-REL-803)" {
    const a = testing.allocator;
    // XFE is past the last column — a row edit never touches the
    // column component, but a lax parse would bless it.
    const bad1 = try wrapSheet(a, "<conditionalFormatting sqref=\"XFE1\"><cfRule type=\"containsBlanks\" priority=\"1\"/></conditionalFormatting>");
    defer a.free(bad1);
    try testing.expectError(error.MalformedXml, applyRowEditToWorksheet(a, bad1, 9, .insert));
    // Row 0 and a row past the grid refuse under a COLUMN edit too.
    const bad2 = try wrapSheet(a, "<dataValidation sqref=\"A0\"><formula1>1</formula1></dataValidation>");
    defer a.free(bad2);
    try testing.expectError(error.MalformedXml, applyColEditToWorksheet(a, bad2, 9, .insert));
    const bad3 = try wrapSheet(a, "<dataValidation sqref=\"A1048577\"><formula1>1</formula1></dataValidation>");
    defer a.free(bad3);
    try testing.expectError(error.MalformedXml, applyColEditToWorksheet(a, bad3, 9, .insert));
    // A lone column or row half is not an area.
    const bad4 = try wrapSheet(a, "<conditionalFormatting sqref=\"A\"><cfRule type=\"containsBlanks\" priority=\"1\"/></conditionalFormatting>");
    defer a.free(bad4);
    try testing.expectError(error.MalformedXml, applyRowEditToWorksheet(a, bad4, 1, .insert));
}

test "non-elements pass through verbatim: a decoy tag inside a comment is prose (S3B-REL-804)" {
    const a = testing.allocator;
    // The comment carries a garbage sqref AND a shiftable mergeCell —
    // neither may be dispatched: no refusal, no rewrite inside the
    // comment; the LIVE siblings still shift.
    const src = try wrapSheet(a, "<!-- <conditionalFormatting sqref=\"NOT-A-REF\"> <mergeCell ref=\"B1:B4\"/> --><conditionalFormatting sqref=\"B1:B4\"><cfRule type=\"containsBlanks\" priority=\"1\"/></conditionalFormatting>");
    defer a.free(src);
    const out = try applyRowEditToWorksheet(a, src, 1, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<!-- <conditionalFormatting sqref=\"NOT-A-REF\"> <mergeCell ref=\"B1:B4\"/> -->") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<conditionalFormatting sqref=\"B2:B5\">") != null);

    // CDATA on the column walker.
    const csrc = try wrapSheet(a, "<x><![CDATA[<dataValidation sqref=\"ZZZ\">]]></x><dataValidation sqref=\"B2\"><formula1>1</formula1></dataValidation>");
    defer a.free(csrc);
    const cout = try applyColEditToWorksheet(a, csrc, 1, .insert);
    defer a.free(cout);
    try testing.expect(std.mem.indexOf(u8, cout, "<![CDATA[<dataValidation sqref=\"ZZZ\">]]>") != null);
    try testing.expect(std.mem.indexOf(u8, cout, "sqref=\"C2\"") != null);
}

test "dataValidation sqref shifts; the dataValidations wrapper and a decoy name do not" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<dataValidations count=\"1\"><dataValidation type=\"list\" sqref=\"B2:B10\"><formula1>\"a,b\"</formula1></dataValidation></dataValidations>");
    defer a.free(src);
    const out = try applyRowEditToWorksheet(a, src, 1, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<dataValidation type=\"list\" sqref=\"B3:B11\">") != null);
    // The plural wrapper's own attributes never carry coordinates.
    try testing.expect(std.mem.indexOf(u8, out, "<dataValidations count=\"1\">") != null);
}

test "conditionalFormatting sqref shifts on the column axis too" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<conditionalFormatting sqref=\"B1:B4 D2\"><cfRule type=\"containsBlanks\" priority=\"1\"/></conditionalFormatting>");
    defer a.free(src);
    const out = try applyColEditToWorksheet(a, src, 1, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "sqref=\"C1:C4 E2\"") != null);
    const dropped = try applyColEditToWorksheet(a, src, 4, .delete);
    defer a.free(dropped);
    try testing.expect(std.mem.indexOf(u8, dropped, "sqref=\"B1:B4\"") != null);
}

test "conditionalFormatting sqref splices by exact attribute name, either quote style (S3B-REL-701)" {
    const a = testing.allocator;
    // A prefixed `x:sqref` decoy BEFORE the real attribute: the
    // substring writer matched `sqref="` inside it and overwrote the
    // wrong value. The exact-name writer leaves it alone.
    const src = try wrapSheet(a, "<conditionalFormatting x:sqref=\"Z9\" sqref=\"B1:B4\"><cfRule type=\"containsBlanks\" priority=\"1\"/></conditionalFormatting>");
    defer a.free(src);
    const out = try applyRowEditToWorksheet(a, src, 1, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "x:sqref=\"Z9\"") != null);
    try testing.expect(std.mem.indexOf(u8, out, " sqref=\"B2:B5\"") != null);

    // Single quotes and Eq whitespace are legal XML `getAttr` reads —
    // the writer must move what the reader read, or the envelope goes
    // stale while the formula sweep moves the bodies.
    const src2 = try wrapSheet(a, "<conditionalFormatting sqref='B1:B4'><cfRule type=\"containsBlanks\" priority=\"1\"/></conditionalFormatting>");
    defer a.free(src2);
    const out2 = try applyRowEditToWorksheet(a, src2, 1, .insert);
    defer a.free(out2);
    try testing.expect(std.mem.indexOf(u8, out2, "sqref='B2:B5'") != null);

    const src3 = try wrapSheet(a, "<dataValidation type=\"list\" sqref = \"B1:B4\"><formula1>\"a\"</formula1></dataValidation>");
    defer a.free(src3);
    const out3 = try applyRowEditToWorksheet(a, src3, 1, .insert);
    defer a.free(out3);
    try testing.expect(std.mem.indexOf(u8, out3, "sqref = \"B2:B5\"") != null);
}

test "sqref whole-column and whole-row areas: absorb across the axis, shift along it (S3B-REL-802)" {
    const a = testing.allocator;
    // `A:A` has no row component: a row edit is absorbed by the area
    // — and must NOT freeze its siblings, which shift normally.
    const src = try wrapSheet(a, "<conditionalFormatting sqref=\"A:A B1:B4\"><cfRule type=\"containsBlanks\" priority=\"1\"/></conditionalFormatting>");
    defer a.free(src);
    const out = try applyRowEditToWorksheet(a, src, 1, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "sqref=\"A:A B2:B5\"") != null);

    // Along its own axis a whole-column area is an interval: insert
    // at A shifts A:B to B:C; delete of the only column collapses.
    const csrc = try wrapSheet(a, "<conditionalFormatting sqref=\"A:B D2\"><cfRule type=\"containsBlanks\" priority=\"1\"/></conditionalFormatting>");
    defer a.free(csrc);
    const cout = try applyColEditToWorksheet(a, csrc, 1, .insert);
    defer a.free(cout);
    try testing.expect(std.mem.indexOf(u8, cout, "sqref=\"B:C E2\"") != null);

    // Whole rows mirror: `1:1` shifts under a row insert above it,
    // absorbs a column edit, collapses when its only row deletes.
    const rsrc = try wrapSheet(a, "<dataValidation type=\"list\" sqref=\"2:3 A9\"><formula1>\"a\"</formula1></dataValidation>");
    defer a.free(rsrc);
    const rout = try applyRowEditToWorksheet(a, rsrc, 1, .insert);
    defer a.free(rout);
    try testing.expect(std.mem.indexOf(u8, rout, "sqref=\"3:4 A10\"") != null);
    const rout2 = try applyColEditToWorksheet(a, rsrc, 1, .insert);
    defer a.free(rout2);
    try testing.expect(std.mem.indexOf(u8, rout2, "sqref=\"2:3 B9\"") != null);
    const rsrc2 = try wrapSheet(a, "<dataValidation type=\"list\" sqref=\"3:3 A9\"><formula1>\"a\"</formula1></dataValidation>");
    defer a.free(rsrc2);
    const rout3 = try applyRowEditToWorksheet(a, rsrc2, 3, .delete);
    defer a.free(rout3);
    try testing.expect(std.mem.indexOf(u8, rout3, "sqref=\"A8\"") != null);
}

test "sqref decodes entities before shifting; garbage refuses the edit whole (S3B-REL-802)" {
    const a = testing.allocator;
    // `B1&#58;B4` MEANS B1:B4 — the shift moves the meaning and emits
    // the plain spelling.
    const src = try wrapSheet(a, "<conditionalFormatting sqref=\"B1&#58;B4\"><cfRule type=\"containsBlanks\" priority=\"1\"/></conditionalFormatting>");
    defer a.free(src);
    const out = try applyRowEditToWorksheet(a, src, 1, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "sqref=\"B2:B5\"") != null);

    // An unchanged entity spelling keeps its bytes: the edit below
    // the range is a no-op and nothing is respelled.
    const out2 = try applyRowEditToWorksheet(a, src, 9, .insert);
    defer a.free(out2);
    try testing.expect(std.mem.indexOf(u8, out2, "sqref=\"B1&#58;B4\"") != null);

    // An area that parses as no sqref form refuses the whole edit —
    // the formula sweep is about to move this rule's bodies, and a
    // frozen envelope beside moved bodies is the skew (the <xm:f>
    // all-or-nothing posture).
    const bad = try wrapSheet(a, "<conditionalFormatting sqref=\"NOT-A-REF B1:B4\"><cfRule type=\"containsBlanks\" priority=\"1\"/></conditionalFormatting>");
    defer a.free(bad);
    try testing.expectError(error.MalformedXml, applyRowEditToWorksheet(a, bad, 1, .insert));
}

test "selection sqref drops only the entry whose range collapses" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<selection activeCell=\"A1\" sqref=\"A1 C3:D3 F8\"/>");
    defer a.free(src);
    // Row 3 delete collapses C3:D3; the siblings survive and shift.
    const out = try applyRowEditToWorksheet(a, src, 3, .delete);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "sqref=\"A1 F7\"") != null);
}

test "selection falls back to activeCell when every sqref entry collapses" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<selection activeCell=\"C3\" sqref=\"C3:D3\"/>");
    defer a.free(src);
    const out = try applyRowEditToWorksheet(a, src, 3, .delete);
    defer a.free(out);
    // Selection is view state, so any schema-valid sqref is correct;
    // rewriting the attribute beats excising the element mid-walk.
    try testing.expect(std.mem.indexOf(u8, out, "sqref=\"C3\"") != null);
}

test "sheet-bare sortState ref shifts" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<sortState ref=\"A2:D10\"/>");
    defer a.free(src);
    const out = try applyRowEditToWorksheet(a, src, 1, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "ref=\"A3:D11\"") != null);
}

test "sheet-bare sortState drops with its body on range collapse" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<sortState ref=\"A4:D4\"><sortCondition ref=\"A4:A4\"/></sortState>");
    defer a.free(src);
    const out = try applyRowEditToWorksheet(a, src, 4, .delete);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "sortState") == null);
    try testing.expect(std.mem.indexOf(u8, out, "sortCondition") == null);
}

test "sortCondition shifts in step with its parent sortState" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<sortState ref=\"A2:D10\"><sortCondition ref=\"B2:B10\"/></sortState>");
    defer a.free(src);
    const out = try applyRowEditToWorksheet(a, src, 1, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<sortState ref=\"A3:D11\"") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<sortCondition ref=\"B3:B11\"") != null);
}

test "sortState nested in open-form autoFilter shifts on a col edit" {
    const a = testing.allocator;
    // The column walker consumes autoFilter children, so this path
    // reaches sortState only via the dispatch added inside that
    // walker — the row axis exercises the top-level dispatch instead.
    const src = try wrapSheet(a, "<autoFilter ref=\"A1:E10\"><sortState ref=\"B2:D10\"><sortCondition ref=\"C2:C10\"/></sortState></autoFilter>");
    defer a.free(src);
    const out = try applyColEditToWorksheet(a, src, 1, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<autoFilter ref=\"B1:F10\"") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<sortState ref=\"C2:E10\"") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<sortCondition ref=\"D2:D10\"") != null);
}

// ─── fuzz targets ───────────────────────────────────────────────────
//
// These walkers are ~3000 LOC of hand-rolled byte-splicing across
// sheet/table/drawing/vml_edit, and until now none of it was fuzzed.
// #125 is the argument for doing so: it found four coordinate-bearing
// elements the walkers neither rewrote nor refused, one of which had
// shipped and survived review. Those were found by reading. Fuzzing
// covers the shapes reading does not reach — truncated tags, unclosed
// elements, attributes that stop mid-quote, coordinates at the u32
// boundary.
//
// Contract: the walker must not panic, hang, or read out of bounds on
// ANY input. Returning a typed error is fine; producing nonsense XML
// from nonsense input is fine. Crashing is not.
//
// Each input is swept across both edit kinds and a few indices rather
// than consuming Smith entropy for them. That keeps the corpus plain
// readable XML, and makes every entry exercise the collapse boundary
// (delete at the index the element occupies), which is where the
// drop-vs-shift branches live.

const fuzz_indices = [_]u32{ 1, 2, 5 };

fn fuzzSheetEditTarget(_: void, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();
    // 4 KiB matches the PRNG harness's scratch bound, so a crash found
    // here reproduces against the same input shape.
    var smith_buf: [4096]u8 = undefined;
    const input = smith_buf[0..smith.slice(&smith_buf)];

    var arena = std.heap.ArenaAllocator.init(std.testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    for (fuzz_indices) |idx| {
        inline for (.{ .insert, .delete }) |kind| {
            if (applyRowEditToWorksheet(a, input, idx, kind)) |out| {
                a.free(out);
            } else |_| {}
            if (applyColEditToWorksheet(a, input, idx, kind)) |out| {
                a.free(out);
            } else |_| {}
        }
    }
}

/// Corpus entries are the shapes that actually broke things, plus the
/// truncations a hand-written test would not think to write.
const sheet_edit_corpus = [_][]const u8{
    "",
    "<worksheet/>",
    "<worksheet><sheetData/></worksheet>",
    // The four elements #125 found unhandled.
    "<sheetViews><sheetView topLeftCell=\"B5\"/></sheetViews>",
    "<selection activeCell=\"B5\" sqref=\"B5 D7:E9\"/>",
    "<sortState ref=\"A2:D10\"><sortCondition ref=\"B2:B10\"/></sortState>",
    // Truncations: every one of these is a place the walker indexes
    // forward after matching a tag name.
    "<c r=",
    "<c r=\"",
    "<c r=\"A",
    "<row r=\"1\"><c r=\"A1\"",
    "<mergeCell ref=\"A1:B2",
    "<autoFilter ref=\"A1:E10\"><filterColumn colId=\"0\"",
    "<sortState ref=\"A1:B2\"><sortCondition",
    "<pane xSplit=\"2\" topLeftCell=\"C1\" state=\"frozen\"",
    "<sheetView topLeftCell=",
    "<selection sqref=\"A1 \"",
    "<selection sqref=\"   \"/>",
    // Coordinate extremes — the shift helpers do bounds arithmetic.
    "<c r=\"XFD1048576\"/>",
    "<c r=\"A0\"/>",
    "<row r=\"0\"/>",
    "<row r=\"4294967296\"/>",
    "<col min=\"0\" max=\"0\"/>",
    "<col min=\"16384\" max=\"16384\"/>",
    "<mergeCell ref=\"A1:A1\"/>",
    "<dimension ref=\"A1:A1\"/>",
    // Malformed but plausible.
    "<c r=\"1A\"/>",
    "<c r=\"\"/>",
    "<c r=\"$A$1\"/>",
    "<autoFilter ref=\"\"/>",
    "<sortState/>",
    "<sortCondition ref=\"\"/>",
    "<<<<>>>>",
    "<c<c<c<c",
};

test "matchTagAt does not read past the end on a bare truncated tag" {
    // Regression, found by the fuzz target below within seconds of it
    // existing. `"<row"` ends exactly where the delimiter would be:
    // the old `after > src.len` guard let the read through.
    try testing.expect(matchTagAt("<row", 0, "row") == null);
    try testing.expect(matchTagAt("<c", 0, "c") == null);
    try testing.expect(matchTagAt("<mergeCell", 0, "mergeCell") == null);
    // Still matches when a delimiter is actually present.
    try testing.expect(matchTagAt("<row>", 0, "row") != null);
    try testing.expect(matchTagAt("<row />", 0, "row") != null);
    // And still rejects a longer name that merely shares the prefix.
    try testing.expect(matchTagAt("<rowspan>", 0, "row") == null);
}

test "unterminated attribute value does not slice backwards" {
    // Regression. `ref="A1:E10` never closes inside the tag, but there
    // are quotes later in the document; the unbounded search matched
    // one of those and produced start > end. Found by the mutation
    // stress in workbook.zig.
    const a = testing.allocator;
    const src =
        \\<worksheet><autoFilter ref="A1:E10><row r="1"><c r="A1"/></row></worksheet>
    ;
    // Contract is "does not crash"; a typed error or passthrough are
    // both acceptable outcomes.
    if (applyRowEditToWorksheet(a, src, 1, .insert)) |out| a.free(out) else |_| {}
    if (applyColEditToWorksheet(a, src, 1, .insert)) |out| a.free(out) else |_| {}
    if (applyRowEditToWorksheet(a, src, 1, .delete)) |out| a.free(out) else |_| {}
    if (applyColEditToWorksheet(a, src, 1, .delete)) |out| a.free(out) else |_| {}
}

test "an over-long A1 reference is refused, not truncated" {
    // Regression. `A` parses as column 1 and the row digits are never
    // bounded, so the ref outgrew the 16-byte scratch buffer. Silent
    // truncation would be worse than an error: a shortened reference
    // still looks like a valid one.
    var buf: [16]u8 = undefined;
    const long_ref = "A99999999999999999";
    try testing.expect(long_ref.len > buf.len);
    try testing.expectError(
        error.MalformedXml,
        shiftSingleA1Col(long_ref, 99, .insert, &buf, false),
    );
    var buf2: [16]u8 = undefined;
    try testing.expectError(
        error.MalformedXml,
        shiftSingleA1Row(long_ref, 99, .insert, &buf2, false),
    );
    // A normal ref still round-trips unchanged when nothing shifts.
    var buf3: [16]u8 = undefined;
    try testing.expectEqualStrings("B7", try shiftSingleA1Col("B7", 99, .insert, &buf3, false));
}

test "fuzz: sheet_edit walkers never crash on adversarial XML" {
    try std.testing.fuzz({}, fuzzSheetEditTarget, .{ .corpus = &sheet_edit_corpus });
}

// ─── extLst / xm:sqref (the last verbatim-passthrough surface) ──────
//
// `docs/plans/refusal-audit.md` §"Method note": the refusal list is
// derived from what the Editor *scans for*, and `<extLst>` blocks
// (`x14:`/`x15:` extensions) "pass through verbatim everywhere" — the
// one surface that method could not reach. These pin the shift.

test "xm:sqref: row insert shifts the extension's range" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<extLst><ext uri=\"{78C0D931}\"><x14:conditionalFormattings>" ++
        "<x14:conditionalFormatting><xm:sqref>A5:A9</xm:sqref>" ++
        "</x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>");
    defer a.free(src);
    const out = try applyRowEditToWorksheet(a, src, 2, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<xm:sqref>A6:A10</xm:sqref>") != null);
}

test "xm:sqref: col insert shifts the extension's range" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<extLst><ext><x14:dataValidations><x14:dataValidation>" ++
        "<xm:sqref>C1:D4</xm:sqref></x14:dataValidation>" ++
        "</x14:dataValidations></ext></extLst>");
    defer a.free(src);
    const out = try applyColEditToWorksheet(a, src, 1, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<xm:sqref>D1:E4</xm:sqref>") != null);
}

test "xm:sqref: a multi-range list shifts every entry" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<extLst><ext><xm:sqref>A5 C7:C9 E2</xm:sqref></ext></extLst>");
    defer a.free(src);
    const out = try applyRowEditToWorksheet(a, src, 3, .insert);
    defer a.free(out);
    // A5 and C7:C9 are at/below the insert point and move; E2 is above
    // it and must not.
    try testing.expect(std.mem.indexOf(u8, out, "<xm:sqref>A6 C8:C10 E2</xm:sqref>") != null);
}

test "xm:sqref: ranges above the edit point are left alone" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<extLst><ext><xm:sqref>A1:A2</xm:sqref></ext></extLst>");
    defer a.free(src);
    const out = try applyRowEditToWorksheet(a, src, 9, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<xm:sqref>A1:A2</xm:sqref>") != null);
}

test "xm:sqref: an unterminated element is emitted unchanged" {
    const a = testing.allocator;
    // No `</xm:sqref>`. Emitting the open tag and moving on beats
    // scanning to end-of-buffer looking for a close that isn't there —
    // the same failure posture every other handler here takes.
    const src = try wrapSheet(a, "<extLst><ext><xm:sqref>A5:A9</ext></extLst>");
    defer a.free(src);
    const out = try applyRowEditToWorksheet(a, src, 2, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "A5:A9") != null);
}

test "xm:sqref: an unparseable ref leaves the whole list untouched" {
    const a = testing.allocator;
    // Partially-shifting a list would be worse than not shifting it:
    // the caller cannot tell which entries moved.
    const src = try wrapSheet(a, "<extLst><ext><xm:sqref>A5 !!bogus!!</xm:sqref></ext></extLst>");
    defer a.free(src);
    const out = try applyRowEditToWorksheet(a, src, 2, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<xm:sqref>A5 !!bogus!!</xm:sqref>") != null);
}

test "xm:sqref: a sheet with no extLst is byte-identical" {
    const a = testing.allocator;
    const src = try wrapSheet(a, "<autoFilter ref=\"B2:D5\"/>");
    defer a.free(src);
    const with = try applyRowEditToWorksheet(a, src, 2, .insert);
    defer a.free(with);
    // The new branch must not perturb sheets that never had one.
    try testing.expect(std.mem.indexOf(u8, with, "xm:sqref") == null);
    try testing.expect(std.mem.indexOf(u8, with, "<autoFilter ref=\"B3:D6\"/>") != null);
}

// ─── extLst / xm:f (the formula carrier the workbook sweep rewrites) ─

test "xm:f: the byte transform leaves the formula body untouched" {
    // Division of labour: this file shifts `xm:sqref`; the workbook
    // sweep rewrites `xm:f`. Both edit the same `<extLst>`, so the
    // transform must not perturb the body the sweep splices.
    const a = testing.allocator;
    const src = try wrapSheet(a, "<extLst><ext><x14:sparklineGroups><x14:sparklineGroup><x14:sparklines>" ++
        "<x14:sparkline><xm:f>Data!A5:A9</xm:f><xm:sqref>B5</xm:sqref></x14:sparkline>" ++
        "</x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst>");
    defer a.free(src);
    const out = try applyRowEditToWorksheet(a, src, 2, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<xm:f>Data!A5:A9</xm:f>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<xm:sqref>B6</xm:sqref>") != null);
}

test "xm:f: the scan walks every carrier in document order" {
    const src = "<a><xm:f>Data!A1:A5</xm:f><xm:sqref>B1</xm:sqref><xm:f>C1&gt;0</xm:f></a>";
    const first = (try nextXmFormula(src, 0)).?;
    try testing.expectEqualStrings("Data!A1:A5", src[first.body_start..first.body_end]);
    const second = (try nextXmFormula(src, first.next)).?;
    try testing.expectEqualStrings("C1&gt;0", src[second.body_start..second.body_end]);
    try testing.expectEqual(@as(?XmFormula, null), try nextXmFormula(src, second.next));
}

test "xm:f: a sheet without the element yields null, not an error" {
    try testing.expectEqual(@as(?XmFormula, null), try nextXmFormula("<worksheet/>", 0));
    // `xm:sqref` is the sibling leaf; `<xm:f` must not match a
    // longer name that happens to share the prefix.
    try testing.expectEqual(@as(?XmFormula, null), try nextXmFormula("<xm:fx>1</xm:fx><xm:sqref>A1</xm:sqref>", 0));
}

test "xm:f: a self-closing element is an empty body" {
    const src = "<x14:sparkline><xm:f/><xm:sqref>B1</xm:sqref></x14:sparkline>";
    const f = (try nextXmFormula(src, 0)).?;
    try testing.expectEqual(f.body_start, f.body_end);
    try testing.expectEqual(f.body_end, f.next);
    try testing.expectEqual(@as(?XmFormula, null), try nextXmFormula(src, f.next));
}

test "xm:f: an unterminated element refuses" {
    // All-or-nothing: unlike `xm:sqref`, a carrier that cannot be
    // rewritten wholly is an error the workbook turns into a refusal
    // of the whole edit, never a silent pass-through.
    try testing.expectError(error.MalformedExtensionXml, nextXmFormula("<xm:f>Data!A1:A5</ext>", 0));
}

test "xm:f: markup inside the body refuses" {
    try testing.expectError(error.MalformedExtensionXml, nextXmFormula("<xm:f>A1<!-- x --></xm:f>", 0));
    try testing.expectError(error.MalformedExtensionXml, nextXmFormula("<xm:f><![CDATA[A1]]></xm:f>", 0));
}

test "xm:f: a truncated open tag at end of input is not a carrier" {
    // `matchTagAt` refuses to read past the end (fuzz finding,
    // 2026-07-27); the scan must inherit that and terminate.
    try testing.expectEqual(@as(?XmFormula, null), try nextXmFormula("<xm:f", 0));
}

test "xm:f: carrier text inside a comment, CDATA or PI is neither a carrier nor a refusal" {
    // Codex S2 r1: a raw substring walk would splice the comment's
    // bytes or refuse on the unclosed decoy. Each construct is
    // stepped over whole; the live carrier after it is found.
    const src = "<!-- <xm:f>Doomed!A1 --><![CDATA[<xm:f>Doomed!A1</xm:f>]]><?pi <xm:f> ?>" ++
        "<xm:f>Data!A1:A5</xm:f>";
    const f = (try nextXmFormula(src, 0)).?;
    try testing.expectEqualStrings("Data!A1:A5", src[f.body_start..f.body_end]);
    try testing.expectEqual(@as(?XmFormula, null), try nextXmFormula(src, f.next));
    // Decoys only: nothing to rewrite, nothing to refuse.
    try testing.expectEqual(@as(?XmFormula, null), try nextXmFormula("<!-- <xm:f>x --><a/>", 0));
}

test "xm:f: an unterminated comment on a sheet that spells <xm:f refuses" {
    // Live or decoy is undecidable once the comment never closes.
    try testing.expectError(error.MalformedExtensionXml, nextXmFormula("<!-- <xm:f>Data!A1</xm:f>", 0));
    // …but a sheet with no `<xm:f` text at all never pays for, or
    // refuses on, its comments.
    try testing.expectEqual(@as(?XmFormula, null), try nextXmFormula("<!-- open forever <a/>", 0));
}
