//! `<table>` part rewriter for OOXML `xl/tables/tableN.xml`.
//! Pure-function; consumed by `pkg/editor.zig`'s row/col edit path
//! after the worksheet's own bytes have been rewritten by
//! `pkg/sheet_edit.zig`. Companion to (not extension of) sheet_edit
//! because table parts live separately — each `<tableParts>` ref
//! in a sheet resolves through rels to its own
//! `xl/tables/tableN.xml`.
//!
//! Walks the `<table>` open tag to shift `ref="A1:Z100"`. Inside
//! the body shifts the inner `<autoFilter ref>` (reusing
//! `sheet_edit.processAutoFilterTagCol` / `Row`, including the
//! `<filterColumn colId>` rebase) and `<sortState ref>` (which the
//! autoFilter lift left as the only caveat). On col edits walks
//! `<tableColumns count="N">` to drop the matching `<tableColumn>`
//! on a delete or insert a synthetic column entry on an insert;
//! `count=` is updated to match.
//!
//! Per ECMA-376 §18.5.1.78 (`CT_TableColumn`), `<tableColumn id>`
//! is a stable **table-unique** field ID, NOT a positional index
//! and NOT workbook-unique (the `<table id>` attribute is the
//! workbook-unique one) — survivors keep their ids; an inserted
//! column claims `max(existing ids) + 1`.
//!
//! `<tableColumn name>` must also be unique within the parent table
//! (same spec section). Synthetic inserts probe `Column<id>`,
//! `Column<id>_2`, ... until finding a free name to avoid clashing
//! with a pre-existing column literally named "Column4" etc.
//!
//! Pre-flight refusals: a table cannot legally collapse to zero
//! columns or zero rows (header + ≥0 data rows). Edits that would
//! do so surface `error.TableCollapseUnsafe` so the Editor can
//! refuse the entire workbook mutation BEFORE the sheet bytes are
//! replaced. Likewise, deleting the header row (the top row of the
//! range when `headerRowCount >= 1`, which is the default) surfaces
//! `error.TableHeaderRowDeleteUnsafe`. Both errors are intended for
//! Editor pre-flight; once a sheet edit has begun, the same code
//! path runs in commit mode and propagates the error to the caller.
//!
//! v1 LIMITATIONS:
//! - `<extLst>` table extensions pass through verbatim; no
//!   coordinate fixups inside `x14:table` / `x15:table` blocks.
//! - `totalsRowCount > 0`: a delete-match against the BR row of
//!   such a table silently loses the totals row's labels +
//!   formulas. zlsx's own writer never emits totals rows; v1 lets
//!   the shift happen rather than refusing. Document if user
//!   reports.
//! - `headerRowCount > 1`: spec-legal but vanishingly rare. v1
//!   refuses any top-row delete regardless of headerRowCount;
//!   future iter can lift this once a corpus fixture appears.

const std = @import("std");
const xlsx = @import("zlsx");
const sheet_edit = @import("sheet_edit.zig");

const Allocator = std.mem.Allocator;
const TagOpen = xlsx.TagOpen;
const getAttr = xlsx.getAttr;
const max_col_1based = xlsx.max_col_1based;
const max_row = xlsx.max_row;

pub const Error = error{
    MalformedTableXml,
    TableCoordinateOverflow,
    TableCollapseUnsafe,
    TableHeaderRowDeleteUnsafe,
    /// Surfaces only when an internal `bufPrint` overflows the
    /// fixed 48-byte ref buffer — an A1 range fits in ~13 bytes
    /// so this is practically unreachable.
    NoSpaceLeft,
} || Allocator.Error;

pub const Axis = enum { row, col };
pub const EditKind = enum { insert, delete };

/// Captured snapshot of the `<table>` open tag's relevant attrs.
/// All bounds are 1-based to mirror sheet_edit's convention; the
/// caller passes 1-based row/col indices.
const TableHeader = struct {
    tag: TagOpen,
    tl_col: u32,
    br_col: u32,
    tl_row: u32,
    br_row: u32,
    /// Defaults to 1 per ECMA-376 §18.5.1.2. Distinguishes
    /// "explicit 0" (legitimate header-less table) from "absent or
    /// 1" (default headered table) so `checkEditSafe` can permit
    /// top-row deletes in the header-less case (REL-A501).
    header_row_count: u32,
    header_row_count_explicit_zero: bool,
};

/// Apply one row OR column edit to a `xl/tables/tableN.xml` body.
/// `idx_1based` matches sheet_edit's convention. Returns a freshly
/// allocated buffer (caller frees).
pub fn applyEditToTable(
    allocator: Allocator,
    src: []const u8,
    axis: Axis,
    idx_1based: u32,
    kind: EditKind,
) Error![]u8 {
    const hdr = parseTableHeader(src) orelse return error.MalformedTableXml;

    // Pre-flight refusals (precede any byte work).
    try checkEditSafe(hdr, axis, idx_1based, kind);

    // Compute the shifted range; encode "outside-range" edits as
    // (new == old) bounds rather than special-casing.
    const new_bounds = try shiftTableBounds(hdr, axis, idx_1based, kind);
    const ek: sheet_edit.RowEditKind = switch (kind) {
        .insert => .insert,
        .delete => .delete,
    };

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

        if (sheet_edit.matchTagAt(src, i, "table")) |t| {
            // Single `<table>` per part, but the walker still
            // accepts it positionally. Rewrite `ref=` only; other
            // attrs stay.
            try emitTableOpenWithRef(allocator, &out, src, t, hdr, new_bounds);
            i = t.after_open;
        } else if (sheet_edit.matchTagAt(src, i, "autoFilter")) |t| {
            // Delegate to the existing sheet_edit handler — same
            // attr/child rewrite the bare-sheet autoFilter gets,
            // including filterColumn colId rebase.
            switch (axis) {
                .col => try sheet_edit.processAutoFilterTagCol(allocator, &out, src, t, idx_1based, ek, &i),
                .row => try sheet_edit.processAutoFilterTagRow(allocator, &out, src, t, idx_1based, ek, &i),
            }
        } else if (sheet_edit.matchTagAt(src, i, "sortState")) |t| {
            try processSortStateTag(allocator, &out, src, t, axis, idx_1based, ek, &i);
        } else if (sheet_edit.matchTagAt(src, i, "tableColumns")) |t| {
            if (axis == .col) {
                try processTableColumnsForCol(allocator, &out, src, t, hdr, idx_1based, kind, &i);
            } else {
                // Row edits don't change the column list.
                try out.appendSlice(allocator, src[t.start..t.after_open]);
                i = t.after_open;
            }
        } else {
            try out.append(allocator, '<');
            i += 1;
        }
    }
    return try out.toOwnedSlice(allocator);
}

// ─── parsing ────────────────────────────────────────────────────

fn parseTableHeader(src: []const u8) ?TableHeader {
    // Find the first `<table` open tag (not `<tableColumn(s)>`).
    var i: usize = 0;
    while (i < src.len) {
        const lt = std.mem.indexOfScalarPos(u8, src, i, '<') orelse return null;
        if (sheet_edit.matchTagAt(src, lt, "table")) |t| {
            const attrs = src[t.start + "<table".len .. t.after_open - 1];
            const ref = getAttr(attrs, "ref") orelse return null;
            const range = parseRange(ref) orelse return null;
            // REL-A501 + REL-A509: distinguish "explicit 0" from
            // "default 1". Malformed (un-parseable) values fall
            // back to the spec default (1) — strictness here would
            // surface MalformedTableXml from a third-party file
            // with `headerRowCount="x"`, which Editor would remap
            // to RowEditUnsafeForSheet anyway.
            var hrc: u32 = 1;
            var hrc_explicit_zero = false;
            if (getAttr(attrs, "headerRowCount")) |v| {
                if (std.fmt.parseInt(u32, v, 10) catch null) |n| {
                    hrc = n;
                    if (n == 0) hrc_explicit_zero = true;
                }
            }
            return .{
                .tag = t,
                .tl_col = range.tl_col,
                .br_col = range.br_col,
                .tl_row = range.tl_row,
                .br_row = range.br_row,
                .header_row_count = hrc,
                .header_row_count_explicit_zero = hrc_explicit_zero,
            };
        }
        i = lt + 1;
    }
    return null;
}

const Range = struct {
    tl_col: u32,
    br_col: u32,
    tl_row: u32,
    br_row: u32,
};

fn parseRange(ref: []const u8) ?Range {
    const colon = std.mem.indexOfScalar(u8, ref, ':');
    if (colon) |c| {
        const tl_col = sheet_edit.parseColFromA1(ref[0..c]) orelse return null;
        const tl_row = sheet_edit.parseRowFromA1(ref[0..c]) orelse return null;
        const br_col = sheet_edit.parseColFromA1(ref[c + 1 ..]) orelse return null;
        const br_row = sheet_edit.parseRowFromA1(ref[c + 1 ..]) orelse return null;
        if (br_col < tl_col or br_row < tl_row) return null;
        return .{ .tl_col = tl_col, .br_col = br_col, .tl_row = tl_row, .br_row = br_row };
    }
    const c = sheet_edit.parseColFromA1(ref) orelse return null;
    const r = sheet_edit.parseRowFromA1(ref) orelse return null;
    return .{ .tl_col = c, .br_col = c, .tl_row = r, .br_row = r };
}

// ─── edit-safety checks ─────────────────────────────────────────

fn checkEditSafe(hdr: TableHeader, axis: Axis, idx_1based: u32, kind: EditKind) Error!void {
    if (idx_1based == 0) return error.MalformedTableXml;
    if (kind == .insert) {
        // Inserts never collapse or remove the header. Overflow is
        // caught by the per-corner shift later.
        return;
    }
    // kind == .delete
    switch (axis) {
        .row => {
            // Header-row delete: refuse when the edit lands on the
            // table's top row AND the table actually has a header
            // (REL-A501). `headerRowCount="0"` (explicit) means a
            // header-less table — top-row delete is a normal data
            // shrink, not a structural break, so allow it.
            if (idx_1based == hdr.tl_row and hdr.header_row_count >= 1 and !hdr.header_row_count_explicit_zero) {
                return error.TableHeaderRowDeleteUnsafe;
            }
            // Collapse: single-row table whose only row is deleted.
            if (hdr.tl_row == hdr.br_row and idx_1based == hdr.tl_row) {
                return error.TableCollapseUnsafe;
            }
        },
        .col => {
            // Collapse: single-column table whose only column is
            // deleted.
            if (hdr.tl_col == hdr.br_col and idx_1based == hdr.tl_col) {
                return error.TableCollapseUnsafe;
            }
        },
    }
}

const NewBounds = struct {
    tl_col: u32,
    br_col: u32,
    tl_row: u32,
    br_row: u32,
};

fn shiftTableBounds(hdr: TableHeader, axis: Axis, idx_1based: u32, kind: EditKind) Error!NewBounds {
    var b: NewBounds = .{
        .tl_col = hdr.tl_col,
        .br_col = hdr.br_col,
        .tl_row = hdr.tl_row,
        .br_row = hdr.br_row,
    };
    switch (axis) {
        .col => switch (kind) {
            .insert => {
                if (hdr.tl_col >= idx_1based) {
                    if (hdr.tl_col >= max_col_1based) return error.TableCoordinateOverflow;
                    b.tl_col = hdr.tl_col + 1;
                }
                if (hdr.br_col >= idx_1based) {
                    if (hdr.br_col >= max_col_1based) return error.TableCoordinateOverflow;
                    b.br_col = hdr.br_col + 1;
                }
            },
            .delete => {
                if (hdr.tl_col > idx_1based) b.tl_col = hdr.tl_col - 1;
                // BR shrinks on delete-match too (range was inclusive
                // of the deleted column, so the upper bound moves
                // down by one).
                if (hdr.br_col >= idx_1based and hdr.br_col > 0) b.br_col = hdr.br_col - 1;
            },
        },
        .row => switch (kind) {
            .insert => {
                if (hdr.tl_row >= idx_1based) {
                    if (hdr.tl_row >= max_row) return error.TableCoordinateOverflow;
                    b.tl_row = hdr.tl_row + 1;
                }
                if (hdr.br_row >= idx_1based) {
                    if (hdr.br_row >= max_row) return error.TableCoordinateOverflow;
                    b.br_row = hdr.br_row + 1;
                }
            },
            .delete => {
                if (hdr.tl_row > idx_1based) b.tl_row = hdr.tl_row - 1;
                if (hdr.br_row >= idx_1based and hdr.br_row > 0) b.br_row = hdr.br_row - 1;
            },
        },
    }
    return b;
}

// ─── emitters ───────────────────────────────────────────────────

fn emitTableOpenWithRef(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    hdr: TableHeader,
    new_bounds: NewBounds,
) !void {
    const unchanged =
        new_bounds.tl_col == hdr.tl_col and new_bounds.br_col == hdr.br_col and
        new_bounds.tl_row == hdr.tl_row and new_bounds.br_row == hdr.br_row;
    if (unchanged) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    var ref_buf: [48]u8 = undefined;
    const new_ref = try formatRange(&ref_buf, new_bounds);
    try sheet_edit.writeWithReplacedAttr(allocator, out, src, t, "<table".len, "ref", new_ref);
}

fn formatRange(buf: *[48]u8, b: NewBounds) ![]const u8 {
    var tl_col_buf: [8]u8 = undefined;
    var br_col_buf: [8]u8 = undefined;
    const tl_col = try formatColLettersLocal(&tl_col_buf, b.tl_col);
    const br_col = try formatColLettersLocal(&br_col_buf, b.br_col);
    if (b.tl_col == b.br_col and b.tl_row == b.br_row) {
        return try std.fmt.bufPrint(buf, "{s}{d}", .{ tl_col, b.tl_row });
    }
    return try std.fmt.bufPrint(buf, "{s}{d}:{s}{d}", .{ tl_col, b.tl_row, br_col, b.br_row });
}

/// 1-based col index → A1 letters (1=A). Local mirror of
/// `pkg/sheet_edit.zig::formatColLetters` (which is private). When a
/// third consumer appears, lift both into a shared `pkg/range_a1.zig`.
fn formatColLettersLocal(buf: *[8]u8, col_1based: u32) ![]const u8 {
    if (col_1based == 0) return error.TableCoordinateOverflow;
    var tmp: [8]u8 = undefined;
    var pos: usize = tmp.len;
    var c = col_1based;
    while (c > 0) {
        c -= 1;
        if (pos == 0) return error.TableCoordinateOverflow;
        pos -= 1;
        tmp[pos] = @intCast('A' + (c % 26));
        c /= 26;
    }
    const len = tmp.len - pos;
    @memcpy(buf[0..len], tmp[pos..]);
    return buf[0..len];
}

fn processSortStateTag(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    axis: Axis,
    idx_1based: u32,
    kind: sheet_edit.RowEditKind,
    i: *usize,
) !void {
    const attrs_full = src[t.start + "<sortState".len .. t.after_open - 1];
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

    // Capture bounds on the edited axis to decide drop-vs-shift.
    var old_tl: u32 = 0;
    var old_br: u32 = 0;
    switch (axis) {
        .col => if (colon) |c| {
            old_tl = sheet_edit.parseColFromA1(ref[0..c]) orelse 0;
            old_br = sheet_edit.parseColFromA1(ref[c + 1 ..]) orelse 0;
        } else {
            old_tl = sheet_edit.parseColFromA1(ref) orelse 0;
            old_br = old_tl;
        },
        .row => if (colon) |c| {
            old_tl = sheet_edit.parseRowFromA1(ref[0..c]) orelse 0;
            old_br = sheet_edit.parseRowFromA1(ref[c + 1 ..]) orelse 0;
        } else {
            old_tl = sheet_edit.parseRowFromA1(ref) orelse 0;
            old_br = old_tl;
        },
    }

    // Full-range collapse: drop the entire sortState (mirrors
    // sheet_edit's autoFilter drop-on-collapse). The sortState's
    // body (sortCondition children) goes with it.
    const drop_entire =
        kind == .delete and old_tl == idx_1based and old_br == idx_1based and old_tl != 0;
    if (drop_entire) {
        if (is_self_closing) {
            i.* = t.after_open;
        } else {
            const close = std.mem.indexOfPos(u8, src, t.after_open, "</sortState>") orelse t.after_open;
            i.* = if (close + "</sortState>".len <= src.len) close + "</sortState>".len else t.after_open;
        }
        return;
    }

    // Compute new ref via the axis-appropriate shifter.
    var tl_buf: [16]u8 = undefined;
    var br_buf: [16]u8 = undefined;
    var new_ref_buf: [40]u8 = undefined;
    var new_ref: []const u8 = ref;
    if (colon) |c| {
        const tl_new = shiftHalf(ref[0..c], axis, idx_1based, kind, &tl_buf, false) catch {
            try out.appendSlice(allocator, src[t.start..t.after_open]);
            i.* = t.after_open;
            return;
        };
        const br_new = shiftHalf(ref[c + 1 ..], axis, idx_1based, kind, &br_buf, true) catch {
            try out.appendSlice(allocator, src[t.start..t.after_open]);
            i.* = t.after_open;
            return;
        };
        new_ref = std.fmt.bufPrint(&new_ref_buf, "{s}:{s}", .{ tl_new, br_new }) catch {
            try out.appendSlice(allocator, src[t.start..t.after_open]);
            i.* = t.after_open;
            return;
        };
    } else {
        const shifted = shiftHalf(ref, axis, idx_1based, kind, &tl_buf, false) catch {
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
        try sheet_edit.writeWithReplacedAttr(allocator, out, src, t, "<sortState".len, "ref", new_ref);
    }
    i.* = t.after_open;
}

/// Wrapper that picks the right axis-specific shifter. Keeps the
/// caller branch-free at each invocation site.
fn shiftHalf(
    ref: []const u8,
    axis: Axis,
    idx_1based: u32,
    kind: sheet_edit.RowEditKind,
    buf: *[16]u8,
    is_br_corner: bool,
) ![]const u8 {
    return switch (axis) {
        .col => try sheet_edit.shiftSingleA1Col(ref, idx_1based, kind, buf, is_br_corner),
        .row => try sheet_edit.shiftSingleA1Row(ref, idx_1based, kind, buf, is_br_corner),
    };
}

// ─── <tableColumns> ─────────────────────────────────────────────

fn processTableColumnsForCol(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    hdr: TableHeader,
    col_1based: u32,
    kind: EditKind,
    i: *usize,
) !void {
    // No-op fast path: edit is outside the table's column span.
    const inside = col_1based >= hdr.tl_col and col_1based <= hdr.br_col;
    const insert_at_left_edge = kind == .insert and col_1based == hdr.tl_col;
    // INSERT at col == tl_col pushes the whole table right by one
    // (TL becomes tl+1), so the column count doesn't change and no
    // synthetic column is added. Same for any insert/delete fully
    // outside the table.
    if (!inside or insert_at_left_edge) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        i.* = t.after_open;
        return;
    }

    // Locate the close tag and the existing count.
    const close_tag = "</tableColumns>";
    const close_pos = std.mem.indexOfPos(u8, src, t.after_open, close_tag) orelse {
        // Malformed — emit open tag verbatim and stop.
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        i.* = t.after_open;
        return;
    };
    const after_close = close_pos + close_tag.len;

    const attrs_full = src[t.start + "<tableColumns".len .. t.after_open - 1];
    const old_count_attr = getAttr(attrs_full, "count");
    const old_count = if (old_count_attr) |c| (std.fmt.parseInt(u32, c, 10) catch null) else null;
    const expected_count = hdr.br_col - hdr.tl_col + 1;
    // REL-A505 + Codex Ticket 651: refuse on any of three
    // independent divergence conditions. The round-1 fix only
    // checked the declared `count` attr against `expected_count`;
    // that misses the case where `count="3"` agrees with a
    // `B2:D5` range but only TWO `<tableColumn>` children are
    // present — the walker would emit `count="4"` after an insert
    // while outputting only THREE children, producing repair-load
    // XML. We now also count actual children and require ALL
    // THREE (declared count, child count, range width) to agree.
    //
    // Refusal happens before this function emits the rewritten
    // `<tableColumns>` open tag. Bytes already buffered by the
    // outer walker (`<table>` open + any preceding autoFilter /
    // sortState) are reclaimed via `applyEditToTable`'s outer
    // `errdefer out.deinit(allocator)` so no caller observes a
    // partial result.
    const child_count = countTableColumnChildren(src, t.after_open, close_pos);
    if (old_count) |c| {
        if (c != expected_count) return error.MalformedTableXml;
    }
    if (child_count != expected_count) return error.MalformedTableXml;
    const total_cols = expected_count;

    // The position within the table at which to insert/delete a
    // `<tableColumn>`. 0-based, mirrors filterColumn colId.
    const target_pos: u32 = col_1based - hdr.tl_col;

    // Rewrite the open tag (count=) first.
    const new_count: u32 = switch (kind) {
        .insert => total_cols + 1,
        .delete => if (total_cols > 0) total_cols - 1 else 0,
    };
    if (new_count == 0) return error.TableCollapseUnsafe;

    var count_buf: [16]u8 = undefined;
    const count_str = try std.fmt.bufPrint(&count_buf, "{d}", .{new_count});
    try sheet_edit.writeWithReplacedAttr(allocator, out, src, t, "<tableColumns".len, "count", count_str);

    // Pre-pass: find the max existing `<tableColumn id>` so the
    // synthetic insert claims `max + 1` without colliding with any
    // not-yet-walked sibling. ids are stable table-unique
    // identifiers (ECMA-376 §18.5.1.78), not positions.
    const max_existing_id = scanMaxTableColumnId(src, t.after_open, close_pos);

    // Walk children, copy/drop/insert.
    var seen_pos: u32 = 0;
    var inserted = false;
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

        if (sheet_edit.matchTagAt(src, j, "tableColumn")) |ct| {
            const ct_attrs_full = src[ct.start + "<tableColumn".len .. ct.after_open - 1];
            const trimmed = std.mem.trimEnd(u8, ct_attrs_full, " \t\r\n");
            const is_self_closing = trimmed.len > 0 and trimmed[trimmed.len - 1] == '/';

            // Locate the full extent of this tableColumn (open +
            // optional body + close). Needed for both drop and
            // copy-then-shift cases.
            const ct_end: usize = if (is_self_closing) ct.after_open else blk: {
                const cend = "</tableColumn>";
                const p = std.mem.indexOfPos(u8, src, ct.after_open, cend) orelse close_pos;
                break :blk if (p + cend.len <= close_pos) p + cend.len else close_pos;
            };

            switch (kind) {
                .insert => {
                    if (!inserted and seen_pos == target_pos) {
                        try emitSyntheticTableColumn(allocator, out, src, t.after_open, close_pos, max_existing_id + 1);
                        inserted = true;
                    }
                    // Copy this tableColumn verbatim — ids stay
                    // stable, position shifts implicitly via the
                    // earlier insert.
                    try out.appendSlice(allocator, src[ct.start..ct_end]);
                    seen_pos += 1;
                },
                .delete => {
                    if (seen_pos == target_pos) {
                        // Drop entirely.
                    } else {
                        try out.appendSlice(allocator, src[ct.start..ct_end]);
                    }
                    seen_pos += 1;
                },
            }
            j = ct_end;
        } else {
            try out.append(allocator, '<');
            j += 1;
        }
    }

    // Insert beyond the last existing column? (target_pos == count)
    if (kind == .insert and !inserted) {
        try emitSyntheticTableColumn(allocator, out, src, t.after_open, close_pos, max_existing_id + 1);
    }

    // Emit close tag and advance.
    try out.appendSlice(allocator, src[close_pos..after_close]);
    i.* = after_close;
}

fn emitSyntheticTableColumn(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    body_start: usize,
    body_end: usize,
    new_id: u32,
) !void {
    // REL-A502: pick a `name` attribute that doesn't collide with
    // any existing `<tableColumn name>` in the same table (ECMA-376
    // §18.5.1.78 requires names unique within the parent table).
    // Try `Column<id>`, then `Column<id>_2`, `Column<id>_3`, ...
    var name_buf: [40]u8 = undefined;
    var suffix: u32 = 1;
    const name = while (true) {
        const candidate = if (suffix == 1)
            try std.fmt.bufPrint(&name_buf, "Column{d}", .{new_id})
        else
            try std.fmt.bufPrint(&name_buf, "Column{d}_{d}", .{ new_id, suffix });
        if (!tableColumnNameTaken(src, body_start, body_end, candidate)) break candidate;
        suffix += 1;
        if (suffix > 1000) return error.MalformedTableXml; // pathological
    };
    var buf: [80]u8 = undefined;
    const written = try std.fmt.bufPrint(&buf, "<tableColumn id=\"{d}\" name=\"{s}\"/>", .{ new_id, name });
    try out.appendSlice(allocator, written);
}

/// Walk `<tableColumn>` siblings looking for one whose `name`
/// attribute equals `candidate`. Used by `emitSyntheticTableColumn`
/// to dodge name collisions on insert (REL-A502).
fn tableColumnNameTaken(src: []const u8, body_start: usize, body_end: usize, candidate: []const u8) bool {
    var k = body_start;
    while (k < body_end) {
        const lt = std.mem.indexOfScalarPos(u8, src, k, '<') orelse return false;
        if (lt >= body_end) return false;
        if (sheet_edit.matchTagAt(src, lt, "tableColumn")) |ct| {
            if (ct.after_open > body_end) return false;
            const attrs = src[ct.start + "<tableColumn".len .. ct.after_open - 1];
            if (getAttr(attrs, "name")) |existing| {
                if (std.mem.eql(u8, existing, candidate)) return true;
            }
            k = ct.after_open;
        } else {
            k = lt + 1;
        }
    }
    return false;
}

/// Count direct `<tableColumn>` children inside a `<tableColumns>`
/// body. Codex Ticket 651 introduced this; Codex Ticket 701 fixed
/// the flat-scan: when a direct `<tableColumn>` is found we MUST
/// advance past its full `<tableColumn>...</tableColumn>` body so a
/// descendant `<tableColumn>` (e.g., nested inside an `<extLst>`
/// extension) doesn't get double-counted as a sibling. Mirrors the
/// `ct_end` advance in `processTableColumnsForCol`.
fn countTableColumnChildren(src: []const u8, body_start: usize, body_end: usize) u32 {
    var n: u32 = 0;
    var k = body_start;
    while (k < body_end) {
        const lt = std.mem.indexOfScalarPos(u8, src, k, '<') orelse return n;
        if (lt >= body_end) return n;
        if (sheet_edit.matchTagAt(src, lt, "tableColumn")) |ct| {
            if (ct.after_open > body_end) return n;
            n += 1;
            const attrs_full = src[ct.start + "<tableColumn".len .. ct.after_open - 1];
            const trimmed = std.mem.trimEnd(u8, attrs_full, " \t\r\n");
            const is_self_closing = trimmed.len > 0 and trimmed[trimmed.len - 1] == '/';
            if (is_self_closing) {
                k = ct.after_open;
            } else {
                const cend = "</tableColumn>";
                const p = std.mem.indexOfPos(u8, src, ct.after_open, cend) orelse body_end;
                k = if (p + cend.len <= body_end) p + cend.len else body_end;
            }
        } else {
            k = lt + 1;
        }
    }
    return n;
}

/// Walk the `<tableColumns>` body for the highest `<tableColumn id>`
/// so a synthetic insert claims `max + 1`. Returns 0 when no
/// numeric ids are found (caller still gets `max + 1 == 1`, which
/// is fine as a fallback).
fn scanMaxTableColumnId(src: []const u8, body_start: usize, body_end: usize) u32 {
    var max_id: u32 = 0;
    var k = body_start;
    while (k < body_end) {
        const lt = std.mem.indexOfScalarPos(u8, src, k, '<') orelse return max_id;
        if (lt >= body_end) return max_id;
        if (sheet_edit.matchTagAt(src, lt, "tableColumn")) |ct| {
            if (ct.after_open > body_end) return max_id;
            const attrs = src[ct.start + "<tableColumn".len .. ct.after_open - 1];
            if (getAttr(attrs, "id")) |idv| {
                if (std.fmt.parseInt(u32, idv, 10) catch null) |n| {
                    if (n > max_id) max_id = n;
                }
            }
            k = ct.after_open;
        } else {
            k = lt + 1;
        }
    }
    return max_id;
}

// ─── tests ──────────────────────────────────────────────────────

const testing = std.testing;

const sample_table =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" displayName="T" ref="C9:E10" totalsRowShown="0">
    \\<autoFilter ref="C9:E10"/>
    \\<tableColumns count="3">
    \\<tableColumn id="1" name="A"/>
    \\<tableColumn id="2" name="B"/>
    \\<tableColumn id="3" name="C"/>
    \\</tableColumns>
    \\<tableStyleInfo name="TS"/>
    \\</table>
;

test "shift table ref on row insert above" {
    const a = testing.allocator;
    const out = try applyEditToTable(a, sample_table, .row, 5, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "ref=\"C10:E11\"") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<autoFilter ref=\"C10:E11\"") != null);
}

test "shift table ref on col insert above" {
    const a = testing.allocator;
    const out = try applyEditToTable(a, sample_table, .col, 1, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "ref=\"D9:F10\"") != null);
    // Insert at col 1 (A) is at the table's left edge — pushes the
    // table right; no synthetic column gets added, count stays 3.
    try testing.expect(std.mem.indexOf(u8, out, "<tableColumns count=\"3\">") != null);
}

test "col insert inside range adds synthetic tableColumn" {
    const a = testing.allocator;
    // Insert at col 4 (D) — between C9 (col 3) and E10 (col 5).
    const out = try applyEditToTable(a, sample_table, .col, 4, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "ref=\"C9:F10\"") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<tableColumns count=\"4\">") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<tableColumn id=\"4\" name=\"Column4\"/>") != null);
}

test "col delete inside range drops matching tableColumn" {
    const a = testing.allocator;
    // Delete at col 4 (D) — drops the middle column (B).
    const out = try applyEditToTable(a, sample_table, .col, 4, .delete);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "ref=\"C9:D10\"") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<tableColumns count=\"2\">") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<tableColumn id=\"2\" name=\"B\"/>") == null);
    try testing.expect(std.mem.indexOf(u8, out, "<tableColumn id=\"1\" name=\"A\"/>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<tableColumn id=\"3\" name=\"C\"/>") != null);
}

test "col delete at left edge drops first tableColumn" {
    const a = testing.allocator;
    const out = try applyEditToTable(a, sample_table, .col, 3, .delete);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "ref=\"C9:D10\"") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<tableColumn id=\"1\" name=\"A\"/>") == null);
}

test "row delete on header row refuses" {
    const a = testing.allocator;
    const r = applyEditToTable(a, sample_table, .row, 9, .delete);
    try testing.expectError(error.TableHeaderRowDeleteUnsafe, r);
}

test "col delete that collapses single-column table refuses" {
    const a = testing.allocator;
    const single_col =
        \\<?xml version="1.0"?>
        \\<table id="1" ref="B2:B5"><autoFilter ref="B2:B5"/><tableColumns count="1"><tableColumn id="1" name="X"/></tableColumns></table>
    ;
    const r = applyEditToTable(a, single_col, .col, 2, .delete);
    try testing.expectError(error.TableCollapseUnsafe, r);
}

test "row delete below header is allowed" {
    const a = testing.allocator;
    const out = try applyEditToTable(a, sample_table, .row, 10, .delete);
    defer a.free(out);
    // tl_row 9, br_row 10. Delete row 10 → tl 9, br 9. Range stays
    // multi-column, so it serialises as C9:E9 (collapse to a single
    // cell only happens when both axes shrink to one).
    try testing.expect(std.mem.indexOf(u8, out, "ref=\"C9:E9\"") != null);
}

test "sortState ref shifts in step" {
    const a = testing.allocator;
    const with_sort =
        \\<?xml version="1.0"?>
        \\<table id="1" ref="A1:C5"><autoFilter ref="A1:C5"><filterColumn colId="0"/></autoFilter><sortState ref="A2:C5"><sortCondition ref="A2:A5"/></sortState><tableColumns count="3"><tableColumn id="1" name="x"/><tableColumn id="2" name="y"/><tableColumn id="3" name="z"/></tableColumns></table>
    ;
    const out = try applyEditToTable(a, with_sort, .row, 1, .insert);
    defer a.free(out);
    // tl_row 1, br_row 5 → insert at 1 → tl 2, br 6.
    try testing.expect(std.mem.indexOf(u8, out, "<table id=\"1\" ref=\"A2:C6\"") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<sortState ref=\"A3:C6\"") != null);
}

test "edits outside table range pass through" {
    const a = testing.allocator;
    // sample_table is C9:E10. Insert at col 10 (J) — far right.
    const out = try applyEditToTable(a, sample_table, .col, 10, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "ref=\"C9:E10\"") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<tableColumns count=\"3\">") != null);
}

test "malformed table xml surfaces typed error" {
    const a = testing.allocator;
    const bad = "<?xml version=\"1.0\"?><notATable/>";
    const r = applyEditToTable(a, bad, .row, 1, .insert);
    try testing.expectError(error.MalformedTableXml, r);
}

test "table with no autoFilter shifts table ref only" {
    const a = testing.allocator;
    const no_af =
        \\<?xml version="1.0"?>
        \\<table id="1" ref="B2:D5"><tableColumns count="3"><tableColumn id="1" name="a"/><tableColumn id="2" name="b"/><tableColumn id="3" name="c"/></tableColumns></table>
    ;
    const out = try applyEditToTable(a, no_af, .row, 3, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "ref=\"B2:D6\"") != null);
}

// REL-A501: headerRowCount="0" tables permit top-row delete.
test "headerRowCount=0 permits top-row delete" {
    const a = testing.allocator;
    const headerless =
        \\<?xml version="1.0"?>
        \\<table id="1" ref="B2:D5" headerRowCount="0"><tableColumns count="3"><tableColumn id="1" name="a"/><tableColumn id="2" name="b"/><tableColumn id="3" name="c"/></tableColumns></table>
    ;
    const out = try applyEditToTable(a, headerless, .row, 2, .delete);
    defer a.free(out);
    // tl_row=2, br_row=5; delete row 2 → tl=2 (stays), br=4. Range B2:D4.
    try testing.expect(std.mem.indexOf(u8, out, "ref=\"B2:D4\"") != null);
}

// REL-A501 (companion): default (absent) headerRowCount still refuses
// top-row delete — the default is 1.
test "absent headerRowCount refuses top-row delete (defaults to 1)" {
    const a = testing.allocator;
    const default_hdr =
        \\<?xml version="1.0"?>
        \\<table id="1" ref="B2:D5"><tableColumns count="3"><tableColumn id="1" name="a"/><tableColumn id="2" name="b"/><tableColumn id="3" name="c"/></tableColumns></table>
    ;
    const r = applyEditToTable(a, default_hdr, .row, 2, .delete);
    try testing.expectError(error.TableHeaderRowDeleteUnsafe, r);
}

// REL-A502: synthetic <tableColumn name=> dodges existing-name collision.
test "synthetic insert avoids name collision with existing Column<id>" {
    const a = testing.allocator;
    // Existing ids 1,2,3; existing names include "Column4" already.
    // After insert at col 4 (D, inside B2:D5 → range becomes B2:E5),
    // synthetic id = max(1,2,3)+1 = 4. Naive name "Column4" collides
    // with the existing "Column4"; the fix should pick "Column4_2".
    const colliding =
        \\<?xml version="1.0"?>
        \\<table id="1" ref="B2:D5"><tableColumns count="3"><tableColumn id="1" name="a"/><tableColumn id="2" name="Column4"/><tableColumn id="3" name="c"/></tableColumns></table>
    ;
    const out = try applyEditToTable(a, colliding, .col, 3, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "ref=\"B2:E5\"") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<tableColumn id=\"4\" name=\"Column4_2\"/>") != null);
    // The original "Column4" survives unchanged.
    try testing.expect(std.mem.indexOf(u8, out, "<tableColumn id=\"2\" name=\"Column4\"/>") != null);
}

// REL-A502 (companion): no collision → uses bare Column<id>.
test "synthetic insert uses bare Column<id> when no collision" {
    const a = testing.allocator;
    const out = try applyEditToTable(a, sample_table, .col, 4, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<tableColumn id=\"4\" name=\"Column4\"/>") != null);
}

// REL-A505: <tableColumns count> attr disagreeing with range → MalformedTableXml.
test "tableColumns count disagreeing with range refuses MalformedTableXml" {
    const a = testing.allocator;
    // Range B2:D5 implies count=3; source claims count=2 (sparse / malformed).
    const sparse =
        \\<?xml version="1.0"?>
        \\<table id="1" ref="B2:D5"><tableColumns count="2"><tableColumn id="1" name="a"/><tableColumn id="2" name="b"/></tableColumns></table>
    ;
    const r = applyEditToTable(a, sparse, .col, 3, .insert);
    try testing.expectError(error.MalformedTableXml, r);
}

// Codex Ticket 651: <tableColumns count> attr matches range but
// the actual <tableColumn> child count differs. Round-1 fix only
// checked the declared count vs range and missed this branch.
test "tableColumns child count disagreeing with range refuses MalformedTableXml" {
    const a = testing.allocator;
    // Range B2:D5 implies 3 cols; declared count=3 agrees; only 2 actual children.
    const sparse_children =
        \\<?xml version="1.0"?>
        \\<table id="1" ref="B2:D5"><tableColumns count="3"><tableColumn id="1" name="a"/><tableColumn id="2" name="b"/></tableColumns></table>
    ;
    const r = applyEditToTable(a, sparse_children, .col, 3, .insert);
    try testing.expectError(error.MalformedTableXml, r);
}

// Codex Ticket 701: nested <tableColumn> descendants (e.g., inside
// an <extLst>) must not be miscounted as direct siblings. A flat
// scan would see 3 tags here while the actual direct child count
// is 2 — refuse before the rewrite walker emits broken bytes.
test "nested <tableColumn> descendant doesn't inflate child count" {
    const a = testing.allocator;
    // 2 direct children; first child has an <extLst> with a nested
    // x14:tableColumn-shaped element (we use the bare tag name to
    // exercise the flat-scan failure mode). count="3", range width
    // 3 (B2:D5), but only 2 direct children → must refuse.
    const nested =
        \\<?xml version="1.0"?>
        \\<table id="1" ref="B2:D5"><tableColumns count="3"><tableColumn id="1" name="a"><extLst><ext><tableColumn id="99" name="ghost"/></ext></extLst></tableColumn><tableColumn id="2" name="b"/></tableColumns></table>
    ;
    const r = applyEditToTable(a, nested, .col, 3, .insert);
    try testing.expectError(error.MalformedTableXml, r);
}

// REL-A509: scanMaxTableColumnId handles sparse / non-contiguous ids.
test "synthetic id is max+1 with sparse ids (1, 5, 99 → 100)" {
    const a = testing.allocator;
    // Range B2:D5 (3 cols); ids deliberately sparse.
    const sparse_ids =
        \\<?xml version="1.0"?>
        \\<table id="1" ref="B2:D5"><tableColumns count="3"><tableColumn id="1" name="a"/><tableColumn id="5" name="b"/><tableColumn id="99" name="c"/></tableColumns></table>
    ;
    const out = try applyEditToTable(a, sparse_ids, .col, 3, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<tableColumn id=\"100\" name=\"Column100\"/>") != null);
}

// REL-A509: sortState collapse-drop on full-range delete-match.
test "sortState drops on collapse" {
    const a = testing.allocator;
    // Single-column sortState that collapses on col-delete.
    const single_col_sort =
        \\<?xml version="1.0"?>
        \\<table id="1" ref="A1:C5"><autoFilter ref="A1:C5"/><sortState ref="B2:B5"><sortCondition ref="B2:B5"/></sortState><tableColumns count="3"><tableColumn id="1" name="x"/><tableColumn id="2" name="y"/><tableColumn id="3" name="z"/></tableColumns></table>
    ;
    const out = try applyEditToTable(a, single_col_sort, .col, 2, .delete);
    defer a.free(out);
    // Table range A1:C5 → A1:B5; sortState was B2:B5 → collapses (drop).
    try testing.expect(std.mem.indexOf(u8, out, "ref=\"A1:B5\"") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<sortState") == null);
    try testing.expect(std.mem.indexOf(u8, out, "</sortState>") == null);
}

// REL-A509: insert beyond the last column (target_pos == count) appends
// the synthetic tableColumn at the end.
test "col insert at right edge appends synthetic tableColumn at end" {
    const a = testing.allocator;
    // Insert at col 6 (F) — exactly one past br_col=5 (E) of C9:E10.
    // Per shiftSingleA1Col semantics: insert at col == br_col+1 with
    // is_br_corner=true keeps br_col unchanged because the gap is at
    // the position AFTER br. So we expect ref unchanged here, no
    // synthetic column added.
    //
    // Instead test the "insert AT br_col" case: col 5 (E). target_pos
    // = 5 - 3 = 2 (the third / last column position).
    const out = try applyEditToTable(a, sample_table, .col, 5, .insert);
    defer a.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "ref=\"C9:F10\"") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<tableColumns count=\"4\">") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<tableColumn id=\"4\" name=\"Column4\"/>") != null);
}

// REL-B524: TablePartRidIterator boundary check is exercised by the
// editor round-trip tests (the iterator lives in workbook.zig). The
// isolation test here is for table_edit only.
