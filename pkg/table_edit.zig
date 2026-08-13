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
//! `renameTableColumn` (C1) is the second entry point: renames one
//! `<tableColumn name>` and routes every `calculatedColumnFormula` /
//! `totalsRowFormula` body through the formula rewriter so
//! structured references (`T[Old]`, bare `[@Old]`) follow the
//! rename. The Editor path finishes the job with the header-cell
//! text and the workbook-wide formula sweep.
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
const store_mod = @import("store.zig");
const wbxml = @import("typed_parts/workbook_xml.zig");
const engine = @import("zlsx_formula");
const coords = @import("zlsx_refs");

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
    /// `renameTableColumn`: no `<tableColumn>` matches the old name
    /// under the symbol layer's fold rule.
    TableColumnNotFound,
    /// `renameTableColumn`: another column already carries the new
    /// name (fold rule — ECMA-376 §18.5.1.78 requires names unique
    /// within the parent table, and Excel compares them
    /// case-insensitively).
    TableColumnNameInUse,
    /// `renameTableColumn`: empty name, or a control byte XML
    /// cannot carry.
    InvalidTableColumnName,
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
            // Delegate, same as autoFilter above: one sortState
            // implementation serves the table, sheet-bare and
            // autoFilter-nested contexts.
            switch (axis) {
                .col => try sheet_edit.processSortStateTagCol(allocator, &out, src, t, idx_1based, ek, &i),
                .row => try sheet_edit.processSortStateTagRow(allocator, &out, src, t, idx_1based, ek, &i),
            }
        } else if (sheet_edit.matchTagAt(src, i, "sortCondition")) |t| {
            switch (axis) {
                .col => try sheet_edit.processSortConditionTagCol(allocator, &out, src, t, idx_1based, ek, &i),
                .row => try sheet_edit.processSortConditionTagRow(allocator, &out, src, t, idx_1based, ek, &i),
            }
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
    // Find the first REAL `<table` open tag (not `<tableColumn(s)>`,
    // and not a decoy inside a comment / CDATA / PI — the aware
    // scanner skips those; Codex #190 r1 F5). Scanner errors on
    // malformed non-element constructs read as "no table here".
    const hit = (wbxml.findTagOpen(src, 0, "table") catch return null) orelse return null;
    const attrs = src[hit.attrs_start..hit.attrs_end];
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
        .tag = .{ .start = hit.open_lt, .after_open = hit.after_tag_close },
        .tl_col = range.tl_col,
        .br_col = range.br_col,
        .tl_row = range.tl_row,
        .br_row = range.br_row,
        .header_row_count = hrc,
        .header_row_count_explicit_zero = hrc_explicit_zero,
    };
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
    // M0: this used to be a hand-rolled duplicate of
    // `pkg/sheet_edit.zig::formatColLetters`, with a comment asking for
    // a shared module once a third consumer appeared. `zlsx_refs` is
    // that module. Unchecked writer — the pre-M0 behaviour accepted
    // out-of-grid columns and only failed when the buffer ran out.
    const len = coords.writeColNumberLetters(buf, col_1based) catch
        return error.TableCoordinateOverflow;
    return buf[0..len];
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

// ─── column rename (C1: structured-ref rewriting) ───────────────

pub const RenameColumnResult = struct {
    /// The rewritten table part. Caller frees.
    bytes: []u8,
    /// 0-based position of the renamed column among the table's
    /// direct `<tableColumn>` children — also its offset from the
    /// range's left edge, which the width check below guarantees.
    column_pos: u32,
    /// 1-based grid coordinates of the header cell whose text
    /// mirrors the column name (ECMA-376 §18.5.1.3: the header row
    /// cell IS the column name); null for a header-less table
    /// (`headerRowCount="0"`).
    header_row: ?u32,
    header_col: ?u32,
    /// The table's declared range (1-based, inclusive), for the
    /// caller's producer-cell scoping: bare structured refs in cell
    /// formulas bind to the table whose range contains the cell.
    tl_row: u32,
    tl_col: u32,
    br_row: u32,
    br_col: u32,
    /// Decoded formula-visible table name (`displayName`, falling
    /// back to `name`) — what structured references spell and what
    /// the caller passes to the workbook-wide formula rewrite.
    /// Caller frees.
    display_name: []u8,

    pub fn deinit(self: *RenameColumnResult, allocator: Allocator) void {
        allocator.free(self.bytes);
        allocator.free(self.display_name);
        self.* = undefined;
    }
};

/// Rename the `<tableColumn>` whose decoded `name` matches
/// `old_name` (the symbol layer's fold rule — same as formula
/// resolution) to `new_name`, and rewrite every
/// `calculatedColumnFormula` / `totalsRowFormula` body in the part
/// so structured references follow the rename. `old_name` /
/// `new_name` are decoded plain text; XML escaping happens at the
/// splice boundaries. Refusals all precede byte work.
///
/// The caller (Editor path) is responsible for the two mutations
/// this function cannot see: the header CELL text in the host
/// sheet (`header_row`/`header_col`) and the workbook-wide formula
/// rewrite (`display_name` + the same edit).
pub fn renameTableColumn(
    allocator: Allocator,
    src: []const u8,
    old_name: []const u8,
    new_name: []const u8,
) Error!RenameColumnResult {
    if (old_name.len == 0) return error.InvalidTableColumnName;
    try validateNewColumnName(new_name);

    const hdr = parseTableHeader(src) orelse return error.MalformedTableXml;
    const table_attrs = src[hdr.tag.start + "<table".len .. hdr.tag.after_open - 1];
    const display_raw = blk: {
        if (getAttr(table_attrs, "displayName")) |v| {
            if (v.len > 0) break :blk v;
        }
        const v = getAttr(table_attrs, "name") orelse return error.MalformedTableXml;
        if (v.len == 0) return error.MalformedTableXml;
        break :blk v;
    };
    // STRING-carrier decode (entities + ST_Xstring) — the codec the
    // engine resolves these attrs with (Codex #190 r1 F1). A part
    // spelling `displayName="Sales_x0041_"` names the table SalesA.
    const display_name = try decodeNameAttr(allocator, .table_name, display_raw);
    errdefer allocator.free(display_name);

    const tc = findTableColumns(src) orelse return error.MalformedTableXml;

    // Pass 1 (pre-flight): locate the target, refuse ambiguity and
    // collisions before any output exists.
    var target_pos: ?u32 = null;
    var count: u32 = 0;
    {
        var it = ColumnIter{ .src = src, .k = tc.body_start, .end = tc.close_pos };
        while (try it.next()) |col| {
            count += 1;
            const attrs = src[col.hit.attrs_start..col.hit.attrs_end];
            const name_raw = getAttr(attrs, "name") orelse return error.MalformedTableXml;
            const decoded = try decodeNameAttr(allocator, .table_column_name, name_raw);
            defer allocator.free(decoded);
            if (try foldedEql(allocator, decoded, old_name)) {
                // Two columns matching the old name means the part
                // already violates §18.5.1.78 uniqueness — corrupt
                // input, not an ambiguity to resolve silently.
                if (target_pos != null) return error.MalformedTableXml;
                target_pos = col.pos;
                // The renamed column itself never collides with its
                // own new spelling (case-respell is a legal rename).
                continue;
            }
            if (try foldedEql(allocator, decoded, new_name)) {
                return error.TableColumnNameInUse;
            }
        }
    }
    const pos = target_pos orelse return error.TableColumnNotFound;

    // Width agreement (the REL-A505 discipline): the header-cell
    // coordinate is tl_col + pos, meaningful only when the column
    // list and the declared range agree.
    const width = hdr.br_col -| hdr.tl_col + 1;
    if (count != width) return error.MalformedTableXml;

    // The authored attr value: ST_Xstring, then XML escaping (both
    // quote characters included, so the splice is safe inside either
    // quote style).
    const encoded_new = try engine.decode.encodeAuthoredString(allocator, new_name);
    defer allocator.free(encoded_new);

    // Pass 2: emit. The SAME iterator that validated pass 1 drives
    // the splice, so the two passes cannot disagree about which
    // elements are direct columns. Every byte between and around
    // columns — comments included — is copied verbatim.
    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(allocator);

    var cursor: usize = 0;
    var it = ColumnIter{ .src = src, .k = tc.body_start, .end = tc.close_pos };
    while (try it.next()) |col| {
        try out.appendSlice(allocator, src[cursor..col.hit.open_lt]);
        if (col.pos == pos) {
            try emitRenamedOpenTag(allocator, &out, src, col.hit, encoded_new, old_name);
        } else {
            try out.appendSlice(allocator, src[col.hit.open_lt..col.hit.after_tag_close]);
        }
        if (col.body_end) |body_end| {
            try emitColumnBodyWithRewrites(
                allocator,
                &out,
                src,
                col.hit.after_tag_close,
                body_end,
                display_name,
                old_name,
                new_name,
            );
            try out.appendSlice(allocator, src[body_end..col.end]);
        }
        cursor = col.end;
    }
    try out.appendSlice(allocator, src[cursor..]);

    const has_header = hdr.header_row_count >= 1;
    return .{
        .bytes = try out.toOwnedSlice(allocator),
        .column_pos = pos,
        .header_row = if (has_header) hdr.tl_row else null,
        .header_col = if (has_header) hdr.tl_col + pos else null,
        .tl_row = hdr.tl_row,
        .tl_col = hdr.tl_col,
        .br_row = hdr.br_row,
        .br_col = hdr.br_col,
        .display_name = display_name,
    };
}

/// The raw `displayName` attribute (falling back to `name`) of a
/// table part's `<table>` open tag — the spelling structured
/// references resolve against, still XML-encoded. Null when the
/// part has no parseable `<table>` tag or neither attribute.
/// Public for the Workbook's locate-table-by-name scan.
pub fn tableDisplayNameRaw(src: []const u8) ?[]const u8 {
    const hdr = parseTableHeader(src) orelse return null;
    const attrs = src[hdr.tag.start + "<table".len .. hdr.tag.after_open - 1];
    if (getAttr(attrs, "displayName")) |v| {
        if (v.len > 0) return v;
    }
    const v = getAttr(attrs, "name") orelse return null;
    return if (v.len > 0) v else null;
}

/// The `<tableColumns>` block: body span and close position. Aware
/// scan on both ends — a decoy inside a comment is not a block.
const TableColumnsSpan = struct {
    body_start: usize,
    close_pos: usize,
};

fn findTableColumns(src: []const u8) ?TableColumnsSpan {
    const hit = (wbxml.findTagOpen(src, 0, "tableColumns") catch return null) orelse return null;
    if (hit.self_closing) return null;
    const close = (wbxml.findClosingTag(src, hit.after_tag_close, "</tableColumns>") catch return null) orelse return null;
    return .{ .body_start = hit.after_tag_close, .close_pos = close };
}

/// Direct `<tableColumn>` children of a `<tableColumns>` body, in
/// document order with 0-based positions. Advances past each
/// column's FULL extent (Codex Ticket 701: a descendant
/// `<tableColumn>` nested in an `<extLst>` extension must not be
/// double-counted as a sibling), skipping comment/CDATA/PI decoys
/// (Codex #190 r1 F5). A column whose extent cannot be established
/// inside the block is a hard `MalformedTableXml`, not a guess.
const ColumnIter = struct {
    src: []const u8,
    k: usize,
    end: usize,
    pos: u32 = 0,

    const Item = struct {
        hit: wbxml.TagHit,
        pos: u32,
        /// Just past the element (self-closing open tag, or the
        /// `</tableColumn>` close).
        end: usize,
        /// The close tag's start for the non-self-closing form.
        body_end: ?usize,
    };

    fn next(self: *ColumnIter) Error!?Item {
        const hit = (wbxml.findTagOpen(self.src[0..self.end], self.k, "tableColumn") catch
            return error.MalformedTableXml) orelse return null;
        const extent = try columnExtent(self.src, hit, self.end);
        self.k = extent.end;
        const item: Item = .{
            .hit = hit,
            .pos = self.pos,
            .end = extent.end,
            .body_end = extent.body_end,
        };
        self.pos += 1;
        return item;
    }
};

const ColumnExtent = struct { end: usize, body_end: ?usize };

/// Full extent of one `<tableColumn>`, depth-tracked: a nested
/// non-self-closing `<tableColumn>` inside an extension block must
/// not terminate the outer column at the inner close (the C2b
/// depth-tracking lesson, raised again as Codex #190 r1 F5). A
/// missing close inside `bound` refuses — silently ending the
/// column at `</tableColumns>` spliced attrs into the wrong
/// element.
fn columnExtent(src: []const u8, ct: wbxml.TagHit, bound: usize) Error!ColumnExtent {
    if (ct.self_closing) return .{ .end = ct.after_tag_close, .body_end = null };
    const cend = "</tableColumn>";
    var depth: usize = 0;
    var i = ct.after_tag_close;
    while (i < bound) {
        const lt = std.mem.indexOfScalarPos(u8, src, i, '<') orelse break;
        if (lt >= bound) break;
        const skip = wbxml.skipNonElement(src[0..bound], lt) catch return error.MalformedTableXml;
        if (skip != lt) {
            i = skip;
            continue;
        }
        if (std.mem.startsWith(u8, src[lt..bound], cend)) {
            if (depth == 0) return .{ .end = lt + cend.len, .body_end = lt };
            depth -= 1;
            i = lt + cend.len;
            continue;
        }
        if (sheet_edit.matchTagAt(src, lt, "tableColumn")) |nested| {
            if (nested.after_open > bound) return error.MalformedTableXml;
            const attrs_full = src[nested.start + "<tableColumn".len .. nested.after_open - 1];
            const trimmed = std.mem.trimEnd(u8, attrs_full, " \t\r\n");
            const self_closing = trimmed.len > 0 and trimmed[trimmed.len - 1] == '/';
            if (!self_closing) depth += 1;
            i = nested.after_open;
            continue;
        }
        i = lt + 1;
    }
    return error.MalformedTableXml;
}

/// Copy one `<tableColumn>` body, rewriting the inner text of every
/// `<calculatedColumnFormula>` / `<totalsRowFormula>` through the
/// formula rewriter (decode-in / escape-out). The part's own
/// formulas are the table's producers, so BARE structured refs
/// (`[Old]`, `[@Old]`) scope to this table via `owning_table`. A
/// body whose rewrite is byte-identical keeps its original bytes —
/// entity spellings included.
fn emitColumnBodyWithRewrites(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    from: usize,
    to: usize,
    display_name: []const u8,
    old_name: []const u8,
    new_name: []const u8,
) Error!void {
    var i = from;
    while (i < to) {
        const lt = std.mem.indexOfScalarPos(u8, src, i, '<') orelse break;
        if (lt >= to) break;
        try out.appendSlice(allocator, src[i..lt]);
        i = lt;

        // Comments / CDATA / PIs are copied verbatim, never matched:
        // a commented `<calculatedColumnFormula>` is not a formula
        // (Codex #190 r1 F5).
        const skip = wbxml.skipNonElement(src[0..to], lt) catch return error.MalformedTableXml;
        if (skip != lt) {
            try out.appendSlice(allocator, src[lt..skip]);
            i = skip;
            continue;
        }

        const t: ?TagOpen, const close_tag: []const u8 = blk: {
            if (sheet_edit.matchTagAt(src, lt, "calculatedColumnFormula")) |t| {
                break :blk .{ t, "</calculatedColumnFormula>" };
            }
            if (sheet_edit.matchTagAt(src, lt, "totalsRowFormula")) |t| {
                break :blk .{ t, "</totalsRowFormula>" };
            }
            break :blk .{ null, "" };
        };
        const tag = t orelse {
            try out.append(allocator, '<');
            i += 1;
            continue;
        };

        const attrs_full = src[tag.start .. tag.after_open - 1];
        const trimmed = std.mem.trimEnd(u8, attrs_full, " \t\r\n");
        if (trimmed.len > 0 and trimmed[trimmed.len - 1] == '/') {
            // Self-closing formula element: no body to rewrite.
            try out.appendSlice(allocator, src[tag.start..tag.after_open]);
            i = tag.after_open;
            continue;
        }
        const close_pos = ((wbxml.findClosingTag(src[0..to], tag.after_open, close_tag) catch
            return error.MalformedTableXml)) orelse return error.MalformedTableXml;

        try out.appendSlice(allocator, src[tag.start..tag.after_open]);
        const body = src[tag.after_open..close_pos];
        try appendRewrittenFormula(allocator, out, body, display_name, old_name, new_name);
        try out.appendSlice(allocator, src[close_pos .. close_pos + close_tag.len]);
        i = close_pos + close_tag.len;
    }
    if (i < to) try out.appendSlice(allocator, src[i..to]);
}

/// Emit the target column's open tag with its `name` attribute
/// value replaced — and, when the column carries a `uniqueName`
/// that mirrors the old name (XML-mapped tables keep the two
/// synchronized; Codex #190 r1 F7), that value too. The splice is
/// quote-style-aware and refuses when the validated attribute
/// cannot be located again (a silent verbatim copy here reported a
/// successful rename that never happened — Codex #190 r1 F2).
fn emitRenamedOpenTag(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    hit: wbxml.TagHit,
    encoded_new: []const u8,
    old_name: []const u8,
) Error!void {
    const name_span = findAttrValueSpan(src, hit.attrs_start, hit.attrs_end, "name") orelse
        return error.MalformedTableXml;

    var unique_span: ?AttrSpan = null;
    if (findAttrValueSpan(src, hit.attrs_start, hit.attrs_end, "uniqueName")) |us| {
        const decoded = try decodeNameAttr(allocator, .table_column_name, src[us.val_start..us.val_end]);
        defer allocator.free(decoded);
        if (try foldedEql(allocator, decoded, old_name)) unique_span = us;
    }

    var spans: [2]AttrSpan = undefined;
    var n: usize = 0;
    spans[n] = name_span;
    n += 1;
    if (unique_span) |us| {
        spans[n] = us;
        n += 1;
        if (spans[0].val_start > spans[1].val_start) std.mem.swap(AttrSpan, &spans[0], &spans[1]);
    }

    var cursor = hit.open_lt;
    for (spans[0..n]) |s| {
        try out.appendSlice(allocator, src[cursor..s.val_start]);
        try out.appendSlice(allocator, encoded_new);
        cursor = s.val_end;
    }
    try out.appendSlice(allocator, src[cursor..hit.after_tag_close]);
}

const AttrSpan = struct { val_start: usize, val_end: usize };

/// The VALUE span of `attr_name` inside an open tag's attribute
/// region, in absolute `src` offsets. Mirrors `getAttr`'s walk —
/// both quote styles, whitespace around `=` — so anything the
/// pre-flight found, the splice finds.
fn findAttrValueSpan(
    src: []const u8,
    attrs_start: usize,
    attrs_end: usize,
    attr_name: []const u8,
) ?AttrSpan {
    var i = attrs_start;
    while (i < attrs_end) {
        while (i < attrs_end and std.ascii.isWhitespace(src[i])) i += 1;
        if (i >= attrs_end) break;
        const name_start = i;
        while (i < attrs_end and src[i] != '=' and !std.ascii.isWhitespace(src[i])) i += 1;
        const this_name = src[name_start..i];
        while (i < attrs_end and (src[i] == '=' or std.ascii.isWhitespace(src[i]))) i += 1;
        if (i >= attrs_end or (src[i] != '"' and src[i] != '\'')) break;
        const quote = src[i];
        i += 1;
        const val_start = i;
        while (i < attrs_end and src[i] != quote) i += 1;
        if (i >= attrs_end) break; // unterminated value
        const val_end = i;
        i += 1;
        if (std.mem.eql(u8, this_name, attr_name)) {
            return .{ .val_start = val_start, .val_end = val_end };
        }
    }
    return null;
}

/// STRING-carrier decode (XML entities + ST_Xstring) — the codec
/// the engine resolves table/column attrs with. A malformed value
/// is a part the engine would refuse; surface it as such.
fn decodeNameAttr(
    allocator: Allocator,
    site: engine.decode.Site,
    raw: []const u8,
) Error![]u8 {
    return engine.decode.decodeAt(allocator, site, raw) catch |err| {
        if (err == error.OutOfMemory) return error.OutOfMemory;
        return error.MalformedTableXml;
    };
}

/// New-name validation (Codex #190 r1 F6): valid UTF-8 and at most
/// 255 Unicode scalars (Excel's `tableColumn@name` cap). Control
/// bytes need no refusal — `encodeAuthoredString` spells them as
/// `_xHHHH_`, the same way Excel does.
fn validateNewColumnName(name: []const u8) Error!void {
    if (name.len == 0) return error.InvalidTableColumnName;
    if (!std.unicode.utf8ValidateSlice(name)) return error.InvalidTableColumnName;
    const scalars = std.unicode.utf8CountCodepoints(name) catch return error.InvalidTableColumnName;
    if (scalars > 255) return error.InvalidTableColumnName;
}

/// Byte contract (same as the cell-formula sweep settled in #188):
/// an UNCHANGED body keeps its source bytes, entity spellings
/// included; a CHANGED body is re-escaped from the decoded rewrite,
/// so an exotic entity spelling elsewhere in that one formula
/// normalizes. Semantic content is identical either way.
fn appendRewrittenFormula(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    body: []const u8,
    display_name: []const u8,
    old_name: []const u8,
    new_name: []const u8,
) Error!void {
    if (body.len == 0) return;
    const decoded = try store_mod.decodeXmlEntities(allocator, body);
    defer allocator.free(decoded);
    const rewritten = xlsx.formula_rewriter.rewriteFormula(allocator, decoded, .{
        .owning_table = display_name,
        .edit = .{ .rename_table_column = .{
            .table = display_name,
            .old = old_name,
            .new = new_name,
        } },
    }) catch |err| switch (err) {
        error.OutOfMemory => return error.OutOfMemory,
        // Old/new emptiness is refused at this function's entry.
        error.InvalidEdit => unreachable,
    };
    defer allocator.free(rewritten);
    if (std.mem.eql(u8, rewritten, decoded)) {
        // No-op rewrite: keep the source bytes, original entity
        // spellings included.
        try out.appendSlice(allocator, body);
        return;
    }
    try appendTextEscaped(allocator, out, rewritten);
}

/// Fold-equality on decoded names — the formula engine's matching
/// rule (`SymbolTable.fold` before every lookup). Invalid UTF-8
/// matches nothing the symbol layer admitted.
fn foldedEql(allocator: Allocator, a: []const u8, b: []const u8) Error!bool {
    return xlsx.casefold.eqlFolded(allocator, a, b) catch |err| {
        if (err == error.OutOfMemory) return error.OutOfMemory;
        return false;
    };
}

/// Element-content escape (3-entity, matching the byte-stable
/// contract everywhere else formulas are spliced).
fn appendTextEscaped(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    text: []const u8,
) Error!void {
    for (text) |b| {
        switch (b) {
            '&' => try out.appendSlice(allocator, "&amp;"),
            '<' => try out.appendSlice(allocator, "&lt;"),
            '>' => try out.appendSlice(allocator, "&gt;"),
            else => try out.append(allocator, b),
        }
    }
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

test "sortCondition ref shifts with its parent sortState" {
    const a = testing.allocator;
    const with_sort =
        \\<?xml version="1.0"?>
        \\<table id="1" ref="A1:C5"><autoFilter ref="A1:C5"><filterColumn colId="0"/></autoFilter><sortState ref="A2:C5"><sortCondition ref="A2:A5"/></sortState><tableColumns count="3"><tableColumn id="1" name="x"/><tableColumn id="2" name="y"/><tableColumn id="3" name="z"/></tableColumns></table>
    ;
    const out = try applyEditToTable(a, with_sort, .row, 1, .insert);
    defer a.free(out);
    // The parent shifts A2:C5 → A3:C6, so the child sort key must
    // shift A2:A5 → A3:A6 in step. Leaving it behind points the sort
    // at the wrong rows with no error surfaced.
    try testing.expect(std.mem.indexOf(u8, out, "<sortCondition ref=\"A3:A6\"") != null);
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

// ─── fuzz target ────────────────────────────────────────────────────
//
// See the note in sheet_edit.zig. This walker additionally mutates
// `<tableColumns count=>` and synthesises column entries, so it has
// arithmetic sheet_edit does not.

fn fuzzTableEditTarget(_: void, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();
    var smith_buf: [4096]u8 = undefined;
    const input = smith_buf[0..smith.slice(&smith_buf)];

    var arena = std.heap.ArenaAllocator.init(std.testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    for ([_]u32{ 1, 2, 5 }) |idx| {
        inline for (.{ .row, .col }) |axis| {
            inline for (.{ .insert, .delete }) |kind| {
                if (applyEditToTable(a, input, axis, idx, kind)) |out| {
                    a.free(out);
                } else |_| {}
            }
        }
    }
}

test "fuzz: applyEditToTable never crashes on adversarial XML" {
    try std.testing.fuzz({}, fuzzTableEditTarget, .{
        .corpus = &[_][]const u8{
            "",
            "<table/>",
            "<table ref=\"A1:C5\"/>",
            "<table id=\"1\" ref=\"A1:C5\"><tableColumns count=\"3\"><tableColumn id=\"1\" name=\"x\"/></tableColumns></table>",
            // count= disagreeing with the actual child count, both ways.
            "<table id=\"1\" ref=\"A1:C5\"><tableColumns count=\"99\"><tableColumn id=\"1\" name=\"x\"/></tableColumns></table>",
            "<table id=\"1\" ref=\"A1:C5\"><tableColumns count=\"0\"><tableColumn id=\"1\" name=\"x\"/></tableColumns></table>",
            "<table id=\"1\" ref=\"A1:C5\"><tableColumns count=\"-1\"/></table>",
            // Nested sortState / sortCondition — the #125 additions.
            "<table id=\"1\" ref=\"A1:C5\"><sortState ref=\"A2:C5\"><sortCondition ref=\"A2:A5\"/></sortState></table>",
            "<table id=\"1\" ref=\"A1:C5\"><autoFilter ref=\"A1:C5\"><filterColumn colId=\"0\"/></autoFilter></table>",
            // headerRowCount drives the refusal branches.
            "<table id=\"1\" ref=\"A1:C5\" headerRowCount=\"0\"/>",
            "<table id=\"1\" ref=\"A1:C5\" headerRowCount=\"9\"/>",
            "<table id=\"1\" ref=\"A1:A1\"/>",
            // Truncations.
            "<table id=\"1\" ref=",
            "<table id=\"1\" ref=\"A1:C5\"><tableColumns",
            "<table id=\"1\" ref=\"A1:C5\"><tableColumn id=",
            "<table id=\"1\" ref=\"A1:C5\"><sortState",
            "<table",
            // Coordinate extremes.
            "<table id=\"1\" ref=\"XFD1048576:XFD1048576\"/>",
            "<table id=\"1\" ref=\"A0:A0\"/>",
            "<table id=\"1\" ref=\"4294967295\"/>",
        },
    });
}

// ─── renameTableColumn ──────────────────────────────────────────

const rename_sample =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="2" name="Sales" displayName="Sales" ref="B2:D6">
    \\<autoFilter ref="B2:D6"/>
    \\<tableColumns count="3">
    \\<tableColumn id="1" name="R&amp;D"/>
    \\<tableColumn id="2" name="Total"><calculatedColumnFormula>SUM(Sales[R&amp;D])*2</calculatedColumnFormula></tableColumn>
    \\<tableColumn id="3" name="Note"><totalsRowFormula>COUNTA([R&amp;D])</totalsRowFormula></tableColumn>
    \\</tableColumns>
    \\<tableStyleInfo name="TS"/>
    \\</table>
;

test "renameTableColumn rewrites the name attr and both formula sites" {
    const a = testing.allocator;
    var r = try renameTableColumn(a, rename_sample, "R&D", "Budget");
    defer r.deinit(a);

    // The attr splice, the qualified calculatedColumnFormula ref,
    // and the BARE totalsRowFormula ref (scoped through
    // owning_table) all follow the rename; every other byte is
    // untouched.
    const expected =
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="2" name="Sales" displayName="Sales" ref="B2:D6">
        \\<autoFilter ref="B2:D6"/>
        \\<tableColumns count="3">
        \\<tableColumn id="1" name="Budget"/>
        \\<tableColumn id="2" name="Total"><calculatedColumnFormula>SUM(Sales[Budget])*2</calculatedColumnFormula></tableColumn>
        \\<tableColumn id="3" name="Note"><totalsRowFormula>COUNTA([Budget])</totalsRowFormula></tableColumn>
        \\</tableColumns>
        \\<tableStyleInfo name="TS"/>
        \\</table>
    ;
    try testing.expectEqualStrings(expected, r.bytes);
    try testing.expectEqualStrings("Sales", r.display_name);
    try testing.expectEqual(@as(u32, 0), r.column_pos);
    // Header cell: top row of B2:D6, first column.
    try testing.expectEqual(@as(?u32, 2), r.header_row);
    try testing.expectEqual(@as(?u32, 2), r.header_col);
}

test "renameTableColumn escapes the new name at both boundaries" {
    const a = testing.allocator;
    var r = try renameTableColumn(a, rename_sample, "R&D", "P&L");
    defer r.deinit(a);
    // Attr boundary: XML attr escape. Formula boundary: entity
    // escape of the rewritten body — where the `&` also makes
    // `columnNeedsBrackets` bracket the part, the canonical
    // printer's spelling for punctuation-carrying names.
    try testing.expect(std.mem.indexOf(u8, r.bytes, "<tableColumn id=\"1\" name=\"P&amp;L\"/>") != null);
    try testing.expect(std.mem.indexOf(u8, r.bytes, "SUM(Sales[[P&amp;L]])*2") != null);
    try testing.expect(std.mem.indexOf(u8, r.bytes, "COUNTA([[P&amp;L]])") != null);
}

test "renameTableColumn matches and collides under the fold rule" {
    const a = testing.allocator;
    // Old-name match is folded: `r&d` finds `R&amp;D`.
    var r = try renameTableColumn(a, rename_sample, "r&d", "Budget");
    defer r.deinit(a);
    try testing.expect(std.mem.indexOf(u8, r.bytes, "name=\"Budget\"") != null);

    // Collision with ANOTHER column is folded too.
    try testing.expectError(
        error.TableColumnNameInUse,
        renameTableColumn(a, rename_sample, "R&D", "total"),
    );

    // A case-respell of the SAME column is a legal rename, not a
    // self-collision.
    var respell = try renameTableColumn(a, rename_sample, "total", "TOTAL");
    defer respell.deinit(a);
    try testing.expect(std.mem.indexOf(u8, respell.bytes, "name=\"TOTAL\"") != null);
    try testing.expectEqual(@as(u32, 1), respell.column_pos);
}

test "renameTableColumn refusals" {
    const a = testing.allocator;
    try testing.expectError(
        error.TableColumnNotFound,
        renameTableColumn(a, rename_sample, "Missing", "X"),
    );
    try testing.expectError(
        error.InvalidTableColumnName,
        renameTableColumn(a, rename_sample, "Total", ""),
    );
    try testing.expectError(
        error.InvalidTableColumnName,
        renameTableColumn(a, rename_sample, "", "X"),
    );
    // Invalid UTF-8 can never match or author a STRING carrier.
    try testing.expectError(
        error.InvalidTableColumnName,
        renameTableColumn(a, rename_sample, "Total", "a\xffb"),
    );
    // Excel caps tableColumn@name at 255 Unicode scalars.
    const long = "x" ** 256;
    try testing.expectError(
        error.InvalidTableColumnName,
        renameTableColumn(a, rename_sample, "Total", long),
    );
    // A control byte is NOT refused — it authors as `_xHHHH_`, the
    // spelling Excel itself uses.
    var ctl = try renameTableColumn(a, rename_sample, "Total", "a\x01b");
    ctl.deinit(a);
    // Column list disagreeing with the declared range width: the
    // header-cell coordinate would be meaningless.
    const skewed =
        "<table id=\"1\" name=\"T\" displayName=\"T\" ref=\"A1:C5\">" ++
        "<tableColumns count=\"2\"><tableColumn id=\"1\" name=\"A\"/>" ++
        "<tableColumn id=\"2\" name=\"B\"/></tableColumns></table>";
    try testing.expectError(
        error.MalformedTableXml,
        renameTableColumn(a, skewed, "A", "X"),
    );
}

test "renameTableColumn: header-less table returns null header cell" {
    const a = testing.allocator;
    const headerless =
        "<table id=\"1\" name=\"T\" displayName=\"T\" ref=\"A1:B3\" headerRowCount=\"0\">" ++
        "<tableColumns count=\"2\"><tableColumn id=\"1\" name=\"In\"/>" ++
        "<tableColumn id=\"2\" name=\"Out\"/></tableColumns></table>";
    var r = try renameTableColumn(a, headerless, "Out", "Result");
    defer r.deinit(a);
    try testing.expectEqual(@as(?u32, null), r.header_row);
    try testing.expectEqual(@as(?u32, null), r.header_col);
    try testing.expectEqual(@as(u32, 1), r.column_pos);
    try testing.expect(std.mem.indexOf(u8, r.bytes, "name=\"Result\"") != null);
}

test "renameTableColumn: nested extLst tableColumn is not a sibling" {
    const a = testing.allocator;
    // Ticket-701 shape: a descendant <tableColumn> inside an
    // extension block must not shift sibling positions — the
    // renamed column here is at pos 1, header col B.
    const nested =
        "<table id=\"1\" name=\"T\" displayName=\"T\" ref=\"A1:B3\">" ++
        "<tableColumns count=\"2\">" ++
        "<tableColumn id=\"1\" name=\"A\"><extLst><ext><tableColumn id=\"9\" name=\"Decoy\"/></ext></extLst></tableColumn>" ++
        "<tableColumn id=\"2\" name=\"B\"/>" ++
        "</tableColumns></table>";
    var r = try renameTableColumn(a, nested, "B", "Renamed");
    defer r.deinit(a);
    try testing.expectEqual(@as(u32, 1), r.column_pos);
    try testing.expectEqual(@as(?u32, 2), r.header_col);
    // The decoy is untouched; the sibling is renamed.
    try testing.expect(std.mem.indexOf(u8, r.bytes, "name=\"Decoy\"") != null);
    try testing.expect(std.mem.indexOf(u8, r.bytes, "name=\"Renamed\"") != null);
}

test "renameTableColumn: ST_Xstring names decode and author through the engine codec" {
    const a = testing.allocator;
    // The engine reads `name="B_x0042_"` as column `BB` — the rename
    // must match on the decoded form (Codex #190 r1 F1).
    const xstr_table =
        "<table id=\"1\" name=\"T\" displayName=\"Sales_x0041_\" ref=\"A1:B3\">" ++
        "<tableColumns count=\"2\"><tableColumn id=\"1\" name=\"B_x0042_\"/>" ++
        "<tableColumn id=\"2\" name=\"Out\"><calculatedColumnFormula>SalesA[BB]*2</calculatedColumnFormula></tableColumn>" ++
        "</tableColumns></table>";
    var r = try renameTableColumn(a, xstr_table, "BB", "New");
    defer r.deinit(a);
    try testing.expectEqualStrings("SalesA", r.display_name);
    try testing.expect(std.mem.indexOf(u8, r.bytes, "name=\"New\"") != null);
    try testing.expect(std.mem.indexOf(u8, r.bytes, "SalesA[New]*2") != null);

    // Authoring is the inverse codec: a new name that LOOKS like an
    // escape is spelled `_x005F_…` so it round-trips.
    var r2 = try renameTableColumn(a, xstr_table, "Out", "_x0041_");
    defer r2.deinit(a);
    try testing.expect(std.mem.indexOf(u8, r2.bytes, "name=\"_x005F_x0041_\"") != null);
}

test "renameTableColumn: single-quoted and spaced attrs still splice" {
    const a = testing.allocator;
    // getAttr accepts both quote styles and whitespace around `=`;
    // the splice must find the same value the pre-flight found — a
    // silent verbatim copy reported success without renaming
    // (Codex #190 r1 F2).
    const quoted =
        "<table id=\"1\" name=\"T\" displayName=\"T\" ref=\"A1:B3\">" ++
        "<tableColumns count=\"2\"><tableColumn id='1' name = 'Old'/>" ++
        "<tableColumn id=\"2\" name=\"Out\"/></tableColumns></table>";
    var r = try renameTableColumn(a, quoted, "Old", "New");
    defer r.deinit(a);
    try testing.expect(std.mem.indexOf(u8, r.bytes, "name = 'New'") != null);
    try testing.expect(std.mem.indexOf(u8, r.bytes, "'Old'") == null);
}

test "renameTableColumn: uniqueName mirroring the old name follows the rename" {
    const a = testing.allocator;
    // XML-mapped tables keep a mirrored uniqueName synchronized
    // (Codex #190 r1 F7); a DISTINCT uniqueName is someone's
    // binding and stays.
    const mapped =
        "<table id=\"1\" name=\"T\" displayName=\"T\" ref=\"A1:B3\">" ++
        "<tableColumns count=\"2\">" ++
        "<tableColumn id=\"1\" name=\"Old\" uniqueName=\"Old\"/>" ++
        "<tableColumn id=\"2\" name=\"Out\" uniqueName=\"2\"/>" ++
        "</tableColumns></table>";
    var r = try renameTableColumn(a, mapped, "Old", "New");
    defer r.deinit(a);
    try testing.expect(std.mem.indexOf(u8, r.bytes, "name=\"New\" uniqueName=\"New\"") != null);
    var r2 = try renameTableColumn(a, mapped, "Out", "Result");
    defer r2.deinit(a);
    try testing.expect(std.mem.indexOf(u8, r2.bytes, "name=\"Result\" uniqueName=\"2\"") != null);
}

test "renameTableColumn: comment decoys are not markup" {
    const a = testing.allocator;
    // A commented `<table>` before the root must not become the
    // header, and a commented `<calculatedColumnFormula>` must not
    // be rewritten (Codex #190 r1 F5).
    const decoyed =
        "<!-- <table id=\"9\" ref=\"Z9:Z9\"> -->" ++
        "<table id=\"1\" name=\"T\" displayName=\"T\" ref=\"A1:B3\">" ++
        "<tableColumns count=\"2\">" ++
        "<tableColumn id=\"1\" name=\"Old\"/>" ++
        "<tableColumn id=\"2\" name=\"Out\">" ++
        "<!-- <calculatedColumnFormula>T[Old]</calculatedColumnFormula> -->" ++
        "<calculatedColumnFormula>T[Old]*2</calculatedColumnFormula></tableColumn>" ++
        "</tableColumns></table>";
    var r = try renameTableColumn(a, decoyed, "Old", "New");
    defer r.deinit(a);
    // Real header (A1:B3, not Z9): header cell A1.
    try testing.expectEqual(@as(?u32, 1), r.header_row);
    try testing.expectEqual(@as(?u32, 1), r.header_col);
    // The commented spelling is untouched; the real formula follows.
    try testing.expect(std.mem.indexOf(u8, r.bytes, "<!-- <calculatedColumnFormula>T[Old]</calculatedColumnFormula> -->") != null);
    try testing.expect(std.mem.indexOf(u8, r.bytes, "<calculatedColumnFormula>T[New]*2</calculatedColumnFormula>") != null);
}

test "renameTableColumn: nested non-self-closing decoy and missing close" {
    const a = testing.allocator;
    // Depth-tracking: a nested NON-self-closing `<tableColumn>`
    // inside an extension must not terminate the outer column at
    // the inner close (Codex #190 r1 F5).
    // The formula sits AFTER the nested decoy, still inside the
    // OUTER column: a walk that ends the outer column at the
    // decoy's close leaves it in a "gap" the splice copies
    // verbatim, so the rewrite silently vanishes.
    const nested =
        "<table id=\"1\" name=\"T\" displayName=\"T\" ref=\"A1:B3\">" ++
        "<tableColumns count=\"2\">" ++
        "<tableColumn id=\"1\" name=\"A\"><extLst><ext>" ++
        "<tableColumn id=\"9\" name=\"Decoy\"></tableColumn>" ++
        "</ext></extLst>" ++
        "<calculatedColumnFormula>T[B]*2</calculatedColumnFormula></tableColumn>" ++
        "<tableColumn id=\"2\" name=\"B\"/>" ++
        "</tableColumns></table>";
    var r = try renameTableColumn(a, nested, "B", "Renamed");
    defer r.deinit(a);
    try testing.expectEqual(@as(u32, 1), r.column_pos);
    try testing.expect(std.mem.indexOf(u8, r.bytes, "name=\"Renamed\"") != null);
    try testing.expect(std.mem.indexOf(u8, r.bytes, "name=\"Decoy\"") != null);
    try testing.expect(std.mem.indexOf(u8, r.bytes, "<calculatedColumnFormula>T[Renamed]*2</calculatedColumnFormula>") != null);

    // A column with no `</tableColumn>` inside the block is a hard
    // refusal, not a guessed extent.
    const truncated =
        "<table id=\"1\" name=\"T\" displayName=\"T\" ref=\"A1:B3\">" ++
        "<tableColumns count=\"2\">" ++
        "<tableColumn id=\"1\" name=\"A\"><calculatedColumnFormula>1</calculatedColumnFormula>" ++
        "<tableColumn id=\"2\" name=\"B\"/>" ++
        "</tableColumns></table>";
    try testing.expectError(
        error.MalformedTableXml,
        renameTableColumn(a, truncated, "A", "X"),
    );
}

test "renameTableColumn: OOM-safe at every allocation site" {
    const helpers = struct {
        fn run(allocator: std.mem.Allocator) !void {
            var r = try renameTableColumn(allocator, rename_sample, "R&D", "Budget");
            r.deinit(allocator);
        }
    };
    try testing.checkAllAllocationFailures(testing.allocator, helpers.run, .{});
}
