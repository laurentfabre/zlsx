//! Formula rewriter (C1 milestone 2, iteration 1).
//!
//! Pure-function rewriter built on top of the C1 M1 tokenizer. Walks
//! a tokenized formula and applies a structural edit (insert / delete
//! rows or columns, or sheet rename) to every A1-style cell or range
//! reference it can recognise. Round-trip property of the underlying
//! tokenizer is preserved for everything we don't touch — whitespace,
//! literals, operators, names, and every opaque kind (`.unknown`,
//! `.external_ref`, `.structured_ref`, the dynamic-array operators)
//! pass through verbatim.
//!
//! Scope (this iter):
//!   - Bare A1 cell refs:                   A1, $A$1, $A1, A$1
//!   - Bare A1 ranges:                       A1:B5, $A$1:$B$5
//!   - Sheet-qualified refs:                 Sheet1!A1, 'My Sheet'!A1:B5
//!   - Apostrophe-escaped quoted sheets:    'It''s'!A1
//!   - 3D spans (all edits):                Sheet1:Sheet3!A1, 'Jan:Mar'!A1
//!   - Full-column / full-row spans:        A:A, $A:$XFD, 1:5, $1:$1
//!     (bare, sheet-qualified, and 3D-span spellings). A `name:name`
//!     pair counts as a full-column span whenever BOTH halves spell an
//!     in-grid column: Excel gives the reference interpretation
//!     precedence over defined names in that position (the 2007 grid
//!     expansion force-renamed names the new columns shadowed), so no
//!     name-registry lookup is needed to disambiguate `Start:End`.
//!   - Insert/delete rows/cols, sheet rename
//!   - Structured table refs, for `rename_table_column` ONLY:
//!     `Table1[Old]` → `Table1[New]` in every spelling the specifier
//!     grammar admits (`[@Old]`, `[[#Data],[Old]]`, `[[A]:[B]]`,
//!     escaped names). The specifier is parsed by
//!     `parser.parseStructuredSpecParts` — the SAME grammar the
//!     engine resolves with — and only the matched column-name
//!     subspans are replaced; every other byte of the token passes
//!     through verbatim. For every OTHER edit the token stays opaque.
//!
//! Out of scope (deferred):
//!   - R1C1, dynamic-array `#`/`@`, external-workbook brackets.
//!     Since M1a the tokenizer gives each of these its own kind
//!     rather than lumping them into `.unknown`;
//!     `isOpaqueQualifier` is where the rewriter names the ones it
//!     must not follow.
//!
//! Rename/delete of a 3D span ENDPOINT is in scope when the caller
//! supplies `RewriteContext.sheet_order` (the workbook's tab order,
//! as of BEFORE the edit). Renaming an endpoint rewrites its
//! spelling (order-independent); deleting an endpoint contracts the
//! span inward to its order-neighbor, with `#REF!` only when the
//! deletion covers the whole span. Without `sheet_order`, delete
//! cannot locate the neighbor and leaves the span untouched —
//! conservative, never corrupting.

const std = @import("std");
const tokenizer = @import("tokenizer.zig");
const parser = @import("parser.zig");
const coords = @import("zlsx_refs");
const casefold = @import("zlsx_casefold");
const assert = std.debug.assert;
const Token = tokenizer.Token;

/// Span describing where a row/col edit applies. Shared payload for
/// all four insert/delete variants so a single capture-group switch
/// can reach `at` / `count` without per-variant duplication.
pub const Span = struct { at: u32, count: u32 };

/// Old/new sheet name pair for `rename_sheet`. Comparison is
/// byte-exact (matches Excel's sheet-name lookup once both sides
/// are decoded — quoting/escape happens at the emit boundary, not
/// the comparison).
pub const Rename = struct { old: []const u8, new: []const u8 };

/// Payload for `rename_table_column`. All three are PLAIN text —
/// decoded XML, and `old`/`new` unescaped (`R&D`, not `R&amp;D` and
/// not a `'`-escaped specifier spelling). Matching uses the same
/// rule the engine resolves with (`casefold.eqlFolded` — the symbol
/// layer folds both table and column names before lookup); the
/// escape-out spelling of `new` happens at the emit boundary.
pub const TableColumnRename = struct {
    /// The table's formula-visible name (`displayName`).
    table: []const u8,
    old: []const u8,
    new: []const u8,
};

/// Structural edit applied to every reference the rewriter recognises.
/// Row / column counts use 1-based positions to match Excel's user-
/// facing conventions; the rewriter converts to 0-based internally
/// only at arithmetic sites.
pub const RewriteEdit = union(enum) {
    /// Insert `count` rows starting at row `at` (1-based).
    /// Refs at row >= `at` shift +count; refs at row < `at` unchanged.
    insert_rows: Span,
    /// Delete `count` rows starting at row `at` (1-based).
    /// Refs at row >= `at + count` shift -count; refs in
    /// [at, at + count) collapse to `#REF!`.
    delete_rows: Span,
    insert_cols: Span,
    delete_cols: Span,
    rename_sheet: Rename,
    /// Delete the sheet named `delete_sheet` from the workbook.
    /// Refs qualified to that sheet (`'Sheet1'!A1`,
    /// `Sheet1!A1:B2`) collapse to `#REF!`. Bare refs are
    /// unaffected — they're scoped to the formula's owning sheet,
    /// which is necessarily some OTHER sheet (the deleted
    /// sheet's own formulas are dropped with the sheet, not
    /// rewritten).
    delete_sheet: []const u8,
    /// Rename a column of the named table: rewrite the column-name
    /// half of every `.structured_ref` token that resolves to it.
    /// Qualified refs (`Table1[Old]`) match on the adjacent table
    /// name; bare refs (`[Old]`, `[@Old]`) match only when
    /// `RewriteContext.owning_table` names the same table — the
    /// engine binds the bare form through a table-producer owner
    /// (`vtStructured`), and the rewriter mirrors that rule.
    rename_table_column: TableColumnRename,
};

pub const RewriteContext = struct {
    /// Sheet that owns these formulas. Bare A1 refs (no sheet
    /// qualifier) are implicitly scoped to this sheet; we apply
    /// row/col edits to bare refs only when `on_sheet` is null OR
    /// equals `target_sheet`. `null` means "unknown — apply
    /// everywhere," matching the spec's permissive default.
    on_sheet: ?[]const u8 = null,
    /// The sheet the row/col edit targets. `null` means "apply to
    /// bare refs everywhere AND every sheet-qualified ref."
    /// Name matching follows Excel's rule via
    /// `casefold.excelSheetNameEql` (Unicode case-fold + NFC): a
    /// qualifier spelled `edited!A2` or `'café'!A2` still scopes
    /// to a sheet named `Edited` / `CAFÉ`, because Excel resolves
    /// sheet names case-insensitively.
    target_sheet: ?[]const u8 = null,
    /// Workbook tab order (sheet names, first tab first), as of
    /// BEFORE the edit is applied: for `delete_sheet` it still
    /// contains the doomed sheet; for `rename_sheet` it spells the
    /// OLD name. Enables 3D span endpoint edits (contracting
    /// `Sheet1:Sheet3` when an endpoint is deleted) and mid-span
    /// membership for `target_sheet` scoping. Name resolution
    /// against this list is Excel's case rule
    /// (`casefold.excelSheetNameEql`), like `target_sheet`. `null`
    /// means "order unknown": rename still rewrites endpoints
    /// (order-independent), delete leaves spans it cannot contract
    /// untouched, and mid-span targets conservatively don't match.
    sheet_order: ?[]const []const u8 = null,
    /// The table this formula's bare structured refs (`[Old]`,
    /// `[@Old]`) belong to, if any: the table part's own
    /// `calculatedColumnFormula` / `totalsRowFormula` bodies, or a
    /// cell inside the table's range — Excel binds the bare form
    /// anywhere inside the range (that is how in-table formulas are
    /// written); the engine's producer-owner rule is a strict
    /// subset, so rewriting here never changes what the engine
    /// binds. Decoded plain text (the `displayName`). With `null`
    /// (the default) bare structured refs are never touched.
    owning_table: ?[]const u8 = null,
    edit: RewriteEdit,
};

pub const Error = error{
    OutOfMemory,
    /// `delete_rows` / `delete_cols` with `count == 0` is meaningless
    /// (no shift, nothing to delete). We refuse it rather than silently
    /// no-op, so callers don't construct an edit they didn't mean.
    InvalidEdit,
};

const MAX_ROWS: u32 = coords.max_row;
const MAX_COLS: u32 = coords.max_col_1based;

// ─── public API ──────────────────────────────────────────────────

/// Rewrite `formula_text` per `ctx`. Allocates the result; caller
/// frees with `allocator.free(out)`. On any error, no memory leaks.
pub fn rewriteFormula(
    allocator: std.mem.Allocator,
    formula_text: []const u8,
    ctx: RewriteContext,
) Error![]u8 {
    try validateEdit(ctx.edit);

    const tokens = try tokenizer.tokenize(allocator, formula_text);
    defer allocator.free(tokens);

    // Owned copies of mutated token text. Tokens that survive
    // unchanged keep their borrowed slices; mutated tokens point
    // into this list. Freed only after `format` has copied every
    // byte into the final output.
    var owned: std.ArrayListUnmanaged([]u8) = .empty;
    defer {
        for (owned.items) |s| allocator.free(s);
        owned.deinit(allocator);
    }

    // Working token slice — same length as `tokens`. We mutate in
    // place: replace `.text` for renamed sheets / shifted refs, or
    // collapse a range's tokens to `#REF!` by overwriting kinds.
    const work = try allocator.alloc(Token, tokens.len);
    defer allocator.free(work);
    @memcpy(work, tokens);

    try applyEdit(allocator, work, &owned, ctx);

    return tokenizer.format(allocator, work);
}

// ─── edit validation ─────────────────────────────────────────────

fn validateEdit(edit: RewriteEdit) Error!void {
    switch (edit) {
        .insert_rows, .delete_rows, .insert_cols, .delete_cols => |spec| {
            // `at == 0` is meaningless in Excel's 1-based addressing;
            // refuse rather than wrap. `count == 0` is a no-op edit
            // but still valid — we accept it (caller may have computed
            // an empty range and we shouldn't punish that path).
            if (spec.at == 0) return error.InvalidEdit;
            assert(spec.at >= 1);
        },
        .rename_sheet => |spec| {
            // An empty old name has no targets; an empty new name
            // produces an unaddressable formula. Refuse both.
            if (spec.old.len == 0) return error.InvalidEdit;
            if (spec.new.len == 0) return error.InvalidEdit;
            assert(spec.old.len >= 1);
            assert(spec.new.len >= 1);
        },
        .delete_sheet => |name| {
            // Empty name has no targets — refuse rather than
            // silently no-op.
            if (name.len == 0) return error.InvalidEdit;
            assert(name.len >= 1);
        },
        .rename_table_column => |spec| {
            // Same policy as rename_sheet: an empty table or old
            // name has no targets, an empty new name produces an
            // unaddressable specifier (`Table1[]` is malformed).
            if (spec.table.len == 0) return error.InvalidEdit;
            if (spec.old.len == 0) return error.InvalidEdit;
            if (spec.new.len == 0) return error.InvalidEdit;
        },
    }
}

// ─── A1 ref structural type ──────────────────────────────────────

const Ref = struct {
    col_abs: bool,
    col: u32, // 1-based, [1, 16384]
    row_abs: bool,
    row: u32, // 1-based, [1, 1_048_576]
};

/// Parse an A1-style cell ref slice into structured form. Returns
/// null for slices that the M1 tokenizer would never label
/// `.cell_ref` (defensive — tokenizer guarantees the shape, but the
/// rewriter keeps the parser explicit so the printer's preconditions
/// stay enforced near the use site).
fn parseRef(s: []const u8) ?Ref {
    if (s.len == 0) return null;
    var i: usize = 0;
    var col_abs = false;
    if (s[i] == '$') {
        col_abs = true;
        i += 1;
    }
    const col_start = i;
    while (i < s.len and isAsciiAlpha(s[i])) : (i += 1) {}
    if (i == col_start) return null;
    const col_letters = s[col_start..i];
    var row_abs = false;
    if (i < s.len and s[i] == '$') {
        row_abs = true;
        i += 1;
    }
    const row_start = i;
    while (i < s.len and isDigit(s[i])) : (i += 1) {}
    if (i == row_start) return null;
    if (i != s.len) return null;
    const col_value = colLettersToNum(col_letters) orelse return null;
    const row_value = std.fmt.parseInt(u32, s[row_start..i], 10) catch return null;
    if (row_value == 0 or row_value > MAX_ROWS) return null;
    if (col_value == 0 or col_value > MAX_COLS) return null;
    return Ref{
        .col_abs = col_abs,
        .col = col_value,
        .row_abs = row_abs,
        .row = row_value,
    };
}

fn formatRef(allocator: std.mem.Allocator, ref: Ref) Error![]u8 {
    assert(ref.col >= 1 and ref.col <= MAX_COLS);
    assert(ref.row >= 1 and ref.row <= MAX_ROWS);

    var buf: [16]u8 = undefined; // worst case: $XFD$1048576 = 12 bytes
    var len: usize = 0;
    if (ref.col_abs) {
        buf[len] = '$';
        len += 1;
    }
    len += writeColLetters(buf[len..], ref.col);
    if (ref.row_abs) {
        buf[len] = '$';
        len += 1;
    }
    len += std.fmt.printInt(buf[len..], ref.row, 10, .lower, .{});

    return allocator.dupe(u8, buf[0..len]);
}

/// Convert column letters ("A", "AA", "XFD", or lowercase variants)
/// to a 1-based column number. Returns null if any byte is not an
/// ASCII letter (caller has already vetted via tokenizer, but be
/// explicit).
fn colLettersToNum(letters: []const u8) ?u32 {
    // M0 adapter over `zlsx_refs`. Policy preserved exactly: at most 3
    // letters, case-insensitive, and NO grid ceiling — an out-of-grid
    // column like "ZZZ" is accepted here and rejected downstream.
    return coords.parseColNumber(letters, .{
        .case = .insensitive,
        .max_letters = 3,
        .bounds = .unchecked,
    }) catch null;
}

/// Write column letters for `col` (1-based) into `buf`. Returns the
/// number of bytes written. `buf.len >= 3` is sufficient — XFD is
/// the largest valid column.
fn writeColLetters(buf: []u8, col: u32) usize {
    assert(col >= 1 and col <= MAX_COLS);
    assert(buf.len >= 3);
    // In grid by the assert above, so the typed constructor cannot fail.
    return coords.writeColLetters(buf, coords.Col.fromOneBased(col) catch unreachable);
}

inline fn isAsciiAlpha(c: u8) bool {
    return (c >= 'A' and c <= 'Z') or (c >= 'a' and c <= 'z');
}

inline fn isDigit(c: u8) bool {
    return c >= '0' and c <= '9';
}

// ─── full-column / full-row spans ────────────────────────────────

const AxisKind = enum { cols, rows };

/// One bound of a full-column / full-row span: `$A` → { abs, 1 },
/// `5` → { !abs, 5 }. `n` is 1-based and in-grid for its axis.
const AxisBound = struct {
    abs: bool,
    n: u32,
};

/// A whole-axis span occupying three tokens: bound, op_range, bound.
/// `first`/`second` keep the WRITTEN order (`C:A` stays reversed);
/// arithmetic normalises internally and writes back positionally.
const AxisRange = struct {
    kind: AxisKind,
    first: AxisBound,
    second: AxisBound,
};

/// Parse one full-column bound: optional `$`, then 1..3 column
/// letters mapping into the grid. Null for anything else.
fn parseColSpec(s: []const u8) ?AxisBound {
    if (s.len == 0) return null;
    var i: usize = 0;
    var abs = false;
    if (s[0] == '$') {
        abs = true;
        i = 1;
    }
    if (i == s.len) return null;
    for (s[i..]) |c| {
        if (!isAsciiAlpha(c)) return null;
    }
    const col = colLettersToNum(s[i..]) orelse return null;
    if (col == 0 or col > MAX_COLS) return null;
    return .{ .abs = abs, .n = col };
}

/// Parse one full-row bound: optional `$`, then digits only (so
/// `1.5` and `1e5` — legal number lexemes — are rejected), value
/// inside the grid. Null for anything else.
fn parseRowSpec(s: []const u8) ?AxisBound {
    if (s.len == 0) return null;
    var i: usize = 0;
    var abs = false;
    if (s[0] == '$') {
        abs = true;
        i = 1;
    }
    if (i == s.len) return null;
    for (s[i..]) |c| {
        if (!isDigit(c)) return null;
    }
    const row = std.fmt.parseInt(u32, s[i..], 10) catch return null;
    if (row == 0 or row > MAX_ROWS) return null;
    return .{ .abs = abs, .n = row };
}

fn formatAxisBound(allocator: std.mem.Allocator, kind: AxisKind, bound: AxisBound) Error![]u8 {
    const cap: u32 = switch (kind) {
        .cols => MAX_COLS,
        .rows => MAX_ROWS,
    };
    assert(bound.n >= 1 and bound.n <= cap);

    var buf: [9]u8 = undefined; // worst case: $1048576 = 8 bytes
    var len: usize = 0;
    if (bound.abs) {
        buf[len] = '$';
        len += 1;
    }
    switch (kind) {
        .cols => len += writeColLetters(buf[len..], bound.n),
        .rows => len += std.fmt.printInt(buf[len..], bound.n, 10, .lower, .{}),
    }
    return allocator.dupe(u8, buf[0..len]);
}

// ─── shift application ───────────────────────────────────────────

const ShiftOutcome = union(enum) {
    unchanged,
    shifted: Ref,
    deleted,
};

fn applyShift(ref: Ref, edit: RewriteEdit) ShiftOutcome {
    switch (edit) {
        .insert_rows => |spec| {
            assert(spec.at >= 1);
            if (spec.count == 0) return .unchanged;
            if (ref.row < spec.at) return .unchanged;
            const new_row = std.math.add(u32, ref.row, spec.count) catch return .unchanged;
            if (new_row > MAX_ROWS) return .unchanged; // off-grid: leave alone
            return .{ .shifted = .{
                .col_abs = ref.col_abs,
                .col = ref.col,
                .row_abs = ref.row_abs,
                .row = new_row,
            } };
        },
        .delete_rows => |spec| {
            assert(spec.at >= 1);
            if (spec.count == 0) return .unchanged;
            if (ref.row < spec.at) return .unchanged;
            const end = std.math.add(u32, spec.at, spec.count) catch return .deleted;
            if (ref.row < end) return .deleted;
            const new_row = ref.row - spec.count;
            assert(new_row >= 1);
            return .{ .shifted = .{
                .col_abs = ref.col_abs,
                .col = ref.col,
                .row_abs = ref.row_abs,
                .row = new_row,
            } };
        },
        .insert_cols => |spec| {
            assert(spec.at >= 1);
            if (spec.count == 0) return .unchanged;
            if (ref.col < spec.at) return .unchanged;
            const new_col = std.math.add(u32, ref.col, spec.count) catch return .unchanged;
            if (new_col > MAX_COLS) return .unchanged;
            return .{ .shifted = .{
                .col_abs = ref.col_abs,
                .col = new_col,
                .row_abs = ref.row_abs,
                .row = ref.row,
            } };
        },
        .delete_cols => |spec| {
            assert(spec.at >= 1);
            if (spec.count == 0) return .unchanged;
            if (ref.col < spec.at) return .unchanged;
            const end = std.math.add(u32, spec.at, spec.count) catch return .deleted;
            if (ref.col < end) return .deleted;
            const new_col = ref.col - spec.count;
            assert(new_col >= 1);
            return .{ .shifted = .{
                .col_abs = ref.col_abs,
                .col = new_col,
                .row_abs = ref.row_abs,
                .row = ref.row,
            } };
        },
        .rename_sheet => return .unchanged,
        // delete_sheet handling lives at the sheet-qualifier level
        // (rewriteSheetQualifiedRefOrRange) — the per-ref shift
        // never sees the sheet name and so can't decide. Treat as
        // unchanged here; the caller has already collapsed the
        // qualified ref to #REF! before applyShift runs.
        .delete_sheet => return .unchanged,
        // Structured refs never contain live A1 refs (the token is
        // opaque to this path); nothing to shift.
        .rename_table_column => return .unchanged,
    }
}

const AxisOutcome = union(enum) {
    unchanged,
    shifted: AxisRange,
    deleted,
};

/// Shift a full-column / full-row span. Edits on the perpendicular
/// axis never touch it (deleting rows cannot reshape `A:A`).
///
/// Deletion uses INTERVAL semantics — the span shrinks around the
/// deleted block and collapses to #REF! only when every spanned
/// column/row is gone — unlike the cell-range path's per-endpoint
/// policy (`A4:A5` → `A4:#REF!`): `A:#REF!` is not a printable
/// reference, so a partially deleted span must stay a span.
fn applyAxisShift(axis: AxisRange, edit: RewriteEdit) AxisOutcome {
    const relevant = switch (edit) {
        .insert_rows, .delete_rows => axis.kind == .rows,
        .insert_cols, .delete_cols => axis.kind == .cols,
        .rename_sheet, .delete_sheet, .rename_table_column => false,
    };
    if (!relevant) return .unchanged;

    const cap: u32 = switch (axis.kind) {
        .cols => MAX_COLS,
        .rows => MAX_ROWS,
    };
    assert(axis.first.n >= 1 and axis.first.n <= cap);
    assert(axis.second.n >= 1 and axis.second.n <= cap);

    switch (edit) {
        .insert_rows, .insert_cols => |spec| {
            assert(spec.at >= 1);
            if (spec.count == 0) return .unchanged;
            // Per-bound, mirroring the cell path: a bound the insert
            // would push off-grid keeps its position (`A:XFD` stays
            // `A:XFD`; `C:XFD` still becomes `D:XFD`).
            var out = axis;
            var changed = false;
            for ([_]*AxisBound{ &out.first, &out.second }) |bound| {
                if (bound.n < spec.at) continue;
                const shifted = std.math.add(u32, bound.n, spec.count) catch continue;
                if (shifted > cap) continue;
                bound.n = shifted;
                changed = true;
            }
            if (!changed) return .unchanged;
            return .{ .shifted = out };
        },
        .delete_rows, .delete_cols => |spec| {
            assert(spec.at >= 1);
            if (spec.count == 0) return .unchanged;
            const lo = @min(axis.first.n, axis.second.n);
            const hi = @max(axis.first.n, axis.second.n);
            // Deleted zone is [at, zone_end); saturate so a huge
            // count reads as "to the end of the grid".
            const zone_end = std.math.add(u32, spec.at, spec.count) catch std.math.maxInt(u32);
            if (spec.at > hi) return .unchanged;
            if (zone_end <= lo) {
                // Zone entirely before the span: both bounds slide.
                var out = axis;
                out.first.n -= spec.count;
                out.second.n -= spec.count;
                assert(out.first.n >= 1 and out.second.n >= 1);
                return .{ .shifted = out };
            }
            if (spec.at <= lo and zone_end > hi) return .deleted;
            // Partial overlap: the surviving interval, in post-delete
            // positions. A bound inside the zone snaps to the zone's
            // edge; a bound past it slides down.
            const new_lo = if (lo < spec.at) lo else spec.at;
            const new_hi = if (hi >= zone_end) hi - spec.count else spec.at - 1;
            assert(new_lo >= 1 and new_lo <= new_hi);
            var out = axis;
            if (axis.first.n <= axis.second.n) {
                out.first.n = new_lo;
                out.second.n = new_hi;
            } else {
                out.first.n = new_hi;
                out.second.n = new_lo;
            }
            return .{ .shifted = out };
        },
        // Filtered out by `relevant` above.
        .rename_sheet, .delete_sheet, .rename_table_column => unreachable,
    }
}

// ─── sheet-name handling ─────────────────────────────────────────

/// Decode a `.sheet_name` token's text (the borrowed slice including
/// the surrounding `'` quotes). Returns the unescaped sheet name
/// owned by `allocator`. Caller frees.
fn decodeQuotedSheet(allocator: std.mem.Allocator, lex: []const u8) Error![]u8 {
    assert(lex.len >= 2);
    assert(lex[0] == '\'');
    // Tokenizer's scanQuotedSheet may return without a closing quote
    // when it hits EOF; defensive — treat the trailing byte as
    // optional. Production-shape tokens always have the closer.
    const has_close = lex[lex.len - 1] == '\'';
    const inner = if (has_close) lex[1 .. lex.len - 1] else lex[1..];

    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(allocator);

    var i: usize = 0;
    while (i < inner.len) : (i += 1) {
        if (inner[i] == '\'' and i + 1 < inner.len and inner[i + 1] == '\'') {
            try out.append(allocator, '\'');
            i += 1;
            continue;
        }
        try out.append(allocator, inner[i]);
    }
    return out.toOwnedSlice(allocator);
}

/// Re-emit a sheet name as a token-text slice, choosing quoted or
/// unquoted form to match Excel's printing rules. Also escapes
/// embedded apostrophes by doubling them. Caller frees.
fn encodeSheetName(allocator: std.mem.Allocator, name: []const u8) Error![]u8 {
    assert(name.len >= 1);

    if (canEmitUnquoted(name)) {
        return allocator.dupe(u8, name);
    }

    // Quoted form: leading `'`, doubled apostrophes inside, trailing `'`.
    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(allocator);

    try out.append(allocator, '\'');
    for (name) |c| {
        if (c == '\'') try out.append(allocator, '\'');
        try out.append(allocator, c);
    }
    try out.append(allocator, '\'');
    return out.toOwnedSlice(allocator);
}

/// True if `name` can be emitted without surrounding quotes. We
/// require a leading letter or underscore and only letter/digit/
/// underscore/period bytes — the same conservative rule Excel uses.
/// Anything else (spaces, punctuation, leading digit) needs quotes.
fn canEmitUnquoted(name: []const u8) bool {
    if (name.len == 0) return false;
    const first = name[0];
    if (!(isAsciiAlpha(first) or first == '_')) return false;
    for (name[1..]) |c| {
        const ok = isAsciiAlpha(c) or isDigit(c) or c == '_' or c == '.';
        if (!ok) return false;
    }
    return true;
}

// ─── structured-ref column rename ────────────────────────────────

/// Rewrite one `.structured_ref` token for `rename_table_column`.
/// Every other edit kind leaves the token opaque, as do specifiers
/// scoped to another table, specifiers without a matching column,
/// and malformed specifiers (which the engine refuses to bind — a
/// rewrite there could only corrupt).
fn rewriteStructuredRef(
    allocator: std.mem.Allocator,
    work: []Token,
    owned: *std.ArrayListUnmanaged([]u8),
    ctx: RewriteContext,
    i: usize,
) Error!void {
    assert(work[i].kind == .structured_ref);
    const spec = switch (ctx.edit) {
        .rename_table_column => |s| s,
        else => return,
    };

    // Scope: whose specifier is this? `Name[…]` — adjacent, the
    // parser's own `kindAt(1)` pairing rule, so `Table1 [x]` with
    // whitespace between is NOT a table specifier — belongs to that
    // name. Anything reached through a `!` is qualified by a sheet
    // or an external workbook and is never a local table specifier
    // (Excel admits neither `Sheet1!Table1[c]` nor `Sheet1![c]`).
    // A bare `[…]` names the owner's own table, which only a
    // table-producer context has (`RewriteContext.owning_table`).
    if (i >= 1 and work[i - 1].kind == .bang) return;
    if (i >= 1 and work[i - 1].kind == .name) {
        if (i >= 2 and work[i - 2].kind == .bang) return;
        if (!try plainNamesMatch(allocator, work[i - 1].text, spec.table)) return;
    } else {
        const owner = ctx.owning_table orelse return;
        if (!try plainNamesMatch(allocator, owner, spec.table)) return;
    }

    if (try rewriteSpecifierColumns(allocator, work[i].text, spec)) |new_text| {
        work[i].text = try registerOwned(allocator, owned, new_text);
    }
}

/// Fold-equality on decoded plain names — the engine's matching rule
/// for both table and column names (the symbol layer folds each side
/// before every lookup, `pkg/workbook.zig::columnIndex`). Invalid
/// UTF-8 cannot match anything the symbol layer admitted, so it
/// reads as "no match" rather than an error.
fn plainNamesMatch(allocator: std.mem.Allocator, a: []const u8, b: []const u8) Error!bool {
    return casefold.eqlFolded(allocator, a, b) catch |err| {
        if (err == error.OutOfMemory) return error.OutOfMemory;
        return false;
    };
}

/// Replace every column-name part of `raw` (a full `[…]` specifier,
/// brackets included) that names `spec.old` with the spelling of
/// `spec.new`. Returns the new token text, or null when nothing
/// matched. The specifier is parsed by the parser's own grammar and
/// only the matched parts' byte spans are replaced — item
/// specifiers, separators, whitespace and every unmatched byte pass
/// through verbatim. A bracketed part keeps its brackets and swaps
/// inner text only; a bare part gains brackets when the new name
/// needs them (`columnNeedsBrackets` — the printer's policy, so the
/// rewriter and the canonical printer agree on which names go bare).
fn rewriteSpecifierColumns(
    allocator: std.mem.Allocator,
    raw: []const u8,
    spec: TableColumnRename,
) Error!?[]u8 {
    const parsed = parser.parseStructuredSpecParts(raw) orelse return null;

    var parts_buf: [2]parser.ColumnPart = undefined;
    const parts: []const parser.ColumnPart = switch (parsed.columns) {
        .none => &.{},
        .one => |p| blk: {
            parts_buf[0] = p;
            break :blk parts_buf[0..1];
        },
        // Source order: `first` precedes `last` in the raw bytes, so
        // the cursor below advances monotonically.
        .range => |r| blk: {
            parts_buf[0] = r.first;
            parts_buf[1] = r.last;
            break :blk parts_buf[0..2];
        },
    };

    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(allocator);
    var cursor: usize = 0;
    var matched = false;
    for (parts) |p| {
        // The grammar hands back slices INTO `raw`; recover each
        // part's offset for the splice.
        const off = @intFromPtr(p.text.ptr) - @intFromPtr(raw.ptr);
        assert(off + p.text.len <= raw.len);
        const decoded = try parser.decodeColumnName(allocator, p.text);
        defer allocator.free(decoded);
        if (!try plainNamesMatch(allocator, decoded, spec.old)) continue;
        matched = true;
        const encoded = try parser.encodeColumnName(allocator, spec.new);
        defer allocator.free(encoded);
        try out.appendSlice(allocator, raw[cursor..off]);
        const wrap = !p.bracketed and parser.columnNeedsBrackets(encoded);
        if (wrap) try out.append(allocator, '[');
        try out.appendSlice(allocator, encoded);
        if (wrap) try out.append(allocator, ']');
        cursor = off + p.text.len;
    }
    if (!matched) {
        out.deinit(allocator);
        return null;
    }
    try out.appendSlice(allocator, raw[cursor..]);
    return try out.toOwnedSlice(allocator);
}

// ─── token walk + rewrite ───────────────────────────────────────

fn applyEdit(
    allocator: std.mem.Allocator,
    work: []Token,
    owned: *std.ArrayListUnmanaged([]u8),
    ctx: RewriteContext,
) Error!void {
    assert(work.len >= 0);

    var i: usize = 0;
    while (i < work.len) {
        // An external-workbook prefix disqualifies everything it
        // qualifies, not just itself: in `[1]Sheet1!A1` the sheet and
        // the ref are inside a workbook we cannot see, yet they
        // tokenize as an ordinary `name bang cell_ref` triple that the
        // qualifier matcher below would happily rewrite. Skip the
        // whole chain.
        if (work[i].kind == .external_ref) {
            i = endOfExternalReference(work, i);
            continue;
        }

        // Structured table specifier. Opaque to every edit except
        // `rename_table_column`, which rewrites the column-name
        // subspans in place; the token count never changes, so the
        // matchers below are unaffected.
        if (work[i].kind == .structured_ref) {
            try rewriteStructuredRef(allocator, work, owned, ctx, i);
            i += 1;
            continue;
        }

        // Detect 3D span: (sheet_name | name) : (sheet_name | name)
        // bang cell_ref [: cell_ref]. Must run before the single-
        // sheet matcher, which would otherwise claim the second
        // endpoint (`Sheet3!A1`) and shift — or rename/collapse —
        // under the wrong scoping.
        if (matchThreeDQualifier(work, i)) |info| {
            try rewriteThreeD(allocator, work, owned, ctx, info);
            i = info.end;
            continue;
        }

        // Detect sheet qualifier: (sheet_name | name) bang cell_ref [: cell_ref]
        const sq = matchSheetQualifier(work, i);
        if (sq) |info| {
            try rewriteSheetQualified(allocator, work, owned, ctx, info);
            i = info.end;
            continue;
        }

        // Bare cell ref or range starting at this token. Skip
        // refs that follow an opaque `!` qualifier — those are
        // scoped to a sheet the rewriter cannot reason about
        // (external workbook, unclassifiable bytes). Mutating them
        // would corrupt formulas pointing at workbooks we don't see.
        if (work[i].kind == .cell_ref) {
            if (i >= 2 and work[i - 1].kind == .bang and isOpaqueQualifier(work[i - 2].kind)) {
                i += 1;
                continue;
            }
            const range_end = if (isRangeAt(work, i)) i + 3 else i + 1;
            try rewriteBareRefOrRange(allocator, work, owned, ctx, i, range_end);
            i = range_end;
            continue;
        }

        // Bare full-column / full-row span (`A:A`, `1:5`) starting at
        // this token. Same opaque-qualifier guard as cell refs: a
        // span after `[..]!` or an unclassifiable `!` belongs to a
        // sheet the rewriter cannot reason about.
        if (work[i].kind == .name or work[i].kind == .number) {
            if (i >= 2 and work[i - 1].kind == .bang and isOpaqueQualifier(work[i - 2].kind)) {
                i += 1;
                continue;
            }
            if (matchAxisRange(work, i)) |axis| {
                if (bareEditApplies(ctx)) {
                    try applyAxisRange(allocator, work, owned, ctx.edit, i, axis);
                }
                i += 3;
                continue;
            }
        }

        i += 1;
    }
}

/// Token kinds that stand in for a sheet the rewriter must not reason
/// about. M1a split the old single `.unknown` bucket into typed kinds
/// — an external-workbook prefix is now `.external_ref` — so the guard
/// has to name each one; keying on `.unknown` alone silently started
/// rewriting external references.
fn isOpaqueQualifier(kind: Token.Kind) bool {
    return switch (kind) {
        .unknown, .external_ref, .structured_ref => true,
        else => false,
    };
}

/// Index just past everything an `.external_ref` at `start` qualifies:
/// `[1]Sheet1!A1:B2`, `[1]Sheet1:Sheet3!A1`, `[1]!Name`,
/// `'[B.xlsx]S'!A1`. Each element is optional, so a bare `[1]`
/// consumes only itself.
fn endOfExternalReference(work: []const Token, start: usize) usize {
    assert(work[start].kind == .external_ref);
    var i = start + 1;
    if (i < work.len and (work[i].kind == .name or work[i].kind == .sheet_name)) {
        i += 1;
        // External 3D span: consume the second endpoint only when a
        // bang follows, so a genuine range operator after `[1]Name`
        // is left for the ordinary walk.
        if (i + 2 < work.len and work[i].kind == .op_range and
            (work[i + 1].kind == .name or work[i + 1].kind == .sheet_name) and
            work[i + 2].kind == .bang)
        {
            i += 2;
        }
    }
    if (i < work.len and work[i].kind == .bang) i += 1;
    if (i < work.len and work[i].kind == .cell_ref) {
        i += 1;
        if (i + 1 < work.len and work[i].kind == .op_range and work[i + 1].kind == .cell_ref) {
            i += 2;
        }
    } else if (matchAxisRange(work, i)) |_| {
        // `[1]Sheet1!A:A` / `[1]Sheet1!1:5` — a whole-axis span
        // inside a workbook we cannot see. Without this arm the walk
        // resumes on the span and the axis matcher rewrites it.
        i += 3;
    } else if (i < work.len and work[i].kind == .name) {
        // `[1]!DefinedName`
        i += 1;
    }
    return i;
}

const SheetQualifierInfo = struct {
    sheet_idx: usize, // index of the sheet token (.sheet_name or .name)
    bang_idx: usize, // index of the `!`
    ref_start: usize, // index of the first cell_ref (or axis bound)
    ref_end: usize, // exclusive — i.e. ref_start+1 (single) or ref_start+3 (range/axis)
    end: usize, // index just past the whole pattern (== ref_end)
    is_range: bool,
    /// Non-null when the ref part is a full-column / full-row span
    /// (`Sheet1!A:A`, `'My Sheet'!1:5`) rather than cell refs.
    axis: ?AxisRange = null,
};

fn matchSheetQualifier(work: []const Token, i: usize) ?SheetQualifierInfo {
    if (i + 2 >= work.len) return null; // need at least sheet, !, ref
    const sheet_kind = work[i].kind;
    if (sheet_kind != .sheet_name and sheet_kind != .name) return null;
    if (work[i + 1].kind != .bang) return null;
    if (work[i + 2].kind != .cell_ref) {
        // Whole-axis tail: `Sheet1!A:A`, `Sheet1!$1:$5`.
        const axis = matchAxisRange(work, i + 2) orelse return null;
        return .{
            .sheet_idx = i,
            .bang_idx = i + 1,
            .ref_start = i + 2,
            .ref_end = i + 5,
            .end = i + 5,
            .is_range = true,
            .axis = axis,
        };
    }
    var end: usize = i + 3;
    var is_range = false;
    if (i + 4 < work.len and work[i + 3].kind == .op_range and work[i + 4].kind == .cell_ref) {
        end = i + 5;
        is_range = true;
    }
    return .{
        .sheet_idx = i,
        .bang_idx = i + 1,
        .ref_start = i + 2,
        .ref_end = if (is_range) i + 5 else i + 3,
        .end = end,
        .is_range = is_range,
    };
}

fn isRangeAt(work: []const Token, i: usize) bool {
    if (i + 2 >= work.len) return false;
    return work[i].kind == .cell_ref and
        work[i + 1].kind == .op_range and
        work[i + 2].kind == .cell_ref;
}

/// Match a full-column (`A:A`, `$A:$XFD`) or full-row (`1:5`,
/// `$1:$1`) span at `i`. Column bounds arrive as `.name` tokens; row
/// bounds as `.number` (bare) or `.name` (`$`-prefixed — the
/// tokenizer folds the marker into an identifier lexeme). Both
/// halves must parse on the SAME axis; a half that spells anything
/// else (a defined name, an off-grid column, a non-integer number)
/// rejects the whole pattern and the tokens pass through untouched.
fn matchAxisRange(work: []const Token, i: usize) ?AxisRange {
    if (i + 2 >= work.len) return null;
    if (work[i + 1].kind != .op_range) return null;
    // A trailing bang means the pair is a 3D span qualifier
    // (`AB:CD!A1`), never a whole-axis reference.
    if (i + 3 < work.len and work[i + 3].kind == .bang) return null;
    const a = work[i];
    const b = work[i + 2];
    if (a.kind == .name and b.kind == .name) {
        if (parseColSpec(a.text)) |first| {
            if (parseColSpec(b.text)) |second| {
                return .{ .kind = .cols, .first = first, .second = second };
            }
        }
    }
    const a_rowish = a.kind == .number or a.kind == .name;
    const b_rowish = b.kind == .number or b.kind == .name;
    if (a_rowish and b_rowish) {
        if (parseRowSpec(a.text)) |first| {
            if (parseRowSpec(b.text)) |second| {
                return .{ .kind = .rows, .first = first, .second = second };
            }
        }
    }
    return null;
}

const ThreeDQualifierInfo = struct {
    first_idx: usize, // index of the first endpoint (.sheet_name or .name)
    second_idx: usize, // index of the second endpoint
    ref_start: usize, // index of the first cell_ref (or axis bound)
    is_range: bool,
    end: usize, // index just past the whole pattern
    /// Non-null when the ref part is a full-column / full-row span
    /// (`Sheet1:Sheet3!A:A`) rather than cell refs.
    axis: ?AxisRange = null,
};

/// Match a 3D span qualifier: (sheet_name | name) op_range
/// (sheet_name | name) bang cell_ref [op_range cell_ref]. The bang
/// disambiguates from a defined-name range (`Start:End` inside SUM)
/// and from full-column refs (`A:A`), neither of which is followed
/// by `!`.
fn matchThreeDQualifier(work: []const Token, i: usize) ?ThreeDQualifierInfo {
    if (i + 4 >= work.len) return null; // need sheet, :, sheet, !, ref
    if (work[i].kind != .sheet_name and work[i].kind != .name) return null;
    if (work[i + 1].kind != .op_range) return null;
    if (work[i + 2].kind != .sheet_name and work[i + 2].kind != .name) return null;
    if (work[i + 3].kind != .bang) return null;
    if (work[i + 4].kind != .cell_ref) {
        // Whole-axis tail: `Sheet1:Sheet3!A:A`, `Sheet1:Sheet3!1:5`.
        const axis = matchAxisRange(work, i + 4) orelse return null;
        return .{
            .first_idx = i,
            .second_idx = i + 2,
            .ref_start = i + 4,
            .is_range = true,
            .end = i + 7,
            .axis = axis,
        };
    }
    var end: usize = i + 5;
    var is_range = false;
    if (i + 6 < work.len and work[i + 5].kind == .op_range and work[i + 6].kind == .cell_ref) {
        end = i + 7;
        is_range = true;
    }
    return .{
        .first_idx = i,
        .second_idx = i + 2,
        .ref_start = i + 4,
        .is_range = is_range,
        .end = end,
    };
}

/// Decode a sheet-qualifier token to its plain-text name: unquote
/// and unescape `.sheet_name`, dupe `.name` verbatim. Caller frees.
fn decodeSheetToken(allocator: std.mem.Allocator, tok: Token) Error![]u8 {
    if (tok.kind == .sheet_name) return decodeQuotedSheet(allocator, tok.text);
    return allocator.dupe(u8, tok.text);
}

/// Decoded 3D span endpoints, in WRITTEN order (a reversed spelling
/// like `Sheet3:Sheet1` keeps its orientation; arithmetic normalises
/// internally and writes back positionally, mirroring `AxisRange`).
const SpanHalves = struct { first: []const u8, second: []const u8 };

/// Split a decoded sheet qualifier into 3D-span halves. A colon is
/// illegal in Excel sheet names, so a decoded qualifier containing
/// exactly one colon with text on both sides can only be the quoted
/// 3D span spelling (`'Jan:Mar'!A1`). Null for anything else.
fn splitQuotedSpan(decoded: []const u8) ?SpanHalves {
    const colon = std.mem.indexOfScalar(u8, decoded, ':') orelse return null;
    const first = decoded[0..colon];
    const second = decoded[colon + 1 ..];
    if (first.len == 0 or second.len == 0) return null;
    if (std.mem.indexOfScalar(u8, second, ':') != null) return null;
    return .{ .first = first, .second = second };
}

/// Position of `name` in the workbook tab order. A byte-exact entry
/// wins outright; otherwise the name resolves with Excel's case rule
/// — but only when that resolution is UNIQUE. Excel forbids
/// case-variant duplicate sheet names, so a well-formed order never
/// has two case-fold matches; a malformed or caller-supplied order
/// that does is ambiguous, and span arithmetic through an arbitrary
/// pick could contract through the wrong interval. Null means "do
/// not run span arithmetic" — the conservative signal.
fn orderIndexOf(order: []const []const u8, name: []const u8) ?usize {
    var folded: ?usize = null;
    var folded_count: usize = 0;
    for (order, 0..) |entry, idx| {
        if (std.mem.eql(u8, entry, name)) return idx;
        if (casefold.excelSheetNameEql(entry, name)) {
            folded = idx;
            folded_count += 1;
        }
    }
    return if (folded_count == 1) folded else null;
}

/// True when `target` is a member of the span `first:second` —
/// an endpoint, or (when `order` is supplied) a mid-span sheet.
/// Without `order`, mid-span membership is undecidable and the
/// target conservatively does not match.
fn spanContainsTarget(
    first: []const u8,
    second: []const u8,
    target: []const u8,
    order: ?[]const []const u8,
) bool {
    if (casefold.excelSheetNameEql(first, target)) return true;
    if (casefold.excelSheetNameEql(second, target)) return true;
    const ord = order orelse return false;
    const ti = orderIndexOf(ord, target) orelse return false;
    const fi = orderIndexOf(ord, first) orelse return false;
    const si = orderIndexOf(ord, second) orelse return false;
    return @min(fi, si) <= ti and ti <= @max(fi, si);
}

/// Row/col target matching for a decoded sheet qualifier. For the
/// quoted 3D span spelling (`'Jan:Mar'!A1`) the target matches
/// either endpoint, or any mid-span member when the workbook order
/// is available. Name comparison follows Excel's case rule — see
/// `RewriteContext.target_sheet`.
fn sheetTargetMatches(decoded: []const u8, target: []const u8, order: ?[]const []const u8) bool {
    const span = splitQuotedSpan(decoded) orelse
        return casefold.excelSheetNameEql(decoded, target);
    return spanContainsTarget(span.first, span.second, target, order);
}

/// What `delete_sheet` does to a 3D span, decided from the decoded
/// endpoint names and the workbook tab order. Endpoint matching is
/// byte-exact, mirroring the pinned single-qualifier rename/delete
/// semantics; order positions resolve with Excel's case rule.
const SpanDeleteOutcome = union(enum) {
    /// Doomed sheet is mid-span, outside the span, or an endpoint we
    /// cannot contract (no order / unresolvable name). A mid-span
    /// member is never spelled in the text, so "unchanged" is exact
    /// for the first two and conservative for the rest.
    unchanged,
    /// The deletion covers the whole span — the qualified ref
    /// collapses to `#REF!`.
    collapse,
    /// The deleted endpoint steps inward to its order-neighbor;
    /// slices borrow from `order` or the caller's decoded names.
    contract: SpanHalves,
};

fn deleteSpanOutcome(
    first: []const u8,
    second: []const u8,
    doomed: []const u8,
    order: ?[]const []const u8,
) SpanDeleteOutcome {
    const first_matches = std.mem.eql(u8, first, doomed);
    const second_matches = std.mem.eql(u8, second, doomed);
    // Both endpoints spell the doomed sheet: the span covers exactly
    // one sheet position, so the deletion covers the whole span —
    // decidable without any order.
    if (first_matches and second_matches) return .collapse;
    if (!first_matches and !second_matches) return .unchanged;

    const ord = order orelse return .unchanged;
    const fi = orderIndexOf(ord, first) orelse return .unchanged;
    const si = orderIndexOf(ord, second) orelse return .unchanged;
    const lo = @min(fi, si);
    const hi = @max(fi, si);
    // Distinct spellings on one order position can only be a
    // case-variant pair (`SHEET1:Sheet1`) — still a single-position
    // span, fully covered by the deletion.
    if (lo == hi) return .collapse;

    // The matched endpoint steps inward: the interval minimum moves
    // up, the maximum moves down. The replacement lands on whichever
    // WRITTEN endpoint spelled the doomed name, so a reversed
    // spelling keeps its orientation.
    var new_first = first;
    var new_second = second;
    if (first_matches) new_first = ord[if (fi == lo) lo + 1 else hi - 1];
    if (second_matches) new_second = ord[if (si == lo) lo + 1 else hi - 1];
    return .{ .contract = .{ .first = new_first, .second = new_second } };
}

/// Rename applied to a span's decoded endpoints, byte-exact per the
/// pinned single-qualifier rename semantics. Needs no sheet order:
/// only endpoints are ever spelled, so a mid-span rename leaves the
/// text unchanged (null).
fn renameSpanEndpoints(first: []const u8, second: []const u8, spec: Rename) ?SpanHalves {
    const first_matches = std.mem.eql(u8, first, spec.old);
    const second_matches = std.mem.eql(u8, second, spec.old);
    if (!first_matches and !second_matches) return null;
    return .{
        .first = if (first_matches) spec.new else first,
        .second = if (second_matches) spec.new else second,
    };
}

/// Encode span halves as ONE token-text slice, Excel-canonically:
/// unquoted `First:Second` when both halves can stand bare, else the
/// pair quoted as a unit (`'Jan Sales:Apr'`) with apostrophes
/// doubled — Excel quotes a 3D span as a whole, never per half.
/// Caller frees.
fn encodeSpanPair(allocator: std.mem.Allocator, halves: SpanHalves) Error![]u8 {
    assert(halves.first.len >= 1);
    assert(halves.second.len >= 1);

    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(allocator);

    const bare = canEmitUnquoted(halves.first) and canEmitUnquoted(halves.second);
    if (!bare) try out.append(allocator, '\'');
    for ([2][]const u8{ halves.first, halves.second }, 0..) |half, idx| {
        if (idx == 1) try out.append(allocator, ':');
        for (half) |c| {
            if (!bare and c == '\'') try out.append(allocator, '\'');
            try out.append(allocator, c);
        }
    }
    if (!bare) try out.append(allocator, '\'');
    return out.toOwnedSlice(allocator);
}

/// Replace the single sheet-qualifier token at `idx` with an encoded
/// span pair. The unquoted pair keeps `.name` despite the embedded
/// colon — the pattern is fully consumed by its rewriter, so the
/// kind is inert; only the text reaches the printer.
fn emitSpanPairToken(
    allocator: std.mem.Allocator,
    work: []Token,
    owned: *std.ArrayListUnmanaged([]u8),
    idx: usize,
    halves: SpanHalves,
) Error!void {
    const merged = try registerOwned(allocator, owned, try encodeSpanPair(allocator, halves));
    work[idx] = .{
        .kind = if (merged[0] == '\'') .sheet_name else .name,
        .text = merged,
    };
}

/// Re-emit rewritten span endpoints into the unquoted 3-token
/// spelling's slots. When either new half needs quoting, the pair
/// merges into ONE `.sheet_name` token (Excel quotes the span as a
/// unit) and the operator + second slots become empty whitespace.
fn emitSpanEndpoints(
    allocator: std.mem.Allocator,
    work: []Token,
    owned: *std.ArrayListUnmanaged([]u8),
    first_idx: usize,
    second_idx: usize,
    halves: SpanHalves,
) Error!void {
    assert(second_idx == first_idx + 2);
    assert(work[first_idx + 1].kind == .op_range);
    if (canEmitUnquoted(halves.first) and canEmitUnquoted(halves.second)) {
        const f = try registerOwned(allocator, owned, try allocator.dupe(u8, halves.first));
        const s = try registerOwned(allocator, owned, try allocator.dupe(u8, halves.second));
        work[first_idx] = .{ .kind = .name, .text = f };
        work[second_idx] = .{ .kind = .name, .text = s };
        return;
    }
    try emitSpanPairToken(allocator, work, owned, first_idx, halves);
    work[first_idx + 1] = .{ .kind = .whitespace, .text = "" };
    work[second_idx] = .{ .kind = .whitespace, .text = "" };
}

/// Collapse tokens [lead, end_exclusive) to a single `#REF!`: the
/// lead token becomes the error literal, the rest become empty
/// whitespace (the printer concatenates `.text` verbatim, so empty
/// text contributes nothing).
fn collapseTokensToRef(work: []Token, lead: usize, end_exclusive: usize) void {
    assert(lead < end_exclusive);
    assert(end_exclusive <= work.len);
    work[lead] = .{ .kind = .error_lit, .text = "#REF!" };
    var k: usize = lead + 1;
    while (k < end_exclusive) : (k += 1) {
        work[k] = .{ .kind = .whitespace, .text = "" };
    }
}

/// Apply an edit to a 3D span (unquoted 3-token spelling). Row/col
/// shifts mutate only the trailing A1 part, scoped by span
/// membership of `target_sheet`. Rename rewrites matching endpoints
/// byte-exact (order-independent); delete contracts the span via
/// `deleteSpanOutcome`.
fn rewriteThreeD(
    allocator: std.mem.Allocator,
    work: []Token,
    owned: *std.ArrayListUnmanaged([]u8),
    ctx: RewriteContext,
    info: ThreeDQualifierInfo,
) Error!void {
    switch (ctx.edit) {
        .insert_rows, .delete_rows, .insert_cols, .delete_cols => {
            const target_match = blk: {
                const t = ctx.target_sheet orelse break :blk true;
                const first = try decodeSheetToken(allocator, work[info.first_idx]);
                defer allocator.free(first);
                const second = try decodeSheetToken(allocator, work[info.second_idx]);
                defer allocator.free(second);
                break :blk spanContainsTarget(first, second, t, ctx.sheet_order);
            };
            if (!target_match) return;

            if (info.axis) |axis| {
                try applyAxisRange(allocator, work, owned, ctx.edit, info.ref_start, axis);
            } else {
                try applyToRefRange(allocator, work, owned, ctx.edit, info.ref_start, info.is_range);
            }
        },
        .rename_sheet => |spec| {
            const first = try decodeSheetToken(allocator, work[info.first_idx]);
            defer allocator.free(first);
            const second = try decodeSheetToken(allocator, work[info.second_idx]);
            defer allocator.free(second);
            const halves = renameSpanEndpoints(first, second, spec) orelse return;
            try emitSpanEndpoints(allocator, work, owned, info.first_idx, info.second_idx, halves);
        },
        .delete_sheet => |doomed| {
            const first = try decodeSheetToken(allocator, work[info.first_idx]);
            defer allocator.free(first);
            const second = try decodeSheetToken(allocator, work[info.second_idx]);
            defer allocator.free(second);
            switch (deleteSpanOutcome(first, second, doomed, ctx.sheet_order)) {
                .unchanged => {},
                .collapse => collapseTokensToRef(work, info.first_idx, info.end),
                .contract => |halves| try emitSpanEndpoints(
                    allocator,
                    work,
                    owned,
                    info.first_idx,
                    info.second_idx,
                    halves,
                ),
            }
        },
        // A 3D span qualifies plain refs, never structured ones.
        .rename_table_column => {},
    }
}

fn rewriteSheetQualified(
    allocator: std.mem.Allocator,
    work: []Token,
    owned: *std.ArrayListUnmanaged([]u8),
    ctx: RewriteContext,
    info: SheetQualifierInfo,
) Error!void {
    assert(info.ref_start < work.len);
    assert(info.bang_idx == info.sheet_idx + 1);

    // Decode sheet name to plain text.
    const decoded = try decodeSheetToken(allocator, work[info.sheet_idx]);
    defer allocator.free(decoded);

    // Sheet rename.
    var current_sheet: []const u8 = decoded;
    var renamed_holder: ?[]u8 = null;
    defer if (renamed_holder) |h| allocator.free(h);

    if (ctx.edit == .rename_sheet) {
        const spec = ctx.edit.rename_sheet;
        if (splitQuotedSpan(decoded)) |halves| {
            // Quoted 3D span spelling ('Jan:Mar'!A1 arrives as ONE
            // token). Endpoints rename byte-exact, mirroring the
            // 3-token spelling in `rewriteThreeD`. `current_sheet`
            // needn't track the result — the row/col targeting below
            // never runs for rename.
            if (renameSpanEndpoints(halves.first, halves.second, spec)) |renamed| {
                try emitSpanPairToken(allocator, work, owned, info.sheet_idx, renamed);
            }
        } else if (std.mem.eql(u8, decoded, spec.old)) {
            const new_lex = try registerOwned(allocator, owned, try encodeSheetName(allocator, spec.new));
            // The encoded form already wraps with quotes when needed,
            // so it goes straight into the token stream.
            work[info.sheet_idx] = .{
                .kind = if (canEmitUnquoted(spec.new)) .name else .sheet_name,
                .text = new_lex,
            };
            // Track the post-rename sheet name for downstream
            // row/col targeting decisions.
            renamed_holder = try allocator.dupe(u8, spec.new);
            current_sheet = renamed_holder.?;
        }
    }

    // Sheet delete: collapse the entire qualified ref to #REF! by
    // rewriting the sheet-qualifier token AND every ref token in
    // the range to `.error_ref = "#REF!"`. The post-format pass
    // squashes consecutive `.error_ref` tokens to a single
    // `#REF!`, but doing the squash here keeps the token stream
    // shape stable. Bare refs and refs to OTHER sheets are
    // unaffected — they're handled by the row/col path below
    // (or skipped via target_match).
    if (ctx.edit == .delete_sheet) {
        const target = ctx.edit.delete_sheet;
        if (splitQuotedSpan(decoded)) |halves| {
            // Quoted 3D span spelling: same outcome logic as the
            // 3-token spelling in `rewriteThreeD`.
            switch (deleteSpanOutcome(halves.first, halves.second, target, ctx.sheet_order)) {
                .unchanged => {},
                .collapse => collapseTokensToRef(work, info.sheet_idx, info.ref_end),
                .contract => |contracted| try emitSpanPairToken(
                    allocator,
                    work,
                    owned,
                    info.sheet_idx,
                    contracted,
                ),
            }
            return;
        }
        if (std.mem.eql(u8, decoded, target)) {
            // Collapse the entire qualified-ref token sequence to a
            // single `#REF!`.
            collapseTokensToRef(work, info.sheet_idx, info.ref_end);
            return;
        }
        // Different sheet — nothing to do (bare-ref / row-col path
        // is also a no-op for delete_sheet, see target_match below).
        return;
    }

    // Row/col edits apply only when the sheet matches the edit's
    // target_sheet. `target_sheet == null` means "apply everywhere."
    // `sheetTargetMatches` also covers the quoted 3D span spelling
    // ('Jan:Mar'!A1 arrives here as ONE .sheet_name token): the
    // target then matches either endpoint — or, with `sheet_order`,
    // any mid-span member.
    const target_match = blk: {
        switch (ctx.edit) {
            // rename/delete_sheet: already handled above.
            // rename_table_column: no rows/cols to shift.
            .rename_sheet, .delete_sheet, .rename_table_column => break :blk false,
            else => {},
        }
        if (ctx.target_sheet) |t| break :blk sheetTargetMatches(current_sheet, t, ctx.sheet_order);
        break :blk true;
    };
    if (!target_match) return;

    if (info.axis) |axis| {
        try applyAxisRange(allocator, work, owned, ctx.edit, info.ref_start, axis);
    } else {
        try applyToRefRange(
            allocator,
            work,
            owned,
            ctx.edit,
            info.ref_start,
            info.is_range,
        );
    }
}

/// Scoping rule shared by every BARE reference (cell or whole-axis):
/// bare refs are scoped to `on_sheet`, so the edit applies only when
/// `target_sheet` matches it (or either side is null = "everywhere").
fn bareEditApplies(ctx: RewriteContext) bool {
    switch (ctx.edit) {
        .rename_sheet => return false, // bare refs have no sheet to rename
        // delete_sheet only collapses qualified refs. Bare refs
        // are scoped to the formula's owning sheet, which is
        // necessarily a sheet that's NOT being deleted (the
        // deleted sheet's own formulas are dropped wholesale by
        // Workbook.deleteSheet, not rewritten).
        .delete_sheet => return false,
        // Bare CELL refs carry no column names; the structured-ref
        // walk has its own scoping (`owning_table`).
        .rename_table_column => return false,
        else => {},
    }
    if (ctx.target_sheet == null) return true;
    if (ctx.on_sheet == null) return true; // permissive default
    return casefold.excelSheetNameEql(ctx.on_sheet.?, ctx.target_sheet.?);
}

fn rewriteBareRefOrRange(
    allocator: std.mem.Allocator,
    work: []Token,
    owned: *std.ArrayListUnmanaged([]u8),
    ctx: RewriteContext,
    start: usize,
    end: usize,
) Error!void {
    assert(start < end);
    assert(end <= work.len);

    if (!bareEditApplies(ctx)) return;

    const is_range = (end - start == 3);
    try applyToRefRange(allocator, work, owned, ctx.edit, start, is_range);
}

/// Apply a row/col edit to a single ref or to a `ref op_range ref`
/// triple. `ref_start` indexes the (first) `.cell_ref` token; if
/// `is_range` is true, `ref_start + 2` is the second endpoint.
fn applyToRefRange(
    allocator: std.mem.Allocator,
    work: []Token,
    owned: *std.ArrayListUnmanaged([]u8),
    edit: RewriteEdit,
    ref_start: usize,
    is_range: bool,
) Error!void {
    assert(ref_start < work.len);
    assert(work[ref_start].kind == .cell_ref);
    if (is_range) {
        assert(ref_start + 2 < work.len);
        assert(work[ref_start + 2].kind == .cell_ref);
    }

    if (edit == .rename_sheet) return;
    // delete_sheet only collapses qualified refs (handled in
    // rewriteSheetQualified). At this point we're processing bare
    // or already-shifted refs, neither of which delete_sheet
    // touches.
    if (edit == .delete_sheet) return;

    const a = parseRef(work[ref_start].text) orelse return;
    const a_out = applyShift(a, edit);

    if (!is_range) {
        switch (a_out) {
            .unchanged => return,
            .shifted => |new_ref| {
                try emitShifted(allocator, owned, &work[ref_start], new_ref);
            },
            .deleted => {
                // Single ref deleted → single `#REF!` token.
                work[ref_start] = .{ .kind = .error_lit, .text = "#REF!" };
            },
        }
        return;
    }

    const b = parseRef(work[ref_start + 2].text) orelse return;
    const b_out = applyShift(b, edit);

    const both_deleted = a_out == .deleted and b_out == .deleted;
    if (both_deleted) {
        // Collapse `A:B` → `#REF!`. Overwrite first slot, then
        // null-out the op and second endpoint with empty `unknown`
        // tokens so the printer emits nothing for them. (Empty
        // `.unknown` text contributes 0 bytes — round-trip safe.)
        work[ref_start] = .{ .kind = .error_lit, .text = "#REF!" };
        work[ref_start + 1] = .{ .kind = .unknown, .text = "" };
        work[ref_start + 2] = .{ .kind = .unknown, .text = "" };
        return;
    }

    switch (a_out) {
        .unchanged => {},
        .shifted => |new_ref| {
            try emitShifted(allocator, owned, &work[ref_start], new_ref);
        },
        .deleted => {
            work[ref_start] = .{ .kind = .error_lit, .text = "#REF!" };
        },
    }
    switch (b_out) {
        .unchanged => {},
        .shifted => |new_ref| {
            try emitShifted(allocator, owned, &work[ref_start + 2], new_ref);
        },
        .deleted => {
            work[ref_start + 2] = .{ .kind = .error_lit, .text = "#REF!" };
        },
    }
}

/// Apply a row/col edit to a full-column / full-row span occupying
/// tokens [ref_start, ref_start+3). A bound that kept its position
/// keeps its borrowed lexeme (case and spelling preserved); a moved
/// bound re-emits canonically with its `$` marker intact. Full
/// deletion collapses the three tokens to one `#REF!`, mirroring
/// the cell-range collapse shape.
fn applyAxisRange(
    allocator: std.mem.Allocator,
    work: []Token,
    owned: *std.ArrayListUnmanaged([]u8),
    edit: RewriteEdit,
    ref_start: usize,
    axis: AxisRange,
) Error!void {
    assert(ref_start + 2 < work.len);
    assert(work[ref_start + 1].kind == .op_range);

    switch (applyAxisShift(axis, edit)) {
        .unchanged => {},
        .shifted => |new_axis| {
            if (new_axis.first.n != axis.first.n) {
                work[ref_start].text = try registerOwned(
                    allocator,
                    owned,
                    try formatAxisBound(allocator, axis.kind, new_axis.first),
                );
            }
            if (new_axis.second.n != axis.second.n) {
                work[ref_start + 2].text = try registerOwned(
                    allocator,
                    owned,
                    try formatAxisBound(allocator, axis.kind, new_axis.second),
                );
            }
        },
        .deleted => {
            work[ref_start] = .{ .kind = .error_lit, .text = "#REF!" };
            work[ref_start + 1] = .{ .kind = .unknown, .text = "" };
            work[ref_start + 2] = .{ .kind = .unknown, .text = "" };
        },
    }
}

/// Allocate the new ref text, register it for later free, and
/// rewrite the token. The allocate-then-append sequence has a
/// leak window if `append` fails after `formatRef` succeeds —
/// `registerOwned` covers it without the `errdefer` being live
/// once the slice has been transferred to `owned`.
fn emitShifted(
    allocator: std.mem.Allocator,
    owned: *std.ArrayListUnmanaged([]u8),
    tok: *Token,
    new_ref: Ref,
) Error!void {
    const new_text = try registerOwned(allocator, owned, try formatRef(allocator, new_ref));
    tok.text = new_text;
}

/// Take ownership of `s` by appending it to `owned`. On allocation
/// failure inside `append`, free `s` and propagate. Returns `s`
/// unchanged on success — caller may use the slice immediately;
/// `owned` is responsible for freeing it later.
///
/// This is the canonical "allocate then transfer" idiom: the
/// `errdefer` is bounded to the function body, so it doesn't fire
/// on errors raised AFTER the function returns successfully.
fn registerOwned(
    allocator: std.mem.Allocator,
    owned: *std.ArrayListUnmanaged([]u8),
    s: []u8,
) Error![]u8 {
    errdefer allocator.free(s);
    try owned.append(allocator, s);
    return s;
}

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

fn expectRewrite(input: []const u8, ctx: RewriteContext, expected: []const u8) !void {
    const out = try rewriteFormula(testing.allocator, input, ctx);
    defer testing.allocator.free(out);
    try testing.expectEqualStrings(expected, out);
}

test "insert_rows shifts refs at or below the insertion point" {
    // Insert one row at row 3: A5 → A6, B2 unchanged.
    try expectRewrite(
        "A5+B2",
        .{ .edit = .{ .insert_rows = .{ .at = 3, .count = 1 } } },
        "A6+B2",
    );
}

test "delete_rows shifts post-range refs and deletes in-range refs" {
    // Delete rows 5..6: A10 (>= 7) shifts -2 → A8.
    try expectRewrite(
        "A10",
        .{ .edit = .{ .delete_rows = .{ .at = 5, .count = 2 } } },
        "A8",
    );
    // A5 is in [5, 7) → #REF!.
    try expectRewrite(
        "A5",
        .{ .edit = .{ .delete_rows = .{ .at = 5, .count = 2 } } },
        "#REF!",
    );
    // A4 is below `at` → unchanged.
    try expectRewrite(
        "A4",
        .{ .edit = .{ .delete_rows = .{ .at = 5, .count = 2 } } },
        "A4",
    );
}

test "insert_cols shifts column letters" {
    // Insert one column at B (col 2): C1 → D1, A1 unchanged.
    try expectRewrite(
        "C1+A1",
        .{ .edit = .{ .insert_cols = .{ .at = 2, .count = 1 } } },
        "D1+A1",
    );
    // Multi-letter column: AA1 (col 27) + insert 1 at A → AB1.
    try expectRewrite(
        "AA1",
        .{ .edit = .{ .insert_cols = .{ .at = 1, .count = 1 } } },
        "AB1",
    );
}

test "delete_cols shifts post-range refs and deletes in-range refs" {
    // Delete cols 2..3 (B, C): D5 → B5.
    try expectRewrite(
        "D5",
        .{ .edit = .{ .delete_cols = .{ .at = 2, .count = 2 } } },
        "B5",
    );
    // B5 is in [2, 4) → #REF!.
    try expectRewrite(
        "B5",
        .{ .edit = .{ .delete_cols = .{ .at = 2, .count = 2 } } },
        "#REF!",
    );
}

test "absolute markers preserved across shifts" {
    // $A$5 + insert_rows{3,1} → $A$6 (both `$`s preserved).
    try expectRewrite(
        "$A$5",
        .{ .edit = .{ .insert_rows = .{ .at = 3, .count = 1 } } },
        "$A$6",
    );
    // Mixed absolute: $A1, A$1.
    try expectRewrite(
        "$A1",
        .{ .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } } },
        "$A2",
    );
    try expectRewrite(
        "A$5",
        .{ .edit = .{ .insert_rows = .{ .at = 3, .count = 1 } } },
        "A$6",
    );
    // Absolute column shifts when columns insert before it.
    try expectRewrite(
        "$C$5",
        .{ .edit = .{ .insert_cols = .{ .at = 2, .count = 1 } } },
        "$D$5",
    );
}

test "ranges shift both endpoints" {
    try expectRewrite(
        "A1:B10",
        .{ .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } } },
        "A2:B11",
    );
    // Both endpoints land in deleted territory → whole range #REF!.
    try expectRewrite(
        "A5:A6",
        .{ .edit = .{ .delete_rows = .{ .at = 5, .count = 2 } } },
        "#REF!",
    );
    // Partial deletion: A4 unchanged (below `at`), A5 deleted.
    try expectRewrite(
        "A4:A5",
        .{ .edit = .{ .delete_rows = .{ .at = 5, .count = 1 } } },
        "A4:#REF!",
    );
}

test "delete_sheet collapses qualified refs to #REF!" {
    // Plain qualified ref to the deleted sheet → #REF!.
    try expectRewrite(
        "Doomed!A1",
        .{ .edit = .{ .delete_sheet = "Doomed" } },
        "#REF!",
    );
    // Qualified ref to a DIFFERENT sheet — leave alone.
    try expectRewrite(
        "Survivor!A1",
        .{ .edit = .{ .delete_sheet = "Doomed" } },
        "Survivor!A1",
    );
    // Bare ref — unaffected (deleted sheet's own formulas are
    // dropped wholesale, not rewritten).
    try expectRewrite(
        "A1",
        .{ .edit = .{ .delete_sheet = "Doomed" } },
        "A1",
    );
    // Range qualified to the deleted sheet → #REF!.
    try expectRewrite(
        "Doomed!A1:B5",
        .{ .edit = .{ .delete_sheet = "Doomed" } },
        "#REF!",
    );
    // Quoted sheet name with apostrophes.
    try expectRewrite(
        "'Bob''s Place'!A1",
        .{ .edit = .{ .delete_sheet = "Bob's Place" } },
        "#REF!",
    );
    // Multiple refs in one formula — only matching qualifier collapses.
    try expectRewrite(
        "Doomed!A1+Survivor!B2",
        .{ .edit = .{ .delete_sheet = "Doomed" } },
        "#REF!+Survivor!B2",
    );
}

test "rename_sheet rewrites quoted sheet names" {
    try expectRewrite(
        "'My Sheet'!A1",
        .{ .edit = .{ .rename_sheet = .{ .old = "My Sheet", .new = "Renamed" } } },
        "Renamed!A1",
    );
    // Old name didn't match — leave alone.
    try expectRewrite(
        "'Other'!A1",
        .{ .edit = .{ .rename_sheet = .{ .old = "My Sheet", .new = "Renamed" } } },
        "'Other'!A1",
    );
    // Unquoted source → quoted destination when new name has spaces.
    try expectRewrite(
        "Sheet1!A1",
        .{ .edit = .{ .rename_sheet = .{ .old = "Sheet1", .new = "New Name" } } },
        "'New Name'!A1",
    );
}

test "apostrophe-escaped sheet names round-trip" {
    // 'It''s'!A1 — embedded apostrophe, decoded as `It's`. Rename to
    // a name with another apostrophe and verify the doubling persists.
    try expectRewrite(
        "'It''s'!A1",
        .{ .edit = .{ .rename_sheet = .{ .old = "It's", .new = "Bob's Place" } } },
        "'Bob''s Place'!A1",
    );
    // No rename, no row/col edit on the qualified ref's sheet:
    // bytes survive verbatim.
    try expectRewrite(
        "'It''s'!A1",
        .{
            .target_sheet = "DifferentSheet",
            .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } },
        },
        "'It''s'!A1",
    );
}

test "row/col target_sheet matches sheet names per Excel case rules" {
    // Excel resolves sheet names without regard to case, so a
    // qualifier spelled in a different case than the workbook's
    // canonical name still refers to the edited sheet and must
    // shift. `casefold.excelSheetNameEql` — Unicode fold + NFC.
    try expectRewrite(
        "edited!A2*2",
        .{
            .target_sheet = "Edited",
            .edit = .{ .insert_rows = .{ .at = 2, .count = 1 } },
        },
        "edited!A3*2",
    );
    // Non-ASCII case pair (quoted spelling pins the decode path).
    try expectRewrite(
        "'café'!A2*2",
        .{
            .target_sheet = "CAFÉ",
            .edit = .{ .insert_rows = .{ .at = 2, .count = 1 } },
        },
        "'café'!A3*2",
    );
    // Quoted-combined 3D span: either endpoint matches, case-folded.
    try expectRewrite(
        "'café:Zeta'!A2",
        .{
            .target_sheet = "CAFÉ",
            .edit = .{ .insert_rows = .{ .at = 2, .count = 1 } },
        },
        "'café:Zeta'!A3",
    );
    // Unquoted 3D span (two name tokens) folds its endpoints too.
    try expectRewrite(
        "Alpha:Zeta!A2",
        .{
            .target_sheet = "ZETA",
            .edit = .{ .insert_rows = .{ .at = 2, .count = 1 } },
        },
        "Alpha:Zeta!A3",
    );
    // Bare refs: on_sheet vs target_sheet folds the same way.
    try expectRewrite(
        "A2+1",
        .{
            .on_sheet = "EDITED",
            .target_sheet = "Edited",
            .edit = .{ .insert_rows = .{ .at = 2, .count = 1 } },
        },
        "A3+1",
    );
    try expectRewrite(
        "A2+1",
        .{
            .on_sheet = "café",
            .target_sheet = "CAFÉ",
            .edit = .{ .insert_rows = .{ .at = 2, .count = 1 } },
        },
        "A3+1",
    );
}

test "sheet-qualified ranges shift both endpoints" {
    try expectRewrite(
        "Sheet1!A1:B10",
        .{ .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } } },
        "Sheet1!A2:B11",
    );
    try expectRewrite(
        "'My Sheet'!$A$5:$B$6",
        .{ .edit = .{ .delete_rows = .{ .at = 5, .count = 2 } } },
        "'My Sheet'!#REF!",
    );
}

test "rewrite: 3D refs shift on row edits" {
    try expectRewrite(
        "SUM(Sheet1:Sheet3!A5)",
        .{ .edit = .{ .insert_rows = .{ .at = 3, .count = 1 } } },
        "SUM(Sheet1:Sheet3!A6)",
    );
    // Range: both endpoints of the cell part shift.
    try expectRewrite(
        "Sheet1:Sheet3!A1:B10",
        .{ .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } } },
        "Sheet1:Sheet3!A2:B11",
    );
    // Deleted-territory collapse mirrors the single-sheet behavior:
    // the span survives, the cell part becomes #REF!.
    try expectRewrite(
        "Sheet1:Sheet3!A5",
        .{ .edit = .{ .delete_rows = .{ .at = 5, .count = 2 } } },
        "Sheet1:Sheet3!#REF!",
    );
    try expectRewrite(
        "Sheet1:Sheet3!A5:A6",
        .{ .edit = .{ .delete_rows = .{ .at = 5, .count = 2 } } },
        "Sheet1:Sheet3!#REF!",
    );
    // Post-range refs shift back.
    try expectRewrite(
        "Sheet1:Sheet3!A10",
        .{ .edit = .{ .delete_rows = .{ .at = 5, .count = 2 } } },
        "Sheet1:Sheet3!A8",
    );
}

test "rewrite: 3D refs shift on column edits" {
    try expectRewrite(
        "Sheet1:Sheet3!C1",
        .{ .edit = .{ .insert_cols = .{ .at = 2, .count = 1 } } },
        "Sheet1:Sheet3!D1",
    );
    try expectRewrite(
        "Sheet1:Sheet3!$B$5",
        .{ .edit = .{ .delete_cols = .{ .at = 2, .count = 2 } } },
        "Sheet1:Sheet3!#REF!",
    );
    // Absolute markers survive the shift.
    try expectRewrite(
        "Sheet1:Sheet3!$C$5:$D$6",
        .{ .edit = .{ .insert_cols = .{ .at = 2, .count = 1 } } },
        "Sheet1:Sheet3!$D$5:$E$6",
    );
}

test "rewrite: 3D quoted span spellings" {
    // Excel's canonical quoted form wraps the WHOLE span in one
    // quote pair — it arrives as a single .sheet_name token.
    try expectRewrite(
        "SUM('Jan Sales:Mar Sales'!A5)",
        .{ .edit = .{ .insert_rows = .{ .at = 3, .count = 1 } } },
        "SUM('Jan Sales:Mar Sales'!A6)",
    );
    // Per-endpoint quoting and the mixed spelling tokenize as
    // separate endpoint tokens; both shift the same way.
    try expectRewrite(
        "'Jan Sales':'Mar Sales'!A5",
        .{ .edit = .{ .insert_rows = .{ .at = 3, .count = 1 } } },
        "'Jan Sales':'Mar Sales'!A6",
    );
    try expectRewrite(
        "'Jan Sales':Mar!A5:B6",
        .{ .edit = .{ .delete_rows = .{ .at = 5, .count = 2 } } },
        "'Jan Sales':Mar!#REF!",
    );
    // Apostrophe-escaped endpoint names decode before comparison.
    try expectRewrite(
        "'It''s':Sheet3!A5",
        .{
            .target_sheet = "It's",
            .edit = .{ .insert_rows = .{ .at = 3, .count = 1 } },
        },
        "'It''s':Sheet3!A6",
    );
}

test "rewrite: 3D target_sheet scopes to span endpoints" {
    // Either endpoint of the span matches a named target.
    try expectRewrite(
        "Sheet1:Sheet3!A5",
        .{
            .target_sheet = "Sheet1",
            .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } },
        },
        "Sheet1:Sheet3!A6",
    );
    try expectRewrite(
        "Sheet1:Sheet3!A5",
        .{
            .target_sheet = "Sheet3",
            .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } },
        },
        "Sheet1:Sheet3!A6",
    );
    // A mid-span sheet contains the edit too, but membership needs
    // the workbook's sheet order, which this layer cannot see: the
    // ref stays put rather than risking a wrong rewrite.
    try expectRewrite(
        "Sheet1:Sheet3!A5",
        .{
            .target_sheet = "Sheet2",
            .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } },
        },
        "Sheet1:Sheet3!A5",
    );
    // The quoted-combined spelling scopes by endpoint the same way.
    try expectRewrite(
        "'Jan:Mar'!A5",
        .{
            .target_sheet = "Jan",
            .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } },
        },
        "'Jan:Mar'!A6",
    );
    try expectRewrite(
        "'Jan:Mar'!A5",
        .{
            .target_sheet = "Feb",
            .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } },
        },
        "'Jan:Mar'!A5",
    );
}

test "rewrite: 3D endpoint rename rewrites span endpoints" {
    const order = [_][]const u8{ "Sheet1", "Sheet2", "Sheet3" };
    // First endpoint — the case the pre-order fallthrough could
    // never reach (only `Sheet3!A1` matched the single-qualifier
    // path). Order-independent: null `sheet_order`.
    try expectRewrite(
        "SUM(Sheet1:Sheet3!A1)",
        .{ .edit = .{ .rename_sheet = .{ .old = "Sheet1", .new = "Start" } } },
        "SUM(Start:Sheet3!A1)",
    );
    // Second endpoint, with the order supplied.
    try expectRewrite(
        "Sheet1:Sheet3!A1:B2",
        .{
            .sheet_order = &order,
            .edit = .{ .rename_sheet = .{ .old = "Sheet3", .new = "End" } },
        },
        "Sheet1:End!A1:B2",
    );
    // Both endpoints spell the old name.
    try expectRewrite(
        "Doomed:Doomed!A1",
        .{ .edit = .{ .rename_sheet = .{ .old = "Doomed", .new = "Kept" } } },
        "Kept:Kept!A1",
    );
    // Mid-span rename: endpoints are the only names ever spelled, so
    // the text is exactly unchanged.
    try expectRewrite(
        "Sheet1:Sheet3!A1",
        .{
            .sheet_order = &order,
            .edit = .{ .rename_sheet = .{ .old = "Sheet2", .new = "Mid" } },
        },
        "Sheet1:Sheet3!A1",
    );
    // Endpoint matching stays byte-exact — the pinned single-
    // qualifier rename semantics extend to span endpoints.
    try expectRewrite(
        "sheet1:Sheet3!A1",
        .{ .edit = .{ .rename_sheet = .{ .old = "Sheet1", .new = "X" } } },
        "sheet1:Sheet3!A1",
    );
    // A new name needing quotes merges the pair into one quoted
    // unit — Excel quotes a 3D span as a whole, never per half.
    try expectRewrite(
        "Sheet1:Sheet3!A1",
        .{ .edit = .{ .rename_sheet = .{ .old = "Sheet3", .new = "Q3 End" } } },
        "'Sheet1:Q3 End'!A1",
    );
    // Apostrophes double inside the merged quoted pair.
    try expectRewrite(
        "Jan:Mar!A1",
        .{ .edit = .{ .rename_sheet = .{ .old = "Mar", .new = "It's" } } },
        "'Jan:It''s'!A1",
    );
    // Quoted-pair spelling unquotes when both halves can stand bare...
    try expectRewrite(
        "'Jan:Mar'!A1",
        .{ .edit = .{ .rename_sheet = .{ .old = "Mar", .new = "Apr" } } },
        "Jan:Apr!A1",
    );
    // ...and stays a quoted pair when one half still needs it.
    try expectRewrite(
        "'Jan Sales:Mar'!A1",
        .{ .edit = .{ .rename_sheet = .{ .old = "Mar", .new = "Apr" } } },
        "'Jan Sales:Apr'!A1",
    );
    // Mixed spelling: renaming the quoted half to a bare-safe name
    // unquotes the whole span.
    try expectRewrite(
        "'Jan Sales':Mar!A1",
        .{ .edit = .{ .rename_sheet = .{ .old = "Jan Sales", .new = "Jan" } } },
        "Jan:Mar!A1",
    );
    // Whole-axis tail: endpoints rewrite the same way.
    try expectRewrite(
        "Sheet1:Sheet3!A:A",
        .{ .edit = .{ .rename_sheet = .{ .old = "Sheet1", .new = "Start" } } },
        "Start:Sheet3!A:A",
    );
}

test "rewrite: 3D endpoint delete contracts the span via sheet order" {
    const order = [_][]const u8{ "Sheet1", "Sheet2", "Sheet3" };
    // First endpoint: the interval minimum steps up to its
    // order-neighbor inside the span.
    try expectRewrite(
        "SUM(Sheet1:Sheet3!A1)",
        .{ .sheet_order = &order, .edit = .{ .delete_sheet = "Sheet1" } },
        "SUM(Sheet2:Sheet3!A1)",
    );
    // Second endpoint: the maximum steps down.
    try expectRewrite(
        "Sheet1:Sheet3!A1:B2",
        .{ .sheet_order = &order, .edit = .{ .delete_sheet = "Sheet3" } },
        "Sheet1:Sheet2!A1:B2",
    );
    // Mid-span delete: never spelled, text exactly unchanged.
    try expectRewrite(
        "Sheet1:Sheet3!A1",
        .{ .sheet_order = &order, .edit = .{ .delete_sheet = "Sheet2" } },
        "Sheet1:Sheet3!A1",
    );
    // Outside the span: unchanged.
    try expectRewrite(
        "Sheet1:Sheet2!A1",
        .{ .sheet_order = &order, .edit = .{ .delete_sheet = "Sheet3" } },
        "Sheet1:Sheet2!A1",
    );
    // Two-sheet span contracts to a single-sheet span; the span FORM
    // is kept (`Sheet2:Sheet2` evaluates identically to `Sheet2`).
    const two = [_][]const u8{ "Sheet1", "Sheet2" };
    try expectRewrite(
        "Sheet1:Sheet2!A1",
        .{ .sheet_order = &two, .edit = .{ .delete_sheet = "Sheet1" } },
        "Sheet2:Sheet2!A1",
    );
    // Single-position span wholly deleted → #REF!, decidable with no
    // order at all.
    try expectRewrite(
        "SUM(Doomed:Doomed!A1)",
        .{ .edit = .{ .delete_sheet = "Doomed" } },
        "SUM(#REF!)",
    );
    // Reversed written order keeps its orientation; the replacement
    // lands on the token that spelled the doomed name.
    try expectRewrite(
        "Sheet3:Sheet1!A1",
        .{ .sheet_order = &order, .edit = .{ .delete_sheet = "Sheet3" } },
        "Sheet2:Sheet1!A1",
    );
    // Quoted-pair spelling contracts too, with canonical emission.
    const months = [_][]const u8{ "Jan", "Feb", "Mar" };
    try expectRewrite(
        "'Jan:Mar'!A1",
        .{ .sheet_order = &months, .edit = .{ .delete_sheet = "Mar" } },
        "Jan:Feb!A1",
    );
    // Contraction onto a name that needs quoting emits a quoted
    // pair; whole-axis tail rewrites the same way.
    const spaced = [_][]const u8{ "Jan", "Feb 2", "Mar" };
    try expectRewrite(
        "Jan:Mar!A:A",
        .{ .sheet_order = &spaced, .edit = .{ .delete_sheet = "Mar" } },
        "'Jan:Feb 2'!A:A",
    );
    // Order lookup resolves with Excel's case rule; only the edit's
    // own name match is byte-exact.
    const shouty = [_][]const u8{ "SHEET1", "SHEET2", "SHEET3" };
    try expectRewrite(
        "Sheet1:Sheet3!A1",
        .{ .sheet_order = &shouty, .edit = .{ .delete_sheet = "Sheet3" } },
        "Sheet1:SHEET2!A1",
    );
    // Byte-exact endpoint matching, pinned: a case-variant spelling
    // of the doomed sheet is NOT an endpoint match.
    try expectRewrite(
        "Sheet1:SHEET3!A1",
        .{ .sheet_order = &order, .edit = .{ .delete_sheet = "Sheet3" } },
        "Sheet1:SHEET3!A1",
    );
    // No order: an endpoint delete cannot locate its neighbor — the
    // span stays untouched (NOT the pre-order `Sheet1:#REF!`
    // fallthrough corruption).
    try expectRewrite(
        "Sheet1:Sheet3!A1",
        .{ .edit = .{ .delete_sheet = "Sheet3" } },
        "Sheet1:Sheet3!A1",
    );
    // Endpoint missing from the order: undecidable, untouched.
    try expectRewrite(
        "Sheet1:Sheet9!A1",
        .{ .sheet_order = &order, .edit = .{ .delete_sheet = "Sheet1" } },
        "Sheet1:Sheet9!A1",
    );
    // Case-variant duplicates in a malformed order: the byte-exact
    // entry wins, so the interval anchors at `sheet1` (index 2) and
    // contraction steps to its true neighbor — never through the
    // case-fold twin's interval (Codex r1 finding 2).
    const dup = [_][]const u8{ "Sheet1", "X", "sheet1", "Y", "Sheet3" };
    try expectRewrite(
        "sheet1:Sheet3!A1",
        .{ .sheet_order = &dup, .edit = .{ .delete_sheet = "sheet1" } },
        "Y:Sheet3!A1",
    );
    // No byte-exact entry and two case-fold candidates: ambiguous,
    // untouched.
    try expectRewrite(
        "SHEET1:Sheet3!A1",
        .{ .sheet_order = &dup, .edit = .{ .delete_sheet = "SHEET1" } },
        "SHEET1:Sheet3!A1",
    );
}

test "rewrite: 3D mid-span target matches with sheet order" {
    const order = [_][]const u8{ "Sheet1", "Sheet2", "Sheet3" };
    // The conservative no-order case is pinned in "3D target_sheet
    // scopes to span endpoints"; with the order, membership is
    // decidable and the shift lands.
    try expectRewrite(
        "Sheet1:Sheet3!A5",
        .{
            .sheet_order = &order,
            .target_sheet = "Sheet2",
            .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } },
        },
        "Sheet1:Sheet3!A6",
    );
    // Quoted spelling gets the same membership rule.
    const months = [_][]const u8{ "Jan", "Feb", "Mar" };
    try expectRewrite(
        "'Jan:Mar'!A5",
        .{
            .sheet_order = &months,
            .target_sheet = "Feb",
            .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } },
        },
        "'Jan:Mar'!A6",
    );
    // A target outside the span still does not match.
    const wide = [_][]const u8{ "Sheet1", "Sheet2", "Sheet3", "Sheet4" };
    try expectRewrite(
        "Sheet1:Sheet3!A5",
        .{
            .sheet_order = &wide,
            .target_sheet = "Sheet4",
            .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } },
        },
        "Sheet1:Sheet3!A5",
    );
}

test "rewrite: 3D refs clamp at the grid caps" {
    // Off-grid shifts follow the leave-alone policy the bare-ref
    // path already has (see "insert beyond grid leaves ref alone").
    try expectRewrite(
        "Sheet1:Sheet3!XFD1",
        .{ .edit = .{ .insert_cols = .{ .at = 1, .count = 1 } } },
        "Sheet1:Sheet3!XFD1",
    );
    try expectRewrite(
        "Sheet1:Sheet3!A1048576",
        .{ .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } } },
        "Sheet1:Sheet3!A1048576",
    );
}

test "rewrite: full-column spans shift on column edits" {
    try expectRewrite(
        "SUM(A:A)",
        .{ .edit = .{ .insert_cols = .{ .at = 1, .count = 1 } } },
        "SUM(B:B)",
    );
    // Insert inside the span: only the bounds at/past the insertion
    // point move.
    try expectRewrite(
        "A:C",
        .{ .edit = .{ .insert_cols = .{ .at = 2, .count = 1 } } },
        "A:D",
    );
    // Delete entirely before the span: both bounds slide down.
    try expectRewrite(
        "C:E",
        .{ .edit = .{ .delete_cols = .{ .at = 1, .count = 1 } } },
        "B:D",
    );
    // Lowercase bounds re-emit canonically once shifted (same
    // normalization trade-off as the cell path).
    try expectRewrite(
        "a:c",
        .{ .edit = .{ .insert_cols = .{ .at = 1, .count = 1 } } },
        "B:D",
    );
}

test "rewrite: full-column spans shrink and collapse on delete" {
    // Every spanned column deleted → the whole span is #REF!.
    try expectRewrite(
        "A:C",
        .{ .edit = .{ .delete_cols = .{ .at = 1, .count = 3 } } },
        "#REF!",
    );
    // Partial deletion SHRINKS the span (interval semantics — there
    // is no printable `A:#REF!`, unlike the cell-range path).
    try expectRewrite(
        "A:C",
        .{ .edit = .{ .delete_cols = .{ .at = 2, .count = 1 } } },
        "A:B",
    );
    // Zone overshoots the span's end: survivors only.
    try expectRewrite(
        "A:C",
        .{ .edit = .{ .delete_cols = .{ .at = 2, .count = 5 } } },
        "A:A",
    );
    // Zone eats the span's head: the survivor lands at the zone's
    // start position.
    try expectRewrite(
        "C:E",
        .{ .edit = .{ .delete_cols = .{ .at = 1, .count = 4 } } },
        "A:A",
    );
    // Reversed spelling keeps its written order positionally.
    try expectRewrite(
        "E:C",
        .{ .edit = .{ .delete_cols = .{ .at = 4, .count = 5 } } },
        "C:C",
    );
}

test "rewrite: full-row spans shift, shrink and collapse on row edits" {
    try expectRewrite(
        "1:5",
        .{ .edit = .{ .insert_rows = .{ .at = 1, .count = 2 } } },
        "3:7",
    );
    try expectRewrite(
        "10:20",
        .{ .edit = .{ .delete_rows = .{ .at = 1, .count = 2 } } },
        "8:18",
    );
    try expectRewrite(
        "1:5",
        .{ .edit = .{ .delete_rows = .{ .at = 2, .count = 2 } } },
        "1:3",
    );
    try expectRewrite(
        "3:5",
        .{ .edit = .{ .delete_rows = .{ .at = 1, .count = 4 } } },
        "1:1",
    );
    try expectRewrite(
        "SUM(1:1)",
        .{ .edit = .{ .delete_rows = .{ .at = 1, .count = 1 } } },
        "SUM(#REF!)",
    );
    // Edit entirely past the span: untouched.
    try expectRewrite(
        "2:3",
        .{ .edit = .{ .insert_rows = .{ .at = 5, .count = 1 } } },
        "2:3",
    );
}

test "rewrite: full-span absolute markers preserved" {
    try expectRewrite(
        "$A:$C",
        .{ .edit = .{ .insert_cols = .{ .at = 1, .count = 1 } } },
        "$B:$D",
    );
    try expectRewrite(
        "$1:$1",
        .{ .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } } },
        "$2:$2",
    );
    // Mixed markers: each bound keeps its own spelling; an unmoved
    // bound keeps its borrowed lexeme verbatim.
    try expectRewrite(
        "$A:C",
        .{ .edit = .{ .insert_cols = .{ .at = 2, .count = 1 } } },
        "$A:D",
    );
    try expectRewrite(
        "1:$5",
        .{ .edit = .{ .delete_rows = .{ .at = 2, .count = 2 } } },
        "1:$3",
    );
}

test "rewrite: full-span perpendicular edits leave the span alone" {
    // Row edits never reshape a column span…
    try expectRewrite(
        "A:A",
        .{ .edit = .{ .insert_rows = .{ .at = 1, .count = 5 } } },
        "A:A",
    );
    try expectRewrite(
        "A:C",
        .{ .edit = .{ .delete_rows = .{ .at = 1, .count = 10 } } },
        "A:C",
    );
    // …and column edits never reshape a row span.
    try expectRewrite(
        "1:5",
        .{ .edit = .{ .insert_cols = .{ .at = 1, .count = 5 } } },
        "1:5",
    );
    try expectRewrite(
        "1:5",
        .{ .edit = .{ .delete_cols = .{ .at = 1, .count = 10 } } },
        "1:5",
    );
}

test "rewrite: full-span sheet-qualified spellings" {
    try expectRewrite(
        "Sheet1!A:A",
        .{ .edit = .{ .insert_cols = .{ .at = 1, .count = 1 } } },
        "Sheet1!B:B",
    );
    // target_sheet scopes qualified spans exactly like qualified cells.
    try expectRewrite(
        "Sheet1!A:A+Sheet2!A:A",
        .{
            .target_sheet = "Sheet1",
            .edit = .{ .insert_cols = .{ .at = 1, .count = 1 } },
        },
        "Sheet1!B:B+Sheet2!A:A",
    );
    try expectRewrite(
        "'My Sheet'!1:5",
        .{ .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } } },
        "'My Sheet'!2:6",
    );
    // Qualifier-level edits reach axis spans too.
    try expectRewrite(
        "Doomed!A:A",
        .{ .edit = .{ .delete_sheet = "Doomed" } },
        "#REF!",
    );
    try expectRewrite(
        "Sheet1!1:5",
        .{ .edit = .{ .rename_sheet = .{ .old = "Sheet1", .new = "Renamed" } } },
        "Renamed!1:5",
    );
    // Bare spans scope to on_sheet, mirroring bare cell refs.
    try expectRewrite(
        "A:A",
        .{
            .on_sheet = "Sheet2",
            .target_sheet = "Sheet1",
            .edit = .{ .insert_cols = .{ .at = 1, .count = 1 } },
        },
        "A:A",
    );
    try expectRewrite(
        "A:A",
        .{
            .on_sheet = "Sheet1",
            .target_sheet = "Sheet1",
            .edit = .{ .insert_cols = .{ .at = 1, .count = 1 } },
        },
        "B:B",
    );
}

test "rewrite: full-span 3D spellings" {
    try expectRewrite(
        "SUM(Sheet1:Sheet3!A:A)",
        .{ .edit = .{ .insert_cols = .{ .at = 1, .count = 1 } } },
        "SUM(Sheet1:Sheet3!B:B)",
    );
    // The span endpoints survive an axis collapse, mirroring the
    // 3D cell-range behavior.
    try expectRewrite(
        "Sheet1:Sheet3!1:5",
        .{ .edit = .{ .delete_rows = .{ .at = 1, .count = 5 } } },
        "Sheet1:Sheet3!#REF!",
    );
    // Endpoint scoping: a named target matches either endpoint; a
    // mid-span target conservatively leaves the ref unchanged.
    try expectRewrite(
        "Sheet1:Sheet3!A:A",
        .{
            .target_sheet = "Sheet3",
            .edit = .{ .insert_cols = .{ .at = 1, .count = 1 } },
        },
        "Sheet1:Sheet3!B:B",
    );
    try expectRewrite(
        "Sheet1:Sheet3!A:A",
        .{
            .target_sheet = "Sheet2",
            .edit = .{ .insert_cols = .{ .at = 1, .count = 1 } },
        },
        "Sheet1:Sheet3!A:A",
    );
    // Quoted-combined span spelling scopes by endpoint the same way.
    try expectRewrite(
        "'Jan:Mar'!A:A",
        .{
            .target_sheet = "Jan",
            .edit = .{ .insert_cols = .{ .at = 1, .count = 1 } },
        },
        "'Jan:Mar'!B:B",
    );
}

test "rewrite: full-span grid caps" {
    // A bound the insert would push off-grid stays clamped at the
    // edge while the other bound still moves — matching Excel's
    // clamp-at-cap for ranges touching the last column/row.
    try expectRewrite(
        "A:XFD",
        .{ .edit = .{ .insert_cols = .{ .at = 1, .count = 1 } } },
        "B:XFD",
    );
    try expectRewrite(
        "XFD:XFD",
        .{ .edit = .{ .insert_cols = .{ .at = 1, .count = 1 } } },
        "XFD:XFD",
    );
    try expectRewrite(
        "1:1048576",
        .{ .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } } },
        "2:1048576",
    );
    try expectRewrite(
        "1048576:1048576",
        .{ .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } } },
        "1048576:1048576",
    );
}

test "rewrite: full-span lookalikes stay untouched" {
    // `name:name` pairs whose halves do NOT both spell in-grid
    // columns are defined-name unions (or malformed input) and pass
    // through byte-identically under every edit.
    for ([_][]const u8{
        "Start:End", // multi-letter names, not columns
        "SUM(Alpha:Omega)",
        "ZZZZ:A", // four letters — off the grid
        "XFE:XFE", // one past the last column
        "1.5:2", // non-integer number lexeme
        "1e3:5", // exponent number lexeme
        "1:B", // mixed axes
        "A:1",
        "0:5", // row 0 does not exist
        "1048577:1048578", // past the last row
        "\u{20AC}!A:A", // opaque `!` qualifier — unknowable sheet
    }) |c| try expectIdentityUnderEveryEdit(c);
    // The flip side of the shadowing rule: halves that DO spell
    // in-grid columns are a column span even when they read like
    // words — Excel's reference interpretation takes precedence
    // over defined names here.
    try expectRewrite(
        "Foo:Bar",
        .{ .edit = .{ .insert_cols = .{ .at = 1, .count = 1 } } },
        "FOP:BAS",
    );
}

test "target_sheet scopes bare refs to on_sheet match" {
    // on_sheet = "Sheet2", target_sheet = "Sheet1" → bare ref unchanged.
    try expectRewrite(
        "A5",
        .{
            .on_sheet = "Sheet2",
            .target_sheet = "Sheet1",
            .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } },
        },
        "A5",
    );
    // Same sheet — ref shifts.
    try expectRewrite(
        "A5",
        .{
            .on_sheet = "Sheet1",
            .target_sheet = "Sheet1",
            .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } },
        },
        "A6",
    );
}

test "target_sheet scopes sheet-qualified refs" {
    // Edit targets Sheet1 only; Sheet2!A5 stays put.
    try expectRewrite(
        "Sheet1!A5+Sheet2!A5",
        .{
            .target_sheet = "Sheet1",
            .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } },
        },
        "Sheet1!A6+Sheet2!A5",
    );
}

test "opaque tokens are not mutated" {
    // External-workbook ref: `.external_ref` since M1a, and the sheet
    // and cell it qualifies are inside a workbook we cannot see.
    try expectRewrite(
        "'[Book.xlsx]Sheet1'!A1+1",
        .{ .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } } },
        "'[Book.xlsx]Sheet1'!A1+1",
    );
    // Dynamic-array spill `A1#` — the `A1` is a live cell_ref and the
    // trailing `#` is `.op_spill`. The ref still shifts; the operator
    // is passed through.
    try expectRewrite(
        "A1#",
        .{ .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } } },
        "A2#",
    );
}

// ─── M1a compat gate: untouched-construct identity ────────────────
//
// Every construct the tokenizer classifies but the rewriter must not
// reach into has to come back byte-identical under EVERY edit
// variant. This is the gate M1a's new token kinds are measured
// against: a kind that is wrong is a licence for the rewriter to
// mutate, and a silent mutation inside a table column name or an
// external workbook reference corrupts a formula rather than failing.

/// The full edit matrix, applied with the most permissive scoping
/// available (`target_sheet = null` — "apply to bare refs everywhere
/// AND every sheet-qualified ref"). Identity under these is the
/// strongest form of the claim.
const all_edits = [_]RewriteEdit{
    .{ .insert_rows = .{ .at = 1, .count = 1 } },
    .{ .delete_rows = .{ .at = 1, .count = 1 } },
    .{ .insert_cols = .{ .at = 1, .count = 1 } },
    .{ .delete_cols = .{ .at = 1, .count = 1 } },
    .{ .rename_sheet = .{ .old = "Sheet1", .new = "Renamed" } },
    .{ .delete_sheet = "Sheet1" },
    // Out-of-scope table: proves structured refs survive the ONE
    // edit that can touch them whenever the table doesn't match,
    // and that no other construct is touched by it at all.
    .{ .rename_table_column = .{ .table = "NoSuchTable", .old = "Nope", .new = "Never" } },
};

fn expectIdentityUnderEveryEdit(input: []const u8) !void {
    for (all_edits) |edit| {
        const out = try rewriteFormula(testing.allocator, input, .{ .edit = edit });
        defer testing.allocator.free(out);
        if (!std.mem.eql(u8, input, out)) {
            std.debug.print(
                "edit {s} mutated an untouchable construct:\n  in:  '{s}'\n  out: '{s}'\n",
                .{ @tagName(edit), input, out },
            );
            return error.TestExpectedEqual;
        }
    }
}

test "compat: structured table refs survive every edit byte-identically" {
    // `Table1[A1]` is the correction the M1a ladder row names. Before
    // M1a the inner `A1` tokenized as a live `.cell_ref`, so
    // insert_rows rewrote a table COLUMN NAME to `A2` — a formula
    // corruption with no error and no diagnostic.
    for ([_][]const u8{
        "Table1[A1]",
        "Table1[Amount]",
        "Table1[#All]",
        "Table1[#Headers]",
        "Table1[@Amount]",
        "Table1[[#Data],[Amount]]",
        "Table1[[#This Row],[Unit Cost]]",
        "[@Amount]",
        "[[Col A]:[Col B]]",
        "Table1[Cost '[USD']]",
        "SUM(Table1[Amount])",
        "SUM(Table1[A1]:Table1[B2])",
        "Sales[[#Data],[Q'[1']]]",
    }) |c| try expectIdentityUnderEveryEdit(c);
}

test "compat: external references survive every edit byte-identically" {
    // Both spellings, and everything each one qualifies — the sheet
    // name and the cell reference belong to a workbook we never see.
    for ([_][]const u8{
        "'[Book.xlsx]Sheet1'!A1",
        "'[Book.xlsx]Sheet1'!A1:B2",
        "'[Book.xlsx]Sheet1'!$A$1",
        "[1]Sheet1!A1",
        "[1]Sheet1!A1:B2",
        "[1]!Total",
        "[12]'My Sheet'!A1",
        // External 3D spans: the whole chain lives in a workbook we
        // cannot see. Before the 3D leg, the walk stopped at the
        // range operator and the single-sheet matcher claimed
        // `Sheet3!A1` — a silent corruption this line now pins.
        "[1]Sheet1:Sheet3!A1",
        "[1]Sheet1:Sheet3!A1:B2",
        "[12]'My Sheet':Other!A1",
        "'[Book.xlsx]Sheet1:Sheet3'!A1",
        // External whole-axis spans: before the full-span leg, the
        // external walk stopped at the bang and the axis matcher
        // rewrote a span inside an unseen workbook.
        "[1]Sheet1!A:A",
        "[1]Sheet1!1:5",
        "'[Book.xlsx]Sheet1'!A:A",
        "'[Book.xlsx]Sheet1'!1:5",
        "[1]Sheet1:Sheet3!A:A",
    }) |c| try expectIdentityUnderEveryEdit(c);
}

test "compat: literals, names and operators survive every edit" {
    for ([_][]const u8{
        "1+2*3",
        "\"A1 inside a string\"",
        "\"a\"\"b\"",
        "TRUE",
        "FALSE()",
        "#N/A",
        "#REF!",
        "#DIV/0!",
        "#GETTING_DATA",
        "#BLOCKED!",
        "#PYTHON!",
        "{1,2;3,4}",
        "MyName.Sub",
        "\\Foo",
        "SUM(MyRange)",
        "  ",
        "",
        "50%",
        "1.5e+10",
        "R1C1",
        "R[-1]C[2]",
        "@SUM(MyRange)",
    }) |c| try expectIdentityUnderEveryEdit(c);
}

test "compat: unicode names survive every edit" {
    // Pre-M1a these shattered into one `.unknown` per byte. Round-trip
    // held, so the identity gate passed for the wrong reason; the
    // kinds are what makes it hold for the right one.
    for ([_][]const u8{
        "Ω",
        "ДАННЫЕ",
        "Größe+données",
        "SUM(日本語)",
        "\u{1D400}",
        "e\u{0301}_total",
        "\u{FF21}1",
    }) |c| try expectIdentityUnderEveryEdit(c);
}

test "compat: recognised refs still rewrite next to untouchable ones" {
    // The identity claim must not degrade into "the rewriter stopped
    // working". A live ref beside an opaque construct still shifts.
    try expectRewrite(
        "Table1[A1]+A1",
        .{ .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } } },
        "Table1[A1]+A2",
    );
    try expectRewrite(
        "'[Book.xlsx]Sheet1'!A1+A1",
        .{ .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } } },
        "'[Book.xlsx]Sheet1'!A1+A2",
    );
    try expectRewrite(
        "[1]Sheet1!A1+Sheet1!A1",
        .{ .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } } },
        "[1]Sheet1!A1+Sheet1!A2",
    );
    try expectRewrite(
        "SUM(Table1[Amount])+B5",
        .{ .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } } },
        "SUM(Table1[Amount])+B6",
    );
    // A sheet rename reaches the local sheet, never the external one.
    try expectRewrite(
        "'[Book.xlsx]Sheet1'!A1+Sheet1!A1",
        .{ .edit = .{ .rename_sheet = .{ .old = "Sheet1", .new = "Renamed" } } },
        "'[Book.xlsx]Sheet1'!A1+Renamed!A1",
    );
}

test "preserves whitespace and surrounding operators" {
    try expectRewrite(
        "  SUM ( A1 ,  B2 )  ",
        .{ .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } } },
        "  SUM ( A2 ,  B3 )  ",
    );
}

test "case-insensitive cell refs preserved as written" {
    // openpyxl emits lowercase. The parser accepts; on shift we
    // re-emit canonical uppercase. Documented trade-off: round-trip
    // for shifted refs normalizes case.
    try expectRewrite(
        "a1",
        .{ .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } } },
        "A2",
    );
    // Unmodified refs keep their original casing (token slice borrowed
    // from input).
    try expectRewrite(
        "a1",
        .{ .edit = .{ .insert_rows = .{ .at = 5, .count = 1 } } },
        "a1",
    );
}

test "invalid edit rejected" {
    // `at == 0` is meaningless in 1-based addressing.
    try testing.expectError(error.InvalidEdit, rewriteFormula(
        testing.allocator,
        "A1",
        .{ .edit = .{ .insert_rows = .{ .at = 0, .count = 1 } } },
    ));
    // Empty rename sides.
    try testing.expectError(error.InvalidEdit, rewriteFormula(
        testing.allocator,
        "A1",
        .{ .edit = .{ .rename_sheet = .{ .old = "", .new = "X" } } },
    ));
    try testing.expectError(error.InvalidEdit, rewriteFormula(
        testing.allocator,
        "A1",
        .{ .edit = .{ .rename_sheet = .{ .old = "X", .new = "" } } },
    ));
}

test "no-op edits leave formula byte-identical" {
    // count == 0: shift is a no-op; formula passes through verbatim.
    try expectRewrite(
        "SUM(A1:B5)",
        .{ .edit = .{ .insert_rows = .{ .at = 3, .count = 0 } } },
        "SUM(A1:B5)",
    );
    // Insert past every existing row → nothing shifts.
    try expectRewrite(
        "A1+B2",
        .{ .edit = .{ .insert_rows = .{ .at = 100, .count = 1 } } },
        "A1+B2",
    );
}

test "rename does not affect bare refs" {
    // rename_sheet has no row/col component; bare A1 passes through.
    try expectRewrite(
        "A1+B2",
        .{ .edit = .{ .rename_sheet = .{ .old = "Sheet1", .new = "X" } } },
        "A1+B2",
    );
}

test "checkAllAllocationFailures: rewrite is leak-safe under OOM" {
    const helpers = struct {
        fn runRename(allocator: std.mem.Allocator) !void {
            const out = try rewriteFormula(
                allocator,
                "'It''s'!$A$5:$B$10+Sheet2!C7",
                .{
                    .target_sheet = "It's",
                    .edit = .{ .rename_sheet = .{
                        .old = "It's",
                        .new = "New's Place",
                    } },
                },
            );
            allocator.free(out);
        }
        fn runShiftAndCollapse(allocator: std.mem.Allocator) !void {
            const out = try rewriteFormula(
                allocator,
                "SUM(A5:A6) + Sheet1!$B$10",
                .{ .edit = .{ .delete_rows = .{ .at = 5, .count = 2 } } },
            );
            allocator.free(out);
        }
        fn runThreeD(allocator: std.mem.Allocator) !void {
            // Named target exercises decodeSheetToken on a quoted AND
            // an unquoted endpoint, plus the #REF! collapse path.
            const out = try rewriteFormula(
                allocator,
                "SUM('Jan Sales':Mar!A5:B6)+Sheet1:Sheet3!C7",
                .{
                    .target_sheet = "Mar",
                    .edit = .{ .delete_rows = .{ .at = 5, .count = 2 } },
                },
            );
            allocator.free(out);
        }
        fn runThreeDEndpointEdits(allocator: std.mem.Allocator) !void {
            // Endpoint rename into a merged quoted pair (both span
            // spellings), then an order-driven contraction — the
            // allocation paths the sheet_order leg added.
            const renamed = try rewriteFormula(
                allocator,
                "SUM('Jan Sales':Mar!A5)+'Jan:Mar'!B2",
                .{ .edit = .{ .rename_sheet = .{ .old = "Mar", .new = "New's" } } },
            );
            allocator.free(renamed);
            const order = [_][]const u8{ "Jan", "Feb", "Mar" };
            const contracted = try rewriteFormula(
                allocator,
                "Jan:Mar!A1+'Jan:Mar'!A1",
                .{ .sheet_order = &order, .edit = .{ .delete_sheet = "Mar" } },
            );
            allocator.free(contracted);
        }
        fn runAxisSpans(allocator: std.mem.Allocator) !void {
            // Bare span shrink (one bound re-emitted) plus a quoted
            // qualified span slide (both bounds re-emitted).
            const out = try rewriteFormula(
                allocator,
                "SUM(A:C)+'My Sheet'!$D:$E",
                .{ .edit = .{ .delete_cols = .{ .at = 2, .count = 1 } } },
            );
            allocator.free(out);
        }
    };

    try testing.checkAllAllocationFailures(testing.allocator, helpers.runRename, .{});
    try testing.checkAllAllocationFailures(testing.allocator, helpers.runShiftAndCollapse, .{});
    try testing.checkAllAllocationFailures(testing.allocator, helpers.runThreeD, .{});
    try testing.checkAllAllocationFailures(testing.allocator, helpers.runThreeDEndpointEdits, .{});
    try testing.checkAllAllocationFailures(testing.allocator, helpers.runAxisSpans, .{});
}

test "many-column shift writes correct multi-letter columns" {
    // Z (26) + insert 1 at A → AA.
    try expectRewrite(
        "Z1",
        .{ .edit = .{ .insert_cols = .{ .at = 1, .count = 1 } } },
        "AA1",
    );
    // ZZ (702) + insert 1 at A → AAA.
    try expectRewrite(
        "ZZ1",
        .{ .edit = .{ .insert_cols = .{ .at = 1, .count = 1 } } },
        "AAA1",
    );
}

test "insert beyond grid leaves ref alone" {
    // XFD (col 16384) + insert_cols{1,1} would push to 16385 — off
    // the grid. Per current policy we leave the ref unchanged; a
    // future iteration may convert to #REF! to match Excel exactly.
    try expectRewrite(
        "XFD1",
        .{ .edit = .{ .insert_cols = .{ .at = 1, .count = 1 } } },
        "XFD1",
    );
}

// ─── rename_table_column ─────────────────────────────────────────

/// Shorthand ctx for the common qualified-ref cases.
fn renameColEdit(table: []const u8, old: []const u8, new: []const u8) RewriteContext {
    return .{ .edit = .{ .rename_table_column = .{ .table = table, .old = old, .new = new } } };
}

test "rename_table_column: qualified single-column spellings" {
    // Bare part inside the specifier.
    try expectRewrite(
        "SUM(Table1[Old])",
        renameColEdit("Table1", "Old", "New"),
        "SUM(Table1[New])",
    );
    // `@` shorthand keeps the `@`.
    try expectRewrite(
        "Table1[@Old]*2",
        renameColEdit("Table1", "Old", "New"),
        "Table1[@New]*2",
    );
    // Bracketed part keeps its brackets, swaps inner text only.
    try expectRewrite(
        "Table1[[#This Row],[Old]]",
        renameColEdit("Table1", "Old", "New"),
        "Table1[[#This Row],[New]]",
    );
    try expectRewrite(
        "SUM(Table1[[#Data],[Old]])",
        renameColEdit("Table1", "Old", "New"),
        "SUM(Table1[[#Data],[New]])",
    );
    // Bracketed single part (`[[Old Col]]` spelling).
    try expectRewrite(
        "Table1[[Old Col]]",
        renameColEdit("Table1", "Old Col", "X"),
        "Table1[[X]]",
    );
    // Two refs to the same table in one formula: both rewritten.
    try expectRewrite(
        "Table1[Old]+Table1[[#Totals],[Old]]",
        renameColEdit("Table1", "Old", "New"),
        "Table1[New]+Table1[[#Totals],[New]]",
    );
}

test "rename_table_column: column ranges rewrite the matching half" {
    try expectRewrite(
        "SUM(Table1[[Old]:[Other]])",
        renameColEdit("Table1", "Old", "New"),
        "SUM(Table1[[New]:[Other]])",
    );
    try expectRewrite(
        "SUM(Table1[[Other]:[Old]])",
        renameColEdit("Table1", "Old", "New"),
        "SUM(Table1[[Other]:[New]])",
    );
    // Both halves name the old column: both rewrite.
    try expectRewrite(
        "SUM(Table1[[Old]:[Old]])",
        renameColEdit("Table1", "Old", "New"),
        "SUM(Table1[[New]:[New]])",
    );
    // Items + range combined form.
    try expectRewrite(
        "Table1[[#Data],[Old]:[Other]]",
        renameColEdit("Table1", "Old", "New"),
        "Table1[[#Data],[New]:[Other]]",
    );
}

test "rename_table_column: decode-in on the OLD name's escapes" {
    // Column literally named `C]x` is spelled `C']x` in the
    // specifier; the edit names it decoded.
    try expectRewrite(
        "Table1[C']x]",
        renameColEdit("Table1", "C]x", "Plain"),
        "Table1[Plain]",
    );
    // Column literally named `#Foo` (escaped `#` — an UNescaped
    // leading `#` would be an item specifier).
    try expectRewrite(
        "Table1['#Foo]",
        renameColEdit("Table1", "#Foo", "New"),
        "Table1[New]",
    );
    // `@` mid-name.
    try expectRewrite(
        "Table1[Rate '@Peak]",
        renameColEdit("Table1", "Rate @Peak", "Flat"),
        "Table1[Flat]",
    );
}

test "rename_table_column: escape-out on the NEW name" {
    // Reserved bytes gain a `'` escape, and the punctuation makes
    // `columnNeedsBrackets` bracket the part — the printer's policy,
    // matching the spelling Excel itself writes for such names.
    try expectRewrite(
        "Table1[Old]",
        renameColEdit("Table1", "Old", "A]B"),
        "Table1[[A']B]]",
    );
    try expectRewrite(
        "Table1[Old]",
        renameColEdit("Table1", "Old", "Q[1]"),
        "Table1[[Q'[1']]]",
    );
    // A space forces the bracketed spelling on a bare part — the
    // printer's own policy, so canonical output agrees.
    try expectRewrite(
        "Table1[Old]",
        renameColEdit("Table1", "Old", "New Col"),
        "Table1[[New Col]]",
    );
    try expectRewrite(
        "Table1[@Old]",
        renameColEdit("Table1", "Old", "New Col"),
        "Table1[@[New Col]]",
    );
    // Already-bracketed parts never need the wrap.
    try expectRewrite(
        "Table1[[#Data],[Old]]",
        renameColEdit("Table1", "Old", "New Col"),
        "Table1[[#Data],[New Col]]",
    );
}

test "rename_table_column: fold-rule matching, byte-exact emit" {
    // Table and column names match under the symbol layer's fold —
    // the same rule resolution uses — while everything emitted is
    // the edit's exact spelling.
    try expectRewrite(
        "TABLE1[OLD]",
        renameColEdit("table1", "old", "New"),
        "TABLE1[New]",
    );
    // Unicode fold (é vs É).
    try expectRewrite(
        "Ventes[CAFÉ]",
        renameColEdit("ventes", "café", "thé"),
        "Ventes[thé]",
    );
}

test "rename_table_column: out-of-scope specifiers survive byte-identically" {
    const edit = renameColEdit("Table1", "Old", "New");
    // Another table's column.
    try expectRewrite("Other[Old]", edit, "Other[Old]");
    // Same table, different column; item specifiers are not columns.
    try expectRewrite("Table1[Older]", edit, "Table1[Older]");
    try expectRewrite("Table1[#All]", renameColEdit("Table1", "#All", "New"), "Table1[#All]");
    // Whitespace breaks the name↔specifier adjacency (the parser's
    // own pairing rule), so this is a bare specifier without an
    // owner — untouched.
    try expectRewrite("Table1 [Old]", edit, "Table1 [Old]");
    // Nothing after a `!` is a local table specifier.
    try expectRewrite("Sheet1!Table1[Old]", edit, "Sheet1!Table1[Old]");
    // Spaces are PART of the name in the specifier grammar (the
    // engine's columnIndex does not trim either): ` Old ` ≠ `Old`.
    try expectRewrite("Table1[ Old ]", edit, "Table1[ Old ]");
    // Malformed specifier: trailing separator never parses, so the
    // token stays opaque even though the table matches.
    try expectRewrite("Table1[Old,]", edit, "Table1[Old,]");
    // A formula with no structured refs at all round-trips.
    try expectRewrite("A1+B2*SUM(C:C)", edit, "A1+B2*SUM(C:C)");
}

test "rename_table_column: bare specifiers scope through owning_table" {
    const rename: RewriteEdit = .{ .rename_table_column = .{
        .table = "Table1",
        .old = "Old",
        .new = "New",
    } };
    // The table part's own formulas (calculatedColumnFormula /
    // totalsRowFormula) and producer cells inside the table's range
    // carry bare specifiers; `owning_table` names their table.
    try expectRewrite(
        "[@Old]*2",
        .{ .owning_table = "Table1", .edit = rename },
        "[@New]*2",
    );
    try expectRewrite(
        "SUM([Old])",
        .{ .owning_table = "Table1", .edit = rename },
        "SUM([New])",
    );
    try expectRewrite(
        "SUM([[Old]:[Other]])",
        .{ .owning_table = "Table1", .edit = rename },
        "SUM([[New]:[Other]])",
    );
    try expectRewrite(
        "[@[Old Col]]",
        .{ .owning_table = "Table1", .edit = .{ .rename_table_column = .{
            .table = "Table1",
            .old = "Old Col",
            .new = "New",
        } } },
        "[@[New]]",
    );
    // No owner (the default): the bare form has no table to match.
    try expectRewrite("[@Old]*2", .{ .edit = rename }, "[@Old]*2");
    // Another table's producer.
    try expectRewrite(
        "[@Old]*2",
        .{ .owning_table = "Other", .edit = rename },
        "[@Old]*2",
    );
    // owning_table scopes BARE forms only — a qualified ref inside
    // a producer formula still matches on its own table name.
    try expectRewrite(
        "Other[Old]+[Old]",
        .{ .owning_table = "Table1", .edit = rename },
        "Other[Old]+[New]",
    );
}

test "rename_table_column: refuses empty names" {
    const cases = [_]RewriteEdit{
        .{ .rename_table_column = .{ .table = "", .old = "Old", .new = "New" } },
        .{ .rename_table_column = .{ .table = "T", .old = "", .new = "New" } },
        .{ .rename_table_column = .{ .table = "T", .old = "Old", .new = "" } },
    };
    for (cases) |edit| {
        try testing.expectError(
            error.InvalidEdit,
            rewriteFormula(testing.allocator, "T[Old]", .{ .edit = edit }),
        );
    }
}

test "rename_table_column: OOM-safe at every allocation site" {
    const helpers = struct {
        fn run(allocator: std.mem.Allocator) !void {
            const out = try rewriteFormula(
                allocator,
                "SUM(Table1[[#Data],[Old]:[Other]])+[@Old]",
                .{
                    .owning_table = "Table1",
                    .edit = .{ .rename_table_column = .{
                        .table = "Table1",
                        .old = "Old",
                        .new = "New Col",
                    } },
                },
            );
            allocator.free(out);
        }
    };
    try testing.checkAllAllocationFailures(testing.allocator, helpers.run, .{});
}
