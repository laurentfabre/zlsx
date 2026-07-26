//! Formula rewriter (C1 milestone 2, iteration 1).
//!
//! Pure-function rewriter built on top of the C1 M1 tokenizer. Walks
//! a tokenized formula and applies a structural edit (insert / delete
//! rows or columns, or sheet rename) to every A1-style cell or range
//! reference it can recognise. Round-trip property of the underlying
//! tokenizer is preserved for everything we don't touch — whitespace,
//! literals, operators, names, and `.unknown` (external-workbook,
//! dynamic-array spill, etc.) bytes pass through verbatim.
//!
//! Scope (this iter):
//!   - Bare A1 cell refs:                   A1, $A$1, $A1, A$1
//!   - Bare A1 ranges:                       A1:B5, $A$1:$B$5
//!   - Sheet-qualified refs:                 Sheet1!A1, 'My Sheet'!A1:B5
//!   - Apostrophe-escaped quoted sheets:    'It''s'!A1
//!   - Insert/delete rows/cols, sheet rename
//!
//! Out of scope (deferred — match tokenizer M1 boundaries):
//!   - R1C1, structured table refs, 3D refs, dynamic-array `#`/`@`,
//!     external-workbook brackets. The tokenizer classifies these
//!     as `.unknown`; we leave them alone.
//!   - Full-column / full-row refs (`A:A`, `1:5`). The tokenizer
//!     reports these as `name op_range name` / `number op_range
//!     number`; this iter does not yet reshape them.

const std = @import("std");
const tokenizer = @import("tokenizer.zig");
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
    target_sheet: ?[]const u8 = null,
    edit: RewriteEdit,
};

pub const Error = error{
    OutOfMemory,
    /// `delete_rows` / `delete_cols` with `count == 0` is meaningless
    /// (no shift, nothing to delete). We refuse it rather than silently
    /// no-op, so callers don't construct an edit they didn't mean.
    InvalidEdit,
};

const MAX_ROWS: u32 = 1_048_576;
const MAX_COLS: u32 = 16_384;

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
    if (letters.len == 0) return null;
    if (letters.len > 3) return null;
    var v: u32 = 0;
    for (letters) |c| {
        const upper: u8 = if (c >= 'a' and c <= 'z') c - ('a' - 'A') else c;
        if (upper < 'A' or upper > 'Z') return null;
        v = v * 26 + @as(u32, upper - 'A' + 1);
    }
    return v;
}

/// Write column letters for `col` (1-based) into `buf`. Returns the
/// number of bytes written. `buf.len >= 3` is sufficient — XFD is
/// the largest valid column.
fn writeColLetters(buf: []u8, col: u32) usize {
    assert(col >= 1 and col <= MAX_COLS);
    assert(buf.len >= 3);
    var stack: [3]u8 = undefined;
    var n: u32 = col;
    var depth: usize = 0;
    while (n > 0) {
        n -= 1;
        stack[depth] = 'A' + @as(u8, @intCast(n % 26));
        depth += 1;
        n /= 26;
    }
    assert(depth >= 1 and depth <= 3);
    var i: usize = 0;
    while (i < depth) : (i += 1) {
        buf[i] = stack[depth - 1 - i];
    }
    return depth;
}

inline fn isAsciiAlpha(c: u8) bool {
    return (c >= 'A' and c <= 'Z') or (c >= 'a' and c <= 'z');
}

inline fn isDigit(c: u8) bool {
    return c >= '0' and c <= '9';
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
        // Detect sheet qualifier: (sheet_name | name) bang cell_ref [: cell_ref]
        const sq = matchSheetQualifier(work, i);
        if (sq) |info| {
            try rewriteSheetQualified(allocator, work, owned, ctx, info);
            i = info.end;
            continue;
        }

        // Bare cell ref or range starting at this token. Skip
        // refs that follow an `.unknown !` qualifier — those are
        // scoped to an opaque (e.g. external-workbook) sheet that
        // the rewriter cannot reason about. Mutating them would
        // corrupt formulas pointing at workbooks we don't see.
        if (work[i].kind == .cell_ref) {
            if (i >= 2 and work[i - 1].kind == .bang and work[i - 2].kind == .unknown) {
                i += 1;
                continue;
            }
            const range_end = if (isRangeAt(work, i)) i + 3 else i + 1;
            try rewriteBareRefOrRange(allocator, work, owned, ctx, i, range_end);
            i = range_end;
            continue;
        }

        i += 1;
    }
}

const SheetQualifierInfo = struct {
    sheet_idx: usize, // index of the sheet token (.sheet_name or .name)
    bang_idx: usize, // index of the `!`
    ref_start: usize, // index of the first cell_ref
    ref_end: usize, // exclusive — i.e. ref_start+1 (single) or ref_start+3 (range)
    end: usize, // index just past the whole pattern (== ref_end)
    is_range: bool,
};

fn matchSheetQualifier(work: []const Token, i: usize) ?SheetQualifierInfo {
    if (i + 2 >= work.len) return null; // need at least sheet, !, ref
    const sheet_kind = work[i].kind;
    if (sheet_kind != .sheet_name and sheet_kind != .name) return null;
    if (work[i + 1].kind != .bang) return null;
    if (work[i + 2].kind != .cell_ref) return null;
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
    const decoded: []u8 = blk: {
        const tok = work[info.sheet_idx];
        if (tok.kind == .sheet_name) {
            break :blk try decodeQuotedSheet(allocator, tok.text);
        }
        break :blk try allocator.dupe(u8, tok.text);
    };
    defer allocator.free(decoded);

    // Sheet rename.
    var current_sheet: []const u8 = decoded;
    var renamed_holder: ?[]u8 = null;
    defer if (renamed_holder) |h| allocator.free(h);

    if (ctx.edit == .rename_sheet) {
        const spec = ctx.edit.rename_sheet;
        if (std.mem.eql(u8, decoded, spec.old)) {
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
        if (std.mem.eql(u8, decoded, target)) {
            // Collapse the entire qualified-ref token sequence to a
            // single `#REF!` by replacing the leading token's text
            // with `#REF!` and zeroing the trailing tokens' text.
            // The printer concatenates `.text` slices verbatim, so
            // empty `.text` contributes nothing to the output.
            const ref_lex = try registerOwned(allocator, owned, try allocator.dupe(u8, "#REF!"));
            work[info.sheet_idx] = .{ .kind = .error_lit, .text = ref_lex };
            var k: usize = info.sheet_idx + 1;
            while (k < info.ref_end) : (k += 1) {
                work[k] = .{ .kind = .whitespace, .text = "" };
            }
            return;
        }
        // Different sheet — nothing to do (bare-ref / row-col path
        // is also a no-op for delete_sheet, see target_match below).
        return;
    }

    // Row/col edits apply only when the sheet matches the edit's
    // target_sheet. `target_sheet == null` means "apply everywhere."
    const target_match = blk: {
        switch (ctx.edit) {
            .rename_sheet, .delete_sheet => break :blk false, // already handled
            else => {},
        }
        if (ctx.target_sheet) |t| break :blk std.mem.eql(u8, current_sheet, t);
        break :blk true;
    };
    if (!target_match) return;

    try applyToRefRange(
        allocator,
        work,
        owned,
        ctx.edit,
        info.ref_start,
        info.is_range,
    );
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

    // Bare refs are scoped to `on_sheet`. Apply the edit only when
    // `target_sheet` matches `on_sheet` (or is null = "everywhere").
    const target_match = blk: {
        switch (ctx.edit) {
            .rename_sheet => break :blk false, // bare refs have no sheet to rename
            // delete_sheet only collapses qualified refs. Bare refs
            // are scoped to the formula's owning sheet, which is
            // necessarily a sheet that's NOT being deleted (the
            // deleted sheet's own formulas are dropped wholesale by
            // Workbook.deleteSheet, not rewritten).
            .delete_sheet => break :blk false,
            else => {},
        }
        if (ctx.target_sheet == null) break :blk true;
        if (ctx.on_sheet == null) break :blk true; // permissive default
        break :blk std.mem.eql(u8, ctx.on_sheet.?, ctx.target_sheet.?);
    };
    if (!target_match) return;

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

test "unknown tokens are not mutated" {
    // External-workbook ref tokenizes as `.unknown`; survives intact.
    try expectRewrite(
        "'[Book.xlsx]Sheet1'!A1+1",
        .{ .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } } },
        "'[Book.xlsx]Sheet1'!A1+1",
    );
    // Dynamic-array spill `A1#` — the `A1` is a cell_ref but the
    // trailing `#` is `.unknown`. Per the tokenizer's contract we
    // SHOULD NOT touch `.unknown`. The bare A1 still shifts.
    try expectRewrite(
        "A1#",
        .{ .edit = .{ .insert_rows = .{ .at = 1, .count = 1 } } },
        "A2#",
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
    };

    try testing.checkAllAllocationFailures(testing.allocator, helpers.runRename, .{});
    try testing.checkAllAllocationFailures(testing.allocator, helpers.runShiftAndCollapse, .{});
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
