//! Typed spreadsheet coordinates — the single owner of A1 parsing,
//! A1 formatting, and the column/row grid bounds (M0 of the tier-D1
//! formula ladder; see `goal_formula.md` §5.1).
//!
//! Why this module exists
//! ---------------------
//! Before M0 the tree carried SIX independent A1 parsers, seven
//! column-letter formatters, and five coordinate structs. They did not
//! agree — the full-reference parsers alone:
//!
//! | site | col base | case | leading-zero row |
//! |---|---|---|---|
//! | `src/xlsx.zig`            | 0-based | upper-only  | accepted |
//! | `pkg/workbook.zig`        | 1-based | insensitive | rejected |
//! | `pkg/sheet_plan.zig`      | 1-based | upper-only  | rejected |
//! | `pkg/embedding_part.zig`  | 0-based | insensitive | rejected |
//! | `pkg/vml_edit.zig`        | 1-based | upper-only  | accepted |
//! | `pkg/workbook_xml_plan.zig` (predicate) | — | insensitive | accepted |
//!
//! Two of those structs are both named `CellRef` and disagree on the
//! column base, so a value passed from one to the other is off by one
//! with no compile error. `pkg/sheet_plan.zig` disagreed with *itself*:
//! `parseA1Corner` returned a 1-based column while `formatCellRef` took
//! a 0-based one. `pkg/table_edit.zig` carried a formatter whose own
//! comment asked for a shared module "when a third consumer appears".
//!
//! The evaluator would have added a seventh dialect. So: one parser,
//! one formatter, and coordinates that carry their base in the type.
//! `refs/import_gate.zig` keeps it that way.
//!
//! Design
//! ------
//! `Col` and `Row` are distinct non-exhaustive enums over `u32`, not
//! bare integers. A `Col` cannot be passed where a `Row` is expected,
//! and neither converts implicitly to an integer — you must ask for
//! `.zeroBased()` or `.oneBased()`, which is where the old off-by-one
//! bugs lived. Canonical internal form is column 0-based, row 1-based,
//! but no caller depends on that: every crossing is an explicit
//! accessor.
//!
//! The parse policy differences above are preserved, not flattened —
//! `CellParseOptions` names them. M0 is a refactor with a
//! byte-identical gate, so each adapter keeps the exact semantics its
//! call sites already had. Unifying the *policies* is a later,
//! behaviour-changing decision; this module only makes them visible.
//!
//! Rooted at top-level `refs/` rather than `src/` or `pkg/` for the
//! same reason as `unicode/`: a file belongs to exactly one module's
//! package tree, and the consumers here span `zlsx` (`src/`),
//! `zlsx_pkg` (`pkg/`), and `zlsx_sheet_plan` (`pkg/sheet_plan.zig`).
//! Anywhere under those and the compile fails with "file exists in
//! modules 'zlsx' and 'zlsx_refs'".

const std = @import("std");
const assert = std.debug.assert;

/// Highest 1-based column Excel accepts (`XFD`).
pub const max_col_1based: u32 = 16_384;

/// Highest 1-based row Excel accepts.
pub const max_row: u32 = 1_048_576;

/// Longest column-letter run in a valid reference (`XFD`).
pub const max_col_letters: usize = 3;

/// Bytes sufficient to format any in-grid cell reference (`XFD1048576`
/// is 10; 16 leaves headroom and matches the legacy buffer contracts).
pub const format_buf_len: usize = 16;

/// Every failure this module can produce. Callers map it onto their own
/// error set at the adapter boundary — that mapping is what keeps M0
/// from changing any public error type.
pub const Error = error{InvalidRef};

/// A column. Distinct type from `Row` so the two cannot be swapped.
/// Canonically 0-based inside; cross with an explicit accessor.
pub const Col = enum(u32) {
    _,

    pub fn fromZeroBased(v: u32) Error!Col {
        if (v >= max_col_1based) return error.InvalidRef;
        return @enumFromInt(v);
    }

    pub fn fromOneBased(v: u32) Error!Col {
        if (v == 0 or v > max_col_1based) return error.InvalidRef;
        return @enumFromInt(v - 1);
    }

    pub fn zeroBased(self: Col) u32 {
        return @intFromEnum(self);
    }

    pub fn oneBased(self: Col) u32 {
        return @intFromEnum(self) + 1;
    }

    pub fn order(a: Col, b: Col) std.math.Order {
        return std.math.order(@intFromEnum(a), @intFromEnum(b));
    }
};

/// A row. Canonically 1-based inside — Excel rows have no zero.
pub const Row = enum(u32) {
    _,

    pub fn fromOneBased(v: u32) Error!Row {
        if (v == 0 or v > max_row) return error.InvalidRef;
        return @enumFromInt(v);
    }

    pub fn oneBased(self: Row) u32 {
        return @intFromEnum(self);
    }

    /// 0-based row index, for array addressing. Named so the crossing
    /// is never silent.
    pub fn zeroBased(self: Row) u32 {
        return @intFromEnum(self) - 1;
    }

    pub fn order(a: Row, b: Row) std.math.Order {
        return std.math.order(@intFromEnum(a), @intFromEnum(b));
    }
};

/// Which halves of a reference are `$`-anchored. Both false is the
/// plain `A1` form, which is what every pre-M0 call site produces and
/// consumes — hence the defaults.
pub const Anchor = struct {
    col: bool = false,
    row: bool = false,

    pub const relative: Anchor = .{};
    pub const absolute: Anchor = .{ .col = true, .row = true };
};

/// One cell coordinate.
pub const Cell = struct {
    col: Col,
    row: Row,
    anchor: Anchor = .{},

    /// Positional equality — ignores anchoring, because `$A$1` and `A1`
    /// address the same cell. Use `eqlExact` when the `$` matters.
    pub fn eql(a: Cell, b: Cell) bool {
        return a.col == b.col and a.row == b.row;
    }

    pub fn eqlExact(a: Cell, b: Cell) bool {
        return a.eql(b) and
            a.anchor.col == b.anchor.col and
            a.anchor.row == b.anchor.row;
    }
};

/// A rectangular span. `first` is the top-left, `last` the
/// bottom-right, component-wise — `normalized` enforces it.
pub const Range = struct {
    first: Cell,
    last: Cell,

    /// Reorder corners so `first <= last` on both axes. Source XML is
    /// allowed to list them either way round.
    pub fn normalized(self: Range) Range {
        // Normalisation is per-axis (Excel treats "B1:A2" as "A1:B2"),
        // so each axis carries its own `$` with it rather than the
        // corner it happened to be written in.
        const col_swap = self.first.col.zeroBased() > self.last.col.zeroBased();
        const row_swap = self.first.row.oneBased() > self.last.row.oneBased();
        const lo_col = if (col_swap) self.last else self.first;
        const hi_col = if (col_swap) self.first else self.last;
        const lo_row = if (row_swap) self.last else self.first;
        const hi_row = if (row_swap) self.first else self.last;
        return .{
            .first = .{
                .col = lo_col.col,
                .row = lo_row.row,
                .anchor = .{ .col = lo_col.anchor.col, .row = lo_row.anchor.row },
            },
            .last = .{
                .col = hi_col.col,
                .row = hi_row.row,
                .anchor = .{ .col = hi_col.anchor.col, .row = hi_row.anchor.row },
            },
        };
    }

    pub fn isNormalized(self: Range) bool {
        return self.first.col.zeroBased() <= self.last.col.zeroBased() and
            self.first.row.oneBased() <= self.last.row.oneBased();
    }

    pub fn rowCount(self: Range) u32 {
        assert(self.isNormalized());
        return self.last.row.oneBased() - self.first.row.oneBased() + 1;
    }

    pub fn colCount(self: Range) u32 {
        assert(self.isNormalized());
        return self.last.col.zeroBased() - self.first.col.zeroBased() + 1;
    }

    pub fn contains(self: Range, c: Cell) bool {
        assert(self.isNormalized());
        return c.col.zeroBased() >= self.first.col.zeroBased() and
            c.col.zeroBased() <= self.last.col.zeroBased() and
            c.row.oneBased() >= self.first.row.oneBased() and
            c.row.oneBased() <= self.last.row.oneBased();
    }

    pub fn overlaps(self: Range, other: Range) bool {
        assert(self.isNormalized() and other.isNormalized());
        return self.first.row.oneBased() <= other.last.row.oneBased() and
            other.first.row.oneBased() <= self.last.row.oneBased() and
            self.first.col.zeroBased() <= other.last.col.zeroBased() and
            other.first.col.zeroBased() <= self.last.col.zeroBased();
    }
};

/// Whether lowercase column letters are accepted. The reader's XML path
/// is upper-only (Excel never emits lowercase in `r=` attributes);
/// user-facing entry points are permissive.
pub const Case = enum { upper_only, insensitive };

/// Whether `A01` parses as row 1 or is rejected. `src/xlsx.zig` has
/// always accepted it; every other site rejects it.
pub const LeadingZeroRow = enum { reject, accept };

/// Whether `$` anchors are accepted. Every pre-M0 call site parses
/// Excel `r=` attributes or merge refs, which never carry `$` — so the
/// default rejects them and those adapters keep their exact behaviour.
/// The formula layer (M2) is what turns this on.
pub const Dollar = enum { reject, accept };

pub const CellParseOptions = struct {
    case: Case = .insensitive,
    leading_zero_row: LeadingZeroRow = .reject,
    dollar: Dollar = .reject,
};

/// Parse an A1 cell reference ("B12", "AA10").
///
/// Rejects: empty input, no letters, no digits, trailing bytes after
/// the digits, row 0, and anything outside the `XFD` / 1048576 grid.
/// Column overflow is caught before it can wrap.
pub fn parseCell(s: []const u8, opts: CellParseOptions) Error!Cell {
    var rest = s;
    var anchor: Anchor = .{};
    if (opts.dollar == .accept and rest.len > 0 and rest[0] == '$') {
        anchor.col = true;
        rest = rest[1..];
    }

    const split = try scanColLetters(rest, opts.case, max_col_1based);
    const col_1based = split.value;
    var i = split.end;

    if (i < rest.len and rest[i] == '$') {
        if (opts.dollar == .reject) return error.InvalidRef;
        anchor.row = true;
        i += 1;
    }

    if (i == rest.len) return error.InvalidRef; // letters but no digits
    if (opts.leading_zero_row == .reject and rest[i] == '0') return error.InvalidRef;

    var row: u32 = 0;
    while (i < rest.len) : (i += 1) {
        const c = rest[i];
        if (c < '0' or c > '9') return error.InvalidRef; // trailing garbage
        row = std.math.mul(u32, row, 10) catch return error.InvalidRef;
        row = std.math.add(u32, row, c - '0') catch return error.InvalidRef;
        if (row > max_row) return error.InvalidRef;
    }
    if (row == 0) return error.InvalidRef;

    return .{
        .col = try Col.fromOneBased(col_1based),
        .row = try Row.fromOneBased(row),
        .anchor = anchor,
    };
}

/// Parse an A1 range ("A1:B2"). A bare cell promotes to a 1x1
/// rectangle. Corners are returned in the order given — call
/// `normalized()` if you need top-left/bottom-right ordering.
pub fn parseRange(s: []const u8, opts: CellParseOptions) Error!Range {
    const colon = std.mem.indexOfScalar(u8, s, ':') orelse {
        const only = try parseCell(s, opts);
        return .{ .first = only, .last = only };
    };
    return .{
        .first = try parseCell(s[0..colon], opts),
        .last = try parseCell(s[colon + 1 ..], opts),
    };
}

pub const ColParseOptions = struct {
    case: Case = .insensitive,
    /// Cap on letter-run length. `null` means "whatever fits the grid".
    /// `src/formula/rewriter.zig` caps at 3 before any grid check.
    max_letters: ?usize = null,
    /// Whether to enforce the `XFD` ceiling. The formula rewriter
    /// deliberately does not — it accepts an out-of-grid column and
    /// lets a later stage reject it.
    bounds: enum { grid, unchecked } = .grid,
};

/// Result of scanning a leading column-letter run.
pub const ColPrefix = struct {
    /// 1-based column number.
    col_1based: u32,
    /// Byte length of the letter run — where the caller resumes.
    letters_len: usize,
};

/// Scan the column-letter PREFIX of `s`, ignoring whatever follows.
/// For callers that only want the column out of a longer reference and
/// deliberately do not validate the tail.
pub fn scanColPrefix(s: []const u8, opts: ColParseOptions) Error!ColPrefix {
    if (opts.max_letters) |cap| {
        if (s.len > cap) return error.InvalidRef;
    }
    const ceiling: ?u32 = switch (opts.bounds) {
        .grid => max_col_1based,
        .unchecked => null,
    };
    const split = try scanColLetters(s, opts.case, ceiling);
    return .{ .col_1based = split.value, .letters_len = split.end };
}

/// Parse a bare column-letter run ("A", "XFD") to a **1-based** column
/// number. Rejects trailing bytes. Returns the raw number rather than a
/// `Col` so that `bounds = .unchecked` callers can still be served.
pub fn parseColNumber(s: []const u8, opts: ColParseOptions) Error!u32 {
    const prefix = try scanColPrefix(s, opts);
    if (prefix.letters_len != s.len) return error.InvalidRef; // trailing bytes
    return prefix.col_1based;
}

/// Grid-checked variant of `parseColNumber`.
pub fn parseCol(s: []const u8, opts: ColParseOptions) Error!Col {
    return Col.fromOneBased(try parseColNumber(s, opts));
}

const ColScan = struct {
    /// 1-based column number accumulated from the letter run.
    value: u32,
    /// Index of the first byte after the letters.
    end: usize,
};

/// Shared letter-run scanner. Stops at the first non-letter.
///
/// `ceiling` bounds the accumulator as it grows, so a long run errors
/// early instead of wrapping. Pass `null` only where a caller genuinely
/// wants out-of-grid columns (the formula rewriter); trapping
/// arithmetic still catches `u32` overflow there, so a pathological
/// all-letters input errors rather than silently aliasing a valid
/// column.
fn scanColLetters(s: []const u8, case: Case, ceiling: ?u32) Error!ColScan {
    var i: usize = 0;
    var value: u32 = 0;
    while (i < s.len) : (i += 1) {
        const c = s[i];
        const upper: u8 = switch (case) {
            .upper_only => c,
            .insensitive => if (c >= 'a' and c <= 'z') c - ('a' - 'A') else c,
        };
        if (upper < 'A' or upper > 'Z') break;
        value = std.math.mul(u32, value, 26) catch return error.InvalidRef;
        value = std.math.add(u32, value, upper - 'A' + 1) catch return error.InvalidRef;
        if (ceiling) |limit| {
            if (value > limit) return error.InvalidRef;
        }
    }
    if (i == 0) return error.InvalidRef; // no letters at all
    return .{ .value = value, .end = i };
}

/// Longest bijective base-26 run a `u32` can produce (26^7 exceeds
/// `maxInt(u32)`), for the unchecked formatting path.
const max_u32_col_letters: usize = 7;

pub const WriteError = error{ColumnIndexOutOfRange};

/// Write the letters for a **1-based** column number into `buf`,
/// returning the byte count. Accepts out-of-grid numbers — the only
/// failure is not fitting in `buf`.
///
/// This exists because two pre-M0 call sites (`pkg/sheet_edit.zig`'s
/// `formatColLetters`, reachable with a shifted index past `XFD`) format
/// without a grid check, and M0 is not allowed to change what they do.
/// Prefer `writeColLetters` — the typed, in-grid path — for new code.
pub fn writeColNumberLetters(buf: []u8, col_1based: u32) WriteError!usize {
    if (col_1based == 0) return error.ColumnIndexOutOfRange;
    var stack: [max_u32_col_letters]u8 = undefined;
    var n: u32 = col_1based;
    var depth: usize = 0;
    while (n > 0) {
        n -= 1;
        stack[depth] = 'A' + @as(u8, @intCast(n % 26));
        depth += 1;
        n /= 26;
    }
    if (depth > buf.len) return error.ColumnIndexOutOfRange;
    var i: usize = 0;
    while (i < depth) : (i += 1) buf[i] = stack[depth - 1 - i];
    return depth;
}

/// Write `col`'s letters into `buf`, returning the byte count.
/// `buf.len >= max_col_letters` is sufficient for any in-grid column.
pub fn writeColLetters(buf: []u8, col: Col) usize {
    assert(buf.len >= max_col_letters);
    // In-grid by construction, and the buffer is large enough, so
    // neither failure mode is reachable.
    return writeColNumberLetters(buf, col.oneBased()) catch unreachable;
}

/// Format `cell` as A1 into `buf`, returning the written slice.
/// Emits `$` for whichever halves `cell.anchor` marks — which is
/// neither, by default, so the pre-M0 call sites are byte-unchanged.
/// `buf.len >= format_buf_len` is sufficient for any in-grid cell,
/// anchored or not (`$XFD$1048576` is 12 bytes).
pub fn formatCell(buf: []u8, cell: Cell) []u8 {
    assert(buf.len >= format_buf_len);
    var n: usize = 0;
    if (cell.anchor.col) {
        buf[n] = '$';
        n += 1;
    }
    n += writeColLetters(buf[n..], cell.col);
    if (cell.anchor.row) {
        buf[n] = '$';
        n += 1;
    }
    n += std.fmt.printInt(buf[n..], cell.row.oneBased(), 10, .lower, .{});
    return buf[0..n];
}

// ─── Tests ───────────────────────────────────────────────────────────

const testing = std.testing;

test "Col/Row bases cross only through explicit accessors" {
    const a = try Col.fromZeroBased(0);
    try testing.expectEqual(@as(u32, 0), a.zeroBased());
    try testing.expectEqual(@as(u32, 1), a.oneBased());

    const xfd = try Col.fromOneBased(max_col_1based);
    try testing.expectEqual(@as(u32, 16_383), xfd.zeroBased());
    try testing.expectEqual(max_col_1based, xfd.oneBased());

    const r = try Row.fromOneBased(1);
    try testing.expectEqual(@as(u32, 1), r.oneBased());
    try testing.expectEqual(@as(u32, 0), r.zeroBased());
}

test "Col/Row constructors reject out-of-grid values" {
    try testing.expectError(error.InvalidRef, Col.fromZeroBased(max_col_1based));
    try testing.expectError(error.InvalidRef, Col.fromOneBased(0));
    try testing.expectError(error.InvalidRef, Col.fromOneBased(max_col_1based + 1));
    try testing.expectError(error.InvalidRef, Row.fromOneBased(0));
    try testing.expectError(error.InvalidRef, Row.fromOneBased(max_row + 1));
}

test "parseCell round-trips the grid corners" {
    const cases = [_]struct { s: []const u8, col0: u32, row: u32 }{
        .{ .s = "A1", .col0 = 0, .row = 1 },
        .{ .s = "B12", .col0 = 1, .row = 12 },
        .{ .s = "Z1", .col0 = 25, .row = 1 },
        .{ .s = "AA1", .col0 = 26, .row = 1 },
        .{ .s = "XFD1", .col0 = 16_383, .row = 1 },
        .{ .s = "A1048576", .col0 = 0, .row = max_row },
        .{ .s = "XFD1048576", .col0 = 16_383, .row = max_row },
    };
    for (cases) |c| {
        const cell = try parseCell(c.s, .{});
        try testing.expectEqual(c.col0, cell.col.zeroBased());
        try testing.expectEqual(c.row, cell.row.oneBased());

        var buf: [format_buf_len]u8 = undefined;
        try testing.expectEqualStrings(c.s, formatCell(&buf, cell));
    }
}

test "parseCell rejects off-grid and malformed input" {
    const bad = [_][]const u8{
        "",      "A",    "1",          "AB",  "A0",
        "XFE1",  "ZZZ1", "A1048577",   "A1 ", " A1",
        "A1:B2", "A-1",  "AAAAAAAAA1", "A1x", "1A",
    };
    for (bad) |s| {
        try testing.expectError(error.InvalidRef, parseCell(s, .{}));
    }
}

test "parse policy: case sensitivity is explicit" {
    try testing.expectError(error.InvalidRef, parseCell("a1", .{ .case = .upper_only }));
    const cell = try parseCell("a1", .{ .case = .insensitive });
    try testing.expectEqual(@as(u32, 0), cell.col.zeroBased());
    try testing.expectEqual(@as(u32, 1), cell.row.oneBased());
}

test "parse policy: leading-zero row is explicit" {
    try testing.expectError(error.InvalidRef, parseCell("A01", .{ .leading_zero_row = .reject }));
    const cell = try parseCell("A01", .{ .leading_zero_row = .accept });
    try testing.expectEqual(@as(u32, 1), cell.row.oneBased());
    // Still no row zero, whatever the policy.
    try testing.expectError(error.InvalidRef, parseCell("A0", .{ .leading_zero_row = .accept }));
    try testing.expectError(error.InvalidRef, parseCell("A00", .{ .leading_zero_row = .accept }));
}

test "parseRange promotes a bare cell and preserves corner order" {
    const r = try parseRange("A1:B2", .{});
    try testing.expectEqual(@as(u32, 0), r.first.col.zeroBased());
    try testing.expectEqual(@as(u32, 1), r.last.col.zeroBased());

    const single = try parseRange("C3", .{});
    try testing.expect(single.first.eql(single.last));
    try testing.expectEqual(@as(u32, 1), single.rowCount());
    try testing.expectEqual(@as(u32, 1), single.colCount());

    const reversed = try parseRange("B2:A1", .{});
    try testing.expect(!reversed.isNormalized());
    const fixed = reversed.normalized();
    try testing.expect(fixed.isNormalized());
    try testing.expectEqual(@as(u32, 0), fixed.first.col.zeroBased());
    try testing.expectEqual(@as(u32, 2), fixed.rowCount());
}

test "Range geometry" {
    const r = (try parseRange("B2:D5", .{})).normalized();
    try testing.expectEqual(@as(u32, 4), r.rowCount());
    try testing.expectEqual(@as(u32, 3), r.colCount());
    try testing.expect(r.contains(try parseCell("C3", .{})));
    try testing.expect(!r.contains(try parseCell("A1", .{})));
    try testing.expect(r.overlaps((try parseRange("D5:E9", .{})).normalized()));
    try testing.expect(!r.overlaps((try parseRange("E1:F2", .{})).normalized()));
}

test "parseColNumber honours the letter cap and the bounds policy" {
    try testing.expectEqual(@as(u32, 1), try parseColNumber("A", .{}));
    try testing.expectEqual(@as(u32, 16_384), try parseColNumber("XFD", .{}));
    try testing.expectError(error.InvalidRef, parseColNumber("XFE", .{}));
    try testing.expectError(error.InvalidRef, parseColNumber("A1", .{}));
    try testing.expectError(error.InvalidRef, parseColNumber("", .{}));

    // The rewriter's contract: cap at 3 letters, no grid ceiling.
    try testing.expectEqual(
        @as(u32, 18_278),
        try parseColNumber("ZZZ", .{ .max_letters = 3, .bounds = .unchecked }),
    );
    try testing.expectError(
        error.InvalidRef,
        parseColNumber("AAAA", .{ .max_letters = 3, .bounds = .unchecked }),
    );
}

test "absolute/relative anchors round-trip, and are opt-in" {
    // Off by default — every pre-M0 call site must keep rejecting `$`.
    for ([_][]const u8{ "$A1", "A$1", "$A$1" }) |s| {
        try testing.expectError(error.InvalidRef, parseCell(s, .{}));
    }

    const cases = [_]struct { s: []const u8, col: bool, row: bool }{
        .{ .s = "A1", .col = false, .row = false },
        .{ .s = "$A1", .col = true, .row = false },
        .{ .s = "A$1", .col = false, .row = true },
        .{ .s = "$A$1", .col = true, .row = true },
        .{ .s = "$XFD$1048576", .col = true, .row = true },
    };
    for (cases) |c| {
        const cell = try parseCell(c.s, .{ .dollar = .accept });
        try testing.expectEqual(c.col, cell.anchor.col);
        try testing.expectEqual(c.row, cell.anchor.row);

        var buf: [format_buf_len]u8 = undefined;
        try testing.expectEqualStrings(c.s, formatCell(&buf, cell));
    }

    // Anchoring does not change which cell is addressed.
    const abs = try parseCell("$B$2", .{ .dollar = .accept });
    const rel = try parseCell("B2", .{});
    try testing.expect(abs.eql(rel));
    try testing.expect(!abs.eqlExact(rel));

    // A lone `$`, or one in the wrong place, is still malformed.
    for ([_][]const u8{ "$", "$$A1", "A$$1", "A1$", "$A" }) |s| {
        try testing.expectError(error.InvalidRef, parseCell(s, .{ .dollar = .accept }));
    }
}

test "normalizing a range carries each axis's anchor with it" {
    // Written backwards on both axes, with the `$`s on the corner that
    // ends up second.
    const r = try parseRange("$B$2:A1", .{ .dollar = .accept });
    const n = r.normalized();
    try testing.expect(n.isNormalized());

    var buf: [format_buf_len]u8 = undefined;
    try testing.expectEqualStrings("A1", formatCell(&buf, n.first));
    try testing.expectEqualStrings("$B$2", formatCell(&buf, n.last));

    // Mixed axes: column reversed, row already ordered.
    const m = (try parseRange("$B1:A$2", .{ .dollar = .accept })).normalized();
    try testing.expectEqualStrings("A1", formatCell(&buf, m.first));
    try testing.expectEqualStrings("$B$2", formatCell(&buf, m.last));
}

test "scanColPrefix stops at the first non-letter and reports the tail" {
    const p = try scanColPrefix("AA10", .{ .case = .upper_only });
    try testing.expectEqual(@as(u32, 27), p.col_1based);
    try testing.expectEqual(@as(usize, 2), p.letters_len);

    // A bare column is a full-length prefix.
    const bare = try scanColPrefix("C", .{ .case = .upper_only });
    try testing.expectEqual(@as(u32, 3), bare.col_1based);
    try testing.expectEqual(@as(usize, 1), bare.letters_len);

    // Still needs at least one letter, and still honours the ceiling.
    try testing.expectError(error.InvalidRef, scanColPrefix("1A", .{}));
    try testing.expectError(error.InvalidRef, scanColPrefix("XFE1", .{}));
}

test "writeColNumberLetters serves out-of-grid columns and bounds by buffer" {
    var buf: [max_u32_col_letters]u8 = undefined;
    // In-grid agrees with the typed path.
    const n = try writeColNumberLetters(&buf, 16_384);
    try testing.expectEqualStrings("XFD", buf[0..n]);
    // Past XFD still formats — the legacy behaviour M0 must preserve.
    const past = try writeColNumberLetters(&buf, 16_385);
    try testing.expectEqualStrings("XFE", buf[0..past]);
    // Zero has no bijective representation.
    try testing.expectError(error.ColumnIndexOutOfRange, writeColNumberLetters(&buf, 0));
    // Too small a buffer is the only other failure.
    var tiny: [1]u8 = undefined;
    try testing.expectError(error.ColumnIndexOutOfRange, writeColNumberLetters(&tiny, 27));
    try testing.expectEqual(@as(usize, 1), try writeColNumberLetters(&tiny, 26));
}

test "writeColLetters covers every in-grid column" {
    var buf: [max_col_letters]u8 = undefined;
    var i: u32 = 0;
    while (i < max_col_1based) : (i += 1) {
        const col = try Col.fromZeroBased(i);
        const n = writeColLetters(&buf, col);
        const back = try parseColNumber(buf[0..n], .{});
        try testing.expectEqual(col.oneBased(), back);
    }
}
