//! Formula value model — types, fidelity rules, coercion, collation.
//!
//! M3a1 of the tier-D1 ladder (`goal_formula.md` §5.3, §5.4, §9). This
//! module is the vocabulary every later layer speaks: the evaluator
//! (M3a2), the registry (M3a2), criteria (M3b), and every public
//! boundary (CLI, C ABI, Python) publish through the one conversion
//! defined here.
//!
//! It deliberately contains **no evaluation**. There is no expression
//! walk, no function dispatch, and no environment — those are M3a2's.
//! What lives here is everything an evaluator must agree with before it
//! can be written:
//!
//!   * `ScalarValue` and `Matrix`, with their invariants enforced at
//!     construction rather than asserted downstream;
//!   * `PublishedScalar` / `PublishedMatrix` and the single mandatory
//!     blank→0 conversion (§5.3a);
//!   * `excel_fp_rules_v1` and `ieee_fp_rules_v1` as **two rule
//!     tables**, so a divergence is a data difference rather than a
//!     branch someone forgot to write (§5.4);
//!   * `parseDecimal(fidelity, ingress, text)` — the ONLY
//!     decimal-text→f64 path in the engine (§5.4);
//!   * the §5.3b shape table and the 7×8 scalar coercion matrix, both
//!     as executable tables rather than prose;
//!   * the §5.3c error-propagation order;
//!   * `collation_v1` — one comparator over full-case-folded code-point
//!     sequences (§5.4b).
//!
//! Fidelity is a mode, not a quality
//! ---------------------------------
//! `.ieee` is a contract, not the absence of one. Both modes share
//! Excel's *value domain* — overflow is `#NUM!` in both, `0/0` is
//! `#DIV/0!` in both — and differ only where §5.4 says they differ.
//! `divergence_points` names every such place, and the "Divergence ×2"
//! gate runs each one under both tables asserting agreement *and*
//! disagreement where each is required.
//!
//! Collation and the module graph
//! ------------------------------
//! `collation_v1` is stated once here as an algorithm over folded code
//! points, and takes the fold as an injected `FoldFn`. That is not
//! indirection for its own sake: the shipped fold lives at
//! `src/unicode/casefold.zig`, which is inside the `zlsx` module's
//! package tree (`src/xlsx.zig:25` imports it relatively), so a named
//! module rooted on the same file would collide the moment `zlsx`
//! imports the formula engine — the failure M0 hit with `refs/` and
//! M1a hit with `unicode/xid.zig`. Injection keeps the *semantics*
//! independent of the build graph; the test section below wires the
//! real fold so the fixtures run against the shipped algorithm.

const std = @import("std");
const assert = std.debug.assert;

// ─── fidelity rule tables (§5.4) ─────────────────────────────────

pub const Fidelity = enum { excel, ieee };

/// Where a decimal string entered the engine. `parseDecimal` is the
/// only decimal-text→f64 path, and the ingress selects which §5.4 rule
/// applies — a second parser diverging is structurally impossible
/// because there is no second parser.
///
/// **Caller→ingress table (normative, §5.4)**:
///
///   * formula literals                    → `.literal`      (N1a)
///   * present numeric `<v>` values only   → `.cache_import` (N1b)
///   * §5.3b arithmetic coercion           → `.text_coercion`
///   * VALUE/NUMBERVALUE/DATEVALUE/TIMEVALUE components → `.function_arg`
///   * criteria operands                   → `.criteria`
///
/// SST indices parse as bounded integers and numeric-looking SST
/// *content* stays text until §5.3b coercion, so there is deliberately
/// no SST-number ingress.
pub const Ingress = enum {
    literal,
    cache_import,
    text_coercion,
    function_arg,
    criteria,

    /// Whether surrounding ASCII spaces are tolerated. Stored forms are
    /// exact; coercion paths see user text, and Excel trims it.
    pub fn trimsSpace(self: Ingress) bool {
        return switch (self) {
            .literal, .cache_import => false,
            .text_coercion, .function_arg, .criteria => true,
        };
    }
};

/// A versioned floating-point rule table. Two instances exist and the
/// engine never branches on `Fidelity` directly — it reads the table,
/// so adding a rule means adding a field, not hunting for `if (excel)`.
pub const FpRules = struct {
    name: []const u8,

    /// N1a — significant decimal digits retained on decimal ingress.
    /// `null` means full binary64-nearest conversion of the whole text.
    /// Applies to every ingress **except** `.cache_import`, which is
    /// N1b: a cached `<v>` is already a binary64 value that some
    /// producer serialized, and re-rounding it would corrupt a value
    /// nobody asked us to reinterpret.
    literal_significant_digits: ?u8,

    /// N2 — snap a near-zero additive result to +0. **Additive scope
    /// only**: never multiplication, division, or function results.
    zero_snap: bool,

    /// N2's threshold, as a binary exponent: `a ± b` snaps when the
    /// result is non-zero and `|r| < 2^shift · max(|a|,|b|)`.
    zero_snap_relative_shift: i32,

    /// N3 — whether a negative zero survives to publication.
    preserve_signed_zero: bool,

    /// N3 — whether subnormal results survive. Both tables preserve
    /// them; the committed manifests agree bit-for-bit on `2^-1074`
    /// and `1E-308/10000000000`, so this is oracle-decided, not
    /// assumed.
    preserve_subnormals: bool,

    /// N4a — a non-finite result is not a value. Shared by both modes:
    /// Excel's value domain is the domain, and `.ieee` changes how
    /// arithmetic rounds, not what a workbook may hold.
    finiteness_is_error: bool,

    pub fn of(f: Fidelity) FpRules {
        return switch (f) {
            .excel => excel_fp_rules_v1,
            .ieee => ieee_fp_rules_v1,
        };
    }
};

/// N2's threshold is **spec-pinned and provisional**: no committed
/// manifest contains a zero-snap case, and the Excel oracle leg is
/// parked, so nothing on disk can decide it. `2^-48` (≈3.6e-15
/// relative) snaps the textbook cases — `1.1-1.0-0.1` and
/// `0.1+0.2-0.3` — and is recorded here as a named constant precisely
/// so pinning it later is a one-line change with a fixture behind it.
pub const zero_snap_relative_shift_v1: i32 = -48;

pub const excel_fp_rules_v1: FpRules = .{
    .name = "excel_fp_rules_v1",
    .literal_significant_digits = 15,
    .zero_snap = true,
    .zero_snap_relative_shift = zero_snap_relative_shift_v1,
    .preserve_signed_zero = false,
    .preserve_subnormals = true,
    .finiteness_is_error = true,
};

pub const ieee_fp_rules_v1: FpRules = .{
    .name = "ieee_fp_rules_v1",
    .literal_significant_digits = null,
    .zero_snap = false,
    .zero_snap_relative_shift = 0,
    .preserve_signed_zero = true,
    .preserve_subnormals = true,
    .finiteness_is_error = true,
};

// ─── errors (§5.3a, §10 plane 1) ─────────────────────────────────

/// The ten spellings Excel documents. These are plane-1 *values* —
/// successful results — not the plane-2 refusals in `parser.zig`.
pub const KnownError = enum {
    null_err,
    div0,
    value,
    ref,
    name,
    num,
    na,
    getting_data,
    spill,
    calc,

    pub fn spelling(self: KnownError) []const u8 {
        return switch (self) {
            .null_err => "#NULL!",
            .div0 => "#DIV/0!",
            .value => "#VALUE!",
            .ref => "#REF!",
            .name => "#NAME?",
            .num => "#NUM!",
            .na => "#N/A",
            .getting_data => "#GETTING_DATA",
            .spill => "#SPILL!",
            .calc => "#CALC!",
        };
    }

    pub fn fromSpelling(text: []const u8) ?KnownError {
        inline for (@typeInfo(KnownError).@"enum".fields) |f| {
            const k: KnownError = @enumFromInt(f.value);
            if (std.mem.eql(u8, k.spelling(), text)) return k;
        }
        return null;
    }
};

/// Rich errors are **preserved, never produced** (§5.3a): a spelling
/// that arrived through the tokenizer's extensible rule round-trips
/// byte-exact, and no evaluation step invents one.
pub const ErrorValue = union(enum) {
    known: KnownError,
    rich: []const u8,

    pub fn spelling(self: ErrorValue) []const u8 {
        return switch (self) {
            .known => |k| k.spelling(),
            .rich => |text| text,
        };
    }

    pub fn eql(a: ErrorValue, b: ErrorValue) bool {
        return switch (a) {
            .known => |k| b == .known and b.known == k,
            .rich => |t| b == .rich and std.mem.eql(u8, t, b.rich),
        };
    }
};

// ─── scalars (§5.3a) ─────────────────────────────────────────────

/// The only `Matrix` element, and **internal to the evaluator**. The
/// public boundary is `PublishedScalar`, which has no blank.
///
/// `blank` ≠ `missing_arg` ≠ `""`. Blank is an empty *cell*;
/// missing_arg is an omitted *argument* (evaluator-layer `Value`, M3a2);
/// `""` is a text value of length zero. Collapsing any pair of them
/// silently changes `COUNTA`, `ISBLANK`, and every criteria match.
pub const ScalarValue = union(enum) {
    /// Invariant: finite. Enforced at construction — see `fromNumber`.
    number: f64,
    /// Borrowed. Lifetime is the run arena's.
    text: []const u8,
    boolean: bool,
    err: ErrorValue,
    blank,

    /// Construct a number. Asserts finiteness: a non-finite value is
    /// not in Excel's domain (N4a), so producing one is a bug in the
    /// caller, not a value to represent. Arithmetic that *can* overflow
    /// goes through `fromArithmetic`, which converts instead.
    pub fn fromNumber(v: f64) ScalarValue {
        assert(std.math.isFinite(v));
        return .{ .number = v };
    }

    /// N4a at the point of production: a non-finite arithmetic result
    /// becomes `#NUM!`, which is Excel's answer and is shared by both
    /// fidelity modes. NaN from `0/0` is the divider's business —
    /// `divide` raises `#DIV/0!` before ever reaching here.
    pub fn fromArithmetic(v: f64) ScalarValue {
        if (!std.math.isFinite(v)) return .{ .err = .{ .known = .num } };
        return .{ .number = v };
    }

    pub fn errorOf(k: KnownError) ScalarValue {
        return .{ .err = .{ .known = k } };
    }

    pub fn isError(self: ScalarValue) bool {
        return self == .err;
    }

    pub fn eql(a: ScalarValue, b: ScalarValue) bool {
        if (@as(std.meta.Tag(ScalarValue), a) != @as(std.meta.Tag(ScalarValue), b)) return false;
        return switch (a) {
            // Bit equality, so `-0` and `+0` are distinguishable here.
            // Semantic numeric equality is `compareNumbers`.
            .number => |v| @as(u64, @bitCast(v)) == @as(u64, @bitCast(b.number)),
            .text => |t| std.mem.eql(u8, t, b.text),
            .boolean => |v| v == b.boolean,
            .err => |e| e.eql(b.err),
            .blank => true,
        };
    }
};

/// What crosses a public boundary. No blank variant: the boundary
/// applies the one mandatory conversion and callers never have to.
pub const PublishedScalar = union(enum) {
    number: f64,
    text: []const u8,
    boolean: bool,
    err: ErrorValue,

    pub fn eql(a: PublishedScalar, b: PublishedScalar) bool {
        if (@as(std.meta.Tag(PublishedScalar), a) != @as(std.meta.Tag(PublishedScalar), b)) return false;
        return switch (a) {
            .number => |v| @as(u64, @bitCast(v)) == @as(u64, @bitCast(b.number)),
            .text => |t| std.mem.eql(u8, t, b.text),
            .boolean => |v| v == b.boolean,
            .err => |e| e.eql(b.err),
        };
    }
};

/// **The one mandatory blank→0 conversion** (§5.3a), shared by Zig,
/// CLI, C, and Python. Every layer calls this; none reimplements it,
/// which is the entire point of it existing.
///
/// Publication is also where N3's signed-zero policy applies: `.excel`
/// normalizes `-0` to `+0`, `.ieee` preserves it bitwise.
pub fn publish(v: ScalarValue, fidelity: Fidelity) PublishedScalar {
    return switch (v) {
        .blank => .{ .number = 0 },
        .number => |n| .{ .number = applySignedZeroPolicy(n, FpRules.of(fidelity)) },
        .text => |t| .{ .text = t },
        .boolean => |b| .{ .boolean = b },
        .err => |e| .{ .err = e },
    };
}

fn applySignedZeroPolicy(v: f64, rules: FpRules) f64 {
    if (rules.preserve_signed_zero) return v;
    if (v == 0) return 0; // `-0 == 0` is true, so this normalizes the sign
    return v;
}

// ─── matrices (§5.3a, §9) ────────────────────────────────────────

/// §9's `max_matrix_cells`.
pub const max_matrix_cells: usize = 4_000_000;

pub const Shape = struct {
    rows: u32,
    cols: u32,

    pub fn cells(self: Shape) usize {
        return @as(usize, self.rows) * @as(usize, self.cols);
    }

    pub fn eql(a: Shape, b: Shape) bool {
        return a.rows == b.rows and a.cols == b.cols;
    }

    pub fn isScalar(self: Shape) bool {
        return self.rows == 1 and self.cols == 1;
    }
};

pub const MatrixError = error{
    /// Zero rows or zero columns. **Unrepresentable as a result**
    /// (§5.3a): the producing function normalizes it to `#CALC!`, which
    /// is Excel's own answer, and it is never persisted or streamed.
    EmptyMatrix,
    /// §9's `max_matrix_cells`.
    TooManyCells,
    OutOfMemory,
};

pub const Matrix = struct {
    rows: u32,
    cols: u32,
    /// `rows * cols` elements, row-major.
    cells: []ScalarValue,

    pub fn init(allocator: std.mem.Allocator, rows: u32, cols: u32) MatrixError!Matrix {
        if (rows == 0 or cols == 0) return error.EmptyMatrix;
        const n = @as(usize, rows) * @as(usize, cols);
        if (n > max_matrix_cells) return error.TooManyCells;
        const cells = try allocator.alloc(ScalarValue, n);
        @memset(cells, .blank);
        return .{ .rows = rows, .cols = cols, .cells = cells };
    }

    pub fn deinit(self: *Matrix, allocator: std.mem.Allocator) void {
        allocator.free(self.cells);
        self.* = undefined;
    }

    pub fn shape(self: Matrix) Shape {
        return .{ .rows = self.rows, .cols = self.cols };
    }

    pub fn at(self: Matrix, r: u32, c: u32) ScalarValue {
        assert(r < self.rows and c < self.cols);
        return self.cells[@as(usize, r) * self.cols + c];
    }

    pub fn set(self: *Matrix, r: u32, c: u32, v: ScalarValue) void {
        assert(r < self.rows and c < self.cols);
        self.cells[@as(usize, r) * self.cols + c] = v;
    }

    /// The top-left element — §5.3b's reduction for an *array* (not a
    /// reference) in scalar context.
    pub fn topLeft(self: Matrix) ScalarValue {
        assert(self.cells.len > 0);
        return self.cells[0];
    }
};

/// §5.3a: a zero-row or zero-column result is not representable, and
/// the answer is `#CALC!` rather than an error the caller has to invent.
pub fn emptyMatrixResult() ScalarValue {
    return ScalarValue.errorOf(.calc);
}

// ─── decimal ingress (§5.4) ──────────────────────────────────────

pub const ParseDecimalError = error{
    /// Not a number under any reading: `"abc"`. §5.3b's answer in
    /// arithmetic is `#VALUE!`, which is Excel's locale-independent
    /// result, so this is a value outcome and not a refusal.
    NotNumeric,
    /// Parses only under some locale: `"1,5"`, `"1 234"`, `"50%"`,
    /// `"$1.00"`. §5.4b: a `FormulaLocaleSensitiveInput` refusal, never
    /// a guessed number and never a guessed `#VALUE!`.
    LocaleSensitive,
};

/// The ONLY decimal-text→f64 path in the engine (§5.4).
///
/// The invariant grammar is deliberately narrow: optional sign, digits
/// with at most one `.`, optional `e`/`E` exponent with optional sign.
/// No grouping separators, no currency, no percent, no infinities or
/// NaN spellings — every one of those is either locale-flavoured or
/// outside Excel's value domain, and guessing at them is how a
/// spreadsheet engine silently changes someone's numbers.
pub fn parseDecimal(
    fidelity: Fidelity,
    ingress: Ingress,
    text: []const u8,
) ParseDecimalError!f64 {
    var s = text;
    if (ingress.trimsSpace()) s = std.mem.trim(u8, s, " \t");
    if (s.len == 0) return error.NotNumeric;

    if (!isInvariantDecimal(s)) {
        return if (parsesUnderSomeLocale(s)) error.LocaleSensitive else error.NotNumeric;
    }

    const full = std.fmt.parseFloat(f64, s) catch return error.NotNumeric;
    // Overflow to infinity is N4a's business, not the parser's — it
    // hands back the value and `fromArithmetic` turns it into `#NUM!`.
    // Returning it unmodified keeps the two rules separable.

    const rules = FpRules.of(fidelity);
    if (ingress == .cache_import) return full; // N1b: full binary64
    const digits = rules.literal_significant_digits orelse return full;
    return roundToSignificantDigits(full, digits);
}

/// N1a. Round-trips through a fixed-precision scientific rendering,
/// which is exactly "keep N significant decimal digits" and needs no
/// hand-rolled decimal arithmetic.
pub fn roundToSignificantDigits(v: f64, digits: u8) f64 {
    assert(digits >= 1 and digits <= 17);
    if (!std.math.isFinite(v) or v == 0) return v;
    var buf: [64]u8 = undefined;
    const rendered = switch (digits) {
        15 => std.fmt.bufPrint(&buf, "{e:.14}", .{v}) catch return v,
        16 => std.fmt.bufPrint(&buf, "{e:.15}", .{v}) catch return v,
        17 => std.fmt.bufPrint(&buf, "{e:.16}", .{v}) catch return v,
        else => return v,
    };
    return std.fmt.parseFloat(f64, rendered) catch v;
}

fn isInvariantDecimal(s: []const u8) bool {
    var i: usize = 0;
    if (i < s.len and (s[i] == '+' or s[i] == '-')) i += 1;

    var int_digits: usize = 0;
    while (i < s.len and isDigit(s[i])) : (i += 1) int_digits += 1;

    var frac_digits: usize = 0;
    if (i < s.len and s[i] == '.') {
        i += 1;
        while (i < s.len and isDigit(s[i])) : (i += 1) frac_digits += 1;
    }
    if (int_digits == 0 and frac_digits == 0) return false;

    if (i < s.len and (s[i] == 'e' or s[i] == 'E')) {
        i += 1;
        if (i < s.len and (s[i] == '+' or s[i] == '-')) i += 1;
        var exp_digits: usize = 0;
        while (i < s.len and isDigit(s[i])) : (i += 1) exp_digits += 1;
        if (exp_digits == 0) return false;
    }
    return i == s.len;
}

/// Would `s` be numeric under *some* locale convention? Used only to
/// separate `LocaleSensitive` from `NotNumeric` — never to produce a
/// value.
///
/// The bias here matters and runs one way: **when in doubt, say no.**
/// A false positive turns a case Excel answers with `#VALUE!` — a
/// perfectly good plane-1 value — into a plane-2 refusal that stops a
/// recalculation. A false negative merely yields `#VALUE!` where a
/// refusal would have been more informative. So the shapes recognised
/// here are specific: a decimal comma, digit groups of exactly three
/// under one consistent separator, and currency/percent affixes around
/// an otherwise-invariant number. `1.2.3` is none of those and is
/// simply not a number.
fn parsesUnderSomeLocale(text: []const u8) bool {
    var s = text;
    var affix = false;

    // Currency signs, as UTF-8. Only ever leading.
    inline for ([_][]const u8{ "$", "£", "¥", "€" }) |sign| {
        if (std.mem.startsWith(u8, s, sign)) {
            s = s[sign.len..];
            affix = true;
        }
    }
    if (s.len > 0 and s[s.len - 1] == '%') {
        s = s[0 .. s.len - 1];
        affix = true;
    }
    if (s.len == 0) return false;
    if (affix and isInvariantDecimal(s)) return true;

    if (decimalCommaParses(s)) return true;
    inline for ([_]u8{ ',', '.', ' ', '\'' }) |sep| {
        if (groupedParses(s, sep)) return true;
    }
    return false;
}

const locale_scratch_len: usize = 64;

/// A single `,` standing in for the decimal point, with no `.` anywhere
/// to contradict the reading.
fn decimalCommaParses(s: []const u8) bool {
    if (s.len > locale_scratch_len) return false;
    if (std.mem.count(u8, s, ",") != 1) return false;
    if (std.mem.indexOfScalar(u8, s, '.') != null) return false;

    var buf: [locale_scratch_len]u8 = undefined;
    @memcpy(buf[0..s.len], s);
    buf[std.mem.indexOfScalar(u8, buf[0..s.len], ',').?] = '.';
    return isInvariantDecimal(buf[0..s.len]);
}

/// Digit groups under one consistent separator: the first group is 1–3
/// digits and every later one is exactly 3. An optional decimal part
/// may follow, marked by whichever of `.` / `,` is not the grouping
/// separator.
fn groupedParses(s: []const u8, sep: u8) bool {
    if (s.len > locale_scratch_len) return false;
    if (std.mem.indexOfScalar(u8, s, sep) == null) return false;

    var buf: [locale_scratch_len]u8 = undefined;
    var n: usize = 0;
    var i: usize = 0;

    if (i < s.len and (s[i] == '+' or s[i] == '-')) {
        buf[n] = s[i];
        n += 1;
        i += 1;
    }

    var run: usize = 0;
    var first_group = true;
    while (i < s.len and (isDigit(s[i]) or s[i] == sep)) : (i += 1) {
        if (s[i] == sep) {
            if (first_group) {
                if (run < 1 or run > 3) return false;
                first_group = false;
            } else if (run != 3) return false;
            run = 0;
            continue;
        }
        buf[n] = s[i];
        n += 1;
        run += 1;
    }
    // A grouped number always ends on a complete final group.
    if (first_group or run != 3) return false;

    const tail = s[i..];
    if (tail.len == 0) return isInvariantDecimal(buf[0..n]);

    // Only a decimal mark may follow, and it cannot be the separator.
    const mark = tail[0];
    if (mark != '.' and mark != ',') return false;
    if (mark == sep) return false;
    if (n + tail.len > buf.len) return false;
    buf[n] = '.';
    n += 1;
    @memcpy(buf[n .. n + tail.len - 1], tail[1..]);
    n += tail.len - 1;
    return isInvariantDecimal(buf[0..n]);
}

inline fn isDigit(c: u8) bool {
    return c >= '0' and c <= '9';
}

// ─── arithmetic under a rule table (§5.4) ────────────────────────

/// `a + b` / `a - b` with N2 applied. Additive scope is the whole
/// point: the same near-zero result out of a multiplication or a
/// function is left alone.
pub fn addSub(rules: FpRules, a: f64, b: f64, op: enum { add, sub }) ScalarValue {
    const r = switch (op) {
        .add => a + b,
        .sub => a - b,
    };
    return ScalarValue.fromArithmetic(applyZeroSnap(rules, r, a, b));
}

pub fn applyZeroSnap(rules: FpRules, r: f64, a: f64, b: f64) f64 {
    if (!rules.zero_snap) return r;
    if (r == 0 or !std.math.isFinite(r)) return r;
    const scale = @max(@abs(a), @abs(b));
    if (scale == 0) return r;
    const threshold = std.math.ldexp(scale, rules.zero_snap_relative_shift);
    if (@abs(r) < threshold) return 0;
    return r;
}

/// Multiplication is **outside** N2's additive scope, so a near-zero
/// product survives in both modes.
pub fn multiply(rules: FpRules, a: f64, b: f64) ScalarValue {
    _ = rules;
    return ScalarValue.fromArithmetic(a * b);
}

/// `0/0` and `x/0` are both `#DIV/0!` — Excel's answer, shared by both
/// modes. Catching it here is what keeps NaN out of `ScalarValue`
/// entirely rather than relying on `fromArithmetic` to notice.
pub fn divide(rules: FpRules, a: f64, b: f64) ScalarValue {
    _ = rules;
    if (b == 0) return ScalarValue.errorOf(.div0);
    return ScalarValue.fromArithmetic(a / b);
}

// ─── serialization (N5) ──────────────────────────────────────────

/// Caller-side buffer size for `formatNumber`. The scientific form
/// never exceeds 23 bytes (17 mantissa digits + sign + point + `e-308`),
/// and the positional form is only ever chosen when it is shorter.
pub const format_buf_len: usize = 32;

/// N5 — the shortest decimal text that round-trips to the same
/// binary64. Identical in both modes: serialization is not where the
/// modes differ.
///
/// Both presentations share Zig's shortest-round-trip digit generation
/// and differ only in layout, so picking the shorter one is genuinely
/// the shortest round-tripping text. The positional form has to be
/// generated into scratch rather than the caller's buffer because it is
/// unbounded in a way the value is not: `5e-324` positionally is 326
/// bytes, and a caller sizing for the *number* would be right to be
/// surprised by that.
pub fn formatNumber(buf: []u8, v: f64) []const u8 {
    assert(std.math.isFinite(v));
    assert(buf.len >= format_buf_len);

    var sci_buf: [format_buf_len]u8 = undefined;
    const sci = std.fmt.bufPrint(&sci_buf, "{e}", .{v}) catch unreachable;

    var dec_buf: [512]u8 = undefined;
    if (std.fmt.bufPrint(&dec_buf, "{d}", .{v})) |dec| {
        if (dec.len <= sci.len and dec.len <= buf.len) {
            @memcpy(buf[0..dec.len], dec);
            return buf[0..dec.len];
        }
    } else |_| {}

    @memcpy(buf[0..sci.len], sci);
    return buf[0..sci.len];
}

/// N5's companion: whether a decimal text denotes a value binary64
/// holds exactly, rather than one it merely rounds to. Used by the
/// cached-value patcher (M5b1) to know when a rewrite is lossless.
pub fn fitsExactlyInF64(text: []const u8) bool {
    if (!isInvariantDecimal(text)) return false;
    const v = std.fmt.parseFloat(f64, text) catch return false;
    if (!std.math.isFinite(v)) return false;
    var buf: [format_buf_len]u8 = undefined;
    const back = formatNumber(&buf, v);
    // Compare the *values*, not the spellings: `1.0` and `1` denote the
    // same exactly-held number, and only the round-trip of the value
    // decides exactness.
    const reparsed = std.fmt.parseFloat(f64, back) catch return false;
    return @as(u64, @bitCast(reparsed)) == @as(u64, @bitCast(v));
}

// ─── collation_v1 (§5.4b) ────────────────────────────────────────

/// A full non-Turkic case fold over UTF-8. Injected rather than
/// imported — see the module header.
pub const FoldFn = *const fn (std.mem.Allocator, []const u8) anyerror![]u8;

/// **ONE comparator, stated once** (§5.4b): lexicographic order of
/// full-non-Turkic-folded code-point sequences. It governs ordinary
/// `=` `<>` `<` `<=` `>` `>=`, SEARCH, wildcards, lookup equality AND
/// ordering, criteria, and SORT/SORTBY.
///
/// Two consequences worth stating out loud, because both are easy to
/// get wrong and neither is negotiable:
///
///   * **Fold-equal strings are EQUAL for ordering too.** `A`/`a` and
///     `ß`/`ss`/`SS` compare equal under `<` and `>`, not merely under
///     `=`. A raw tie-break would make them unequal, so it is *not*
///     part of the semantic order; SORT/SORTBY use stable source
///     position as a private, non-semantic tie-break instead.
///   * **No Unicode normalization.** Code points as stored. This is
///     deliberately unlike `casefold.excelSheetNameEql`, which applies
///     NFC because sheet-name dedup wants composed/decomposed
///     equivalence. Text comparison does not.
///
/// Byte order over folded UTF-8 *is* code-point order — that is a
/// property of the encoding — so the comparison needs no decoding step.
pub const Collation = struct {
    version: []const u8 = "collation_v1",
    fold: FoldFn,

    pub fn compare(
        self: Collation,
        allocator: std.mem.Allocator,
        a: []const u8,
        b: []const u8,
    ) !std.math.Order {
        const fa = try self.fold(allocator, a);
        defer allocator.free(fa);
        const fb = try self.fold(allocator, b);
        defer allocator.free(fb);
        return std.mem.order(u8, fa, fb);
    }

    pub fn eql(
        self: Collation,
        allocator: std.mem.Allocator,
        a: []const u8,
        b: []const u8,
    ) !bool {
        return (try self.compare(allocator, a, b)) == .eq;
    }
};

/// SORT/SORTBY's tie-break among fold-equal elements. Named so it is
/// visibly *not* part of the comparator.
pub const sort_tiebreak_policy = "stable source position (private, non-semantic)";

/// §5.4b registry match policies. Every text function carries one
/// explicitly; there is no family-wide default hiding a wrong answer.
pub const MatchPolicy = enum {
    /// `=`, SEARCH, wildcard consumers.
    folded,
    /// FIND and SUBSTITUTE are case-SENSITIVE, like EXACT. CODE and
    /// UNICODE operate on raw units.
    raw,
    /// TEXTBEFORE / TEXTAFTER / TEXTSPLIT, via their `match_mode` arg.
    arg_selected,
};

/// Cross-type ordering (§5.3b comparison column): number < text <
/// logical, and no cross-type pair is ever equal.
pub const CrossTypeRank = enum(u2) { number = 0, text = 1, logical = 2 };

pub fn crossTypeRank(v: ScalarValue) ?CrossTypeRank {
    return switch (v) {
        .number => .number,
        .text => .text,
        .boolean => .logical,
        // Blank adopts the other operand's type; errors propagate.
        .blank, .err => null,
    };
}

// ─── shape rules (§5.3b, table 1) ────────────────────────────────

pub const Dialect = enum { dynamic_array, legacy };

pub const ShapeContext = enum {
    scalar_where_array,
    binary_op,
    reference_where_scalar,
    array_where_scalar,
    reference_in_value,
    at_single_cell_reference,
    at_multi_cell_reference,
    at_array,
};

pub const ShapeRule = enum {
    lift_1x1,
    broadcast,
    /// The reference spills or iterates per the function's DA-awareness.
    no_implicit_intersection,
    /// Same-row → row-projected element; same-column → column-projected;
    /// otherwise `#VALUE!`.
    implicit_intersection,
    /// Arrays reduce to their top-left element — **not** row/column
    /// intersection. Excel distinguishes references from arrays and so
    /// does this table.
    top_left_reduction,
    dereference,
    dereference_with_intersection,
    /// `=@A1` yields A1's value regardless of the evaluation site:
    /// Excel's single-item exception precedes intersection.
    reference_unchanged,
    row_col_intersection,
    spill_or_iterate,
};

/// The §5.3b shape table, as a lookup rather than prose.
pub fn shapeRule(ctx: ShapeContext, dialect: Dialect) ShapeRule {
    return switch (ctx) {
        .scalar_where_array => .lift_1x1,
        .binary_op => .broadcast,
        .reference_where_scalar => switch (dialect) {
            .dynamic_array => .no_implicit_intersection,
            .legacy => .implicit_intersection,
        },
        .array_where_scalar => switch (dialect) {
            .dynamic_array => .spill_or_iterate,
            .legacy => .top_left_reduction,
        },
        .reference_in_value => switch (dialect) {
            .dynamic_array => .dereference,
            .legacy => .dereference_with_intersection,
        },
        // The three `@` rows are dialect-independent: `@` is the
        // explicit spelling of what legacy did implicitly.
        .at_single_cell_reference => .reference_unchanged,
        .at_multi_cell_reference => .row_col_intersection,
        .at_array => .top_left_reduction,
    };
}

/// Broadcast two operand shapes. `null` means incompatible — both
/// extents greater than one and unequal — which §5.3b resolves as an
/// elementwise `#N/A` fill, not a refusal.
pub fn broadcastShape(a: Shape, b: Shape) ?Shape {
    const rows = broadcastExtent(a.rows, b.rows) orelse return null;
    const cols = broadcastExtent(a.cols, b.cols) orelse return null;
    return .{ .rows = rows, .cols = cols };
}

fn broadcastExtent(a: u32, b: u32) ?u32 {
    if (a == b) return a;
    if (a == 1) return b;
    if (b == 1) return a;
    return null;
}

/// The value an incompatible broadcast fills its result with.
pub fn incompatibleBroadcastFill() ScalarValue {
    return ScalarValue.errorOf(.na);
}

// ─── scalar coercion matrix (§5.3b, table 2) ─────────────────────

/// Where a scalar came from. Provenance, not type: `numeric_text` and
/// `non_numeric_text` are both `.text` values, and they coerce
/// differently in six of the eight contexts.
pub const Provenance = enum {
    blank_cell,
    empty_text,
    numeric_text,
    locale_text,
    non_numeric_text,
    boolean,
    error_value,
};

pub const CoercionContext = enum {
    arithmetic,
    comparison,
    concat,
    direct_fn_arg_numeric,
    via_range_aggregate,
    lookup_key,
    criteria_operand,
    sort_element,
};

/// What the matrix says happens. These are dispositions, not values:
/// several cells describe a *policy* the consuming layer applies (an
/// aggregate skipping an element, a sort placing one) rather than a
/// conversion producing a scalar.
pub const Disposition = enum {
    /// Coerces to 0.
    zero,
    /// Yields `""`.
    empty_text_result,
    /// `#VALUE!` — a plane-1 value, not a refusal.
    value_error,
    /// Parsed through the invariant grammar.
    coerced_number,
    /// Used as its own text.
    as_text,
    /// `"TRUE"` / `"FALSE"`.
    as_text_bool,
    /// `FormulaLocaleSensitiveInput` — a plane-2 refusal.
    locale_refusal,
    /// Excluded from the aggregate entirely.
    skipped,
    /// Present in the range but not counted.
    ignored,
    /// Excel does **not** coerce numeric text found in a range.
    not_coerced_ignored,
    /// Ignored via ranges, counted as a direct argument — Excel's split,
    /// pinned per aggregate.
    ignored_in_range_counted_direct,
    /// number < text < logical; never cross-equal.
    cross_type_order,
    /// Compared as text under `collation_v1`.
    text_compare,
    matches_blank,
    text_match,
    exact_text_unless_numeric_key,
    empty_criterion,
    blank_rules,
    logical_match,
    logical_rules,
    /// TRUE → 1, FALSE → 0.
    one_or_zero,
    sorts_first,
    text_order,
    logical_order,
    propagates,
    propagates_unless_skip_class,
    pinned_position,
};

/// The §5.3b scalar coercion matrix — 7 provenances × 8 contexts, every
/// cell a fixture. Transcribed as a table so the normative document and
/// the implementation cannot drift apart without a test failing.
pub fn disposition(p: Provenance, c: CoercionContext) Disposition {
    return switch (p) {
        .blank_cell => switch (c) {
            .arithmetic => .zero,
            .comparison => .cross_type_order,
            .concat => .empty_text_result,
            .direct_fn_arg_numeric => .zero,
            .via_range_aggregate => .skipped,
            .lookup_key => .matches_blank,
            .criteria_operand => .blank_rules,
            .sort_element => .sorts_first,
        },
        .empty_text => switch (c) {
            .arithmetic => .value_error,
            .comparison => .text_compare,
            .concat => .as_text,
            .direct_fn_arg_numeric => .value_error,
            .via_range_aggregate => .skipped,
            .lookup_key => .text_match,
            .criteria_operand => .empty_criterion,
            .sort_element => .text_order,
        },
        .numeric_text => switch (c) {
            .arithmetic => .coerced_number,
            .comparison => .cross_type_order,
            .concat => .as_text,
            .direct_fn_arg_numeric => .coerced_number,
            .via_range_aggregate => .not_coerced_ignored,
            .lookup_key => .exact_text_unless_numeric_key,
            .criteria_operand => .coerced_number,
            .sort_element => .text_order,
        },
        .locale_text => switch (c) {
            .arithmetic => .locale_refusal,
            .comparison => .text_compare,
            .concat => .as_text,
            .direct_fn_arg_numeric => .locale_refusal,
            .via_range_aggregate => .ignored,
            .lookup_key => .text_match,
            .criteria_operand => .locale_refusal,
            .sort_element => .text_order,
        },
        .non_numeric_text => switch (c) {
            .arithmetic => .value_error,
            .comparison => .text_compare,
            .concat => .as_text,
            .direct_fn_arg_numeric => .value_error,
            .via_range_aggregate => .ignored,
            .lookup_key => .text_match,
            .criteria_operand => .text_match,
            .sort_element => .text_order,
        },
        .boolean => switch (c) {
            .arithmetic => .one_or_zero,
            .comparison => .cross_type_order,
            .concat => .as_text_bool,
            .direct_fn_arg_numeric => .one_or_zero,
            .via_range_aggregate => .ignored_in_range_counted_direct,
            .lookup_key => .logical_match,
            .criteria_operand => .logical_rules,
            .sort_element => .logical_order,
        },
        .error_value => switch (c) {
            .arithmetic => .propagates,
            .comparison => .propagates,
            .concat => .propagates,
            .direct_fn_arg_numeric => .propagates,
            .via_range_aggregate => .propagates_unless_skip_class,
            .lookup_key => .propagates,
            .criteria_operand => .propagates,
            .sort_element => .pinned_position,
        },
    };
}

/// Classify a scalar's provenance for the matrix. Text splits three
/// ways, which is the whole reason provenance exists as a concept
/// separate from the union tag.
///
/// `null` for an actual number: the matrix has no row for one, because
/// there is nothing to coerce. Returning a row anyway would invite a
/// caller to look up a disposition that the normative table never
/// stated.
pub fn provenanceOf(v: ScalarValue) ?Provenance {
    return switch (v) {
        .number => null,
        .blank => .blank_cell,
        .boolean => .boolean,
        .err => .error_value,
        .text => |t| blk: {
            if (t.len == 0) break :blk .empty_text;
            const trimmed = std.mem.trim(u8, t, " \t");
            if (isInvariantDecimal(trimmed)) break :blk .numeric_text;
            if (parsesUnderSomeLocale(trimmed)) break :blk .locale_text;
            break :blk .non_numeric_text;
        },
    };
}

/// The behavioural half of the matrix's two numeric columns. Returns
/// either a number, a plane-1 value (`#VALUE!` or a propagated error),
/// or a plane-2 refusal signal the caller turns into
/// `FormulaLocaleSensitiveInput`.
pub const NumericCoercion = union(enum) {
    number: f64,
    value: ScalarValue,
    locale_refusal,
};

pub fn coerceToNumber(
    v: ScalarValue,
    fidelity: Fidelity,
    ingress: Ingress,
) NumericCoercion {
    return switch (v) {
        .number => |n| .{ .number = n },
        .blank => .{ .number = 0 },
        .boolean => |b| .{ .number = if (b) 1 else 0 },
        .err => .{ .value = v },
        .text => |t| blk: {
            const n = parseDecimal(fidelity, ingress, t) catch |err| switch (err) {
                error.NotNumeric => break :blk .{ .value = ScalarValue.errorOf(.value) },
                error.LocaleSensitive => break :blk .locale_refusal,
            };
            break :blk .{ .number = n };
        },
    };
}

// ─── error propagation order (§5.3c) ─────────────────────────────

/// A function's propagation class. Provenance-aware **per function**,
/// never family-wide: `COUNT` ignores errors in ranges while `COUNTA`
/// counts them, and the two live in the same family.
pub const PropagationClass = enum {
    propagate,
    /// ISERROR / IFERROR / IFNA — they look at an error without
    /// becoming one.
    observe,
    /// Elementwise array ops keep errors per cell.
    per_element,
    per_function_provenance,
};

/// §5.3c: operator operands evaluate left-to-right and the first error
/// encountered is the result.
pub fn propagateBinary(lhs: ScalarValue, rhs: ScalarValue) ?ErrorValue {
    if (lhs == .err) return lhs.err;
    if (rhs == .err) return rhs.err;
    return null;
}

/// Eager function arguments evaluate in declaration order; first error
/// wins unless the registry's class says otherwise.
pub fn firstError(values: []const ScalarValue) ?ErrorValue {
    for (values) |v| {
        if (v == .err) return v.err;
    }
    return null;
}

// ─── divergence points (§5.4, the "Divergence ×2" gate) ──────────

/// Whether the two rule tables must agree or must differ at a point.
/// Stating it per point is what makes the gate meaningful: a test that
/// only checked for differences would pass if the modes diverged
/// everywhere, which would be just as wrong.
pub const Divergence = enum { must_differ, must_agree };

pub const DivergencePoint = struct {
    /// §5.4's rule identifier.
    rule: []const u8,
    what: []const u8,
    expect: Divergence,
    /// Whether a committed oracle manifest decides this point, or it
    /// rests on §5.4's text alone. Labelled rather than blurred.
    evidence: enum { oracle, spec_pinned },
};

/// Every place §5.4 says the modes do or do not diverge. The gate walks
/// this list; adding a rule without a row here fails the count check.
pub const divergence_points = [_]DivergencePoint{
    .{ .rule = "N1a", .what = "17-significant-digit literal ingress", .expect = .must_differ, .evidence = .spec_pinned },
    .{ .rule = "N1b", .what = "cached <v> import is full binary64", .expect = .must_agree, .evidence = .spec_pinned },
    .{ .rule = "N2", .what = "near-zero additive result", .expect = .must_differ, .evidence = .spec_pinned },
    .{ .rule = "N2", .what = "near-zero product is outside additive scope", .expect = .must_agree, .evidence = .spec_pinned },
    .{ .rule = "N3", .what = "signed zero at publication", .expect = .must_differ, .evidence = .spec_pinned },
    .{ .rule = "N3", .what = "subnormals preserved", .expect = .must_agree, .evidence = .oracle },
    .{ .rule = "N4a", .what = "overflow is #NUM!", .expect = .must_agree, .evidence = .oracle },
    .{ .rule = "N4a", .what = "division by zero is #DIV/0!", .expect = .must_agree, .evidence = .oracle },
    .{ .rule = "N5", .what = "shortest-round-trip serialization", .expect = .must_agree, .evidence = .spec_pinned },
};

// ─── tests ───────────────────────────────────────────────────────
//
// The concrete fold is imported HERE, in the test section, and nowhere
// else. A file-scope `const` referenced only from a `test` block is not
// resolved in a non-test build — verified against Zig 0.16.0 — so a
// module built from this file without declaring `zlsx_casefold`
// compiles fine. That is what lets the fixtures run against the shipped
// algorithm while `collation_v1` stays injection-based, and it is what
// keeps `src/unicode/casefold.zig` from having to belong to two module
// trees at once (see the module header).

const casefold = @import("zlsx_casefold");
const testing = std.testing;

fn shippedFold(allocator: std.mem.Allocator, s: []const u8) anyerror![]u8 {
    return casefold.foldString(allocator, s);
}

const collation_v1: Collation = .{ .fold = &shippedFold };

fn bits(v: f64) u64 {
    return @bitCast(v);
}

// ─── scalars, matrices, publication (§5.3a) ──────────────────────

test "ScalarValue: blank, empty text, and zero are three different things" {
    const blank: ScalarValue = .blank;
    const empty: ScalarValue = .{ .text = "" };
    const zero = ScalarValue.fromNumber(0);

    try testing.expect(!blank.eql(empty));
    try testing.expect(!blank.eql(zero));
    try testing.expect(!empty.eql(zero));
    try testing.expectEqual(Provenance.blank_cell, provenanceOf(blank).?);
    try testing.expectEqual(Provenance.empty_text, provenanceOf(empty).?);
    try testing.expect(provenanceOf(zero) == null);
}

test "ScalarValue: a non-finite arithmetic result is #NUM!, not a number" {
    try testing.expect(ScalarValue.fromArithmetic(std.math.inf(f64)).eql(ScalarValue.errorOf(.num)));
    try testing.expect(ScalarValue.fromArithmetic(-std.math.inf(f64)).eql(ScalarValue.errorOf(.num)));
    try testing.expect(ScalarValue.fromArithmetic(std.math.nan(f64)).eql(ScalarValue.errorOf(.num)));
    try testing.expect(ScalarValue.fromArithmetic(1.5).eql(ScalarValue.fromNumber(1.5)));
    // The invariant holds for every finite input, including the edges.
    for ([_]f64{ 0, -0.0, 5e-324, 1.7976931348623157e308, -1.7976931348623157e308 }) |v| {
        const s = ScalarValue.fromArithmetic(v);
        try testing.expect(s == .number);
        try testing.expect(std.math.isFinite(s.number));
    }
}

test "ScalarValue: number equality is bitwise, so -0 and +0 are distinguishable" {
    try testing.expect(!ScalarValue.fromNumber(-0.0).eql(ScalarValue.fromNumber(0.0)));
}

test "errors: the frozen ten round-trip through their spellings" {
    inline for (@typeInfo(KnownError).@"enum".fields) |f| {
        const k: KnownError = @enumFromInt(f.value);
        try testing.expectEqual(k, KnownError.fromSpelling(k.spelling()).?);
    }
    try testing.expect(KnownError.fromSpelling("#BLOCKED!") == null);
}

test "errors: rich errors are preserved and compare by spelling" {
    const a: ErrorValue = .{ .rich = "#BLOCKED!" };
    const b: ErrorValue = .{ .rich = "#BLOCKED!" };
    const c: ErrorValue = .{ .rich = "#PYTHON!" };
    try testing.expect(a.eql(b));
    try testing.expect(!a.eql(c));
    try testing.expect(!a.eql(.{ .known = .value }));
    try testing.expectEqualStrings("#BLOCKED!", a.spelling());
}

test "Matrix: invariants are enforced at construction" {
    try testing.expectError(error.EmptyMatrix, Matrix.init(testing.allocator, 0, 3));
    try testing.expectError(error.EmptyMatrix, Matrix.init(testing.allocator, 3, 0));
    try testing.expectError(error.TooManyCells, Matrix.init(testing.allocator, 4_000_001, 2));

    var m = try Matrix.init(testing.allocator, 2, 3);
    defer m.deinit(testing.allocator);
    try testing.expectEqual(@as(usize, 6), m.cells.len);
    // Fresh cells are blank, not zero — the two are not the same value.
    try testing.expect(m.at(0, 0).eql(.blank));

    m.set(1, 2, ScalarValue.fromNumber(7));
    try testing.expect(m.at(1, 2).eql(ScalarValue.fromNumber(7)));
    m.set(0, 0, ScalarValue.fromNumber(1));
    try testing.expect(m.topLeft().eql(ScalarValue.fromNumber(1)));
}

test "Matrix: an empty result is #CALC!, never an empty matrix" {
    try testing.expect(emptyMatrixResult().eql(ScalarValue.errorOf(.calc)));
}

test "publish: the one mandatory blank→0 conversion, in both modes" {
    inline for ([_]Fidelity{ .excel, .ieee }) |f| {
        try testing.expect(publish(.blank, f).eql(.{ .number = 0 }));
        try testing.expect(publish(.{ .text = "x" }, f).eql(.{ .text = "x" }));
        try testing.expect(publish(.{ .boolean = true }, f).eql(.{ .boolean = true }));
        try testing.expect(publish(ScalarValue.errorOf(.na), f).eql(.{ .err = .{ .known = .na } }));
    }
}

test "publish: N3 signed-zero policy differs by mode" {
    const neg_zero = ScalarValue.fromNumber(-0.0);
    try testing.expectEqual(bits(0.0), bits(publish(neg_zero, .excel).number));
    try testing.expectEqual(bits(-0.0), bits(publish(neg_zero, .ieee).number));
    // A non-zero value is untouched by the policy in either mode.
    try testing.expectEqual(bits(-1.5), bits(publish(ScalarValue.fromNumber(-1.5), .excel).number));
}

// ─── parseDecimal (§5.4) ─────────────────────────────────────────

test "parseDecimal: the invariant grammar" {
    const ok = [_][]const u8{
        "0",      "1",     "-1",  "+1",     "1.5",                  "-1.5",
        ".5",     "5.",    "1e5", "1E+5",   "1.5e-10",              "0.0",
        "000123", "1e308", "-0",  "1E-308", "12345678901234567890",
    };
    for (ok) |s| {
        _ = parseDecimal(.ieee, .literal, s) catch |err| {
            std.debug.print("expected `{s}` to parse, got {t}\n", .{ s, err });
            return error.ShouldHaveParsed;
        };
    }
}

test "parseDecimal: not numeric under any reading" {
    const cases = [_][]const u8{ "abc", "", "e5", "1e", "1e+", "--1", "1.2.3", "#N/A", "1x" };
    for (cases) |s| {
        try testing.expectError(error.NotNumeric, parseDecimal(.ieee, .text_coercion, s));
    }
}

test "parseDecimal: locale-flavoured text refuses, never guesses" {
    // §5.3b: never a guessed number and never a guessed `#VALUE!`.
    const cases = [_][]const u8{ "1,5", "1 234", "1'234.5", "1.234,56", "50%", "$1.00" };
    for (cases) |s| {
        try testing.expectError(error.LocaleSensitive, parseDecimal(.ieee, .text_coercion, s));
    }
}

test "parseDecimal: the locale classifier errs toward #VALUE!, not toward refusing" {
    // A false positive here would turn a case Excel answers with a
    // plane-1 `#VALUE!` into a plane-2 refusal that stops a whole
    // recalculation. These are the shapes that tempt a looser
    // classifier and must NOT be read as locale-flavoured.
    const not_numeric = [_][]const u8{
        "1.2.3", // not groups of three under either reading
        "1,2,3,4", // ditto
        "1,23,456", // first group fine, second is two digits
        "1 23", // grouped space, but the group is two digits
        "..",
        ",,",
        "1..2",
        "1,,2",
        "%",
        "$",
        "1%%",
        "%1",
        "e",
        "1e5e5",
    };
    for (not_numeric) |s| {
        try testing.expectError(error.NotNumeric, parseDecimal(.ieee, .text_coercion, s));
    }

    // …and these genuinely are locale-flavoured.
    const locale = [_][]const u8{
        "1,5",   "-1,5",      "1.234,56",  "1,234.56",
        "1 234", "1 234 567", "1'234",     "1'234.5",
        "12%",   "-12,5%",    "$1,234.50",
    };
    for (locale) |s| {
        try testing.expectError(error.LocaleSensitive, parseDecimal(.ieee, .text_coercion, s));
    }
}

test "parseDecimal: space handling follows the ingress, not the text" {
    // Stored forms are exact; coercion paths see user text.
    try testing.expectError(error.NotNumeric, parseDecimal(.ieee, .literal, " 1 "));
    try testing.expectError(error.NotNumeric, parseDecimal(.ieee, .cache_import, " 1 "));
    try testing.expectEqual(@as(f64, 1), try parseDecimal(.ieee, .text_coercion, " 1 "));
    try testing.expectEqual(@as(f64, 1), try parseDecimal(.ieee, .function_arg, "\t1"));
    try testing.expectEqual(@as(f64, 1), try parseDecimal(.ieee, .criteria, "1 "));
}

test "parseDecimal: N1a rounds on ingress under .excel, N1b never does" {
    const text = "1.2345678901234567";
    const full = try std.fmt.parseFloat(f64, text);

    // `.ieee` keeps every digit, on every ingress.
    inline for (@typeInfo(Ingress).@"enum".fields) |f| {
        const ing: Ingress = @enumFromInt(f.value);
        try testing.expectEqual(bits(full), bits(try parseDecimal(.ieee, ing, text)));
    }

    // `.excel` keeps 15 significant digits — except on `.cache_import`,
    // which is N1b: a cached <v> is already binary64 and re-rounding it
    // would corrupt a value nobody asked us to reinterpret.
    const rounded = try std.fmt.parseFloat(f64, "1.23456789012346");
    try testing.expectEqual(bits(rounded), bits(try parseDecimal(.excel, .literal, text)));
    try testing.expectEqual(bits(rounded), bits(try parseDecimal(.excel, .text_coercion, text)));
    try testing.expectEqual(bits(rounded), bits(try parseDecimal(.excel, .function_arg, text)));
    try testing.expectEqual(bits(rounded), bits(try parseDecimal(.excel, .criteria, text)));
    try testing.expectEqual(bits(full), bits(try parseDecimal(.excel, .cache_import, text)));
    try testing.expect(bits(rounded) != bits(full));
}

test "parseDecimal: the paired all-ingress fixture in both modes" {
    // §5.4: "Paired fixtures push identical decimal text through every
    // path in both fidelity modes; a second parser diverging is
    // structurally impossible." There is one parser, so the fixture
    // proves the ingress table rather than parser agreement.
    const texts = [_][]const u8{ "0", "1", "-2.5", "1e-5", "0.1", "1.2345678901234567", "9007199254740993" };
    for (texts) |text| {
        inline for ([_]Fidelity{ .excel, .ieee }) |fid| {
            var seen: ?u64 = null;
            inline for (@typeInfo(Ingress).@"enum".fields) |f| {
                const ing: Ingress = @enumFromInt(f.value);
                const v = try parseDecimal(fid, ing, text);
                try testing.expect(std.math.isFinite(v));
                // Every ingress but `.cache_import` shares one answer.
                if (ing != .cache_import) {
                    if (seen) |prev| {
                        try testing.expectEqual(prev, bits(v));
                    } else seen = bits(v);
                }
            }
        }
    }
}

// ─── arithmetic under the rule tables (§5.4) ─────────────────────

test "N2: zero-snap is additive-scope only" {
    const ex = excel_fp_rules_v1;
    const ie = ieee_fp_rules_v1;

    // The textbook case: 1.1 - 1.0 - 0.1.
    const step1_ex = addSub(ex, 1.1, 1.0, .sub);
    const residue_ex = addSub(ex, step1_ex.number, 0.1, .sub);
    try testing.expectEqual(@as(f64, 0), residue_ex.number);

    const step1_ie = addSub(ie, 1.1, 1.0, .sub);
    const residue_ie = addSub(ie, step1_ie.number, 0.1, .sub);
    try testing.expect(residue_ie.number != 0);

    // …and the counterexamples: a product and a quotient of the same
    // magnitude are NOT snapped, in either mode.
    const tiny = 1e-20;
    try testing.expect(multiply(ex, tiny, tiny).number != 0);
    try testing.expect(divide(ex, tiny, 1e20).number != 0);
    try testing.expect(multiply(ie, tiny, tiny).number != 0);

    // A large result is never near zero, so no mode touches it.
    try testing.expectEqual(bits(addSub(ex, 0.1, 0.2, .add).number), bits(addSub(ie, 0.1, 0.2, .add).number));
}

test "N4a: overflow is #NUM! and division by zero is #DIV/0!, in both modes" {
    inline for ([_]FpRules{ excel_fp_rules_v1, ieee_fp_rules_v1 }) |rules| {
        try testing.expect(multiply(rules, 1e308, 10).eql(ScalarValue.errorOf(.num)));
        try testing.expect(divide(rules, 1, 0).eql(ScalarValue.errorOf(.div0)));
        try testing.expect(divide(rules, 0, 0).eql(ScalarValue.errorOf(.div0)));
        try testing.expect(divide(rules, -1, 0).eql(ScalarValue.errorOf(.div0)));
        // Underflow to zero is a value, not an error.
        try testing.expect(multiply(rules, 1e-300, 1e-300).eql(ScalarValue.fromNumber(0)));
    }
}

// ─── N5 serialization ────────────────────────────────────────────

test "N5: the shortest round-tripping form, identical in both modes" {
    const cases = [_]f64{
        0,     1,      -1,     1.5,       0.1,                0.30000000000000004,
        1e308, 5e-324, 1e-318, 1.0 / 3.0, 9007199254740993.0, 2.2250738585072014e-308,
    };
    for (cases) |v| {
        var buf: [format_buf_len]u8 = undefined;
        const text = formatNumber(&buf, v);
        const back = try std.fmt.parseFloat(f64, text);
        try testing.expectEqual(bits(v), bits(back));
        // Shorter than 32 bytes for every value, including the
        // subnormals whose positional form runs to 326 bytes.
        try testing.expect(text.len <= format_buf_len);
    }
}

test "N5: fitsExactlyInF64" {
    try testing.expect(fitsExactlyInF64("1"));
    try testing.expect(fitsExactlyInF64("1.5"));
    try testing.expect(fitsExactlyInF64("0.5"));
    try testing.expect(fitsExactlyInF64("-0"));
    try testing.expect(!fitsExactlyInF64("abc"));
    try testing.expect(!fitsExactlyInF64("1,5"));
}

// ─── collation_v1 (§5.4b) ────────────────────────────────────────

test "collation_v1: A/a fold-equal, and equal for ordering too" {
    try testing.expect(try collation_v1.eql(testing.allocator, "A", "a"));
    try testing.expect(try collation_v1.eql(testing.allocator, "HELLO", "hello"));
    try testing.expectEqual(
        std.math.Order.eq,
        try collation_v1.compare(testing.allocator, "Apple", "APPLE"),
    );
    // Fold-equal must be EQUAL under `<`/`>` as well — a raw tie-break
    // would make them unequal and is therefore not part of the order.
    try testing.expect((try collation_v1.compare(testing.allocator, "a", "A")) != .lt);
    try testing.expect((try collation_v1.compare(testing.allocator, "a", "A")) != .gt);
}

test "collation_v1: ß, ss, and SS are one string" {
    try testing.expect(try collation_v1.eql(testing.allocator, "ß", "ss"));
    try testing.expect(try collation_v1.eql(testing.allocator, "ß", "SS"));
    try testing.expect(try collation_v1.eql(testing.allocator, "straße", "STRASSE"));
    try testing.expectEqual(
        std.math.Order.eq,
        try collation_v1.compare(testing.allocator, "Straße", "strasse"),
    );
}

test "collation_v1: astral code points fold and order" {
    // U+1D400 MATHEMATICAL BOLD CAPITAL A has no fold; U+10400 DESERET
    // CAPITAL LONG I folds to U+10428.
    try testing.expect(try collation_v1.eql(testing.allocator, "\u{10400}", "\u{10428}"));
    try testing.expect(!try collation_v1.eql(testing.allocator, "\u{1D400}", "\u{1D41A}"));
    // Byte order over folded UTF-8 is code-point order.
    try testing.expectEqual(
        std.math.Order.lt,
        try collation_v1.compare(testing.allocator, "z", "\u{10428}"),
    );
}

test "collation_v1: no Unicode normalization is applied" {
    // Precomposed é vs e + combining acute are DIFFERENT text, unlike
    // sheet-name dedup, which normalizes on purpose.
    const precomposed = "caf\u{00E9}";
    const decomposed = "cafe\u{0301}";
    try testing.expect(!try collation_v1.eql(testing.allocator, precomposed, decomposed));
    try testing.expect(casefold.excelSheetNameEql(precomposed, decomposed));
}

test "collation_v1: ordering is lexicographic over folded sequences" {
    try testing.expectEqual(std.math.Order.lt, try collation_v1.compare(testing.allocator, "a", "b"));
    try testing.expectEqual(std.math.Order.gt, try collation_v1.compare(testing.allocator, "B", "a"));
    try testing.expectEqual(std.math.Order.lt, try collation_v1.compare(testing.allocator, "ab", "abc"));
    try testing.expectEqual(std.math.Order.eq, try collation_v1.compare(testing.allocator, "", ""));
    try testing.expectEqual(std.math.Order.lt, try collation_v1.compare(testing.allocator, "", "a"));
    try testing.expectEqualStrings("collation_v1", collation_v1.version);
}

test "collation_v1: cross-type ranking is number < text < logical" {
    try testing.expectEqual(CrossTypeRank.number, crossTypeRank(ScalarValue.fromNumber(1)).?);
    try testing.expectEqual(CrossTypeRank.text, crossTypeRank(.{ .text = "a" }).?);
    try testing.expectEqual(CrossTypeRank.logical, crossTypeRank(.{ .boolean = true }).?);
    try testing.expect(@intFromEnum(CrossTypeRank.number) < @intFromEnum(CrossTypeRank.text));
    try testing.expect(@intFromEnum(CrossTypeRank.text) < @intFromEnum(CrossTypeRank.logical));
    // Blank adopts the other operand's type; errors propagate. Neither
    // has a rank of its own.
    try testing.expect(crossTypeRank(.blank) == null);
    try testing.expect(crossTypeRank(ScalarValue.errorOf(.na)) == null);
}

// ─── shape rules (§5.3b, table 1) ────────────────────────────────

test "§5.3b: the shape table, every row under both dialects" {
    const Case = struct { ctx: ShapeContext, da: ShapeRule, legacy: ShapeRule };
    const rows = [_]Case{
        .{ .ctx = .scalar_where_array, .da = .lift_1x1, .legacy = .lift_1x1 },
        .{ .ctx = .binary_op, .da = .broadcast, .legacy = .broadcast },
        .{ .ctx = .reference_where_scalar, .da = .no_implicit_intersection, .legacy = .implicit_intersection },
        .{ .ctx = .array_where_scalar, .da = .spill_or_iterate, .legacy = .top_left_reduction },
        .{ .ctx = .reference_in_value, .da = .dereference, .legacy = .dereference_with_intersection },
        .{ .ctx = .at_single_cell_reference, .da = .reference_unchanged, .legacy = .reference_unchanged },
        .{ .ctx = .at_multi_cell_reference, .da = .row_col_intersection, .legacy = .row_col_intersection },
        .{ .ctx = .at_array, .da = .top_left_reduction, .legacy = .top_left_reduction },
    };
    for (rows) |r| {
        try testing.expectEqual(r.da, shapeRule(r.ctx, .dynamic_array));
        try testing.expectEqual(r.legacy, shapeRule(r.ctx, .legacy));
    }
    // The table covers every context exactly once.
    try testing.expectEqual(@typeInfo(ShapeContext).@"enum".fields.len, rows.len);
}

test "§5.3b: an array in scalar context reduces top-left, a reference intersects" {
    // The distinction Excel draws and the table preserves: arrays and
    // references are not the same thing in legacy dialects.
    try testing.expectEqual(ShapeRule.top_left_reduction, shapeRule(.array_where_scalar, .legacy));
    try testing.expectEqual(ShapeRule.implicit_intersection, shapeRule(.reference_where_scalar, .legacy));
    try testing.expect(shapeRule(.array_where_scalar, .legacy) != shapeRule(.reference_where_scalar, .legacy));
}

test "§5.3b: broadcast, and the elementwise #N/A fill when it fails" {
    const s = struct {
        fn sh(r: u32, c: u32) Shape {
            return .{ .rows = r, .cols = c };
        }
    };
    try testing.expect(broadcastShape(s.sh(1, 1), s.sh(3, 4)).?.eql(s.sh(3, 4)));
    try testing.expect(broadcastShape(s.sh(3, 1), s.sh(1, 4)).?.eql(s.sh(3, 4)));
    try testing.expect(broadcastShape(s.sh(2, 2), s.sh(2, 2)).?.eql(s.sh(2, 2)));
    try testing.expect(broadcastShape(s.sh(3, 1), s.sh(2, 1)) == null);
    try testing.expect(broadcastShape(s.sh(1, 3), s.sh(1, 2)) == null);
    try testing.expect(incompatibleBroadcastFill().eql(ScalarValue.errorOf(.na)));
}

// ─── scalar coercion matrix (§5.3b, table 2) ─────────────────────

test "§5.3b: every cell of the 7×8 scalar coercion matrix" {
    const Cell = struct { p: Provenance, c: CoercionContext, d: Disposition };
    const cells = [_]Cell{
        // blank cell
        .{ .p = .blank_cell, .c = .arithmetic, .d = .zero },
        .{ .p = .blank_cell, .c = .comparison, .d = .cross_type_order },
        .{ .p = .blank_cell, .c = .concat, .d = .empty_text_result },
        .{ .p = .blank_cell, .c = .direct_fn_arg_numeric, .d = .zero },
        .{ .p = .blank_cell, .c = .via_range_aggregate, .d = .skipped },
        .{ .p = .blank_cell, .c = .lookup_key, .d = .matches_blank },
        .{ .p = .blank_cell, .c = .criteria_operand, .d = .blank_rules },
        .{ .p = .blank_cell, .c = .sort_element, .d = .sorts_first },
        // "" text
        .{ .p = .empty_text, .c = .arithmetic, .d = .value_error },
        .{ .p = .empty_text, .c = .comparison, .d = .text_compare },
        .{ .p = .empty_text, .c = .concat, .d = .as_text },
        .{ .p = .empty_text, .c = .direct_fn_arg_numeric, .d = .value_error },
        .{ .p = .empty_text, .c = .via_range_aggregate, .d = .skipped },
        .{ .p = .empty_text, .c = .lookup_key, .d = .text_match },
        .{ .p = .empty_text, .c = .criteria_operand, .d = .empty_criterion },
        .{ .p = .empty_text, .c = .sort_element, .d = .text_order },
        // numeric text
        .{ .p = .numeric_text, .c = .arithmetic, .d = .coerced_number },
        .{ .p = .numeric_text, .c = .comparison, .d = .cross_type_order },
        .{ .p = .numeric_text, .c = .concat, .d = .as_text },
        .{ .p = .numeric_text, .c = .direct_fn_arg_numeric, .d = .coerced_number },
        .{ .p = .numeric_text, .c = .via_range_aggregate, .d = .not_coerced_ignored },
        .{ .p = .numeric_text, .c = .lookup_key, .d = .exact_text_unless_numeric_key },
        .{ .p = .numeric_text, .c = .criteria_operand, .d = .coerced_number },
        .{ .p = .numeric_text, .c = .sort_element, .d = .text_order },
        // locale-flavoured text
        .{ .p = .locale_text, .c = .arithmetic, .d = .locale_refusal },
        .{ .p = .locale_text, .c = .comparison, .d = .text_compare },
        .{ .p = .locale_text, .c = .concat, .d = .as_text },
        .{ .p = .locale_text, .c = .direct_fn_arg_numeric, .d = .locale_refusal },
        .{ .p = .locale_text, .c = .via_range_aggregate, .d = .ignored },
        .{ .p = .locale_text, .c = .lookup_key, .d = .text_match },
        .{ .p = .locale_text, .c = .criteria_operand, .d = .locale_refusal },
        .{ .p = .locale_text, .c = .sort_element, .d = .text_order },
        // non-numeric text
        .{ .p = .non_numeric_text, .c = .arithmetic, .d = .value_error },
        .{ .p = .non_numeric_text, .c = .comparison, .d = .text_compare },
        .{ .p = .non_numeric_text, .c = .concat, .d = .as_text },
        .{ .p = .non_numeric_text, .c = .direct_fn_arg_numeric, .d = .value_error },
        .{ .p = .non_numeric_text, .c = .via_range_aggregate, .d = .ignored },
        .{ .p = .non_numeric_text, .c = .lookup_key, .d = .text_match },
        .{ .p = .non_numeric_text, .c = .criteria_operand, .d = .text_match },
        .{ .p = .non_numeric_text, .c = .sort_element, .d = .text_order },
        // boolean
        .{ .p = .boolean, .c = .arithmetic, .d = .one_or_zero },
        .{ .p = .boolean, .c = .comparison, .d = .cross_type_order },
        .{ .p = .boolean, .c = .concat, .d = .as_text_bool },
        .{ .p = .boolean, .c = .direct_fn_arg_numeric, .d = .one_or_zero },
        .{ .p = .boolean, .c = .via_range_aggregate, .d = .ignored_in_range_counted_direct },
        .{ .p = .boolean, .c = .lookup_key, .d = .logical_match },
        .{ .p = .boolean, .c = .criteria_operand, .d = .logical_rules },
        .{ .p = .boolean, .c = .sort_element, .d = .logical_order },
        // error
        .{ .p = .error_value, .c = .arithmetic, .d = .propagates },
        .{ .p = .error_value, .c = .comparison, .d = .propagates },
        .{ .p = .error_value, .c = .concat, .d = .propagates },
        .{ .p = .error_value, .c = .direct_fn_arg_numeric, .d = .propagates },
        .{ .p = .error_value, .c = .via_range_aggregate, .d = .propagates_unless_skip_class },
        .{ .p = .error_value, .c = .lookup_key, .d = .propagates },
        .{ .p = .error_value, .c = .criteria_operand, .d = .propagates },
        .{ .p = .error_value, .c = .sort_element, .d = .pinned_position },
    };

    const provenances = @typeInfo(Provenance).@"enum".fields.len;
    const contexts = @typeInfo(CoercionContext).@"enum".fields.len;
    try testing.expectEqual(provenances * contexts, cells.len);

    // Every cell asserted, and every (provenance, context) pair covered
    // exactly once — so neither the table nor this fixture can lose a
    // row without the other noticing.
    var seen = std.AutoHashMap(struct { Provenance, CoercionContext }, void).init(testing.allocator);
    defer seen.deinit();
    for (cells) |cell| {
        try testing.expectEqual(cell.d, disposition(cell.p, cell.c));
        const key = .{ cell.p, cell.c };
        try testing.expect(!seen.contains(key));
        try seen.put(key, {});
    }
    try testing.expectEqual(cells.len, seen.count());
}

test "§5.3b: provenance classification splits text three ways" {
    const Case = struct { v: ScalarValue, p: Provenance };
    const cases = [_]Case{
        .{ .v = .blank, .p = .blank_cell },
        .{ .v = .{ .text = "" }, .p = .empty_text },
        .{ .v = .{ .text = "1" }, .p = .numeric_text },
        .{ .v = .{ .text = "-2.5e3" }, .p = .numeric_text },
        .{ .v = .{ .text = " 42 " }, .p = .numeric_text },
        .{ .v = .{ .text = "1,5" }, .p = .locale_text },
        .{ .v = .{ .text = "50%" }, .p = .locale_text },
        .{ .v = .{ .text = "abc" }, .p = .non_numeric_text },
        .{ .v = .{ .boolean = true }, .p = .boolean },
        .{ .v = ScalarValue.errorOf(.na), .p = .error_value },
    };
    for (cases) |c| {
        try testing.expectEqual(c.p, provenanceOf(c.v).?);
    }
}

test "§5.3b: coerceToNumber implements the two numeric columns" {
    try testing.expectEqual(@as(f64, 0), coerceToNumber(.blank, .excel, .text_coercion).number);
    try testing.expectEqual(@as(f64, 1), coerceToNumber(.{ .boolean = true }, .excel, .text_coercion).number);
    try testing.expectEqual(@as(f64, 0), coerceToNumber(.{ .boolean = false }, .excel, .text_coercion).number);
    try testing.expectEqual(@as(f64, 2), coerceToNumber(.{ .text = "2" }, .excel, .text_coercion).number);
    // `"1"+1 = 2` — the §5.3b worked example.
    try testing.expectEqual(@as(f64, 1), coerceToNumber(.{ .text = "1" }, .excel, .text_coercion).number);

    // Non-numeric text is `#VALUE!` — a value, not a refusal.
    const bad = coerceToNumber(.{ .text = "abc" }, .excel, .text_coercion);
    try testing.expect(bad.value.eql(ScalarValue.errorOf(.value)));
    // Empty text likewise.
    try testing.expect(coerceToNumber(.{ .text = "" }, .excel, .text_coercion).value.eql(ScalarValue.errorOf(.value)));

    // Locale-flavoured text is a refusal, and specifically NOT `#VALUE!`.
    try testing.expect(coerceToNumber(.{ .text = "1,5" }, .excel, .text_coercion) == .locale_refusal);

    // Errors propagate unchanged, rich spellings included.
    const rich: ScalarValue = .{ .err = .{ .rich = "#BLOCKED!" } };
    try testing.expect(coerceToNumber(rich, .excel, .text_coercion).value.eql(rich));
}

// ─── error propagation order (§5.3c) ─────────────────────────────

test "§5.3c: operands evaluate left-to-right and the first error wins" {
    const left = ScalarValue.errorOf(.value);
    const right = ScalarValue.errorOf(.na);
    try testing.expect(propagateBinary(left, right).?.eql(.{ .known = .value }));
    try testing.expect(propagateBinary(ScalarValue.fromNumber(1), right).?.eql(.{ .known = .na }));
    try testing.expect(propagateBinary(ScalarValue.fromNumber(1), ScalarValue.fromNumber(2)) == null);
}

test "§5.3c: eager arguments propagate in declaration order" {
    const args = [_]ScalarValue{
        ScalarValue.fromNumber(1),
        .blank,
        ScalarValue.errorOf(.ref),
        ScalarValue.errorOf(.div0),
    };
    try testing.expect(firstError(&args).?.eql(.{ .known = .ref }));
    try testing.expect(firstError(&.{ScalarValue.fromNumber(1)}) == null);
    try testing.expect(firstError(&.{}) == null);
}

test "§5.3c: propagation classes are per function, never family-wide" {
    // The classes exist so COUNT and COUNTA can disagree inside one
    // family — COUNT ignores errors in ranges, COUNTA counts them.
    try testing.expect(PropagationClass.propagate != PropagationClass.per_function_provenance);
    try testing.expectEqual(@as(usize, 4), @typeInfo(PropagationClass).@"enum".fields.len);
}

// ─── the "Divergence ×2" gate (§5.4) ─────────────────────────────

const ProbeFn = *const fn (FpRules, []u8) ScalarValue;

fn probeLiteralIngress(rules: FpRules, _: []u8) ScalarValue {
    const digits = rules.literal_significant_digits orelse 17;
    return ScalarValue.fromNumber(roundToSignificantDigits(1.2345678901234567, digits));
}

fn probeCacheImport(_: FpRules, _: []u8) ScalarValue {
    // N1b is ingress-selected, not rules-selected, so both tables land
    // on the same value by construction.
    return ScalarValue.fromNumber(std.fmt.parseFloat(f64, "1.2345678901234567") catch unreachable);
}

fn probeAdditiveResidue(rules: FpRules, _: []u8) ScalarValue {
    const step = addSub(rules, 1.1, 1.0, .sub);
    return addSub(rules, step.number, 0.1, .sub);
}

fn probeNearZeroProduct(rules: FpRules, _: []u8) ScalarValue {
    return multiply(rules, 1e-20, 1e-20);
}

fn probeSignedZero(rules: FpRules, _: []u8) ScalarValue {
    const published = publish(ScalarValue.fromNumber(-0.0), if (rules.preserve_signed_zero) .ieee else .excel);
    return ScalarValue.fromNumber(published.number);
}

fn probeSubnormal(rules: FpRules, _: []u8) ScalarValue {
    return divide(rules, 1e-308, 1e10);
}

fn probeOverflow(rules: FpRules, _: []u8) ScalarValue {
    return multiply(rules, 1e308, 10);
}

fn probeDivZero(rules: FpRules, _: []u8) ScalarValue {
    return divide(rules, 1, 0);
}

fn probeSerialization(rules: FpRules, scratch: []u8) ScalarValue {
    _ = rules;
    return .{ .text = formatNumber(scratch, 0.30000000000000004) };
}

const divergence_probes = [_]ProbeFn{
    probeLiteralIngress,
    probeCacheImport,
    probeAdditiveResidue,
    probeNearZeroProduct,
    probeSignedZero,
    probeSubnormal,
    probeOverflow,
    probeDivZero,
    probeSerialization,
};

test "Divergence ×2: every §5.4 point runs under both rule tables" {
    // The gate asserts agreement AND disagreement where each is
    // required. A test that only looked for differences would pass if
    // the modes diverged everywhere, which is just as wrong as not
    // diverging at all.
    try testing.expectEqual(divergence_points.len, divergence_probes.len);

    var differed: usize = 0;
    var agreed: usize = 0;
    for (divergence_points, divergence_probes) |point, probe| {
        var buf_excel: [format_buf_len]u8 = undefined;
        var buf_ieee: [format_buf_len]u8 = undefined;
        const under_excel = probe(excel_fp_rules_v1, &buf_excel);
        const under_ieee = probe(ieee_fp_rules_v1, &buf_ieee);
        const same = under_excel.eql(under_ieee);

        switch (point.expect) {
            .must_differ => {
                if (same) {
                    std.debug.print("{s} ({s}): expected the modes to differ, both gave the same value\n", .{ point.rule, point.what });
                    return error.DivergenceMissing;
                }
                differed += 1;
            },
            .must_agree => {
                if (!same) {
                    std.debug.print("{s} ({s}): expected the modes to agree, they differed\n", .{ point.rule, point.what });
                    return error.UnexpectedDivergence;
                }
                agreed += 1;
            },
        }
    }
    // Both halves of the gate actually fired.
    try testing.expect(differed >= 3);
    try testing.expect(agreed >= 5);
}

test "Divergence ×2: the evidence label is honest about what the oracle decides" {
    var oracle_backed: usize = 0;
    for (divergence_points) |p| {
        if (p.evidence == .oracle) oracle_backed += 1;
    }
    // Subnormals, overflow, and division by zero are the three the
    // committed manifests decide. Everything else rests on §5.4's text,
    // and says so rather than implying oracle backing.
    try testing.expectEqual(@as(usize, 3), oracle_backed);
}

// ─── oracle ties (tests/oracle/fixtures) ─────────────────────────

const oracle_excel = @embedFile("oracle_hand_spec_excel");
const oracle_ieee = @embedFile("oracle_hand_spec_ieee");
const oracle_libreoffice = @embedFile("oracle_libreoffice_suite");

/// Compute a manifest formula with this module's primitives. Only the
/// value-layer cases: M3a1 has no evaluator, so anything needing
/// function dispatch or a cell reference is skipped rather than faked.
fn computeOracleCase(formula: []const u8, rules: FpRules) ?ScalarValue {
    if (std.mem.eql(u8, formula, "1+1")) return addSub(rules, 1, 1, .add);
    if (std.mem.eql(u8, formula, "0.1+0.2")) return addSub(rules, 0.1, 0.2, .add);
    if (std.mem.eql(u8, formula, "1/3")) return divide(rules, 1, 3);
    if (std.mem.eql(u8, formula, "1/0")) return divide(rules, 1, 0);
    if (std.mem.eql(u8, formula, "1E+308*10")) return multiply(rules, 1e308, 10);
    if (std.mem.eql(u8, formula, "1E-308/10000000000")) return divide(rules, 1e-308, 1e10);
    if (std.mem.eql(u8, formula, "1+1E-016")) return addSub(rules, 1, 1e-16, .add);
    if (std.mem.eql(u8, formula, "10%")) return divide(rules, 10, 100);
    if (std.mem.eql(u8, formula, "-0")) return ScalarValue.fromNumber(-0.0);
    if (std.mem.eql(u8, formula, "2^-1074")) return ScalarValue.fromArithmetic(std.math.pow(f64, 2, -1074));
    if (std.mem.eql(u8, formula, "2^3^2")) return ScalarValue.fromArithmetic(std.math.pow(f64, std.math.pow(f64, 2, 3), 2));
    if (std.mem.eql(u8, formula, "-1^2")) return ScalarValue.fromArithmetic(std.math.pow(f64, -1, 2));
    return null;
}

/// `-0` is the one row where the committed excel-fidelity manifest and
/// §5.4 disagree, and the disagreement is *about the adapter*, not
/// about zlsx: LibreOffice writes `-0` into its cached value, while
/// §5.4 says `.excel` normalizes the sign at publication. LibreOffice
/// is not Excel and the Excel leg is parked (M1b), so nothing on disk
/// can settle it. Named here rather than silently skipped.
const excel_adapter_divergences = [_][]const u8{"-0"};

fn isAdapterDivergent(formula: []const u8) bool {
    for (excel_adapter_divergences) |f| {
        if (std.mem.eql(u8, f, formula)) return true;
    }
    return false;
}

const OracleTie = struct { checked: usize, skipped_adapter: usize };

fn tieOracleManifest(json: []const u8) !OracleTie {
    const doc = try std.json.parseFromSlice(std.json.Value, testing.allocator, json, .{});
    defer doc.deinit();

    const fidelity_text = doc.value.object.get("fidelity").?.string;
    const exact = std.mem.eql(u8, fidelity_text, "ieee");
    const rules = if (exact) ieee_fp_rules_v1 else excel_fp_rules_v1;

    var tie: OracleTie = .{ .checked = 0, .skipped_adapter = 0 };
    for (doc.value.object.get("cells").?.array.items) |cell| {
        const obj = cell.object;
        const formula = (obj.get("formula") orelse continue).string;
        const computed = computeOracleCase(formula, rules) orelse continue;

        if (!exact and isAdapterDivergent(formula)) {
            tie.skipped_adapter += 1;
            continue;
        }

        const kind = (obj.get("kind") orelse continue).string;
        if (std.mem.eql(u8, kind, "error")) {
            const spelling = obj.get("error_spelling").?.string;
            if (computed != .err) return error.ExpectedError;
            try testing.expectEqualStrings(spelling, computed.err.spelling());
            tie.checked += 1;
            continue;
        }
        if (!std.mem.eql(u8, kind, "number")) continue;
        const bits_text = (obj.get("bits") orelse continue).string;
        const expected_bits = try std.fmt.parseInt(u64, bits_text[2..], 16);
        const expected: f64 = @bitCast(expected_bits);
        if (computed != .number) return error.ExpectedNumber;

        // `ieee` manifests record raw arithmetic, so they pin bits.
        // `excel` manifests carry §5.4 display-rounded values — the
        // LibreOffice adapter writes `0.3` for `0.1+0.2` — so they pin
        // to 15 significant digits, which is exactly what that rounding
        // preserves.
        const agrees = if (exact)
            expected_bits == bits(computed.number)
        else
            expected_bits == bits(computed.number) or
                @abs(computed.number - expected) <= 1e-15 * @abs(expected);

        if (!agrees) {
            std.debug.print(
                "oracle tie failed for `{s}` ({s}): recorded 0x{X:0>16}, computed 0x{X:0>16}\n",
                .{ formula, fidelity_text, expected_bits, bits(computed.number) },
            );
            return error.OracleTieMismatch;
        }
        tie.checked += 1;
    }
    return tie;
}

test "oracle: the rule tables reproduce every value-layer manifest cell" {
    const ieee = try tieOracleManifest(oracle_ieee);
    const excel = try tieOracleManifest(oracle_excel);
    const lo = try tieOracleManifest(oracle_libreoffice);

    // Guard against a vacuous pass.
    try testing.expect(ieee.checked >= 5);
    try testing.expect(excel.checked >= 4);
    try testing.expect(lo.checked >= 8);

    // The one adapter divergence is skipped deliberately, once per
    // excel-fidelity manifest that contains it.
    try testing.expectEqual(@as(usize, 0), ieee.skipped_adapter);
    try testing.expectEqual(@as(usize, 1), lo.skipped_adapter);
}

test "oracle: subnormals survive in both modes, bit for bit" {
    // The manifests agree on `2^-1074` and `1E-308/10000000000` across
    // fidelities, which is what makes N3's subnormal row oracle-backed
    // rather than assumed.
    inline for ([_]FpRules{ excel_fp_rules_v1, ieee_fp_rules_v1 }) |rules| {
        try testing.expectEqual(@as(u64, 0x00000000000316A2), bits(divide(rules, 1e-308, 1e10).number));
        try testing.expectEqual(@as(u64, 0x0000000000000001), bits(std.math.pow(f64, 2, -1074)));
    }
}

// ─── allocation failure ──────────────────────────────────────────

fn collateAndPublish(allocator: std.mem.Allocator, a: []const u8, b: []const u8) !void {
    const order = try collation_v1.compare(allocator, a, b);
    _ = order;
    var m = try Matrix.init(allocator, 2, 2);
    defer m.deinit(allocator);
    m.set(0, 0, .{ .text = a });
    _ = publish(m.at(0, 0), .excel);
}

test "allocation failure leaks nothing" {
    const pairs = [_][2][]const u8{
        .{ "Straße", "STRASSE" },
        .{ "\u{10400}", "\u{10428}" },
        .{ "abc", "ABD" },
        .{ "", "a" },
    };
    for (pairs) |p| {
        try testing.checkAllAllocationFailures(testing.allocator, collateAndPublish, .{ p[0], p[1] });
    }
}

// ─── fuzz ────────────────────────────────────────────────────────
//
// Contract: constructing and coercing a scalar must never panic and
// must never yield a non-finite number. Every path that can produce one
// — decimal ingress, arithmetic, coercion — is exercised against
// arbitrary bytes.

fn fuzzValueTarget(_: void, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();
    var smith_buf: [512]u8 = undefined;
    const input = smith_buf[0..smith.slice(&smith_buf)];

    inline for ([_]Fidelity{ .excel, .ieee }) |fid| {
        inline for (@typeInfo(Ingress).@"enum".fields) |f| {
            const ing: Ingress = @enumFromInt(f.value);
            if (parseDecimal(fid, ing, input)) |v| {
                // A parse that succeeds yields a real number. Infinity
                // is reachable by overflow and is N4a's business, but
                // NaN must never come out of the grammar.
                try std.testing.expect(!std.math.isNan(v));
                if (std.math.isFinite(v)) {
                    const s = ScalarValue.fromArithmetic(v);
                    try std.testing.expect(s == .number);
                    try std.testing.expect(std.math.isFinite(s.number));
                    var buf: [format_buf_len]u8 = undefined;
                    const text = formatNumber(&buf, s.number);
                    const back = std.fmt.parseFloat(f64, text) catch return;
                    try std.testing.expectEqual(bits(s.number), bits(back));
                }
            } else |_| {}
        }

        // Coercion over the same bytes as a text scalar.
        const scalar: ScalarValue = .{ .text = input };
        _ = provenanceOf(scalar);
        switch (coerceToNumber(scalar, fid, .text_coercion)) {
            .number => |n| try std.testing.expect(std.math.isFinite(n)),
            .value => |v| try std.testing.expect(v == .err),
            .locale_refusal => {},
        }
    }
}

test "fuzz: scalar construction and coercion never panic or go non-finite" {
    try std.testing.fuzz({}, fuzzValueTarget, .{
        .corpus = &[_][]const u8{
            "1",                "0",                   "-0",                 "1.5",      ".5",    "5.",
            "1e309",            "-1e309",              "1e-400",             "1,5",      "50%",   "$1.00",
            "abc",              "",                    " 1 ",                "1e",       "--1",   "1.2.3",
            "9007199254740993", "0.1",                 "1.2345678901234567", "1 234,56", "'''''", "\xFF\xFE",
            "1e2147483648",     "0000000000000000001",
        },
    });
}

fn fuzzArithmeticTarget(_: void, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();
    var raw: [16]u8 = undefined;
    const n = smith.slice(&raw);
    if (n < 16) return;

    const a: f64 = @bitCast(std.mem.readInt(u64, raw[0..8], .little));
    const b: f64 = @bitCast(std.mem.readInt(u64, raw[8..16], .little));
    // Non-finite operands are not values; the engine never holds one,
    // so the arithmetic contract is only defined over finite inputs.
    if (!std.math.isFinite(a) or !std.math.isFinite(b)) return;

    inline for ([_]FpRules{ excel_fp_rules_v1, ieee_fp_rules_v1 }) |rules| {
        for ([_]ScalarValue{
            addSub(rules, a, b, .add),
            addSub(rules, a, b, .sub),
            multiply(rules, a, b),
            divide(rules, a, b),
        }) |result| {
            switch (result) {
                .number => |v| try std.testing.expect(std.math.isFinite(v)),
                .err => |e| try std.testing.expect(e == .known),
                else => return error.UnexpectedResultKind,
            }
        }
    }
}

test "fuzz: arithmetic never yields a non-finite number in either mode" {
    try std.testing.fuzz({}, fuzzArithmeticTarget, .{
        .corpus = &[_][]const u8{
            &[_]u8{0} ** 16,
            &[_]u8{0xFF} ** 16,
            &[_]u8{ 0, 0, 0, 0, 0, 0, 0xF0, 0x7F, 0, 0, 0, 0, 0, 0, 0xF0, 0x7F },
        },
    });
}
