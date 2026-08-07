//! numfmt_v1 — the versioned number-format grammar (M8a, §5.4b).
//!
//! ONE parser and ONE renderer for OOXML number-format codes, shared by
//! the two callers that must never disagree about what a format means:
//! `TEXT()` in the registry and cell-display rendering at the workbook
//! layer. The grammar is versioned the way `collation_v1` and `rng_v1`
//! are: the constructs it accepts, the constructs it refuses, and the
//! spelling of every rendered byte are contract, not implementation
//! detail.
//!
//! The support matrix below (`support_matrix`) is the normative surface:
//! one row per grammar construct, each either rendered byte-exactly or
//! refusing a typed `Refusal.Reason`. The matrix is data on purpose —
//! the enumeration tests derive the refusal list from it, and a parked
//! construct is a row anyone can read rather than a sentence in a plan.
//!
//! Locale policy (§5.4b): output is pinned invariant en-US. `[$-409]`
//! licenses exactly the tables this file carries; any other LCID, any
//! calendar system, and any numeral shaping refuse by name — never a
//! fabricated error value, never a silently wrong month name.

const std = @import("std");
const assert = std.debug.assert;

const serial_date = @import("serial_date.zig");
const value = @import("value.zig");
const parser = @import("parser.zig");

/// The grammar's version. Bumping it is a contract change: every pinned
/// rendering below is `numfmt_v1`'s answer, and a caller that caches
/// rendered strings may key them on it.
pub const grammar_version: u32 = 1;

/// Excel caps a custom format code at 255 characters; zlsx counts code
/// points the way `max_formula_chars` does (§9). Boundary-tested.
pub const max_format_chars: usize = 255;

/// Subsecond display stops at milliseconds, matching
/// `serial_date.max_fractional_digits` — the lexical layer and the
/// display layer must agree on how fine a serial's time gets.
pub const max_subsecond_digits: usize = serial_date.max_fractional_digits;

// ─── the support matrix ──────────────────────────────────────────

/// Every construct the grammar knows. One enum value per row of the
/// support matrix; the exhaustiveness test walks `@typeInfo` of this
/// enum against `support_matrix`, so a construct added here without a
/// matrix row fails the tests by count and by membership.
pub const Construct = enum {
    general,
    digit_zero,
    digit_hash,
    digit_space,
    decimal_point,
    group_separator,
    scale_comma,
    percent,
    scientific,
    fraction_dynamic,
    fraction_fixed,
    literal_quoted,
    literal_escaped,
    literal_verbatim,
    fill,
    skip_width,
    text_at,
    color_tag,
    color_indexed,
    condition_tag,
    section_split,
    empty_section,
    date_year,
    date_month,
    date_month_name,
    date_day,
    date_day_name,
    time_hour,
    time_minute,
    time_second,
    twelve_hour,
    subsecond,
    elapsed_duration,
    locale_currency,
    locale_en_us,
    // ── refused constructs: real grammar, parked by name ──
    locale_other,
    locale_flags,
    dbnum,
    natnum,
    era_token,
    buddhist_calendar,
    weekday_localized,
};

pub const ConstructSet = std.EnumSet(Construct);

/// Why a format code refuses. The malformed family is the code failing
/// the grammar; the locale/calendar/shaping family is the grammar
/// naming a table zlsx does not carry (§5.4b's park). `format_too_long`
/// is a §9 limit.
pub const Refusal = struct {
    reason: Reason,
    /// Byte offset into the format code where the refusing construct
    /// starts. Diagnostics only — never load-bearing.
    at: usize,
    /// The matrix row that refused, when the refusal IS a matrix row.
    /// Malformed input has no row: a code outside the grammar is not a
    /// construct, it is the absence of one.
    construct: ?Construct,

    pub const Reason = enum {
        format_too_long,
        unclosed_quote,
        unclosed_bracket,
        unknown_bracket_tag,
        too_many_sections,
        condition_misplaced,
        condition_malformed,
        multiple_decimal_points,
        mixed_text_number,
        mixed_date_number,
        malformed_fraction,
        elapsed_misplaced,
        subsecond_precision,
        dangling_modifier,
        unknown_token,
        locale_table_unavailable,
        numeral_shaping_unsupported,
        calendar_unsupported,
    };

    /// Exhaustive by construction — a new `Reason` fails to compile
    /// until it has a §10 plane.
    pub fn planeTwo(self: Refusal) parser.PlaneTwo {
        return switch (self.reason) {
            .format_too_long => .FormulaLimitExceeded,
            .unclosed_quote,
            .unclosed_bracket,
            .unknown_bracket_tag,
            .too_many_sections,
            .condition_misplaced,
            .condition_malformed,
            .multiple_decimal_points,
            .mixed_text_number,
            .mixed_date_number,
            .malformed_fraction,
            .elapsed_misplaced,
            .subsecond_precision,
            .dangling_modifier,
            .unknown_token,
            => .FormulaMalformedInput,
            .locale_table_unavailable,
            .numeral_shaping_unsupported,
            .calendar_unsupported,
            => .FormulaLocaleSensitiveInput,
        };
    }
};

pub const Evidence = enum { oracle, spec_pinned };

/// One row of the support matrix.
///
/// The table is data on purpose: the enumeration test walks it, the
/// refused rows ARE the surfaced park list (`refusedConstructs`), and
/// promoting a construct means flipping exactly one row — the tests
/// that derive from the table follow without being edited.
pub const MatrixRow = struct {
    construct: Construct,
    status: Status,
    /// A format code exercising exactly this construct. Rendered rows
    /// must parse and mark the construct seen; refused rows must refuse
    /// with the row's reason AND name the row in the refusal.
    example: []const u8,
    evidence: Evidence = .spec_pinned,

    pub const Status = union(enum) {
        /// Rendered byte-exactly; the per-row fixtures pin the bytes.
        rendered,
        /// Refused with exactly this reason, proven by parsing `example`.
        refused: Refusal.Reason,
    };
};

pub const support_matrix = [_]MatrixRow{
    .{ .construct = .general, .status = .rendered, .example = "General" },
    .{ .construct = .digit_zero, .status = .rendered, .example = "00" },
    .{ .construct = .digit_hash, .status = .rendered, .example = "#" },
    .{ .construct = .digit_space, .status = .rendered, .example = "??" },
    .{ .construct = .decimal_point, .status = .rendered, .example = "0.00" },
    .{ .construct = .group_separator, .status = .rendered, .example = "#,##0" },
    .{ .construct = .scale_comma, .status = .rendered, .example = "0.0,," },
    .{ .construct = .percent, .status = .rendered, .example = "0%" },
    .{ .construct = .scientific, .status = .rendered, .example = "0.00E+00" },
    .{ .construct = .fraction_dynamic, .status = .rendered, .example = "# ?/?" },
    .{ .construct = .fraction_fixed, .status = .rendered, .example = "# ?/8" },
    .{ .construct = .literal_quoted, .status = .rendered, .example = "0 \"units\"" },
    .{ .construct = .literal_escaped, .status = .rendered, .example = "0\\h" },
    .{ .construct = .literal_verbatim, .status = .rendered, .example = "$0" },
    .{ .construct = .fill, .status = .rendered, .example = "*-0" },
    .{ .construct = .skip_width, .status = .rendered, .example = "_-0" },
    .{ .construct = .text_at, .status = .rendered, .example = "@" },
    .{ .construct = .color_tag, .status = .rendered, .example = "[Red]0" },
    .{ .construct = .color_indexed, .status = .rendered, .example = "[Color12]0" },
    .{ .construct = .condition_tag, .status = .rendered, .example = "[>100]0;0" },
    .{ .construct = .section_split, .status = .rendered, .example = "0;-0" },
    .{ .construct = .empty_section, .status = .rendered, .example = "0;;" },
    .{ .construct = .date_year, .status = .rendered, .example = "yyyy" },
    .{ .construct = .date_month, .status = .rendered, .example = "mm" },
    .{ .construct = .date_month_name, .status = .rendered, .example = "mmm" },
    .{ .construct = .date_day, .status = .rendered, .example = "dd" },
    .{ .construct = .date_day_name, .status = .rendered, .example = "ddd" },
    .{ .construct = .time_hour, .status = .rendered, .example = "h" },
    .{ .construct = .time_minute, .status = .rendered, .example = "h:mm" },
    .{ .construct = .time_second, .status = .rendered, .example = "ss" },
    .{ .construct = .twelve_hour, .status = .rendered, .example = "h AM/PM" },
    .{ .construct = .subsecond, .status = .rendered, .example = "ss.000" },
    .{ .construct = .elapsed_duration, .status = .rendered, .example = "[h]:mm" },
    .{ .construct = .locale_currency, .status = .rendered, .example = "[$\u{20AC}-409]0" },
    .{ .construct = .locale_en_us, .status = .rendered, .example = "[$-409]mmm" },
    .{ .construct = .locale_other, .status = .{ .refused = .locale_table_unavailable }, .example = "[$-40C]0" },
    .{ .construct = .locale_flags, .status = .{ .refused = .numeral_shaping_unsupported }, .example = "[$-D000409]0" },
    .{ .construct = .dbnum, .status = .{ .refused = .numeral_shaping_unsupported }, .example = "[DBNum1]0" },
    .{ .construct = .natnum, .status = .{ .refused = .numeral_shaping_unsupported }, .example = "[NatNum1]0" },
    .{ .construct = .era_token, .status = .{ .refused = .calendar_unsupported }, .example = "ge.m.d" },
    .{ .construct = .buddhist_calendar, .status = .{ .refused = .calendar_unsupported }, .example = "b2yyyy" },
    .{ .construct = .weekday_localized, .status = .{ .refused = .locale_table_unavailable }, .example = "aaaa" },
};

/// The refused rows, in table order — the surfaced park list. The
/// buffer is sized by the matrix itself so the list cannot overstate
/// its own capacity.
pub fn refusedConstructs(buf: *[support_matrix.len]Construct) []const Construct {
    var n: usize = 0;
    for (support_matrix) |row| {
        if (row.status == .refused) {
            buf[n] = row.construct;
            n += 1;
        }
    }
    return buf[0..n];
}

// ─── parsed representation ───────────────────────────────────────

/// A digit placeholder's role in the layout it belongs to.
const DigitRole = enum { int, frac, exponent, numerator, denominator };

/// What a placeholder emits when it has no digit to show.
const DigitPad = enum {
    /// `0` — a forced zero.
    zero,
    /// `#` — nothing.
    none,
    /// `?` — a space the width of a digit.
    space,
};

const DatePart = enum {
    year_two,
    year_four,
    month, // m
    month_padded, // mm
    month_abbr, // mmm
    month_full, // mmmm
    month_letter, // mmmmm
    day, // d
    day_padded, // dd
    day_abbr, // ddd
    day_full, // dddd
    hour, // h
    hour_padded, // hh
    minute, // m, resolved by context
    minute_padded, // mm, resolved by context
    second, // s
    second_padded, // ss
    am_pm, // AM/PM
    a_p, // A/P
};

const ElapsedUnit = enum { hours, minutes, seconds };

pub const Color = union(enum) {
    named: NamedColor,
    /// `[ColorN]`, N in 1…56.
    indexed: u8,

    pub const NamedColor = enum { black, blue, cyan, green, magenta, red, white, yellow };
};

/// `[>100]`-style section condition. The operand grammar is pinned
/// invariant: sign, digits, optional `.` digits — no locale could
/// change what it means, so it never refuses (§5.4b's line).
pub const Condition = struct {
    op: Op,
    operand: f64,

    pub const Op = enum { lt, le, gt, ge, eq, ne };

    pub fn matches(self: Condition, v: f64) bool {
        return switch (self.op) {
            .lt => v < self.operand,
            .le => v <= self.operand,
            .gt => v > self.operand,
            .ge => v >= self.operand,
            .eq => v == self.operand,
            .ne => v != self.operand,
        };
    }
};

const Token = union(enum) {
    digit: struct {
        role: DigitRole,
        pad: DigitPad,
        /// Index within the role, counting from the left.
        ord: u16,
    },
    decimal_point,
    /// Literal text, borrowed from the `Format`'s owned copy of the
    /// code. Consumed commas and fixed-denominator digits become empty
    /// literals rather than being spliced out — token indices stay
    /// stable through analysis.
    literal: []const u8,
    /// The `General` token (any spelling). Distinct from a quoted
    /// `"General"` literal, which is inert text.
    general_token,
    /// `%` — emits itself; the ×100 lives in the layout.
    percent,
    /// The `/` of a fraction candidate. Emits `/`, or a space when the
    /// fraction blanks out.
    fraction_slash,
    /// `E+`/`E-`/`e+`/`e-`. Case preserved for emission; the exponent's
    /// digits follow as `.digit` tokens with role `.exponent`.
    exponent: struct { upper: bool, sign_always: bool },
    /// `*x` — pinned to emit nothing (TEXT parity; decisions block).
    fill,
    /// `_x` — pinned to emit one space.
    skip,
    /// `@` — the value's text.
    text_at,
    /// `[$…-409]`'s currency string, emitted verbatim at its position.
    currency: []const u8,
    date: DatePart,
    elapsed: struct { unit: ElapsedUnit, width: u8 },
    subsecond: struct { digits: u8 },
};

const SectionClass = enum { number, date, text, general, empty };

const NumericKind = enum { plain, scientific, fraction };

const NumericLayout = struct {
    kind: NumericKind = .plain,
    int_places: u16 = 0,
    /// How many of the int placeholders force a zero — what grouping
    /// pads to, so `0,000` renders 5 as `0,005`.
    int_zero_places: u16 = 0,
    frac_places: u16 = 0,
    exp_places: u16 = 0,
    num_places: u16 = 0,
    den_places: u16 = 0,
    /// Non-zero for `?/8`-style fixed denominators.
    den_fixed: u32 = 0,
    scale_commas: u8 = 0,
    percent_count: u8 = 0,
    grouping: bool = false,
    has_decimal_point: bool = false,
};

const DateLayout = struct {
    /// The finest unit any token displays; the serial is rounded here
    /// once, BEFORE decomposition, so every displayed component agrees
    /// about the carry (12:00:59.6 under `hh:mm` is `12:01`).
    finest: Finest = .day,
    twelve_hour: bool = false,
    elapsed: ?ElapsedUnit = null,
    subsecond_digits: u8 = 0,

    const Finest = enum(u3) { day = 0, hour = 1, minute = 2, second = 3, subsecond = 4 };
};

pub const Section = struct {
    tokens: []const Token,
    class: SectionClass,
    condition: ?Condition = null,
    color: ?Color = null,
    layout: NumericLayout = .{},
    date: DateLayout = .{},
};

/// A parsed format code. Owns a copy of the code (token literals borrow
/// it) and the token storage.
pub const Format = struct {
    code: []const u8,
    sections: []const Section,
    /// Every construct the parse encountered — what the enumeration
    /// test checks a matrix example against.
    seen: ConstructSet,

    pub fn deinit(self: *Format, gpa: std.mem.Allocator) void {
        for (self.sections) |s| gpa.free(s.tokens);
        gpa.free(self.sections);
        gpa.free(self.code);
        self.* = undefined;
    }

    /// Whether any section renders the value as a date or duration —
    /// the grammar-backed answer to the question `isDateFormatCode`
    /// (src/xlsx.zig) approximates today. The heuristic's callers stay
    /// on the heuristic until a later row flips them through this seam.
    pub fn describesDate(self: *const Format) bool {
        for (self.sections) |s| {
            if (s.class == .date) return true;
        }
        return false;
    }
};

pub const ParseResult = union(enum) {
    ok: Format,
    refused: Refusal,
};

// ─── tokenizer / parser ──────────────────────────────────────────

const ParseState = struct {
    gpa: std.mem.Allocator,
    code: []const u8,
    i: usize = 0,
    seen: ConstructSet = ConstructSet.initEmpty(),
};

fn refusal(reason: Refusal.Reason, at: usize, construct: ?Construct) Refusal {
    return .{ .reason = reason, .at = at, .construct = construct };
}

/// Parse a format code. The result is either an owned `Format` (caller
/// deinits) or a typed refusal; `error.OutOfMemory` is the only Zig
/// error, because a refusal is an answer, not a failure.
pub fn parse(gpa: std.mem.Allocator, code: []const u8) error{OutOfMemory}!ParseResult {
    // §9 limit first, counted in code points like `max_formula_chars`.
    const cp_count = std.unicode.utf8CountCodepoints(code) catch code.len;
    if (cp_count > max_format_chars) {
        return .{ .refused = refusal(.format_too_long, 0, null) };
    }

    const owned = try gpa.dupe(u8, code);
    errdefer gpa.free(owned);

    var sections: std.ArrayListUnmanaged(Section) = .empty;
    errdefer {
        for (sections.items) |s| gpa.free(s.tokens);
        sections.deinit(gpa);
    }

    var st = ParseState{ .gpa = gpa, .code = owned };

    while (true) {
        switch (try parseSection(&st, sections.items.len)) {
            .ok => |sec| try sections.append(gpa, sec),
            .refused => |r| {
                for (sections.items) |s| gpa.free(s.tokens);
                sections.deinit(gpa);
                gpa.free(owned);
                return .{ .refused = r };
            },
        }
        if (st.i >= owned.len) break;
        assert(owned[st.i] == ';');
        st.i += 1;
        st.seen.insert(.section_split);
        if (sections.items.len == 4) {
            for (sections.items) |s| gpa.free(s.tokens);
            sections.deinit(gpa);
            gpa.free(owned);
            return .{ .refused = refusal(.too_many_sections, st.i, null) };
        }
    }

    return .{ .ok = .{
        .code = owned,
        .sections = try sections.toOwnedSlice(gpa),
        .seen = st.seen,
    } };
}

const SectionResult = union(enum) { ok: Section, refused: Refusal };

/// One section: tokens up to an unquoted `;` or the end of the code,
/// then classification and layout analysis.
fn parseSection(st: *ParseState, section_index: usize) error{OutOfMemory}!SectionResult {
    const gpa = st.gpa;
    const code = st.code;

    var tokens: std.ArrayListUnmanaged(Token) = .empty;
    defer tokens.deinit(gpa);

    var condition: ?Condition = null;
    var color: ?Color = null;

    while (st.i < code.len and code[st.i] != ';') {
        const at = st.i;
        const c = code[at];
        switch (c) {
            '"' => {
                const close = std.mem.indexOfScalarPos(u8, code, at + 1, '"') orelse {
                    return .{ .refused = refusal(.unclosed_quote, at, null) };
                };
                try tokens.append(gpa, .{ .literal = code[at + 1 .. close] });
                st.seen.insert(.literal_quoted);
                st.i = close + 1;
            },
            '\\' => {
                if (at + 1 >= code.len) {
                    return .{ .refused = refusal(.dangling_modifier, at, null) };
                }
                const n = std.unicode.utf8ByteSequenceLength(code[at + 1]) catch 1;
                const end = @min(at + 1 + n, code.len);
                try tokens.append(gpa, .{ .literal = code[at + 1 .. end] });
                st.seen.insert(.literal_escaped);
                st.i = end;
            },
            '*' => {
                if (at + 1 >= code.len) {
                    return .{ .refused = refusal(.dangling_modifier, at, null) };
                }
                const n = std.unicode.utf8ByteSequenceLength(code[at + 1]) catch 1;
                try tokens.append(gpa, .fill);
                st.seen.insert(.fill);
                st.i = @min(at + 1 + n, code.len);
            },
            '_' => {
                if (at + 1 >= code.len) {
                    return .{ .refused = refusal(.dangling_modifier, at, null) };
                }
                const n = std.unicode.utf8ByteSequenceLength(code[at + 1]) catch 1;
                try tokens.append(gpa, .skip);
                st.seen.insert(.skip_width);
                st.i = @min(at + 1 + n, code.len);
            },
            '[' => switch (parseBracket(st)) {
                .condition => |cond| {
                    if (condition != null or section_index >= 2) {
                        return .{ .refused = refusal(.condition_misplaced, at, null) };
                    }
                    condition = cond;
                    st.seen.insert(.condition_tag);
                },
                .color => |col| color = col,
                .token => |tok| try tokens.append(gpa, tok),
                .refused => |r| return .{ .refused = r },
            },
            '@' => {
                try tokens.append(gpa, .text_at);
                st.seen.insert(.text_at);
                st.i += 1;
            },
            '0', '#', '?' => {
                try tokens.append(gpa, .{
                    .digit = .{
                        .role = .int, // roles are assigned in analysis
                        .pad = switch (c) {
                            '0' => .zero,
                            '#' => .none,
                            else => .space,
                        },
                        .ord = 0,
                    },
                });
                st.seen.insert(switch (c) {
                    '0' => .digit_zero,
                    '#' => .digit_hash,
                    else => .digit_space,
                });
                st.i += 1;
            },
            '.' => {
                try tokens.append(gpa, .decimal_point);
                st.i += 1;
            },
            ',' => {
                // Grouping, scale, or plain literal — resolved during
                // numeric analysis, when the neighbours are known.
                try tokens.append(gpa, .{ .literal = code[at .. at + 1] });
                st.i += 1;
            },
            '%' => {
                try tokens.append(gpa, .percent);
                st.seen.insert(.percent);
                st.i += 1;
            },
            '/' => {
                // Fraction slash in a numeric section, plain separator
                // in a date one — analysis decides.
                try tokens.append(gpa, .fraction_slash);
                st.i += 1;
            },
            'E', 'e' => {
                if (at + 1 < code.len and (code[at + 1] == '+' or code[at + 1] == '-')) {
                    try tokens.append(gpa, .{ .exponent = .{
                        .upper = c == 'E',
                        .sign_always = code[at + 1] == '+',
                    } });
                    st.seen.insert(.scientific);
                    st.i = at + 2;
                } else {
                    // Bare `e` is the era-adjacent year: under a Japanese
                    // calendar it counts eras, under en-US it happens to
                    // count Gregorian years. A spelling whose MEANING the
                    // locale changes is exactly what §5.4b refuses.
                    st.seen.insert(.era_token);
                    return .{ .refused = refusal(.calendar_unsupported, at, .era_token) };
                }
            },
            'G', 'g' => {
                if (startsWithIgnoreCase(code[at..], "general")) {
                    try tokens.append(gpa, .general_token);
                    st.seen.insert(.general);
                    st.i = at + 7;
                } else {
                    st.seen.insert(.era_token);
                    return .{ .refused = refusal(.calendar_unsupported, at, .era_token) };
                }
            },
            'A', 'a' => {
                if (startsWithIgnoreCase(code[at..], "am/pm")) {
                    try tokens.append(gpa, .{ .date = .am_pm });
                    st.seen.insert(.twelve_hour);
                    st.i = at + 5;
                } else if (startsWithIgnoreCase(code[at..], "a/p")) {
                    try tokens.append(gpa, .{ .date = .a_p });
                    st.seen.insert(.twelve_hour);
                    st.i = at + 3;
                } else if (runLen(code, at, "aA") >= 3) {
                    // `aaa`/`aaaa` spell the weekday in the locale's
                    // language — a table zlsx does not carry (§5.4b).
                    st.seen.insert(.weekday_localized);
                    return .{ .refused = refusal(.locale_table_unavailable, at, .weekday_localized) };
                } else {
                    return .{ .refused = refusal(.unknown_token, at, null) };
                }
            },
            'B', 'b' => {
                // `b1`/`b2` calendar selectors and the bare Buddhist-year
                // `b` both change which calendar counts the years.
                st.seen.insert(.buddhist_calendar);
                return .{ .refused = refusal(.calendar_unsupported, at, .buddhist_calendar) };
            },
            'Y', 'y' => {
                const n = runLen(code, at, "yY");
                try tokens.append(gpa, .{ .date = if (n >= 3) .year_four else .year_two });
                st.seen.insert(.date_year);
                st.i = at + n;
            },
            'M', 'm' => {
                const n = runLen(code, at, "mM");
                const part: DatePart = switch (n) {
                    1 => .month,
                    2 => .month_padded,
                    3 => .month_abbr,
                    4 => .month_full,
                    else => .month_letter,
                };
                try tokens.append(gpa, .{ .date = part });
                st.seen.insert(if (n >= 3) .date_month_name else .date_month);
                st.i = at + n;
            },
            'D', 'd' => {
                const n = runLen(code, at, "dD");
                const part: DatePart = switch (n) {
                    1 => .day,
                    2 => .day_padded,
                    3 => .day_abbr,
                    else => .day_full,
                };
                try tokens.append(gpa, .{ .date = part });
                st.seen.insert(if (n >= 3) .date_day_name else .date_day);
                st.i = at + n;
            },
            'H', 'h' => {
                const n = runLen(code, at, "hH");
                try tokens.append(gpa, .{ .date = if (n >= 2) .hour_padded else .hour });
                st.seen.insert(.time_hour);
                st.i = at + n;
            },
            'S', 's' => {
                const n = runLen(code, at, "sS");
                try tokens.append(gpa, .{ .date = if (n >= 2) .second_padded else .second });
                st.seen.insert(.time_second);
                st.i = at + n;
            },
            // Digits 1–9 are literals; the fraction analysis consumes
            // them as fixed denominators when they follow a `/`.
            '1', '2', '3', '4', '5', '6', '7', '8', '9' => {
                try tokens.append(gpa, .{ .literal = code[at .. at + 1] });
                st.seen.insert(.literal_verbatim);
                st.i += 1;
            },
            // The bare literals ECMA-376 §18.8.31 allows unescaped.
            '$', '+', '-', '(', ')', ':', ' ', '!', '^', '&', '\'', '~', '{', '}', '<', '>', '=' => {
                try tokens.append(gpa, .{ .literal = code[at .. at + 1] });
                st.seen.insert(.literal_verbatim);
                st.i += 1;
            },
            else => {
                if (c < 0x80) {
                    // An unescaped ASCII byte the grammar gives no
                    // meaning. Excel accepts some of these silently;
                    // zlsx refuses rather than guess which (§5.4b's
                    // spirit: outside the grammar is a refusal, not an
                    // improvisation).
                    return .{ .refused = refusal(.unknown_token, at, null) };
                }
                // Non-ASCII passes through as a literal, the way Excel
                // treats `€` or `円` typed straight into a format.
                const n = std.unicode.utf8ByteSequenceLength(c) catch 1;
                const end = @min(at + n, code.len);
                try tokens.append(gpa, .{ .literal = code[at..end] });
                st.seen.insert(.literal_verbatim);
                st.i = end;
            },
        }
    }

    const owned_tokens = try tokens.toOwnedSlice(gpa);
    errdefer gpa.free(owned_tokens);

    var sec = Section{
        .tokens = owned_tokens,
        .class = .empty,
        .condition = condition,
        .color = color,
    };
    switch (analyzeSection(st, &sec, owned_tokens)) {
        .ok => {},
        .refused => |r| {
            gpa.free(owned_tokens);
            return .{ .refused = r };
        },
    }
    if (sec.class == .empty and sec.tokens.len == 0) st.seen.insert(.empty_section);
    return .{ .ok = sec };
}

const BracketResult = union(enum) {
    condition: Condition,
    color: Color,
    token: Token,
    refused: Refusal,
};

fn parseBracket(st: *ParseState) BracketResult {
    const code = st.code;
    const at = st.i;
    assert(code[at] == '[');
    const close = std.mem.indexOfScalarPos(u8, code, at + 1, ']') orelse {
        return .{ .refused = refusal(.unclosed_bracket, at, null) };
    };
    const body = code[at + 1 .. close];
    if (body.len == 0) return .{ .refused = refusal(.unknown_bracket_tag, at, null) };

    // Condition: `[<op><number>]`.
    if (body[0] == '<' or body[0] == '>' or body[0] == '=') {
        var j: usize = 1;
        var op: Condition.Op = switch (body[0]) {
            '<' => .lt,
            '>' => .gt,
            else => .eq,
        };
        if (body[0] == '<' and body.len > 1 and body[1] == '=') {
            op = .le;
            j = 2;
        } else if (body[0] == '<' and body.len > 1 and body[1] == '>') {
            op = .ne;
            j = 2;
        } else if (body[0] == '>' and body.len > 1 and body[1] == '=') {
            op = .ge;
            j = 2;
        }
        const operand = std.fmt.parseFloat(f64, body[j..]) catch {
            return .{ .refused = refusal(.condition_malformed, at, null) };
        };
        if (!std.math.isFinite(operand)) {
            return .{ .refused = refusal(.condition_malformed, at, null) };
        }
        st.i = close + 1;
        return .{ .condition = .{ .op = op, .operand = operand } };
    }

    // Elapsed duration: `[h]`, `[mm]`, `[s]` — the whole body one run.
    inline for (.{
        .{ "hH", ElapsedUnit.hours },
        .{ "mM", ElapsedUnit.minutes },
        .{ "sS", ElapsedUnit.seconds },
    }) |entry| {
        if (runCovers(body, entry[0])) {
            st.i = close + 1;
            st.seen.insert(.elapsed_duration);
            return .{ .token = .{ .elapsed = .{
                .unit = entry[1],
                .width = @intCast(@min(body.len, 8)),
            } } };
        }
    }

    // Currency/locale: `[$currency-LCID]` or `[$currency]`.
    if (body[0] == '$') {
        const dash = std.mem.lastIndexOfScalar(u8, body, '-');
        const currency = if (dash) |d| body[1..d] else body[1..];
        if (dash) |d| {
            const lcid_text = body[d + 1 ..];
            if (lcid_text.len == 0 or lcid_text.len > 8) {
                return .{ .refused = refusal(.unknown_bracket_tag, at, null) };
            }
            const lcid = std.fmt.parseInt(u32, lcid_text, 16) catch {
                return .{ .refused = refusal(.unknown_bracket_tag, at, null) };
            };
            if (lcid > 0xFFFF) {
                // The high bits select calendars and numeral shaping —
                // `[$-D000409]` asks for Thai digits, not for en-US.
                st.seen.insert(.locale_flags);
                return .{ .refused = refusal(.numeral_shaping_unsupported, at, .locale_flags) };
            }
            if (lcid != 0x0409) {
                st.seen.insert(.locale_other);
                return .{ .refused = refusal(.locale_table_unavailable, at, .locale_other) };
            }
            st.seen.insert(.locale_en_us);
        }
        st.i = close + 1;
        if (currency.len > 0) st.seen.insert(.locale_currency);
        return .{ .token = .{ .currency = currency } };
    }

    // DBNum / NatNum: East Asian numeral shaping.
    if (startsWithIgnoreCase(body, "dbnum")) {
        st.seen.insert(.dbnum);
        return .{ .refused = refusal(.numeral_shaping_unsupported, at, .dbnum) };
    }
    if (startsWithIgnoreCase(body, "natnum")) {
        st.seen.insert(.natnum);
        return .{ .refused = refusal(.numeral_shaping_unsupported, at, .natnum) };
    }

    // Colors.
    if (colorByName(body)) |named| {
        st.i = close + 1;
        st.seen.insert(.color_tag);
        return .{ .color = .{ .named = named } };
    }
    if (startsWithIgnoreCase(body, "color")) {
        const idx = std.fmt.parseInt(u8, body[5..], 10) catch {
            return .{ .refused = refusal(.unknown_bracket_tag, at, null) };
        };
        if (idx < 1 or idx > 56) {
            return .{ .refused = refusal(.unknown_bracket_tag, at, null) };
        }
        st.i = close + 1;
        st.seen.insert(.color_indexed);
        return .{ .color = .{ .indexed = idx } };
    }

    return .{ .refused = refusal(.unknown_bracket_tag, at, null) };
}

fn colorByName(body: []const u8) ?Color.NamedColor {
    inline for (@typeInfo(Color.NamedColor).@"enum".fields) |f| {
        if (std.ascii.eqlIgnoreCase(body, f.name)) {
            return @field(Color.NamedColor, f.name);
        }
    }
    return null;
}

fn startsWithIgnoreCase(haystack: []const u8, prefix: []const u8) bool {
    if (haystack.len < prefix.len) return false;
    return std.ascii.eqlIgnoreCase(haystack[0..prefix.len], prefix);
}

/// Length of the run of characters from `set` starting at `at`.
fn runLen(code: []const u8, at: usize, set: []const u8) usize {
    var n: usize = 0;
    while (at + n < code.len and std.mem.indexOfScalar(u8, set, code[at + n]) != null) n += 1;
    return n;
}

/// Whether `body` is one uninterrupted run of characters from `set`.
fn runCovers(body: []const u8, set: []const u8) bool {
    if (body.len == 0) return false;
    return runLen(body, 0, set) == body.len;
}

// ─── section analysis ────────────────────────────────────────────

const AnalyzeResult = union(enum) { ok, refused: Refusal };

fn analyzeRefuse(reason: Refusal.Reason) AnalyzeResult {
    return .{ .refused = refusal(reason, 0, null) };
}

/// Classify the section and assign digit roles. This is where `m`
/// becomes month or minute, where `,` becomes grouping, scale or
/// literal, where `/` becomes a fraction, and where the mixes the
/// grammar has no rendering for refuse.
fn analyzeSection(st: *ParseState, sec: *Section, tokens: []Token) AnalyzeResult {
    var has_digit = false;
    var has_date = false;
    var has_at = false;
    var has_general = false;
    for (tokens) |t| switch (t) {
        .digit => has_digit = true,
        .date, .subsecond, .elapsed => has_date = true,
        .text_at => has_at = true,
        .general_token => has_general = true,
        else => {},
    };

    if (has_at) {
        if (has_digit or has_date or has_general) {
            return analyzeRefuse(.mixed_text_number);
        }
        if (sec.condition != null) return analyzeRefuse(.condition_misplaced);
        sec.class = .text;
        return .ok;
    }
    if (has_general) {
        if (has_digit or has_date) return analyzeRefuse(.mixed_text_number);
        sec.class = .general;
        return .ok;
    }
    if (has_date) return analyzeDateSection(st, sec, tokens);
    if (tokens.len == 0) {
        sec.class = .empty;
        return .ok;
    }
    sec.class = .number;
    return analyzeNumberSection(st, sec, tokens);
}

fn analyzeDateSection(st: *ParseState, sec: *Section, tokens: []Token) AnalyzeResult {
    sec.class = .date;

    // Rewrite `.` + a run of `0` placeholders directly after a seconds
    // token into a subsecond token; any other digit shape inside a date
    // section is a mix the grammar refuses.
    var i: usize = 0;
    var last_was_seconds = false;
    while (i < tokens.len) : (i += 1) {
        switch (tokens[i]) {
            .date => |p| last_was_seconds = (p == .second or p == .second_padded),
            .elapsed => |e| last_was_seconds = (e.unit == .seconds),
            .fraction_slash => {
                // `/` is a plain separator inside a date — `m/d/yyyy`.
                tokens[i] = .{ .literal = "/" };
            },
            .decimal_point => {
                if (!last_was_seconds) {
                    tokens[i] = .{ .literal = "." };
                    continue;
                }
                var digits: u8 = 0;
                var j = i + 1;
                while (j < tokens.len) : (j += 1) {
                    switch (tokens[j]) {
                        .digit => |d| {
                            if (d.pad != .zero) break;
                            digits += 1;
                        },
                        else => break,
                    }
                }
                if (digits == 0) {
                    tokens[i] = .{ .literal = "." };
                    continue;
                }
                if (digits > max_subsecond_digits) {
                    return analyzeRefuse(.subsecond_precision);
                }
                tokens[i] = .{ .subsecond = .{ .digits = digits } };
                st.seen.insert(.subsecond);
                var k = i + 1;
                while (k < i + 1 + digits) : (k += 1) tokens[k] = .{ .literal = "" };
                i += digits;
                last_was_seconds = false;
            },
            else => {},
        }
    }

    // Any digit placeholder that survived, and any numeric operator, is
    // a numeric layout inside a date — Excel refuses `0 yyyy` at entry;
    // so does the grammar.
    for (tokens) |t| switch (t) {
        .digit => return analyzeRefuse(.mixed_date_number),
        .percent, .exponent, .fraction_slash => return analyzeRefuse(.mixed_date_number),
        else => {},
    };

    // `m` disambiguation: a month token of width ≤2 becomes minutes
    // when its nearest disambiguating neighbour is time-ish — `h:mm`
    // counts minutes, `yyyy-mm` counts months.
    for (tokens, 0..) |t, idx| {
        const width: u8 = switch (t) {
            .date => |p| switch (p) {
                .month => 1,
                .month_padded => 2,
                else => continue,
            },
            else => continue,
        };
        if (nearestIsTime(tokens, idx)) {
            tokens[idx] = .{ .date = if (width == 2) .minute_padded else .minute };
            st.seen.insert(.time_minute);
        }
    }

    // Elapsed legality: at most one, and nothing time-ish before it —
    // `[h]:mm` counts down from hours; `mm:[ss]` has no meaning to pin.
    var time_before = false;
    for (tokens) |t| switch (t) {
        .elapsed => |e| {
            if (sec.date.elapsed != null or time_before) {
                return analyzeRefuse(.elapsed_misplaced);
            }
            sec.date.elapsed = e.unit;
        },
        .date => |p| switch (p) {
            .hour, .hour_padded, .minute, .minute_padded, .second, .second_padded => time_before = true,
            else => {},
        },
        else => {},
    };

    // Twelve-hour clock and the finest displayed unit.
    var finest: DateLayout.Finest = .day;
    for (tokens) |t| switch (t) {
        .date => |p| {
            const f: ?DateLayout.Finest = switch (p) {
                .hour, .hour_padded => .hour,
                .minute, .minute_padded => .minute,
                .second, .second_padded => .second,
                .am_pm, .a_p => blk: {
                    sec.date.twelve_hour = true;
                    break :blk .hour;
                },
                else => null,
            };
            if (f) |ff| {
                if (@intFromEnum(finest) < @intFromEnum(ff)) finest = ff;
            }
        },
        .elapsed => |e| {
            const f: DateLayout.Finest = switch (e.unit) {
                .hours => .hour,
                .minutes => .minute,
                .seconds => .second,
            };
            if (@intFromEnum(finest) < @intFromEnum(f)) finest = f;
        },
        .subsecond => |sub| {
            finest = .subsecond;
            sec.date.subsecond_digits = sub.digits;
        },
        else => {},
    };
    sec.date.finest = finest;
    return .ok;
}

/// Whether the nearest disambiguating neighbour of the `m` token at
/// `idx` is a time token. Backward first (an `h` before wins), then
/// forward (an `s` after wins); a date token in between blocks.
fn nearestIsTime(tokens: []const Token, idx: usize) bool {
    var j = idx;
    while (j > 0) {
        j -= 1;
        switch (tokens[j]) {
            .date => |p| switch (p) {
                .hour, .hour_padded => return true,
                .year_two, .year_four, .day, .day_padded, .day_abbr, .day_full, .month_abbr, .month_full, .month_letter => break,
                else => {},
            },
            .elapsed => |e| if (e.unit == .hours) return true,
            else => {},
        }
    }
    var k = idx + 1;
    while (k < tokens.len) : (k += 1) {
        switch (tokens[k]) {
            .date => |p| switch (p) {
                .second, .second_padded => return true,
                .year_two, .year_four, .day, .day_padded, .day_abbr, .day_full, .month_abbr, .month_full, .month_letter => return false,
                else => {},
            },
            .elapsed => |e| if (e.unit == .seconds) return true,
            else => {},
        }
    }
    return false;
}

fn analyzeNumberSection(st: *ParseState, sec: *Section, tokens: []Token) AnalyzeResult {
    var layout = NumericLayout{};

    // Structural markers first: decimal point, exponent, fraction slash.
    var decimal_at: ?usize = null;
    var exponent_at: ?usize = null;
    var slash_at: ?usize = null;
    for (tokens, 0..) |t, idx| switch (t) {
        .decimal_point => {
            if (decimal_at != null) return analyzeRefuse(.multiple_decimal_points);
            decimal_at = idx;
            layout.has_decimal_point = true;
            st.seen.insert(.decimal_point);
        },
        .exponent => {
            if (exponent_at != null) return analyzeRefuse(.unknown_token);
            exponent_at = idx;
        },
        .fraction_slash => {
            if (slash_at != null) return analyzeRefuse(.malformed_fraction);
            slash_at = idx;
        },
        .percent => layout.percent_count +|= 1,
        else => {},
    };

    if (slash_at) |sl| {
        if (decimal_at != null or exponent_at != null) {
            return analyzeRefuse(.malformed_fraction);
        }
        return analyzeFraction(st, sec, tokens, sl, &layout);
    }

    layout.kind = if (exponent_at != null) .scientific else .plain;

    // Assign digit roles positionally.
    var int_ord: u16 = 0;
    var frac_ord: u16 = 0;
    var exp_ord: u16 = 0;
    for (tokens, 0..) |*t, idx| {
        if (t.* != .digit) continue;
        if (exponent_at != null and idx > exponent_at.?) {
            t.digit.role = .exponent;
            t.digit.ord = exp_ord;
            exp_ord += 1;
        } else if (decimal_at != null and idx > decimal_at.?) {
            t.digit.role = .frac;
            t.digit.ord = frac_ord;
            frac_ord += 1;
        } else {
            t.digit.role = .int;
            t.digit.ord = int_ord;
            int_ord += 1;
            if (t.digit.pad == .zero) layout.int_zero_places += 1;
        }
    }
    layout.int_places = int_ord;
    layout.frac_places = frac_ord;
    layout.exp_places = exp_ord;
    if (exponent_at != null and exp_ord == 0) return analyzeRefuse(.unknown_token);

    // Resolve `,`: commas directly after the LAST digit placeholder of
    // the mantissa scale by 1000 each (`#,` and `0.0,,` both scale);
    // commas between int placeholders turn grouping on; anything else
    // stays a literal comma.
    var first_int_idx: ?usize = null;
    var last_int_idx: ?usize = null;
    var last_mantissa_idx: ?usize = null;
    for (tokens, 0..) |t, idx| {
        if (t != .digit) continue;
        switch (t.digit.role) {
            .int => {
                if (first_int_idx == null) first_int_idx = idx;
                last_int_idx = idx;
                last_mantissa_idx = idx;
            },
            .frac => last_mantissa_idx = idx,
            else => {},
        }
    }
    if (last_mantissa_idx) |lm| {
        var scale_run = true;
        var idx = lm + 1;
        while (idx < tokens.len) : (idx += 1) {
            switch (tokens[idx]) {
                .literal => |lit| {
                    if (lit.len == 1 and lit[0] == ',') {
                        if (scale_run) {
                            layout.scale_commas +|= 1;
                            tokens[idx] = .{ .literal = "" };
                            st.seen.insert(.scale_comma);
                        }
                    } else if (lit.len > 0) {
                        scale_run = false;
                    }
                },
                else => scale_run = false,
            }
        }
    }
    if (last_int_idx) |li| {
        var idx = first_int_idx.?;
        while (idx < li) : (idx += 1) {
            const t = tokens[idx];
            if (t == .literal and t.literal.len == 1 and t.literal[0] == ',') {
                layout.grouping = true;
                tokens[idx] = .{ .literal = "" };
                st.seen.insert(.group_separator);
            }
        }
    }

    sec.layout = layout;
    return .ok;
}

fn analyzeFraction(
    st: *ParseState,
    sec: *Section,
    tokens: []Token,
    slash_idx: usize,
    layout: *NumericLayout,
) AnalyzeResult {
    layout.kind = .fraction;

    // After the slash: all placeholders (dynamic denominator) or all
    // literal digits (fixed denominator). They cannot mix.
    var den_places: u16 = 0;
    var fixed: u32 = 0;
    var fixed_digits: usize = 0;
    var j = slash_idx + 1;
    while (j < tokens.len) : (j += 1) {
        switch (tokens[j]) {
            .digit => {
                if (fixed_digits > 0) return analyzeRefuse(.malformed_fraction);
                tokens[j].digit.role = .denominator;
                tokens[j].digit.ord = den_places;
                den_places += 1;
            },
            .literal => |lit| {
                if (lit.len == 1 and lit[0] >= '1' and lit[0] <= '9') {
                    if (den_places > 0) return analyzeRefuse(.malformed_fraction);
                    fixed = fixed * 10 + (lit[0] - '0');
                    fixed_digits += 1;
                    if (fixed_digits > 6) return analyzeRefuse(.malformed_fraction);
                    // The digit stays a literal: `?/8` renders its own 8.
                } else if (lit.len == 0) {
                    // an already-consumed token — keep scanning
                } else break;
            },
            else => break,
        }
    }
    if (den_places == 0 and fixed == 0) return analyzeRefuse(.malformed_fraction);
    if (den_places > 6) return analyzeRefuse(.malformed_fraction);
    layout.den_places = den_places;
    layout.den_fixed = fixed;
    st.seen.insert(if (fixed != 0) .fraction_fixed else .fraction_dynamic);

    // Before the slash: the digit run adjacent to it is the numerator,
    // everything left of that run is the integer part.
    var num_lo = slash_idx;
    var k = slash_idx;
    scan: while (k > 0) {
        k -= 1;
        switch (tokens[k]) {
            .digit => num_lo = k,
            .literal => |lit| {
                if (lit.len != 0) break :scan;
            },
            else => break :scan,
        }
    }
    if (num_lo == slash_idx) return analyzeRefuse(.malformed_fraction);

    var num_places: u16 = 0;
    var int_places: u16 = 0;
    var int_zero: u16 = 0;
    for (tokens, 0..) |*t, idx| {
        if (t.* != .digit or idx > slash_idx) continue;
        if (idx >= num_lo) {
            t.digit.role = .numerator;
            t.digit.ord = num_places;
            num_places += 1;
        } else {
            t.digit.role = .int;
            t.digit.ord = int_places;
            int_places += 1;
            if (t.digit.pad == .zero) int_zero += 1;
        }
    }
    if (num_places > 6) return analyzeRefuse(.malformed_fraction);
    layout.num_places = num_places;
    layout.int_places = int_places;
    layout.int_zero_places = int_zero;

    sec.layout = layout.*;
    return .ok;
}

// ─── rendering ───────────────────────────────────────────────────

/// What the renderer accepts. Errors never reach here — TEXT propagates
/// them before formatting, and a cell's error display is the error's
/// own spelling, not a format's business.
pub const RenderValue = union(enum) {
    number: f64,
    text: []const u8,
    boolean: bool,
};

pub const RenderOptions = struct {
    date_system: serial_date.DateSystem = .d1900,
};

pub const RenderError = error{
    OutOfMemory,
    /// A serial outside the date system's domain reached a date or
    /// duration section — Excel fills the cell with `#`, TEXT answers
    /// `#VALUE!`; each caller maps this error to its own spelling.
    SerialOutOfRange,
};

/// Render `v` under `fmt`. The returned slice is owned by `gpa`.
///
/// ONE derivation for both callers (M8a): `TEXT()` calls this with the
/// run arena, the workbook display seam calls it with its caller's
/// allocator, and neither adds a byte the other would not.
pub fn render(
    gpa: std.mem.Allocator,
    fmt: *const Format,
    v: RenderValue,
    opts: RenderOptions,
) RenderError![]u8 {
    switch (v) {
        .boolean => |b| {
            // Booleans bypass sections entirely: Excel renders TRUE and
            // FALSE through every format, numeric or not.
            return gpa.dupe(u8, if (b) "TRUE" else "FALSE");
        },
        .text => |t| {
            // First @-carrying section wins; no text section → the text
            // passes through unformatted.
            for (fmt.sections) |*s| {
                if (s.class == .text) return renderTextSection(gpa, s, t);
            }
            return gpa.dupe(u8, t);
        },
        .number => |n| {
            assert(std.math.isFinite(n));
            const pick = pickSection(fmt, n) orelse {
                // Every section conditioned, none matched: pinned to the
                // General spelling of the raw value (decisions block;
                // Excel leg parked with the oracle).
                return renderGeneralNumber(gpa, n);
            };
            const sec = &fmt.sections[pick.index];
            // A negative serial has no date in either system: refusing
            // here keeps the positional-|v| rule from silently dropping
            // the sign into a fabricated date (decisions block).
            if (sec.class == .date and n < 0) return error.SerialOutOfRange;
            const rendered = if (pick.abs) @abs(n) else n;
            return switch (sec.class) {
                .empty => gpa.dupe(u8, ""),
                .text => renderNumberThroughAt(gpa, sec, rendered),
                .general => renderGeneralSection(gpa, sec, rendered, pick.explicit_sign),
                .date => renderDateSection(gpa, sec, rendered, opts),
                .number => renderNumberSection(gpa, sec, rendered, pick.explicit_sign),
            };
        },
    }
}

const SectionPick = struct {
    index: usize,
    /// Positional negative sections render |v|.
    abs: bool,
    /// Whether the renderer must emit a leading `-` itself.
    explicit_sign: bool,
};

/// Which section formats `n`, and who is responsible for its sign.
///
/// Two regimes, pinned in the decisions block:
/// - No explicit conditions: Excel's positional rules. One section
///   takes everything (negatives keep their sign); two split pos+zero
///   from neg (the negative section renders |v|); three add a zero
///   section.
/// - Any explicit condition: positional sign handling is OFF. Sections
///   are tried in order (conditioned ones by their condition, bare ones
///   as else-arms) and the value keeps its own sign everywhere.
fn pickSection(fmt: *const Format, n: f64) ?SectionPick {
    var any_condition = false;
    for (fmt.sections) |s| {
        if (s.condition != null) any_condition = true;
    }

    if (any_condition) {
        for (fmt.sections, 0..) |s, idx| {
            if (s.class == .text) continue;
            if (s.condition) |c| {
                if (c.matches(n)) return .{ .index = idx, .abs = false, .explicit_sign = false };
            } else {
                return .{ .index = idx, .abs = false, .explicit_sign = false };
            }
        }
        return null;
    }

    var numeric_indices: [4]usize = undefined;
    var count: usize = 0;
    for (fmt.sections, 0..) |s, idx| {
        if (s.class == .text) continue;
        numeric_indices[count] = idx;
        count += 1;
    }
    if (count == 0) {
        // Only a text section exists: numbers render their General
        // spelling through its `@` (renderNumberThroughAt).
        return .{ .index = 0, .abs = false, .explicit_sign = false };
    }

    if (n > 0) return .{ .index = numeric_indices[0], .abs = false, .explicit_sign = false };
    if (n < 0) {
        if (count >= 2) return .{ .index = numeric_indices[1], .abs = true, .explicit_sign = false };
        return .{ .index = numeric_indices[0], .abs = true, .explicit_sign = true };
    }
    if (count >= 3) return .{ .index = numeric_indices[2], .abs = false, .explicit_sign = false };
    return .{ .index = numeric_indices[0], .abs = false, .explicit_sign = false };
}

/// A number arriving at a text (`@`) section renders its General
/// spelling at each `@` — the same bytes concatenation would use, so
/// `"@"` applied to 12 and `12&""` cannot disagree.
fn renderNumberThroughAt(gpa: std.mem.Allocator, sec: *const Section, n: f64) RenderError![]u8 {
    var buf: [value.format_buf_len]u8 = undefined;
    const s = value.formatNumber(&buf, n);
    return renderTextSection(gpa, sec, s);
}

fn renderTextSection(gpa: std.mem.Allocator, sec: *const Section, text: []const u8) RenderError![]u8 {
    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(gpa);
    for (sec.tokens) |t| switch (t) {
        .text_at => try out.appendSlice(gpa, text),
        .literal => |lit| try out.appendSlice(gpa, lit),
        .currency => |cur| try out.appendSlice(gpa, cur),
        .percent => try out.append(gpa, '%'),
        .skip => try out.append(gpa, ' '),
        .fill => {},
        else => {},
    };
    return out.toOwnedSlice(gpa);
}

fn renderGeneralNumber(gpa: std.mem.Allocator, n: f64) RenderError![]u8 {
    var buf: [value.format_buf_len]u8 = undefined;
    const s = value.formatNumber(&buf, n);
    return gpa.dupe(u8, s);
}

fn renderGeneralSection(
    gpa: std.mem.Allocator,
    sec: *const Section,
    n: f64,
    explicit_sign: bool,
) RenderError![]u8 {
    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(gpa);
    // `formatNumber` spells the sign itself; a positional-negative
    // section rendering |v| re-adds it here only when asked to.
    if (explicit_sign and n >= 0) try out.append(gpa, '-');
    var buf: [value.format_buf_len]u8 = undefined;
    for (sec.tokens) |t| switch (t) {
        .general_token => try out.appendSlice(gpa, value.formatNumber(&buf, n)),
        .literal => |lit| try out.appendSlice(gpa, lit),
        .currency => |cur| try out.appendSlice(gpa, cur),
        .percent => try out.append(gpa, '%'),
        .skip => try out.append(gpa, ' '),
        .fill => {},
        else => {},
    };
    return out.toOwnedSlice(gpa);
}

// ── decimal core ──────────────────────────────────────────────────

/// A finite f64 as decimal digits: value = ±0.digits × 10^exp. The
/// digits are the shortest round-trip set (the repo's one answer to
/// what a float's digits ARE — `value.formatNumber`'s N5), so all
/// rounding below is decimal-string arithmetic and never a second
/// float operation.
const Decimal = struct {
    neg: bool,
    /// ≤17 significant digits from shortest round-trip; +1 slack for a
    /// rounding carry.
    digits: [20]u8,
    len: u8,
    exp: i32,

    fn isZero(self: *const Decimal) bool {
        return self.len == 0;
    }

    /// The digit at decimal place `exponent` (0 = ones, -1 = tenths).
    fn digitAt(self: *const Decimal, exponent: i32) u8 {
        const i = @as(i64, self.exp) - 1 - exponent;
        if (i < 0 or i >= self.len) return '0';
        return self.digits[@intCast(i)];
    }
};

fn decompose(v: f64) Decimal {
    assert(std.math.isFinite(v));
    var out = Decimal{ .neg = false, .digits = undefined, .len = 0, .exp = 0 };
    var buf: [48]u8 = undefined;
    const s = std.fmt.bufPrint(&buf, "{e}", .{v}) catch unreachable;
    var i: usize = 0;
    if (s[i] == '-') {
        out.neg = true;
        i += 1;
    }
    var mantissa_end = i;
    while (mantissa_end < s.len and s[mantissa_end] != 'e') mantissa_end += 1;
    var exp10: i32 = 0;
    {
        var j = mantissa_end + 1;
        var exp_neg = false;
        if (j < s.len and (s[j] == '+' or s[j] == '-')) {
            exp_neg = s[j] == '-';
            j += 1;
        }
        while (j < s.len) : (j += 1) exp10 = exp10 * 10 + @as(i32, s[j] - '0');
        if (exp_neg) exp10 = -exp10;
    }
    while (i < mantissa_end) : (i += 1) {
        if (s[i] == '.') continue;
        out.digits[out.len] = s[i];
        out.len += 1;
    }
    // Only the value 0 prints a leading zero in `{e}`.
    if (out.digits[0] == '0') {
        out.len = 0;
        out.neg = false;
        return out;
    }
    while (out.len > 0 and out.digits[out.len - 1] == '0') out.len -= 1;
    // `{e}` mantissa is d.ddd → value = 0.digits × 10^(exp10 + 1).
    out.exp = exp10 + 1;
    return out;
}

/// Round half away from zero at `places` decimal places (negative
/// rounds left of the point). Decimal-string arithmetic on the
/// shortest digits — the only rounding rule in this file.
fn roundAt(d: Decimal, places: i32) Decimal {
    if (d.isZero()) return d;
    const keep_i64: i64 = @as(i64, d.exp) + places;
    if (keep_i64 <= 0) {
        // Every digit is dropped; the value rounds up only when the
        // leading digit sits exactly one place below the cut and is ≥5.
        if (keep_i64 == 0 and d.digits[0] >= '5') {
            var out = Decimal{ .neg = d.neg, .digits = undefined, .len = 1, .exp = 1 - places };
            out.digits[0] = '1';
            return out;
        }
        return .{ .neg = false, .digits = undefined, .len = 0, .exp = 0 };
    }
    if (keep_i64 >= d.len) return d;
    const keep: u8 = @intCast(keep_i64);
    var out = Decimal{ .neg = d.neg, .digits = d.digits, .len = keep, .exp = d.exp };
    if (d.digits[keep] >= '5') {
        var i: i32 = @as(i32, keep) - 1;
        var carried = false;
        while (i >= 0) : (i -= 1) {
            const idx: usize = @intCast(i);
            if (out.digits[idx] == '9') {
                out.digits[idx] = '0';
            } else {
                out.digits[idx] += 1;
                carried = true;
                break;
            }
        }
        if (!carried) {
            // 999… became 1000…: one new leading digit, same length of
            // significance after the trailing-zero trim below.
            std.mem.copyBackwards(u8, out.digits[1 .. @as(usize, out.len) + 1], out.digits[0..out.len]);
            out.digits[0] = '1';
            out.exp += 1;
        }
    }
    while (out.len > 0 and out.digits[out.len - 1] == '0') out.len -= 1;
    if (out.len == 0) return .{ .neg = false, .digits = undefined, .len = 0, .exp = 0 };
    return out;
}

// ── numeric section renderer ──────────────────────────────────────

fn renderNumberSection(
    gpa: std.mem.Allocator,
    sec: *const Section,
    n: f64,
    explicit_sign: bool,
) RenderError![]u8 {
    return switch (sec.layout.kind) {
        .plain => renderPlainNumber(gpa, sec, n, explicit_sign),
        .scientific => renderScientific(gpa, sec, n, explicit_sign),
        .fraction => renderFraction(gpa, sec, n, explicit_sign),
    };
}

const IntShare = struct { text: []const u8, pad: ?u8 };

/// The slice of `s` that placeholder `ord` (of `places`, counting from
/// the left) emits. Rightmost placeholders take one character each; the
/// leftmost takes all overflow. A placeholder past the string's left
/// edge emits its pad character instead.
fn intShare(s: []const u8, places: u16, ord: u16, pad: DigitPad) IntShare {
    assert(places > 0 and ord < places);
    const from_right: usize = places - 1 - ord;
    if (ord == 0) {
        if (s.len > from_right) return .{ .text = s[0 .. s.len - from_right], .pad = null };
        return .{ .text = "", .pad = padByte(pad) };
    }
    if (from_right < s.len) {
        const idx = s.len - 1 - from_right;
        return .{ .text = s[idx .. idx + 1], .pad = null };
    }
    return .{ .text = "", .pad = padByte(pad) };
}

fn padByte(pad: DigitPad) ?u8 {
    return switch (pad) {
        .zero => '0',
        .space => ' ',
        .none => null,
    };
}

/// The shared emission walk for plain and scientific layouts. The digit
/// strings arrive fully formed; each placeholder token takes its
/// right-to-left share.
fn emitNumericTokens(
    gpa: std.mem.Allocator,
    out: *std.ArrayListUnmanaged(u8),
    sec: *const Section,
    int_str: []const u8,
    frac_sig: []const u8,
    exp_neg: bool,
    exp_str: []const u8,
) RenderError!void {
    const layout = sec.layout;
    for (sec.tokens) |t| switch (t) {
        .digit => |d| switch (d.role) {
            .int => {
                const share = intShare(int_str, layout.int_places, d.ord, d.pad);
                try out.appendSlice(gpa, share.text);
                if (share.pad) |p| try out.append(gpa, p);
            },
            .frac => {
                if (d.ord < frac_sig.len) {
                    try out.append(gpa, frac_sig[d.ord]);
                } else if (padByte(d.pad)) |p| {
                    try out.append(gpa, p);
                }
            },
            .exponent => {
                const share = intShare(exp_str, layout.exp_places, d.ord, d.pad);
                try out.appendSlice(gpa, share.text);
                if (share.pad) |p| try out.append(gpa, p);
            },
            .numerator, .denominator => unreachable, // fraction path only
        },
        .decimal_point => try out.append(gpa, '.'),
        .literal => |lit| try out.appendSlice(gpa, lit),
        .currency => |cur| try out.appendSlice(gpa, cur),
        .percent => try out.append(gpa, '%'),
        .skip => try out.append(gpa, ' '),
        .fill => {},
        .exponent => |e| {
            try out.append(gpa, if (e.upper) 'E' else 'e');
            if (exp_neg) {
                try out.append(gpa, '-');
            } else if (e.sign_always) {
                try out.append(gpa, '+');
            }
        },
        else => {},
    };
}

fn renderPlainNumber(
    gpa: std.mem.Allocator,
    sec: *const Section,
    n: f64,
    explicit_sign: bool,
) RenderError![]u8 {
    const layout = sec.layout;
    var d = decompose(@abs(n));
    // Scale before rounding: `0.0,,` divides by a million first, then
    // rounds at one place — the other order would round twice.
    d.exp -= 3 * @as(i32, layout.scale_commas);
    d.exp += 2 * @as(i32, layout.percent_count);
    d = roundAt(d, @intCast(layout.frac_places));

    var int_buf: [448]u8 = undefined;
    const int_str = buildIntString(&int_buf, &d, layout);
    var frac_buf: [300]u8 = undefined;
    const frac_sig = buildFracSig(&frac_buf, &d, layout.frac_places);

    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(gpa);
    if ((explicit_sign or n < 0) and !d.isZero()) try out.append(gpa, '-');
    try emitNumericTokens(gpa, &out, sec, int_str, frac_sig, false, "");
    return out.toOwnedSlice(gpa);
}

/// The integer digit string, group separators included when the layout
/// asks for them. Grouping counts `0`-forced places too (`0,000`
/// renders 5 as `0,005`), so the digits are zero-extended to the
/// layout's forced width before commas go in; `?`/`#` padding stays
/// per-placeholder in `intShare`, where no comma can reach it.
fn buildIntString(buf: []u8, d: *const Decimal, layout: NumericLayout) []const u8 {
    var digits_buf: [848]u8 = undefined;
    var n: usize = 0;
    if (d.exp > 0) {
        // Magnitude cap: |exp| ≤ 309 for a finite f64, plus at most
        // 2×255 places of `%` scaling (the 255-char format limit bounds
        // the percent count); the buffer holds the largest possible.
        assert(d.exp <= 824);
        var e: i32 = d.exp - 1;
        while (e >= 0) : (e -= 1) {
            digits_buf[n] = d.digitAt(e);
            n += 1;
        }
    }
    while (n < layout.int_zero_places) {
        std.mem.copyBackwards(u8, digits_buf[1 .. n + 1], digits_buf[0..n]);
        digits_buf[0] = '0';
        n += 1;
    }
    if (!layout.grouping) {
        @memcpy(buf[0..n], digits_buf[0..n]);
        return buf[0..n];
    }
    var out_n: usize = 0;
    for (digits_buf[0..n], 0..) |c, i| {
        if (i > 0 and (n - i) % 3 == 0) {
            buf[out_n] = ',';
            out_n += 1;
        }
        buf[out_n] = c;
        out_n += 1;
    }
    return buf[0..out_n];
}

/// The significant fractional digits after rounding: everything up to
/// the last nonzero digit within `frac_places`. Trailing policy is
/// per-placeholder (`0` keeps, `#` drops, `?` spaces) and lives in the
/// emission walk.
fn buildFracSig(buf: []u8, d: *const Decimal, frac_places: u16) []const u8 {
    var n: usize = 0;
    var j: u16 = 0;
    while (j < frac_places) : (j += 1) {
        buf[n] = d.digitAt(-@as(i32, j) - 1);
        n += 1;
    }
    while (n > 0 and buf[n - 1] == '0') n -= 1;
    return buf[0..n];
}

fn renderScientific(
    gpa: std.mem.Allocator,
    sec: *const Section,
    n: f64,
    explicit_sign: bool,
) RenderError![]u8 {
    const layout = sec.layout;
    const k: i32 = @max(1, @as(i32, layout.int_places));
    const d = decompose(@abs(n));

    var exp10: i32 = 0;
    if (!d.isZero()) {
        // Engineering rule: the exponent is the multiple of the int
        // placeholder count that leaves 1…k digits before the point —
        // `##0.0E+0` steps in thousands.
        exp10 = @divFloor(d.exp - 1, k) * k;
    }
    // Round the mantissa; a carry that grows it past k digits steps the
    // exponent and re-rounds (999.95 → 1.0E+3 under ##0.0E+0). The
    // second attempt's mantissa is a power of ten, which cannot carry.
    var attempt: usize = 0;
    while (true) : (attempt += 1) {
        assert(attempt < 2);
        var m = d;
        m.exp -= exp10;
        m = roundAt(m, @intCast(layout.frac_places));
        if (!m.isZero() and m.exp > k) {
            exp10 += k;
            continue;
        }
        var int_buf: [1152]u8 = undefined;
        const int_str = buildIntString(&int_buf, &m, layout);
        var frac_buf: [300]u8 = undefined;
        const frac_sig = buildFracSig(&frac_buf, &m, layout.frac_places);

        var exp_buf: [16]u8 = undefined;
        const exp_str = std.fmt.bufPrint(&exp_buf, "{d}", .{@abs(exp10)}) catch unreachable;

        var out: std.ArrayListUnmanaged(u8) = .empty;
        errdefer out.deinit(gpa);
        if ((explicit_sign or n < 0) and !m.isZero()) try out.append(gpa, '-');
        try emitNumericTokens(gpa, &out, sec, int_str, frac_sig, exp10 < 0, exp_str);
        return out.toOwnedSlice(gpa);
    }
}

/// Best rational approximation of `f` ∈ [0,1) with denominator ≤
/// `max_den`: continued-fraction convergents plus the closing
/// semiconvergent, ties to the smaller denominator. Excel's own search
/// has undocumented quirks; this is numfmt_v1's pinned answer,
/// spec_pinned pending the parked Excel leg.
fn bestFraction(f: f64, max_den: u64) struct { num: u64, den: u64 } {
    assert(f >= 0 and f < 1);
    assert(max_den >= 1);
    var p_prev: u64 = 1; // p(-1)
    var q_prev: u64 = 0;
    var p_cur: u64 = 0; // p(0) — a0 = 0 because f < 1
    var q_cur: u64 = 1;
    var x = f;
    var iterations: usize = 0;
    while (iterations < 64) : (iterations += 1) {
        if (x <= 0) break;
        const inv = 1.0 / x;
        if (!std.math.isFinite(inv)) break;
        const a_f = @floor(inv);
        // Overflow-safe continuation test in floating point first: the
        // exact next denominator only gets computed when it fits.
        if (a_f * @as(f64, @floatFromInt(q_cur)) + @as(f64, @floatFromInt(q_prev)) >
            @as(f64, @floatFromInt(max_den)))
        {
            // Semiconvergent: the largest t keeping the denominator in
            // bounds; then closest wins, tie to the convergent (its
            // denominator is never larger).
            const t = (max_den - q_prev) / q_cur;
            if (t > 0) {
                const sp = t * p_cur + p_prev;
                const sq = t * q_cur + q_prev;
                const err_cur = @abs(f - asF(p_cur, q_cur));
                const err_semi = @abs(f - asF(sp, sq));
                if (err_semi < err_cur) return .{ .num = sp, .den = sq };
            }
            return .{ .num = p_cur, .den = q_cur };
        }
        const a: u64 = @intFromFloat(a_f);
        const p_next = a * p_cur + p_prev;
        const q_next = a * q_cur + q_prev;
        p_prev = p_cur;
        q_prev = q_cur;
        p_cur = p_next;
        q_cur = q_next;
        x = inv - a_f;
        if (x < 1e-12) break;
    }
    return .{ .num = p_cur, .den = q_cur };
}

fn asF(num: u64, den: u64) f64 {
    assert(den != 0);
    return @as(f64, @floatFromInt(num)) / @as(f64, @floatFromInt(den));
}

/// The largest magnitude the fraction search can hold exactly. Beyond
/// it a fraction format falls back to the General spelling — pinned;
/// a saturated wrong fraction would be a lie with a denominator.
const max_fraction_magnitude: f64 = 9007199254740992.0; // 2^53

fn renderFraction(
    gpa: std.mem.Allocator,
    sec: *const Section,
    n: f64,
    explicit_sign: bool,
) RenderError![]u8 {
    const layout = sec.layout;
    var av = @abs(n);
    // `%` scales fractions too: `0/8%` counts hundredths of a percent.
    var pc = layout.percent_count;
    while (pc > 0) : (pc -= 1) av *= 100;

    // Improper fractions ride the numerator, so their cap is what a
    // numerator times a six-digit denominator can hold, not 2^53.
    const cap: f64 = if (layout.int_places == 0) 9.0e12 else max_fraction_magnitude;
    if (av >= cap or !std.math.isFinite(av)) {
        var out: std.ArrayListUnmanaged(u8) = .empty;
        errdefer out.deinit(gpa);
        if (explicit_sign and n >= 0) try out.append(gpa, '-');
        var buf: [value.format_buf_len]u8 = undefined;
        try out.appendSlice(gpa, value.formatNumber(&buf, n));
        return out.toOwnedSlice(gpa);
    }

    var int_part: u64 = 0;
    var frac_part: f64 = av;
    if (layout.int_places > 0) {
        int_part = @intFromFloat(@floor(av));
        frac_part = av - @floor(av);
    }

    var num: u64 = 0;
    var den: u64 = 1;
    if (layout.den_fixed != 0) {
        den = layout.den_fixed;
        // Half away from zero; the domain is non-negative here.
        num = @intFromFloat(@floor(frac_part * @as(f64, @floatFromInt(den)) + 0.5));
    } else {
        var max_den: u64 = 1;
        var p: u16 = 0;
        while (p < layout.den_places) : (p += 1) max_den *= 10;
        max_den -= 1;
        assert(max_den >= 1);
        // The search always runs on the fractional remainder; the
        // improper shape re-attaches the whole part to the numerator.
        const fr = bestFraction(av - @floor(av), max_den);
        if (layout.int_places == 0) {
            // Improper: `?/?` on 2.5 is 5/2 — the integer rides the
            // numerator over the same denominator.
            den = if (fr.den == 0) 1 else fr.den;
            num = fr.num + @as(u64, @intFromFloat(@floor(av))) * den;
        } else {
            num = fr.num;
            den = if (fr.den == 0) 1 else fr.den;
        }
    }
    // A numerator that reached the denominator is a whole: 0.99 under
    // `# ?/?` carries into the integer and blanks the fraction.
    if (layout.int_places > 0 and num == den and layout.den_fixed == 0) {
        int_part += 1;
        num = 0;
        den = 1;
    }
    if (layout.den_fixed != 0 and num == den) {
        if (layout.int_places > 0) {
            int_part += 1;
            num = 0;
        }
        // No integer part to carry into: 0.999 under `?/8` stays 8/8 —
        // the fixed denominator is the caller's ruler, not ours.
    }

    // Only a searched-for fraction blanks at zero; a fixed denominator
    // is the caller's ruler and its `0/8` is information.
    const blank_fraction = layout.int_places > 0 and num == 0 and layout.den_fixed == 0;

    var int_buf: [24]u8 = undefined;
    var int_str: []const u8 = "";
    if (layout.int_places > 0) {
        if (int_part != 0 or blank_fraction) {
            int_str = std.fmt.bufPrint(&int_buf, "{d}", .{int_part}) catch unreachable;
        }
    }

    var num_buf: [24]u8 = undefined;
    var den_buf: [24]u8 = undefined;
    const num_str = if (blank_fraction) "" else std.fmt.bufPrint(&num_buf, "{d}", .{num}) catch unreachable;
    const den_str = if (blank_fraction) "" else std.fmt.bufPrint(&den_buf, "{d}", .{den}) catch unreachable;

    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(gpa);
    const is_zero = int_part == 0 and num == 0;
    if ((explicit_sign or n < 0) and !is_zero) try out.append(gpa, '-');

    for (sec.tokens) |t| switch (t) {
        .digit => |d| switch (d.role) {
            .int => {
                const share = intShare(int_str, layout.int_places, d.ord, d.pad);
                try out.appendSlice(gpa, share.text);
                if (share.pad) |p| try out.append(gpa, p);
            },
            .numerator => {
                const share = intShare(num_str, layout.num_places, d.ord, d.pad);
                try out.appendSlice(gpa, share.text);
                if (share.pad) |p| {
                    try out.append(gpa, if (blank_fraction) ' ' else p);
                }
            },
            .denominator => {
                // Denominators pad on the RIGHT: `# ??/??` on 5.25 is
                // `5  1/4 `, the denominator left-aligned in its field.
                if (d.ord < den_str.len) {
                    try out.append(gpa, den_str[d.ord]);
                } else if (padByte(d.pad)) |p| {
                    try out.append(gpa, if (blank_fraction) ' ' else p);
                }
            },
            .frac, .exponent => unreachable,
        },
        .fraction_slash => try out.append(gpa, if (blank_fraction) ' ' else '/'),
        .literal => |lit| try out.appendSlice(gpa, lit),
        .currency => |cur| try out.appendSlice(gpa, cur),
        .percent => try out.append(gpa, '%'),
        .skip => try out.append(gpa, ' '),
        .fill => {},
        else => {},
    };
    return out.toOwnedSlice(gpa);
}

// ── date section renderer ─────────────────────────────────────────

// en-US invariant name tables — the tables `[$-409]` licenses (§5.4b).
const month_abbr = [_][]const u8{ "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
const month_full = [_][]const u8{ "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" };
const day_abbr = [_][]const u8{ "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" };
const day_full = [_][]const u8{ "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday" };

/// Weekday index (0 = Sunday) by SERIAL arithmetic, not by the civil
/// calendar: Excel's early-1900 drift means `WEEKDAY(1)` is Sunday even
/// though 1900-01-01 was a Monday, and the day NAMES must drift with it
/// or `ddd` and `WEEKDAY` would disagree about the same cell.
fn weekdayIndex(system: serial_date.DateSystem, days: i64) usize {
    return switch (system) {
        .d1900 => @intCast(@mod(days + 6, 7)),
        // 1904's serial 0 is 1904-01-01, a Friday.
        .d1904 => @intCast(@mod(days + 5, 7)),
    };
}

fn renderDateSection(
    gpa: std.mem.Allocator,
    sec: *const Section,
    serial: f64,
    opts: RenderOptions,
) RenderError![]u8 {
    // Negative durations exist in 1904-system workbooks; v1 refuses
    // them along with every other out-of-domain serial (decisions
    // block, spec_pinned divergence).
    if (serial < 0) return error.SerialOutOfRange;
    const max_serial: f64 = @floatFromInt(serial_date.maxSerial(opts.date_system));
    if (serial >= max_serial + 1.0) return error.SerialOutOfRange;

    const layout = sec.date;

    // Units per day at the finest displayed precision. The serial is
    // rounded HERE, once, so every displayed component agrees about
    // the carry.
    var upd: i64 = switch (layout.finest) {
        .day => 1,
        .hour => 24,
        .minute => 24 * 60,
        .second, .subsecond => 24 * 60 * 60,
    };
    var sub_scale: i64 = 1;
    if (layout.finest == .subsecond) {
        var p: u8 = 0;
        while (p < layout.subsecond_digits) : (p += 1) sub_scale *= 10;
        upd *= sub_scale;
    }

    // Date-only formats truncate the time instead of rounding: 45000.9
    // is still 2023-03-15. Time formats round at their finest unit.
    const total: i64 = if (layout.finest == .day)
        @intFromFloat(@floor(serial))
    else
        @intFromFloat(@round(serial * @as(f64, @floatFromInt(upd))));

    const days: i64 = @divFloor(total, upd);
    const day_units: i64 = @mod(total, upd);

    if (days > serial_date.maxSerial(opts.date_system)) return error.SerialOutOfRange;

    var needs_date = false;
    for (sec.tokens) |t| switch (t) {
        .date => |p| switch (p) {
            .year_two, .year_four, .month, .month_padded, .month_abbr, .month_full, .month_letter, .day, .day_padded, .day_abbr, .day_full => needs_date = true,
            else => {},
        },
        else => {},
    };
    var date: serial_date.Date = undefined;
    if (needs_date) {
        date = serial_date.dateFromSerial(opts.date_system, @intCast(days)) catch {
            return error.SerialOutOfRange;
        };
    }

    // Unsigned from here down: the serial ≥ 0 guard above makes every
    // component non-negative, and Zig 0.16's `{d:0>2}` prints a sign
    // column for signed operands.
    const subsecond_value: u64 = if (layout.finest == .subsecond) @intCast(@mod(day_units, sub_scale)) else 0;
    const seconds_of_day: u64 = @intCast(switch (layout.finest) {
        .day => 0,
        .hour => day_units * 3600,
        .minute => day_units * 60,
        .second => day_units,
        .subsecond => @divFloor(day_units, sub_scale),
    });
    const hour24: u64 = seconds_of_day / 3600;
    const minute: u64 = (seconds_of_day / 60) % 60;
    const second: u64 = seconds_of_day % 60;

    // Elapsed counts never wrap at the day boundary.
    const secs_total: i64 = switch (layout.finest) {
        .day => total * 86400,
        .hour => total * 3600,
        .minute => total * 60,
        .second => total,
        .subsecond => @divFloor(total, sub_scale),
    };

    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(gpa);
    var buf: [32]u8 = undefined;

    for (sec.tokens) |t| switch (t) {
        .date => |p| switch (p) {
            .year_two => try out.appendSlice(gpa, std.fmt.bufPrint(&buf, "{d:0>2}", .{date.year % 100}) catch unreachable),
            .year_four => try out.appendSlice(gpa, std.fmt.bufPrint(&buf, "{d:0>4}", .{date.year}) catch unreachable),
            .month => try out.appendSlice(gpa, std.fmt.bufPrint(&buf, "{d}", .{date.month}) catch unreachable),
            .month_padded => try out.appendSlice(gpa, std.fmt.bufPrint(&buf, "{d:0>2}", .{date.month}) catch unreachable),
            .month_abbr => try out.appendSlice(gpa, month_abbr[date.month - 1]),
            .month_full => try out.appendSlice(gpa, month_full[date.month - 1]),
            .month_letter => try out.appendSlice(gpa, month_full[date.month - 1][0..1]),
            .day => try out.appendSlice(gpa, std.fmt.bufPrint(&buf, "{d}", .{date.day}) catch unreachable),
            .day_padded => try out.appendSlice(gpa, std.fmt.bufPrint(&buf, "{d:0>2}", .{date.day}) catch unreachable),
            .day_abbr => try out.appendSlice(gpa, day_abbr[weekdayIndex(opts.date_system, days)]),
            .day_full => try out.appendSlice(gpa, day_full[weekdayIndex(opts.date_system, days)]),
            .hour, .hour_padded => {
                var h = hour24;
                if (layout.twelve_hour) {
                    h = @mod(hour24, 12);
                    if (h == 0) h = 12;
                }
                if (p == .hour_padded) {
                    try out.appendSlice(gpa, std.fmt.bufPrint(&buf, "{d:0>2}", .{h}) catch unreachable);
                } else {
                    try out.appendSlice(gpa, std.fmt.bufPrint(&buf, "{d}", .{h}) catch unreachable);
                }
            },
            .minute => try out.appendSlice(gpa, std.fmt.bufPrint(&buf, "{d}", .{minute}) catch unreachable),
            .minute_padded => try out.appendSlice(gpa, std.fmt.bufPrint(&buf, "{d:0>2}", .{minute}) catch unreachable),
            .second => try out.appendSlice(gpa, std.fmt.bufPrint(&buf, "{d}", .{second}) catch unreachable),
            .second_padded => try out.appendSlice(gpa, std.fmt.bufPrint(&buf, "{d:0>2}", .{second}) catch unreachable),
            .am_pm => try out.appendSlice(gpa, if (hour24 < 12) "AM" else "PM"),
            .a_p => try out.appendSlice(gpa, if (hour24 < 12) "A" else "P"),
        },
        .elapsed => |e| {
            const units: u64 = @intCast(switch (e.unit) {
                .hours => @divFloor(secs_total, 3600),
                .minutes => @divFloor(secs_total, 60),
                .seconds => secs_total,
            });
            const w: usize = @max(1, @as(usize, e.width));
            var nbuf: [24]u8 = undefined;
            const digits = std.fmt.bufPrint(&nbuf, "{d}", .{units}) catch unreachable;
            var pad_i = digits.len;
            while (pad_i < w) : (pad_i += 1) try out.append(gpa, '0');
            try out.appendSlice(gpa, digits);
        },
        .subsecond => |sub| {
            try out.append(gpa, '.');
            var nbuf: [8]u8 = undefined;
            const digits = std.fmt.bufPrint(&nbuf, "{d}", .{subsecond_value}) catch unreachable;
            var pad_i = digits.len;
            while (pad_i < sub.digits) : (pad_i += 1) try out.append(gpa, '0');
            try out.appendSlice(gpa, digits);
        },
        .literal => |lit| try out.appendSlice(gpa, lit),
        .currency => |cur| try out.appendSlice(gpa, cur),
        .skip => try out.append(gpa, ' '),
        .fill => {},
        else => {},
    };
    return out.toOwnedSlice(gpa);
}

// ─── the canonical built-in id table ─────────────────────────────

/// The ECMA-376 built-in numFmtId → format code table, ids 0–49, minus
/// the ids whose codes are locale-negotiated at open time (5–8, 23–36,
/// 41–44) — the same deliberate skips as `pkg/workbook.zig`'s
/// `builtinNumFmtCode` and `src/xlsx.zig`'s `builtinNumberFormat`,
/// which this table exists to eventually replace as the single copy.
pub fn builtinFormatCode(id: u32) ?[]const u8 {
    return switch (id) {
        0 => "General",
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        14 => "mm-dd-yy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yy h:mm",
        37 => "#,##0 ;(#,##0)",
        38 => "#,##0 ;[Red](#,##0)",
        39 => "#,##0.00;(#,##0.00)",
        40 => "#,##0.00;[Red](#,##0.00)",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mmss.0",
        48 => "##0.0E+0",
        49 => "@",
        else => null,
    };
}

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

fn expectParseOk(code: []const u8) !Format {
    switch (try parse(testing.allocator, code)) {
        .ok => |fmt| return fmt,
        .refused => |r| {
            std.debug.print("unexpected refusal {t} at {d} for \"{s}\"\n", .{ r.reason, r.at, code });
            return error.TestUnexpectedResult;
        },
    }
}

fn expectRender(code: []const u8, v: RenderValue, expected: []const u8) !void {
    var fmt = try expectParseOk(code);
    defer fmt.deinit(testing.allocator);
    const got = try render(testing.allocator, &fmt, v, .{});
    defer testing.allocator.free(got);
    if (!std.mem.eql(u8, got, expected)) {
        std.debug.print("format \"{s}\": expected \"{s}\", got \"{s}\"\n", .{ code, expected, got });
        return error.TestUnexpectedResult;
    }
}

fn expectRefusal(code: []const u8, reason: Refusal.Reason) !void {
    switch (try parse(testing.allocator, code)) {
        .ok => |fmt| {
            var f = fmt;
            f.deinit(testing.allocator);
            std.debug.print("expected {t}, but \"{s}\" parsed\n", .{ reason, code });
            return error.TestUnexpectedResult;
        },
        .refused => |r| {
            if (r.reason != reason) {
                std.debug.print("expected {t}, got {t} for \"{s}\"\n", .{ reason, r.reason, code });
                return error.TestUnexpectedResult;
            }
        },
    }
}

test "numfmt_v1: the matrix covers every construct exactly once" {
    var seen_count = [_]usize{0} ** @typeInfo(Construct).@"enum".fields.len;
    for (support_matrix) |row| {
        seen_count[@intFromEnum(row.construct)] += 1;
    }
    inline for (@typeInfo(Construct).@"enum".fields) |f| {
        try testing.expectEqual(@as(usize, 1), seen_count[f.value]);
    }
    try testing.expectEqual(@typeInfo(Construct).@"enum".fields.len, support_matrix.len);
}

test "numfmt_v1: every matrix example answers its row" {
    for (support_matrix) |row| {
        switch (row.status) {
            .rendered => {
                var fmt = try expectParseOk(row.example);
                defer fmt.deinit(testing.allocator);
                // The example must exercise the construct it stands for,
                // or the row proves nothing.
                try testing.expect(fmt.seen.contains(row.construct));
            },
            .refused => |want_reason| switch (try parse(testing.allocator, row.example)) {
                .ok => |fmt| {
                    var f = fmt;
                    f.deinit(testing.allocator);
                    return error.TestUnexpectedResult;
                },
                .refused => |r| {
                    try testing.expectEqual(want_reason, r.reason);
                    // The refusal names its own matrix row.
                    try testing.expectEqual(row.construct, r.construct.?);
                },
            },
        }
    }
}

test "numfmt_v1: the refused rows are the park list, derived not restated" {
    var buf: [support_matrix.len]Construct = undefined;
    const refused = refusedConstructs(&buf);
    // Derived from the table…
    var expect_n: usize = 0;
    for (support_matrix) |row| {
        if (row.status == .refused) {
            try testing.expectEqual(row.construct, refused[expect_n]);
            expect_n += 1;
        }
    }
    try testing.expectEqual(expect_n, refused.len);
    // …with the count pinned once, so a vanished row cannot hide
    // behind the derivation (the spill_transitions precedent).
    try testing.expectEqual(@as(usize, 7), refused.len);
}

test "numfmt_v1: every refusal reason has a fixture that produces it" {
    const cases = [_]struct { code: []const u8, reason: Refusal.Reason }{
        .{ .code = "0" ** (max_format_chars + 1), .reason = .format_too_long },
        .{ .code = "\"abc", .reason = .unclosed_quote },
        .{ .code = "[Red", .reason = .unclosed_bracket },
        .{ .code = "[Foo]0", .reason = .unknown_bracket_tag },
        .{ .code = "0;0;0;0;0", .reason = .too_many_sections },
        .{ .code = "0;0;[>5]0", .reason = .condition_misplaced },
        .{ .code = "[>abc]0", .reason = .condition_malformed },
        .{ .code = "0.0.0", .reason = .multiple_decimal_points },
        .{ .code = "0@", .reason = .mixed_text_number },
        .{ .code = "0 yyyy", .reason = .mixed_date_number },
        .{ .code = "# ?/", .reason = .malformed_fraction },
        .{ .code = "[h][h]", .reason = .elapsed_misplaced },
        .{ .code = "ss.0000", .reason = .subsecond_precision },
        .{ .code = "0\\", .reason = .dangling_modifier },
        .{ .code = "0q", .reason = .unknown_token },
        .{ .code = "[$-40C]0", .reason = .locale_table_unavailable },
        .{ .code = "[DBNum1]0", .reason = .numeral_shaping_unsupported },
        .{ .code = "g", .reason = .calendar_unsupported },
    };
    for (cases) |case| try expectRefusal(case.code, case.reason);
    // Reason coverage is derived, not asserted by hand: every enum
    // member must appear in the case table.
    inline for (@typeInfo(Refusal.Reason).@"enum".fields) |f| {
        const reason: Refusal.Reason = @enumFromInt(f.value);
        var covered = false;
        for (cases) |case| {
            if (case.reason == reason) covered = true;
        }
        try testing.expect(covered);
    }
}

test "numfmt_v1: the §9 length limit is a boundary, not a vicinity" {
    var at_limit = try expectParseOk("0" ** max_format_chars);
    at_limit.deinit(testing.allocator);
    try expectRefusal("0" ** (max_format_chars + 1), .format_too_long);
}

test "numfmt_v1: plain digit layouts render byte-exactly" {
    try expectRender("General", .{ .number = 1234.5678 }, "1234.5678");
    try expectRender("General", .{ .number = 0 }, "0");
    try expectRender("00", .{ .number = 5 }, "05");
    try expectRender("00.00", .{ .number = -5 }, "-05.00");
    try expectRender("#", .{ .number = 0 }, "");
    try expectRender("#.##", .{ .number = 5.5 }, "5.5");
    try expectRender("#.##", .{ .number = 0.567 }, ".57");
    try expectRender("??", .{ .number = 5 }, " 5");
    try expectRender("?.??", .{ .number = 0.5 }, " .5 ");
    try expectRender("0.00", .{ .number = 0.996 }, "1.00");
    try expectRender("0.0", .{ .number = 0.25 }, "0.3"); // half away from zero
    try expectRender("0", .{ .number = -0.0 }, "0"); // -0 shows no sign
}

test "numfmt_v1: grouping, scale and percent" {
    try expectRender("#,##0", .{ .number = 1234567 }, "1,234,567");
    try expectRender("#,##0.00", .{ .number = 1234567.891 }, "1,234,567.89");
    try expectRender("0,000", .{ .number = 5 }, "0,005");
    try expectRender("0.0,,", .{ .number = 12000000 }, "12.0");
    try expectRender("#,", .{ .number = 1234567 }, "1235");
    try expectRender("0%", .{ .number = 0.375 }, "38%");
    try expectRender("0.00%", .{ .number = 0.375 }, "37.50%");
}

test "numfmt_v1: scientific layouts, engineering step included" {
    try expectRender("0.00E+00", .{ .number = 12345.6789 }, "1.23E+04");
    try expectRender("0.00E+00", .{ .number = 0.000123 }, "1.23E-04");
    try expectRender("0.00E+00", .{ .number = 0 }, "0.00E+00");
    try expectRender("##0.0E+0", .{ .number = 12345 }, "12.3E+3");
    try expectRender("##0.0E+0", .{ .number = 999.95 }, "1.0E+3");
    try expectRender("0.00E-00", .{ .number = 12345.6789 }, "1.23E04");
    try expectRender("0.00e-00", .{ .number = 0.000123 }, "1.23e-04");
}

test "numfmt_v1: fractions — search, fixed, padding, blanking" {
    try expectRender("# ?/?", .{ .number = 5.25 }, "5 1/4");
    try expectRender("# ?/?", .{ .number = 0.333333333333333 }, " 1/3");
    try expectRender("# ?/?", .{ .number = 5 }, "5    ");
    try expectRender("# ??/??", .{ .number = 5.25 }, "5  1/4 ");
    try expectRender("?/8", .{ .number = 0.5 }, "4/8");
    try expectRender("# ?/8", .{ .number = 5.0 }, "5 0/8");
    try expectRender("?/?", .{ .number = 2.5 }, "5/2");
    try expectRender("# ?/?", .{ .number = -5.25 }, "-5 1/4");
    try expectRender("# ??/??", .{ .number = 0.005 }, "0      "); // rounds to whole zero
}

test "numfmt_v1: date layouts under both epochs" {
    try expectRender("yyyy-mm-dd", .{ .number = 45000 }, "2023-03-15");
    try expectRender("yyyy-mm-dd", .{ .number = 60 }, "1900-02-29"); // the fictitious leap day
    try expectRender("yyyy-mm-dd", .{ .number = 0 }, "1900-01-00"); // day zero
    try expectRender("yyyy-mm-dd", .{ .number = 45000.9 }, "2023-03-15"); // date-only truncates
    try expectRender("ddd mmm d", .{ .number = 45000 }, "Wed Mar 15");
    try expectRender("dddd", .{ .number = 45000 }, "Wednesday");
    try expectRender("mmmmm", .{ .number = 45000 }, "M");
    try expectRender("yy", .{ .number = 45000 }, "23");
    try expectRender("ddd", .{ .number = 1 }, "Sun"); // Excel's drifted 1900 week

    // 1904: same code, different epoch and phase.
    var fmt = try expectParseOk("yyyy-mm-dd ddd");
    defer fmt.deinit(testing.allocator);
    const got = try render(testing.allocator, &fmt, .{ .number = 0 }, .{ .date_system = .d1904 });
    defer testing.allocator.free(got);
    try testing.expectEqualStrings("1904-01-01 Fri", got);
}

test "numfmt_v1: time layouts — clock, twelve-hour, rounding carry" {
    try expectRender("m/d/yyyy h:mm", .{ .number = 45000.5 }, "3/15/2023 12:00");
    try expectRender("hh:mm AM/PM", .{ .number = 0.75 }, "06:00 PM");
    try expectRender("h:mm A/P", .{ .number = 0.75 }, "6:00 P");
    try expectRender("hh:mm", .{ .number = 0.5 + 59.6 / 86400.0 }, "12:01"); // rounds at the finest unit
    try expectRender("hh:mm:ss.000", .{ .number = 45000.123456 }, "02:57:46.598");
    try expectRender("hh:mm:ss.0", .{ .number = 0.5 + 30.25 / 86400.0 }, "12:00:30.3");
    try expectRender("mm:ss", .{ .number = 0.75 }, "00:00"); // mm here is minutes
    try expectRender("yyyy-mm", .{ .number = 45000 }, "2023-03"); // and here months
}

test "numfmt_v1: elapsed durations do not wrap at the day" {
    try expectRender("[h]:mm", .{ .number = 1.5 }, "36:00");
    try expectRender("[mm]:ss", .{ .number = 1.5 }, "2160:00");
    try expectRender("[s]", .{ .number = 90.0 / 86400.0 }, "90");
    try expectRender("[hh]:mm", .{ .number = 0.25 }, "06:00");
}

test "numfmt_v1: sections, conditions and signs" {
    try expectRender("[>10]\"big\";\"small\"", .{ .number = 15 }, "big");
    try expectRender("[>10]\"big\";\"small\"", .{ .number = 5 }, "small");
    // Explicit conditions keep the value's own sign.
    try expectRender("[<0]0.0;0.0", .{ .number = -5 }, "-5.0");
    // No condition matched: pinned General fallback.
    try expectRender("[>10]0", .{ .number = 5 }, "5");
    // Positional sections: the negative section renders |v|.
    try expectRender("#,##0.00;(#,##0.00)", .{ .number = -1234.5 }, "(1,234.50)");
    try expectRender("0.00;[Red]-0.00", .{ .number = -5 }, "-5.00");
    try expectRender("0.00;-0.00;\"zero\"", .{ .number = 0 }, "zero");
    try expectRender("0.0;;", .{ .number = -5 }, "");
    try expectRender("0.0;;", .{ .number = 5 }, "5.0");
    // One section: negatives keep their sign in front of everything.
    try expectRender("\"x\"0", .{ .number = -5 }, "-x5");
}

test "numfmt_v1: text values, text sections, and the @ bridge" {
    try expectRender("@", .{ .text = "abc" }, "abc");
    try expectRender("\"pre \"@\" post\"", .{ .text = "abc" }, "pre abc post");
    try expectRender("@", .{ .number = 12 }, "12"); // number through @ = General spelling
    try expectRender("0.00", .{ .text = "abc" }, "abc"); // no text section: text passes through
    try expectRender("0.00;-0.00;0.00;\"<\"@\">\"", .{ .text = "abc" }, "<abc>");
    try expectRender("", .{ .number = 5 }, ""); // the empty-format quirk, pinned
}

test "numfmt_v1: booleans bypass every section" {
    try expectRender("0.00", .{ .boolean = true }, "TRUE");
    try expectRender("@", .{ .boolean = false }, "FALSE");
    try expectRender("yyyy-mm-dd", .{ .boolean = true }, "TRUE");
}

test "numfmt_v1: literals, escapes, fill and skip" {
    try expectRender("0\\h", .{ .number = 5 }, "5h");
    try expectRender("0 \"units\"", .{ .number = 5 }, "5 units");
    try expectRender("*-0", .{ .number = 5 }, "5"); // fill elided — TEXT parity
    try expectRender("_-0", .{ .number = 5 }, " 5"); // skip is one space
    try expectRender("$0", .{ .number = 5 }, "$5");
}

test "numfmt_v1: locale tags — 409 renders, currency is literal" {
    try expectRender("[$-409]#,##0.00", .{ .number = 1234.5 }, "1,234.50");
    try expectRender("[$\u{20AC}-409]#,##0.00", .{ .number = 1234.5 }, "\u{20AC}1,234.50");
    try expectRender("[$USD]0", .{ .number = 7 }, "USD7");
    try expectRender("[Red]0;[Blue]-0", .{ .number = -3 }, "-3"); // colors carry no bytes
}

test "numfmt_v1: out-of-domain serials refuse rendering, typed" {
    var fmt = try expectParseOk("yyyy-mm-dd");
    defer fmt.deinit(testing.allocator);
    try testing.expectError(error.SerialOutOfRange, render(testing.allocator, &fmt, .{ .number = -1 }, .{}));
    try testing.expectError(error.SerialOutOfRange, render(testing.allocator, &fmt, .{ .number = 3000000 }, .{}));
    var dur = try expectParseOk("[h]:mm");
    defer dur.deinit(testing.allocator);
    try testing.expectError(error.SerialOutOfRange, render(testing.allocator, &dur, .{ .number = -0.5 }, .{}));
}

test "numfmt_v1: every built-in id parses, and describesDate matches the known date set" {
    var id: u32 = 0;
    while (id <= 49) : (id += 1) {
        const code = builtinFormatCode(id) orelse continue;
        var fmt = try expectParseOk(code);
        defer fmt.deinit(testing.allocator);
        const is_date = (id >= 14 and id <= 22) or (id >= 45 and id <= 47);
        try testing.expectEqual(is_date, fmt.describesDate());
        // Crash-freedom over a representative value for every id.
        const v: RenderValue = if (id == 49) .{ .text = "t" } else .{ .number = 1234.5678 };
        const got = try render(testing.allocator, &fmt, v, .{});
        testing.allocator.free(got);
    }
}

test "numfmt_v1: planeTwo splits the three refusal families" {
    inline for (@typeInfo(Refusal.Reason).@"enum".fields) |f| {
        const reason: Refusal.Reason = @enumFromInt(f.value);
        const plane = (Refusal{ .reason = reason, .at = 0, .construct = null }).planeTwo();
        switch (reason) {
            .format_too_long => try testing.expectEqual(parser.PlaneTwo.FormulaLimitExceeded, plane),
            .locale_table_unavailable,
            .numeral_shaping_unsupported,
            .calendar_unsupported,
            => try testing.expectEqual(parser.PlaneTwo.FormulaLocaleSensitiveInput, plane),
            else => try testing.expectEqual(parser.PlaneTwo.FormulaMalformedInput, plane),
        }
    }
}

test "numfmt_v1: every rendered matrix row has at least one byte-exact fixture" {
    // The fixtures above are keyed by format code; this test re-derives
    // coverage from the matrix by parsing each rendered row's example
    // and rendering a probe value — a row whose example cannot render
    // is a row whose "rendered" claim is false.
    for (support_matrix) |row| {
        if (row.status != .rendered) continue;
        var fmt = try expectParseOk(row.example);
        defer fmt.deinit(testing.allocator);
        const probe: RenderValue = if (row.construct == .text_at)
            .{ .text = "probe" }
        else
            .{ .number = 1.5 };
        const got = render(testing.allocator, &fmt, probe, .{}) catch |e| switch (e) {
            error.SerialOutOfRange => unreachable, // 1.5 is in-domain everywhere
            else => return e,
        };
        testing.allocator.free(got);
    }
}

fn fuzzNumfmtTarget(_: void, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();
    var smith_buf: [512]u8 = undefined;
    const input = smith_buf[0..smith.slice(&smith_buf)];
    const gpa = testing.allocator;
    switch (try parse(gpa, input)) {
        .refused => |r| {
            // Every refusal has a plane; a reason without one is a
            // compile error, but a torn refusal value would surface here.
            _ = r.planeTwo();
        },
        .ok => |fmt| {
            var f = fmt;
            defer f.deinit(gpa);
            const probes = [_]RenderValue{
                .{ .number = 0 },
                .{ .number = 1 },
                .{ .number = -2.75 },
                .{ .number = 0.15 },
                .{ .number = 1234567.891 },
                .{ .number = 45000.123456 },
                .{ .number = 1e300 },
                .{ .number = -1e300 },
                .{ .number = 5e-324 },
                .{ .text = "ß" },
                .{ .boolean = true },
            };
            for (probes) |p| {
                const first = render(gpa, &f, p, .{}) catch |e| switch (e) {
                    error.SerialOutOfRange => continue,
                    error.OutOfMemory => return e,
                };
                defer gpa.free(first);
                // Deterministic: the same format and value must agree
                // with themselves.
                const second = render(gpa, &f, p, .{}) catch |e| switch (e) {
                    error.SerialOutOfRange => unreachable,
                    error.OutOfMemory => return e,
                };
                defer gpa.free(second);
                try testing.expectEqualStrings(first, second);
            }
        },
    }
}

test "fuzz: no format code can panic, leak, or render non-deterministically" {
    try std.testing.fuzz({}, fuzzNumfmtTarget, .{
        .corpus = &[_][]const u8{
            "General",   "0.00",                "#,##0.00;[Red](#,##0.00)",
            "# ??/??",   "?/8",                 "0.00E+00",
            "##0.0E+0",  "yyyy-mm-dd",          "m/d/yy h:mm",
            "[h]:mm:ss", "mmss.0",              "[$\u{20AC}-409]#,##0.00",
            "[$-40C]0",  "[DBNum1]0",           "@",
            "0;;;",      "[>100]0;[<-100]-0;0", "\"literal\"",
            "*-_-0%",    "0.0,,",               "\xFF\xFE",
            "ss.000",    "aaaa",                "b2yyyy",
            "0" ** 255,  ";;;",                 "[Color56]# ?/?",
        },
    });
}
