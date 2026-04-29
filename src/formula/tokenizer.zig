//! Formula tokenizer + loss-preserving printer (C1 milestone 1).
//!
//! Walks an Excel A1-style formula expression and emits a flat
//! stream of tokens whose `.text` slices borrow from the input.
//! `format(tokens) == input` byte-for-byte for any input the
//! tokenizer accepts — that round-trip property is what lets the
//! later C1 rewriter mutate refs without disturbing surrounding
//! whitespace, comments, locale-specific separators, or operator
//! formatting.
//!
//! Scope: A1 refs (absolute / mixed / relative), sheet qualifiers
//! (bare + quoted), ranges, names (functions + named ranges
//! collapse to one kind; the parser disambiguates by lookahead),
//! number / string / bool / error literals, every Excel binary
//! and unary operator, array constants, and arbitrary whitespace.
//! R1C1 notation, structured table refs, 3D refs, dynamic-array
//! operators (#, @), and external-workbook brackets are out of
//! scope for milestone 1; bytes that don't match any rule fall
//! through as `.unknown` so round-trip stays lossless.
//!
//! Public API:
//!   tokenize(allocator, text) -> []Token
//!   format(allocator, tokens) -> []u8

const std = @import("std");

pub const Token = struct {
    kind: Kind,
    /// Borrowed slice into the tokenizer's input buffer.
    text: []const u8,

    pub const Kind = enum {
        // References + identifiers
        cell_ref, // A1, $A$1, $A1, A$1, AA10
        name, // function name OR named range; parser decides via lookahead for `(`
        sheet_name, // 'Quoted Sheet'
        // Literals
        number, // 1, 1.5, .5, 1.5e10, 1E+5
        string, // "..." with "" escape
        bool_lit, // TRUE / FALSE (case-insensitive)
        error_lit, // #N/A, #REF!, #DIV/0!, #NAME?, #NUM!, #VALUE!, #NULL!, #GETTING_DATA, #SPILL!, #CALC!
        // Operators
        op_plus,
        op_minus,
        op_mul,
        op_div,
        op_pow,
        op_concat,
        op_eq,
        op_ne,
        op_lt,
        op_gt,
        op_le,
        op_ge,
        op_percent,
        op_range, // :
        // Punctuation
        paren_open,
        paren_close,
        array_open, // {
        array_close, // }
        arg_sep, // , or ;
        bang, // ! (sheet qualifier separator)
        whitespace,
        /// Bytes that don't match any rule. Preserved verbatim so
        /// callers can round-trip even partially malformed inputs;
        /// rewriters MUST NOT mutate `.unknown` tokens.
        unknown,
    };
};

pub const Error = error{OutOfMemory};

/// Tokenize a formula expression. The returned slice is owned by
/// the caller; each token's `.text` borrows from `input` and stays
/// valid as long as `input` does. Free with `allocator.free(out)`.
pub fn tokenize(allocator: std.mem.Allocator, input: []const u8) Error![]Token {
    var out: std.ArrayListUnmanaged(Token) = .{};
    errdefer out.deinit(allocator);

    var i: usize = 0;
    while (i < input.len) {
        const start = i;
        const c = input[i];
        const tok: Token = switch (c) {
            ' ', '\t', '\n', '\r' => blk: {
                while (i < input.len and isWhitespace(input[i])) : (i += 1) {}
                break :blk .{ .kind = .whitespace, .text = input[start..i] };
            },
            '"' => blk: {
                i = scanString(input, start);
                break :blk .{ .kind = .string, .text = input[start..i] };
            },
            '\'' => blk: {
                i = scanQuotedSheet(input, start);
                break :blk .{ .kind = .sheet_name, .text = input[start..i] };
            },
            '#' => blk: {
                if (matchKnownError(input, start)) |end| {
                    i = end;
                    break :blk .{ .kind = .error_lit, .text = input[start..end] };
                }
                // Anything else starting with `#` (the dynamic-array
                // spill operator `A1#`, an unrecognized error name)
                // is treated as `.unknown` so the rewriter knows
                // not to touch it. Round-trip is still preserved.
                i += 1;
                break :blk .{ .kind = .unknown, .text = input[start..i] };
            },
            '0'...'9' => blk: {
                i = scanNumber(input, start);
                break :blk .{ .kind = .number, .text = input[start..i] };
            },
            '.' => blk: {
                if (start + 1 < input.len and isDigit(input[start + 1])) {
                    i = scanNumber(input, start);
                    break :blk .{ .kind = .number, .text = input[start..i] };
                }
                i += 1;
                break :blk .{ .kind = .unknown, .text = input[start..i] };
            },
            '$', 'A'...'Z', 'a'...'z', '_' => blk: {
                i = scanIdent(input, start);
                const lex = input[start..i];
                if (isCellRef(lex)) break :blk .{ .kind = .cell_ref, .text = lex };
                if (isBoolLit(lex)) break :blk .{ .kind = .bool_lit, .text = lex };
                break :blk .{ .kind = .name, .text = lex };
            },
            '+' => blk: {
                i += 1;
                break :blk .{ .kind = .op_plus, .text = input[start..i] };
            },
            '-' => blk: {
                i += 1;
                break :blk .{ .kind = .op_minus, .text = input[start..i] };
            },
            '*' => blk: {
                i += 1;
                break :blk .{ .kind = .op_mul, .text = input[start..i] };
            },
            '/' => blk: {
                i += 1;
                break :blk .{ .kind = .op_div, .text = input[start..i] };
            },
            '^' => blk: {
                i += 1;
                break :blk .{ .kind = .op_pow, .text = input[start..i] };
            },
            '&' => blk: {
                i += 1;
                break :blk .{ .kind = .op_concat, .text = input[start..i] };
            },
            '%' => blk: {
                i += 1;
                break :blk .{ .kind = .op_percent, .text = input[start..i] };
            },
            ':' => blk: {
                i += 1;
                break :blk .{ .kind = .op_range, .text = input[start..i] };
            },
            ',', ';' => blk: {
                i += 1;
                break :blk .{ .kind = .arg_sep, .text = input[start..i] };
            },
            '(' => blk: {
                i += 1;
                break :blk .{ .kind = .paren_open, .text = input[start..i] };
            },
            ')' => blk: {
                i += 1;
                break :blk .{ .kind = .paren_close, .text = input[start..i] };
            },
            '{' => blk: {
                i += 1;
                break :blk .{ .kind = .array_open, .text = input[start..i] };
            },
            '}' => blk: {
                i += 1;
                break :blk .{ .kind = .array_close, .text = input[start..i] };
            },
            '!' => blk: {
                i += 1;
                break :blk .{ .kind = .bang, .text = input[start..i] };
            },
            '=' => blk: {
                i += 1;
                break :blk .{ .kind = .op_eq, .text = input[start..i] };
            },
            '<' => blk: {
                if (start + 1 < input.len and input[start + 1] == '=') {
                    i += 2;
                    break :blk .{ .kind = .op_le, .text = input[start..i] };
                }
                if (start + 1 < input.len and input[start + 1] == '>') {
                    i += 2;
                    break :blk .{ .kind = .op_ne, .text = input[start..i] };
                }
                i += 1;
                break :blk .{ .kind = .op_lt, .text = input[start..i] };
            },
            '>' => blk: {
                if (start + 1 < input.len and input[start + 1] == '=') {
                    i += 2;
                    break :blk .{ .kind = .op_ge, .text = input[start..i] };
                }
                i += 1;
                break :blk .{ .kind = .op_gt, .text = input[start..i] };
            },
            else => blk: {
                i += 1;
                break :blk .{ .kind = .unknown, .text = input[start..i] };
            },
        };
        try out.append(allocator, tok);
    }

    return out.toOwnedSlice(allocator);
}

/// Concatenate `tokens` back into the source text. With a faithful
/// tokenize() result this returns a byte-for-byte copy of the
/// original input. Caller frees the returned slice.
pub fn format(allocator: std.mem.Allocator, tokens: []const Token) Error![]u8 {
    var total: usize = 0;
    for (tokens) |t| total += t.text.len;
    const out = try allocator.alloc(u8, total);
    var p: usize = 0;
    for (tokens) |t| {
        @memcpy(out[p .. p + t.text.len], t.text);
        p += t.text.len;
    }
    return out;
}

// ─── lex helpers ─────────────────────────────────────────────────

inline fn isWhitespace(c: u8) bool {
    return c == ' ' or c == '\t' or c == '\n' or c == '\r';
}

inline fn isDigit(c: u8) bool {
    return c >= '0' and c <= '9';
}

inline fn isAsciiAlpha(c: u8) bool {
    return (c >= 'A' and c <= 'Z') or (c >= 'a' and c <= 'z');
}

inline fn isIdentByte(c: u8) bool {
    // Excel name characters: letters (any case), digits, `_`, `.`,
    // `?`, `\`. The `$` is treated separately (only valid as a
    // cell-ref absolute marker, not in named ranges).
    return isAsciiAlpha(c) or isDigit(c) or c == '_' or c == '.' or c == '?' or c == '\\';
}

/// Read past the leading `"` and return the index of the byte after
/// the closing `"`. Doubled `""` is an escaped quote; embedded
/// newlines are tolerated. An unterminated string consumes to EOF.
fn scanString(input: []const u8, start: usize) usize {
    var i = start + 1;
    while (i < input.len) : (i += 1) {
        if (input[i] == '"') {
            if (i + 1 < input.len and input[i + 1] == '"') {
                i += 1; // skip the escaped quote; the outer +=1 advances past the second
                continue;
            }
            return i + 1;
        }
    }
    return input.len;
}

/// Same shape as scanString but for `'...'` quoted sheet names.
fn scanQuotedSheet(input: []const u8, start: usize) usize {
    var i = start + 1;
    while (i < input.len) : (i += 1) {
        if (input[i] == '\'') {
            if (i + 1 < input.len and input[i + 1] == '\'') {
                i += 1;
                continue;
            }
            return i + 1;
        }
    }
    return input.len;
}

/// Match a known Excel error literal at `start` (where `input[start]`
/// is `#`). Returns the index past the literal on hit, null on miss.
/// Strict whitelist — anything else starting with `#` (the dynamic-
/// array spill operator `A1#`, future operators, malformed bytes)
/// is left to the caller to classify as `.unknown`.
///
/// The recognised set is the seven canonical errors plus the four
/// dynamic-array errors Excel 365 introduced. Listed in length-
/// descending order so the longest match wins (e.g. `#GETTING_DATA`
/// before any single-letter prefix).
fn matchKnownError(input: []const u8, start: usize) ?usize {
    const literals = [_][]const u8{
        "#GETTING_DATA",
        "#DIV/0!",
        "#VALUE!",
        "#NAME?",
        "#NULL!",
        "#SPILL!",
        "#CALC!",
        "#REF!",
        "#NUM!",
        "#N/A",
    };
    inline for (literals) |lit| {
        if (start + lit.len <= input.len and
            std.mem.eql(u8, input[start .. start + lit.len], lit))
        {
            return start + lit.len;
        }
    }
    return null;
}

/// Excel number grammar: `[0-9]+(\.[0-9]+)?([eE][+-]?[0-9]+)?` or
/// `\.[0-9]+([eE][+-]?[0-9]+)?`. Caller guarantees `input[start]`
/// is `0-9` or `.` followed by a digit.
fn scanNumber(input: []const u8, start: usize) usize {
    var i = start;
    // Integer part (may be absent only if we're entering at `.`).
    while (i < input.len and isDigit(input[i])) : (i += 1) {}
    // Optional fractional part.
    if (i < input.len and input[i] == '.') {
        i += 1;
        while (i < input.len and isDigit(input[i])) : (i += 1) {}
    }
    // Optional exponent.
    if (i < input.len and (input[i] == 'e' or input[i] == 'E')) {
        const exp_start = i;
        var j = i + 1;
        if (j < input.len and (input[j] == '+' or input[j] == '-')) j += 1;
        // The exponent must have at least one digit; otherwise the
        // `e`/`E` is part of an identifier (e.g. `EXP`), not a
        // number. Roll back if no digit follows.
        if (j < input.len and isDigit(input[j])) {
            i = j;
            while (i < input.len and isDigit(input[i])) : (i += 1) {}
        } else {
            i = exp_start;
        }
    }
    return i;
}

/// Read identifier characters. Includes a leading `$` (cell-ref
/// absolute marker) and an inner `$` (between letters and digits
/// in `$A$1`). Stops at any byte that can't appear in a name or
/// cell ref.
fn scanIdent(input: []const u8, start: usize) usize {
    var i = start;
    if (i < input.len and input[i] == '$') i += 1;
    // First name char must be a letter or `_` for non-ref names;
    // for cell refs the leading char is a letter (after optional
    // `$`). We accept both here and let isCellRef classify.
    while (i < input.len) {
        const c = input[i];
        if (c == '$') {
            // Inner `$` is only valid as a cell-ref marker between
            // column letters and the row number. Accept eagerly;
            // isCellRef enforces the structural shape.
            i += 1;
            continue;
        }
        if (isIdentByte(c)) {
            i += 1;
            continue;
        }
        break;
    }
    return i;
}

/// True if `s` is a syntactically valid A1-style cell ref:
/// `\$?[A-Za-z]+\$?[0-9]+` with column letters (case-insensitive)
/// in [A, XFD] and row in [1, 1048576]. Excel treats `a1` and `A1`
/// as the same reference; openpyxl-authored formulas occasionally
/// emit lowercase, so we accept both to spare callers a normalize
/// step the future rewriter would otherwise need.
fn isCellRef(s: []const u8) bool {
    if (s.len == 0) return false;
    var i: usize = 0;
    if (s[i] == '$') i += 1;
    const col_start = i;
    while (i < s.len and isAsciiAlpha(s[i])) : (i += 1) {}
    const col_len = i - col_start;
    if (col_len == 0 or col_len > 3) return false;
    if (i < s.len and s[i] == '$') i += 1;
    const row_start = i;
    while (i < s.len and isDigit(s[i])) : (i += 1) {}
    const row_len = i - row_start;
    if (row_len == 0 or row_len > 7) return false;
    if (i != s.len) return false;
    return columnInRange(s[col_start .. col_start + col_len]) and
        rowInRange(s[row_start .. row_start + row_len]);
}

fn columnInRange(letters: []const u8) bool {
    // Convert A=1, AA=27, ... up to XFD=16384. Case-insensitive.
    var v: u32 = 0;
    for (letters) |c| {
        const upper: u8 = if (c >= 'a' and c <= 'z') c - ('a' - 'A') else c;
        v = v * 26 + @as(u32, upper - 'A' + 1);
    }
    return v >= 1 and v <= 16384;
}

fn rowInRange(digits: []const u8) bool {
    var v: u64 = 0;
    for (digits) |c| {
        v = v * 10 + @as(u64, c - '0');
        if (v > 1_048_576) return false;
    }
    return v >= 1 and v <= 1_048_576;
}

fn isBoolLit(s: []const u8) bool {
    return std.ascii.eqlIgnoreCase(s, "TRUE") or std.ascii.eqlIgnoreCase(s, "FALSE");
}

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

fn expectKinds(input: []const u8, expected: []const Token.Kind) !void {
    const tokens = try tokenize(testing.allocator, input);
    defer testing.allocator.free(tokens);
    if (tokens.len != expected.len) {
        std.debug.print("expected {d} tokens, got {d} for input '{s}':\n", .{ expected.len, tokens.len, input });
        for (tokens) |t| std.debug.print("  {s} = '{s}'\n", .{ @tagName(t.kind), t.text });
        return error.TestExpectedEqual;
    }
    for (tokens, 0..) |t, i| {
        if (t.kind != expected[i]) {
            std.debug.print("token {d}: expected {s}, got {s} ('{s}') for input '{s}'\n", .{ i, @tagName(expected[i]), @tagName(t.kind), t.text, input });
            return error.TestExpectedEqual;
        }
    }
}

fn expectRoundTrip(input: []const u8) !void {
    const tokens = try tokenize(testing.allocator, input);
    defer testing.allocator.free(tokens);
    const back = try format(testing.allocator, tokens);
    defer testing.allocator.free(back);
    try testing.expectEqualSlices(u8, input, back);
}

test "simple cell refs" {
    try expectKinds("A1", &.{.cell_ref});
    try expectKinds("$A$1", &.{.cell_ref});
    try expectKinds("$A1", &.{.cell_ref});
    try expectKinds("A$1", &.{.cell_ref});
    try expectKinds("XFD1048576", &.{.cell_ref});
    try expectKinds("AA10", &.{.cell_ref});
}

test "identifiers that look like cell refs but aren't" {
    // Column out of range (XFE > XFD).
    try expectKinds("XFE1", &.{.name});
    // Row out of range (1048577 > 1048576).
    try expectKinds("A1048577", &.{.name});
}

test "lowercase + mixed-case A1 refs classify as cell_ref" {
    // Excel cell refs are case-insensitive; openpyxl-authored
    // formulas occasionally emit lowercase. The tokenizer accepts
    // both so the rewriter doesn't need a normalize pass.
    try expectKinds("a1", &.{.cell_ref});
    try expectKinds("$b$2", &.{.cell_ref});
    try expectKinds("Aa10", &.{.cell_ref});
    try expectKinds("xfd1048576", &.{.cell_ref});
}

test "ranges" {
    try expectKinds("A1:B5", &.{ .cell_ref, .op_range, .cell_ref });
    try expectKinds("$A$1:$B$5", &.{ .cell_ref, .op_range, .cell_ref });
    // Column-only range (full-column ref) — `A:A` parses as
    // `name : name` because `A` alone isn't a cell ref. Round-trip
    // still preserved; the rewriter recognises the pattern.
    try expectKinds("A:A", &.{ .name, .op_range, .name });
}

test "sheet-qualified refs" {
    try expectKinds("Sheet1!A1", &.{ .name, .bang, .cell_ref });
    try expectKinds("'My Sheet'!A1", &.{ .sheet_name, .bang, .cell_ref });
    try expectKinds("'It''s'!A1", &.{ .sheet_name, .bang, .cell_ref });
}

test "function calls" {
    try expectKinds("SUM(A1:A10)", &.{
        .name, .paren_open, .cell_ref, .op_range, .cell_ref, .paren_close,
    });
    try expectKinds(
        "IF(A1>0,\"Yes\",\"No\")",
        &.{ .name, .paren_open, .cell_ref, .op_gt, .number, .arg_sep, .string, .arg_sep, .string, .paren_close },
    );
}

test "string literals with doubled-quote escape" {
    try expectKinds("\"hello\"", &.{.string});
    try expectKinds("\"a\"\"b\"", &.{.string});
    try expectRoundTrip("\"contains \"\" inside\"");
}

test "number literals" {
    try expectKinds("1", &.{.number});
    try expectKinds("1.5", &.{.number});
    try expectKinds(".5", &.{.number});
    try expectKinds("1.5e10", &.{.number});
    try expectKinds("1E+5", &.{.number});
    try expectKinds("1E-5", &.{.number});
    // Identifier that starts with `E` but isn't a number — `EXP` is
    // a function name, and the digit-less exponent rolls back.
    try expectKinds("EXP", &.{.name});
}

test "boolean literals" {
    try expectKinds("TRUE", &.{.bool_lit});
    try expectKinds("FALSE", &.{.bool_lit});
    try expectKinds("True", &.{.bool_lit});
    try expectKinds("false", &.{.bool_lit});
}

test "error literals — strict whitelist" {
    // Canonical seven.
    try expectKinds("#N/A", &.{.error_lit});
    try expectKinds("#REF!", &.{.error_lit});
    try expectKinds("#DIV/0!", &.{.error_lit});
    try expectKinds("#NAME?", &.{.error_lit});
    try expectKinds("#NUM!", &.{.error_lit});
    try expectKinds("#VALUE!", &.{.error_lit});
    try expectKinds("#NULL!", &.{.error_lit});
    // Excel 365 dynamic-array errors.
    try expectKinds("#SPILL!", &.{.error_lit});
    try expectKinds("#CALC!", &.{.error_lit});
    try expectKinds("#GETTING_DATA", &.{.error_lit});
}

test "non-error # falls back to .unknown" {
    // Dynamic-array spill operator `A1#` is out of scope for
    // milestone 1; the `#` must classify as `.unknown` so the
    // rewriter can refuse to touch the formula. (Round-trip
    // still preserves the bytes verbatim.)
    try expectKinds("A1#", &.{ .cell_ref, .unknown });
    try expectRoundTrip("A1#");
    // Unknown error name `#FOO` — same treatment: just `#` as
    // unknown, then `FOO` as a name.
    try expectKinds("#FOO", &.{ .unknown, .name });
    try expectRoundTrip("#FOO");
}

test "operators" {
    try expectKinds("1+2", &.{ .number, .op_plus, .number });
    try expectKinds("1-2", &.{ .number, .op_minus, .number });
    try expectKinds("1*2", &.{ .number, .op_mul, .number });
    try expectKinds("1/2", &.{ .number, .op_div, .number });
    try expectKinds("1^2", &.{ .number, .op_pow, .number });
    try expectKinds("1&2", &.{ .number, .op_concat, .number });
    try expectKinds("1=2", &.{ .number, .op_eq, .number });
    try expectKinds("1<>2", &.{ .number, .op_ne, .number });
    try expectKinds("1<=2", &.{ .number, .op_le, .number });
    try expectKinds("1>=2", &.{ .number, .op_ge, .number });
    try expectKinds("1<2", &.{ .number, .op_lt, .number });
    try expectKinds("1>2", &.{ .number, .op_gt, .number });
    try expectKinds("50%", &.{ .number, .op_percent });
}

test "array constants" {
    try expectKinds(
        "{1,2,3;4,5,6}",
        &.{ .array_open, .number, .arg_sep, .number, .arg_sep, .number, .arg_sep, .number, .arg_sep, .number, .arg_sep, .number, .array_close },
    );
}

test "whitespace preserved" {
    try expectKinds("A1 + B2", &.{ .cell_ref, .whitespace, .op_plus, .whitespace, .cell_ref });
    try expectRoundTrip("  SUM ( A1 ,  B2 )  ");
}

test "round-trip preserves bytes verbatim" {
    const cases = [_][]const u8{
        "A1",
        "SUM(A1:A10)",
        "IF(A1>0,\"yes\",\"no\")",
        "Sheet1!A1+'Other Sheet'!$B$2",
        "{1,2;3,4}",
        "1.5e+10",
        "#REF!+#N/A",
        "MyName.Sub",
        "  ", // whitespace only
        "", // empty
        "A1:B2,C3:D4", // union of ranges
        "(A1:A10) (B5:F5)", // intersection (whitespace operator)
    };
    for (cases) |c| try expectRoundTrip(c);
}

test "unknown bytes are preserved" {
    // `@` isn't part of any milestone-1 production. Round-trip
    // still works: it falls into `.unknown` and the formatter
    // emits it back verbatim.
    try expectKinds("@A1", &.{ .unknown, .cell_ref });
    try expectRoundTrip("@A1");
}

test "named ranges and function calls share the .name kind" {
    // The tokenizer doesn't disambiguate; lookahead for `(` is the
    // parser's job. Both produce `.name` tokens.
    try expectKinds("MyRange", &.{.name});
    try expectKinds("SUM", &.{.name});
    try expectKinds("My_Range.v2", &.{.name});
}

test "mixed bag round-trip" {
    try expectRoundTrip("=IF(SUM($A$1:$A$10)>0,\"yes\",\"no\")");
    try expectRoundTrip("Sheet1!A1+Sheet2!$B$2-'Some Sheet'!C3");
    try expectRoundTrip("MIN(1,2,3) - MAX({1;2;3})");
}
