//! Formula tokenizer + loss-preserving printer.
//!
//! Walks an Excel A1-style formula expression and emits a flat
//! stream of tokens whose `.text` slices borrow from the input.
//! `format(tokens) == input` byte-for-byte for any input the
//! tokenizer accepts — that round-trip property is what lets the
//! C1 rewriter mutate refs without disturbing surrounding
//! whitespace, comments, locale-specific separators, or operator
//! formatting, and it is what the M2 parser will re-derive spans
//! from.
//!
//! Scope (M1a of the tier-D1 ladder; `goal_formula.md` §5.2)
//! -------------------------------------------------------
//! A1 refs (absolute / mixed / relative), sheet qualifiers (bare +
//! quoted), ranges, full-row / full-column / 3D shapes, names
//! (functions + named ranges collapse to one kind; the parser
//! disambiguates by lookahead), number / string / bool literals,
//! **extensible** error literals, every Excel binary and unary
//! operator, the dynamic-array operators `#` and `@`, structured
//! table references with their `'` escapes, external-workbook
//! references (recognised, then refused by kind so no caller can
//! mistake one for a live ref), array constants, and arbitrary
//! whitespace.
//!
//! Identifiers are Unicode. Excel names and sheet names are not
//! ASCII — `=Ω!Σ` and `=ДАННЫЕ` are ordinary formulas — so the
//! grammar is UAX #31 with Excel's additions:
//!
//!   start    = `XID_Start` ∪ {`_`, `\`}
//!   continue = `XID_Continue` ∪ {`_`, `.`}
//!
//! `\` is **start-only**: `a\b` is refused. Identifiers are capped
//! at 255 code points (Excel's defined-name limit) and invalid UTF-8
//! is refused. The tables live in `unicode/xid.zig`.
//!
//! R1C1 notation is recognised and refused — evaluation stays
//! A1-only. A clean whole atom (`R1C1`, `R[-1]C[2]`) additionally
//! merges into one `.r1c1_ref` token so the C1 rewriter can shift it
//! as a unit; unclean spellings keep the pre-merge fragment
//! tokenization byte-for-byte. LAMBDA/LET parameter prefixes
//! (`_xlpm.`) tokenize as ordinary names; the parser refuses them
//! at M2.
//!
//! Bytes that don't match any rule still fall through as `.unknown`
//! so round-trip stays lossless on partially malformed input.
//! Rewriters MUST NOT mutate `.unknown` or `.external_ref` tokens.
//! `.structured_ref` is opaque to every edit except
//! `rename_table_column`, which re-parses the specifier through the
//! parser's own grammar and replaces only the column-name subspans.
//!
//! Refusals vs round-trip
//! ----------------------
//! The tokenizer never fails on malformed input (only on OOM):
//! failing would break the round-trip contract that every downstream
//! byte-identity gate depends on. Constructs that a later stage must
//! reject are reported out-of-band as `Refusal`s, which `scan`
//! returns alongside the tokens. Each maps onto a §10 plane-2 error
//! at M2 — see `Refusal.Reason`.
//!
//! Public API:
//!   scan(allocator, text)     -> Scan { tokens, refusals }
//!   tokenize(allocator, text) -> []Token   (kinds only; no refusals)
//!   format(allocator, tokens) -> []u8

const std = @import("std");
const coords = @import("zlsx_refs");
const xid = @import("zlsx_xid");

pub const Token = struct {
    kind: Kind,
    /// Borrowed slice into the tokenizer's input buffer.
    text: []const u8,

    pub const Kind = enum {
        // References + identifiers
        cell_ref, // A1, $A$1, $A1, A$1, AA10
        name, // function name OR named range; parser decides via lookahead for `(`
        sheet_name, // 'Quoted Sheet'
        /// A `[...]` structured-table item specifier, brackets
        /// included: `[Col]`, `[#Headers]`, `[@Col]`, `[[Col A]:[Col B]]`.
        /// Opaque and single — the inner text is NOT tokenized, which
        /// is what stops a rewriter from mistaking `Table1[A1]`'s
        /// column name for a live cell reference.
        structured_ref,
        /// An external-workbook reference prefix: the quoted
        /// `'[Book.xlsx]Sheet1'` form or the `[1]` index form.
        /// Always accompanied by a `.external_reference` refusal.
        external_ref,
        /// A complete R1C1-style reference — `R1C1`, `R12C`, `R1C[2]`,
        /// `R[-1]C[2]`, `R[2]`, `C[2]` — as ONE token, bracketed
        /// relative parts included. Always accompanied by a
        /// `.r1c1_reference` refusal: evaluation still refuses the
        /// construct (v1 is A1-only). The kind exists for the C1
        /// rewriter, which must shift the atom as a unit — fragment
        /// scanning left its `C5`-shaped column part looking like a
        /// live A1 cell ref, which the A1 path then row-shifted.
        /// Decompose with `parseR1C1AtomText` — the same grammar the
        /// scanner merged with, never a second scanner.
        r1c1_ref,
        // Literals
        number, // 1, 1.5, .5, 1.5e10, 1E+5
        string, // "..." with "" escape
        bool_lit, // TRUE / FALSE (case-insensitive)
        /// `#N/A`, `#REF!`, … and any other `#…!` / `#…?` lexeme.
        /// Bytes preserved verbatim so rich and future error spellings
        /// round-trip; `isKnownErrorLiteral` separates the frozen ten.
        error_lit,
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
        /// Dynamic-array spill operator: the `#` in `A1#`.
        op_spill,
        /// Implicit-intersection operator `@` (the printed form of
        /// `_xlfn.SINGLE`).
        op_at,
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

/// Excel's defined-name length limit, in code points (not bytes).
pub const max_identifier_codepoints: usize = 255;

/// Upper bound on an extensible error literal, in bytes including the
/// leading `#` and the trailing `!`/`?`. Bounded so a stray `#` cannot
/// make the scanner run to a terminator hundreds of bytes away and
/// swallow a whole subexpression into one token.
pub const max_error_literal_bytes: usize = 64;

/// A construct the tokenizer recognised and preserved, but which a
/// later stage must reject. Reported out-of-band precisely so the
/// byte-for-byte round-trip survives.
pub const Refusal = struct {
    reason: Reason,
    /// Byte offset of the offending span within the tokenizer input.
    offset: usize,
    /// Byte length of the offending span.
    len: usize,

    /// Each variant names the §10 plane-2 error the M2 parser raises
    /// for it. The tokenizer classifies; it does not map — the
    /// mapping needs the diagnostic sink that arrives with the parser.
    pub const Reason = enum {
        /// → `FormulaMalformedInput`
        invalid_utf8,
        /// → `FormulaLimitExceeded` (255 code points)
        identifier_too_long,
        /// → `FormulaMalformedInput`: `\` starts identifiers, it does
        /// not continue them.
        backslash_after_start,
        /// → `FormulaUnsupportedConstruct`
        r1c1_reference,
        /// → `FormulaUnsupportedConstruct`
        external_reference,
        /// → `FormulaMalformedInput`
        unterminated_string,
        /// → `FormulaMalformedInput`
        unterminated_sheet_name,
        /// → `FormulaMalformedInput`
        unterminated_structured_ref,
    };
};

/// Tokens plus every refusal raised while producing them, in
/// detection order. Detection order is *mostly* left-to-right but not
/// guaranteed to be: a length refusal is attributed to the start of
/// the identifier it measures, which can precede a refusal raised
/// while scanning that same identifier.
pub const Scan = struct {
    tokens: []Token,
    refusals: []Refusal,

    pub fn deinit(self: *Scan, allocator: std.mem.Allocator) void {
        allocator.free(self.tokens);
        allocator.free(self.refusals);
        self.* = undefined;
    }

    /// First refusal with `reason`, or null.
    pub fn find(self: Scan, reason: Refusal.Reason) ?Refusal {
        for (self.refusals) |r| {
            if (r.reason == reason) return r;
        }
        return null;
    }
};

pub const Error = error{OutOfMemory};

/// Tokenize `input`, reporting refusals. The returned slices are owned
/// by the caller; each token's `.text` borrows from `input` and stays
/// valid as long as `input` does. Free with `scan_result.deinit(alloc)`.
pub fn scan(allocator: std.mem.Allocator, input: []const u8) Error!Scan {
    var s: Scanner = .{ .allocator = allocator, .input = input };
    errdefer {
        s.tokens.deinit(allocator);
        s.refusals.deinit(allocator);
    }
    while (s.i < input.len) try s.next();

    const refusals = try s.refusals.toOwnedSlice(allocator);
    errdefer allocator.free(refusals);
    return .{ .tokens = try s.tokens.toOwnedSlice(allocator), .refusals = refusals };
}

/// Tokenize `input`, discarding refusals. Kept as the pre-M1a
/// signature for callers that only consume kinds — the rewriter is
/// one: it mutates `.cell_ref` / `.sheet_name` and passes everything
/// else through untouched, so a refused construct is already inert to
/// it. Callers that must *reject* input want `scan` instead.
pub fn tokenize(allocator: std.mem.Allocator, input: []const u8) Error![]Token {
    const result = try scan(allocator, input);
    allocator.free(result.refusals);
    return result.tokens;
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

/// True if `text` is one of the ten error spellings Excel documents.
/// Everything else reaching `.error_lit` came in through the
/// extensible rule and must be preserved byte-exact rather than
/// interpreted.
pub fn isKnownErrorLiteral(text: []const u8) bool {
    for (known_error_literals) |lit| {
        if (std.mem.eql(u8, text, lit)) return true;
    }
    return false;
}

// ─── scanner ─────────────────────────────────────────────────────

const Scanner = struct {
    allocator: std.mem.Allocator,
    input: []const u8,
    i: usize = 0,
    tokens: std.ArrayListUnmanaged(Token) = .empty,
    refusals: std.ArrayListUnmanaged(Refusal) = .empty,

    fn emit(self: *Scanner, kind: Token.Kind, start: usize, end: usize) Error!void {
        self.i = end;
        try self.tokens.append(self.allocator, .{ .kind = kind, .text = self.input[start..end] });
    }

    fn refuse(self: *Scanner, reason: Refusal.Reason, offset: usize, len: usize) Error!void {
        // Coalesce a run of invalid UTF-8 bytes into one refusal —
        // a truncated multi-byte sequence would otherwise report once
        // per byte and bury every other diagnostic.
        if (reason == .invalid_utf8 and self.refusals.items.len > 0) {
            const last = &self.refusals.items[self.refusals.items.len - 1];
            if (last.reason == .invalid_utf8 and last.offset + last.len == offset) {
                last.len += len;
                return;
            }
        }
        try self.refusals.append(self.allocator, .{ .reason = reason, .offset = offset, .len = len });
    }

    fn next(self: *Scanner) Error!void {
        const input = self.input;
        const start = self.i;
        switch (input[start]) {
            ' ', '\t', '\n', '\r' => {
                var j = start;
                while (j < input.len and isWhitespace(input[j])) : (j += 1) {}
                try self.emit(.whitespace, start, j);
            },
            '"' => {
                const str = scanString(input, start);
                if (!str.terminated) {
                    try self.refuse(.unterminated_string, start, str.end - start);
                }
                try self.emit(.string, start, str.end);
            },
            '\'' => try self.quotedSheet(start),
            '#' => {
                if (matchErrorLiteral(input, start)) |end| {
                    try self.emit(.error_lit, start, end);
                } else {
                    // No terminator and no whitelist hit: this is the
                    // dynamic-array spill operator. Pre-M1a it was
                    // `.unknown`; the kind is the correction, the bytes
                    // are unchanged.
                    try self.emit(.op_spill, start, start + 1);
                }
            },
            '@' => try self.emit(.op_at, start, start + 1),
            '[' => try self.bracketed(start),
            '0'...'9' => try self.emit(.number, start, scanNumber(input, start)),
            '.' => {
                if (start + 1 < input.len and isDigit(input[start + 1])) {
                    try self.emit(.number, start, scanNumber(input, start));
                } else {
                    try self.emit(.unknown, start, start + 1);
                }
            },
            '$', 'A'...'Z', 'a'...'z', '_', '\\' => try self.identifier(start),
            '+' => try self.emit(.op_plus, start, start + 1),
            '-' => try self.emit(.op_minus, start, start + 1),
            '*' => try self.emit(.op_mul, start, start + 1),
            '/' => try self.emit(.op_div, start, start + 1),
            '^' => try self.emit(.op_pow, start, start + 1),
            '&' => try self.emit(.op_concat, start, start + 1),
            '%' => try self.emit(.op_percent, start, start + 1),
            ':' => try self.emit(.op_range, start, start + 1),
            ',', ';' => try self.emit(.arg_sep, start, start + 1),
            '(' => try self.emit(.paren_open, start, start + 1),
            ')' => try self.emit(.paren_close, start, start + 1),
            '{' => try self.emit(.array_open, start, start + 1),
            '}' => try self.emit(.array_close, start, start + 1),
            ']' => try self.emit(.unknown, start, start + 1), // unmatched; `[` consumes matched pairs
            '!' => try self.emit(.bang, start, start + 1),
            '=' => try self.emit(.op_eq, start, start + 1),
            '<' => {
                if (start + 1 < input.len and input[start + 1] == '=') {
                    try self.emit(.op_le, start, start + 2);
                } else if (start + 1 < input.len and input[start + 1] == '>') {
                    try self.emit(.op_ne, start, start + 2);
                } else {
                    try self.emit(.op_lt, start, start + 1);
                }
            },
            '>' => {
                if (start + 1 < input.len and input[start + 1] == '=') {
                    try self.emit(.op_ge, start, start + 2);
                } else {
                    try self.emit(.op_gt, start, start + 1);
                }
            },
            else => |c| {
                if (c < 0x80) {
                    try self.emit(.unknown, start, start + 1);
                    return;
                }
                // Non-ASCII. A codepoint that can start an identifier
                // enters the identifier path; anything else (emoji,
                // NBSP, currency symbols) is preserved as `.unknown`
                // one whole codepoint at a time, so a multi-byte
                // sequence is never split across tokens.
                const cp = decodeAt(input, start) orelse {
                    try self.refuse(.invalid_utf8, start, 1);
                    try self.emit(.unknown, start, start + 1);
                    return;
                };
                if (xid.isXidStart(cp.value)) {
                    try self.identifier(start);
                } else {
                    try self.emit(.unknown, start, start + cp.len);
                }
            },
        }
    }

    fn quotedSheet(self: *Scanner, start: usize) Error!void {
        const input = self.input;
        const quoted = scanQuotedSheet(input, start);
        const end = quoted.end;
        const lex = input[start..end];
        if (!quoted.terminated) {
            try self.refuse(.unterminated_sheet_name, start, end - start);
        }
        // External-workbook refs (`'[Book.xlsx]Sheet1'!A1`) carry the
        // workbook in brackets inside the quotes. Typed as
        // `.external_ref` so no caller treats the sheet or the range
        // that follows as local.
        if (std.mem.indexOfScalar(u8, lex, '[') != null) {
            try self.refuse(.external_reference, start, end - start);
            try self.emit(.external_ref, start, end);
            return;
        }
        try self.emit(.sheet_name, start, end);
    }

    fn bracketed(self: *Scanner, start: usize) Error!void {
        const input = self.input;
        const br = scanBracketed(input, start);
        if (!br.terminated) {
            try self.refuse(.unterminated_structured_ref, start, br.end - start);
            try self.emit(.unknown, start, br.end);
            return;
        }
        if (isExternalIndex(input, start, br.end)) {
            try self.refuse(.external_reference, start, br.end - start);
            try self.emit(.external_ref, start, br.end);
            return;
        }
        try self.emit(.structured_ref, start, br.end);
    }

    /// Scan and classify an identifier-shaped lexeme.
    ///
    /// Classification precedence is normative (§5.2): cell reference →
    /// R1C1-shape rejection → TRUE/FALSE → function/name. The
    /// call/qualifier lookahead runs ahead of all four: a name-shaped
    /// lexeme followed by `(` is always a function call (that flips
    /// `TRUE(` from a literal into the zero-argument TRUE function, and
    /// `LOG10(` from an A1 ref into the LOG10 function), and a trailing
    /// `!` means a sheet qualifier.
    fn identifier(self: *Scanner, start: usize) Error!void {
        const input = self.input;
        const end = try self.scanIdentifier(start);
        const lex = input[start..end];

        const followed_by_call = end < input.len and input[end] == '(';
        const followed_by_qual = end < input.len and input[end] == '!';
        if (followed_by_call or followed_by_qual) {
            try self.emit(.name, start, end);
            return;
        }
        if (isCellRef(lex)) {
            try self.emit(.cell_ref, start, end);
            return;
        }
        if (isR1C1Shape(lex)) {
            // A clean whole atom merges into one `.r1c1_ref` — which
            // may run past the lexeme when a bare col letter grows a
            // bracket (`R1C` + `[2]`). Anything unclean keeps the
            // pre-merge fragment behavior byte-for-byte.
            if (scanR1C1Atom(input, start)) |scanned| {
                try self.refuse(.r1c1_reference, start, scanned.end - start);
                try self.emit(.r1c1_ref, start, scanned.end);
                return;
            }
            try self.refuse(.r1c1_reference, start, end - start);
            try self.emit(.name, start, end);
            return;
        }
        // The bracketed relative forms (`R[-1]C`, `C[2]`) split at `[`,
        // so the lexeme alone is just `R` or `C` — which are also legal
        // full-column references (`R:R`). Only the trailing bracket
        // tells them apart.
        if (lex.len == 1 and isRowColLetter(lex[0]) and
            end < input.len and input[end] == '[')
        {
            if (scanR1C1Atom(input, start)) |scanned| {
                try self.refuse(.r1c1_reference, start, scanned.end - start);
                try self.emit(.r1c1_ref, start, scanned.end);
                return;
            }
            try self.refuse(.r1c1_reference, start, end - start);
        }
        if (isBoolLit(lex)) {
            try self.emit(.bool_lit, start, end);
            return;
        }
        try self.emit(.name, start, end);
    }

    /// Read one identifier lexeme, returning the end index. Records
    /// the length and backslash-placement refusals; classification is
    /// the caller's.
    fn scanIdentifier(self: *Scanner, start: usize) Error!usize {
        const input = self.input;
        var i = start;
        var codepoints: usize = 0;

        const leading_dollar = input[i] == '$';
        if (leading_dollar) {
            i += 1;
            codepoints += 1;
        }

        if (leading_dollar and i < input.len and isDigit(input[i])) {
            // `$1` / `$1048576` — the absolute form of a full-row
            // reference (`$1:$1`). Digits cannot *start* an identifier,
            // so without this branch the `$` would strand as a lone
            // lexeme and split a construct that must stay whole.
            while (i < input.len and isDigit(input[i])) : (i += 1) codepoints += 1;
        } else {
            var at_start = true;
            while (i < input.len) {
                const c = input[i];
                if (c == '$') {
                    // Inner `$` is only valid as a cell-ref marker
                    // between the column letters and the row number.
                    // Accept eagerly; `isCellRef` enforces the
                    // structural shape. It is not an identifier
                    // character, so it neither opens nor closes the
                    // start position.
                    i += 1;
                    codepoints += 1;
                    continue;
                }
                if (c < 0x80) {
                    const ok = if (at_start) isAsciiIdentStart(c) else isAsciiIdentContinue(c);
                    if (!ok) {
                        if (c == '\\' and !at_start) {
                            // `\` starts identifiers, it does not
                            // continue them: `\Foo` is a name, `a\b`
                            // is not.
                            try self.refuse(.backslash_after_start, i, 1);
                        }
                        break;
                    }
                    i += 1;
                    codepoints += 1;
                    at_start = false;
                    continue;
                }
                // A malformed sequence just ends the identifier; the
                // main dispatch reaches the same byte next and raises
                // the single `invalid_utf8` refusal for it.
                const cp = decodeAt(input, i) orelse break;
                const ok = if (at_start) xid.isXidStart(cp.value) else xid.isXidContinue(cp.value);
                if (!ok) break;
                i += cp.len;
                codepoints += 1;
                at_start = false;
            }
        }

        if (codepoints > max_identifier_codepoints) {
            try self.refuse(.identifier_too_long, start, i - start);
        }
        return i;
    }
};

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

inline fn isRowColLetter(c: u8) bool {
    return c == 'R' or c == 'r' or c == 'C' or c == 'c';
}

/// ASCII half of `XID_Start ∪ {_, \}`. Kept as a byte test rather
/// than a table probe because it is the hot path — a test asserts it
/// agrees with `zlsx_xid` across all 128 ASCII codepoints.
inline fn isAsciiIdentStart(c: u8) bool {
    return isAsciiAlpha(c) or c == '_' or c == '\\';
}

/// ASCII half of `XID_Continue ∪ {_, .}`. Note `?` is absent: Excel
/// does not accept it in defined names, and the pre-M1a predicate that
/// did was one of the corrections this milestone lands.
inline fn isAsciiIdentContinue(c: u8) bool {
    return isAsciiAlpha(c) or isDigit(c) or c == '_' or c == '.';
}

const Decoded = struct { value: u21, len: usize };

/// Decode the UTF-8 sequence at `input[i]`. Null on any malformation
/// — truncated tail, bad continuation byte, overlong encoding, or a
/// surrogate — all of which `utf8Decode` already rejects.
fn decodeAt(input: []const u8, i: usize) ?Decoded {
    const len = std.unicode.utf8ByteSequenceLength(input[i]) catch return null;
    if (i + len > input.len) return null;
    const cp = std.unicode.utf8Decode(input[i .. i + len]) catch return null;
    return .{ .value = cp, .len = len };
}

const Delimited = struct { end: usize, terminated: bool };

/// Read past the leading `"` and return the index of the byte after
/// the closing `"`. Doubled `""` is an escaped quote; embedded
/// newlines are tolerated. An unterminated string consumes to EOF and
/// reports `terminated = false` — `"""` is unterminated even though it
/// ends in a quote, because the middle pair is an escape.
fn scanString(input: []const u8, start: usize) Delimited {
    return scanDelimited(input, start, '"');
}

/// Same shape as scanString but for `'...'` quoted sheet names.
fn scanQuotedSheet(input: []const u8, start: usize) Delimited {
    return scanDelimited(input, start, '\'');
}

fn scanDelimited(input: []const u8, start: usize, quote: u8) Delimited {
    var i = start + 1;
    while (i < input.len) : (i += 1) {
        if (input[i] == quote) {
            if (i + 1 < input.len and input[i + 1] == quote) {
                i += 1; // skip the escaped quote; the outer +=1 advances past the second
                continue;
            }
            return .{ .end = i + 1, .terminated = true };
        }
    }
    return .{ .end = input.len, .terminated = false };
}

const Bracketed = struct { end: usize, terminated: bool };

/// Scan a `[...]` span from its opening bracket. Nested brackets are
/// tracked by depth (`[[Col A]:[Col B]]` is one specifier), and `'`
/// escapes the byte after it — that is how a column name containing
/// `[`, `]`, `#`, or `@` is written.
fn scanBracketed(input: []const u8, start: usize) Bracketed {
    var depth: usize = 0;
    var i = start;
    while (i < input.len) {
        switch (input[i]) {
            '\'' => i = @min(i + 2, input.len),
            '[' => {
                depth += 1;
                i += 1;
            },
            ']' => {
                i += 1;
                depth -= 1;
                if (depth == 0) return .{ .end = i, .terminated = true };
            },
            else => i += 1,
        }
    }
    return .{ .end = input.len, .terminated = false };
}

/// True if `input[start..end]` is the `[1]` workbook-index prefix of an
/// external reference rather than a structured-table specifier. The
/// index form is all digits AND is immediately followed by the sheet
/// or defined name it qualifies — a bare `[1]` with nothing after it
/// stays a structured ref, since that is what a table column named `1`
/// would look like.
fn isExternalIndex(input: []const u8, start: usize, end: usize) bool {
    const inner = input[start + 1 .. end - 1];
    if (inner.len == 0) return false;
    for (inner) |c| {
        if (!isDigit(c)) return false;
    }
    if (end >= input.len) return false;
    const c = input[end];
    return isAsciiIdentStart(c) or c == '\'' or c == '!' or c >= 0x80;
}

/// The ten error spellings Excel documents, in length-descending order
/// so the longest match wins (`#GETTING_DATA` before any prefix).
const known_error_literals = [_][]const u8{
    "#GETTING_DATA",
    "#DIV/0!",
    "#VALUE!",
    "#SPILL!",
    "#NAME?",
    "#NULL!",
    "#CALC!",
    "#REF!",
    "#NUM!",
    "#N/A",
};

/// Match an error literal at `start` (where `input[start]` is `#`).
/// Returns the index past the literal on hit, null on miss.
///
/// Two rules, tried in order:
///
///  1. the frozen ten above — these include `#N/A` and
///     `#GETTING_DATA`, which carry no terminator;
///  2. the **extensible** rule: `#` + a bounded run of
///     `[A-Za-z0-9_/.]` + `!` or `?`. Excel keeps adding errors
///     (`#SPILL!`, `#CALC!`, `#BLOCKED!`, `#PYTHON!`, `#BUSY!`…) and
///     a closed whitelist silently mangles every future one into
///     operator-plus-name debris. Unknown spellings tokenize whole and
///     round-trip byte-exact; `isKnownErrorLiteral` is how a caller
///     asks whether it may interpret one.
///
/// A `#` matching neither rule is the spill operator.
fn matchErrorLiteral(input: []const u8, start: usize) ?usize {
    inline for (known_error_literals) |lit| {
        if (start + lit.len <= input.len and
            std.mem.eql(u8, input[start .. start + lit.len], lit))
        {
            return start + lit.len;
        }
    }

    const limit = @min(input.len, start + max_error_literal_bytes);
    var i = start + 1;
    while (i < limit) : (i += 1) {
        const c = input[i];
        if (c == '!' or c == '?') {
            if (i == start + 1) return null; // `#!` — no body
            return i + 1;
        }
        if (!isErrorBodyByte(c)) return null;
    }
    return null;
}

inline fn isErrorBodyByte(c: u8) bool {
    return isAsciiAlpha(c) or isDigit(c) or c == '_' or c == '/' or c == '.';
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

/// True if `s` is a syntactically valid A1-style cell ref:
/// `\$?[A-Za-z]+\$?[0-9]+` with column letters (case-insensitive)
/// in [A, XFD] and row in [1, 1048576]. Excel treats `a1` and `A1`
/// as the same reference; openpyxl-authored formulas occasionally
/// emit lowercase, so we accept both to spare callers a normalize
/// step the rewriter would otherwise need.
fn isCellRef(s: []const u8) bool {
    if (s.len == 0) return false;
    var i: usize = 0;
    if (s[i] == '$') i += 1;
    const col_start = i;
    while (i < s.len and isAsciiAlpha(s[i])) : (i += 1) {}
    const col_len = i - col_start;
    if (col_len == 0 or col_len > coords.max_col_letters) return false;
    if (i < s.len and s[i] == '$') i += 1;
    const row_start = i;
    while (i < s.len and isDigit(s[i])) : (i += 1) {}
    const row_len = i - row_start;
    if (row_len == 0 or row_len > 7) return false;
    if (i != s.len) return false;
    return columnInRange(s[col_start .. col_start + col_len]) and
        rowInRange(s[row_start .. row_start + row_len]);
}

/// One axis component of an R1C1 atom. The three forms are spelling,
/// not just value — the rewriter preserves each part's form (`R5`
/// stays digits, `R[5]` stays bracketed, a trailing bare `C` stays
/// bare until its offset has to change).
pub const R1C1Part = struct {
    pub const Form = enum {
        /// `R5` — absolute 1-based position in `value`. The scanner
        /// checks shape only, not grid bounds: `R0C1` carries 0.
        digits,
        /// `R[-2]` — relative offset in `value`.
        bracket,
        /// Trailing bare `R`/`C` — relative offset 0. `value` is 0.
        bare,
    };
    form: Form,
    value: i64,
};

/// A decomposed `.r1c1_ref` token. Exactly one of the two single-part
/// spellings the scanner can produce is null-row (`C[2]`); `R[2]` is
/// null-col. Both parts present for the cell forms.
pub const R1C1Atom = struct {
    row: ?R1C1Part,
    col: ?R1C1Part,
    /// Byte offset (within the atom's text) where the col part
    /// begins; equals the text length when `col` is null. Lets the
    /// rewriter splice one axis while keeping the other axis's bytes
    /// — spelling, letter case — verbatim.
    col_start: usize,
};

const ScannedR1C1 = struct { atom: R1C1Atom, end: usize };

/// The one R1C1 grammar. The scanner calls it to decide whether the
/// bytes at `start` form a complete, well-formed atom to merge into a
/// single `.r1c1_ref` token; the rewriter calls `parseR1C1AtomText`
/// (below) to decompose that token's text. One implementation on both
/// sides is what guarantees they can never disagree (#188 r6: never
/// mirror a parser with a second scanner).
///
/// Accepted atoms — the reachable-and-emittable set, nothing more:
///   - row digits + col part:      `R1C1`, `R5C`, `R1C[2]`
///   - row bracket + optional col: `R[2]`, `R[1]C`, `R[1]C2`, `R[-1]C[2]`
///   - col bracket alone:          `C[2]`
/// Rejected (returns null → caller falls back to the pre-merge
/// fragment behavior): a bare row letter with no bracket (`RC`,
/// `RC4` — those read as A1), digits that run past 7, a malformed
/// bracket body (`R[foo]`, `R[]`, unterminated), a digits part that
/// is really a longer lexeme (`R[1]C2x`, `R[1]C$2` — the col part is
/// simply not consumed), and any atom followed directly by `[`
/// (`R1C1[2]` could be a specifier on a foreign-produced table name;
/// leave those bytes exactly as they always tokenized).
fn scanR1C1Atom(input: []const u8, start: usize) ?ScannedR1C1 {
    var i = start;
    var row: ?R1C1Part = null;
    if (i < input.len and (input[i] == 'R' or input[i] == 'r')) {
        // A malformed row part sinks the whole atom — nothing R1C1
        // starts to the right of a broken `R[…`.
        row = scanR1C1Part(input, &i, "Cc") orelse return null;
    }
    var col: ?R1C1Part = null;
    const col_mark = i;
    if (i < input.len and (input[i] == 'C' or input[i] == 'c')) {
        // A failed col ATTEMPT sinks the whole atom (Codex #192 F2):
        // merging just the row prefix of `R[8]C[99999999]` or
        // `R[1]C$2` would hand the rewriter a live atom that is a
        // strict prefix of a construct it couldn't read, and the
        // leftover bytes would rewrite under their A1 reading. The
        // full fallback keeps every such spelling byte-identical to
        // its pre-merge tokenization.
        col = scanR1C1Part(input, &i, "") orelse return null;
    }
    // Shape gate: which part combinations are R1C1 rather than a name
    // or an A1 ref. `RC4` is column RC row 4 (cell-ref precedence,
    // §5.2) and `R5` is a plain name, so a bare row letter never
    // anchors an atom and a digits row needs a col part.
    if (row) |r| switch (r.form) {
        .bare => return null,
        .digits => if (col == null) return null,
        .bracket => {},
    } else {
        const c = col orelse return null;
        if (c.form != .bracket) return null; // bare `C` / `C5`: name / A1
    }
    // A `[` directly after the atom means these bytes may spell
    // something else entirely (`R1C1[2]` on a foreign-produced table
    // named `R1C1`) — leave them exactly as they always tokenized.
    if (i < input.len and input[i] == '[') return null;
    return .{
        .atom = .{
            .row = row,
            .col = col,
            .col_start = (if (col != null) col_mark else i) - start,
        },
        .end = i,
    };
}

/// Scan one `R`/`C` part at `i.*` (caller has already matched the
/// letter): letter alone (bare), letter + digits, or letter +
/// `[signed digits]`. Advances `i.*` past what it consumed. Returns
/// null — with `i.*` unspecified — for a malformed part: a bad
/// bracket body, more than 7 digits, or a digits/bare spelling whose
/// next byte keeps the lexeme going (`R1C2x`, `R[1]Cx` — that is a
/// name, not a part). `boundary_exempt` names bytes allowed to follow
/// a digits/bare spelling anyway: the row part exempts `Cc` because
/// its digits are legitimately followed by the col letter.
fn scanR1C1Part(input: []const u8, i: *usize, boundary_exempt: []const u8) ?R1C1Part {
    std.debug.assert(i.* < input.len);
    var j = i.* + 1;
    if (j < input.len and input[j] == '[') {
        j += 1;
        const neg = j < input.len and input[j] == '-';
        if (neg) j += 1;
        const digit_start = j;
        while (j < input.len and isDigit(input[j])) : (j += 1) {}
        const n_digits = j - digit_start;
        if (n_digits == 0 or n_digits > 7) return null;
        if (j >= input.len or input[j] != ']') return null;
        const magnitude = std.fmt.parseInt(i64, input[digit_start..j], 10) catch unreachable;
        i.* = j + 1;
        return .{ .form = .bracket, .value = if (neg) -magnitude else magnitude };
    }
    const digit_start = j;
    while (j < input.len and isDigit(input[j])) : (j += 1) {}
    const n_digits = j - digit_start;
    if (n_digits > 7) return null;
    if (j < input.len) {
        const next = input[j];
        const exempt = std.mem.indexOfScalar(u8, boundary_exempt, next) != null;
        if (!exempt and (next == '$' or next >= 0x80 or isAsciiIdentContinue(next))) {
            return null;
        }
    }
    i.* = j;
    if (n_digits == 0) return .{ .form = .bare, .value = 0 };
    return .{
        .form = .digits,
        .value = std.fmt.parseInt(i64, input[digit_start..j], 10) catch unreachable,
    };
}

/// Decompose a `.r1c1_ref` token's text. The token was produced by
/// `scanR1C1Atom`, so this cannot fail on scanner output; the return
/// stays optional because the rewriter treats "doesn't re-parse" as
/// "don't touch" rather than trusting and asserting.
pub fn parseR1C1AtomText(text: []const u8) ?R1C1Atom {
    const scanned = scanR1C1Atom(text, 0) orelse return null;
    if (scanned.end != text.len) return null;
    return scanned.atom;
}

/// True for the bracket-free R1C1 forms — `R1C1`, `R12C`, `RC4`.
///
/// Bare `RC` is deliberately excluded: `RC` is also column 471, so
/// `RC:RC` is a legal A1 full-column reference, and in A1 mode that
/// reading is the live one. The bracketed relative forms are caught by
/// the caller's `[` lookahead instead, since `[` ends the lexeme.
fn isR1C1Shape(s: []const u8) bool {
    if (s.len < 2) return false;
    var i: usize = 0;
    if (s[i] != 'R' and s[i] != 'r') return false;
    i += 1;
    const r_start = i;
    while (i < s.len and isDigit(s[i])) : (i += 1) {}
    const r_digits = i - r_start;
    if (i >= s.len) return false;
    if (s[i] != 'C' and s[i] != 'c') return false;
    i += 1;
    const c_start = i;
    while (i < s.len and isDigit(s[i])) : (i += 1) {}
    const c_digits = i - c_start;
    if (i != s.len) return false;
    return r_digits + c_digits > 0;
}

fn columnInRange(letters: []const u8) bool {
    // M0 adapter over `zlsx_refs` (A=1 … XFD=16384, case-insensitive).
    // The shared scanner uses trapping arithmetic, which also closes a
    // latent overflow panic here: the old loop multiplied unchecked, so
    // a long enough letter run would have trapped on `v * 26` rather
    // than answering `false`.
    _ = coords.parseColNumber(letters, .{ .case = .insensitive }) catch return false;
    return true;
}

fn rowInRange(digits: []const u8) bool {
    var v: u64 = 0;
    for (digits) |c| {
        v = v * 10 + @as(u64, c - '0');
        if (v > coords.max_row) return false;
    }
    return v >= 1 and v <= coords.max_row;
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

/// Round-trip plus the single-token assertion: `input` must tokenize
/// to exactly one token of `kind` covering every byte.
fn expectSingle(input: []const u8, kind: Token.Kind) !void {
    try expectKinds(input, &.{kind});
    try expectRoundTrip(input);
}

fn expectRefusals(input: []const u8, expected: []const Refusal.Reason) !void {
    var result = try scan(testing.allocator, input);
    defer result.deinit(testing.allocator);
    if (result.refusals.len != expected.len) {
        std.debug.print("expected {d} refusals, got {d} for input '{s}':\n", .{ expected.len, result.refusals.len, input });
        for (result.refusals) |r| std.debug.print("  {s} @{d}+{d}\n", .{ @tagName(r.reason), r.offset, r.len });
        return error.TestExpectedEqual;
    }
    for (result.refusals, 0..) |r, i| {
        try testing.expectEqual(expected[i], r.reason);
    }
    // A refusal never costs the round-trip.
    const back = try format(testing.allocator, result.tokens);
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

test "full-row and full-column references stay structurally whole" {
    // §5.2 promises the M2 parser these shapes; the tokenizer's job is
    // to hand them over without splitting a construct in two.
    try expectKinds("1:1", &.{ .number, .op_range, .number });
    try expectKinds("2:3", &.{ .number, .op_range, .number });
    try expectKinds("$1:$1", &.{ .name, .op_range, .name });
    try expectKinds("$A:$A", &.{ .name, .op_range, .name });
    try expectKinds("A:XFD", &.{ .name, .op_range, .name });
    try expectKinds("SUM(1:1)", &.{ .name, .paren_open, .number, .op_range, .number, .paren_close });
    try expectKinds("COUNTBLANK(2:3)", &.{ .name, .paren_open, .number, .op_range, .number, .paren_close });
    for ([_][]const u8{ "1:1", "$1:$1", "$A:$A", "SUM(1:1)", "COUNTBLANK(2:3)" }) |c| {
        try expectRoundTrip(c);
    }
}

test "3D references" {
    try expectKinds("Sheet1:Sheet3!A1", &.{ .name, .op_range, .name, .bang, .cell_ref });
    try expectKinds("'My Sheet:Other'!A1", &.{ .sheet_name, .bang, .cell_ref });
    try expectKinds(
        "SUM(Sheet1:Sheet3!A1:B2)",
        &.{ .name, .paren_open, .name, .op_range, .name, .bang, .cell_ref, .op_range, .cell_ref, .paren_close },
    );
    try expectRoundTrip("SUM(Sheet1:Sheet3!A1:B2)");
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

test "error literals — the frozen ten" {
    // Canonical seven.
    try expectSingle("#N/A", .error_lit);
    try expectSingle("#REF!", .error_lit);
    try expectSingle("#DIV/0!", .error_lit);
    try expectSingle("#NAME?", .error_lit);
    try expectSingle("#NUM!", .error_lit);
    try expectSingle("#VALUE!", .error_lit);
    try expectSingle("#NULL!", .error_lit);
    // Excel 365 dynamic-array errors.
    try expectSingle("#SPILL!", .error_lit);
    try expectSingle("#CALC!", .error_lit);
    try expectSingle("#GETTING_DATA", .error_lit);

    for ([_][]const u8{
        "#N/A",    "#REF!",  "#DIV/0!", "#NAME?", "#NUM!",
        "#VALUE!", "#NULL!", "#SPILL!", "#CALC!", "#GETTING_DATA",
    }) |lit| {
        try testing.expect(isKnownErrorLiteral(lit));
    }
}

test "error literals — extensible rule preserves unknown spellings" {
    // Errors Excel has shipped since the whitelist was frozen. Each
    // must survive as ONE token with its bytes intact — the closed
    // whitelist shredded them into `unknown` + `name` + operator.
    for ([_][]const u8{
        "#BLOCKED!",  "#CONNECT!", "#UNKNOWN!", "#FIELD!",
        "#CALC!",     "#PYTHON!",  "#BUSY!",    "#EXTERNAL!",
        "#TIMEOUT!",  "#PRIVACY!", "#PENDING?", "#PARSE.ERROR!",
        "#A1/B2_C3!",
    }) |lit| {
        try expectSingle(lit, .error_lit);
    }
    // …but they are not the known ten, so no caller may interpret them.
    try testing.expect(!isKnownErrorLiteral("#BLOCKED!"));
    try testing.expect(!isKnownErrorLiteral("#PYTHON!"));
    try testing.expect(isKnownErrorLiteral("#REF!"));

    // In-context round-trip.
    try expectKinds("IFERROR(A1,#BLOCKED!)", &.{
        .name, .paren_open, .cell_ref, .arg_sep, .error_lit, .paren_close,
    });
    try expectRoundTrip("IFERROR(A1,#BLOCKED!)");
}

test "error literals — extensible rule boundaries" {
    // No body → not an error literal, so the `#` is the spill operator.
    try expectKinds("#!", &.{ .op_spill, .bang });
    try expectKinds("#?", &.{ .op_spill, .unknown });
    // No terminator → spill operator plus a name.
    try expectKinds("#FOO", &.{ .op_spill, .name });
    try expectRoundTrip("#FOO");
    // A byte outside the body set stops the match before any
    // terminator, so `#FOO BAR!` never becomes one giant token.
    try expectKinds("#FOO BAR!", &.{ .op_spill, .name, .whitespace, .name, .bang });
    try expectRoundTrip("#FOO BAR!");
    // Length bound: a body longer than the cap cannot swallow the rest
    // of the expression looking for a terminator.
    const long = "#" ++ ("A" ** (max_error_literal_bytes + 4)) ++ "!";
    try expectKinds(long, &.{ .op_spill, .name, .bang });
    try expectRoundTrip(long);
    // …and one byte inside the cap still matches.
    const fits = "#" ++ ("A" ** (max_error_literal_bytes - 2)) ++ "!";
    try expectSingle(fits, .error_lit);
}

test "dynamic-array operators are typed, not unknown" {
    // Pre-M1a both were `.unknown`; the bytes are identical, the kinds
    // are the correction. The scope note this file used to carry
    // ("dynamic-array operators (#, @) … out of scope") went with them.
    try expectKinds("A1#", &.{ .cell_ref, .op_spill });
    try expectRoundTrip("A1#");
    try expectKinds("SUM(A1#)", &.{ .name, .paren_open, .cell_ref, .op_spill, .paren_close });
    try expectKinds("@A1", &.{ .op_at, .cell_ref });
    try expectRoundTrip("@A1");
    try expectKinds("@SUM(A1:A2)", &.{
        .op_at, .name, .paren_open, .cell_ref, .op_range, .cell_ref, .paren_close,
    });
    try expectKinds("Sheet1!A1#", &.{ .name, .bang, .cell_ref, .op_spill });
    try expectKinds("_xlfn.SINGLE(A1:A5)", &.{
        .name, .paren_open, .cell_ref, .op_range, .cell_ref, .paren_close,
    });
    // Neither operator is an error literal.
    try expectRefusals("A1#+@B2", &.{});
}

test "structured refs: Table1[A1] is one opaque token" {
    // The correction the ladder names. Pre-M1a this tokenized as
    // `name` `unknown` `cell_ref` `unknown` — and that inner
    // `.cell_ref` is live to the rewriter, so an insert-rows edit on
    // the sheet would have rewritten a table COLUMN NAME into `A2`,
    // corrupting the formula. One opaque token closes the hole.
    try expectKinds("Table1[A1]", &.{ .name, .structured_ref });
    try expectRoundTrip("Table1[A1]");

    var result = try scan(testing.allocator, "Table1[A1]");
    defer result.deinit(testing.allocator);
    try testing.expectEqualStrings("Table1", result.tokens[0].text);
    try testing.expectEqualStrings("[A1]", result.tokens[1].text);
    try testing.expectEqual(@as(usize, 0), result.refusals.len);
}

test "structured refs: item specifiers, ranges, escapes" {
    try expectKinds("Table1[Amount]", &.{ .name, .structured_ref });
    try expectKinds("Table1[#All]", &.{ .name, .structured_ref });
    try expectKinds("Table1[#Headers]", &.{ .name, .structured_ref });
    try expectKinds("Table1[#Totals]", &.{ .name, .structured_ref });
    try expectKinds("Table1[#Data]", &.{ .name, .structured_ref });
    try expectKinds("Table1[@Amount]", &.{ .name, .structured_ref });
    // `[@]` this-row with no table name (inside the table itself).
    try expectKinds("[@Amount]", &.{.structured_ref});
    // Nested brackets: a column range, and combined forms.
    try expectSingle("[[Col A]:[Col B]]", .structured_ref);
    try expectKinds("Table1[[#Data],[Amount]]", &.{ .name, .structured_ref });
    try expectKinds("Table1[[#This Row],[Amount]]", &.{ .name, .structured_ref });
    // The `'` escapes: `'[`, `']`, `'#`, `'@` name a column whose text
    // contains the special character. The escaped bracket must not
    // change the depth count.
    try expectSingle("[Cost '[USD']]", .structured_ref);
    try expectSingle("[Item '#1]", .structured_ref);
    try expectSingle("[Rate '@Peak]", .structured_ref);
    try expectSingle("[Col ']]", .structured_ref);
    // In context.
    try expectKinds("SUM(Table1[Amount])", &.{
        .name, .paren_open, .name, .structured_ref, .paren_close,
    });
    try expectRoundTrip("=SUM(Sales[[#Data],[Q'[1']]])+Table1[@[Unit Cost]]");
}

test "structured refs: unterminated bracket refuses without losing bytes" {
    try expectKinds("Table1[Amount", &.{ .name, .unknown });
    try expectRefusals("Table1[Amount", &.{.unterminated_structured_ref});
    try expectRoundTrip("Table1[Amount");
    // An escape at the very end must not read past the buffer.
    try expectRoundTrip("Table1[Amount'");
    try expectRefusals("Table1[Amount'", &.{.unterminated_structured_ref});
    // A stray closing bracket is unknown, not a depth underflow.
    try expectKinds("A1]", &.{ .cell_ref, .unknown });
    try expectRoundTrip("]]]");
}

test "external refs are typed and refused" {
    // Quoted form.
    try expectKinds("'[Book.xlsx]Sheet1'!A1", &.{ .external_ref, .bang, .cell_ref });
    try expectRefusals("'[Book.xlsx]Sheet1'!A1", &.{.external_reference});
    try expectRoundTrip("'[Book.xlsx]Sheet1'!A1+1");
    // Workbook-index form — `[1]Sheet1!A1` and `[1]!DefinedName`.
    try expectKinds("[1]Sheet1!A1", &.{ .external_ref, .name, .bang, .cell_ref });
    try expectRefusals("[1]Sheet1!A1", &.{.external_reference});
    try expectKinds("[1]!Total", &.{ .external_ref, .bang, .name });
    try expectKinds("[12]'My Sheet'!A1", &.{ .external_ref, .sheet_name, .bang, .cell_ref });
    try expectRoundTrip("[1]Sheet1!A1*2");
    // A bare `[1]` with nothing after it is a table column named `1`,
    // not an external workbook.
    try expectKinds("Table1[1]", &.{ .name, .structured_ref });
    try expectRefusals("Table1[1]", &.{});
    // Plain quoted sheet without `[` stays `.sheet_name`.
    try expectKinds("'My Sheet'!A1", &.{ .sheet_name, .bang, .cell_ref });
    try expectRefusals("'My Sheet'!A1", &.{});
}

test "R1C1 shapes are recognised only to reject them" {
    try expectRefusals("R1C1", &.{.r1c1_reference});
    try expectRefusals("R12C4", &.{.r1c1_reference});
    try expectRefusals("r1c1", &.{.r1c1_reference});
    try expectRefusals("R1C", &.{.r1c1_reference});
    // Bracketed relative forms merge with the letter into one atom —
    // one construct, one refusal.
    try expectRefusals("R[-1]C", &.{.r1c1_reference});
    try expectRefusals("C[2]", &.{.r1c1_reference});
    try expectRefusals("R[-1]C[2]", &.{.r1c1_reference});
    // Refused, but still round-tripped.
    try expectRoundTrip("R1C1");
    try expectRoundTrip("R[-1]C[2]");

    // Not R1C1: `RC1` is column RC row 1, a live A1 reference, and the
    // cell-ref rule outranks the R1C1 rule (§5.2 precedence).
    try expectKinds("RC1", &.{.cell_ref});
    try expectRefusals("RC1", &.{});
    // Bare `RC` is ambiguous with column RC (`RC:RC`) — left alone.
    try expectRefusals("RC:RC", &.{});
    // Ordinary names that merely start with R or C.
    try expectRefusals("Rate", &.{});
    try expectRefusals("Cost", &.{});
    try expectRefusals("R:R", &.{});
    try expectRefusals("SUM(C:C)", &.{});
    // A sheet named R1C1 is a name, not a reference.
    try expectKinds("R1C1!A1", &.{ .name, .bang, .cell_ref });
    try expectRefusals("R1C1!A1", &.{});
}

test "R1C1: clean whole atoms merge into one .r1c1_ref token" {
    try expectKinds("R1C1", &.{.r1c1_ref});
    try expectKinds("R12C4", &.{.r1c1_ref});
    try expectKinds("r1c1", &.{.r1c1_ref});
    try expectKinds("R1C", &.{.r1c1_ref});
    try expectKinds("R1C[2]", &.{.r1c1_ref});
    try expectKinds("R[-1]C", &.{.r1c1_ref});
    try expectKinds("R[-1]C[2]", &.{.r1c1_ref});
    // Pre-merge, `[1]` in here scanned as an external-workbook index
    // and `C2` as a live A1 cell ref the rewriter would shift.
    try expectKinds("R[1]C2", &.{.r1c1_ref});
    try expectRefusals("R[1]C2", &.{.r1c1_reference});
    // Whole-row / whole-col relative forms.
    try expectKinds("R[2]", &.{.r1c1_ref});
    try expectKinds("C[-3]", &.{.r1c1_ref});
    try expectKinds("R[0]", &.{.r1c1_ref});
    // Ranges pair two atoms around the ordinary range operator.
    try expectKinds("R1C1:R2C2", &.{ .r1c1_ref, .op_range, .r1c1_ref });
    // Sheet-qualified: the qualifier lexes exactly as before.
    try expectKinds("Sheet2!R5C3", &.{ .name, .bang, .r1c1_ref });
    try expectRoundTrip("SUM(R[-2]C7,R1C1:R2C2)");
}

test "R1C1: unclean spellings keep the pre-merge fragment tokenization" {
    // Digit-tailed atom followed by `[`: could be a specifier on a
    // foreign-produced table named `R1C1`.
    try expectKinds("R1C1[2]", &.{ .name, .structured_ref });
    try expectRefusals("R1C1[2]", &.{.r1c1_reference});
    // Malformed bracket bodies.
    try expectKinds("R[foo]", &.{ .name, .structured_ref });
    try expectRefusals("R[foo]", &.{.r1c1_reference});
    try expectKinds("R[]", &.{ .name, .structured_ref });
    // A failed col ATTEMPT sinks the whole atom (Codex #192 F2):
    // merging just `R[1]` would leave bytes that rewrite under
    // their A1 reading. These keep the full pre-merge fragment
    // stream — where `[1]` scans as an external-workbook index,
    // whose chain-skip is what protects the tail fragments.
    try expectKinds("R[1]C2x", &.{ .name, .external_ref, .name });
    try expectRefusals("R[1]C2x", &.{ .r1c1_reference, .external_reference });
    try expectKinds("R[1]C$2", &.{ .name, .external_ref, .cell_ref });
    try expectKinds("R[8]C[99999999]", &.{ .name, .external_ref, .name, .structured_ref });
    try expectRefusals("R[8]C[99999999]", &.{ .r1c1_reference, .external_reference, .r1c1_reference });
}

test "R1C1: parseR1C1AtomText decomposes exactly what the scanner merged" {
    const abs_abs = parseR1C1AtomText("R5C3").?;
    try testing.expectEqual(R1C1Part.Form.digits, abs_abs.row.?.form);
    try testing.expectEqual(@as(i64, 5), abs_abs.row.?.value);
    try testing.expectEqual(R1C1Part.Form.digits, abs_abs.col.?.form);
    try testing.expectEqual(@as(i64, 3), abs_abs.col.?.value);

    const rel_bare = parseR1C1AtomText("R[-1]C").?;
    try testing.expectEqual(R1C1Part.Form.bracket, rel_bare.row.?.form);
    try testing.expectEqual(@as(i64, -1), rel_bare.row.?.value);
    try testing.expectEqual(R1C1Part.Form.bare, rel_bare.col.?.form);

    try testing.expectEqual(@as(usize, 2), abs_abs.col_start); // "R5"|"C3"

    const col_only = parseR1C1AtomText("c[12]").?;
    try testing.expect(col_only.row == null);
    try testing.expectEqual(@as(i64, 12), col_only.col.?.value);
    try testing.expectEqual(@as(usize, 0), col_only.col_start);

    const row_only = parseR1C1AtomText("R[0]").?;
    try testing.expect(row_only.col == null);
    try testing.expectEqual(@as(i64, 0), row_only.row.?.value);
    try testing.expectEqual(@as(usize, 4), row_only.col_start); // == text.len

    // The gate set: spellings that are NOT atoms.
    try testing.expect(parseR1C1AtomText("RC4") == null); // A1: column RC
    try testing.expect(parseR1C1AtomText("RC") == null);
    try testing.expect(parseR1C1AtomText("R5") == null); // plain name
    try testing.expect(parseR1C1AtomText("C5") == null); // A1 cell
    try testing.expect(parseR1C1AtomText("R") == null);
    try testing.expect(parseR1C1AtomText("C") == null);
    try testing.expect(parseR1C1AtomText("R[8]C[99999999]") == null); // >7 digits
    try testing.expect(parseR1C1AtomText("R1C1[2]") == null);
    try testing.expect(parseR1C1AtomText("R[1]x") == null); // trailing bytes
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
    // `~` isn't part of any production. Round-trip still works: it
    // falls into `.unknown` and the formatter emits it back verbatim.
    try expectKinds("~A1", &.{ .unknown, .cell_ref });
    try expectRoundTrip("~A1");
}

test "named ranges and function calls share the .name kind" {
    // The tokenizer doesn't disambiguate; lookahead for `(` is the
    // parser's job. Both produce `.name` tokens.
    try expectKinds("MyRange", &.{.name});
    try expectKinds("SUM", &.{.name});
    try expectKinds("My_Range.v2", &.{.name});
}

test "A1-shaped function names + sheet refs classify as .name not .cell_ref" {
    // `LOG10` matches the A1 shape (column LOG = 8509, row 10) but
    // is a built-in function. The trailing `(` flips it to `.name`.
    try expectKinds("LOG10(A1)", &.{
        .name, .paren_open, .cell_ref, .paren_close,
    });
    // `Q1` is a valid A1 ref AND a common sheet name. Trailing `!`
    // means it's the sheet qualifier — `.name`, not `.cell_ref`.
    try expectKinds("Q1!A1", &.{ .name, .bang, .cell_ref });
    // Without the trailing delimiter, `Q1` stays a cell ref.
    try expectKinds("Q1", &.{.cell_ref});
    try expectKinds("Q1+B2", &.{ .cell_ref, .op_plus, .cell_ref });
}

test "backslash-started defined names" {
    // Excel allows `\` as a name starter. openpyxl tokenizes
    // `\Foo` as a RANGE (defined name); we should match.
    try expectKinds("\\Foo", &.{.name});
    try expectRoundTrip("\\Foo+1");
    try expectRefusals("\\Foo", &.{});
}

test "backslash is start-only" {
    // `a\b` is not one name — the grammar puts `\` in the start set
    // only. The lexeme ends at the backslash, which then starts a
    // fresh identifier (`\b`), so the bytes survive as two adjacent
    // names; the refusal is what the parser acts on at M2.
    try expectKinds("a\\b", &.{ .name, .name });
    try expectRefusals("a\\b", &.{.backslash_after_start});
    try expectRoundTrip("a\\b");
    var split = try scan(testing.allocator, "a\\b");
    defer split.deinit(testing.allocator);
    try testing.expectEqualStrings("a", split.tokens[0].text);
    try testing.expectEqualStrings("\\b", split.tokens[1].text);
    // Doubled backslash mid-name is still one refusal per occurrence.
    try expectRefusals("ab\\\\cd", &.{ .backslash_after_start, .backslash_after_start });
    // A backslash starting the SECOND name is fine.
    try expectRefusals("\\a+\\b", &.{});
}

test "TRUE()/FALSE() are function calls, not literals" {
    // Zero-argument TRUE() and FALSE() are valid Excel functions.
    // The trailing `(` flips them from `.bool_lit` to `.name`.
    try expectKinds("TRUE()", &.{ .name, .paren_open, .paren_close });
    try expectKinds("FALSE()", &.{ .name, .paren_open, .paren_close });
    // No trailing `(` — back to literal classification.
    try expectKinds("TRUE", &.{.bool_lit});
    try expectKinds("FALSE", &.{.bool_lit});
}

test "mixed bag round-trip" {
    try expectRoundTrip("=IF(SUM($A$1:$A$10)>0,\"yes\",\"no\")");
    try expectRoundTrip("Sheet1!A1+Sheet2!$B$2-'Some Sheet'!C3");
    try expectRoundTrip("MIN(1,2,3) - MAX({1;2;3})");
}

// ─── Unicode identifiers ─────────────────────────────────────────

test "unicode identifiers: non-ASCII names are single tokens" {
    // Every one of these is an ordinary Excel defined name. Pre-M1a
    // the ASCII-only predicate shattered each into one `.unknown` per
    // BYTE — `Ω` became two tokens, and a rewriter walking the stream
    // saw noise where a name belongs.
    try expectSingle("Ω", .name);
    try expectSingle("Σ", .name);
    try expectSingle("ДАННЫЕ", .name);
    try expectSingle("Größe", .name);
    try expectSingle("café", .name);
    try expectSingle("日本語", .name);
    try expectSingle("데이터", .name);
    try expectSingle("שם", .name);
    try expectSingle("اسم", .name);
    // Continue-only characters after a valid start.
    try expectSingle("Ω1", .name);
    try expectSingle("données_2024", .name);
    try expectSingle("Ω.Σ", .name);
    // In context.
    try expectKinds("SUM(ДАННЫЕ)", &.{ .name, .paren_open, .name, .paren_close });
    try expectKinds("Ω!A1", &.{ .name, .bang, .cell_ref });
    try expectKinds("Größe+1", &.{ .name, .op_plus, .number });
    try expectRoundTrip("=SUM(données!Ω1:Ω9)*Größe");
}

test "unicode identifiers: astral starts round-trip" {
    // U+1D400 MATHEMATICAL BOLD CAPITAL A — a 4-byte XID_Start.
    try expectSingle("\u{1D400}", .name);
    try expectSingle("\u{1D400}\u{1D401}", .name);
    // U+10000 LINEAR B SYLLABLE B008 A.
    try expectSingle("\u{10000}bc", .name);
    // U+20000 CJK extension B.
    try expectSingle("\u{20000}", .name);
    try expectKinds("SUM(\u{1D400}1:\u{1D400}9)", &.{
        .name, .paren_open, .name, .op_range, .name, .paren_close,
    });
    try expectRoundTrip("\u{1D400}+\u{20000}");
    try expectRefusals("\u{1D400}", &.{});
}

test "unicode identifiers: combining marks continue but do not start" {
    // `e` + COMBINING ACUTE: one name, the mark continues it.
    try expectSingle("e\u{0301}", .name);
    try expectSingle("cafe\u{0301}_total", .name);
    // A leading combining mark cannot start an identifier — it is not
    // a name, and the codepoint is preserved whole (not per byte).
    try expectKinds("\u{0301}abc", &.{ .unknown, .name });
    try expectRoundTrip("\u{0301}abc");
    var result = try scan(testing.allocator, "\u{0301}abc");
    defer result.deinit(testing.allocator);
    try testing.expectEqualStrings("\u{0301}", result.tokens[0].text);
}

test "unicode identifiers: cell-looking names" {
    // Fullwidth `Ａ1` is NOT the A1 reference — the column letters must
    // be ASCII. It is a name, and must stay one.
    try expectSingle("\u{FF21}1", .name);
    // A name that would be a cell ref if its non-ASCII head were
    // dropped: the head keeps it a name.
    try expectSingle("ΩA1", .name);
    // Sheet named like a ref, in Cyrillic.
    try expectKinds("Ф1!A1", &.{ .name, .bang, .cell_ref });
}

test "unicode identifiers: non-identifier codepoints stay unknown" {
    // Symbols, currency, punctuation and format controls are not
    // identifier characters at any position.
    for ([_][]const u8{ "€", "©", "😀", "\u{00A0}", "\u{200B}", "−" }) |s| {
        try expectKinds(s, &.{.unknown});
        try expectRoundTrip(s);
    }
    // …and one whole codepoint per token, never one per byte.
    var result = try scan(testing.allocator, "😀");
    defer result.deinit(testing.allocator);
    try testing.expectEqual(@as(usize, 1), result.tokens.len);
    try testing.expectEqualStrings("😀", result.tokens[0].text);
    // A symbol inside a name terminates it rather than joining it.
    try expectKinds("a€b", &.{ .name, .unknown, .name });
}

test "unicode identifiers: ASCII predicates agree with the XID tables" {
    // The hot path uses byte tests instead of table probes; if the two
    // ever disagree, an ASCII name would tokenize differently from the
    // identical codepoint reached through the Unicode branch.
    var c: u8 = 0;
    while (c < 0x80) : (c += 1) {
        const cp: u21 = c;
        try testing.expectEqual(
            xid.isXidStart(cp) or c == '_' or c == '\\',
            isAsciiIdentStart(c),
        );
        try testing.expectEqual(
            xid.isXidContinue(cp) or c == '.',
            isAsciiIdentContinue(c),
        );
    }
}

test "invalid UTF-8 is refused, never dropped" {
    const cases = [_][]const u8{
        "\xFF",
        "\xC0\x80", // overlong
        "\xE2\x82", // truncated 3-byte
        "\xED\xA0\x80", // surrogate half
        "A1+\xFF",
        "SUM(\xC3)",
        "\xF0\x9F", // truncated 4-byte
    };
    for (cases) |c| {
        var result = try scan(testing.allocator, c);
        defer result.deinit(testing.allocator);
        try testing.expect(result.find(.invalid_utf8) != null);
        const back = try format(testing.allocator, result.tokens);
        defer testing.allocator.free(back);
        try testing.expectEqualSlices(u8, c, back);
    }
    // A run of bad bytes coalesces into ONE refusal rather than burying
    // every other diagnostic under per-byte noise.
    try expectRefusals("\xFF\xFF\xFF\xFF", &.{.invalid_utf8});
    // Invalid bytes *inside* a name end it and refuse.
    try expectRefusals("ab\xFFcd", &.{.invalid_utf8});
}

test "identifier length limit" {
    // 255 code points is Excel's defined-name cap; the limit counts
    // code points, not bytes, so a Cyrillic name of 255 characters
    // (510 bytes) is legal and a 256-character one is not.
    const ascii_ok = "a" ** max_identifier_codepoints;
    try expectRefusals(ascii_ok, &.{});
    try expectSingle(ascii_ok, .name);

    const ascii_over = "a" ** (max_identifier_codepoints + 1);
    try expectRefusals(ascii_over, &.{.identifier_too_long});
    try expectSingle(ascii_over, .name); // still one token, still round-trips

    const cyrillic_ok = "Д" ** max_identifier_codepoints;
    try testing.expectEqual(@as(usize, 2 * max_identifier_codepoints), cyrillic_ok.len);
    try expectRefusals(cyrillic_ok, &.{});

    const cyrillic_over = "Д" ** (max_identifier_codepoints + 1);
    try expectRefusals(cyrillic_over, &.{.identifier_too_long});

    // Astral: 4 bytes per code point, so the byte length is 4× the cap
    // and a byte-counting implementation would have refused at 64.
    const astral_ok = "\u{1D400}" ** max_identifier_codepoints;
    try testing.expectEqual(@as(usize, 4 * max_identifier_codepoints), astral_ok.len);
    try expectRefusals(astral_ok, &.{});
    const astral_over = "\u{1D400}" ** (max_identifier_codepoints + 1);
    try expectRefusals(astral_over, &.{.identifier_too_long});
}

test "unterminated string and sheet name refuse" {
    try expectRefusals("\"abc", &.{.unterminated_string});
    try expectRoundTrip("\"abc");
    try expectRefusals("SUM(\"x", &.{.unterminated_string});
    try expectRefusals("'My Sheet", &.{.unterminated_sheet_name});
    try expectRoundTrip("'My Sheet");
    // A lone quote is unterminated too — and must not underflow.
    try expectRefusals("\"", &.{.unterminated_string});
    try expectRefusals("'", &.{.unterminated_sheet_name});
    // Terminated forms refuse nothing.
    try expectRefusals("\"abc\"", &.{});
    try expectRefusals("'My Sheet'!A1", &.{});
}

test "corrections: constructs the pre-M1a tokenizer got wrong" {
    // Each row is a construct whose BYTES round-tripped before but
    // whose KINDS were wrong. Round-trip was never the gap; the gap was
    // that a wrong kind is a licence for the rewriter to mutate.
    const Case = struct { input: []const u8, kinds: []const Token.Kind };
    const cases = [_]Case{
        // `[A1]` was `unknown cell_ref unknown` — a live ref inside a
        // table column name.
        .{ .input = "Table1[A1]", .kinds = &.{ .name, .structured_ref } },
        // `#` was `.unknown` in both operator positions.
        .{ .input = "A1#", .kinds = &.{ .cell_ref, .op_spill } },
        .{ .input = "@A1", .kinds = &.{ .op_at, .cell_ref } },
        // The external prefix was `.unknown`, indistinguishable from
        // arbitrary junk.
        .{ .input = "'[Book.xlsx]Sheet1'!A1", .kinds = &.{ .external_ref, .bang, .cell_ref } },
        // `?` was accepted in names; Excel does not allow it.
        .{ .input = "FOO?", .kinds = &.{ .name, .unknown } },
        .{ .input = "A?1", .kinds = &.{ .name, .unknown, .number } },
        // Non-ASCII names were one `.unknown` per byte.
        .{ .input = "Ω", .kinds = &.{.name} },
    };
    for (cases) |c| {
        try expectKinds(c.input, c.kinds);
        try expectRoundTrip(c.input);
    }
}

test "compat: every previously-recognized construct keeps its kinds" {
    // The M1a gate. These are the constructs the pre-M1a tokenizer
    // classified, with the kinds it classified them as. Any drift here
    // is a compatibility break in the rewriter, which reads kinds to
    // decide what it may touch.
    const Case = struct { input: []const u8, kinds: []const Token.Kind };
    const cases = [_]Case{
        .{ .input = "A1", .kinds = &.{.cell_ref} },
        .{ .input = "$A$1", .kinds = &.{.cell_ref} },
        .{ .input = "$A1", .kinds = &.{.cell_ref} },
        .{ .input = "A$1", .kinds = &.{.cell_ref} },
        .{ .input = "XFD1048576", .kinds = &.{.cell_ref} },
        .{ .input = "a1", .kinds = &.{.cell_ref} },
        .{ .input = "XFE1", .kinds = &.{.name} },
        .{ .input = "A1048577", .kinds = &.{.name} },
        .{ .input = "A1:B5", .kinds = &.{ .cell_ref, .op_range, .cell_ref } },
        .{ .input = "A:A", .kinds = &.{ .name, .op_range, .name } },
        .{ .input = "Sheet1!A1", .kinds = &.{ .name, .bang, .cell_ref } },
        .{ .input = "'My Sheet'!A1", .kinds = &.{ .sheet_name, .bang, .cell_ref } },
        .{ .input = "'It''s'!A1", .kinds = &.{ .sheet_name, .bang, .cell_ref } },
        .{ .input = "SUM(A1:A10)", .kinds = &.{ .name, .paren_open, .cell_ref, .op_range, .cell_ref, .paren_close } },
        .{ .input = "\"a\"\"b\"", .kinds = &.{.string} },
        .{ .input = "1.5e10", .kinds = &.{.number} },
        .{ .input = ".5", .kinds = &.{.number} },
        .{ .input = "1E+5", .kinds = &.{.number} },
        .{ .input = "EXP", .kinds = &.{.name} },
        .{ .input = "TRUE", .kinds = &.{.bool_lit} },
        .{ .input = "false", .kinds = &.{.bool_lit} },
        .{ .input = "TRUE()", .kinds = &.{ .name, .paren_open, .paren_close } },
        .{ .input = "#N/A", .kinds = &.{.error_lit} },
        .{ .input = "#DIV/0!", .kinds = &.{.error_lit} },
        .{ .input = "#GETTING_DATA", .kinds = &.{.error_lit} },
        .{ .input = "50%", .kinds = &.{ .number, .op_percent } },
        .{ .input = "1<>2", .kinds = &.{ .number, .op_ne, .number } },
        .{ .input = "1<=2", .kinds = &.{ .number, .op_le, .number } },
        .{ .input = "1>=2", .kinds = &.{ .number, .op_ge, .number } },
        .{ .input = "{1,2;3,4}", .kinds = &.{ .array_open, .number, .arg_sep, .number, .arg_sep, .number, .arg_sep, .number, .array_close } },
        .{ .input = "A1 + B2", .kinds = &.{ .cell_ref, .whitespace, .op_plus, .whitespace, .cell_ref } },
        .{ .input = "MyName.Sub", .kinds = &.{.name} },
        .{ .input = "My_Range.v2", .kinds = &.{.name} },
        .{ .input = "\\Foo", .kinds = &.{.name} },
        .{ .input = "LOG10(A1)", .kinds = &.{ .name, .paren_open, .cell_ref, .paren_close } },
        .{ .input = "Q1!A1", .kinds = &.{ .name, .bang, .cell_ref } },
        .{ .input = "Q1", .kinds = &.{.cell_ref} },
        .{ .input = "(A1:A10) (B5:F5)", .kinds = &.{
            .paren_open,  .cell_ref,   .op_range, .cell_ref, .paren_close,
            .whitespace,  .paren_open, .cell_ref, .op_range, .cell_ref,
            .paren_close,
        } },
    };
    for (cases) |c| {
        try expectKinds(c.input, c.kinds);
        try expectRoundTrip(c.input);
        try expectRefusals(c.input, &.{});
    }
}

test "scan and tokenize agree on the token stream" {
    const inputs = [_][]const u8{
        "SUM(Table1[@Amount])", "'[B.xlsx]S'!A1",
        "Ω+#BLOCKED!",
        "a\\b",                 "\xFF",
    };
    for (inputs) |input| {
        var result = try scan(testing.allocator, input);
        defer result.deinit(testing.allocator);
        const tokens = try tokenize(testing.allocator, input);
        defer testing.allocator.free(tokens);
        try testing.expectEqual(result.tokens.len, tokens.len);
        for (result.tokens, tokens) |a, b| {
            try testing.expectEqual(a.kind, b.kind);
            try testing.expectEqualStrings(a.text, b.text);
        }
    }
}

test "allocation failure leaks nothing" {
    const inputs = [_][]const u8{
        "=SUM(Sheet1!A1:B2)+Table1[@Col]&\"x\"",
        "'[Book.xlsx]Sheet1'!A1#+@Ω\xFF",
    };
    for (inputs) |input| {
        try testing.checkAllAllocationFailures(testing.allocator, struct {
            fn run(allocator: std.mem.Allocator, text: []const u8) !void {
                var result = try scan(allocator, text);
                defer result.deinit(allocator);
                const back = try format(allocator, result.tokens);
                defer allocator.free(back);
            }
        }.run, .{input});
    }
}

// ─── fuzz ────────────────────────────────────────────────────────
//
// Contract: the tokenizer must not panic, over-read, or lose a byte on
// ANY input — including invalid UTF-8, unterminated everything, and
// pathological bracket nesting. Refusals are fine; a lost byte is not.
// Runs as a seed-corpus smoke test under `zig build test` and becomes
// coverage-guided under `zig build fuzz`.

fn fuzzTokenizerTarget(_: void, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();
    var smith_buf: [4096]u8 = undefined;
    const input = smith_buf[0..smith.slice(&smith_buf)];

    var result = scan(std.testing.allocator, input) catch return;
    defer result.deinit(std.testing.allocator);

    // 1. Round-trip: concatenating the tokens reproduces the input.
    const back = format(std.testing.allocator, result.tokens) catch return;
    defer std.testing.allocator.free(back);
    try std.testing.expectEqualSlices(u8, input, back);

    // 2. Tokens tile the input exactly — contiguous, in order, no
    //    overlap. Equality of the concatenation alone would not catch a
    //    token whose text pointed outside its span.
    var offset: usize = 0;
    for (result.tokens) |t| {
        try std.testing.expect(t.text.len > 0);
        try std.testing.expect(t.text.ptr == input.ptr + offset);
        offset += t.text.len;
    }
    try std.testing.expectEqual(input.len, offset);

    // 3. Every refusal points at a real span inside the input.
    for (result.refusals) |r| {
        try std.testing.expect(r.offset + r.len <= input.len);
    }

    // 4. Kind invariants that later stages rely on.
    for (result.tokens) |t| {
        switch (t.kind) {
            .error_lit => try std.testing.expect(t.text[0] == '#'),
            .op_spill => try std.testing.expectEqualStrings("#", t.text),
            .op_at => try std.testing.expectEqualStrings("@", t.text),
            .structured_ref => {
                try std.testing.expect(t.text[0] == '[');
                try std.testing.expect(t.text[t.text.len - 1] == ']');
            },
            .cell_ref => try std.testing.expect(isCellRef(t.text)),
            .string => try std.testing.expect(t.text[0] == '"'),
            .sheet_name => {
                try std.testing.expect(t.text[0] == '\'');
                try std.testing.expect(std.mem.indexOfScalar(u8, t.text, '[') == null);
            },
            else => {},
        }
    }
}

test "fuzz: tokenizer round-trips and tiles any input" {
    try std.testing.fuzz({}, fuzzTokenizerTarget, .{
        .corpus = &[_][]const u8{
            "=SUM(A1:B2)",
            "Table1[[#Data],[Amount]]",
            "'[Book.xlsx]Sheet1'!A1",
            "#BLOCKED!",
            "#GETTING_DATA",
            "#",
            "@",
            "[",
            "]",
            "[[[[[[[[",
            "]]]]]]]]",
            "['''''",
            "Table1['",
            "\"unterminated",
            "'unterminated",
            "R1C1",
            "R[-1]C[2]",
            "\\a\\b\\c",
            "Ω+Σ",
            "\u{1D400}",
            "e\u{0301}",
            "\xFF\xFE",
            "\xC0\x80",
            "\xED\xA0\x80",
            "\xF0\x9F\x98",
            "$1:$1",
            "1:1",
            "A1#+@B2",
            "#!?!?!?",
            "..........",
            "$$$$$$",
        },
    });
}

fn fuzzIdentifierTarget(_: void, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();
    // A narrow alphabet concentrates the search on the identifier
    // grammar's edges — start-vs-continue, `$` placement, the
    // backslash rule — which random bytes reach only rarely. Picking
    // BYTES out of the multi-byte characters is deliberate: it
    // manufactures lone continuation bytes and truncated sequences,
    // which is exactly the UTF-8 boundary the scanner must survive.
    const alphabet = "aZ0_.\\$!(:Ω\u{0301}\u{1D400}\xFF";
    var smith_buf: [512]u8 = undefined;
    const raw = smith_buf[0..smith.slice(&smith_buf)];

    var buf: [4096]u8 = undefined;
    var n: usize = 0;
    for (raw) |b| {
        const pick = alphabet[b % alphabet.len];
        if (n == buf.len) break;
        buf[n] = pick;
        n += 1;
    }
    const input = buf[0..n];

    var result = scan(std.testing.allocator, input) catch return;
    defer result.deinit(std.testing.allocator);
    const back = format(std.testing.allocator, result.tokens) catch return;
    defer std.testing.allocator.free(back);
    try std.testing.expectEqualSlices(u8, input, back);
}

test "fuzz: identifier grammar edges" {
    try std.testing.fuzz({}, fuzzIdentifierTarget, .{
        .corpus = &[_][]const u8{ "", "a", "\\", "$", "a\\b", "_x1", "Z$9" },
    });
}
