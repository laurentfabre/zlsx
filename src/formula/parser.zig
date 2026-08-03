//! Formula parser — AST, canonical printer, typed refusals.
//!
//! Consumes the M1a tokenizer's stream (`tokenizer.zig`) and produces a
//! flat, index-addressed AST plus a canonical printer with the property
//! this milestone is named for:
//!
//!   parse(print(parse(x))) ≡ parse(x)      (structural equality)
//!
//! The printer is *canonical*, not byte-preserving: insignificant
//! whitespace is dropped and booleans normalise to `TRUE`/`FALSE`.
//! Byte fidelity is the tokenizer's contract (`format(tokenize(x)) == x`)
//! and stays there — the rewriter is the byte-splice path, the parser is
//! the semantic one. Everything a canonical form cannot re-derive is
//! carried in the AST verbatim: number spellings, string bodies, error
//! spellings, sheet-name quoting, structured-ref column names, and the
//! layered `_xlfn.` / `_xlfn._xlws.` prefixes.
//!
//! Scope (M2 of the tier-D1 ladder; `goal_formula.md` §5.2, §5.9, §9, §10)
//! ---------------------------------------------------------------------
//! Precedence chain, loosest to tightest:
//!
//!   comparisons < `&` < `+` `-` < `*` `/` < `^` < `%` < unary `±`
//!     < `,` (union) < ` ` (intersection) < `:` (range)
//!
//! `^` is left-associative (`2^3^2 = 64`) and unary `±` binds tighter
//! than it (`-1^2 = 1`) — both pinned against the committed oracle
//! manifests in `tests/oracle/fixtures/`, not against an assumption.
//!
//! Also here: structured-reference item specifiers, array constants with
//! a rectangularity refusal, leading-`=` stripping, the `_xlfn.` prefix
//! layers, full-row / full-column references, the §9 parse limits, and
//! the §10 plane-2 mapping for every refusal the tokenizer reports
//! out-of-band.
//!
//! Refusals
//! --------
//! The parser refuses; it does not fabricate. `parse` returns a
//! `Parsed` union — `.ok` with an AST, or `.refused` with a typed
//! `Refusal` naming the construct, its span, and the §10 plane-2 error
//! it raises. Zig error sets cannot carry a payload and a refusal
//! without a span is not actionable, so the union is the channel;
//! `error.OutOfMemory` remains the only true error.
//!
//! Public API:
//!   parse(allocator, input, opts) -> Parsed { ok: Ast, refused: Refusal }
//!   Ast.print(allocator)          -> []u8      (canonical text)
//!   structurallyEqual(a, b)       -> bool

const std = @import("std");
const coords = @import("zlsx_refs");
const tokenizer = @import("tokenizer.zig");

const Token = tokenizer.Token;
const assert = std.debug.assert;

// ─── limits (§9) ─────────────────────────────────────────────────

/// The §9 rows that bind at parse time. Every one is a named, typed
/// refusal (`FormulaLimitExceeded`) with a below/at/above boundary test.
///
/// **Dominance at defaults** (proven by `test "limits: the dominance
/// order at defaults"`): a code point is at most 4 UTF-8 bytes and a
/// token is at least one code point, so `max_formula_chars` is the
/// binding limit for flat input — `max_formula_utf8_bytes`,
/// `max_tokens`, and `max_ast_nodes` can only be reached by lowering
/// them. They are still enforced, and still boundary-tested, because a
/// caller may lower any of them.
pub const Limits = struct {
    /// Unicode code points, Excel-aligned.
    max_formula_chars: usize = 8_192,
    /// Conservative zlsx cap on UTF-8 source bytes. Excel's documented
    /// "16 384 bytes" counts its internal *compiled* representation,
    /// which zlsx does not model; no equivalence is claimed.
    max_formula_utf8_bytes: usize = 32_768,
    max_tokens: usize = 16_384,
    max_ast_nodes: usize = 16_384,
    /// Nesting of any grammar production that recurses: parenthesised
    /// groups, call argument lists, array constants, and prefix
    /// operator chains.
    max_parse_depth: usize = 256,
    /// Function-call depth (Excel's documented limit).
    max_fn_nesting: usize = 64,
    max_args: usize = 255,
};

pub const LimitKind = enum {
    formula_chars,
    formula_utf8_bytes,
    tokens,
    ast_nodes,
    parse_depth,
    fn_nesting,
    args,
};

// ─── refusals (§10) ──────────────────────────────────────────────

/// §10 plane 2 — the refusal plane. Excel *error values* are plane 1
/// and are successful results; nothing here is an error value.
///
/// The full taxonomy is spelled out so the mapping lives in one place
/// and an exhaustive switch catches an addition. M2 raises only the
/// subset `raisableAtM2` names; the rest arrive with their milestones,
/// and `test "plane 2: M2 raises exactly its documented subset"` pins
/// that. At M5b2 this enum merges into `pkg/workbook.Error`.
pub const PlaneTwo = enum {
    FormulaUnsupportedFunction,
    FormulaUnsupportedConstruct,
    FormulaPrecisionAsDisplayed,
    FormulaMalformedInput,
    FormulaLocaleSensitiveInput,
    FormulaDataTableUnsupported,
    FormulaSignedWorkbook,
    FormulaStaleEmbeddings,
    FormulaAnchorRequired,
    FormulaCycle,
    FormulaDynamicRefUnstable,
    FormulaSpillPersistUnsupported,
    FormulaResultNotRepresentable,
    FormulaLimitExceeded,
};

/// The plane-2 errors a *parse* can produce. Everything else in
/// `PlaneTwo` needs a workbook, an evaluator, or a transaction.
pub const raisable_at_m2 = [_]PlaneTwo{
    .FormulaUnsupportedConstruct,
    .FormulaMalformedInput,
    .FormulaLocaleSensitiveInput,
    .FormulaLimitExceeded,
};

/// Why a parse refused. The first eight mirror
/// `tokenizer.Refusal.Reason` one-for-one — M1a classifies out-of-band
/// so its byte-for-byte round-trip survives, and M2 is where those
/// classifications become refusals with a §10 error attached.
pub const Reason = enum {
    // Lifted from the tokenizer, in its declaration order.
    invalid_utf8,
    identifier_too_long,
    backslash_after_start,
    r1c1_reference,
    external_reference,
    unterminated_string,
    unterminated_sheet_name,
    unterminated_structured_ref,

    // The parser's own.
    /// Nothing but whitespace (or nothing at all) after the optional
    /// leading `=`.
    empty_formula,
    /// `==…` — one leading `=` is stripped, a second is not an operand.
    double_equals,
    /// A token that cannot start or continue an expression here.
    unexpected_token,
    /// Input ran out mid-expression.
    unexpected_end,
    /// `(` with no `)`, or `{` with no `}`.
    unbalanced_delimiter,
    /// Tokens left over after a complete expression.
    trailing_input,
    /// `;` as an argument separator. The stored `<f>` grammar is
    /// en-US: `,` separates arguments and `;` separates array rows.
    /// A `;` anywhere else is locale-sensitive input.
    locale_separator,
    /// An array constant whose rows differ in length.
    ragged_array,
    /// An element position with nothing in it: `{}`, `{1,}`, `{1;}`.
    empty_array,
    /// An array constant element that is not a literal: a reference,
    /// a call, or a nested `{…}`.
    array_element_not_constant,
    /// A `[…]` specifier that does not match the item-specifier
    /// grammar.
    malformed_structured_ref,
    /// LAMBDA / LET — parsed, then refused (v1 non-goal).
    lambda_let_unsupported,
    /// `_xlpm.` — a LAMBDA/LET parameter name.
    xlpm_parameter,
    /// A §9 limit; `Refusal.limit` names which.
    limit_exceeded,
};

pub const Refusal = struct {
    reason: Reason,
    /// Byte offset of the offending span within the parser input.
    offset: u32,
    /// Byte length of the offending span. Zero is legal: an
    /// `unexpected_end` points one past the last byte.
    len: u32,
    /// Set exactly when `reason == .limit_exceeded`.
    limit: ?LimitKind = null,

    /// The §10 plane-2 error this refusal raises. Exhaustive by
    /// construction — a new `Reason` fails to compile until it is
    /// mapped.
    pub fn planeTwo(self: Refusal) PlaneTwo {
        return switch (self.reason) {
            .invalid_utf8,
            .backslash_after_start,
            .unterminated_string,
            .unterminated_sheet_name,
            .unterminated_structured_ref,
            .empty_formula,
            .double_equals,
            .unexpected_token,
            .unexpected_end,
            .unbalanced_delimiter,
            .trailing_input,
            .ragged_array,
            .empty_array,
            .array_element_not_constant,
            .malformed_structured_ref,
            => .FormulaMalformedInput,

            .r1c1_reference,
            .external_reference,
            .lambda_let_unsupported,
            .xlpm_parameter,
            => .FormulaUnsupportedConstruct,

            .locale_separator => .FormulaLocaleSensitiveInput,

            .identifier_too_long, .limit_exceeded => .FormulaLimitExceeded,
        };
    }
};

/// The §10 error M2 raises for a construct the tokenizer refused
/// out-of-band. Kept as its own function — and exhaustive over
/// `tokenizer.Refusal.Reason` — so a new tokenizer refusal cannot land
/// without a plane-2 decision.
pub fn reasonFromTokenizer(r: tokenizer.Refusal.Reason) Reason {
    return switch (r) {
        .invalid_utf8 => .invalid_utf8,
        .identifier_too_long => .identifier_too_long,
        .backslash_after_start => .backslash_after_start,
        .r1c1_reference => .r1c1_reference,
        .external_reference => .external_reference,
        .unterminated_string => .unterminated_string,
        .unterminated_sheet_name => .unterminated_sheet_name,
        .unterminated_structured_ref => .unterminated_structured_ref,
    };
}

// ─── name resolution (§5.9) ──────────────────────────────────────

/// How a name is used at its site. The parser decides this
/// syntactically — a name immediately followed by `(` is a call — and
/// resolution (M4b3) reads it to pick a resolution order.
pub const NameUse = enum { call, value };

/// §5.9 value-position resolution order, in the order it is tried.
/// Sheet-scoped names shadow workbook-scoped ones.
pub const ValueScope = enum {
    sheet_scoped_name,
    workbook_name,
    table,
    builtin_xlnm,
    /// Provably nowhere: `#NAME?`, a plane-1 error value.
    name_error,
};

pub const value_resolution_order = [_]ValueScope{
    .sheet_scoped_name,
    .workbook_name,
    .table,
    .builtin_xlnm,
    .name_error,
};

/// §5.9 call-position resolution order. Unregistered is a plane-2
/// refusal (`FormulaUnsupportedFunction`), not `#NAME?` — zlsx refuses
/// rather than inventing an error value for a function it simply does
/// not implement.
pub const CallStage = enum {
    strip_layered_prefixes,
    registry,
    unsupported_function,
};

pub const call_resolution_order = [_]CallStage{
    .strip_layered_prefixes,
    .registry,
    .unsupported_function,
};

/// All §5.9 matching is case-folded over the decoded symbol layer
/// (M4b1). At M2 there is no workbook, so the parser records the
/// spelling and the classification and resolves nothing.
pub const name_matching_policy = "case-folded over the decoded symbol layer (§5.9)";

/// The layered prefixes Excel writes ahead of a post-2007 function
/// name so older readers do not silently mis-evaluate it.
pub const Prefix = struct {
    /// `_xlfn.` — the function postdates the consumer's registry.
    xlfn: bool = false,
    /// `_xlws.` — additionally worksheet-only. Only ever layered
    /// *inside* `_xlfn.`, as `_xlfn._xlws.FILTER`.
    xlws: bool = false,

    pub fn eql(a: Prefix, b: Prefix) bool {
        return a.xlfn == b.xlfn and a.xlws == b.xlws;
    }
};

// ─── AST ─────────────────────────────────────────────────────────

pub const Index = u32;

/// Half-open byte span into `Ast.source`.
pub const Span = struct {
    start: u32,
    end: u32,

    pub fn eql(a: Span, b: Span) bool {
        return a.start == b.start and a.end == b.end;
    }
};

/// A run of child indices inside `Ast.extra`.
pub const ExtraSlice = struct { start: u32, len: u32 };

pub const Tag = enum {
    number,
    string,
    boolean,
    error_lit,
    missing_arg,
    array,
    ref_cell,
    ref_full_col,
    ref_full_row,
    name,
    qualified,
    structured,
    call,
    paren,
    unary,
    postfix,
    binary,
};

pub const UnaryOp = enum {
    plus,
    minus,
    /// Implicit intersection. Two surface forms, one operator — see
    /// `Unary.form`.
    implicit_intersection,
};

/// Which spelling an implicit intersection came from. `@` and
/// `_xlfn.SINGLE(x)` are the same operator; the AST unifies them and
/// this field is what lets the printer hand back the original.
pub const SingleForm = enum { at_operator, xlfn_single };

pub const PostfixOp = enum {
    /// `x%` — divide by 100.
    percent,
    /// `A1#` — the spilled range of a dynamic array.
    spill,
};

pub const BinaryOp = enum {
    range,
    intersect,
    union_op,
    pow,
    mul,
    div,
    add,
    sub,
    concat,
    eq,
    ne,
    lt,
    gt,
    le,
    ge,

    /// Canonical spelling. The intersection operator is a single space.
    pub fn text(self: BinaryOp) []const u8 {
        return switch (self) {
            .range => ":",
            .intersect => " ",
            .union_op => ",",
            .pow => "^",
            .mul => "*",
            .div => "/",
            .add => "+",
            .sub => "-",
            .concat => "&",
            .eq => "=",
            .ne => "<>",
            .lt => "<",
            .gt => ">",
            .le => "<=",
            .ge => ">=",
        };
    }
};

/// One end of a full-column reference (`$A:$B`).
pub const ColBound = struct { col: coords.Col, absolute: bool };

/// One end of a full-row reference (`$1:$5`).
pub const RowBound = struct { row: coords.Row, absolute: bool };

/// The item specifiers a structured reference may carry. `@` sets
/// `this_row` — it *is* `#This Row`; `Structured.at_shorthand` records
/// only which of the two spellings the source used.
pub const ItemSet = packed struct(u8) {
    all: bool = false,
    data: bool = false,
    headers: bool = false,
    totals: bool = false,
    this_row: bool = false,
    _pad: u3 = 0,

    pub fn bits(self: ItemSet) u8 {
        return @bitCast(self);
    }

    pub fn count(self: ItemSet) usize {
        return @popCount(self.bits());
    }
};

/// The column half of a structured reference. Names are kept **raw** —
/// with their `'` escapes intact — because that is what makes the
/// printer allocation-free and byte-exact; `decodeColumnName` resolves
/// the escapes for the symbol layer at M4b1.
pub const ColumnSelector = union(enum) {
    none,
    one: []const u8,
    range: struct { first: []const u8, last: []const u8 },
};

/// A qualifying sheet, or sheet span for a 3D reference. Raw text: the
/// quoted form keeps its quotes and its `''` escapes, because the
/// printer must hand back the spelling that was written.
///
/// A *quoted* 3D span is one token (`'Q1:Q4'!A1` — Excel puts the whole
/// span inside one quote pair), so it arrives here as `first` alone
/// with `last == null`. Splitting it is name resolution's job (M4b3),
/// not the parser's.
pub const SheetSpec = struct {
    first: []const u8,
    last: ?[]const u8 = null,
    quoted: bool = false,

    pub fn eql(a: SheetSpec, b: SheetSpec) bool {
        if (a.quoted != b.quoted) return false;
        if (!std.mem.eql(u8, a.first, b.first)) return false;
        if ((a.last == null) != (b.last == null)) return false;
        if (a.last) |al| return std.mem.eql(u8, al, b.last.?);
        return true;
    }
};

pub const Node = union(Tag) {
    /// Source spelling, unconverted. `parseDecimal` is M3a1's — a
    /// parser that converts here would have to pick a rounding policy
    /// before §5.4 pins one.
    number: struct { span: Span, text: []const u8 },
    /// Including the delimiting quotes and any `""` escapes. Decoding
    /// is the M4b1 boundary's.
    string: struct { span: Span, text: []const u8 },
    boolean: struct { span: Span, value: bool },
    error_lit: struct {
        span: Span,
        text: []const u8,
        /// One of the frozen ten, versus a spelling that reached the
        /// tokenizer through the extensible rule.
        known: bool,
    },
    /// An omitted argument: `IF(A1,,2)`. Zero-length span at the
    /// position where the argument would have been.
    missing_arg: struct { span: Span },
    array: struct {
        span: Span,
        rows: u32,
        cols: u32,
        /// `rows * cols` indices, row-major.
        elems: ExtraSlice,
    },
    ref_cell: struct { span: Span, cell: coords.Cell, text: []const u8 },
    ref_full_col: struct { span: Span, first: ColBound, last: ColBound },
    ref_full_row: struct { span: Span, first: RowBound, last: RowBound },
    name: struct {
        span: Span,
        /// Original spelling, prefixes included. What the printer emits.
        raw: []const u8,
        /// `raw` with the layered `_xlfn.` / `_xlws.` prefixes removed.
        bare: []const u8,
        prefix: Prefix,
        use: NameUse,
        /// `_xlnm.Print_Area` and friends. The `_xlnm.` is part of the
        /// name, not a prefix layer, so it stays in `bare`.
        builtin_xlnm: bool,
    },
    qualified: struct { span: Span, sheet: SheetSpec, target: Index },
    structured: struct {
        span: Span,
        /// Raw table name, or null for the bare `[…]` same-table form.
        table: ?[]const u8,
        items: ItemSet,
        at_shorthand: bool,
        columns: ColumnSelector,
    },
    call: struct { span: Span, callee: Index, args: ExtraSlice },
    paren: struct { span: Span, child: Index },
    unary: struct {
        span: Span,
        op: UnaryOp,
        child: Index,
        form: SingleForm = .at_operator,
        /// The callee spelling for `form == .xlfn_single`, so the
        /// printer can reproduce `_xlfn.SINGLE(…)` exactly.
        callee_raw: []const u8 = "",
    },
    postfix: struct { span: Span, op: PostfixOp, child: Index },
    binary: struct { span: Span, op: BinaryOp, lhs: Index, rhs: Index },

    pub fn span(self: Node) Span {
        return switch (self) {
            inline else => |payload| payload.span,
        };
    }
};

pub const Ast = struct {
    /// The whole parser input, borrowed. Every text slice in every node
    /// points into it.
    source: []const u8,
    /// The expression body — `source` minus any stripped leading `=`
    /// and its surrounding whitespace.
    body: Span,
    nodes: []const Node,
    extra: []const Index,
    root: Index,

    pub fn deinit(self: *Ast, allocator: std.mem.Allocator) void {
        allocator.free(self.nodes);
        allocator.free(self.extra);
        self.* = undefined;
    }

    pub fn node(self: Ast, i: Index) Node {
        return self.nodes[i];
    }

    pub fn children(self: Ast, s: ExtraSlice) []const Index {
        return self.extra[s.start .. s.start + s.len];
    }

    /// Canonical text for the whole tree. Caller frees.
    pub fn print(self: Ast, allocator: std.mem.Allocator) error{OutOfMemory}![]u8 {
        var out: std.ArrayListUnmanaged(u8) = .empty;
        errdefer out.deinit(allocator);
        try self.printNode(allocator, &out, self.root);
        return out.toOwnedSlice(allocator);
    }

    fn printNode(
        self: Ast,
        allocator: std.mem.Allocator,
        out: *std.ArrayListUnmanaged(u8),
        i: Index,
    ) error{OutOfMemory}!void {
        switch (self.nodes[i]) {
            .number => |n| try out.appendSlice(allocator, n.text),
            .string => |n| try out.appendSlice(allocator, n.text),
            .boolean => |n| try out.appendSlice(allocator, if (n.value) "TRUE" else "FALSE"),
            .error_lit => |n| try out.appendSlice(allocator, n.text),
            .missing_arg => {},
            .array => |n| {
                try out.append(allocator, '{');
                const elems = self.children(n.elems);
                var r: u32 = 0;
                while (r < n.rows) : (r += 1) {
                    if (r != 0) try out.append(allocator, ';');
                    var c: u32 = 0;
                    while (c < n.cols) : (c += 1) {
                        if (c != 0) try out.append(allocator, ',');
                        try self.printNode(allocator, out, elems[r * n.cols + c]);
                    }
                }
                try out.append(allocator, '}');
            },
            .ref_cell => |n| try out.appendSlice(allocator, n.text),
            .ref_full_col => |n| {
                try printColBound(allocator, out, n.first);
                try out.append(allocator, ':');
                try printColBound(allocator, out, n.last);
            },
            .ref_full_row => |n| {
                try printRowBound(allocator, out, n.first);
                try out.append(allocator, ':');
                try printRowBound(allocator, out, n.last);
            },
            .name => |n| try out.appendSlice(allocator, n.raw),
            .qualified => |n| {
                try out.appendSlice(allocator, n.sheet.first);
                if (n.sheet.last) |last| {
                    try out.append(allocator, ':');
                    try out.appendSlice(allocator, last);
                }
                try out.append(allocator, '!');
                try self.printNode(allocator, out, n.target);
            },
            .structured => |n| {
                if (n.table) |t| try out.appendSlice(allocator, t);
                try printStructuredSpec(allocator, out, n.items, n.at_shorthand, n.columns);
            },
            .call => |n| {
                try self.printNode(allocator, out, n.callee);
                try out.append(allocator, '(');
                for (self.children(n.args), 0..) |arg, k| {
                    if (k != 0) try out.append(allocator, ',');
                    try self.printNode(allocator, out, arg);
                }
                try out.append(allocator, ')');
            },
            .paren => |n| {
                try out.append(allocator, '(');
                try self.printNode(allocator, out, n.child);
                try out.append(allocator, ')');
            },
            .unary => |n| {
                switch (n.op) {
                    .plus => try out.append(allocator, '+'),
                    .minus => try out.append(allocator, '-'),
                    .implicit_intersection => switch (n.form) {
                        .at_operator => try out.append(allocator, '@'),
                        .xlfn_single => {
                            try out.appendSlice(allocator, n.callee_raw);
                            try out.append(allocator, '(');
                        },
                    },
                }
                try self.printNode(allocator, out, n.child);
                if (n.op == .implicit_intersection and n.form == .xlfn_single) {
                    try out.append(allocator, ')');
                }
            },
            .postfix => |n| {
                try self.printNode(allocator, out, n.child);
                try out.append(allocator, switch (n.op) {
                    .percent => '%',
                    .spill => '#',
                });
            },
            .binary => |n| {
                try self.printNode(allocator, out, n.lhs);
                try out.appendSlice(allocator, n.op.text());
                try self.printNode(allocator, out, n.rhs);
            },
        }
    }
};

fn printColBound(
    allocator: std.mem.Allocator,
    out: *std.ArrayListUnmanaged(u8),
    b: ColBound,
) error{OutOfMemory}!void {
    if (b.absolute) try out.append(allocator, '$');
    var buf: [coords.format_buf_len]u8 = undefined;
    const n = coords.writeColLetters(&buf, b.col);
    try out.appendSlice(allocator, buf[0..n]);
}

fn printRowBound(
    allocator: std.mem.Allocator,
    out: *std.ArrayListUnmanaged(u8),
    b: RowBound,
) error{OutOfMemory}!void {
    if (b.absolute) try out.append(allocator, '$');
    var buf: [16]u8 = undefined;
    const s = std.fmt.bufPrint(&buf, "{d}", .{b.row.oneBased()}) catch unreachable;
    try out.appendSlice(allocator, s);
}

/// True when a raw structured-ref column name has to be wrapped in its
/// own `[…]`.
///
/// The bare form is reserved for *simple* names — letters, digits, `_`,
/// `.`, and non-ASCII. That is stricter than the grammar strictly needs
/// (only `,` and `:` are genuinely ambiguous), and deliberately so:
/// Excel brackets anything carrying a space or a punctuation byte, and
/// a canonical form that produced `Table1[@Col A]` where Excel writes
/// `Table1[@[Col A]]` would be a spelling no other reader expects.
fn columnNeedsBrackets(raw: []const u8) bool {
    if (raw.len == 0) return true;
    for (raw) |c| {
        const simple = (c >= 'a' and c <= 'z') or (c >= 'A' and c <= 'Z') or
            (c >= '0' and c <= '9') or c == '_' or c == '.' or c >= 0x80;
        if (!simple) return true;
    }
    return false;
}

fn printStructuredSpec(
    allocator: std.mem.Allocator,
    out: *std.ArrayListUnmanaged(u8),
    items: ItemSet,
    at_shorthand: bool,
    columns: ColumnSelector,
) error{OutOfMemory}!void {
    // `@` prints in place of the `#This Row` it stands for.
    var printed_items = items;
    if (at_shorthand) printed_items.this_row = false;

    const col_parts: usize = switch (columns) {
        .none => 0,
        .one => 1,
        .range => 2,
    };
    const parts = printed_items.count() + col_parts;

    try out.append(allocator, '[');
    if (at_shorthand) try out.append(allocator, '@');

    // One part needs no inner brackets; two or more are always
    // bracketed and comma-joined, which is the only form that can carry
    // a separator unambiguously.
    const wrap = parts > 1;
    var written: usize = 0;
    inline for (.{
        .{ "all", "#All" },
        .{ "data", "#Data" },
        .{ "headers", "#Headers" },
        .{ "totals", "#Totals" },
        .{ "this_row", "#This Row" },
    }) |pair| {
        if (@field(printed_items, pair[0])) {
            if (written != 0) try out.append(allocator, ',');
            if (wrap) try out.append(allocator, '[');
            try out.appendSlice(allocator, pair[1]);
            if (wrap) try out.append(allocator, ']');
            written += 1;
        }
    }
    switch (columns) {
        .none => {},
        .one => |raw| {
            if (written != 0) try out.append(allocator, ',');
            if (wrap or columnNeedsBrackets(raw)) {
                try out.append(allocator, '[');
                try out.appendSlice(allocator, raw);
                try out.append(allocator, ']');
            } else {
                try out.appendSlice(allocator, raw);
            }
        },
        .range => |r| {
            if (written != 0) try out.append(allocator, ',');
            try out.append(allocator, '[');
            try out.appendSlice(allocator, r.first);
            try out.appendSlice(allocator, "]:[");
            try out.appendSlice(allocator, r.last);
            try out.append(allocator, ']');
        },
    }
    try out.append(allocator, ']');
}

/// Resolve the `'` escapes in a raw structured-ref column name. The
/// symbol layer (M4b1) matches on the decoded form; the AST keeps the
/// raw one. Caller frees.
pub fn decodeColumnName(
    allocator: std.mem.Allocator,
    raw: []const u8,
) error{OutOfMemory}![]u8 {
    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(allocator);
    var i: usize = 0;
    while (i < raw.len) : (i += 1) {
        if (raw[i] == '\'' and i + 1 < raw.len) {
            i += 1;
        }
        try out.append(allocator, raw[i]);
    }
    return out.toOwnedSlice(allocator);
}

// ─── structural equality ─────────────────────────────────────────

/// Structural equality of two trees. Spans are deliberately **not**
/// compared: the printer's output has different offsets from the source
/// it was printed from, and the round-trip property is about structure.
/// Everything a canonical print can carry is compared, including the
/// spellings the AST preserves verbatim.
pub fn structurallyEqual(a: Ast, b: Ast) bool {
    return nodesEqual(a, a.root, b, b.root);
}

fn nodesEqual(a: Ast, ai: Index, b: Ast, bi: Index) bool {
    const x = a.nodes[ai];
    const y = b.nodes[bi];
    if (@as(Tag, x) != @as(Tag, y)) return false;
    switch (x) {
        .number => |n| return std.mem.eql(u8, n.text, y.number.text),
        .string => |n| return std.mem.eql(u8, n.text, y.string.text),
        .boolean => |n| return n.value == y.boolean.value,
        .error_lit => |n| return n.known == y.error_lit.known and
            std.mem.eql(u8, n.text, y.error_lit.text),
        .missing_arg => return true,
        .array => |n| {
            const m = y.array;
            if (n.rows != m.rows or n.cols != m.cols) return false;
            const xs = a.children(n.elems);
            const ys = b.children(m.elems);
            if (xs.len != ys.len) return false;
            for (xs, ys) |p, q| {
                if (!nodesEqual(a, p, b, q)) return false;
            }
            return true;
        },
        .ref_cell => |n| return n.cell.eqlExact(y.ref_cell.cell),
        .ref_full_col => |n| {
            const m = y.ref_full_col;
            return n.first.col == m.first.col and n.first.absolute == m.first.absolute and
                n.last.col == m.last.col and n.last.absolute == m.last.absolute;
        },
        .ref_full_row => |n| {
            const m = y.ref_full_row;
            return n.first.row == m.first.row and n.first.absolute == m.first.absolute and
                n.last.row == m.last.row and n.last.absolute == m.last.absolute;
        },
        .name => |n| {
            const m = y.name;
            return std.mem.eql(u8, n.raw, m.raw) and
                std.mem.eql(u8, n.bare, m.bare) and
                n.prefix.eql(m.prefix) and n.use == m.use and
                n.builtin_xlnm == m.builtin_xlnm;
        },
        .qualified => |n| {
            const m = y.qualified;
            return n.sheet.eql(m.sheet) and nodesEqual(a, n.target, b, m.target);
        },
        .structured => |n| {
            const m = y.structured;
            if ((n.table == null) != (m.table == null)) return false;
            if (n.table) |t| {
                if (!std.mem.eql(u8, t, m.table.?)) return false;
            }
            if (n.items.bits() != m.items.bits()) return false;
            if (n.at_shorthand != m.at_shorthand) return false;
            if (@as(std.meta.Tag(ColumnSelector), n.columns) !=
                @as(std.meta.Tag(ColumnSelector), m.columns)) return false;
            return switch (n.columns) {
                .none => true,
                .one => |raw| std.mem.eql(u8, raw, m.columns.one),
                .range => |r| std.mem.eql(u8, r.first, m.columns.range.first) and
                    std.mem.eql(u8, r.last, m.columns.range.last),
            };
        },
        .call => |n| {
            const m = y.call;
            if (!nodesEqual(a, n.callee, b, m.callee)) return false;
            const xs = a.children(n.args);
            const ys = b.children(m.args);
            if (xs.len != ys.len) return false;
            for (xs, ys) |p, q| {
                if (!nodesEqual(a, p, b, q)) return false;
            }
            return true;
        },
        .paren => |n| return nodesEqual(a, n.child, b, y.paren.child),
        .unary => |n| {
            const m = y.unary;
            return n.op == m.op and n.form == m.form and
                std.mem.eql(u8, n.callee_raw, m.callee_raw) and
                nodesEqual(a, n.child, b, m.child);
        },
        .postfix => |n| {
            const m = y.postfix;
            return n.op == m.op and nodesEqual(a, n.child, b, m.child);
        },
        .binary => |n| {
            const m = y.binary;
            return n.op == m.op and nodesEqual(a, n.lhs, b, m.lhs) and
                nodesEqual(a, n.rhs, b, m.rhs);
        },
    }
}

// ─── entry point ─────────────────────────────────────────────────

pub const LeadingEq = enum {
    /// The stored `<f>` form: the body only, a leading `=` is an
    /// operator with a missing left operand.
    forbid,
    /// A standalone formula: exactly one optional `=` after leading
    /// whitespace is stripped. `==` refuses.
    optional,
};

pub const Options = struct {
    leading_eq: LeadingEq = .optional,
    limits: Limits = .{},
};

pub const Parsed = union(enum) {
    ok: Ast,
    refused: Refusal,

    pub fn deinit(self: *Parsed, allocator: std.mem.Allocator) void {
        switch (self.*) {
            .ok => |*ast| ast.deinit(allocator),
            .refused => {},
        }
        self.* = undefined;
    }
};

pub const Error = error{OutOfMemory};

/// Parse `input`. The returned AST borrows every text slice from
/// `input`, which must outlive it; free with `parsed.deinit(allocator)`.
pub fn parse(
    allocator: std.mem.Allocator,
    input: []const u8,
    opts: Options,
) Error!Parsed {
    if (input.len > opts.limits.max_formula_utf8_bytes) {
        return refuseLimit(.formula_utf8_bytes, 0, input.len);
    }
    const cps = countCodepoints(input);
    if (cps > opts.limits.max_formula_chars) {
        return refuseLimit(.formula_chars, 0, input.len);
    }

    var scanned = try tokenizer.scan(allocator, input);
    defer scanned.deinit(allocator);

    // A construct M1a preserved but classified as unacceptable becomes a
    // refusal here, before any parse, so the diagnostic names the
    // construct rather than whatever syntax error it happens to cause.
    // First in detection order wins.
    if (scanned.refusals.len > 0) {
        const first = scanned.refusals[0];
        return .{ .refused = .{
            .reason = reasonFromTokenizer(first.reason),
            .offset = @intCast(first.offset),
            .len = @intCast(first.len),
        } };
    }

    if (scanned.tokens.len > opts.limits.max_tokens) {
        return refuseLimit(.tokens, 0, input.len);
    }

    var p: Parser = .{
        .allocator = allocator,
        .input = input,
        .tokens = scanned.tokens,
        .limits = opts.limits,
    };
    errdefer {
        p.nodes.deinit(allocator);
        p.extra.deinit(allocator);
    }

    const body_start = p.stripPreamble(opts.leading_eq) catch |err| switch (err) {
        error.OutOfMemory => return error.OutOfMemory,
        error.Refused => return .{ .refused = p.refusal.? },
    };

    const root = p.parseFormula() catch |err| switch (err) {
        error.OutOfMemory => return error.OutOfMemory,
        error.Refused => {
            p.nodes.deinit(allocator);
            p.extra.deinit(allocator);
            return .{ .refused = p.refusal.? };
        },
    };

    // No lost bytes. The cursor only moves forward (speculative matches
    // rewind `bytes_consumed` with it), every step adds the token's
    // length, and a successful parse must have consumed every token —
    // `parseFormula` refuses `trailing_input` otherwise. Since the
    // tokenizer tiles the input exactly, the total is a proof that no
    // byte was skipped past.
    assert(p.bytes_consumed == input.len);

    const nodes = try p.nodes.toOwnedSlice(allocator);
    errdefer allocator.free(nodes);
    const extra = try p.extra.toOwnedSlice(allocator);

    return .{ .ok = .{
        .source = input,
        .body = .{ .start = @intCast(body_start), .end = @intCast(input.len) },
        .nodes = nodes,
        .extra = extra,
        .root = root,
    } };
}

fn refuseLimit(kind: LimitKind, offset: usize, len: usize) Parsed {
    return .{ .refused = .{
        .reason = .limit_exceeded,
        .offset = @intCast(offset),
        .len = @intCast(len),
        .limit = kind,
    } };
}

/// Count code points without requiring valid UTF-8: every byte that is
/// not a continuation byte starts one. Exact for well-formed input, and
/// well-defined for the malformed input the tokenizer is also required
/// to survive — the limit check runs before tokenization, so it cannot
/// assume validity.
fn countCodepoints(s: []const u8) usize {
    var n: usize = 0;
    for (s) |b| {
        if (b & 0xC0 != 0x80) n += 1;
    }
    return n;
}

// ─── parser ──────────────────────────────────────────────────────

const ParseError = error{ OutOfMemory, Refused };

const Mark = struct { ti: usize, bytes: usize, last_end: usize };

const Parser = struct {
    allocator: std.mem.Allocator,
    input: []const u8,
    tokens: []const Token,
    limits: Limits,

    ti: usize = 0,
    bytes_consumed: usize = 0,
    /// Byte offset just past the last consumed token. Node spans end
    /// here rather than at the cursor, so a node's span never includes
    /// trailing trivia the next level speculatively skipped.
    last_end: usize = 0,
    depth: usize = 0,
    fn_depth: usize = 0,
    /// Inside a call's argument list, `,` ends the argument instead of
    /// forming a reference union. A parenthesised group resets it —
    /// `SUM((A1,B2))` is how a union reaches a function.
    in_arg_list: bool = false,

    nodes: std.ArrayListUnmanaged(Node) = .empty,
    extra: std.ArrayListUnmanaged(Index) = .empty,
    refusal: ?Refusal = null,

    // ── cursor ──

    fn cur(self: *const Parser) ?Token {
        if (self.ti >= self.tokens.len) return null;
        return self.tokens[self.ti];
    }

    fn at(self: *const Parser, n: usize) ?Token {
        if (self.ti + n >= self.tokens.len) return null;
        return self.tokens[self.ti + n];
    }

    fn kindAt(self: *const Parser, n: usize) ?Token.Kind {
        const t = self.at(n) orelse return null;
        return t.kind;
    }

    fn offset(self: *const Parser) usize {
        if (self.ti >= self.tokens.len) return self.input.len;
        const t = self.tokens[self.ti];
        return @intFromPtr(t.text.ptr) - @intFromPtr(self.input.ptr);
    }

    fn advance(self: *Parser) Token {
        const t = self.tokens[self.ti];
        self.ti += 1;
        self.bytes_consumed += t.text.len;
        self.last_end = (@intFromPtr(t.text.ptr) - @intFromPtr(self.input.ptr)) + t.text.len;
        return t;
    }

    fn mark(self: *const Parser) Mark {
        return .{ .ti = self.ti, .bytes = self.bytes_consumed, .last_end = self.last_end };
    }

    fn reset(self: *Parser, m: Mark) void {
        self.ti = m.ti;
        self.bytes_consumed = m.bytes;
        self.last_end = m.last_end;
    }

    /// Consume whitespace as trivia. Only ever called where whitespace
    /// cannot be the intersection operator — before an operand, and
    /// speculatively around an infix operator that a failed match
    /// rewinds.
    fn skipTrivia(self: *Parser) void {
        while (self.cur()) |t| {
            if (t.kind != .whitespace) break;
            _ = self.advance();
        }
    }

    fn spanFrom(self: *const Parser, start: usize) Span {
        return .{
            .start = @intCast(start),
            .end = @intCast(@max(start, self.last_end)),
        };
    }

    fn bytesSince(self: *const Parser, start: usize) usize {
        return if (self.last_end > start) self.last_end - start else 0;
    }

    // ── refusal ──

    fn fail(self: *Parser, reason: Reason, off: usize, len: usize) ParseError {
        if (self.refusal == null) {
            self.refusal = .{
                .reason = reason,
                .offset = @intCast(off),
                .len = @intCast(len),
            };
        }
        return error.Refused;
    }

    fn failHere(self: *Parser, reason: Reason) ParseError {
        const off = self.offset();
        const len = if (self.cur()) |t| t.text.len else 0;
        return self.fail(reason, off, len);
    }

    fn failLimit(self: *Parser, kind: LimitKind, off: usize, len: usize) ParseError {
        if (self.refusal == null) {
            self.refusal = .{
                .reason = .limit_exceeded,
                .offset = @intCast(off),
                .len = @intCast(len),
                .limit = kind,
            };
        }
        return error.Refused;
    }

    // ── node construction ──

    fn addNode(self: *Parser, n: Node) ParseError!Index {
        if (self.nodes.items.len >= self.limits.max_ast_nodes) {
            return self.failLimit(.ast_nodes, 0, self.input.len);
        }
        const i: Index = @intCast(self.nodes.items.len);
        try self.nodes.append(self.allocator, n);
        return i;
    }

    fn addExtra(self: *Parser, items: []const Index) ParseError!ExtraSlice {
        const start: u32 = @intCast(self.extra.items.len);
        try self.extra.appendSlice(self.allocator, items);
        return .{ .start = start, .len = @intCast(items.len) };
    }

    fn enter(self: *Parser) ParseError!void {
        self.depth += 1;
        if (self.depth > self.limits.max_parse_depth) {
            return self.failLimit(.parse_depth, self.offset(), 0);
        }
    }

    fn leave(self: *Parser) void {
        assert(self.depth > 0);
        self.depth -= 1;
    }

    // ── preamble ──

    /// Strip leading whitespace and, when allowed, exactly one `=`.
    /// Returns the byte offset where the body begins.
    fn stripPreamble(self: *Parser, mode: LeadingEq) ParseError!usize {
        self.skipTrivia();
        if (mode == .optional) {
            if (self.cur()) |t| {
                if (t.kind == .op_eq) {
                    _ = self.advance();
                    self.skipTrivia();
                    if (self.cur()) |next| {
                        if (next.kind == .op_eq) {
                            return self.fail(.double_equals, self.offset(), next.text.len);
                        }
                    }
                }
            }
        }
        if (self.ti >= self.tokens.len) {
            return self.fail(.empty_formula, self.input.len, 0);
        }
        return self.offset();
    }

    fn parseFormula(self: *Parser) ParseError!Index {
        const root = try self.parseComparison();
        self.skipTrivia();
        if (self.ti < self.tokens.len) {
            return self.failHere(.trailing_input);
        }
        return root;
    }

    // ── precedence ladder ──
    //
    // Loosest first. Each level parses the next-tighter level for its
    // operands, so the chain itself *is* the precedence table:
    //
    //   comparison < concat < additive < multiplicative < power
    //     < percent < unary < union < intersection < range < primary

    fn parseComparison(self: *Parser) ParseError!Index {
        var lhs = try self.parseConcat();
        while (self.cur()) |t| {
            const op: BinaryOp = switch (t.kind) {
                .op_eq => .eq,
                .op_ne => .ne,
                .op_lt => .lt,
                .op_gt => .gt,
                .op_le => .le,
                .op_ge => .ge,
                else => break,
            };
            _ = self.advance();
            const rhs = try self.parseConcat();
            lhs = try self.combine(op, lhs, rhs);
        }
        return lhs;
    }

    fn parseConcat(self: *Parser) ParseError!Index {
        var lhs = try self.parseAdditive();
        while (self.cur()) |t| {
            if (t.kind != .op_concat) break;
            _ = self.advance();
            const rhs = try self.parseAdditive();
            lhs = try self.combine(.concat, lhs, rhs);
        }
        return lhs;
    }

    fn parseAdditive(self: *Parser) ParseError!Index {
        var lhs = try self.parseMultiplicative();
        while (self.cur()) |t| {
            const op: BinaryOp = switch (t.kind) {
                .op_plus => .add,
                .op_minus => .sub,
                else => break,
            };
            _ = self.advance();
            const rhs = try self.parseMultiplicative();
            lhs = try self.combine(op, lhs, rhs);
        }
        return lhs;
    }

    fn parseMultiplicative(self: *Parser) ParseError!Index {
        var lhs = try self.parsePower();
        while (self.cur()) |t| {
            const op: BinaryOp = switch (t.kind) {
                .op_mul => .mul,
                .op_div => .div,
                else => break,
            };
            _ = self.advance();
            const rhs = try self.parsePower();
            lhs = try self.combine(op, lhs, rhs);
        }
        return lhs;
    }

    /// Left-associative: `2^3^2` is `(2^3)^2` = 64, which is what the
    /// oracle manifests record. A right-associative `^` would give 512.
    fn parsePower(self: *Parser) ParseError!Index {
        var lhs = try self.parsePercent();
        while (self.cur()) |t| {
            if (t.kind != .op_pow) break;
            _ = self.advance();
            const rhs = try self.parsePercent();
            lhs = try self.combine(.pow, lhs, rhs);
        }
        return lhs;
    }

    fn parsePercent(self: *Parser) ParseError!Index {
        var e = try self.parseUnary();
        while (self.cur()) |t| {
            if (t.kind != .op_percent) break;
            _ = self.advance();
            const start = self.nodes.items[e].span().start;
            e = try self.addNode(.{ .postfix = .{
                .span = self.spanFrom(start),
                .op = .percent,
                .child = e,
            } });
        }
        return e;
    }

    /// Unary `±` binds tighter than `^`, so `-1^2` is `(-1)^2` = 1 —
    /// again the oracle's recorded value, not an assumption. Prefix
    /// chains are right-associative and recurse, hence the depth guard.
    fn parseUnary(self: *Parser) ParseError!Index {
        self.skipTrivia();
        const t = self.cur() orelse return self.fail(.unexpected_end, self.input.len, 0);
        const op: UnaryOp = switch (t.kind) {
            .op_plus => .plus,
            .op_minus => .minus,
            else => return self.parseUnion(),
        };
        try self.enter();
        defer self.leave();
        const start = self.offset();
        _ = self.advance();
        const child = try self.parseUnary();
        return self.addNode(.{ .unary = .{
            .span = self.spanFrom(start),
            .op = op,
            .child = child,
        } });
    }

    /// `,` — reference union. Also where the `@` implicit-intersection
    /// prefix binds: its operand reaches down to the intersection level,
    /// so `@A1:A10` intersects the whole range rather than just `A1`.
    fn parseUnion(self: *Parser) ParseError!Index {
        self.skipTrivia();
        if (self.cur()) |t| {
            if (t.kind == .op_at) {
                try self.enter();
                defer self.leave();
                const start = self.offset();
                _ = self.advance();
                const child = try self.parseIntersection();
                return self.addNode(.{ .unary = .{
                    .span = self.spanFrom(start),
                    .op = .implicit_intersection,
                    .child = child,
                    .form = .at_operator,
                } });
            }
        }
        var lhs = try self.parseIntersection();
        while (self.cur()) |t| {
            if (t.kind != .arg_sep) break;
            // Outside `{…}` the stored grammar is en-US: `,` is the
            // separator, `;` is not a separator at all.
            if (t.text[0] == ';') return self.failHere(.locale_separator);
            if (self.in_arg_list) break;
            _ = self.advance();
            const rhs = try self.parseIntersection();
            lhs = try self.combine(.union_op, lhs, rhs);
        }
        return lhs;
    }

    /// The intersection operator is a space, so this level is also the
    /// only one that consumes whitespace *after* an operand: whitespace
    /// followed by something that cannot start a primary is trivia.
    /// `+` and `-` are deliberately excluded from the starter set —
    /// `A1 -1` is subtraction in Excel, not an intersection with `-1`.
    fn parseIntersection(self: *Parser) ParseError!Index {
        var lhs = try self.parseRange();
        while (true) {
            if (self.cur()) |t| {
                if (t.kind != .whitespace) break;
            } else break;

            self.skipTrivia();
            if (self.cur()) |next| {
                if (startsPrimary(next.kind)) {
                    const rhs = try self.parseRange();
                    lhs = try self.combine(.intersect, lhs, rhs);
                    continue;
                }
            }
            // Not an intersection. The whitespace stays consumed as
            // trivia — it is accounted for either way, and rewinding it
            // would only make a later level consume it again.
            break;
        }
        return lhs;
    }

    fn parseRange(self: *Parser) ParseError!Index {
        var lhs = try self.parsePrimary();
        while (true) {
            const m = self.mark();
            self.skipTrivia();
            const t = self.cur() orelse {
                self.reset(m);
                break;
            };
            if (t.kind != .op_range) {
                self.reset(m);
                break;
            }
            _ = self.advance();
            const rhs = try self.parsePrimary();
            lhs = try self.combine(.range, lhs, rhs);
        }
        return lhs;
    }

    fn combine(self: *Parser, op: BinaryOp, lhs: Index, rhs: Index) ParseError!Index {
        const start = self.nodes.items[lhs].span().start;
        const end = self.nodes.items[rhs].span().end;
        return self.addNode(.{ .binary = .{
            .span = .{ .start = start, .end = end },
            .op = op,
            .lhs = lhs,
            .rhs = rhs,
        } });
    }

    // ── primaries ──

    fn parsePrimary(self: *Parser) ParseError!Index {
        self.skipTrivia();
        const start = self.offset();
        const sheet = try self.tryParseSheetPrefix();
        var node = try self.parseCore(sheet != null);
        if (sheet) |s| {
            node = try self.addNode(.{ .qualified = .{
                .span = self.spanFrom(start),
                .sheet = s,
                .target = node,
            } });
        }
        // `A1#` — adjacency is required, so no trivia skip here.
        if (self.cur()) |t| {
            if (t.kind == .op_spill) {
                _ = self.advance();
                node = try self.addNode(.{ .postfix = .{
                    .span = self.spanFrom(start),
                    .op = .spill,
                    .child = node,
                } });
            }
        }
        return node;
    }

    /// `Sheet1!`, `'My Sheet'!`, or the 3D span `Sheet1:Sheet3!`.
    /// Adjacency is required throughout — Excel's stored form never
    /// spaces a qualifier, and requiring it keeps this lookahead from
    /// stealing the `:` of a full-column reference.
    fn tryParseSheetPrefix(self: *Parser) ParseError!?SheetSpec {
        const t = self.cur() orelse return null;
        const quoted = switch (t.kind) {
            .sheet_name => true,
            .name => false,
            else => return null,
        };
        if (self.kindAt(1)) |k1| {
            if (k1 == .bang) {
                _ = self.advance();
                _ = self.advance();
                return .{ .first = t.text, .last = null, .quoted = quoted };
            }
            if (k1 == .op_range) {
                const k2 = self.kindAt(2) orelse return null;
                if (k2 != .name and k2 != .sheet_name) return null;
                const k3 = self.kindAt(3) orelse return null;
                if (k3 != .bang) return null;
                const last = self.at(2).?.text;
                _ = self.advance();
                _ = self.advance();
                _ = self.advance();
                _ = self.advance();
                return .{ .first = t.text, .last = last, .quoted = quoted };
            }
        }
        return null;
    }

    fn parseCore(self: *Parser, qualified: bool) ParseError!Index {
        if (try self.tryFullColSpan()) |n| return n;
        if (try self.tryFullRowSpan()) |n| return n;

        const t = self.cur() orelse return self.fail(.unexpected_end, self.input.len, 0);
        const start = self.offset();
        switch (t.kind) {
            .number => {
                if (qualified) return self.failHere(.unexpected_token);
                _ = self.advance();
                return self.addNode(.{ .number = .{
                    .span = self.spanFrom(start),
                    .text = t.text,
                } });
            },
            .string => {
                if (qualified) return self.failHere(.unexpected_token);
                _ = self.advance();
                return self.addNode(.{ .string = .{
                    .span = self.spanFrom(start),
                    .text = t.text,
                } });
            },
            .bool_lit => {
                if (qualified) return self.failHere(.unexpected_token);
                _ = self.advance();
                return self.addNode(.{ .boolean = .{
                    .span = self.spanFrom(start),
                    .value = t.text[0] == 'T' or t.text[0] == 't',
                } });
            },
            .error_lit => {
                _ = self.advance();
                return self.addNode(.{ .error_lit = .{
                    .span = self.spanFrom(start),
                    .text = t.text,
                    .known = tokenizer.isKnownErrorLiteral(t.text),
                } });
            },
            .cell_ref => {
                _ = self.advance();
                const cell = coords.parseCell(t.text, .{
                    .case = .insensitive,
                    .leading_zero_row = .reject,
                    .dollar = .accept,
                }) catch return self.fail(.unexpected_token, start, t.text.len);
                return self.addNode(.{ .ref_cell = .{
                    .span = self.spanFrom(start),
                    .cell = cell,
                    .text = t.text,
                } });
            },
            .structured_ref => {
                if (qualified) return self.failHere(.unexpected_token);
                return self.parseStructured(null, start);
            },
            .name => return self.parseNameLike(qualified, start),
            .paren_open => {
                if (qualified) return self.failHere(.unexpected_token);
                return self.parseParen(start);
            },
            .array_open => {
                if (qualified) return self.failHere(.unexpected_token);
                return self.parseArray(start);
            },
            .arg_sep => {
                if (t.text[0] == ';') return self.failHere(.locale_separator);
                return self.failHere(.unexpected_token);
            },
            .external_ref => return self.fail(.external_reference, start, t.text.len),
            else => return self.failHere(.unexpected_token),
        }
    }

    fn parseNameLike(self: *Parser, qualified: bool, start: usize) ParseError!Index {
        const t = self.cur().?;
        // A structured reference's table name: `Table1[…]`, adjacent.
        if (!qualified) {
            if (self.kindAt(1)) |k1| {
                if (k1 == .structured_ref) {
                    _ = self.advance();
                    return self.parseStructured(t.text, start);
                }
            }
        }
        // A call: `SUM(`, adjacent. Adjacency matches the tokenizer's
        // own normative classification (§5.2) — `SUM (A1)` keeps the
        // kinds M1a assigned it and parses as an intersection. Excel's
        // stored `<f>` never spaces a call, so the divergence is
        // unreachable through the workbook path.
        if (!qualified) {
            if (self.kindAt(1)) |k1| {
                if (k1 == .paren_open) return self.parseCall(start);
            }
        }
        _ = self.advance();
        return self.addNode(try self.nameNode(t.text, self.spanFrom(start), .value));
    }

    fn nameNode(self: *Parser, raw: []const u8, span: Span, use: NameUse) ParseError!Node {
        const split = splitPrefix(raw);
        if (split.xlpm) {
            return self.fail(.xlpm_parameter, span.start, raw.len);
        }
        return .{ .name = .{
            .span = span,
            .raw = raw,
            .bare = split.bare,
            .prefix = split.prefix,
            .use = use,
            .builtin_xlnm = asciiStartsWithIgnoreCase(split.bare, "_xlnm."),
        } };
    }

    fn parseCall(self: *Parser, start: usize) ParseError!Index {
        try self.enter();
        defer self.leave();
        self.fn_depth += 1;
        defer self.fn_depth -= 1;
        if (self.fn_depth > self.limits.max_fn_nesting) {
            return self.failLimit(.fn_nesting, start, 0);
        }

        const name_tok = self.advance();
        const callee_span = self.spanFrom(start);
        const split = splitPrefix(name_tok.text);
        if (split.xlpm) return self.fail(.xlpm_parameter, start, name_tok.text.len);
        if (asciiEqlIgnoreCase(split.bare, "LAMBDA") or asciiEqlIgnoreCase(split.bare, "LET")) {
            return self.fail(.lambda_let_unsupported, start, name_tok.text.len);
        }
        const callee = try self.addNode(try self.nameNode(name_tok.text, callee_span, .call));

        assert(self.cur().?.kind == .paren_open);
        _ = self.advance();

        const outer_arg_list = self.in_arg_list;
        self.in_arg_list = true;
        defer self.in_arg_list = outer_arg_list;

        var args: std.ArrayListUnmanaged(Index) = .empty;
        defer args.deinit(self.allocator);

        self.skipTrivia();
        const empty = if (self.cur()) |t| t.kind == .paren_close else false;
        if (!empty) {
            while (true) {
                if (args.items.len >= self.limits.max_args) {
                    return self.failLimit(.args, start, 0);
                }
                const arg = try self.parseArgument();
                try args.append(self.allocator, arg);
                self.skipTrivia();
                const t = self.cur() orelse
                    return self.fail(.unbalanced_delimiter, start, 1);
                if (t.kind == .arg_sep) {
                    if (t.text[0] == ';') return self.failHere(.locale_separator);
                    _ = self.advance();
                    continue;
                }
                if (t.kind == .paren_close) break;
                return self.failHere(.unexpected_token);
            }
        }
        const close = self.cur() orelse return self.fail(.unbalanced_delimiter, start, 1);
        if (close.kind != .paren_close) return self.failHere(.unbalanced_delimiter);
        _ = self.advance();

        const span = self.spanFrom(start);
        // `_xlfn.SINGLE(x)` IS `@x` — one operator, two spellings. The
        // AST unifies them so downstream sees one shape; `form` and
        // `callee_raw` are what let the printer hand back the original.
        if (split.prefix.xlfn and !split.prefix.xlws and
            asciiEqlIgnoreCase(split.bare, "SINGLE") and args.items.len == 1)
        {
            return self.addNode(.{ .unary = .{
                .span = span,
                .op = .implicit_intersection,
                .child = args.items[0],
                .form = .xlfn_single,
                .callee_raw = name_tok.text,
            } });
        }

        const extra = try self.addExtra(args.items);
        return self.addNode(.{ .call = .{
            .span = span,
            .callee = callee,
            .args = extra,
        } });
    }

    /// One argument, which may be omitted: `IF(A1,,2)`.
    fn parseArgument(self: *Parser) ParseError!Index {
        self.skipTrivia();
        const off = self.offset();
        if (self.cur()) |t| {
            const omitted = switch (t.kind) {
                .paren_close => true,
                .arg_sep => t.text[0] != ';',
                else => false,
            };
            if (omitted) {
                return self.addNode(.{ .missing_arg = .{
                    .span = .{ .start = @intCast(off), .end = @intCast(off) },
                } });
            }
        }
        return self.parseComparison();
    }

    fn parseParen(self: *Parser, start: usize) ParseError!Index {
        try self.enter();
        defer self.leave();
        _ = self.advance();

        const outer_arg_list = self.in_arg_list;
        self.in_arg_list = false;
        defer self.in_arg_list = outer_arg_list;

        const child = try self.parseComparison();
        self.skipTrivia();
        const t = self.cur() orelse return self.fail(.unbalanced_delimiter, start, 1);
        if (t.kind != .paren_close) return self.failHere(.unbalanced_delimiter);
        _ = self.advance();
        return self.addNode(.{ .paren = .{
            .span = self.spanFrom(start),
            .child = child,
        } });
    }

    /// `{1,2;3,4}` — `,` separates columns, `;` separates rows. The
    /// separators are re-disambiguated here: outside braces a `;` is
    /// locale-sensitive input, inside them it is the row separator of
    /// the one array grammar Excel stores.
    fn parseArray(self: *Parser, start: usize) ParseError!Index {
        try self.enter();
        defer self.leave();
        _ = self.advance();

        var elems: std.ArrayListUnmanaged(Index) = .empty;
        defer elems.deinit(self.allocator);

        var rows: u32 = 0;
        var cols: u32 = 0;
        var row_len: u32 = 0;

        while (true) {
            const e = try self.parseArrayElement();
            try elems.append(self.allocator, e);
            row_len += 1;

            self.skipTrivia();
            const t = self.cur() orelse return self.fail(.unbalanced_delimiter, start, 1);
            if (t.kind == .arg_sep) {
                _ = self.advance();
                if (t.text[0] == ',') continue;
                // `;` — end of row.
                rows += 1;
                if (rows == 1) {
                    cols = row_len;
                } else if (row_len != cols) {
                    return self.fail(.ragged_array, start, self.bytesSince(start));
                }
                row_len = 0;
                continue;
            }
            if (t.kind == .array_close) {
                _ = self.advance();
                rows += 1;
                if (rows == 1) {
                    cols = row_len;
                } else if (row_len != cols) {
                    return self.fail(.ragged_array, start, self.bytesSince(start));
                }
                break;
            }
            // Inside braces the only things that may follow an element
            // are a separator and the closing brace. Anything else means
            // the "element" was an expression, which the array-constant
            // grammar does not admit.
            return self.failHere(.array_element_not_constant);
        }

        assert(elems.items.len == @as(usize, rows) * @as(usize, cols));
        const extra = try self.addExtra(elems.items);
        return self.addNode(.{ .array = .{
            .span = self.spanFrom(start),
            .rows = rows,
            .cols = cols,
            .elems = extra,
        } });
    }

    /// Array elements are literals only: a signed number, a string, a
    /// boolean, or an error literal. No references, no calls, no
    /// nesting — Excel's array-constant grammar, refused rather than
    /// half-supported.
    fn parseArrayElement(self: *Parser) ParseError!Index {
        self.skipTrivia();
        const start = self.offset();
        const t = self.cur() orelse return self.fail(.unexpected_end, self.input.len, 0);

        // `{}` and `{1,}` both land here with nothing to read.
        if (t.kind == .array_close) return self.fail(.empty_array, start, t.text.len);
        // `{{1}}` — the grammar has exactly one level.
        if (t.kind == .array_open) return self.fail(.array_element_not_constant, start, t.text.len);

        switch (t.kind) {
            .op_plus, .op_minus => {
                const op: UnaryOp = if (t.kind == .op_plus) .plus else .minus;
                _ = self.advance();
                self.skipTrivia();
                const num = self.cur() orelse return self.fail(.unexpected_end, self.input.len, 0);
                if (num.kind != .number) {
                    return self.fail(.array_element_not_constant, self.offset(), num.text.len);
                }
                const num_start = self.offset();
                _ = self.advance();
                const child = try self.addNode(.{ .number = .{
                    .span = self.spanFrom(num_start),
                    .text = num.text,
                } });
                return self.addNode(.{ .unary = .{
                    .span = self.spanFrom(start),
                    .op = op,
                    .child = child,
                } });
            },
            .number => {
                _ = self.advance();
                return self.addNode(.{ .number = .{
                    .span = self.spanFrom(start),
                    .text = t.text,
                } });
            },
            .string => {
                _ = self.advance();
                return self.addNode(.{ .string = .{
                    .span = self.spanFrom(start),
                    .text = t.text,
                } });
            },
            .bool_lit => {
                _ = self.advance();
                return self.addNode(.{ .boolean = .{
                    .span = self.spanFrom(start),
                    .value = t.text[0] == 'T' or t.text[0] == 't',
                } });
            },
            .error_lit => {
                _ = self.advance();
                return self.addNode(.{ .error_lit = .{
                    .span = self.spanFrom(start),
                    .text = t.text,
                    .known = tokenizer.isKnownErrorLiteral(t.text),
                } });
            },
            else => return self.fail(.array_element_not_constant, start, t.text.len),
        }
    }

    // ── references ──

    fn tryFullColSpan(self: *Parser) ParseError!?Index {
        const m = self.mark();
        const start = self.offset();
        const a = self.cur() orelse return null;
        if (a.kind != .name) return null;
        const first = parseColBound(a.text) orelse return null;
        _ = self.advance();
        self.skipTrivia();
        const colon = self.cur() orelse {
            self.reset(m);
            return null;
        };
        if (colon.kind != .op_range) {
            self.reset(m);
            return null;
        }
        _ = self.advance();
        self.skipTrivia();
        const b = self.cur() orelse {
            self.reset(m);
            return null;
        };
        if (b.kind != .name) {
            self.reset(m);
            return null;
        }
        const last = parseColBound(b.text) orelse {
            self.reset(m);
            return null;
        };
        _ = self.advance();
        return try self.addNode(.{ .ref_full_col = .{
            .span = self.spanFrom(start),
            .first = first,
            .last = last,
        } });
    }

    fn tryFullRowSpan(self: *Parser) ParseError!?Index {
        const m = self.mark();
        const start = self.offset();
        const a = self.cur() orelse return null;
        if (a.kind != .name and a.kind != .number) return null;
        const first = parseRowBound(a.text) orelse return null;
        _ = self.advance();
        self.skipTrivia();
        const colon = self.cur() orelse {
            self.reset(m);
            return null;
        };
        if (colon.kind != .op_range) {
            self.reset(m);
            return null;
        }
        _ = self.advance();
        self.skipTrivia();
        const b = self.cur() orelse {
            self.reset(m);
            return null;
        };
        if (b.kind != .name and b.kind != .number) {
            self.reset(m);
            return null;
        }
        const last = parseRowBound(b.text) orelse {
            self.reset(m);
            return null;
        };
        _ = self.advance();
        return try self.addNode(.{ .ref_full_row = .{
            .span = self.spanFrom(start),
            .first = first,
            .last = last,
        } });
    }

    // ── structured references ──

    fn parseStructured(self: *Parser, table: ?[]const u8, start: usize) ParseError!Index {
        const t = self.cur().?;
        assert(t.kind == .structured_ref);
        _ = self.advance();
        const spec = parseStructuredSpec(t.text) orelse
            return self.fail(.malformed_structured_ref, start, self.bytesSince(start));
        return self.addNode(.{ .structured = .{
            .span = self.spanFrom(start),
            .table = table,
            .items = spec.items,
            .at_shorthand = spec.at_shorthand,
            .columns = spec.columns,
        } });
    }
};

// ─── lexeme classification ───────────────────────────────────────

/// Token kinds that can begin a primary expression. `+` and `-` are
/// excluded on purpose — see `parseIntersection`.
fn startsPrimary(k: Token.Kind) bool {
    return switch (k) {
        .cell_ref,
        .name,
        .sheet_name,
        .structured_ref,
        .number,
        .string,
        .bool_lit,
        .error_lit,
        .paren_open,
        .array_open,
        .op_at,
        => true,
        else => false,
    };
}

/// `A`, `$A`, `xfd` — one end of a full-column reference.
fn parseColBound(text: []const u8) ?ColBound {
    var rest = text;
    var absolute = false;
    if (rest.len > 0 and rest[0] == '$') {
        absolute = true;
        rest = rest[1..];
    }
    if (rest.len == 0 or rest.len > coords.max_col_letters) return null;
    const col = coords.parseCol(rest, .{ .case = .insensitive }) catch return null;
    return .{ .col = col, .absolute = absolute };
}

/// `1`, `$1048576` — one end of a full-row reference. Rejects anything
/// a number literal could also be: `1.5`, `1e3`, a leading zero.
fn parseRowBound(text: []const u8) ?RowBound {
    var rest = text;
    var absolute = false;
    if (rest.len > 0 and rest[0] == '$') {
        absolute = true;
        rest = rest[1..];
    }
    if (rest.len == 0 or rest[0] == '0') return null;
    var value: u32 = 0;
    for (rest) |c| {
        if (c < '0' or c > '9') return null;
        value = std.math.mul(u32, value, 10) catch return null;
        value = std.math.add(u32, value, c - '0') catch return null;
        if (value > coords.max_row) return null;
    }
    const row = coords.Row.fromOneBased(value) catch return null;
    return .{ .row = row, .absolute = absolute };
}

const PrefixSplit = struct {
    prefix: Prefix,
    bare: []const u8,
    xlpm: bool,
};

/// Peel the layered `_xlfn.` / `_xlfn._xlws.` prefixes. Matching is
/// ASCII case-insensitive (§5.9 case-folds); `raw` keeps the original
/// spelling regardless.
fn splitPrefix(raw: []const u8) PrefixSplit {
    var rest = raw;
    var prefix: Prefix = .{};
    if (asciiStartsWithIgnoreCase(rest, "_xlpm.")) {
        return .{ .prefix = prefix, .bare = rest, .xlpm = true };
    }
    if (asciiStartsWithIgnoreCase(rest, "_xlfn.")) {
        prefix.xlfn = true;
        rest = rest["_xlfn.".len..];
        if (asciiStartsWithIgnoreCase(rest, "_xlws.")) {
            prefix.xlws = true;
            rest = rest["_xlws.".len..];
        }
    }
    return .{ .prefix = prefix, .bare = rest, .xlpm = false };
}

fn asciiLower(c: u8) u8 {
    return if (c >= 'A' and c <= 'Z') c + 32 else c;
}

fn asciiEqlIgnoreCase(a: []const u8, b: []const u8) bool {
    if (a.len != b.len) return false;
    for (a, b) |x, y| {
        if (asciiLower(x) != asciiLower(y)) return false;
    }
    return true;
}

fn asciiStartsWithIgnoreCase(s: []const u8, needle: []const u8) bool {
    if (s.len < needle.len) return false;
    return asciiEqlIgnoreCase(s[0..needle.len], needle);
}

// ─── structured-reference specifier grammar ──────────────────────

const StructuredSpec = struct {
    items: ItemSet,
    at_shorthand: bool,
    columns: ColumnSelector,
};

const SpecPart = struct { text: []const u8, bracketed: bool };

/// Parse a `[…]` specifier, brackets included. Null on anything the
/// item-specifier grammar does not admit — the caller turns that into
/// `malformed_structured_ref`.
///
/// Grammar (§5.2): `#All #Data #Headers #Totals #This Row`, `@`, a
/// single column, a `[a]:[b]` column range, and the comma-joined
/// combined forms. `'` escapes the byte after it, which is how a column
/// name carries `[`, `]`, `#`, `@`, or `'` itself.
fn parseStructuredSpec(raw: []const u8) ?StructuredSpec {
    if (raw.len < 2 or raw[0] != '[' or raw[raw.len - 1] != ']') return null;
    var inner = raw[1 .. raw.len - 1];

    var at_shorthand = false;
    if (inner.len > 0 and inner[0] == '@') {
        at_shorthand = true;
        inner = inner[1..];
    }

    var parts: [8]SpecPart = undefined;
    var seps: [8]u8 = undefined;
    var n: usize = 0;

    var i: usize = 0;
    while (i < inner.len) {
        if (n == parts.len) return null;
        if (inner[i] == '[') {
            const close = matchBracket(inner, i) orelse return null;
            parts[n] = .{ .text = inner[i + 1 .. close], .bracketed = true };
            i = close + 1;
        } else {
            const end = scanBareSpecPart(inner, i);
            if (end == i) return null;
            parts[n] = .{ .text = inner[i..end], .bracketed = false };
            i = end;
        }
        if (i < inner.len) {
            if (inner[i] != ',' and inner[i] != ':') return null;
            seps[n] = inner[i];
            i += 1;
            if (i == inner.len) return null; // trailing separator
        } else {
            seps[n] = 0;
        }
        n += 1;
    }

    var items: ItemSet = .{};
    var columns: ColumnSelector = .none;
    if (at_shorthand) items.this_row = true;

    var k: usize = 0;
    while (k < n) {
        const p = parts[k];
        if (isItemSpecifier(p.text)) {
            if (seps[k] == ':') return null; // `#Data:` is not a range
            if (!setItem(&items, p.text)) return null;
            k += 1;
            continue;
        }
        // An unescaped leading `#` means the author wrote an item
        // specifier; an unrecognised one is malformed, not a column. A
        // column genuinely named `#Foo` is written `['#Foo]`.
        if (p.text.len > 0 and p.text[0] == '#') return null;
        // A column, or a column range. Either way it is the last thing
        // the specifier may contain.
        if (std.meta.activeTag(columns) != .none) return null;
        if (seps[k] == ':') {
            if (k + 1 >= n) return null;
            const q = parts[k + 1];
            if (isItemSpecifier(q.text)) return null;
            if (seps[k + 1] != 0) return null; // nothing may follow a range
            columns = .{ .range = .{ .first = p.text, .last = q.text } };
            k += 2;
            continue;
        }
        columns = .{ .one = p.text };
        if (seps[k] != 0) return null; // nothing may follow a column
        k += 1;
    }

    if (n == 0 and !at_shorthand) return null; // `[]`
    return .{ .items = items, .at_shorthand = at_shorthand, .columns = columns };
}

fn matchBracket(s: []const u8, open: usize) ?usize {
    var depth: usize = 0;
    var i = open;
    while (i < s.len) {
        switch (s[i]) {
            '\'' => i = @min(i + 2, s.len),
            '[' => {
                depth += 1;
                i += 1;
            },
            ']' => {
                depth -= 1;
                if (depth == 0) return i;
                i += 1;
            },
            else => i += 1,
        }
    }
    return null;
}

fn scanBareSpecPart(s: []const u8, from: usize) usize {
    var i = from;
    while (i < s.len) {
        switch (s[i]) {
            '\'' => i = @min(i + 2, s.len),
            ',', ':' => return i,
            '[', ']' => return i,
            else => i += 1,
        }
    }
    return i;
}

const item_specifiers = [_][]const u8{ "#All", "#Data", "#Headers", "#Totals", "#This Row" };

fn isItemSpecifier(text: []const u8) bool {
    if (text.len == 0 or text[0] != '#') return false;
    for (item_specifiers) |spec| {
        if (asciiEqlIgnoreCase(text, spec)) return true;
    }
    return false;
}

fn setItem(items: *ItemSet, text: []const u8) bool {
    if (asciiEqlIgnoreCase(text, "#All")) {
        if (items.all) return false;
        items.all = true;
    } else if (asciiEqlIgnoreCase(text, "#Data")) {
        if (items.data) return false;
        items.data = true;
    } else if (asciiEqlIgnoreCase(text, "#Headers")) {
        if (items.headers) return false;
        items.headers = true;
    } else if (asciiEqlIgnoreCase(text, "#Totals")) {
        if (items.totals) return false;
        items.totals = true;
    } else if (asciiEqlIgnoreCase(text, "#This Row")) {
        if (items.this_row) return false;
        items.this_row = true;
    } else return false;
    return true;
}

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

/// Parse, asserting success. Caller frees.
fn expectOk(text: []const u8) !Ast {
    var parsed = try parse(testing.allocator, text, .{});
    switch (parsed) {
        .ok => |ast| return ast,
        .refused => |r| {
            std.debug.print(
                "unexpected refusal {t} (limit {?}) at {d}+{d} for `{s}`\n",
                .{ r.reason, r.limit, r.offset, r.len, text },
            );
            parsed.deinit(testing.allocator);
            return error.UnexpectedRefusal;
        },
    }
}

fn expectRefused(text: []const u8, reason: Reason) !Refusal {
    var parsed = try parse(testing.allocator, text, .{});
    switch (parsed) {
        .ok => {
            parsed.deinit(testing.allocator);
            std.debug.print("expected refusal {t} but `{s}` parsed\n", .{ reason, text });
            return error.UnexpectedSuccess;
        },
        .refused => |r| {
            try testing.expectEqual(reason, r.reason);
            return r;
        },
    }
}

/// The M2 gate: canonical text, and the parse→print→parse structural
/// equality it exists to support.
fn expectPrints(text: []const u8, expected: []const u8) !void {
    var ast = try expectOk(text);
    defer ast.deinit(testing.allocator);
    const out = try ast.print(testing.allocator);
    defer testing.allocator.free(out);
    try testing.expectEqualStrings(expected, out);

    var again = try parse(testing.allocator, out, .{});
    defer again.deinit(testing.allocator);
    switch (again) {
        .ok => |second| {
            try testing.expect(structurallyEqual(ast, second));
            // Printing is idempotent: the canonical form is a fixed
            // point, not merely a form that happens to re-parse.
            const twice = try second.print(testing.allocator);
            defer testing.allocator.free(twice);
            try testing.expectEqualStrings(out, twice);
        },
        .refused => |r| {
            std.debug.print("canonical form `{s}` refused: {t}\n", .{ out, r.reason });
            return error.CanonicalFormRefused;
        },
    }
}

/// Round-trip without pinning the canonical spelling — for corpus
/// sweeps where the point is the property, not the text.
fn expectRoundTrip(ast: Ast) !void {
    const out = try ast.print(testing.allocator);
    defer testing.allocator.free(out);
    var again = try parse(testing.allocator, out, .{});
    defer again.deinit(testing.allocator);
    switch (again) {
        .ok => |second| try testing.expect(structurallyEqual(ast, second)),
        .refused => |r| {
            std.debug.print("canonical form `{s}` refused: {t}\n", .{ out, r.reason });
            return error.CanonicalFormRefused;
        },
    }
}

fn tagOf(ast: Ast, i: Index) Tag {
    return @as(Tag, ast.nodes[i]);
}

// ─── precedence: pinned against the committed oracle manifests ───

const oracle_excel = @embedFile("oracle_hand_spec_excel");
const oracle_ieee = @embedFile("oracle_hand_spec_ieee");
const oracle_libreoffice = @embedFile("oracle_libreoffice_suite");

/// Evaluate the arithmetic fragment of the grammar. Deliberately tiny —
/// it exists so a *parse shape* can be checked against a recorded
/// binary64 result, which is the only way to pin precedence and
/// associativity without inventing an expectation. Anything it does not
/// understand (references, calls, comparisons, text) folds to null and
/// the caller skips it.
fn foldNumeric(ast: Ast, i: Index) ?f64 {
    return switch (ast.nodes[i]) {
        .number => |n| std.fmt.parseFloat(f64, n.text) catch null,
        .paren => |n| foldNumeric(ast, n.child),
        .unary => |n| switch (n.op) {
            .plus => foldNumeric(ast, n.child),
            .minus => if (foldNumeric(ast, n.child)) |v| -v else null,
            .implicit_intersection => null,
        },
        .postfix => |n| switch (n.op) {
            .percent => if (foldNumeric(ast, n.child)) |v| v / 100.0 else null,
            .spill => null,
        },
        .binary => |n| blk: {
            const a = foldNumeric(ast, n.lhs) orelse break :blk null;
            const b = foldNumeric(ast, n.rhs) orelse break :blk null;
            break :blk switch (n.op) {
                .add => a + b,
                .sub => a - b,
                .mul => a * b,
                .div => a / b,
                .pow => std.math.pow(f64, a, b),
                else => null,
            };
        },
        else => null,
    };
}

const OracleSweep = struct { parsed: usize, folded: usize };

/// Sweep one committed manifest: every formula must parse and
/// round-trip, and every foldable number cell must reproduce the
/// recorded result.
///
/// How strictly depends on the manifest's own `fidelity` field:
///
///   * `ieee`  — exact binary64 bits. These manifests record raw
///     arithmetic, so any difference is a real disagreement.
///   * `excel` — agreement to 15 significant decimal digits, which is
///     Excel's own display rule (§5.4). An excel-fidelity manifest
///     records `0.1+0.2` as `0.3`, so demanding exact bits there would
///     be testing zlsx against a rounding policy M3a1 has not landed
///     yet. The tolerance is nowhere near wide enough to blur a
///     precedence question — the cases that discriminate differ by
///     factors of 8 and by sign.
fn sweepOracleManifest(json: []const u8) !OracleSweep {
    const doc = try std.json.parseFromSlice(std.json.Value, testing.allocator, json, .{});
    defer doc.deinit();

    const fidelity = doc.value.object.get("fidelity").?.string;
    const exact = std.mem.eql(u8, fidelity, "ieee");

    var stats: OracleSweep = .{ .parsed = 0, .folded = 0 };
    const cells = doc.value.object.get("cells").?.array.items;
    for (cells) |cell| {
        const obj = cell.object;
        const formula = (obj.get("formula") orelse continue).string;

        var ast = try expectOk(formula);
        defer ast.deinit(testing.allocator);
        try expectRoundTrip(ast);
        stats.parsed += 1;

        // Only value-carrying number cells can pin a shape.
        const kind = (obj.get("kind") orelse continue).string;
        if (!std.mem.eql(u8, kind, "number")) continue;
        const bits_text = (obj.get("bits") orelse continue).string;
        const folded = foldNumeric(ast, ast.root) orelse continue;

        const expected_bits = try std.fmt.parseInt(u64, bits_text[2..], 16);
        const expected: f64 = @bitCast(expected_bits);
        const got_bits: u64 = @bitCast(folded);

        const agrees = if (exact)
            expected_bits == got_bits
        else
            expected_bits == got_bits or
                @abs(folded - expected) <= 1e-15 * @abs(expected);

        if (!agrees) {
            std.debug.print(
                "oracle mismatch for `{s}` ({s}): recorded 0x{X:0>16}, parse shape yields 0x{X:0>16}\n",
                .{ formula, fidelity, expected_bits, got_bits },
            );
            return error.OracleShapeMismatch;
        }
        stats.folded += 1;
    }
    return stats;
}

test "precedence: every oracle-manifest formula parses, round-trips, and folds to its recorded bits" {
    // The manifests are the M1b deliverable; M2 reads them, never
    // writes them. A shape that disagrees with a recorded binary64
    // result is a precedence bug, and this is what catches it.
    const excel = try sweepOracleManifest(oracle_excel);
    const ieee = try sweepOracleManifest(oracle_ieee);
    const lo = try sweepOracleManifest(oracle_libreoffice);

    // Guard against a vacuous pass: if the manifests stop carrying
    // foldable arithmetic, this test stops proving anything.
    try testing.expect(excel.parsed >= 11);
    try testing.expect(excel.folded >= 4);
    try testing.expect(ieee.folded >= 5);
    try testing.expect(lo.parsed >= 18);
}

test "precedence: `^` is left-associative — the oracle records 2^3^2 = 64" {
    var ast = try expectOk("2^3^2");
    defer ast.deinit(testing.allocator);

    // (2^3)^2 — a right-associative `^` would give 2^(3^2) = 512.
    const root = ast.nodes[ast.root].binary;
    try testing.expectEqual(BinaryOp.pow, root.op);
    try testing.expectEqual(Tag.binary, tagOf(ast, root.lhs));
    try testing.expectEqual(Tag.number, tagOf(ast, root.rhs));
    try testing.expectEqual(BinaryOp.pow, ast.nodes[root.lhs].binary.op);
    try testing.expectEqual(@as(?f64, 64.0), foldNumeric(ast, ast.root));
}

test "precedence: unary ± binds tighter than `^` — the oracle records -1^2 = 1" {
    var ast = try expectOk("-1^2");
    defer ast.deinit(testing.allocator);

    // (-1)^2. If `^` bound tighter the result would be -1.
    const root = ast.nodes[ast.root].binary;
    try testing.expectEqual(BinaryOp.pow, root.op);
    try testing.expectEqual(Tag.unary, tagOf(ast, root.lhs));
    try testing.expectEqual(UnaryOp.minus, ast.nodes[root.lhs].unary.op);
    try testing.expectEqual(@as(?f64, 1.0), foldNumeric(ast, ast.root));
}

test "precedence: `%` sits between unary ± and `^`" {
    // `-1%` is `(-1)%` — unary is tighter.
    {
        var ast = try expectOk("-1%");
        defer ast.deinit(testing.allocator);
        const root = ast.nodes[ast.root].postfix;
        try testing.expectEqual(PostfixOp.percent, root.op);
        try testing.expectEqual(Tag.unary, tagOf(ast, root.child));
        try testing.expectEqual(@as(?f64, -0.01), foldNumeric(ast, ast.root));
    }
    // `2^3%` is `2^(3%)` — `%` is tighter than `^`.
    {
        var ast = try expectOk("2^3%");
        defer ast.deinit(testing.allocator);
        const root = ast.nodes[ast.root].binary;
        try testing.expectEqual(BinaryOp.pow, root.op);
        try testing.expectEqual(Tag.postfix, tagOf(ast, root.rhs));
    }
}

test "precedence: the full arithmetic chain" {
    const Case = struct { input: []const u8, canonical: []const u8, value: ?f64 };
    const cases = [_]Case{
        // `*` `/` over `+` `-`
        .{ .input = "1+2*3", .canonical = "1+2*3", .value = 7 },
        .{ .input = "(1+2)*3", .canonical = "(1+2)*3", .value = 9 },
        // `^` over `*`
        .{ .input = "2*3^2", .canonical = "2*3^2", .value = 18 },
        // `+` `-` left-associative
        .{ .input = "1-2-3", .canonical = "1-2-3", .value = -4 },
        .{ .input = "8/4/2", .canonical = "8/4/2", .value = 1 },
        // unary chains are right-associative
        .{ .input = "--1", .canonical = "--1", .value = 1 },
        .{ .input = "-+-1", .canonical = "-+-1", .value = 1 },
    };
    for (cases) |c| {
        try expectPrints(c.input, c.canonical);
        var ast = try expectOk(c.input);
        defer ast.deinit(testing.allocator);
        try testing.expectEqual(c.value, foldNumeric(ast, ast.root));
    }
}

test "precedence: `&` is looser than arithmetic and tighter than comparisons" {
    // Spec-pinned (§5.2), not oracle-pinned: no committed manifest
    // mixes `&` with `+` or a comparison. Recorded as a fixture so the
    // chain is still a checked artifact rather than a comment.
    {
        var ast = try expectOk("1+2&\"x\"");
        defer ast.deinit(testing.allocator);
        const root = ast.nodes[ast.root].binary;
        try testing.expectEqual(BinaryOp.concat, root.op);
        try testing.expectEqual(BinaryOp.add, ast.nodes[root.lhs].binary.op);
    }
    {
        var ast = try expectOk("\"a\"&\"b\"=\"ab\"");
        defer ast.deinit(testing.allocator);
        const root = ast.nodes[ast.root].binary;
        try testing.expectEqual(BinaryOp.eq, root.op);
        try testing.expectEqual(BinaryOp.concat, ast.nodes[root.lhs].binary.op);
    }
}

test "precedence: reference operators bind tighter than every arithmetic operator" {
    // `:` > ` ` > `,` > unary ±  (§5.2)
    {
        var ast = try expectOk("SUM(A1:B2 B1:C3)");
        defer ast.deinit(testing.allocator);
        const call = ast.nodes[ast.root].call;
        const arg = ast.children(call.args)[0];
        try testing.expectEqual(BinaryOp.intersect, ast.nodes[arg].binary.op);
        // Each side of the intersection is a whole range.
        const lhs = ast.nodes[arg].binary.lhs;
        try testing.expectEqual(BinaryOp.range, ast.nodes[lhs].binary.op);
    }
    {
        // `,` is looser than ` ` but tighter than unary.
        var ast = try expectOk("(A1:A5 B1:B5,C1)");
        defer ast.deinit(testing.allocator);
        const inner = ast.nodes[ast.nodes[ast.root].paren.child].binary;
        try testing.expectEqual(BinaryOp.union_op, inner.op);
        try testing.expectEqual(BinaryOp.intersect, ast.nodes[inner.lhs].binary.op);
    }
    {
        // `-A1:B2` is `-(A1:B2)`, not `(-A1):B2`.
        var ast = try expectOk("-A1:B2");
        defer ast.deinit(testing.allocator);
        const root = ast.nodes[ast.root].unary;
        try testing.expectEqual(UnaryOp.minus, root.op);
        try testing.expectEqual(BinaryOp.range, ast.nodes[root.child].binary.op);
    }
}

test "intersection: whitespace is an operator only between operands" {
    try expectPrints("(A1:A10) (B5:F5)", "(A1:A10) (B5:F5)");
    try expectPrints("A1 + B2", "A1+B2");
    try expectPrints("  SUM( A1 , B2 )  ", "SUM(A1,B2)");
    try expectPrints("1 * 2", "1*2");
    // `A1 -1` is subtraction. Excel resolves the ambiguity that way and
    // so does the ladder: `+`/`-` are excluded from the primary-starter
    // set that triggers an intersection.
    {
        var ast = try expectOk("A1 -1");
        defer ast.deinit(testing.allocator);
        try testing.expectEqual(BinaryOp.sub, ast.nodes[ast.root].binary.op);
    }
    // Whitespace around `:` is trivia, not an intersection.
    try expectPrints("A1 : B2", "A1:B2");
}

// ─── references ──────────────────────────────────────────────────

test "references: cells carry typed coordinates and their anchors" {
    const Case = struct { input: []const u8, col: u32, row: u32, abs_col: bool, abs_row: bool };
    const cases = [_]Case{
        .{ .input = "A1", .col = 0, .row = 1, .abs_col = false, .abs_row = false },
        .{ .input = "$A$1", .col = 0, .row = 1, .abs_col = true, .abs_row = true },
        .{ .input = "$A1", .col = 0, .row = 1, .abs_col = true, .abs_row = false },
        .{ .input = "A$1", .col = 0, .row = 1, .abs_col = false, .abs_row = true },
        .{ .input = "XFD1048576", .col = 16383, .row = 1048576, .abs_col = false, .abs_row = false },
        .{ .input = "b12", .col = 1, .row = 12, .abs_col = false, .abs_row = false },
    };
    for (cases) |c| {
        var ast = try expectOk(c.input);
        defer ast.deinit(testing.allocator);
        const ref = ast.nodes[ast.root].ref_cell;
        try testing.expectEqual(c.col, ref.cell.col.zeroBased());
        try testing.expectEqual(c.row, ref.cell.row.oneBased());
        try testing.expectEqual(c.abs_col, ref.cell.anchor.col);
        try testing.expectEqual(c.abs_row, ref.cell.anchor.row);
        try expectRoundTrip(ast);
    }
}

test "references: full columns and full rows are structurally recognized" {
    {
        var ast = try expectOk("A:C");
        defer ast.deinit(testing.allocator);
        const r = ast.nodes[ast.root].ref_full_col;
        try testing.expectEqual(@as(u32, 0), r.first.col.zeroBased());
        try testing.expectEqual(@as(u32, 2), r.last.col.zeroBased());
        try testing.expect(!r.first.absolute);
    }
    {
        var ast = try expectOk("$A:$XFD");
        defer ast.deinit(testing.allocator);
        const r = ast.nodes[ast.root].ref_full_col;
        try testing.expect(r.first.absolute and r.last.absolute);
        try testing.expectEqual(@as(u32, 16383), r.last.col.zeroBased());
    }
    {
        var ast = try expectOk("1:1");
        defer ast.deinit(testing.allocator);
        const r = ast.nodes[ast.root].ref_full_row;
        try testing.expectEqual(@as(u32, 1), r.first.row.oneBased());
        try testing.expectEqual(@as(u32, 1), r.last.row.oneBased());
    }
    {
        var ast = try expectOk("$1:$1048576");
        defer ast.deinit(testing.allocator);
        const r = ast.nodes[ast.root].ref_full_row;
        try testing.expect(r.first.absolute and r.last.absolute);
        try testing.expectEqual(@as(u32, 1_048_576), r.last.row.oneBased());
    }
    // Canonical spellings, including case normalisation of the letters.
    try expectPrints("a:c", "A:C");
    try expectPrints("2:3", "2:3");
    try expectPrints("SUM(1:1)", "SUM(1:1)");
    try expectPrints("COUNTBLANK(2:3)", "COUNTBLANK(2:3)");
    try expectPrints("SUM(Sheet1!A:A)", "SUM(Sheet1!A:A)");
}

test "references: out-of-grid spans are not full-span references" {
    // `XFE` and row 1 048 577 are outside the grid, so neither end is a
    // bound. The construct stays a generic range over names and reaches
    // name resolution (§5.9), which is where `#NAME?` is provable.
    {
        var ast = try expectOk("XFE:XFE");
        defer ast.deinit(testing.allocator);
        try testing.expectEqual(Tag.binary, tagOf(ast, ast.root));
        try testing.expectEqual(BinaryOp.range, ast.nodes[ast.root].binary.op);
    }
    {
        var ast = try expectOk("1:1048577");
        defer ast.deinit(testing.allocator);
        try testing.expectEqual(Tag.binary, tagOf(ast, ast.root));
    }
}

test "references: sheet qualifiers, quoting, and 3D spans" {
    {
        var ast = try expectOk("Sheet1!A1");
        defer ast.deinit(testing.allocator);
        const q = ast.nodes[ast.root].qualified;
        try testing.expectEqualStrings("Sheet1", q.sheet.first);
        try testing.expect(q.sheet.last == null);
        try testing.expect(!q.sheet.quoted);
        try testing.expectEqual(Tag.ref_cell, tagOf(ast, q.target));
    }
    {
        var ast = try expectOk("'My Sheet'!A1");
        defer ast.deinit(testing.allocator);
        const q = ast.nodes[ast.root].qualified;
        try testing.expectEqualStrings("'My Sheet'", q.sheet.first);
        try testing.expect(q.sheet.quoted);
    }
    {
        var ast = try expectOk("Sheet1:Sheet3!A1");
        defer ast.deinit(testing.allocator);
        const q = ast.nodes[ast.root].qualified;
        try testing.expectEqualStrings("Sheet1", q.sheet.first);
        try testing.expectEqualStrings("Sheet3", q.sheet.last.?);
    }
    // Quoting is preserved verbatim, escapes included.
    try expectPrints("'It''s'!A1", "'It''s'!A1");
    try expectPrints("Q1!A1", "Q1!A1");
    try expectPrints("Sheet1!#REF!", "Sheet1!#REF!");
    try expectPrints("Sheet1!MyName", "Sheet1!MyName");
    // The range operator applies to the qualified reference.
    try expectPrints("Sheet1!A1:B2", "Sheet1!A1:B2");
}

test "references: the spill operator is a postfix on a reference" {
    {
        var ast = try expectOk("A1#");
        defer ast.deinit(testing.allocator);
        const p = ast.nodes[ast.root].postfix;
        try testing.expectEqual(PostfixOp.spill, p.op);
        try testing.expectEqual(Tag.ref_cell, tagOf(ast, p.child));
    }
    try expectPrints("SUM(A1#)", "SUM(A1#)");
    try expectPrints("Sheet1!A1#", "Sheet1!A1#");
    // A `#` with nothing to spill is not an operand.
    _ = try expectRefused("#", .unexpected_token);
}

// ─── structured references ───────────────────────────────────────

test "structured refs: the item-specifier grammar" {
    const Case = struct {
        input: []const u8,
        canonical: []const u8,
        table: ?[]const u8,
        items: ItemSet,
        at: bool,
    };
    const cases = [_]Case{
        .{ .input = "Table1[Amount]", .canonical = "Table1[Amount]", .table = "Table1", .items = .{}, .at = false },
        .{ .input = "Table1[[Amount]]", .canonical = "Table1[Amount]", .table = "Table1", .items = .{}, .at = false },
        .{ .input = "Table1[#All]", .canonical = "Table1[#All]", .table = "Table1", .items = .{ .all = true }, .at = false },
        .{ .input = "Table1[#Data]", .canonical = "Table1[#Data]", .table = "Table1", .items = .{ .data = true }, .at = false },
        .{ .input = "Table1[#Headers]", .canonical = "Table1[#Headers]", .table = "Table1", .items = .{ .headers = true }, .at = false },
        .{ .input = "Table1[#Totals]", .canonical = "Table1[#Totals]", .table = "Table1", .items = .{ .totals = true }, .at = false },
        .{ .input = "Table1[#This Row]", .canonical = "Table1[#This Row]", .table = "Table1", .items = .{ .this_row = true }, .at = false },
        .{ .input = "Table1[@]", .canonical = "Table1[@]", .table = "Table1", .items = .{ .this_row = true }, .at = true },
        .{ .input = "Table1[@Amount]", .canonical = "Table1[@Amount]", .table = "Table1", .items = .{ .this_row = true }, .at = true },
        .{ .input = "Table1[@[Col A]]", .canonical = "Table1[@[Col A]]", .table = "Table1", .items = .{ .this_row = true }, .at = true },
        .{ .input = "Table1[Col A]", .canonical = "Table1[[Col A]]", .table = "Table1", .items = .{}, .at = false },
        .{ .input = "Table1[[#Headers],[Amount]]", .canonical = "Table1[[#Headers],[Amount]]", .table = "Table1", .items = .{ .headers = true }, .at = false },
        // Item specifiers are a *set*: the canonical form orders them
        // `#All #Data #Headers #Totals #This Row` regardless of how they
        // were written.
        .{ .input = "Table1[[#Headers],[#Data],[Amount]]", .canonical = "Table1[[#Data],[#Headers],[Amount]]", .table = "Table1", .items = .{ .headers = true, .data = true }, .at = false },
        .{ .input = "Table1[[#This Row],[Amount]]", .canonical = "Table1[[#This Row],[Amount]]", .table = "Table1", .items = .{ .this_row = true }, .at = false },
        .{ .input = "Table1[[Col A]:[Col B]]", .canonical = "Table1[[Col A]:[Col B]]", .table = "Table1", .items = .{}, .at = false },
        .{ .input = "Table1[[#Data],[Col A]:[Col B]]", .canonical = "Table1[[#Data],[Col A]:[Col B]]", .table = "Table1", .items = .{ .data = true }, .at = false },
        // Bare form: the table the formula already sits in.
        .{ .input = "[Amount]", .canonical = "[Amount]", .table = null, .items = .{}, .at = false },
        .{ .input = "[@Amount]", .canonical = "[@Amount]", .table = null, .items = .{ .this_row = true }, .at = true },
        // The M1a correction: a column that looks like a cell reference.
        .{ .input = "Table1[A1]", .canonical = "Table1[A1]", .table = "Table1", .items = .{}, .at = false },
    };
    for (cases) |c| {
        var ast = try expectOk(c.input);
        defer ast.deinit(testing.allocator);
        const s = ast.nodes[ast.root].structured;
        if (c.table) |t| {
            try testing.expectEqualStrings(t, s.table.?);
        } else {
            try testing.expect(s.table == null);
        }
        try testing.expectEqual(c.items.bits(), s.items.bits());
        try testing.expectEqual(c.at, s.at_shorthand);
        try expectPrints(c.input, c.canonical);
    }
}

test "structured refs: item order is canonicalized, not preserved" {
    // `[#Headers]` and `[#Data]` name row bands; writing them in the
    // other order selects the same cells. Canonicalising the order is
    // what makes two spellings of one selection compare equal.
    var a = try expectOk("Table1[[#Headers],[#Data],[Amount]]");
    defer a.deinit(testing.allocator);
    var b = try expectOk("Table1[[#Data],[#Headers],[Amount]]");
    defer b.deinit(testing.allocator);
    try testing.expect(structurallyEqual(a, b));
}

test "structured refs: `'` escapes survive as raw text and decode on demand" {
    var ast = try expectOk("Table1[Total ']' Cost]");
    defer ast.deinit(testing.allocator);
    const s = ast.nodes[ast.root].structured;
    try testing.expectEqualStrings("Total ']' Cost", s.columns.one);

    const decoded = try decodeColumnName(testing.allocator, s.columns.one);
    defer testing.allocator.free(decoded);
    try testing.expectEqualStrings("Total ] Cost", decoded);

    try expectRoundTrip(ast);
}

test "structured refs: a column name carrying a separator keeps its brackets" {
    // `,` and `:` are the only bytes the bare form cannot carry.
    try expectPrints("Table1[[a,b]]", "Table1[[a,b]]");
    try expectPrints("Table1[[a:b]]", "Table1[[a:b]]");
}

test "structured refs: malformed specifiers refuse" {
    const cases = [_][]const u8{
        "Table1[]", // no specifier at all
        "Table1[#Nope]", // not an item specifier
        "Table1[[#Headers],[#Headers]]", // repeated item
        "Table1[[a],[b]]", // two columns, no range
        "Table1[[a]:[b],[c]]", // something after a range
        "Table1[[#Data]:[#Headers]]", // items are not range ends
        "Table1[a,]", // trailing separator
    };
    for (cases) |c| {
        _ = try expectRefused(c, .malformed_structured_ref);
    }
}

// ─── array constants ─────────────────────────────────────────────

test "array constants: `,` columns and `;` rows, rectangular" {
    {
        var ast = try expectOk("{1,2;3,4}");
        defer ast.deinit(testing.allocator);
        const a = ast.nodes[ast.root].array;
        try testing.expectEqual(@as(u32, 2), a.rows);
        try testing.expectEqual(@as(u32, 2), a.cols);
        try testing.expectEqual(@as(usize, 4), ast.children(a.elems).len);
    }
    {
        var ast = try expectOk("{1;2;3}");
        defer ast.deinit(testing.allocator);
        const a = ast.nodes[ast.root].array;
        try testing.expectEqual(@as(u32, 3), a.rows);
        try testing.expectEqual(@as(u32, 1), a.cols);
    }
    try expectPrints("{1,2;3,4}", "{1,2;3,4}");
    try expectPrints("{ 1 , 2 ; 3 , 4 }", "{1,2;3,4}");
    try expectPrints("{-1,+2}", "{-1,+2}");
    try expectPrints("{\"a\",TRUE,#N/A}", "{\"a\",TRUE,#N/A}");
    try expectPrints("{true}", "{TRUE}");
}

test "array constants: rectangularity is required" {
    _ = try expectRefused("{1,2;3}", .ragged_array);
    _ = try expectRefused("{1;2,3}", .ragged_array);
    _ = try expectRefused("{}", .empty_array);
    _ = try expectRefused("{1,}", .empty_array);
}

test "array constants: elements are literals only" {
    const cases = [_][]const u8{
        "{A1}", // reference
        "{SUM(1)}", // call
        "{{1}}", // nesting
        "{1+1}", // expression
        "{-TRUE}", // sign on a non-number
    };
    for (cases) |c| {
        _ = try expectRefused(c, .array_element_not_constant);
    }
}

// ─── leading `=` ─────────────────────────────────────────────────

test "leading `=`: exactly one is stripped, and only when allowed" {
    try expectPrints("=1+1", "1+1");
    try expectPrints("  =  SUM(A1)  ", "SUM(A1)");
    // The stripped `=` is not part of the body.
    {
        var ast = try expectOk("=A1");
        defer ast.deinit(testing.allocator);
        try testing.expectEqual(@as(u32, 1), ast.body.start);
        try testing.expectEqual(Tag.ref_cell, tagOf(ast, ast.root));
    }
    _ = try expectRefused("==1", .double_equals);
    _ = try expectRefused("= =1", .double_equals);
    _ = try expectRefused("", .empty_formula);
    _ = try expectRefused("   ", .empty_formula);
    _ = try expectRefused("=", .empty_formula);

    // The stored `<f>` form is body-only: there the `=` is an operator
    // with nothing on its left.
    var parsed = try parse(testing.allocator, "=1+1", .{ .leading_eq = .forbid });
    defer parsed.deinit(testing.allocator);
    try testing.expectEqual(Reason.unexpected_token, parsed.refused.reason);
}

// ─── prefixes and name resolution (§5.9) ─────────────────────────

test "prefixes: `_xlfn.` and `_xlfn._xlws.` peel, and the spelling survives" {
    {
        var ast = try expectOk("_xlfn.IFS(A1,1)");
        defer ast.deinit(testing.allocator);
        const callee = ast.nodes[ast.nodes[ast.root].call.callee].name;
        try testing.expectEqualStrings("_xlfn.IFS", callee.raw);
        try testing.expectEqualStrings("IFS", callee.bare);
        try testing.expect(callee.prefix.xlfn and !callee.prefix.xlws);
        try testing.expectEqual(NameUse.call, callee.use);
    }
    {
        var ast = try expectOk("_xlfn._xlws.FILTER(A1:A5,B1:B5)");
        defer ast.deinit(testing.allocator);
        const callee = ast.nodes[ast.nodes[ast.root].call.callee].name;
        try testing.expectEqualStrings("FILTER", callee.bare);
        try testing.expect(callee.prefix.xlfn and callee.prefix.xlws);
    }
    // Original spelling preserved through the canonical printer.
    try expectPrints("_xlfn.IFS(A1,1)", "_xlfn.IFS(A1,1)");
    try expectPrints("_xlfn._xlws.FILTER(A1:A5,B1:B5)", "_xlfn._xlws.FILTER(A1:A5,B1:B5)");
    // A prefix on a value-position name peels the same way.
    {
        var ast = try expectOk("_xlfn.MyName");
        defer ast.deinit(testing.allocator);
        const n = ast.nodes[ast.root].name;
        try testing.expectEqual(NameUse.value, n.use);
        try testing.expect(n.prefix.xlfn);
    }
}

test "prefixes: `_xlfn.SINGLE` and `@` are one operator with two spellings" {
    var a = try expectOk("@A1:A10");
    defer a.deinit(testing.allocator);
    var b = try expectOk("_xlfn.SINGLE(A1:A10)");
    defer b.deinit(testing.allocator);

    try testing.expectEqual(Tag.unary, tagOf(a, a.root));
    try testing.expectEqual(Tag.unary, tagOf(b, b.root));
    try testing.expectEqual(UnaryOp.implicit_intersection, a.nodes[a.root].unary.op);
    try testing.expectEqual(UnaryOp.implicit_intersection, b.nodes[b.root].unary.op);
    try testing.expectEqual(SingleForm.at_operator, a.nodes[a.root].unary.form);
    try testing.expectEqual(SingleForm.xlfn_single, b.nodes[b.root].unary.form);

    // Same operator, same operand — the trees differ only in the
    // spelling each one has to hand back.
    try testing.expect(nodesEqual(
        a,
        a.nodes[a.root].unary.child,
        b,
        b.nodes[b.root].unary.child,
    ));
    try testing.expect(!structurallyEqual(a, b));

    try expectPrints("@A1:A10", "@A1:A10");
    try expectPrints("_xlfn.SINGLE(A1:A10)", "_xlfn.SINGLE(A1:A10)");
    // `@` reaches through the whole reference, not just its first end.
    try testing.expectEqual(BinaryOp.range, a.nodes[a.nodes[a.root].unary.child].binary.op);
}

test "prefixes: `_xlpm.`, LAMBDA, and LET refuse" {
    _ = try expectRefused("_xlpm.x", .xlpm_parameter);
    _ = try expectRefused("_xlpm.x+1", .xlpm_parameter);
    _ = try expectRefused("LAMBDA(x,x+1)", .lambda_let_unsupported);
    _ = try expectRefused("_xlfn.LAMBDA(x,x+1)", .lambda_let_unsupported);
    _ = try expectRefused("LET(x,1,x)", .lambda_let_unsupported);
    _ = try expectRefused("_xlfn.LET(x,1,x)", .lambda_let_unsupported);
}

test "names: `_xlnm.` builtins are flagged, ordinary names are not" {
    {
        var ast = try expectOk("_xlnm.Print_Area");
        defer ast.deinit(testing.allocator);
        try testing.expect(ast.nodes[ast.root].name.builtin_xlnm);
    }
    {
        var ast = try expectOk("MyName");
        defer ast.deinit(testing.allocator);
        try testing.expect(!ast.nodes[ast.root].name.builtin_xlnm);
    }
}

test "names: call versus value position is decided syntactically" {
    var ast = try expectOk("SUM(MyRange)");
    defer ast.deinit(testing.allocator);
    const call = ast.nodes[ast.root].call;
    try testing.expectEqual(NameUse.call, ast.nodes[call.callee].name.use);
    const arg = ast.children(call.args)[0];
    try testing.expectEqual(NameUse.value, ast.nodes[arg].name.use);
}

test "§5.9: the resolution orders are the spec artifact M2 owes" {
    // Value position: sheet-scoped shadows workbook, then tables, then
    // the `_xlnm.` builtins, then a provable `#NAME?`.
    try testing.expectEqualSlices(ValueScope, &.{
        .sheet_scoped_name,
        .workbook_name,
        .table,
        .builtin_xlnm,
        .name_error,
    }, &value_resolution_order);

    // Call position: peel the layered prefixes first, then the
    // registry, then a plane-2 refusal — never a fabricated `#NAME?`.
    try testing.expectEqualSlices(CallStage, &.{
        .strip_layered_prefixes,
        .registry,
        .unsupported_function,
    }, &call_resolution_order);

    try testing.expect(std.mem.indexOf(u8, name_matching_policy, "case-folded") != null);
}

test "names: prefix matching is case-folded, the raw spelling is not" {
    var ast = try expectOk("_XLFN.IFS(A1,1)");
    defer ast.deinit(testing.allocator);
    const callee = ast.nodes[ast.nodes[ast.root].call.callee].name;
    try testing.expect(callee.prefix.xlfn);
    try testing.expectEqualStrings("_XLFN.IFS", callee.raw);
    try testing.expectEqualStrings("IFS", callee.bare);
}

// ─── calls ───────────────────────────────────────────────────────

test "calls: arity, omitted arguments, and nesting" {
    try expectPrints("TRUE()", "TRUE()");
    try expectPrints("IF(A1,,2)", "IF(A1,,2)");
    try expectPrints("IF(,,)", "IF(,,)");
    try expectPrints("SUM(SUM(1),2)", "SUM(SUM(1),2)");
    {
        var ast = try expectOk("IF(A1,,2)");
        defer ast.deinit(testing.allocator);
        const args = ast.children(ast.nodes[ast.root].call.args);
        try testing.expectEqual(@as(usize, 3), args.len);
        try testing.expectEqual(Tag.missing_arg, tagOf(ast, args[1]));
    }
    {
        var ast = try expectOk("TRUE()");
        defer ast.deinit(testing.allocator);
        try testing.expectEqual(@as(usize, 0), ast.children(ast.nodes[ast.root].call.args).len);
    }
}

test "calls: `,` separates arguments; a union needs its own parentheses" {
    {
        var ast = try expectOk("SUM(A1,B2)");
        defer ast.deinit(testing.allocator);
        try testing.expectEqual(@as(usize, 2), ast.children(ast.nodes[ast.root].call.args).len);
    }
    {
        var ast = try expectOk("SUM((A1,B2))");
        defer ast.deinit(testing.allocator);
        const args = ast.children(ast.nodes[ast.root].call.args);
        try testing.expectEqual(@as(usize, 1), args.len);
        const inner = ast.nodes[args[0]].paren.child;
        try testing.expectEqual(BinaryOp.union_op, ast.nodes[inner].binary.op);
    }
    // …and the arg-list context is restored on the way out.
    {
        var ast = try expectOk("SUM((A1,B2),C3)");
        defer ast.deinit(testing.allocator);
        try testing.expectEqual(@as(usize, 2), ast.children(ast.nodes[ast.root].call.args).len);
    }
}

test "calls: a name is a call only when `(` is adjacent" {
    // The tokenizer's classification precedence is normative (§5.2):
    // `LOG10` with a space before `(` is a cell reference, and the
    // parser must not disagree with the kinds M1a assigned. Excel's
    // stored `<f>` never spaces a call, so this cannot be reached
    // through the workbook path.
    var ast = try expectOk("LOG10 (A1)");
    defer ast.deinit(testing.allocator);
    try testing.expectEqual(BinaryOp.intersect, ast.nodes[ast.root].binary.op);
    try testing.expectEqual(Tag.ref_cell, tagOf(ast, ast.nodes[ast.root].binary.lhs));
}

// ─── malformed input ─────────────────────────────────────────────

test "refusals: the parser's own" {
    _ = try expectRefused("(1", .unbalanced_delimiter);
    _ = try expectRefused("SUM(1", .unbalanced_delimiter);
    _ = try expectRefused("{1,2", .unbalanced_delimiter);
    _ = try expectRefused("1+", .unexpected_end);
    _ = try expectRefused("1 2 )", .trailing_input);
    _ = try expectRefused("1)", .trailing_input);
    _ = try expectRefused("*1", .unexpected_token);
    _ = try expectRefused(",1", .unexpected_token);
}

test "refusals: `;` outside braces is locale-sensitive input" {
    const r = try expectRefused("SUM(1;2)", .locale_separator);
    try testing.expectEqual(PlaneTwo.FormulaLocaleSensitiveInput, r.planeTwo());
    _ = try expectRefused("1;2", .locale_separator);
}

test "refusals: every tokenizer refusal maps onto a §10 plane-2 error" {
    const Case = struct { input: []const u8, reason: Reason, plane: PlaneTwo };
    const cases = [_]Case{
        .{ .input = "\xFF", .reason = .invalid_utf8, .plane = .FormulaMalformedInput },
        .{ .input = "a\\b", .reason = .backslash_after_start, .plane = .FormulaMalformedInput },
        .{ .input = "R1C1", .reason = .r1c1_reference, .plane = .FormulaUnsupportedConstruct },
        .{ .input = "R[-1]C[2]", .reason = .r1c1_reference, .plane = .FormulaUnsupportedConstruct },
        .{ .input = "'[Book.xlsx]Sheet1'!A1", .reason = .external_reference, .plane = .FormulaUnsupportedConstruct },
        .{ .input = "[1]Sheet1!A1", .reason = .external_reference, .plane = .FormulaUnsupportedConstruct },
        .{ .input = "\"unterminated", .reason = .unterminated_string, .plane = .FormulaMalformedInput },
        .{ .input = "'unterminated", .reason = .unterminated_sheet_name, .plane = .FormulaMalformedInput },
        .{ .input = "Table1['", .reason = .unterminated_structured_ref, .plane = .FormulaMalformedInput },
    };
    for (cases) |c| {
        const r = try expectRefused(c.input, c.reason);
        try testing.expectEqual(c.plane, r.planeTwo());
    }

    // The long-identifier refusal needs a 256-code-point name.
    var long: std.ArrayListUnmanaged(u8) = .empty;
    defer long.deinit(testing.allocator);
    try long.appendNTimes(testing.allocator, 'a', 256);
    const r = try expectRefused(long.items, .identifier_too_long);
    try testing.expectEqual(PlaneTwo.FormulaLimitExceeded, r.planeTwo());
}

test "refusals: the tokenizer→parser reason map is total" {
    // Exhaustive over `tokenizer.Refusal.Reason` — a new tokenizer
    // refusal cannot land without a plane-2 decision here.
    inline for (@typeInfo(tokenizer.Refusal.Reason).@"enum".fields) |f| {
        const tok_reason: tokenizer.Refusal.Reason = @enumFromInt(f.value);
        const reason = reasonFromTokenizer(tok_reason);
        // Every mapped reason must itself be classified in §10.
        const plane = (Refusal{ .reason = reason, .offset = 0, .len = 0 }).planeTwo();
        var found = false;
        for (raisable_at_m2) |p| {
            if (p == plane) found = true;
        }
        try testing.expect(found);
    }
}

test "plane 2: M2 raises exactly its documented subset" {
    var seen = std.EnumSet(PlaneTwo).initEmpty();
    inline for (@typeInfo(Reason).@"enum".fields) |f| {
        const reason: Reason = @enumFromInt(f.value);
        seen.insert((Refusal{ .reason = reason, .offset = 0, .len = 0 }).planeTwo());
    }
    var expected = std.EnumSet(PlaneTwo).initEmpty();
    for (raisable_at_m2) |p| expected.insert(p);
    try testing.expect(seen.eql(expected));
}

// ─── §9 parse limits ─────────────────────────────────────────────

fn repeated(allocator: std.mem.Allocator, prefix: []const u8, unit: []const u8, n: usize, suffix: []const u8) ![]u8 {
    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(allocator);
    try out.appendSlice(allocator, prefix);
    var i: usize = 0;
    while (i < n) : (i += 1) try out.appendSlice(allocator, unit);
    try out.appendSlice(allocator, suffix);
    return out.toOwnedSlice(allocator);
}

fn nestedParens(allocator: std.mem.Allocator, n: usize) ![]u8 {
    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(allocator);
    try out.appendNTimes(allocator, '(', n);
    try out.append(allocator, '1');
    try out.appendNTimes(allocator, ')', n);
    return out.toOwnedSlice(allocator);
}

fn nestedCalls(allocator: std.mem.Allocator, n: usize) ![]u8 {
    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(allocator);
    var i: usize = 0;
    while (i < n) : (i += 1) try out.appendSlice(allocator, "SUM(");
    try out.append(allocator, '1');
    try out.appendNTimes(allocator, ')', n);
    return out.toOwnedSlice(allocator);
}

/// Parse with `limits` and report whether it survived, so a boundary
/// test reads as below / at / above.
fn parsesUnder(text: []const u8, limits: Limits) !bool {
    var parsed = try parse(testing.allocator, text, .{ .limits = limits });
    defer parsed.deinit(testing.allocator);
    return parsed == .ok;
}

fn limitRefusal(text: []const u8, limits: Limits) !LimitKind {
    var parsed = try parse(testing.allocator, text, .{ .limits = limits });
    defer parsed.deinit(testing.allocator);
    try testing.expectEqual(Reason.limit_exceeded, parsed.refused.reason);
    return parsed.refused.limit.?;
}

test "limits: max_parse_depth — below, at, above" {
    const defaults: Limits = .{};
    const n = defaults.max_parse_depth;

    const below = try nestedParens(testing.allocator, n - 1);
    defer testing.allocator.free(below);
    try testing.expect(try parsesUnder(below, defaults));

    const at = try nestedParens(testing.allocator, n);
    defer testing.allocator.free(at);
    try testing.expect(try parsesUnder(at, defaults));

    const above = try nestedParens(testing.allocator, n + 1);
    defer testing.allocator.free(above);
    try testing.expectEqual(LimitKind.parse_depth, try limitRefusal(above, defaults));
}

test "limits: max_fn_nesting — below, at, above" {
    const defaults: Limits = .{};
    const n = defaults.max_fn_nesting;

    const below = try nestedCalls(testing.allocator, n - 1);
    defer testing.allocator.free(below);
    try testing.expect(try parsesUnder(below, defaults));

    const at = try nestedCalls(testing.allocator, n);
    defer testing.allocator.free(at);
    try testing.expect(try parsesUnder(at, defaults));

    const above = try nestedCalls(testing.allocator, n + 1);
    defer testing.allocator.free(above);
    try testing.expectEqual(LimitKind.fn_nesting, try limitRefusal(above, defaults));
}

test "limits: max_args — below, at, above" {
    const defaults: Limits = .{};
    const n = defaults.max_args;

    // `SUM(1,1,…,1)` with exactly k arguments.
    const below = try repeated(testing.allocator, "SUM(1", ",1", n - 2, ")");
    defer testing.allocator.free(below);
    try testing.expect(try parsesUnder(below, defaults));

    const at = try repeated(testing.allocator, "SUM(1", ",1", n - 1, ")");
    defer testing.allocator.free(at);
    try testing.expect(try parsesUnder(at, defaults));

    const above = try repeated(testing.allocator, "SUM(1", ",1", n, ")");
    defer testing.allocator.free(above);
    try testing.expectEqual(LimitKind.args, try limitRefusal(above, defaults));
}

test "limits: max_formula_chars — below, at, above" {
    const defaults: Limits = .{};
    const n = defaults.max_formula_chars;

    // `1` + `+1` × k  ⇒  1 + 2k code points.
    const below = try repeated(testing.allocator, "1", "+1", (n - 3) / 2, "");
    defer testing.allocator.free(below);
    try testing.expect(countCodepoints(below) < n);
    try testing.expect(try parsesUnder(below, defaults));

    const at = try repeated(testing.allocator, "11", "+1", (n - 2) / 2, "");
    defer testing.allocator.free(at);
    try testing.expectEqual(n, countCodepoints(at));
    try testing.expect(try parsesUnder(at, defaults));

    const above = try repeated(testing.allocator, "111", "+1", (n - 1) / 2, "");
    defer testing.allocator.free(above);
    try testing.expect(countCodepoints(above) > n);
    try testing.expectEqual(LimitKind.formula_chars, try limitRefusal(above, defaults));
}

test "limits: max_formula_utf8_bytes — below, at, above" {
    // Dominated at defaults (see the dominance test), so the boundary is
    // exercised at a lowered ceiling. The check runs before
    // tokenization, so any bytes will do.
    const limits: Limits = .{ .max_formula_utf8_bytes = 8 };

    try testing.expect(try parsesUnder("1+1", limits));
    try testing.expect(try parsesUnder("1+1+1+1", limits)); // exactly 8
    try testing.expectEqual(
        LimitKind.formula_utf8_bytes,
        try limitRefusal("1+1+1+1+1", limits),
    );
}

test "limits: max_tokens — below, at, above" {
    // `1+1` is three tokens; `1+1+1` is five.
    const limits: Limits = .{ .max_tokens = 5 };
    try testing.expect(try parsesUnder("1+1", limits));
    try testing.expect(try parsesUnder("1+1+1", limits));
    try testing.expectEqual(LimitKind.tokens, try limitRefusal("1+1+1+1", limits));
}

test "limits: max_ast_nodes — below, at, above" {
    // `1+1` is three nodes (two literals, one operator); `1+1+1` is five.
    const limits: Limits = .{ .max_ast_nodes = 5 };
    try testing.expect(try parsesUnder("1+1", limits));
    try testing.expect(try parsesUnder("1+1+1", limits));
    try testing.expectEqual(LimitKind.ast_nodes, try limitRefusal("1+1+1+1", limits));
}

test "limits: the dominance order at defaults" {
    const d: Limits = .{};
    // A code point is at most four UTF-8 bytes, so an input that
    // exceeds the byte cap necessarily exceeds the code-point cap
    // first. Equality is reachable; excess is not.
    try testing.expectEqual(d.max_formula_chars * 4, d.max_formula_utf8_bytes);
    // Every token spans at least one code point, and every node is
    // built from at least one token, so both counts are bounded by the
    // code-point cap.
    try testing.expect(d.max_tokens > d.max_formula_chars);
    try testing.expect(d.max_ast_nodes >= d.max_tokens);
}

test "limits: a refusal names the limit it hit" {
    const above = try nestedParens(testing.allocator, 300);
    defer testing.allocator.free(above);
    var parsed = try parse(testing.allocator, above, .{});
    defer parsed.deinit(testing.allocator);
    try testing.expectEqual(PlaneTwo.FormulaLimitExceeded, parsed.refused.planeTwo());
    try testing.expectEqual(LimitKind.parse_depth, parsed.refused.limit.?);
}

// ─── the M1a fixture corpus ──────────────────────────────────────

const tokenizer_source = @embedFile("tokenizer.zig");

/// Pull every `.input = "…"` row out of the M1a fixture tables. Reading
/// them out of the tokenizer's own source is the point: a hand-copied
/// list would drift the moment M1a gains a fixture, and this milestone
/// owes a round-trip over *every* one of them.
const FixtureIterator = struct {
    src: []const u8,
    i: usize = 0,
    buf: [512]u8 = undefined,

    const marker = ".input = \"";

    fn next(self: *FixtureIterator) ?[]const u8 {
        const rel = std.mem.indexOfPos(u8, self.src, self.i, marker) orelse return null;
        var p = rel + marker.len;
        var n: usize = 0;
        while (p < self.src.len and self.src[p] != '"') {
            if (self.src[p] == '\\') {
                const decoded = decodeEscape(self.src, p) catch {
                    self.i = p + 1;
                    return self.next();
                };
                if (n + decoded.len > self.buf.len) return null;
                @memcpy(self.buf[n .. n + decoded.len], decoded.bytes[0..decoded.len]);
                n += decoded.len;
                p = decoded.next;
                continue;
            }
            if (n == self.buf.len) return null;
            self.buf[n] = self.src[p];
            n += 1;
            p += 1;
        }
        self.i = p + 1;
        return self.buf[0..n];
    }
};

const Escape = struct { bytes: [4]u8, len: usize, next: usize };

fn decodeEscape(src: []const u8, at: usize) !Escape {
    if (at + 1 >= src.len) return error.BadEscape;
    const c = src[at + 1];
    const simple: ?u8 = switch (c) {
        'n' => '\n',
        't' => '\t',
        'r' => '\r',
        '\\' => '\\',
        '"' => '"',
        '\'' => '\'',
        else => null,
    };
    if (simple) |b| return .{ .bytes = .{ b, 0, 0, 0 }, .len = 1, .next = at + 2 };
    if (c == 'x') {
        if (at + 3 >= src.len) return error.BadEscape;
        const b = try std.fmt.parseInt(u8, src[at + 2 .. at + 4], 16);
        return .{ .bytes = .{ b, 0, 0, 0 }, .len = 1, .next = at + 4 };
    }
    if (c == 'u') {
        if (at + 3 >= src.len or src[at + 2] != '{') return error.BadEscape;
        const close = std.mem.indexOfScalarPos(u8, src, at + 3, '}') orelse return error.BadEscape;
        const cp = try std.fmt.parseInt(u21, src[at + 3 .. close], 16);
        var out: Escape = .{ .bytes = undefined, .len = 0, .next = close + 1 };
        out.len = std.unicode.utf8Encode(cp, &out.bytes) catch return error.BadEscape;
        return out;
    }
    return error.BadEscape;
}

test "M1a corpus: every tokenizer fixture either refuses or round-trips" {
    var it: FixtureIterator = .{ .src = tokenizer_source };
    var count: usize = 0;
    var parsed_count: usize = 0;
    while (it.next()) |fixture| {
        count += 1;
        var parsed = try parse(testing.allocator, fixture, .{});
        defer parsed.deinit(testing.allocator);
        switch (parsed) {
            .refused => |r| {
                // A refusal must be typed and must point inside the
                // input — "it did not crash" is not a contract.
                try testing.expect(r.offset + r.len <= fixture.len);
                _ = r.planeTwo();
            },
            .ok => |ast| {
                parsed_count += 1;
                try expectRoundTrip(ast);
                try verifySpans(ast, ast.root);
            },
        }
    }
    // The extractor must actually be finding the tables.
    try testing.expect(count >= 40);
    try testing.expect(parsed_count >= 30);
}

test "M1a corpus: every compat-suite fixture parses" {
    // The compat suite is M1a's own gate — the constructs the pre-M1a
    // tokenizer already classified. Every one of them is a well-formed
    // formula, so a refusal here is an M2 regression, not a fixture
    // that was always meant to fail.
    const start = std.mem.indexOf(
        u8,
        tokenizer_source,
        "test \"compat: every previously-recognized construct keeps its kinds\"",
    ).?;
    const rest = tokenizer_source[start..];
    const end = std.mem.indexOf(u8, rest, "\n}\n").?;

    var it: FixtureIterator = .{ .src = rest[0..end] };
    var count: usize = 0;
    while (it.next()) |fixture| {
        count += 1;
        var ast = try expectOk(fixture);
        defer ast.deinit(testing.allocator);
        try expectRoundTrip(ast);
        try verifySpans(ast, ast.root);
    }
    try testing.expectEqual(@as(usize, 38), count);
}

// ─── span discipline ─────────────────────────────────────────────

/// Every node's span must sit inside its parent's and inside the
/// source. This is the structural half of "no lost bytes"; the byte
/// count asserted inside `parse` is the other half.
fn verifySpans(ast: Ast, i: Index) error{TestUnexpectedResult}!void {
    const parent = ast.nodes[i].span();
    try testing.expect(parent.start <= parent.end);
    try testing.expect(parent.end <= ast.source.len);

    switch (ast.nodes[i]) {
        .number, .string, .boolean, .error_lit, .missing_arg => {},
        .ref_cell, .ref_full_col, .ref_full_row, .name, .structured => {},
        .array => |n| for (ast.children(n.elems)) |c| try verifyChild(ast, parent, c),
        .qualified => |n| try verifyChild(ast, parent, n.target),
        .call => |n| {
            try verifyChild(ast, parent, n.callee);
            for (ast.children(n.args)) |c| try verifyChild(ast, parent, c);
        },
        .paren => |n| try verifyChild(ast, parent, n.child),
        .unary => |n| try verifyChild(ast, parent, n.child),
        .postfix => |n| try verifyChild(ast, parent, n.child),
        .binary => |n| {
            try verifyChild(ast, parent, n.lhs);
            try verifyChild(ast, parent, n.rhs);
            // Siblings are in source order and do not overlap.
            try testing.expect(ast.nodes[n.lhs].span().end <= ast.nodes[n.rhs].span().start);
        },
    }
}

fn verifyChild(ast: Ast, parent: Span, child: Index) error{TestUnexpectedResult}!void {
    const s = ast.nodes[child].span();
    try testing.expect(s.start >= parent.start);
    try testing.expect(s.end <= parent.end);
    try verifySpans(ast, child);
}

test "spans: nesting and ordering hold over a mixed corpus" {
    const inputs = [_][]const u8{
        "=SUM(Sheet1!A1:B2)+Table1[@Col]&\"x\"",
        "IF(A1>0,{1,2;3,4},-B2%)",
        "@A1:A10",
        "_xlfn._xlws.FILTER(A:A,B:B)",
        "(A1:A10) (B5:F5)",
        "'My Sheet'!$A$1:$C$9",
        "1:1",
    };
    for (inputs) |input| {
        var ast = try expectOk(input);
        defer ast.deinit(testing.allocator);
        try verifySpans(ast, ast.root);
        try expectRoundTrip(ast);
    }
}

// ─── allocation failure ──────────────────────────────────────────

fn parseAndPrint(allocator: std.mem.Allocator, text: []const u8) !void {
    var parsed = try parse(allocator, text, .{});
    defer parsed.deinit(allocator);
    switch (parsed) {
        .ok => |ast| {
            const out = try ast.print(allocator);
            defer allocator.free(out);
            var again = try parse(allocator, out, .{});
            defer again.deinit(allocator);
        },
        .refused => {},
    }
}

test "allocation failure leaks nothing" {
    const inputs = [_][]const u8{
        "=SUM(Sheet1!A1:B2)+Table1[@Col]&\"x\"",
        "IF(A1>0,{1,2;3,4},-B2%)",
        "_xlfn._xlws.FILTER(A:A,B:B)",
        "'[Book.xlsx]Sheet1'!A1",
        "SUM(SUM(SUM(SUM(1))))",
        "{1,2;3}",
        "Table1[[#Headers],[Col A]:[Col B]]",
    };
    for (inputs) |input| {
        try testing.checkAllAllocationFailures(testing.allocator, parseAndPrint, .{input});
    }
}

// ─── fuzz ────────────────────────────────────────────────────────
//
// Contract: the parser must not panic on ANY input, and a successful
// parse must account for every byte — the `bytes_consumed` assertion
// inside `parse` proves that directly, and the span walk proves the
// tree agrees. A refusal is fine; a silently dropped construct is not.
// Runs as a seed-corpus smoke test under `zig build test` and becomes
// coverage-guided under `zig build fuzz`.

fn fuzzParserTarget(_: void, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();
    var smith_buf: [4096]u8 = undefined;
    const input = smith_buf[0..smith.slice(&smith_buf)];

    var parsed = parse(std.testing.allocator, input, .{}) catch return;
    defer parsed.deinit(std.testing.allocator);

    switch (parsed) {
        .refused => |r| {
            // A refusal points at a real span and carries a §10 error.
            try std.testing.expect(r.offset <= input.len);
            try std.testing.expect(@as(usize, r.offset) + r.len <= input.len);
            try std.testing.expect((r.limit != null) == (r.reason == .limit_exceeded));
            _ = r.planeTwo();
        },
        .ok => |ast| {
            try verifySpans(ast, ast.root);
            const out = ast.print(std.testing.allocator) catch return;
            defer std.testing.allocator.free(out);

            var again = parse(std.testing.allocator, out, .{}) catch return;
            defer again.deinit(std.testing.allocator);
            switch (again) {
                .ok => |second| {
                    try std.testing.expect(structurallyEqual(ast, second));
                    const twice = second.print(std.testing.allocator) catch return;
                    defer std.testing.allocator.free(twice);
                    try std.testing.expectEqualSlices(u8, out, twice);
                },
                .refused => return error.CanonicalFormRefused,
            }
        },
    }
}

test "fuzz: the parser never panics and never loses a byte" {
    try std.testing.fuzz({}, fuzzParserTarget, .{
        .corpus = &[_][]const u8{
            "=SUM(A1:B2)",
            "=1+2*3^4%",
            "-1^2",
            "2^3^2",
            "@A1:A10",
            "_xlfn._xlws.FILTER(A:A,B:B)",
            "Table1[[#Headers],[Col A]:[Col B]]",
            "Table1[@]",
            "[@Col]",
            "{1,2;3,4}",
            "{}",
            "{1,2;3}",
            "IF(A1,,2)",
            "SUM((A1,B2))",
            "(A1:A10) (B5:F5)",
            "'My Sheet'!A1",
            "Sheet1:Sheet3!A1",
            "1:1",
            "$1:$1048576",
            "A:A",
            "A1#",
            "#",
            "==1",
            "",
            "(((((((((",
            ")))))))))",
            "SUM(SUM(SUM(1)))",
            "1;2",
            "R1C1",
            "'[Book.xlsx]S'!A1",
            "\xFF\xFE",
            "Ω+Σ",
            "a\\b",
        },
    });
}

fn fuzzOperatorSoupTarget(_: void, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();
    // A narrow alphabet concentrates the search on the precedence
    // ladder's seams — the reference operators, the two postfixes, the
    // separators that change meaning inside braces — which random bytes
    // reach only by accident.
    const alphabet = "AB12+-*/^%&<>=:,;()[]{}@#\"' ";
    var smith_buf: [256]u8 = undefined;
    const raw = smith_buf[0..smith.slice(&smith_buf)];

    var buf: [256]u8 = undefined;
    var n: usize = 0;
    for (raw) |b| {
        if (n == buf.len) break;
        buf[n] = alphabet[b % alphabet.len];
        n += 1;
    }

    var parsed = parse(std.testing.allocator, buf[0..n], .{}) catch return;
    defer parsed.deinit(std.testing.allocator);
    if (parsed == .refused) return;

    try verifySpans(parsed.ok, parsed.ok.root);
    const out = parsed.ok.print(std.testing.allocator) catch return;
    defer std.testing.allocator.free(out);
    var again = parse(std.testing.allocator, out, .{}) catch return;
    defer again.deinit(std.testing.allocator);
    if (again == .refused) return error.CanonicalFormRefused;
    try std.testing.expect(structurallyEqual(parsed.ok, again.ok));
}

test "fuzz: operator soup" {
    try std.testing.fuzz({}, fuzzOperatorSoupTarget, .{
        .corpus = &[_][]const u8{
            "A1:B2 C1:D2,E1",
            "1%^2%",
            "{1;2},{3}",
            "((A1))#",
            "@@A1",
            "A1 A2 A3",
        },
    });
}
