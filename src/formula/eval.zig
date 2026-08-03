//! The evaluator — M2's AST walked over M3a1's value model
//! (`goal_formula.md` §5.3, §5.6a, §7).
//!
//! M3a2 of the tier-D1 ladder. Everything the earlier rows built as
//! *tables* becomes executable here, and the point of this file is that
//! it adds no semantics of its own: the shape rules come from
//! `value.shapeRule`, the coercions from `value.coerceToNumber`, the
//! arithmetic from `value.addSub`/`multiply`/`divide`, the propagation
//! order from `value.propagateBinary`. Where this file has an opinion it
//! is an *evaluation* opinion — what to evaluate, in what order, and how
//! often — and each one is pinned by a fixture below.
//!
//! Laziness is a contract, not an optimisation
//! -------------------------------------------
//! §5.3a fixes which forms evaluate which arms, and the difference is
//! observable in two ways that outlive any performance argument:
//! volatile draws and dependency capture. `IF(A1,RAND(),RAND())` draws
//! **once**, and `IF(TRUE,A1,B1)` reads A1 and not B1. Meanwhile the
//! dependency GRAPH still carries a static syntactic edge to B1 — that
//! is the §5.3a "static vs runtime split", and `staticDependencies`
//! exists precisely so the two can be compared rather than conflated.
//! Every branch position of every lazy form is fixtured for value AND
//! for draws AND for captured dependencies; a form that quietly became
//! eager would still return the right number and would still fail.
//!
//! `IFS` and `SWITCH` are eager because Excel evaluates all their arms,
//! and that is observable through exactly the same two instruments. An
//! engine that "helpfully" short-circuited them would diverge on any
//! workbook containing a volatile function.
//!
//! What the evaluator never produces
//! --------------------------------
//! A non-finite number and a zero-dimension array are both
//! unrepresentable (§5.3a, §5.4 N4a). Neither is prevented by comment:
//! every arithmetic result goes through `ScalarValue.fromArithmetic`,
//! which converts overflow to `#NUM!`; every matrix comes from
//! `Matrix.init`, which refuses an empty shape; and the empty shape is
//! normalized to `#CALC!` at the producing function's boundary rather
//! than being handed on. `test "fuzz: no evaluation escapes with a
//! non-finite number or an empty matrix"` is the gate.
//!
//! The environment seam
//! --------------------
//! Nothing here reads a workbook. Cells, ranges, blank counts, per-cell
//! dialect and spill shapes all arrive through `env.EvalEnv`, and every
//! read goes through a `CallCtx`/`Evaluator` accessor that records the
//! dependency as it goes — so runtime capture cannot drift out of sync
//! with what was actually read.

const std = @import("std");
const assert = std.debug.assert;

const coords = @import("zlsx_refs");
const parser = @import("parser.zig");
const value = @import("value.zig");
const env = @import("env.zig");
const registry = @import("registry.zig");

// ─── evaluator-layer values (§5.3a) ──────────────────────────────

/// One or more rectangular areas. Multi-area sets come from the union
/// operator (`(A1:A2,C1:C2)`); §10 refuses one as a top-level *result*
/// while functions that take them keep working.
pub const Reference = struct {
    areas: []const env.RangeRef,

    pub fn single(self: Reference) ?env.RangeRef {
        return if (self.areas.len == 1) self.areas[0] else null;
    }
};

/// `ScalarValue` plus the three things only an expression can be
/// (§5.3a). `array` holds a `Matrix` of `ScalarValue`, so it is
/// non-recursive by construction rather than by rule.
pub const Value = union(enum) {
    scalar: value.ScalarValue,
    /// An omitted argument: `IF(A1,,2)`. Distinct from blank and from
    /// `""`; collapsing any pair changes `COUNTA` and every criteria
    /// match.
    missing_arg,
    array: value.Matrix,
    reference: Reference,

    pub fn err(k: value.KnownError) Value {
        return .{ .scalar = value.ScalarValue.errorOf(k) };
    }

    pub fn num(v: f64) Value {
        return .{ .scalar = value.ScalarValue.fromNumber(v) };
    }

    pub fn boolean(b: bool) Value {
        return .{ .scalar = .{ .boolean = b } };
    }
};

/// Plane-2 refusals the *evaluator* can raise. Plane-1 error values are
/// results and never appear here (§10).
pub const EvalError = error{
    OutOfMemory,
    /// Text that parses only under some locale reached a numeric
    /// context — never a guessed number, never a guessed `#VALUE!`.
    LocaleSensitiveInput,
    /// A call to a name that is not in the registry.
    UnsupportedFunction,
    /// Supported-in-v1 but not yet implemented at this ladder row. The
    /// set is enumerated and tested, so a later row deletes entries
    /// rather than discovering them.
    NotYetImplemented,
    /// A registered function called with an argument count it does not
    /// accept — a formula Excel could not have written.
    MalformedInput,
    /// A site-dependent construct evaluated without an `EvalSite`.
    AnchorRequired,
    /// A top-level multi-area result (§10, v1).
    ResultNotRepresentable,
    /// Any §9 limit.
    LimitExceeded,
} || env.Error;

/// The §10 plane-2 error a refusal raises. Exhaustive by construction —
/// a new `EvalError` fails to compile until it is mapped.
pub fn planeTwo(e: EvalError) parser.PlaneTwo {
    return switch (e) {
        error.OutOfMemory => .FormulaLimitExceeded,
        error.LocaleSensitiveInput => .FormulaLocaleSensitiveInput,
        error.UnsupportedFunction => .FormulaUnsupportedFunction,
        error.NotYetImplemented => .FormulaUnsupportedConstruct,
        error.AnchorRequired => .FormulaAnchorRequired,
        error.ResultNotRepresentable => .FormulaResultNotRepresentable,
        error.LimitExceeded => .FormulaLimitExceeded,
        // The environment's failures are invariant violations by the
        // time they reach here: the evaluator only ever builds a
        // `SheetIndex` from `resolveSheet`, and it converts shape and
        // grid failures into `#VALUE!` at the call that caused them.
        error.MalformedInput, error.UnknownSheet, error.ShapeMismatch, error.RefOutOfGrid => .FormulaMalformedInput,
    };
}

/// Constructs this row parses and refuses rather than evaluating. Named
/// as data so the later row that implements one deletes a line here and
/// watches a test fail until it does.
pub const not_yet_implemented = [_][]const u8{
    "3D sheet spans (M4b3)",
};

// ─── run instruments ─────────────────────────────────────────────

/// The single seam every volatile draw passes through.
///
/// The counter lives in the seam rather than in a test double on
/// purpose: "zero draws in the dead branch" is then a statement about
/// the evaluator, not about how carefully a fixture was wired. M3b
/// replaces the callback with `rng_v1` seeded from `RunInputs`; the
/// counter and its meaning do not change.
pub const DrawSource = struct {
    ctx: *anyopaque,
    draw_fn: *const fn (ctx: *anyopaque) f64,
    count: u64 = 0,

    pub fn draw(self: *DrawSource) f64 {
        self.count += 1;
        const v = self.draw_fn(self.ctx);
        assert(std.math.isFinite(v));
        return v;
    }

    /// A fixed source. Deterministic, so a fixture asserting a *count*
    /// is not also asserting a value it does not care about.
    pub fn constant(slot: *f64) DrawSource {
        const H = struct {
            fn draw(ctx: *anyopaque) f64 {
                return @as(*f64, @ptrCast(@alignCast(ctx))).*;
            }
        };
        return .{ .ctx = slot, .draw_fn = H.draw };
    }
};

/// What a run actually read. **Runtime** capture — laziness governs it,
/// which is exactly why it is not the graph's edge set (§5.3a). Compare
/// with `staticDependencies` to see the split.
pub const DependencyLog = struct {
    allocator: std.mem.Allocator,
    cells: std.ArrayListUnmanaged(env.CellRef) = .empty,
    areas: std.ArrayListUnmanaged(env.RangeRef) = .empty,

    pub fn init(allocator: std.mem.Allocator) DependencyLog {
        return .{ .allocator = allocator };
    }

    pub fn deinit(self: *DependencyLog) void {
        self.cells.deinit(self.allocator);
        self.areas.deinit(self.allocator);
        self.* = undefined;
    }

    pub fn noteCell(self: *DependencyLog, cell: env.CellRef) error{OutOfMemory}!void {
        for (self.cells.items) |c| if (c.eql(cell)) return;
        try self.cells.append(self.allocator, cell);
    }

    pub fn noteArea(self: *DependencyLog, area: env.RangeRef) error{OutOfMemory}!void {
        for (self.areas.items) |a| if (a.eql(area)) return;
        try self.areas.append(self.allocator, area);
    }

    pub fn hasCell(self: DependencyLog, cell: env.CellRef) bool {
        for (self.cells.items) |c| if (c.eql(cell)) return true;
        return false;
    }

    pub fn hasArea(self: DependencyLog, area: env.RangeRef) bool {
        for (self.areas.items) |a| if (a.eql(area)) return true;
        return false;
    }
};

/// Where a site-dependent construct is being evaluated. `@`
/// intersection needs it; standalone eval without one refuses rather
/// than guessing a row (§5.5).
pub const EvalSite = struct {
    row: coords.Row,
    col: coords.Col,
};

pub const Options = struct {
    /// The sheet unqualified references resolve against. Required at
    /// every layer (§5.5) — there is no "current sheet" default.
    current_sheet: env.SheetIndex,
    /// `collation_v1` with its fold injected. No default: the shipped
    /// fold lives in the `zlsx` package tree, and a module-local
    /// stand-in would be a second, quietly different comparator (§5.4b).
    collation: value.Collation,
    draws: *DrawSource,
    fidelity: value.Fidelity = .excel,
    /// Standalone eval only. Recalc asks `EvalEnv.dialectOf` per stored
    /// cell and passes what it gets (§5.3b).
    dialect: value.Dialect = .dynamic_array,
    site: ?EvalSite = null,
    limits: parser.Limits = .{},
    /// Recursion bound for the **expression-tree walk** — a stack limit.
    ///
    /// Three different depths exist and none substitutes for another.
    /// `Limits.max_parse_depth` (256) counts recursing grammar
    /// productions, so `1+1+1+…` is depth 1 however long it gets. §9's
    /// `max_eval_depth` (512) counts *dependency-closure* recursion —
    /// cell to cell, M5a's graph, not an expression at all. This one
    /// counts AST nodes on the stack, where that same flat sum is a
    /// left-leaning tree as deep as it is long.
    ///
    /// Left-associative operator chains are folded iteratively (see
    /// `Evaluator.binary`), so in practice this bound is reached only
    /// through parenthesis nesting, which `max_parse_depth` already
    /// bounds far lower. It is the backstop for a hand-built AST that
    /// never went through the parser. A §9 limit: it raises
    /// `FormulaLimitExceeded`, a defined outcome rather than a crash.
    ///
    /// It lives here rather than in `Limits` because `Limits` is M2's
    /// parse-limit struct; §9's aggregate-limit consolidation is M3b's.
    max_expr_depth: usize = 1024,
};

// ─── the evaluator ───────────────────────────────────────────────

pub const Evaluator = struct {
    /// Per-run arena (§5.3a). Every text and matrix a run produces
    /// belongs to it; nothing here frees individually.
    arena: std.mem.Allocator,
    environment: env.EvalEnv,
    opts: Options,
    deps: DependencyLog,

    ast: parser.Ast = undefined,
    /// The sheet references resolve against right now — `opts
    /// .current_sheet` except inside a qualified reference's target.
    sheet: env.SheetIndex,
    depth: usize = 0,

    pub fn init(arena: std.mem.Allocator, environment: env.EvalEnv, opts: Options) Evaluator {
        return .{
            .arena = arena,
            .environment = environment,
            .opts = opts,
            .deps = DependencyLog.init(arena),
            .sheet = opts.current_sheet,
        };
    }

    pub fn deinit(self: *Evaluator) void {
        self.deps.deinit();
        self.* = undefined;
    }

    /// Evaluate a parsed formula. The result borrows the arena.
    pub fn evaluate(self: *Evaluator, ast: parser.Ast) EvalError!Value {
        self.ast = ast;
        self.sheet = self.opts.current_sheet;
        self.depth = 0;
        const v = try self.evalNode(ast.root);
        // §10: a multi-area reference is not a representable result,
        // even though functions may take one as an argument.
        if (v == .reference and v.reference.areas.len != 1) {
            return error.ResultNotRepresentable;
        }
        // A formula's *result* is a value. A reference reaching the top
        // dereferences through the same §5.3b row as a reference reaching
        // an operator — which is why `=A1:A3` spills under `dynamic_array`
        // and intersects under `legacy` without either being special-cased.
        if (v == .reference) return operandValue(try self.operandOf(v));
        return v;
    }

    // ─── environment reads (dependency capture lives here) ───────

    pub fn readCell(self: *Evaluator, cell: env.CellRef) EvalError!value.ScalarValue {
        try self.deps.noteCell(cell);
        return self.environment.cellValue(cell);
    }

    pub fn readRange(self: *Evaluator, area: env.RangeRef) EvalError!env.RangeIterator {
        try self.deps.noteArea(area);
        return self.environment.rangeIterator(area);
    }

    pub fn readBlankCount(self: *Evaluator, area: env.RangeRef, class: env.BlankClass) EvalError!u64 {
        try self.deps.noteArea(area);
        return self.environment.logicalBlankCount(area, class);
    }

    // ─── node dispatch ───────────────────────────────────────────

    fn evalNode(self: *Evaluator, i: parser.Index) EvalError!Value {
        if (self.depth >= self.opts.max_expr_depth) return error.LimitExceeded;
        self.depth += 1;
        defer self.depth -= 1;

        return switch (self.ast.node(i)) {
            .number => |n| .{ .scalar = try self.literalNumber(n.text) },
            .string => |n| .{ .scalar = .{ .text = try self.unquoteString(n.text) } },
            .boolean => |n| Value.boolean(n.value),
            .error_lit => |n| .{ .scalar = .{ .err = errorLiteral(n.text, n.known) } },
            .missing_arg => .missing_arg,
            .array => |n| try self.arrayConstant(n.rows, n.cols, n.elems),
            .ref_cell => |n| try self.cellReference(n.cell),
            .ref_full_col => |n| try self.fullColReference(n.first, n.last),
            .ref_full_row => |n| try self.fullRowReference(n.first, n.last),
            // §5.9: at M3a2 there is no name table, so every name
            // provably resolves nowhere and the terminal stage of the
            // value-position order applies. M4b3 inserts the earlier
            // stages ahead of this; it does not change the fallthrough.
            .name, .structured => Value.err(.name),
            .qualified => |n| try self.qualified(n.sheet, n.target),
            .call => |n| try self.call(n.callee, n.args),
            .paren => |n| try self.evalNode(n.child),
            .unary => |n| try self.unary(n.op, n.child),
            .postfix => |n| try self.postfix(n.op, n.child),
            .binary => |n| try self.binary(n.op, n.lhs, n.rhs),
        };
    }

    // ─── literals ────────────────────────────────────────────────

    fn literalNumber(self: *Evaluator, text: []const u8) EvalError!value.ScalarValue {
        const n = value.parseDecimal(self.opts.fidelity, .literal, text) catch |e| switch (e) {
            error.LocaleSensitive => return error.LocaleSensitiveInput,
            // The tokenizer only emits invariant-grammar numbers, so
            // this is unreachable through `parse`; a hand-built AST gets
            // Excel's answer for unparseable text rather than a panic.
            error.NotNumeric => return value.ScalarValue.errorOf(.value),
        };
        // `1E+999` parses to infinity, which N4a turns into `#NUM!`
        // rather than letting a non-finite number exist.
        return value.ScalarValue.fromArithmetic(n);
    }

    /// A string literal carries its delimiters and its `""` escapes.
    /// Unescaping is the *literal* rule, not M4b1's carrier decode —
    /// those are different layers over different bytes.
    fn unquoteString(self: *Evaluator, text: []const u8) EvalError![]const u8 {
        assert(text.len >= 2 and text[0] == '"' and text[text.len - 1] == '"');
        const body = text[1 .. text.len - 1];
        if (std.mem.indexOfScalar(u8, body, '"') == null) return body; // borrow
        var out: std.ArrayListUnmanaged(u8) = .empty;
        try out.ensureTotalCapacity(self.arena, body.len);
        var k: usize = 0;
        while (k < body.len) : (k += 1) {
            out.appendAssumeCapacity(body[k]);
            if (body[k] == '"') k += 1; // skip the second of a `""` pair
        }
        return out.items;
    }

    fn errorLiteral(text: []const u8, known: bool) value.ErrorValue {
        if (known) {
            if (value.KnownError.fromSpelling(text)) |k| return .{ .known = k };
        }
        // Preserved, never produced (§5.3a): a spelling that arrived
        // through the tokenizer's extensible rule round-trips byte-exact.
        return .{ .rich = text };
    }

    fn arrayConstant(
        self: *Evaluator,
        rows: u32,
        cols: u32,
        elems: parser.ExtraSlice,
    ) EvalError!Value {
        var m = try self.newMatrix(rows, cols);
        const children = self.ast.children(elems);
        for (children, 0..) |child, k| {
            const v = try self.evalNode(child);
            // The parser admits only literals inside `{…}`, so every
            // element is a scalar by construction.
            assert(v == .scalar);
            m.set(@intCast(k / cols), @intCast(k % cols), v.scalar);
        }
        return .{ .array = m };
    }

    fn newMatrix(self: *Evaluator, rows: u32, cols: u32) EvalError!value.Matrix {
        return value.Matrix.init(self.arena, rows, cols) catch |e| switch (e) {
            error.OutOfMemory => error.OutOfMemory,
            // §9's cap. Not a value outcome: no error value means "this
            // array was too large to exist".
            error.TooManyCells => error.LimitExceeded,
            // Callers that can produce an empty shape normalize it to
            // `#CALC!` before reaching here (§5.3a).
            error.EmptyMatrix => unreachable,
        };
    }

    // ─── references ──────────────────────────────────────────────

    /// Every reference a run *evaluates* is a runtime dependency, read or
    /// not: capturing at the point of construction rather than at the
    /// point of dereference is what makes an eager form that holds a
    /// reference it never looks at still depend on it (§5.3a).
    fn refValue(self: *Evaluator, area: env.RangeRef) EvalError!Value {
        if (area.isSingleCell()) {
            try self.deps.noteCell(area.topLeft());
        } else {
            try self.deps.noteArea(area);
        }
        const areas = try self.arena.alloc(env.RangeRef, 1);
        areas[0] = area;
        return .{ .reference = .{ .areas = areas } };
    }

    fn cellReference(self: *Evaluator, cell: coords.Cell) EvalError!Value {
        return self.refValue(.{
            .sheet = self.sheet,
            .range = .{ .first = cell, .last = cell },
        });
    }

    fn fullColReference(self: *Evaluator, first: parser.ColBound, last: parser.ColBound) EvalError!Value {
        return self.refValue(.{ .sheet = self.sheet, .range = fullColRange(first, last) });
    }

    fn fullRowReference(self: *Evaluator, first: parser.RowBound, last: parser.RowBound) EvalError!Value {
        return self.refValue(.{ .sheet = self.sheet, .range = fullRowRange(first, last) });
    }

    fn qualified(self: *Evaluator, spec: parser.SheetSpec, target: parser.Index) EvalError!Value {
        // A quoted 3D span arrives as one token (`'Q1:Q4'!A1`), so a
        // null `last` is not proof of a single sheet. Splitting it is
        // name resolution's job (M4b3), and until that lands a span is
        // refused rather than silently read as one strangely named sheet.
        if (spec.last != null) return error.NotYetImplemented;
        const name = try self.unquoteSheetName(spec);
        if (std.mem.indexOfScalar(u8, name, ':') != null) return error.NotYetImplemented;
        const idx = (try self.environment.resolveSheet(name)) orelse
            return Value.err(.ref); // a deleted or unknown sheet
        const saved = self.sheet;
        self.sheet = idx;
        defer self.sheet = saved;
        return self.evalNode(target);
    }

    fn unquoteSheetName(self: *Evaluator, spec: parser.SheetSpec) EvalError![]const u8 {
        if (!spec.quoted) return spec.first;
        const raw = spec.first;
        assert(raw.len >= 2 and raw[0] == '\'' and raw[raw.len - 1] == '\'');
        const body = raw[1 .. raw.len - 1];
        if (std.mem.indexOfScalar(u8, body, '\'') == null) return body;
        var out: std.ArrayListUnmanaged(u8) = .empty;
        try out.ensureTotalCapacity(self.arena, body.len);
        var k: usize = 0;
        while (k < body.len) : (k += 1) {
            out.appendAssumeCapacity(body[k]);
            if (body[k] == '\'') k += 1;
        }
        return out.items;
    }

    /// Evaluate a node that must denote a reference. Used by the three
    /// reference operators, which combine areas rather than values.
    fn evalAsReference(self: *Evaluator, i: parser.Index) EvalError!?Reference {
        const v = try self.evalNode(i);
        return switch (v) {
            .reference => |r| r,
            else => null,
        };
    }

    // ─── operators ───────────────────────────────────────────────

    fn unary(self: *Evaluator, op: parser.UnaryOp, child: parser.Index) EvalError!Value {
        const v = try self.evalNode(child);
        return switch (op) {
            // Excel's unary plus is a genuine no-op: `=+"abc"` is
            // `"abc"`, not `#VALUE!`.
            .plus => v,
            .minus => try self.mapScalars(v, negate),
            .implicit_intersection => try self.implicitIntersection(v),
        };
    }

    fn postfix(self: *Evaluator, op: parser.PostfixOp, child: parser.Index) EvalError!Value {
        const v = try self.evalNode(child);
        return switch (op) {
            .percent => try self.mapScalars(v, percent),
            .spill => try self.spill(v),
        };
    }

    fn spill(self: *Evaluator, v: Value) EvalError!Value {
        const r = switch (v) {
            .reference => |r| r,
            else => return Value.err(.ref),
        };
        const area = r.single() orelse return Value.err(.ref);
        const anchor = area.topLeft();
        try self.deps.noteCell(anchor);
        const shape = (try self.environment.spillShape(anchor)) orelse
            return Value.err(.ref); // `A1#` against a non-anchor
        const last_row = coords.Row.fromOneBased(anchor.row.oneBased() + shape.rows - 1) catch
            return Value.err(.ref);
        const last_col = coords.Col.fromZeroBased(anchor.col.zeroBased() + shape.cols - 1) catch
            return Value.err(.ref);
        return self.refValue(.{ .sheet = anchor.sheet, .range = .{
            .first = .{ .col = anchor.col, .row = anchor.row },
            .last = .{ .col = last_col, .row = last_row },
        } });
    }

    /// `@expr` and legacy implicit intersection are the same operator;
    /// §5.3b gives it three rows and this is all three.
    fn implicitIntersection(self: *Evaluator, v: Value) EvalError!Value {
        switch (v) {
            .reference => |r| {
                const area = r.single() orelse return Value.err(.value);
                if (area.isSingleCell()) {
                    assert(value.shapeRule(.at_single_cell_reference, self.opts.dialect) == .reference_unchanged);
                    // `=@A1` is A1 regardless of the evaluation site:
                    // the single-item exception precedes intersection.
                    return v;
                }
                assert(value.shapeRule(.at_multi_cell_reference, self.opts.dialect) == .row_col_intersection);
                return self.rowColIntersection(area);
            },
            .array => |m| {
                assert(value.shapeRule(.at_array, self.opts.dialect) == .top_left_reduction);
                return .{ .scalar = m.topLeft() };
            },
            else => return v,
        }
    }

    /// §5.3b's intersection rule, shared by `@` and by legacy
    /// dereference. One implementation because the table says the two
    /// spellings mean the same thing.
    fn rowColIntersection(self: *Evaluator, area: env.RangeRef) EvalError!Value {
        const site = self.opts.site orelse return error.AnchorRequired;
        const r = area.range;
        const spans_row = site.row.oneBased() >= r.first.row.oneBased() and
            site.row.oneBased() <= r.last.row.oneBased();
        const spans_col = site.col.zeroBased() >= r.first.col.zeroBased() and
            site.col.zeroBased() <= r.last.col.zeroBased();

        const cell: env.CellRef = if (r.colCount() == 1 and spans_row)
            .{ .sheet = area.sheet, .row = site.row, .col = r.first.col }
        else if (r.rowCount() == 1 and spans_col)
            .{ .sheet = area.sheet, .row = r.first.row, .col = site.col }
        else if (spans_row and spans_col)
            .{ .sheet = area.sheet, .row = site.row, .col = site.col }
        else
            return Value.err(.value);

        return .{ .scalar = try self.readCell(cell) };
    }

    /// Evaluate a binary operator, descending its left spine
    /// **iteratively**.
    ///
    /// Every value operator in the grammar is left-associative, so
    /// `1+1+1+…` parses as a left-leaning tree exactly as deep as it is
    /// long. Recursing into it would put a 5 000-term sum — a formula
    /// well inside Excel's 8 192-character limit — into the call stack,
    /// where it does not fit. Folding the spine on the heap instead
    /// leaves recursion depth bounded by *parenthesis* nesting, which
    /// §9 already bounds at 256.
    ///
    /// The order is unchanged: §5.3c says operands evaluate left to
    /// right and the first error wins, and folding a left spine from its
    /// deepest step outward visits operands in exactly that order.
    fn binary(self: *Evaluator, op: parser.BinaryOp, lhs: parser.Index, rhs: parser.Index) EvalError!Value {
        if (isReferenceOp(op)) return self.referenceOperator(op, lhs, rhs);

        var spine: std.ArrayListUnmanaged(SpineStep) = .empty;
        try spine.append(self.arena, .{ .op = op, .rhs = rhs });
        var head = lhs;
        while (self.ast.node(head) == .binary) {
            const b = self.ast.node(head).binary;
            // A reference operator is a different evaluation entirely
            // (it combines areas, not values), so the spine stops there.
            if (isReferenceOp(b.op)) break;
            if (spine.items.len >= self.opts.limits.max_ast_nodes) return error.LimitExceeded;
            try spine.append(self.arena, .{ .op = b.op, .rhs = b.rhs });
            head = b.lhs;
        }

        var acc = try self.evalNode(head);
        var k = spine.items.len;
        while (k > 0) {
            k -= 1;
            const step = spine.items[k];
            const r = try self.evalNode(step.rhs);
            acc = try self.applyBinary(step.op, acc, r);
        }
        return acc;
    }

    fn referenceOperator(
        self: *Evaluator,
        op: parser.BinaryOp,
        lhs: parser.Index,
        rhs: parser.Index,
    ) EvalError!Value {
        const l = (try self.evalAsReference(lhs)) orelse return Value.err(.value);
        const r = (try self.evalAsReference(rhs)) orelse return Value.err(.value);
        switch (op) {
            .range => {
                const a = l.single() orelse return Value.err(.value);
                const b = r.single() orelse return Value.err(.value);
                if (a.sheet != b.sheet) return Value.err(.value);
                return self.refValue(.{ .sheet = a.sheet, .range = boundingBox(a.range, b.range) });
            },
            .intersect => {
                const a = l.single() orelse return Value.err(.value);
                const b = r.single() orelse return Value.err(.value);
                if (a.sheet != b.sheet) return Value.err(.null_err);
                const hit = intersectRanges(a.range, b.range) orelse
                    return Value.err(.null_err); // Excel's answer for a null intersection
                return self.refValue(.{ .sheet = a.sheet, .range = hit });
            },
            .union_op => {
                const areas = try self.arena.alloc(env.RangeRef, l.areas.len + r.areas.len);
                @memcpy(areas[0..l.areas.len], l.areas);
                @memcpy(areas[l.areas.len..], r.areas);
                return .{ .reference = .{ .areas = areas } };
            },
            else => unreachable,
        }
    }

    // ─── shape (§5.3b) ───────────────────────────────────────────

    /// A value seen through the shape table: everything an elementwise
    /// operator needs, and nothing it does not.
    const Operand = struct {
        shape: value.Shape,
        src: union(enum) { scalar: value.ScalarValue, matrix: value.Matrix },

        /// `null` where the operand has no element at this position —
        /// the incompatible half of a broadcast (§5.3b).
        fn at(self: Operand, r: u32, c: u32) ?value.ScalarValue {
            const rr = if (self.shape.rows == 1) 0 else r;
            const cc = if (self.shape.cols == 1) 0 else c;
            if (rr >= self.shape.rows or cc >= self.shape.cols) return null;
            return switch (self.src) {
                .scalar => |s| s,
                .matrix => |m| m.at(rr, cc),
            };
        }
    };

    fn scalarOperand(s: value.ScalarValue) Operand {
        return .{ .shape = .{ .rows = 1, .cols = 1 }, .src = .{ .scalar = s } };
    }

    fn operandValue(op: Operand) Value {
        return switch (op.src) {
            .scalar => |s| .{ .scalar = s },
            .matrix => |m| .{ .array = m },
        };
    }

    /// Bring a value into a value context, applying the dialect-indexed
    /// row of §5.3b that governs it.
    fn operandOf(self: *Evaluator, v: Value) EvalError!Operand {
        switch (v) {
            .scalar => |s| return scalarOperand(s),
            // A blank operand: `IF(A1,,2)+1` is `0+1`. Not `""`.
            .missing_arg => return scalarOperand(.blank),
            .array => |m| return .{ .shape = m.shape(), .src = .{ .matrix = m } },
            .reference => |r| {
                const area = r.single() orelse return scalarOperand(value.ScalarValue.errorOf(.value));
                return switch (value.shapeRule(.reference_in_value, self.opts.dialect)) {
                    .dereference => self.materialize(area),
                    .dereference_with_intersection => blk: {
                        if (area.isSingleCell()) break :blk self.materialize(area);
                        const hit = try self.rowColIntersection(area);
                        break :blk scalarOperand(hit.scalar);
                    },
                    else => unreachable,
                };
            },
        }
    }

    /// Dereference an area into the operand it denotes. A single cell
    /// stays a scalar; anything larger becomes a dense matrix filled
    /// from one sparse pass.
    fn materialize(self: *Evaluator, area: env.RangeRef) EvalError!Operand {
        if (area.isSingleCell()) {
            return scalarOperand(try self.readCell(area.topLeft()));
        }
        var m = try self.newMatrix(area.range.rowCount(), area.range.colCount());
        var it = try self.readRange(area);
        while (try it.next()) |e| {
            const dr = e.row.oneBased() - area.range.first.row.oneBased();
            const dc = e.col.zeroBased() - area.range.first.col.zeroBased();
            m.set(dr, dc, e.value);
        }
        return .{ .shape = m.shape(), .src = .{ .matrix = m } };
    }

    /// The §5.3b broadcast: dims of size one stretch, the result is
    /// (max rows × max cols), and a position no operand can supply is
    /// filled elementwise with `#N/A` — Excel's answer, not a refusal.
    fn broadcast(self: *Evaluator, a: Operand, b: Operand, op: parser.BinaryOp) EvalError!Value {
        const compatible = value.broadcastShape(a.shape, b.shape);
        const shape = compatible orelse value.Shape{
            .rows = @max(a.shape.rows, b.shape.rows),
            .cols = @max(a.shape.cols, b.shape.cols),
        };
        if (shape.isScalar()) {
            return .{ .scalar = try self.applyBinaryScalar(op, a.at(0, 0).?, b.at(0, 0).?) };
        }
        var m = try self.newMatrix(shape.rows, shape.cols);
        var r: u32 = 0;
        while (r < shape.rows) : (r += 1) {
            var c: u32 = 0;
            while (c < shape.cols) : (c += 1) {
                const av = a.at(r, c) orelse {
                    m.set(r, c, value.incompatibleBroadcastFill());
                    continue;
                };
                const bv = b.at(r, c) orelse {
                    m.set(r, c, value.incompatibleBroadcastFill());
                    continue;
                };
                // `per_element`: an error stays in its cell rather than
                // taking the whole array with it (§5.3c).
                m.set(r, c, try self.applyBinaryScalar(op, av, bv));
            }
        }
        return .{ .array = m };
    }

    fn applyBinary(self: *Evaluator, op: parser.BinaryOp, l: Value, r: Value) EvalError!Value {
        const a = try self.operandOf(l);
        const b = try self.operandOf(r);
        return self.broadcast(a, b, op);
    }

    fn applyBinaryScalar(
        self: *Evaluator,
        op: parser.BinaryOp,
        a: value.ScalarValue,
        b: value.ScalarValue,
    ) EvalError!value.ScalarValue {
        // §5.3c: left to right, first error wins — for every operator,
        // before any coercion is attempted.
        if (value.propagateBinary(a, b)) |e| return .{ .err = e };

        const rules = value.FpRules.of(self.opts.fidelity);
        switch (op) {
            .add, .sub, .mul, .div, .pow => {
                const x = switch (try self.toNumber(a)) {
                    .n => |n| n,
                    .err => |e| return e,
                };
                const y = switch (try self.toNumber(b)) {
                    .n => |n| n,
                    .err => |e| return e,
                };
                return switch (op) {
                    .add => value.addSub(rules, x, y, .add),
                    .sub => value.addSub(rules, x, y, .sub),
                    .mul => value.multiply(rules, x, y),
                    .div => value.divide(rules, x, y),
                    .pow => value.ScalarValue.fromArithmetic(std.math.pow(f64, x, y)),
                    else => unreachable,
                };
            },
            .concat => {
                const x = try self.toText(a);
                const y = try self.toText(b);
                const out = try self.arena.alloc(u8, x.len + y.len);
                @memcpy(out[0..x.len], x);
                @memcpy(out[x.len..], y);
                return .{ .text = out };
            },
            .eq, .ne, .lt, .gt, .le, .ge => {
                const ord = try self.compare(a, b);
                return .{ .boolean = switch (op) {
                    .eq => ord == .eq,
                    .ne => ord != .eq,
                    .lt => ord == .lt,
                    .gt => ord == .gt,
                    .le => ord != .gt,
                    .ge => ord != .lt,
                    else => unreachable,
                } };
            },
            .range, .intersect, .union_op => unreachable,
        }
    }

    const Num = union(enum) { n: f64, err: value.ScalarValue };

    /// The numeric column of the §5.3b coercion matrix, behaviourally.
    /// `.text_coercion` is the ingress, which is what makes `" 1 "+1`
    /// work while a stored `<v>` of `" 1 "` does not (§5.4 decision 5).
    fn toNumber(self: *Evaluator, s: value.ScalarValue) EvalError!Num {
        return switch (value.coerceToNumber(s, self.opts.fidelity, .text_coercion)) {
            .number => |n| .{ .n = n },
            .value => |v| .{ .err = v },
            .locale_refusal => error.LocaleSensitiveInput,
        };
    }

    fn toText(self: *Evaluator, s: value.ScalarValue) EvalError![]const u8 {
        return switch (s) {
            .text => |t| t,
            .blank => "",
            .boolean => |b| if (b) "TRUE" else "FALSE",
            .number => |n| blk: {
                const buf = try self.arena.alloc(u8, value.format_buf_len);
                break :blk value.formatNumber(buf, n);
            },
            .err => unreachable, // propagated before any coercion
        };
    }

    /// Ordinary comparison (§5.3b): blank adopts the other operand's
    /// type, cross-type pairs order number < text < logical and are
    /// never equal, text compares case-INsensitively under
    /// `collation_v1`, and numbers compare through N2.
    pub fn compare(self: *Evaluator, a: value.ScalarValue, b: value.ScalarValue) EvalError!std.math.Order {
        // Ordering is defined over values, not over errors: propagation
        // decides an error's fate before a comparison ever sees it.
        assert(a != .err and b != .err);
        const x = adoptBlank(a, b);
        const y = adoptBlank(b, a);

        const rx = value.crossTypeRank(x).?;
        const ry = value.crossTypeRank(y).?;
        if (rx != ry) return std.math.order(@intFromEnum(rx), @intFromEnum(ry));

        return switch (x) {
            .number => |n| self.compareNumbers(n, y.number),
            .boolean => |v| std.math.order(@intFromBool(v), @intFromBool(y.boolean)),
            .text => |t| self.opts.collation.compare(self.arena, t, y.text) catch |e| {
                if (e == error.OutOfMemory) return error.OutOfMemory;
                // The injected fold is `collation_v1`'s; nothing else it
                // could report is a value outcome.
                return error.MalformedInput;
            },
            .blank, .err => unreachable,
        };
    }

    /// A comparison is a subtraction against zero, which puts it inside
    /// N2's additive scope. Both committed manifests decide this:
    /// `(0.1+0.2)=0.3` is TRUE under `excel` and FALSE under `ieee`.
    fn compareNumbers(self: *Evaluator, a: f64, b: f64) std.math.Order {
        const rules = value.FpRules.of(self.opts.fidelity);
        const d = value.applyZeroSnap(rules, a - b, a, b);
        return std.math.order(d, 0);
    }

    fn adoptBlank(v: value.ScalarValue, other: value.ScalarValue) value.ScalarValue {
        if (v != .blank) return v;
        return switch (other) {
            .text => .{ .text = "" },
            .boolean => .{ .boolean = false },
            // Blank vs blank compares as two zeros, which makes
            // `=A1=B1` TRUE for two empty cells — Excel's answer.
            .number, .blank, .err => value.ScalarValue.fromNumber(0),
        };
    }

    /// Apply a scalar transform elementwise, keeping the shape. Used by
    /// the unary operators, which broadcast over an array exactly as a
    /// binary operator does.
    fn mapScalars(
        self: *Evaluator,
        v: Value,
        comptime f: fn (*Evaluator, value.ScalarValue) EvalError!value.ScalarValue,
    ) EvalError!Value {
        const op = try self.operandOf(v);
        if (op.shape.isScalar()) return .{ .scalar = try f(self, op.at(0, 0).?) };
        var m = try self.newMatrix(op.shape.rows, op.shape.cols);
        var r: u32 = 0;
        while (r < op.shape.rows) : (r += 1) {
            var c: u32 = 0;
            while (c < op.shape.cols) : (c += 1) {
                m.set(r, c, try f(self, op.at(r, c).?));
            }
        }
        return .{ .array = m };
    }

    /// Negation, not `0 - x`: the difference is `-0`, which §5.4 N3
    /// treats as a distinct value until publication normalizes it.
    fn negate(self: *Evaluator, s: value.ScalarValue) EvalError!value.ScalarValue {
        if (s == .err) return s;
        return switch (try self.toNumber(s)) {
            .n => |n| value.ScalarValue.fromArithmetic(-n),
            .err => |e| e,
        };
    }

    fn percent(self: *Evaluator, s: value.ScalarValue) EvalError!value.ScalarValue {
        if (s == .err) return s;
        const rules = value.FpRules.of(self.opts.fidelity);
        return switch (try self.toNumber(s)) {
            .n => |n| value.divide(rules, n, 100),
            .err => |e| e,
        };
    }

    // ─── calls (§5.3a per-form contracts, §7 registry) ───────────

    fn call(self: *Evaluator, callee: parser.Index, args: parser.ExtraSlice) EvalError!Value {
        const name_node = self.ast.node(callee);
        if (name_node != .name) return error.NotYetImplemented;
        const f = registry.lookup(name_node.name.bare) orelse return error.UnsupportedFunction;

        const arg_nodes = self.ast.children(args);
        if (arg_nodes.len < f.arity.min) return error.MalformedInput;
        if (f.arity.max) |max| {
            if (arg_nodes.len > max) return error.MalformedInput;
        }
        if (arg_nodes.len > self.opts.limits.max_args) return error.LimitExceeded;

        return switch (f.form) {
            .plain => self.callPlain(f, arg_nodes),
            .if_form => self.formIf(arg_nodes),
            .choose_form => self.formChoose(arg_nodes),
            .iferror_form => self.formIfError(arg_nodes, .any_error),
            .ifna_form => self.formIfError(arg_nodes, .na_only),
        };
    }

    /// Eager dispatch: every argument evaluates, in declaration order,
    /// before the implementation sees any of them — and each one is
    /// brought to its slot's coercion class on the way, so an
    /// implementation never re-derives what the table already said.
    ///
    /// An array reaching a scalar slot is the one place a *dialect*
    /// changes the answer: §5.3b reduces it to its top-left under
    /// `legacy` and iterates under `dynamic_array`, so `SQRT({4,9})` is
    /// `2` in one and `{2,3}` in the other.
    fn callPlain(self: *Evaluator, f: *const registry.Function, arg_nodes: []const parser.Index) EvalError!Value {
        const n = arg_nodes.len;
        const args = try self.arena.alloc(Value, n);
        const ops = try self.arena.alloc(?Operand, n);
        var lift: ?value.Shape = null;

        for (arg_nodes, 0..) |node, k| {
            const raw = try self.evalNode(node);
            const class = f.coercion.at(k);
            ops[k] = null;
            if (!class.isScalarClass()) {
                args[k] = if (class == .reference and raw != .reference)
                    Value.err(.value)
                else
                    raw;
                continue;
            }
            const op = try self.operandOf(raw);
            ops[k] = op;
            if (op.shape.isScalar()) continue;
            switch (value.shapeRule(.array_where_scalar, self.opts.dialect)) {
                // `Operand.at(0, 0)` already *is* the top-left, so the
                // scalar path below needs no separate branch.
                .top_left_reduction => {},
                .spill_or_iterate => {
                    // A mixed signature — a range slot beside a scalar
                    // one — is M7a's decision table, not this row's.
                    if (!f.liftable()) return error.NotYetImplemented;
                    const cur = lift orelse value.Shape{ .rows = 1, .cols = 1 };
                    lift = .{
                        .rows = @max(cur.rows, op.shape.rows),
                        .cols = @max(cur.cols, op.shape.cols),
                    };
                },
                else => unreachable,
            }
        }

        if (lift) |shape| {
            var m = try self.newMatrix(shape.rows, shape.cols);
            var r: u32 = 0;
            while (r < shape.rows) : (r += 1) {
                var c: u32 = 0;
                while (c < shape.cols) : (c += 1) {
                    for (ops, 0..) |maybe_op, k| {
                        const op = maybe_op orelse continue;
                        const el = op.at(r, c) orelse value.incompatibleBroadcastFill();
                        args[k] = .{ .scalar = try self.coerceSlot(f.coercion.at(k), el) };
                    }
                    const one = try self.propagateAndInvoke(f, args);
                    // A liftable function's slots are all scalar classes,
                    // so a per-element call cannot produce an array.
                    assert(one == .scalar);
                    m.set(r, c, one.scalar);
                }
            }
            return .{ .array = m };
        }

        for (ops, 0..) |maybe_op, k| {
            const op = maybe_op orelse continue;
            args[k] = .{ .scalar = try self.coerceSlot(f.coercion.at(k), op.at(0, 0).?) };
        }
        return self.propagateAndInvoke(f, args);
    }

    /// The behavioural half of a coercion class. An error passes through
    /// untouched: whether it propagates is the propagation class's
    /// decision, taken next, and coercing it first would lose it.
    fn coerceSlot(self: *Evaluator, class: registry.CoercionClass, s: value.ScalarValue) EvalError!value.ScalarValue {
        if (s == .err) return s;
        return switch (class) {
            // `fromArithmetic`, not `fromNumber`: `"1E+999"` parses to
            // infinity, and N4a's answer for that is `#NUM!`.
            .number => switch (try self.toNumber(s)) {
                .n => |x| value.ScalarValue.fromArithmetic(x),
                .err => |e| e,
            },
            .logical => switch (try self.toLogical(s)) {
                .b => |b| .{ .boolean = b },
                .err => |e| e,
            },
            .text => .{ .text = try self.toText(s) },
            .value_any, .aggregate, .reference, .lazy_any => s,
        };
    }

    fn propagateAndInvoke(self: *Evaluator, f: *const registry.Function, args: []const Value) EvalError!Value {
        // §5.3c: first error wins **unless the registry says otherwise**,
        // provenance-aware per function. `COUNT` and `COUNTA` differ here
        // and they are in the same family.
        switch (f.propagation) {
            .propagate => {
                for (args) |a| {
                    if (a == .scalar and a.scalar == .err) return a;
                }
            },
            .observe, .per_element, .per_function_provenance => {},
        }
        return self.invoke(f, args);
    }

    /// The one place a function result becomes a value. An empty shape
    /// is normalized to `#CALC!` **here**, at the producing function's
    /// boundary (§5.3a), rather than travelling as a matrix nobody can
    /// represent.
    pub fn invoke(self: *Evaluator, f: *const registry.Function, args: []const Value) EvalError!Value {
        const impl = f.impl orelse unreachable; // `.plain` implies an impl
        const v = impl(.{ .ev = self }, args) catch |e| switch (e) {
            error.EmptyMatrix => return .{ .scalar = value.emptyMatrixResult() },
            error.TooManyCells => return error.LimitExceeded,
            else => |other| return other,
        };
        if (v == .array) assert(v.array.rows > 0 and v.array.cols > 0);
        return v;
    }

    // ─── the lazy forms (§5.3a) ──────────────────────────────────

    /// `IF(cond, then, else?)`.
    ///
    /// A scalar condition takes exactly one arm — no runtime draw and no
    /// dependency from the other. An **array** condition switches the
    /// whole form to per-element masking: both arms evaluate, all three
    /// broadcast to the mask's shape, and an error in one element stays
    /// in that element.
    fn formIf(self: *Evaluator, arg_nodes: []const parser.Index) EvalError!Value {
        const cond = try self.evalNode(arg_nodes[0]);
        const mask = try self.operandOf(cond);

        if (mask.shape.isScalar()) {
            const taken = switch (try self.toLogical(mask.at(0, 0).?)) {
                .b => |b| b,
                .err => |e| return .{ .scalar = e },
            };
            if (taken) return self.evalNode(arg_nodes[1]);
            // `IF(FALSE,1)` is FALSE, not blank — Excel's answer for the
            // omitted third argument.
            if (arg_nodes.len < 3) return Value.boolean(false);
            return self.evalNode(arg_nodes[2]);
        }

        const then_op = try self.operandOf(try self.evalNode(arg_nodes[1]));
        const else_op = if (arg_nodes.len >= 3)
            try self.operandOf(try self.evalNode(arg_nodes[2]))
        else
            scalarOperand(.{ .boolean = false });

        const shape = maskShape(mask.shape, then_op.shape, else_op.shape);
        var m = try self.newMatrix(shape.rows, shape.cols);
        var r: u32 = 0;
        while (r < shape.rows) : (r += 1) {
            var c: u32 = 0;
            while (c < shape.cols) : (c += 1) {
                const cv = mask.at(r, c) orelse {
                    m.set(r, c, value.incompatibleBroadcastFill());
                    continue;
                };
                const picked = switch (try self.toLogical(cv)) {
                    .b => |b| if (b) then_op.at(r, c) else else_op.at(r, c),
                    .err => |e| {
                        m.set(r, c, e);
                        continue;
                    },
                };
                m.set(r, c, picked orelse value.incompatibleBroadcastFill());
            }
        }
        return .{ .array = m };
    }

    /// `CHOOSE(index, v1, …)`. Lazy for a scalar index, per-element
    /// masking for an array one — the same split as `IF`, because §5.3a
    /// gives them the same contract.
    fn formChoose(self: *Evaluator, arg_nodes: []const parser.Index) EvalError!Value {
        const idx = try self.evalNode(arg_nodes[0]);
        const sel = try self.operandOf(idx);
        const arms = arg_nodes[1..];

        if (sel.shape.isScalar()) {
            const pick = switch (try self.chooseIndex(sel.at(0, 0).?, arms.len)) {
                .n => |n| n,
                .err => |e| return .{ .scalar = e },
            };
            return self.evalNode(arms[pick]);
        }

        const ops = try self.arena.alloc(Operand, arms.len);
        for (arms, ops) |node, *slot| slot.* = try self.operandOf(try self.evalNode(node));

        var shape = sel.shape;
        for (ops) |o| shape = maskShape(shape, o.shape, shape);
        var m = try self.newMatrix(shape.rows, shape.cols);
        var r: u32 = 0;
        while (r < shape.rows) : (r += 1) {
            var c: u32 = 0;
            while (c < shape.cols) : (c += 1) {
                const cv = sel.at(r, c) orelse {
                    m.set(r, c, value.incompatibleBroadcastFill());
                    continue;
                };
                switch (try self.chooseIndex(cv, arms.len)) {
                    .n => |n| m.set(r, c, ops[n].at(r, c) orelse value.incompatibleBroadcastFill()),
                    .err => |e| m.set(r, c, e),
                }
            }
        }
        return .{ .array = m };
    }

    const Pick = union(enum) { n: usize, err: value.ScalarValue };

    fn chooseIndex(self: *Evaluator, s: value.ScalarValue, arms: usize) EvalError!Pick {
        if (s == .err) return .{ .err = s };
        const n = switch (try self.toNumber(s)) {
            .n => |n| n,
            .err => |e| return .{ .err = e },
        };
        // Excel truncates toward zero, then requires 1 ≤ index ≤ arms.
        const t = std.math.trunc(n);
        if (t < 1 or t > @as(f64, @floatFromInt(arms))) {
            return .{ .err = value.ScalarValue.errorOf(.value) };
        }
        return .{ .n = @as(usize, @intFromFloat(t)) - 1 };
    }

    const CatchClass = enum { any_error, na_only };

    /// `IFERROR(value, fallback)` / `IFNA(value, fallback)`: the value
    /// argument evaluates, and the fallback **only on error**. Observing
    /// an error without becoming one is the `observe` propagation class.
    fn formIfError(
        self: *Evaluator,
        arg_nodes: []const parser.Index,
        class: CatchClass,
    ) EvalError!Value {
        const v = try self.evalNode(arg_nodes[0]);
        const op = try self.operandOf(v);

        var any = false;
        var r: u32 = 0;
        outer: while (r < op.shape.rows) : (r += 1) {
            var c: u32 = 0;
            while (c < op.shape.cols) : (c += 1) {
                if (caught(op.at(r, c).?, class)) {
                    any = true;
                    break :outer;
                }
            }
        }
        // No error anywhere: the fallback is never evaluated, so it
        // draws nothing and contributes no runtime dependency.
        if (!any) return v;

        const fallback = try self.operandOf(try self.evalNode(arg_nodes[1]));
        if (op.shape.isScalar()) {
            return .{ .scalar = fallback.at(0, 0) orelse value.incompatibleBroadcastFill() };
        }
        var m = try self.newMatrix(op.shape.rows, op.shape.cols);
        r = 0;
        while (r < op.shape.rows) : (r += 1) {
            var c: u32 = 0;
            while (c < op.shape.cols) : (c += 1) {
                const cur = op.at(r, c).?;
                m.set(r, c, if (caught(cur, class))
                    (fallback.at(r, c) orelse value.incompatibleBroadcastFill())
                else
                    cur);
            }
        }
        return .{ .array = m };
    }

    fn caught(s: value.ScalarValue, class: CatchClass) bool {
        if (s != .err) return false;
        return switch (class) {
            .any_error => true,
            .na_only => s.err == .known and s.err.known == .na,
        };
    }

    fn maskShape(a: value.Shape, b: value.Shape, c: value.Shape) value.Shape {
        return .{
            .rows = @max(a.rows, @max(b.rows, c.rows)),
            .cols = @max(a.cols, @max(b.cols, c.cols)),
        };
    }

    const Logical = union(enum) { b: bool, err: value.ScalarValue };

    /// Coerce to a condition. Text never coerces — `IF("TRUE",…)` is
    /// `#VALUE!` in Excel, which is why this is not `toNumber`.
    pub fn toLogical(self: *Evaluator, s: value.ScalarValue) EvalError!Logical {
        _ = self;
        return switch (s) {
            .boolean => |b| .{ .b = b },
            .number => |n| .{ .b = n != 0 },
            .blank => .{ .b = false },
            .err => .{ .err = s },
            .text => .{ .err = value.ScalarValue.errorOf(.value) },
        };
    }
};

/// The area `$A:$B` denotes. A free function because the evaluator and
/// the static walk must agree on it exactly, and two copies of a grid
/// bound is one too many.
fn fullColRange(first: parser.ColBound, last: parser.ColBound) coords.Range {
    const lo = @min(first.col.zeroBased(), last.col.zeroBased());
    const hi = @max(first.col.zeroBased(), last.col.zeroBased());
    return .{
        .first = .{
            .col = coords.Col.fromZeroBased(lo) catch unreachable,
            .row = coords.Row.fromOneBased(1) catch unreachable,
            .anchor = .{ .col = first.absolute },
        },
        .last = .{
            .col = coords.Col.fromZeroBased(hi) catch unreachable,
            .row = coords.Row.fromOneBased(coords.max_row) catch unreachable,
            .anchor = .{ .col = last.absolute },
        },
    };
}

/// The area `$1:$5` denotes.
fn fullRowRange(first: parser.RowBound, last: parser.RowBound) coords.Range {
    const lo = @min(first.row.oneBased(), last.row.oneBased());
    const hi = @max(first.row.oneBased(), last.row.oneBased());
    return .{
        .first = .{
            .col = coords.Col.fromZeroBased(0) catch unreachable,
            .row = coords.Row.fromOneBased(lo) catch unreachable,
            .anchor = .{ .row = first.absolute },
        },
        .last = .{
            .col = coords.Col.fromZeroBased(coords.max_col_1based - 1) catch unreachable,
            .row = coords.Row.fromOneBased(hi) catch unreachable,
            .anchor = .{ .row = last.absolute },
        },
    };
}

/// One `(operator, right operand)` step of a left spine.
const SpineStep = struct { op: parser.BinaryOp, rhs: parser.Index };

fn isReferenceOp(op: parser.BinaryOp) bool {
    return switch (op) {
        .range, .intersect, .union_op => true,
        else => false,
    };
}

fn boundingBox(a: coords.Range, b: coords.Range) coords.Range {
    return (coords.Range{
        .first = .{
            .col = coords.Col.fromZeroBased(@min(a.first.col.zeroBased(), b.first.col.zeroBased())) catch unreachable,
            .row = coords.Row.fromOneBased(@min(a.first.row.oneBased(), b.first.row.oneBased())) catch unreachable,
        },
        .last = .{
            .col = coords.Col.fromZeroBased(@max(a.last.col.zeroBased(), b.last.col.zeroBased())) catch unreachable,
            .row = coords.Row.fromOneBased(@max(a.last.row.oneBased(), b.last.row.oneBased())) catch unreachable,
        },
    }).normalized();
}

fn intersectRanges(a: coords.Range, b: coords.Range) ?coords.Range {
    if (!a.overlaps(b)) return null;
    return .{
        .first = .{
            .col = coords.Col.fromZeroBased(@max(a.first.col.zeroBased(), b.first.col.zeroBased())) catch unreachable,
            .row = coords.Row.fromOneBased(@max(a.first.row.oneBased(), b.first.row.oneBased())) catch unreachable,
        },
        .last = .{
            .col = coords.Col.fromZeroBased(@min(a.last.col.zeroBased(), b.last.col.zeroBased())) catch unreachable,
            .row = coords.Row.fromOneBased(@min(a.last.row.oneBased(), b.last.row.oneBased())) catch unreachable,
        },
    };
}

/// The cell one side of a `:` denotes, or null when it denotes
/// something the static walk cannot resolve without a workbook (a name,
/// a call, a structured reference) — those become areas at M4b3.
fn rangeEndpoint(ast: parser.Ast, i: parser.Index) ?coords.Cell {
    return switch (ast.node(i)) {
        .ref_cell => |n| n.cell,
        .paren => |n| rangeEndpoint(ast, n.child),
        else => null,
    };
}

// ─── static dependencies (§5.3a static-vs-runtime split) ─────────

/// Every reference the *text* mentions, whatever laziness does at run
/// time. The dependency graph is built from this, so a cell that only a
/// dead branch reads is still an edge and still triggers a recalc —
/// M5a's correctness rests on the two being different functions.
pub fn staticDependencies(
    allocator: std.mem.Allocator,
    ast: parser.Ast,
    current_sheet: env.SheetIndex,
    environment: env.EvalEnv,
    out: *DependencyLog,
) EvalError!void {
    try staticWalk(allocator, ast, ast.root, current_sheet, environment, out);
}

fn staticWalk(
    allocator: std.mem.Allocator,
    ast: parser.Ast,
    i: parser.Index,
    sheet: env.SheetIndex,
    environment: env.EvalEnv,
    out: *DependencyLog,
) EvalError!void {
    switch (ast.node(i)) {
        .number, .string, .boolean, .error_lit, .missing_arg, .name, .structured => {},
        .array => |n| for (ast.children(n.elems)) |c| {
            try staticWalk(allocator, ast, c, sheet, environment, out);
        },
        .ref_cell => |n| try out.noteCell(.{ .sheet = sheet, .row = n.cell.row, .col = n.cell.col }),
        .ref_full_col => |n| try out.noteArea(.{ .sheet = sheet, .range = fullColRange(n.first, n.last) }),
        .ref_full_row => |n| try out.noteArea(.{ .sheet = sheet, .range = fullRowRange(n.first, n.last) }),
        .qualified => |n| {
            // An unresolvable sheet contributes no edge — there is
            // nothing to depend on — and a 3D span is M4b3's.
            if (n.sheet.last != null) return;
            const idx = (try environment.resolveSheet(n.sheet.first)) orelse return;
            try staticWalk(allocator, ast, n.target, idx, environment, out);
        },
        .call => |n| {
            try staticWalk(allocator, ast, n.callee, sheet, environment, out);
            // Every arm, including the ones evaluation will skip. That
            // is the whole point.
            for (ast.children(n.args)) |c| {
                try staticWalk(allocator, ast, c, sheet, environment, out);
            }
        },
        .paren => |n| try staticWalk(allocator, ast, n.child, sheet, environment, out),
        .unary => |n| try staticWalk(allocator, ast, n.child, sheet, environment, out),
        .postfix => |n| try staticWalk(allocator, ast, n.child, sheet, environment, out),
        .binary => |n| {
            // `A1:B2` is two endpoints in the grammar and one area in
            // the graph. Recording only the endpoints would leave every
            // interior cell without a static edge, so the area is
            // recorded through the same bounding box the evaluator uses.
            if (n.op == .range) {
                if (rangeEndpoint(ast, n.lhs)) |a| {
                    if (rangeEndpoint(ast, n.rhs)) |b| {
                        try out.noteArea(.{ .sheet = sheet, .range = boundingBox(
                            .{ .first = a, .last = a },
                            .{ .first = b, .last = b },
                        ) });
                    }
                }
            }
            try staticWalk(allocator, ast, n.lhs, sheet, environment, out);
            try staticWalk(allocator, ast, n.rhs, sheet, environment, out);
        },
    }
}

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

/// The shipped fold, wired in the test section only — the same
/// arrangement `value.zig` uses, and for the same reason: a named module
/// rooted on `src/unicode/casefold.zig` collides with `zlsx` the moment
/// the engine is imported from `src/`, so the semantics take the fold as
/// a parameter and only the tests name it.
const casefold = @import("zlsx_casefold");

fn shippedFold(allocator: std.mem.Allocator, s: []const u8) anyerror![]u8 {
    return casefold.foldString(allocator, s);
}

test {
    // Force analysis of the files this module owns. Zig runs the tests
    // of an imported file only once something in it is referenced, and
    // a registry nobody touched is a registry whose tests never ran.
    _ = registry;
    _ = env;
}

fn cellOf(a1: []const u8) coords.Cell {
    return coords.parseCell(a1, .{ .dollar = .accept }) catch unreachable;
}

const Harness = struct {
    gpa: std.mem.Allocator,
    arena_state: std.heap.ArenaAllocator,
    fake: env.Fake,
    asts: std.ArrayListUnmanaged(parser.Ast) = .empty,
    draw_value: f64 = 0.5,
    draws: DrawSource = undefined,
    sheet: env.SheetIndex = undefined,
    ev: Evaluator = undefined,
    have_ev: bool = false,

    fn init(h: *Harness, gpa: std.mem.Allocator) !void {
        h.* = .{
            .gpa = gpa,
            .arena_state = std.heap.ArenaAllocator.init(gpa),
            .fake = env.Fake.init(gpa),
        };
        h.draws = DrawSource.constant(&h.draw_value);
        h.sheet = try h.fake.addSheet("Sheet1");
    }

    fn deinit(h: *Harness) void {
        if (h.have_ev) h.ev.deinit();
        for (h.asts.items) |*a| a.deinit(h.gpa);
        h.asts.deinit(h.gpa);
        h.fake.deinit();
        h.arena_state.deinit();
    }

    fn arena(h: *Harness) std.mem.Allocator {
        return h.arena_state.allocator();
    }

    fn options(h: *Harness) Options {
        return .{
            .current_sheet = h.sheet,
            .collation = .{ .fold = shippedFold },
            .draws = &h.draws,
        };
    }

    fn parse(h: *Harness, src: []const u8) !parser.Ast {
        return h.parseWith(src, .{});
    }

    fn parseWith(h: *Harness, src: []const u8, opts: parser.Options) !parser.Ast {
        var parsed = try parser.parse(h.gpa, src, opts);
        if (parsed == .refused) {
            parsed.deinit(h.gpa);
            return error.ParseRefused;
        }
        try h.asts.append(h.gpa, parsed.ok);
        return parsed.ok;
    }

    fn eval(h: *Harness, src: []const u8) !Value {
        return h.evalOpts(src, h.options());
    }

    fn evalOpts(h: *Harness, src: []const u8, opts: Options) !Value {
        // The evaluator's limits and the parser's are the same `Limits`,
        // so a test that raises one raises both.
        const ast = try h.parseWith(src, .{ .limits = opts.limits });
        if (h.have_ev) h.ev.deinit();
        h.ev = Evaluator.init(h.arena(), h.fake.evalEnv(), opts);
        h.have_ev = true;
        return h.ev.evaluate(ast);
    }

    fn scalar(h: *Harness, src: []const u8) !value.ScalarValue {
        const v = try h.eval(src);
        try testing.expect(v == .scalar);
        return v.scalar;
    }

    fn put(h: *Harness, a1: []const u8, v: value.ScalarValue) !void {
        try h.fake.putA1(h.sheet, .stored, a1, v);
    }
};

fn num(v: f64) value.ScalarValue {
    return value.ScalarValue.fromNumber(v);
}

// ─── the §5.3b tables, applied ───────────────────────────────────

test "operators: the committed hand-spec rows, evaluated end to end" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();

    try testing.expectEqual(@as(f64, 2), (try h.scalar("1+1")).number);
    try testing.expectEqual(@as(f64, 64), (try h.scalar("2^3^2")).number);
    try testing.expectEqual(@as(f64, 1), (try h.scalar("-1^2")).number);
    try testing.expectEqual(@as(f64, 2), (try h.scalar("TRUE()+1")).number);
    try testing.expectEqual(@as(f64, 0.1), (try h.scalar("10%")).number);
    try testing.expectEqualStrings("1x", (try h.scalar("1&\"x\"")).text);
    try testing.expect((try h.scalar("\"a\"<\"B\"")).boolean);
    try testing.expectEqual(value.KnownError.div0, (try h.scalar("1/0")).err.known);
    try testing.expectEqual(value.KnownError.num, (try h.scalar("SQRT(-1)")).err.known);
    try testing.expectEqual(value.KnownError.value, (try h.scalar("\"a\"+1")).err.known);
    try testing.expectEqual(value.KnownError.num, (try h.scalar("1E+308*10")).err.known);
}

test "coercion: the arithmetic column of the §5.3b matrix, behaviourally" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try h.put("A1", .{ .text = "1" }); // numeric text
    try h.put("A2", .{ .text = "" }); // empty text
    try h.put("A3", .{ .boolean = true });
    try h.put("A4", .{ .text = "abc" });
    // A5 is a true blank.

    try testing.expectEqual(@as(f64, 2), (try h.scalar("A1+1")).number);
    try testing.expectEqual(value.KnownError.value, (try h.scalar("A2+1")).err.known);
    try testing.expectEqual(@as(f64, 2), (try h.scalar("A3+1")).number);
    try testing.expectEqual(value.KnownError.value, (try h.scalar("A4+1")).err.known);
    try testing.expectEqual(@as(f64, 1), (try h.scalar("A5+1")).number); // blank is 0

    // `&` sees the same values as text, blank as `""`.
    try testing.expectEqualStrings("TRUEx", (try h.scalar("A3&\"x\"")).text);
    try testing.expectEqualStrings("x", (try h.scalar("A5&\"x\"")).text);
}

test "coercion: locale-flavoured text refuses instead of guessing" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try h.put("A1", .{ .text = "1,5" });
    // Neither a guessed 1.5 nor a guessed `#VALUE!`: a plane-2 refusal,
    // which is the whole point of the locale classifier.
    try testing.expectError(error.LocaleSensitiveInput, h.scalar("A1+1"));
    try testing.expectEqual(
        parser.PlaneTwo.FormulaLocaleSensitiveInput,
        planeTwo(error.LocaleSensitiveInput),
    );
}

test "comparison: case-insensitive text, cross-type order, blank adoption" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try h.put("A1", .{ .text = "Straße" });
    try h.put("A2", .{ .text = "STRASSE" });
    // A3 blank, A4 blank.
    try h.put("B1", num(0));
    try h.put("B2", .{ .text = "x" });

    // `collation_v1`: fold-equal is equal, for ordering as well.
    try testing.expect((try h.scalar("A1=A2")).boolean);
    try testing.expect(!(try h.scalar("A1<A2")).boolean);
    try testing.expect((try h.scalar("A1<=A2")).boolean);

    // number < text < logical, never cross-equal.
    try testing.expect((try h.scalar("1<\"a\"")).boolean);
    try testing.expect((try h.scalar("\"a\"<TRUE()")).boolean);
    try testing.expect(!(try h.scalar("1=\"1\"")).boolean);

    // Blank adopts the other operand's type: zero against a number,
    // `""` against text, FALSE against a boolean.
    try testing.expect((try h.scalar("A3=B1")).boolean);
    try testing.expect((try h.scalar("A3=A4")).boolean);
    try testing.expect(!(try h.scalar("A3=B2")).boolean);
    try testing.expect((try h.scalar("A3=FALSE()")).boolean);
}

test "comparison: N2 puts `(0.1+0.2)=0.3` on different sides in the two modes" {
    // Both committed manifests decide this row, in opposite directions:
    // the LibreOffice excel-fidelity suite records TRUE, the hand-spec
    // ieee suite records FALSE. A comparison is a subtraction against
    // zero, so it is inside N2's additive scope — and that is the only
    // reading under which both manifests can be satisfied at once.
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();

    var excel = h.options();
    excel.fidelity = .excel;
    var ieee = h.options();
    ieee.fidelity = .ieee;

    try testing.expect((try h.evalOpts("(0.1+0.2)=0.3", excel)).scalar.boolean);
    try testing.expect(!(try h.evalOpts("(0.1+0.2)=0.3", ieee)).scalar.boolean);
}

test "broadcast: dims of size one stretch, incompatible positions fill #N/A" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();

    {
        const v = try h.eval("{1,2}+{10,20}");
        try testing.expectEqual(@as(u32, 1), v.array.rows);
        try testing.expectEqual(@as(f64, 11), v.array.at(0, 0).number);
        try testing.expectEqual(@as(f64, 22), v.array.at(0, 1).number);
    }
    {
        // 2×1 against 1×2 broadcasts to 2×2.
        const v = try h.eval("{1;2}+{10,20}");
        try testing.expectEqual(@as(u32, 2), v.array.rows);
        try testing.expectEqual(@as(u32, 2), v.array.cols);
        try testing.expectEqual(@as(f64, 11), v.array.at(0, 0).number);
        try testing.expectEqual(@as(f64, 21), v.array.at(0, 1).number);
        try testing.expectEqual(@as(f64, 12), v.array.at(1, 0).number);
        try testing.expectEqual(@as(f64, 22), v.array.at(1, 1).number);
    }
    {
        // Both extents greater than one and unequal: elementwise `#N/A`
        // at the positions no operand can supply — not a refusal.
        const v = try h.eval("{1;2;3}+{10;20}");
        try testing.expectEqual(@as(u32, 3), v.array.rows);
        try testing.expectEqual(@as(f64, 11), v.array.at(0, 0).number);
        try testing.expectEqual(@as(f64, 22), v.array.at(1, 0).number);
        try testing.expectEqual(value.KnownError.na, v.array.at(2, 0).err.known);
    }
}

test "shape: an array in a scalar slot follows the dialect, not the function" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();

    var da = h.options();
    da.dialect = .dynamic_array;
    var legacy = h.options();
    legacy.dialect = .legacy;

    // §5.3b: `spill_or_iterate` under DA…
    const lifted = try h.evalOpts("SQRT({4,9})", da);
    try testing.expectEqual(@as(u32, 2), lifted.array.cols);
    try testing.expectEqual(@as(f64, 2), lifted.array.at(0, 0).number);
    try testing.expectEqual(@as(f64, 3), lifted.array.at(0, 1).number);

    // …and `top_left_reduction` under legacy. Arrays are NOT references:
    // this is a reduction, not an intersection.
    const reduced = try h.evalOpts("SQRT({4,9})", legacy);
    try testing.expectEqual(@as(f64, 2), reduced.scalar.number);
}

test "shape: legacy dereference intersects, DA dereference does not" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try h.put("A1", num(1));
    try h.put("A2", num(2));
    try h.put("A3", num(3));

    var legacy = h.options();
    legacy.dialect = .legacy;
    legacy.site = .{ .row = try coords.Row.fromOneBased(2), .col = try coords.Col.fromZeroBased(3) };

    // The evaluation row falls inside A1:A3, so the column projects to A2.
    try testing.expectEqual(@as(f64, 4), (try h.evalOpts("A1:A3*2", legacy)).scalar.number);

    // Off-axis: no intersection, `#VALUE!` — not a guess and not a spill.
    var off = legacy;
    off.site = .{ .row = try coords.Row.fromOneBased(9), .col = try coords.Col.fromZeroBased(3) };
    try testing.expectEqual(value.KnownError.value, (try h.evalOpts("A1:A3*2", off)).scalar.err.known);

    // The same text under DA dereferences the whole area instead.
    var da = h.options();
    da.dialect = .dynamic_array;
    const spilled = try h.evalOpts("A1:A3*2", da);
    try testing.expectEqual(@as(u32, 3), spilled.array.rows);
    try testing.expectEqual(@as(f64, 6), spilled.array.at(2, 0).number);
}

test "shape: the three `@` rows" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try h.put("A1", num(1));
    try h.put("A2", num(2));
    try h.put("A3", num(3));

    var opts = h.options();
    opts.site = .{ .row = try coords.Row.fromOneBased(2), .col = try coords.Col.fromZeroBased(5) };

    // Single-cell reference: unchanged, whatever the evaluation site is.
    try testing.expectEqual(@as(f64, 1), (try h.evalOpts("@A1+0", opts)).scalar.number);
    // Multi-cell: row/column intersection with the site.
    try testing.expectEqual(@as(f64, 2), (try h.evalOpts("@A1:A3", opts)).scalar.number);
    // Array: top-left, not an intersection.
    try testing.expectEqual(@as(f64, 1), (try h.evalOpts("@{1,2;3,4}", opts)).scalar.number);

    // Site-less evaluation of a site-dependent construct refuses.
    try testing.expectError(error.AnchorRequired, h.evalOpts("@A1:A3", h.options()));
    try testing.expectEqual(parser.PlaneTwo.FormulaAnchorRequired, planeTwo(error.AnchorRequired));
}

test "references: `:` spans, ` ` intersects, `,` unions" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try h.put("A1", num(1));
    try h.put("B1", num(2));
    try h.put("A2", num(4));
    try h.put("B2", num(8));

    try testing.expectEqual(@as(f64, 15), (try h.scalar("SUM(A1:B2)")).number);
    // The intersection of A1:B2 and B1:B9 is B1:B2.
    try testing.expectEqual(@as(f64, 10), (try h.scalar("SUM(A1:B2 B1:B9)")).number);
    // A disjoint intersection is `#NULL!`.
    try testing.expectEqual(value.KnownError.null_err, (try h.scalar("SUM(A1:A2 B1:B2)")).err.known);
    // A union is a multi-area reference an aggregate can consume…
    try testing.expectEqual(@as(f64, 13), (try h.scalar("SUM((A1:A2,B2:B2))")).number);
    // …but not a representable result (§10).
    try testing.expectError(error.ResultNotRepresentable, h.eval("(A1:A2,B2:B2)"));
}

// ─── §5.3a per-form evaluation contracts ─────────────────────────

const Expect = union(enum) {
    number: f64,
    boolean: bool,
    err: value.KnownError,
    array: Arr,

    const Arr = struct { rows: u32, cols: u32, cells: []const Expect };
};

fn expectValue(exp: Expect, v: Value) !void {
    switch (exp) {
        .number => |n| {
            try testing.expect(v == .scalar);
            try testing.expectApproxEqAbs(n, v.scalar.number, 1e-12);
        },
        .boolean => |b| {
            try testing.expect(v == .scalar);
            try testing.expectEqual(b, v.scalar.boolean);
        },
        .err => |e| {
            try testing.expect(v == .scalar);
            try testing.expect(v.scalar == .err);
            try testing.expectEqual(e, v.scalar.err.known);
        },
        .array => |a| {
            try testing.expect(v == .array);
            try testing.expectEqual(a.rows, v.array.rows);
            try testing.expectEqual(a.cols, v.array.cols);
            for (a.cells, 0..) |cell, k| {
                const r: u32 = @intCast(k / a.cols);
                const c: u32 = @intCast(k % a.cols);
                try expectValue(cell, .{ .scalar = v.array.at(r, c) });
            }
        },
    }
}

/// One §5.3a contract: the value, the volatile draws, and both halves of
/// the dependency split — in every branch position.
const FormCase = struct {
    formula: []const u8,
    expect: Expect,
    draws: u64,
    /// Read at run time. Laziness governs this list.
    runtime: []const []const u8,
    /// Provably NOT read at run time. The half a "value is right" test
    /// would never notice going wrong.
    absent: []const []const u8 = &.{},
    /// Syntactic edges, which laziness never governs.
    static: []const []const u8,
};

const form_cases = [_]FormCase{
    // ── IF: scalar selector takes exactly one arm ──
    .{ .formula = "IF(TRUE,A1,B1)", .expect = .{ .number = 10 }, .draws = 0, .runtime = &.{"A1"}, .absent = &.{"B1"}, .static = &.{ "A1", "B1" } },
    .{ .formula = "IF(FALSE,A1,B1)", .expect = .{ .number = 20 }, .draws = 0, .runtime = &.{"B1"}, .absent = &.{"A1"}, .static = &.{ "A1", "B1" } },
    .{ .formula = "IF(A1>5,A1,B1)", .expect = .{ .number = 10 }, .draws = 0, .runtime = &.{"A1"}, .absent = &.{"B1"}, .static = &.{ "A1", "B1" } },
    // The omitted third argument is FALSE, not blank.
    .{ .formula = "IF(FALSE,A1)", .expect = .{ .boolean = false }, .draws = 0, .runtime = &.{}, .absent = &.{"A1"}, .static = &.{"A1"} },
    // Zero draws in the dead branch, in either branch position.
    .{ .formula = "IF(TRUE,RAND(),RAND())", .expect = .{ .number = 0.5 }, .draws = 1, .runtime = &.{}, .static = &.{} },
    .{ .formula = "IF(FALSE,RAND(),RAND())", .expect = .{ .number = 0.5 }, .draws = 1, .runtime = &.{}, .static = &.{} },
    // A condition that errors takes neither arm.
    .{ .formula = "IF(C1,RAND(),RAND())", .expect = .{ .err = .na }, .draws = 0, .runtime = &.{"C1"}, .static = &.{"C1"} },
    // ── IF: an ARRAY selector switches the form to per-element masking ──
    .{
        .formula = "IF({TRUE;FALSE},A1,B1)",
        .expect = .{ .array = .{ .rows = 2, .cols = 1, .cells = &.{ .{ .number = 10 }, .{ .number = 20 } } } },
        .draws = 0,
        // BOTH arms evaluate now, so both are captured at run time.
        .runtime = &.{ "A1", "B1" },
        .static = &.{ "A1", "B1" },
    },
    .{ .formula = "IF({TRUE;FALSE},RAND(),RAND())", .expect = .{ .array = .{ .rows = 2, .cols = 1, .cells = &.{ .{ .number = 0.5 }, .{ .number = 0.5 } } } }, .draws = 2, .runtime = &.{}, .static = &.{} },

    // ── IFERROR / IFNA: the fallback only on error ──
    .{ .formula = "IFERROR(A1,RAND())", .expect = .{ .number = 10 }, .draws = 0, .runtime = &.{"A1"}, .static = &.{"A1"} },
    .{ .formula = "IFERROR(C1,RAND())", .expect = .{ .number = 0.5 }, .draws = 1, .runtime = &.{"C1"}, .static = &.{"C1"} },
    .{ .formula = "IFERROR(1/0,B1)", .expect = .{ .number = 20 }, .draws = 0, .runtime = &.{"B1"}, .static = &.{"B1"} },
    .{ .formula = "IFNA(C1,B1)", .expect = .{ .number = 20 }, .draws = 0, .runtime = &.{ "C1", "B1" }, .static = &.{ "C1", "B1" } },
    // The distinction that makes IFNA a separate form: a `#DIV/0!` is
    // not caught, so the fallback is never evaluated.
    .{ .formula = "IFNA(1/0,RAND())", .expect = .{ .err = .div0 }, .draws = 0, .runtime = &.{}, .static = &.{} },
    .{ .formula = "IFNA(1/0,B1)", .expect = .{ .err = .div0 }, .draws = 0, .runtime = &.{}, .absent = &.{"B1"}, .static = &.{"B1"} },

    // ── CHOOSE: lazy over its arms, in each arm position ──
    .{ .formula = "CHOOSE(1,A1,B1)", .expect = .{ .number = 10 }, .draws = 0, .runtime = &.{"A1"}, .absent = &.{"B1"}, .static = &.{ "A1", "B1" } },
    .{ .formula = "CHOOSE(2,A1,B1)", .expect = .{ .number = 20 }, .draws = 0, .runtime = &.{"B1"}, .absent = &.{"A1"}, .static = &.{ "A1", "B1" } },
    .{ .formula = "CHOOSE(1,RAND(),RAND())", .expect = .{ .number = 0.5 }, .draws = 1, .runtime = &.{}, .static = &.{} },
    .{ .formula = "CHOOSE(2,RAND(),RAND())", .expect = .{ .number = 0.5 }, .draws = 1, .runtime = &.{}, .static = &.{} },
    // Out of range takes no arm at all.
    .{ .formula = "CHOOSE(3,A1,B1)", .expect = .{ .err = .value }, .draws = 0, .runtime = &.{}, .absent = &.{ "A1", "B1" }, .static = &.{ "A1", "B1" } },

    // ── IFS and SWITCH are EAGER. Excel evaluates every arm and so do we. ──
    .{ .formula = "IFS(FALSE,RAND(),TRUE,RAND())", .expect = .{ .number = 0.5 }, .draws = 2, .runtime = &.{}, .static = &.{} },
    .{ .formula = "IFS(FALSE,A1,TRUE,B1)", .expect = .{ .number = 20 }, .draws = 0, .runtime = &.{ "A1", "B1" }, .static = &.{ "A1", "B1" } },
    .{ .formula = "IFS(FALSE,A1,FALSE,B1)", .expect = .{ .err = .na }, .draws = 0, .runtime = &.{ "A1", "B1" }, .static = &.{ "A1", "B1" } },
    .{ .formula = "SWITCH(1,1,RAND(),2,RAND())", .expect = .{ .number = 0.5 }, .draws = 2, .runtime = &.{}, .static = &.{} },
    .{ .formula = "SWITCH(2,1,A1,2,B1)", .expect = .{ .number = 20 }, .draws = 0, .runtime = &.{ "A1", "B1" }, .static = &.{ "A1", "B1" } },
    .{ .formula = "SWITCH(3,1,A1,2,B1,A1)", .expect = .{ .number = 10 }, .draws = 0, .runtime = &.{ "A1", "B1" }, .static = &.{ "A1", "B1" } },

    // ── AND / OR are eager too: no short-circuit, so the draw happens ──
    .{ .formula = "AND(FALSE,RAND()>2)", .expect = .{ .boolean = false }, .draws = 1, .runtime = &.{}, .static = &.{} },
    .{ .formula = "OR(TRUE,RAND()>2)", .expect = .{ .boolean = true }, .draws = 1, .runtime = &.{}, .static = &.{} },
    .{ .formula = "AND(A1>5,B1>5)", .expect = .{ .boolean = true }, .draws = 0, .runtime = &.{ "A1", "B1" }, .static = &.{ "A1", "B1" } },
};

test "forms: every §5.3a contract, in value, draws, and both dependency halves" {
    for (form_cases) |c| {
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        try h.put("A1", num(10));
        try h.put("B1", num(20));
        try h.put("C1", value.ScalarValue.errorOf(.na));

        const v = h.eval(c.formula) catch |e| {
            std.debug.print("form case `{s}` refused: {t}\n", .{ c.formula, e });
            return e;
        };
        expectValue(c.expect, v) catch |e| {
            std.debug.print("form case `{s}`: wrong value\n", .{c.formula});
            return e;
        };

        if (h.draws.count != c.draws) {
            std.debug.print(
                "form case `{s}`: expected {d} volatile draws, saw {d}\n",
                .{ c.formula, c.draws, h.draws.count },
            );
            return error.WrongDrawCount;
        }

        for (c.runtime) |a1| {
            const cell = cellOf(a1);
            if (!h.ev.deps.hasCell(.{ .sheet = h.sheet, .row = cell.row, .col = cell.col })) {
                std.debug.print("form case `{s}`: {s} was not captured at run time\n", .{ c.formula, a1 });
                return error.MissingRuntimeDependency;
            }
        }
        for (c.absent) |a1| {
            const cell = cellOf(a1);
            if (h.ev.deps.hasCell(.{ .sheet = h.sheet, .row = cell.row, .col = cell.col })) {
                std.debug.print("form case `{s}`: {s} was read despite being in a dead arm\n", .{ c.formula, a1 });
                return error.UnexpectedRuntimeDependency;
            }
        }

        // The static half: every arm, always, whatever laziness did.
        var static = DependencyLog.init(h.arena());
        const ast = try h.parse(c.formula);
        try staticDependencies(h.arena(), ast, h.sheet, h.fake.evalEnv(), &static);
        for (c.static) |a1| {
            const cell = cellOf(a1);
            if (!static.hasCell(.{ .sheet = h.sheet, .row = cell.row, .col = cell.col })) {
                std.debug.print("form case `{s}`: {s} is missing from the static edges\n", .{ c.formula, a1 });
                return error.MissingStaticDependency;
            }
        }
        try testing.expectEqual(c.static.len, static.cells.items.len);
    }
}

test "forms: the static edge set is a superset of what any run reads" {
    // §5.3a's static-vs-runtime split, stated as an invariant over the
    // whole fixture table rather than case by case: laziness may only
    // ever *remove* reads, never invent one the text does not mention.
    for (form_cases) |c| {
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        try h.put("A1", num(10));
        try h.put("B1", num(20));
        try h.put("C1", value.ScalarValue.errorOf(.na));

        _ = try h.eval(c.formula);
        var static = DependencyLog.init(h.arena());
        const ast = try h.parse(c.formula);
        try staticDependencies(h.arena(), ast, h.sheet, h.fake.evalEnv(), &static);

        for (h.ev.deps.cells.items) |cell| {
            if (!static.hasCell(cell)) {
                std.debug.print("`{s}`: read a cell no static edge covers\n", .{c.formula});
                return error.RuntimeDependencyOutsideStatic;
            }
        }
        for (h.ev.deps.areas.items) |area| {
            if (!static.hasArea(area)) {
                std.debug.print("`{s}`: read an area no static edge covers\n", .{c.formula});
                return error.RuntimeDependencyOutsideStatic;
            }
        }
        try testing.expect(h.ev.deps.cells.items.len <= static.cells.items.len);
    }
}

test "dependencies: an area is one static edge, not two endpoints" {
    // `A1:B2` is two endpoints in the grammar and one area in the graph.
    // Recording only the endpoints would leave B1 and A2 without an edge,
    // so an edit to either would never trigger a recalculation.
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try h.put("A1", num(1));
    try h.put("B2", num(2));

    const src = "SUM(A1:B2)+SUM(A:A)+SUM(2:3)";
    _ = try h.eval(src);
    var static = DependencyLog.init(h.arena());
    const ast = try h.parse(src);
    try staticDependencies(h.arena(), ast, h.sheet, h.fake.evalEnv(), &static);

    for (h.ev.deps.areas.items) |area| {
        try testing.expect(static.hasArea(area));
    }
    // The rectangle, the whole column, and the two-row band.
    try testing.expectEqual(@as(usize, 3), static.areas.items.len);
    try testing.expect(static.areas.items[0].range.contains(cellOf("B1")));
    try testing.expectEqual(@as(u32, coords.max_row), static.areas.items[1].range.last.row.oneBased());
    try testing.expectEqual(@as(u32, coords.max_col_1based - 1), static.areas.items[2].range.last.col.zeroBased());
}

// ─── §5.3c propagation classes ───────────────────────────────────

test "propagation: COUNT, COUNTA, and COUNTBLANK disagree on purpose" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try h.put("A1", num(1));
    try h.put("A2", .{ .text = "" }); // an `=""` result
    try h.put("A3", value.ScalarValue.errorOf(.na));
    try h.put("A4", .{ .text = "7" }); // numeric text, in a range
    // A5 is a true blank.

    // Numbers only. The error in A3 is neither counted nor propagated,
    // and numeric text in a range is NOT coerced.
    try testing.expectEqual(@as(f64, 1), (try h.scalar("COUNT(A1:A5)")).number);
    // Everything that is not a true blank, errors and `""` included.
    try testing.expectEqual(@as(f64, 4), (try h.scalar("COUNTA(A1:A5)")).number);
    // True blanks plus the `""`.
    try testing.expectEqual(@as(f64, 2), (try h.scalar("COUNTBLANK(A1:A5)")).number);
    // And ISBLANK, which is FALSE for the same `""` COUNTBLANK counted.
    try testing.expect(!(try h.scalar("ISBLANK(A2)")).boolean);
    try testing.expect((try h.scalar("ISBLANK(A5)")).boolean);

    // A direct argument coerces where a range element does not.
    try testing.expectEqual(@as(f64, 1), (try h.scalar("COUNT(\"7\")")).number);
    try testing.expectEqual(@as(f64, 0), (try h.scalar("COUNT(\"x\")")).number);
}

test "propagation: SUM propagates from inside a range, COUNT does not" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try h.put("A1", num(1));
    try h.put("A2", value.ScalarValue.errorOf(.div0));
    try h.put("A3", num(2));
    try h.put("A4", .{ .boolean = true });
    try h.put("A5", .{ .text = "9" });

    try testing.expectEqual(value.KnownError.div0, (try h.scalar("SUM(A1:A3)")).err.known);
    try testing.expectEqual(@as(f64, 2), (try h.scalar("COUNT(A1:A3)")).number);
    // Booleans and numeric text found in a range are ignored, so only
    // A3's 2 counts — a coercing implementation would answer 12.
    try testing.expectEqual(@as(f64, 2), (try h.scalar("SUM(A3:A5)")).number);
    // …but a direct boolean argument is 1, and direct numeric text coerces.
    try testing.expectEqual(@as(f64, 10), (try h.scalar("SUM(TRUE(),\"9\")")).number);
}

test "propagation: `observe` looks at an error without becoming one" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try h.put("A1", value.ScalarValue.errorOf(.div0));
    try h.put("A2", num(3));

    try testing.expect((try h.scalar("ISERROR(A1)")).boolean);
    try testing.expect(!(try h.scalar("ISERROR(A2)")).boolean);
    try testing.expect((try h.scalar("ISNUMBER(A2)")).boolean);
    // A `propagate` function in the same expression still propagates.
    try testing.expectEqual(value.KnownError.div0, (try h.scalar("SQRT(A1)")).err.known);
    // …and `observe` wrapping it does not.
    try testing.expect((try h.scalar("ISERROR(SQRT(A1))")).boolean);
}

test "propagation: operands evaluate left to right and the first error wins" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try h.put("A1", value.ScalarValue.errorOf(.div0));
    try h.put("A2", value.ScalarValue.errorOf(.na));

    try testing.expectEqual(value.KnownError.div0, (try h.scalar("A1+A2")).err.known);
    try testing.expectEqual(value.KnownError.na, (try h.scalar("A2+A1")).err.known);
    // Including through a coercion that would otherwise have failed.
    try testing.expectEqual(value.KnownError.div0, (try h.scalar("A1+\"x\"")).err.known);
}

test "propagation: per_element keeps an error in its own cell" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try h.put("A1", num(1));
    try h.put("A2", value.ScalarValue.errorOf(.na));
    try h.put("A3", num(3));

    var da = h.options();
    da.dialect = .dynamic_array;
    const v = try h.evalOpts("A1:A3*2", da);
    try testing.expectEqual(@as(f64, 2), v.array.at(0, 0).number);
    try testing.expectEqual(value.KnownError.na, v.array.at(1, 0).err.known);
    try testing.expectEqual(@as(f64, 6), v.array.at(2, 0).number);
}

// ─── boundaries and refusals ─────────────────────────────────────

fn emptyMatrixImpl(ctx: registry.CallCtx, args: []const Value) registry.FnError!Value {
    _ = ctx;
    _ = args;
    // What FILTER does at M7a when nothing matches.
    return error.EmptyMatrix;
}

test "boundary: an empty matrix normalizes to #CALC! at the producing call" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    h.ev = Evaluator.init(h.arena(), h.fake.evalEnv(), h.options());
    h.have_ev = true;

    const f = registry.Function{
        .name = "__EMPTY",
        .arity = .{ .min = 0, .max = 0, .fixed = &.{}, .rest = &.{} },
        .coercion = .{ .fixed = &.{}, .rest = &.{} },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = emptyMatrixImpl,
    };
    const v = try h.ev.invoke(&f, &.{});
    // Excel's own answer, and it never travels as a matrix nobody can
    // represent.
    try testing.expectEqual(value.KnownError.calc, v.scalar.err.known);
}

test "refusals: the plane-2 map is total and says what §10 says" {
    // Exhaustive by construction — a new `EvalError` fails to compile
    // until it is mapped — so this pins the mapping rather than its
    // existence.
    try testing.expectEqual(parser.PlaneTwo.FormulaUnsupportedFunction, planeTwo(error.UnsupportedFunction));
    try testing.expectEqual(parser.PlaneTwo.FormulaUnsupportedConstruct, planeTwo(error.NotYetImplemented));
    try testing.expectEqual(parser.PlaneTwo.FormulaMalformedInput, planeTwo(error.MalformedInput));
    try testing.expectEqual(parser.PlaneTwo.FormulaResultNotRepresentable, planeTwo(error.ResultNotRepresentable));
    try testing.expectEqual(parser.PlaneTwo.FormulaLimitExceeded, planeTwo(error.LimitExceeded));
    try testing.expectEqual(@as(usize, 1), not_yet_implemented.len);
}

test "refusals: an unregistered call refuses, an unresolvable name is #NAME?" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();

    // §7: unregistered calls refuse rather than inventing `#NAME?` for a
    // function zlsx simply does not implement.
    try testing.expectError(error.UnsupportedFunction, h.eval("VLOOKUP(1,A1:B2,2)"));
    try testing.expectError(error.UnsupportedFunction, h.eval("NOTAFUNCTION()"));
    // §5.9: a value-position name that provably resolves nowhere is a
    // plane-1 `#NAME?`, which is a successful result.
    try testing.expectEqual(value.KnownError.name, (try h.scalar("SomeName+1")).err.known);
    // Wrong arity is a formula Excel could not have written.
    try testing.expectError(error.MalformedInput, h.eval("SQRT(1,2)"));
    try testing.expectError(error.MalformedInput, h.eval("IFERROR(1)"));
}

test "refusals: constructs whose milestone has not landed refuse by name" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    _ = try h.fake.addSheet("Q1");
    // A 3D span is M4b3's; it refuses rather than being read as one
    // strangely named sheet.
    try testing.expectError(error.NotYetImplemented, h.eval("'Q1:Q4'!A1"));
}

test "references: an unknown sheet is #REF!, a known one resolves" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    const other = try h.fake.addSheet("Data");
    try h.fake.putA1(other, .stored, "A1", num(41));

    try testing.expectEqual(@as(f64, 42), (try h.scalar("Data!A1+1")).number);
    try testing.expectEqual(value.KnownError.ref, (try h.scalar("Missing!A1")).err.known);
}

test "references: `A1#` resolves through the anchor, or is #REF!" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    const a1 = cellOf("A1");
    try h.fake.put(h.sheet, .stored, .{
        .row = a1.row,
        .col = a1.col,
        .v = num(1),
        .spill = .{ .rows = 3, .cols = 1 },
    });
    try h.put("A2", num(2));
    try h.put("A3", num(3));
    try h.put("B1", num(9));

    try testing.expectEqual(@as(f64, 6), (try h.scalar("SUM(A1#)")).number);
    // A cell that is not an anchor has no spilled range.
    try testing.expectEqual(value.KnownError.ref, (try h.scalar("B1#")).err.known);
}

test "literals: string escapes, extensible error spellings, and blanks" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();

    try testing.expectEqualStrings("a\"b", (try h.scalar("\"a\"\"b\"")).text);
    try testing.expectEqualStrings("", (try h.scalar("\"\"")).text);
    // An error literal is a value: `#N/A` is one of the frozen ten.
    try testing.expectEqual(value.KnownError.na, (try h.scalar("#N/A")).err.known);
    // Blank is neither `""` nor zero-that-counts.
    try testing.expect((try h.scalar("ISBLANK(Z99)")).boolean);
}

test "limits: a flat expression is not a deep one" {
    // A 400-term sum is one grammar production deep and 400 nodes deep.
    // Bounding the walk by the parser's nesting limit would refuse this,
    // which is why the two limits are separate fields.
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();

    var src: std.ArrayListUnmanaged(u8) = .empty;
    defer src.deinit(testing.allocator);
    try src.appendSlice(testing.allocator, "1");
    for (0..399) |_| try src.appendSlice(testing.allocator, "+1");
    try testing.expectEqual(@as(f64, 400), (try h.scalar(src.items)).number);

    // Even a 5 000-term chain, which does not fit in the call stack if
    // the spine is recursed rather than folded.
    var long: std.ArrayListUnmanaged(u8) = .empty;
    defer long.deinit(testing.allocator);
    try long.appendSlice(testing.allocator, "1");
    for (0..4999) |_| try long.appendSlice(testing.allocator, "+1");
    var wide = h.options();
    wide.limits.max_formula_chars = 65_536;
    wide.limits.max_tokens = 65_536;
    wide.limits.max_ast_nodes = 65_536;
    try testing.expectEqual(@as(f64, 5000), (try h.evalOpts(long.items, wide)).scalar.number);

    // Genuine *depth* — parenthesis nesting — still refuses with a typed
    // limit rather than running the stack out. The parser's bound is
    // raised here so the evaluator's is the one that fires.
    var nested: std.ArrayListUnmanaged(u8) = .empty;
    defer nested.deinit(testing.allocator);
    for (0..600) |_| try nested.append(testing.allocator, '(');
    try nested.append(testing.allocator, '1');
    for (0..600) |_| try nested.append(testing.allocator, ')');
    var deep = wide;
    deep.limits.max_parse_depth = 4096;
    deep.max_expr_depth = 100;
    try testing.expectError(error.LimitExceeded, h.evalOpts(nested.items, deep));
    // …and the same text evaluates once the evaluator is allowed the depth.
    deep.max_expr_depth = 4096;
    try testing.expectEqual(@as(f64, 1), (try h.evalOpts(nested.items, deep)).scalar.number);
}

test "limits: a matrix beyond §9's cap is a limit, not an error value" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    var da = h.options();
    da.dialect = .dynamic_array;
    // Two whole columns broadcast to 2 097 152 cells, which fits; five
    // do not. There is no error value meaning "too large to exist".
    try testing.expectError(error.LimitExceeded, h.evalOpts("A:E*B:F", da));
}

// ─── oracle ties (tests/oracle/fixtures) ─────────────────────────

const oracle_excel = @embedFile("oracle_hand_spec_excel");
const oracle_ieee = @embedFile("oracle_hand_spec_ieee");
const oracle_libreoffice = @embedFile("oracle_libreoffice_suite");

/// Rows where the committed *adapter* disagrees with Excel, named rather
/// than silently skipped. Both are excel-fidelity only.
///
///   * `-0` — §5.4 normalizes the sign at publication; LibreOffice
///     records the negative zero. M3a1 already listed this one.
///   * `SQRT(-1)` — Excel answers `#NUM!` (and the hand-spec manifest
///     records exactly that); LibreOffice answers `#VALUE!`. The two
///     excel-fidelity manifests disagree with each other, which is
///     itself the proof that this is the adapter and not the rule.
const excel_adapter_divergences = [_][]const u8{ "-0", "SQRT(-1)" };

fn isAdapterDivergent(adapter: []const u8, formula: []const u8) bool {
    if (!std.mem.eql(u8, adapter, "libreoffice")) {
        // The hand-spec manifests ARE the Excel statement; only a real
        // adapter can diverge from it.
        return std.mem.eql(u8, formula, "-0");
    }
    for (excel_adapter_divergences) |f| {
        if (std.mem.eql(u8, f, formula)) return true;
    }
    return false;
}

const Tie = struct { checked: usize, skipped_adapter: usize, skipped_excluded: usize };

fn recordedValue(obj: std.json.ObjectMap) !?value.ScalarValue {
    const kind = (obj.get("kind") orelse return null).string;
    if (std.mem.eql(u8, kind, "number")) {
        const bits_text = (obj.get("bits") orelse return null).string;
        const raw = try std.fmt.parseInt(u64, bits_text[2..], 16);
        return value.ScalarValue.fromNumber(@bitCast(raw));
    }
    if (std.mem.eql(u8, kind, "text")) return .{ .text = (obj.get("text") orelse return null).string };
    if (std.mem.eql(u8, kind, "boolean")) return .{ .boolean = (obj.get("boolean") orelse return null).bool };
    if (std.mem.eql(u8, kind, "error")) {
        const spelling = (obj.get("error_spelling") orelse return null).string;
        const known = value.KnownError.fromSpelling(spelling) orelse return error.UnknownErrorSpelling;
        return value.ScalarValue.errorOf(known);
    }
    return null; // blank
}

fn tieOracleManifest(json: []const u8) !Tie {
    const doc = try std.json.parseFromSlice(std.json.Value, testing.allocator, json, .{});
    defer doc.deinit();

    const fidelity_text = doc.value.object.get("fidelity").?.string;
    const exact = std.mem.eql(u8, fidelity_text, "ieee");
    const adapter = doc.value.object.get("provenance").?.object.get("adapter").?.string;
    const cells = doc.value.object.get("cells").?.array.items;

    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    var fake = env.Fake.init(testing.allocator);
    defer fake.deinit();

    // Seed the environment with every recorded value, so a formula that
    // depends on another cell reads what the oracle recorded there —
    // that is what makes `C1+A2` and `C2*2` real dependency rows rather
    // than arithmetic on constants.
    var sheet_names: [8][]const u8 = undefined;
    var sheet_count: usize = 0;
    for (cells) |cell| {
        const obj = cell.object;
        const sheet_name = obj.get("sheet").?.string;
        var idx: ?env.SheetIndex = null;
        for (sheet_names[0..sheet_count], 0..) |n, i| {
            if (std.mem.eql(u8, n, sheet_name)) idx = env.SheetIndex.fromInt(@intCast(i));
        }
        if (idx == null) {
            sheet_names[sheet_count] = sheet_name;
            sheet_count += 1;
            idx = try fake.addSheet(sheet_name);
        }
        const v = (try recordedValue(obj)) orelse continue;
        try fake.putA1(idx.?, .stored, obj.get("ref").?.string, v);
    }

    var tie: Tie = .{ .checked = 0, .skipped_adapter = 0, .skipped_excluded = 0 };
    var draw_value: f64 = 0.5;
    var draws = DrawSource.constant(&draw_value);

    for (cells) |cell| {
        const obj = cell.object;
        const formula = (obj.get("formula") orelse continue).string;
        if (obj.get("excluded") != null) {
            // A volatile formula has no recorded value to tie to.
            tie.skipped_excluded += 1;
            continue;
        }
        if (!exact and isAdapterDivergent(adapter, formula)) {
            tie.skipped_adapter += 1;
            continue;
        }
        const expected = (try recordedValue(obj)) orelse continue;

        const sheet_name = obj.get("sheet").?.string;
        var sheet: env.SheetIndex = undefined;
        for (sheet_names[0..sheet_count], 0..) |n, i| {
            if (std.mem.eql(u8, n, sheet_name)) sheet = env.SheetIndex.fromInt(@intCast(i));
        }

        var parsed = try parser.parse(testing.allocator, formula, .{});
        defer parsed.deinit(testing.allocator);
        if (parsed == .refused) {
            std.debug.print("oracle: `{s}` did not parse\n", .{formula});
            return error.OracleParseRefused;
        }

        const fidelity: value.Fidelity = if (exact) .ieee else .excel;
        var ev = Evaluator.init(arena_state.allocator(), fake.evalEnv(), .{
            .current_sheet = sheet,
            .collation = .{ .fold = shippedFold },
            .draws = &draws,
            .fidelity = fidelity,
        });
        defer ev.deinit();
        const got = ev.evaluate(parsed.ok) catch |e| {
            std.debug.print("oracle: `{s}` refused with {t}\n", .{ formula, e });
            return e;
        };
        if (got != .scalar) {
            std.debug.print("oracle: `{s}` produced a non-scalar\n", .{formula});
            return error.OracleShapeMismatch;
        }

        // Publication is where the modes are allowed to differ, so the
        // comparison happens on the published value rather than on an
        // internal one no boundary would ever hand out.
        const published = value.publish(got.scalar, fidelity);
        const want = value.publish(expected, fidelity);

        const agrees = switch (want) {
            // `ieee` manifests record raw arithmetic and pin bits;
            // `excel` manifests carry §5.4 display-rounded values and pin
            // to 15 significant digits (M3a1 decision 9).
            .number => |w| published == .number and (if (exact)
                @as(u64, @bitCast(w)) == @as(u64, @bitCast(published.number))
            else
                @as(u64, @bitCast(w)) == @as(u64, @bitCast(published.number)) or
                    @abs(published.number - w) <= 1e-15 * @abs(w)),
            else => value.PublishedScalar.eql(want, published),
        };
        if (!agrees) {
            std.debug.print("oracle tie failed for `{s}` ({s})\n", .{ formula, fidelity_text });
            return error.OracleTieMismatch;
        }
        tie.checked += 1;
    }
    return tie;
}

test "oracle: the evaluator reproduces every manifest cell it can decide" {
    const ieee = try tieOracleManifest(oracle_ieee);
    const excel = try tieOracleManifest(oracle_excel);
    const lo = try tieOracleManifest(oracle_libreoffice);

    // Exact counts, not lower bounds: a row that silently stopped being
    // checked is exactly the failure a `>=` would hide.
    try testing.expectEqual(@as(usize, 7), ieee.checked);
    try testing.expectEqual(@as(usize, 0), ieee.skipped_adapter);
    try testing.expectEqual(@as(usize, 11), excel.checked);
    try testing.expectEqual(@as(usize, 0), excel.skipped_adapter);
    try testing.expectEqual(@as(usize, 20), lo.checked);
    // `-0` and `SQRT(-1)`, both named above.
    try testing.expectEqual(@as(usize, 2), lo.skipped_adapter);
    // `RAND()`, which the extractor marked volatile.
    try testing.expectEqual(@as(usize, 1), lo.skipped_excluded);
}

test "oracle: the dependency rows evaluate through the environment" {
    // The LibreOffice suite's `A1*10`, `C1+A2`, and `C2*2` are the only
    // committed rows that read another cell, which makes them the tie
    // that proves the evaluator is reading through `EvalEnv` at all.
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try h.put("A1", num(1));
    try h.put("A2", num(2));
    try h.put("C1", num(10));
    try h.put("C2", num(12));

    try testing.expectEqual(@as(f64, 10), (try h.scalar("A1*10")).number);
    try testing.expectEqual(@as(f64, 12), (try h.scalar("C1+A2")).number);
    try testing.expectEqual(@as(f64, 24), (try h.scalar("C2*2")).number);
}

// ─── allocation failure ──────────────────────────────────────────

fn evalUnderOom(allocator: std.mem.Allocator, src: []const u8) !void {
    var arena_state = std.heap.ArenaAllocator.init(allocator);
    defer arena_state.deinit();
    var fake = env.Fake.init(allocator);
    defer fake.deinit();
    const sheet = try fake.addSheet("S");
    try fake.putA1(sheet, .stored, "A1", num(2));
    try fake.putA1(sheet, .stored, "A2", .{ .text = "abc" });
    try fake.putA1(sheet, .stored, "B1", num(3));

    var parsed = try parser.parse(allocator, src, .{});
    defer parsed.deinit(allocator);
    if (parsed == .refused) return error.ParseRefused;

    var draw_value: f64 = 0.25;
    var draws = DrawSource.constant(&draw_value);
    var ev = Evaluator.init(arena_state.allocator(), fake.evalEnv(), .{
        .current_sheet = sheet,
        .collation = .{ .fold = shippedFold },
        .draws = &draws,
    });
    defer ev.deinit();
    _ = ev.evaluate(parsed.ok) catch |e| switch (e) {
        error.OutOfMemory => return error.OutOfMemory,
        else => return,
    };

    var static = DependencyLog.init(arena_state.allocator());
    try staticDependencies(arena_state.allocator(), parsed.ok, sheet, fake.evalEnv(), &static);
}

test "checkAllAllocationFailures: evaluation leaks nothing under OOM" {
    const sources = [_][]const u8{
        "A1+B1",
        "A1:B1*2",
        "IF(A1>1,A1&\"x\",RAND())",
        "SUM(A1:B1,TRUE(),\"9\")",
        "{1,2;3,4}+{10;20}",
        "A2<\"ABC\"",
        "IFERROR(1/0,A1&A2)",
        "COUNTBLANK(A1:B9)",
    };
    for (sources) |src| {
        try testing.checkAllAllocationFailures(testing.allocator, evalUnderOom, .{src});
    }
}

// ─── fuzz (§8.1: eval no-panic / no-leak / non-finite) ───────────

fn assertRepresentable(v: Value) !void {
    switch (v) {
        .scalar => |s| {
            if (s == .number) try std.testing.expect(std.math.isFinite(s.number));
        },
        .missing_arg => {},
        .array => |m| {
            // A zero-dimension array is not a representable result; the
            // producing function normalizes it to `#CALC!` instead.
            try std.testing.expect(m.rows > 0 and m.cols > 0);
            try std.testing.expect(m.cells.len == @as(usize, m.rows) * m.cols);
            for (m.cells) |s| {
                if (s == .number) try std.testing.expect(std.math.isFinite(s.number));
            }
        },
        .reference => |r| try std.testing.expect(r.areas.len > 0),
    }
}

fn fuzzEvalTarget(_: void, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();
    var smith_buf: [256]u8 = undefined;
    const input = smith_buf[0..smith.slice(&smith_buf)];

    var parsed = parser.parse(std.testing.allocator, input, .{}) catch return;
    defer parsed.deinit(std.testing.allocator);
    if (parsed == .refused) return;

    var arena_state = std.heap.ArenaAllocator.init(std.testing.allocator);
    defer arena_state.deinit();
    var fake = env.Fake.init(std.testing.allocator);
    defer fake.deinit();
    const sheet = try fake.addSheet("S");
    try fake.putA1(sheet, .stored, "A1", num(2));
    try fake.putA1(sheet, .stored, "A2", .{ .text = "1e308" });
    try fake.putA1(sheet, .stored, "A3", value.ScalarValue.errorOf(.na));
    try fake.putA1(sheet, .stored, "B1", num(-1));

    var draw_value: f64 = 0.5;
    var draws = DrawSource.constant(&draw_value);

    inline for ([_]value.Fidelity{ .excel, .ieee }) |fid| {
        inline for ([_]value.Dialect{ .dynamic_array, .legacy }) |dia| {
            var ev = Evaluator.init(arena_state.allocator(), fake.evalEnv(), .{
                .current_sheet = sheet,
                .collation = .{ .fold = shippedFold },
                .draws = &draws,
                .fidelity = fid,
                .dialect = dia,
                .site = .{
                    .row = coords.Row.fromOneBased(2) catch unreachable,
                    .col = coords.Col.fromZeroBased(1) catch unreachable,
                },
            });
            defer ev.deinit();
            if (ev.evaluate(parsed.ok)) |v| {
                try assertRepresentable(v);
            } else |_| {
                // A typed refusal is a legitimate outcome; a panic is not.
            }
        }
    }
}

test "fuzz: no evaluation escapes with a non-finite number or an empty matrix" {
    try std.testing.fuzz({}, fuzzEvalTarget, .{
        .corpus = &[_][]const u8{
            "1+1",
            "1E+308*10",
            "1E+308+1E+308",
            "-1E+308-1E+308",
            "1/0",
            "0/0",
            "SQRT(-1)",
            "2^1024",
            "(-8)^(1/3)",
            "A1*A2",
            "A2+1",
            "A2*A2",
            "{1e308,1e308}*{10,10}",
            "IF({TRUE;FALSE},A1,B1)",
            "IFERROR(1/0,RAND())",
            "SUM(A1:B9)",
            "CHOOSE(2,A1,B1)",
            "@A1:A3",
            "A1:B2 B1:B9",
            "\"a\"<\"B\"",
            "10%%%",
            "SQRT({4,9})",
        },
    });
}
