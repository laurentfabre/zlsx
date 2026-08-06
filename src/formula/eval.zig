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
const run_inputs = @import("run_inputs.zig");
/// §5.6d's draw schedule. A leaf below this file and below the iteration
/// engine both, because a draw's key is half in-body (which callsite,
/// which element — this file's) and half run-level (whose body, which
/// pass — the engine's).
const draw_schedule = @import("draws.zig");
/// Test-only, for M4f: the ST_Xstring codec, so a round-trip can start
/// from text a formula PRODUCED rather than from a hand-written string.
/// A file-scope const referenced only from a `test` block is not
/// resolved in a non-test build, so this costs a plain build nothing.
const decode = @import("decode.zig");
/// §5.9's resolution drivers and §5.6g's 3D matrix. The evaluator asks
/// this file *what the rule is* and never restates one; the symbol
/// layer that implements the tiers imports it too, which is why the
/// rules live below both rather than inside either.
const name_rules = @import("names.zig");

// ─── evaluator-layer values (§5.3a) ──────────────────────────────

/// One or more rectangular areas. Multi-area sets come from the union
/// operator (`(A1:A2,C1:C2)`); §10 refuses one as a top-level *result*
/// while functions that take them keep working.
pub const Reference = struct {
    areas: []const env.RangeRef,
    /// Whether the areas are the members of a 3D sheet span (§5.6g)
    /// rather than an authored union.
    ///
    /// The two are the same shape and not the same thing. `A1:B1` after
    /// a union takes the bounding box of two areas; after a *span* it
    /// takes one box per member sheet, because the span's members are
    /// one reference repeated across sheets and not several references
    /// side by side. Nothing else in the evaluator branches on it —
    /// aggregation over N areas is aggregation over N areas.
    three_d: bool = false,

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
    /// A construct v1 refuses on purpose rather than has not reached: a
    /// 3D span outside the frozen eligible list or inside an array or
    /// intersection context (§5.6g), a name the `CT_DefinedName`
    /// inventory or its own body disqualifies (§5.9), a table
    /// reference (M7b evaluates one; §5.9's order only has to reach
    /// it). Distinct from `NotYetImplemented` because the two answer
    /// different questions — "never in v1" versus "not at this row" —
    /// and only the second has an enumerated membership a later row
    /// deletes from.
    UnsupportedConstruct,
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
        error.NotYetImplemented, error.UnsupportedConstruct => .FormulaUnsupportedConstruct,
        error.AnchorRequired => .FormulaAnchorRequired,
        error.ResultNotRepresentable => .FormulaResultNotRepresentable,
        error.LimitExceeded => .FormulaLimitExceeded,
        // The environment's failures are invariant violations by the
        // time they reach here: the evaluator only ever builds a
        // `SheetIndex` from `resolveSheet`, and it converts shape and
        // grid failures into `#VALUE!` at the call that caused them.
        error.MalformedInput, error.UnknownSheet, error.ShapeMismatch, error.RefOutOfGrid => .FormulaMalformedInput,
        // M4b3, and the same split M4a's `MetadataRefused` uses: the
        // interface can only say the resolution refused, and the typed
        // reason travels with the resolver that raised it
        // (`symbols.NameResolution.last_refusal`). Every reason it can
        // carry is an unsupported construct, so nothing is lost here.
        error.NameRefused => .FormulaUnsupportedConstruct,
        // M4a. The precise plane — unsupported construct for a rich
        // value, malformed input for a broken part — travels with the
        // resolver's own `Refusal` (`metadata.CellDialectResolver.
        // last_refusal`), which is what the report carries. Here the
        // detail is already gone, so this maps to the class that ALWAYS
        // refuses (§5.7.7): collapsing onto a mark-eligible class could
        // let `.keep_stale_and_mark` suppress a refusal the taxonomy
        // says must stand.
        error.MetadataRefused => .FormulaMalformedInput,
    };
}

/// Constructs this row parses and refuses rather than evaluating. Named
/// as data so the later row that implements one deletes a line here and
/// watches a test fail until it does.
///
/// **Empty as of M4b3**, which deleted its one entry — 3D sheet spans —
/// and now evaluates them (§5.6g). The array stays, and so does the
/// distinction it exists for: `UnsupportedConstruct` is "never in v1"
/// and this list is "not at this row yet". A construct that lands here
/// again has a milestone attached to it, in writing.
pub const not_yet_implemented = [_][]const u8{};

// ─── run instruments ─────────────────────────────────────────────

/// The single seam every volatile draw passes through.
///
/// The counter lives in the seam rather than in a test double on
/// purpose: "zero draws in the dead branch" is then a statement about
/// the evaluator, not about how carefully a fixture was wired. M3b
/// replaced the callback with `rng_v1` seeded from `RunInputs`; the
/// counter and its meaning did not change.
///
/// M5a2 adds §5.6d's **schedule**. Without one this is what it always
/// was — every reach draws. With one, a draw is first a *lookup*: the
/// evaluator maintains `key`'s in-body terms (which callsite, which
/// element) and whoever runs the evaluator maintains the run's terms
/// (whose body, which pass), so a second reach at the same key returns
/// the same number. That is the property §5.6e needs and cannot get any
/// other way: a graph rebuild re-walks bodies that already drew, and if
/// they drew again a discovery pass would change an answer.
pub const DrawSource = struct {
    ctx: *anyopaque,
    draw_fn: *const fn (ctx: *anyopaque) f64,
    count: u64 = 0,
    /// §5.6d's memo, when the caller wired one. Null is not a degraded
    /// mode: a single non-iterating evaluation reaches every callsite
    /// once, so there is nothing for a memo to decide.
    schedule: ?*draw_schedule.Schedule = null,
    /// Where the memo lives. Required exactly when `schedule` is set.
    gpa: ?std.mem.Allocator = null,
    /// The key the evaluator is currently at. `path`, `component` and
    /// `pass` belong to whoever drives the evaluator; `callsite` and
    /// `element` are maintained here.
    key: draw_schedule.Key = .{},

    /// A draw can now fail, because a memo has to be stored somewhere.
    /// The alternative — swallowing the allocation failure and drawing
    /// afresh — would make an out-of-memory condition silently change a
    /// result, which is exactly the class of bug the memo exists to
    /// prevent.
    pub fn draw(self: *DrawSource) error{OutOfMemory}!f64 {
        self.count += 1;
        if (self.schedule) |s| {
            return s.valueFor(self.gpa.?, self.key, self.ctx, self.draw_fn);
        }
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
///
/// Declared with the rest of a run's inputs, and re-exported here
/// because every caller of the evaluator already has this namespace
/// open.
pub const EvalSite = run_inputs.EvalSite;

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
    /// §5.4d's compatibility version, and therefore what a *character*
    /// is to `LEN`/`MID`/`FIND`/`SEARCH`/`REPLACE`. Workbook-derived:
    /// recalc passes what `CalcState` parsed, and the default is CV1
    /// because a workbook with no compatibility metadata IS CV1.
    text_compat: run_inputs.CompatibilityVersion = .cv1,
    /// §5.4a's epoch, and therefore what every serial in the workbook
    /// MEANS. Workbook-derived (`workbookPr@date1904`), never
    /// caller-set: a caller who could change it would silently redate
    /// every cell in the file by four years and a day.
    date_system: run_inputs.DateSystem = .d1900,
    /// The instant `NOW()` and `TODAY()` report, as Unix milliseconds.
    /// It is an INPUT rather than a clock read, which is what makes a
    /// recalc reproducible from `RunInputs` alone and what lets every
    /// volatile-date fixture pin an exact answer.
    now_utc_ms: i64 = 0,
    /// The fixed civil offset those two apply. Zero — UTC — at every
    /// layer, because the stdlib has no portable local-timezone resolver
    /// and zlsx is stdlib-only; a caller that wants local time passes
    /// the offset. TZif is M10+.
    utc_offset_min: i32 = 0,
    /// §5.4b's code page, which `CHAR` and `CODE` resolve through. The
    /// v1 enum is closed, so this is a seam rather than a choice — but
    /// it is the seam a second profile lands in, and it belongs to the
    /// run rather than to the workbook.
    platform_profile: run_inputs.PlatformProfile = .windows_1252,
    /// §5.9 name resolution. Optional, and null is not a degraded mode
    /// but a *stated* one: with no symbol layer every spelling provably
    /// resolves nowhere, which is the terminal stage of the
    /// value-position order and exactly what M3a2 shipped. Standalone
    /// eval against a workbook wires `symbols.NameResolution`.
    names: ?env.NameResolver = null,
    /// Whether the formula is a declared array formula — legacy CSE
    /// (`<f t="array">`) or dynamic-array. §5.6g forbids a 3D span
    /// inside one, and the refusal is pre-eval, so the flag has to be
    /// *declared* rather than discovered from the result's shape.
    array_formula: bool = false,
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
    /// §5.6d's invocation-path root — which owner's body this is.
    ///
    /// A recalc passes `draw_schedule.Key.ofCell(…)` for each stored cell it
    /// evaluates; a standalone `evaluate` leaves the constant root,
    /// because a formula with no cell has no coordinate to be keyed by
    /// and inventing one would collide with whatever really lives there.
    draw_path: u64 = draw_schedule.Key.root,
    /// The §9 byte budget, when the caller wired one — normally by
    /// building the run arena over `Budget.allocator(.run_arena)`.
    ///
    /// It is optional because a budget is a *policy*, and the evaluator
    /// is usable without one; when present, an exhausted category turns
    /// what the allocator could only call `OutOfMemory` into the typed
    /// `FormulaLimitExceeded` the caller asked for.
    budget: ?*run_inputs.Budget = null,
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
    /// How many defined-name bodies are open on the stack. The interim
    /// guard §5.9 asks for (M5a replaces it with graph nodes, where a
    /// cycle through two names is a cycle like any other); until then
    /// `A = B`, `B = A` has to stop somewhere with a §9 limit rather
    /// than a stack overflow.
    name_depth: usize = 0,
    /// Which §5.6g rule an `error.UnsupportedConstruct` came from, when
    /// it came from the 3D matrix. Same split as the two resolvers: the
    /// error type says a construct was refused, the typed reason stays
    /// with whatever raised it, and a report reads it from here.
    last_three_d: ?name_rules.Refusal = null,

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
        self.name_depth = 0;
        // The run-level terms of §5.6d's key (`component`, `pass`)
        // belong to whoever drives this evaluator and are left alone;
        // the body-level ones start where this body starts.
        self.opts.draws.key.path = self.opts.draw_path;
        self.opts.draws.key.callsite = 0;
        self.opts.draws.key.element = 0;
        // §5.6g's context legality, **before** anything is evaluated —
        // which is what makes the same check usable pre-persist, where
        // there is no result yet to inspect.
        if (name_rules.checkThreeD(ast, .{ .array_formula = self.opts.array_formula })) |r| {
            self.last_three_d = r;
            return error.UnsupportedConstruct;
        }
        const v = self.evalNode(ast.root) catch |e| return self.mapBudget(e);
        // §10: a multi-area reference is not a representable result,
        // even though functions may take one as an argument.
        if (v == .reference and v.reference.areas.len != 1) {
            return error.ResultNotRepresentable;
        }
        // A formula's *result* is a value. A reference reaching the top
        // dereferences through the same §5.3b row as a reference reaching
        // an operator — which is why `=A1:A3` spills under `dynamic_array`
        // and intersects under `legacy` without either being special-cased.
        if (v == .reference) {
            const op = self.operandOf(v) catch |e| return self.mapBudget(e);
            return operandValue(op);
        }
        return v;
    }

    /// A budget that ran out is a **refusal**, not an allocation
    /// failure. `std.mem.Allocator` has only one way to say no, so the
    /// budget records which category tripped and the distinction is
    /// recovered here — otherwise a caller could not tell "your formula
    /// asked for too much" from "this machine is out of memory", and
    /// only one of those is actionable.
    fn mapBudget(self: *Evaluator, e: EvalError) EvalError {
        if (e != error.OutOfMemory) return e;
        const b = self.opts.budget orelse return e;
        if (b.tripped == null) return e;
        return error.LimitExceeded;
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
            // §5.9: with no symbol layer wired, every name provably
            // resolves nowhere and the terminal stage of the
            // value-position order applies. M4b3 inserted the earlier
            // stages ahead of it and did not change the fallthrough.
            //
            // A structured reference still lands there whatever the
            // symbol layer says: `Table[Col]` needs the table's
            // geometry, which is M7b's, and answering `#NAME?` is the
            // answer M3a2 gave — not a new claim about tables.
            // The node index travels with the spelling: §5.6d's
            // invocation path descends by the *occurrence* of the
            // reference that expanded, which is what makes the two `N`s
            // of `A1=N+N` two paths rather than one.
            .name => |n| try self.namedValue(n.raw, i),
            .structured => Value.err(.name),
            .qualified => |n| try self.qualified(n.sheet, n.target),
            .call => |n| try self.call(n.callee, n.args, i),
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
        const buf = try self.allocText(body.len);
        var n: usize = 0;
        var k: usize = 0;
        while (k < body.len) : (k += 1) {
            buf[n] = body[k];
            n += 1;
            if (body[k] == '"') k += 1; // skip the second of a `""` pair
        }
        return buf[0..n];
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

    /// Text the run produced, charged to §9's `string payload` budget.
    /// Borrowed text — a string literal pointing into the source, a
    /// cell's stored bytes — costs nothing and is not charged, because
    /// the run did not create it.
    fn allocText(self: *Evaluator, n: usize) EvalError![]u8 {
        if (self.opts.budget) |b| {
            b.charge(.string_payload, n) catch return error.LimitExceeded;
        }
        return self.arena.alloc(u8, n);
    }

    fn newMatrix(self: *Evaluator, rows: u32, cols: u32) EvalError!value.Matrix {
        if (self.opts.budget) |b| {
            // A cell count, not a byte count: §9 bounds `live matrix
            // cells` so the limit means the same thing whatever
            // `@sizeOf(ScalarValue)` happens to be. The run arena frees
            // nothing individually, so "live" is "allocated this run".
            b.charge(.matrix_cells, @as(u64, rows) * @as(u64, cols)) catch
                return error.LimitExceeded;
        }
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
        // null `last` is not proof of a single sheet — `isSpan` knows
        // both spellings, and a sheet name cannot contain a colon.
        const name = try self.unquoteSheetName(spec);
        if (name_rules.isSpan(spec)) return self.threeDReference(spec, name, target);
        const idx = (try self.environment.resolveSheet(name)) orelse
            return Value.err(.ref); // a deleted or unknown sheet
        const saved = self.sheet;
        self.sheet = idx;
        defer self.sheet = saved;
        return self.evalNode(target);
    }

    /// §5.6g: one reference over an inclusive span of sheets, in
    /// workbook order, as one area per member.
    ///
    /// Eligibility and context legality were settled before the walk
    /// began (`evaluate`), so what is left here is arithmetic: resolve
    /// the two endpoints, expand between them, and evaluate the target
    /// once per member sheet. A multi-area reference is a shape the
    /// aggregates already consume — the union operator produces one —
    /// so the six eligible functions need no 3D-specific code.
    fn threeDReference(
        self: *Evaluator,
        spec: parser.SheetSpec,
        unquoted: []const u8,
        target: parser.Index,
    ) EvalError!Value {
        const ends = name_rules.splitSpan(spec, unquoted) orelse
            return Value.err(.ref);
        const first = try self.environment.resolveSheet(ends.first);
        const last = try self.environment.resolveSheet(ends.last);
        const span = name_rules.expandSpan(
            if (first) |f| f.toInt() else null,
            if (last) |l| l.toInt() else null,
        );
        const members = switch (span) {
            // A deleted endpoint, or two that have swapped. Excel
            // leaves the spelling in place rather than repairing the
            // span, and every cell it reached reads `#REF!`.
            .ref_error => return Value.err(.ref),
            .members => |m| m,
        };

        var areas: std.ArrayListUnmanaged(env.RangeRef) = .empty;
        try areas.ensureTotalCapacity(self.arena, members.last - members.first + 1);
        var s = members.first;
        while (s <= members.last) : (s += 1) {
            const saved = self.sheet;
            self.sheet = env.SheetIndex.fromInt(s);
            const r = self.evalAsReference(target) catch |e| {
                self.sheet = saved;
                return e;
            };
            self.sheet = saved;
            // The target of a 3D qualifier is a reference by grammar;
            // anything else is a tree no parser built.
            const ref = r orelse return Value.err(.value);
            try areas.appendSlice(self.arena, ref.areas);
        }
        return .{ .reference = .{ .areas = areas.items, .three_d = true } };
    }

    // ─── names (§5.9) ────────────────────────────────────────────

    /// A spelling in value position.
    ///
    /// The order it resolves in is §5.9's, walked by the symbol layer
    /// over M2's exported array; what arrives here is one of three
    /// answers. A body is expanded inline — the interim shape, guarded
    /// by depth, until M5a makes bodies graph nodes.
    fn namedValue(self: *Evaluator, spelling: []const u8, at: parser.Index) EvalError!Value {
        const resolver = self.opts.names orelse return Value.err(.name);
        const binding = try resolver.resolveName(self.sheet, spelling);
        return switch (binding) {
            .not_found => Value.err(.name),
            // §5.9's order reaches the table tier so a table can shadow
            // an `_xlnm.` builtin; evaluating one is M7b's.
            .table => error.UnsupportedConstruct,
            .body => |b| self.expandName(b.text, b.scope, at),
        };
    }

    fn expandName(
        self: *Evaluator,
        body: []const u8,
        scope: ?env.SheetIndex,
        at: parser.Index,
    ) EvalError!Value {
        if (self.name_depth >= name_rules.max_name_expansion_depth) {
            return error.LimitExceeded;
        }
        self.name_depth += 1;
        defer self.name_depth -= 1;

        // §5.6d: expansion descends the invocation path by the
        // occurrence that expanded. The row is 0 — a defined name
        // materializes nowhere, unlike a table producer, which is M7b's
        // and is the only thing that will ever pass a non-zero row.
        const outer_path = self.opts.draws.key.path;
        self.opts.draws.key.path = draw_schedule.Key.descend(outer_path, at, 0);
        defer self.opts.draws.key.path = outer_path;

        // Parsed into the run arena: a name body is a formula, and the
        // AST it produces lives exactly as long as the run does.
        const parsed = try parser.parse(self.arena, body, .{ .limits = self.opts.limits });
        const ast = switch (parsed) {
            .ok => |a| a,
            // A body that will not parse is a statement about the
            // workbook, not about this evaluation.
            .refused => return error.MalformedInput,
        };
        // §5.6g holds inside a name body too — expansion must not be a
        // way to smuggle a 3D span past the pre-eval check.
        if (name_rules.checkThreeD(ast, .{ .array_formula = self.opts.array_formula })) |r| {
            self.last_three_d = r;
            return error.UnsupportedConstruct;
        }

        const saved_ast = self.ast;
        const saved_sheet = self.sheet;
        self.ast = ast;
        // A sheet-scoped name resolves its unqualified halves against
        // its own sheet; a workbook-scoped one has no sheet of its own,
        // so the referencing sheet stands. Relative bodies refused
        // before they got here, which is what keeps this from being a
        // guess about where the name was authored.
        if (scope) |s| self.sheet = s;
        defer {
            self.ast = saved_ast;
            self.sheet = saved_sheet;
        }
        return self.evalNode(ast.root);
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
                // §5.6g: `Sheet1:Sheet3!A1:B1` parses as the span's
                // `A1` ranged with a bare `B1`, so the span's members
                // are on the left and the second endpoint is on the
                // right. It is one box per member sheet — the span
                // repeats one reference across sheets, so the endpoint
                // repeats with it.
                if (l.three_d and !r.three_d) {
                    const b = r.single() orelse return Value.err(.value);
                    const areas = try self.arena.alloc(env.RangeRef, l.areas.len);
                    for (l.areas, areas) |member, *out| {
                        out.* = .{
                            .sheet = member.sheet,
                            .range = boundingBox(member.range, b.range),
                        };
                        try self.deps.noteArea(out.*);
                    }
                    return .{ .reference = .{ .areas = areas, .three_d = true } };
                }
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
                return .{ .reference = .{
                    .areas = areas,
                    .three_d = l.three_d or r.three_d,
                } };
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
                const out = try self.allocText(x.len + y.len);
                @memcpy(out[0..x.len], x);
                @memcpy(out[x.len..], y);
                // §9's cell-text cap applies to text a FORMULA produced,
                // whichever construct produced it. Before M4f no formula
                // could reach it — a literal is bounded by the 8 192-byte
                // formula length — but `REPT` can, and `&` is how two of
                // them are put together. A cap `CONCAT` enforced and `&`
                // did not would be one rule with two answers.
                return registry.cappedText(out);
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

    /// Public since M4f: `CONCAT` and `TEXTJOIN` walk ranges themselves,
    /// so they need the same "what does this value look like as text"
    /// rule the dispatcher applies to a `.text` slot. A second copy in
    /// the registry would be a second answer to how `TRUE` prints.
    pub fn toText(self: *Evaluator, s: value.ScalarValue) EvalError![]const u8 {
        return switch (s) {
            .text => |t| t,
            .blank => "",
            .boolean => |b| if (b) "TRUE" else "FALSE",
            .number => |n| blk: {
                const buf = try self.allocText(value.format_buf_len);
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

    fn call(self: *Evaluator, callee: parser.Index, args: parser.ExtraSlice, site: parser.Index) EvalError!Value {
        const name_node = self.ast.node(callee);
        if (name_node != .name) return error.NotYetImplemented;
        const f = registry.lookup(name_node.name.bare) orelse return error.UnsupportedFunction;

        // §5.6d's callsite ordinal. Set around the whole call, so a
        // nested call restores it on the way out and the implementation
        // always draws under its own index — `RAND()+RAND()` is two
        // callsites, and `SUM(RAND(),RAND())` is two more.
        const outer_site = self.opts.draws.key.callsite;
        const outer_element = self.opts.draws.key.element;
        self.opts.draws.key.callsite = site;
        self.opts.draws.key.element = 0;
        defer {
            self.opts.draws.key.callsite = outer_site;
            self.opts.draws.key.element = outer_element;
        }

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
                    // The mixed-signature answer (M7a): lift over the
                    // SCALAR slots that carry arrays, hold every
                    // non-scalar slot fixed — `VLOOKUP({1;2},T,2)` is
                    // two lookups down one table. `liftable()` is no
                    // longer a gate, only the statement that a function
                    // with nothing but scalar slots lifts wholesale.
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
                    // §5.6d's element ordinal. One call site, N results:
                    // without this every element of a lifted volatile
                    // would share a key and the array would be constant.
                    self.opts.draws.key.element = r * shape.cols + c;
                    for (ops, 0..) |maybe_op, k| {
                        const op = maybe_op orelse continue;
                        const el = op.at(r, c) orelse value.incompatibleBroadcastFill();
                        args[k] = .{ .scalar = try self.coerceSlot(f.coercion.at(k), el) };
                    }
                    const one = try self.propagateAndInvoke(f, args);
                    // §5.3b's nested-array rule: a per-element result
                    // that is itself an array reduces to its top-left —
                    // `SEQUENCE({1,2})` is `{1,1}` — and a per-element
                    // reference dereferences the same way.
                    m.set(r, c, switch (one) {
                        .scalar => |s| s,
                        .array => |a| a.topLeft(),
                        .reference => blk: {
                            const rop = try self.operandOf(one);
                            break :blk rop.at(0, 0).?;
                        },
                        .missing_arg => unreachable,
                    });
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
            // A criterion is classified by `criteria.parse`, not coerced:
            // turning `">5"` into a number here would throw the operator
            // away before anything had read it.
            .criteria => s,
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

    /// A reference an implementation computed rather than a walk found.
    ///
    /// The **only** way `INDIRECT` and `OFFSET` produce one, and public
    /// for exactly that reason: routing them through `refValue` is what
    /// makes a dynamic reference note its dependency the same way a
    /// written one does, which is in turn what lets §5.6e see it at all.
    /// A second construction path would be a reference nothing captured.
    pub fn computedReference(self: *Evaluator, area: env.RangeRef) EvalError!Value {
        return self.refValue(area);
    }

    /// §7's `INDIRECT`: the area a piece of *text* denotes.
    ///
    /// Everything it can refuse is a value (`#REF!`) except one thing
    /// that is a construct: R1C1. The tokenizer refuses `R1C1` wherever
    /// it appears in a formula (v1 is A1-only), and `INDIRECT(t,FALSE)`
    /// is a request for the same construct by another spelling — so it
    /// gets the same answer rather than a `#REF!` that would imply the
    /// text was malformed. It was not; it was R1C1.
    pub fn referenceFromText(self: *Evaluator, text: []const u8) EvalError!Value {
        const split = splitSheetPrefix(text) orelse return Value.err(.ref);
        var sheet = self.sheet;
        if (split.sheet) |raw| {
            const name = unquoteSheetSpelling(raw) orelse return Value.err(.ref);
            sheet = (try self.environment.resolveSheet(name)) orelse return Value.err(.ref);
        }
        const range = parseA1Area(split.rest) orelse return Value.err(.ref);
        return self.refValue(.{ .sheet = sheet, .range = range.normalized() });
    }

    /// §7's `OFFSET`: an area displaced from another area, optionally
    /// resized. `rows`/`cols`/`height`/`width` have already been coerced
    /// through the numeric column and truncated toward zero by the
    /// caller, which is where Excel truncates them too.
    pub fn offsetReference(
        self: *Evaluator,
        base: env.RangeRef,
        row_delta: i64,
        col_delta: i64,
        height: i64,
        width: i64,
    ) EvalError!Value {
        // Microsoft documents height and width as positive numbers.
        // Excel 365 has since grown an undocumented negative-extent
        // behaviour that extends in the opposite direction; with the
        // Excel oracle leg parked there is no evidence for it here, and
        // inventing it would be a claim about Excel this repo cannot
        // back. The documented contract is what ships.
        if (height <= 0 or width <= 0) return Value.err(.value);

        const first_row = @as(i64, base.range.first.row.oneBased()) + row_delta;
        const first_col = @as(i64, base.range.first.col.zeroBased()) + col_delta;
        const last_row = first_row + height - 1;
        const last_col = first_col + width - 1;
        // "If rows and cols offset reference over the edge of the
        // worksheet, OFFSET returns #REF!" — and so does an extent that
        // runs off it.
        if (first_row < 1 or last_row > coords.max_row) return Value.err(.ref);
        if (first_col < 0 or last_col >= coords.max_col_1based) return Value.err(.ref);

        return self.refValue(.{
            .sheet = base.sheet,
            .range = .{
                .first = .{
                    .row = coords.Row.fromOneBased(@intCast(first_row)) catch unreachable,
                    .col = coords.Col.fromZeroBased(@intCast(first_col)) catch unreachable,
                },
                .last = .{
                    .row = coords.Row.fromOneBased(@intCast(last_row)) catch unreachable,
                    .col = coords.Col.fromZeroBased(@intCast(last_col)) catch unreachable,
                },
            },
        });
    }
};

const SheetSplit = struct { sheet: ?[]const u8, rest: []const u8 };

/// Split `Sheet1!A1` / `'My Sheet'!A1:B2` / `A1`.
///
/// Null means the text cannot be a reference at all — an unterminated
/// quote, or a `!` inside a bare (unquoted) name, which no sheet name
/// may contain.
fn splitSheetPrefix(text: []const u8) ?SheetSplit {
    if (text.len == 0) return null;
    if (text[0] == '\'') {
        var i: usize = 1;
        while (i < text.len) : (i += 1) {
            if (text[i] != '\'') continue;
            // `''` is one literal quote inside the name, not the end.
            if (i + 1 < text.len and text[i + 1] == '\'') {
                i += 1;
                continue;
            }
            if (i + 1 >= text.len or text[i + 1] != '!') return null;
            return .{ .sheet = text[0 .. i + 1], .rest = text[i + 2 ..] };
        }
        return null; // unterminated
    }
    const bang = std.mem.indexOfScalar(u8, text, '!') orelse
        return .{ .sheet = null, .rest = text };
    return .{ .sheet = text[0..bang], .rest = text[bang + 1 ..] };
}

/// The sheet name a (possibly quoted) spelling denotes. Returns a
/// borrowed slice; `''` inside a quoted name is the one escape, and a
/// name containing one is refused rather than unescaped into an arena
/// this function does not own — Excel sheet names may contain `'`, but
/// only as a doubled literal, and `INDIRECT` over one is rare enough
/// that a `#REF!` is a better answer than an allocation on this path.
fn unquoteSheetSpelling(raw: []const u8) ?[]const u8 {
    if (raw.len == 0) return null;
    if (raw[0] != '\'') return if (std.mem.indexOfScalar(u8, raw, '\'') == null) raw else null;
    if (raw.len < 2 or raw[raw.len - 1] != '\'') return null;
    const inner = raw[1 .. raw.len - 1];
    if (inner.len == 0) return null;
    if (std.mem.indexOfScalar(u8, inner, '\'') != null) return null;
    return inner;
}

/// The area an A1 spelling denotes: a cell, a rectangle, a whole column
/// span or a whole row span. Null for anything else.
fn parseA1Area(s: []const u8) ?coords.Range {
    if (s.len == 0) return null;
    if (coords.parseRange(s, .{ .dollar = .accept })) |r| {
        return r;
    } else |_| {}

    // `A:A` and `1:1`. `parseRange` cannot take them — neither half is
    // a cell — and they are the two spellings `INDIRECT` is most often
    // handed after a rectangle.
    const colon = std.mem.indexOfScalar(u8, s, ':') orelse return null;
    const left = stripDollar(s[0..colon]);
    const right = stripDollar(s[colon + 1 ..]);
    if (left.len == 0 or right.len == 0) return null;

    if (std.ascii.isDigit(left[0]) and std.ascii.isDigit(right[0])) {
        const lo = std.fmt.parseInt(u32, left, 10) catch return null;
        const hi = std.fmt.parseInt(u32, right, 10) catch return null;
        if (lo == 0 or hi == 0 or lo > coords.max_row or hi > coords.max_row) return null;
        return fullRowRange(
            .{ .row = coords.Row.fromOneBased(lo) catch return null, .absolute = false },
            .{ .row = coords.Row.fromOneBased(hi) catch return null, .absolute = false },
        );
    }
    const lo = coords.parseCol(left, .{}) catch return null;
    const hi = coords.parseCol(right, .{}) catch return null;
    return fullColRange(
        .{ .col = lo, .absolute = false },
        .{ .col = hi, .absolute = false },
    );
}

fn stripDollar(s: []const u8) []const u8 {
    return if (s.len > 0 and s[0] == '$') s[1..] else s;
}

/// The area `$A:$B` denotes. A free function because the evaluator and
/// the static walk must agree on it exactly, and two copies of a grid
/// bound is one too many.
pub fn fullColRange(first: parser.ColBound, last: parser.ColBound) coords.Range {
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
pub fn fullRowRange(first: parser.RowBound, last: parser.RowBound) coords.Range {
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
/// arrangement `value.zig` uses, and for the same reason: the semantics
/// take the fold as a parameter so a second, quietly different
/// comparator cannot appear, and only the tests name the concrete one.
/// (Before M4f there was a build reason too, since the file then lived
/// inside the `zlsx` package tree.)
const casefold = @import("zlsx_casefold");
const rng = @import("rng.zig");

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
    text: []const u8,
    err: value.KnownError,
    /// A blank element inside an expected array (M7a: SORT pins §5.3b's
    /// blanks-first row against a range with an absent cell).
    blank,
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
        .text => |t| {
            try testing.expect(v == .scalar);
            try testing.expect(v.scalar == .text);
            try testing.expectEqualStrings(t, v.scalar.text);
        },
        .err => |e| {
            try testing.expect(v == .scalar);
            try testing.expect(v.scalar == .err);
            try testing.expectEqual(e, v.scalar.err.known);
        },
        .blank => {
            try testing.expect(v == .scalar);
            try testing.expect(v.scalar == .blank);
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
    // The dependency half of the same fact: an unevaluated fallback is
    // not read, which a draw counter alone would not have shown.
    .{ .formula = "IFERROR(A1,B1)", .expect = .{ .number = 10 }, .draws = 0, .runtime = &.{"A1"}, .absent = &.{"B1"}, .static = &.{ "A1", "B1" } },
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

// ─── M4c: the F1a-1 batch (§7, twenty names) ─────────────────────
//
// Oracle-first, and honest about how little the oracle decides here.
// One cell of one committed manifest touches this batch — `TRUE()+1`,
// recorded 2 by both the hand-spec excel suite and the LibreOffice one
// (§8.2). Everything else rests on the spec, so every row says which it
// is and `evidence` is checked against the manifests rather than
// trusted. That is the discipline §5.6g's 3D matrix shipped under at
// M4b3, applied to functions.

/// Where an expected value comes from. Same two-member shape as
/// `value.DivergencePoint.evidence`, and for the same reason: a fixture
/// that cannot say whether a machine or a paragraph decided it is a
/// fixture nobody can re-check.
const Evidence = enum { oracle, spec_pinned };

/// One F1a-1 fixture. `func` is the inventory name the row is a fixture
/// FOR — the coverage test derives the batch from the frozen TSV and
/// fails if any of the twenty has no row here.
const F1a1Case = struct {
    func: []const u8,
    formula: []const u8,
    expect: Expect,
    evidence: Evidence = .spec_pinned,
    /// Why the spec says so, where the answer is one a reasonable
    /// reading could get wrong.
    note: []const u8 = "",
};

/// The environment every F1a-1 fixture reads. One cell per §5.3b
/// provenance row, so a fixture can name the provenance it means
/// instead of building a bespoke sheet.
fn putF1a1Cells(h: *Harness) !void {
    try h.put("A1", num(10)); // number
    try h.put("A2", .{ .text = "abc" }); // non-numeric text
    try h.put("A3", .{ .text = "" }); // `=""`, which is text and not blank
    try h.put("A4", .{ .boolean = true }); // logical
    try h.put("A5", value.ScalarValue.errorOf(.div0)); // an error that is not #N/A
    try h.put("A6", value.ScalarValue.errorOf(.na)); // #N/A itself
    // A7 is a true blank — deliberately not stored.
    try h.put("A8", .{ .text = "7" }); // numeric text
}

const f1a1_cases = [_]F1a1Case{
    // ── TRUE / FALSE: the two names M3a2 added to this batch ──
    // The one row a committed manifest decides. `TRUE()` is a CALL, not
    // a boolean literal — which is what put it in the inventory (M3a2
    // decision 2) — so this is also the proof that the registry answers
    // for it.
    .{ .func = "TRUE", .formula = "TRUE()+1", .expect = .{ .number = 2 }, .evidence = .oracle },
    .{ .func = "TRUE", .formula = "TRUE()", .expect = .{ .boolean = true } },
    .{ .func = "FALSE", .formula = "FALSE()", .expect = .{ .boolean = false } },
    .{ .func = "FALSE", .formula = "FALSE()+1", .expect = .{ .number = 1 } },

    // ── NA: the only function whose whole job is to produce an error ──
    .{ .func = "NA", .formula = "NA()", .expect = .{ .err = .na } },
    .{ .func = "NA", .formula = "ISNA(NA())", .expect = .{ .boolean = true } },

    // ── NOT / AND / OR ──
    .{ .func = "NOT", .formula = "NOT(TRUE())", .expect = .{ .boolean = false } },
    .{ .func = "NOT", .formula = "NOT(A1)", .expect = .{ .boolean = false }, .note = "a non-zero number is a TRUE condition" },
    .{ .func = "NOT", .formula = "NOT(0)", .expect = .{ .boolean = true } },
    .{ .func = "NOT", .formula = "NOT(A2)", .expect = .{ .err = .value }, .note = "text never coerces to a condition" },
    .{ .func = "AND", .formula = "AND(TRUE(),FALSE())", .expect = .{ .boolean = false } },
    .{ .func = "AND", .formula = "AND(A1,A4)", .expect = .{ .boolean = true } },
    .{ .func = "AND", .formula = "AND(A2)", .expect = .{ .err = .value }, .note = "a direct text argument is #VALUE!" },
    .{ .func = "AND", .formula = "AND(A2:A3)", .expect = .{ .err = .value }, .note = "text in a range is ignored, and ignoring everything leaves no logical value" },
    .{ .func = "OR", .formula = "OR(TRUE(),FALSE())", .expect = .{ .boolean = true } },
    .{ .func = "OR", .formula = "OR(0,0)", .expect = .{ .boolean = false } },

    // ── the IS-family: `observe`, so an error is data ──
    .{ .func = "ISBLANK", .formula = "ISBLANK(A7)", .expect = .{ .boolean = true } },
    .{ .func = "ISBLANK", .formula = "ISBLANK(A3)", .expect = .{ .boolean = false }, .note = "`=\"\"` is text, not blank" },
    .{ .func = "ISNUMBER", .formula = "ISNUMBER(A1)", .expect = .{ .boolean = true } },
    .{ .func = "ISNUMBER", .formula = "ISNUMBER(A8)", .expect = .{ .boolean = false }, .note = "numeric text is text" },
    .{ .func = "ISTEXT", .formula = "ISTEXT(A2)", .expect = .{ .boolean = true } },
    .{ .func = "ISTEXT", .formula = "ISTEXT(A3)", .expect = .{ .boolean = true }, .note = "`\"\"` is text of length zero" },
    .{ .func = "ISTEXT", .formula = "ISTEXT(A7)", .expect = .{ .boolean = false }, .note = "a blank cell is not text" },
    .{ .func = "ISTEXT", .formula = "ISTEXT(A1)", .expect = .{ .boolean = false } },
    .{ .func = "ISLOGICAL", .formula = "ISLOGICAL(A4)", .expect = .{ .boolean = true } },
    .{ .func = "ISLOGICAL", .formula = "ISLOGICAL(A1)", .expect = .{ .boolean = false }, .note = "1 and TRUE are not the same value" },
    .{ .func = "ISERROR", .formula = "ISERROR(A5)", .expect = .{ .boolean = true } },
    .{ .func = "ISERROR", .formula = "ISERROR(A6)", .expect = .{ .boolean = true }, .note = "#N/A is an error to ISERROR" },
    .{ .func = "ISERROR", .formula = "ISERROR(A1)", .expect = .{ .boolean = false } },
    // The one distinction that makes ISERR a function rather than a
    // synonym, stated in both directions beside ISNA's mirror image.
    .{ .func = "ISERR", .formula = "ISERR(A5)", .expect = .{ .boolean = true } },
    .{ .func = "ISERR", .formula = "ISERR(A6)", .expect = .{ .boolean = false }, .note = "every error EXCEPT #N/A" },
    .{ .func = "ISERR", .formula = "ISERR(A1)", .expect = .{ .boolean = false } },
    .{ .func = "ISNA", .formula = "ISNA(A6)", .expect = .{ .boolean = true } },
    .{ .func = "ISNA", .formula = "ISNA(A5)", .expect = .{ .boolean = false } },

    // ── N and T: Excel's own conversion tables, not the coercion classes ──
    .{ .func = "N", .formula = "N(A1)", .expect = .{ .number = 10 } },
    .{ .func = "N", .formula = "N(A4)", .expect = .{ .number = 1 }, .note = "TRUE is 1, FALSE is 0" },
    .{ .func = "N", .formula = "N(FALSE())", .expect = .{ .number = 0 } },
    .{ .func = "N", .formula = "N(A2)", .expect = .{ .number = 0 }, .note = "anything else is 0, never #VALUE!" },
    .{ .func = "N", .formula = "N(A7)", .expect = .{ .number = 0 } },
    // The row most likely to be got wrong by reaching for the `.number`
    // coercion class: that class coerces numeric text and would answer 7.
    .{ .func = "N", .formula = "N(A8)", .expect = .{ .number = 0 }, .note = "N does NOT coerce numeric text" },
    .{ .func = "N", .formula = "N(\"7\")", .expect = .{ .number = 0 } },
    .{ .func = "N", .formula = "N(A5)", .expect = .{ .err = .div0 }, .note = "an error is the error — N propagates" },
    .{ .func = "T", .formula = "T(A2)", .expect = .{ .text = "abc" } },
    .{ .func = "T", .formula = "T(A1)", .expect = .{ .text = "" }, .note = "a number does not format; T(1) is \"\", not \"1\"" },
    .{ .func = "T", .formula = "T(A4)", .expect = .{ .text = "" } },
    .{ .func = "T", .formula = "T(A7)", .expect = .{ .text = "" } },
    .{ .func = "T", .formula = "T(A5)", .expect = .{ .err = .div0 } },
    .{ .func = "T", .formula = "T(A3)", .expect = .{ .text = "" } },

    // ── the conditional forms ──
    .{ .func = "IF", .formula = "IF(TRUE(),A1,A2)", .expect = .{ .number = 10 } },
    .{ .func = "IF", .formula = "IF(A2,1,2)", .expect = .{ .err = .value }, .note = "text is not a condition" },
    .{ .func = "IFS", .formula = "IFS(FALSE(),1,TRUE(),2)", .expect = .{ .number = 2 } },
    .{ .func = "IFS", .formula = "IFS(FALSE(),1,FALSE(),2)", .expect = .{ .err = .na }, .note = "nothing matched" },
    .{ .func = "SWITCH", .formula = "SWITCH(2,1,\"a\",2,\"b\")", .expect = .{ .text = "b" } },
    .{ .func = "SWITCH", .formula = "SWITCH(9,1,\"a\",2,\"b\")", .expect = .{ .err = .na } },
    .{ .func = "SWITCH", .formula = "SWITCH(9,1,\"a\",2,\"b\",\"z\")", .expect = .{ .text = "z" }, .note = "a trailing odd argument is the default" },
    .{ .func = "IFERROR", .formula = "IFERROR(A5,99)", .expect = .{ .number = 99 } },
    .{ .func = "IFERROR", .formula = "IFERROR(A6,99)", .expect = .{ .number = 99 }, .note = "IFERROR catches #N/A too" },
    .{ .func = "IFERROR", .formula = "IFERROR(A1,99)", .expect = .{ .number = 10 } },
    .{ .func = "IFNA", .formula = "IFNA(A6,99)", .expect = .{ .number = 99 } },
    .{ .func = "IFNA", .formula = "IFNA(A5,99)", .expect = .{ .err = .div0 }, .note = "IFNA catches #N/A ONLY — this is why it is a separate form" },
    .{ .func = "IFNA", .formula = "IFNA(A1,99)", .expect = .{ .number = 10 } },
};

test "M4c: every F1a-1 fixture evaluates to what the oracle or the spec says" {
    for (f1a1_cases) |c| {
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        try putF1a1Cells(&h);

        const v = h.eval(c.formula) catch |e| {
            std.debug.print("F1a-1 `{s}` refused: {t}\n", .{ c.formula, e });
            return e;
        };
        expectValue(c.expect, v) catch |e| {
            std.debug.print("F1a-1 `{s}` ({s}): wrong value\n", .{ c.formula, c.func });
            return e;
        };
    }
}

test "M4c: all twenty frozen names resolve, and each has a fixture" {
    // The batch is read from `function_inventory_v1.tsv`, never written
    // down here: §7 makes the file the count source, so a 21st row would
    // fail this test rather than silently ship unfixtured.
    var it = registry.inventory();
    var batch: usize = 0;
    while (it.next()) |e| {
        if (!std.mem.eql(u8, e.milestone, "M4c")) continue;
        batch += 1;

        if (registry.lookup(e.name) == null) {
            std.debug.print("F1a-1 name does not resolve: {s}\n", .{e.name});
            return error.UnregisteredBatchFunction;
        }
        var fixtures: usize = 0;
        for (f1a1_cases) |c| {
            if (std.mem.eql(u8, c.func, e.name)) fixtures += 1;
        }
        if (fixtures == 0) {
            std.debug.print("F1a-1 name has no fixture: {s}\n", .{e.name});
            return error.UnfixturedBatchFunction;
        }
    }
    try testing.expectEqual(@as(usize, 20), batch);

    // …and no fixture names something outside the batch, which would
    // make the coverage count above pass for the wrong reason.
    for (f1a1_cases) |c| {
        var found = false;
        var it2 = registry.inventory();
        while (it2.next()) |e| {
            if (std.mem.eql(u8, e.name, c.func) and std.mem.eql(u8, e.milestone, "M4c")) found = true;
        }
        if (!found) {
            std.debug.print("fixture names a function outside F1a-1: {s}\n", .{c.func});
            return error.FixtureOutsideBatch;
        }
    }
}

/// Whether any committed manifest records a cell whose formula is this
/// text. The point of asking the files rather than a list: a row
/// labelled `.oracle` that no manifest contains is a claim of evidence
/// that does not exist, and it is exactly the claim nobody re-checks.
fn manifestsDecide(formula: []const u8) !bool {
    for ([_][]const u8{ oracle_excel, oracle_ieee, oracle_libreoffice }) |json| {
        const doc = try std.json.parseFromSlice(std.json.Value, testing.allocator, json, .{});
        defer doc.deinit();
        for (doc.value.object.get("cells").?.array.items) |cell| {
            const f = cell.object.get("formula") orelse continue;
            if (std.mem.eql(u8, f.string, formula)) return true;
        }
    }
    return false;
}

test "M4c: the evidence label on every fixture is true of the committed manifests" {
    var oracle_rows: usize = 0;
    for (f1a1_cases) |c| {
        const decided = try manifestsDecide(c.formula);
        switch (c.evidence) {
            .oracle => {
                if (!decided) {
                    std.debug.print("`{s}` claims oracle evidence no manifest holds\n", .{c.formula});
                    return error.UnbackedOracleClaim;
                }
                oracle_rows += 1;
            },
            .spec_pinned => {
                if (decided) {
                    std.debug.print("`{s}` is decided by a manifest but ships spec-pinned\n", .{c.formula});
                    return error.UnderstatedEvidence;
                }
            },
        }
    }
    // Stated as a number so the balance cannot drift silently: the
    // committed manifests touch this batch exactly once. When the parked
    // Excel leg runs (§8.2) and the suite grows F1a-1 rows, this count
    // moves and the row that moves it is the row that re-labels.
    try testing.expectEqual(@as(usize, 1), oracle_rows);
}

test "M4c: error order in every multi-argument name of the batch (§5.3c)" {
    // §5.3c: eager arguments evaluate in declaration order and the first
    // error wins unless the class says otherwise. Every case below is
    // run in both argument orders, because a fixture with one error in
    // it proves propagation and says nothing about order.
    const Case = struct { formula: []const u8, expect: value.KnownError };
    const cases = [_]Case{
        // AND / OR: `propagate`, and the error found first is the answer.
        .{ .formula = "AND(A5,A6)", .expect = .div0 },
        .{ .formula = "AND(A6,A5)", .expect = .na },
        .{ .formula = "OR(A5,A6)", .expect = .div0 },
        .{ .formula = "OR(A6,A5)", .expect = .na },
        // …including from inside a range, in §5.6a's iteration order.
        .{ .formula = "AND(A5:A6)", .expect = .div0 },
        // IF: the condition is the only argument evaluated before a
        // branch is chosen, so its error is the whole result and the
        // arms never run.
        .{ .formula = "IF(A5,A6,A6)", .expect = .div0 },
        .{ .formula = "IF(A6,A5,A5)", .expect = .na },
        // IFS is eager: every arm evaluates, but the FIRST condition's
        // error still decides, because conditions are read in order.
        .{ .formula = "IFS(A5,1,A6,2)", .expect = .div0 },
        .{ .formula = "IFS(A6,1,A5,2)", .expect = .na },
        // SWITCH: the subject first, then each candidate in order.
        .{ .formula = "SWITCH(A5,1,2)", .expect = .div0 },
        .{ .formula = "SWITCH(1,A5,2,A6,3)", .expect = .div0 },
        .{ .formula = "SWITCH(1,A6,2,A5,3)", .expect = .na },
        // IFERROR/IFNA are `observe`, which INVERTS the rule: the first
        // argument's error is caught rather than propagated, so the
        // answer is the fallback's error and not the first one.
        .{ .formula = "IFERROR(A5,A6)", .expect = .na },
        .{ .formula = "IFERROR(A6,A5)", .expect = .div0 },
        // IFNA only catches #N/A, so a #DIV/0! first argument wins after
        // all — the same two cells, the opposite answer from IFERROR.
        .{ .formula = "IFNA(A5,A6)", .expect = .div0 },
        .{ .formula = "IFNA(A6,A5)", .expect = .div0 },
    };

    for (cases) |c| {
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        try putF1a1Cells(&h);

        const s = h.scalar(c.formula) catch |e| {
            std.debug.print("error-order case `{s}` refused: {t}\n", .{ c.formula, e });
            return e;
        };
        if (s != .err or s.err != .known or s.err.known != c.expect) {
            std.debug.print("error-order case `{s}`: expected {s}\n", .{ c.formula, c.expect.spelling() });
            return error.WrongErrorOrder;
        }
    }

    // Every multi-argument name in the batch appears above. The list is
    // derived from the registry's arity rather than typed out, so a
    // function that gains an argument later cannot slip past unordered.
    var it = registry.inventory();
    while (it.next()) |e| {
        if (!std.mem.eql(u8, e.milestone, "M4c")) continue;
        const f = registry.lookup(e.name).?;
        const multi = f.arity.max == null or f.arity.max.? > 1;
        if (!multi) continue;
        var covered = false;
        for (cases) |c| {
            if (std.mem.startsWith(u8, c.formula, e.name) and c.formula[e.name.len] == '(') covered = true;
        }
        if (!covered) {
            std.debug.print("multi-argument name with no error-order fixture: {s}\n", .{e.name});
            return error.MissingErrorOrderFixture;
        }
    }
}

test "M4c: the five conditional names, each proving its own §5.3a contract" {
    // The ladder row groups IF, IFS, SWITCH, IFERROR and IFNA together
    // as "the forms". They do NOT share a contract: §5.3a defers three
    // of them and declares the other two eager, so the proof for each
    // pair is the opposite of the proof for the other. Both are draw
    // counts, and a draw count is the one instrument that cannot be
    // satisfied by a right answer arrived at wrongly.
    const Case = struct { formula: []const u8, draws: u64, absent: []const []const u8 = &.{} };
    const lazy = [_]Case{
        // Three arms written, one evaluated.
        .{ .formula = "IF(TRUE(),RAND(),RAND())", .draws = 1 },
        .{ .formula = "IF(FALSE(),RAND(),RAND())", .draws = 1 },
        .{ .formula = "IF(TRUE(),1,A1)", .draws = 0, .absent = &.{"A1"} },
        // The fallback is not evaluated when there is nothing to catch.
        .{ .formula = "IFERROR(1,RAND())", .draws = 0 },
        .{ .formula = "IFERROR(1,A1)", .draws = 0, .absent = &.{"A1"} },
        .{ .formula = "IFNA(1,RAND())", .draws = 0 },
        // …nor when the error is one this form does not catch.
        .{ .formula = "IFNA(A5,RAND())", .draws = 0 },
        .{ .formula = "IFNA(A5,A1)", .draws = 0, .absent = &.{"A1"} },
    };
    const eager = [_]Case{
        // Both arms drawn, both conditions drawn: four RAND()s, four
        // draws, whatever the first condition answered.
        .{ .formula = "IFS(TRUE(),RAND(),TRUE(),RAND())", .draws = 2 },
        .{ .formula = "IFS(FALSE(),RAND(),TRUE(),RAND())", .draws = 2 },
        .{ .formula = "SWITCH(1,1,RAND(),2,RAND())", .draws = 2 },
        .{ .formula = "SWITCH(2,1,RAND(),2,RAND())", .draws = 2 },
    };

    for (lazy ++ eager) |c| {
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        try putF1a1Cells(&h);

        _ = h.eval(c.formula) catch |e| {
            std.debug.print("form case `{s}` refused: {t}\n", .{ c.formula, e });
            return e;
        };
        if (h.draws.count != c.draws) {
            std.debug.print(
                "`{s}`: expected {d} draws, saw {d}\n",
                .{ c.formula, c.draws, h.draws.count },
            );
            return error.WrongDrawCount;
        }
        for (c.absent) |a1| {
            const cell = cellOf(a1);
            if (h.ev.deps.hasCell(.{ .sheet = h.sheet, .row = cell.row, .col = cell.col })) {
                std.debug.print("`{s}`: {s} was read from an arm that never runs\n", .{ c.formula, a1 });
                return error.UnexpectedRuntimeDependency;
            }
        }
    }

    // And the same fact stated where it is easy to check: the eager pair
    // draw once per written RAND(), the lazy trio do not.
    for (eager) |c| try testing.expect(c.draws == 2);
}

// ─── M4c fuzz: no argument shape panics, leaks, or evaluates twice ──

const f1a1_names = [_][]const u8{
    "AND",     "FALSE", "IF",      "IFERROR",   "IFNA", "IFS",
    "ISBLANK", "ISERR", "ISERROR", "ISLOGICAL", "ISNA", "ISNUMBER",
    "ISTEXT",  "N",     "NA",      "NOT",       "OR",   "SWITCH",
    "T",       "TRUE",
};

/// Argument *shapes*, not values: the fuzz is about what a slot can be
/// handed, so the list spans scalars, every provenance, references,
/// multi-area sets, arrays of both orientations, an omitted argument, a
/// nested call, and the two constructs that produce a plane-2 refusal.
const f1a1_arg_shapes = [_][]const u8{
    "A1",            "A2",    "A3",     "A5",
    "A7",            "A8",    "\"\"",   "\"7\"",
    "TRUE()",        "0",     "-1",     "1/0",
    "NA()",          "{1,2}", "{1;2}",  "A1:A8",
    "(A1:A2,A5:A6)", "",      "@A1:A8", "1E+308*10",
    "A1:B2 B1:B9",   "N(A8)", "S!A1",   "\"1,5\"",
};

fn valuesAgree(a: Value, b: Value) bool {
    if (@as(std.meta.Tag(Value), a) != @as(std.meta.Tag(Value), b)) return false;
    return switch (a) {
        .scalar => |s| s.eql(b.scalar),
        .missing_arg => true,
        .array => |m| blk: {
            if (m.rows != b.array.rows or m.cols != b.array.cols) break :blk false;
            for (m.cells, b.array.cells) |x, y| {
                if (!x.eql(y)) break :blk false;
            }
            break :blk true;
        },
        .reference => |r| blk: {
            if (r.areas.len != b.reference.areas.len) break :blk false;
            if (r.three_d != b.reference.three_d) break :blk false;
            for (r.areas, b.reference.areas) |x, y| {
                if (!std.meta.eql(x, y)) break :blk false;
            }
            break :blk true;
        },
    };
}

fn fuzzF1a1Env(fake: *env.Fake) !env.SheetIndex {
    const sheet = try fake.addSheet("S");
    try fake.putA1(sheet, .stored, "A1", num(10));
    try fake.putA1(sheet, .stored, "A2", .{ .text = "abc" });
    try fake.putA1(sheet, .stored, "A3", .{ .text = "" });
    try fake.putA1(sheet, .stored, "A4", .{ .boolean = true });
    try fake.putA1(sheet, .stored, "A5", value.ScalarValue.errorOf(.div0));
    try fake.putA1(sheet, .stored, "A6", value.ScalarValue.errorOf(.na));
    try fake.putA1(sheet, .stored, "A8", .{ .text = "7" });
    return sheet;
}

fn fuzzF1a1Target(_: void, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();

    // Build a call to one of the twenty out of the shape alphabet. The
    // grammar is fixed and the arguments are not, because the property
    // under test is about argument shapes reaching a registered slot —
    // malformed *text* is `fuzzEvalTarget`'s job.
    var buf: [512]u8 = undefined;
    var w: usize = 0;
    const name = f1a1_names[smith.index(f1a1_names.len)];
    @memcpy(buf[w..][0..name.len], name);
    w += name.len;
    buf[w] = '(';
    w += 1;
    var n: usize = 0;
    while (n < 6 and !smith.eos()) : (n += 1) {
        const arg = f1a1_arg_shapes[smith.index(f1a1_arg_shapes.len)];
        if (w + arg.len + 2 > buf.len) break;
        if (n > 0) {
            buf[w] = ',';
            w += 1;
        }
        @memcpy(buf[w..][0..arg.len], arg);
        w += arg.len;
    }
    buf[w] = ')';
    w += 1;
    const src = buf[0..w];

    var parsed = parser.parse(std.testing.allocator, src, .{}) catch return;
    defer parsed.deinit(std.testing.allocator);
    if (parsed == .refused) return;

    var arena_state = std.heap.ArenaAllocator.init(std.testing.allocator);
    defer arena_state.deinit();
    var fake = env.Fake.init(std.testing.allocator);
    defer fake.deinit();
    const sheet = try fuzzF1a1Env(&fake);

    var draw_value: f64 = 0.5;
    var draws = DrawSource.constant(&draw_value);
    const opts: Options = .{
        .current_sheet = sheet,
        .collation = .{ .fold = shippedFold },
        .draws = &draws,
        .site = .{
            .row = coords.Row.fromOneBased(2) catch unreachable,
            .col = coords.Col.fromZeroBased(1) catch unreachable,
        },
    };

    // Two evaluators, both alive: results borrow the arena, and the
    // arena outlives both. "Evaluates two ways" is the failure a single
    // run cannot see — a registry entry that reads uninitialized
    // padding, or an implementation whose answer depends on where the
    // arena happened to be.
    var first = Evaluator.init(arena_state.allocator(), fake.evalEnv(), opts);
    defer first.deinit();
    var second = Evaluator.init(arena_state.allocator(), fake.evalEnv(), opts);
    defer second.deinit();

    if (first.evaluate(parsed.ok)) |a| {
        try assertRepresentable(a);
        const b = second.evaluate(parsed.ok) catch return error.NondeterministicEvaluation;
        try assertRepresentable(b);
        if (!valuesAgree(a, b)) {
            std.debug.print("`{s}` evaluated two ways\n", .{src});
            return error.NondeterministicEvaluation;
        }
    } else |e| {
        // A typed refusal is a legitimate outcome — but it has to be the
        // same refusal twice.
        if (second.evaluate(parsed.ok)) |_| {
            return error.NondeterministicEvaluation;
        } else |e2| {
            if (e != e2) return error.NondeterministicEvaluation;
        }
    }
}

test "fuzz: no F1a-1 argument shape panics, leaks, or evaluates two ways" {
    // The generator's alphabet is a second copy of the batch, so it gets
    // the same treatment as every other copy in this row: checked
    // against the file, not maintained by hand. A name the fuzzer never
    // builds is a name the fuzzer never covers, and nothing else would
    // have said so.
    var it = registry.inventory();
    var batch: usize = 0;
    while (it.next()) |e| {
        if (!std.mem.eql(u8, e.milestone, "M4c")) continue;
        batch += 1;
        var present = false;
        for (f1a1_names) |n| {
            if (std.mem.eql(u8, n, e.name)) present = true;
        }
        if (!present) {
            std.debug.print("fuzz alphabet is missing {s}\n", .{e.name});
            return error.FuzzAlphabetIncomplete;
        }
    }
    try testing.expectEqual(batch, f1a1_names.len);

    try std.testing.fuzz({}, fuzzF1a1Target, .{
        .corpus = &[_][]const u8{
            "IF(A1,A2,A5)",
            "IFS(A5,A1,A6,A2)",
            "SWITCH(A1,A1,A2,A5)",
            "IFERROR(A5,A6)",
            "IFNA(A6,{1,2})",
            "AND(A1:A8)",
            "OR((A1:A2,A5:A6))",
            "NOT(@A1:A8)",
            "ISBLANK(A7)",
            "ISERR(1/0)",
            "ISNA(NA())",
            "ISLOGICAL(TRUE())",
            "ISTEXT(\"\")",
            "ISNUMBER({1,2})",
            "ISERROR(A1:B2 B1:B9)",
            "N(A8)",
            "T(A2)",
            "N(\"1,5\")",
            "T({1;2})",
            "NA()",
            "TRUE()",
            "FALSE()",
        },
    });
}

// ─── M4d: the F1a-2 batch (§7, seventeen numeric names) ──────────
//
// Oracle-first, and the oracle decides one cell of it: `SQRT(-1)`,
// which both excel-fidelity manifests record and which they record
// DIFFERENTLY. Everything else in seventeen functions rests on the spec,
// so every row says which it is and the label is checked against the
// files rather than trusted — the discipline M4c shipped, with one
// addition this batch forced. F1a-1 held no volatile, so "does a
// manifest contain this formula?" was a two-valued question there. The
// LibreOffice suite records a `RAND()` cell and marks it
// `"excluded": "volatile_formula"`; a cell that is recorded and
// excluded decides nothing, and reading it as evidence would let this
// row claim oracle backing for a value no oracle ever wrote down. So
// M4d asks a three-valued question instead.

/// One F1a-2 fixture. `func` is the inventory name the row is a fixture
/// FOR — the coverage test derives the batch from the frozen TSV and
/// fails if any of the seventeen has no row here.
const F1a2Case = struct {
    func: []const u8,
    formula: []const u8,
    expect: Expect,
    evidence: Evidence = .spec_pinned,
    /// Why the spec says so, where the answer is one a reasonable
    /// reading could get wrong.
    note: []const u8 = "",
};

/// The environment every F1a-2 fixture reads. The A-column repeats
/// M4c's provenance rows so the two batches can be compared cell for
/// cell; the B column carries the three values that make the fidelity
/// modes disagree, and they are **stored** rather than written as
/// literals on purpose — N1a would round a 17-digit literal to 15 at
/// ingress under `.excel`, and then the divergence under test would be
/// the parser's rather than the function's.
fn putF1a2Cells(h: *Harness) !void {
    try h.put("A1", num(10)); // number
    try h.put("A2", .{ .text = "abc" }); // non-numeric text
    try h.put("A4", .{ .boolean = true }); // logical
    try h.put("A5", value.ScalarValue.errorOf(.div0)); // an error that is not #N/A
    try h.put("A6", value.ScalarValue.errorOf(.na)); // #N/A itself
    // A7 is a true blank — deliberately not stored.
    try h.put("A8", .{ .text = "7" }); // numeric text
    // One ULP below 2.5 and below 3: the 15-digit decimal view rounds
    // each up to the boundary, raw binary64 does not.
    try h.put("B1", num(2.4999999999999996));
    try h.put("B2", num(2.9999999999999996));
    // Negative and small: every truncating mode sends it to `-0`.
    try h.put("B3", num(-0.4));
}

const f1a2_cases = [_]F1a2Case{
    // ── ABS / SIGN / INT: the shape-preserving trio ──
    .{ .func = "ABS", .formula = "ABS(-3)", .expect = .{ .number = 3 } },
    .{ .func = "ABS", .formula = "ABS(A1)", .expect = .{ .number = 10 } },
    .{ .func = "ABS", .formula = "ABS(A5)", .expect = .{ .err = .div0 }, .note = "numeric functions propagate; none of this batch observes" },
    .{ .func = "ABS", .formula = "ABS(A8)", .expect = .{ .number = 7 }, .note = "the `.number` class coerces numeric text, unlike N" },
    .{ .func = "ABS", .formula = "ABS(A2)", .expect = .{ .err = .value }, .note = "non-numeric text is #VALUE! at the coercion, not in the impl" },
    .{ .func = "ABS", .formula = "ABS(A7)", .expect = .{ .number = 0 }, .note = "a blank coerces to 0" },
    .{ .func = "SIGN", .formula = "SIGN(-4)", .expect = .{ .number = -1 } },
    .{ .func = "SIGN", .formula = "SIGN(0)", .expect = .{ .number = 0 } },
    .{ .func = "SIGN", .formula = "SIGN(A1)", .expect = .{ .number = 1 } },
    .{ .func = "SIGN", .formula = "SIGN(A4)", .expect = .{ .number = 1 }, .note = "TRUE coerces to 1" },
    .{ .func = "INT", .formula = "INT(2.9)", .expect = .{ .number = 2 } },
    .{ .func = "INT", .formula = "INT(-2.5)", .expect = .{ .number = -3 }, .note = "INT floors; TRUNC truncates. The pair below is the same argument" },
    .{ .func = "INT", .formula = "INT(A1)", .expect = .{ .number = 10 } },

    // ── the rounding family ──
    .{ .func = "TRUNC", .formula = "TRUNC(-2.5)", .expect = .{ .number = -2 }, .note = "toward zero, where INT(-2.5) is -3" },
    .{ .func = "TRUNC", .formula = "TRUNC(2.9)", .expect = .{ .number = 2 } },
    .{ .func = "TRUNC", .formula = "TRUNC(3.14159,2)", .expect = .{ .number = 3.14 } },
    .{ .func = "TRUNC", .formula = "TRUNC(-3.14159,2)", .expect = .{ .number = -3.14 } },
    .{ .func = "TRUNC", .formula = "TRUNC(1234.5,-2)", .expect = .{ .number = 1200 }, .note = "a negative digit count rounds to the left of the point" },
    .{ .func = "ROUND", .formula = "ROUND(2.5,0)", .expect = .{ .number = 3 }, .note = "half AWAY from zero — not the banker's rounding IEEE defaults to" },
    .{ .func = "ROUND", .formula = "ROUND(-2.5,0)", .expect = .{ .number = -3 } },
    .{ .func = "ROUND", .formula = "ROUND(2.4,0)", .expect = .{ .number = 2 } },
    .{ .func = "ROUND", .formula = "ROUND(3.14159,2)", .expect = .{ .number = 3.14 } },
    .{ .func = "ROUND", .formula = "ROUND(1234.5,-2)", .expect = .{ .number = 1200 } },
    .{ .func = "ROUND", .formula = "ROUND(1250,-2)", .expect = .{ .number = 1300 } },
    .{ .func = "ROUND", .formula = "ROUND(A1,2.9)", .expect = .{ .number = 10 }, .note = "the digit count truncates toward zero: 2.9 digits is 2" },
    .{ .func = "ROUNDUP", .formula = "ROUNDUP(2.1,0)", .expect = .{ .number = 3 } },
    .{ .func = "ROUNDUP", .formula = "ROUNDUP(-2.1,0)", .expect = .{ .number = -3 }, .note = "away from zero in BOTH directions — not ceiling" },
    .{ .func = "ROUNDUP", .formula = "ROUNDUP(3.14159,2)", .expect = .{ .number = 3.15 } },
    .{ .func = "ROUNDUP", .formula = "ROUNDUP(0,0)", .expect = .{ .number = 0 }, .note = "zero has no significant digit to round away from" },
    .{ .func = "ROUNDDOWN", .formula = "ROUNDDOWN(2.9,0)", .expect = .{ .number = 2 } },
    .{ .func = "ROUNDDOWN", .formula = "ROUNDDOWN(-2.9,0)", .expect = .{ .number = -2 }, .note = "toward zero in both directions — not floor" },
    .{ .func = "ROUNDDOWN", .formula = "ROUNDDOWN(3.14159,2)", .expect = .{ .number = 3.14 } },
    // The place value pushed past everything the value holds, from both
    // ends. These are the two branches `roundAt` takes before it scales.
    .{ .func = "ROUND", .formula = "ROUND(A1,400)", .expect = .{ .number = 10 }, .note = "a place below the last significant digit cannot change the value" },
    .{ .func = "ROUND", .formula = "ROUND(A1,-400)", .expect = .{ .number = 0 }, .note = "a place far above the leading digit removes every one of them" },
    .{ .func = "ROUNDUP", .formula = "ROUNDUP(A1,-400)", .expect = .{ .err = .num }, .note = "…and rounding AWAY from zero at that place overflows (N4a)" },
    .{ .func = "ROUNDDOWN", .formula = "ROUNDDOWN(A1,-400)", .expect = .{ .number = 0 } },

    // ── MOD: §5.4's N4 names its sign rule specifically ──
    .{ .func = "MOD", .formula = "MOD(5,3)", .expect = .{ .number = 2 } },
    .{ .func = "MOD", .formula = "MOD(-5,3)", .expect = .{ .number = 1 }, .note = "the result takes the DIVISOR's sign, so this is 1 and not -2" },
    .{ .func = "MOD", .formula = "MOD(5,-3)", .expect = .{ .number = -1 } },
    .{ .func = "MOD", .formula = "MOD(-5,-3)", .expect = .{ .number = -2 } },
    .{ .func = "MOD", .formula = "MOD(6,3)", .expect = .{ .number = 0 } },
    .{ .func = "MOD", .formula = "MOD(5,0)", .expect = .{ .err = .div0 }, .note = "a zero divisor, spelled the way division spells it" },

    // ── POWER: the function spelling of `^` ──
    .{ .func = "POWER", .formula = "POWER(2,10)", .expect = .{ .number = 1024 } },
    .{ .func = "POWER", .formula = "POWER(-2,3)", .expect = .{ .number = -8 } },
    .{ .func = "POWER", .formula = "POWER(2,0.5)", .expect = .{ .number = 1.4142135623730951 } },
    .{ .func = "POWER", .formula = "POWER(-8,1/3)", .expect = .{ .err = .num }, .note = "a negative base with a fractional exponent has no real root" },
    .{ .func = "POWER", .formula = "POWER(2,1024)", .expect = .{ .err = .num }, .note = "overflow is #NUM! in both modes (N4a)" },

    // ── SQRT: the row's one oracle-decided cell ──
    .{
        .func = "SQRT",
        .formula = "SQRT(-1)",
        .expect = .{ .err = .num },
        .evidence = .oracle,
        .note = "the hand-spec excel manifest records #NUM!; LibreOffice records #VALUE!, and that disagreement is a NAMED adapter divergence",
    },
    .{ .func = "SQRT", .formula = "SQRT(9)", .expect = .{ .number = 3 } },
    .{ .func = "SQRT", .formula = "SQRT(0)", .expect = .{ .number = 0 } },
    .{ .func = "SQRT", .formula = "SQRT(2)", .expect = .{ .number = 1.4142135623730951 } },

    // ── the exponential/logarithm family ──
    .{ .func = "EXP", .formula = "EXP(0)", .expect = .{ .number = 1 } },
    .{ .func = "EXP", .formula = "EXP(1)", .expect = .{ .number = 2.718281828459045 } },
    .{ .func = "EXP", .formula = "EXP(1000)", .expect = .{ .err = .num }, .note = "overflow, reached through fromArithmetic rather than a magnitude test" },
    .{ .func = "EXP", .formula = "EXP(-1000)", .expect = .{ .number = 0 }, .note = "underflow is a representable 0, and therefore a value" },
    .{ .func = "LN", .formula = "LN(1)", .expect = .{ .number = 0 } },
    .{ .func = "LN", .formula = "LN(A1)", .expect = .{ .number = 2.302585092994046 } },
    .{ .func = "LN", .formula = "LN(0)", .expect = .{ .err = .num }, .note = "LN(0) is #NUM!, not -infinity: N4a has no room for one" },
    .{ .func = "LN", .formula = "LN(-1)", .expect = .{ .err = .num } },
    .{ .func = "LOG10", .formula = "LOG10(1000)", .expect = .{ .number = 3 } },
    .{ .func = "LOG10", .formula = "LOG10(1)", .expect = .{ .number = 0 } },
    .{ .func = "LOG10", .formula = "LOG10(0)", .expect = .{ .err = .num } },
    .{ .func = "LOG", .formula = "LOG(100)", .expect = .{ .number = 2 }, .note = "the base defaults to 10, which makes LOG(x) and LOG10(x) one operation" },
    .{ .func = "LOG", .formula = "LOG(100,10)", .expect = .{ .number = 2 } },
    .{ .func = "LOG", .formula = "LOG(8,2)", .expect = .{ .number = 3 } },
    .{ .func = "LOG", .formula = "LOG(0)", .expect = .{ .err = .num } },
    .{ .func = "LOG", .formula = "LOG(10,-2)", .expect = .{ .err = .num }, .note = "a non-positive base is a domain error" },
    .{ .func = "LOG", .formula = "LOG(10,1)", .expect = .{ .err = .div0 }, .note = "base 1 divides by LN(1) — the family's one non-#NUM! failure" },

    // ── PI: the batch's only constant ──
    .{ .func = "PI", .formula = "PI()", .expect = .{ .number = 3.141592653589793 } },
    .{ .func = "PI", .formula = "PI()*2", .expect = .{ .number = 6.283185307179586 } },

    // ── the two volatiles, under the harness's fixed draw of 0.5 ──
    .{
        .func = "RAND",
        .formula = "RAND()",
        .expect = .{ .number = 0.5 },
        .note = "the LibreOffice suite RECORDS this formula and excludes it as volatile; an excluded cell is not evidence",
    },
    .{ .func = "RANDBETWEEN", .formula = "RANDBETWEEN(1,1)", .expect = .{ .number = 1 }, .note = "a range of one is deterministic whatever the draw returns" },
    .{ .func = "RANDBETWEEN", .formula = "RANDBETWEEN(0,10)", .expect = .{ .number = 5 } },
    .{ .func = "RANDBETWEEN", .formula = "RANDBETWEEN(-3,-3)", .expect = .{ .number = -3 } },
    .{ .func = "RANDBETWEEN", .formula = "RANDBETWEEN(5,1)", .expect = .{ .err = .num }, .note = "an empty range is #NUM!" },
    .{ .func = "RANDBETWEEN", .formula = "RANDBETWEEN(1.5,3.5)", .expect = .{ .number = 3 }, .note = "non-integer bounds move INWARD, so every result is inside [bottom, top]" },
    .{ .func = "RANDBETWEEN", .formula = "RANDBETWEEN(1.2,1.8)", .expect = .{ .err = .num }, .note = "…and a range holding no integer is empty after they move" },
};

test "M4d: every F1a-2 fixture evaluates to what the oracle or the spec says" {
    for (f1a2_cases) |c| {
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        try putF1a2Cells(&h);

        const v = h.eval(c.formula) catch |e| {
            std.debug.print("F1a-2 `{s}` refused: {t}\n", .{ c.formula, e });
            return e;
        };
        expectValue(c.expect, v) catch |e| {
            std.debug.print("F1a-2 `{s}` ({s}): wrong value\n", .{ c.formula, c.func });
            return e;
        };
    }
}

test "M4d: all seventeen frozen names resolve, and each has a fixture" {
    // The batch is read from `function_inventory_v1.tsv`, never written
    // down here: §7 makes the file the count source, so an eighteenth
    // row fails this test rather than shipping unfixtured.
    var it = registry.inventory();
    var batch: usize = 0;
    while (it.next()) |e| {
        if (!std.mem.eql(u8, e.milestone, "M4d")) continue;
        batch += 1;

        if (registry.lookup(e.name) == null) {
            std.debug.print("F1a-2 name does not resolve: {s}\n", .{e.name});
            return error.UnregisteredBatchFunction;
        }
        var fixtures: usize = 0;
        for (f1a2_cases) |c| {
            if (std.mem.eql(u8, c.func, e.name)) fixtures += 1;
        }
        if (fixtures == 0) {
            std.debug.print("F1a-2 name has no fixture: {s}\n", .{e.name});
            return error.UnfixturedBatchFunction;
        }
    }
    try testing.expectEqual(@as(usize, 17), batch);

    // …and no fixture names something outside the batch, which would
    // make the coverage count above pass for the wrong reason.
    for (f1a2_cases) |c| {
        var found = false;
        var it2 = registry.inventory();
        while (it2.next()) |e| {
            if (std.mem.eql(u8, e.name, c.func) and std.mem.eql(u8, e.milestone, "M4d")) found = true;
        }
        if (!found) {
            std.debug.print("fixture names a function outside F1a-2: {s}\n", .{c.func});
            return error.FixtureOutsideBatch;
        }
    }
}

/// What the committed manifests say about a formula. **Three** answers,
/// where M4c needed two: a cell can be recorded and still decide
/// nothing, which is exactly what `"excluded": "volatile_formula"`
/// means. F1a-1 held no volatile so the distinction never arose; F1a-2
/// registers `RAND` and `RANDBETWEEN`, and the LibreOffice suite records
/// a `RAND()` cell it deliberately excludes.
const ManifestVerdict = enum { silent, decided, excluded };

fn manifestVerdict(formula: []const u8) !ManifestVerdict {
    var verdict: ManifestVerdict = .silent;
    for ([_][]const u8{ oracle_excel, oracle_ieee, oracle_libreoffice }) |json| {
        const doc = try std.json.parseFromSlice(std.json.Value, testing.allocator, json, .{});
        defer doc.deinit();
        for (doc.value.object.get("cells").?.array.items) |cell| {
            const f = cell.object.get("formula") orelse continue;
            if (!std.mem.eql(u8, f.string, formula)) continue;
            // A recorded value outranks an excluded one: if any manifest
            // decides the formula, it is decided.
            if (cell.object.get("excluded") == null) return .decided;
            verdict = .excluded;
        }
    }
    return verdict;
}

test "M4d: the evidence label on every fixture is true of the committed manifests" {
    var oracle_rows: usize = 0;
    var excluded_rows: usize = 0;
    for (f1a2_cases) |c| {
        switch (try manifestVerdict(c.formula)) {
            .decided => {
                if (c.evidence != .oracle) {
                    std.debug.print("`{s}` is decided by a manifest but ships spec-pinned\n", .{c.formula});
                    return error.UnderstatedEvidence;
                }
                oracle_rows += 1;
            },
            .excluded => {
                // Recorded, and recorded as undecidable. Reading this as
                // evidence is the specific mistake this arm exists to
                // make impossible.
                if (c.evidence != .spec_pinned) {
                    std.debug.print("`{s}` claims evidence from an EXCLUDED manifest cell\n", .{c.formula});
                    return error.ExcludedCellClaimedAsEvidence;
                }
                excluded_rows += 1;
            },
            .silent => {
                if (c.evidence != .spec_pinned) {
                    std.debug.print("`{s}` claims oracle evidence no manifest holds\n", .{c.formula});
                    return error.UnbackedOracleClaim;
                }
            },
        }
    }
    // Stated as numbers so the balance cannot drift silently. One
    // oracle-decided cell in seventeen functions — `SQRT(-1)` — and one
    // recorded-but-excluded cell, `RAND()`. When the parked Excel leg
    // runs (§8.2) and the suite grows F1a-2 rows, these counts move and
    // the row that moves them is the row that re-labels.
    try testing.expectEqual(@as(usize, 1), oracle_rows);
    try testing.expectEqual(@as(usize, 1), excluded_rows);
}

/// The error a manifest recorded for a formula, or null if it recorded
/// no such cell (or recorded a non-error there).
fn recordedErrorFor(json: []const u8, formula: []const u8) !?value.KnownError {
    const doc = try std.json.parseFromSlice(std.json.Value, testing.allocator, json, .{});
    defer doc.deinit();
    for (doc.value.object.get("cells").?.array.items) |cell| {
        const f = cell.object.get("formula") orelse continue;
        if (!std.mem.eql(u8, f.string, formula)) continue;
        const spelling = cell.object.get("error_spelling") orelse continue;
        return value.KnownError.fromSpelling(spelling.string);
    }
    return null;
}

test "M4d: SQRT(-1) is real evidence, and the two excel manifests still disagree" {
    // The batch's one oracle cell is also a named adapter divergence,
    // and those two facts are not in tension — they are the same fact.
    // Both files claim `"fidelity": "excel"`; they record different
    // errors for the same formula; so at most one of them is Excel, and
    // the row that averaged them would produce an answer neither adapter
    // gave. This test asserts the DISAGREEMENT, not merely the skip.
    const hand = try recordedErrorFor(oracle_excel, "SQRT(-1)");
    const lo = try recordedErrorFor(oracle_libreoffice, "SQRT(-1)");
    try testing.expectEqual(value.KnownError.num, hand.?);
    try testing.expectEqual(value.KnownError.value, lo.?);
    try testing.expect(hand.? != lo.?);
    // The ieee manifest is silent, so there is no third answer.
    try testing.expect((try recordedErrorFor(oracle_ieee, "SQRT(-1)")) == null);

    // The row is named in the divergence list rather than skipped
    // anonymously — `-0` and this one, and the tie test asserts the
    // skip count is exactly two.
    var listed = false;
    for (excel_adapter_divergences) |f| {
        if (std.mem.eql(u8, f, "SQRT(-1)")) listed = true;
    }
    try testing.expect(listed);

    // zlsx answers the hand-spec's `#NUM!` — Excel's answer — and does
    // so in both modes, because a radicand's domain is not a
    // floating-point-rules question.
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    for ([_]value.Fidelity{ .excel, .ieee }) |mode| {
        var opts = h.options();
        opts.fidelity = mode;
        const v = try h.evalOpts("SQRT(-1)", opts);
        try testing.expectEqual(value.KnownError.num, v.scalar.err.known);
    }
}

/// One row where `excel_fp_rules_v1` and `ieee_fp_rules_v1` answer
/// differently. Compared as **published** values, bit for bit: N3's
/// signed-zero policy applies at publication, so a comparison anywhere
/// else would miss half of these.
const F1a2FidelityCase = struct {
    formula: []const u8,
    excel: f64,
    ieee: f64,
    why: []const u8,
};

const f1a2_fidelity_cases = [_]F1a2FidelityCase{
    // N1a's 15 significant digits are not only an ingress rule: under
    // `excel_fp_rules_v1` they are what the value IS, so a rounding
    // decision is taken on the decimal a user sees. B1 is one ULP below
    // 2.5 and B2 one ULP below 3.
    .{ .formula = "ROUND(B1,0)", .excel = 3, .ieee = 2, .why = "the 15-digit view of B1 is exactly 2.5, and 2.5 rounds away from zero" },
    .{ .formula = "INT(B2)", .excel = 3, .ieee = 2, .why = "the same view, floored" },
    .{ .formula = "TRUNC(B2)", .excel = 3, .ieee = 2, .why = "…and truncated" },
    .{ .formula = "ROUNDDOWN(B2,0)", .excel = 3, .ieee = 2, .why = "TRUNC(x) and ROUNDDOWN(x,0) agree in both modes" },
    // N3: a truncating mode produces `-0`, which `.excel` normalizes at
    // publication and `.ieee` preserves bitwise.
    .{ .formula = "ROUNDDOWN(B3,0)", .excel = 0, .ieee = -0.0, .why = "rounding -0.4 toward zero yields -0" },
    .{ .formula = "TRUNC(B3)", .excel = 0, .ieee = -0.0, .why = "the same value by the other spelling" },
    .{ .formula = "ROUND(B3,0)", .excel = 0, .ieee = -0.0, .why = "half-away-from-zero also leaves the sign behind" },
    .{ .formula = "ROUNDDOWN(-1,-400)", .excel = 0, .ieee = -0.0, .why = "the collapse branch keeps the sign too, which is where a bare `0` would have lost it" },
};

test "M4d: both fidelity modes are fixtured wherever the two rule tables disagree" {
    for (f1a2_fidelity_cases) |c| {
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        try putF1a2Cells(&h);

        inline for ([_]value.Fidelity{ .excel, .ieee }) |mode| {
            var opts = h.options();
            opts.fidelity = mode;
            const v = h.evalOpts(c.formula, opts) catch |e| {
                std.debug.print("fidelity case `{s}` refused: {t}\n", .{ c.formula, e });
                return e;
            };
            try testing.expect(v == .scalar);
            const got = value.publish(v.scalar, mode);
            const want = if (mode == .excel) c.excel else c.ieee;
            try testing.expect(got == .number);
            if (@as(u64, @bitCast(got.number)) != @as(u64, @bitCast(want))) {
                std.debug.print(
                    "`{s}` under {s}: expected {d} (0x{X:0>16}), got {d} (0x{X:0>16}) — {s}\n",
                    .{
                        c.formula,
                        @tagName(mode),
                        want,
                        @as(u64, @bitCast(want)),
                        got.number,
                        @as(u64, @bitCast(got.number)),
                        c.why,
                    },
                );
                return error.FidelityMismatch;
            }
        }

        // A table of "divergences" that agreed everywhere would pass a
        // per-mode check and prove nothing, so every row must actually
        // differ — the same both-directions gate M3a1's Divergence ×2
        // applies to the rule tables themselves.
        if (@as(u64, @bitCast(c.excel)) == @as(u64, @bitCast(c.ieee))) {
            std.debug.print("`{s}` is listed as a divergence but does not diverge\n", .{c.formula});
            return error.NonDivergentFidelityCase;
        }
    }
}

test "M4d: the batch agrees across modes everywhere it is not listed as diverging" {
    // The other half of the gate. Most of F1a-2 is mode-independent, and
    // saying so is what makes the short list above meaningful: if the
    // modes differed everywhere, the divergence table would be a
    // sampling rather than an enumeration.
    for (f1a2_cases) |c| {
        var listed = false;
        for (f1a2_fidelity_cases) |d| {
            if (std.mem.eql(u8, d.formula, c.formula)) listed = true;
        }
        if (listed) continue;

        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        try putF1a2Cells(&h);

        var excel_opts = h.options();
        excel_opts.fidelity = .excel;
        const a = value.publish((try h.evalOpts(c.formula, excel_opts)).scalar, .excel);
        var ieee_opts = h.options();
        ieee_opts.fidelity = .ieee;
        const b = value.publish((try h.evalOpts(c.formula, ieee_opts)).scalar, .ieee);
        if (!value.PublishedScalar.eql(a, b)) {
            std.debug.print("`{s}` diverges between modes but is not listed\n", .{c.formula});
            return error.UnlistedFidelityDivergence;
        }
    }
}

test "M4d: POWER is the function spelling of `^`, and answers identically" {
    // A workbook must not get two answers for one operation. Stated over
    // the cases most likely to separate them — a negative base, a
    // fractional exponent, an overflow, and `0^0`, which Excel and
    // LibreOffice do not agree about and no committed manifest records.
    // Whatever the operator answers there, POWER answers; changing it is
    // an operator-level decision, and M4d has no evidence to make one.
    const pairs = [_][2][]const u8{
        .{ "POWER(2,10)", "2^10" },
        .{ "POWER(-2,3)", "(-2)^3" },
        .{ "POWER(2,0.5)", "2^0.5" },
        .{ "POWER(0,0)", "0^0" },
        .{ "POWER(-8,1/3)", "(-8)^(1/3)" },
        .{ "POWER(2,1024)", "2^1024" },
        .{ "POWER(10,-2)", "10^-2" },
        .{ "POWER(A1,2)", "A1^2" },
    };
    for (pairs) |pair| {
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        try putF1a2Cells(&h);

        inline for ([_]value.Fidelity{ .excel, .ieee }) |mode| {
            var opts = h.options();
            opts.fidelity = mode;
            const fn_v = try h.evalOpts(pair[0], opts);
            const op_v = try h.evalOpts(pair[1], opts);
            const a = value.publish(fn_v.scalar, mode);
            const b = value.publish(op_v.scalar, mode);
            if (!value.PublishedScalar.eql(a, b)) {
                std.debug.print("`{s}` and `{s}` disagree under {s}\n", .{ pair[0], pair[1], @tagName(mode) });
                return error.PowerOperatorMismatch;
            }
        }
    }
}

test "M4d: TRUNC(x) is ROUNDDOWN(x,0), and LOG(x) is LOG10(x)" {
    // Two pairs the table declares to be one operation under two
    // spellings — the optional-argument defaults. An equivalence a
    // reader would otherwise have to infer from two implementations.
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try putF1a2Cells(&h);

    const subjects = [_][]const u8{ "2.9", "-2.9", "0", "A1", "B2", "B3", "1234.5" };
    for (subjects) |s| {
        var one: [64]u8 = undefined;
        var two: [64]u8 = undefined;
        const a = try h.scalar(try std.fmt.bufPrint(&one, "TRUNC({s})", .{s}));
        const b = try h.scalar(try std.fmt.bufPrint(&two, "ROUNDDOWN({s},0)", .{s}));
        try testing.expect(a.eql(b));
    }
    for ([_][]const u8{ "100", "A1", "0.5", "1" }) |s| {
        var one: [64]u8 = undefined;
        var two: [64]u8 = undefined;
        const a = try h.scalar(try std.fmt.bufPrint(&one, "LOG({s})", .{s}));
        const b = try h.scalar(try std.fmt.bufPrint(&two, "LOG10({s})", .{s}));
        try testing.expect(a.eql(b));
    }
}

test "M4d: error order in every multi-argument name of the batch (§5.3c)" {
    // §5.3c: eager arguments evaluate in declaration order and the first
    // error wins. Every case runs in both argument orders, because a
    // fixture with one error in it proves propagation and says nothing
    // about order.
    const Case = struct { formula: []const u8, expect: value.KnownError };
    const cases = [_]Case{
        .{ .formula = "ROUND(A5,A6)", .expect = .div0 },
        .{ .formula = "ROUND(A6,A5)", .expect = .na },
        .{ .formula = "ROUNDUP(A5,A6)", .expect = .div0 },
        .{ .formula = "ROUNDUP(A6,A5)", .expect = .na },
        .{ .formula = "ROUNDDOWN(A5,A6)", .expect = .div0 },
        .{ .formula = "ROUNDDOWN(A6,A5)", .expect = .na },
        .{ .formula = "TRUNC(A5,A6)", .expect = .div0 },
        .{ .formula = "TRUNC(A6,A5)", .expect = .na },
        .{ .formula = "MOD(A5,A6)", .expect = .div0 },
        .{ .formula = "MOD(A6,A5)", .expect = .na },
        .{ .formula = "POWER(A5,A6)", .expect = .div0 },
        .{ .formula = "POWER(A6,A5)", .expect = .na },
        .{ .formula = "LOG(A5,A6)", .expect = .div0 },
        .{ .formula = "LOG(A6,A5)", .expect = .na },
        .{ .formula = "RANDBETWEEN(A5,A6)", .expect = .div0 },
        .{ .formula = "RANDBETWEEN(A6,A5)", .expect = .na },
        // Propagation runs BEFORE the implementation, so an argument's
        // error beats a domain failure the implementation would have
        // raised. Three functions, three domain failures, one answer.
        .{ .formula = "MOD(A6,0)", .expect = .na },
        .{ .formula = "LOG(A6,1)", .expect = .na },
        .{ .formula = "RANDBETWEEN(A6,-1)", .expect = .na },
        // …and from inside a range, in §5.6a's iteration order. A range
        // in a scalar slot lifts, so this is the per-element answer of
        // the first cell rather than a fold.
        .{ .formula = "MOD(A5,A5)", .expect = .div0 },
    };

    for (cases) |c| {
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        try putF1a2Cells(&h);

        const s = h.scalar(c.formula) catch |e| {
            std.debug.print("error-order case `{s}` refused: {t}\n", .{ c.formula, e });
            return e;
        };
        if (s != .err or s.err != .known or s.err.known != c.expect) {
            std.debug.print("error-order case `{s}`: expected {s}\n", .{ c.formula, c.expect.spelling() });
            return error.WrongErrorOrder;
        }
    }

    // Every multi-argument name in the batch appears above, with the
    // list derived from the registry's own arity rather than typed out —
    // so a function that gains an argument later cannot slip past
    // unordered.
    var it = registry.inventory();
    while (it.next()) |e| {
        if (!std.mem.eql(u8, e.milestone, "M4d")) continue;
        const f = registry.lookup(e.name).?;
        const multi = f.arity.max == null or f.arity.max.? > 1;
        if (!multi) continue;
        var covered = false;
        for (cases) |c| {
            if (std.mem.startsWith(u8, c.formula, e.name) and c.formula[e.name.len] == '(') covered = true;
        }
        if (!covered) {
            std.debug.print("multi-argument name with no error-order fixture: {s}\n", .{e.name});
            return error.MissingErrorOrderFixture;
        }
    }
}

// ─── M4d draw KATs (§5.6d: multi-callsite, and lazy branches) ────

test "M4d: draw counts — two RAND() in one formula, none in a dead branch" {
    // §5.6d's own words: two `RAND()` in one cell are distinct call
    // sites. The count is the instrument, because a result cannot tell a
    // draw that happened from one that did not — under a constant source
    // both look the same.
    const Case = struct { formula: []const u8, draws: u64 };
    const cases = [_]Case{
        .{ .formula = "RAND()", .draws = 1 },
        .{ .formula = "RAND()+RAND()", .draws = 2 },
        .{ .formula = "RAND()+RAND()+RAND()", .draws = 3 },
        .{ .formula = "RANDBETWEEN(1,6)", .draws = 1 },
        .{ .formula = "RANDBETWEEN(1,6)+RANDBETWEEN(1,6)", .draws = 2 },
        .{ .formula = "RAND()+RANDBETWEEN(1,6)", .draws = 2 },
        // One draw per CALL, not per argument: RANDBETWEEN reads two
        // arguments and draws once.
        .{ .formula = "RANDBETWEEN(A1,A1)", .draws = 1 },
        // The batch's fourteen stable names draw nothing, however
        // arithmetic they look.
        .{ .formula = "ABS(-1)+SQRT(4)+PI()+EXP(0)+SIGN(3)", .draws = 0 },
        .{ .formula = "ROUND(PI(),2)+MOD(7,3)+POWER(2,3)+LOG(8,2)", .draws = 0 },
        // A dead lazy branch draws ZERO — the property M4c proved for
        // the three deferring forms, re-proved here for the volatiles
        // this row adds.
        .{ .formula = "IF(TRUE(),1,RANDBETWEEN(1,6))", .draws = 0 },
        .{ .formula = "IF(FALSE(),RANDBETWEEN(1,6),1)", .draws = 0 },
        .{ .formula = "IF(TRUE(),RAND(),RANDBETWEEN(1,6))", .draws = 1 },
        .{ .formula = "IF(FALSE(),RAND(),RANDBETWEEN(1,6))", .draws = 1 },
        .{ .formula = "IFERROR(1,RANDBETWEEN(1,6))", .draws = 0 },
        .{ .formula = "IFNA(1,RAND()+RANDBETWEEN(1,6))", .draws = 0 },
        .{ .formula = "IFNA(A5,RANDBETWEEN(1,6))", .draws = 0 },
        // …and an eager form draws every arm it writes, which is the
        // same fact stated so that "optimizing" a short-circuit in would
        // fail rather than pass.
        .{ .formula = "IFS(TRUE(),RAND(),TRUE(),RANDBETWEEN(1,6))", .draws = 2 },
        .{ .formula = "SWITCH(1,1,RAND(),2,RANDBETWEEN(1,6))", .draws = 2 },
        // A propagated error reaches the implementation never, so the
        // volatile never draws either.
        .{ .formula = "RANDBETWEEN(A5,A6)", .draws = 0 },
        // A lift draws once per element, because each element is its own
        // invocation.
        .{ .formula = "RANDBETWEEN({1,1},{6,6})", .draws = 2 },
    };

    for (cases) |c| {
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        try putF1a2Cells(&h);

        _ = h.eval(c.formula) catch |e| {
            std.debug.print("draw case `{s}` refused: {t}\n", .{ c.formula, e });
            return e;
        };
        if (h.draws.count != c.draws) {
            std.debug.print(
                "`{s}`: expected {d} draws, saw {d}\n",
                .{ c.formula, c.draws, h.draws.count },
            );
            return error.WrongDrawCount;
        }
    }
}

test "M4d KAT: the draw sequence is reproducible from RunInputs alone" {
    // §5.6d's callsite-keyed schedule is M5a2's and needs a graph this
    // row does not have. What M4d can state — and must, because every
    // count above rests on it — is that the sequence a formula consumes
    // is a function of `RunInputs` and of nothing else: not a clock, not
    // an entropy source, not evaluation history.
    const inputs: run_inputs.RunInputs = .{ .now_utc_ms = 0, .rng_seed = 0x5EED_1A2, .limits = .{} };
    try inputs.validate();

    // The stream this seed names, taken directly from `rng_v1`. Every
    // expectation below is against these, so the KAT ties the EVALUATOR
    // to the generator rather than to itself.
    var direct = rng.Rng.fromRunInputs(inputs);
    const d0 = direct.nextFloat();
    const d1 = direct.nextFloat();
    try testing.expect(d0 != d1);

    {
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        var generator = rng.Rng.fromRunInputs(inputs);
        var source = generator.drawSource();
        var opts = h.options();
        opts.draws = &source;

        // Two evaluations of one formula, off one source: the draws come
        // out in order, and they are the generator's own first two.
        const a = try h.evalOpts("RAND()", opts);
        const b = try h.evalOpts("RAND()", opts);
        try testing.expectEqual(@as(u64, @bitCast(d0)), @as(u64, @bitCast(a.scalar.number)));
        try testing.expectEqual(@as(u64, @bitCast(d1)), @as(u64, @bitCast(b.scalar.number)));
        try testing.expectEqual(@as(u64, 2), source.count);
    }

    {
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        var generator = rng.Rng.fromRunInputs(inputs);
        var source = generator.drawSource();
        var opts = h.options();
        opts.draws = &source;

        // The dead arm does not merely produce nothing — it CONSUMES
        // nothing, so the live arm receives the stream's first value. A
        // draw that happened and was discarded would satisfy a count on
        // the result and fail this.
        const v = try h.evalOpts("IF(TRUE(),RAND(),RAND())", opts);
        try testing.expectEqual(@as(u64, @bitCast(d0)), @as(u64, @bitCast(v.scalar.number)));
        try testing.expectEqual(@as(u64, 1), source.count);
    }

    {
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        var generator = rng.Rng.fromRunInputs(inputs);
        var source = generator.drawSource();
        var opts = h.options();
        opts.draws = &source;

        // Two call sites in ONE formula draw twice, and get different
        // numbers. Subtraction rather than addition: a sum cannot tell
        // `0.3 + 0.7` from `0.5 + 0.5`, and the property under test is
        // that the two sites were handed different values.
        const v = try h.evalOpts("RAND()-RAND()", opts);
        try testing.expectEqual(@as(u64, 2), source.count);
        try testing.expect(v.scalar.number != 0);
    }

    // Same inputs, same answer; different seed, different answer. The
    // second half is what makes the first mean "derived from the seed"
    // rather than "constant".
    const again = try seededResult(inputs, "RAND()-RAND()");
    const first = try seededResult(inputs, "RAND()-RAND()");
    try testing.expectEqual(@as(u64, @bitCast(first)), @as(u64, @bitCast(again)));
    var other = inputs;
    other.rng_seed = inputs.rng_seed +% 1;
    const different = try seededResult(other, "RAND()-RAND()");
    try testing.expect(@as(u64, @bitCast(different)) != @as(u64, @bitCast(first)));

    // RANDBETWEEN rides the same seam, so its result is a function of
    // the same inputs — and stays inside the range it was given.
    const roll = try seededResult(inputs, "RANDBETWEEN(1,6)");
    try testing.expectEqual(@as(f64, @floor(1 + d0 * 6)), roll);
    try testing.expect(roll >= 1 and roll <= 6);
}

/// Evaluate one formula against a generator built from `inputs` and
/// nothing else, and hand back the number. The whole point of the helper
/// is that it takes no other parameter: reproducibility is a property of
/// the argument list.
fn seededResult(inputs: run_inputs.RunInputs, formula: []const u8) !f64 {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    var generator = rng.Rng.fromRunInputs(inputs);
    var source = generator.drawSource();
    var opts = h.options();
    opts.draws = &source;
    const v = try h.evalOpts(formula, opts);
    return v.scalar.number;
}

test "M4d: RANDBETWEEN stays inside its range across the whole draw interval" {
    // The mapping from a `[0,1)` draw to an inclusive integer range is
    // the one place a single draw can leave the range: a draw just under
    // 1 against a wide span. Swept rather than sampled, over the ranges
    // whose ends are the ones that go wrong.
    const Range = struct { lo: []const u8, hi: []const u8, lo_n: f64, hi_n: f64 };
    const ranges = [_]Range{
        .{ .lo = "1", .hi = "6", .lo_n = 1, .hi_n = 6 },
        .{ .lo = "0", .hi = "1", .lo_n = 0, .hi_n = 1 },
        .{ .lo = "-5", .hi = "5", .lo_n = -5, .hi_n = 5 },
        .{ .lo = "-1", .hi = "-1", .lo_n = -1, .hi_n = -1 },
        .{ .lo = "1", .hi = "1000000000", .lo_n = 1, .hi_n = 1000000000 },
    };
    // Including the largest double strictly below 1, which is what
    // `nextFloat` can actually return and what the `@min` clamp exists
    // for.
    const draws = [_]f64{ 0, 0.5, 0.999999, 1 - 0x1.0p-53 };

    for (ranges) |r| {
        for (draws) |d| {
            var h: Harness = undefined;
            try h.init(testing.allocator);
            defer h.deinit();
            h.draw_value = d;

            var buf: [64]u8 = undefined;
            const src = try std.fmt.bufPrint(&buf, "RANDBETWEEN({s},{s})", .{ r.lo, r.hi });
            const s = try h.scalar(src);
            try testing.expect(s == .number);
            if (s.number < r.lo_n or s.number > r.hi_n) {
                std.debug.print("`{s}` with draw {d} produced {d}\n", .{ src, d, s.number });
                return error.RandBetweenOutOfRange;
            }
            // …and it is an integer, which is the other half of the
            // function's name.
            try testing.expectEqual(s.number, @floor(s.number));
        }
    }
}

// ─── M4d fuzz: no argument shape panics, leaks, or evaluates twice ──

const f1a2_names = [_][]const u8{
    "ABS",       "EXP",     "INT",   "LN",   "LOG",         "LOG10",
    "MOD",       "PI",      "POWER", "RAND", "RANDBETWEEN", "ROUND",
    "ROUNDDOWN", "ROUNDUP", "SIGN",  "SQRT", "TRUNC",
};

/// Argument *shapes* and magnitudes. F1a-1's alphabet was about
/// provenance; this one adds the numeric extremes, because the failure
/// mode a numeric batch has and a predicate batch does not is a result
/// binary64 cannot hold — and `assertRepresentable` is watching for
/// exactly that.
const f1a2_arg_shapes = [_][]const u8{
    "A1",      "A2",            "A5",     "A7",
    "A8",      "0",             "-0",     "1",
    "-1",      "2.5",           "-2.5",   "0.5",
    "308",     "-308",          "400",    "-400",
    "1E+308",  "-1E+308",       "1E-308", "2.2250738585072014E-308",
    "1/0",     "1E+308*10",     "{1,2}",  "{1;2}",
    "A1:A8",   "(A1:A2,A5:A6)", "",       "@A1:A8",
    "\"1,5\"", "S!A1",          "PI()",   "16",
};

fn fuzzF1a2Env(fake: *env.Fake) !env.SheetIndex {
    const sheet = try fake.addSheet("S");
    try fake.putA1(sheet, .stored, "A1", num(10));
    try fake.putA1(sheet, .stored, "A2", .{ .text = "abc" });
    try fake.putA1(sheet, .stored, "A5", value.ScalarValue.errorOf(.div0));
    try fake.putA1(sheet, .stored, "A6", value.ScalarValue.errorOf(.na));
    try fake.putA1(sheet, .stored, "A8", .{ .text = "7" });
    return sheet;
}

fn fuzzF1a2Target(_: void, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();

    var buf: [512]u8 = undefined;
    var w: usize = 0;
    const name = f1a2_names[smith.index(f1a2_names.len)];
    @memcpy(buf[w..][0..name.len], name);
    w += name.len;
    buf[w] = '(';
    w += 1;
    var n: usize = 0;
    while (n < 4 and !smith.eos()) : (n += 1) {
        const arg = f1a2_arg_shapes[smith.index(f1a2_arg_shapes.len)];
        if (w + arg.len + 2 > buf.len) break;
        if (n > 0) {
            buf[w] = ',';
            w += 1;
        }
        @memcpy(buf[w..][0..arg.len], arg);
        w += arg.len;
    }
    buf[w] = ')';
    w += 1;
    const src = buf[0..w];

    var parsed = parser.parse(std.testing.allocator, src, .{}) catch return;
    defer parsed.deinit(std.testing.allocator);
    if (parsed == .refused) return;

    var arena_state = std.heap.ArenaAllocator.init(std.testing.allocator);
    defer arena_state.deinit();
    var fake = env.Fake.init(std.testing.allocator);
    defer fake.deinit();
    const sheet = try fuzzF1a2Env(&fake);

    // Both rule tables, because a numeric batch's answer may depend on
    // which one is in force and "no shape produces a non-finite number"
    // has to hold under each.
    for ([_]value.Fidelity{ .excel, .ieee }) |mode| {
        var draw_value: f64 = 0.5;
        var draws = DrawSource.constant(&draw_value);
        const opts: Options = .{
            .current_sheet = sheet,
            .collation = .{ .fold = shippedFold },
            .draws = &draws,
            .fidelity = mode,
            .site = .{
                .row = coords.Row.fromOneBased(2) catch unreachable,
                .col = coords.Col.fromZeroBased(1) catch unreachable,
            },
        };

        var first = Evaluator.init(arena_state.allocator(), fake.evalEnv(), opts);
        defer first.deinit();
        var second = Evaluator.init(arena_state.allocator(), fake.evalEnv(), opts);
        defer second.deinit();

        if (first.evaluate(parsed.ok)) |a| {
            try assertRepresentable(a);
            const b = second.evaluate(parsed.ok) catch return error.NondeterministicEvaluation;
            try assertRepresentable(b);
            if (!valuesAgree(a, b)) {
                std.debug.print("`{s}` evaluated two ways\n", .{src});
                return error.NondeterministicEvaluation;
            }
        } else |e| {
            if (second.evaluate(parsed.ok)) |_| {
                return error.NondeterministicEvaluation;
            } else |e2| {
                if (e != e2) return error.NondeterministicEvaluation;
            }
        }
    }
}

test "fuzz: no F1a-2 argument shape panics, leaks, or evaluates two ways" {
    // The generator's alphabet is a second copy of the batch, so it gets
    // the same treatment as every other copy in this row: checked
    // against the file rather than maintained by hand.
    var it = registry.inventory();
    var batch: usize = 0;
    while (it.next()) |e| {
        if (!std.mem.eql(u8, e.milestone, "M4d")) continue;
        batch += 1;
        var present = false;
        for (f1a2_names) |n| {
            if (std.mem.eql(u8, n, e.name)) present = true;
        }
        if (!present) {
            std.debug.print("fuzz alphabet is missing {s}\n", .{e.name});
            return error.FuzzAlphabetIncomplete;
        }
    }
    try testing.expectEqual(batch, f1a2_names.len);

    try std.testing.fuzz({}, fuzzF1a2Target, .{
        .corpus = &[_][]const u8{
            "ABS(-1E+308)",
            "SIGN(-0)",
            "INT(A1:A8)",
            "ROUND(A1,400)",
            "ROUND(A1,-400)",
            "ROUNDUP(A1,-400)",
            "ROUNDDOWN(1E-308,308)",
            "TRUNC(2.2250738585072014E-308,400)",
            "MOD(1E+308,1E-308)",
            "MOD(A1,0)",
            "POWER(1E+308,2)",
            "POWER(-1,0.5)",
            "EXP(1E+308)",
            "LN(0)",
            "LOG(A1,1)",
            "LOG(0,0)",
            "LOG10(-1)",
            "SQRT(-1)",
            "PI()",
            "RAND()",
            "RANDBETWEEN(1E+308,-1E+308)",
            "RANDBETWEEN(-1E+308,1E+308)",
            "RANDBETWEEN({1,2},{3,4})",
        },
    });
}

test "M4d: every name against every argument shape, exhaustively and in both modes" {
    // The `fuzz` step is Linux-only (`build.zig:249` — coverage-guided
    // fuzzing is broken upstream on macOS and Windows), so on the other
    // two platforms the target above runs its corpus and nothing else.
    // The property it exists to prove — no shape panics, leaks, produces
    // a non-finite number, or evaluates two ways — is small enough here
    // to prove by ENUMERATION instead of by search: seventeen names, one
    // and two arguments, every shape in the alphabet, both rule tables.
    // A sweep that always runs beats a search that runs on one platform.
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    var fake = env.Fake.init(testing.allocator);
    defer fake.deinit();
    const sheet = try fuzzF1a2Env(&fake);

    var checked: usize = 0;
    var buf: [256]u8 = undefined;
    for (f1a2_names) |name| {
        for (f1a2_arg_shapes) |first_arg| {
            // The one-argument form belongs to `(name, first_arg)` and is
            // built here rather than inside the pair loop — thirty-two
            // identical rebuilds would have cost half the sweep's time
            // and covered nothing the first one did not.
            const one = std.fmt.bufPrint(&buf, "{s}({s})", .{ name, first_arg }) catch continue;
            try sweepShape(&arena_state, &fake, sheet, one, &checked);
            for (f1a2_arg_shapes) |second_arg| {
                var pair_buf: [256]u8 = undefined;
                const two = std.fmt.bufPrint(&pair_buf, "{s}({s},{s})", .{ name, first_arg, second_arg }) catch continue;
                try sweepShape(&arena_state, &fake, sheet, two, &checked);
            }
        }
    }
    // A sweep that silently stopped enumerating would still pass, so the
    // count is asserted as a floor rather than left to the loops.
    try testing.expect(checked > 10_000);
}

/// One swept call: parse it, evaluate it twice under each rule table,
/// and hold the four results to the same contract the fuzz target holds
/// its inputs to. A malformed shape or an arity the registry rejects is
/// not a finding — the property is about calls that *reach* an
/// implementation.
fn sweepShape(
    arena_state: *std.heap.ArenaAllocator,
    fake: *env.Fake,
    sheet: env.SheetIndex,
    src: []const u8,
    checked: *usize,
) !void {
    var parsed = parser.parse(testing.allocator, src, .{}) catch return;
    defer parsed.deinit(testing.allocator);
    if (parsed == .refused) return;

    for ([_]value.Fidelity{ .excel, .ieee }) |mode| {
        var draw_value: f64 = 0.5;
        var draws = DrawSource.constant(&draw_value);
        const opts: Options = .{
            .current_sheet = sheet,
            .collation = .{ .fold = shippedFold },
            .draws = &draws,
            .fidelity = mode,
            .site = .{
                .row = coords.Row.fromOneBased(2) catch unreachable,
                .col = coords.Col.fromZeroBased(1) catch unreachable,
            },
        };
        var first = Evaluator.init(arena_state.allocator(), fake.evalEnv(), opts);
        defer first.deinit();
        var second = Evaluator.init(arena_state.allocator(), fake.evalEnv(), opts);
        defer second.deinit();

        if (first.evaluate(parsed.ok)) |a| {
            assertRepresentable(a) catch |e| {
                std.debug.print("`{s}` under {s} produced an unrepresentable value\n", .{ src, @tagName(mode) });
                return e;
            };
            const b = second.evaluate(parsed.ok) catch {
                std.debug.print("`{s}` succeeded then refused\n", .{src});
                return error.NondeterministicEvaluation;
            };
            try assertRepresentable(b);
            if (!valuesAgree(a, b)) {
                std.debug.print("`{s}` evaluated two ways\n", .{src});
                return error.NondeterministicEvaluation;
            }
        } else |e| {
            if (second.evaluate(parsed.ok)) |_| {
                std.debug.print("`{s}` refused then succeeded\n", .{src});
                return error.NondeterministicEvaluation;
            } else |e2| {
                if (e != e2) return error.NondeterministicEvaluation;
            }
        }
        checked.* += 1;
    }
}

// ─── M4e: the F1b batch (§7, twenty-two names) ───────────────────
//
// Oracle-first, and the oracle decides **nothing** here. Not one of the
// three committed manifests contains a formula calling any of the
// twenty-two — §8.2's evidence is eighteen operator and literal cells
// plus `SQRT(-1)`, and no aggregate, criteria, lookup or position name
// appears in any of them. So the whole batch ships `spec_pinned`, its
// oracle-row count is pinned at **zero**, and the label is checked
// against the files in both directions exactly as M4c and M4d checked
// theirs. A row claiming evidence no manifest holds fails; so does a
// row shipping spec-pinned that a manifest decides; and so does a row
// reading an EXCLUDED cell as evidence, which is M4d's third verdict
// (decision 2) applied to a batch that needs it for a different reason
// — this one holds no volatile, and the check is still three-valued
// because the checker is shared rather than re-derived per row.

/// One F1b fixture. `func` is the inventory name the row is a fixture
/// FOR — the coverage test derives the batch from the frozen TSV and
/// fails if any of the twenty-two has no row here.
const F1bCase = struct {
    func: []const u8,
    formula: []const u8,
    expect: Expect,
    evidence: Evidence = .spec_pinned,
    /// Why the spec says so, where the answer is one a reasonable
    /// reading could get wrong.
    note: []const u8 = "",
};

/// The environment every F1b fixture reads.
///
/// The A column repeats M4c's and M4d's provenance rows cell for cell,
/// so the three batches can be compared against the same §5.3b table.
/// Everything else is shaped for this batch: one column to aggregate,
/// one range holding every population the three counting functions
/// disagree about, and the tables the five lookups search.
fn putF1bCells(h: *Harness) !void {
    try h.put("A1", num(10)); // number
    try h.put("A2", .{ .text = "abc" }); // non-numeric text
    try h.put("A3", .{ .text = "" }); // `=""`, which is text and not blank
    try h.put("A4", .{ .boolean = true }); // logical
    try h.put("A5", value.ScalarValue.errorOf(.div0)); // an error that is not #N/A
    try h.put("A6", value.ScalarValue.errorOf(.na)); // #N/A itself
    // A7 is a true blank — deliberately not stored.
    try h.put("A8", .{ .text = "7" }); // numeric text

    // A plain numeric column: what an aggregate is for.
    try h.put("C1", num(1));
    try h.put("C2", num(2));
    try h.put("C3", num(3));
    try h.put("C4", num(4));

    // §5.3c's spine, as ONE range. Every population the three counting
    // functions disagree about is here, so all three can be fixtured
    // side by side over the same cells rather than each over a range
    // built to flatter it.
    try h.put("D1", num(1)); // a number
    try h.put("D2", .{ .text = "x" }); // text
    try h.put("D3", value.ScalarValue.errorOf(.div0)); // an error
    // D4 is a true blank.
    try h.put("D5", .{ .text = "" }); // `=""`

    // A negative zero, **stored** rather than written as the literal
    // `-0`: N1a rounds a literal at ingress under `.excel`, and the
    // divergence under test has to be the function's rather than the
    // parser's (M4d decision 5).
    try h.put("E1", num(-0.0));

    // The lookup table: ascending keys down F, two result columns.
    try h.put("F1", num(10));
    try h.put("F2", num(20));
    try h.put("F3", num(30));
    try h.put("G1", .{ .text = "ten" });
    try h.put("G2", .{ .text = "twenty" });
    try h.put("G3", .{ .text = "thirty" });
    try h.put("H1", num(100));
    try h.put("H2", num(200));
    try h.put("H3", num(300));

    // The same table transposed, for HLOOKUP — one function under two
    // axes, so it gets the same data laid out the other way.
    try h.put("F5", num(10));
    try h.put("G5", num(20));
    try h.put("H5", num(30));
    try h.put("F6", .{ .text = "ten" });
    try h.put("G6", .{ .text = "twenty" });
    try h.put("H6", .{ .text = "thirty" });

    // Fold-equality's own table: `ß` folds to `ss`, so `STRASSE` and
    // `Straße` are ONE key under `collation_v1` (§5.4b).
    try h.put("J1", .{ .text = "Straße" });
    try h.put("J2", .{ .text = "beta" });
    try h.put("K1", num(1));
    try h.put("K2", num(2));

    // Descending, for `MATCH`'s type −1.
    try h.put("N1", num(30));
    try h.put("N2", num(20));
    try h.put("N3", num(10));

    // Duplicates, for the two search orders.
    try h.put("P1", num(1));
    try h.put("P2", num(2));
    try h.put("P3", num(2));
    try h.put("P4", num(3));
}

const f1b_cases = [_]F1bCase{
    // ── the aggregates ──
    .{ .func = "SUM", .formula = "SUM(C1:C4)", .expect = .{ .number = 10 } },
    .{ .func = "SUM", .formula = "SUM(D1:D5)", .expect = .{ .err = .div0 }, .note = "SUM propagates an error found in a range; COUNT beside it does not" },
    .{ .func = "SUM", .formula = "SUM(TRUE())", .expect = .{ .number = 1 }, .note = "a DIRECT logical coerces" },
    .{ .func = "SUM", .formula = "SUM(A4)", .expect = .{ .number = 0 }, .note = "…and the same logical found in a range is ignored — §5.3b's split, one line apart" },
    .{ .func = "AVERAGE", .formula = "AVERAGE(C1:C4)", .expect = .{ .number = 2.5 } },
    .{ .func = "AVERAGE", .formula = "AVERAGE(1,2,3)", .expect = .{ .number = 2 } },
    .{ .func = "AVERAGE", .formula = "AVERAGE(1,)", .expect = .{ .number = 0.5 }, .note = "an omitted argument is 0 AND is counted, so the denominator is 2" },
    .{ .func = "AVERAGE", .formula = "AVERAGE(A2)", .expect = .{ .err = .div0 }, .note = "text in a range is ignored, and an average over nothing is a division by zero" },
    .{ .func = "AVERAGE", .formula = "AVERAGE(A6)", .expect = .{ .err = .na } },
    .{ .func = "AVERAGE", .formula = "AVERAGE(\"abc\")", .expect = .{ .err = .value }, .note = "a DIRECT text argument is #VALUE! where the same text in a range is skipped" },
    .{ .func = "MIN", .formula = "MIN(C1:C4)", .expect = .{ .number = 1 } },
    .{ .func = "MIN", .formula = "MIN(A2)", .expect = .{ .number = 0 }, .note = "no numbers anywhere is 0 — not #DIV/0!, and not #VALUE!" },
    .{ .func = "MIN", .formula = "MIN(J1:J2)", .expect = .{ .number = 0 }, .note = "MIN never compares text (§5.4b): a column of it answers 0" },
    .{ .func = "MIN", .formula = "MIN(\"abc\")", .expect = .{ .err = .value } },
    .{ .func = "MIN", .formula = "MIN(A6)", .expect = .{ .err = .na } },
    .{ .func = "MAX", .formula = "MAX(C1:C4)", .expect = .{ .number = 4 } },
    .{ .func = "MAX", .formula = "MAX(-5,-3)", .expect = .{ .number = -3 } },
    .{ .func = "MAX", .formula = "MAX(A4)", .expect = .{ .number = 0 } },
    .{ .func = "MAX", .formula = "MAX(TRUE())", .expect = .{ .number = 1 } },
    .{ .func = "SUMPRODUCT", .formula = "SUMPRODUCT(C1:C4,C1:C4)", .expect = .{ .number = 30 } },
    .{ .func = "SUMPRODUCT", .formula = "SUMPRODUCT(C1:C4)", .expect = .{ .number = 10 }, .note = "one array is a sum" },
    .{ .func = "SUMPRODUCT", .formula = "SUMPRODUCT({1,2},{3,4})", .expect = .{ .number = 11 } },
    .{ .func = "SUMPRODUCT", .formula = "SUMPRODUCT(C1:C4,F1:F3)", .expect = .{ .err = .value }, .note = "identical dimensions required — SUMPRODUCT does not broadcast" },
    .{ .func = "SUMPRODUCT", .formula = "SUMPRODUCT(A4)", .expect = .{ .number = 0 }, .note = "a logical contributes 0, where SUM(TRUE()) is 1" },
    .{ .func = "SUMPRODUCT", .formula = "SUMPRODUCT(D1:D5)", .expect = .{ .err = .div0 } },

    // ── the three-way COUNT split (§5.3c), over the one range ──
    .{ .func = "COUNT", .formula = "COUNT(D1:D5)", .expect = .{ .number = 1 }, .note = "numbers only; the error is neither counted nor propagated" },
    .{ .func = "COUNT", .formula = "COUNT(\"1\")", .expect = .{ .number = 1 }, .note = "a direct argument coerces" },
    .{ .func = "COUNT", .formula = "COUNT(A8)", .expect = .{ .number = 0 }, .note = "…and the same numeric text in a range is NOT coerced" },
    .{ .func = "COUNTA", .formula = "COUNTA(D1:D5)", .expect = .{ .number = 4 }, .note = "everything that is not a true blank — the error counts, and so does `\"\"`" },
    .{ .func = "COUNTA", .formula = "COUNTA(A7)", .expect = .{ .number = 0 } },
    .{ .func = "COUNTBLANK", .formula = "COUNTBLANK(D1:D5)", .expect = .{ .number = 2 }, .note = "true blanks PLUS `\"\"` — the third question, and the `.countblank_class` population" },

    // ── criteria ──
    .{ .func = "COUNTIF", .formula = "COUNTIF(C1:C4,\">2\")", .expect = .{ .number = 2 } },
    .{ .func = "COUNTIF", .formula = "COUNTIF(D1:D5,\"#DIV/0!\")", .expect = .{ .number = 1 }, .note = "criteria CAN match errors, which is a third answer again" },
    .{ .func = "COUNTIF", .formula = "COUNTIF(D1:D5,\"\")", .expect = .{ .number = 2 }, .note = "an empty criterion is the COUNTBLANK population, not the COUNTA one" },
    .{ .func = "SUMIF", .formula = "SUMIF(C1:C4,\">2\")", .expect = .{ .number = 7 } },
    .{ .func = "SUMIF", .formula = "SUMIF(F1:F3,\">=20\",H1:H3)", .expect = .{ .number = 500 } },
    .{ .func = "AVERAGEIF", .formula = "AVERAGEIF(C1:C4,\">2\")", .expect = .{ .number = 3.5 } },
    .{ .func = "AVERAGEIF", .formula = "AVERAGEIF(F1:F3,\">=20\",H1:H3)", .expect = .{ .number = 250 }, .note = "the average range is PROJECTED from its top-left (§5.6a), as SUMIF's is" },
    .{ .func = "AVERAGEIF", .formula = "AVERAGEIF(C1:C4,\">99\")", .expect = .{ .err = .div0 }, .note = "where SUMIF answers 0, an average over no match is #DIV/0!" },

    // ── VLOOKUP / HLOOKUP: one function under two axes ──
    .{ .func = "VLOOKUP", .formula = "VLOOKUP(20,F1:H3,2,FALSE)", .expect = .{ .text = "twenty" } },
    .{ .func = "VLOOKUP", .formula = "VLOOKUP(30,F1:H3,3)", .expect = .{ .number = 300 } },
    .{ .func = "VLOOKUP", .formula = "VLOOKUP(25,F1:H3,2)", .expect = .{ .text = "twenty" }, .note = "the omitted fourth argument is TRUE — approximate is Excel's default" },
    .{ .func = "VLOOKUP", .formula = "VLOOKUP(25,F1:H3,2,FALSE)", .expect = .{ .err = .na }, .note = "…and the same key exactly is #N/A" },
    .{ .func = "VLOOKUP", .formula = "VLOOKUP(5,F1:H3,2)", .expect = .{ .err = .na }, .note = "nothing is ≤ 5, so approximate has no candidate either" },
    .{ .func = "VLOOKUP", .formula = "VLOOKUP(20,F1:H3,4)", .expect = .{ .err = .ref }, .note = "past the table is #REF!" },
    .{ .func = "VLOOKUP", .formula = "VLOOKUP(20,F1:H3,0)", .expect = .{ .err = .value }, .note = "…and below its first column is #VALUE!. Two mistakes, two spellings" },
    .{ .func = "VLOOKUP", .formula = "VLOOKUP(\"STRASSE\",J1:K2,2,FALSE)", .expect = .{ .number = 1 }, .note = "fold-equal is EQUAL: `ß` folds to `ss`, so this is one key (§5.4b)" },
    .{ .func = "HLOOKUP", .formula = "HLOOKUP(20,F5:H6,2,FALSE)", .expect = .{ .text = "twenty" } },
    .{ .func = "HLOOKUP", .formula = "HLOOKUP(25,F5:H6,2)", .expect = .{ .text = "twenty" } },
    .{ .func = "HLOOKUP", .formula = "HLOOKUP(5,F5:H6,2)", .expect = .{ .err = .na } },
    .{ .func = "HLOOKUP", .formula = "HLOOKUP(20,F5:H6,3)", .expect = .{ .err = .ref } },

    // ── MATCH / XMATCH: the position, not the value ──
    .{ .func = "MATCH", .formula = "MATCH(20,F1:F3,0)", .expect = .{ .number = 2 } },
    .{ .func = "MATCH", .formula = "MATCH(25,F1:F3,1)", .expect = .{ .number = 2 }, .note = "type 1 is the largest value ≤ the key, over an ascending vector" },
    .{ .func = "MATCH", .formula = "MATCH(5,F1:F3,1)", .expect = .{ .err = .na } },
    .{ .func = "MATCH", .formula = "MATCH(25,N1:N3,-1)", .expect = .{ .number = 1 }, .note = "type −1 is the smallest value ≥ the key, over a DESCENDING vector" },
    .{ .func = "MATCH", .formula = "MATCH(35,N1:N3,-1)", .expect = .{ .err = .na } },
    .{ .func = "MATCH", .formula = "MATCH(\"t*n\",G1:G3,0)", .expect = .{ .number = 1 }, .note = "exact match honours `*` and `?` — the criteria matcher's wildcards, shared" },
    .{ .func = "MATCH", .formula = "MATCH(\"STRASSE\",J1:J2,0)", .expect = .{ .number = 1 } },
    .{ .func = "MATCH", .formula = "MATCH(20,F1:H3,0)", .expect = .{ .err = .na }, .note = "a 2-D array has no single axis to answer a position along" },
    .{ .func = "XMATCH", .formula = "XMATCH(20,F1:F3)", .expect = .{ .number = 2 }, .note = "XMATCH's default is EXACT where MATCH's is approximate" },
    .{ .func = "XMATCH", .formula = "XMATCH(25,F1:F3,-1)", .expect = .{ .number = 2 }, .note = "mode −1 is MATCH's type 1: the signs are opposite between the two names" },
    .{ .func = "XMATCH", .formula = "XMATCH(25,F1:F3,1)", .expect = .{ .number = 3 } },
    .{ .func = "XMATCH", .formula = "XMATCH(25,F1:F3,0)", .expect = .{ .err = .na } },
    .{ .func = "XMATCH", .formula = "XMATCH(\"t*\",G1:G3,0)", .expect = .{ .err = .na }, .note = "mode 0 makes the wildcards INERT — XMATCH asks for them by name in mode 2" },
    .{ .func = "XMATCH", .formula = "XMATCH(\"t*\",G1:G3,2)", .expect = .{ .number = 1 } },
    .{ .func = "XMATCH", .formula = "XMATCH(2,P1:P4,0,1)", .expect = .{ .number = 2 }, .note = "first-to-last finds the first of two equal keys…" },
    .{ .func = "XMATCH", .formula = "XMATCH(2,P1:P4,0,-1)", .expect = .{ .number = 3 }, .note = "…and last-to-first finds the second" },
    .{ .func = "XMATCH", .formula = "XMATCH(20,F1:F3,3)", .expect = .{ .err = .value }, .note = "an unlisted mode is #VALUE!, not the nearest listed one" },

    // ── XLOOKUP ──
    .{ .func = "XLOOKUP", .formula = "XLOOKUP(20,F1:F3,G1:G3)", .expect = .{ .text = "twenty" } },
    .{ .func = "XLOOKUP", .formula = "XLOOKUP(99,F1:F3,G1:G3)", .expect = .{ .err = .na } },
    .{ .func = "XLOOKUP", .formula = "XLOOKUP(99,F1:F3,G1:G3,\"none\")", .expect = .{ .text = "none" }, .note = "if_not_found replaces the #N/A of a failed MATCH, and nothing else" },
    .{ .func = "XLOOKUP", .formula = "XLOOKUP(25,F1:F3,G1:G3,\"none\",-1)", .expect = .{ .text = "twenty" } },
    .{ .func = "XLOOKUP", .formula = "XLOOKUP(25,F1:F3,G1:G3,\"none\",1)", .expect = .{ .text = "thirty" } },
    .{ .func = "XLOOKUP", .formula = "XLOOKUP(20,F1:F3,G1:G2)", .expect = .{ .err = .value }, .note = "the return range is indexed along the lookup vector's axis and must match it" },

    // ── INDEX ──
    .{ .func = "INDEX", .formula = "INDEX(F1:H3,2,2)", .expect = .{ .text = "twenty" } },
    .{ .func = "INDEX", .formula = "INDEX(F1:F3,2)", .expect = .{ .number = 20 }, .note = "one index runs along a vector's own axis" },
    .{ .func = "INDEX", .formula = "INDEX({10,20,30},2)", .expect = .{ .number = 20 } },
    .{ .func = "INDEX", .formula = "INDEX(F1:H3,4,1)", .expect = .{ .err = .ref } },
    .{ .func = "INDEX", .formula = "INDEX(F1:H3,1,4)", .expect = .{ .err = .ref } },
    .{ .func = "INDEX", .formula = "INDEX(F1:H3,-1,1)", .expect = .{ .err = .value }, .note = "past the array is #REF!, before it is #VALUE! — VLOOKUP's split again" },

    // ── CHOOSE ──
    .{ .func = "CHOOSE", .formula = "CHOOSE(2,10,20,30)", .expect = .{ .number = 20 } },
    .{ .func = "CHOOSE", .formula = "CHOOSE(1,10,20)", .expect = .{ .number = 10 } },
    .{ .func = "CHOOSE", .formula = "CHOOSE(1.9,10,20)", .expect = .{ .number = 10 }, .note = "the index truncates toward zero before the bound is checked" },
    .{ .func = "CHOOSE", .formula = "CHOOSE(0,10)", .expect = .{ .err = .value } },
    .{ .func = "CHOOSE", .formula = "CHOOSE(2,10)", .expect = .{ .err = .value } },

    // ── position: what a reference IS ──
    .{ .func = "ROW", .formula = "ROW(F2)", .expect = .{ .number = 2 } },
    .{ .func = "ROW", .formula = "ROW(F1:H3)", .expect = .{ .number = 1 }, .note = "an area answers from its top-left; spilling the whole column is M7a's" },
    .{ .func = "ROW", .formula = "ROW(1)", .expect = .{ .err = .value }, .note = "the `.reference` coercion class refuses a value before the impl runs" },
    .{ .func = "COLUMN", .formula = "COLUMN(F2)", .expect = .{ .number = 6 } },
    .{ .func = "COLUMN", .formula = "COLUMN(F1:H3)", .expect = .{ .number = 6 } },
    .{ .func = "ROWS", .formula = "ROWS(F1:H3)", .expect = .{ .number = 3 } },
    .{ .func = "ROWS", .formula = "ROWS(A1)", .expect = .{ .number = 1 } },
    .{ .func = "ROWS", .formula = "ROWS({1,2;3,4})", .expect = .{ .number = 2 } },
    .{ .func = "COLUMNS", .formula = "COLUMNS(F1:H3)", .expect = .{ .number = 3 } },
    .{ .func = "COLUMNS", .formula = "COLUMNS({1,2;3,4})", .expect = .{ .number = 2 } },
};

test "M4e: every F1b fixture evaluates to what the oracle or the spec says" {
    for (f1b_cases) |c| {
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        try putF1bCells(&h);

        const v = h.eval(c.formula) catch |e| {
            std.debug.print("F1b `{s}` refused: {t}\n", .{ c.formula, e });
            return e;
        };
        expectValue(c.expect, v) catch |e| {
            std.debug.print("F1b `{s}` ({s}): wrong value\n", .{ c.formula, c.func });
            return e;
        };
    }
}

test "M4e: all twenty-two frozen names resolve, and each has a fixture" {
    var it = registry.inventory();
    var batch: usize = 0;
    while (it.next()) |e| {
        if (!std.mem.eql(u8, e.milestone, "M4e")) continue;
        batch += 1;

        if (registry.lookup(e.name) == null) {
            std.debug.print("F1b name does not resolve: {s}\n", .{e.name});
            return error.UnregisteredBatchFunction;
        }
        var fixtures: usize = 0;
        for (f1b_cases) |c| {
            if (std.mem.eql(u8, c.func, e.name)) fixtures += 1;
        }
        if (fixtures == 0) {
            std.debug.print("F1b name has no fixture: {s}\n", .{e.name});
            return error.UnfixturedBatchFunction;
        }
    }
    try testing.expectEqual(@as(usize, 22), batch);

    // …and no fixture names something outside the batch.
    for (f1b_cases) |c| {
        var found = false;
        var it2 = registry.inventory();
        while (it2.next()) |e| {
            if (std.mem.eql(u8, e.name, c.func) and std.mem.eql(u8, e.milestone, "M4e")) found = true;
        }
        if (!found) {
            std.debug.print("fixture names a function outside F1b: {s}\n", .{c.func});
            return error.FixtureOutsideBatch;
        }
    }
}

test "M4e: the evidence label on every fixture is true of the committed manifests" {
    // Same three-valued checker M4d built (`silent` / `decided` /
    // `excluded`), shared rather than re-derived — which is why this
    // row gets the excluded-cell guard for free even though it holds no
    // volatile of its own.
    var oracle_rows: usize = 0;
    var excluded_rows: usize = 0;
    for (f1b_cases) |c| {
        switch (try manifestVerdict(c.formula)) {
            .decided => {
                if (c.evidence != .oracle) {
                    std.debug.print("`{s}` is decided by a manifest but ships spec-pinned\n", .{c.formula});
                    return error.UnderstatedEvidence;
                }
                oracle_rows += 1;
            },
            .excluded => {
                if (c.evidence != .spec_pinned) {
                    std.debug.print("`{s}` claims evidence from an EXCLUDED manifest cell\n", .{c.formula});
                    return error.ExcludedCellClaimedAsEvidence;
                }
                excluded_rows += 1;
            },
            .silent => {
                if (c.evidence != .spec_pinned) {
                    std.debug.print("`{s}` claims oracle evidence no manifest holds\n", .{c.formula});
                    return error.UnbackedOracleClaim;
                }
            },
        }
    }
    // **Zero** in twenty-two functions, stated as a number so it cannot
    // drift silently. When the parked Excel leg runs (§8.2) and the
    // suite grows F1b rows, this count moves and the row that moves it
    // is the row that re-labels.
    try testing.expectEqual(@as(usize, 0), oracle_rows);
    try testing.expectEqual(@as(usize, 0), excluded_rows);

    // Said the other way too, over the manifests rather than over the
    // fixtures: not one committed cell calls an F1b name at all. That
    // is what makes "the oracle decides nothing here" a fact about the
    // files instead of a claim about the table above.
    for ([_][]const u8{ oracle_excel, oracle_ieee, oracle_libreoffice }) |json| {
        const doc = try std.json.parseFromSlice(std.json.Value, testing.allocator, json, .{});
        defer doc.deinit();
        for (doc.value.object.get("cells").?.array.items) |cell| {
            const f = cell.object.get("formula") orelse continue;
            if (f == .null) continue;
            var it = registry.inventory();
            while (it.next()) |e| {
                if (!std.mem.eql(u8, e.milestone, "M4e")) continue;
                if (std.mem.indexOf(u8, f.string, e.name) == null) continue;
                std.debug.print("a manifest cell calls {s}: `{s}`\n", .{ e.name, f.string });
                return error.UnexpectedOracleCoverage;
            }
        }
    }
}

test "M4e: §5.3c's spine — the three counting functions over ONE range" {
    // The row's whole point, in five cells. `COUNT` ignores the error,
    // `COUNTA` counts it, `COUNTBLANK` answers a third question, and
    // `SUM` — the fourth name over the same range — propagates it. Four
    // functions, one range, four answers: the class is per function and
    // never per family, which is exactly what §5.3c says and what a
    // family-wide rule would get three-quarters wrong.
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try putF1bCells(&h);

    try testing.expectEqual(@as(f64, 1), (try h.scalar("COUNT(D1:D5)")).number);
    try testing.expectEqual(@as(f64, 4), (try h.scalar("COUNTA(D1:D5)")).number);
    try testing.expectEqual(@as(f64, 2), (try h.scalar("COUNTBLANK(D1:D5)")).number);
    try testing.expectEqual(value.KnownError.div0, (try h.scalar("SUM(D1:D5)")).err.known);

    // The three populations partition nothing and overlap on purpose:
    // `""` is counted by COUNTA **and** by COUNTBLANK, which is the
    // pair a single "is it empty" predicate could not have produced.
    try testing.expectEqual(@as(f64, 1), (try h.scalar("COUNTA(D5)")).number);
    try testing.expectEqual(@as(f64, 1), (try h.scalar("COUNTBLANK(D5)")).number);
    try testing.expectEqual(@as(f64, 0), (try h.scalar("COUNT(D5)")).number);
    // …while `ISBLANK` disagrees with both, over the same cell.
    try testing.expectEqual(false, (try h.scalar("ISBLANK(D5)")).boolean);
    try testing.expectEqual(true, (try h.scalar("ISBLANK(D4)")).boolean);

    // And the classes are what the registry says they are, name by
    // name — `per_function_provenance` for the three that inspect
    // provenance, `propagate` for the one that does not, and `observe`
    // for the one that looks at an error without becoming it.
    try testing.expectEqual(
        value.PropagationClass.per_function_provenance,
        registry.lookup("COUNT").?.propagation,
    );
    try testing.expectEqual(
        value.PropagationClass.per_function_provenance,
        registry.lookup("COUNTA").?.propagation,
    );
    try testing.expectEqual(
        value.PropagationClass.per_function_provenance,
        registry.lookup("COUNTBLANK").?.propagation,
    );
    try testing.expectEqual(value.PropagationClass.propagate, registry.lookup("SUM").?.propagation);
    try testing.expectEqual(value.PropagationClass.observe, registry.lookup("CHOOSE").?.propagation);
}

test "M4e: the 3D eligible list is the frozen six, and BOTH directions are fixtured" {
    // §5.6g freezes the list at exactly six, and every member of it is
    // an M4e name — which is what makes this the row that can finally
    // fixture it end to end. M4b3 could only run the three the registry
    // already held (its decision 14); `AVERAGE`, `MIN` and `MAX`
    // arrived here.
    try testing.expectEqual(@as(usize, 6), name_rules.three_d_eligible.len);
    for (name_rules.three_d_eligible) |name| {
        try testing.expect(registry.lookup(name) != null);
        var tagged = false;
        var it = registry.inventory();
        while (it.next()) |e| {
            if (std.mem.eql(u8, e.name, name)) tagged = std.mem.eql(u8, e.milestone, "M4e");
        }
        if (!tagged) {
            std.debug.print("3D-eligible name is not an M4e row: {s}\n", .{name});
            return error.EligibleOutsideBatch;
        }
    }

    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    const s2 = try h.fake.addSheet("Sheet2");
    try h.fake.putA1(h.sheet, .stored, "A1", num(4));
    try h.fake.putA1(s2, .stored, "A1", num(6));

    // Direction one: each of the six aggregates a span. Six functions,
    // six different right answers over the same two cells — which is
    // also the proof they aggregate the span rather than one member.
    const Span = struct { name: []const u8, want: f64 };
    const spans = [_]Span{
        .{ .name = "SUM", .want = 10 },
        .{ .name = "COUNT", .want = 2 },
        .{ .name = "COUNTA", .want = 2 },
        .{ .name = "AVERAGE", .want = 5 },
        .{ .name = "MIN", .want = 4 },
        .{ .name = "MAX", .want = 6 },
    };
    for (spans) |s| {
        var buf: [64]u8 = undefined;
        const src = try std.fmt.bufPrint(&buf, "{s}(Sheet1:Sheet2!A1)", .{s.name});
        const got = h.scalar(src) catch |e| {
            std.debug.print("eligible 3D consumer `{s}` refused: {t}\n", .{ src, e });
            return e;
        };
        try testing.expectEqual(s.want, got.number);
    }
    // The list above is not a second copy of the frozen one: every
    // frozen member must appear here, so a seventh eligible name cannot
    // ship without a span fixture.
    try testing.expectEqual(name_rules.three_d_eligible.len, spans.len);
    for (name_rules.three_d_eligible) |name| {
        var covered = false;
        for (spans) |s| {
            if (std.mem.eql(u8, s.name, name)) covered = true;
        }
        if (!covered) {
            std.debug.print("eligible name has no 3D span fixture: {s}\n", .{name});
            return error.UnfixturedEligibleFunction;
        }
    }

    // Direction two: every OTHER name in the batch refuses a span, and
    // refuses it **typed** — `UnsupportedConstruct` carrying §5.6g's
    // own reason, not a generic failure. The list is derived from the
    // inventory and the arity, so a name added to the batch later
    // cannot ship without landing in one direction or the other.
    var it = registry.inventory();
    var refused: usize = 0;
    while (it.next()) |e| {
        if (!std.mem.eql(u8, e.milestone, "M4e")) continue;
        if (name_rules.threeDEligible(e.name)) continue;
        const f = registry.lookup(e.name).?;

        // The span goes in the first slot and the rest is padded to the
        // minimum arity. §5.6g refuses BEFORE evaluation, so what the
        // padding would have evaluated to never matters.
        var pad: [64]u8 = undefined;
        var used: usize = 0;
        var k: usize = 1;
        while (k < @max(f.arity.min, 1)) : (k += 1) {
            @memcpy(pad[used..][0..2], ",1");
            used += 2;
        }
        var buf: [128]u8 = undefined;
        const src = try std.fmt.bufPrint(&buf, "{s}(Sheet1:Sheet2!A1{s})", .{ e.name, pad[0..used] });

        _ = h.eval(src) catch |err| {
            if (err != error.UnsupportedConstruct) {
                std.debug.print("`{s}` refused a span as {t}, not UnsupportedConstruct\n", .{ src, err });
                return err;
            }
            try testing.expectEqual(
                name_rules.Refusal.Reason.three_d_ineligible_function,
                h.ev.last_three_d.?.reason,
            );
            refused += 1;
            continue;
        };
        std.debug.print("ineligible name accepted a 3D span: {s}\n", .{src});
        return error.IneligibleFunctionAcceptedSpan;
    }
    // Twenty-two names, six eligible, sixteen refusing — counted rather
    // than asserted, so the two directions cannot both be describing
    // the same subset.
    try testing.expectEqual(@as(usize, 16), refused);
}

test "M4e: MIN and MAX are outside the comparator, the five lookups are inside it" {
    // §5.4b's one named exception (plan revision 15, change 8), proved
    // behaviourally rather than by reading the flag back.
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try putF1bCells(&h);
    // Ascending under the fold — `Beta` sorts between them, which the
    // raw code points would also give; the case is what makes the
    // ordering below a statement about `collation_v1`.
    try h.put("Q1", .{ .text = "alpha" });
    try h.put("Q2", .{ .text = "Beta" });
    try h.put("Q3", .{ .text = "gamma" });

    // The exception, at its sharpest. Under the comparator's cross-type
    // ranking (number < text < logical) any text outranks every number,
    // so a MAX that used it would answer "gamma" here. It answers 4,
    // because it never compares text at all.
    try testing.expectEqual(@as(f64, 4), (try h.scalar("MAX(C1:C4,Q1:Q3)")).number);
    try testing.expectEqual(@as(f64, 1), (try h.scalar("MIN(C1:C4,Q1:Q3)")).number);
    // And with nothing numeric anywhere the answer is 0 — not the
    // first string, and not an error.
    try testing.expectEqual(@as(f64, 0), (try h.scalar("MAX(Q1:Q3)")).number);
    try testing.expectEqual(@as(f64, 0), (try h.scalar("MIN(Q1:Q3)")).number);
    // A DIRECT text argument still refuses, which is the other half of
    // §5.4b's sentence: text is not compared, and it is not silently
    // skipped when it was written out by hand either.
    try testing.expectEqual(value.KnownError.value, (try h.scalar("MAX(\"gamma\")")).err.known);

    // The five lookups take EQUALITY from the comparator: `ß` folds to
    // `ss`, so `STRASSE` and `Straße` are one key. Every one of the
    // five, because the flag is per function and the fold has to be
    // reached the same way by each.
    try testing.expectEqual(@as(f64, 1), (try h.scalar("VLOOKUP(\"STRASSE\",J1:K2,2,FALSE)")).number);
    try testing.expectEqual(@as(f64, 1), (try h.scalar("MATCH(\"straSSe\",J1:J2,0)")).number);
    try testing.expectEqual(@as(f64, 1), (try h.scalar("XMATCH(\"Strasse\",J1:J2,0)")).number);
    try testing.expectEqual(@as(f64, 1), (try h.scalar("XLOOKUP(\"STRASSE\",J1:J2,K1:K2)")).number);
    try h.put("R1", .{ .text = "Straße" });
    try h.put("S1", .{ .text = "beta" });
    try h.put("R2", num(1));
    try h.put("S2", num(2));
    try testing.expectEqual(@as(f64, 1), (try h.scalar("HLOOKUP(\"STRASSE\",R1:S2,2,FALSE)")).number);

    // …and ORDERING from it too, which is the half an equality-only
    // fixture would miss. `bz` sorts after `Beta` and before `gamma`,
    // case-insensitively, so the largest key ≤ it is the second.
    try testing.expectEqual(@as(f64, 2), (try h.scalar("MATCH(\"bz\",Q1:Q3,1)")).number);
    try testing.expectEqual(@as(f64, 2), (try h.scalar("MATCH(\"BETA\",Q1:Q3,1)")).number);
    try testing.expectEqual(@as(f64, 2), (try h.scalar("XMATCH(\"bz\",Q1:Q3,-1)")).number);
    // Nothing sorts at or below `a`, since `a` is a proper prefix of
    // `alpha` and therefore strictly less than it.
    try testing.expectEqual(value.KnownError.na, (try h.scalar("MATCH(\"a\",Q1:Q3,1)")).err.known);
    // An ordered match never crosses a type boundary: a numeric key
    // finds nothing in a column of text, rather than the whole column.
    try testing.expectEqual(value.KnownError.na, (try h.scalar("MATCH(5,Q1:Q3,1)")).err.known);
}

test "M4e: CHOOSE is lazy for a scalar selector and masks per element for an array one" {
    // §5.3a assigns CHOOSE its fixtures at this row. The instrument is
    // the draw COUNT, because a result cannot tell a draw that happened
    // from one that did not — under a constant source both look alike.
    const Case = struct { formula: []const u8, draws: u64, expect: Expect };
    const cases = [_]Case{
        // Every arm position of a three-arm call: one draw each, which
        // is two dead arms at zero every time.
        .{ .formula = "CHOOSE(1,RAND(),RAND(),RAND())", .draws = 1, .expect = .{ .number = 0.5 } },
        .{ .formula = "CHOOSE(2,RAND(),RAND(),RAND())", .draws = 1, .expect = .{ .number = 0.5 } },
        .{ .formula = "CHOOSE(3,RAND(),RAND(),RAND())", .draws = 1, .expect = .{ .number = 0.5 } },
        // A dead arm at exactly zero, stated on its own: the volatile
        // is in the arm nobody took, so the whole formula draws none.
        .{ .formula = "CHOOSE(1,7,RAND())", .draws = 0, .expect = .{ .number = 7 } },
        .{ .formula = "CHOOSE(2,RAND(),7)", .draws = 0, .expect = .{ .number = 7 } },
        // An out-of-range selector takes NO arm, so nothing draws — and
        // the answer is `#VALUE!` rather than the nearest arm.
        .{ .formula = "CHOOSE(0,RAND(),RAND())", .draws = 0, .expect = .{ .err = .value } },
        .{ .formula = "CHOOSE(3,RAND(),RAND())", .draws = 0, .expect = .{ .err = .value } },
        // …and neither does a selector that is itself an error.
        .{ .formula = "CHOOSE(1/0,RAND(),RAND())", .draws = 0, .expect = .{ .err = .div0 } },
        // Two call sites in one formula are two call sites (§5.6d).
        .{ .formula = "CHOOSE(1,RAND())+CHOOSE(1,RAND())", .draws = 2, .expect = .{ .number = 1 } },
        // An ARRAY selector switches the whole form to per-element
        // masking: every arm evaluates, so both volatiles draw, and the
        // result is the mask's shape rather than one arm's.
        .{
            .formula = "CHOOSE({1,2},RAND(),RAND())",
            .draws = 2,
            .expect = .{ .array = .{
                .rows = 1,
                .cols = 2,
                .cells = &.{ .{ .number = 0.5 }, .{ .number = 0.5 } },
            } },
        },
        // Per-element errors stay per element: the second selector is
        // out of range and only its cell is `#VALUE!`.
        .{
            .formula = "CHOOSE({1,9},10,20)",
            .draws = 0,
            .expect = .{ .array = .{
                .rows = 1,
                .cols = 2,
                .cells = &.{ .{ .number = 10 }, .{ .err = .value } },
            } },
        },
    };

    for (cases) |c| {
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        try putF1bCells(&h);

        h.draws.count = 0;
        const v = h.eval(c.formula) catch |e| {
            std.debug.print("CHOOSE case `{s}` refused: {t}\n", .{ c.formula, e });
            return e;
        };
        expectValue(c.expect, v) catch |e| {
            std.debug.print("CHOOSE case `{s}`: wrong value\n", .{c.formula});
            return e;
        };
        if (h.draws.count != c.draws) {
            std.debug.print(
                "CHOOSE case `{s}`: expected {d} draws, counted {d}\n",
                .{ c.formula, c.draws, h.draws.count },
            );
            return error.WrongDrawCount;
        }
    }
}

const F1bFidelityCase = struct {
    formula: []const u8,
    excel: f64,
    ieee: f64,
    why: []const u8,
};

/// Where `excel_fp_rules_v1` and `ieee_fp_rules_v1` part in F1b. Two
/// places, and neither is the one M4d found: this batch never rounds to
/// a decimal place, so `decimalView` is not involved at all.
const f1b_fidelity_cases = [_]F1bFidelityCase{
    // N2's zero-snap is **additive-scope**, and an aggregate is a chain
    // of additions — so the rule M4d could only reach through a
    // rounding argument reaches this batch through its own fold.
    .{ .formula = "SUM(0.1,0.2,-0.3)", .excel = 0, .ieee = 5.551115123125783e-17, .why = "the running total lands within 2^-48 of zero relative to its operands" },
    .{ .formula = "AVERAGE(0.1,0.2,-0.3)", .excel = 0, .ieee = 1.850371707708594e-17, .why = "the same total, divided by three — the snap happens before the division" },
    .{ .formula = "SUMPRODUCT({0.1,0.2,-0.3},{1,1,1})", .excel = 0, .ieee = 5.551115123125783e-17, .why = "…and again through SUMPRODUCT's accumulator, which is the same addition" },
    // N3: a negative zero survives publication under `.ieee` and is
    // normalized under `.excel`. MIN and MAX are the only names in the
    // batch that can produce one, because they RETURN an input rather
    // than computing a result — an accumulator would have added it to
    // `+0` and lost the sign on the way.
    .{ .formula = "MIN(E1)", .excel = 0, .ieee = -0.0, .why = "MIN returns the stored -0 itself" },
    .{ .formula = "MAX(E1)", .excel = 0, .ieee = -0.0, .why = "and so does MAX, over the same one-cell range" },
};

test "M4e: both fidelity modes are fixtured wherever the two rule tables disagree" {
    for (f1b_fidelity_cases) |c| {
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        try putF1bCells(&h);

        inline for ([_]value.Fidelity{ .excel, .ieee }) |mode| {
            var opts = h.options();
            opts.fidelity = mode;
            const v = h.evalOpts(c.formula, opts) catch |e| {
                std.debug.print("fidelity case `{s}` refused: {t}\n", .{ c.formula, e });
                return e;
            };
            try testing.expect(v == .scalar);
            const got = value.publish(v.scalar, mode);
            const want = if (mode == .excel) c.excel else c.ieee;
            try testing.expect(got == .number);
            if (@as(u64, @bitCast(got.number)) != @as(u64, @bitCast(want))) {
                std.debug.print(
                    "`{s}` under {s}: expected {d} (0x{X:0>16}), got {d} (0x{X:0>16}) — {s}\n",
                    .{
                        c.formula,
                        @tagName(mode),
                        want,
                        @as(u64, @bitCast(want)),
                        got.number,
                        @as(u64, @bitCast(got.number)),
                        c.why,
                    },
                );
                return error.FidelityMismatch;
            }
        }

        // A table of "divergences" that agreed everywhere would pass a
        // per-mode check and prove nothing.
        if (@as(u64, @bitCast(c.excel)) == @as(u64, @bitCast(c.ieee))) {
            std.debug.print("`{s}` is listed as a divergence but does not diverge\n", .{c.formula});
            return error.NonDivergentFidelityCase;
        }
    }
}

test "M4e: the batch agrees across modes everywhere it is not listed as diverging" {
    // The converse half. Most of F1b is mode-independent — a lookup
    // returns a stored value unchanged, and a position is an integer —
    // and saying so is what makes the short list above an enumeration
    // rather than a sampling.
    for (f1b_cases) |c| {
        var listed = false;
        for (f1b_fidelity_cases) |d| {
            if (std.mem.eql(u8, d.formula, c.formula)) listed = true;
        }
        if (listed) continue;

        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        try putF1bCells(&h);

        var excel_opts = h.options();
        excel_opts.fidelity = .excel;
        const a = try h.evalOpts(c.formula, excel_opts);
        var ieee_opts = h.options();
        ieee_opts.fidelity = .ieee;
        const b = try h.evalOpts(c.formula, ieee_opts);
        // Arrays reach this table too (`CHOOSE`'s masked form does not,
        // but `INDEX`'s slices could), so agreement is checked over
        // whatever shape came back rather than over a scalar.
        if (a == .array or b == .array) {
            try testing.expect(a == .array and b == .array);
            try testing.expectEqual(a.array.rows, b.array.rows);
            try testing.expectEqual(a.array.cols, b.array.cols);
            for (a.array.cells, b.array.cells) |x, y| {
                if (!value.PublishedScalar.eql(value.publish(x, .excel), value.publish(y, .ieee))) {
                    std.debug.print("`{s}` diverges between modes but is not listed\n", .{c.formula});
                    return error.UnlistedFidelityDivergence;
                }
            }
            continue;
        }
        if (!value.PublishedScalar.eql(value.publish(a.scalar, .excel), value.publish(b.scalar, .ieee))) {
            std.debug.print("`{s}` diverges between modes but is not listed\n", .{c.formula});
            return error.UnlistedFidelityDivergence;
        }
    }
}

test "M4e: error order in every multi-argument name of the batch (§5.3c)" {
    // Every case runs in both argument orders, because a fixture with
    // one error in it proves propagation and says nothing about order.
    // Half of this batch does not propagate at all, and for those the
    // pair proves the opposite: the class overrides the rule in BOTH
    // directions, which a single-order fixture could not distinguish
    // from a lucky argument list.
    const Case = struct { formula: []const u8, expect: Expect, note: []const u8 = "" };
    const cases = [_]Case{
        // The propagating aggregates: first error in §5.6a's order.
        .{ .formula = "SUM(A5,A6)", .expect = .{ .err = .div0 } },
        .{ .formula = "SUM(A6,A5)", .expect = .{ .err = .na } },
        .{ .formula = "AVERAGE(A5,A6)", .expect = .{ .err = .div0 } },
        .{ .formula = "AVERAGE(A6,A5)", .expect = .{ .err = .na } },
        .{ .formula = "MIN(A5,A6)", .expect = .{ .err = .div0 } },
        .{ .formula = "MIN(A6,A5)", .expect = .{ .err = .na } },
        .{ .formula = "MAX(A5,A6)", .expect = .{ .err = .div0 } },
        .{ .formula = "MAX(A6,A5)", .expect = .{ .err = .na } },
        .{ .formula = "SUMPRODUCT(A5,A6)", .expect = .{ .err = .div0 }, .note = "the error pass runs argument by argument BEFORE the products" },
        .{ .formula = "SUMPRODUCT(A6,A5)", .expect = .{ .err = .na } },
        // The counting family: neither order propagates, which is the
        // whole point of the class.
        .{ .formula = "COUNT(A5,A6)", .expect = .{ .number = 0 } },
        .{ .formula = "COUNT(A6,A5)", .expect = .{ .number = 0 } },
        .{ .formula = "COUNTA(A5,A6)", .expect = .{ .number = 2 }, .note = "COUNTA counts the very errors COUNT ignores" },
        .{ .formula = "COUNTA(A6,A5)", .expect = .{ .number = 2 } },
        .{ .formula = "COUNTIF(A5,A6)", .expect = .{ .number = 0 } },
        .{ .formula = "COUNTIF(A6,A5)", .expect = .{ .number = 0 } },
        .{ .formula = "COUNTIF(A5,A5)", .expect = .{ .number = 1 }, .note = "…and a criterion CAN match one, which is why neither of the two above is a refusal" },
        .{ .formula = "SUMIF(A5,A6)", .expect = .{ .number = 0 } },
        .{ .formula = "SUMIF(A6,A5)", .expect = .{ .number = 0 } },
        .{ .formula = "AVERAGEIF(A5,A6)", .expect = .{ .err = .div0 }, .note = "#DIV/0! from the division over no match, NOT from A5" },
        .{ .formula = "AVERAGEIF(A6,A5)", .expect = .{ .err = .div0 } },
        // The lookups, whose key slot is a reference the dispatcher's
        // scan cannot see — so the order below is the implementation's.
        .{ .formula = "VLOOKUP(A5,F1:H3,A6)", .expect = .{ .err = .div0 } },
        .{ .formula = "VLOOKUP(A6,F1:H3,A5)", .expect = .{ .err = .na } },
        .{ .formula = "HLOOKUP(A5,F5:H6,A6)", .expect = .{ .err = .div0 } },
        .{ .formula = "HLOOKUP(A6,F5:H6,A5)", .expect = .{ .err = .na } },
        .{ .formula = "MATCH(A5,F1:F3,A6)", .expect = .{ .err = .div0 } },
        .{ .formula = "MATCH(A6,F1:F3,A5)", .expect = .{ .err = .na } },
        .{ .formula = "XMATCH(A5,F1:F3,A6)", .expect = .{ .err = .div0 } },
        .{ .formula = "XMATCH(A6,F1:F3,A5)", .expect = .{ .err = .na } },
        // Slots 0 and 4 — the key and the match mode — with slots 1..3
        // masked, so this pair really is about declaration order.
        .{ .formula = "XLOOKUP(A5,F1:F3,G1:G3,\"x\",A6)", .expect = .{ .err = .div0 } },
        .{ .formula = "XLOOKUP(A6,F1:F3,G1:G3,\"x\",A5)", .expect = .{ .err = .na } },
        // …and `if_not_found` is a value this call may RETURN rather
        // than one it propagates: an error there does not spoil a hit,
        // and on a miss it is the answer.
        .{ .formula = "XLOOKUP(20,F1:F3,G1:G3,A5)", .expect = .{ .text = "twenty" } },
        .{ .formula = "XLOOKUP(99,F1:F3,G1:G3,A5)", .expect = .{ .err = .div0 }, .note = "returned, not propagated — the same cell, two outcomes" },
        .{ .formula = "XLOOKUP(20,F1:F3,G1:G3,C1:C4)", .expect = .{ .text = "twenty" }, .note = "a RANGE fallback is fine too, which propagating it would have made #VALUE!" },
        // …and an error found INSIDE a lookup table does not propagate
        // at all. It is a value the lookup may return, which is the
        // other half of `per_function_provenance`.
        .{ .formula = "MATCH(20,A5:A6,0)", .expect = .{ .err = .na }, .note = "#N/A from a failed match, not #DIV/0! from the table" },
        .{ .formula = "VLOOKUP(10,F1:H3,3)", .expect = .{ .number = 100 } },
        // INDEX propagates from its index slots in declaration order,
        // and returns an error found in the array as the value it is.
        .{ .formula = "INDEX(F1:H3,A5,A6)", .expect = .{ .err = .div0 } },
        .{ .formula = "INDEX(F1:H3,A6,A5)", .expect = .{ .err = .na } },
        .{ .formula = "INDEX(A5:A6,1)", .expect = .{ .err = .div0 }, .note = "the element IS the error — propagation never entered into it" },
        // CHOOSE observes, and laziness decides WHICH error you get.
        .{ .formula = "CHOOSE(1,A5,A6)", .expect = .{ .err = .div0 } },
        .{ .formula = "CHOOSE(2,A5,A6)", .expect = .{ .err = .na } },
        .{ .formula = "CHOOSE(A5,1,2)", .expect = .{ .err = .div0 }, .note = "an erroring selector takes no arm at all" },
    };

    for (cases) |c| {
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        try putF1bCells(&h);

        const v = h.eval(c.formula) catch |e| {
            std.debug.print("error-order case `{s}` refused: {t}\n", .{ c.formula, e });
            return e;
        };
        expectValue(c.expect, v) catch |e| {
            std.debug.print("error-order case `{s}`: wrong value\n", .{c.formula});
            return e;
        };
    }

    // Every multi-argument name in the batch appears above, with the
    // list derived from the registry's own arity rather than typed out
    // — so a function that gains an argument later cannot slip past
    // unordered. Each must appear at least TWICE: one order is not an
    // order.
    var it = registry.inventory();
    var covered_names: usize = 0;
    while (it.next()) |e| {
        if (!std.mem.eql(u8, e.milestone, "M4e")) continue;
        const f = registry.lookup(e.name).?;
        const multi = f.arity.max == null or f.arity.max.? > 1;
        if (!multi) continue;
        var seen: usize = 0;
        for (cases) |c| {
            if (std.mem.startsWith(u8, c.formula, e.name) and c.formula[e.name.len] == '(') seen += 1;
        }
        if (seen < 2) {
            std.debug.print("multi-argument name with {d} error-order fixtures: {s}\n", .{ seen, e.name });
            return error.MissingErrorOrderFixture;
        }
        covered_names += 1;
    }
    // Seventeen of the twenty-two take more than one argument; the five
    // that do not are COUNTBLANK, ROW, ROWS, COLUMN and COLUMNS.
    try testing.expectEqual(@as(usize, 17), covered_names);
}

test "M4e: ROW and COLUMN answer from the site, or refuse for want of one" {
    // The registry's only site-dependent rows. `AnchorRequired` already
    // exists for `@` (§5.3b); these are the first *functions* to raise
    // it, and the alternative — guessing 1 — would be a number nobody
    // could tell from a real answer.
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try putF1bCells(&h);

    try testing.expectError(error.AnchorRequired, h.eval("ROW()"));
    try testing.expectError(error.AnchorRequired, h.eval("COLUMN()"));
    try testing.expectEqual(parser.PlaneTwo.FormulaAnchorRequired, planeTwo(error.AnchorRequired));

    var opts = h.options();
    opts.site = .{
        .row = coords.Row.fromOneBased(7) catch unreachable,
        .col = coords.Col.fromZeroBased(2) catch unreachable,
    };
    try testing.expectEqual(@as(f64, 7), (try h.evalOpts("ROW()", opts)).scalar.number);
    try testing.expectEqual(@as(f64, 3), (try h.evalOpts("COLUMN()", opts)).scalar.number);
    // An explicit reference ignores the site entirely, which is what
    // makes the zero-argument form the special case rather than the
    // rule.
    try testing.expectEqual(@as(f64, 2), (try h.evalOpts("ROW(F2)", opts)).scalar.number);
    try testing.expectEqual(@as(f64, 6), (try h.evalOpts("COLUMN(F2)", opts)).scalar.number);
    // ROWS and COLUMNS never consult it: they answer about a shape.
    try testing.expectEqual(@as(f64, 3), (try h.scalar("ROWS(F1:H3)")).number);
    try testing.expectEqual(@as(f64, 3), (try h.scalar("COLUMNS(F1:H3)")).number);
}

test "M4e: the two names that return an ARRAY return the right shape" {
    // `INDEX` with a zero index and `XLOOKUP` with a 2-D return range
    // are the batch's only array producers. Both answers are Excel's,
    // and both are shapes this evaluator already carries — the spilling
    // question they raise belongs to M7a, not to what they compute.
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try putF1bCells(&h);

    const cases = [_]struct { formula: []const u8, expect: Expect }{
        // A whole column of the table.
        .{ .formula = "INDEX(F1:H3,0,2)", .expect = .{ .array = .{
            .rows = 3,
            .cols = 1,
            .cells = &.{ .{ .text = "ten" }, .{ .text = "twenty" }, .{ .text = "thirty" } },
        } } },
        // A whole row — and the form a single index takes on a 2-D
        // array, which is the same thing written more briefly.
        .{ .formula = "INDEX(F1:H3,2,0)", .expect = .{ .array = .{
            .rows = 1,
            .cols = 3,
            .cells = &.{ .{ .number = 20 }, .{ .text = "twenty" }, .{ .number = 200 } },
        } } },
        .{ .formula = "INDEX(F1:H3,2)", .expect = .{ .array = .{
            .rows = 1,
            .cols = 3,
            .cells = &.{ .{ .number = 20 }, .{ .text = "twenty" }, .{ .number = 200 } },
        } } },
        // The whole array, when both indices are zero.
        .{ .formula = "COLUMNS(INDEX(F1:H3,0,0))", .expect = .{ .number = 3 } },
        // XLOOKUP's 2-D return: the row at the match, across the range.
        .{ .formula = "XLOOKUP(20,F1:F3,G1:H3)", .expect = .{ .array = .{
            .rows = 1,
            .cols = 2,
            .cells = &.{ .{ .text = "twenty" }, .{ .number = 200 } },
        } } },
    };
    for (cases) |c| {
        const v = h.eval(c.formula) catch |e| {
            std.debug.print("array case `{s}` refused: {t}\n", .{ c.formula, e });
            return e;
        };
        expectValue(c.expect, v) catch |e| {
            std.debug.print("array case `{s}`: wrong shape or value\n", .{c.formula});
            return e;
        };
    }
}

test "M4e: a rectangle beyond §9's cap is a limit, not a hang" {
    // Lookups materialize their table, which is what buys them random
    // access along an axis (`registry.Grid`). The bound on that is
    // §9's `max_matrix_cells`, so an absurd range is a **limit** —
    // a defined plane-2 outcome — rather than a run that never ends.
    // Making a whole-column lookup fast is M7b2's row; making an
    // impossible one refuse is this one's.
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try putF1bCells(&h);

    try testing.expectError(error.LimitExceeded, h.eval("MATCH(1,A1:XFD100000,0)"));
    try testing.expectError(error.LimitExceeded, h.eval("VLOOKUP(1,A1:XFD100000,2)"));
    try testing.expectError(error.LimitExceeded, h.eval("SUMPRODUCT(A1:XFD100000)"));
    try testing.expectEqual(parser.PlaneTwo.FormulaLimitExceeded, planeTwo(error.LimitExceeded));
    // The aggregates walk sparsely instead and are unbothered by the
    // same range — which is the difference the two access patterns
    // make. `SUM` reaches D3's error and answers WITH it: an error
    // value is a result, and therefore the proof the walk finished.
    try testing.expectEqual(value.KnownError.div0, (try h.scalar("SUM(A1:XFD100000)")).err.known);
    try testing.expectEqual(@as(f64, 10), (try h.scalar("SUM(C1:C1000000)")).number);
    try testing.expectEqual(@as(f64, 4), (try h.scalar("COUNT(C1:C1000000)")).number);
}

// ─── M4e fuzz: no argument shape panics, leaks, or evaluates twice ──

const f1b_names = [_][]const u8{
    "AVERAGE", "AVERAGEIF", "CHOOSE",     "COLUMN",     "COLUMNS",
    "COUNT",   "COUNTA",    "COUNTBLANK", "COUNTIF",    "HLOOKUP",
    "INDEX",   "MATCH",     "MAX",        "MIN",        "ROW",
    "ROWS",    "SUM",       "SUMIF",      "SUMPRODUCT", "VLOOKUP",
    "XLOOKUP", "XMATCH",
};

/// Argument shapes. F1a-1's alphabet was about provenance and F1a-2's
/// about numeric extremes; this one is about **shape**, because that is
/// the failure mode a batch of aggregates and lookups has and the other
/// two did not: a table with no rows, a vector that is really a
/// rectangle, an index past the end, a multi-area set where one
/// rectangle was expected, and a criterion that is not a criterion.
const f1b_arg_shapes = [_][]const u8{
    "A1",            "A2",     "A5",     "A7",
    "A8",            "C1:C4",  "D1:D5",  "F1:H3",
    "(A1:A2,C1:C4)", "{1,2}",  "{1;2}",  "{1,2;3,4}",
    "0",             "1",      "-1",     "2.5",
    "\"\"",          "\">2\"", "\"t*\"", "\"abc\"",
    "",              "TRUE()", "1/0",    "1E+308",
};

fn fuzzF1bEnv(fake: *env.Fake) !env.SheetIndex {
    const sheet = try fake.addSheet("S");
    try fake.putA1(sheet, .stored, "A1", num(10));
    try fake.putA1(sheet, .stored, "A2", .{ .text = "abc" });
    try fake.putA1(sheet, .stored, "A5", value.ScalarValue.errorOf(.div0));
    try fake.putA1(sheet, .stored, "A6", value.ScalarValue.errorOf(.na));
    try fake.putA1(sheet, .stored, "A8", .{ .text = "7" });
    try fake.putA1(sheet, .stored, "C1", num(1));
    try fake.putA1(sheet, .stored, "C2", num(2));
    try fake.putA1(sheet, .stored, "C3", num(3));
    try fake.putA1(sheet, .stored, "C4", num(4));
    try fake.putA1(sheet, .stored, "D1", num(1));
    try fake.putA1(sheet, .stored, "D2", .{ .text = "x" });
    try fake.putA1(sheet, .stored, "D3", value.ScalarValue.errorOf(.div0));
    try fake.putA1(sheet, .stored, "D5", .{ .text = "" });
    try fake.putA1(sheet, .stored, "F1", num(10));
    try fake.putA1(sheet, .stored, "F2", num(20));
    try fake.putA1(sheet, .stored, "F3", num(30));
    try fake.putA1(sheet, .stored, "G1", .{ .text = "ten" });
    try fake.putA1(sheet, .stored, "G2", .{ .text = "twenty" });
    try fake.putA1(sheet, .stored, "G3", .{ .text = "thirty" });
    try fake.putA1(sheet, .stored, "H1", num(100));
    try fake.putA1(sheet, .stored, "H2", num(200));
    try fake.putA1(sheet, .stored, "H3", num(300));
    return sheet;
}

/// Build `NAME(a, b, b, …)`, padded to the function's own minimum
/// arity. Without the padding a three-argument name like `XLOOKUP`
/// would only ever be swept at arities it rejects outright, and the
/// sweep would prove nothing about the one code path it exists to
/// reach.
fn buildF1bCall(buf: []u8, name: []const u8, args: []const []const u8) ?[]const u8 {
    if (args.len == 0) return null;
    const f = registry.lookup(name).?;
    const min = @max(@as(usize, f.arity.min), args.len);
    var w: usize = 0;
    if (name.len + 1 > buf.len) return null;
    @memcpy(buf[w..][0..name.len], name);
    w += name.len;
    buf[w] = '(';
    w += 1;
    var k: usize = 0;
    while (k < min) : (k += 1) {
        const arg = if (k < args.len) args[k] else args[args.len - 1];
        if (w + arg.len + 2 > buf.len) return null;
        if (k > 0) {
            buf[w] = ',';
            w += 1;
        }
        @memcpy(buf[w..][0..arg.len], arg);
        w += arg.len;
    }
    if (w + 1 > buf.len) return null;
    buf[w] = ')';
    w += 1;
    return buf[0..w];
}

fn fuzzF1bTarget(_: void, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();

    var picked: [6][]const u8 = undefined;
    var n: usize = 0;
    while (n < picked.len and !smith.eos()) : (n += 1) {
        picked[n] = f1b_arg_shapes[smith.index(f1b_arg_shapes.len)];
    }
    if (n == 0) return;
    const name = f1b_names[smith.index(f1b_names.len)];
    var buf: [512]u8 = undefined;
    const src = buildF1bCall(&buf, name, picked[0..n]) orelse return;

    var parsed = parser.parse(std.testing.allocator, src, .{}) catch return;
    defer parsed.deinit(std.testing.allocator);
    if (parsed == .refused) return;

    var arena_state = std.heap.ArenaAllocator.init(std.testing.allocator);
    defer arena_state.deinit();
    var fake = env.Fake.init(std.testing.allocator);
    defer fake.deinit();
    const sheet = try fuzzF1bEnv(&fake);

    for ([_]value.Fidelity{ .excel, .ieee }) |mode| {
        var draw_value: f64 = 0.5;
        var draws = DrawSource.constant(&draw_value);
        const opts: Options = .{
            .current_sheet = sheet,
            .collation = .{ .fold = shippedFold },
            .draws = &draws,
            .fidelity = mode,
            .site = .{
                .row = coords.Row.fromOneBased(2) catch unreachable,
                .col = coords.Col.fromZeroBased(1) catch unreachable,
            },
        };

        var first = Evaluator.init(arena_state.allocator(), fake.evalEnv(), opts);
        defer first.deinit();
        var second = Evaluator.init(arena_state.allocator(), fake.evalEnv(), opts);
        defer second.deinit();

        if (first.evaluate(parsed.ok)) |a| {
            try assertRepresentable(a);
            const b = second.evaluate(parsed.ok) catch return error.NondeterministicEvaluation;
            try assertRepresentable(b);
            if (!valuesAgree(a, b)) {
                std.debug.print("`{s}` evaluated two ways\n", .{src});
                return error.NondeterministicEvaluation;
            }
        } else |e| {
            if (second.evaluate(parsed.ok)) |_| {
                return error.NondeterministicEvaluation;
            } else |e2| {
                if (e != e2) return error.NondeterministicEvaluation;
            }
        }
    }
}

test "fuzz: no F1b argument shape panics, leaks, or evaluates two ways" {
    // The generator's alphabet is a second copy of the batch, so it is
    // checked against the file rather than maintained by hand.
    var it = registry.inventory();
    var batch: usize = 0;
    while (it.next()) |e| {
        if (!std.mem.eql(u8, e.milestone, "M4e")) continue;
        batch += 1;
        var present = false;
        for (f1b_names) |n| {
            if (std.mem.eql(u8, n, e.name)) present = true;
        }
        if (!present) {
            std.debug.print("fuzz alphabet is missing {s}\n", .{e.name});
            return error.FuzzAlphabetIncomplete;
        }
    }
    try testing.expectEqual(batch, f1b_names.len);

    try std.testing.fuzz({}, fuzzF1bTarget, .{
        .corpus = &[_][]const u8{
            "SUM(C1:C4)",
            "AVERAGE(A7)",
            "AVERAGE((A1:A2,C1:C4))",
            "MIN(D1:D5)",
            "MAX({1,2;3,4})",
            "SUMPRODUCT(C1:C4,C1:C4)",
            "SUMPRODUCT(C1:C4,F1:H3)",
            "COUNT(D1:D5)",
            "COUNTA(D1:D5)",
            "COUNTBLANK(D1:D5)",
            "COUNTIF(D1:D5,\"\")",
            "SUMIF(C1:C4,\">2\",F1:H3)",
            "AVERAGEIF(C1:C4,\">99\")",
            "VLOOKUP(\"t*\",F1:H3,3,FALSE)",
            "VLOOKUP(1E+308,F1:H3,1E+308)",
            "HLOOKUP(0,F1:H3,-1)",
            "MATCH(\"\",D1:D5,0)",
            "MATCH(2.5,F1:H3,-1)",
            "XMATCH(1,C1:C4,2,-2)",
            "XMATCH(1,C1:C4,1E+308,1E+308)",
            "XLOOKUP(1,C1:C4,F1:H3,\"\",0,1)",
            "XLOOKUP(1,F1:H3,C1:C4,1/0)",
            "INDEX(F1:H3,0,0)",
            "INDEX(F1:H3,1E+308,-1)",
            "CHOOSE({1,2},C1:C4,D1:D5)",
            "ROW((A1:A2,C1:C4))",
            "COLUMNS({1,2;3,4})",
        },
    });
}

test "M4e: every name against every argument shape, exhaustively and in both modes" {
    // The `fuzz` step is Linux-only (`build.zig` — coverage-guided
    // fuzzing is broken upstream on macOS and Windows), so on the other
    // two platforms the target above runs its corpus and nothing else.
    // The property it exists to prove is small enough to prove by
    // ENUMERATION instead: twenty-two names, every shape in the
    // alphabet at one and two arguments — each padded to the name's own
    // minimum arity so the three-argument lookups actually run — both
    // rule tables, every input evaluated twice.
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    var fake = env.Fake.init(testing.allocator);
    defer fake.deinit();
    const sheet = try fuzzF1bEnv(&fake);

    var checked: usize = 0;
    for (f1b_names) |name| {
        for (f1b_arg_shapes) |first_arg| {
            var buf: [256]u8 = undefined;
            if (buildF1bCall(&buf, name, &.{first_arg})) |one| {
                try sweepShape(&arena_state, &fake, sheet, one, &checked);
            }
            for (f1b_arg_shapes) |second_arg| {
                var pair_buf: [256]u8 = undefined;
                const two = buildF1bCall(&pair_buf, name, &.{ first_arg, second_arg }) orelse continue;
                try sweepShape(&arena_state, &fake, sheet, two, &checked);
            }
            // A lookup materializes its table into the arena, and this
            // sweep runs tens of thousands of them. Reclaiming per
            // shape keeps the enumeration's memory bounded by its
            // widest single call rather than by its length.
            _ = arena_state.reset(.retain_capacity);
        }
    }
    // A sweep that silently stopped enumerating would still pass, so
    // the count is asserted as a floor rather than left to the loops.
    try testing.expect(checked > 20_000);
}

// ─── M4f: the F1c-text batch (§7, nineteen names) ────────────────
//
// Oracle-first, and — as at M4c, M4d and M4e — the three committed
// manifests decide nothing here: §8.2's evidence is eighteen operator
// and literal cells plus `SQRT(-1)`, and no text name appears in any of
// them. So the batch ships `spec_pinned`, its oracle-row count is pinned
// at **zero**, and the same three-valued checker guards it.
//
// What is different about this row is that a fixture is not fully
// specified by a formula and a result any more. Two of them can hold at
// once — `LEN("😀")` is 2 under CV1 and 1 under CV2, and both are
// right — so the case carries the compatibility version it was written
// for, and the version-free cases are asserted under BOTH.

/// One F1c fixture. `cv` is null for the majority of rows, which mean
/// the same thing under either version and are run under both.
const F1cCase = struct {
    func: []const u8,
    formula: []const u8,
    expect: Expect,
    cv: ?run_inputs.CompatibilityVersion = null,
    evidence: Evidence = .spec_pinned,
    note: []const u8 = "",
};

/// The environment every F1c fixture reads.
fn putF1cCells(h: *Harness) !void {
    // §5.3b's provenance column again, cell for cell with M4c/M4d/M4e,
    // so a text function's coercion can be compared against the same
    // table the numeric batches were.
    try h.put("A1", num(10));
    try h.put("A2", .{ .text = "abc" });
    try h.put("A3", .{ .text = "" });
    try h.put("A4", .{ .boolean = true });
    try h.put("A5", value.ScalarValue.errorOf(.div0));
    // A7 is a true blank — deliberately not stored.
    try h.put("A8", .{ .text = "7" });

    // A column to join, holding one of each thing a range can hold.
    try h.put("C1", .{ .text = "a" });
    try h.put("C2", .{ .text = "b" });
    // C3 is blank, which CONCAT contributes nothing for and TEXTJOIN
    // skips only when asked to.
    try h.put("C4", num(3));

    // The astral cell: one code point, two UTF-16 code units. Every
    // per-version fixture in the table reads it.
    try h.put("E1", .{ .text = "a\u{1F600}b" });
    // The length-changing casing cell.
    try h.put("E2", .{ .text = "Straße" });
    // A cell whose fold expands, for SEARCH's positional map.
    try h.put("E3", .{ .text = "aßb" });
}

const f1c_cases = [_]F1cCase{
    // ── LEN: the function the compatibility version is about ──
    .{ .func = "LEN", .formula = "LEN(\"hello\")", .expect = .{ .number = 5 } },
    .{ .func = "LEN", .formula = "LEN(\"\")", .expect = .{ .number = 0 } },
    // Blank is the empty string to a text slot, not an error.
    .{ .func = "LEN", .formula = "LEN(A7)", .expect = .{ .number = 0 } },
    // A number coerces through §5.3b's text column before it is
    // measured: `LEN(10)` is 2, not `#VALUE!`.
    .{ .func = "LEN", .formula = "LEN(A1)", .expect = .{ .number = 2 } },
    .{ .func = "LEN", .formula = "LEN(A4)", .expect = .{ .number = 4 }, .note = "TRUE prints as four characters" },
    .{ .func = "LEN", .formula = "LEN(\"café\")", .expect = .{ .number = 4 }, .note = "characters, not the five bytes" },
    .{ .func = "LEN", .formula = "LEN(E1)", .expect = .{ .number = 4 }, .cv = .cv1 },
    .{ .func = "LEN", .formula = "LEN(E1)", .expect = .{ .number = 3 }, .cv = .cv2 },
    // Not grapheme clustering, under either version.
    .{ .func = "LEN", .formula = "LEN(\"e\u{0301}\")", .expect = .{ .number = 2 } },

    // ── LEFT / RIGHT / MID ──
    .{ .func = "LEFT", .formula = "LEFT(\"hello\",2)", .expect = .{ .text = "he" } },
    .{ .func = "LEFT", .formula = "LEFT(\"hello\")", .expect = .{ .text = "h" }, .note = "the omitted count is 1" },
    .{ .func = "LEFT", .formula = "LEFT(\"hello\",)", .expect = .{ .text = "" }, .note = "an omitted ARGUMENT is blank, and blank is 0" },
    .{ .func = "LEFT", .formula = "LEFT(\"hello\",99)", .expect = .{ .text = "hello" } },
    .{ .func = "LEFT", .formula = "LEFT(\"hello\",-1)", .expect = .{ .err = .value } },
    .{ .func = "LEFT", .formula = "LEFT(\"hello\",2.9)", .expect = .{ .text = "he" }, .note = "a fractional count truncates" },
    .{ .func = "LEFT", .formula = "LEFT(E1,2)", .expect = .{ .text = "a\u{1F600}" }, .cv = .cv2 },
    .{ .func = "RIGHT", .formula = "RIGHT(\"hello\",3)", .expect = .{ .text = "llo" } },
    .{ .func = "RIGHT", .formula = "RIGHT(\"hello\")", .expect = .{ .text = "o" } },
    .{ .func = "RIGHT", .formula = "RIGHT(\"hello\",99)", .expect = .{ .text = "hello" } },
    .{ .func = "RIGHT", .formula = "RIGHT(\"hello\",0)", .expect = .{ .text = "" } },
    .{ .func = "MID", .formula = "MID(\"hello\",2,3)", .expect = .{ .text = "ell" } },
    .{ .func = "MID", .formula = "MID(\"hello\",1,99)", .expect = .{ .text = "hello" } },
    .{ .func = "MID", .formula = "MID(\"hello\",9,2)", .expect = .{ .text = "" }, .note = "past the end is empty, not an error" },
    .{ .func = "MID", .formula = "MID(\"hello\",0,2)", .expect = .{ .err = .value }, .note = "the start is 1-based" },
    .{ .func = "MID", .formula = "MID(\"hello\",1,-1)", .expect = .{ .err = .value } },
    .{ .func = "MID", .formula = "MID(E1,2,1)", .expect = .{ .text = "\u{1F600}" }, .cv = .cv2 },

    // ── REPLACE ──
    .{ .func = "REPLACE", .formula = "REPLACE(\"abcdef\",2,3,\"XY\")", .expect = .{ .text = "aXYef" } },
    .{ .func = "REPLACE", .formula = "REPLACE(\"abc\",1,0,\"X\")", .expect = .{ .text = "Xabc" }, .note = "replacing nothing inserts" },
    .{ .func = "REPLACE", .formula = "REPLACE(\"abc\",9,3,\"X\")", .expect = .{ .text = "abcX" }, .note = "past the end appends" },
    .{ .func = "REPLACE", .formula = "REPLACE(\"abc\",0,1,\"X\")", .expect = .{ .err = .value } },
    .{ .func = "REPLACE", .formula = "REPLACE(\"abc\",1,-1,\"X\")", .expect = .{ .err = .value } },

    // ── FIND: case-sensitive, no wildcards ──
    .{ .func = "FIND", .formula = "FIND(\"l\",\"hello\")", .expect = .{ .number = 3 } },
    .{ .func = "FIND", .formula = "FIND(\"l\",\"hello\",4)", .expect = .{ .number = 4 } },
    .{ .func = "FIND", .formula = "FIND(\"L\",\"hello\")", .expect = .{ .err = .value }, .note = "case-SENSITIVE, which is the whole difference from SEARCH" },
    .{ .func = "FIND", .formula = "FIND(\"*\",\"a*b\")", .expect = .{ .number = 2 }, .note = "no wildcards: a star is a star" },
    .{ .func = "FIND", .formula = "FIND(\"z\",\"hello\")", .expect = .{ .err = .value } },
    .{ .func = "FIND", .formula = "FIND(\"h\",\"hello\",0)", .expect = .{ .err = .value } },
    .{ .func = "FIND", .formula = "FIND(\"\",\"hello\")", .expect = .{ .number = 1 } },
    .{ .func = "FIND", .formula = "FIND(\"b\",E1)", .expect = .{ .number = 4 }, .cv = .cv1 },
    .{ .func = "FIND", .formula = "FIND(\"b\",E1)", .expect = .{ .number = 3 }, .cv = .cv2 },

    // ── SEARCH: folded, wildcards active ──
    .{ .func = "SEARCH", .formula = "SEARCH(\"L\",\"hello\")", .expect = .{ .number = 3 }, .note = "case-INsensitive" },
    .{ .func = "SEARCH", .formula = "SEARCH(\"l*o\",\"hello\")", .expect = .{ .number = 3 } },
    .{ .func = "SEARCH", .formula = "SEARCH(\"h?l\",\"hello\")", .expect = .{ .number = 1 }, .note = "`?` is one character, so h-e-l matches at 1" },
    .{ .func = "SEARCH", .formula = "SEARCH(\"h?o\",\"hello\")", .expect = .{ .err = .value } },
    .{ .func = "SEARCH", .formula = "SEARCH(\"~*\",\"a*b\")", .expect = .{ .number = 2 }, .note = "the escape survives, same as in a criterion" },
    .{ .func = "SEARCH", .formula = "SEARCH(\"z\",\"hello\")", .expect = .{ .err = .value } },
    .{ .func = "SEARCH", .formula = "SEARCH(\"l\",\"hello\",4)", .expect = .{ .number = 4 } },
    .{ .func = "SEARCH", .formula = "SEARCH(\"SS\",E3)", .expect = .{ .number = 2 }, .note = "the fold expands, and the position is the caller's" },
    .{ .func = "SEARCH", .formula = "SEARCH(\"b\",E1)", .expect = .{ .number = 4 }, .cv = .cv1 },
    .{ .func = "SEARCH", .formula = "SEARCH(\"b\",E1)", .expect = .{ .number = 3 }, .cv = .cv2 },

    // ── casing ──
    .{ .func = "UPPER", .formula = "UPPER(\"hello\")", .expect = .{ .text = "HELLO" } },
    .{ .func = "UPPER", .formula = "UPPER(E2)", .expect = .{ .text = "STRASSE" }, .note = "ß→SS is length-changing; a simple mapping cannot do it" },
    .{ .func = "UPPER", .formula = "UPPER(\"café\")", .expect = .{ .text = "CAFÉ" } },
    .{ .func = "LOWER", .formula = "LOWER(\"HeLLo\")", .expect = .{ .text = "hello" } },
    .{ .func = "LOWER", .formula = "LOWER(\"ΟΔΟΣ\")", .expect = .{ .text = "οδο\u{03C2}" }, .note = "Final_Sigma: ς at the end of a word" },
    .{ .func = "LOWER", .formula = "LOWER(\"I\")", .expect = .{ .text = "i" }, .note = "locale-neutral; a Turkish Excel answers ı" },

    // ── TRIM ──
    .{ .func = "TRIM", .formula = "TRIM(\"  a  b  \")", .expect = .{ .text = "a b" } },
    .{ .func = "TRIM", .formula = "TRIM(\"abc\")", .expect = .{ .text = "abc" } },
    .{ .func = "TRIM", .formula = "TRIM(\"   \")", .expect = .{ .text = "" } },
    .{ .func = "TRIM", .formula = "TRIM(\"a\u{00A0}b\")", .expect = .{ .text = "a\u{00A0}b" }, .note = "a non-breaking space is not a space; CLEAN is M8c" },

    // ── EXACT ──
    .{ .func = "EXACT", .formula = "EXACT(\"abc\",\"abc\")", .expect = .{ .boolean = true } },
    .{ .func = "EXACT", .formula = "EXACT(\"abc\",\"ABC\")", .expect = .{ .boolean = false }, .note = "the one comparison that does not fold" },
    .{ .func = "EXACT", .formula = "EXACT(\"ß\",\"SS\")", .expect = .{ .boolean = false }, .note = "…so fold-equality does not reach it either" },
    .{ .func = "EXACT", .formula = "EXACT(\"\",A7)", .expect = .{ .boolean = true } },

    // ── SUBSTITUTE ──
    .{ .func = "SUBSTITUTE", .formula = "SUBSTITUTE(\"a-b-c\",\"-\",\"+\")", .expect = .{ .text = "a+b+c" } },
    .{ .func = "SUBSTITUTE", .formula = "SUBSTITUTE(\"a-b-c\",\"-\",\"+\",2)", .expect = .{ .text = "a-b+c" } },
    .{ .func = "SUBSTITUTE", .formula = "SUBSTITUTE(\"a-b\",\"-\",\"+\",9)", .expect = .{ .text = "a-b" }, .note = "an instance that is not there changes nothing" },
    .{ .func = "SUBSTITUTE", .formula = "SUBSTITUTE(\"aAa\",\"a\",\"x\")", .expect = .{ .text = "xAx" }, .note = "case-SENSITIVE" },
    .{ .func = "SUBSTITUTE", .formula = "SUBSTITUTE(\"abc\",\"\",\"x\")", .expect = .{ .text = "abc" }, .note = "an empty needle matches nowhere rather than everywhere" },
    .{ .func = "SUBSTITUTE", .formula = "SUBSTITUTE(\"abc\",\"b\",\"\")", .expect = .{ .text = "ac" } },
    .{ .func = "SUBSTITUTE", .formula = "SUBSTITUTE(\"a-b\",\"-\",\"+\",0)", .expect = .{ .err = .value } },

    // ── REPT ──
    .{ .func = "REPT", .formula = "REPT(\"ab\",3)", .expect = .{ .text = "ababab" } },
    .{ .func = "REPT", .formula = "REPT(\"ab\",0)", .expect = .{ .text = "" } },
    .{ .func = "REPT", .formula = "REPT(\"ab\",-1)", .expect = .{ .err = .value } },
    .{ .func = "REPT", .formula = "REPT(\"ab\",2.9)", .expect = .{ .text = "abab" } },
    .{ .func = "REPT", .formula = "REPT(\"ab\",100000)", .expect = .{ .err = .value }, .note = "§9's cell cap, as Excel's own #VALUE! rather than a refusal" },

    // ── VALUE ──
    .{ .func = "VALUE", .formula = "VALUE(\"42\")", .expect = .{ .number = 42 } },
    .{ .func = "VALUE", .formula = "VALUE(\" -2.5e3 \")", .expect = .{ .number = -2500 } },
    .{ .func = "VALUE", .formula = "VALUE(\"abc\")", .expect = .{ .err = .value } },
    .{ .func = "VALUE", .formula = "VALUE(A8)", .expect = .{ .number = 7 } },

    // ── CHAR / CODE ──
    .{ .func = "CHAR", .formula = "CHAR(65)", .expect = .{ .text = "A" } },
    .{ .func = "CHAR", .formula = "CHAR(0)", .expect = .{ .err = .value } },
    .{ .func = "CHAR", .formula = "CHAR(256)", .expect = .{ .err = .value } },
    .{ .func = "CHAR", .formula = "CHAR(128)", .expect = .{ .text = "€" }, .note = "windows-1252, not Latin-1: 0x80 is the euro sign" },
    .{ .func = "CHAR", .formula = "CHAR(233)", .expect = .{ .text = "é" } },
    .{ .func = "CODE", .formula = "CODE(\"A\")", .expect = .{ .number = 65 } },
    .{ .func = "CODE", .formula = "CODE(\"abc\")", .expect = .{ .number = 97 }, .note = "the first character only" },
    .{ .func = "CODE", .formula = "CODE(\"\")", .expect = .{ .err = .value } },
    .{ .func = "CODE", .formula = "CODE(\"€\")", .expect = .{ .number = 128 } },
    .{ .func = "CODE", .formula = "CODE(\"\u{1F600}\")", .expect = .{ .number = 63 }, .note = "outside the code page: Excel's `?` substitution" },

    // ── joining ──
    .{ .func = "CONCAT", .formula = "CONCAT(\"a\",\"b\",\"c\")", .expect = .{ .text = "abc" } },
    .{ .func = "CONCAT", .formula = "CONCAT(C1:C4)", .expect = .{ .text = "ab3" }, .note = "a RANGE, which is the whole difference from CONCATENATE" },
    .{ .func = "CONCAT", .formula = "CONCAT(\"x\",A5)", .expect = .{ .err = .div0 }, .note = "an error in an argument is the answer" },
    .{ .func = "CONCATENATE", .formula = "CONCATENATE(\"a\",\"b\")", .expect = .{ .text = "ab" } },
    .{ .func = "CONCATENATE", .formula = "CONCATENATE(\"n=\",A1)", .expect = .{ .text = "n=10" } },
    .{ .func = "TEXTJOIN", .formula = "TEXTJOIN(\"-\",TRUE,\"a\",\"b\")", .expect = .{ .text = "a-b" } },
    .{ .func = "TEXTJOIN", .formula = "TEXTJOIN(\"-\",TRUE,C1:C4)", .expect = .{ .text = "a-b-3" }, .note = "the blank is skipped" },
    .{ .func = "TEXTJOIN", .formula = "TEXTJOIN(\"-\",FALSE,C1:C4)", .expect = .{ .text = "a-b--3" }, .note = "…and kept, which is what the flag is for" },
    .{ .func = "TEXTJOIN", .formula = "TEXTJOIN(\"\",TRUE,C1:C4)", .expect = .{ .text = "ab3" } },
};

fn f1cOptions(h: *Harness, cv: run_inputs.CompatibilityVersion) Options {
    var opts = h.options();
    opts.text_compat = cv;
    return opts;
}

test "M4f: every F1c fixture evaluates to what the spec says, under its version" {
    for (f1c_cases) |c| {
        // A case with no version means the same thing under both, and is
        // asserted under both — which is how a rule that only LOOKS
        // version-free gets caught.
        const versions: []const run_inputs.CompatibilityVersion =
            if (c.cv) |v| &.{v} else &.{ .cv1, .cv2 };
        for (versions) |cv| {
            var h: Harness = undefined;
            try h.init(testing.allocator);
            defer h.deinit();
            try putF1cCells(&h);

            const v = h.evalOpts(c.formula, f1cOptions(&h, cv)) catch |e| {
                std.debug.print("F1c `{s}` ({t}) refused: {t}\n", .{ c.formula, cv, e });
                return e;
            };
            expectValue(c.expect, v) catch |e| {
                std.debug.print("F1c `{s}` ({s}, {t}): wrong value\n", .{ c.formula, c.func, cv });
                return e;
            };
        }
    }
}

test "M4f: all nineteen frozen names resolve, and each has a fixture" {
    var it = registry.inventory();
    var batch: usize = 0;
    while (it.next()) |e| {
        if (!std.mem.eql(u8, e.milestone, "M4f")) continue;
        batch += 1;

        if (registry.lookup(e.name) == null) {
            std.debug.print("F1c name does not resolve: {s}\n", .{e.name});
            return error.UnregisteredBatchFunction;
        }
        var fixtures: usize = 0;
        for (f1c_cases) |c| {
            if (std.mem.eql(u8, c.func, e.name)) fixtures += 1;
        }
        if (fixtures == 0) {
            std.debug.print("F1c name has no fixture: {s}\n", .{e.name});
            return error.UnfixturedBatchFunction;
        }
    }
    try testing.expectEqual(@as(usize, 19), batch);

    for (f1c_cases) |c| {
        var found = false;
        var it2 = registry.inventory();
        while (it2.next()) |e| {
            if (std.mem.eql(u8, e.name, c.func) and std.mem.eql(u8, e.milestone, "M4f")) found = true;
        }
        if (!found) {
            std.debug.print("fixture names a function outside F1c: {s}\n", .{c.func});
            return error.FixtureOutsideBatch;
        }
    }
}

test "M4f: the evidence label on every fixture is true of the committed manifests" {
    var oracle_rows: usize = 0;
    var excluded_rows: usize = 0;
    for (f1c_cases) |c| {
        switch (try manifestVerdict(c.formula)) {
            .decided => {
                if (c.evidence != .oracle) {
                    std.debug.print("`{s}` is decided by a manifest but ships spec-pinned\n", .{c.formula});
                    return error.UnderstatedEvidence;
                }
                oracle_rows += 1;
            },
            .excluded => {
                if (c.evidence != .spec_pinned) {
                    std.debug.print("`{s}` claims evidence from an EXCLUDED manifest cell\n", .{c.formula});
                    return error.ExcludedCellClaimedAsEvidence;
                }
                excluded_rows += 1;
            },
            .silent => {
                if (c.evidence != .spec_pinned) {
                    std.debug.print("`{s}` claims oracle evidence no manifest holds\n", .{c.formula});
                    return error.UnbackedOracleClaim;
                }
            },
        }
    }
    // Zero again, and for the same reason — but this row's zero carries
    // more weight than M4e's did. `UPPER("ß")` is the case §5.4b marks
    // oracle-pinned, and the Excel adapter that would pin it is parked
    // (§8.2). So it ships spec-pinned and FLAGGED: the count below is
    // what has to move when that leg runs, and this is the row that
    // will move it.
    try testing.expectEqual(@as(usize, 0), oracle_rows);
    try testing.expectEqual(@as(usize, 0), excluded_rows);
}

test "M4f: the compatibility version is exactly the seven names that index" {
    // Registry data, checked in both directions: a name that counts or
    // positions in §5.4d's units must declare it, and a name that
    // operates on whole strings must not.
    const indexing = [_][]const u8{
        "LEN", "LEFT", "RIGHT", "MID", "REPLACE", "FIND", "SEARCH",
    };
    var it = registry.inventory();
    while (it.next()) |e| {
        if (!std.mem.eql(u8, e.milestone, "M4f")) continue;
        var listed = false;
        for (indexing) |n| {
            if (std.mem.eql(u8, n, e.name)) listed = true;
        }
        if (listed != registry.lookup(e.name).?.cv_sensitive) {
            std.debug.print("{s}: cv_sensitive={} but the list says {}\n", .{
                e.name,
                registry.lookup(e.name).?.cv_sensitive,
                listed,
            });
            return error.CvFlagMismatch;
        }
    }
}

test "M4f: the same call, two versions, two right answers" {
    // The compatibility version end to end, over one cell, in every
    // function that can see it. Read as a block this is the row's
    // headline: `a😀b` is four characters to CV1 and three to CV2, and
    // every position downstream of that shifts with it.
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try putF1cCells(&h);

    const Row = struct { formula: []const u8, cv1: f64, cv2: f64 };
    for ([_]Row{
        .{ .formula = "LEN(E1)", .cv1 = 4, .cv2 = 3 },
        .{ .formula = "FIND(\"b\",E1)", .cv1 = 4, .cv2 = 3 },
        .{ .formula = "SEARCH(\"B\",E1)", .cv1 = 4, .cv2 = 3 },
        .{ .formula = "LEN(MID(E1,2,2))", .cv1 = 2, .cv2 = 2 },
    }) |r| {
        const a = try h.evalOpts(r.formula, f1cOptions(&h, .cv1));
        try testing.expectEqual(r.cv1, a.scalar.number);
        const b = try h.evalOpts(r.formula, f1cOptions(&h, .cv2));
        try testing.expectEqual(r.cv2, b.scalar.number);
    }

    // LEFT and RIGHT are not in §5.4d's list of five, and they are
    // version-dependent all the same: a count of characters cannot mean
    // one thing in `MID` and another in `LEFT` within one workbook.
    // Recorded as a decision rather than left as an inconsistency.
    try testing.expectEqualStrings(
        "a",
        (try h.evalOpts("LEFT(E1,1)", f1cOptions(&h, .cv1))).scalar.text,
    );
    try testing.expectEqualStrings(
        "b",
        (try h.evalOpts("RIGHT(E1,1)", f1cOptions(&h, .cv1))).scalar.text,
    );
}

test "M4f: a CV1 index into an astral character refuses rather than halving it" {
    // §5.4d's one unrepresentable case. Excel hands back a lone
    // surrogate; UTF-8 has no such thing, so this is a typed refusal —
    // and it is a refusal only under CV1, because under CV2 there is no
    // half to ask for.
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try putF1cCells(&h);

    for ([_][]const u8{ "MID(E1,2,1)", "LEFT(E1,2)", "RIGHT(E1,2)", "REPLACE(E1,2,1,\"x\")" }) |f| {
        try testing.expectError(
            error.ResultNotRepresentable,
            h.evalOpts(f, f1cOptions(&h, .cv1)),
        );
        // The same formula under CV2 is an ordinary answer.
        _ = try h.evalOpts(f, f1cOptions(&h, .cv2));
    }
}

test "M4f: the match policy is registry data, and behaviour agrees with it" {
    // §5.4b: every text function's policy is explicit. The four `.raw`
    // rows are the case-sensitive ones, and the pairing is asserted
    // against what the functions actually do rather than only against
    // the table.
    const raw = [_][]const u8{ "FIND", "SUBSTITUTE", "EXACT", "CODE" };
    var it = registry.inventory();
    while (it.next()) |e| {
        if (!std.mem.eql(u8, e.milestone, "M4f")) continue;
        var listed = false;
        for (raw) |n| {
            if (std.mem.eql(u8, n, e.name)) listed = true;
        }
        const want: value.MatchPolicy = if (listed) .raw else .folded;
        try testing.expectEqual(want, registry.lookup(e.name).?.match_policy);
    }

    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try putF1cCells(&h);
    // The same question, asked of the raw function and the folded one.
    try testing.expectEqual(value.KnownError.value, (try h.scalar("FIND(\"L\",\"hello\")")).err.known);
    try testing.expectEqual(@as(f64, 3), (try h.scalar("SEARCH(\"L\",\"hello\")")).number);
    // …and `SEARCH` is the only M4f name inside the comparator, which is
    // what `collation_sensitive` is for.
    var it2 = registry.inventory();
    while (it2.next()) |e| {
        if (!std.mem.eql(u8, e.milestone, "M4f")) continue;
        const want = std.mem.eql(u8, e.name, "SEARCH");
        try testing.expectEqual(want, registry.lookup(e.name).?.collation_sensitive);
    }
}

test "M4f: casing_v1 through UPPER and LOWER, including what folding cannot do" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try putF1cCells(&h);

    // Length-changing, and the reason `casing_v1` is not the fold: the
    // fold's answer for `ß` is lowercase `ss`.
    try testing.expectEqualStrings("STRASSE", (try h.scalar("UPPER(\"Straße\")")).text);
    try testing.expectEqual(@as(f64, 7), (try h.scalar("LEN(UPPER(\"Straße\"))")).number);
    try testing.expectEqual(@as(f64, 6), (try h.scalar("LEN(\"Straße\")")).number);
    // Ligature, astral, and the dotted I.
    try testing.expectEqualStrings("FFI", (try h.scalar("UPPER(\"ﬃ\")")).text);
    try testing.expectEqualStrings("\u{10400}", (try h.scalar("UPPER(\"\u{10428}\")")).text);
    try testing.expectEqualStrings("i\u{0307}", (try h.scalar("LOWER(\"İ\")")).text);
    // Final_Sigma in both positions, from one formula each.
    try testing.expectEqualStrings("οδο\u{03C2}", (try h.scalar("LOWER(\"ΟΔΟΣ\")")).text);
    try testing.expectEqualStrings("α\u{03C3}α", (try h.scalar("LOWER(\"ΑΣΑ\")")).text);
    // Uppercasing is not injective, so it does not round-trip — stated
    // here so nobody later "fixes" it.
    try testing.expectEqualStrings("strasse", (try h.scalar("LOWER(UPPER(\"Straße\"))")).text);
}

test "M4f: text a formula produced can be written back out (the ST_Xstring codec)" {
    // `CHAR` is the first function in the ladder that can produce a C0
    // control, which XML cannot carry literally and ST_Xstring exists to
    // escape. The codec is M4b1's; this is the row that first NEEDS it,
    // so the round-trip is asserted from a formula result rather than
    // from a hand-written string.
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();

    for ([_][]const u8{ "CHAR(1)", "CHAR(13)", "CHAR(9)", "CHAR(65)&CHAR(2)" }) |src| {
        const produced = (try h.scalar(src)).text;
        const encoded = try decode.encodeAuthoredString(testing.allocator, produced);
        defer testing.allocator.free(encoded);
        const back = try decode.decodeCarrier(testing.allocator, .string, encoded);
        defer testing.allocator.free(back);
        try testing.expectEqualStrings(produced, back);
    }

    // A literal `_x0041_` a formula produced must come back as those
    // seven characters rather than as `A` — the escape of the escape.
    const literal = (try h.scalar("\"_x0041_\"")).text;
    const encoded = try decode.encodeAuthoredString(testing.allocator, literal);
    defer testing.allocator.free(encoded);
    try testing.expect(std.mem.indexOf(u8, encoded, "_x005F_") != null);
    const back = try decode.decodeCarrier(testing.allocator, .string, encoded);
    defer testing.allocator.free(back);
    try testing.expectEqualStrings("_x0041_", back);
}

test "M4f: §9's cell cap is Excel's #VALUE!, and it is checked before the allocation" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();

    // At the cap, under it, and over it. The over-cap call must not
    // allocate a gigabyte on its way to the answer, which is why the
    // check is on the count rather than on the result.
    try testing.expectEqual(@as(f64, 32767), (try h.scalar("LEN(REPT(\"a\",32767))")).number);
    try testing.expectEqual(value.KnownError.value, (try h.scalar("REPT(\"a\",32768)")).err.known);
    try testing.expectEqual(value.KnownError.value, (try h.scalar("REPT(\"ab\",1000000000)")).err.known);
    // Concatenation reaches the same cap by a different road. Before
    // M4f no formula could: a literal is bounded by the 8 192-byte
    // formula length, so `REPT` is what made `&` able to overflow a cell
    // and the cap had to move to the operator with it.
    try testing.expectEqual(
        value.KnownError.value,
        (try h.scalar("REPT(\"a\",20000)&REPT(\"b\",20000)&\"\"")).err.known,
    );

    // …and TEXTJOIN's dense walk stops at the cap rather than emitting a
    // delimiter per row of a whole column. The answer is the same
    // `#VALUE!`; what is under test is that it arrives.
    try putF1cCells(&h);
    try testing.expectEqual(
        value.KnownError.value,
        (try h.scalar("TEXTJOIN(\"-\",FALSE,C:C)")).err.known,
    );
    // With empties skipped the same call is the three cells that are
    // there, because then a blank and an absence are the same thing.
    try testing.expectEqualStrings("a-b-3", (try h.scalar("TEXTJOIN(\"-\",TRUE,C:C)")).text);
}

test "M4f: the batch agrees across both rule tables, everywhere" {
    // F1c-text computes no floating-point value, so — unlike M4d and
    // M4e — there is nothing here for the two rule tables to disagree
    // about. Asserted rather than assumed: `VALUE` parses numbers, and a
    // parse is exactly where a divergence would hide.
    for (f1c_cases) |c| {
        const cv = c.cv orelse .cv1;
        var excel: Harness = undefined;
        try excel.init(testing.allocator);
        defer excel.deinit();
        try putF1cCells(&excel);
        var opts_e = f1cOptions(&excel, cv);
        opts_e.fidelity = .excel;

        var ieee: Harness = undefined;
        try ieee.init(testing.allocator);
        defer ieee.deinit();
        try putF1cCells(&ieee);
        var opts_i = f1cOptions(&ieee, cv);
        opts_i.fidelity = .ieee;

        const a = excel.evalOpts(c.formula, opts_e) catch continue;
        const b = try ieee.evalOpts(c.formula, opts_i);
        if (a == .scalar and b == .scalar and !a.scalar.eql(b.scalar)) {
            std.debug.print("`{s}` differs between rule tables\n", .{c.formula});
            return error.UnexpectedFidelityDivergence;
        }
    }
}

test "M4f: every name against every argument shape, in both versions" {
    // M4d's and M4e's sweep, run over this batch: each name padded to
    // its own minimum arity so a function with three required arguments
    // is actually reached, and each input evaluated twice so nothing is
    // order-dependent. Refusals are allowed — a wrong SHAPE is a legal
    // answer — but a panic, a leak or a nondeterministic result is not.
    const shapes = [_][]const u8{ "\"abc\"", "1", "TRUE", "A5", "A7", "C1:C4", "{1,2}" };
    var checked: usize = 0;
    for ([_]run_inputs.CompatibilityVersion{ .cv1, .cv2 }) |cv| {
        var it = registry.inventory();
        while (it.next()) |e| {
            if (!std.mem.eql(u8, e.milestone, "M4f")) continue;
            const f = registry.lookup(e.name).?;
            for (shapes) |s| {
                var buf: [256]u8 = undefined;
                var w: std.Io.Writer = .fixed(&buf);
                try w.print("{s}({s}", .{ e.name, s });
                // Pad by repeating the last shape, so every name is put
                // in front of its own implementation rather than
                // refused on arity before it gets there.
                var k: usize = 1;
                while (k < f.arity.min) : (k += 1) try w.print(",{s}", .{s});
                try w.writeAll(")");
                const src = w.buffered();

                var h: Harness = undefined;
                try h.init(testing.allocator);
                defer h.deinit();
                try putF1cCells(&h);
                const first = h.evalOpts(src, f1cOptions(&h, cv));

                var h2: Harness = undefined;
                try h2.init(testing.allocator);
                defer h2.deinit();
                try putF1cCells(&h2);
                const second = h2.evalOpts(src, f1cOptions(&h2, cv));

                if (first) |a| {
                    const b = second catch {
                        std.debug.print("`{s}` succeeded then refused\n", .{src});
                        return error.NondeterministicEvaluation;
                    };
                    if (a == .scalar and b == .scalar and !a.scalar.eql(b.scalar)) {
                        std.debug.print("`{s}` gave two answers\n", .{src});
                        return error.NondeterministicEvaluation;
                    }
                } else |err| {
                    if (second) |_| {
                        std.debug.print("`{s}` refused then succeeded\n", .{src});
                        return error.NondeterministicEvaluation;
                    } else |err2| {
                        if (err != err2) return error.NondeterministicEvaluation;
                    }
                }
                checked += 1;
            }
        }
    }
    // Nineteen names × seven shapes × two versions.
    try testing.expectEqual(@as(usize, 19 * 7 * 2), checked);
}

// ─── M4g: the F1c-date batch (§7, fifteen names) ─────────────────
//
// Oracle-first, and — as at every F-batch so far — the three committed
// manifests decide nothing: §8.2's evidence is eighteen operator and
// literal cells plus `SQRT(-1)`. So the batch ships `spec_pinned` with
// its oracle-row count pinned at **zero**, guarded by the same
// three-valued checker.
//
// What is different here is the second axis. M4f's fixtures carried a
// compatibility version; these carry an EPOCH, and the two epochs
// disagree about the value of every date rather than about its length.
// A fixture with no epoch means the same thing under both and is
// asserted under both.

/// 2020-01-01 in each system — the anchor every dated fixture below is
/// written against, so a reader can check one number instead of forty.
const serial_2020_1900: f64 = 43831;
const serial_2020_1904: f64 = 42369;
/// 2020-01-01T12:00:00Z as Unix milliseconds, for the two volatiles.
const now_2020_noon_ms: i64 = 1_577_880_000_000;

const F1dCase = struct {
    func: []const u8,
    formula: []const u8,
    expect: Expect,
    system: ?run_inputs.DateSystem = null,
    evidence: Evidence = .spec_pinned,
    note: []const u8 = "",
};

const f1d_cases = [_]F1dCase{
    // ── DATE: the constructor, and its two overflow rules ──
    .{ .func = "DATE", .formula = "DATE(2020,1,1)", .expect = .{ .number = serial_2020_1900 }, .system = .d1900 },
    .{ .func = "DATE", .formula = "DATE(2020,1,1)", .expect = .{ .number = serial_2020_1904 }, .system = .d1904 },
    .{ .func = "DATE", .formula = "DATE(2020,13,1)", .expect = .{ .number = 44197 }, .system = .d1900, .note = "month 13 is January of the next year" },
    .{ .func = "DATE", .formula = "DATE(2020,1,32)", .expect = .{ .number = 43862 }, .system = .d1900, .note = "day 32 is the 1st of February" },
    .{ .func = "DATE", .formula = "DATE(2020,0,1)", .expect = .{ .number = 43800 }, .system = .d1900, .note = "month 0 is December of the previous year" },
    .{ .func = "DATE", .formula = "DATE(20,1,1)", .expect = .{ .number = 7306 }, .system = .d1900, .note = "1920-01-01: a numeric year below 1900 is 1900+y — NOT the text grammar's window" },
    .{ .func = "DATE", .formula = "DATE(-1,1,1)", .expect = .{ .err = .num } },
    .{ .func = "DATE", .formula = "DATE(10000,1,1)", .expect = .{ .err = .num } },

    // ── the readers ──
    .{ .func = "YEAR", .formula = "YEAR(43831)", .expect = .{ .number = 2020 }, .system = .d1900 },
    .{ .func = "YEAR", .formula = "YEAR(42369)", .expect = .{ .number = 2020 }, .system = .d1904 },
    .{ .func = "YEAR", .formula = "YEAR(-1)", .expect = .{ .err = .num }, .note = "no date precedes the epoch" },
    .{ .func = "MONTH", .formula = "MONTH(43831)", .expect = .{ .number = 1 }, .system = .d1900 },
    .{ .func = "MONTH", .formula = "MONTH(43862)", .expect = .{ .number = 2 }, .system = .d1900 },
    .{ .func = "DAY", .formula = "DAY(43831)", .expect = .{ .number = 1 }, .system = .d1900 },
    .{ .func = "DAY", .formula = "DAY(43861)", .expect = .{ .number = 31 }, .system = .d1900 },
    // The fractional part is truncated, not rounded: 43831.9 is still
    // the 1st.
    .{ .func = "DAY", .formula = "DAY(43831.9)", .expect = .{ .number = 1 }, .system = .d1900 },

    // ── §5.4a's two invented days, read back ──
    .{ .func = "DAY", .formula = "DAY(0)", .expect = .{ .number = 0 }, .system = .d1900, .note = "1900-01-00: a day number no calendar has" },
    .{ .func = "MONTH", .formula = "MONTH(0)", .expect = .{ .number = 1 }, .system = .d1900 },
    .{ .func = "YEAR", .formula = "YEAR(0)", .expect = .{ .number = 1900 }, .system = .d1900 },
    .{ .func = "DAY", .formula = "DAY(60)", .expect = .{ .number = 29 }, .system = .d1900, .note = "1900-02-29, which never happened" },
    .{ .func = "MONTH", .formula = "MONTH(60)", .expect = .{ .number = 2 }, .system = .d1900 },
    .{ .func = "DAY", .formula = "DAY(59)", .expect = .{ .number = 28 }, .system = .d1900, .note = "the day before the gap" },
    .{ .func = "DAY", .formula = "DAY(61)", .expect = .{ .number = 1 }, .system = .d1900, .note = "and the day after it: 1900-03-01" },
    // Under 1904 the same serials are ordinary days in 1904, which is
    // the clearest statement that the epoch is not a display setting.
    .{ .func = "DAY", .formula = "DAY(0)", .expect = .{ .number = 1 }, .system = .d1904 },
    .{ .func = "YEAR", .formula = "YEAR(0)", .expect = .{ .number = 1904 }, .system = .d1904 },

    // ── the clock: epoch-independent by construction ──
    .{ .func = "TIME", .formula = "TIME(13,30,0)", .expect = .{ .number = 0.5625 } },
    .{ .func = "TIME", .formula = "TIME(0,0,0)", .expect = .{ .number = 0 } },
    .{ .func = "TIME", .formula = "TIME(25,0,0)", .expect = .{ .number = 0.041666666666666664 }, .note = "25 hours wraps into the next day" },
    .{ .func = "TIME", .formula = "TIME(-1,0,0)", .expect = .{ .err = .num } },
    .{ .func = "TIME", .formula = "TIME(0,0,86400)", .expect = .{ .err = .num }, .note = "the range check is on the ARGUMENT, before the sum" },
    .{ .func = "HOUR", .formula = "HOUR(0.5625)", .expect = .{ .number = 13 } },
    .{ .func = "MINUTE", .formula = "MINUTE(0.5625)", .expect = .{ .number = 30 } },
    .{ .func = "SECOND", .formula = "SECOND(0.5625)", .expect = .{ .number = 0 } },
    .{ .func = "HOUR", .formula = "HOUR(43831.5)", .expect = .{ .number = 12 }, .note = "a serial's whole part is irrelevant to its clock" },
    .{ .func = "SECOND", .formula = "SECOND(0.9999999)", .expect = .{ .number = 0 }, .note = "rounding reaches midnight, and 24:00:00 is not a time" },
    .{ .func = "HOUR", .formula = "HOUR(-0.5)", .expect = .{ .err = .num } },
    // A serial between the two maxima is a date under 1900 and out of
    // range under 1904, and the clock has to agree with the calendar
    // about which: `YEAR` refusing while `HOUR` answered would be two
    // answers about one cell.
    .{ .func = "HOUR", .formula = "HOUR(2957500.5)", .expect = .{ .number = 12 }, .system = .d1900 },
    .{ .func = "HOUR", .formula = "HOUR(2957500.5)", .expect = .{ .err = .num }, .system = .d1904 },
    .{ .func = "YEAR", .formula = "YEAR(2957500)", .expect = .{ .err = .num }, .system = .d1904 },

    // ── WEEKDAY ──
    .{ .func = "WEEKDAY", .formula = "WEEKDAY(43831)", .expect = .{ .number = 4 }, .system = .d1900, .note = "2020-01-01 was a Wednesday" },
    .{ .func = "WEEKDAY", .formula = "WEEKDAY(42369)", .expect = .{ .number = 4 }, .system = .d1904 },
    .{ .func = "WEEKDAY", .formula = "WEEKDAY(43831,2)", .expect = .{ .number = 3 }, .system = .d1900, .note = "Monday-first, one-based" },
    .{ .func = "WEEKDAY", .formula = "WEEKDAY(43831,3)", .expect = .{ .number = 2 }, .system = .d1900, .note = "Monday-first, ZERO-based" },
    .{ .func = "WEEKDAY", .formula = "WEEKDAY(1)", .expect = .{ .number = 1 }, .system = .d1900, .note = "Excel counts serials: serial 1 is a Sunday though 1900-01-01 was a Monday" },
    .{ .func = "WEEKDAY", .formula = "WEEKDAY(43831,0)", .expect = .{ .err = .num } },
    .{ .func = "WEEKDAY", .formula = "WEEKDAY(43831,18)", .expect = .{ .err = .num } },

    // ── EDATE / EOMONTH ──
    .{ .func = "EDATE", .formula = "EDATE(43831,1)", .expect = .{ .number = 43862 }, .system = .d1900 },
    .{ .func = "EDATE", .formula = "EDATE(43831,-1)", .expect = .{ .number = 43800 }, .system = .d1900 },
    .{ .func = "EDATE", .formula = "EDATE(43861,1)", .expect = .{ .number = 43890 }, .system = .d1900, .note = "Jan 31 + 1 month clamps to Feb 29 in a leap year" },
    .{ .func = "EDATE", .formula = "EDATE(43831,0)", .expect = .{ .number = 43831 }, .system = .d1900 },
    .{ .func = "EOMONTH", .formula = "EOMONTH(43831,0)", .expect = .{ .number = 43861 }, .system = .d1900 },
    .{ .func = "EOMONTH", .formula = "EOMONTH(43831,1)", .expect = .{ .number = 43890 }, .system = .d1900, .note = "February 2020 had 29 days" },
    .{ .func = "EOMONTH", .formula = "EOMONTH(43831,-1)", .expect = .{ .number = 43830 }, .system = .d1900 },

    // ── the parses ──
    .{ .func = "DATEVALUE", .formula = "DATEVALUE(\"2020-01-01\")", .expect = .{ .number = serial_2020_1900 }, .system = .d1900 },
    .{ .func = "DATEVALUE", .formula = "DATEVALUE(\"2020-01-01\")", .expect = .{ .number = serial_2020_1904 }, .system = .d1904 },
    .{ .func = "DATEVALUE", .formula = "DATEVALUE(\"1-Jan-2020\")", .expect = .{ .number = serial_2020_1900 }, .system = .d1900 },
    .{ .func = "DATEVALUE", .formula = "DATEVALUE(\"January 15, 2020\")", .expect = .{ .number = 43845 }, .system = .d1900 },
    .{ .func = "DATEVALUE", .formula = "DATEVALUE(\"1/15/2020\")", .expect = .{ .number = 43845 }, .system = .d1900, .note = "15 is not a month, so only one reading exists" },
    .{ .func = "DATEVALUE", .formula = "DATEVALUE(\"hello\")", .expect = .{ .err = .value } },
    .{ .func = "TIMEVALUE", .formula = "TIMEVALUE(\"13:30\")", .expect = .{ .number = 0.5625 } },
    .{ .func = "TIMEVALUE", .formula = "TIMEVALUE(\"1:30 PM\")", .expect = .{ .number = 0.5625 } },
    .{ .func = "TIMEVALUE", .formula = "TIMEVALUE(\"12:00 AM\")", .expect = .{ .number = 0 } },
    .{ .func = "TIMEVALUE", .formula = "TIMEVALUE(\"hello\")", .expect = .{ .err = .value } },

    // ── the volatiles, at a pinned instant ──
    .{ .func = "TODAY", .formula = "TODAY()", .expect = .{ .number = serial_2020_1900 }, .system = .d1900 },
    .{ .func = "TODAY", .formula = "TODAY()", .expect = .{ .number = serial_2020_1904 }, .system = .d1904 },
    .{ .func = "NOW", .formula = "NOW()", .expect = .{ .number = 43831.5 }, .system = .d1900 },
    .{ .func = "NOW", .formula = "NOW()", .expect = .{ .number = 42369.5 }, .system = .d1904 },
};

fn f1dOptions(h: *Harness, system: run_inputs.DateSystem) Options {
    var opts = h.options();
    opts.date_system = system;
    opts.now_utc_ms = now_2020_noon_ms;
    return opts;
}

test "M4g: every F1c-date fixture evaluates to what the spec says, under its epoch" {
    for (f1d_cases) |c| {
        const systems: []const run_inputs.DateSystem =
            if (c.system) |s| &.{s} else &.{ .d1900, .d1904 };
        for (systems) |system| {
            var h: Harness = undefined;
            try h.init(testing.allocator);
            defer h.deinit();

            const v = h.evalOpts(c.formula, f1dOptions(&h, system)) catch |e| {
                std.debug.print("F1c-date `{s}` ({t}) refused: {t}\n", .{ c.formula, system, e });
                return e;
            };
            expectValue(c.expect, v) catch |e| {
                std.debug.print("F1c-date `{s}` ({s}, {t}): wrong value\n", .{ c.formula, c.func, system });
                return e;
            };
        }
    }
}

test "M4g: all fifteen frozen names resolve, and each has a fixture" {
    var it = registry.inventory();
    var batch: usize = 0;
    while (it.next()) |e| {
        if (!std.mem.eql(u8, e.milestone, "M4g")) continue;
        batch += 1;

        if (registry.lookup(e.name) == null) {
            std.debug.print("F1c-date name does not resolve: {s}\n", .{e.name});
            return error.UnregisteredBatchFunction;
        }
        var fixtures: usize = 0;
        for (f1d_cases) |c| {
            if (std.mem.eql(u8, c.func, e.name)) fixtures += 1;
        }
        if (fixtures == 0) {
            std.debug.print("F1c-date name has no fixture: {s}\n", .{e.name});
            return error.UnfixturedBatchFunction;
        }
    }
    try testing.expectEqual(@as(usize, 15), batch);

    for (f1d_cases) |c| {
        var found = false;
        var it2 = registry.inventory();
        while (it2.next()) |e| {
            if (std.mem.eql(u8, e.name, c.func) and std.mem.eql(u8, e.milestone, "M4g")) found = true;
        }
        if (!found) {
            std.debug.print("fixture names a function outside F1c-date: {s}\n", .{c.func});
            return error.FixtureOutsideBatch;
        }
    }
}

test "M4g: the evidence label on every fixture is true of the committed manifests" {
    var oracle_rows: usize = 0;
    var excluded_rows: usize = 0;
    for (f1d_cases) |c| {
        switch (try manifestVerdict(c.formula)) {
            .decided => {
                if (c.evidence != .oracle) return error.UnderstatedEvidence;
                oracle_rows += 1;
            },
            .excluded => {
                if (c.evidence != .spec_pinned) return error.ExcludedCellClaimedAsEvidence;
                excluded_rows += 1;
            },
            .silent => {
                if (c.evidence != .spec_pinned) return error.UnbackedOracleClaim;
            },
        }
    }
    try testing.expectEqual(@as(usize, 0), oracle_rows);
    try testing.expectEqual(@as(usize, 0), excluded_rows);
}

test "M4g: the epoch flag is exactly the names a date system can reach" {
    // Registry data in both directions, as M4f did for the
    // compatibility version. `TIME`, `HOUR`, `MINUTE` and `SECOND` are
    // the interesting absences: a fraction of a day is the same
    // fraction under either epoch, and flagging them would claim a
    // dependence that does not exist.
    const epoch = [_][]const u8{
        "DATE",  "YEAR",    "MONTH",     "DAY",   "WEEKDAY",
        "EDATE", "EOMONTH", "DATEVALUE", "TODAY", "NOW",
    };
    var it = registry.inventory();
    while (it.next()) |e| {
        if (!std.mem.eql(u8, e.milestone, "M4g")) continue;
        var listed = false;
        for (epoch) |n| {
            if (std.mem.eql(u8, n, e.name)) listed = true;
        }
        if (listed != registry.lookup(e.name).?.epoch_sensitive) {
            std.debug.print("{s}: epoch_sensitive={} but the list says {}\n", .{
                e.name,
                registry.lookup(e.name).?.epoch_sensitive,
                listed,
            });
            return error.EpochFlagMismatch;
        }
    }

    // …and the two volatiles are the only volatile rows in the batch.
    var it2 = registry.inventory();
    while (it2.next()) |e| {
        if (!std.mem.eql(u8, e.milestone, "M4g")) continue;
        const want: registry.Volatility = if (std.mem.eql(u8, e.name, "TODAY") or
            std.mem.eql(u8, e.name, "NOW")) .volatile_fn else .stable;
        try testing.expectEqual(want, registry.lookup(e.name).?.volatility);
    }
}

test "M4g: the same serial, two epochs, two dates" {
    // The row's headline, in one cell: serial 43831 is 2020-01-01 in a
    // 1900 workbook and 2024-01-02 in a 1904 one. Nothing about the
    // value changed — only what the file says it means.
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();

    try testing.expectEqual(
        @as(f64, 2020),
        (try h.evalOpts("YEAR(43831)", f1dOptions(&h, .d1900))).scalar.number,
    );
    try testing.expectEqual(
        @as(f64, 2024),
        (try h.evalOpts("YEAR(43831)", f1dOptions(&h, .d1904))).scalar.number,
    );
    // The gap between the systems is 1462 days — four years and the
    // 1900 phantom — and it holds for every date the 1904 system can
    // express at all. Below serial 1462 there is no 1904 counterpart:
    // those dates precede its epoch and simply are not dates there,
    // which is the same fact seen from the other side.
    for ([_]f64{ 2000, 43831, 90000 }) |serial| {
        var buf: [64]u8 = undefined;
        const a = try std.fmt.bufPrint(&buf, "YEAR({d})*10000+MONTH({d})*100+DAY({d})", .{ serial, serial, serial });
        const under_1900 = (try h.evalOpts(a, f1dOptions(&h, .d1900))).scalar.number;
        var buf2: [64]u8 = undefined;
        const b = try std.fmt.bufPrint(&buf2, "YEAR({d})*10000+MONTH({d})*100+DAY({d})", .{ serial - 1462, serial - 1462, serial - 1462 });
        const under_1904 = (try h.evalOpts(b, f1dOptions(&h, .d1904))).scalar.number;
        try testing.expectEqual(under_1900, under_1904);
    }
}

test "M4g: TODAY and NOW come from RunInputs, not from a clock" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();

    // Determinism first: the same inputs give the same answer, which is
    // what makes a volatile function reproducible at all.
    var opts = f1dOptions(&h, .d1900);
    const first = (try h.evalOpts("NOW()", opts)).scalar.number;
    const second = (try h.evalOpts("NOW()", opts)).scalar.number;
    try testing.expectEqual(first, second);
    try testing.expectEqual(@as(f64, 43831.5), first);
    // TODAY is NOW floored, asserted as a relationship rather than as
    // two independent constants.
    try testing.expectEqual(
        @floor(first),
        (try h.evalOpts("TODAY()", opts)).scalar.number,
    );

    // The offset is a civil offset and moves the clock, not the epoch.
    opts.utc_offset_min = 90;
    try testing.expectEqual(
        @as(f64, 43831.5 + 90.0 / 1440.0),
        (try h.evalOpts("NOW()", opts)).scalar.number,
    );

    // …and it can move the DAY, which is the case a UTC-only engine
    // gets wrong. 23:30 UTC plus 60 minutes is tomorrow.
    var late = f1dOptions(&h, .d1900);
    late.now_utc_ms = now_2020_noon_ms + 11 * 3_600_000 + 30 * 60_000; // 23:30Z
    try testing.expectEqual(
        @as(f64, 43831),
        (try h.evalOpts("TODAY()", late)).scalar.number,
    );
    late.utc_offset_min = 60;
    try testing.expectEqual(
        @as(f64, 43832),
        (try h.evalOpts("TODAY()", late)).scalar.number,
    );
}

test "M4g: DATEVALUE refuses a locale-ordered date and errors on a non-date" {
    // §5.4b's split, end to end. The two outcomes are different KINDS
    // of answer — one is a refusal the caller must handle, the other is
    // a value a formula can catch with IFERROR — and collapsing them
    // would either hide an ambiguity or invent an error.
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    const opts = f1dOptions(&h, .d1900);

    try testing.expectError(
        error.LocaleSensitiveInput,
        h.evalOpts("DATEVALUE(\"1/2/2020\")", opts),
    );
    try testing.expectError(
        error.LocaleSensitiveInput,
        h.evalOpts("DATEVALUE(\"12-11-2020\")", opts),
    );
    // Unambiguous by value, in either field order.
    try testing.expectEqual(
        @as(f64, 43845),
        (try h.evalOpts("DATEVALUE(\"15/1/2020\")", opts)).scalar.number,
    );
    // Not a date is `#VALUE!`, and IFERROR proves it is a value rather
    // than a refusal.
    try testing.expectEqualStrings(
        "no",
        (try h.evalOpts("IFERROR(DATEVALUE(\"hello\"),\"no\")", opts)).scalar.text,
    );
    // A refusal is NOT catchable, which is the other half of the split.
    try testing.expectError(
        error.LocaleSensitiveInput,
        h.evalOpts("IFERROR(DATEVALUE(\"1/2/2020\"),\"no\")", opts),
    );
}

test "M4g: the batch agrees across both rule tables" {
    // Dates are integer arithmetic over a serial, so the two fidelity
    // modes have nothing to disagree about — asserted rather than
    // assumed, because `TIME` divides and `NOW` carries a fraction.
    for (f1d_cases) |c| {
        const system = c.system orelse .d1900;
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();

        var opts_e = f1dOptions(&h, system);
        opts_e.fidelity = .excel;
        var opts_i = f1dOptions(&h, system);
        opts_i.fidelity = .ieee;

        const a = h.evalOpts(c.formula, opts_e) catch continue;
        const b = try h.evalOpts(c.formula, opts_i);
        if (a == .scalar and b == .scalar and !a.scalar.eql(b.scalar)) {
            std.debug.print("`{s}` differs between rule tables\n", .{c.formula});
            return error.UnexpectedFidelityDivergence;
        }
    }
}

test "M4g: every name against every argument shape, in both epochs" {
    const shapes = [_][]const u8{ "43831", "0", "\"abc\"", "TRUE", "A5", "A7", "{1,2}" };
    var checked: usize = 0;
    for ([_]run_inputs.DateSystem{ .d1900, .d1904 }) |system| {
        var it = registry.inventory();
        while (it.next()) |e| {
            if (!std.mem.eql(u8, e.milestone, "M4g")) continue;
            const f = registry.lookup(e.name).?;
            for (shapes) |s| {
                var buf: [256]u8 = undefined;
                var w: std.Io.Writer = .fixed(&buf);
                try w.print("{s}({s}", .{ e.name, s });
                var k: usize = 1;
                while (k < f.arity.min) : (k += 1) try w.print(",{s}", .{s});
                try w.writeAll(")");
                const src = w.buffered();

                var h: Harness = undefined;
                try h.init(testing.allocator);
                defer h.deinit();
                try putF1cCells(&h);
                const first = h.evalOpts(src, f1dOptions(&h, system));

                var h2: Harness = undefined;
                try h2.init(testing.allocator);
                defer h2.deinit();
                try putF1cCells(&h2);
                const second = h2.evalOpts(src, f1dOptions(&h2, system));

                if (first) |a| {
                    const b = second catch return error.NondeterministicEvaluation;
                    if (a == .scalar and b == .scalar and !a.scalar.eql(b.scalar)) {
                        std.debug.print("`{s}` gave two answers\n", .{src});
                        return error.NondeterministicEvaluation;
                    }
                } else |err| {
                    if (second) |_| {
                        return error.NondeterministicEvaluation;
                    } else |err2| {
                        if (err != err2) return error.NondeterministicEvaluation;
                    }
                }
                checked += 1;
            }
        }
    }
    // Fifteen names × seven shapes × two epochs. `TODAY` and `NOW` take
    // no arguments and refuse every shape on arity, which is itself the
    // assertion that a zero-argument row cannot quietly grow one.
    try testing.expectEqual(@as(usize, 15 * 7 * 2), checked);
}

// ─── M7a: the F2-DA batch (§7, seven names; §5.8a) ───────────────

/// The battery's grid. Column A is SORT's numeric subject with a
/// deliberate hole (A5 — §5.3b pins blanks sorting FIRST); column B is
/// the collation ladder (`A`/`a` fold-equal, `ss`/`ß` fold-equal);
/// D holds an error ABOVE a number so "errors pinned last" is proved
/// against source order; E is FILTER's include with an error in it.
fn putF2Cells(h: *Harness) !void {
    try h.put("A1", num(3));
    try h.put("A2", num(1));
    try h.put("A3", num(2));
    try h.put("A4", num(2));
    try h.put("B1", .{ .text = "b" });
    try h.put("B2", .{ .text = "A" });
    try h.put("B3", .{ .text = "a" });
    try h.put("B4", .{ .text = "ss" });
    try h.put("B5", .{ .text = "ß" });
    try h.put("C1", .{ .boolean = true });
    try h.put("C2", value.ScalarValue.errorOf(.ref));
    try h.put("D1", value.ScalarValue.errorOf(.ref));
    try h.put("D2", num(7));
    try h.put("E1", num(1));
    try h.put("E2", value.ScalarValue.errorOf(.div0));
    try h.put("E3", num(0));
    try h.put("E4", num(1));
}

const F2Case = struct {
    func: []const u8,
    formula: []const u8,
    expect: Expect,
    evidence: Evidence = .spec_pinned,
    note: []const u8 = "",
};

const f2_cases = [_]F2Case{
    // ── SEQUENCE: the closed form, both dimensions, CHOOSE's trunc ──
    .{ .func = "SEQUENCE", .formula = "SEQUENCE(2,3)", .expect = .{ .array = .{ .rows = 2, .cols = 3, .cells = &.{ .{ .number = 1 }, .{ .number = 2 }, .{ .number = 3 }, .{ .number = 4 }, .{ .number = 5 }, .{ .number = 6 } } } } },
    .{ .func = "SEQUENCE", .formula = "SEQUENCE(3)", .expect = .{ .array = .{ .rows = 3, .cols = 1, .cells = &.{ .{ .number = 1 }, .{ .number = 2 }, .{ .number = 3 } } } } },
    .{ .func = "SEQUENCE", .formula = "SEQUENCE(2,2,10,-1)", .expect = .{ .array = .{ .rows = 2, .cols = 2, .cells = &.{ .{ .number = 10 }, .{ .number = 9 }, .{ .number = 8 }, .{ .number = 7 } } } } },
    .{ .func = "SEQUENCE", .formula = "SEQUENCE(2.9,1.9)", .expect = .{ .array = .{ .rows = 2, .cols = 1, .cells = &.{ .{ .number = 1 }, .{ .number = 2 } } } }, .note = "dimensions truncate toward zero" },
    .{ .func = "SEQUENCE", .formula = "SEQUENCE(1,2,0.5,0.25)", .expect = .{ .array = .{ .rows = 1, .cols = 2, .cells = &.{ .{ .number = 0.5 }, .{ .number = 0.75 } } } } },
    .{ .func = "SEQUENCE", .formula = "SEQUENCE(0)", .expect = .{ .err = .calc }, .note = "a zero extent is the empty rectangle (§5.3a)" },
    .{ .func = "SEQUENCE", .formula = "SEQUENCE(-1)", .expect = .{ .err = .value } },

    // ── RANDARRAY under the harness's constant 0.5 source: the VALUE
    //    rows; the draw schedule is the KATs' below ──
    .{ .func = "RANDARRAY", .formula = "RANDARRAY()", .expect = .{ .array = .{ .rows = 1, .cols = 1, .cells = &.{.{ .number = 0.5 }} } } },
    .{ .func = "RANDARRAY", .formula = "RANDARRAY(2,2)", .expect = .{ .array = .{ .rows = 2, .cols = 2, .cells = &.{ .{ .number = 0.5 }, .{ .number = 0.5 }, .{ .number = 0.5 }, .{ .number = 0.5 } } } } },
    .{ .func = "RANDARRAY", .formula = "RANDARRAY(1,1,10,20)", .expect = .{ .array = .{ .rows = 1, .cols = 1, .cells = &.{.{ .number = 15 }} } } },
    .{ .func = "RANDARRAY", .formula = "RANDARRAY(1,1,1,10,TRUE())", .expect = .{ .array = .{ .rows = 1, .cols = 1, .cells = &.{.{ .number = 6 }} } }, .note = "RANDBETWEEN's scaling: floor(0.5·10)+1" },
    .{ .func = "RANDARRAY", .formula = "RANDARRAY(1,1,2,1)", .expect = .{ .err = .value }, .note = "min above max" },
    .{ .func = "RANDARRAY", .formula = "RANDARRAY(1,1,1.2,1.8,TRUE())", .expect = .{ .err = .num }, .note = "empty only after the bounds moved inward — RANDBETWEEN's row" },
    .{ .func = "RANDARRAY", .formula = "RANDARRAY(0)", .expect = .{ .err = .calc } },
    .{ .func = "RANDARRAY", .formula = "RANDARRAY(-1)", .expect = .{ .err = .value } },

    // ── TRANSPOSE ──
    .{ .func = "TRANSPOSE", .formula = "TRANSPOSE({1,2,3})", .expect = .{ .array = .{ .rows = 3, .cols = 1, .cells = &.{ .{ .number = 1 }, .{ .number = 2 }, .{ .number = 3 } } } } },
    .{ .func = "TRANSPOSE", .formula = "TRANSPOSE({1,2;3,4})", .expect = .{ .array = .{ .rows = 2, .cols = 2, .cells = &.{ .{ .number = 1 }, .{ .number = 3 }, .{ .number = 2 }, .{ .number = 4 } } } } },
    .{ .func = "TRANSPOSE", .formula = "TRANSPOSE(5)", .expect = .{ .array = .{ .rows = 1, .cols = 1, .cells = &.{.{ .number = 5 }} } } },
    .{ .func = "TRANSPOSE", .formula = "TRANSPOSE(A1:A3)", .expect = .{ .array = .{ .rows = 1, .cols = 3, .cells = &.{ .{ .number = 3 }, .{ .number = 1 }, .{ .number = 2 } } } } },
    .{ .func = "TRANSPOSE", .formula = "TRANSPOSE((A1:A2,B1:B2))", .expect = .{ .err = .value }, .note = "a multi-area union is not one rectangle" },

    // ── FILTER ──
    .{ .func = "FILTER", .formula = "FILTER(A1:A4,{1;0;1;0})", .expect = .{ .array = .{ .rows = 2, .cols = 1, .cells = &.{ .{ .number = 3 }, .{ .number = 2 } } } } },
    .{ .func = "FILTER", .formula = "FILTER(A1:A4,{0;0;0;0})", .expect = .{ .err = .calc }, .note = "nothing matched and no third argument: the empty matrix meets the call boundary" },
    .{ .func = "FILTER", .formula = "FILTER(A1:A4,{0;0;0;0},\"none\")", .expect = .{ .text = "none" } },
    .{ .func = "FILTER", .formula = "FILTER({1;2},{0;0},{9,8})", .expect = .{ .array = .{ .rows = 1, .cols = 2, .cells = &.{ .{ .number = 9 }, .{ .number = 8 } } } }, .note = "if_empty may itself be an array" },
    .{ .func = "FILTER", .formula = "FILTER(A1:A4,{1;0;1})", .expect = .{ .err = .value }, .note = "include length mismatch" },
    .{ .func = "FILTER", .formula = "FILTER({1,2,3},{0,1,1})", .expect = .{ .array = .{ .rows = 1, .cols = 2, .cells = &.{ .{ .number = 2 }, .{ .number = 3 } } } }, .note = "a row include names columns" },
    .{ .func = "FILTER", .formula = "FILTER(A1:A4,{1;\"x\";1;0})", .expect = .{ .err = .value }, .note = "text is not a condition (§5.3b's logical column)" },
    .{ .func = "FILTER", .formula = "FILTER(A1:A4,E1:E4)", .expect = .{ .err = .div0 }, .note = "an error in the include is the whole answer" },
    .{ .func = "FILTER", .formula = "FILTER({1,2;3,4},{1;0})", .expect = .{ .array = .{ .rows = 1, .cols = 2, .cells = &.{ .{ .number = 1 }, .{ .number = 2 } } } } },

    // ── SORT ──
    .{ .func = "SORT", .formula = "SORT({3;1;2})", .expect = .{ .array = .{ .rows = 3, .cols = 1, .cells = &.{ .{ .number = 1 }, .{ .number = 2 }, .{ .number = 3 } } } } },
    .{ .func = "SORT", .formula = "SORT({3;1;2},1,-1)", .expect = .{ .array = .{ .rows = 3, .cols = 1, .cells = &.{ .{ .number = 3 }, .{ .number = 2 }, .{ .number = 1 } } } } },
    .{ .func = "SORT", .formula = "SORT({1,3;2,1},2)", .expect = .{ .array = .{ .rows = 2, .cols = 2, .cells = &.{ .{ .number = 2 }, .{ .number = 1 }, .{ .number = 1 }, .{ .number = 3 } } } }, .note = "sort rows by the second column" },
    .{ .func = "SORT", .formula = "SORT({3,1,2},1,1,TRUE())", .expect = .{ .array = .{ .rows = 1, .cols = 3, .cells = &.{ .{ .number = 1 }, .{ .number = 2 }, .{ .number = 3 } } } }, .note = "by_col sorts columns" },
    .{ .func = "SORT", .formula = "SORT(A1:A5)", .expect = .{ .array = .{ .rows = 5, .cols = 1, .cells = &.{ .blank, .{ .number = 1 }, .{ .number = 2 }, .{ .number = 2 }, .{ .number = 3 } } } }, .note = "§5.3b: blanks sort first" },
    .{ .func = "SORT", .formula = "SORT(B1:B5)", .expect = .{ .array = .{ .rows = 5, .cols = 1, .cells = &.{ .{ .text = "A" }, .{ .text = "a" }, .{ .text = "b" }, .{ .text = "ss" }, .{ .text = "ß" } } } }, .note = "fold-equal elements are EQUAL; source order is the only tie-break" },
    .{ .func = "SORT", .formula = "SORT(D1:D2)", .expect = .{ .array = .{ .rows = 2, .cols = 1, .cells = &.{ .{ .number = 7 }, .{ .err = .ref } } } }, .note = "errors pinned last, from above the number in source order" },
    .{ .func = "SORT", .formula = "SORT({2;1},2)", .expect = .{ .err = .value }, .note = "sort_index out of bounds" },
    .{ .func = "SORT", .formula = "SORT({2;1},1,0)", .expect = .{ .err = .value }, .note = "an order is 1 or -1" },
    .{ .func = "SORT", .formula = "SORT({TRUE;\"x\";2})", .expect = .{ .array = .{ .rows = 3, .cols = 1, .cells = &.{ .{ .number = 2 }, .{ .text = "x" }, .{ .boolean = true } } } }, .note = "number < text < logical" },
    .{ .func = "SORT", .formula = "SORT({1,2;1,1;0,5},{1,2})", .expect = .{ .array = .{ .rows = 3, .cols = 2, .cells = &.{ .{ .number = 0 }, .{ .number = 5 }, .{ .number = 1 }, .{ .number = 1 }, .{ .number = 1 }, .{ .number = 2 } } } }, .note = "two levels: column 1, then column 2" },
    .{ .func = "SORT", .formula = "SORT({1,2;2,1;1,1},{1,2},-1)", .expect = .{ .array = .{ .rows = 3, .cols = 2, .cells = &.{ .{ .number = 2 }, .{ .number = 1 }, .{ .number = 1 }, .{ .number = 2 }, .{ .number = 1 }, .{ .number = 1 } } } }, .note = "a single order broadcasts over every level" },

    // ── SORTBY ──
    .{ .func = "SORTBY", .formula = "SORTBY({\"a\";\"b\";\"c\"},{3;1;2})", .expect = .{ .array = .{ .rows = 3, .cols = 1, .cells = &.{ .{ .text = "b" }, .{ .text = "c" }, .{ .text = "a" } } } } },
    .{ .func = "SORTBY", .formula = "SORTBY({\"a\";\"b\";\"c\"},{3;1;2},-1)", .expect = .{ .array = .{ .rows = 3, .cols = 1, .cells = &.{ .{ .text = "a" }, .{ .text = "c" }, .{ .text = "b" } } } } },
    .{ .func = "SORTBY", .formula = "SORTBY({1;2;3},{1;1;1},1,{3;2;1},1)", .expect = .{ .array = .{ .rows = 3, .cols = 1, .cells = &.{ .{ .number = 3 }, .{ .number = 2 }, .{ .number = 1 } } } }, .note = "primary all-ties; the second pair decides" },
    .{ .func = "SORTBY", .formula = "SORTBY({5,4},{2,1})", .expect = .{ .array = .{ .rows = 1, .cols = 2, .cells = &.{ .{ .number = 4 }, .{ .number = 5 } } } }, .note = "a row by-vector sorts columns" },
    .{ .func = "SORTBY", .formula = "SORTBY({1;2},{1,2})", .expect = .{ .err = .value }, .note = "the by-vector must run the array's way" },
    .{ .func = "SORTBY", .formula = "SORTBY({1;2},{1;2},5)", .expect = .{ .err = .value } },

    // ── UNIQUE ──
    .{ .func = "UNIQUE", .formula = "UNIQUE({1;2;1;3})", .expect = .{ .array = .{ .rows = 3, .cols = 1, .cells = &.{ .{ .number = 1 }, .{ .number = 2 }, .{ .number = 3 } } } } },
    .{ .func = "UNIQUE", .formula = "UNIQUE({1;2;1;3},FALSE(),TRUE())", .expect = .{ .array = .{ .rows = 2, .cols = 1, .cells = &.{ .{ .number = 2 }, .{ .number = 3 } } } }, .note = "exactly_once" },
    .{ .func = "UNIQUE", .formula = "UNIQUE({1,2;1,2;3,4})", .expect = .{ .array = .{ .rows = 2, .cols = 2, .cells = &.{ .{ .number = 1 }, .{ .number = 2 }, .{ .number = 3 }, .{ .number = 4 } } } }, .note = "a row is the unit of equality" },
    .{ .func = "UNIQUE", .formula = "UNIQUE({1,1,2},TRUE())", .expect = .{ .array = .{ .rows = 1, .cols = 2, .cells = &.{ .{ .number = 1 }, .{ .number = 2 } } } }, .note = "by_col dedups columns" },
    .{ .func = "UNIQUE", .formula = "UNIQUE(B2:B3)", .expect = .{ .array = .{ .rows = 1, .cols = 1, .cells = &.{.{ .text = "A" }} } }, .note = "`A`/`a` fold-equal: keep-first" },
    .{ .func = "UNIQUE", .formula = "UNIQUE(B4:B5)", .expect = .{ .array = .{ .rows = 1, .cols = 1, .cells = &.{.{ .text = "ss" }} } }, .note = "`ss`/`ß` fold-equal: keep-first" },
    .{ .func = "UNIQUE", .formula = "UNIQUE({1;1},FALSE(),TRUE())", .expect = .{ .err = .calc }, .note = "no singleton rows: the empty rectangle" },
    .{ .func = "UNIQUE", .formula = "UNIQUE({1;TRUE})", .expect = .{ .array = .{ .rows = 2, .cols = 1, .cells = &.{ .{ .number = 1 }, .{ .boolean = true } } } }, .note = "cross-type pairs are never equal" },
};

test "M7a: every F2-DA fixture evaluates to what the spec says" {
    for (f2_cases) |c| {
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        try putF2Cells(&h);

        const v = h.eval(c.formula) catch |e| {
            std.debug.print("F2-DA case `{s}` refused: {t}\n", .{ c.formula, e });
            return e;
        };
        expectValue(c.expect, v) catch |e| {
            std.debug.print("F2-DA case `{s}` ({s})\n", .{ c.formula, c.note });
            return e;
        };
    }
}

test "M7a: all seven frozen names resolve, and each has a fixture" {
    var it = registry.inventory();
    var batch: usize = 0;
    while (it.next()) |e| {
        if (!std.mem.eql(u8, e.milestone, "M7a")) continue;
        batch += 1;

        if (registry.lookup(e.name) == null) {
            std.debug.print("F2-DA name does not resolve: {s}\n", .{e.name});
            return error.UnregisteredBatchFunction;
        }
        var fixtures: usize = 0;
        for (f2_cases) |c| {
            if (std.mem.eql(u8, c.func, e.name)) fixtures += 1;
        }
        if (fixtures == 0) {
            std.debug.print("F2-DA name has no fixture: {s}\n", .{e.name});
            return error.UnfixturedBatchFunction;
        }
    }
    try testing.expectEqual(@as(usize, 7), batch);

    for (f2_cases) |c| {
        var found = false;
        var it2 = registry.inventory();
        while (it2.next()) |e| {
            if (std.mem.eql(u8, e.name, c.func) and std.mem.eql(u8, e.milestone, "M7a")) found = true;
        }
        if (!found) {
            std.debug.print("fixture names a function outside F2-DA: {s}\n", .{c.func});
            return error.FixtureOutsideBatch;
        }
    }
}

test "M7a: the evidence label on every fixture is true of the committed manifests" {
    var oracle_rows: usize = 0;
    var excluded_rows: usize = 0;
    for (f2_cases) |c| {
        switch (try manifestVerdict(c.formula)) {
            .decided => {
                if (c.evidence != .oracle) return error.UnderstatedEvidence;
                oracle_rows += 1;
            },
            .excluded => {
                if (c.evidence != .spec_pinned) return error.ExcludedCellClaimedAsEvidence;
                excluded_rows += 1;
            },
            .silent => {
                if (c.evidence != .spec_pinned) return error.UnbackedOracleClaim;
            },
        }
    }
    // The committed manifests predate every F2-DA name; the parked
    // Excel leg is what would move these.
    try testing.expectEqual(@as(usize, 0), oracle_rows);
    try testing.expectEqual(@as(usize, 0), excluded_rows);
}

test "M7a: the mixed-signature lift — scalar slots lift, whole slots hold" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try h.put("A1", num(1));
    try h.put("B1", num(10));
    try h.put("A2", num(2));
    try h.put("B2", num(20));

    const Case = struct { formula: []const u8, expect: Expect };
    const cases = [_]Case{
        // `INDEX(A1:B2,{1;2},1)` is two INDEXes into ONE table: the
        // scalar row slot lifts, the `.aggregate` table holds. Before
        // M7a this was the `error.NotYetImplemented` at the
        // `.spill_or_iterate` arm.
        .{ .formula = "INDEX(A1:B2,{1;2},1)", .expect = .{ .array = .{ .rows = 2, .cols = 1, .cells = &.{ .{ .number = 1 }, .{ .number = 2 } } } } },
        // Both lifted slots at once, broadcast against each other.
        .{ .formula = "INDEX(A1:B2,{1;2},{1;2})", .expect = .{ .array = .{ .rows = 2, .cols = 1, .cells = &.{ .{ .number = 1 }, .{ .number = 20 } } } } },
        // §5.3b's nested-array rule: a per-element result that is
        // itself an array reduces to its top-left — `SEQUENCE({1,2})`
        // is `{1,1}`.
        .{ .formula = "SEQUENCE({1,2})", .expect = .{ .array = .{ .rows = 1, .cols = 2, .cells = &.{ .{ .number = 1 }, .{ .number = 1 } } } } },
    };
    for (cases) |c| {
        const v = h.eval(c.formula) catch |e| {
            std.debug.print("mixed-lift case `{s}` refused: {t}\n", .{ c.formula, e });
            return e;
        };
        expectValue(c.expect, v) catch |e| {
            std.debug.print("mixed-lift case `{s}`\n", .{c.formula});
            return e;
        };
    }

    // Under `legacy` the same text reduces the array in the scalar slot
    // to its top-left instead of lifting (§5.3b `array where scalar
    // expected`).
    var legacy = h.options();
    legacy.dialect = .legacy;
    try expectValue(.{ .number = 1 }, try h.evalOpts("INDEX(A1:B2,{1;2},1)", legacy));
}

test "M7a: a DA native is a DA native under either dialect" {
    // Dialect changes how an ARRAY MEETS A SCALAR SLOT, not what a
    // matrix-producing native returns: `SEQUENCE(3)` is 3×1 under
    // `legacy` too (a legacy CSE consumes it by declared range, §5.6h).
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    var legacy = h.options();
    legacy.dialect = .legacy;
    try expectValue(
        .{ .array = .{ .rows = 3, .cols = 1, .cells = &.{ .{ .number = 1 }, .{ .number = 2 }, .{ .number = 3 } } } },
        try h.evalOpts("SEQUENCE(3)", legacy),
    );
    // `@SEQUENCE(3)` is §5.3b's own citation for `@` over an array:
    // top-left, under both dialects.
    try expectValue(.{ .number = 1 }, try h.eval("@SEQUENCE(3)"));
    try expectValue(.{ .number = 1 }, try h.evalOpts("@SEQUENCE(3)", legacy));
}

test "M7a: a rectangle past §9's matrix cap is a LIMIT, not a value" {
    // §5.8a's last row: > `max_matrix_cells` → `FormulaLimitExceeded`.
    // The cap is enforced by construction at `Matrix.init`, so the
    // refusal fires before a single cell is allocated.
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try testing.expectError(error.LimitExceeded, h.eval("SEQUENCE(1,4000001)"));
    // A dimension the cap can never fit refuses identically — the u32
    // cast is guarded by the same rule, not by luck.
    try testing.expectError(error.LimitExceeded, h.eval("SEQUENCE(99999999999)"));
    try testing.expectError(error.LimitExceeded, h.eval("RANDARRAY(4000001)"));
    // The boundary itself is a value: exactly `max_matrix_cells` cells.
    const v = try h.eval("SEQUENCE(2000,2000)");
    try testing.expect(v == .array);
    try testing.expectEqual(@as(f64, 1), v.array.at(0, 0).number);
    try testing.expectEqual(@as(f64, 4_000_000), v.array.at(1999, 1999).number);
}

// ─── M7a draw KATs (§5.6d: RANDARRAY element ordinals) ───────────

test "M7a: draw counts — one per element, zero before a refused range" {
    const Case = struct { formula: []const u8, draws: u64 };
    const cases = [_]Case{
        .{ .formula = "SEQUENCE(3)", .draws = 0 },
        .{ .formula = "RANDARRAY()", .draws = 1 },
        .{ .formula = "RANDARRAY(2,3)", .draws = 6 },
        .{ .formula = "RANDARRAY(2,2,1,4,TRUE())", .draws = 4 },
        // The bound check precedes the loop: a refused range consumed
        // NOTHING, so the schedule cannot have been perturbed by an
        // answer that never existed.
        .{ .formula = "RANDARRAY(1,1,2,1)", .draws = 0 },
        // Laziness still governs: a dead branch's rectangle draws zero.
        .{ .formula = "IF(TRUE(),1,RANDARRAY(2,2))", .draws = 0 },
        // An aggregate slot is eager: SORT's subject draws before SORT
        // reads it.
        .{ .formula = "SORT(RANDARRAY(2,2))", .draws = 4 },
    };
    for (cases) |c| {
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();

        _ = h.eval(c.formula) catch |e| {
            std.debug.print("draw case `{s}` refused: {t}\n", .{ c.formula, e });
            return e;
        };
        if (h.draws.count != c.draws) {
            std.debug.print("`{s}`: expected {d} draws, saw {d}\n", .{ c.formula, c.draws, h.draws.count });
            return error.WrongDrawCount;
        }
    }
}

test "M7a KAT: RANDARRAY draws by element ordinal, reproducibly from the seed" {
    const inputs: run_inputs.RunInputs = .{ .now_utc_ms = 0, .rng_seed = 0xDA7A_A44A, .limits = .{} };
    try inputs.validate();

    var first: [6]f64 = undefined;
    {
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        var generator = rng.Rng.fromRunInputs(inputs);
        var source = generator.drawSource();
        var opts = h.options();
        opts.draws = &source;

        const v = try h.evalOpts("RANDARRAY(2,3)", opts);
        try testing.expect(v == .array);
        try testing.expectEqual(@as(u32, 2), v.array.rows);
        try testing.expectEqual(@as(u32, 3), v.array.cols);
        // One draw per element, counted — the instrument a right answer
        // arrived at wrongly cannot satisfy.
        try testing.expectEqual(@as(u64, 6), source.count);
        for (&first, 0..) |*slot, i| slot.* = v.array.cells[i].number;
        // Distinct ordinals, distinct draws: under a per-element key no
        // two of this seed's six elements collide.
        for (first, 0..) |a, i| {
            for (first[i + 1 ..]) |b| try testing.expect(a != b);
        }
    }
    {
        // Same seed ⇒ the SAME array, bit for bit.
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        var generator = rng.Rng.fromRunInputs(inputs);
        var source = generator.drawSource();
        var opts = h.options();
        opts.draws = &source;

        const v = try h.evalOpts("RANDARRAY(2,3)", opts);
        for (first, 0..) |a, i| {
            try testing.expectEqual(
                @as(u64, @bitCast(a)),
                @as(u64, @bitCast(v.array.cells[i].number)),
            );
        }
    }
    {
        // A different seed names a different stream.
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        var generator = rng.Rng.init(inputs.rng_seed + 1);
        var source = generator.drawSource();
        var opts = h.options();
        opts.draws = &source;

        const v = try h.evalOpts("RANDARRAY(2,3)", opts);
        var differs = false;
        for (first, 0..) |a, i| {
            if (@as(u64, @bitCast(a)) != @as(u64, @bitCast(v.array.cells[i].number))) differs = true;
        }
        try testing.expect(differs);
    }
}

test "M7a KAT: a decided element key generates nothing on re-evaluation" {
    // §5.6d's memo at the element level: the second evaluation of the
    // same text meets six DECIDED keys and reuses every one — the
    // rebuild-reuse property RANDARRAY's shape passes will lean on.
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    var sched: draw_schedule.Schedule = .{};
    defer sched.deinit(testing.allocator);
    var generator = rng.Rng.init(0x5EED_DA);
    var source = generator.drawSource();
    source.schedule = &sched;
    source.gpa = testing.allocator;
    var opts = h.options();
    opts.draws = &source;

    const a = try h.evalOpts("RANDARRAY(2,3)", opts);
    try testing.expectEqual(@as(u64, 6), sched.generated);
    try testing.expectEqual(@as(u64, 0), sched.reused);

    const b = try h.evalOpts("RANDARRAY(2,3)", opts);
    try testing.expectEqual(@as(u64, 6), sched.generated);
    try testing.expectEqual(@as(u64, 6), sched.reused);
    for (a.array.cells, b.array.cells) |x, y| {
        try testing.expectEqual(@as(u64, @bitCast(x.number)), @as(u64, @bitCast(y.number)));
    }
}

// ─── M7b2: the F2-criteria batch (§7, six names; §5.6a) ──────────

/// The battery's grid: three aligned criteria columns — A numeric with
/// a deliberate hole at A5 (logical blanks travel as RUNS through the
/// aligned cursor, and a grid without one would never prove it), B
/// text, C numeric — and the aggregation column D beside them. E mixes
/// types for the restriction rows, F holds §5.3c's two errors, G a
/// logical.
fn putF2cCells(h: *Harness) !void {
    try h.put("A1", num(1));
    try h.put("A2", num(2));
    try h.put("A3", num(3));
    try h.put("A4", num(4));
    try h.put("B1", .{ .text = "x" });
    try h.put("B2", .{ .text = "y" });
    try h.put("B3", .{ .text = "x" });
    try h.put("B4", .{ .text = "y" });
    try h.put("B5", .{ .text = "x" });
    try h.put("C1", num(10));
    try h.put("C2", num(20));
    try h.put("C3", num(30));
    try h.put("C4", num(40));
    try h.put("C5", num(50));
    try h.put("D1", num(100));
    try h.put("D2", num(200));
    try h.put("D3", num(300));
    try h.put("D4", num(400));
    try h.put("D5", num(500));
    try h.put("E1", .{ .text = "x" });
    try h.put("E2", num(5));
    try h.put("F1", value.ScalarValue.errorOf(.div0));
    try h.put("F2", value.ScalarValue.errorOf(.na));
    try h.put("G1", .{ .boolean = true });
}

const f2c_cases = [_]F2Case{
    // ── SUMIFS: one criterion, several, and §5.6a's two refashionings
    //    of `#VALUE!` — mismatched dimensions and a multi-area union ──
    .{ .func = "SUMIFS", .formula = "SUMIFS(D1:D5,A1:A5,\">2\")", .expect = .{ .number = 700 } },
    .{ .func = "SUMIFS", .formula = "SUMIFS(D1:D5,A1:A5,\">1\",B1:B5,\"x\")", .expect = .{ .number = 300 }, .note = "two criteria, one aligned pass" },
    .{ .func = "SUMIFS", .formula = "SUMIFS(D1:D5,A1:A5,\">0\",B1:B5,\"y\",C1:C5,\"<40\")", .expect = .{ .number = 200 }, .note = "§5.6a's 3+-criteria fixture" },
    .{ .func = "SUMIFS", .formula = "SUMIFS(D:D,A:A,\">2\")", .expect = .{ .number = 700 }, .note = "whole columns ride the same blank runs the benches time" },
    .{ .func = "SUMIFS", .formula = "SUMIFS(D:D,A:A,\">2\",B:B,\"y\")", .expect = .{ .number = 400 } },
    .{ .func = "SUMIFS", .formula = "SUMIFS(D1:D4,A1:A5,\">2\")", .expect = .{ .err = .value }, .note = "`*IFS` requires equal dimensions — no SUMIF projection here (§5.6a)" },
    .{ .func = "SUMIFS", .formula = "SUMIFS((D1:D2,D3:D5),A1:A5,\">0\")", .expect = .{ .err = .value }, .note = "a multi-area union is not one rectangle" },
    .{ .func = "SUMIFS", .formula = "SUMIFS(B1:B5,A1:A5,\">0\")", .expect = .{ .number = 0 }, .note = "text in the sum range contributes nothing" },
    .{ .func = "SUMIFS", .formula = "SUMIFS(D1:D5,B1:B5,\"x*\")", .expect = .{ .number = 900 }, .note = "wildcards reach the family through `criteria.parse`, nowhere else" },
    .{ .func = "SUMIFS", .formula = "SUMIFS(D1:D5,B1:B5,\"X\")", .expect = .{ .number = 900 }, .note = "`collation_v1`: `X` fold-equals `x`" },

    // ── COUNTIFS: §5.3c's criterion kinds, one each — direct number,
    //    text, bool, error, `\"\"`, `<>`, and the type restriction ──
    .{ .func = "COUNTIFS", .formula = "COUNTIFS(A1:A5,\">1\")", .expect = .{ .number = 3 } },
    .{ .func = "COUNTIFS", .formula = "COUNTIFS(A1:A5,\">1\",B1:B5,\"x\")", .expect = .{ .number = 1 } },
    .{ .func = "COUNTIFS", .formula = "COUNTIFS(A1:A5,2)", .expect = .{ .number = 1 }, .note = "direct number" },
    .{ .func = "COUNTIFS", .formula = "COUNTIFS(G1:G1,TRUE())", .expect = .{ .number = 1 }, .note = "bool criterion, bool cell" },
    .{ .func = "COUNTIFS", .formula = "COUNTIFS(A1:A5,\"\")", .expect = .{ .number = 1 }, .note = "an empty criterion is the COUNTBLANK population — A5 is the hole" },
    .{ .func = "COUNTIFS", .formula = "COUNTIFS(A1:A5,\"<>\")", .expect = .{ .number = 4 }, .note = "…and `<>` is its complement" },
    .{ .func = "COUNTIFS", .formula = "COUNTIFS(A:A,\"<>\")", .expect = .{ .number = 4 }, .note = "the whole column answers the same, through runs" },
    .{ .func = "COUNTIFS", .formula = "COUNTIFS(1:1,\">0\")", .expect = .{ .number = 3 }, .note = "full row: 1, 10, 100 — text, error and TRUE all fail the numeric restriction" },
    .{ .func = "COUNTIFS", .formula = "COUNTIFS(E1:E2,5)", .expect = .{ .number = 1 }, .note = "a numeric criterion sees only numbers" },
    .{ .func = "COUNTIFS", .formula = "COUNTIFS(E1:E2,\"x\")", .expect = .{ .number = 1 }, .note = "a text criterion sees only text" },
    .{ .func = "COUNTIFS", .formula = "COUNTIFS(B1:B5,1)", .expect = .{ .number = 0 }, .note = "type-restricted: no text cell satisfies a numeric criterion" },
    .{ .func = "COUNTIFS", .formula = "COUNTIFS(F1:F2,\"#DIV/0!\")", .expect = .{ .number = 1 }, .note = "the text spelling of an error is that error as a criterion" },
    .{ .func = "COUNTIFS", .formula = "COUNTIFS(A1:A5,F1)", .expect = .{ .number = 0 }, .note = "an error criterion matches error cells, of which A has none" },
    .{ .func = "COUNTIFS", .formula = "COUNTIFS(A1:A5,\">0\",A1:E1,\">0\")", .expect = .{ .err = .value }, .note = "5×1 beside 1×5: same count, different shape, still `#VALUE!`" },

    // ── AVERAGEIFS ──
    .{ .func = "AVERAGEIFS", .formula = "AVERAGEIFS(D1:D5,B1:B5,\"x\")", .expect = .{ .number = 300 } },
    .{ .func = "AVERAGEIFS", .formula = "AVERAGEIFS(D1:D5,A1:A5,\">1\",B1:B5,\"x\")", .expect = .{ .number = 300 } },
    .{ .func = "AVERAGEIFS", .formula = "AVERAGEIFS(D1:D5,A1:A5,\">99\")", .expect = .{ .err = .div0 }, .note = "an average over no match is `#DIV/0!` — AVERAGEIF's rule, verbatim" },
    .{ .func = "AVERAGEIFS", .formula = "AVERAGEIFS(D1:D4,A1:A5,\">0\")", .expect = .{ .err = .value } },

    // ── MINIFS / MAXIFS: the fold `ScanResult` grew for ──
    .{ .func = "MINIFS", .formula = "MINIFS(D1:D5,B1:B5,\"y\")", .expect = .{ .number = 200 } },
    .{ .func = "MAXIFS", .formula = "MAXIFS(D1:D5,B1:B5,\"y\")", .expect = .{ .number = 400 } },
    .{ .func = "MINIFS", .formula = "MINIFS(D1:D5,A1:A5,\">1\",C1:C5,\"<45\")", .expect = .{ .number = 200 } },
    .{ .func = "MINIFS", .formula = "MINIFS(D1:D5,A1:A5,\">99\")", .expect = .{ .number = 0 }, .note = "MIN over no match is 0 through the `*IFS` door, unlike AVERAGE's `#DIV/0!`" },
    .{ .func = "MAXIFS", .formula = "MAXIFS(D1:D5,A1:A5,\">99\")", .expect = .{ .number = 0 } },
    .{ .func = "MINIFS", .formula = "MINIFS(B1:B5,A1:A5,\">0\")", .expect = .{ .number = 0 }, .note = "matches WITHOUT a number are still the empty numeric set" },
    .{ .func = "MAXIFS", .formula = "MAXIFS(B1:B5,A1:A5,\">0\")", .expect = .{ .number = 0 } },
    .{ .func = "MAXIFS", .formula = "MAXIFS(D:D,B:B,\"x\")", .expect = .{ .number = 500 } },
    .{ .func = "MINIFS", .formula = "MINIFS(D1:D4,A1:A5,\">0\")", .expect = .{ .err = .value } },
    .{ .func = "MAXIFS", .formula = "MAXIFS(D1:D4,A1:A5,\">0\")", .expect = .{ .err = .value } },

    // ── ADDRESS: four abs modes × two styles, sheet text quoted the
    //    way the tokenizer reads it back ──
    .{ .func = "ADDRESS", .formula = "ADDRESS(2,3)", .expect = .{ .text = "$C$2" } },
    .{ .func = "ADDRESS", .formula = "ADDRESS(2,3,2)", .expect = .{ .text = "C$2" } },
    .{ .func = "ADDRESS", .formula = "ADDRESS(2,3,3)", .expect = .{ .text = "$C2" } },
    .{ .func = "ADDRESS", .formula = "ADDRESS(2,3,4)", .expect = .{ .text = "C2" } },
    .{ .func = "ADDRESS", .formula = "ADDRESS(2,3,1,FALSE())", .expect = .{ .text = "R2C3" } },
    .{ .func = "ADDRESS", .formula = "ADDRESS(2,3,2,FALSE())", .expect = .{ .text = "R2C[3]" }, .note = "the relative half is the bracketed one" },
    .{ .func = "ADDRESS", .formula = "ADDRESS(2,3,3,FALSE())", .expect = .{ .text = "R[2]C3" } },
    .{ .func = "ADDRESS", .formula = "ADDRESS(2,3,4,FALSE())", .expect = .{ .text = "R[2]C[3]" } },
    .{ .func = "ADDRESS", .formula = "ADDRESS(2,3,1,TRUE(),\"Sheet1\")", .expect = .{ .text = "Sheet1!$C$2" } },
    .{ .func = "ADDRESS", .formula = "ADDRESS(2,3,1,TRUE(),\"EXCEL SHEET\")", .expect = .{ .text = "'EXCEL SHEET'!$C$2" }, .note = "a space forces the quotes" },
    .{ .func = "ADDRESS", .formula = "ADDRESS(2,3,1,FALSE(),\"[Book1]Sheet1\")", .expect = .{ .text = "'[Book1]Sheet1'!R2C3" }, .note = "Excel's own documented example" },
    .{ .func = "ADDRESS", .formula = "ADDRESS(1,1,1,TRUE(),\"\")", .expect = .{ .text = "!$A$1" }, .note = "the empty sheet text keeps its bare `!`" },
    .{ .func = "ADDRESS", .formula = "ADDRESS(1,1,1,TRUE(),\"O'Brien\")", .expect = .{ .text = "'O''Brien'!$A$1" }, .note = "an embedded quote doubles" },
    .{ .func = "ADDRESS", .formula = "ADDRESS(2.9,3.9)", .expect = .{ .text = "$C$2" }, .note = "coordinates truncate toward zero" },
    .{ .func = "ADDRESS", .formula = "ADDRESS(1048576,16384)", .expect = .{ .text = "$XFD$1048576" }, .note = "the grid's far corner spells" },
    .{ .func = "ADDRESS", .formula = "ADDRESS(1048577,1)", .expect = .{ .err = .value }, .note = "…and one row past it does not" },
    .{ .func = "ADDRESS", .formula = "ADDRESS(1,16385)", .expect = .{ .err = .value } },
    .{ .func = "ADDRESS", .formula = "ADDRESS(0,1)", .expect = .{ .err = .value } },
    .{ .func = "ADDRESS", .formula = "ADDRESS(1,0)", .expect = .{ .err = .value } },
    .{ .func = "ADDRESS", .formula = "ADDRESS(1,1,0)", .expect = .{ .err = .value } },
    .{ .func = "ADDRESS", .formula = "ADDRESS(1,1,5)", .expect = .{ .err = .value } },
    .{ .func = "ADDRESS", .formula = "ADDRESS(1,1,,FALSE())", .expect = .{ .err = .value }, .note = "an elided abs became blank and then 0 (`LEFT(a,)`'s rule), and 0 is not a mode" },
    .{ .func = "ADDRESS", .formula = "ADDRESS(2,3,1,,\"S\")", .expect = .{ .text = "S!R2C3" }, .note = "an elided a1 is blank, and blank is FALSE — omission's TRUE needs the slot absent" },
    .{ .func = "ADDRESS", .formula = "INDIRECT(ADDRESS(2,3))", .expect = .{ .number = 20 }, .note = "the round trip: what ADDRESS spells, INDIRECT reads" },
    .{ .func = "ADDRESS", .formula = "ADDRESS({1;2},1)", .expect = .{ .array = .{ .rows = 2, .cols = 1, .cells = &.{ .{ .text = "$A$1" }, .{ .text = "$A$2" } } } }, .note = "all-scalar slots lift elementwise (M7a)" },
};

test "M7b2: every F2-criteria fixture evaluates to what the spec says" {
    for (f2c_cases) |c| {
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        try putF2cCells(&h);

        const v = h.eval(c.formula) catch |e| {
            std.debug.print("F2-criteria case `{s}` refused: {t}\n", .{ c.formula, e });
            return e;
        };
        expectValue(c.expect, v) catch |e| {
            std.debug.print("F2-criteria case `{s}` ({s})\n", .{ c.formula, c.note });
            return e;
        };
    }
}

test "M7b2: all six frozen names resolve, and each has a fixture" {
    var it = registry.inventory();
    var batch: usize = 0;
    while (it.next()) |e| {
        if (!std.mem.eql(u8, e.milestone, "M7b2")) continue;
        batch += 1;

        if (registry.lookup(e.name) == null) {
            std.debug.print("F2-criteria name does not resolve: {s}\n", .{e.name});
            return error.UnregisteredBatchFunction;
        }
        var fixtures: usize = 0;
        for (f2c_cases) |c| {
            if (std.mem.eql(u8, c.func, e.name)) fixtures += 1;
        }
        if (fixtures == 0) {
            std.debug.print("F2-criteria name has no fixture: {s}\n", .{e.name});
            return error.UnfixturedBatchFunction;
        }
    }
    try testing.expectEqual(@as(usize, 6), batch);

    for (f2c_cases) |c| {
        var found = false;
        var it2 = registry.inventory();
        while (it2.next()) |e| {
            if (std.mem.eql(u8, e.name, c.func) and std.mem.eql(u8, e.milestone, "M7b2")) found = true;
        }
        if (!found) {
            std.debug.print("fixture names a function outside F2-criteria: {s}\n", .{c.func});
            return error.FixtureOutsideBatch;
        }
    }
}

test "M7b2: the evidence label on every fixture is true of the committed manifests" {
    var oracle_rows: usize = 0;
    var excluded_rows: usize = 0;
    for (f2c_cases) |c| {
        switch (try manifestVerdict(c.formula)) {
            .decided => {
                if (c.evidence != .oracle) return error.UnderstatedEvidence;
                oracle_rows += 1;
            },
            .excluded => {
                if (c.evidence != .spec_pinned) return error.ExcludedCellClaimedAsEvidence;
                excluded_rows += 1;
            },
            .silent => {
                if (c.evidence != .spec_pinned) return error.UnbackedOracleClaim;
            },
        }
    }
    // The committed manifests predate every F2-criteria name; the
    // parked Excel leg is what would move these.
    try testing.expectEqual(@as(usize, 0), oracle_rows);
    try testing.expectEqual(@as(usize, 0), excluded_rows);
}

test "M7b2: error order in every name of the batch (§5.3c)" {
    // Every case runs in both argument orders, because a fixture with
    // one error in it proves propagation and says nothing about order.
    // The five folds are `per_function_provenance` — an error in a
    // range is a value a criterion may MATCH, never one the call
    // propagates — and ADDRESS is `.propagate`, taking §5.3c's
    // declaration order from the dispatcher.
    const Case = struct { formula: []const u8, expect: Expect, note: []const u8 = "" };
    const cases = [_]Case{
        .{ .formula = "SUMIFS(A1:A1,F1:F1,F2)", .expect = .{ .number = 0 }, .note = "criterion #N/A against cell #DIV/0!: no match, no propagation" },
        .{ .formula = "SUMIFS(F1:F1,A1:A1,1)", .expect = .{ .number = 0 }, .note = "an error in the sum range is ignored by the numeric fold, not propagated" },
        .{ .formula = "SUMIFS(D1:D1,F1:F1,F1)", .expect = .{ .number = 100 }, .note = "…and a criterion CAN match one" },
        .{ .formula = "COUNTIFS(F1:F1,F2)", .expect = .{ .number = 0 } },
        .{ .formula = "COUNTIFS(F2:F2,F1)", .expect = .{ .number = 0 } },
        .{ .formula = "COUNTIFS(F1:F1,F1)", .expect = .{ .number = 1 } },
        .{ .formula = "COUNTIFS(A1:A5,1/0)", .expect = .{ .number = 0 }, .note = "a computed error criterion is still a criterion" },
        .{ .formula = "AVERAGEIFS(F1:F1,A1:A1,1)", .expect = .{ .err = .div0 }, .note = "#DIV/0! from the empty numeric set, NOT from F1" },
        .{ .formula = "AVERAGEIFS(D1:D1,F1:F1,F2)", .expect = .{ .err = .div0 }, .note = "#DIV/0! from no match — same answer, other route" },
        .{ .formula = "MINIFS(F1:F1,A1:A1,1)", .expect = .{ .number = 0 } },
        .{ .formula = "MINIFS(D1:D1,F1:F1,F2)", .expect = .{ .number = 0 } },
        .{ .formula = "MAXIFS(F1:F1,A1:A1,1)", .expect = .{ .number = 0 } },
        .{ .formula = "MAXIFS(D1:D1,F1:F1,F2)", .expect = .{ .number = 0 } },
        .{ .formula = "ADDRESS(F1,F2)", .expect = .{ .err = .div0 }, .note = "`.propagate`: declaration order, first error wins" },
        .{ .formula = "ADDRESS(F2,F1)", .expect = .{ .err = .na } },
    };

    for (cases) |c| {
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        try putF2cCells(&h);

        const v = h.eval(c.formula) catch |e| {
            std.debug.print("error-order case `{s}` refused: {t}\n", .{ c.formula, e });
            return e;
        };
        expectValue(c.expect, v) catch |e| {
            std.debug.print("error-order case `{s}`: wrong value ({s})\n", .{ c.formula, c.note });
            return e;
        };
    }

    // Every name in the batch appears above, derived from the
    // inventory rather than typed out, at least twice: one order is
    // not an order.
    var it = registry.inventory();
    var covered_names: usize = 0;
    while (it.next()) |e| {
        if (!std.mem.eql(u8, e.milestone, "M7b2")) continue;
        var seen: usize = 0;
        for (cases) |c| {
            if (std.mem.startsWith(u8, c.formula, e.name) and c.formula[e.name.len] == '(') seen += 1;
        }
        if (seen < 2) {
            std.debug.print("batch name with {d} error-order fixtures: {s}\n", .{ seen, e.name });
            return error.MissingErrorOrderFixture;
        }
        covered_names += 1;
    }
    try testing.expectEqual(@as(usize, 6), covered_names);
}

test "M7b2: an unpaired criteria tail refuses, in every fold" {
    // A lone trailing range is an argument list Excel refuses at
    // entry, the same way it refuses a wrong arity — `MalformedInput`,
    // a plane-2 refusal, not a value. The registry's min/max cannot
    // say it (the tail is unbounded), so the implementations do.
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try putF2cCells(&h);

    try testing.expectError(error.MalformedInput, h.eval("SUMIFS(D1:D5,A1:A5,\">1\",B1:B5)"));
    try testing.expectError(error.MalformedInput, h.eval("COUNTIFS(A1:A5,\">1\",B1:B5)"));
    try testing.expectError(error.MalformedInput, h.eval("AVERAGEIFS(D1:D5,A1:A5,\">1\",B1:B5)"));
    try testing.expectError(error.MalformedInput, h.eval("MINIFS(D1:D5,A1:A5,\">1\",B1:B5)"));
    try testing.expectError(error.MalformedInput, h.eval("MAXIFS(D1:D5,A1:A5,\">1\",B1:B5)"));
}

test "M7b2: every criteria name against every argument shape, exhaustively and in both modes" {
    // M4e's enumeration, pointed at the six new names: every shape in
    // the alphabet at one and two arguments, each padded to the name's
    // own minimum arity (`SUMIFS` runs at three, not at the two it
    // would reject), both rule tables, every input evaluated twice.
    const f2c_names = [_][]const u8{
        "SUMIFS", "COUNTIFS", "AVERAGEIFS", "MINIFS", "MAXIFS", "ADDRESS",
    };

    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    var fake = env.Fake.init(testing.allocator);
    defer fake.deinit();
    const sheet = try fuzzF1bEnv(&fake);

    var checked: usize = 0;
    for (f2c_names) |name| {
        for (f1b_arg_shapes) |first_arg| {
            var buf: [256]u8 = undefined;
            if (buildF1bCall(&buf, name, &.{first_arg})) |one| {
                try sweepShape(&arena_state, &fake, sheet, one, &checked);
            }
            for (f1b_arg_shapes) |second_arg| {
                var pair_buf: [256]u8 = undefined;
                const two = buildF1bCall(&pair_buf, name, &.{ first_arg, second_arg }) orelse continue;
                try sweepShape(&arena_state, &fake, sheet, two, &checked);
            }
            _ = arena_state.reset(.retain_capacity);
        }
    }
    // A sweep that silently stopped enumerating would still pass, so
    // the count is asserted as a floor rather than left to the loops.
    try testing.expect(checked > 5_000);
}

// ─── M7b3: the F2-stats batch (§7, eleven names; one collection) ──

/// The battery's grid. A is the population set — VAR.P/STDEV.P land
/// exact, MODE.SNGL's winner lives here, and RANK.EQ's ties do. B is
/// the sample set (VAR.S/STDEV.S exact). C is the percentile set:
/// n − 1 = 4, so every k that is a multiple of 1/8 interpolates
/// exactly and the knot fixtures compare equal rather than close. D is
/// three tied pairs for MODE.SNGL's tie-break, E the range-restriction
/// mix with a hole at E4, F §5.3c's two errors in M7b2's spelling.
fn putF2sCells(h: *Harness) !void {
    try h.put("A1", num(2));
    try h.put("A2", num(4));
    try h.put("A3", num(4));
    try h.put("A4", num(4));
    try h.put("A5", num(5));
    try h.put("A6", num(5));
    try h.put("A7", num(7));
    try h.put("A8", num(9));
    try h.put("B1", num(1));
    try h.put("B2", num(3));
    try h.put("B3", num(5));
    try h.put("C1", num(15));
    try h.put("C2", num(20));
    try h.put("C3", num(35));
    try h.put("C4", num(40));
    try h.put("C5", num(50));
    try h.put("D1", num(5));
    try h.put("D2", num(5));
    try h.put("D3", num(2));
    try h.put("D4", num(2));
    try h.put("D5", num(9));
    try h.put("D6", num(9));
    try h.put("E1", .{ .text = "x" });
    try h.put("E2", num(7));
    try h.put("E3", .{ .boolean = true });
    try h.put("F1", value.ScalarValue.errorOf(.div0));
    try h.put("F2", value.ScalarValue.errorOf(.na));
}

const f2s_cases = [_]F2Case{
    // ── MEDIAN: both parities, the range/direct split, the union ──
    .{ .func = "MEDIAN", .formula = "MEDIAN(B1:B3)", .expect = .{ .number = 3 } },
    .{ .func = "MEDIAN", .formula = "MEDIAN(A1:A8)", .expect = .{ .number = 4.5 }, .note = "the even case averages the two middles" },
    .{ .func = "MEDIAN", .formula = "MEDIAN(A:A)", .expect = .{ .number = 4.5 }, .note = "the whole column rides the same blank runs" },
    .{ .func = "MEDIAN", .formula = "MEDIAN(1,2,3,4)", .expect = .{ .number = 2.5 } },
    .{ .func = "MEDIAN", .formula = "MEDIAN(E1:E4)", .expect = .{ .number = 7 }, .note = "text, logicals and the hole are invisible through a range" },
    .{ .func = "MEDIAN", .formula = "MEDIAN(TRUE(),2)", .expect = .{ .number = 1.5 }, .note = "…and a direct logical coerces — SUM's split verbatim" },
    .{ .func = "MEDIAN", .formula = "MEDIAN(\"3\",1)", .expect = .{ .number = 2 }, .note = "direct numeric text coerces" },
    .{ .func = "MEDIAN", .formula = "MEDIAN(\"x\")", .expect = .{ .err = .value }, .note = "direct non-numeric text is `#VALUE!`, not ignored" },
    .{ .func = "MEDIAN", .formula = "MEDIAN(E1:E1)", .expect = .{ .err = .num }, .note = "no numbers anywhere is `#NUM!`" },
    .{ .func = "MEDIAN", .formula = "MEDIAN((A1:A2,A5:A6))", .expect = .{ .number = 4.5 }, .note = "a multi-area union is one collection here, where the criteria family answers `#VALUE!`" },
    .{ .func = "MEDIAN", .formula = "MEDIAN(1.7E+308,1.7E+308)", .expect = .{ .err = .num }, .note = "the even case overflows through its addition, to `#NUM!`" },

    // ── MODE.SNGL: the count, the tie, the two `#N/A` routes ──
    .{ .func = "MODE.SNGL", .formula = "MODE.SNGL(A1:A8)", .expect = .{ .number = 4 } },
    .{ .func = "MODE.SNGL", .formula = "MODE.SNGL(D1:D6)", .expect = .{ .number = 5 }, .note = "three tied pairs: §5.6a's first encounter wins, pinned pending the parked Excel leg" },
    .{ .func = "MODE.SNGL", .formula = "MODE.SNGL(2,2,3,3)", .expect = .{ .number = 2 }, .note = "the tie-break through direct arguments" },
    .{ .func = "MODE.SNGL", .formula = "MODE.SNGL(B1:B3)", .expect = .{ .err = .na }, .note = "all distinct: a value seen once is not a mode" },
    .{ .func = "MODE.SNGL", .formula = "MODE.SNGL(E1:E4)", .expect = .{ .err = .na }, .note = "…and neither is the collection of one" },

    // ── the four moment names: `.P` exact on A, `.S` exact on B ──
    .{ .func = "VAR.P", .formula = "VAR.P(A1:A8)", .expect = .{ .number = 4 } },
    .{ .func = "STDEV.P", .formula = "STDEV.P(A1:A8)", .expect = .{ .number = 2 } },
    .{ .func = "VAR.S", .formula = "VAR.S(B1:B3)", .expect = .{ .number = 4 } },
    .{ .func = "STDEV.S", .formula = "STDEV.S(B1:B3)", .expect = .{ .number = 2 } },
    .{ .func = "VAR.P", .formula = "VAR.P(5)", .expect = .{ .number = 0 }, .note = "a population of one has variance 0…" },
    .{ .func = "STDEV.P", .formula = "STDEV.P(5)", .expect = .{ .number = 0 } },
    .{ .func = "VAR.S", .formula = "VAR.S(5)", .expect = .{ .err = .div0 }, .note = "…and a sample of one divides by zero, which is the whole `.P`/`.S` distinction" },
    .{ .func = "STDEV.S", .formula = "STDEV.S(5)", .expect = .{ .err = .div0 } },
    .{ .func = "VAR.S", .formula = "VAR.S(E1:E4)", .expect = .{ .err = .div0 }, .note = "one number through the range restriction is still a sample of one" },
    .{ .func = "VAR.P", .formula = "VAR.P(E1:E1)", .expect = .{ .err = .div0 }, .note = "no numbers anywhere: what the division would have said" },
    .{ .func = "VAR.P", .formula = "VAR.P(TRUE(),3)", .expect = .{ .number = 1 }, .note = "the direct-argument coercion reaches the moments too" },

    // ── PERCENTILE.INC: the five knots, and between them (§7) ──
    .{ .func = "PERCENTILE.INC", .formula = "PERCENTILE.INC(C1:C5,0)", .expect = .{ .number = 15 } },
    .{ .func = "PERCENTILE.INC", .formula = "PERCENTILE.INC(C1:C5,0.25)", .expect = .{ .number = 20 } },
    .{ .func = "PERCENTILE.INC", .formula = "PERCENTILE.INC(C1:C5,0.5)", .expect = .{ .number = 35 } },
    .{ .func = "PERCENTILE.INC", .formula = "PERCENTILE.INC(C1:C5,0.75)", .expect = .{ .number = 40 } },
    .{ .func = "PERCENTILE.INC", .formula = "PERCENTILE.INC(C1:C5,1)", .expect = .{ .number = 50 } },
    .{ .func = "PERCENTILE.INC", .formula = "PERCENTILE.INC(C1:C5,0.125)", .expect = .{ .number = 17.5 }, .note = "between the 0 and 0.25 knots: rank 0.5, half-way up the first gap" },
    .{ .func = "PERCENTILE.INC", .formula = "PERCENTILE.INC(C1:C5,0.375)", .expect = .{ .number = 27.5 }, .note = "between 0.25 and 0.5" },
    .{ .func = "PERCENTILE.INC", .formula = "PERCENTILE.INC(C1:C5,0.875)", .expect = .{ .number = 45 }, .note = "between 0.75 and 1" },
    .{ .func = "PERCENTILE.INC", .formula = "PERCENTILE.INC(B1:B3,0.75)", .expect = .{ .number = 4 }, .note = "an even gap count interpolates exactly too" },
    .{ .func = "PERCENTILE.INC", .formula = "PERCENTILE.INC(A2:A2,0.3)", .expect = .{ .number = 4 }, .note = "a single number answers every k" },
    .{ .func = "PERCENTILE.INC", .formula = "PERCENTILE.INC(C1:C5,-0.1)", .expect = .{ .err = .num } },
    .{ .func = "PERCENTILE.INC", .formula = "PERCENTILE.INC(C1:C5,1.1)", .expect = .{ .err = .num } },
    .{ .func = "PERCENTILE.INC", .formula = "PERCENTILE.INC(E1:E1,0.5)", .expect = .{ .err = .num }, .note = "the empty collection is `#NUM!` before k is judged" },

    // ── QUARTILE.INC: the same five knots by their other name ──
    .{ .func = "QUARTILE.INC", .formula = "QUARTILE.INC(C1:C5,0)", .expect = .{ .number = 15 } },
    .{ .func = "QUARTILE.INC", .formula = "QUARTILE.INC(C1:C5,1)", .expect = .{ .number = 20 }, .note = "q1 IS the 0.25 knot — same cell, same answer as PERCENTILE.INC" },
    .{ .func = "QUARTILE.INC", .formula = "QUARTILE.INC(C1:C5,2)", .expect = .{ .number = 35 } },
    .{ .func = "QUARTILE.INC", .formula = "QUARTILE.INC(C1:C5,3)", .expect = .{ .number = 40 } },
    .{ .func = "QUARTILE.INC", .formula = "QUARTILE.INC(C1:C5,4)", .expect = .{ .number = 50 } },
    .{ .func = "QUARTILE.INC", .formula = "QUARTILE.INC(C1:C5,1.9)", .expect = .{ .number = 20 }, .note = "quart truncates toward zero — CHOOSE/ADDRESS's house rule, pinned pending the parked Excel leg" },
    .{ .func = "QUARTILE.INC", .formula = "QUARTILE.INC(C1:C5,5)", .expect = .{ .err = .num } },
    .{ .func = "QUARTILE.INC", .formula = "QUARTILE.INC(C1:C5,-1)", .expect = .{ .err = .num } },

    // ── RANK.EQ: ties share the top rank, in both directions ──
    .{ .func = "RANK.EQ", .formula = "RANK.EQ(4,A1:A8)", .expect = .{ .number = 5 }, .note = "descending by default: 9, 7 and the two 5s outrank every 4" },
    .{ .func = "RANK.EQ", .formula = "RANK.EQ(4,A1:A8,0)", .expect = .{ .number = 5 } },
    .{ .func = "RANK.EQ", .formula = "RANK.EQ(4,A1:A8,1)", .expect = .{ .number = 2 }, .note = "ascending: the three 4s share rank 2, the top of the tie" },
    .{ .func = "RANK.EQ", .formula = "RANK.EQ(5,A1:A8,2)", .expect = .{ .number = 5 }, .note = "any nonzero order is ascending — Excel reads the slot as a logical" },
    .{ .func = "RANK.EQ", .formula = "RANK.EQ(9,A1:A8)", .expect = .{ .number = 1 } },
    .{ .func = "RANK.EQ", .formula = "RANK.EQ(2,A1:A8,1)", .expect = .{ .number = 1 } },
    .{ .func = "RANK.EQ", .formula = "RANK.EQ(A2,A1:A8)", .expect = .{ .number = 5 }, .note = "the number slot dereferences like any scalar slot" },
    .{ .func = "RANK.EQ", .formula = "RANK.EQ(7,E1:E4)", .expect = .{ .number = 1 }, .note = "non-numbers are invisible to the collection it ranks against" },
    .{ .func = "RANK.EQ", .formula = "RANK.EQ(6,A1:A8)", .expect = .{ .err = .na }, .note = "absent from the collection is `#N/A`" },
    .{ .func = "RANK.EQ", .formula = "RANK.EQ(7,E1:E1)", .expect = .{ .err = .na }, .note = "…including absent from the empty one" },

    // ── LARGE / SMALL: the shared sorted view, read from both ends ──
    .{ .func = "LARGE", .formula = "LARGE(A1:A8,1)", .expect = .{ .number = 9 } },
    .{ .func = "LARGE", .formula = "LARGE(A1:A8,3)", .expect = .{ .number = 5 }, .note = "duplicates occupy ranks: 9, 7, then the first 5" },
    .{ .func = "LARGE", .formula = "LARGE(A1:A8,8)", .expect = .{ .number = 2 } },
    .{ .func = "SMALL", .formula = "SMALL(A1:A8,1)", .expect = .{ .number = 2 } },
    .{ .func = "SMALL", .formula = "SMALL(A1:A8,4)", .expect = .{ .number = 4 }, .note = "2 and the three 4s" },
    .{ .func = "LARGE", .formula = "LARGE(A1:A8,1.9)", .expect = .{ .number = 9 }, .note = "k truncates toward zero — the house rule, pinned pending the parked Excel leg" },
    .{ .func = "SMALL", .formula = "SMALL(A1:A8,2.9)", .expect = .{ .number = 4 } },
    .{ .func = "LARGE", .formula = "LARGE(A1:A8,0)", .expect = .{ .err = .num } },
    .{ .func = "LARGE", .formula = "LARGE(A1:A8,9)", .expect = .{ .err = .num }, .note = "one past the collection" },
    .{ .func = "SMALL", .formula = "SMALL(A1:A8,0.5)", .expect = .{ .err = .num }, .note = "0.5 truncates to 0, and 0 is below the first rank" },
    .{ .func = "SMALL", .formula = "SMALL(E1:E1,1)", .expect = .{ .err = .num }, .note = "the empty collection has no smallest" },
    .{ .func = "LARGE", .formula = "LARGE(5,1)", .expect = .{ .number = 5 }, .note = "a direct scalar is a one-element collection" },
    .{ .func = "SMALL", .formula = "SMALL(E1:E4,1)", .expect = .{ .number = 7 } },
    .{ .func = "LARGE", .formula = "LARGE(A1:A8,{1;2})", .expect = .{ .array = .{ .rows = 2, .cols = 1, .cells = &.{ .{ .number = 9 }, .{ .number = 7 } } } }, .note = "the scalar slot lifts, the collection holds (M7a)" },
};

test "M7b3: every F2-stats fixture evaluates to what the spec says" {
    for (f2s_cases) |c| {
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        try putF2sCells(&h);

        const v = h.eval(c.formula) catch |e| {
            std.debug.print("F2-stats case `{s}` refused: {t}\n", .{ c.formula, e });
            return e;
        };
        expectValue(c.expect, v) catch |e| {
            std.debug.print("F2-stats case `{s}` ({s})\n", .{ c.formula, c.note });
            return e;
        };
    }
}

test "M7b3: all eleven frozen names resolve, and each has a fixture" {
    var it = registry.inventory();
    var batch: usize = 0;
    while (it.next()) |e| {
        if (!std.mem.eql(u8, e.milestone, "M7b3")) continue;
        batch += 1;

        if (registry.lookup(e.name) == null) {
            std.debug.print("F2-stats name does not resolve: {s}\n", .{e.name});
            return error.UnregisteredBatchFunction;
        }
        var fixtures: usize = 0;
        for (f2s_cases) |c| {
            if (std.mem.eql(u8, c.func, e.name)) fixtures += 1;
        }
        if (fixtures == 0) {
            std.debug.print("F2-stats name has no fixture: {s}\n", .{e.name});
            return error.UnfixturedBatchFunction;
        }
    }
    try testing.expectEqual(@as(usize, 11), batch);

    for (f2s_cases) |c| {
        var found = false;
        var it2 = registry.inventory();
        while (it2.next()) |e| {
            if (std.mem.eql(u8, e.name, c.func) and std.mem.eql(u8, e.milestone, "M7b3")) found = true;
        }
        if (!found) {
            std.debug.print("fixture names a function outside F2-stats: {s}\n", .{c.func});
            return error.FixtureOutsideBatch;
        }
    }
}

test "M7b3: the evidence label on every fixture is true of the committed manifests" {
    var oracle_rows: usize = 0;
    var excluded_rows: usize = 0;
    for (f2s_cases) |c| {
        switch (try manifestVerdict(c.formula)) {
            .decided => {
                if (c.evidence != .oracle) return error.UnderstatedEvidence;
                oracle_rows += 1;
            },
            .excluded => {
                if (c.evidence != .spec_pinned) return error.ExcludedCellClaimedAsEvidence;
                excluded_rows += 1;
            },
            .silent => {
                if (c.evidence != .spec_pinned) return error.UnbackedOracleClaim;
            },
        }
    }
    // The committed manifests predate every F2-stats name; the parked
    // Excel leg is what would move these.
    try testing.expectEqual(@as(usize, 0), oracle_rows);
    try testing.expectEqual(@as(usize, 0), excluded_rows);
}

test "M7b3: error order in every name of the batch (§5.3c)" {
    // Every case runs in both argument orders, because a fixture with
    // one error in it proves propagation and says nothing about order.
    // An error anywhere in the batch's inputs is the fold's error —
    // the opposite side of §5.3c's line from the criteria family,
    // where an error is a value a criterion may match — and the ORDER
    // is declaration order everywhere: from the dispatcher for the
    // variadic six (every slot `.aggregate`, the collector reads
    // §5.6a's order inside), from the implementations for the fixed
    // five, whose collection is a reference the dispatcher cannot see
    // an error inside — the lookups' arrangement, opposite verdict.
    const Case = struct { formula: []const u8, expect: Expect, note: []const u8 = "" };
    const cases = [_]Case{
        .{ .formula = "MEDIAN(F1,F2)", .expect = .{ .err = .div0 }, .note = "declaration order, first error wins" },
        .{ .formula = "MEDIAN(F2,F1)", .expect = .{ .err = .na } },
        .{ .formula = "MEDIAN(F1:F2)", .expect = .{ .err = .div0 }, .note = "inside a range the collector reads §5.6a's order" },
        .{ .formula = "MODE.SNGL(F1,F2)", .expect = .{ .err = .div0 } },
        .{ .formula = "MODE.SNGL(F2,F1)", .expect = .{ .err = .na } },
        .{ .formula = "STDEV.P(F1,F2)", .expect = .{ .err = .div0 } },
        .{ .formula = "STDEV.P(F2,F1)", .expect = .{ .err = .na } },
        .{ .formula = "STDEV.S(F1,F2)", .expect = .{ .err = .div0 } },
        .{ .formula = "STDEV.S(F2,F1)", .expect = .{ .err = .na } },
        .{ .formula = "VAR.P(F1,F2)", .expect = .{ .err = .div0 } },
        .{ .formula = "VAR.P(F2,F1)", .expect = .{ .err = .na } },
        .{ .formula = "VAR.S(F1,F2)", .expect = .{ .err = .div0 } },
        .{ .formula = "VAR.S(F2,F1)", .expect = .{ .err = .na } },
        .{ .formula = "PERCENTILE.INC(F1,F2)", .expect = .{ .err = .div0 } },
        .{ .formula = "PERCENTILE.INC(F2,F1)", .expect = .{ .err = .na } },
        .{ .formula = "QUARTILE.INC(F1,F2)", .expect = .{ .err = .div0 } },
        .{ .formula = "QUARTILE.INC(F2,F1)", .expect = .{ .err = .na } },
        .{ .formula = "LARGE(F1,F2)", .expect = .{ .err = .div0 } },
        .{ .formula = "LARGE(F2,F1)", .expect = .{ .err = .na } },
        .{ .formula = "LARGE(F1:F2,1)", .expect = .{ .err = .div0 }, .note = "an error inside the collection is the fold's error — the anti-criteria pin" },
        .{ .formula = "SMALL(F1,F2)", .expect = .{ .err = .div0 } },
        .{ .formula = "SMALL(F2,F1)", .expect = .{ .err = .na } },
        .{ .formula = "RANK.EQ(F1,F2)", .expect = .{ .err = .div0 } },
        .{ .formula = "RANK.EQ(F2,F1)", .expect = .{ .err = .na } },
        .{ .formula = "RANK.EQ(F2,A1:A8,F1)", .expect = .{ .err = .na }, .note = "declaration order across the collection slot: the first SCALAR error wins" },
        .{ .formula = "RANK.EQ(4,F1:F1)", .expect = .{ .err = .div0 }, .note = "…and an error inside it is found by the collector" },
    };

    for (cases) |c| {
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        try putF2sCells(&h);

        const v = h.eval(c.formula) catch |e| {
            std.debug.print("error-order case `{s}` refused: {t}\n", .{ c.formula, e });
            return e;
        };
        expectValue(c.expect, v) catch |e| {
            std.debug.print("error-order case `{s}`: wrong value ({s})\n", .{ c.formula, c.note });
            return e;
        };
    }

    // Every name in the batch appears above, derived from the
    // inventory rather than typed out, at least twice: one order is
    // not an order.
    var it = registry.inventory();
    var covered_names: usize = 0;
    while (it.next()) |e| {
        if (!std.mem.eql(u8, e.milestone, "M7b3")) continue;
        var seen: usize = 0;
        for (cases) |c| {
            if (std.mem.startsWith(u8, c.formula, e.name) and c.formula[e.name.len] == '(') seen += 1;
        }
        if (seen < 2) {
            std.debug.print("batch name with {d} error-order fixtures: {s}\n", .{ seen, e.name });
            return error.MissingErrorOrderFixture;
        }
        covered_names += 1;
    }
    try testing.expectEqual(@as(usize, 11), covered_names);
}

test "M7b3: every stats name against every argument shape, exhaustively and in both modes" {
    // M4e's enumeration, pointed at the eleven new names: every shape
    // in the alphabet at one and two arguments, each padded to the
    // name's own minimum arity (`PERCENTILE.INC` runs at two, not the
    // one it would reject), both rule tables, every input evaluated
    // twice.
    const f2s_names = [_][]const u8{
        "MEDIAN", "MODE.SNGL",    "STDEV.P",        "STDEV.S",
        "VAR.P",  "VAR.S",        "RANK.EQ",        "LARGE",
        "SMALL",  "QUARTILE.INC", "PERCENTILE.INC",
    };

    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    var fake = env.Fake.init(testing.allocator);
    defer fake.deinit();
    const sheet = try fuzzF1bEnv(&fake);

    var checked: usize = 0;
    for (f2s_names) |name| {
        for (f1b_arg_shapes) |first_arg| {
            var buf: [256]u8 = undefined;
            if (buildF1bCall(&buf, name, &.{first_arg})) |one| {
                try sweepShape(&arena_state, &fake, sheet, one, &checked);
            }
            for (f1b_arg_shapes) |second_arg| {
                var pair_buf: [256]u8 = undefined;
                const two = buildF1bCall(&pair_buf, name, &.{ first_arg, second_arg }) orelse continue;
                try sweepShape(&arena_state, &fake, sheet, two, &checked);
            }
            _ = arena_state.reset(.retain_capacity);
        }
    }
    // A sweep that silently stopped enumerating would still pass, so
    // the count is asserted as a floor rather than left to the loops —
    // eleven names to M7b2's six, and the floor scales with them.
    try testing.expect(checked > 9_000);
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
    try testing.expectEqual(parser.PlaneTwo.FormulaUnsupportedConstruct, planeTwo(error.UnsupportedConstruct));
    try testing.expectEqual(parser.PlaneTwo.FormulaUnsupportedConstruct, planeTwo(error.NameRefused));
    // M4b3 deleted the list's one entry. It stays empty until a row
    // needs it again, and the row that does writes a milestone next to
    // whatever it adds.
    try testing.expectEqual(@as(usize, 0), not_yet_implemented.len);
}

test "refusals: an unregistered call refuses, an unresolvable name is #NAME?" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();

    // §7: unregistered calls refuse rather than inventing `#NAME?` for a
    // function zlsx simply does not implement.
    // `TEXT` is frozen in the inventory for M8a and unregistered
    // today; `VLOOKUP` stood here until M4e registered it, `SUMIFS`
    // until M7b2, `MEDIAN` until M7b3.
    try testing.expectError(error.UnsupportedFunction, h.eval("TEXT(A1,\"0\")"));
    try testing.expectError(error.UnsupportedFunction, h.eval("NOTAFUNCTION()"));
    // §5.9: a value-position name that provably resolves nowhere is a
    // plane-1 `#NAME?`, which is a successful result.
    try testing.expectEqual(value.KnownError.name, (try h.scalar("SomeName+1")).err.known);
    // Wrong arity is a formula Excel could not have written.
    try testing.expectError(error.MalformedInput, h.eval("SQRT(1,2)"));
    try testing.expectError(error.MalformedInput, h.eval("IFERROR(1)"));
}

test "refusals: a 3D span no longer refuses as not-yet-implemented" {
    // The test the deletion is watched by. Until M4b3 this expected
    // `NotYetImplemented`, which was the list's one entry; a span that
    // refuses that way again means the entry came back without the
    // milestone note the list exists to carry.
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    _ = try h.fake.addSheet("Q1");

    // Ineligible consumer: a refusal, but a §5.6g one.
    try testing.expectError(error.UnsupportedConstruct, h.eval("COUNTBLANK('Q1:Q4'!A1)"));
    try testing.expectEqual(
        name_rules.Refusal.Reason.three_d_ineligible_function,
        h.ev.last_three_d.?.reason,
    );
    // And an eligible one evaluates. `Q4` does not exist, so the span
    // is `#REF!` — a value, which is the proof it was evaluated at all.
    try testing.expectEqual(value.KnownError.ref, (try h.scalar("SUM('Q1:Q4'!A1)")).err.known);

    // Nothing anywhere in the evaluator answers `NotYetImplemented` for
    // a span any more, in either spelling.
    for ([_][]const u8{ "SUM(Q1:Q4!A1)", "'Q1:Q4'!A1", "SUM('Q1:Q4'!A1:B2)" }) |src| {
        const r = h.eval(src);
        if (r) |_| {} else |e| try testing.expect(e != error.NotYetImplemented);
    }
}

// ─── 3D references (§5.6g, M4b3) ─────────────────────────────────

test "3D: an eligible function aggregates a span, inclusively and in workbook order" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    // Sheet1 is index 0; the span below covers all three.
    const s2 = try h.fake.addSheet("Sheet2");
    const s3 = try h.fake.addSheet("Sheet3");
    try h.fake.putA1(h.sheet, .stored, "A1", num(1));
    try h.fake.putA1(s2, .stored, "A1", num(10));
    try h.fake.putA1(s3, .stored, "A1", num(100));

    try testing.expectEqual(@as(f64, 111), (try h.scalar("SUM(Sheet1:Sheet3!A1)")).number);
    // Inclusive at both ends, and a one-sheet span is a span.
    try testing.expectEqual(@as(f64, 110), (try h.scalar("SUM(Sheet2:Sheet3!A1)")).number);
    try testing.expectEqual(@as(f64, 10), (try h.scalar("SUM(Sheet2:Sheet2!A1)")).number);
    // Case folding is the symbol layer's — the in-memory fake matches
    // sheet names exactly, and `symbols.resolveSheetSpan` is where the
    // folded span lookup is proven.
    // A range target expands per member, not once.
    try h.fake.putA1(s3, .stored, "B1", num(1000));
    try testing.expectEqual(@as(f64, 1111), (try h.scalar("SUM(Sheet1:Sheet3!A1:B1)")).number);
}

test "3D: the other two eligible functions this row can run agree with SUM" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    const s2 = try h.fake.addSheet("Sheet2");
    try h.fake.putA1(h.sheet, .stored, "A1", num(1));
    try h.fake.putA1(s2, .stored, "A1", .{ .text = "x" });

    // AVERAGE, MIN and MAX are M4e's — their eligibility is fixtured in
    // `names.zig` against the frozen matrix, and their oracle legs land
    // with the functions. COUNT and COUNTA are registered here.
    try testing.expectEqual(@as(f64, 1), (try h.scalar("COUNT(Sheet1:Sheet2!A1)")).number);
    try testing.expectEqual(@as(f64, 2), (try h.scalar("COUNTA(Sheet1:Sheet2!A1)")).number);
}

test "3D: a missing or reordered endpoint pins #REF!, and evaluates nothing" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    const s2 = try h.fake.addSheet("Sheet2");
    try h.fake.putA1(h.sheet, .stored, "A1", num(1));
    try h.fake.putA1(s2, .stored, "A1", num(10));

    // A deleted endpoint. Excel leaves the spelling in place rather
    // than repairing the span.
    try testing.expectEqual(
        value.KnownError.ref,
        (try h.scalar("SUM(Sheet1:Gone!A1)")).err.known,
    );
    try testing.expectEqual(
        value.KnownError.ref,
        (try h.scalar("SUM(Gone:Sheet2!A1)")).err.known,
    );
    // Reordered endpoints are not silently normalized: `Sheet2:Sheet1`
    // is a span someone's edit broke, and summing between them would
    // answer a question no one asked.
    try testing.expectEqual(
        value.KnownError.ref,
        (try h.scalar("SUM(Sheet2:Sheet1!A1)")).err.known,
    );
}

test "3D: array and intersection contexts refuse before anything is evaluated" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    _ = try h.fake.addSheet("Sheet2");

    var opts = h.options();
    opts.array_formula = true;
    try testing.expectError(
        error.UnsupportedConstruct,
        h.evalOpts("SUM(Sheet1:Sheet2!A1)", opts),
    );
    try testing.expectEqual(
        name_rules.Refusal.Reason.three_d_in_array_context,
        h.ev.last_three_d.?.reason,
    );

    try testing.expectError(error.UnsupportedConstruct, h.eval("SUM(@Sheet1:Sheet2!A1)"));
    try testing.expectEqual(
        name_rules.Refusal.Reason.three_d_in_intersection_context,
        h.ev.last_three_d.?.reason,
    );
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
        // M4c: every new impl, plus the shapes that make one of them
        // allocate — a range walk, a multi-area set, and the lifted
        // array path through a scalar slot.
        "ISERR(1/0)",
        "ISNA(A1:B1)",
        "ISLOGICAL((A1:A2,B1:B1))",
        "ISTEXT(A2)",
        "N(A2)&T(A2)",
        "N(A1:B1)",
        "IFS(ISERR(1/0),T(A2),TRUE(),N(A1))",
        "SWITCH(N(A1),2,T(A2),\"z\")",
        "AND(ISNA(NA()),OR(ISTEXT(A2),ISLOGICAL(A1)))",
        "NOT(ISBLANK(A1))",
        // M4d: every new impl at least once, plus the shapes that make
        // one of them allocate — the lifted array path through a scalar
        // slot, a multi-area set, and a range walk.
        "ABS(A1)+SIGN(A1)+INT(A1)",
        "ROUND(A1,2)+ROUNDUP(A1,2)+ROUNDDOWN(A1,2)",
        "TRUNC(A1)+TRUNC(A1,1)",
        "MOD(A1,B1)+POWER(A1,B1)",
        "EXP(A1)+LN(A1)+LOG(A1)+LOG10(A1)+SQRT(A1)+PI()",
        "LOG(A1,B1)",
        "RANDBETWEEN(A1,B1)+RAND()",
        "ABS(A1:B1)",
        "ROUND(A1:B1,0)",
        "RANDBETWEEN(A1:B1,A1:B1)",
        "TRUNC((A1:B1,A1:A1),0)",
        "IFERROR(SQRT(-A1),LOG(A2))",
        // M4e: every new impl at least once, plus the shapes that make
        // one of them allocate — a materialized lookup table, a folded
        // wildcard match, a criteria scan over a multi-area set, and
        // the two names that build an array to return.
        "AVERAGE(A1:B1,A2,TRUE())",
        "MIN(A1:B1)+MAX(A1:B1)",
        "SUMPRODUCT(A1:B1,A1:B1)",
        "SUMPRODUCT({1,2},{3,4})",
        "AVERAGEIF(A1:B1,\">1\",A1:B1)",
        "COUNTIF((A1:A2,B1:B1),\">1\")",
        "VLOOKUP(2,A1:B1,2)",
        "VLOOKUP(\"a*\",A2:B2,1,FALSE)",
        "HLOOKUP(2,A1:B1,1)",
        "MATCH(\"ABC\",A2:B2,0)",
        "XMATCH(2,A1:B1,-1,-1)",
        "XLOOKUP(2,A1:B1,A1:B1,\"none\")",
        "XLOOKUP(2,A1:A1,A1:B1)",
        "INDEX(A1:B1,1,2)",
        "INDEX(A1:B1,0,0)",
        "ROW(A1)+COLUMN(B1)+ROWS(A1:B1)+COLUMNS(A1:B1)",
        "CHOOSE(2,A1,A2)",
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

// ─── M3b: run inputs, the byte budget, and criteria end to end ───

test "budget: an exhausted category is a refusal, not an allocation failure" {
    // `std.mem.Allocator` has one way to say no; §9 has five reasons.
    // The mapping is what recovers the difference, and it must hold for
    // every category — including the three whose charge sites arrive
    // with later rows.
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();

    inline for (@typeInfo(run_inputs.Category).@"enum".fields) |f| {
        const cat: run_inputs.Category = @enumFromInt(f.value);
        var budget = run_inputs.Budget.init(testing.allocator, .{});
        var opts = h.options();
        opts.budget = &budget;
        var ev = Evaluator.init(h.arena(), h.fake.evalEnv(), opts);
        defer ev.deinit();

        // Untripped: a genuine allocator failure stays one.
        try testing.expectEqual(EvalError.OutOfMemory, ev.mapBudget(error.OutOfMemory));
        budget.charge(cat, budget.limits.get(cat) + 1) catch {};
        try testing.expectEqual(@as(?run_inputs.Category, cat), budget.tripped);
        try testing.expectEqual(EvalError.LimitExceeded, ev.mapBudget(error.OutOfMemory));
        // …and every other error passes through untouched.
        try testing.expectEqual(EvalError.UnsupportedFunction, ev.mapBudget(error.UnsupportedFunction));
    }
    try testing.expectEqual(parser.PlaneTwo.FormulaLimitExceeded, planeTwo(error.LimitExceeded));
}

/// Evaluate `src` against a budget whose `cat` limit is `limit`.
fn evalUnderBudget(
    gpa: std.mem.Allocator,
    cat: run_inputs.Category,
    limit: u64,
    src: []const u8,
) !void {
    var budget = run_inputs.Budget.init(gpa, .{});
    budget.limits.set(cat, limit);

    var arena_state = std.heap.ArenaAllocator.init(budget.allocator(.run_arena));
    defer arena_state.deinit();
    var fake = env.Fake.init(gpa);
    defer fake.deinit();
    const sheet = try fake.addSheet("S");
    var r: usize = 1;
    while (r <= 32) : (r += 1) {
        var buf: [8]u8 = undefined;
        const ref = std.fmt.bufPrint(&buf, "A{d}", .{r}) catch unreachable;
        try fake.putA1(sheet, .stored, ref, num(@floatFromInt(r)));
    }

    var draw_value: f64 = 0.5;
    var draws = DrawSource.constant(&draw_value);
    var ev = Evaluator.init(arena_state.allocator(), fake.evalEnv(), .{
        .current_sheet = sheet,
        .collation = .{ .fold = shippedFold },
        .draws = &draws,
        .budget = &budget,
    });
    defer ev.deinit();

    var parsed = try parser.parse(gpa, src, .{});
    defer parsed.deinit(gpa);
    if (parsed == .refused) return error.MalformedInput;
    _ = try ev.evaluate(parsed.ok);
}

test "budget: matrix cells refuse below, at, and above the limit" {
    // `A1:A32*2` materializes 32 cells and produces 32 more.
    const needed: u64 = 64;
    try evalUnderBudget(testing.allocator, .matrix_cells, needed + 1, "A1:A32*2");
    try evalUnderBudget(testing.allocator, .matrix_cells, needed, "A1:A32*2");
    try testing.expectError(
        error.LimitExceeded,
        evalUnderBudget(testing.allocator, .matrix_cells, needed - 1, "A1:A32*2"),
    );
}

test "budget: string payload refuses when concatenation outgrows it" {
    // `"aaaa…"&"bbbb…"` allocates the joined text, and nothing else here
    // is charged to the string budget.
    const src = "\"aaaaaaaaaa\"&\"bbbbbbbbbb\"";
    try evalUnderBudget(testing.allocator, .string_payload, 20, src);
    try testing.expectError(
        error.LimitExceeded,
        evalUnderBudget(testing.allocator, .string_payload, 19, src),
    );
}

test "budget: the run arena refuses when the whole run outgrows it" {
    // Generous enough to finish, then one that is not. The failure is a
    // typed limit rather than an anonymous OOM, which is the point.
    try evalUnderBudget(testing.allocator, .run_arena, 1 << 20, "SUM(A1:A32)+1");
    try testing.expectError(
        error.LimitExceeded,
        evalUnderBudget(testing.allocator, .run_arena, 64, "A1:A32*2"),
    );
}

test "rng: a draw sequence is reproducible from RunInputs alone" {
    const H = struct {
        fn run(seed: u64, out: *[2]f64) !void {
            var h: Harness = undefined;
            try h.init(testing.allocator);
            defer h.deinit();

            const inputs = run_inputs.RunInputs{
                .now_utc_ms = 0,
                .rng_seed = seed,
                .limits = .{},
            };
            try inputs.validate();

            var generator = rng.Rng.init(inputs.rng_seed);
            var source = generator.drawSource();
            var opts = h.options();
            opts.draws = &source;
            opts.fidelity = inputs.fidelity;
            opts.dialect = inputs.dialect;

            const v = try h.evalOpts("RAND()&\"|\"&RAND()", opts);
            // Two draws, in order, from one formula.
            try testing.expectEqual(@as(u64, 2), source.count);

            var it = std.mem.splitScalar(u8, v.scalar.text, '|');
            out[0] = try std.fmt.parseFloat(f64, it.next().?);
            out[1] = try std.fmt.parseFloat(f64, it.next().?);
        }
    };

    var first: [2]f64 = undefined;
    var again: [2]f64 = undefined;
    var other: [2]f64 = undefined;
    try H.run(0xC0FFEE, &first);
    try H.run(0xC0FFEE, &again);
    try H.run(0xC0FFEF, &other);

    // Equal RunInputs ⇒ equal output, bit for bit. Nothing else went
    // into the run: no clock, no entropy source.
    try testing.expectEqual(@as(u64, @bitCast(first[0])), @as(u64, @bitCast(again[0])));
    try testing.expectEqual(@as(u64, @bitCast(first[1])), @as(u64, @bitCast(again[1])));
    // A different seed is a different run.
    try testing.expect(first[0] != other[0]);
    // And the two draws within a run are distinct, which is what makes
    // "the seam is called once per draw" observable at all.
    try testing.expect(first[0] != first[1]);

    // The stream is `rng_v1`'s, not something the evaluator invented.
    var direct = rng.Rng.init(0xC0FFEE);
    try testing.expectEqual(direct.nextFloat(), first[0]);
    try testing.expectEqual(direct.nextFloat(), first[1]);
}

test "criteria: COUNTIF and SUMIF evaluate through the aligned cursor" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try h.put("A1", .{ .text = "apple" });
    try h.put("A2", .{ .text = "pear" });
    try h.put("A3", .{ .text = "APPLE" });
    try h.put("B1", num(10));
    try h.put("B2", num(20));
    try h.put("B3", num(30));

    // Case-insensitive under `collation_v1`, like every other match.
    try testing.expectEqual(@as(f64, 2), (try h.scalar("COUNTIF(A1:A3,\"apple\")")).number);
    try testing.expectEqual(@as(f64, 1), (try h.scalar("COUNTIF(A1:A3,\"pear\")")).number);
    try testing.expectEqual(@as(f64, 3), (try h.scalar("COUNTIF(A1:A3,\"*\")")).number);
    try testing.expectEqual(@as(f64, 2), (try h.scalar("COUNTIF(B1:B3,\">15\")")).number);
    // A numeric criterion is type-restricted: the text column has no
    // numbers in it, whatever the cross-type order would say.
    try testing.expectEqual(@as(f64, 0), (try h.scalar("COUNTIF(A1:A3,\">15\")")).number);

    // Three-argument SUMIF projects the sum range from its top-left.
    try testing.expectEqual(@as(f64, 40), (try h.scalar("SUMIF(A1:A3,\"apple\",B1:B3)")).number);
    try testing.expectEqual(@as(f64, 40), (try h.scalar("SUMIF(A1:A3,\"apple\",B1)")).number);
    // Two-argument SUMIF sums the criteria range itself.
    try testing.expectEqual(@as(f64, 50), (try h.scalar("SUMIF(B1:B3,\">15\")")).number);

    // A criterion may arrive as a reference rather than a literal.
    try h.put("D1", .{ .text = "pear" });
    try testing.expectEqual(@as(f64, 1), (try h.scalar("COUNTIF(A1:A3,D1)")).number);
    // A non-reference in a reference slot is `#VALUE!`, not a refusal.
    try testing.expectEqual(value.KnownError.value, (try h.scalar("COUNTIF(1,\"x\")")).err.known);
}

test "criteria: the ranges a COUNTIF reads are captured as dependencies" {
    // `criteria.scan` reads through `EvalEnv` directly rather than
    // through `Evaluator.readRange`, so this is the test that the
    // M3a2 capture-at-construction rule covers it anyway.
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try h.put("A1", .{ .text = "x" });
    try h.put("B1", num(1));

    _ = try h.eval("SUMIF(A1:A3,\"x\",B1:B3)");
    const a = coords.parseRange("A1:A3", .{}) catch unreachable;
    const b = coords.parseRange("B1:B3", .{}) catch unreachable;
    try testing.expect(h.ev.deps.hasArea(.{ .sheet = h.sheet, .range = a.normalized() }));
    try testing.expect(h.ev.deps.hasArea(.{ .sheet = h.sheet, .range = b.normalized() }));
}

test "criteria: a locale-flavoured criterion refuses rather than guessing" {
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try h.put("A1", num(2));
    try testing.expectError(error.LocaleSensitiveInput, h.eval("COUNTIF(A1:A3,\">1,5\")"));
}

// ─── M5a2: the two reference-producing rows (§7) ─────────────────
//
// Oracle-first and, again, honest about how little the oracle decides:
// no committed manifest records an `INDIRECT` or an `OFFSET` cell, so
// every row below is spec-pinned and the evidence gate asserts it in
// both directions. The parked Excel leg (§8.2) is what would move these
// labels, and the count is stated so that moving them is a visible edit.
//
// These are the fixpoint's test subjects (§5.6e), which is why they ship
// complete at this row rather than at M7b with the rest of the reference
// family: M6 exposes a public CLI, and a half-function behind it is a
// promise the ladder would have to take back.

/// One M5a2 fixture. `func` is the inventory name the row is a fixture
/// FOR, so the coverage test can derive the batch from the frozen TSV.
const RefCase = struct {
    func: []const u8,
    formula: []const u8,
    expect: Expect,
    evidence: Evidence = .spec_pinned,
    note: []const u8 = "",
};

/// The world every M5a2 fixture reads: a 3×3 block on Sheet1 whose
/// values encode their own coordinates, so a displaced reference names
/// which cell it landed on rather than only that it landed somewhere.
fn putRefCells(h: *Harness) !void {
    try h.put("A1", num(11));
    try h.put("A2", num(21));
    try h.put("A3", num(31));
    try h.put("B1", num(12));
    try h.put("B2", num(22));
    try h.put("B3", num(32));
    try h.put("C1", num(13));
    try h.put("C2", num(23));
    try h.put("C3", num(33));
    try h.put("D1", .{ .text = "B2" });
    try h.put("D2", .{ .text = "Sheet1!C3" });
    try h.put("D3", .{ .text = "not a reference" });
}

const ref_cases = [_]RefCase{
    // ── INDIRECT: text becomes an area ──
    .{ .func = "INDIRECT", .formula = "INDIRECT(\"B2\")", .expect = .{ .number = 22 } },
    .{ .func = "INDIRECT", .formula = "INDIRECT(D1)", .expect = .{ .number = 22 }, .note = "the spelling may itself come from a cell — which is what makes the reference dynamic" },
    .{ .func = "INDIRECT", .formula = "INDIRECT(\"$B$2\")", .expect = .{ .number = 22 }, .note = "anchors are part of the spelling and change nothing about where it points" },
    .{ .func = "INDIRECT", .formula = "SUM(INDIRECT(\"A1:B2\"))", .expect = .{ .number = 66 }, .note = "11+12+21+22" },
    .{ .func = "INDIRECT", .formula = "SUM(INDIRECT(\"B:B\"))", .expect = .{ .number = 66 }, .note = "a whole column: 12+22+32" },
    .{ .func = "INDIRECT", .formula = "SUM(INDIRECT(\"2:2\"))", .expect = .{ .number = 66 }, .note = "a whole row: 21+22+23" },
    .{ .func = "INDIRECT", .formula = "INDIRECT(\"Sheet1!C3\")", .expect = .{ .number = 33 } },
    .{ .func = "INDIRECT", .formula = "INDIRECT(D2)", .expect = .{ .number = 33 } },
    .{ .func = "INDIRECT", .formula = "INDIRECT(\"'Sheet1'!C3\")", .expect = .{ .number = 33 }, .note = "a quoted sheet name is the same sheet" },
    .{ .func = "INDIRECT", .formula = "INDIRECT(\"B2\",TRUE())", .expect = .{ .number = 22 } },
    // Every way of denoting nothing is `#REF!` — a value, because the
    // formula is well-formed and simply points nowhere.
    .{ .func = "INDIRECT", .formula = "INDIRECT(\"not a reference\")", .expect = .{ .err = .ref } },
    .{ .func = "INDIRECT", .formula = "INDIRECT(D3)", .expect = .{ .err = .ref } },
    .{ .func = "INDIRECT", .formula = "INDIRECT(\"\")", .expect = .{ .err = .ref } },
    .{ .func = "INDIRECT", .formula = "INDIRECT(\"Nowhere!A1\")", .expect = .{ .err = .ref }, .note = "an unknown sheet is #REF!, the same answer a deleted one gives" },
    .{ .func = "INDIRECT", .formula = "INDIRECT(\"A0\")", .expect = .{ .err = .ref }, .note = "row 0 is outside the grid" },
    .{ .func = "INDIRECT", .formula = "INDIRECT(\"A1048577\")", .expect = .{ .err = .ref }, .note = "one row past the last one" },
    .{ .func = "INDIRECT", .formula = "INDIRECT(\"XFE1\")", .expect = .{ .err = .ref }, .note = "one column past XFD" },
    .{ .func = "INDIRECT", .formula = "INDIRECT(A1)", .expect = .{ .err = .ref }, .note = "a number is not a reference spelling" },
    .{ .func = "INDIRECT", .formula = "INDIRECT(\"B2\")+1", .expect = .{ .number = 23 }, .note = "the reference dereferences into an operator like any other" },
    .{ .func = "INDIRECT", .formula = "ROW(INDIRECT(\"C3\"))", .expect = .{ .number = 3 }, .note = "…and satisfies a `.reference` slot, which a value could not" },

    // ── OFFSET: an area displaced, and optionally resized ──
    .{ .func = "OFFSET", .formula = "OFFSET(A1,1,1)", .expect = .{ .number = 22 } },
    .{ .func = "OFFSET", .formula = "OFFSET(A1,0,0)", .expect = .{ .number = 11 }, .note = "a zero displacement is the reference itself" },
    .{ .func = "OFFSET", .formula = "OFFSET(C3,-2,-2)", .expect = .{ .number = 11 }, .note = "negative displacements move up and left" },
    .{ .func = "OFFSET", .formula = "OFFSET(A1,1.9,1.9)", .expect = .{ .number = 22 }, .note = "Excel TRUNCATES the displacement toward zero rather than rounding it" },
    .{ .func = "OFFSET", .formula = "OFFSET(C3,-1.9,-1.9)", .expect = .{ .number = 22 }, .note = "…toward zero in the negative direction too, so -1.9 is -1" },
    .{ .func = "OFFSET", .formula = "SUM(OFFSET(A1,0,0,2,2))", .expect = .{ .number = 66 } },
    .{ .func = "OFFSET", .formula = "SUM(OFFSET(A1:B2,1,1))", .expect = .{ .number = 110 }, .note = "an omitted extent keeps the BASE's shape, so this is the 2x2 at B2: 22+23+32+33" },
    .{ .func = "OFFSET", .formula = "ROWS(OFFSET(A1,0,0,3,1))", .expect = .{ .number = 3 } },
    .{ .func = "OFFSET", .formula = "COLUMNS(OFFSET(A1,0,0,1,3))", .expect = .{ .number = 3 } },
    .{ .func = "OFFSET", .formula = "OFFSET(A1,-1,0)", .expect = .{ .err = .ref }, .note = "off the top edge" },
    .{ .func = "OFFSET", .formula = "OFFSET(A1,0,-1)", .expect = .{ .err = .ref }, .note = "off the left edge" },
    .{ .func = "OFFSET", .formula = "OFFSET(A1048576,1,0)", .expect = .{ .err = .ref }, .note = "off the bottom edge" },
    .{ .func = "OFFSET", .formula = "OFFSET(A1,0,0,2,0)", .expect = .{ .err = .value }, .note = "Microsoft documents height and width as POSITIVE; zero is not" },
    .{ .func = "OFFSET", .formula = "OFFSET(A1,0,0,-2,1)", .expect = .{ .err = .value }, .note = "and neither is negative — Excel 365's undocumented reverse-extent behaviour is not a claim this repo can back while the Excel oracle leg is parked" },
    .{ .func = "OFFSET", .formula = "OFFSET(\"A1\",0,0)", .expect = .{ .err = .value }, .note = "the first slot is a reference; text is #VALUE!, not a spelling to parse — that is INDIRECT's job" },
    .{ .func = "OFFSET", .formula = "SUM(OFFSET(INDIRECT(\"A1\"),1,1,2,2))", .expect = .{ .number = 110 }, .note = "the two compose: OFFSET's reference slot takes INDIRECT's product" },
};

test "M5a2: every reference fixture evaluates to what the spec says" {
    for (ref_cases) |c| {
        var h: Harness = undefined;
        try h.init(testing.allocator);
        defer h.deinit();
        try putRefCells(&h);

        const v = h.eval(c.formula) catch |e| {
            std.debug.print("M5a2 `{s}` refused: {t}\n", .{ c.formula, e });
            return e;
        };
        expectValue(c.expect, v) catch |e| {
            std.debug.print("M5a2 `{s}` ({s}): wrong value\n", .{ c.formula, c.func });
            return e;
        };
    }
}

test "M5a2: INDIRECT's R1C1 request is a refused CONSTRUCT, not a #REF! value" {
    // The whole point of the distinction: `#REF!` says "this text names
    // nothing", and `R1C1` names something perfectly well — something v1
    // refuses everywhere else, including in a written formula, where the
    // tokenizer has a refusal reason for exactly it. Answering `#REF!`
    // here would let a caller conclude the text was malformed.
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try putRefCells(&h);

    try testing.expectError(error.UnsupportedConstruct, h.eval("INDIRECT(\"R2C2\",FALSE())"));
    // Even when the spelling would have been a perfectly good A1
    // reference: it is the REQUEST that is refused, not the text.
    try testing.expectError(error.UnsupportedConstruct, h.eval("INDIRECT(\"B2\",FALSE())"));
    // …and the plane it reaches a caller through is the one every other
    // R1C1 refusal uses.
    try testing.expectEqual(
        parser.PlaneTwo.FormulaUnsupportedConstruct,
        planeTwo(error.UnsupportedConstruct),
    );
    // A written R1C1 reference refuses at the tokenizer, which is the
    // other half of "the same answer by another spelling".
    var refused = try parser.parse(testing.allocator, "R2C2", .{});
    defer refused.deinit(testing.allocator);
    try testing.expect(refused == .refused);
}

test "M5a2: a dynamic reference is captured as a dependency, which is what §5.6e reads" {
    // The property the outer loop rests on. `INDIRECT("C3")` mentions no
    // coordinate a walk over the text could find, so unless the runtime
    // read is logged, §5.6e has nothing to recondense on and the graph
    // would order a cell before something it actually depends on.
    var h: Harness = undefined;
    try h.init(testing.allocator);
    defer h.deinit();
    try putRefCells(&h);

    _ = try h.eval("INDIRECT(\"C3\")");
    try testing.expect(h.ev.deps.hasCell(.{
        .sheet = h.sheet,
        .row = try coords.Row.fromOneBased(3),
        .col = try coords.Col.fromZeroBased(2),
    }));
    // …and the static walk over the same text finds nothing, which is
    // the asymmetry the outer loop exists to close.
    const ast = try h.parse("INDIRECT(\"C3\")");
    var statics = DependencyLog.init(testing.allocator);
    defer statics.deinit();
    try staticDependencies(testing.allocator, ast, h.sheet, h.fake.evalEnv(), &statics);
    try testing.expectEqual(@as(usize, 0), statics.cells.items.len);
    try testing.expectEqual(@as(usize, 0), statics.areas.items.len);

    // An area behaves the same way, through `readRange` rather than
    // `readCell`.
    _ = try h.eval("SUM(OFFSET(A1,1,1,2,2))");
    try testing.expect(h.ev.deps.hasArea(.{
        .sheet = h.sheet,
        .range = (try coords.parseRange("B2:C3", .{})).normalized(),
    }));
}

test "M5a2: both names resolve, and each has a fixture" {
    var it = registry.inventory();
    var batch: usize = 0;
    while (it.next()) |e| {
        if (!std.mem.eql(u8, e.milestone, "M5a2")) continue;
        batch += 1;

        if (registry.lookup(e.name) == null) {
            std.debug.print("M5a2 name does not resolve: {s}\n", .{e.name});
            return error.UnregisteredBatchFunction;
        }
        var fixtures: usize = 0;
        for (ref_cases) |c| {
            if (std.mem.eql(u8, c.func, e.name)) fixtures += 1;
        }
        if (fixtures == 0) {
            std.debug.print("M5a2 name has no fixture: {s}\n", .{e.name});
            return error.UnfixturedBatchFunction;
        }
    }
    try testing.expectEqual(@as(usize, 2), batch);

    for (ref_cases) |c| {
        var found = false;
        var it2 = registry.inventory();
        while (it2.next()) |e| {
            if (std.mem.eql(u8, e.name, c.func) and std.mem.eql(u8, e.milestone, "M5a2")) found = true;
        }
        if (!found) {
            std.debug.print("fixture names a function outside M5a2: {s}\n", .{c.func});
            return error.FixtureOutsideBatch;
        }
    }
}

test "M5a2: the evidence label on every fixture is true of the committed manifests" {
    var oracle_rows: usize = 0;
    var excluded_rows: usize = 0;
    for (ref_cases) |c| {
        switch (try manifestVerdict(c.formula)) {
            .decided => {
                if (c.evidence != .oracle) {
                    std.debug.print("`{s}` is decided by a manifest but ships spec-pinned\n", .{c.formula});
                    return error.UnderstatedEvidence;
                }
                oracle_rows += 1;
            },
            .excluded => {
                if (c.evidence != .spec_pinned) {
                    std.debug.print("`{s}` claims evidence from an EXCLUDED cell\n", .{c.formula});
                    return error.ExcludedCellClaimedAsEvidence;
                }
                excluded_rows += 1;
            },
            .silent => {
                if (c.evidence != .spec_pinned) {
                    std.debug.print("`{s}` claims oracle evidence no manifest holds\n", .{c.formula});
                    return error.UnbackedOracleClaim;
                }
            },
        }
    }
    // Stated as numbers so the balance cannot drift silently: the
    // committed manifests contain no reference-producing cell at all.
    // When the parked Excel leg runs and the suite grows one, these move
    // and the row that moves them is the row that re-labels.
    try testing.expectEqual(@as(usize, 0), oracle_rows);
    try testing.expectEqual(@as(usize, 0), excluded_rows);
}
