//! The function registry — metadata framework, the frozen v1 inventory,
//! and the minimal function set M3a2 needs to exercise both
//! (`goal_formula.md` §7, §5.3a, §5.3c).
//!
//! Two things live here and they are deliberately different kinds of
//! thing.
//!
//! **The inventory** (`function_inventory_v1.tsv`) is *committed data*:
//! 175 frozen v1 function names, each tagged with the ladder row that
//! ships it. It is the authoritative count source — every F-batch PR
//! regenerates its number from this file rather than from prose, so the
//! ladder's counts and the implementation cannot drift apart without a
//! test noticing. Adding a row is a ladder change.
//!
//! **The table** (`functions`) is *code*: the functions that actually
//! evaluate today. Each entry declares five things with no defaults —
//! name, arity (with per-slot laziness), coercion classes, volatility,
//! propagation class — because a default is how a wrong answer gets
//! shipped quietly. `test "registry: the five required fields have no
//! defaults"` reads `@typeInfo` and fails if anyone adds one.
//!
//! Why laziness is metadata and not code
//! -------------------------------------
//! §5.3a fixes per form which arms evaluate. Recording that in the table
//! rather than inside each implementation means the evaluator can assert
//! the invariant that ties them together: a `.plain` form must have every
//! slot eager, and a lazy slot must belong to a form the evaluator knows
//! how to defer. Both directions are checked at comptime, so "I made IF
//! lazy" and "the table still says eager" cannot coexist.
//!
//! Why propagation is per function and never per family
//! ---------------------------------------------------
//! §5.3c is explicit and the examples are neighbours: `COUNT` ignores
//! errors found in ranges, `COUNTA` counts them, and `COUNTBLANK`
//! answers a third question entirely. A family-wide rule would get two
//! of those three wrong, so the class is a required field on every entry.
//!
//! Scope note. M3a2 built the framework and registered whichever
//! functions it needed to exercise it. **M4c closes F1a-1** — the
//! twenty names the inventory tags `M4c`, each with a fixture, each
//! pinned to the oracle where a committed manifest decides it and
//! labelled `spec_pinned` where none does. **M4d closes F1a-2**, the
//! seventeen numeric names, and two of them — `SQRT` and `RAND` —
//! were already here as M3a2's framework subjects. That row *pins*
//! them rather than adding them: they are held to the same five-field
//! check, the same fixture-per-name coverage, and the same
//! evidence labelling as the fifteen the row writes. **M4e closes
//! F1b** — the twenty-two aggregate, criteria, lookup and
//! position names — and with it the Core gate of 59, which is M4c's
//! twenty plus M4d's seventeen plus this row's twenty-two, regenerated
//! from the inventory rather than written down. Seven of the
//! twenty-two (`SUM`, `COUNT`, `COUNTA`, `COUNTBLANK`, `COUNTIF`,
//! `SUMIF`, `CHOOSE`) were the framework's borrowed subjects; M4e pins
//! them where they stand, exactly as M4d pinned `SQRT` and `RAND`.
//! Nothing in the table is now unpinned.

const std = @import("std");
const assert = std.debug.assert;

const value = @import("value.zig");
const env = @import("env.zig");
const eval = @import("eval.zig");
const criteria = @import("criteria.zig");

const Value = eval.Value;

/// What an implementation may fail with: everything the evaluator can,
/// plus the two matrix failures the call boundary converts (`EmptyMatrix`
/// → `#CALC!`, `TooManyCells` → a §9 limit).
pub const FnError = eval.EvalError || value.MatrixError;

/// What a plain implementation is handed. A thin handle on the
/// evaluator rather than a copy of its fields: every environment read
/// must go through the accessors so runtime dependency capture cannot be
/// bypassed by an implementation that reaches for `EvalEnv` directly.
pub const CallCtx = struct {
    ev: *eval.Evaluator,

    pub fn arena(self: CallCtx) std.mem.Allocator {
        return self.ev.arena;
    }

    pub fn fidelity(self: CallCtx) value.Fidelity {
        return self.ev.opts.fidelity;
    }

    pub fn rules(self: CallCtx) value.FpRules {
        return value.FpRules.of(self.ev.opts.fidelity);
    }

    /// The single volatile-draw seam. Counted, so "no draw in the dead
    /// branch" is a property of the evaluator rather than of a fixture.
    pub fn draw(self: CallCtx) f64 {
        return self.ev.opts.draws.draw();
    }

    /// Where the formula lives, for the two functions whose answer *is*
    /// their position. Optional because standalone evaluation has no
    /// stored cell (§5.3b): `ROW()` with no anchor raises
    /// `error.AnchorRequired` rather than guessing 1.
    pub fn site(self: CallCtx) ?eval.EvalSite {
        return self.ev.opts.site;
    }

    /// `collation_v1`, for the five lookup names that take their
    /// equality and their ordering from it (§5.4b). Reached through the
    /// context rather than rebuilt, because a second comparator is the
    /// failure the single-comparator rule exists to prevent.
    pub fn collation(self: CallCtx) value.Collation {
        return self.ev.opts.collation;
    }
};

pub const Impl = *const fn (ctx: CallCtx, args: []const Value) FnError!Value;

/// Whether a slot's argument is evaluated before the implementation runs.
pub const Laziness = enum {
    eager,
    /// Evaluated only if the form's contract reaches it. Laziness governs
    /// runtime evaluation and volatile draws **only** — the dependency
    /// graph still carries a static edge for the arm (§5.3a).
    lazy,
};

/// What a slot expects, and therefore what the dispatcher coerces it to
/// before the implementation sees it.
pub const CoercionClass = enum {
    /// Coerced through the §5.3b numeric column; text that parses only
    /// under a locale refuses rather than guessing.
    number,
    text,
    /// Excel does not coerce text to a condition: `IF("TRUE",…)` is
    /// `#VALUE!`.
    logical,
    /// Handed over untouched. The implementation inspects the `Value`
    /// itself — `ISBLANK` needs the reference, not its contents.
    value_any,
    /// A range, array, or scalar folded by the implementation, which
    /// applies the `via range` column of the coercion matrix itself.
    /// Never lifted elementwise.
    aggregate,
    /// Must be a reference; anything else is `#VALUE!`.
    reference,
    /// A criterion. Collapsed to a scalar like the numeric classes, but
    /// **not** coerced: `criteria.parse` classifies it, and coercing
    /// `">5"` to a number first would destroy the operator.
    criteria,
    /// A slot the form defers. The dispatcher never evaluates it.
    lazy_any,

    /// Whether a slot of this class participates in elementwise lifting
    /// over an array argument (§5.3b `array where scalar expected`).
    pub fn isScalarClass(self: CoercionClass) bool {
        return switch (self) {
            .number, .text, .logical, .criteria => true,
            .value_any, .aggregate, .reference, .lazy_any => false,
        };
    }
};

/// Arity plus per-slot laziness. `fixed` covers the leading slots;
/// `rest` cycles over everything beyond them, which is what lets `IFS`
/// declare its alternating condition/value tail as a two-slot cycle.
pub const Arity = struct {
    min: u8,
    /// `null` = bounded only by §9's `max_args`.
    max: ?u8,
    fixed: []const Laziness,
    rest: []const Laziness,

    pub fn at(self: Arity, i: usize) Laziness {
        if (i < self.fixed.len) return self.fixed[i];
        assert(self.rest.len > 0);
        return self.rest[(i - self.fixed.len) % self.rest.len];
    }
};

/// The coercion class of each slot, shaped exactly like `Arity`. The two
/// are separate fields because §7 lists them as separate metadata; a
/// comptime check below proves they can never disagree about how many
/// slots there are.
pub const Coercion = struct {
    fixed: []const CoercionClass,
    rest: []const CoercionClass,

    pub fn at(self: Coercion, i: usize) CoercionClass {
        if (i < self.fixed.len) return self.fixed[i];
        assert(self.rest.len > 0);
        return self.rest[(i - self.fixed.len) % self.rest.len];
    }
};

/// `ca` is excluded on purpose (§7): it is a cell-scheduling attribute,
/// not a property of the function.
pub const Volatility = enum {
    stable,
    /// Redrawn every recalculation; the callsite-keyed schedule is M5a2's.
    volatile_fn,
};

/// Which evaluation contract the form follows. Everything except
/// `.plain` is implemented by the evaluator, because deferring an arm
/// means holding an AST index rather than a value.
pub const Form = enum {
    plain,
    if_form,
    choose_form,
    iferror_form,
    ifna_form,
};

pub const Function = struct {
    // ── the five required fields (§7). No defaults, by design. ──
    name: []const u8,
    arity: Arity,
    coercion: Coercion,
    volatility: Volatility,
    propagation: value.PropagationClass,

    // ── the rest of §7's row, where a default is the honest answer ──
    form: Form = .plain,
    /// Present exactly when `form == .plain` (asserted at comptime).
    impl: ?Impl = null,
    /// Returns a reference rather than a value (`INDIRECT`, `OFFSET`).
    reference_producing: bool = false,
    /// Consumes arrays itself instead of being lifted over them (M7a).
    da_aware: bool = false,
    /// Depends on `collation_v1`, and therefore on §5.4b's fold.
    collation_sensitive: bool = false,
    /// Depends on the platform profile (code pages, `CHAR`/`CODE`).
    platform_sensitive: bool = false,
    /// Depends on a §5.4d compatibility version.
    cv_sensitive: bool = false,

    /// Whether an array in a scalar slot may be iterated elementwise.
    /// Mixed signatures (a range slot beside a scalar one) are M7a's
    /// decision table, not this row's.
    pub fn liftable(self: Function) bool {
        for (self.coercion.fixed) |c| if (!c.isScalarClass()) return false;
        for (self.coercion.rest) |c| if (!c.isScalarClass()) return false;
        return self.coercion.fixed.len + self.coercion.rest.len > 0;
    }
};

// ─── the table ───────────────────────────────────────────────────

const eager1 = [_]Laziness{.eager};
const eager2 = [_]Laziness{ .eager, .eager };
const lazy1 = [_]Laziness{.lazy};
const none_l = [_]Laziness{};
const none_c = [_]CoercionClass{};

// F1a-2's two signatures. Named rather than spelled inline at fifteen
// call sites, so "this slot takes a number" is one thing to read and one
// thing to get wrong.
const num1 = [_]CoercionClass{.number};
const num2 = [_]CoercionClass{ .number, .number };

pub const functions = [_]Function{
    // ── zero-argument literals and the volatile probe ──
    .{
        .name = "TRUE",
        .arity = .{ .min = 0, .max = 0, .fixed = &none_l, .rest = &none_l },
        .coercion = .{ .fixed = &none_c, .rest = &none_c },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnTrue,
    },
    .{
        .name = "FALSE",
        .arity = .{ .min = 0, .max = 0, .fixed = &none_l, .rest = &none_l },
        .coercion = .{ .fixed = &none_c, .rest = &none_c },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnFalse,
    },
    .{
        .name = "NA",
        .arity = .{ .min = 0, .max = 0, .fixed = &none_l, .rest = &none_l },
        .coercion = .{ .fixed = &none_c, .rest = &none_c },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnNa,
    },
    .{
        .name = "RAND",
        .arity = .{ .min = 0, .max = 0, .fixed = &none_l, .rest = &none_l },
        .coercion = .{ .fixed = &none_c, .rest = &none_c },
        .volatility = .volatile_fn,
        .propagation = .propagate,
        .impl = fnRand,
    },

    // ── scalar numerics: the liftable shape ──
    //
    // `SQRT` is an F1a-2 row that M3a2 wrote early, as `RAND` above is.
    // M4d pins both; neither moved, because moving a row to prove it
    // belongs to a batch would be proving it about the file rather than
    // about the table. The M4d tests read the inventory instead.
    .{
        .name = "SQRT",
        .arity = .{ .min = 1, .max = 1, .fixed = &eager1, .rest = &none_l },
        .coercion = .{ .fixed = &[_]CoercionClass{.number}, .rest = &none_c },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnSqrt,
    },
    .{
        .name = "NOT",
        .arity = .{ .min = 1, .max = 1, .fixed = &eager1, .rest = &none_l },
        .coercion = .{ .fixed = &[_]CoercionClass{.logical}, .rest = &none_c },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnNot,
    },

    // ── M4d / F1a-2: the numeric batch ──
    //
    // Fifteen rows, plus `SQRT` and `RAND` above. Every one of them is
    // `.number`-classed and `.propagate`: the dispatcher coerces, the
    // propagation pass returns on the first error, and the
    // implementation therefore starts from a finite f64 it did not have
    // to re-derive. The two facts that are *not* uniform are the two the
    // tests state name by name — which rows are volatile, and which take
    // more than one argument.
    .{
        .name = "ABS",
        .arity = .{ .min = 1, .max = 1, .fixed = &eager1, .rest = &none_l },
        .coercion = .{ .fixed = &num1, .rest = &none_c },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnAbs,
    },
    .{
        .name = "SIGN",
        .arity = .{ .min = 1, .max = 1, .fixed = &eager1, .rest = &none_l },
        .coercion = .{ .fixed = &num1, .rest = &none_c },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnSign,
    },
    .{
        .name = "INT",
        .arity = .{ .min = 1, .max = 1, .fixed = &eager1, .rest = &none_l },
        .coercion = .{ .fixed = &num1, .rest = &none_c },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnInt,
    },
    // ROUND/ROUNDUP/ROUNDDOWN take the digit count as a required second
    // argument; TRUNC's is optional and defaults to 0, which is the only
    // difference between `TRUNC(x)` and `ROUNDDOWN(x,0)`.
    .{
        .name = "ROUND",
        .arity = .{ .min = 2, .max = 2, .fixed = &eager2, .rest = &none_l },
        .coercion = .{ .fixed = &num2, .rest = &none_c },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnRound,
    },
    .{
        .name = "ROUNDUP",
        .arity = .{ .min = 2, .max = 2, .fixed = &eager2, .rest = &none_l },
        .coercion = .{ .fixed = &num2, .rest = &none_c },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnRoundUp,
    },
    .{
        .name = "ROUNDDOWN",
        .arity = .{ .min = 2, .max = 2, .fixed = &eager2, .rest = &none_l },
        .coercion = .{ .fixed = &num2, .rest = &none_c },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnRoundDown,
    },
    .{
        .name = "TRUNC",
        .arity = .{ .min = 1, .max = 2, .fixed = &eager2, .rest = &none_l },
        .coercion = .{ .fixed = &num2, .rest = &none_c },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnTrunc,
    },
    .{
        .name = "MOD",
        .arity = .{ .min = 2, .max = 2, .fixed = &eager2, .rest = &none_l },
        .coercion = .{ .fixed = &num2, .rest = &none_c },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnMod,
    },
    .{
        .name = "POWER",
        .arity = .{ .min = 2, .max = 2, .fixed = &eager2, .rest = &none_l },
        .coercion = .{ .fixed = &num2, .rest = &none_c },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnPower,
    },
    .{
        .name = "EXP",
        .arity = .{ .min = 1, .max = 1, .fixed = &eager1, .rest = &none_l },
        .coercion = .{ .fixed = &num1, .rest = &none_c },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnExp,
    },
    .{
        .name = "LN",
        .arity = .{ .min = 1, .max = 1, .fixed = &eager1, .rest = &none_l },
        .coercion = .{ .fixed = &num1, .rest = &none_c },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnLn,
    },
    .{
        .name = "LOG",
        .arity = .{ .min = 1, .max = 2, .fixed = &eager2, .rest = &none_l },
        .coercion = .{ .fixed = &num2, .rest = &none_c },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnLog,
    },
    .{
        .name = "LOG10",
        .arity = .{ .min = 1, .max = 1, .fixed = &eager1, .rest = &none_l },
        .coercion = .{ .fixed = &num1, .rest = &none_c },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnLog10,
    },
    .{
        .name = "PI",
        .arity = .{ .min = 0, .max = 0, .fixed = &none_l, .rest = &none_l },
        .coercion = .{ .fixed = &none_c, .rest = &none_c },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnPi,
    },
    // The batch's second volatile, and the only one M4d adds. It draws
    // exactly once per call from the same counted seam `RAND` uses —
    // see `fnRandBetween` for why a rejection sampler would have been
    // the wrong instrument here.
    .{
        .name = "RANDBETWEEN",
        .arity = .{ .min = 2, .max = 2, .fixed = &eager2, .rest = &none_l },
        .coercion = .{ .fixed = &num2, .rest = &none_c },
        .volatility = .volatile_fn,
        .propagation = .propagate,
        .impl = fnRandBetween,
    },

    // ── the `observe` class: looking at an error without becoming one ──
    .{
        .name = "ISERROR",
        .arity = .{ .min = 1, .max = 1, .fixed = &eager1, .rest = &none_l },
        .coercion = .{ .fixed = &[_]CoercionClass{.value_any}, .rest = &none_c },
        .volatility = .stable,
        .propagation = .observe,
        .impl = fnIsError,
    },
    .{
        .name = "ISNUMBER",
        .arity = .{ .min = 1, .max = 1, .fixed = &eager1, .rest = &none_l },
        .coercion = .{ .fixed = &[_]CoercionClass{.value_any}, .rest = &none_c },
        .volatility = .stable,
        .propagation = .observe,
        .impl = fnIsNumber,
    },
    .{
        .name = "ISBLANK",
        .arity = .{ .min = 1, .max = 1, .fixed = &eager1, .rest = &none_l },
        .coercion = .{ .fixed = &[_]CoercionClass{.value_any}, .rest = &none_c },
        .volatility = .stable,
        .propagation = .observe,
        .impl = fnIsBlank,
    },
    .{
        .name = "ISERR",
        .arity = .{ .min = 1, .max = 1, .fixed = &eager1, .rest = &none_l },
        .coercion = .{ .fixed = &[_]CoercionClass{.value_any}, .rest = &none_c },
        .volatility = .stable,
        .propagation = .observe,
        .impl = fnIsErr,
    },
    .{
        .name = "ISNA",
        .arity = .{ .min = 1, .max = 1, .fixed = &eager1, .rest = &none_l },
        .coercion = .{ .fixed = &[_]CoercionClass{.value_any}, .rest = &none_c },
        .volatility = .stable,
        .propagation = .observe,
        .impl = fnIsNa,
    },
    .{
        .name = "ISLOGICAL",
        .arity = .{ .min = 1, .max = 1, .fixed = &eager1, .rest = &none_l },
        .coercion = .{ .fixed = &[_]CoercionClass{.value_any}, .rest = &none_c },
        .volatility = .stable,
        .propagation = .observe,
        .impl = fnIsLogical,
    },
    .{
        .name = "ISTEXT",
        .arity = .{ .min = 1, .max = 1, .fixed = &eager1, .rest = &none_l },
        .coercion = .{ .fixed = &[_]CoercionClass{.value_any}, .rest = &none_c },
        .volatility = .stable,
        .propagation = .observe,
        .impl = fnIsText,
    },

    // ── N and T: they inspect the value like the IS-family and then
    //    *become* something else, so they read `.value_any` and
    //    `.propagate`. Not `.number`/`.text`: those classes coerce, and
    //    Excel's `N("7")` is 0 where the numeric class would answer 7. ──
    .{
        .name = "N",
        .arity = .{ .min = 1, .max = 1, .fixed = &eager1, .rest = &none_l },
        .coercion = .{ .fixed = &[_]CoercionClass{.value_any}, .rest = &none_c },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnN,
    },
    .{
        .name = "T",
        .arity = .{ .min = 1, .max = 1, .fixed = &eager1, .rest = &none_l },
        .coercion = .{ .fixed = &[_]CoercionClass{.value_any}, .rest = &none_c },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnT,
    },

    // ── eager logical folds. Excel does NOT short-circuit these. ──
    .{
        .name = "AND",
        .arity = .{ .min = 1, .max = null, .fixed = &none_l, .rest = &eager1 },
        .coercion = .{ .fixed = &none_c, .rest = &[_]CoercionClass{.aggregate} },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnAnd,
    },
    .{
        .name = "OR",
        .arity = .{ .min = 1, .max = null, .fixed = &none_l, .rest = &eager1 },
        .coercion = .{ .fixed = &none_c, .rest = &[_]CoercionClass{.aggregate} },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnOr,
    },
    .{
        .name = "IFS",
        .arity = .{ .min = 2, .max = null, .fixed = &none_l, .rest = &eager2 },
        .coercion = .{
            .fixed = &none_c,
            .rest = &[_]CoercionClass{ .value_any, .value_any },
        },
        .volatility = .stable,
        .propagation = .observe,
        .impl = fnIfs,
    },
    .{
        .name = "SWITCH",
        .arity = .{ .min = 3, .max = null, .fixed = &eager1, .rest = &eager1 },
        .coercion = .{
            .fixed = &[_]CoercionClass{.value_any},
            .rest = &[_]CoercionClass{.value_any},
        },
        .volatility = .stable,
        .propagation = .observe,
        .collation_sensitive = true,
        .impl = fnSwitch,
    },

    // ── aggregates: the three-way COUNT split §5.3c names by hand ──
    .{
        .name = "SUM",
        .arity = .{ .min = 1, .max = null, .fixed = &none_l, .rest = &eager1 },
        .coercion = .{ .fixed = &none_c, .rest = &[_]CoercionClass{.aggregate} },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnSum,
    },
    .{
        .name = "COUNT",
        .arity = .{ .min = 1, .max = null, .fixed = &none_l, .rest = &eager1 },
        .coercion = .{ .fixed = &none_c, .rest = &[_]CoercionClass{.aggregate} },
        .volatility = .stable,
        .propagation = .per_function_provenance,
        .impl = fnCount,
    },
    .{
        .name = "COUNTA",
        .arity = .{ .min = 1, .max = null, .fixed = &none_l, .rest = &eager1 },
        .coercion = .{ .fixed = &none_c, .rest = &[_]CoercionClass{.aggregate} },
        .volatility = .stable,
        .propagation = .per_function_provenance,
        .impl = fnCountA,
    },
    .{
        .name = "COUNTBLANK",
        .arity = .{ .min = 1, .max = 1, .fixed = &eager1, .rest = &none_l },
        .coercion = .{ .fixed = &[_]CoercionClass{.reference}, .rest = &none_c },
        .volatility = .stable,
        .propagation = .per_function_provenance,
        .impl = fnCountBlank,
    },
    .{
        .name = "AVERAGE",
        .arity = .{ .min = 1, .max = null, .fixed = &none_l, .rest = &eager1 },
        .coercion = .{ .fixed = &none_c, .rest = &[_]CoercionClass{.aggregate} },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnAverage,
    },
    // MIN and MAX are **not** `collation_sensitive`, and that is a
    // decision rather than an omission: §5.4b names them as the
    // comparator's one exception, because they never compare text at
    // all — a direct text argument coerces or is `#VALUE!`, text found
    // in a range is ignored, and no numbers anywhere is 0.
    .{
        .name = "MIN",
        .arity = .{ .min = 1, .max = null, .fixed = &none_l, .rest = &eager1 },
        .coercion = .{ .fixed = &none_c, .rest = &[_]CoercionClass{.aggregate} },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnMin,
    },
    .{
        .name = "MAX",
        .arity = .{ .min = 1, .max = null, .fixed = &none_l, .rest = &eager1 },
        .coercion = .{ .fixed = &none_c, .rest = &[_]CoercionClass{.aggregate} },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnMax,
    },
    .{
        .name = "SUMPRODUCT",
        .arity = .{ .min = 1, .max = null, .fixed = &none_l, .rest = &eager1 },
        .coercion = .{ .fixed = &none_c, .rest = &[_]CoercionClass{.aggregate} },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnSumProduct,
    },

    // ── criteria: the two shapes §5.6a's alignment rule distinguishes ──
    .{
        .name = "COUNTIF",
        .arity = .{ .min = 2, .max = 2, .fixed = &eager2, .rest = &none_l },
        .coercion = .{
            .fixed = &[_]CoercionClass{ .reference, .criteria },
            .rest = &none_c,
        },
        .volatility = .stable,
        .propagation = .per_function_provenance,
        .collation_sensitive = true,
        .impl = fnCountIf,
    },
    .{
        .name = "SUMIF",
        .arity = .{ .min = 2, .max = 3, .fixed = &[_]Laziness{ .eager, .eager, .eager }, .rest = &none_l },
        .coercion = .{
            .fixed = &[_]CoercionClass{ .reference, .criteria, .reference },
            .rest = &none_c,
        },
        .volatility = .stable,
        .propagation = .per_function_provenance,
        .collation_sensitive = true,
        .impl = fnSumIf,
    },
    .{
        .name = "AVERAGEIF",
        .arity = .{ .min = 2, .max = 3, .fixed = &[_]Laziness{ .eager, .eager, .eager }, .rest = &none_l },
        .coercion = .{
            .fixed = &[_]CoercionClass{ .reference, .criteria, .reference },
            .rest = &none_c,
        },
        .volatility = .stable,
        .propagation = .per_function_provenance,
        .collation_sensitive = true,
        .impl = fnAverageIf,
    },

    // ── lookups: the five names that take BOTH their equality and
    //    their ordering from `collation_v1` (§5.4b), which is exactly
    //    what MIN and MAX above do not.
    //
    //    All five are `per_function_provenance` rather than
    //    `propagate`, for a reason §5.3c states and this is the first
    //    batch to meet: an error found INSIDE a lookup table is not the
    //    lookup's error — it is a value the lookup may or may not
    //    return. The dispatcher's first-error scan cannot draw that
    //    line (it sees evaluated scalars, and a table arrives as a
    //    reference), and it cannot see an error in the KEY slot either,
    //    which is `.value_any` and therefore also often a reference. So
    //    §5.3c's declaration order is taken by `lookupPropagate` below,
    //    where both halves are one rule. ──
    .{
        .name = "VLOOKUP",
        .arity = .{
            .min = 3,
            .max = 4,
            .fixed = &[_]Laziness{ .eager, .eager, .eager, .eager },
            .rest = &none_l,
        },
        .coercion = .{
            .fixed = &[_]CoercionClass{ .value_any, .aggregate, .number, .logical },
            .rest = &none_c,
        },
        .volatility = .stable,
        .propagation = .per_function_provenance,
        .collation_sensitive = true,
        .impl = fnVLookup,
    },
    .{
        .name = "HLOOKUP",
        .arity = .{
            .min = 3,
            .max = 4,
            .fixed = &[_]Laziness{ .eager, .eager, .eager, .eager },
            .rest = &none_l,
        },
        .coercion = .{
            .fixed = &[_]CoercionClass{ .value_any, .aggregate, .number, .logical },
            .rest = &none_c,
        },
        .volatility = .stable,
        .propagation = .per_function_provenance,
        .collation_sensitive = true,
        .impl = fnHLookup,
    },
    .{
        .name = "MATCH",
        .arity = .{ .min = 2, .max = 3, .fixed = &[_]Laziness{ .eager, .eager, .eager }, .rest = &none_l },
        .coercion = .{
            .fixed = &[_]CoercionClass{ .value_any, .aggregate, .number },
            .rest = &none_c,
        },
        .volatility = .stable,
        .propagation = .per_function_provenance,
        .collation_sensitive = true,
        .impl = fnMatch,
    },
    .{
        .name = "XLOOKUP",
        .arity = .{
            .min = 3,
            .max = 6,
            .fixed = &[_]Laziness{ .eager, .eager, .eager, .eager, .eager, .eager },
            .rest = &none_l,
        },
        .coercion = .{
            .fixed = &[_]CoercionClass{
                .value_any, .aggregate, .aggregate,
                .value_any, .number,    .number,
            },
            .rest = &none_c,
        },
        .volatility = .stable,
        .propagation = .per_function_provenance,
        .collation_sensitive = true,
        .impl = fnXLookup,
    },
    .{
        .name = "XMATCH",
        .arity = .{
            .min = 2,
            .max = 4,
            .fixed = &[_]Laziness{ .eager, .eager, .eager, .eager },
            .rest = &none_l,
        },
        .coercion = .{
            .fixed = &[_]CoercionClass{ .value_any, .aggregate, .number, .number },
            .rest = &none_c,
        },
        .volatility = .stable,
        .propagation = .per_function_provenance,
        .collation_sensitive = true,
        .impl = fnXMatch,
    },
    .{
        .name = "INDEX",
        .arity = .{ .min = 2, .max = 3, .fixed = &[_]Laziness{ .eager, .eager, .eager }, .rest = &none_l },
        .coercion = .{
            .fixed = &[_]CoercionClass{ .aggregate, .number, .number },
            .rest = &none_c,
        },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnIndex,
    },

    // ── position: what a reference IS, rather than what it holds. The
    //    zero-argument forms are the only site-dependent rows in the
    //    registry, which is why `CallCtx.site` exists. ──
    .{
        .name = "ROW",
        .arity = .{ .min = 0, .max = 1, .fixed = &eager1, .rest = &none_l },
        .coercion = .{ .fixed = &[_]CoercionClass{.reference}, .rest = &none_c },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnRow,
    },
    .{
        .name = "COLUMN",
        .arity = .{ .min = 0, .max = 1, .fixed = &eager1, .rest = &none_l },
        .coercion = .{ .fixed = &[_]CoercionClass{.reference}, .rest = &none_c },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnColumn,
    },
    .{
        .name = "ROWS",
        .arity = .{ .min = 1, .max = 1, .fixed = &eager1, .rest = &none_l },
        .coercion = .{ .fixed = &[_]CoercionClass{.aggregate}, .rest = &none_c },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnRows,
    },
    .{
        .name = "COLUMNS",
        .arity = .{ .min = 1, .max = 1, .fixed = &eager1, .rest = &none_l },
        .coercion = .{ .fixed = &[_]CoercionClass{.aggregate}, .rest = &none_c },
        .volatility = .stable,
        .propagation = .propagate,
        .impl = fnColumns,
    },

    // ── the lazy forms. No `impl`: deferring an arm means holding an
    //    AST index, which only the evaluator has. ──
    .{
        .name = "IF",
        .arity = .{ .min = 2, .max = 3, .fixed = &[_]Laziness{ .eager, .lazy, .lazy }, .rest = &none_l },
        .coercion = .{
            .fixed = &[_]CoercionClass{ .value_any, .lazy_any, .lazy_any },
            .rest = &none_c,
        },
        .volatility = .stable,
        .propagation = .observe,
        .form = .if_form,
    },
    .{
        .name = "CHOOSE",
        .arity = .{ .min = 2, .max = null, .fixed = &eager1, .rest = &lazy1 },
        .coercion = .{
            .fixed = &[_]CoercionClass{.value_any},
            .rest = &[_]CoercionClass{.lazy_any},
        },
        .volatility = .stable,
        .propagation = .observe,
        .form = .choose_form,
    },
    .{
        .name = "IFERROR",
        .arity = .{ .min = 2, .max = 2, .fixed = &[_]Laziness{ .eager, .lazy }, .rest = &none_l },
        .coercion = .{
            .fixed = &[_]CoercionClass{ .value_any, .lazy_any },
            .rest = &none_c,
        },
        .volatility = .stable,
        .propagation = .observe,
        .form = .iferror_form,
    },
    .{
        .name = "IFNA",
        .arity = .{ .min = 2, .max = 2, .fixed = &[_]Laziness{ .eager, .lazy }, .rest = &none_l },
        .coercion = .{
            .fixed = &[_]CoercionClass{ .value_any, .lazy_any },
            .rest = &none_c,
        },
        .volatility = .stable,
        .propagation = .observe,
        .form = .ifna_form,
    },
};

/// §5.9 call-position resolution: case-folded over the decoded symbol
/// layer. Function names are ASCII, so the fold is `std.ascii`'s and
/// needs no allocation.
pub fn lookup(name: []const u8) ?*const Function {
    for (&functions) |*f| {
        if (std.ascii.eqlIgnoreCase(f.name, name)) return f;
    }
    return null;
}

comptime {
    for (functions) |f| {
        // The two per-slot tables describe the same slots. Parallel
        // arrays are only safe when they cannot disagree.
        if (f.arity.fixed.len != f.coercion.fixed.len) {
            @compileError(f.name ++ ": arity.fixed and coercion.fixed disagree on slot count");
        }
        if (f.arity.rest.len != f.coercion.rest.len) {
            @compileError(f.name ++ ": arity.rest and coercion.rest disagree on slot count");
        }
        if (f.arity.max == null and f.arity.rest.len == 0) {
            @compileError(f.name ++ ": unbounded arity with no repeating slot");
        }
        // A form the evaluator does not defer must not claim a lazy
        // slot, and a form it does defer must not carry an impl the
        // dispatcher would call eagerly.
        const has_lazy = blk: {
            for (f.arity.fixed) |l| if (l == .lazy) break :blk true;
            for (f.arity.rest) |l| if (l == .lazy) break :blk true;
            break :blk false;
        };
        if (f.form == .plain) {
            if (has_lazy) @compileError(f.name ++ ": plain form with a lazy slot");
            if (f.impl == null) @compileError(f.name ++ ": plain form without an impl");
        } else {
            if (!has_lazy) @compileError(f.name ++ ": deferring form with no lazy slot");
            if (f.impl != null) @compileError(f.name ++ ": deferring form must not carry an impl");
        }
    }
}

// ─── implementations ─────────────────────────────────────────────

fn fnTrue(ctx: CallCtx, args: []const Value) FnError!Value {
    _ = ctx;
    _ = args;
    return Value.boolean(true);
}

fn fnFalse(ctx: CallCtx, args: []const Value) FnError!Value {
    _ = ctx;
    _ = args;
    return Value.boolean(false);
}

fn fnNa(ctx: CallCtx, args: []const Value) FnError!Value {
    _ = ctx;
    _ = args;
    return Value.err(.na);
}

fn fnRand(ctx: CallCtx, args: []const Value) FnError!Value {
    _ = args;
    return Value.num(ctx.draw());
}

fn fnSqrt(ctx: CallCtx, args: []const Value) FnError!Value {
    _ = ctx;
    // The `.number` class already coerced, and `.propagate` already
    // returned on an error, so this is a number.
    const n = args[0].scalar.number;
    // Excel's answer for a negative radicand is `#NUM!`, not `#VALUE!`
    // and not a NaN — N4a has no room for one.
    if (n < 0) return Value.err(.num);
    return Value.num(@sqrt(n));
}

// ── M4d / F1a-2: the numeric implementations ──
//
// Three rules hold across every one of them, so they are stated here
// once instead of fifteen times.
//
// **A number leaves through `fromArithmetic`, never `fromNumber`.** N4a
// says a non-finite result is not a value; `fromNumber` *asserts*
// finiteness, so `EXP(1000)` through it is a panic, and through
// `fromArithmetic` it is `#NUM!` — Excel's own answer, and the same one
// under both rule tables because §5.4 gives the two modes one value
// domain.
//
// **The fidelity mode enters in exactly one place.** N2's zero-snap is
// additive-scope only, so it never applies here. N3's signed-zero policy
// applies at *publication*, so a `-0` a truncation produces travels
// intact and `value.publish` normalizes it under `.excel`. What the mode
// does decide inside this batch is which decimal value is being operated
// on — see `decimalView`.
//
// **The dispatcher has already coerced and propagated.** A `.number`
// slot arrives as a finite `.number` scalar or the implementation was
// never reached, so these functions assert that instead of re-deriving
// it (§5.3b's coercion is the table's job, §5.3c's first-error-wins is
// `propagateAndInvoke`'s).

/// Read a `.number`-classed slot. The assertions are the contract with
/// the table above stated where it is used: change a slot's class and
/// this fires immediately rather than producing a quietly wrong number.
fn numArg(args: []const Value, i: usize) f64 {
    const s = args[i].scalar;
    assert(s == .number);
    assert(std.math.isFinite(s.number));
    return s.number;
}

/// The one way a computed number leaves this batch.
fn arith(x: f64) Value {
    return .{ .scalar = value.ScalarValue.fromArithmetic(x) };
}

/// The value a fidelity mode considers `n` to *be*.
///
/// `.excel` reads the 15 significant digits N1a keeps, `.ieee` reads the
/// binary64 itself — `literal_significant_digits` is `null` there
/// precisely so this is a no-op. It is why `ROUND(2.675, 2)` is `2.68`
/// in one mode and `2.67` in the other: 2.675 is really
/// 2.67499999999999982, and only one of the two modes is looking at that.
fn decimalView(rules: value.FpRules, n: f64) f64 {
    const sig = rules.literal_significant_digits orelse return n;
    return value.roundToSignificantDigits(n, sig);
}

const RoundMode = enum {
    /// ROUND — half away from zero.
    half_away,
    /// ROUNDUP — away from zero.
    away,
    /// ROUNDDOWN and TRUNC — toward zero.
    toward,
};

/// `x · 10^d`, split so no *intermediate* leaves binary64's range even
/// where `10^d` alone would. `roundAt`'s guards bound the result; this
/// keeps the journey to it representable too.
fn scaleByPowerOfTen(x: f64, d: f64) f64 {
    assert(@abs(d) < 400); // `roundAt` bounds `d` well inside this
    var acc = x;
    var left = d;
    while (left > 300) : (left -= 300) acc *= 1e300;
    while (left < -300) : (left += 300) acc *= 1e-300;
    return acc * std.math.pow(f64, 10, left);
}

/// The shared body of ROUND / ROUNDUP / ROUNDDOWN / TRUNC.
///
/// Excel rounds at a **decimal** place and binary64 has none, so the
/// implementation is scale → round → unscale. Two things make that more
/// than three lines.
///
/// **The scale factor need not be representable.** `10^d` overflows
/// above `d = 308` and underflows below `d = -324`, while `d` itself
/// arrives as an arbitrary f64 the caller typed. Both extremes are
/// decided *before* any scaling, by comparing the requested place with
/// the value's own decimal exponent: a place below the last significant
/// digit cannot change the value, and a place well above the leading one
/// removes all of them. What is left in between is bounded —
/// `|n·10^d| < 10^17` by construction — so the scaling never has to
/// represent a quantity binary64 cannot hold.
///
/// **The modes disagree about half-way cases**, through `decimalView`
/// and nowhere else.
fn roundAt(rules: value.FpRules, n: f64, digits: f64, mode: RoundMode) value.ScalarValue {
    assert(std.math.isFinite(n) and std.math.isFinite(digits));
    // Zero has no significant digit to round at any place. Returning `n`
    // rather than a literal 0 keeps a `-0` argument's sign, which is a
    // value under `ieee_fp_rules_v1`.
    if (n == 0) return value.ScalarValue.fromNumber(n);

    // Excel truncates the digit count toward zero rather than rounding
    // it: `ROUND(x, 2.9)` is `ROUND(x, 2)`.
    const d = std.math.trunc(digits);
    // The decimal exponent of the leading significant digit.
    const e = @floor(@log10(@abs(n)));

    // Below the last significant digit: binary64 carries at most 17 of
    // them, so there is nothing left down there to round away.
    if (d >= 17 - e) return value.ScalarValue.fromNumber(n);

    // Two or more decades above the leading digit: `|n| ≤ 10^(e+1)` and
    // the place is `10^(-d) ≥ 10^(e+2)`, so the scaled magnitude is at
    // most 0.1 and all three modes answer without scaling anything.
    if (d + e <= -2) {
        return switch (mode) {
            // `copysign` rather than a bare zero: `ROUNDDOWN(-1, -100)`
            // is `-0` before publication — preserved by `.ieee`,
            // normalized by `.excel`.
            .half_away, .toward => value.ScalarValue.fromNumber(std.math.copysign(@as(f64, 0), n)),
            // Away from zero at a place that large overflows, and N4a
            // has an answer for that. Guarded rather than computed,
            // because `10^(-d)` is the one power of ten here that a
            // caller can push past the representable range.
            .away => if (-d > 308)
                value.ScalarValue.errorOf(.num)
            else
                value.ScalarValue.fromArithmetic(
                    std.math.copysign(std.math.pow(f64, 10, -d), n),
                ),
        };
    }

    const scaled = decimalView(rules, scaleByPowerOfTen(n, d));
    const placed = switch (mode) {
        // Zig's `@round` is half-away-from-zero, which is Excel's rule.
        .half_away => @round(scaled),
        .away => std.math.copysign(@ceil(@abs(scaled)), scaled),
        .toward => std.math.trunc(scaled),
    };
    return value.ScalarValue.fromArithmetic(scaleByPowerOfTen(placed, -d));
}

fn fnAbs(ctx: CallCtx, args: []const Value) FnError!Value {
    _ = ctx;
    // `@abs(-0.0)` is `+0.0`, which is the answer in both modes: the
    // absolute value of a signed zero has no sign left to preserve.
    return arith(@abs(numArg(args, 0)));
}

fn fnSign(ctx: CallCtx, args: []const Value) FnError!Value {
    _ = ctx;
    const n = numArg(args, 0);
    // `-0 == 0` is true, so a negative zero answers 0: SIGN reports the
    // sign of a quantity, and the quantity is zero.
    if (n > 0) return Value.num(1);
    if (n < 0) return Value.num(-1);
    return Value.num(0);
}

fn fnInt(ctx: CallCtx, args: []const Value) FnError!Value {
    // Floor, not truncation: `INT(-2.5)` is `-3` where `TRUNC(-2.5)` is
    // `-2`. The two functions existing separately is the reason to say
    // which is which.
    return arith(@floor(decimalView(ctx.rules(), numArg(args, 0))));
}

fn fnRound(ctx: CallCtx, args: []const Value) FnError!Value {
    return .{ .scalar = roundAt(ctx.rules(), numArg(args, 0), numArg(args, 1), .half_away) };
}

fn fnRoundUp(ctx: CallCtx, args: []const Value) FnError!Value {
    return .{ .scalar = roundAt(ctx.rules(), numArg(args, 0), numArg(args, 1), .away) };
}

fn fnRoundDown(ctx: CallCtx, args: []const Value) FnError!Value {
    return .{ .scalar = roundAt(ctx.rules(), numArg(args, 0), numArg(args, 1), .toward) };
}

fn fnTrunc(ctx: CallCtx, args: []const Value) FnError!Value {
    // The digit count is optional and 0 when omitted, which is the only
    // thing separating `TRUNC(x)` from `ROUNDDOWN(x, 0)` — a fixture
    // states that equivalence rather than leaving it to be inferred.
    const digits = if (args.len >= 2) numArg(args, 1) else 0;
    return .{ .scalar = roundAt(ctx.rules(), numArg(args, 0), digits, .toward) };
}

fn fnMod(ctx: CallCtx, args: []const Value) FnError!Value {
    _ = ctx;
    const n = numArg(args, 0);
    const d = numArg(args, 1);
    // Excel's answer for a zero divisor, and the reason this is not
    // `@mod`: that builtin has no defined result here.
    if (d == 0) return Value.err(.div0);
    // §5.4's N4 names MOD's **sign** rule as a per-function quirk, and
    // names only that: the result takes the DIVISOR's sign, so
    // `MOD(-5,3)` is 1 and `MOD(5,-3)` is -1. That is floored modulus,
    // written out rather than delegated so the overflow stays visible —
    // an extreme ratio sends `d · floor(n/d)` to infinity, and N4a
    // answers `#NUM!`. The quotient is deliberately NOT read through
    // `decimalView`: N4 scopes MOD's quirk to the sign, and widening it
    // here would be inventing an Excel behaviour no manifest recorded.
    return arith(n - d * @floor(n / d));
}

fn fnPower(ctx: CallCtx, args: []const Value) FnError!Value {
    _ = ctx;
    // POWER is the function spelling of `^` and runs the same
    // arithmetic: `applyBinaryScalar`'s `.pow` arm is this same
    // `std.math.pow` through this same `fromArithmetic`. A workbook must
    // not get two answers for one operation, so the identity is a
    // fixture rather than a comment.
    return arith(std.math.pow(f64, numArg(args, 0), numArg(args, 1)));
}

fn fnExp(ctx: CallCtx, args: []const Value) FnError!Value {
    _ = ctx;
    // `EXP(1000)` overflows, and `#NUM!` is reached through
    // `fromArithmetic` rather than through a magnitude test — the
    // boundary is the representation's, not a guessed one.
    return arith(@exp(numArg(args, 0)));
}

fn fnLn(ctx: CallCtx, args: []const Value) FnError!Value {
    _ = ctx;
    const n = numArg(args, 0);
    // The whole non-positive half is out of domain, zero included:
    // `LN(0)` is `#NUM!` and not `-∞`, because N4a has no room for one.
    if (n <= 0) return Value.err(.num);
    return arith(@log(n));
}

fn fnLog10(ctx: CallCtx, args: []const Value) FnError!Value {
    _ = ctx;
    const n = numArg(args, 0);
    if (n <= 0) return Value.err(.num);
    return arith(@log10(n));
}

fn fnLog(ctx: CallCtx, args: []const Value) FnError!Value {
    _ = ctx;
    const n = numArg(args, 0);
    if (n <= 0) return Value.err(.num);
    // The base is optional and 10 when omitted, which makes `LOG(x)` and
    // `LOG10(x)` one operation under two spellings.
    const base = if (args.len >= 2) numArg(args, 1) else 10;
    if (base <= 0) return Value.err(.num);
    // Base 1 divides by `LN(1)`, and that failure is spelled `#DIV/0!`
    // rather than `#NUM!` — the one place this family answers with
    // something other than a domain error, which is why it is its own
    // line and its own fixture.
    if (base == 1) return Value.err(.div0);
    // The two bases with an exact primitive get it: `LOG(8,2)` is 3 and
    // `LOG(100,10)` is 2 exactly, where `@log(8)/@log(2)` is under no
    // obligation to be.
    if (base == 10) return arith(@log10(n));
    if (base == 2) return arith(@log2(n));
    return arith(@log(n) / @log(base));
}

fn fnPi(ctx: CallCtx, args: []const Value) FnError!Value {
    _ = ctx;
    _ = args;
    // The binary64 nearest π. Excel documents 15 digits and stores this
    // same double, so there is nothing to round.
    return Value.num(std.math.pi);
}

fn fnRandBetween(ctx: CallCtx, args: []const Value) FnError!Value {
    const bottom = numArg(args, 0);
    const top = numArg(args, 1);
    // Non-integer bounds move INWARD — ceil the bottom, floor the top.
    // The alternative reading, truncating both toward zero, lets
    // `RANDBETWEEN(1.5, 3.5)` answer 1, a value outside the interval the
    // caller wrote. No committed manifest decides this, so the fixture
    // ships `spec_pinned` and pins the invariant a reader can check
    // instead: every result lies within `[bottom, top]`.
    const lo = @ceil(bottom);
    const hi = @floor(top);
    // Excel's answer for an empty range, including a range that is only
    // empty after the bounds moved in.
    if (lo > hi) return Value.err(.num);

    // ONE draw, always. A rejection sampler would be exactly uniform and
    // would draw a data-dependent number of times — which would make the
    // draw counter, the instrument every §5.6d KAT is built on, unable
    // to state anything. Scaling a single draw is off perfect uniformity
    // by at most one part in 2^53.
    const span = hi - lo + 1;
    if (!std.math.isFinite(span)) return Value.err(.num);
    const u = ctx.draw();
    assert(u >= 0 and u < 1);
    // `@min` covers the one case scaling cannot: a draw just under 1
    // against a span large enough for the product to round up to it.
    return arith(@min(hi, lo + @floor(u * span)));
}

fn fnNot(ctx: CallCtx, args: []const Value) FnError!Value {
    _ = ctx;
    return Value.boolean(!args[0].scalar.boolean);
}

fn fnIsError(ctx: CallCtx, args: []const Value) FnError!Value {
    return Value.boolean(try observedScalar(ctx, args[0]) == .err);
}

fn fnIsNumber(ctx: CallCtx, args: []const Value) FnError!Value {
    return Value.boolean(try observedScalar(ctx, args[0]) == .number);
}

/// Whether an error value is `#N/A` — asked by **spelling**, not by
/// enum tag. `ErrorValue` has two arms and only one of them carries a
/// `KnownError`; matching on the tag alone would answer FALSE for an
/// `#N/A` that reached us through the extensible-literal rule as a rich
/// spelling. `ISNA` is a question about the error a user sees, and what
/// they see is the spelling.
fn isNaError(s: value.ScalarValue) bool {
    if (s != .err) return false;
    return std.mem.eql(u8, s.err.spelling(), value.KnownError.na.spelling());
}

fn fnIsErr(ctx: CallCtx, args: []const Value) FnError!Value {
    const s = try observedScalar(ctx, args[0]);
    // Every error except `#N/A` — the one distinction that makes `ISERR`
    // a separate function from `ISERROR` rather than a synonym.
    return Value.boolean(s == .err and !isNaError(s));
}

fn fnIsNa(ctx: CallCtx, args: []const Value) FnError!Value {
    return Value.boolean(isNaError(try observedScalar(ctx, args[0])));
}

fn fnIsLogical(ctx: CallCtx, args: []const Value) FnError!Value {
    return Value.boolean(try observedScalar(ctx, args[0]) == .boolean);
}

fn fnIsText(ctx: CallCtx, args: []const Value) FnError!Value {
    // `""` is text and a blank cell is not — the same three-way split
    // `COUNTA`, `COUNTBLANK`, and `ISBLANK` are built on.
    return Value.boolean(try observedScalar(ctx, args[0]) == .text);
}

/// `N(value)` — Excel's own conversion table, which is **not** the
/// `.number` coercion class: a number is itself, `TRUE`/`FALSE` are 1
/// and 0, an error is the error, and *everything else* — text numeric
/// or not, and blank — is 0. `N("7")` is 0; the numeric class would say
/// 7, and that is the whole reason this slot is `.value_any`.
fn fnN(ctx: CallCtx, args: []const Value) FnError!Value {
    const s = try observedScalar(ctx, args[0]);
    return switch (s) {
        .number => |n| Value.num(n),
        .boolean => |b| Value.num(if (b) 1 else 0),
        .err => .{ .scalar = s },
        .text, .blank => Value.num(0),
    };
}

/// `T(value)` — the mirror of `N`: text is itself, an error is the
/// error, everything else is `""`. A number does NOT format; `T(1)` is
/// `""`, not `"1"`.
fn fnT(ctx: CallCtx, args: []const Value) FnError!Value {
    const s = try observedScalar(ctx, args[0]);
    return switch (s) {
        .text => .{ .scalar = s },
        .err => .{ .scalar = s },
        .number, .boolean, .blank => .{ .scalar = .{ .text = "" } },
    };
}

fn fnIsBlank(ctx: CallCtx, args: []const Value) FnError!Value {
    switch (args[0]) {
        .reference => |r| {
            const area = r.single() orelse return Value.err(.value);
            if (!area.isSingleCell()) return Value.err(.value);
            // Through the class, not through `cellValue`: `ISBLANK` and
            // `COUNTBLANK` must answer from the same merged view, and
            // the class is what makes them differ about `=""`.
            const blanks = try ctx.ev.readBlankCount(area, .isblank_class);
            return Value.boolean(blanks == 1);
        },
        .missing_arg => return Value.boolean(true),
        .scalar => |s| return Value.boolean(s == .blank),
        .array => |m| return Value.boolean(m.topLeft() == .blank),
    }
}

/// What an `observe`-class predicate sees. A reference dereferences; an
/// array reduces to its top-left, which is the §5.3b row for an array in
/// a scalar context.
fn observedScalar(ctx: CallCtx, v: Value) FnError!value.ScalarValue {
    return switch (v) {
        .scalar => |s| s,
        .missing_arg => .blank,
        .array => |m| m.topLeft(),
        .reference => |r| blk: {
            const area = r.single() orelse break :blk value.ScalarValue.errorOf(.value);
            if (!area.isSingleCell()) break :blk value.ScalarValue.errorOf(.value);
            break :blk try ctx.ev.readCell(area.topLeft());
        },
    };
}

// ── aggregate walking ──

/// Visit every scalar an argument list contributes, telling the visitor
/// whether each arrived directly or through a range/array — the
/// distinction the `via range` column of §5.3b turns on, and the reason
/// `SUM(TRUE)` is 1 while `SUM(A1)` with `A1=TRUE` is 0.
///
/// Unoccupied cells are not visited: every aggregate class in the matrix
/// skips or ignores a blank found in a range, so materialising 1 048 576
/// of them would change nothing but the running time.
fn visitArgs(ctx: CallCtx, args: []const Value, acc: anytype) FnError!void {
    for (args) |a| switch (a) {
        .missing_arg => {
            if (!try acc.visit(.blank, false)) return;
        },
        .scalar => |s| {
            if (!try acc.visit(s, false)) return;
        },
        .array => |m| for (m.cells) |s| {
            if (!try acc.visit(s, true)) return;
        },
        .reference => |r| for (r.areas) |area| {
            var it = try ctx.ev.readRange(area);
            while (try it.next()) |e| {
                if (!try acc.visit(e.value, true)) return;
            }
        },
    };
}

fn fnSum(ctx: CallCtx, args: []const Value) FnError!Value {
    const Acc = struct {
        ctx: CallCtx,
        total: f64 = 0,
        failed: ?value.ScalarValue = null,

        fn visit(self: *@This(), s: value.ScalarValue, via_range: bool) FnError!bool {
            if (s == .err) {
                // `SUM` propagates, including from inside a range — and
                // the first error in §5.6a's iteration order is the one
                // that wins.
                self.failed = s;
                return false;
            }
            const n: f64 = if (via_range) blk: {
                // Text and booleans found in a range are ignored; numeric
                // text is NOT coerced (Excel's rule, §5.3b).
                if (s != .number) return true;
                break :blk s.number;
            } else switch (value.coerceToNumber(s, self.ctx.fidelity(), .function_arg)) {
                .number => |n| n,
                .value => |v| {
                    self.failed = v;
                    return false;
                },
                .locale_refusal => return error.LocaleSensitiveInput,
            };
            const r = value.addSub(self.ctx.rules(), self.total, n, .add);
            if (r == .err) {
                // An overflowing running total is `#NUM!`, and it stops
                // the fold rather than saturating quietly.
                self.failed = r;
                return false;
            }
            self.total = r.number;
            return true;
        }
    };
    var acc: Acc = .{ .ctx = ctx };
    try visitArgs(ctx, args, &acc);
    if (acc.failed) |f| return .{ .scalar = f };
    return .{ .scalar = value.ScalarValue.fromArithmetic(acc.total) };
}

fn fnCount(ctx: CallCtx, args: []const Value) FnError!Value {
    const Acc = struct {
        ctx: CallCtx,
        n: f64 = 0,

        fn visit(self: *@This(), s: value.ScalarValue, via_range: bool) FnError!bool {
            if (via_range) {
                // Numbers only. An error in a range is neither counted
                // nor propagated — §5.3c names `COUNT` specifically.
                if (s == .number) self.n += 1;
                return true;
            }
            // A direct argument coerces, so `COUNT("1")` is 1. A direct
            // error is still not a number and still does not propagate.
            switch (value.coerceToNumber(s, self.ctx.fidelity(), .function_arg)) {
                .number => self.n += 1,
                .value, .locale_refusal => {},
            }
            return true;
        }
    };
    var acc: Acc = .{ .ctx = ctx };
    try visitArgs(ctx, args, &acc);
    return Value.num(acc.n);
}

fn fnCountA(ctx: CallCtx, args: []const Value) FnError!Value {
    const Acc = struct {
        n: f64 = 0,

        fn visit(self: *@This(), s: value.ScalarValue, via_range: bool) FnError!bool {
            _ = via_range;
            // Everything that is not a true blank, which includes error
            // values and includes `""` — the three-way split that makes
            // `COUNTA`, `COUNTBLANK`, and `ISBLANK` disagree on purpose.
            if (s != .blank) self.n += 1;
            return true;
        }
    };
    var acc: Acc = .{};
    try visitArgs(ctx, args, &acc);
    return Value.num(acc.n);
}

fn fnCountBlank(ctx: CallCtx, args: []const Value) FnError!Value {
    const r = switch (args[0]) {
        .reference => |r| r,
        else => return Value.err(.value),
    };
    var total: u64 = 0;
    for (r.areas) |area| {
        total += try ctx.ev.readBlankCount(area, .countblank_class);
    }
    return Value.num(@floatFromInt(total));
}

fn fnAverage(ctx: CallCtx, args: []const Value) FnError!Value {
    const Acc = struct {
        ctx: CallCtx,
        total: f64 = 0,
        n: f64 = 0,
        failed: ?value.ScalarValue = null,

        fn visit(self: *@This(), s: value.ScalarValue, via_range: bool) FnError!bool {
            if (s == .err) {
                self.failed = s;
                return false;
            }
            // `SUM`'s split exactly: a range contributes numbers only,
            // a direct argument coerces. The denominator follows the
            // numerator, which is why `AVERAGE(1,)` is 0.5 and not 1.
            const x: f64 = if (via_range) blk: {
                if (s != .number) return true;
                break :blk s.number;
            } else switch (value.coerceToNumber(s, self.ctx.fidelity(), .function_arg)) {
                .number => |n| n,
                .value => |v| {
                    self.failed = v;
                    return false;
                },
                .locale_refusal => return error.LocaleSensitiveInput,
            };
            const r = value.addSub(self.ctx.rules(), self.total, x, .add);
            if (r == .err) {
                self.failed = r;
                return false;
            }
            self.total = r.number;
            self.n += 1;
            return true;
        }
    };
    var acc: Acc = .{ .ctx = ctx };
    try visitArgs(ctx, args, &acc);
    if (acc.failed) |f| return .{ .scalar = f };
    // Nothing numeric anywhere is `#DIV/0!`, spelled by the division
    // rather than tested for: an average of nothing IS a division by
    // zero, and routing it through `divide` keeps one answer for it.
    return .{ .scalar = value.divide(ctx.rules(), acc.total, acc.n) };
}

const Extreme = enum { min, max };

/// `MIN` and `MAX` — outside `collation_v1`'s comparator by §5.4b,
/// because they never compare text: a direct text argument coerces or
/// is `#VALUE!`, text found in a range is ignored, and no numbers
/// anywhere is 0 rather than an error.
fn foldExtreme(ctx: CallCtx, args: []const Value, mode: Extreme) FnError!Value {
    const Acc = struct {
        ctx: CallCtx,
        mode: Extreme,
        best: ?f64 = null,
        failed: ?value.ScalarValue = null,

        fn visit(self: *@This(), s: value.ScalarValue, via_range: bool) FnError!bool {
            if (s == .err) {
                self.failed = s;
                return false;
            }
            const x: f64 = if (via_range) blk: {
                if (s != .number) return true;
                break :blk s.number;
            } else switch (value.coerceToNumber(s, self.ctx.fidelity(), .function_arg)) {
                .number => |n| n,
                .value => |v| {
                    self.failed = v;
                    return false;
                },
                .locale_refusal => return error.LocaleSensitiveInput,
            };
            const cur = self.best orelse {
                self.best = x;
                return true;
            };
            // Deliberately not `@min`/`@max`: those are IEEE
            // minNum/maxNum, which treat `-0` and `+0` as
            // interchangeable and would decide — silently, and in
            // whichever direction the hardware felt like — the one case
            // N3 makes observable between the two rule tables.
            const take = switch (self.mode) {
                .min => x < cur,
                .max => x > cur,
            };
            if (take) self.best = x;
            return true;
        }
    };
    var acc: Acc = .{ .ctx = ctx, .mode = mode };
    try visitArgs(ctx, args, &acc);
    if (acc.failed) |f| return .{ .scalar = f };
    return .{ .scalar = value.ScalarValue.fromArithmetic(acc.best orelse 0) };
}

fn fnMin(ctx: CallCtx, args: []const Value) FnError!Value {
    return foldExtreme(ctx, args, .min);
}

fn fnMax(ctx: CallCtx, args: []const Value) FnError!Value {
    return foldExtreme(ctx, args, .max);
}

fn fnSumProduct(ctx: CallCtx, args: []const Value) FnError!Value {
    const grids = try ctx.arena().alloc(Grid, args.len);
    for (args, grids) |a, *g| g.* = (try gridOf(ctx, a)) orelse return Value.err(.value);

    // Excel requires identical dimensions here and does not broadcast:
    // `SUMPRODUCT(A1:A3, 2)` is `#VALUE!`, not three doubled cells.
    const rows = grids[0].rows;
    const cols = grids[0].cols;
    for (grids[1..]) |g| {
        if (g.rows != rows or g.cols != cols) return Value.err(.value);
    }

    // Errors first, argument by argument and row-major within each —
    // §5.6a's iteration order, which is the order `SUM` propagates in.
    // Folding this into the product loop below would have made the
    // answer depend on position instead of on declaration order.
    for (grids) |g| {
        var r: u32 = 0;
        while (r < rows) : (r += 1) {
            var c: u32 = 0;
            while (c < cols) : (c += 1) {
                const s = g.at(r, c);
                if (s == .err) return .{ .scalar = s };
            }
        }
    }

    var total: f64 = 0;
    var r: u32 = 0;
    while (r < rows) : (r += 1) {
        var c: u32 = 0;
        while (c < cols) : (c += 1) {
            var product: f64 = 1;
            for (grids) |g| {
                const s = g.at(r, c);
                // Anything that is not a number contributes zero — the
                // array rule, and the reason `SUMPRODUCT({TRUE})` is 0
                // where `SUM(TRUE)` is 1.
                const x: f64 = if (s == .number) s.number else 0;
                const p = value.multiply(ctx.rules(), product, x);
                if (p == .err) return .{ .scalar = p };
                product = p.number;
            }
            const t = value.addSub(ctx.rules(), total, product, .add);
            if (t == .err) return .{ .scalar = t };
            total = t.number;
        }
    }
    return .{ .scalar = value.ScalarValue.fromArithmetic(total) };
}

const LogicalFold = enum { all, any };

fn foldLogical(ctx: CallCtx, args: []const Value, mode: LogicalFold) FnError!Value {
    const Acc = struct {
        result: bool,
        mode: LogicalFold,
        seen: bool = false,
        failed: ?value.ScalarValue = null,

        fn visit(self: *@This(), s: value.ScalarValue, via_range: bool) FnError!bool {
            if (s == .err) {
                self.failed = s;
                return false;
            }
            const b: bool = switch (s) {
                .boolean => |v| v,
                .number => |n| n != 0,
                // Text and blanks are ignored in ranges; a direct text
                // argument is `#VALUE!`.
                .text => {
                    if (via_range) return true;
                    self.failed = value.ScalarValue.errorOf(.value);
                    return false;
                },
                .blank => return true,
                .err => unreachable,
            };
            self.seen = true;
            switch (self.mode) {
                .all => self.result = self.result and b,
                .any => self.result = self.result or b,
            }
            return true;
        }
    };
    var acc: Acc = .{ .result = mode == .all, .mode = mode };
    try visitArgs(ctx, args, &acc);
    if (acc.failed) |f| return .{ .scalar = f };
    // No logical value anywhere is `#VALUE!`, not a silent TRUE.
    if (!acc.seen) return Value.err(.value);
    return Value.boolean(acc.result);
}

fn fnAnd(ctx: CallCtx, args: []const Value) FnError!Value {
    return foldLogical(ctx, args, .all);
}

fn fnOr(ctx: CallCtx, args: []const Value) FnError!Value {
    return foldLogical(ctx, args, .any);
}

/// `IFS(cond1, val1, …)` — **eager**. Every arm has already been
/// evaluated by the time this runs, which is exactly what Excel does and
/// what makes the difference observable through volatile draws.
fn fnIfs(ctx: CallCtx, args: []const Value) FnError!Value {
    var i: usize = 0;
    while (i + 1 < args.len) : (i += 2) {
        const cond = try observedScalar(ctx, args[i]);
        switch (try ctx.ev.toLogical(cond)) {
            .b => |b| if (b) return args[i + 1],
            .err => |e| return .{ .scalar = e },
        }
    }
    // Excel's answer when nothing matched.
    return Value.err(.na);
}

/// `SWITCH(expr, match1, result1, …, [default])` — also eager.
fn fnSwitch(ctx: CallCtx, args: []const Value) FnError!Value {
    const subject = try observedScalar(ctx, args[0]);
    if (subject == .err) return .{ .scalar = subject };

    var i: usize = 1;
    while (i + 1 < args.len) : (i += 2) {
        const candidate = try observedScalar(ctx, args[i]);
        if (candidate == .err) return .{ .scalar = candidate };
        if ((try ctx.ev.compare(subject, candidate)) == .eq) return args[i + 1];
    }
    // A trailing odd argument is the default.
    if (i < args.len) return args[i];
    return Value.err(.na);
}

// ── criteria functions ──

fn criteriaContext(ctx: CallCtx) criteria.Context {
    return .{
        .allocator = ctx.arena(),
        .collation = ctx.ev.opts.collation,
        .fidelity = ctx.fidelity(),
    };
}

fn singleArea(v: eval.Value) ?env.RangeRef {
    return switch (v) {
        .reference => |r| r.single(),
        else => null,
    };
}

/// `criteria.Error` reaches the evaluator's taxonomy. `BadFold` is the
/// only member without a home there, and it is an injected-fold failure
/// rather than a value outcome.
fn mapCriteriaError(e: anyerror) FnError {
    return switch (e) {
        error.OutOfMemory => error.OutOfMemory,
        error.LocaleSensitiveInput => error.LocaleSensitiveInput,
        error.ShapeMismatch => error.ShapeMismatch,
        error.RefOutOfGrid => error.RefOutOfGrid,
        error.UnknownSheet => error.UnknownSheet,
        else => error.MalformedInput,
    };
}

/// Run one aligned pass. Both `COUNTIF` and `SUMIF` are this call plus a
/// choice of which number to return — which is the point of §5.6a
/// putting the alignment in one place.
fn runScan(
    ctx: CallCtx,
    areas: []const env.RangeRef,
    mode: env.AlignMode,
    criterion: criteria.Criterion,
) FnError!?criteria.ScanResult {
    const cursors = try ctx.arena().alloc(usize, areas.len);
    const scratch = try ctx.arena().alloc(value.ScalarValue, areas.len);
    var out: criteria.ScanResult = .{};
    criteria.scan(
        criteriaContext(ctx),
        ctx.ev.environment,
        areas,
        mode,
        &.{criterion},
        cursors,
        scratch,
        &out,
    ) catch |e| {
        // §5.6a: unequal dimensions under `.require_equal` are `#VALUE!`,
        // a value outcome rather than a refusal.
        if (e == error.ShapeMismatch or e == error.RefOutOfGrid) return null;
        return mapCriteriaError(e);
    };
    return out;
}

fn fnCountIf(ctx: CallCtx, args: []const Value) FnError!Value {
    const area = singleArea(args[0]) orelse return Value.err(.value);
    const criterion = criteria.parse(args[1].scalar, ctx.fidelity()) catch |e| return mapCriteriaError(e);
    const areas = [_]env.RangeRef{area};
    const out = (try runScan(ctx, &areas, .require_equal, criterion)) orelse return Value.err(.value);
    return Value.num(@floatFromInt(out.matched));
}

fn fnSumIf(ctx: CallCtx, args: []const Value) FnError!Value {
    const area = singleArea(args[0]) orelse return Value.err(.value);
    const criterion = criteria.parse(args[1].scalar, ctx.fidelity()) catch |e| return mapCriteriaError(e);

    if (args.len < 3) {
        // Two-argument form: the criteria range is also the sum range.
        const areas = [_]env.RangeRef{area};
        const out = (try runScan(ctx, &areas, .require_equal, criterion)) orelse return Value.err(.value);
        return Value.num(out.numeric_total);
    }
    const sum_area = singleArea(args[2]) orelse return Value.err(.value);
    // §5.6a: the sum range is PROJECTED from its top-left using the
    // criteria range's dimensions. The ranges need not be written the
    // same size, and Excel documents the projection rather than the
    // shape check `*IFS` uses.
    const areas = [_]env.RangeRef{ area, sum_area };
    const out = (try runScan(ctx, &areas, .project_from_first, criterion)) orelse return Value.err(.value);
    return Value.num(out.numeric_total);
}

fn fnAverageIf(ctx: CallCtx, args: []const Value) FnError!Value {
    const area = singleArea(args[0]) orelse return Value.err(.value);
    const criterion = criteria.parse(args[1].scalar, ctx.fidelity()) catch |e| return mapCriteriaError(e);

    const out = if (args.len < 3) blk: {
        const areas = [_]env.RangeRef{area};
        break :blk (try runScan(ctx, &areas, .require_equal, criterion)) orelse
            return Value.err(.value);
    } else blk: {
        const avg_area = singleArea(args[2]) orelse return Value.err(.value);
        // The same projection `SUMIF` uses, and for the same reason
        // (§5.6a): Excel documents the average range as projected from
        // its top-left rather than required to be the same size.
        const areas = [_]env.RangeRef{ area, avg_area };
        break :blk (try runScan(ctx, &areas, .project_from_first, criterion)) orelse
            return Value.err(.value);
    };
    // An average over no matching number is `#DIV/0!`. This is the one
    // place `AVERAGEIF` is not `SUMIF` with a division bolted on: a
    // total of nothing is 0 and an average of nothing is not.
    return .{ .scalar = value.divide(ctx.rules(), out.numeric_total, @floatFromInt(out.numeric_count)) };
}

// ─── lookups (§5.4b: equality AND ordering from `collation_v1`) ──

/// A rectangular view of one argument, with random access.
///
/// Aggregate walking (`visitArgs`) is a fold and needs no coordinates;
/// a lookup is a search along an axis and needs them, so a reference is
/// materialized here rather than streamed. The cell cap is §9's, which
/// is why a whole-column lookup is a **limit** rather than a hang —
/// making it fast is M7b2's row, not this one.
const Grid = struct {
    rows: u32,
    cols: u32,
    src: union(enum) { matrix: value.Matrix, scalar: value.ScalarValue },

    fn at(self: Grid, r: u32, c: u32) value.ScalarValue {
        return switch (self.src) {
            .scalar => |s| s,
            .matrix => |m| m.at(r, c),
        };
    }
};

/// `null` where the argument is not one rectangle — a multi-area union,
/// which every lookup in this batch answers `#VALUE!` for. A 3D span
/// never reaches here: §5.6g refuses one at an ineligible function
/// before evaluation starts.
fn gridOf(ctx: CallCtx, v: Value) FnError!?Grid {
    switch (v) {
        .scalar => |s| return Grid{ .rows = 1, .cols = 1, .src = .{ .scalar = s } },
        .missing_arg => return Grid{ .rows = 1, .cols = 1, .src = .{ .scalar = .blank } },
        .array => |m| return Grid{ .rows = m.rows, .cols = m.cols, .src = .{ .matrix = m } },
        .reference => |r| {
            const area = r.single() orelse return null;
            const shape = area.shape();
            var m = try value.Matrix.init(ctx.arena(), shape.rows, shape.cols);
            var it = try ctx.ev.readRange(area);
            while (try it.next()) |e| {
                m.set(
                    e.row.oneBased() - area.range.first.row.oneBased(),
                    e.col.zeroBased() - area.range.first.col.zeroBased(),
                    e.value,
                );
            }
            return Grid{ .rows = shape.rows, .cols = shape.cols, .src = .{ .matrix = m } };
        },
    }
}

/// One axis of a grid. `MATCH`, `XMATCH` and `XLOOKUP` search a vector;
/// `VLOOKUP` and `HLOOKUP` search the first column or row of a table.
const Vector = struct {
    grid: Grid,
    /// Down a column, rather than across a row.
    down: bool,
    len: u32,

    /// Always the FIRST column (or row) of the grid — a lookup table's
    /// key axis is its first, and a standalone vector has only one.
    fn at(self: Vector, i: u32) value.ScalarValue {
        return if (self.down) self.grid.at(i, 0) else self.grid.at(0, i);
    }
};

/// A grid that is a vector, whichever way it runs. `null` for a 2-D
/// one, which `MATCH` has no single axis to search.
fn vectorOf(g: Grid) ?Vector {
    if (g.cols == 1) return .{ .grid = g, .down = true, .len = g.rows };
    if (g.rows == 1) return .{ .grid = g, .down = false, .len = g.cols };
    return null;
}

/// Excel's match semantics, spelled once for the five names that share
/// them. `MATCH`'s type and `XMATCH`'s mode are two spellings of this.
const MatchMode = enum {
    /// Equality, with `*`/`?`/`~` inert.
    exact,
    /// Equality, with the wildcards active — what `MATCH(…,0)` and an
    /// unsorted `VLOOKUP` do, and what `XMATCH(…,2)` asks for by name.
    wildcard,
    /// The largest value ≤ the key (`MATCH` type 1, `XMATCH` mode −1).
    exact_or_smaller,
    /// The smallest value ≥ the key (`MATCH` type −1, `XMATCH` mode 1).
    exact_or_larger,
};

const SearchOrder = enum { first_to_last, last_to_first };

/// Lookup equality is `collation_v1`'s, reached through the criteria
/// matcher rather than through a second comparator: §5.4b names lookup
/// equality and criteria in the same sentence, the matcher is already
/// type-restricted the way a lookup is (a numeric key never matches a
/// text cell), and the wildcard engine is already there.
///
/// The lookup VALUE is not a criterion, though, which is why this is
/// not `criteria.parse`: that would read `VLOOKUP("<5",…)` as "less
/// than 5" where Excel looks for the literal text `<5`.
fn lookupCriterion(s: value.ScalarValue, wildcards: bool) criteria.Criterion {
    return switch (s) {
        .number => |n| .{ .relation = .eq, .operand = .{ .number = n } },
        .boolean => |b| .{ .relation = .eq, .operand = .{ .boolean = b } },
        .err => |e| .{ .relation = .eq, .operand = .{ .err = e } },
        .blank => .{ .relation = .eq, .operand = .empty },
        .text => |t| .{
            .relation = .eq,
            .operand = .{ .text = t },
            .is_pattern = wildcards and
                (criteria.hasWildcards(t) or std.mem.indexOfScalar(u8, t, '~') != null),
            .source = t,
        },
    };
}

/// Where the key is, along the vector — or `null` for no match.
fn scanVector(
    ctx: CallCtx,
    vec: Vector,
    key: value.ScalarValue,
    mode: MatchMode,
    order: SearchOrder,
) FnError!?u32 {
    const criterion = lookupCriterion(key, mode == .wildcard);
    // A blank key in an ordered scan compares as 0 — the same adoption
    // `Evaluator.compare` applies, hoisted here because the type
    // restriction below has to ask about the adopted type, not the
    // written one.
    const ordered_key = if (key == .blank) value.ScalarValue.fromNumber(0) else key;

    var best: ?u32 = null;
    var i: u32 = 0;
    while (i < vec.len) : (i += 1) {
        const idx = if (order == .last_to_first) vec.len - 1 - i else i;
        const cell = vec.at(idx);
        switch (mode) {
            .exact, .wildcard => {
                const hit = criteria.matches(criteriaContext(ctx), criterion, cell) catch |e|
                    return mapCriteriaError(e);
                if (hit) return idx;
            },
            .exact_or_smaller, .exact_or_larger => {
                // An empty cell is not a candidate, and an error in the
                // lookup vector is not one either: the pass is looking
                // for a value, and `#N/A` in a column is not an answer
                // to "where is 5". Neither propagates — the error the
                // caller sees is `#N/A`, from the failed match.
                if (cell == .blank or cell == .err) continue;
                if (ordered_key == .err) continue;
                // Excel's ordered match never crosses a type boundary,
                // so the total cross-type order (number < text <
                // logical) decides which cells are candidates and
                // nothing else.
                const cr = value.crossTypeRank(cell) orelse continue;
                const kr = value.crossTypeRank(ordered_key) orelse continue;
                if (cr != kr) continue;

                const rel = try ctx.ev.compare(cell, ordered_key);
                const eligible = switch (mode) {
                    .exact_or_smaller => rel != .gt,
                    .exact_or_larger => rel != .lt,
                    else => unreachable,
                };
                if (!eligible) continue;
                const b = best orelse {
                    best = idx;
                    continue;
                };
                const against = try ctx.ev.compare(cell, vec.at(b));
                const better = switch (mode) {
                    .exact_or_smaller => against == .gt,
                    .exact_or_larger => against == .lt,
                    else => unreachable,
                };
                // Ties go to the LAST position in array order, whichever
                // direction the scan ran — so `search_mode` changes an
                // exact match's answer and leaves an ordered one alone,
                // which is what makes the two search orders agree on
                // sorted data the way Excel's binary search does.
                if (better or (against == .eq and idx > b)) best = idx;
            },
        }
    }
    return best;
}

/// §5.3c's declaration order, taken by the implementation rather than
/// by the dispatcher — which is what `per_function_provenance` means
/// for these five names.
///
/// Two things `propagateAndInvoke` cannot do for a lookup. It cannot
/// see an error in the KEY slot, because that slot is `.value_any` and
/// a key usually arrives as a reference rather than as a scalar; and it
/// must NOT propagate an error found inside the lookup TABLE, because
/// that error is a value the lookup may or may not return. `tables` is
/// the bitmask of slots holding one, and every other slot is read in
/// declaration order through the same reduction the key needs.
fn lookupPropagate(ctx: CallCtx, args: []const Value, tables: u8) FnError!?value.ScalarValue {
    for (args, 0..) |a, k| {
        if (k < 8 and (tables >> @intCast(k)) & 1 == 1) continue;
        const s = try observedScalar(ctx, a);
        if (s == .err) return s;
    }
    return null;
}

const LookupAxis = enum { vertical, horizontal };

fn fnVLookup(ctx: CallCtx, args: []const Value) FnError!Value {
    return lookupTable(ctx, args, .vertical);
}

fn fnHLookup(ctx: CallCtx, args: []const Value) FnError!Value {
    return lookupTable(ctx, args, .horizontal);
}

/// `VLOOKUP` and `HLOOKUP` are one function under two axes — the table
/// is transposed and nothing else changes, which is why they share an
/// implementation rather than two that could drift.
fn lookupTable(ctx: CallCtx, args: []const Value, axis: LookupAxis) FnError!Value {
    // Slot 1 is the table; every other slot propagates, in order.
    if (try lookupPropagate(ctx, args, 0b0010)) |e| return .{ .scalar = e };
    const key = try observedScalar(ctx, args[0]);
    const table = (try gridOf(ctx, args[1])) orelse return Value.err(.value);

    // The result axis. Below 1 is a formula Excel calls `#VALUE!`;
    // beyond the table is `#REF!` — two different mistakes, and Excel
    // spells them differently.
    const wanted = std.math.trunc(numArg(args, 2));
    const limit: u32 = if (axis == .vertical) table.cols else table.rows;
    if (wanted < 1) return Value.err(.value);
    if (wanted > @as(f64, @floatFromInt(limit))) return Value.err(.ref);
    const offset: u32 = @as(u32, @intFromFloat(wanted)) - 1;

    // Excel's default is approximate, which is the default most people
    // did not want; it is still the default.
    const approximate = if (args.len >= 4) args[3].scalar.boolean else true;
    const keys: Vector = if (axis == .vertical)
        .{ .grid = table, .down = true, .len = table.rows }
    else
        .{ .grid = table, .down = false, .len = table.cols };

    const hit = (try scanVector(
        ctx,
        keys,
        key,
        if (approximate) .exact_or_smaller else .wildcard,
        .first_to_last,
    )) orelse return Value.err(.na);

    return .{
        .scalar = if (axis == .vertical) table.at(hit, offset) else table.at(offset, hit),
    };
}

fn fnMatch(ctx: CallCtx, args: []const Value) FnError!Value {
    if (try lookupPropagate(ctx, args, 0b0010)) |e| return .{ .scalar = e };
    const key = try observedScalar(ctx, args[0]);
    const g = (try gridOf(ctx, args[1])) orelse return Value.err(.value);
    // `MATCH` answers a position along one axis, and a 2-D array has no
    // single one to answer along.
    const vec = vectorOf(g) orelse return Value.err(.na);

    const t = std.math.trunc(if (args.len >= 3) numArg(args, 2) else 1);
    const mode: MatchMode = if (t > 0)
        .exact_or_smaller
    else if (t < 0)
        .exact_or_larger
    else
        .wildcard;

    const hit = (try scanVector(ctx, vec, key, mode, .first_to_last)) orelse
        return Value.err(.na);
    return Value.num(@floatFromInt(hit + 1));
}

/// `XMATCH`'s mode pair, shared with `XLOOKUP`. An unlisted mode is
/// `#VALUE!` rather than a nearest-neighbour reading of what the caller
/// might have meant.
const XModes = struct { mode: MatchMode, order: SearchOrder };

fn xModes(args: []const Value, mode_slot: usize, order_slot: usize) ?XModes {
    const m = std.math.trunc(if (args.len > mode_slot) numArg(args, mode_slot) else 0);
    const mode: MatchMode = if (m == 0)
        .exact
    else if (m == -1)
        .exact_or_smaller
    else if (m == 1)
        .exact_or_larger
    else if (m == 2)
        .wildcard
    else
        return null;

    // The two binary modes document a requirement on the *caller* (a
    // sorted array), not a different answer: over sorted input a linear
    // pass in the same direction lands on the same element, so they map
    // onto the two scan orders rather than onto a second algorithm.
    const s = std.math.trunc(if (args.len > order_slot) numArg(args, order_slot) else 1);
    const order: SearchOrder = if (s == 1 or s == 2)
        .first_to_last
    else if (s == -1 or s == -2)
        .last_to_first
    else
        return null;

    return .{ .mode = mode, .order = order };
}

fn fnXMatch(ctx: CallCtx, args: []const Value) FnError!Value {
    if (try lookupPropagate(ctx, args, 0b0010)) |e| return .{ .scalar = e };
    const key = try observedScalar(ctx, args[0]);
    const g = (try gridOf(ctx, args[1])) orelse return Value.err(.value);
    const vec = vectorOf(g) orelse return Value.err(.na);
    const m = xModes(args, 2, 3) orelse return Value.err(.value);

    const hit = (try scanVector(ctx, vec, key, m.mode, m.order)) orelse
        return Value.err(.na);
    return Value.num(@floatFromInt(hit + 1));
}

fn fnXLookup(ctx: CallCtx, args: []const Value) FnError!Value {
    // Slots 1 and 2 are the arrays — and slot 3, `if_not_found`, is
    // masked with them. It is not an array; it is a value this call
    // may RETURN, which puts it on the same side of the line for the
    // same reason: propagating it would make a successful lookup fail
    // over a fallback nobody reached. `XLOOKUP(20,…,A5)` is the hit,
    // and `XLOOKUP(99,…,A5)` is A5 — returned, not propagated.
    if (try lookupPropagate(ctx, args, 0b1110)) |e| return .{ .scalar = e };
    const key = try observedScalar(ctx, args[0]);
    const lg = (try gridOf(ctx, args[1])) orelse return Value.err(.value);
    const vec = vectorOf(lg) orelse return Value.err(.value);
    const rg = (try gridOf(ctx, args[2])) orelse return Value.err(.value);
    const m = xModes(args, 4, 5) orelse return Value.err(.value);

    // The return range is indexed along the axis the lookup vector runs
    // on, and must be the same length on it.
    const along: u32 = if (vec.down) rg.rows else rg.cols;
    if (along != vec.len) return Value.err(.value);

    const hit = (try scanVector(ctx, vec, key, m.mode, m.order)) orelse {
        // `if_not_found` replaces the `#N/A` a failed match would be,
        // and only that: it is not an `IFERROR`, so an error anywhere
        // else has already propagated (§5.3c) and never reaches here.
        if (args.len >= 4) return args[3];
        return Value.err(.na);
    };

    // A vector return range yields the element; a 2-D one yields the
    // whole row (or column) at the match, which is Excel's answer and
    // an array this evaluator already carries.
    const across: u32 = if (vec.down) rg.cols else rg.rows;
    if (across == 1) {
        return .{ .scalar = if (vec.down) rg.at(hit, 0) else rg.at(0, hit) };
    }
    var m2 = try value.Matrix.init(
        ctx.arena(),
        if (vec.down) 1 else across,
        if (vec.down) across else 1,
    );
    var k: u32 = 0;
    while (k < across) : (k += 1) {
        if (vec.down) m2.set(0, k, rg.at(hit, k)) else m2.set(k, 0, rg.at(k, hit));
    }
    return .{ .array = m2 };
}

fn fnIndex(ctx: CallCtx, args: []const Value) FnError!Value {
    const g = (try gridOf(ctx, args[0])) orelse return Value.err(.value);
    const first = std.math.trunc(numArg(args, 1));

    var want_row = first;
    var want_col: f64 = if (args.len >= 3) std.math.trunc(numArg(args, 2)) else 0;
    if (args.len < 3) {
        // One index runs along a vector's own axis; on a 2-D array it
        // selects a whole row, which is what a bare `0` column means
        // below.
        if (g.rows == 1) {
            want_row = 1;
            want_col = first;
        } else if (g.cols == 1) {
            want_col = 1;
        }
    }

    if (want_row < 0 or want_col < 0) return Value.err(.value);
    if (want_row > @as(f64, @floatFromInt(g.rows))) return Value.err(.ref);
    if (want_col > @as(f64, @floatFromInt(g.cols))) return Value.err(.ref);
    const r: u32 = @intFromFloat(want_row);
    const c: u32 = @intFromFloat(want_col);

    if (r > 0 and c > 0) return .{ .scalar = g.at(r - 1, c - 1) };
    // A zero index means "the whole axis" — the form that makes
    // `INDEX(A1:B3,0,2)` a column rather than an error.
    const rows: u32 = if (r == 0) g.rows else 1;
    const cols: u32 = if (c == 0) g.cols else 1;
    if (rows == 1 and cols == 1) {
        return .{ .scalar = g.at(if (r == 0) 0 else r - 1, if (c == 0) 0 else c - 1) };
    }
    var m = try value.Matrix.init(ctx.arena(), rows, cols);
    var i: u32 = 0;
    while (i < rows) : (i += 1) {
        var j: u32 = 0;
        while (j < cols) : (j += 1) {
            m.set(i, j, g.at(if (r == 0) i else r - 1, if (c == 0) j else c - 1));
        }
    }
    return .{ .array = m };
}

// ── position: what a reference IS ──

/// The first area of a reference argument. The `.reference` coercion
/// class has already turned a non-reference into `#VALUE!` and
/// `.propagate` has already returned it, so the `else` arm is
/// unreachable in practice and stated anyway — an exhaustive switch is
/// cheaper than a comment claiming it cannot happen.
fn firstArea(v: Value) ?env.RangeRef {
    return switch (v) {
        .reference => |r| if (r.areas.len == 0) null else r.areas[0],
        else => null,
    };
}

fn fnRow(ctx: CallCtx, args: []const Value) FnError!Value {
    if (args.len == 0) {
        const s = ctx.site() orelse return error.AnchorRequired;
        return Value.num(@floatFromInt(s.row.oneBased()));
    }
    const area = firstArea(args[0]) orelse return Value.err(.value);
    return Value.num(@floatFromInt(area.range.first.row.oneBased()));
}

fn fnColumn(ctx: CallCtx, args: []const Value) FnError!Value {
    if (args.len == 0) {
        const s = ctx.site() orelse return error.AnchorRequired;
        return Value.num(@floatFromInt(s.col.zeroBased() + 1));
    }
    const area = firstArea(args[0]) orelse return Value.err(.value);
    return Value.num(@floatFromInt(area.range.first.col.zeroBased() + 1));
}

const Axis = enum { rows, cols };

/// How far an argument reaches along one axis. No cell is read: the
/// answer is a property of the shape, which is why `ROWS(A1:A9)` is 9
/// whether or not anything is stored in it.
fn spanOf(v: Value, axis: Axis) u32 {
    return switch (v) {
        .reference => |r| blk: {
            if (r.areas.len == 0) break :blk 1;
            const s = r.areas[0].shape();
            break :blk if (axis == .rows) s.rows else s.cols;
        },
        .array => |m| if (axis == .rows) m.rows else m.cols,
        .scalar, .missing_arg => 1,
    };
}

fn fnRows(ctx: CallCtx, args: []const Value) FnError!Value {
    _ = ctx;
    return Value.num(@floatFromInt(spanOf(args[0], .rows)));
}

fn fnColumns(ctx: CallCtx, args: []const Value) FnError!Value {
    _ = ctx;
    return Value.num(@floatFromInt(spanOf(args[0], .cols)));
}

// ─── the frozen inventory (committed data) ───────────────────────

/// The authoritative v1 function list. Data, not code: every F-batch PR
/// regenerates its count from this file (§7).
pub const inventory_v1 = @embedFile("function_inventory_v1.tsv");

/// Frozen at M3a2. Asserted against the file, so the two cannot drift.
pub const inventory_count: usize = 175;

pub const InventoryEntry = struct {
    name: []const u8,
    milestone: []const u8,
    /// `-` where the ladder row names no F-batch label.
    batch: []const u8,
};

pub const InventoryIterator = struct {
    rest: []const u8,

    pub fn next(self: *InventoryIterator) ?InventoryEntry {
        while (self.rest.len > 0) {
            const end = std.mem.indexOfScalar(u8, self.rest, '\n') orelse self.rest.len;
            const line = self.rest[0..end];
            self.rest = self.rest[@min(end + 1, self.rest.len)..];
            if (line.len == 0 or line[0] == '#') continue;
            var parts = std.mem.splitScalar(u8, line, '\t');
            const name = parts.next() orelse continue;
            const milestone = parts.next() orelse continue;
            const batch = parts.next() orelse continue;
            return .{ .name = name, .milestone = milestone, .batch = batch };
        }
        return null;
    }
};

pub fn inventory() InventoryIterator {
    return .{ .rest = inventory_v1 };
}

pub fn inInventory(name: []const u8) bool {
    var it = inventory();
    while (it.next()) |e| {
        if (std.mem.eql(u8, e.name, name)) return true;
    }
    return false;
}

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

test "inventory: the frozen count is what the file holds" {
    var it = inventory();
    var n: usize = 0;
    while (it.next()) |_| n += 1;
    try testing.expectEqual(inventory_count, n);
}

test "inventory: names are unique, sorted, and fully tagged" {
    var it = inventory();
    var previous: []const u8 = "";
    while (it.next()) |e| {
        // Strictly ascending: uniqueness and order in one assertion.
        try testing.expect(std.mem.order(u8, previous, e.name) == .lt);
        previous = e.name;
        try testing.expect(e.name.len > 0);
        try testing.expect(e.milestone.len > 0);
        try testing.expect(e.batch.len > 0);
        try testing.expectEqual(@as(u8, 'M'), e.milestone[0]);
        for (e.name) |c| {
            try testing.expect(std.ascii.isUpper(c) or std.ascii.isDigit(c) or c == '.');
        }
    }
}

test "inventory: per-milestone counts reproduce the ladder" {
    // The ladder's per-row counts and this file are the same fact
    // written twice; if they ever disagree, the file wins and this test
    // says so out loud.
    const Expected = struct { milestone: []const u8, n: usize };
    const expected = [_]Expected{
        .{ .milestone = "M4c", .n = 20 },
        .{ .milestone = "M4d", .n = 17 },
        .{ .milestone = "M4e", .n = 22 },
        .{ .milestone = "M4f", .n = 19 },
        .{ .milestone = "M4g", .n = 15 },
        .{ .milestone = "M5a2", .n = 2 },
        .{ .milestone = "M7a", .n = 7 },
        .{ .milestone = "M7b2", .n = 6 },
        .{ .milestone = "M7b3", .n = 11 },
        .{ .milestone = "M8a", .n = 1 },
        .{ .milestone = "M8b", .n = 1 },
        .{ .milestone = "M8c", .n = 19 },
        .{ .milestone = "M9c1", .n = 7 },
        .{ .milestone = "M9c2", .n = 8 },
        .{ .milestone = "M9d", .n = 20 },
    };

    var total: usize = 0;
    for (expected) |e| {
        var it = inventory();
        var n: usize = 0;
        while (it.next()) |entry| {
            if (std.mem.eql(u8, entry.milestone, e.milestone)) n += 1;
        }
        try testing.expectEqual(e.n, n);
        total += e.n;
    }
    // Every row belongs to one of the listed milestones — a typo in a
    // milestone tag would otherwise hide a function from the ladder.
    try testing.expectEqual(inventory_count, total);

    // The "Core gate" figure in the ladder is M4c + M4d + M4e.
    try testing.expectEqual(@as(usize, 59), expected[0].n + expected[1].n + expected[2].n);
}

// ─── M4c: the F1a-1 batch, against the frozen inventory ──────────

/// The batch this row ships, named the way the inventory names it. Both
/// halves are checked: a row tagged `M4c` that the registry cannot
/// answer is a missing function, and a registered function tagged `M4c`
/// that the file does not list is a substitution. §7 makes the file the
/// count source, so neither number is written down here.
const m4c_milestone = "M4c";
const m4c_batch = "F1a-1";

fn isM4c(e: InventoryEntry) bool {
    return std.mem.eql(u8, e.milestone, m4c_milestone) and
        std.mem.eql(u8, e.batch, m4c_batch);
}

test "M4c: the batch's size is regenerated from the inventory, never from prose" {
    // The ladder row says "20". This test does not: it counts the file,
    // and the ladder's number is checked against the same file by
    // `inventory: per-milestone counts reproduce the ladder`. If someone
    // adds a 21st F1a-1 row, that test fails on the ladder's count and
    // this one keeps holding the registry to whatever the file now says.
    var it = inventory();
    var counted: usize = 0;
    while (it.next()) |e| {
        if (isM4c(e)) counted += 1;
    }
    try testing.expect(counted > 0);

    // Every `M4c`/`F1a-1` row resolves. This is the "no additions and no
    // substitutions" gate in the direction the file decides.
    var it2 = inventory();
    var resolved: usize = 0;
    while (it2.next()) |e| {
        if (!isM4c(e)) continue;
        const f = lookup(e.name) orelse {
            std.debug.print("F1a-1 name not registered: {s}\n", .{e.name});
            return error.UnregisteredBatchFunction;
        };
        // `lookup` folds case, so a row spelled `Isna` would resolve to
        // `ISNA`; the table's own spelling is what ships.
        try testing.expectEqualStrings(e.name, f.name);
        resolved += 1;
    }
    try testing.expectEqual(counted, resolved);

    // …and in the other direction: a registered function whose inventory
    // row says `M4c` must be one the file lists under this batch. A
    // function this row invented would be registered, present in the
    // inventory under some *other* milestone, and caught by neither
    // half without this.
    var registered_m4c: usize = 0;
    for (&functions) |*f| {
        var it3 = inventory();
        while (it3.next()) |e| {
            if (!std.mem.eql(u8, e.name, f.name)) continue;
            if (std.mem.eql(u8, e.milestone, m4c_milestone)) {
                if (!isM4c(e)) {
                    std.debug.print("{s}: tagged {s} but batch {s}\n", .{ f.name, e.milestone, e.batch });
                    return error.BatchTagMismatch;
                }
                registered_m4c += 1;
            }
            break;
        }
    }
    try testing.expectEqual(counted, registered_m4c);
}

test "M4c: every F1a-1 row declares all five fields, and none of them by default" {
    // The struct-level guarantee is below (`the five required fields
    // have no defaults`); this is the row-level one. A field with no
    // default cannot be omitted, so "declared" is the compiler's job —
    // what is left to check is that what was declared says something.
    var it = inventory();
    var seen: usize = 0;
    while (it.next()) |e| {
        if (!isM4c(e)) continue;
        const f = lookup(e.name).?;
        seen += 1;

        try testing.expect(f.name.len > 0);
        try testing.expectEqualStrings(e.name, f.name);

        // Arity: a real range, with a repeating slot iff it is unbounded.
        if (f.arity.max) |max| try testing.expect(f.arity.min <= max);
        if (f.arity.max == null) try testing.expect(f.arity.rest.len > 0);
        // Coercion: the two per-slot tables address the same slots, and
        // a lazy slot lines up with the class that defers.
        try testing.expectEqual(f.arity.fixed.len, f.coercion.fixed.len);
        try testing.expectEqual(f.arity.rest.len, f.coercion.rest.len);
        for (f.arity.fixed, f.coercion.fixed) |l, c| try testing.expectEqual(l == .lazy, c == .lazy_any);
        for (f.arity.rest, f.coercion.rest) |l, c| try testing.expectEqual(l == .lazy, c == .lazy_any);
        // Volatility: nothing in F1a-1 is redrawn. Stated per row rather
        // than assumed, because `.stable` is also what a forgotten field
        // would have looked like if it had a default — which is why it
        // does not have one.
        try testing.expectEqual(Volatility.stable, f.volatility);
        // Propagation: §5.3c's classes, and F1a-1 uses exactly two of
        // them. `per_element` is the operators' and
        // `per_function_provenance` belongs to the counting family.
        try testing.expect(f.propagation == .observe or f.propagation == .propagate);
    }
    try testing.expect(seen > 0);
}

test "M4c: TRUE and FALSE carry the same five fields as every other row" {
    // They joined F1a-1 at M3a2 (decision 2) rather than being written
    // for it, so the row proves they are ordinary entries instead of
    // assuming it. Zero-argument is the shape most likely to have been
    // waved through.
    for ([_][]const u8{ "TRUE", "FALSE" }) |name| {
        const f = lookup(name).?;
        try testing.expectEqualStrings(name, f.name);
        try testing.expectEqual(@as(u8, 0), f.arity.min);
        try testing.expectEqual(@as(?u8, 0), f.arity.max);
        try testing.expectEqual(@as(usize, 0), f.arity.fixed.len);
        try testing.expectEqual(@as(usize, 0), f.coercion.fixed.len);
        try testing.expectEqual(Volatility.stable, f.volatility);
        try testing.expectEqual(value.PropagationClass.propagate, f.propagation);
        try testing.expectEqual(Form.plain, f.form);
        try testing.expect(f.impl != null);
        // A no-slot signature is not liftable: there is nothing to lift.
        try testing.expect(!f.liftable());
        try testing.expect(inInventory(name));
    }
}

test "M4c: the batch's propagation classes are the ones §5.3c assigns" {
    // Stated name by name because §5.3c's whole point is that the class
    // is per function: `ISERR` observes and `N` propagates, and they sit
    // two rows apart in the same table.
    const observers = [_][]const u8{
        "ISBLANK",  "ISERR",  "ISERROR", "ISLOGICAL", "ISNA",
        "ISNUMBER", "ISTEXT", "IF",      "IFERROR",   "IFNA",
        "IFS",      "SWITCH",
    };
    for (observers) |name| {
        try testing.expectEqual(value.PropagationClass.observe, lookup(name).?.propagation);
    }
    const propagators = [_][]const u8{ "AND", "OR", "NOT", "N", "T", "NA", "TRUE", "FALSE" };
    for (propagators) |name| {
        try testing.expectEqual(value.PropagationClass.propagate, lookup(name).?.propagation);
    }
    // Together they are the whole batch — so a name added to one list
    // and forgotten in the other cannot pass.
    var it = inventory();
    var n: usize = 0;
    while (it.next()) |e| {
        if (isM4c(e)) n += 1;
    }
    try testing.expectEqual(n, observers.len + propagators.len);
}

test "M4c: the lazy forms are exactly the three §5.3a defers" {
    // IFS and SWITCH are in this batch and are **eager** (§5.3a): Excel
    // evaluates every arm, observably. So they carry an `impl` and no
    // lazy slot, and IF/IFERROR/IFNA carry a form and no impl.
    for ([_][]const u8{ "IF", "IFERROR", "IFNA" }) |name| {
        const f = lookup(name).?;
        try testing.expect(f.form != .plain);
        try testing.expect(f.impl == null);
        var lazy_slots: usize = 0;
        for (f.arity.fixed) |l| {
            if (l == .lazy) lazy_slots += 1;
        }
        try testing.expect(lazy_slots > 0);
    }
    for ([_][]const u8{ "IFS", "SWITCH", "AND", "OR" }) |name| {
        const f = lookup(name).?;
        try testing.expectEqual(Form.plain, f.form);
        try testing.expect(f.impl != null);
        for (f.arity.fixed) |l| try testing.expectEqual(Laziness.eager, l);
        for (f.arity.rest) |l| try testing.expectEqual(Laziness.eager, l);
    }
}

// ─── M4d: the F1a-2 batch, against the frozen inventory ──────────

const m4d_milestone = "M4d";
const m4d_batch = "F1a-2";

fn isM4d(e: InventoryEntry) bool {
    return std.mem.eql(u8, e.milestone, m4d_milestone) and
        std.mem.eql(u8, e.batch, m4d_batch);
}

test "M4d: the batch's size is regenerated from the inventory, never from prose" {
    // The ladder row says "~17". This test does not say 17: it counts
    // the file. `inventory: per-milestone counts reproduce the ladder`
    // holds the ladder's own number to the same file, so an eighteenth
    // F1a-2 row fails there and is held to the registry here.
    var it = inventory();
    var counted: usize = 0;
    while (it.next()) |e| {
        if (isM4d(e)) counted += 1;
    }
    try testing.expect(counted > 0);

    // Direction one — no omissions: every `M4d`/`F1a-2` row resolves,
    // under the file's own spelling and not merely case-insensitively.
    var it2 = inventory();
    var resolved: usize = 0;
    while (it2.next()) |e| {
        if (!isM4d(e)) continue;
        const f = lookup(e.name) orelse {
            std.debug.print("F1a-2 name not registered: {s}\n", .{e.name});
            return error.UnregisteredBatchFunction;
        };
        try testing.expectEqualStrings(e.name, f.name);
        resolved += 1;
    }
    try testing.expectEqual(counted, resolved);

    // Direction two — no substitutions: a registered function whose
    // inventory row says `M4d` must be one the file lists under this
    // batch. A name this row invented would be registered and tagged to
    // some other milestone, and neither half above would have noticed.
    var registered: usize = 0;
    for (&functions) |*f| {
        var it3 = inventory();
        while (it3.next()) |e| {
            if (!std.mem.eql(u8, e.name, f.name)) continue;
            if (std.mem.eql(u8, e.milestone, m4d_milestone)) {
                if (!isM4d(e)) {
                    std.debug.print("{s}: tagged {s} but batch {s}\n", .{ f.name, e.milestone, e.batch });
                    return error.BatchTagMismatch;
                }
                registered += 1;
            }
            break;
        }
    }
    try testing.expectEqual(counted, registered);
}

test "M4d: every F1a-2 row declares all five fields, and none of them by default" {
    var it = inventory();
    var seen: usize = 0;
    var volatiles: usize = 0;
    while (it.next()) |e| {
        if (!isM4d(e)) continue;
        const f = lookup(e.name).?;
        seen += 1;

        try testing.expect(f.name.len > 0);
        try testing.expectEqualStrings(e.name, f.name);

        // Arity: a real range, bounded — nothing in F1a-2 is variadic,
        // so every row states a maximum and none carries a repeating
        // slot.
        try testing.expect(f.arity.max != null);
        try testing.expect(f.arity.min <= f.arity.max.?);
        try testing.expectEqual(@as(usize, 0), f.arity.rest.len);
        try testing.expectEqual(@as(usize, 0), f.coercion.rest.len);
        // Coercion: the two per-slot tables address the same slots, and
        // every declared slot of this batch is the numeric class. A
        // `.value_any` here would silently skip the coercion the
        // implementations rely on.
        try testing.expectEqual(f.arity.fixed.len, f.coercion.fixed.len);
        for (f.arity.fixed) |l| try testing.expectEqual(Laziness.eager, l);
        for (f.coercion.fixed) |c| try testing.expectEqual(CoercionClass.number, c);
        // Volatility: counted rather than assumed to be `.stable`,
        // because `.stable` is also what a forgotten field would look
        // like if the field had a default — which is why it has none.
        if (f.volatility == .volatile_fn) volatiles += 1;
        // Propagation: F1a-2 uses exactly one of §5.3c's four classes.
        // Nothing in it observes an error; a numeric function handed one
        // becomes it.
        try testing.expectEqual(value.PropagationClass.propagate, f.propagation);
        // Every row is eager and plain: no member of this batch defers
        // an argument, so none may carry a form.
        try testing.expectEqual(Form.plain, f.form);
        try testing.expect(f.impl != null);
    }
    try testing.expect(seen > 0);
    try testing.expectEqual(@as(usize, 2), volatiles);
}

test "M4d: SQRT and RAND carry the same five fields as the fifteen the row writes" {
    // They were registered at M3a2 as framework subjects rather than
    // written for this batch, so the row proves they are ordinary
    // entries instead of assuming it — the same treatment M4c gave TRUE
    // and FALSE, and for the same reason: an early row is the one most
    // likely to have been waved through.
    const sqrt = lookup("SQRT").?;
    try testing.expectEqualStrings("SQRT", sqrt.name);
    try testing.expectEqual(@as(u8, 1), sqrt.arity.min);
    try testing.expectEqual(@as(?u8, 1), sqrt.arity.max);
    try testing.expectEqual(@as(usize, 1), sqrt.coercion.fixed.len);
    try testing.expectEqual(CoercionClass.number, sqrt.coercion.at(0));
    try testing.expectEqual(Volatility.stable, sqrt.volatility);
    try testing.expectEqual(value.PropagationClass.propagate, sqrt.propagation);
    try testing.expect(sqrt.liftable());

    const rand = lookup("RAND").?;
    try testing.expectEqualStrings("RAND", rand.name);
    try testing.expectEqual(@as(u8, 0), rand.arity.min);
    try testing.expectEqual(@as(?u8, 0), rand.arity.max);
    try testing.expectEqual(@as(usize, 0), rand.coercion.fixed.len);
    try testing.expectEqual(Volatility.volatile_fn, rand.volatility);
    try testing.expectEqual(value.PropagationClass.propagate, rand.propagation);
    // No slot is nothing to lift, which is also true of PI.
    try testing.expect(!rand.liftable());
    try testing.expect(!lookup("PI").?.liftable());

    // Both are in the frozen file under this row, which is what makes
    // "pinned here" a statement about the inventory rather than a claim.
    var it = inventory();
    var found: usize = 0;
    while (it.next()) |e| {
        if (!isM4d(e)) continue;
        if (std.mem.eql(u8, e.name, "SQRT") or std.mem.eql(u8, e.name, "RAND")) found += 1;
    }
    try testing.expectEqual(@as(usize, 2), found);
}

test "M4d: the multi-argument names are derived from arity, not listed" {
    // §5.3c's error-order gate applies to every name taking more than
    // one argument, and `eval.zig` derives that set from the registry.
    // Here is the same set stated by hand — the two must agree, or one
    // of them is describing a table that no longer exists.
    const expected = [_][]const u8{
        "LOG",   "MOD",       "POWER",   "RANDBETWEEN",
        "ROUND", "ROUNDDOWN", "ROUNDUP", "TRUNC",
    };
    var it = inventory();
    var multi: usize = 0;
    while (it.next()) |e| {
        if (!isM4d(e)) continue;
        const f = lookup(e.name).?;
        if (f.arity.max.? <= 1) continue;
        multi += 1;
        var named = false;
        for (expected) |n| {
            if (std.mem.eql(u8, n, e.name)) named = true;
        }
        if (!named) {
            std.debug.print("multi-argument F1a-2 name not in the list: {s}\n", .{e.name});
            return error.UnlistedMultiArgFunction;
        }
    }
    try testing.expectEqual(expected.len, multi);

    // The two optional-argument rows, which are the only ones whose
    // minimum and maximum differ.
    try testing.expectEqual(@as(u8, 1), lookup("TRUNC").?.arity.min);
    try testing.expectEqual(@as(?u8, 2), lookup("TRUNC").?.arity.max);
    try testing.expectEqual(@as(u8, 1), lookup("LOG").?.arity.min);
    try testing.expectEqual(@as(?u8, 2), lookup("LOG").?.arity.max);
}

// ─── M4e: the F1b batch, against the frozen inventory ────────────

const m4e_milestone = "M4e";
const m4e_batch = "F1b";

fn isM4e(e: InventoryEntry) bool {
    return std.mem.eql(u8, e.milestone, m4e_milestone) and
        std.mem.eql(u8, e.batch, m4e_batch);
}

test "M4e: the batch's size is regenerated from the inventory, never from prose" {
    // The ladder row says "~22". This test does not say 22: it counts
    // the file, and `inventory: per-milestone counts reproduce the
    // ladder` holds the ladder's own number to the same file.
    var it = inventory();
    var counted: usize = 0;
    while (it.next()) |e| {
        if (isM4e(e)) counted += 1;
    }
    try testing.expect(counted > 0);

    // Direction one — no omissions.
    var it2 = inventory();
    var resolved: usize = 0;
    while (it2.next()) |e| {
        if (!isM4e(e)) continue;
        const f = lookup(e.name) orelse {
            std.debug.print("F1b name not registered: {s}\n", .{e.name});
            return error.UnregisteredBatchFunction;
        };
        try testing.expectEqualStrings(e.name, f.name);
        resolved += 1;
    }
    try testing.expectEqual(counted, resolved);

    // Direction two — no substitutions.
    var registered: usize = 0;
    for (&functions) |*f| {
        var it3 = inventory();
        while (it3.next()) |e| {
            if (!std.mem.eql(u8, e.name, f.name)) continue;
            if (std.mem.eql(u8, e.milestone, m4e_milestone)) {
                if (!isM4e(e)) {
                    std.debug.print("{s}: tagged {s} but batch {s}\n", .{ f.name, e.milestone, e.batch });
                    return error.BatchTagMismatch;
                }
                registered += 1;
            }
            break;
        }
    }
    try testing.expectEqual(counted, registered);
}

test "M4e: the Core gate is counted from the file, not read from the ladder" {
    // §7 makes the TSV the count source, and the ladder's "Core gate
    // 59" is the sum of three of its milestones. Counting it here from
    // the file means the ladder's figure and the registry meet at the
    // data: a twenty-third F1b row moves this number, and the only way
    // to keep the prose true is to change the prose.
    const rows = [_][]const u8{ "M4c", "M4d", "M4e" };
    var core: usize = 0;
    for (rows) |m| {
        var it = inventory();
        while (it.next()) |e| {
            if (std.mem.eql(u8, e.milestone, m)) core += 1;
        }
    }
    try testing.expectEqual(@as(usize, 59), core);

    // …and every one of the fifty-nine resolves, which is what makes
    // "the gate closes at M4e" a statement about the registry rather
    // than about three counts that happen to add up.
    var closed: usize = 0;
    for (rows) |m| {
        var it = inventory();
        while (it.next()) |e| {
            if (!std.mem.eql(u8, e.milestone, m)) continue;
            if (lookup(e.name) == null) {
                std.debug.print("Core-gate name does not resolve: {s} ({s})\n", .{ e.name, m });
                return error.CoreGateIncomplete;
            }
            closed += 1;
        }
    }
    try testing.expectEqual(core, closed);
}

test "M4e: every F1b row declares all five fields, and none of them by default" {
    var it = inventory();
    var seen: usize = 0;
    while (it.next()) |e| {
        if (!isM4e(e)) continue;
        const f = lookup(e.name).?;
        seen += 1;

        try testing.expect(f.name.len > 0);
        try testing.expectEqualStrings(e.name, f.name);

        // Arity: a real range, with a repeating slot iff unbounded.
        if (f.arity.max) |max| try testing.expect(f.arity.min <= max);
        if (f.arity.max == null) try testing.expect(f.arity.rest.len > 0);
        // Coercion: the two per-slot tables address the same slots, and
        // a lazy slot lines up with the class that defers. `CHOOSE` is
        // the only row here with one.
        try testing.expectEqual(f.arity.fixed.len, f.coercion.fixed.len);
        try testing.expectEqual(f.arity.rest.len, f.coercion.rest.len);
        for (f.arity.fixed, f.coercion.fixed) |l, c| try testing.expectEqual(l == .lazy, c == .lazy_any);
        for (f.arity.rest, f.coercion.rest) |l, c| try testing.expectEqual(l == .lazy, c == .lazy_any);
        // Volatility: nothing in F1b is redrawn. Stated per row rather
        // than assumed, because `.stable` is also what a forgotten
        // field would look like if the field had a default.
        try testing.expectEqual(Volatility.stable, f.volatility);
        // Propagation: F1b uses three of §5.3c's four classes, and
        // `per_element` — the operators' — is not one of them.
        try testing.expect(f.propagation != .per_element);
        // No member of this batch is lifted elementwise: every one of
        // them consumes a range, an array, or a raw value itself.
        // `da_aware` spilling is M7a's decision, not this row's.
        try testing.expect(!f.liftable());
        try testing.expect(!f.da_aware);
        try testing.expect(!f.reference_producing);
    }
    try testing.expect(seen > 0);
}

test "M4e: the seven framework subjects carry the same five fields as the fifteen the row writes" {
    // SUM, COUNT, COUNTA, COUNTBLANK, COUNTIF, SUMIF and CHOOSE were
    // registered at M3a2 to exercise the framework, not because their
    // row had run (M4c decision 12). M4e pins them where they stand —
    // the same treatment M4d gave SQRT and RAND, and for the same
    // reason: the rows written first are the ones most likely to have
    // been waved through.
    const pinned = [_][]const u8{
        "SUM",     "COUNT", "COUNTA", "COUNTBLANK",
        "COUNTIF", "SUMIF", "CHOOSE",
    };
    for (pinned) |name| {
        const f = lookup(name).?;
        try testing.expectEqualStrings(name, f.name);
        try testing.expect(f.arity.min >= 1);
        try testing.expectEqual(f.arity.fixed.len, f.coercion.fixed.len);
        try testing.expectEqual(f.arity.rest.len, f.coercion.rest.len);
        try testing.expectEqual(Volatility.stable, f.volatility);
        try testing.expect(f.propagation != .per_element);
        // Their table rows have NOT moved: `lookup` finds them where
        // M3a2 put them, and the proof they belong to this batch is the
        // inventory, not their position in the array.
        try testing.expect(inInventory(name));
        var it = inventory();
        var tagged = false;
        while (it.next()) |e| {
            if (std.mem.eql(u8, e.name, name)) tagged = isM4e(e);
        }
        if (!tagged) {
            std.debug.print("{s} is pinned by M4e but the file does not tag it F1b\n", .{name});
            return error.PinnedOutsideBatch;
        }
    }

    // `CHOOSE` is the batch's one deferring form, and the only one of
    // the seven that carries no `impl` — laziness lives in the
    // evaluator because deferring an arm means holding an AST index.
    const choose = lookup("CHOOSE").?;
    try testing.expectEqual(Form.choose_form, choose.form);
    try testing.expect(choose.impl == null);
    try testing.expectEqual(Laziness.lazy, choose.arity.at(1));
    try testing.expectEqual(Laziness.eager, choose.arity.at(0));
}

test "M4e: the batch's propagation classes are the ones §5.3c assigns" {
    // Stated name by name, because §5.3c's whole point is that the
    // class is per function and never per family: `COUNT` and `SUM`
    // are neighbours in the same table and carry different classes,
    // and so do `COUNTA` and `AVERAGE`.
    // `per_function_provenance` is the class for a function whose
    // answer depends on WHERE an error was found, not merely on
    // whether there was one. The counting family is §5.3c's own
    // example; the five lookups join it because an error inside a
    // lookup table is a value the lookup may return rather than an
    // error the lookup becomes.
    const provenance = [_][]const u8{
        "COUNT", "COUNTA",    "COUNTBLANK", "COUNTIF",
        "SUMIF", "AVERAGEIF", "VLOOKUP",    "HLOOKUP",
        "MATCH", "XLOOKUP",   "XMATCH",
    };
    for (provenance) |name| {
        try testing.expectEqual(
            value.PropagationClass.per_function_provenance,
            lookup(name).?.propagation,
        );
    }
    const propagators = [_][]const u8{
        "SUM",   "AVERAGE", "MIN",    "MAX",  "SUMPRODUCT",
        "INDEX", "ROW",     "COLUMN", "ROWS", "COLUMNS",
    };
    for (propagators) |name| {
        try testing.expectEqual(value.PropagationClass.propagate, lookup(name).?.propagation);
    }
    // `CHOOSE` observes: it takes an index and hands back an arm, and
    // an error in an arm it did not take is not its error.
    try testing.expectEqual(value.PropagationClass.observe, lookup("CHOOSE").?.propagation);

    // Together the three lists are the whole batch, so a name added to
    // one and forgotten in another cannot pass.
    var it = inventory();
    var n: usize = 0;
    while (it.next()) |e| {
        if (isM4e(e)) n += 1;
    }
    try testing.expectEqual(n, provenance.len + propagators.len + 1);
}

test "M4e: MIN and MAX are out of the comparator, the five lookups are in it" {
    // §5.4b's one named exception (plan revision 15, change 8). The
    // split is registry data rather than a comment, because the flag is
    // what a later collation change reads to know what it affects.
    for ([_][]const u8{ "MIN", "MAX" }) |name| {
        try testing.expect(!lookup(name).?.collation_sensitive);
    }
    for ([_][]const u8{ "VLOOKUP", "HLOOKUP", "MATCH", "XLOOKUP", "XMATCH" }) |name| {
        try testing.expect(lookup(name).?.collation_sensitive);
    }
    // …and the whole batch, in both directions, so a sixth lookup
    // added later cannot ship uncollated and a third extreme cannot
    // ship collated.
    const sensitive = [_][]const u8{
        "COUNTIF", "SUMIF", "AVERAGEIF", "VLOOKUP",
        "HLOOKUP", "MATCH", "XLOOKUP",   "XMATCH",
    };
    var it = inventory();
    while (it.next()) |e| {
        if (!isM4e(e)) continue;
        var listed = false;
        for (sensitive) |s| {
            if (std.mem.eql(u8, s, e.name)) listed = true;
        }
        if (listed != lookup(e.name).?.collation_sensitive) {
            std.debug.print("{s}: collation_sensitive={} but the list says {}\n", .{
                e.name,
                lookup(e.name).?.collation_sensitive,
                listed,
            });
            return error.CollationFlagMismatch;
        }
    }
}

test "M4e: the multi-argument names are derived from arity, not listed" {
    // The same gate M4d applied, over a batch where most names take
    // more than one argument rather than most taking one. The list is
    // checked against the registry's own arity so a function that
    // gains an argument later cannot slip past unordered.
    const expected = [_][]const u8{
        "AVERAGE", "AVERAGEIF", "CHOOSE", "COUNT",      "COUNTA",
        "COUNTIF", "HLOOKUP",   "INDEX",  "MATCH",      "MAX",
        "MIN",     "SUM",       "SUMIF",  "SUMPRODUCT", "VLOOKUP",
        "XLOOKUP", "XMATCH",
    };
    var it = inventory();
    var multi: usize = 0;
    while (it.next()) |e| {
        if (!isM4e(e)) continue;
        const f = lookup(e.name).?;
        // Unbounded arity is multi-argument by construction.
        if (f.arity.max) |max| {
            if (max <= 1) continue;
        }
        multi += 1;
        var named = false;
        for (expected) |n| {
            if (std.mem.eql(u8, n, e.name)) named = true;
        }
        if (!named) {
            std.debug.print("multi-argument F1b name not in the list: {s}\n", .{e.name});
            return error.UnlistedMultiArgFunction;
        }
    }
    try testing.expectEqual(expected.len, multi);

    // The five single-argument rows: the three position names that take
    // one reference, and `COUNTBLANK`. `ROW`/`COLUMN` reach zero, which
    // is what makes them the registry's only site-dependent rows.
    try testing.expectEqual(@as(u8, 0), lookup("ROW").?.arity.min);
    try testing.expectEqual(@as(?u8, 1), lookup("ROW").?.arity.max);
    try testing.expectEqual(@as(u8, 0), lookup("COLUMN").?.arity.min);
    try testing.expectEqual(@as(?u8, 1), lookup("COLUMN").?.arity.max);
    for ([_][]const u8{ "ROWS", "COLUMNS", "COUNTBLANK" }) |name| {
        try testing.expectEqual(@as(u8, 1), lookup(name).?.arity.min);
        try testing.expectEqual(@as(?u8, 1), lookup(name).?.arity.max);
    }
}

test "registry: the five required fields have no defaults" {
    // "Declares" is enforced by the type, not by review: a field with a
    // default can be omitted, and an omitted propagation class is how
    // `COUNTA` quietly becomes `COUNT`.
    const required = [_][]const u8{ "name", "arity", "coercion", "volatility", "propagation" };
    var found: usize = 0;
    inline for (required) |want| {
        inline for (@typeInfo(Function).@"struct".fields) |f| {
            if (comptime std.mem.eql(u8, f.name, want)) {
                found += 1;
                try testing.expect(f.default_value_ptr == null);
            }
        }
    }
    try testing.expectEqual(required.len, found);

    // …and they are the ONLY fields without one. Checked in this
    // direction too because §7 names five: a sixth mandatory field is a
    // spec change, and a five that quietly became four is the failure
    // this test exists to catch. Both are `@typeInfo` questions, and
    // only asking one of them would leave half the claim untested.
    var mandatory: usize = 0;
    inline for (@typeInfo(Function).@"struct".fields) |f| {
        if (f.default_value_ptr == null) mandatory += 1;
    }
    try testing.expectEqual(required.len, mandatory);
}

test "registry: every entry declares all five fields coherently" {
    try testing.expect(functions.len > 0);
    for (&functions) |*f| {
        try testing.expect(f.name.len > 0);
        try testing.expect(f.arity.min <= f.arity.max orelse 255);
        try testing.expectEqual(f.arity.fixed.len, f.coercion.fixed.len);
        try testing.expectEqual(f.arity.rest.len, f.coercion.rest.len);
        // Volatility and propagation are enums with no "unset" member,
        // so declaring them is all there is to check — but a lazy slot
        // must line up with a `lazy_any` class, or the dispatcher would
        // evaluate an arm the form means to defer.
        for (f.arity.fixed, f.coercion.fixed) |l, c| {
            try testing.expectEqual(l == .lazy, c == .lazy_any);
        }
        for (f.arity.rest, f.coercion.rest) |l, c| {
            try testing.expectEqual(l == .lazy, c == .lazy_any);
        }
        _ = f.volatility;
        _ = f.propagation;
    }
}

test "registry: every implemented function is in the frozen inventory" {
    for (&functions) |*f| {
        if (!inInventory(f.name)) {
            std.debug.print("registered but not frozen: {s}\n", .{f.name});
            return error.UnfrozenFunction;
        }
    }
}

test "registry: lookup is case-insensitive and rejects unknown names" {
    try testing.expect(lookup("sum") != null);
    try testing.expect(lookup("SuM") != null);
    try testing.expectEqualStrings("SUM", lookup("sum").?.name);
    // A name the frozen inventory holds and no row has reached yet.
    // `VLOOKUP` stood here until M4e registered it, which is what this
    // line is for: the example has to be a function that is genuinely
    // still ahead of the ladder, and every batch that lands moves it.
    try testing.expect(lookup("SUMIFS") == null); // frozen, M7b2
    try testing.expect(inInventory("SUMIFS"));
    try testing.expect(lookup("NOTAFUNCTION") == null);
}

test "registry: RAND and RANDBETWEEN are the only volatile rows" {
    // Both directions, because either alone passes for the wrong reason:
    // naming the two proves they are volatile, counting proves nothing
    // else quietly became so. Every volatile row is a cell the M5a2
    // schedule has to re-key, so the set is not a detail.
    const expected = [_][]const u8{ "RAND", "RANDBETWEEN" };
    for (expected) |name| {
        try testing.expectEqual(Volatility.volatile_fn, lookup(name).?.volatility);
    }
    var volatiles: usize = 0;
    for (&functions) |*f| {
        if (f.volatility != .volatile_fn) continue;
        volatiles += 1;
        var named = false;
        for (expected) |n| {
            if (std.mem.eql(u8, n, f.name)) named = true;
        }
        if (!named) {
            std.debug.print("unexpected volatile row: {s}\n", .{f.name});
            return error.UnexpectedVolatileFunction;
        }
    }
    try testing.expectEqual(expected.len, volatiles);
}

test "registry: liftability follows the coercion classes" {
    // Scalar-class signatures lift over arrays; anything holding a range
    // slot is M7a's decision table, not this row's.
    try testing.expect(lookup("SQRT").?.liftable());
    try testing.expect(lookup("NOT").?.liftable());
    try testing.expect(!lookup("SUM").?.liftable());
    try testing.expect(!lookup("COUNTBLANK").?.liftable());
    try testing.expect(!lookup("IF").?.liftable());
    try testing.expect(!lookup("TRUE").?.liftable());
}

test "registry: the four propagation classes are all represented" {
    var seen = std.EnumSet(value.PropagationClass).initEmpty();
    for (&functions) |*f| seen.insert(f.propagation);
    // `per_element` belongs to the elementwise operators rather than to
    // a named function, so three of the four appear in this table.
    try testing.expect(seen.contains(.propagate));
    try testing.expect(seen.contains(.observe));
    try testing.expect(seen.contains(.per_function_provenance));
    try testing.expectEqual(value.PropagationClass.per_function_provenance, lookup("COUNT").?.propagation);
    try testing.expectEqual(value.PropagationClass.per_function_provenance, lookup("COUNTA").?.propagation);
    try testing.expectEqual(value.PropagationClass.propagate, lookup("SUM").?.propagation);
    try testing.expectEqual(value.PropagationClass.observe, lookup("ISERROR").?.propagation);
}

test "registry: arity and coercion address the same slots at any index" {
    const ifs = lookup("IFS").?;
    // The alternating tail is a two-slot cycle, so slot 4 is a condition
    // again — the property a single `rest` slot could not express.
    try testing.expectEqual(Laziness.eager, ifs.arity.at(0));
    try testing.expectEqual(Laziness.eager, ifs.arity.at(5));
    const choose = lookup("CHOOSE").?;
    try testing.expectEqual(Laziness.eager, choose.arity.at(0));
    try testing.expectEqual(Laziness.lazy, choose.arity.at(1));
    try testing.expectEqual(Laziness.lazy, choose.arity.at(9));
    try testing.expectEqual(CoercionClass.lazy_any, choose.coercion.at(9));
}
