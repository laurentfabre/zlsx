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
//! labelled `spec_pinned` where none does. The functions the framework
//! borrowed from later rows (SUM, COUNT, COUNTA, COUNTBLANK, COUNTIF,
//! SUMIF, SQRT, RAND, CHOOSE) are still M4d/M4e's to pin; they are in
//! the table because the framework needed them, not because their row
//! has run.

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
    try testing.expect(lookup("VLOOKUP") == null); // frozen, not yet implemented
    try testing.expect(lookup("NOTAFUNCTION") == null);
}

test "registry: the volatile set is exactly what the fixtures rely on" {
    var volatiles: usize = 0;
    for (&functions) |*f| {
        if (f.volatility == .volatile_fn) volatiles += 1;
    }
    try testing.expectEqual(@as(usize, 1), volatiles);
    try testing.expectEqual(Volatility.volatile_fn, lookup("RAND").?.volatility);
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
