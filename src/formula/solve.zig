//! The shared deterministic solver contract (M9c1, `goal_formula.md`
//! §5.5 / the M9c1 ladder row).
//!
//! One Newton driver, used by every function that answers "which rate
//! makes this equation zero" — `RATE` at this row, the `IRR`/`XIRR`
//! pair at M9c2. It ships *before* the first consumer because the
//! contract is the deliverable: two functions with two hand-rolled
//! loops would disagree about iteration caps, convergence, domain
//! guards, or what an iteration costs, and each disagreement would be
//! an answer that moves when an implementation is touched.
//!
//! The contract, pinned:
//!
//!   - **At most 128 iterations.** Excel documents 20 for `RATE` at a
//!     tolerance of 1e-7; Newton doubles its correct digits per
//!     iteration near a root, so 128 at the tighter tolerance below is
//!     generous for everything Excel converges on, and a hard number a
//!     fixture can count.
//!   - **Tolerance 1e-10 on the step.** `|Δx| ≤ 1e-10` is convergence.
//!     Tighter than Excel's documented 1e-7 — a *recorded* divergence
//!     in the conservative direction: anything that converges under
//!     Excel's test converges under this one, to more digits.
//!   - **Root selection is the Newton path from the guess.** No
//!     bracketing, no restart sweep: the root found is the one the
//!     caller's guess descends to, which is Excel's own contract
//!     ("if RATE does not converge, try different values for guess").
//!   - **A step clamped at the domain edge never converges.** The
//!     domain is an open half-line (`x > domain_min`; −1 for a rate,
//!     where `(1+r)^n` lives). A raw step that lands outside is pulled
//!     to the midpoint between the iterate and the boundary — and that
//!     iteration is disqualified from the convergence test, because the
//!     equation asked to leave the domain and a shrinking leash is not
//!     a shrinking step. A rootless residual that chases the boundary
//!     (all-positive cashflows) therefore runs its 128 iterations and
//!     answers `NoConvergence`, never a fake root: without the rule,
//!     thirty-some halvings shrink `|Δx|` below any tolerance while
//!     the residual still stands at its plateau.
//!   - **`#NUM!` on domain or convergence failure** — both surface as
//!     errors here and the registry maps them onto the same Excel
//!     answer; they stay distinct errors because a *fixture* wants to
//!     say which one it is pinning.
//!   - **Every iteration polls, then charges 4 units** to the shared
//!     `WorkBudget` when the caller wired one. The poll comes first —
//!     a token set mid-solve is observed before the iteration's work,
//!     within one iteration rather than within one 65 536-unit stride
//!     (§5.5's bound, met with five orders of magnitude to spare).
//!     Exhaustion (`LimitExceeded`) and cancellation (`Cancelled`) are
//!     refusals that abort the run; `NoConvergence` is a *value* the
//!     caller turns into `#NUM!` — the M5b-round split between the
//!     semantic outcome and the resource ceiling, held here too.
//!
//! Determinism: the loop reads no clock, draws nothing, and allocates
//! nothing; equal `(residual, guess, domain_min)` walk equal iterates.
//! The budget changes *whether* the walk completes, never where it
//! goes — and the poller is cancellation, outside identity like every
//! other cancel seam.

const std = @import("std");

const run_inputs = @import("run_inputs.zig");

pub const max_iterations: u32 = 128;
pub const tolerance: f64 = 1e-10;

/// The two solver outcomes that are not an answer, beside the two
/// budget refusals. `BadDomain` covers a non-finite residual or
/// derivative, a zero derivative (Newton has no step), and a guess
/// already outside the domain — everything where the *equation* stopped
/// cooperating; `NoConvergence` is 128 honest iterations without a
/// converging step.
pub const Error = error{ BadDomain, NoConvergence } || run_inputs.WorkBudgetError;

/// A residual and its derivative at one point — what one iteration
/// evaluates, and the unit the 4-unit charge prices.
pub const Fdf = struct { f: f64, df: f64 };

/// Newton–Raphson under the contract above. `ctx`/`fdf` rather than a
/// closure because Zig has none: the residual's coefficients live in
/// the caller's frame and arrive by pointer.
pub fn newton(
    ctx: anytype,
    comptime fdf: fn (@TypeOf(ctx), f64) Fdf,
    guess: f64,
    domain_min: f64,
    work: ?*run_inputs.WorkBudget,
) Error!f64 {
    if (!std.math.isFinite(guess) or guess <= domain_min) return Error.BadDomain;

    var x = guess;
    var iter: u32 = 0;
    while (iter < max_iterations) : (iter += 1) {
        if (work) |w| {
            try w.poll();
            try w.charge(run_inputs.WorkBudget.solver_iteration_units);
        }
        const e = fdf(ctx, x);
        if (!std.math.isFinite(e.f) or !std.math.isFinite(e.df) or e.df == 0) {
            return Error.BadDomain;
        }
        var next = x - e.f / e.df;
        if (!std.math.isFinite(next)) return Error.BadDomain;
        var clamped = false;
        if (next <= domain_min) {
            next = (x + domain_min) / 2;
            clamped = true;
        }
        if (!clamped and @abs(next - x) <= tolerance) return next;
        x = next;
    }
    return Error.NoConvergence;
}

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

const Quadratic = struct {
    // f(x) = x² − c, root at √c.
    c: f64,
    fn eval(self: *const Quadratic, x: f64) Fdf {
        return .{ .f = x * x - self.c, .df = 2 * x };
    }
};

test "newton: converges to the root the guess descends to" {
    var q: Quadratic = .{ .c = 2 };
    const r = try newton(&q, Quadratic.eval, 1.0, -std.math.inf(f64), null);
    try testing.expectApproxEqAbs(std.math.sqrt2, r, 1e-12);
    // The other basin: a negative guess selects the negative root —
    // root selection IS the Newton path, stated as a test.
    const l = try newton(&q, Quadratic.eval, -1.0, -std.math.inf(f64), null);
    try testing.expectApproxEqAbs(-std.math.sqrt2, l, 1e-12);
}

const Drifter = struct {
    // f ≡ 1 with slope 1: every step walks −1 and no step ever
    // shrinks. The clean picture of "no root, honest iterations".
    fn eval(_: *const Drifter, _: f64) Fdf {
        return .{ .f = 1, .df = 1 };
    }
};

test "newton: a rootless residual runs its 128 iterations and refuses" {
    var d: Drifter = .{};
    var w: run_inputs.WorkBudget = .{};
    try testing.expectError(
        Error.NoConvergence,
        newton(&d, Drifter.eval, 0.0, -std.math.inf(f64), &w),
    );
    // Exactly 128 iterations, 4 units each — the budget is the
    // iteration count's receipt.
    try testing.expectEqual(
        128 * run_inputs.WorkBudget.solver_iteration_units,
        w.used,
    );
}

test "newton: a clamped step is disqualified from convergence" {
    // f(x) = x + 2 with slope 1 wants x = −2, outside an open domain
    // at −1. Every raw step lands below the boundary, every clamped
    // step halves toward it, and |Δx| is under any tolerance within
    // ~40 halvings — the contract's rule is what keeps this from
    // "converging" to a point where f = 1.
    const Wall = struct {
        fn eval(_: *const @This(), x: f64) Fdf {
            return .{ .f = x + 2, .df = 1 };
        }
    };
    var s: Wall = .{};
    try testing.expectError(
        Error.NoConvergence,
        newton(&s, Wall.eval, 0.1, -1.0, null),
    );
}

test "newton: domain failures name themselves" {
    var q: Quadratic = .{ .c = 2 };
    // A guess on or below the boundary never iterates.
    try testing.expectError(Error.BadDomain, newton(&q, Quadratic.eval, -1.0, -1.0, null));
    // A zero derivative has no Newton step: x = 0 on x² − c.
    try testing.expectError(Error.BadDomain, newton(&q, Quadratic.eval, 0.0, -std.math.inf(f64), null));
}

test "newton: the budget draws 4 units per iteration and exhaustion aborts mid-solve" {
    var d: Drifter = .{};
    // Room for exactly three iterations; the fourth charge refuses.
    var w: run_inputs.WorkBudget = .{ .limit = 3 * run_inputs.WorkBudget.solver_iteration_units };
    try testing.expectError(
        Error.LimitExceeded,
        newton(&d, Drifter.eval, 0.0, -std.math.inf(f64), &w),
    );
    try testing.expectEqual(3 * run_inputs.WorkBudget.solver_iteration_units, w.used);
    try testing.expect(w.tripped);
}

test "newton: a cancel token is observed before the iteration's work" {
    const TripAt = struct {
        n: usize = 0,
        at: usize,
        fn check(ctx: ?*const anyopaque) error{Cancelled}!void {
            const self: *const @This() = @ptrCast(@alignCast(ctx.?));
            @constCast(self).n += 1;
            if (self.n >= self.at) return error.Cancelled;
        }
    };
    var d: Drifter = .{};
    var trip: TripAt = .{ .at = 6 };
    var w: run_inputs.WorkBudget = .{ .poller = .{ .ctx = &trip, .check_fn = TripAt.check } };
    try testing.expectError(
        Error.Cancelled,
        newton(&d, Drifter.eval, 0.0, -std.math.inf(f64), &w),
    );
    // Five iterations of work were charged; the sixth poll refused
    // before its iteration spent anything — mid-solve, observed before
    // commit, and the meter shows exactly where.
    try testing.expectEqual(5 * run_inputs.WorkBudget.solver_iteration_units, w.used);
    try testing.expectEqual(@as(usize, 6), trip.n);
}
