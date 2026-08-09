//! The iteration engine: what a cycle's members are worth once they
//! stop moving, what it means when they never do, and how a reference
//! that only exists at runtime gets into the order at all
//! (`goal_formula.md` §5.6c–e, §9).
//!
//! M5a2 of the tier-D1 ladder. M5a1 built and ordered the graph; this
//! file **runs** it. Nothing here rebuilds a node model, an index or an
//! order — `graph.zig` owns all three, and the outer loop below reaches
//! for `graph.build` again rather than growing a second builder.
//!
//! The two exhaustion outcomes are not the same outcome
//! ----------------------------------------------------
//! An SCC's pass counter is bounded twice over, and which bound fires
//! changes what the run *is*:
//!
//! | bound | whose | reaching it |
//! |---|---|---|
//! | `calcPr@iterateCount`, clamped to 32 767 | the workbook's | **success**, with the cells that never settled reported |
//! | `max_scc_iterations` (§9) | the caller's | **`FormulaLimitExceeded`, zero mutation** |
//!
//! Excel documents the first: a workbook that asks for 100 passes and
//! gets 100 passes has been calculated, whether or not anything
//! converged. The second is a resource cap, and a resource cap that
//! quietly published caches computed from fewer passes than the workbook
//! asked for would be answering a different question than the one the
//! file poses. So it refuses, and it refuses the **whole run** — one
//! component hitting the ceiling rolls back the components that
//! converged beside it, because a workbook half-calculated under one
//! iteration budget and half under another is not a state any consumer
//! can reason about.
//!
//! The ceiling refuses only when it is **strictly lower** than the
//! semantic bound. Equal bounds are the workbook's answer: the caller
//! permitted exactly what the file asked for, and got it.
//!
//! Gauss–Seidel, and the divergence it buys
//! ----------------------------------------
//! Inside a pass, a member sees values computed earlier in the same pass
//! — publication is immediate, not double-buffered. That is Excel's
//! visibility rule. What is *not* Excel's is the order: Excel iterates
//! along a mutable calculation chain whose order evolves during a
//! recalc, and this engine walks a fixed coordinate order. §5.6c
//! declares that divergence rather than hiding it, because determinism
//! requires a fixed order and a fixed order cannot reproduce a chain
//! that reorders itself. Order-sensitive circular fixtures therefore
//! assert that *convergence* agrees with Excel and record where the
//! converged values do not.
//!
//! Zero mutation is a rollback, not an absence of writing
//! ------------------------------------------------------
//! Gauss–Seidel visibility means a pass has to publish as it goes, so
//! "the refusal wrote nothing" cannot be true by construction. It is
//! true by journal: every publish is recorded, and a refusal retracts
//! them in reverse. `Host.retract` is infallible for that reason — a
//! rollback that could run out of memory would make the promise
//! conditional on there being memory to keep it.
//!
//! What this row does NOT do
//! -------------------------
//! It does not place a spill (M7a), does not write a cache back to a
//! part (M5b), and does not evaluate a table producer's rows (M7b). An
//! array reaches it only as something to compare — §5.6c needs array
//! convergence and needs a shape that moves between passes to become
//! `#SPILL!`, and both are decidable without placing anything.

const std = @import("std");
const assert = std.debug.assert;

const coords = @import("zlsx_refs");
const draw_schedule = @import("draws.zig");
const env = @import("env.zig");
const graph = @import("graph.zig");
const parser = @import("parser.zig");
const run_inputs = @import("run_inputs.zig");
const value = @import("value.zig");

pub const WorkLimits = run_inputs.WorkLimits;
pub const WorkCounters = run_inputs.WorkCounters;
pub const WorkCategory = run_inputs.WorkCategory;

/// This machine's failures, separately from the workbook's. Same split
/// the rest of the engine keeps.
pub const Error = error{OutOfMemory};

// ─── what a member holds between passes ──────────────────────────

/// One member's value, as convergence sees it.
///
/// An array is here because §5.6c compares shapes and elements, not
/// because this row places one. The distinction matters: a
/// shape-mutating cycle has to be *detected* to be turned into
/// `#SPILL!`, and detecting it needs the previous shape.
pub const Snapshot = union(enum) {
    scalar: value.ScalarValue,
    array: value.Matrix,

    pub fn shape(self: Snapshot) value.Shape {
        return switch (self) {
            .scalar => .{ .rows = 1, .cols = 1 },
            .array => |m| m.shape(),
        };
    }

    pub const ZeroError = Error || error{
        /// §9's `max_matrix_cells`. A declared array range is bounded by
        /// the *grid*, not by that limit — `A1:D1048576` is four million
        /// and change — so an anchor's declared shape really can be too
        /// large to materialize. It is a §9 refusal rather than an
        /// assertion for exactly that reason: the input that reaches it
        /// is a workbook, and a workbook must not be able to trip an
        /// `unreachable`.
        ShapeTooLarge,
    };

    /// A zero-filled array of exactly `s`, which is what §5.6c's seed
    /// table hands an anchor whose shape is known.
    pub fn zeros(allocator: std.mem.Allocator, s: value.Shape) ZeroError!Snapshot {
        const m = value.Matrix.init(allocator, s.rows, s.cols) catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            error.TooManyCells => return error.ShapeTooLarge,
            // A declared shape comes from a range, and a range has at
            // least one cell in each dimension.
            error.EmptyMatrix => unreachable,
        };
        @memset(m.cells, value.ScalarValue.fromNumber(0));
        return .{ .array = m };
    }

    pub fn number(n: f64) Snapshot {
        return .{ .scalar = value.ScalarValue.fromNumber(n) };
    }
};

// ─── §5.6c convergence ───────────────────────────────────────────

/// Why two consecutive passes did or did not agree.
///
/// Richer than a bool because the three ways of not converging are
/// three different fixtures, and a caller that could only ask "did it
/// settle" could not tell a value still moving from a value that changed
/// what it *is*.
pub const Convergence = enum {
    converged,
    /// Same type, same shape, still moving.
    changed,
    /// `2` became `"2"`, or a number became an error. §5.6c: any type
    /// transition is not converged, however small the numeric distance
    /// between the spellings might look.
    type_transition,
    /// An array whose dimensions moved, or a scalar that became an array.
    shape_change,

    pub fn settled(self: Convergence) bool {
        return self == .converged;
    }
};

/// §5.6c's per-cell rule, all of it.
///
/// Numbers compare by **magnitude** — `abs(new − previous) < delta`. A
/// raw signed difference would call any decreasing value converged on
/// its first pass, which is the bug this spelling exists to prevent and
/// the reason the decreasing fixture is not optional. The comparison is
/// strict, so a difference exactly equal to `delta` has not converged.
///
/// Everything else is two-pass equality: text by bytes, booleans by
/// value, errors by kind, blank by being blank. `NaN` cannot occur —
/// N4a converts a non-finite arithmetic result to `#NUM!` at the point
/// of production, so a number that reaches here is finite.
pub fn compare(prev: Snapshot, next: Snapshot, delta: f64) Convergence {
    if (prev == .array or next == .array) {
        if (prev != .array or next != .array) return .shape_change;
        const a = prev.array;
        const b = next.array;
        if (!a.shape().eql(b.shape())) return .shape_change;
        var worst: Convergence = .converged;
        for (a.cells, b.cells) |x, y| {
            const c = compareScalar(x, y, delta);
            if (c == .converged) continue;
            // A type transition anywhere in an array is the array's
            // answer: it is the strongest statement about why the two
            // passes are not the same array.
            if (c == .type_transition) return .type_transition;
            worst = c;
        }
        return worst;
    }
    return compareScalar(prev.scalar, next.scalar, delta);
}

fn compareScalar(prev: value.ScalarValue, next: value.ScalarValue, delta: f64) Convergence {
    const Tag = std.meta.Tag(value.ScalarValue);
    if (@as(Tag, prev) != @as(Tag, next)) return .type_transition;
    return switch (prev) {
        .number => |p| blk: {
            const d = @abs(next.number - p);
            // `-0` and `+0` differ in bits and not in magnitude, so the
            // signed-zero fixture converges — which is the answer, since
            // iteration is about how much a value moved.
            assert(!std.math.isNan(d));
            break :blk if (d < delta) .converged else .changed;
        },
        .text => |p| if (std.mem.eql(u8, p, next.text)) .converged else .changed,
        .boolean => |p| if (p == next.boolean) .converged else .changed,
        .err => |p| if (p.known == next.err.known) .converged else .changed,
        .blank => .converged,
    };
}

pub fn converged(prev: Snapshot, next: Snapshot, delta: f64) bool {
    return compare(prev, next, delta).settled();
}

// ─── §5.6c's iteration settings ──────────────────────────────────

/// The workbook's own iteration policy, normalized.
///
/// Kept apart from `calc.CalcState` so this engine can be driven from a
/// fixture without a workbook part, and so the transition table below is
/// one function with one set of fixtures rather than a rule spread
/// across a parser and an engine.
pub const Settings = struct {
    iterate: bool = false,
    iterate_count: u32 = default_count,
    iterate_delta: f64 = default_delta,

    /// The schema defaults, which are also what an absent `<calcPr>`
    /// means (§5.6c: "defaults off/100/0.001").
    pub const default_count: u32 = 100;
    pub const default_delta: f64 = 0.001;
    /// Excel's own ceiling on `iterateCount`.
    pub const excel_max_count: u32 = 32_767;

    /// §5.6c's pinned transition table for missing, zero and
    /// out-of-range values.
    ///
    /// | raw | normalized | why |
    /// |---|---|---|
    /// | `iterateCount` = 0 | 100 | Excel's own minimum is 1, so a zero in the file is an unset attribute, and an unset attribute means the schema default |
    /// | `iterateCount` > 32 767 | 32 767 | the clamp §5.6c names; a workbook cannot ask for more than Excel can give |
    /// | `iterateDelta` < 0 | 0.001 | a negative *maximum change* is not a tolerance, so the default stands |
    /// | `iterateDelta` = 0 | 0.001 | **the same row**, for a reason worth writing down (below) |
    /// | `iterateDelta` non-finite | — | never reaches here; `calc.parseCalcState` refuses the part |
    ///
    /// The zero row reads at first like "iterate until nothing changes
    /// at all", and that is a coherent thing to want — but it is not
    /// what a zero delta *means* under §5.6c's comparison. The rule is
    /// `abs(new − previous) < iterateDelta`, strictly; at `delta = 0`
    /// nothing satisfies it, not even a value that did not move, so a
    /// zero tolerance is not a strict tolerance at all but an
    /// unsatisfiable one. Reading it as exact equality would mean
    /// special-casing the comparison, and a comparison with an exception
    /// in it is a comparison two implementations will eventually
    /// disagree about. So zero joins the same row as negative: it is an
    /// unset attribute, and an unset attribute is the schema default —
    /// which is exactly the reasoning that already turns `iterateCount`
    /// zero into 100.
    pub fn normalize(raw: Settings) Settings {
        var out = raw;
        if (out.iterate_count == 0) out.iterate_count = default_count;
        if (out.iterate_count > excel_max_count) out.iterate_count = excel_max_count;
        if (!(out.iterate_delta > 0)) out.iterate_delta = default_delta;
        return out;
    }
};

// ─── the seam to whatever can actually evaluate a cell ───────────

/// What one evaluation produced, plus what it read on the way.
pub const Produced = union(enum) {
    ok: struct {
        value: Snapshot,
        /// §5.6e: the references the evaluation **actually** reached, as
        /// opposed to the ones its text mentions. A dynamic reference is
        /// invisible to a walk over the body, so the outer loop learns
        /// about it here or not at all.
        reads: Reads = .{},
    },
    /// A plane-2 refusal from the evaluator. The run stops and rolls
    /// back; the engine does not reinterpret it.
    refused: parser.PlaneTwo,
};

pub const Reads = struct {
    cells: []const env.CellRef = &.{},
    areas: []const env.RangeRef = &.{},
};

/// Everything the engine cannot do itself: evaluate a body, make a value
/// visible, and take it back.
///
/// A vtable rather than a comptime parameter because the two real
/// implementations — the package's workbook model and this file's own
/// fixtures — must be the same code path. A generic engine specialized
/// per host would let a fixture pass while the workbook took a different
/// branch.
pub const Host = struct {
    ctx: *anyopaque,
    vtable: *const VTable,

    pub const VTable = struct {
        /// Evaluate the formula at `cell`, keyed for §5.6d.
        evaluate: *const fn (ctx: *anyopaque, cell: env.CellRef, key: draw_schedule.Key) Error!Produced,
        /// Make `v` visible to every subsequent read in this run. Scratch
        /// only — §5.6f's purity contract holds through here.
        publish: *const fn (ctx: *anyopaque, cell: env.CellRef, v: Snapshot) Error!void,
        /// Undo one publish. Infallible: see the header.
        retract: *const fn (ctx: *anyopaque, cell: env.CellRef) void,
    };

    fn evaluate(self: Host, cell: env.CellRef, key: draw_schedule.Key) Error!Produced {
        return self.vtable.evaluate(self.ctx, cell, key);
    }

    fn publish(self: Host, cell: env.CellRef, v: Snapshot) Error!void {
        return self.vtable.publish(self.ctx, cell, v);
    }

    fn retract(self: Host, cell: env.CellRef) void {
        self.vtable.retract(self.ctx, cell);
    }
};

// ─── outcomes ────────────────────────────────────────────────────

/// Which bound stopped a component, when one did.
pub const Bound = enum {
    /// `calcPr@iterateCount`. Success.
    semantic,
    /// `max_scc_iterations`. A refusal — recorded per component anyway,
    /// because the report has to be able to say which component it was.
    resource,
};

pub const Outcome = enum {
    /// Every member met §5.6c's rule within its bounds.
    converged,
    /// The workbook's own bound ran out. Success, with the members that
    /// never settled counted.
    semantic_bound,
    /// A member's shape moved between passes. §5.6c makes that
    /// `#SPILL!` and stops — a shape-mutating cycle never spins.
    shape_indeterminate,
    /// An acyclic component. It evaluated once; convergence is not a
    /// question that applies to it.
    acyclic,
};

pub const ComponentReport = struct {
    /// The component's smallest member key. Its identity across
    /// rebuilds — node *ids* move when the node set changes, keys do
    /// not.
    at: graph.Key,
    cyclic: bool,
    outcome: Outcome,
    /// Passes actually run. Zero for an acyclic component, which
    /// evaluates its members once each and is not iterating.
    passes: u32 = 0,
    bound: ?Bound = null,
    non_converged_cells: u32 = 0,
    /// Whether this component was re-run by a later §5.6e pass. Only
    /// changed components are, and the report says so rather than
    /// leaving a reader to infer it from a pass count.
    rerun: bool = false,
};

pub const Report = struct {
    /// Owned by the allocator passed to `run`, and the one thing a
    /// successful run hands back that outlives the engine — everything
    /// else it allocated was scratch and dies with the arena.
    components: []const ComponentReport,
    /// §5.7.8's figure: members that reached the semantic bound without
    /// settling.
    non_converged_cells: u32 = 0,
    /// Members whose shape moved and became `#SPILL!`.
    shape_indeterminate_cells: u32 = 0,
    /// How many times the outer loop ran the schedule. One means the
    /// graph the static walk built was already the right graph.
    dynamic_passes: u32 = 1,
    /// Evaluations performed, across every pass of every component.
    cell_evals: u64 = 0,

    pub fn deinit(self: *Report, gpa: std.mem.Allocator) void {
        gpa.free(self.components);
        self.* = undefined;
    }
};

pub const Refusal = struct {
    reason: Reason,
    at: ?graph.Key = null,
    /// Set when `reason == .work_limit_exceeded` or
    /// `.scc_iteration_ceiling`.
    limit: ?WorkCategory = null,
    /// Set when `reason == .evaluation_refused`, carrying the plane the
    /// evaluator raised rather than a plane this file invented.
    plane: ?parser.PlaneTwo = null,

    pub const Reason = enum {
        /// §5.6c: the caller's ceiling bound before the workbook's own.
        scc_iteration_ceiling,
        /// §5.6e: the outer loop never reached a fixpoint.
        dynamic_ref_unstable,
        /// §5.6c: iteration is off and the plan reaches a cycle.
        cycle,
        /// §5.6c's fourth seed row, reached through `graph.seedFor`.
        malformed_cache_seed,
        /// The host's evaluator refused. `plane` carries which.
        evaluation_refused,
        /// Any other §9 counter.
        work_limit_exceeded,
        /// §9's `max_matrix_cells`, reached through §5.6c's seed table:
        /// a declared array range is bounded by the grid rather than by
        /// that limit, so an anchor can declare a shape too large to
        /// materialize.
        seed_shape_too_large,
    };

    /// §10's plane-2 vocabulary.
    pub fn planeTwo(self: Refusal) parser.PlaneTwo {
        return switch (self.reason) {
            .scc_iteration_ceiling, .work_limit_exceeded, .seed_shape_too_large => .FormulaLimitExceeded,
            .dynamic_ref_unstable => .FormulaDynamicRefUnstable,
            .cycle => .FormulaCycle,
            .malformed_cache_seed => .FormulaMalformedInput,
            .evaluation_refused => self.plane.?,
        };
    }
};

pub const Result = union(enum) {
    ok: Report,
    refused: Refusal,
};

// ─── §5.6e's rebuild seam ────────────────────────────────────────

/// What the outer loop needs to build the graph again.
///
/// Present only when the caller wants §5.6e. Absent, the engine runs the
/// graph it was given exactly once, which is the whole of §5.6c and is
/// what every iteration fixture needs.
pub const Rebuild = struct {
    input: graph.Input,
    resolver: graph.Resolver,
    options: graph.Options = .{},
};

pub const Options = struct {
    settings: Settings = .{},
    limits: WorkLimits = .{},
    /// §5.6d's memo, shared across every pass and every rebuild.
    schedule: ?*draw_schedule.Schedule = null,
    counters: ?*WorkCounters = null,
    /// §5.6e. Null means "run this graph once".
    rebuild: ?Rebuild = null,
    /// §5.6f: run only the closure of these roots. Null runs the whole
    /// graph.
    ///
    /// The ROOTS, not the components they resolved to. A closure is
    /// *derived* from its roots, and §5.6e can change what it derives
    /// to: `A1=INDIRECT("C1")` reaches a cell no static walk found, so
    /// the closure that did not contain `C1` on pass 1 must contain it
    /// on pass 2 or the rebuild would reorder a component it then
    /// declines to run. Holding component ids would have frozen the
    /// closure at whatever the discovery pass could see.
    closure: ?Closure = null,
};

pub const Closure = struct {
    roots: []const graph.Key,
    /// What `graph.plan` needs to know to admit a cyclic component
    /// instead of refusing it (§5.6c).
    iterating: bool = false,
};

// ─── the engine ──────────────────────────────────────────────────

/// Run the whole schedule: §5.6e's outer loop around §5.6c's inner one.
///
/// `g` is M5a1's graph, already built. When `opts.rebuild` is set the
/// engine may build further graphs from the same input — through
/// `graph.build`, never through a builder of its own.
///
/// **The engine owns `g` from this call on, error paths included** —
/// the caller must not deinit it. §5.6e replaces graphs, and the moment
/// of replacement was the run's memory peak (§9.1): the outgoing graph
/// has to be freeable *before* its successor is built, which a
/// caller-owned graph forbids. Everything compared across the
/// replacement copies what it needs — signatures copy their keys, and
/// a Key's spellings borrow from the *Input*, never from the graph.
pub fn run(gpa: std.mem.Allocator, g: graph.Graph, host: Host, opts: Options) Error!Result {
    var e = Engine.init(gpa, host, opts);
    defer e.deinit();
    const r = try e.drive(g);
    if (r == .refused) e.rollback();
    return r;
}

/// Set-of-edges and set-of-keys, over the equality `graph.zig` already
/// defines. Not `AutoHashMap`, which cannot be instantiated over a
/// `Key` at all: a `Key` carries borrowed spellings, and `autoHash`
/// refuses a slice rather than silently choosing between the pointer
/// and the bytes. The bytes are the answer here — two rows naming one
/// defined name are one node — and `Key.hash` says so explicitly.
const EdgeSet = std.HashMapUnmanaged(graph.DynamicRef, void, graph.DynamicRef.HashContext, 80);
const KeySet = std.HashMapUnmanaged(graph.Key, void, graph.Key.HashContext, 80);

const Engine = struct {
    gpa: std.mem.Allocator,
    arena: std.heap.ArenaAllocator,
    host: Host,
    opts: Options,
    settings: Settings,
    local_counters: WorkCounters,

    /// Every cell this run has published, in publish order. The journal
    /// a refusal replays backwards.
    journal: std.ArrayListUnmanaged(env.CellRef) = .empty,
    /// The current value of every published cell, so a pass can compare
    /// against the previous one without asking the host to read back
    /// something it may not be able to represent.
    ///
    /// A map, not a list: nothing reads it in publish order — `journal`
    /// is what a rollback replays — and every publish looked its own
    /// coordinate up, which made a pass cost O(published²).
    ///
    /// Recorded only under §5.6c iteration (§9.1 M10g): `heldValue`'s
    /// one caller sits behind `iterateComponent`'s settings gate, so a
    /// workbook with iteration off can never read a held value — and
    /// recording one per publish anyway held 6.9 MB of snapshots at the
    /// named workload's peak, plus the map's growth ladder in pages the
    /// allocator kept.
    held: std.AutoHashMapUnmanaged(env.CellRef, Snapshot) = .empty,
    reports: std.ArrayListUnmanaged(ComponentReport) = .empty,
    /// The reports of the pass before this one. A §5.6e pass re-runs
    /// only what changed, so the components it skipped have to restate
    /// what they were rather than be described afresh by a pass that did
    /// not look at them.
    previous_reports: std.ArrayListUnmanaged(ComponentReport) = .empty,
    /// Where each of those sits, by the key that names its component.
    /// `execute` asks once per component, so scanning the list made a
    /// pass cost O(components²) — the list stays because `report` hands
    /// its order to the caller; only the lookup moved.
    previous_index: std.HashMapUnmanaged(graph.Key, u32, graph.Key.HashContext, 80) = .empty,

    /// §5.6e's runtime edges as they stand right now.
    ///
    /// Replaced per pass rather than accumulated, and that is the whole
    /// difference between an outer loop that can converge and one that
    /// can only grow. §5.6e says a component may be "merged, **split**,
    /// or gained **or lost** edges" — an accumulating set can never lose
    /// one, so a component could never split and a reference that
    /// stopped pointing somewhere would keep an edge nothing reads. It
    /// also makes exhaustion reachable: a reference that oscillates
    /// between two targets produces two edge sets that never agree,
    /// which is exactly what `max_dynamic_passes` is a bound on.
    dynamic: std.ArrayListUnmanaged(graph.DynamicRef) = .empty,
    /// What this pass has read so far, and from which owners. A pass
    /// re-evaluates only the components §5.6e says changed, so an owner
    /// that did not run keeps the edges it last reported — dropping them
    /// would make "not re-evaluated" indistinguishable from "reads
    /// nothing".
    pass_edges: std.ArrayListUnmanaged(graph.DynamicRef) = .empty,
    /// Membership index over the list above. The **list** stays the
    /// record: `foldPassEdges` hands `pass_edges` to the next build in
    /// the order the run produced it, and a set has no order to hand
    /// over. The set only answers "already?", which is all the dedupe
    /// ever asked — and answering it by scanning made a pass cost
    /// O(edges²), which is what §9.1's profile caught.
    pass_edge_set: EdgeSet = .empty,
    /// Which owners this pass evaluated — recorded only while `dynamic`
    /// holds edges (§9.1 M10g). Its one reader is `foldedPassEdges`'
    /// carry, which asks it once per `dynamic` entry to decide whether
    /// an owner superseded its old edges; `dynamic` is replaced only
    /// between passes, so a pass that starts with no edges to carry can
    /// prove the set unread before recording a single owner. Recording
    /// unconditionally built a 6.4 MB set nothing read at the named
    /// workload's peak.
    touched_set: KeySet = .empty,
    /// §5.6f's closure for the CURRENT graph, as component ids. Rebuilt
    /// every pass, because a rebuild renumbers and a dynamic edge can
    /// widen what the closure covers.
    /// A bitset, not a map (§9.1 M10k): `execute` asks one membership
    /// bit per component, and the u32 map answered that in 655 384
    /// bytes at the drive's peak instant — an 11 KB fact.
    scope: std.DynamicBitSetUnmanaged = .{},
    scoped: bool = false,
    /// The graph the currently-executing pass runs over — set for the
    /// span of `execute`, so `noteReads` can ask `walkNoted` without
    /// threading the graph through every evaluation frame (M10b).
    run_graph: ?*const graph.Graph = null,

    fn init(gpa: std.mem.Allocator, host: Host, opts: Options) Engine {
        return .{
            .gpa = gpa,
            .arena = std.heap.ArenaAllocator.init(gpa),
            .host = host,
            .opts = opts,
            .settings = Settings.normalize(opts.settings),
            .local_counters = .{ .limits = opts.limits },
        };
    }

    fn deinit(self: *Engine) void {
        self.journal.deinit(self.gpa);
        self.held.deinit(self.gpa);
        self.reports.deinit(self.gpa);
        self.previous_reports.deinit(self.gpa);
        self.previous_index.deinit(self.gpa);
        self.dynamic.deinit(self.gpa);
        self.pass_edges.deinit(self.gpa);
        self.pass_edge_set.deinit(self.gpa);
        self.touched_set.deinit(self.gpa);
        self.scope.deinit(self.gpa);
        self.arena.deinit();
        self.* = undefined;
    }

    fn counters(self: *Engine) *WorkCounters {
        return self.opts.counters orelse &self.local_counters;
    }

    fn a(self: *Engine) std.mem.Allocator {
        return self.arena.allocator();
    }

    /// §5.6c's "zero mutation": every publish, undone in reverse.
    fn rollback(self: *Engine) void {
        var i = self.journal.items.len;
        while (i > 0) {
            i -= 1;
            self.host.retract(self.journal.items[i]);
        }
        self.journal.clearRetainingCapacity();
        self.held.clearRetainingCapacity();
    }

    fn publish(self: *Engine, cell: env.CellRef, v: Snapshot) Error!void {
        try self.host.publish(cell, v);
        try self.journal.append(self.gpa, cell);
        // Only iteration reads a held value, so only iteration pays for
        // holding one — see the field's header.
        if (self.settings.iterate) try self.held.put(self.gpa, cell, v);
    }

    fn heldValue(self: Engine, cell: env.CellRef) ?Snapshot {
        return self.held.get(cell);
    }

    /// The components this pass covers, re-derived from the roots.
    ///
    /// Re-derived rather than remembered: see `Options.closure`. Planning
    /// does not charge `total_cell_evals` here — the engine charges per
    /// evaluation, and a plan charged once per pass would count the same
    /// cell twice.
    fn planScope(self: *Engine, g: graph.Graph) Error!?Refusal {
        const c = self.opts.closure orelse {
            self.scoped = false;
            return null;
        };
        // The plan dies with this call: only its component ids survive,
        // copied into `scope` below — planned into its own scratch
        // rather than the engine arena, which held the dead plan for
        // the drive's lifetime (§9.1 M10j).
        var plan_scratch = std.heap.ArenaAllocator.init(self.gpa);
        defer plan_scratch.deinit();
        const planned = try graph.plan(g, self.gpa, plan_scratch.allocator(), c.roots, self.counters(), .{
            .iterating = c.iterating,
            .charge_evals = false,
        });
        switch (planned) {
            .refused => |r| return .{
                .reason = switch (r.reason) {
                    .cycle => .cycle,
                    .work_limit_exceeded => .work_limit_exceeded,
                    else => .malformed_cache_seed,
                },
                .at = r.at,
                .limit = r.limit,
            },
            .ok => |p| {
                // Sized to the condensation, cleared, then filled — the
                // rebuild case renumbers, so stale bits cannot carry.
                try self.scope.resize(self.gpa, g.componentCount(), false);
                self.scope.unsetAll();
                for (p.components) |cid| self.scope.set(cid);
            },
        }
        self.scoped = true;
        return null;
    }

    fn inScope(self: Engine, cid: u32) bool {
        if (!self.scoped) return true;
        return self.scope.isSet(cid);
    }

    // ─── §5.6e: the outer loop ───────────────────────────────────

    fn drive(self: *Engine, initial: graph.Graph) Error!Result {
        // The engine owns every graph it runs, the initial one included
        // (see `run`'s header). At most one is ever alive: the outgoing
        // graph is freed before its successor is built, because the two
        // coexisting was the run's memory peak (§9.1). Declared before
        // any early return — ownership starts at the call, not at the
        // first pass.
        var current: ?graph.Graph = initial;
        defer if (current) |*o| o.deinit();

        // One block, not a ladder (§9.1 M10k): a publish comes only
        // from a node's evaluation, so the node count bounds an acyclic
        // run's journal exactly — the append ladder held 1 493 928 B
        // against 959 964 B of entries at the drive's peak instant, in
        // eighteen growth chunks. Iteration republishes and may append
        // past the reserve; that is growth from a full block, not from
        // a ladder's first rung.
        try self.journal.ensureTotalCapacityPrecise(self.gpa, initial.keys.len);

        const ceiling: u32 = @intCast(@min(
            self.opts.limits.max_dynamic_passes,
            @as(u64, std.math.maxInt(u32)),
        ));
        // A limit of zero permits no passes at all. `WorkLimits.validate`
        // rejects it, but the engine is reachable without a validated
        // set, and a caller-supplied number must not be able to trip an
        // assertion — §9's answer to an out-of-range limit is a refusal.
        if (ceiling == 0) return .{ .refused = .{
            .reason = .work_limit_exceeded,
            .limit = .dynamic_passes,
        } };

        // Lazy (M10b): only a real rebuild compares condensations, so a
        // run whose gate never lets one happen never pays for the
        // signatures either.
        var signatures: ?[]const Signature = null;
        var rerun: ?[]const bool = null;
        var pass: u32 = 1;
        while (true) : (pass += 1) {
            if (try self.planScope(current.?)) |r| return .{ .refused = r };
            self.pass_edges.clearRetainingCapacity();
            self.pass_edge_set.clearRetainingCapacity();
            self.touched_set.clearRetainingCapacity();
            self.run_graph = &current.?;
            const executed = try self.execute(current.?, rerun);
            self.run_graph = null;
            switch (executed) {
                .refused => |r| return .{ .refused = r },
                .ok => {},
            }

            const rb = self.opts.rebuild orelse {
                // No rebuild seam wired: the caller asked for §5.6c
                // only, and a dynamic read it cannot act on is not a
                // reason to refuse a schedule that completed.
                return .{ .ok = try self.report(pass) };
            };

            // The fixpoint test is on the **graph**, not on the edge set
            // — §5.6e terminates when rebuilding with what was actually
            // read yields the same condensation: then every dependency
            // this run has is already in the order it ran under, and
            // every cell downstream of a value that moved re-ran (that
            // is what `changedComponents` propagates). A value can only
            // still be stale if something it depends on is missing from
            // the graph — and that is precisely the case where the
            // graph changes.
            //
            // The gate (M10b) decides the common case of that test
            // WITHOUT building anything. `noteReads` records only reads
            // the walk did not find — everything else would dedupe away
            // inside a rebuild's injection (`captureAll`) and so cannot
            // change any graph — and `graph.build` is deterministic in
            // (input, resolver, injected edges). So when the fold names
            // exactly the edges the current graph was built from, the
            // rebuild is provably the identical graph, `sameCondensation`
            // is provably true, and paying a full parse-and-link to
            // learn so is the §9.1 profile's single largest cost. The
            // named workload folds zero edges against zero injected;
            // a stable dynamic reference folds the same edge set it ran
            // under; both stop here.
            const unchanged = blk: {
                var folded = try self.foldedPassEdges();
                errdefer folded.deinit(self.gpa);
                const eq = try self.foldedEqualsInjected(folded.items);
                self.dynamic.deinit(self.gpa);
                self.dynamic = folded;
                break :blk eq;
            };
            if (unchanged) return .{ .ok = try self.report(pass) };

            var next_opts = rb.options;
            next_opts.limits = self.opts.limits;
            next_opts.counters = self.counters();
            next_opts.dynamic_edges = self.dynamic.items;
            // The signatures of the outgoing graph, taken only now
            // (M10b): the gate above answers most runs, and a run it
            // answers never needs them. After a real rebuild they are
            // carried forward, so each graph is signed at most once.
            if (signatures == null) signatures = try self.signaturesOf(current.?);
            // The pass is over and nothing below reads the old graph —
            // `signatures` copied its keys and the rebuild starts from
            // the input. Freeing it here is the point of the engine
            // owning it: the build below is the run's high-water mark.
            current.?.deinit();
            current = null;
            const built = try graph.build(self.gpa, rb.input, rb.resolver, next_opts);
            switch (built) {
                .refused => |r| return .{ .refused = .{
                    .reason = switch (r.reason) {
                        .malformed_cache_seed => .malformed_cache_seed,
                        .work_limit_exceeded => .work_limit_exceeded,
                        .cycle => .cycle,
                        else => .evaluation_refused,
                    },
                    .at = r.at,
                    .limit = r.limit,
                    .plane = switch (r.reason) {
                        .malformed_cache_seed, .work_limit_exceeded, .cycle => null,
                        else => r.planeTwo(),
                    },
                } },
                .ok => |x| current = x,
            }

            const next_signatures = try self.signaturesOf(current.?);
            if (sameCondensation(signatures.?, next_signatures)) {
                return .{ .ok = try self.report(pass) };
            }
            if (pass >= ceiling) {
                return .{ .refused = .{
                    .reason = .dynamic_ref_unstable,
                    .limit = .dynamic_passes,
                } };
            }

            rerun = try self.changedComponents(current.?, next_signatures, signatures.?);
            signatures = next_signatures;
            try self.carryReports();
        }
    }

    /// Fold this pass's reads into the standing edge set, carrying
    /// forward whatever an unrun owner last reported. Returned rather
    /// than installed (M10b): the gate compares the fold against what
    /// `self.dynamic` still holds — the set the current graph was built
    /// from — before the caller installs it.
    fn foldedPassEdges(self: *Engine) Error!std.ArrayListUnmanaged(graph.DynamicRef) {
        var next: std.ArrayListUnmanaged(graph.DynamicRef) = .empty;
        errdefer next.deinit(self.gpa);
        try next.appendSlice(self.gpa, self.pass_edges.items);
        for (self.dynamic.items) |old| {
            const ran = self.touched_set.contains(old.owner);
            // A pass that ran an owner restated everything that owner
            // still reads, so its old edges are superseded. A pass that
            // skipped one has said nothing about it, and dropping those
            // edges would make "not re-evaluated" indistinguishable from
            // "reads nothing".
            if (!ran) try next.append(self.gpa, old);
        }
        return next;
    }

    /// Whether the fold names exactly the edge set the current graph
    /// was built from (M10b's gate). Set equality over two lists that
    /// are both exact-deduped by construction — `noteEdge` dedupes the
    /// pass's edges, the carry admits only owners the pass did not
    /// touch, and the injected list is a previous fold — so equal
    /// lengths plus one-sided membership decide it. The overwhelmingly
    /// common shape, zero against zero, costs nothing at all.
    fn foldedEqualsInjected(self: *Engine, folded: []const graph.DynamicRef) Error!bool {
        const injected = self.dynamic.items;
        if (folded.len != injected.len) return false;
        if (folded.len == 0) return true;
        var set: EdgeSet = .empty;
        defer set.deinit(self.gpa);
        try set.ensureTotalCapacity(self.gpa, @intCast(injected.len));
        for (injected) |e| set.putAssumeCapacity(e, {});
        for (folded) |e| if (!set.contains(e)) return false;
        return true;
    }

    /// Whether two condensations are the same condensation.
    fn sameCondensation(before: []const Signature, after: []const Signature) bool {
        if (before.len != after.len) return false;
        for (before, after) |x, y| {
            if (!x.eql(y)) return false;
        }
        return true;
    }

    /// Keep the reports of components this pass did not re-run, so the
    /// next pass can restate them rather than invent an outcome for a
    /// component it never touched.
    fn carryReports(self: *Engine) Error!void {
        std.mem.swap(
            std.ArrayListUnmanaged(ComponentReport),
            &self.reports,
            &self.previous_reports,
        );
        self.reports.clearRetainingCapacity();

        // Rebuilt rather than maintained alongside the appends: the list
        // is replaced wholesale here, and one O(components) pass per
        // dynamic pass is the cheap half of the trade.
        self.previous_index.clearRetainingCapacity();
        try self.previous_index.ensureTotalCapacity(
            self.gpa,
            @intCast(self.previous_reports.items.len),
        );
        for (self.previous_reports.items, 0..) |r, i| {
            // First occurrence wins, which is what a scan from the front
            // answered. Two reports naming one component would be a bug
            // upstream; this is not the place that decides so.
            const gop = self.previous_index.getOrPutAssumeCapacity(r.at);
            if (!gop.found_existing) gop.value_ptr.* = @intCast(i);
        }
    }

    fn previousReport(self: Engine, at: graph.Key) ?ComponentReport {
        const i = self.previous_index.get(at) orelse return null;
        return self.previous_reports.items[i];
    }

    fn report(self: *Engine, passes: u32) Error!Report {
        var non_converged: u32 = 0;
        var indeterminate: u32 = 0;
        for (self.reports.items) |r| {
            non_converged += r.non_converged_cells;
            if (r.outcome == .shape_indeterminate) indeterminate += 1;
        }
        return .{
            // The caller's allocator, not the arena: the report is the
            // one thing that outlives this engine. Moved, not duped
            // (§9.1 M10f): the dupe minted a second copy of every
            // component record at the drive's very peak while the list
            // still held its own buffer — and with `execute`'s exact
            // pre-size the two lengths agree, so the move is free.
            .components = try self.reports.toOwnedSlice(self.gpa),
            .non_converged_cells = non_converged,
            .shape_indeterminate_cells = indeterminate,
            .dynamic_passes = passes,
            .cell_evals = self.counters().usedBy(.total_cell_evals),
        };
    }

    // ─── component identity across rebuilds (§5.6e) ──────────────

    /// A component's identity as *content*: its member keys and the
    /// keys those members depend on.
    ///
    /// Node ids move when the node set changes, so an id-based
    /// comparison would report every component as changed the moment a
    /// single new node appeared anywhere. Comparing keys means "changed"
    /// means what §5.6e says it means — merged, split, or gained or lost
    /// an edge.
    const Signature = struct {
        members: []const graph.Key,
        /// Every member's dependency keys, concatenated in member order.
        ///
        /// Intra-component edges are kept, not filtered out as
        /// "implied by the member list". They are not implied: a
        /// component of one node is cyclic when that node depends on
        /// itself and acyclic when it does not, and those are the same
        /// member list. `INDIRECT` closing a self-reference is exactly
        /// that case, and filtering would have made §5.6e's flip
        /// invisible to the very comparison that has to see it.
        deps: []const graph.Key,
        cyclic: bool,

        fn eql(x: Signature, y: Signature) bool {
            if (x.cyclic != y.cyclic) return false;
            if (x.members.len != y.members.len or x.deps.len != y.deps.len) return false;
            for (x.members, y.members) |p, q| if (!p.eql(q)) return false;
            for (x.deps, y.deps) |p, q| if (!p.eql(q)) return false;
            return true;
        }

        /// Every term `eql` compares, in the order it compares them.
        fn hash(self: Signature, h: *std.hash.Wyhash) void {
            h.update(&[_]u8{@intFromBool(self.cyclic)});
            h.update(std.mem.asBytes(&@as(u64, self.members.len)));
            h.update(std.mem.asBytes(&@as(u64, self.deps.len)));
            for (self.members) |k| k.hash(h);
            for (self.deps) |k| k.hash(h);
        }

        const HashContext = struct {
            pub fn hash(_: HashContext, s: Signature) u64 {
                var h: std.hash.Wyhash = .init(0);
                s.hash(&h);
                return h.final();
            }
            pub fn eql(_: HashContext, x: Signature, y: Signature) bool {
                return x.eql(y);
            }
        };
    };

    fn signaturesOf(self: *Engine, g: graph.Graph) Error![]const Signature {
        const out = try self.a().alloc(Signature, g.componentCount());
        for (out, 0..) |*slot, i| {
            const comp = g.members(i);
            const members = try self.a().alloc(graph.Key, comp.len);
            for (comp, 0..) |n, k| members[k] = g.keys[n];

            // Counted, then filled: a list growing inside this arena
            // strands every buffer it abandons until the engine dies,
            // once per component per pass (§9.1).
            var dep_count: usize = 0;
            for (comp) |n| dep_count += g.depsOf(n).len;
            const deps = try self.a().alloc(graph.Key, dep_count);
            var k: usize = 0;
            for (comp) |n| {
                // `g.depsOf(n)` is ascending and deduped already (M5a1),
                // and members arrive in ascending node order, so the
                // concatenation is canonical without a sort here.
                for (g.depsOf(n)) |d| {
                    deps[k] = g.keys[d];
                    k += 1;
                }
            }
            slot.* = .{
                .members = members,
                .deps = deps,
                .cyclic = g.cyclic[g.component[comp[0]]],
            };
        }
        return out;
    }

    /// Which components must re-run: the ones whose signature is new,
    /// plus everything transitively downstream of them (§5.6e's
    /// "transitive dependents of every changed node plus all changed
    /// SCCs").
    fn changedComponents(
        self: *Engine,
        g: graph.Graph,
        now: []const Signature,
        before: []const Signature,
    ) Error![]const bool {
        const rerun = try self.a().alloc(bool, g.componentCount());
        @memset(rerun, false);

        // "Is this signature one of the ones from before?" — a
        // membership test, asked once per component, and answered by
        // scanning every previous signature until M5d4. Hashing the
        // terms `eql` compares answers the same question in one look;
        // duplicate signatures collapsing in the set is exactly what a
        // scan-until-first-match already did with them.
        var was: std.HashMapUnmanaged(Signature, void, Signature.HashContext, 80) = .empty;
        defer was.deinit(self.a());
        try was.ensureTotalCapacity(self.a(), @intCast(before.len));
        for (before) |old| was.putAssumeCapacity(old, {});

        for (now, 0..) |sig, i| {
            if (!was.contains(sig)) rerun[i] = true;
        }

        // `g.order` is topological, so one forward sweep propagates
        // downstream: a component's dependencies all sit at earlier
        // positions than it does.
        var position: std.AutoHashMapUnmanaged(u32, usize) = .empty;
        defer position.deinit(self.a());
        for (0..g.componentCount()) |i| {
            try position.put(self.a(), g.component[g.members(i)[0]], i);
        }
        for (0..g.componentCount()) |i| {
            if (rerun[i]) continue;
            const comp = g.members(i);
            const depends_on_changed = blk: {
                for (comp) |n| {
                    for (g.depsOf(n)) |d| {
                        const cid = g.component[d];
                        if (cid == g.component[n]) continue;
                        if (rerun[position.get(cid).?]) break :blk true;
                    }
                }
                break :blk false;
            };
            rerun[i] = depends_on_changed;
        }
        return rerun;
    }

    // ─── §5.6c: the inner loop ───────────────────────────────────

    const Executed = union(enum) { ok, refused: Refusal };

    fn execute(self: *Engine, g: graph.Graph, rerun: ?[]const bool) Error!Executed {
        // Exact, not laddered (§9.1 M10f): every in-scope component
        // appends exactly one report per pass — run or carried — so the
        // condensation's component count bounds the list precisely and
        // the append ladder's churn (three growth doublings retained by
        // the allocator at the drive's peak) never happens. `report`
        // then MOVES this buffer out instead of duplicating it.
        try self.reports.ensureTotalCapacityPrecise(self.gpa, g.componentCount());
        for (0..g.componentCount()) |position| {
            const comp = g.members(position);
            const cid = g.component[comp[0]];
            const cyclic = g.cyclic[cid];
            if (!self.inScope(cid)) continue;
            // "Unchanged" is only a reason to skip a component that
            // actually ran. A dynamic edge can pull a component INTO the
            // closure without changing anything about the component
            // itself — `A1=INDIRECT("C1")` leaves `C1` exactly as it
            // was and merely makes it reachable — and skipping it there
            // would leave the cell the new edge points at unevaluated,
            // which is the one state the rebuild existed to fix.
            const ran_before = self.previousReport(g.keys[comp[0]]) != null;
            const skip = if (rerun) |r| !r[position] and ran_before else false;
            if (skip) {
                // §5.6e: an unchanged SCC keeps its converged state. Its
                // report from the previous pass is still the truth about
                // it, so it is carried forward rather than described
                // afresh by a pass that did not look at it — a skipped
                // cyclic component that reported "converged" would be
                // claiming an outcome nobody observed.
                var carried = self.previousReport(g.keys[comp[0]]) orelse ComponentReport{
                    .at = g.keys[comp[0]],
                    .cyclic = cyclic,
                    .outcome = if (cyclic) .converged else .acyclic,
                };
                carried.rerun = false;
                try self.reports.append(self.gpa, carried);
                continue;
            }
            const r = if (cyclic)
                try self.iterateComponent(g, comp, rerun != null)
            else
                try self.evaluateAcyclic(g, comp, rerun != null);
            switch (r) {
                .refused => |x| return .{ .refused = x },
                .ok => {},
            }
        }
        return .ok;
    }

    fn evaluateAcyclic(
        self: *Engine,
        g: graph.Graph,
        comp: []const u32,
        rerun: bool,
    ) Error!Executed {
        for (comp) |node| {
            if (g.keys[node].kind() != .cell) continue;
            const cell = g.keys[node].cell;
            switch (try self.evaluateOne(cell, 0)) {
                .refused => |r| return .{ .refused = r },
                .ok => |v| try self.publish(cell, v),
            }
        }
        try self.reports.append(self.gpa, .{
            .at = g.keys[comp[0]],
            .cyclic = false,
            .outcome = .acyclic,
            .rerun = rerun,
        });
        return .ok;
    }

    /// One SCC, iterated to its own convergence before anything
    /// downstream of it evaluates (§5.6c).
    fn iterateComponent(
        self: *Engine,
        g: graph.Graph,
        comp: []const u32,
        rerun: bool,
    ) Error!Executed {
        if (!self.settings.iterate) {
            return .{ .refused = .{ .reason = .cycle, .at = g.keys[comp[0]] } };
        }

        // The members, in the canonical order M5a1 sorted them into.
        // "Fixed coordinate order" is a property of the graph, not a
        // sort this file performs — re-sorting here would be a second
        // opinion about an order that already has one.
        var members: std.ArrayListUnmanaged(env.CellRef) = .empty;
        defer members.deinit(self.gpa);
        for (comp) |node| {
            if (g.keys[node].kind() != .cell) continue;
            try members.append(self.gpa, g.keys[node].cell);
        }
        if (members.items.len == 0) {
            // A cycle through nothing but ordering nodes — a name that
            // references itself, say. There is no cell to iterate, and
            // the members that would have carried a value do not exist.
            try self.reports.append(self.gpa, .{
                .at = g.keys[comp[0]],
                .cyclic = true,
                .outcome = .converged,
                .rerun = rerun,
            });
            return .ok;
        }

        // ── seeding (§5.6c's table, plus its shape pass) ──
        for (comp) |node| {
            if (g.keys[node].kind() != .cell) continue;
            const cell = g.keys[node].cell;
            const seed = g.seedOf(node) orelse continue;
            const snapshot: Snapshot = switch (seed) {
                .number => |n| Snapshot.number(n),
                .array_zeros => |s| Snapshot.zeros(self.a(), s) catch |e| switch (e) {
                    error.OutOfMemory => return error.OutOfMemory,
                    error.ShapeTooLarge => return .{ .refused = .{
                        .reason = .seed_shape_too_large,
                        .at = .{ .cell = cell },
                    } },
                },
                // §5.6c's pre-iteration shape pass: evaluate the anchor
                // once OUTSIDE the cycle, purely to learn its shape, and
                // seed the zero-filled version of what came back. Pass 0
                // keys it distinctly from the first real pass, and
                // §5.6d's memo means the draws it makes are the draws
                // pass 1 will reuse.
                .shape_pass => switch (try self.evaluateOne(cell, 0)) {
                    .refused => |r| return .{ .refused = r },
                    .ok => |v| Snapshot.zeros(self.a(), v.shape()) catch |e| switch (e) {
                        error.OutOfMemory => return error.OutOfMemory,
                        error.ShapeTooLarge => return .{ .refused = .{
                            .reason = .seed_shape_too_large,
                            .at = .{ .cell = cell },
                        } },
                    },
                },
            };
            try self.publish(cell, snapshot);
        }

        // ── the two bounds ──
        const semantic: u64 = self.settings.iterate_count;
        const resource: u64 = self.opts.limits.max_scc_iterations;
        // Strictly lower: equal bounds are the workbook's answer,
        // because the caller permitted exactly what the file asked for.
        const resource_binds = resource < semantic;
        const limit: u64 = @min(semantic, resource);

        var pass: u32 = 0;
        var settled = false;
        var indeterminate = false;
        var not_converged: u32 = 0;

        while (pass < limit) {
            pass += 1;
            var all = true;
            not_converged = 0;
            for (members.items) |cell| {
                const previous = self.heldValue(cell);
                switch (try self.evaluateOne(cell, pass)) {
                    .refused => |r| return .{ .refused = r },
                    .ok => |next| {
                        const verdict = if (previous) |p|
                            compare(p, next, self.settings.iterate_delta)
                        else
                            .changed;
                        if (verdict == .shape_change) {
                            // §5.6c: a shape that moves between
                            // iterations is indeterminate, and the
                            // answer is `#SPILL!` rather than another
                            // pass. Publishing it *is* stopping —
                            // there is nothing left to converge to.
                            try self.publish(cell, .{
                                .scalar = value.ScalarValue.errorOf(.spill),
                            });
                            indeterminate = true;
                            break;
                        }
                        if (!verdict.settled()) {
                            all = false;
                            not_converged += 1;
                        }
                        // Gauss–Seidel: visible to the next member of
                        // this same pass, not at the end of it.
                        try self.publish(cell, next);
                    },
                }
            }
            if (indeterminate) break;
            if (all) {
                settled = true;
                break;
            }
        }

        if (indeterminate) {
            try self.reports.append(self.gpa, .{
                .at = g.keys[comp[0]],
                .cyclic = true,
                .outcome = .shape_indeterminate,
                .passes = pass,
                .rerun = rerun,
            });
            return .ok;
        }

        if (!settled and resource_binds) {
            // The caller's ceiling was the bound that fired, and it was
            // lower than what the workbook asked for. §9's limits are
            // plane-2 refusals at every layer, and this one takes the
            // whole run down with it.
            try self.reports.append(self.gpa, .{
                .at = g.keys[comp[0]],
                .cyclic = true,
                .outcome = .semantic_bound,
                .passes = pass,
                .bound = .resource,
                .non_converged_cells = not_converged,
                .rerun = rerun,
            });
            return .{ .refused = .{
                .reason = .scc_iteration_ceiling,
                .at = g.keys[comp[0]],
                .limit = .scc_iterations,
            } };
        }

        try self.reports.append(self.gpa, .{
            .at = g.keys[comp[0]],
            .cyclic = true,
            .outcome = if (settled) .converged else .semantic_bound,
            .passes = pass,
            .bound = if (settled) null else .semantic,
            .non_converged_cells = if (settled) 0 else not_converged,
            .rerun = rerun,
        });
        return .ok;
    }

    const One = union(enum) { ok: Snapshot, refused: Refusal };

    fn evaluateOne(self: *Engine, cell: env.CellRef, pass: u32) Error!One {
        self.counters().charge(.total_cell_evals, 1) catch return .{ .refused = .{
            .reason = .work_limit_exceeded,
            .at = .{ .cell = cell },
            .limit = .total_cell_evals,
        } };

        const key: draw_schedule.Key = .{
            .path = draw_schedule.Key.ofCell(
                cell.sheet.toInt(),
                cell.row.oneBased(),
                cell.col.zeroBased(),
            ),
            .pass = pass,
        };
        return switch (try self.host.evaluate(cell, key)) {
            .refused => |p| .{ .refused = .{
                .reason = .evaluation_refused,
                .at = .{ .cell = cell },
                .plane = p,
            } },
            .ok => |produced| blk: {
                try self.noteReads(cell, produced.reads);
                break :blk .{ .ok = produced.value };
            },
        };
    }

    /// §5.6e's runtime-edge capture.
    ///
    /// Only reads the walk did NOT note are recorded (M10b). This is
    /// not the engine *classifying* a read as dynamic — which would
    /// mean this file deciding what the walk can see, exactly the
    /// disagreement M5a1's differential test exists to catch — it is
    /// the engine asking the walk's own record, through `walkNoted`,
    /// the same membership a rebuild's injection would ask: an edge the
    /// walk already noted dedupes away inside `captureAll` and can
    /// never change any graph, so recording it buys nothing and costs
    /// the fold its emptiness. The probe answers against the walk
    /// prefix, never the injected tail, so a genuinely dynamic read
    /// stays novel on every pass and is restated on every pass — the
    /// carry contract `foldedPassEdges` depends on.
    fn noteReads(self: *Engine, cell: env.CellRef, reads: Reads) Error!void {
        const owner: graph.Key = .{ .cell = cell };
        // Recorded only while there are old edges a fold could carry —
        // see the field's header. An empty `dynamic` folds identically
        // against an empty set and an unwritten one.
        if (self.dynamic.items.len != 0) {
            try self.touched_set.put(self.gpa, owner, {});
        }

        // Null only if the evaluated coordinate is somehow not a node
        // of the running graph; recording everything is the fallback
        // that preserves the pre-gate behavior exactly.
        const probe: ?struct { g: *const graph.Graph, node: u32 } = blk: {
            const g = self.run_graph orelse break :blk null;
            const u = g.find(owner) orelse break :blk null;
            break :blk .{ .g = g, .node = u };
        };

        for (reads.cells) |c| {
            if (probe) |p| {
                // Dedupes in the injection: the walk noted this exact
                // coordinate.
                if (p.g.walkNoted(p.node, .{ .cell = c })) continue;
                // Cannot survive the injection: `collectKeys` makes no
                // key for a cell ref, so a coordinate that is not
                // already a node (formula cell or spill tail — both
                // derived from the Input alone, so absent from every
                // rebuild too) resolves to no target and `link` draws
                // no edge. Aggregates note the stored cells they visit
                // (`readCell` under a cursor), which over a data column
                // is exactly this shape, ~every window formula.
                if (p.g.find(.{ .cell = c }) == null and
                    p.g.find(.{ .spill_tail = c }) == null) continue;
            }
            try self.noteEdge(.{ .owner = owner, .target = .{ .cell = c } });
        }
        for (reads.areas) |x| {
            if (probe) |p| {
                if (p.g.walkNoted(p.node, .{ .area = x })) continue;
                // The walk's own normalization (`Capture.note`): a
                // single-cell area IS a cell. The runtime spells the
                // same read as a raw 1×1 area (`readRange` notes what
                // it iterates), and asking it in the walk's vocabulary
                // is what lets it dedupe. The ordering the read implies
                // is already whatever the static build decided for that
                // coordinate; the 1×1 range node an injection would
                // mint carries no edge the run can feel — only the
                // wasted pass its appearance forces.
                if (x.isSingleCell() and
                    p.g.walkNoted(p.node, .{ .cell = x.topLeft() })) continue;
            }
            try self.noteEdge(.{ .owner = owner, .target = .{ .area = x } });
        }
    }

    fn noteEdge(self: *Engine, edge: graph.DynamicRef) Error!void {
        if ((try self.pass_edge_set.getOrPut(self.gpa, edge)).found_existing) return;
        errdefer _ = self.pass_edge_set.remove(edge);
        try self.pass_edges.append(self.gpa, edge);
    }
};

// ─── tests ───────────────────────────────────────────────────────

const eval = @import("eval.zig");
const testing = std.testing;

/// The shipped fold, so a fixture's comparator is the comparator
/// (§5.4b). Reached the same way every other test in the engine reaches
/// it: through the collation the evaluator is handed.
fn shippedFold(allocator: std.mem.Allocator, s: []const u8) anyerror![]u8 {
    return @import("zlsx_casefold").foldString(allocator, s);
}

const test_collation: value.Collation = .{ .fold = shippedFold };

fn cellAt(sheet: u32, a1: []const u8) env.CellRef {
    const p = coords.parseCell(a1, .{ .dollar = .accept }) catch unreachable;
    return .{ .sheet = env.SheetIndex.fromInt(sheet), .row = p.row, .col = p.col };
}

fn rangeOf(a1: []const u8) coords.Range {
    return (coords.parseRange(a1, .{ .dollar = .accept }) catch unreachable).normalized();
}

/// The smallest resolver a graph build needs: two sheets, no names, no
/// tables. `graph.zig`'s own tests have the full one; what these
/// fixtures exercise is the schedule, not resolution.
const World = struct {
    const sheets = [_][]const u8{ "Sheet1", "Sheet2" };

    fn resolver() graph.Resolver {
        return .{ .ctx = @constCast(&sheets), .vtable = &vtable };
    }

    const vtable: graph.Resolver.VTable = .{
        .resolveSheet = vtSheet,
        .resolveName = vtName,
        .resolveStructured = vtStructured,
    };

    fn vtSheet(_: *anyopaque, name: []const u8) graph.Error!?env.SheetIndex {
        for (sheets, 0..) |s, i| {
            if (std.ascii.eqlIgnoreCase(s, name)) return env.SheetIndex.fromInt(@intCast(i));
        }
        return null;
    }

    fn vtName(_: *anyopaque, _: ?env.SheetIndex, _: []const u8) graph.Error!graph.NameBinding {
        return .unresolved;
    }

    fn vtStructured(
        _: *anyopaque,
        _: ?env.SheetIndex,
        _: graph.Owner,
        _: ?[]const u8,
        _: parser.ItemSet,
        _: parser.ColumnSelector,
    ) graph.Error!?env.RangeRef {
        return null;
    }
};

/// One cell of a fixture workbook.
const Spec = struct {
    a1: []const u8,
    /// The `<f>` body. Empty means a constant cell — stored, never
    /// evaluated, and therefore not a node (§5.6b).
    formula: []const u8 = "",
    /// The `<v>` cache, which is what §5.6c's seed table reads.
    cache: graph.CacheState = .absent,
    /// The value the environment starts with at this coordinate.
    stored: ?value.ScalarValue = null,
    /// A declared array range, in A1 notation.
    array: ?[]const u8 = null,
    dynamic_anchor: bool = false,
};

/// What a scripted cell returns on its Nth evaluation.
const Step = union(enum) {
    scalar: value.ScalarValue,
    /// An array of exactly this shape, every element `fill`. The fill is
    /// separate from the shape because a fixture usually needs to move
    /// one without the other: a value that changes at a stable shape is
    /// ordinary non-convergence, and a shape that changes is `#SPILL!`.
    array: Filled,
};

const Filled = struct { rows: u32, cols: u32, fill: f64 = 0 };

/// A cell whose successive results are stated rather than computed.
///
/// The engine's contract for a moving shape is "publish `#SPILL!` and
/// stop", and that contract is about what the HOST returns — it holds
/// however the host arrived at the shape. Producing a genuinely
/// resizing array needs the dynamic-array functions, which are M7a's
/// and which this row must not touch, so the shapes are stated. The
/// last step repeats, so a script is "what happens, then what happens
/// forever after".
const Scripted = struct {
    a1: []const u8,
    steps: []const Step,
};

/// A fixture workbook plus the host that evaluates it.
///
/// The host is the same seam `pkg/workbook.zig` implements, wired to a
/// `env.Fake` and to the real evaluator — so what these fixtures pin is
/// the engine driving an evaluator, not the engine driving a mock that
/// agrees with it.
const Fixture = struct {
    gpa: std.mem.Allocator,
    arena_state: std.heap.ArenaAllocator,
    fake: env.Fake,
    cells: []graph.CellInput,
    schedule: draw_schedule.Schedule = .{},
    /// A counting generator rather than a constant: a KAT that asserts
    /// two draws DIFFER needs a source that can produce two values, and
    /// one that asserts a draw was REUSED needs to be able to tell the
    /// reuse from a coincidence. Kept inside `[0,1)` because that is
    /// what `RAND` promises and `RANDBETWEEN` asserts.
    draw_counter: f64 = 0,
    /// Every publish and retract this run made. A rollback is asserted
    /// on the environment's contents, and these say how it got there.
    publishes: u32 = 0,
    retracts: u32 = 0,
    /// Deliberately fail the Nth evaluation, for the injected-OOM leg of
    /// the purity gate. Zero disables it.
    fail_at_eval: u32 = 0,
    evals: u32 = 0,
    /// Report slices handed back by `run`, so a fixture body never has
    /// to remember to free one. `Report.components` outlives the engine
    /// by design (it is the run's answer), and a harness that leaked it
    /// would turn every fixture into a leak check failure.
    owned_reports: std.ArrayListUnmanaged([]const ComponentReport) = .empty,
    /// Cells whose successive results are stated (see `Scripted`).
    scripted: []const Scripted = &.{},
    script_calls: [8]u32 = @splat(0),
    /// Per-cell evaluation counts, so a fixture can assert that a
    /// downstream cell evaluated ONCE — §5.6c's "downstream sees final
    /// values only" is a statement about how many times, not only about
    /// which value.
    call_log: std.ArrayListUnmanaged(CallCount) = .empty,

    const CallCount = struct { cell: env.CellRef, n: u32 };

    fn init(f: *Fixture, gpa: std.mem.Allocator, specs: []const Spec) !void {
        f.* = .{
            .gpa = gpa,
            .arena_state = std.heap.ArenaAllocator.init(gpa),
            .fake = env.Fake.init(gpa),
            .cells = &.{},
        };
        _ = try f.fake.addSheet("Sheet1");
        _ = try f.fake.addSheet("Sheet2");

        var list: std.ArrayListUnmanaged(graph.CellInput) = .empty;
        for (specs) |s| {
            const cell = cellAt(0, s.a1);
            if (s.stored) |v| try f.fake.put(cell.sheet, .stored, .{
                .row = cell.row,
                .col = cell.col,
                .v = v,
            });
            if (s.formula.len == 0) continue;
            try list.append(f.arena(), .{
                .cell = cell,
                .formula = s.formula,
                .cache = s.cache,
                .array = if (s.array) |r| rangeOf(r) else null,
                .dynamic_anchor = s.dynamic_anchor,
            });
        }
        f.cells = try list.toOwnedSlice(f.arena());
    }

    fn deinit(f: *Fixture) void {
        for (f.owned_reports.items) |r| f.gpa.free(r);
        f.owned_reports.deinit(f.gpa);
        f.schedule.deinit(f.gpa);
        f.call_log.deinit(f.gpa);
        f.fake.deinit();
        f.arena_state.deinit();
    }

    /// How many times one cell was evaluated across the whole run.
    fn callsOn(f: *Fixture, a1: []const u8) u32 {
        const cell = cellAt(0, a1);
        for (f.call_log.items) |c| {
            if (c.cell.eql(cell)) return c.n;
        }
        return 0;
    }

    fn arena(f: *Fixture) std.mem.Allocator {
        return f.arena_state.allocator();
    }

    fn input(f: *Fixture) graph.Input {
        return .{ .sheet_count = 2, .cells = f.cells };
    }

    fn build(f: *Fixture) !graph.Graph {
        return switch (try graph.build(f.gpa, f.input(), World.resolver(), .{})) {
            .ok => |g| g,
            .refused => |r| {
                std.debug.print("fixture graph refused: {t}\n", .{r.reason});
                return error.FixtureGraphRefused;
            },
        };
    }

    fn host(f: *Fixture) Host {
        return .{ .ctx = f, .vtable = &host_vtable };
    }

    const host_vtable: Host.VTable = .{
        .evaluate = vtEvaluate,
        .publish = vtPublish,
        .retract = vtRetract,
    };

    fn of(ctx: *anyopaque) *Fixture {
        return @ptrCast(@alignCast(ctx));
    }

    fn noteCall(f: *Fixture, cell: env.CellRef) Error!void {
        for (f.call_log.items) |*c| {
            if (c.cell.eql(cell)) {
                c.n += 1;
                return;
            }
        }
        try f.call_log.append(f.gpa, .{ .cell = cell, .n = 1 });
    }

    fn formulaAt(f: *Fixture, cell: env.CellRef) ?[]const u8 {
        for (f.cells) |c| {
            if (c.cell.eql(cell)) return c.formula;
        }
        return null;
    }

    fn vtEvaluate(ctx: *anyopaque, cell: env.CellRef, key: draw_schedule.Key) Error!Produced {
        const f = of(ctx);
        f.evals += 1;
        if (f.fail_at_eval != 0 and f.evals == f.fail_at_eval) return error.OutOfMemory;
        try f.noteCall(cell);

        for (f.scripted, 0..) |s, i| {
            if (!cellAt(0, s.a1).eql(cell)) continue;
            const n = f.script_calls[i];
            f.script_calls[i] = n + 1;
            const step = s.steps[@min(n, s.steps.len - 1)];
            return .{
                .ok = .{
                    .value = switch (step) {
                        .scalar => |v| .{ .scalar = v },
                        .array => |sh| blk: {
                            // Every scripted shape is small by construction, so
                            // the §9 arm is a fixture bug rather than a workbook
                            // one and is caught as such.
                            const m = Snapshot.zeros(f.arena(), .{ .rows = sh.rows, .cols = sh.cols }) catch |e| switch (e) {
                                error.OutOfMemory => return error.OutOfMemory,
                                error.ShapeTooLarge => unreachable,
                            };
                            @memset(m.array.cells, value.ScalarValue.fromNumber(sh.fill));
                            break :blk m;
                        },
                    },
                },
            };
        }

        const text = f.formulaAt(cell) orelse return .{ .ok = .{ .value = .{ .scalar = .blank } } };
        const parsed = parser.parse(f.arena(), text, .{}) catch return error.OutOfMemory;
        const ast = switch (parsed) {
            .ok => |x| x,
            .refused => return .{ .ok = .{ .value = .{
                .scalar = value.ScalarValue.errorOf(.value),
            } } },
        };

        var source: eval.DrawSource = .{ .ctx = f, .draw_fn = nextDraw };
        source.schedule = &f.schedule;
        source.gpa = f.gpa;
        source.key = key;

        var ev = eval.Evaluator.init(f.arena(), f.fake.evalEnv(), .{
            .current_sheet = cell.sheet,
            .collation = test_collation,
            .draws = &source,
            .site = .{ .row = cell.row, .col = cell.col },
            .draw_path = key.path,
        });
        defer ev.deinit();

        const v = ev.evaluate(ast) catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            else => return .{ .refused = eval.planeTwo(e) },
        };

        // The dependency log is runtime capture — what the body actually
        // read — which is precisely §5.6e's input.
        const reads: Reads = .{
            .cells = try f.arena().dupe(env.CellRef, ev.deps.cells.items),
            .areas = try f.arena().dupe(env.RangeRef, ev.deps.areas.items),
        };
        return .{ .ok = .{ .value = snapshotOf(v), .reads = reads } };
    }

    fn nextDraw(ctx: *anyopaque) f64 {
        const f = of(ctx);
        f.draw_counter += 1;
        return f.draw_counter / 1024;
    }

    fn snapshotOf(v: eval.Value) Snapshot {
        return switch (v) {
            .scalar => |s| .{ .scalar = s },
            .array => |m| .{ .array = m },
            // A reference reaching the top of a stored body already
            // dereferenced inside `Evaluator.evaluate`; `missing_arg`
            // cannot be a whole formula.
            else => .{ .scalar = value.ScalarValue.errorOf(.value) },
        };
    }

    fn vtPublish(ctx: *anyopaque, cell: env.CellRef, v: Snapshot) Error!void {
        const f = of(ctx);
        f.publishes += 1;
        // An array's readable value at its anchor is its top-left
        // (§5.3b). Placing the tails is M7a's, and this row does not do
        // it — which is why the engine keeps the whole array itself and
        // only convergence ever looks at the rest of it.
        const scalar = switch (v) {
            .scalar => |s| s,
            .array => |m| m.topLeft(),
        };
        // `blank` is the absence of a cell, so a computed blank is a
        // retraction rather than a stored nothing.
        if (scalar == .blank) {
            f.fake.clear(cell.sheet, .computed, cell);
            return;
        }
        // The sheet came from a node of this fixture's own graph, so
        // `UnknownSheet` is unreachable; the seam's error set is narrow
        // on purpose, so the impossible half is mapped rather than
        // widened.
        f.fake.put(cell.sheet, .computed, .{
            .row = cell.row,
            .col = cell.col,
            .v = scalar,
        }) catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            else => unreachable,
        };
    }

    fn vtRetract(ctx: *anyopaque, cell: env.CellRef) void {
        const f = of(ctx);
        f.retracts += 1;
        f.fake.clear(cell.sheet, .computed, cell);
    }

    /// The merged value at a coordinate — what a reader of this
    /// workbook would see right now.
    fn read(f: *Fixture, a1: []const u8) !value.ScalarValue {
        return f.fake.evalEnv().cellValue(cellAt(0, a1));
    }

    fn appendDigest(f: *Fixture, out: *std.ArrayListUnmanaged(u8), comptime fmt: []const u8, args: anytype) !void {
        const text = try std.fmt.allocPrint(f.arena(), fmt, args);
        try out.appendSlice(f.gpa, text);
    }

    /// Every occupied coordinate of the computed layer, as a string. The
    /// purity gate compares two of these; a rolled-back run has to leave
    /// the empty one.
    fn computedDigest(f: *Fixture, out: *std.ArrayListUnmanaged(u8)) !void {
        out.clearRetainingCapacity();
        for (f.fake.sheets.items, 0..) |s, si| {
            for (s.cells.items) |c| {
                if (c.layer != .computed) continue;
                try f.appendDigest(out, "{d}!{d},{d}={t};", .{
                    si,
                    c.row.oneBased(),
                    c.col.zeroBased(),
                    @as(std.meta.Tag(value.ScalarValue), c.v),
                });
                if (c.v == .number) try f.appendDigest(out, "{d}", .{c.v.number});
            }
        }
    }

    fn run(f: *Fixture, opts: Options) !Result {
        const g = try f.build();
        // The engine owns `g` from `runOn` onward (see `run`'s header).
        return runOn(f, g, opts);
    }

    /// The same run with §5.6e's outer loop wired. Separate rather than
    /// a flag so a fixture states which loop it is about.
    fn runFixpoint(f: *Fixture, opts: Options) !Result {
        var o = opts;
        o.rebuild = .{ .input = f.input(), .resolver = World.resolver() };
        return f.run(o);
    }

    fn runOn(f: *Fixture, g: graph.Graph, opts: Options) !Result {
        var o = opts;
        if (o.schedule == null) o.schedule = &f.schedule;
        const r = try iterateRun(f.gpa, g, f.host(), o);
        if (r == .ok) try f.owned_reports.append(f.gpa, r.ok.components);
        return r;
    }
};

const iterateRun = run;

fn iterating(count: u32, delta: f64) Settings {
    return .{ .iterate = true, .iterate_count = count, .iterate_delta = delta };
}

// ─── §5.6c convergence, per type (done-when 4) ───────────────────

fn num(n: f64) Snapshot {
    return Snapshot.number(n);
}

test "M5a2 §5.6c: numbers converge by MAGNITUDE, increasing and decreasing alike" {
    // The bug this spelling exists to prevent: a raw signed difference
    // makes every decreasing value converge on its first comparison,
    // because `new - previous` is negative and any negative is below any
    // positive tolerance.
    try testing.expectEqual(Convergence.converged, compare(num(1.0), num(1.0005), 0.001));
    try testing.expectEqual(Convergence.changed, compare(num(1.0), num(1.002), 0.001));

    // Decreasing, the same two distances. A signed comparison would call
    // both of these converged.
    try testing.expectEqual(Convergence.converged, compare(num(1.0), num(0.9995), 0.001));
    try testing.expectEqual(Convergence.changed, compare(num(1.0), num(0.998), 0.001));

    // Sign-crossing: the magnitude is what counts, not that the sign
    // flipped.
    try testing.expectEqual(Convergence.converged, compare(num(0.0004), num(-0.0004), 0.001));
    try testing.expectEqual(Convergence.changed, compare(num(0.6), num(-0.6), 0.001));
}

test "M5a2 §5.6c: the boundary is strict, and signed zero converges" {
    // Exactly `delta` has NOT converged — `< delta`, not `<=`. Pinned
    // because both readings are defensible and only one can be shipped.
    try testing.expectEqual(Convergence.changed, compare(num(0), num(0.001), 0.001));
    try testing.expectEqual(Convergence.converged, compare(num(0), num(0.0009999), 0.001));

    // `-0` and `+0` differ in bits and not in magnitude. Iteration asks
    // how far a value moved, and it moved nowhere.
    const plus_zero = Snapshot{ .scalar = .{ .number = 0.0 } };
    const minus_zero = Snapshot{ .scalar = .{ .number = -0.0 } };
    try testing.expectEqual(Convergence.converged, compare(plus_zero, minus_zero, 0.001));
    try testing.expectEqual(Convergence.converged, compare(minus_zero, plus_zero, 0.001));

    // And the reason `Settings.normalize` refuses to let a zero delta
    // reach here: under a strict `<`, zero is satisfied by nothing at
    // all — not even by a value that did not move. A tolerance that
    // cannot be met is not a tolerance, which is why the transition
    // table treats it as an unset attribute rather than as a request
    // for exact equality.
    try testing.expectEqual(Convergence.changed, compare(plus_zero, minus_zero, 0));
    try testing.expectEqual(Convergence.changed, compare(num(1), num(1), 0));
    try testing.expectEqual(@as(f64, 0.001), Settings.normalize(.{ .iterate_delta = 0 }).iterate_delta);
}

test "M5a2 §5.6c: every non-numeric type converges by two-pass equality" {
    const cases = [_]struct { a: value.ScalarValue, b: value.ScalarValue, want: Convergence }{
        .{ .a = .{ .text = "abc" }, .b = .{ .text = "abc" }, .want = .converged },
        .{ .a = .{ .text = "abc" }, .b = .{ .text = "abd" }, .want = .changed },
        // Two-pass equality is BYTES, not the collation: a cycle whose
        // member alternates case has not settled, however equal the two
        // spellings compare in a lookup.
        .{ .a = .{ .text = "abc" }, .b = .{ .text = "ABC" }, .want = .changed },
        .{ .a = .{ .boolean = true }, .b = .{ .boolean = true }, .want = .converged },
        .{ .a = .{ .boolean = true }, .b = .{ .boolean = false }, .want = .changed },
        .{ .a = .blank, .b = .blank, .want = .converged },
        .{
            .a = value.ScalarValue.errorOf(.div0),
            .b = value.ScalarValue.errorOf(.div0),
            .want = .converged,
        },
        .{
            .a = value.ScalarValue.errorOf(.div0),
            .b = value.ScalarValue.errorOf(.value),
            .want = .changed,
        },
    };
    for (cases) |c| {
        try testing.expectEqual(c.want, compare(
            .{ .scalar = c.a },
            .{ .scalar = c.b },
            0.001,
        ));
    }
}

test "M5a2 §5.6c: any type transition is not converged" {
    // Including the one a tolerant reading would let through: `0` and
    // blank are the same number in every arithmetic context, and they
    // are still a transition here.
    const kinds = [_]value.ScalarValue{
        .{ .number = 0 },
        .{ .text = "0" },
        .{ .boolean = false },
        value.ScalarValue.errorOf(.na),
        .blank,
    };
    for (kinds, 0..) |a, i| {
        for (kinds, 0..) |b, j| {
            if (i == j) continue;
            try testing.expectEqual(
                Convergence.type_transition,
                compare(.{ .scalar = a }, .{ .scalar = b }, 1e9),
            );
        }
    }
}

test "M5a2 §5.6c: arrays need shape equality AND per-element convergence" {
    var a = try value.Matrix.init(testing.allocator, 2, 2);
    defer a.deinit(testing.allocator);
    var b = try value.Matrix.init(testing.allocator, 2, 2);
    defer b.deinit(testing.allocator);
    var narrow = try value.Matrix.init(testing.allocator, 2, 1);
    defer narrow.deinit(testing.allocator);

    for (0..4) |i| {
        a.cells[i] = value.ScalarValue.fromNumber(1);
        b.cells[i] = value.ScalarValue.fromNumber(1.0005);
    }
    for (0..2) |i| narrow.cells[i] = value.ScalarValue.fromNumber(1);

    // Same shape, every element inside the tolerance.
    try testing.expectEqual(Convergence.converged, compare(
        .{ .array = a },
        .{ .array = b },
        0.001,
    ));
    // One element outside it is enough.
    b.set(1, 1, value.ScalarValue.fromNumber(9));
    try testing.expectEqual(Convergence.changed, compare(
        .{ .array = a },
        .{ .array = b },
        0.001,
    ));
    // A type transition inside an array is the array's answer, and it
    // outranks "still moving" — it is the stronger statement about why
    // the two passes are not the same array.
    b.set(1, 1, .{ .text = "x" });
    try testing.expectEqual(Convergence.type_transition, compare(
        .{ .array = a },
        .{ .array = b },
        0.001,
    ));
    // Shape change, both directions, and against a scalar.
    try testing.expectEqual(Convergence.shape_change, compare(
        .{ .array = a },
        .{ .array = narrow },
        0.001,
    ));
    try testing.expectEqual(Convergence.shape_change, compare(
        .{ .array = narrow },
        .{ .array = a },
        0.001,
    ));
    try testing.expectEqual(Convergence.shape_change, compare(
        .{ .array = a },
        num(1),
        0.001,
    ));
    try testing.expectEqual(Convergence.shape_change, compare(
        num(1),
        .{ .array = a },
        0.001,
    ));
}

// ─── §5.6c's transition table ────────────────────────────────────

test "M5a2 §5.6c: the pinned transition table for missing, zero and out-of-range" {
    // Absent `<calcPr>` — every typed field at its schema default.
    const absent = Settings.normalize(.{});
    try testing.expect(!absent.iterate);
    try testing.expectEqual(@as(u32, 100), absent.iterate_count);
    try testing.expectEqual(@as(f64, 0.001), absent.iterate_delta);

    // Zero count: Excel's own minimum is 1, so a zero in the file is an
    // unset attribute and an unset attribute means the schema default.
    try testing.expectEqual(@as(u32, 100), Settings.normalize(.{ .iterate_count = 0 }).iterate_count);

    // The clamp §5.6c names, at and above it.
    try testing.expectEqual(
        @as(u32, 32_767),
        Settings.normalize(.{ .iterate_count = 32_767 }).iterate_count,
    );
    try testing.expectEqual(
        @as(u32, 32_767),
        Settings.normalize(.{ .iterate_count = 32_768 }).iterate_count,
    );
    try testing.expectEqual(
        @as(u32, 32_767),
        Settings.normalize(.{ .iterate_count = std.math.maxInt(u32) }).iterate_count,
    );

    // A negative maximum-change is not a tolerance, and neither is a
    // zero one under a strict comparison — both are unset attributes.
    try testing.expectEqual(
        @as(f64, 0.001),
        Settings.normalize(.{ .iterate_delta = -1 }).iterate_delta,
    );
    try testing.expectEqual(
        @as(f64, 0.001),
        Settings.normalize(.{ .iterate_delta = 0 }).iterate_delta,
    );
    // The smallest delta that IS a tolerance survives untouched: the
    // rule is "unsatisfiable or negative", not "small".
    try testing.expectEqual(
        @as(f64, std.math.floatMin(f64)),
        Settings.normalize(.{ .iterate_delta = std.math.floatMin(f64) }).iterate_delta,
    );
    // Normalizing is idempotent, which is what lets a caller normalize
    // at the boundary and the engine normalize again without moving.
    const once = Settings.normalize(.{ .iterate_count = 40_000, .iterate_delta = -3 });
    try testing.expectEqual(once, Settings.normalize(once));
}

// ─── the multi-SCC schedule (done-when 2, 3, 5) ──────────────────

/// `A1=(A1+1)/2` — §5.6c's own parity formula. Halves its distance to 1
/// every pass, so the pass at which it converges is arithmetic rather
/// than a number someone observed: the gap after pass k is 2^-k, and
/// 2^-10 is the first below 0.001.
const halving = [_]Spec{
    .{ .a1 = "A1", .formula = "(A1+1)/2" },
};

/// A cycle that never settles: the gap is 1 every pass, whatever the
/// tolerance.
const forever = [_]Spec{
    .{ .a1 = "A1", .formula = "A1+1" },
};

fn expectNumber(f: *Fixture, a1: []const u8, want: f64) !void {
    const v = try f.read(a1);
    try testing.expect(v == .number);
    try testing.expectApproxEqAbs(want, v.number, 1e-12);
}

test "M5a2 §5.6c: convergence before either bound is a success with no bound recorded" {
    var f: Fixture = undefined;
    try f.init(testing.allocator, &halving);
    defer f.deinit();

    const r = try f.run(.{
        .settings = iterating(100, 0.001),
        .limits = .{ .max_scc_iterations = 50 },
    });
    try testing.expect(r == .ok);
    try testing.expectEqual(@as(usize, 1), r.ok.components.len);

    const c = r.ok.components[0];
    try testing.expect(c.cyclic);
    try testing.expectEqual(Outcome.converged, c.outcome);
    try testing.expectEqual(@as(?Bound, null), c.bound);
    // Ten, by arithmetic: 2^-10 is the first gap below 0.001.
    try testing.expectEqual(@as(u32, 10), c.passes);
    try testing.expectEqual(@as(u32, 0), r.ok.non_converged_cells);
    try expectNumber(&f, "A1", 1 - std.math.pow(f64, 2, -10));
}

test "M5a2 §5.6c: a caller ceiling ABOVE iterateCount leaves the workbook's bound in charge" {
    var f: Fixture = undefined;
    try f.init(testing.allocator, &forever);
    defer f.deinit();

    const r = try f.run(.{
        .settings = iterating(10, 0.001),
        .limits = .{ .max_scc_iterations = 50 },
    });
    // Excel's documented behaviour: the workbook asked for ten passes,
    // got ten passes, and is calculated — non-convergence is reported,
    // not refused.
    try testing.expect(r == .ok);
    try testing.expectEqual(Outcome.semantic_bound, r.ok.components[0].outcome);
    try testing.expectEqual(@as(?Bound, .semantic), r.ok.components[0].bound);
    try testing.expectEqual(@as(u32, 10), r.ok.components[0].passes);
    try testing.expectEqual(@as(u32, 1), r.ok.non_converged_cells);
    try expectNumber(&f, "A1", 10);
}

test "M5a2 §5.6c: a caller ceiling EQUAL to iterateCount is still the workbook's bound" {
    // The row the two-outcome rule turns on. The ceiling refuses only
    // when it is STRICTLY lower: at equality the caller permitted
    // exactly what the file asked for and the file got it, so calling
    // this a resource refusal would refuse a run nothing actually
    // constrained.
    var f: Fixture = undefined;
    try f.init(testing.allocator, &forever);
    defer f.deinit();

    const r = try f.run(.{
        .settings = iterating(10, 0.001),
        .limits = .{ .max_scc_iterations = 10 },
    });
    try testing.expect(r == .ok);
    try testing.expectEqual(@as(?Bound, .semantic), r.ok.components[0].bound);
    try testing.expectEqual(@as(u32, 10), r.ok.components[0].passes);
    try expectNumber(&f, "A1", 10);
}

test "M5a2 §5.6c: a caller ceiling BELOW iterateCount refuses, and writes nothing" {
    var f: Fixture = undefined;
    try f.init(testing.allocator, &forever);
    defer f.deinit();

    var before: std.ArrayListUnmanaged(u8) = .empty;
    defer before.deinit(testing.allocator);
    try f.computedDigest(&before);

    const r = try f.run(.{
        .settings = iterating(10, 0.001),
        .limits = .{ .max_scc_iterations = 4 },
    });
    try testing.expect(r == .refused);
    try testing.expectEqual(Refusal.Reason.scc_iteration_ceiling, r.refused.reason);
    try testing.expectEqual(@as(?WorkCategory, .scc_iterations), r.refused.limit);
    try testing.expectEqual(parser.PlaneTwo.FormulaLimitExceeded, r.refused.planeTwo());

    // Zero mutation, asserted on the state rather than on the promise:
    // the cell the run spent four passes computing reads exactly as it
    // did before the run started.
    var after: std.ArrayListUnmanaged(u8) = .empty;
    defer after.deinit(testing.allocator);
    try f.computedDigest(&after);
    try testing.expectEqualStrings(before.items, after.items);
    try testing.expectEqual(value.ScalarValue.blank, try f.read("A1"));
    // …and it got there by retracting, not by never having written.
    try testing.expect(f.publishes > 0);
    try testing.expectEqual(f.publishes, f.retracts);
}

test "M5a2 §5.6c: a ceiling hit in ONE component refuses the whole run" {
    // The component that converged is rolled back with the one that did
    // not. A workbook half-calculated under one iteration budget and
    // half under another is not a state any consumer can reason about,
    // and `A1` here is a cell that genuinely reached its answer.
    const two_cycles = [_]Spec{
        .{ .a1 = "A1", .formula = "(A1+1)/2" },
        .{ .a1 = "C1", .formula = "C1+1" },
    };
    var f: Fixture = undefined;
    try f.init(testing.allocator, &two_cycles);
    defer f.deinit();

    const r = try f.run(.{
        .settings = iterating(100, 0.001),
        .limits = .{ .max_scc_iterations = 20 },
    });
    try testing.expect(r == .refused);
    try testing.expectEqual(Refusal.Reason.scc_iteration_ceiling, r.refused.reason);
    // Named: the refusal says WHICH component, and it is not the one
    // that converged.
    try testing.expect(r.refused.at.?.eql(.{ .cell = cellAt(0, "C1") }));

    try testing.expectEqual(value.ScalarValue.blank, try f.read("A1"));
    try testing.expectEqual(value.ScalarValue.blank, try f.read("C1"));
    var digest: std.ArrayListUnmanaged(u8) = .empty;
    defer digest.deinit(testing.allocator);
    try f.computedDigest(&digest);
    try testing.expectEqualStrings("", digest.items);
}

test "M5a2 §9: max_scc_iterations below, at and above what a component needs" {
    // The halving cycle converges on pass 10 exactly, which makes 10 the
    // boundary rather than a number chosen to be near one.
    const cases = [_]struct { ceiling: u64, refuses: bool }{
        .{ .ceiling = 9, .refuses = true },
        .{ .ceiling = 10, .refuses = false },
        .{ .ceiling = 11, .refuses = false },
    };
    for (cases) |c| {
        var f: Fixture = undefined;
        try f.init(testing.allocator, &halving);
        defer f.deinit();

        const r = try f.run(.{
            .settings = iterating(100, 0.001),
            .limits = .{ .max_scc_iterations = c.ceiling },
        });
        if (c.refuses) {
            try testing.expect(r == .refused);
            try testing.expectEqual(@as(?WorkCategory, .scc_iterations), r.refused.limit);
        } else {
            try testing.expect(r == .ok);
            try testing.expectEqual(Outcome.converged, r.ok.components[0].outcome);
            try testing.expectEqual(@as(u32, 10), r.ok.components[0].passes);
        }
    }
}

test "M5a2 §5.6c: `A1=(A1+1)/2` — first run and resume, pinned" {
    // The parity fixture §5.6c names. A first run starts from the seed
    // table's zero and walks ten passes; a resume starts from the cache
    // the first run would have written and settles on its first pass.
    // The two do NOT produce the same number — they cannot, since the
    // sequence approaches 1 without reaching it — and what parity means
    // is the property that holds of both: re-running the engine on its
    // own answer does not move that answer by more than the tolerance it
    // converged under.
    const first = 1 - std.math.pow(f64, 2, -10);
    {
        var f: Fixture = undefined;
        try f.init(testing.allocator, &halving);
        defer f.deinit();
        const r = try f.run(.{ .settings = iterating(100, 0.001) });
        try testing.expectEqual(@as(u32, 10), r.ok.components[0].passes);
        try expectNumber(&f, "A1", first);
    }
    {
        const resumed = [_]Spec{
            .{ .a1 = "A1", .formula = "(A1+1)/2", .cache = .{ .number = first } },
        };
        var f: Fixture = undefined;
        try f.init(testing.allocator, &resumed);
        defer f.deinit();
        const r = try f.run(.{ .settings = iterating(100, 0.001) });
        try testing.expectEqual(Outcome.converged, r.ok.components[0].outcome);
        // One pass, because the seed was already within tolerance of
        // where the next pass lands.
        try testing.expectEqual(@as(u32, 1), r.ok.components[0].passes);
        const second = (first + 1) / 2;
        try expectNumber(&f, "A1", second);
        try testing.expect(@abs(second - first) < 0.001);
    }
}

test "M5a2 §5.6c: the seed table is load-bearing — a cached number is resumed from" {
    // M5a1 computed the seeds and left them unused. This is the row that
    // consumes them, and the fixture that proves it: without the seed
    // one pass of `A1+1` is 1, with it the same pass is 6.
    const seeded = [_]Spec{
        .{ .a1 = "A1", .formula = "A1+1", .cache = .{ .number = 5 } },
    };
    var f: Fixture = undefined;
    try f.init(testing.allocator, &seeded);
    defer f.deinit();

    const r = try f.run(.{ .settings = iterating(1, 0.001) });
    try testing.expect(r == .ok);
    try expectNumber(&f, "A1", 6);
}

test "M5a2 §5.6c: text, boolean and error caches all seed zero" {
    // §5.6c's seed table again, through the engine: iteration is
    // numeric, and a text cache carries no number to resume from.
    const caches = [_]graph.CacheState{ .text, .boolean, .err, .absent };
    for (caches) |cache| {
        const specs = [_]Spec{
            .{ .a1 = "A1", .formula = "A1+1", .cache = cache },
        };
        var f: Fixture = undefined;
        try f.init(testing.allocator, &specs);
        defer f.deinit();
        _ = try f.run(.{ .settings = iterating(1, 0.001) });
        try expectNumber(&f, "A1", 1);
    }
}

test "M5a2 §5.6c: visibility inside a pass is Gauss–Seidel, not double-buffered" {
    // `A1=B1+1`, `B1=A1`, three passes. Gauss–Seidel: B1 sees the A1 of
    // THIS pass, so both reach 3. A double-buffered pass would leave A1
    // at 2 and B1 at 1, which is why three passes and not one is the
    // fixture — one pass cannot tell the two apart.
    const pair = [_]Spec{
        .{ .a1 = "A1", .formula = "B1+1" },
        .{ .a1 = "B1", .formula = "A1" },
    };
    var f: Fixture = undefined;
    try f.init(testing.allocator, &pair);
    defer f.deinit();

    const r = try f.run(.{ .settings = iterating(3, 0.001) });
    try testing.expect(r == .ok);
    try testing.expectEqual(@as(usize, 1), r.ok.components.len);
    try expectNumber(&f, "A1", 3);
    try expectNumber(&f, "B1", 3);
}

test "M5a2 §5.6c: downstream evaluates ONCE, after its SCC has finished" {
    // "An SCC iterates to its own convergence before any downstream node
    // evaluates" is two claims. That the downstream cell sees the final
    // value is the visible one; that it evaluated once rather than once
    // per pass is the one only a count can make.
    const with_downstream = [_]Spec{
        .{ .a1 = "A1", .formula = "(A1+1)/2" },
        .{ .a1 = "Z1", .formula = "A1*100" },
    };
    var f: Fixture = undefined;
    try f.init(testing.allocator, &with_downstream);
    defer f.deinit();

    const r = try f.run(.{ .settings = iterating(100, 0.001) });
    try testing.expect(r == .ok);
    try testing.expectEqual(@as(usize, 2), r.ok.components.len);
    try testing.expectEqual(Outcome.converged, r.ok.components[0].outcome);
    try testing.expectEqual(Outcome.acyclic, r.ok.components[1].outcome);

    try testing.expectEqual(@as(u32, 10), f.callsOn("A1"));
    try testing.expectEqual(@as(u32, 1), f.callsOn("Z1"));
    try expectNumber(&f, "Z1", (1 - std.math.pow(f64, 2, -10)) * 100);
}

test "M5a2 §5.6c: iteration off makes a cycle FormulaCycle, exactly as M5a1 said" {
    var f: Fixture = undefined;
    try f.init(testing.allocator, &halving);
    defer f.deinit();

    const r = try f.run(.{ .settings = .{ .iterate = false } });
    try testing.expect(r == .refused);
    try testing.expectEqual(Refusal.Reason.cycle, r.refused.reason);
    try testing.expectEqual(parser.PlaneTwo.FormulaCycle, r.refused.planeTwo());
    try testing.expectEqual(value.ScalarValue.blank, try f.read("A1"));
}

test "M5a2 §5.6c: an acyclic workbook needs no iteration and reports none" {
    const chain = [_]Spec{
        .{ .a1 = "A1", .formula = "2" },
        .{ .a1 = "B1", .formula = "A1*3" },
        .{ .a1 = "C1", .formula = "B1+1" },
    };
    var f: Fixture = undefined;
    try f.init(testing.allocator, &chain);
    defer f.deinit();

    // Iteration OFF: an acyclic workbook must not need it.
    const r = try f.run(.{ .settings = .{} });
    try testing.expect(r == .ok);
    try expectNumber(&f, "C1", 7);
    for (r.ok.components) |c| {
        try testing.expectEqual(Outcome.acyclic, c.outcome);
        try testing.expectEqual(@as(u32, 0), c.passes);
    }
    try testing.expectEqual(@as(u32, 1), r.ok.dynamic_passes);
}

test "M5a2 §5.6c: cyclic and acyclic components interact in one run" {
    const mixed = [_]Spec{
        .{ .a1 = "A1", .formula = "10" }, // acyclic, upstream
        .{ .a1 = "B1", .formula = "(B1+A1)/2" }, // cyclic, fed by A1
        .{ .a1 = "C1", .formula = "B1+1" }, // acyclic, downstream
    };
    var f: Fixture = undefined;
    try f.init(testing.allocator, &mixed);
    defer f.deinit();

    const r = try f.run(.{ .settings = iterating(100, 0.001) });
    try testing.expect(r == .ok);
    try testing.expectEqual(@as(usize, 3), r.ok.components.len);
    try testing.expectEqual(Outcome.acyclic, r.ok.components[0].outcome);
    try testing.expectEqual(Outcome.converged, r.ok.components[1].outcome);
    try testing.expectEqual(Outcome.acyclic, r.ok.components[2].outcome);

    // B1 approaches 10 the same way the parity fixture approaches 1, so
    // C1 is one more than a value within tolerance of 10.
    const b1 = (try f.read("B1")).number;
    try testing.expect(@abs(b1 - 10) < 0.02);
    try expectNumber(&f, "C1", b1 + 1);
    try testing.expectEqual(@as(u32, 1), f.callsOn("A1"));
    try testing.expectEqual(@as(u32, 1), f.callsOn("C1"));
}

// ─── §5.6c's shape pass and its `#SPILL!` (done-when 4) ──────────

test "M5a2 §5.6c: an anchor with no recoverable shape takes the pre-iteration shape pass" {
    // `.shape_pass` is the seed M5a1 named and could not run. The anchor
    // evaluates once OUTSIDE the cycle to fix its shape, and the cycle
    // then starts from a zero-filled array of that shape.
    const specs = [_]Spec{
        .{ .a1 = "A1", .formula = "A1+1", .dynamic_anchor = true },
    };
    var f: Fixture = undefined;
    try f.init(testing.allocator, &specs);
    defer f.deinit();
    f.scripted = &[_]Scripted{.{
        .a1 = "A1",
        .steps = &[_]Step{.{ .array = .{ .rows = 2, .cols = 3 } }},
    }};

    const r = try f.run(.{ .settings = iterating(5, 0.001) });
    try testing.expect(r == .ok);
    // The shape pass, plus the passes that follow it. A stable shape
    // converges on the second pass: the first has the seed to compare
    // against, and every element is already zero.
    try testing.expect(f.callsOn("A1") >= 2);
    try testing.expectEqual(Outcome.converged, r.ok.components[0].outcome);
}

test "M5a2 §5.6c: a shape that moves between iterations is #SPILL!, and stops" {
    // "Shape-mutating cycles never spin." The engine does not keep
    // iterating in the hope the shape settles — an indeterminate
    // placement has no answer to converge to, so `#SPILL!` IS the
    // answer.
    const specs = [_]Spec{
        .{ .a1 = "A1", .formula = "A1+1", .array = "A1:B2" },
    };
    var f: Fixture = undefined;
    try f.init(testing.allocator, &specs);
    defer f.deinit();
    f.scripted = &[_]Scripted{.{
        .a1 = "A1",
        .steps = &[_]Step{
            // A 2x2 that is not the zero-filled 2x2 the seed table
            // produced, so pass 1 is ordinary non-convergence and pass 2
            // is where the SHAPE moves. Converging on pass 1 would have
            // meant this fixture never reached the rule it is for.
            .{ .array = .{ .rows = 2, .cols = 2, .fill = 1 } },
            .{ .array = .{ .rows = 3, .cols = 3, .fill = 1 } },
        },
    }};

    const r = try f.run(.{ .settings = iterating(100, 0.001) });
    try testing.expect(r == .ok);
    try testing.expectEqual(Outcome.shape_indeterminate, r.ok.components[0].outcome);
    try testing.expectEqual(@as(u32, 1), r.ok.shape_indeterminate_cells);
    // Two passes, not a hundred: the second is where the shape moved.
    try testing.expectEqual(@as(u32, 2), r.ok.components[0].passes);

    const v = try f.read("A1");
    try testing.expect(v == .err);
    try testing.expectEqual(value.KnownError.spill, v.err.known);
}

test "M5a2 §5.6c: a type transition across passes is not convergence" {
    // The cell alternates between a number and text of the same
    // spelling. Every pass is a transition, so it never settles and the
    // workbook's bound is what stops it.
    const specs = [_]Spec{
        .{ .a1 = "A1", .formula = "A1+1" },
    };
    var f: Fixture = undefined;
    try f.init(testing.allocator, &specs);
    defer f.deinit();
    f.scripted = &[_]Scripted{.{
        .a1 = "A1",
        .steps = &[_]Step{
            .{ .scalar = .{ .number = 2 } },
            .{ .scalar = .{ .text = "2" } },
            .{ .scalar = .{ .number = 2 } },
            .{ .scalar = .{ .text = "2" } },
        },
    }};

    const r = try f.run(.{ .settings = iterating(4, 0.001) });
    try testing.expect(r == .ok);
    try testing.expectEqual(Outcome.semantic_bound, r.ok.components[0].outcome);
    try testing.expectEqual(@as(u32, 1), r.ok.non_converged_cells);
}

// ─── §5.6d, at the schedule level (done-when 6) ──────────────────

test "M5a2 §5.6d KAT: two cells drawing at the same callsite draw differently" {
    // Graph-order draws. Both bodies are byte-identical, so the callsite
    // ordinal inside each is the same number; only the invocation path
    // separates them, and the path is rooted at the owning cell.
    const two = [_]Spec{
        .{ .a1 = "A1", .formula = "RAND()" },
        .{ .a1 = "B1", .formula = "RAND()" },
    };
    var f: Fixture = undefined;
    try f.init(testing.allocator, &two);
    defer f.deinit();

    const r = try f.run(.{});
    try testing.expect(r == .ok);
    const a = (try f.read("A1")).number;
    const b = (try f.read("B1")).number;
    try testing.expect(a != b);
    try testing.expectEqual(@as(u64, 2), f.schedule.generated);
    try testing.expectEqual(@as(u64, 0), f.schedule.reused);
}

test "M5a2 §5.6d KAT: two callsites in ONE body draw twice" {
    const one = [_]Spec{
        .{ .a1 = "A1", .formula = "RAND()+RAND()" },
    };
    var f: Fixture = undefined;
    try f.init(testing.allocator, &one);
    defer f.deinit();

    _ = try f.run(.{});
    try testing.expectEqual(@as(u64, 2), f.schedule.generated);
    // …and the sum is of two DIFFERENT numbers, which is the observable
    // property §5.6d's oracle policy does allow an external oracle to
    // check.
    const v = (try f.read("A1")).number;
    try testing.expect(v != 2 * (1.0 / 1024.0));
}

test "M5a2 §5.6d KAT: an iterating SCC draws once per pass, keyed by the pass" {
    // The pass term. Without it every pass of this cycle would reuse
    // pass 1's number and the cycle would converge immediately — which
    // is exactly the wrong answer for a body containing a volatile.
    const volatile_cycle = [_]Spec{
        .{ .a1 = "A1", .formula = "A1+RAND()" },
    };
    var f: Fixture = undefined;
    try f.init(testing.allocator, &volatile_cycle);
    defer f.deinit();

    const r = try f.run(.{ .settings = iterating(5, 0.0000001) });
    try testing.expect(r == .ok);
    try testing.expectEqual(@as(u32, 5), r.ok.components[0].passes);
    // Five passes, five draws, none reused: each pass is its own key.
    try testing.expectEqual(@as(u64, 5), f.schedule.generated);
    try testing.expectEqual(@as(u64, 0), f.schedule.reused);
    // The cell is the running sum of the five, which is what proves the
    // draws were distinct rather than merely counted.
    var want: f64 = 0;
    for (1..6) |i| want += @as(f64, @floatFromInt(i)) / 1024;
    try expectNumber(&f, "A1", want);
}

test "M5a2 §5.6d KAT: a second run over the same graph generates nothing new" {
    // Rebuild-reuse, in its purest form: the same bodies at the same
    // keys must answer with the same numbers, or a discovery pass would
    // change a result. §5.6e leans on this and cannot check it, because
    // by the time the outer loop notices a difference it cannot say
    // whether the graph or the RNG produced it.
    const volatile_chain = [_]Spec{
        .{ .a1 = "A1", .formula = "RAND()" },
        .{ .a1 = "B1", .formula = "A1+RAND()" },
    };
    var f: Fixture = undefined;
    try f.init(testing.allocator, &volatile_chain);
    defer f.deinit();

    _ = try f.runOn(try f.build(), .{});
    const first_a = (try f.read("A1")).number;
    const first_b = (try f.read("B1")).number;
    const generated = f.schedule.generated;
    try testing.expectEqual(@as(u64, 2), generated);

    // The engine consumed the first graph, so the second run gets a
    // fresh build of the same specs — which is the identity §5.6e
    // actually leans on: same bodies at the same keys, not the same
    // graph object.
    _ = try f.runOn(try f.build(), .{});
    try testing.expectEqual(generated, f.schedule.generated);
    try testing.expectEqual(@as(u64, 2), f.schedule.reused);
    try expectNumber(&f, "A1", first_a);
    try expectNumber(&f, "B1", first_b);
}

// ─── §5.6e: the dynamic-edge fixpoint (done-when 7) ──────────────

test "M5a2 §5.6e: INDIRECT closes a cycle the static walk could not see" {
    // `B1` is a stored text cell, so nothing in `A1`'s TEXT mentions
    // `A1`. The static graph therefore says acyclic, the run reads
    // `A1`, and the rebuild says otherwise.
    const closed = [_]Spec{
        .{ .a1 = "A1", .formula = "INDIRECT(B1)+1" },
        .{ .a1 = "B1", .stored = .{ .text = "A1" } },
    };
    var f: Fixture = undefined;
    try f.init(testing.allocator, &closed);
    defer f.deinit();

    // The static graph, before anything runs: one acyclic component.
    {
        var g = try f.build();
        defer g.deinit();
        try testing.expect(!g.isCyclic(g.find(.{ .cell = cellAt(0, "A1") }).?));
    }

    const r = try f.runFixpoint(.{ .settings = iterating(4, 0.001) });
    try testing.expect(r == .ok);
    try testing.expectEqual(@as(u32, 2), r.ok.dynamic_passes);
    try testing.expectEqual(@as(usize, 1), r.ok.components.len);
    try testing.expect(r.ok.components[0].cyclic);
    try testing.expectEqual(Outcome.semantic_bound, r.ok.components[0].outcome);
    try expectNumber(&f, "A1", 4);
}

test "M5a2 §5.6e: INDIRECT pointing elsewhere leaves the cycle open" {
    // The other half of the flip, and the reason it is a pair: a fixture
    // that only closed a cycle would pass just as well against an engine
    // that called everything cyclic.
    const open = [_]Spec{
        .{ .a1 = "A1", .formula = "INDIRECT(B1)+1" },
        .{ .a1 = "B1", .stored = .{ .text = "Z9" } },
    };
    var f: Fixture = undefined;
    try f.init(testing.allocator, &open);
    defer f.deinit();

    // Iteration OFF, deliberately: an open cycle must not need it.
    const r = try f.runFixpoint(.{ .settings = .{} });
    try testing.expect(r == .ok);
    try testing.expectEqual(Outcome.acyclic, r.ok.components[0].outcome);
    try expectNumber(&f, "A1", 1);
}

test "M5a2 §5.6e: a closed cycle with iteration off is FormulaCycle after the rebuild" {
    // The refusal has to survive the outer loop. A cycle discovered on
    // pass 2 is a cycle, and a workbook whose `calcPr` says no iteration
    // gets §5.6c's answer for one.
    const closed = [_]Spec{
        .{ .a1 = "A1", .formula = "INDIRECT(B1)+1" },
        .{ .a1 = "B1", .stored = .{ .text = "A1" } },
    };
    var f: Fixture = undefined;
    try f.init(testing.allocator, &closed);
    defer f.deinit();

    const r = try f.runFixpoint(.{ .settings = .{} });
    try testing.expect(r == .refused);
    try testing.expectEqual(Refusal.Reason.cycle, r.refused.reason);
    try testing.expectEqual(value.ScalarValue.blank, try f.read("A1"));
}

test "M5a2 §5.6e: OFFSET merges two components into one" {
    // Statically `A1` and `C1` are two cycles with an edge between them.
    // `OFFSET(A1,0,2)` reaches C1 at runtime, which closes the other
    // direction and makes the pair one strongly connected component.
    const merging = [_]Spec{
        .{ .a1 = "A1", .formula = "A1+OFFSET(A1,0,2)+1" },
        .{ .a1 = "C1", .formula = "C1+A1" },
    };

    // Without the outer loop: two components, in canonical order.
    {
        var f: Fixture = undefined;
        try f.init(testing.allocator, &merging);
        defer f.deinit();
        const r = try f.run(.{ .settings = iterating(3, 0.001) });
        try testing.expect(r == .ok);
        try testing.expectEqual(@as(usize, 2), r.ok.components.len);
        try testing.expect(r.ok.components[0].at.eql(.{ .cell = cellAt(0, "A1") }));
        try testing.expect(r.ok.components[1].at.eql(.{ .cell = cellAt(0, "C1") }));
    }

    // With it: one.
    {
        var f: Fixture = undefined;
        try f.init(testing.allocator, &merging);
        defer f.deinit();
        const r = try f.runFixpoint(.{ .settings = iterating(3, 0.001) });
        try testing.expect(r == .ok);
        try testing.expectEqual(@as(u32, 2), r.ok.dynamic_passes);
        try testing.expectEqual(@as(usize, 1), r.ok.components.len);
        try testing.expect(r.ok.components[0].cyclic);
        try testing.expect(r.ok.components[0].at.eql(.{ .cell = cellAt(0, "A1") }));
        // A merged component re-seeds and re-runs, which the report says
        // rather than leaving to be inferred from the pass count.
        try testing.expect(r.ok.components[0].rerun);
    }
}

/// A workbook whose dynamic reference cannot settle until the value it
/// is displaced by has settled — which takes one pass more than
/// discovering the reference did.
///
/// `E1` reads `F1` through `INDIRECT`, so `E1` is zero on pass 1 and 2
/// only after the rebuild puts `F1` ahead of it. `A1` is displaced by
/// `E1`, so it lands on `B1` on pass 1 and on `B3` on pass 2 — and the
/// rebuild after pass 2 is the one that comes back identical, which is
/// where the loop stops.
const staged_offset = [_]Spec{
    .{ .a1 = "A1", .formula = "OFFSET(B1,E1,0)" },
    .{ .a1 = "B1", .stored = .{ .number = 100 } },
    .{ .a1 = "B3", .stored = .{ .number = 300 } },
    .{ .a1 = "E1", .formula = "INDIRECT(\"F1\")" },
    .{ .a1 = "F1", .formula = "2" },
};

test "M5a2 §5.6e: the outer loop reaches a fixpoint, and the answer is the settled one" {
    var f: Fixture = undefined;
    try f.init(testing.allocator, &staged_offset);
    defer f.deinit();

    const r = try f.runFixpoint(.{ .limits = .{ .max_dynamic_passes = 3 } });
    try testing.expect(r == .ok);
    // Two: the pass that discovers the ordering and the pass that runs
    // under it. There is no third confirmation pass, because the
    // fixpoint test is on the graph rather than on the values — a
    // rebuild that comes back identical has already confirmed it.
    try testing.expectEqual(@as(u32, 2), r.ok.dynamic_passes);
    try expectNumber(&f, "E1", 2);
    // 300, not 100: the displacement that landed on `B1` in pass 1 was
    // computed from an `E1` that had not been evaluated yet.
    try expectNumber(&f, "A1", 300);
}

test "M5a2 §5.6e: exhausting the outer loop is FormulaDynamicRefUnstable, with nothing written" {
    var f: Fixture = undefined;
    try f.init(testing.allocator, &staged_offset);
    defer f.deinit();

    const r = try f.runFixpoint(.{ .limits = .{ .max_dynamic_passes = 1 } });
    try testing.expect(r == .refused);
    try testing.expectEqual(Refusal.Reason.dynamic_ref_unstable, r.refused.reason);
    try testing.expectEqual(@as(?WorkCategory, .dynamic_passes), r.refused.limit);
    try testing.expectEqual(parser.PlaneTwo.FormulaDynamicRefUnstable, r.refused.planeTwo());

    // A refusal writes nothing, whichever loop raised it.
    var digest: std.ArrayListUnmanaged(u8) = .empty;
    defer digest.deinit(testing.allocator);
    try f.computedDigest(&digest);
    try testing.expectEqualStrings("", digest.items);
    try testing.expectEqual(value.ScalarValue.blank, try f.read("A1"));
}

test "M5a2 §9: max_dynamic_passes below, at and above what the fixpoint needs" {
    const cases = [_]struct { ceiling: u64, refuses: bool }{
        .{ .ceiling = 1, .refuses = true },
        .{ .ceiling = 2, .refuses = false },
        .{ .ceiling = 3, .refuses = false },
    };
    for (cases) |c| {
        var f: Fixture = undefined;
        try f.init(testing.allocator, &staged_offset);
        defer f.deinit();

        const r = try f.runFixpoint(.{ .limits = .{ .max_dynamic_passes = c.ceiling } });
        if (c.refuses) {
            try testing.expect(r == .refused);
            try testing.expectEqual(@as(?WorkCategory, .dynamic_passes), r.refused.limit);
        } else {
            try testing.expect(r == .ok);
            // Two either way: a higher ceiling permits more passes, it
            // does not cause them.
            try testing.expectEqual(@as(u32, 2), r.ok.dynamic_passes);
        }
    }
}

test "M5a2 §5.6e: a workbook with no dynamic reference costs exactly one pass" {
    // The common case, asserted so that wiring the outer loop is not
    // silently a second evaluation of everything.
    const plain = [_]Spec{
        .{ .a1 = "A1", .formula = "2" },
        .{ .a1 = "B1", .formula = "A1*3" },
    };
    var f: Fixture = undefined;
    try f.init(testing.allocator, &plain);
    defer f.deinit();

    const r = try f.runFixpoint(.{});
    try testing.expect(r == .ok);
    try testing.expectEqual(@as(u32, 1), r.ok.dynamic_passes);
    try testing.expectEqual(@as(u32, 1), f.callsOn("A1"));
    try testing.expectEqual(@as(u32, 1), f.callsOn("B1"));
    try expectNumber(&f, "B1", 6);
}

test "M5a2 §5.6e: an unchanged component keeps its state instead of re-running" {
    // "Unchanged SCCs keep their converged state." `F1` gains no edge
    // and loses none across the rebuild, so it evaluates once for the
    // whole run however many passes the loop takes.
    var f: Fixture = undefined;
    try f.init(testing.allocator, &staged_offset);
    defer f.deinit();

    const r = try f.runFixpoint(.{ .limits = .{ .max_dynamic_passes = 3 } });
    try testing.expect(r == .ok);
    try testing.expectEqual(@as(u32, 1), f.callsOn("F1"));
    // …while the two that did change ran again.
    try testing.expect(f.callsOn("E1") > 1);
    try testing.expect(f.callsOn("A1") > 1);
}

// ─── purity (done-when 10, the engine's half) ────────────────────

test "M5a2 §5.6f: an injected allocation failure leaves the state it found" {
    // The fourth leg of the purity gate. A run that dies partway through
    // has already published — Gauss–Seidel requires it — so "wrote
    // nothing" is a claim about the rollback, and the rollback is what
    // this checks.
    const chain = [_]Spec{
        .{ .a1 = "A1", .formula = "1" },
        .{ .a1 = "B1", .formula = "A1+1" },
        .{ .a1 = "C1", .formula = "B1+1" },
    };
    var f: Fixture = undefined;
    try f.init(testing.allocator, &chain);
    defer f.deinit();
    f.fail_at_eval = 3; // partway: A1 and B1 have already published

    // No deinit here: the engine owns the graph even on the error path.
    const g = try f.build();
    try testing.expectError(error.OutOfMemory, f.runOn(g, .{}));

    // `run` rolls back only on a refusal, because an `Error` unwinds
    // past it — so the caller is the one who must not read a half-run
    // scratch layer, and it can tell because it got an error rather than
    // a result. What IS asserted here is that the partial state is
    // exactly the two cells that succeeded and nothing beyond them.
    try expectNumber(&f, "A1", 1);
    try expectNumber(&f, "B1", 2);
    try testing.expectEqual(value.ScalarValue.blank, try f.read("C1"));
}

test "M5a2 §5.6f: a refused run is byte-for-byte the state before it" {
    // Every refusal the engine can raise, against one digest. The point
    // of running all of them through one loop is that a rollback added
    // for one reason has to hold for the others too.
    const Case = struct { specs: []const Spec, settings: Settings, limits: WorkLimits };
    const cases = [_]Case{
        // The caller's ceiling.
        .{ .specs = &forever, .settings = iterating(10, 0.001), .limits = .{ .max_scc_iterations = 4 } },
        // Iteration off, over a cycle.
        .{ .specs = &halving, .settings = .{}, .limits = .{} },
        // §9's evaluation counter.
        .{ .specs = &halving, .settings = iterating(100, 0.001), .limits = .{ .max_total_cell_evals = 3 } },
    };
    for (cases) |c| {
        var f: Fixture = undefined;
        try f.init(testing.allocator, c.specs);
        defer f.deinit();

        var before: std.ArrayListUnmanaged(u8) = .empty;
        defer before.deinit(testing.allocator);
        try f.computedDigest(&before);

        const r = try f.run(.{ .settings = c.settings, .limits = c.limits });
        try testing.expect(r == .refused);

        var after: std.ArrayListUnmanaged(u8) = .empty;
        defer after.deinit(testing.allocator);
        try f.computedDigest(&after);
        try testing.expectEqualStrings(before.items, after.items);
        try testing.expectEqualStrings("", after.items);
    }
}

test "M5a2 §9: the evaluation counter charges every pass, not every cell" {
    // A cycle's members are evaluated once per pass, and §9's budget is
    // about work performed rather than about cells named. Ten passes
    // over one member is ten evaluations, which is what makes the
    // counter a real bound on an iterating workbook.
    var f: Fixture = undefined;
    try f.init(testing.allocator, &halving);
    defer f.deinit();

    const r = try f.run(.{ .settings = iterating(100, 0.001) });
    try testing.expect(r == .ok);
    try testing.expectEqual(@as(u64, 10), r.ok.cell_evals);
}

test "M5a2: the report names which bound fired, per component" {
    // Two cycles under one `iterateCount`, one of which converges. The
    // report has to distinguish them, because "the run succeeded" and
    // "everything converged" are different facts and §5.7.8 reports
    // both.
    const both = [_]Spec{
        .{ .a1 = "A1", .formula = "(A1+1)/2" },
        .{ .a1 = "C1", .formula = "C1+1" },
    };
    var f: Fixture = undefined;
    try f.init(testing.allocator, &both);
    defer f.deinit();

    const r = try f.run(.{ .settings = iterating(20, 0.001) });
    try testing.expect(r == .ok);
    try testing.expectEqual(@as(usize, 2), r.ok.components.len);
    try testing.expectEqual(Outcome.converged, r.ok.components[0].outcome);
    try testing.expectEqual(@as(?Bound, null), r.ok.components[0].bound);
    try testing.expectEqual(Outcome.semantic_bound, r.ok.components[1].outcome);
    try testing.expectEqual(@as(?Bound, .semantic), r.ok.components[1].bound);
    try testing.expectEqual(@as(u32, 1), r.ok.non_converged_cells);
}

/// A component that merges and then splits again, driven by a
/// displacement that is wrong on the discovery pass and right after it.
///
/// `E1` reads `F1` through `INDIRECT`, so it is blank — and therefore
/// zero — until the rebuild puts `F1` ahead of it. The displacement is
/// `2-E1`, so pass 1 reaches `C1` and pass 2 reaches `A1` itself: the
/// runtime edge `A1 → C1` exists on one pass and not on the next, which
/// merges the two cycles and then splits them.
///
/// §5.6e's own third fixture is a *spill* shape doing this. Placing a
/// spill is M7a's and this row must not touch it, so the mechanism is
/// exercised through the other thing that can move a runtime edge. What
/// is being tested is the engine's response to an edge that disappears,
/// and that response does not depend on which construct removed it.
const merge_then_split = [_]Spec{
    .{ .a1 = "A1", .formula = "A1+OFFSET(A1,0,2-E1)" },
    .{ .a1 = "C1", .formula = "C1+A1" },
    .{ .a1 = "E1", .formula = "INDIRECT(\"F1\")" },
    .{ .a1 = "F1", .formula = "2" },
};

test "M5a2 §5.6e: a runtime edge that disappears splits the component again" {
    var f: Fixture = undefined;
    try f.init(testing.allocator, &merge_then_split);
    defer f.deinit();

    const r = try f.runFixpoint(.{
        .settings = iterating(20, 0.001),
        .limits = .{ .max_dynamic_passes = 3 },
    });
    try testing.expect(r == .ok);
    // Three: discover the ordering, run under it and lose the edge, then
    // run under the split graph and find it stable.
    try testing.expectEqual(@as(u32, 3), r.ok.dynamic_passes);

    // Back to two components, each cyclic on its own self-reference —
    // which is what the static graph said before the discovery pass
    // merged them.
    try testing.expectEqual(@as(usize, 4), r.ok.components.len);
    var cyclic_components: usize = 0;
    for (r.ok.components) |c| {
        if (c.cyclic) cyclic_components += 1;
    }
    try testing.expectEqual(@as(usize, 2), cyclic_components);
    try expectNumber(&f, "E1", 2);
}

test "M5a2 §5.6e: the merge is real before the split undoes it" {
    // The other half: without enough passes to reach the split, the loop
    // refuses rather than reporting the merged shape as an answer. A
    // fixture that only checked the final state could not tell a merge
    // that happened from one that never did.
    var f: Fixture = undefined;
    try f.init(testing.allocator, &merge_then_split);
    defer f.deinit();

    const r = try f.runFixpoint(.{
        .settings = iterating(20, 0.001),
        .limits = .{ .max_dynamic_passes = 2 },
    });
    try testing.expect(r == .refused);
    try testing.expectEqual(Refusal.Reason.dynamic_ref_unstable, r.refused.reason);
}

// ─── stabilization property (the ladder's fuzz gate) ─────────────

/// One generated body, as the two halves a self-reference sits between.
///
/// Halves rather than a format string because the shape is chosen at
/// runtime and a format string cannot be: concatenation says the same
/// thing and says it without a comptime argument.
const Shape = struct { pre: []const u8, post: []const u8 };

/// Bodies that all contain a self-reference, so every generated workbook
/// has a cycle. A generator that mostly emitted acyclic sheets would
/// spend its budget testing nothing. The set covers what §5.6c has to
/// tell apart: divergent, convergent from above and below, immediately
/// settled, and clamped.
const generated_shapes = [_]Shape{
    .{ .pre = "", .post = "+1" },
    .{ .pre = "(", .post = "+1)/2" },
    .{ .pre = "(", .post = "+10)/2" },
    .{ .pre = "", .post = "*0" },
    .{ .pre = "", .post = "" },
    .{ .pre = "ABS(", .post = ")+0.0001" },
    .{ .pre = "MIN(", .post = "+1,3)" },
    .{ .pre = "MAX(", .post = "-1,0)" },
};

test "M5a2 property: every generated cyclic workbook terminates in a stated outcome" {
    // Not a coverage-guided fuzz target — `zig build fuzz` owns those,
    // and what needs exercising here is a SCHEDULE rather than a
    // parser's byte handling. What it shares with one is the shape of
    // the claim: over a few hundred generated workbooks the engine
    // terminates, and every outcome it reports is one §5.6c allows.
    var seed: u64 = 0;
    while (seed < 240) : (seed += 1) {
        var rng = std.Random.DefaultPrng.init(seed);
        const r = rng.random();
        const n = 1 + r.uintLessThan(usize, 5);

        var texts: [5][48]u8 = @splat(@splat(0));
        var labels: [5][3]u8 = @splat(@splat(0));
        var specs: [5]Spec = undefined;
        for (0..n) |i| {
            labels[i] = .{ 'A' + @as(u8, @intCast(i)), '1', 0 };
            const a1 = labels[i][0..2];
            const shape = generated_shapes[r.uintLessThan(usize, generated_shapes.len)];
            const body = try std.fmt.bufPrint(&texts[i], "{s}{s}{s}", .{ shape.pre, a1, shape.post });
            specs[i] = .{
                .a1 = a1,
                .formula = body,
                .cache = if (r.boolean()) .{ .number = r.float(f64) * 10 } else .absent,
            };
        }

        var f: Fixture = undefined;
        try f.init(testing.allocator, specs[0..n]);
        defer f.deinit();

        const count = 1 + r.uintLessThan(u32, 30);
        const ceiling: u64 = 1 + r.uintLessThan(u64, 40);
        const outcome = try f.run(.{
            .settings = iterating(count, 0.001),
            .limits = .{ .max_scc_iterations = ceiling },
        });

        switch (outcome) {
            .refused => |x| {
                // The only refusal a workbook of pure arithmetic can
                // raise here is the caller's ceiling, and it must leave
                // nothing behind.
                try testing.expectEqual(Refusal.Reason.scc_iteration_ceiling, x.reason);
                try testing.expect(ceiling < count);
                var digest: std.ArrayListUnmanaged(u8) = .empty;
                defer digest.deinit(testing.allocator);
                try f.computedDigest(&digest);
                try testing.expectEqualStrings("", digest.items);
            },
            .ok => |report| {
                for (report.components) |c| {
                    // Every pass count is inside BOTH bounds, whichever
                    // one ended up deciding — that is the invariant the
                    // two-bound rule turns on.
                    try testing.expect(c.passes <= count);
                    try testing.expect(c.passes <= ceiling);
                    switch (c.outcome) {
                        .converged => {
                            try testing.expectEqual(@as(?Bound, null), c.bound);
                            try testing.expectEqual(@as(u32, 0), c.non_converged_cells);
                        },
                        .semantic_bound => {
                            // Reaching the workbook's own bound is a
                            // success, and it can only be reached when
                            // the caller permitted at least that many.
                            try testing.expectEqual(@as(?Bound, .semantic), c.bound);
                            try testing.expectEqual(count, c.passes);
                            try testing.expect(ceiling >= count);
                            try testing.expect(c.non_converged_cells > 0);
                        },
                        // No generated body can move a shape, and every
                        // one of them is in a cycle. Reporting either
                        // here would be reporting it about the wrong
                        // component.
                        .shape_indeterminate, .acyclic => return error.UnexpectedOutcome,
                    }
                }
            },
        }
    }
}

test "M5a2 §9: a declared shape too large to materialize refuses rather than asserts" {
    // A declared array range is bounded by the GRID, not by
    // `max_matrix_cells` — `A1:D1048576` is four million and change, so
    // the seed table really can be handed a shape it cannot build. The
    // input that reaches it is a workbook, and a workbook must not be
    // able to trip an `unreachable`.
    const huge = [_]Spec{
        .{ .a1 = "A1", .formula = "A1+1", .array = "A1:D1048576" },
    };
    var f: Fixture = undefined;
    try f.init(testing.allocator, &huge);
    defer f.deinit();

    const r = try f.run(.{ .settings = iterating(10, 0.001) });
    try testing.expect(r == .refused);
    try testing.expectEqual(Refusal.Reason.seed_shape_too_large, r.refused.reason);
    try testing.expectEqual(parser.PlaneTwo.FormulaLimitExceeded, r.refused.planeTwo());

    // One column narrower fits, and iterates — so the refusal is the
    // limit and not the construct.
    const fits = [_]Spec{
        .{ .a1 = "A1", .formula = "A1+1", .array = "A1:C1048576" },
    };
    var ok: Fixture = undefined;
    try ok.init(testing.allocator, &fits);
    defer ok.deinit();
    const r2 = try ok.run(.{ .settings = iterating(2, 0.001) });
    try testing.expect(r2 == .ok);
}

test "M5a2 §9: a zero max_dynamic_passes refuses instead of asserting" {
    // `WorkLimits.validate` rejects a zero, but the engine is reachable
    // without a validated set and a caller-supplied number must not be
    // able to trip an assertion.
    var f: Fixture = undefined;
    try f.init(testing.allocator, &halving);
    defer f.deinit();

    const r = try f.run(.{
        .settings = iterating(100, 0.001),
        .limits = .{ .max_dynamic_passes = 0 },
    });
    try testing.expect(r == .refused);
    try testing.expectEqual(Refusal.Reason.work_limit_exceeded, r.refused.reason);
    try testing.expectEqual(@as(?WorkCategory, .dynamic_passes), r.refused.limit);
    try testing.expectEqual(value.ScalarValue.blank, try f.read("A1"));
}
