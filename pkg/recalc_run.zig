//! §5.7's recalc pipeline, from a model build to a committed file, and
//! the two entry points a caller actually has (M5d2).
//!
//! What this file is
//! -----------------
//! Every stage of §5.7 existed before this row and none of them were
//! wired to each other. `WorkbookEnv.build` models a workbook,
//! `graph.build`/`iterate.run` evaluate one, `resolved.project`/`patch`
//! turn values into bytes, `recalc_txn.prepare` stages those bytes into
//! one candidate, and `PartStore.saveCommitted` publishes a candidate
//! atomically. This is the function that calls them in order, plus the
//! three gates §5.7 puts between them.
//!
//! Why the ordering in `saveWithRecalc` is not a style choice
//! ----------------------------------------------------------
//! §5.7.9 makes it normative: prepare fully → serialize the output bytes
//! **from the prepared, unswapped state** → write temp → `File.sync` →
//! final cancellation poll → rename → swap → directory fsync. Each
//! adjacency is load bearing.
//!
//!   * Serializing from the *unswapped* candidate is what lets a rename
//!     failure leave memory untouched. If the swap came first there
//!     would be no un-swapped state left to fall back to, and a failed
//!     save would have mutated the workbook it failed to write.
//!   * The last poll sits immediately before the rename, so a
//!     cancellation can never commit a file. Everything after it —
//!     rename, swap, dir fsync — is non-cancellable, and the swap is
//!     `Candidate.swap`, which cannot fail. `PartStore.CommitHook`
//!     enforces that by taking a `void`-returning function.
//!   * The directory fsync comes *after* the swap because it is allowed
//!     to fail and the swap is not. A failure there is a
//!     `durability_warning` on a successful save — the rename already
//!     committed, so reporting an error would contradict the state on
//!     disk.
//!
//! `recalculate()` is the same pipeline with the file half removed: it
//! swaps as the final pipeline operation and never opens anything.
//!
//! The three gates
//! ---------------
//!   * **Logical-view gate.** §5.7.1 builds the model over the logical
//!     view — stored cells plus staged deltas plus appends. Deltas are
//!     modeled; appends and fresh-emit bodies are not cells the source
//!     part has, and §5.8b's approved mutation set cannot insert one.
//!     So a workbook carrying either refuses *before* the model is
//!     built rather than recalculating a view that silently omits them.
//!   * **Pre-M7 gate.** `resolved.patch`'s: a non-1×1 result, a
//!     dynamic-array anchor or a legacy CSE array is
//!     `FormulaSpillPersistUnsupported`. It runs over the whole
//!     projection before a byte is written, so a refusal has produced
//!     nothing to roll back — and it runs before `prepare`, so the
//!     workbook has no candidate either.
//!   * **Embedding staleness preflight** (§5.7.1 step 3b). After
//!     staging, before the candidate: a staged cell whose bytes changed
//!     and which lands inside any coverage refuses with
//!     `FormulaStaleEmbeddings`. v1 has nowhere to record "this vector
//!     is stale" — `zlsxER1` has no status field and `hashes.bin` no
//!     sentinel — so the alternative to refusing is committing changed
//!     values under hashes that no longer describe them.
//!
//! The no-formula rule
//! -------------------
//! A workbook with no formula cells at all is left byte-identical, calc
//! state included. §5.7.6's truthful producer state says what zlsx's
//! caches are worth, and a run that computed nothing produced none;
//! writing `fullCalcOnLoad="1"` there would ask a consumer to
//! recalculate a workbook with nothing to calculate, and would make a
//! recalc-on-save of a formula-free file differ from a plain save. The
//! identity is a test, not a remark.

const std = @import("std");
const assert = std.debug.assert;
const Allocator = std.mem.Allocator;

const control = @import("zlsx_control");
const coords = @import("zlsx_refs");
const engine = @import("zlsx_formula");
const zlsx = @import("zlsx");

const embedding_part = @import("embedding_part.zig");
const recalc_txn = @import("recalc_txn.zig");
const store_mod = @import("store.zig");
const workbook_mod = @import("workbook.zig");

const Workbook = workbook_mod.Workbook;
const Worksheet = workbook_mod.Worksheet;
const WorkbookEnv = workbook_mod.WorkbookEnv;
const Error = workbook_mod.Error;

pub const Control = control.Control;
pub const Report = recalc_txn.Report;
pub const RunInputs = engine.run_inputs.RunInputs;
/// §5.8c's authoring dialect (M7c, Zig-only until M9a2). Lives at the
/// projection seam — the same layer setCell publications ride — and is
/// re-exported here so a caller can spell it without importing the
/// engine tree.
pub const FormulaWrite = engine.resolved.FormulaWrite;

/// §5.4b's comparator with the shipped fold wired in.
///
/// `value.Collation` takes its fold as a parameter so the *semantics*
/// stay independent of the build graph, but a recalc entry point has to
/// supply one, and there is exactly one right answer: the fold in the
/// `zlsx` module's `unicode/` tree, which this layer already imports.
/// A caller may still override it; the default is not a placeholder.
pub const collation_v1: engine.value.Collation = .{ .fold = &foldString };

fn foldString(allocator: Allocator, s: []const u8) anyerror![]u8 {
    return zlsx.casefold.foldString(allocator, s);
}

/// M9a1: the rule versions `zlsx_engine_fingerprint()` names, read from
/// the engine so the exported identity cannot drift from the code that
/// implements each rule. The C ABI cannot import the engine directly —
/// this seam is the one place it already reaches recalc semantics.
pub const rule_versions = .{
    .excel_fp = engine.value.excel_fp_rules_v1.name,
    .rng = engine.rng.version,
    .collation = collation_v1.version,
};

/// M9a1: §5.3a's publish seam, re-exported for the C ABI. Blank never
/// crosses a public boundary; this is the one mandatory conversion.
pub const publish = engine.value.publish;
pub const PublishedScalar = engine.value.PublishedScalar;

/// Everything about a recalc that is not a `RunInputs`.
///
/// The split follows §5.5: `RunInputs` is what makes a run reproducible
/// — clock, seed, fidelity, resource limits, and the cancel/deadline
/// pair that is deliberately outside the fingerprint. This struct is the
/// policy around it.
pub const Options = struct {
    /// §5.4b, injected. See `collation_v1`.
    collation: engine.value.Collation = collation_v1,
    /// §9's bound on *modeling the workbook*.
    limits: engine.decode.Limits = .{},
    /// §9's bound on the *shape of a formula*.
    parse_limits: engine.parser.Limits = .{},
    /// §9's bound on *work done*, including §5.6c's two iteration
    /// ceilings.
    work_limits: engine.iterate.WorkLimits = .{},
    /// §5.7.7's policy for constructs this engine does not implement.
    on_unsupported: recalc_txn.OnUnsupported = .refuse,
    /// §5.7.4's counted retention.
    max_retained_generations: usize = recalc_txn.default_max_retained_generations,
    max_retained_bytes: u64 = recalc_txn.default_max_retained_bytes,
    /// M9a2's refusal seam (decision M9a1-4): when non-null, a Plane-2
    /// refusal is MOVED here — census included — before the error
    /// returns, and the caller owns it (`Refusal.deinit` with the
    /// workbook's allocator, which is what allocated it). Untouched on
    /// success, cancellation and non-refusal errors. Null keeps the
    /// pre-M9a2 behaviour: the census dies with the refusal and only
    /// the error name crosses.
    refusal_out: ?*recalc_txn.Refusal = null,
};

// ─── the entry points (§12.1) ────────────────────────────────────

/// §5.7's in-memory transaction: prepare, then swap as the final
/// pipeline operation. No file is involved and none is opened.
///
/// On any refusal or cancellation the workbook is exactly as it was —
/// the candidate is abandoned before this returns, and every gate above
/// it runs before a candidate exists at all.
pub fn recalculate(
    wb: *Workbook,
    gpa: Allocator,
    io: std.Io,
    run: RunInputs,
    opts: Options,
) Error!Report {
    var prepared = try prepare(wb, gpa, io, run, opts);
    switch (prepared) {
        .none => |r| return r,
        .refused => |r| return takeRefusal(wb, opts, r),
        .ok => |*candidate| {
            candidate.swap(wb);
            return candidate.takeReport();
        },
    }
}

/// The `.refused` arm's single exit: name the error, then either move
/// the refusal — census and all — into the caller's `refusal_out` slot
/// or free it. Returning the error *value* keeps both callers a
/// one-liner.
fn takeRefusal(wb: *Workbook, opts: Options, refusal: recalc_txn.Refusal) workbook_mod.Error {
    var r = refusal;
    const e = r.toWorkbookError();
    if (opts.refusal_out) |slot| slot.* = r else r.deinit(wb.allocator);
    return e;
}

/// §5.7.9's file transaction, in the order §5.7.9 makes normative.
///
/// The bytes written are serialized from the prepared candidate, not
/// from the workbook: the swap has not happened yet and must not, so
/// that any failure before the rename leaves both the destination's
/// prior bytes and the workbook's memory untouched. The swap runs inside
/// the commit region, between the rename and the directory fsync.
pub fn saveWithRecalc(
    wb: *Workbook,
    gpa: Allocator,
    io: std.Io,
    path: []const u8,
    run: RunInputs,
    opts: Options,
) Error!Report {
    var prepared = try prepare(wb, gpa, io, run, opts);
    switch (prepared) {
        .refused => |r| return takeRefusal(wb, opts, r),
        // Nothing to recalculate, so nothing to prepare — and a save
        // that writes the staged state is precisely what "byte-identical
        // to a plain save" means.
        .none => |r| {
            var watch: control.Watch = .init(io, controlOf(run));
            const commit = try wb.store.saveControlled(io, path, watch.poller());
            var out = r;
            if (commit.durability_warning) out.durability.warn(commit.durability_errno);
            return out;
        },
        .ok => |*candidate| {
            var swapper: Swapper = .{ .wb = wb, .candidate = candidate };
            var watch: control.Watch = .init(io, controlOf(run));

            const commit = candidate.next.saveCommitted(
                io,
                path,
                watch.poller(),
                swapper.hook(),
            ) catch |err| {
                // Pre-commit, by construction: `saveCommitted` cannot
                // return an error after `finish`, and `finish` is the
                // rename. So the candidate is still un-swapped and the
                // destination still holds whatever it held before.
                assert(!swapper.fired);
                candidate.abandon();
                return err;
            };
            assert(swapper.fired);

            var out = candidate.takeReport();
            // §5.7.9's dormant slot, and the one mutation permitted
            // after the swap: two scalar stores into memory the report
            // already owned before the transaction began.
            if (commit.durability_warning) out.durability.warn(commit.durability_errno);
            return out;
        },
    }
}

/// The swap, as something `PartStore` can call without knowing what a
/// candidate is.
///
/// `fired` is not bookkeeping for its own sake: it is how the two
/// assertions above state §5.7.9's rule — an error implies the hook did
/// not run, and a success implies it did — in a form a Debug build
/// enforces rather than a comment a reader has to trust.
const Swapper = struct {
    wb: *Workbook,
    candidate: *recalc_txn.Candidate,
    fired: bool = false,

    fn hook(self: *Swapper) store_mod.CommitHook {
        return .{ .ctx = self, .call = call };
    }

    fn call(ctx: ?*anyopaque) void {
        const self: *Swapper = @ptrCast(@alignCast(ctx.?));
        self.candidate.swap(self.wb);
        self.fired = true;
    }
};

/// §5.10's control, as `RunInputs` carries it. Both fields are outside
/// `EffectiveRunInputs` by construction, so threading them here cannot
/// change what a run fingerprints as.
fn controlOf(run: RunInputs) Control {
    return .{ .cancel = run.cancel, .deadline = run.deadline };
}

// ─── prepare ─────────────────────────────────────────────────────

/// What `prepare` produced. Three arms rather than `recalc_txn.Result`'s
/// two: a workbook with no formulas has no candidate *and* no refusal,
/// and collapsing that into either would make a no-op either a mutation
/// or an error.
pub const Prepared = union(enum) {
    ok: recalc_txn.Candidate,
    refused: recalc_txn.Refusal,
    /// Nothing to recalculate. The report is complete and empty.
    none: Report,
};

/// Run §5.7 steps 1–4 and hand back the candidate, un-swapped.
///
/// Public because M5d3's `writerSaveWithRecalc` composes the same
/// candidate across a different serialization, and because a caller that
/// wants to inspect a recalc before committing it has nowhere else to
/// stand.
///
/// **The `.ok` arm is an obligation**: exactly one of `swap` or
/// `abandon` must be called on the candidate. Neither leaks a whole
/// generation; both is a safety-checked assertion.
pub fn prepare(
    wb: *Workbook,
    gpa: Allocator,
    io: std.Io,
    run: RunInputs,
    opts: Options,
) Error!Prepared {
    run.validate() catch |e| return switch (e) {
        error.LimitOutOfRange => Error.FormulaLimitExceeded,
        error.UtcOffsetOutOfRange => Error.FormulaMalformedInput,
    };

    var watch: control.Watch = .init(io, controlOf(run));
    watch.poller().check() catch return Error.Cancelled;

    // Gate one, before the model: the logical view has to be one this
    // pipeline can both read and write back.
    try logicalViewGate(wb);

    var model = switch (try WorkbookEnv.build(gpa, wb, .{
        .collation = opts.collation,
        .fidelity = run.fidelity,
        .limits = opts.limits,
    })) {
        .ok => |m| m,
        .refused => |r| return censusRefusal(wb, run, opts, r.planeTwo(), null),
    };
    defer model.deinit();

    var arena = std.heap.ArenaAllocator.init(gpa);
    defer arena.deinit();
    const a = arena.allocator();

    var bridge: workbook_mod.GraphBridge = .{ .model = &model, .gpa = gpa };
    const input = try bridge.buildInput(a);

    // §5.7's no-formula rule. Decided on the model rather than on the
    // output bytes: "this workbook has nothing to calculate" is a
    // property of the file, and a run that merely happened to change no
    // byte is a different (and much later) statement.
    if (input.cells.len == 0) return .{ .none = .{ .resolved = run.effective(.recalc) } };

    var g = switch (try engine.graph.build(gpa, input, bridge.resolver(), .{
        .parse_limits = opts.parse_limits,
        .limits = opts.work_limits,
    })) {
        .ok => |x| x,
        .refused => |r| return censusRefusal(wb, run, opts, r.planeTwo(), null),
    };
    defer g.deinit();

    // Every formula cell is a root. A recalc is the whole workbook by
    // definition (§5.7.1), and stating it as roots rather than as "run
    // the whole graph" is what keeps §5.6e's rebuild honest: a dynamic
    // reference discovered on pass two widens the closure these roots
    // derive, and a frozen component list could not.
    const roots = try a.alloc(engine.graph.Key, input.cells.len);
    for (input.cells, roots) |c, *k| k.* = .{ .cell = c.cell };

    var rng: engine.rng.Rng = .init(run.rng_seed);
    var draws = rng.drawSource();
    var schedule: engine.draws.Schedule = .{};
    defer schedule.deinit(gpa);
    draws.schedule = &schedule;
    draws.gpa = gpa;

    const eval_opts: workbook_mod.EvaluateOptions = .{
        .collation = opts.collation,
        .fidelity = run.fidelity,
        .limits = opts.limits,
        .parse_limits = opts.parse_limits,
        .work_limits = opts.work_limits,
        .now_utc_ms = run.now_utc_ms,
        .utc_offset_min = run.utc_offset_min,
        .platform_profile = run.platform_profile,
        // Workbook-derived, never caller-set (§5.4a, §5.4d): the same
        // text is a different serial under each epoch and a different
        // character count under each compatibility version.
        .date_system = model.calc.date_system,
        .text_compat = switch (model.calc.text_compat) {
            .v1 => .cv1,
            .v2 => .cv2,
        },
        .draws = &draws,
    };

    var driver: Driver = .{
        .wb = wb,
        .model = &model,
        .arena = a,
        .gpa = gpa,
        .opts = eval_opts,
        .watch = &watch,
    };
    defer driver.deinit();

    var counters: engine.graph.WorkCounters = .{ .limits = opts.work_limits };
    switch (try engine.graph.plan(g, a, roots, &counters, .{
        .iterating = model.calc.iterate,
    })) {
        .ok => {},
        .refused => |r| return censusRefusal(wb, run, opts, r.planeTwo(), null),
    }

    const outcome = try engine.iterate.run(gpa, g, driver.host(), .{
        .limits = opts.work_limits,
        .settings = .{
            .iterate = model.calc.iterate,
            .iterate_count = model.calc.iterate_count,
            .iterate_delta = model.calc.iterate_delta,
        },
        .schedule = &schedule,
        .counters = &counters,
        .closure = .{ .roots = roots, .iterating = model.calc.iterate },
        .rebuild = .{ .input = input, .resolver = bridge.resolver() },
    });
    if (driver.failure) |e| return e;
    var iter_report = switch (outcome) {
        .ok => |r| r,
        .refused => |r| return censusRefusal(
            wb,
            run,
            opts,
            driver.refused_plane orelse r.planeTwo(),
            driver.refused_at,
        ),
    };
    defer iter_report.deinit(gpa);

    watch.poller().check() catch return Error.Cancelled;

    // §5.7.3 step 3 and §5.7.1 step 3b, in that order: stage the bytes,
    // then ask whether staging them invalidated a coverage.
    var staged = try stage(wb, gpa, a, &model, &driver, run, opts);
    switch (staged) {
        .refused => |plane| return censusRefusal(wb, run, opts, plane, null),
        .ok => |*s| {
            if (try staleCoverage(wb, s.cells)) {
                return censusRefusal(wb, run, opts, .FormulaStaleEmbeddings, null);
            }

            watch.poller().check() catch return Error.Cancelled;

            var result = try recalc_txn.prepare(wb, s.parts, &.{}, .{
                .max_retained_generations = opts.max_retained_generations,
                .max_retained_bytes = opts.max_retained_bytes,
                .on_unsupported = opts.on_unsupported,
                .cancel = run.cancel,
                .resolved = run.effective(.recalc),
            });
            switch (result) {
                .refused => |r| return .{ .refused = r },
                .ok => |*candidate| {
                    candidate.report.cells_written = s.cells_written;
                    candidate.report.passes = totalPasses(iter_report);
                    candidate.report.non_converged_cells = iter_report.non_converged_cells;
                    candidate.report.dynamic_passes = iter_report.dynamic_passes;
                    return .{ .ok = candidate.* };
                },
            }
        },
    }
}

fn totalPasses(r: engine.iterate.Report) u32 {
    var n: u32 = 0;
    for (r.components) |c| n +|= c.passes;
    return n;
}

/// A plane-2 refusal, routed through §5.7.7's census so the caller's
/// `on_unsupported` policy is the thing that decides.
///
/// Handing it to `recalc_txn.prepare` rather than returning it directly
/// is what makes `.keep_stale_and_mark` mean something: a refusal from
/// an eligible plane becomes a mark-only run there, and one from any
/// other plane refuses whatever the caller asked for. Deciding here
/// would put the eligibility table in two places.
fn censusRefusal(
    wb: *Workbook,
    run: RunInputs,
    opts: Options,
    plane: engine.decode.PlaneTwo,
    at: ?engine.env.CellRef,
) Error!Prepared {
    const entry: recalc_txn.Unsupported = .{
        .plane = plane,
        .sheet = if (at) |c| c.sheet.toInt() else 0,
        .row = if (at) |c| c.row.oneBased() else 0,
        .col = if (at) |c| c.col.zeroBased() else 0,
    };
    const result = try recalc_txn.prepare(wb, &.{}, &.{entry}, .{
        .max_retained_generations = opts.max_retained_generations,
        .max_retained_bytes = opts.max_retained_bytes,
        .on_unsupported = opts.on_unsupported,
        .cancel = run.cancel,
        .resolved = run.effective(.recalc),
    });
    return switch (result) {
        .ok => |c| .{ .ok = c },
        .refused => |r| .{ .refused = r },
    };
}

// ─── gate one: the logical view (§5.7.1) ─────────────────────────

/// A recalc reads the logical view and writes back through §5.8b's
/// approved mutation set, and the two do not cover the same ground.
///
/// Staged *deltas* are modeled (`WorkbookEnv.build` inserts them at the
/// staged layer) and land on `<c>` elements the part already has, so
/// they patch. Appended rows and fresh-emit bodies are cells the part
/// does not have; `resolved.patch` refuses to insert one, and a model
/// that quietly omitted them would compute `SUM(A:A)` over a column the
/// caller has already added to. Refusing here — before the model, before
/// any candidate — is the only answer that is both truthful and
/// zero-mutation.
fn logicalViewGate(wb: *Workbook) Error!void {
    for (wb.worksheets) |*ws| {
        if (ws.appended_rows.items.len > 0) return Error.SheetHasUnsavedAppends;
        if (ws.body.items.len > 0) return Error.SheetHasUnsavedAppends;
    }
}

// ─── the evaluation driver (§5.6c/§5.6e host) ────────────────────

/// One published value, kept alongside the model's computed layer.
///
/// The layer holds values because that is what a read needs; this holds
/// the *shape* and the *dialect* as well, because §5.7.3's pre-M7 gate
/// is a statement about those two and both are gone by the time a
/// publication reaches the patcher.
const Published = struct {
    cell: engine.env.CellRef,
    value: engine.value.ScalarValue,
    shape: engine.value.Shape,
    /// False between `evaluate` (which is where the shape is knowable)
    /// and `publish` (which is where the value is). A cell the engine
    /// evaluated but chose not to publish must not be staged: the entry
    /// exists to hold its shape, and treating its placeholder as a
    /// result would cache a zero nothing computed.
    has_value: bool = false,
    /// The whole array the cell last published, when it published one —
    /// arena-backed, so it outlives the run. `value` is the narrowed
    /// anchor scalar; §5.6h's slave synthesis (M7b1) is the one reader
    /// of the rest, because a legacy CSE's tails are not evaluable
    /// nodes and their file caches are maintained from exactly this.
    matrix: ?engine.value.Matrix = null,
};

/// The package's `iterate.Host` for a whole-workbook recalc.
///
/// Mirrors `Workbook.evaluateClosure`'s driver and differs in exactly one
/// way: it keeps what it published. The engine's own journal is private
/// and would not answer the question anyway — it records that a cell was
/// published, not what shape the value had before `scalarOf` narrowed it.
const Driver = struct {
    wb: *Workbook,
    model: *WorkbookEnv,
    arena: Allocator,
    gpa: Allocator,
    opts: workbook_mod.EvaluateOptions,
    /// §5.5's seam through the evaluation phase. The engine takes no
    /// poller of its own, and one cell is the natural interval: below it
    /// there is nothing to interrupt, and above it a large workbook
    /// would be one unpollable stretch — exactly the shape M5d1 removed
    /// from the archive layer.
    watch: *const control.Watch,

    published: std.ArrayListUnmanaged(Published) = .empty,
    /// Where each cell sits in `published`.
    ///
    /// The list keeps its order — `stage` walks it per sheet and the
    /// projection is handed the result — but a shape note and a publish
    /// each looked their own cell up by scanning it, which made a run
    /// cost O(cells²) and was, after M5d4's other fixes, the largest
    /// single cost left in the pipeline.
    published_at: std.AutoHashMapUnmanaged(engine.env.CellRef, u32) = .empty,
    /// A plane-2 refusal with the detail the engine's own `PlaneTwo`
    /// cannot carry.
    refused_plane: ?engine.decode.PlaneTwo = null,
    refused_at: ?engine.env.CellRef = null,
    /// A package failure that is not a refusal at all.
    failure: ?Error = null,

    fn deinit(self: *Driver) void {
        self.published.deinit(self.gpa);
        self.published_at.deinit(self.gpa);
    }

    /// Append one entry and record where it landed.
    fn track(self: *Driver, p: Published) error{OutOfMemory}!void {
        try self.published.append(self.gpa, p);
        errdefer _ = self.published.pop();
        try self.published_at.put(self.gpa, p.cell, @intCast(self.published.items.len - 1));
    }

    fn host(self: *Driver) engine.iterate.Host {
        return .{ .ctx = self, .vtable = &host_vtable };
    }

    const host_vtable: engine.iterate.Host.VTable = .{
        .evaluate = vtEvaluate,
        .publish = vtPublish,
        .retract = vtRetract,
    };

    fn of(ctx: *anyopaque) *Driver {
        return @ptrCast(@alignCast(ctx));
    }

    fn vtEvaluate(
        ctx: *anyopaque,
        cell: engine.env.CellRef,
        key: engine.draws.Key,
    ) error{OutOfMemory}!engine.iterate.Produced {
        const self = of(ctx);

        // Before the work, not after: a cell already cancelled when its
        // turn comes should not be evaluated at all. `refused` is how
        // the engine is told to stop and roll back — it has no
        // cancellation arm — and `failure` is what turns that back into
        // `error.Cancelled` for the caller, checked before the outcome.
        self.watch.poller().check() catch {
            self.failure = Error.Cancelled;
            return .{ .refused = .FormulaLimitExceeded };
        };

        const text = (self.model.formulaAt(cell) catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            else => {
                self.failure = Error.SheetNotFound;
                return .{ .ok = .{ .value = .{ .scalar = .blank } } };
            },
        }) orelse return .{ .ok = .{ .value = .{ .scalar = .blank } } };

        const one = self.wb.evaluateOne(self.arena, self.model, cell, text, self.opts, key) catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            else => {
                self.failure = e;
                return .{ .ok = .{ .value = .{ .scalar = .blank } } };
            },
        };
        return switch (one) {
            .refused => |r| blk: {
                const plane = Workbook.planeOfRefusal(r);
                self.refused_plane = plane;
                self.refused_at = cell;
                break :blk .{ .refused = plane };
            },
            .ok => |x| blk: {
                // Recorded here rather than in `publish` because a cell
                // can be evaluated and never published (a refusal
                // between the two), and the gate still wants its shape.
                self.noteShape(cell, x.shape) catch return error.OutOfMemory;
                break :blk .{ .ok = .{ .value = x.value, .reads = x.reads } };
            },
        };
    }

    fn noteShape(self: *Driver, cell: engine.env.CellRef, shape: engine.value.Shape) !void {
        if (self.published_at.get(cell)) |i| {
            self.published.items[i].shape = shape;
            return;
        }
        try self.track(.{
            .cell = cell,
            .value = .blank,
            .shape = shape,
            .has_value = false,
        });
    }

    fn vtPublish(
        ctx: *anyopaque,
        cell: engine.env.CellRef,
        v: engine.iterate.Snapshot,
    ) error{OutOfMemory}!void {
        const self = of(ctx);
        // §5.8a (M7a): the model half decides placement — a spilled
        // array's anchor reads its top-left with owned tails placed, a
        // blocked one reads `#SPILL!` with the class recorded. The
        // shape in `published` is still `evaluate`'s, which is what the
        // pre-M7 persistence gate reads.
        const scalar = self.model.publishResult(cell, v) catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            else => blk: {
                self.failure = Error.SheetNotFound;
                break :blk engine.value.ScalarValue.blank;
            },
        };
        const matrix: ?engine.value.Matrix = switch (v) {
            .array => |m| m,
            .scalar => null,
        };
        if (self.published_at.get(cell)) |i| {
            const p = &self.published.items[i];
            p.value = scalar;
            p.has_value = true;
            p.matrix = matrix;
            // `shape` is deliberately not touched: the entry `evaluate`
            // made carries the result's shape, and that — not the
            // scalar the anchor's coordinate reads — is what the
            // pre-M7 gate refuses on.
            return;
        }
        try self.track(.{
            .cell = cell,
            .value = scalar,
            .shape = v.shape(),
            .has_value = true,
            .matrix = matrix,
        });
    }

    /// One publish, undone. A refused run replays every one of these in
    /// reverse (§5.6c's "zero mutation"), so the record has to shrink
    /// with the layer — a staged set that outlived a rollback would
    /// write values the run decided not to keep.
    fn vtRetract(ctx: *anyopaque, cell: engine.env.CellRef) void {
        const self = of(ctx);
        self.model.retractResult(cell);
        const found = self.published_at.fetchRemove(cell) orelse return;
        const at = found.value;
        _ = self.published.orderedRemove(at);
        // A rollback replays the journal backwards, so the entry being
        // dropped is normally the last one and this repairs nothing. It
        // is a loop because "normally" is not "always": a cell whose
        // shape was noted but which never published leaves an entry
        // behind the one being removed.
        for (self.published.items[at..]) |p| {
            if (self.published_at.getPtr(p.cell)) |slot| slot.* -= 1;
        }
    }
};

// ─── §5.7.3 step 3: values become bytes ──────────────────────────

/// One staged cell, for the embedding preflight's benefit.
///
/// It carries the *part* rather than the sheet index because that is
/// what a coverage names, and only the cells whose bytes actually moved:
/// a publication that matched the cached value produced no edit, and an
/// unchanged byte cannot invalidate a hash.
const StagedCell = struct {
    part: []const u8,
    row: u32,
    col: u32,
};

const Staged = struct {
    parts: []const recalc_txn.StagedPart,
    cells: []const StagedCell,
    cells_written: u32,
};

const StageResult = union(enum) {
    ok: Staged,
    /// The plane alone, so the caller routes it through §5.7.7's census
    /// like every other plane-2 refusal rather than deciding here.
    refused: engine.decode.PlaneTwo,
};

/// The declared range of a legacy CSE anchor, from the scan: a
/// `<f t="array" ref="…">` whose ref parses and starts at the anchor.
/// Anything else answers null and the patcher's own gate names the
/// refusal pre-mutation.
fn cseDeclaredRange(cells: []const engine.decode.SheetCell, cell: engine.env.CellRef) ?coords.Range {
    const sc = sheetCellAt(cells, cell.row, cell.col) orelse return null;
    const f = sc.formula orelse return null;
    const kind = f.kind orelse return null;
    if (!std.mem.eql(u8, kind, "array")) return null;
    const raw = f.ref orelse return null;
    const range = (coords.parseRange(raw, .{
        .dollar = .accept,
        .case = .insensitive,
    }) catch return null).normalized();
    if (range.first.row.oneBased() != cell.row.oneBased() or
        range.first.col.zeroBased() != cell.col.zeroBased()) return null;
    return range;
}

fn sheetCellAt(cells: []const engine.decode.SheetCell, row: coords.Row, col: coords.Col) ?engine.decode.SheetCell {
    var lo: usize = 0;
    var hi: usize = cells.len;
    while (lo < hi) {
        const mid = lo + (hi - lo) / 2;
        const c = cells[mid];
        if (c.row.oneBased() < row.oneBased() or
            (c.row == row and c.col.zeroBased() < col.zeroBased()))
        {
            lo = mid + 1;
        } else if (c.row == row and c.col == col) {
            return c;
        } else {
            hi = mid;
        }
    }
    return null;
}

/// §5.6h at the staging boundary (M7b1): let D be the declared range
/// and R the anchor's published array. Only a 1×1 R broadcasts — it
/// fills every cell of D (an error R propagates the same way); a
/// non-scalar R places by coordinate, `#N/A` wherever D extends beyond
/// R in either dimension, surplus of R truncated. Applied to the FILE's
/// slave caches, which §5.7.3 says "keep `<f>` any-shape, gain `<v>`".
fn synthesizeCseSlaves(
    a: Allocator,
    pubs: *std.ArrayListUnmanaged(engine.resolved.Publication),
    driver: *const Driver,
    p: Published,
    decl: coords.Range,
    run: RunInputs,
    sheet_idx: u32,
) error{OutOfMemory}!void {
    const rows_d = decl.rowCount();
    const cols_d = decl.colCount();
    var i: u32 = 0;
    while (i < rows_d) : (i += 1) {
        var j: u32 = 0;
        while (j < cols_d) : (j += 1) {
            if (i == 0 and j == 0) continue; // the anchor's own cell
            const row = coords.Row.fromOneBased(decl.first.row.oneBased() + i) catch continue;
            const col = coords.Col.fromZeroBased(decl.first.col.zeroBased() + j) catch continue;
            // A covered cell with a formula of its own publishes for
            // itself — the graph never made it a tail (M7a).
            const target: engine.env.CellRef = .{
                .sheet = engine.env.SheetIndex.fromInt(sheet_idx),
                .row = row,
                .col = col,
            };
            if (driver.published_at.get(target) != null) continue;
            const elem: engine.value.ScalarValue = if (p.matrix) |m|
                (if (m.rows == 1 and m.cols == 1)
                    // A 1×1 R broadcasts — the same rule as a scalar,
                    // spelled per §5.6h so a one-element array does not
                    // fill the rest of D with `#N/A`.
                    m.at(0, 0)
                else if (i < m.rows and j < m.cols)
                    m.at(i, j)
                else
                    engine.value.ScalarValue.errorOf(.na))
            else
                p.value;
            try pubs.append(a, .{
                .row = row,
                .col = col,
                .result = engine.value.publish(elem, run.fidelity),
                .origin = .computed,
                .shape = .{ .rows = 1, .cols = 1 },
                .dialect = .legacy,
            });
        }
    }
}

fn stage(
    wb: *Workbook,
    gpa: Allocator,
    a: Allocator,
    model: *WorkbookEnv,
    driver: *Driver,
    run: RunInputs,
    opts: Options,
) Error!StageResult {
    var parts: std.ArrayListUnmanaged(recalc_txn.StagedPart) = .empty;
    var touched: std.ArrayListUnmanaged(StagedCell) = .empty;
    var written: u32 = 0;

    var sheet_idx: u32 = 0;
    while (sheet_idx < wb.sheetCount()) : (sheet_idx += 1) {
        var any = false;
        for (driver.published.items) |p| {
            if (p.has_value and p.cell.sheet.toInt() == sheet_idx) {
                any = true;
                break;
            }
        }
        if (!any) continue;

        const ws = try wb.sheet(sheet_idx);
        const part_name = try ws.resolvePartName();
        const part = (try wb.store.part(part_name)) orelse return Error.MissingSheetPart;

        // The same scan the model was built from, re-run because the
        // model keeps *cells* and the projection needs *spans*. Same
        // bytes, same options — `resolved.project` requires the accepted
        // scan of the source it is patching, and a second scan under
        // different options would be a second opinion about the document.
        // It runs BEFORE the publications are built since M7b1: §5.6h's
        // slave synthesis reads each anchor's declared range from it.
        var scan = switch (try engine.decode.scanSheet(gpa, part.bytes, model.strings.items, .{
            .limits = opts.limits,
            .fidelity = run.fidelity,
            .date_system = model.calc.date_system,
        })) {
            .ok => |s| s,
            .refused => |r| return .{ .refused = r.planeTwo() },
        };
        defer scan.deinit();

        var pubs: std.ArrayListUnmanaged(engine.resolved.Publication) = .empty;
        for (driver.published.items) |p| {
            if (!p.has_value) continue;
            if (p.cell.sheet.toInt() != sheet_idx) continue;
            const dialect = model.evalEnv().dialectOf(p.cell) catch |e| switch (e) {
                error.OutOfMemory => return Error.OutOfMemory,
                // A dialect the metadata layer could not resolve is a
                // pre-mutation refusal, never a guess (§5.3b).
                else => return .{ .refused = .FormulaUnsupportedConstruct },
            };
            try pubs.append(a, .{
                .row = p.cell.row,
                .col = p.cell.col,
                .result = engine.value.publish(p.value, run.fidelity),
                .origin = .computed,
                .shape = p.shape,
                .dialect = dialect,
                // §5.8b: what the model placed travels with the
                // publication — the patcher consumes the outcome, never
                // re-decides it. Tail publications land with the first
                // committed transition reference; until one exists the
                // anchor's own refusal precedes every tail on every
                // reachable path.
                .role = if (model.spillOutcomeOf(p.cell)) |o| .{ .da_anchor = o } else .plain,
            });
            // A legacy CSE's slaves are not evaluable nodes; their file
            // caches are maintained here, from the anchor's own array,
            // by §5.6h's placement rule — never by re-decision.
            if (dialect == .legacy) {
                if (cseDeclaredRange(scan.cells, p.cell)) |decl| {
                    try synthesizeCseSlaves(a, &pubs, driver, p, decl, run, sheet_idx);
                }
            }
        }
        if (pubs.items.len == 0) continue;

        var deltas: engine.resolved.StagedDeltas = .{ .publications = pubs.items };
        var projected = engine.resolved.project(gpa, part.bytes, &scan, &deltas) catch |e| switch (e) {
            error.OutOfMemory => return Error.OutOfMemory,
            // The set is fresh per sheet and consumed once; a second
            // consume is a bookkeeping bug, not a workbook statement.
            error.DeltasAlreadyConsumed => unreachable,
        };
        switch (projected) {
            .refused => |r| return .{ .refused = r.planeTwo() },
            .ok => |*projection| {
                defer projection.deinit();
                var patched = switch (try engine.resolved.patch(projection, gpa)) {
                    .ok => |p| p,
                    // Gate two: the pre-M7 spill gate, plus every shape
                    // a patch cannot be confined over. Nothing has been
                    // written — `patch` runs both passes before it
                    // produces a byte — and no candidate exists yet.
                    .refused => |r| return .{ .refused = r.planeTwo() },
                };
                defer patched.deinit();

                if (patched.edits.len == 0) continue;

                try parts.append(a, .{
                    .name = try a.dupe(u8, part_name),
                    .bytes = try a.dupe(u8, patched.bytes),
                });
                var last: ?engine.resolved.CellSite = null;
                for (patched.edits) |e| {
                    if (last) |l| {
                        if (l.row == e.cell.row and l.col == e.cell.col) continue;
                    }
                    last = e.cell;
                    written += 1;
                    try touched.append(a, .{
                        .part = part_name,
                        .row = e.cell.row,
                        .col = e.cell.col,
                    });
                }
            },
        }
    }

    return .{ .ok = .{
        .parts = try parts.toOwnedSlice(a),
        .cells = try touched.toOwnedSlice(a),
        .cells_written = written,
    } };
}

// ─── gate three: the embedding preflight (§5.7.1 step 3b) ────────

/// Whether any staged cell landed inside a coverage.
///
/// **Any overlap, not any hash mismatch.** A coverage's hash binds to a
/// canonical row payload, and a changed `<v>` inside that row changes
/// the payload by construction — the type letter, the canonical number,
/// or the folded text. Recomputing the hash to confirm what the edit
/// already proves would add a second canonicalizer to keep in sync with
/// the one the embedding arc shipped, and would answer the same question.
///
/// Only cells the patch actually rewrote are considered. A recalc that
/// republished the value a cell already carried produced no edit, so it
/// is not a mutation and cannot make a vector stale — which is what
/// makes recalculating an embedded workbook twice legal.
fn staleCoverage(wb: *Workbook, cells: []const StagedCell) Error!bool {
    if (cells.len == 0) return false;

    // A workbook whose embedding index will not parse is not one this
    // preflight can clear, and it is not this row's to diagnose either:
    // `embeddings()` already refuses it by name.
    //
    // The call populates the workbook's lazy coverage cache, which
    // borrows the index / vec / hashes part bytes. A recalc never
    // rewrites any of those three, and §5.7.4's whole-generation
    // retention keeps the bytes alive regardless, so the cache stays
    // valid across the swap that may follow.
    const state = try wb.embeddings();
    const view = switch (state) {
        .present => |v| v,
        // A stripped index leaves a recovery record and no coverages, so
        // there is nothing a staged cell could overlap.
        .stripped, .absent => return false,
    };

    for (view.coverages) |cv| {
        const cov = cv.coverage;
        var buf: [512]u8 = undefined;
        const want = std.fmt.bufPrint(&buf, "xl/{s}", .{cov.worksheet_target}) catch continue;
        for (cells) |c| {
            if (!std.mem.eql(u8, c.part, want)) continue;
            const r = cov.parsed_range;
            if (c.row < r.first.row or c.row > r.last.row) continue;
            if (c.col < r.first.col or c.col > r.last.col) continue;
            return true;
        }
    }
    return false;
}

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;
const PartStore = @import("store.zig").PartStore;

test {
    // The public surface here is reached from two forwarder bodies in
    // `workbook.zig`, so nothing else analyses it.
    testing.refAllDecls(@This());
}

const ns_main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const ct_sheet = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const ct_workbook = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const ct_rels = "application/vnd.openxmlformats-package.relationships+xml";
const ct_metadata = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml";
const rel_worksheet = ns_r ++ "/worksheet";
const rel_officedoc = ns_r ++ "/officeDocument";

const sheet_part = "xl/worksheets/sheet1.xml";

/// `B1 = A1+1` with a cache that says otherwise. A recalc has to have
/// something to change, or every identity below is trivially true.
const sheet_stale =
    "<worksheet xmlns=\"" ++ ns_main ++ "\"><sheetData><row r=\"1\">" ++
    "<c r=\"A1\"><v>1</v></c><c r=\"B1\"><f>A1+1</f><v>999</v></c>" ++
    "</row></sheetData></worksheet>";

/// The same two cells with no formula anywhere.
const sheet_no_formula =
    "<worksheet xmlns=\"" ++ ns_main ++ "\"><sheetData><row r=\"1\">" ++
    "<c r=\"A1\"><v>1</v></c><c r=\"B1\"><v>999</v></c>" ++
    "</row></sheetData></worksheet>";

const metadata_dynamic_array =
    "<metadata xmlns=\"" ++ ns_main ++ "\" xmlns:xda=\"http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray\">" ++
    "<metadataTypes count=\"1\"><metadataType name=\"XLDAPR\" minSupportedVersion=\"120000\" copy=\"1\" pasteAll=\"1\" pasteValues=\"1\" merge=\"1\" splitFirst=\"1\" rowColShift=\"1\" clearFormats=\"1\" clearComments=\"1\" assign=\"1\" coerce=\"1\" cellMeta=\"1\"/></metadataTypes>" ++
    "<futureMetadata name=\"XLDAPR\" count=\"1\"><bk><extLst><ext uri=\"{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}\">" ++
    "<xda:dynamicArrayProperties fDynamic=\"1\" fCollapsed=\"0\"/></ext></extLst></bk></futureMetadata>" ++
    "<cellMetadata count=\"1\"><bk><rc t=\"1\" v=\"0\"/></bk></cellMetadata></metadata>";

const Fixture = struct {
    sheet: []const u8 = sheet_stale,
    calc_pr: []const u8 = "<calcPr calcId=\"191029\"/>",
    /// `xl/metadata.xml`, for the dynamic-array dialect. Absent means
    /// every cell is legacy, which is what a workbook without the part
    /// actually is.
    metadata: ?[]const u8 = null,
    /// An embedding index over `Sheet1!B1:B1`, for the staleness
    /// preflight. The vectors are one f32 lane of zeros — the preflight
    /// reads coverage geometry, never a vector.
    embeddings: bool = false,
    /// `xl/calcChain.xml` plus its rel and content type, so §5.7.5's
    /// removal has something to remove.
    calc_chain: bool = false,
};

fn writeFixture(gpa: Allocator, io: std.Io, dir: []const u8, name: []const u8, f: Fixture) ![]u8 {
    const path = try std.fs.path.join(gpa, &.{ dir, name });
    errdefer gpa.free(path);

    var store = try PartStore.fresh(gpa, io);
    defer store.deinit();

    try store.addPart("_rels/.rels", ct_rels, "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" ++
        "<Relationship Id=\"rId1\" Type=\"" ++ rel_officedoc ++ "\" Target=\"xl/workbook.xml\"/>" ++
        "</Relationships>");

    const wb_xml = try std.fmt.allocPrint(
        gpa,
        "<workbook xmlns=\"" ++ ns_main ++ "\" xmlns:r=\"" ++ ns_r ++ "\">" ++
            "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>{s}</workbook>",
        .{f.calc_pr},
    );
    defer gpa.free(wb_xml);
    try store.addPart("xl/workbook.xml", ct_workbook, wb_xml);

    var rels: std.ArrayListUnmanaged(u8) = .empty;
    defer rels.deinit(gpa);
    try rels.appendSlice(gpa, "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" ++
        "<Relationship Id=\"rId1\" Type=\"" ++ rel_worksheet ++ "\" Target=\"worksheets/sheet1.xml\"/>");
    if (f.calc_chain) {
        try rels.appendSlice(gpa, "<Relationship Id=\"rId2\" Type=\"" ++
            recalc_txn.calc_chain_rel_type ++ "\" Target=\"calcChain.xml\"/>");
    }
    try rels.appendSlice(gpa, "</Relationships>");
    try store.addPart("xl/_rels/workbook.xml.rels", ct_rels, rels.items);

    try store.addPart(sheet_part, ct_sheet, f.sheet);
    if (f.calc_chain) {
        try store.addPart(
            "xl/calcChain.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml",
            "<calcChain xmlns=\"" ++ ns_main ++ "\"><c r=\"B1\" i=\"1\"/></calcChain>",
        );
    }
    if (f.metadata) |m| try store.addPart("xl/metadata.xml", ct_metadata, m);
    if (f.embeddings) try addEmbeddingParts(gpa, &store);

    try store.save(io, path);
    return path;
}

/// A minimal but real embedding index covering `Sheet1!B1:B1`.
fn addEmbeddingParts(gpa: Allocator, store: *PartStore) !void {
    const index =
        "<embeddings xmlns=\"" ++ embedding_part.INDEX_NAMESPACE ++ "\" version=\"1\" model=\"test-model\" dim=\"2\" dtype=\"f32\" hash_algo=\"xxh3-64\">" ++
        "<coverage id=\"c1\" worksheet_target=\"worksheets/sheet1.xml\" range=\"B1:B1\" column=\"B\" count=\"1\" include_formulas=\"true\" vec_rId=\"rId1\" hash_rId=\"rId2\"/>" ++
        "</embeddings>";
    try store.addPart(embedding_part.INDEX_PART_NAME, "application/xml", index);
    try store.addPart(embedding_part.INDEX_RELS_PART_NAME, ct_rels, "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" ++
        "<Relationship Id=\"rId1\" Type=\"" ++ embedding_part.REL_TYPE_VEC ++ "\" Target=\"c1/vec.bin\"/>" ++
        "<Relationship Id=\"rId2\" Type=\"" ++ embedding_part.REL_TYPE_HASH ++ "\" Target=\"c1/hashes.bin\"/>" ++
        "</Relationships>");

    // One f32 lane pair of zeros. The preflight reads coverage geometry
    // and never a vector, but the parser validates both headers against
    // the declared count, so they have to be real.
    var vec: [embedding_part.VEC_HEADER_BYTES + 8]u8 = @splat(0);
    _ = embedding_part.writeVecHeader(&vec, .{
        .version = embedding_part.WIRE_VERSION,
        .count = 1,
        .dim = 2,
        .dtype = .f32,
    });
    try store.addPart("xl/zlsxEmbeddings/c1/vec.bin", "application/octet-stream", &vec);

    var hashes: [embedding_part.HASH_HEADER_BYTES + 8]u8 = @splat(0);
    _ = embedding_part.writeHashHeader(&hashes, .{
        .version = embedding_part.WIRE_VERSION,
        .count = 1,
    });
    try store.addPart("xl/zlsxEmbeddings/c1/hashes.bin", "application/octet-stream", &hashes);
    _ = gpa;
}

/// The sentinel is part of the type: `realPathFileAlloc` allocates
/// `len + 1` and a `[]u8` return would free one byte short of what it
/// asked for. The allocator notices.
fn tmpPath(gpa: Allocator, io: std.Io, tmp: *testing.TmpDir) ![:0]u8 {
    return tmp.dir.realPathFileAlloc(io, ".", gpa);
}

fn readAll(gpa: Allocator, io: std.Io, path: []const u8) ![]u8 {
    var f = try std.Io.Dir.cwd().openFile(io, path, .{});
    defer f.close(io);
    const size = (try f.stat(io)).size;
    const buf = try gpa.alloc(u8, @intCast(size));
    errdefer gpa.free(buf);
    var r = f.reader(io, &.{});
    try r.interface.readSliceAll(buf);
    return buf;
}

/// The run every test below uses unless it needs a different one. The
/// clock and the seed are fixed because §5.5 says a run is reproducible
/// from its inputs, and a test that read either from the environment
/// would be asserting something else.
const fixed_run: RunInputs = .{
    .now_utc_ms = 1_700_000_000_000,
    .rng_seed = 0x5EED_1A2,
    .limits = .{},
};

fn cellCache(ws: *Worksheet, ref: []const u8) ![]const u8 {
    const view = try ws.ensureParsed();
    for (view.rows) |row| {
        for (row.cells) |c| {
            if (std.mem.eql(u8, c.ref, ref)) return c.raw_value orelse "";
        }
    }
    return error.CellNotFound;
}

// ─── the pipeline actually runs ──────────────────────────────────

test "recalculate: a stale cache becomes the value the formula says" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmpPath(a, io, &tmp);
    defer a.free(dir);
    const path = try writeFixture(a, io, dir, "in.xlsx", .{});
    defer a.free(path);

    var wb = try Workbook.open(a, io, path);
    defer wb.deinit();

    var report = try wb.recalculate(a, io, fixed_run, .{});
    defer report.deinit(a);

    try testing.expectEqual(@as(u32, 1), report.sheets_patched);
    try testing.expectEqual(@as(u32, 1), report.cells_written);
    try testing.expect(!report.kept_stale);
    // The swap happened: one generation retained, and the workbook's own
    // typed view now reads through the new bytes.
    try testing.expectEqual(@as(usize, 1), wb.retained.items.len);
    try testing.expectEqualStrings("2", try cellCache(try wb.sheet(0), "B1"));

    // §5.7.6's truthful producer state, at the byte level.
    const part = (try wb.store.part("xl/workbook.xml")).?;
    try testing.expect(std.mem.indexOf(u8, part.bytes, "calcId=\"0\"") != null);
    try testing.expect(std.mem.indexOf(u8, part.bytes, "fullCalcOnLoad=\"1\"") != null);
}

test "saveWithRecalc: the file holds what the recalc computed" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmpPath(a, io, &tmp);
    defer a.free(dir);
    const path = try writeFixture(a, io, dir, "in.xlsx", .{});
    defer a.free(path);
    const out = try std.fs.path.join(a, &.{ dir, "out.xlsx" });
    defer a.free(out);

    var wb = try Workbook.open(a, io, path);
    defer wb.deinit();

    var report = try wb.saveWithRecalc(a, io, out, fixed_run, .{});
    defer report.deinit(a);
    try testing.expectEqual(@as(u32, 1), report.cells_written);
    try testing.expect(!report.durability.warning);
    try testing.expectEqual(@as(usize, 1), wb.retained.items.len);

    var reopened = try Workbook.open(a, io, out);
    defer reopened.deinit();
    try testing.expectEqualStrings("2", try cellCache(try reopened.sheet(0), "B1"));
}

test "the pipeline removes the calc chain, part, rel and content type" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmpPath(a, io, &tmp);
    defer a.free(dir);
    const src = try writeFixture(a, io, dir, "in.xlsx", .{ .calc_chain = true });
    defer a.free(src);
    const out = try std.fs.path.join(a, &.{ dir, "out.xlsx" });
    defer a.free(out);

    var wb = try Workbook.open(a, io, src);
    defer wb.deinit();
    var report = try wb.saveWithRecalc(a, io, out, fixed_run, .{});
    defer report.deinit(a);
    try testing.expect(report.calc_chain_removed);

    // §5.7.5 is about the *file*, so the assertion is about the file:
    // reopened from disk, the part is gone and nothing still points at
    // it.
    var reopened = try Workbook.open(a, io, out);
    defer reopened.deinit();
    try testing.expect((try reopened.store.part("xl/calcChain.xml")) == null);
    const wb_rels = (try reopened.store.part("xl/_rels/workbook.xml.rels")).?;
    try testing.expect(std.mem.indexOf(u8, wb_rels.bytes, "calcChain.xml") == null);
    const ct = (try reopened.store.part("[Content_Types].xml")).?;
    try testing.expect(std.mem.indexOf(u8, ct.bytes, "calcChain") == null);
}

// ─── done-when 2: determinism ────────────────────────────────────

/// One recalc through `entry`, saved to `name`, returned as bytes.
fn runOnce(
    a: Allocator,
    io: std.Io,
    dir: []const u8,
    src: []const u8,
    name: []const u8,
    comptime file_transaction: bool,
) ![]u8 {
    const out = try std.fs.path.join(a, &.{ dir, name });
    defer a.free(out);

    var wb = try Workbook.open(a, io, src);
    defer wb.deinit();

    if (file_transaction) {
        var report = try wb.saveWithRecalc(a, io, out, fixed_run, .{});
        report.deinit(a);
    } else {
        var report = try wb.recalculate(a, io, fixed_run, .{});
        report.deinit(a);
        _ = try wb.store.saveControlled(io, out, .none);
    }
    return readAll(a, io, out);
}

test "determinism: equal RunInputs and equal bytes give byte-equal output" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmpPath(a, io, &tmp);
    defer a.free(dir);
    const src = try writeFixture(a, io, dir, "in.xlsx", .{});
    defer a.free(src);

    inline for (.{ true, false }) |file_transaction| {
        const first = try runOnce(a, io, dir, src, "a.xlsx", file_transaction);
        defer a.free(first);
        const second = try runOnce(a, io, dir, src, "b.xlsx", file_transaction);
        defer a.free(second);
        try testing.expectEqualSlices(u8, first, second);
    }

    // And the two entry points agree with each other, which is the
    // statement that `saveWithRecalc` really is `recalculate` plus a
    // file rather than a second pipeline.
    const via_file = try runOnce(a, io, dir, src, "c.xlsx", true);
    defer a.free(via_file);
    const via_memory = try runOnce(a, io, dir, src, "d.xlsx", false);
    defer a.free(via_memory);
    try testing.expectEqualSlices(u8, via_file, via_memory);
}

// ─── done-when 3: scoped idempotence ─────────────────────────────

test "idempotence: recalculating a recalculated workbook changes nothing" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmpPath(a, io, &tmp);
    defer a.free(dir);
    const src = try writeFixture(a, io, dir, "in.xlsx", .{});
    defer a.free(src);

    const once = try std.fs.path.join(a, &.{ dir, "once.xlsx" });
    defer a.free(once);
    const twice = try std.fs.path.join(a, &.{ dir, "twice.xlsx" });
    defer a.free(twice);

    {
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        var r = try wb.saveWithRecalc(a, io, once, fixed_run, .{});
        r.deinit(a);
    }
    {
        var wb = try Workbook.open(a, io, once);
        defer wb.deinit();
        var r = try wb.saveWithRecalc(a, io, twice, fixed_run, .{});
        defer r.deinit(a);
        // Nothing left to write: the caches already say what the
        // formulas say, and §5.7.3's transitions leave a correct cell
        // alone. The calc-state attributes §5.7.3 pins are already at
        // their post-recalc values, so the second pass is a no-op there
        // too — which is what makes the byte comparison below possible.
        try testing.expectEqual(@as(u32, 0), r.cells_written);
        try testing.expectEqual(@as(u32, 0), r.sheets_patched);
    }

    const first = try readAll(a, io, once);
    defer a.free(first);
    const second = try readAll(a, io, twice);
    defer a.free(second);
    try testing.expectEqualSlices(u8, first, second);
}

test "idempotence: the first pass changes exactly the sheet and the calc state" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmpPath(a, io, &tmp);
    defer a.free(dir);
    const src = try writeFixture(a, io, dir, "in.xlsx", .{});
    defer a.free(src);
    const out = try std.fs.path.join(a, &.{ dir, "out.xlsx" });
    defer a.free(out);

    var wb = try Workbook.open(a, io, src);
    defer wb.deinit();
    var r = try wb.saveWithRecalc(a, io, out, fixed_run, .{});
    r.deinit(a);

    var before = try Workbook.open(a, io, src);
    defer before.deinit();
    var after = try Workbook.open(a, io, out);
    defer after.deinit();

    const names = try before.store.partNames();
    for (names) |n| {
        const b = (try before.store.part(n)).?;
        const c = (try after.store.part(n)) orelse return error.PartDisappeared;
        if (std.mem.eql(u8, n, sheet_part)) {
            try testing.expect(!std.mem.eql(u8, b.bytes, c.bytes));
        } else if (std.mem.eql(u8, n, "xl/workbook.xml")) {
            // The one part that may differ, and only in the two
            // attributes §5.7.6 pins.
            try testing.expect(std.mem.indexOf(u8, c.bytes, "calcId=\"0\"") != null);
        } else {
            try testing.expectEqualSlices(u8, b.bytes, c.bytes);
        }
    }
}

// ─── done-when 4: no-formula identity ────────────────────────────

test "no-formula identity: saveWithRecalc equals a plain staged-state save" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmpPath(a, io, &tmp);
    defer a.free(dir);
    const src = try writeFixture(a, io, dir, "in.xlsx", .{ .sheet = sheet_no_formula });
    defer a.free(src);

    const recalced = try std.fs.path.join(a, &.{ dir, "recalced.xlsx" });
    defer a.free(recalced);
    const plain = try std.fs.path.join(a, &.{ dir, "plain.xlsx" });
    defer a.free(plain);

    {
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        var r = try wb.saveWithRecalc(a, io, recalced, fixed_run, .{});
        defer r.deinit(a);
        try testing.expectEqual(@as(u32, 0), r.cells_written);
        try testing.expectEqual(@as(u32, 0), r.sheets_patched);
        // Nothing to compute means nothing to retain: a no-op recalc
        // does not spend a generation.
        try testing.expectEqual(@as(usize, 0), wb.retained.items.len);
    }
    {
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        try wb.store.save(io, plain);
    }

    const x = try readAll(a, io, recalced);
    defer a.free(x);
    const y = try readAll(a, io, plain);
    defer a.free(y);
    try testing.expectEqualSlices(u8, x, y);
}

// ─── done-when 5: §5.7.9's ordering, proven ──────────────────────

test "ordering: an injected rename failure leaves memory AND the destination alone" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const base = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmpPath(a, base, &tmp);
    defer a.free(dir);
    const src = try writeFixture(a, base, dir, "in.xlsx", .{});
    defer a.free(src);
    // A destination that already exists, because "prior bytes intact" is
    // the interesting half of the promise and an absent file cannot show
    // it.
    const dest = try writeFixture(a, base, dir, "dest.xlsx", .{ .sheet = sheet_no_formula });
    defer a.free(dest);

    const before = try readAll(a, base, dest);
    defer a.free(before);

    var wb = try Workbook.open(a, base, src);
    defer wb.deinit();

    const io = control.inject.wrap(base, .{ .fail_rename = true });
    try testing.expectError(error.AccessDenied, wb.saveWithRecalc(a, io, dest, fixed_run, .{}));

    // Memory: no generation retained, and B1 still reads the stale cache.
    try testing.expectEqual(@as(usize, 0), wb.retained.items.len);
    try testing.expectEqualStrings("999", try cellCache(try wb.sheet(0), "B1"));

    // Destination: the prior bytes, and no debris beside them.
    const after = try readAll(a, base, dest);
    defer a.free(after);
    try testing.expectEqualSlices(u8, before, after);
    try expectNoTempFiles(base, dir);
}

test "ordering: an injected sync failure is pre-commit too" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const base = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmpPath(a, base, &tmp);
    defer a.free(dir);
    const src = try writeFixture(a, base, dir, "in.xlsx", .{});
    defer a.free(src);
    const dest = try std.fs.path.join(a, &.{ dir, "absent.xlsx" });
    defer a.free(dest);

    var wb = try Workbook.open(a, base, src);
    defer wb.deinit();

    // The FIRST `File.sync` is the temp file's, which sits above the
    // rename. The destination never existed, so the other half of
    // §5.7.9's promise applies: it is still absent.
    const io = control.inject.wrap(base, .{ .fail_file_sync_at = 1 });
    try testing.expectError(error.InputOutput, wb.saveWithRecalc(a, io, dest, fixed_run, .{}));

    try testing.expectEqual(@as(usize, 0), wb.retained.items.len);
    try testing.expectEqualStrings("999", try cellCache(try wb.sheet(0), "B1"));
    try testing.expectError(error.FileNotFound, std.Io.Dir.cwd().access(base, dest, .{}));
    try expectNoTempFiles(base, dir);
}

test "ordering: a post-rename dir-fsync failure is a warning on a committed save" {
    // No directory fsync exists on Windows: `syncDir` returns clean
    // before the injected second sync can fire — same skip as the two
    // M5d1 siblings (`atomic_file.zig`, `store.zig`).
    if (@import("builtin").os.tag == .windows) return error.SkipZigTest;

    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const base = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmpPath(a, base, &tmp);
    defer a.free(dir);
    const src = try writeFixture(a, base, dir, "in.xlsx", .{});
    defer a.free(src);
    const out = try std.fs.path.join(a, &.{ dir, "out.xlsx" });
    defer a.free(out);

    var wb = try Workbook.open(a, base, src);
    defer wb.deinit();

    // The SECOND `File.sync` is the directory's, which runs after the
    // rename has committed and after the swap.
    const io = control.inject.wrap(base, .{
        .fail_file_sync_at = 2,
        .file_sync_error = error.InputOutput,
    });
    var report = try wb.saveWithRecalc(a, io, out, fixed_run, .{});
    defer report.deinit(a);

    // Success, with the dormant slot flipped.
    try testing.expect(report.durability.warning);
    try testing.expectEqual(@as(i32, @intFromEnum(std.posix.E.IO)), report.durability.err_code);
    // And the swap already applied — that is the ordering claim.
    try testing.expectEqual(@as(usize, 1), wb.retained.items.len);
    try testing.expectEqualStrings("2", try cellCache(try wb.sheet(0), "B1"));

    var reopened = try Workbook.open(a, base, out);
    defer reopened.deinit();
    try testing.expectEqualStrings("2", try cellCache(try reopened.sheet(0), "B1"));
}

fn expectNoTempFiles(io: std.Io, dir: []const u8) !void {
    var d = try std.Io.Dir.cwd().openDir(io, dir, .{ .iterate = true });
    defer d.close(io);
    var it = d.iterate();
    while (try it.next(io)) |entry| {
        if (std.mem.startsWith(u8, entry.name, ".ztmp")) return error.TempFileLeftBehind;
    }
}

// ─── done-when 6: the commit span is non-cancellable ─────────────

test "cancellation: a cancel armed after the final poll cannot prevent the commit" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const base = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmpPath(a, base, &tmp);
    defer a.free(dir);
    const src = try writeFixture(a, base, dir, "in.xlsx", .{});
    defer a.free(src);

    // Pass one: count the polls. A deadline forces a clock read at every
    // one of them, so the injected `now` IS the poll counter — the same
    // mechanism M5d1 measures §5.5's bound with.
    const far: std.Io.Timestamp = .{ .nanoseconds = std.math.maxInt(i64) };
    var polls: u64 = 0;
    {
        const io = control.inject.wrap(base, .{});
        const out = try std.fs.path.join(a, &.{ dir, "count.xlsx" });
        defer a.free(out);
        var wb = try Workbook.open(a, base, src);
        defer wb.deinit();
        var run = fixed_run;
        run.deadline = far;
        var r = try wb.saveWithRecalc(a, io, out, run, .{});
        r.deinit(a);
        polls = control.inject.state.now_calls;
        try testing.expect(polls > 0);
    }

    // Pass two: the same save, with the cancel flag armed by the LAST
    // clock read — the one the final pre-rename poll performs. `check`
    // reads the token before the clock, so that poll passes; there is no
    // poll after it, so the commit region runs to completion. Setting the
    // flag one read earlier would have refused instead, which is what
    // makes this a statement about placement rather than about luck.
    var flag: u8 = 0;
    const out = try std.fs.path.join(a, &.{ dir, "committed.xlsx" });
    defer a.free(out);
    var wb = try Workbook.open(a, base, src);
    defer wb.deinit();

    const io = control.inject.wrap(base, .{ .trip_at = polls, .trip_flag = &flag });
    var run = fixed_run;
    run.deadline = far;
    run.cancel = .{ .flag = &flag };

    var report = try wb.saveWithRecalc(a, io, out, run, .{});
    defer report.deinit(a);

    // The poll count is identical, which is what makes `trip_at` land
    // where the first pass said it would.
    try testing.expectEqual(polls, control.inject.state.now_calls);
    try testing.expectEqual(@as(u8, 1), flag);
    // Committed: swapped in memory and published on disk, with a token
    // that is triggered right now.
    try testing.expectEqual(@as(usize, 1), wb.retained.items.len);
    try testing.expectEqualStrings("2", try cellCache(try wb.sheet(0), "B1"));
    var reopened = try Workbook.open(a, base, out);
    defer reopened.deinit();
    try testing.expectEqualStrings("2", try cellCache(try reopened.sheet(0), "B1"));
}

test "cancellation: a cancel one poll earlier refuses, and changes nothing" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const base = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmpPath(a, base, &tmp);
    defer a.free(dir);
    const src = try writeFixture(a, base, dir, "in.xlsx", .{});
    defer a.free(src);

    const far: std.Io.Timestamp = .{ .nanoseconds = std.math.maxInt(i64) };
    var polls: u64 = 0;
    {
        const io = control.inject.wrap(base, .{});
        const out = try std.fs.path.join(a, &.{ dir, "count.xlsx" });
        defer a.free(out);
        var wb = try Workbook.open(a, base, src);
        defer wb.deinit();
        var run = fixed_run;
        run.deadline = far;
        var r = try wb.saveWithRecalc(a, io, out, run, .{});
        r.deinit(a);
        polls = control.inject.state.now_calls;
    }

    var flag: u8 = 0;
    const out = try std.fs.path.join(a, &.{ dir, "never.xlsx" });
    defer a.free(out);
    var wb = try Workbook.open(a, base, src);
    defer wb.deinit();

    const io = control.inject.wrap(base, .{ .trip_at = polls - 1, .trip_flag = &flag });
    var run = fixed_run;
    run.deadline = far;
    run.cancel = .{ .flag = &flag };

    try testing.expectError(error.Cancelled, wb.saveWithRecalc(a, io, out, run, .{}));
    try testing.expectEqual(@as(usize, 0), wb.retained.items.len);
    try testing.expectEqualStrings("999", try cellCache(try wb.sheet(0), "B1"));
    try testing.expectError(error.FileNotFound, std.Io.Dir.cwd().access(base, out, .{}));
    try expectNoTempFiles(base, dir);
}

test "cancellation: a token already up before the run does nothing at all" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmpPath(a, io, &tmp);
    defer a.free(dir);
    const src = try writeFixture(a, io, dir, "in.xlsx", .{});
    defer a.free(src);

    var wb = try Workbook.open(a, io, src);
    defer wb.deinit();

    var flag: u8 = 1;
    var run = fixed_run;
    run.cancel = .{ .flag = &flag };

    try testing.expectError(error.Cancelled, wb.recalculate(a, io, run, .{}));
    try testing.expectEqual(@as(usize, 0), wb.retained.items.len);
    try testing.expectEqualStrings("999", try cellCache(try wb.sheet(0), "B1"));
}

test "cancellation: a token that fires between two cells rolls the run back" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const base = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmpPath(a, base, &tmp);
    defer a.free(dir);
    // Two formula cells, so there is a "between" for the cancel to
    // arrive in.
    const src = try writeFixture(a, base, dir, "in.xlsx", .{
        .sheet = "<worksheet xmlns=\"" ++ ns_main ++ "\"><sheetData><row r=\"1\">" ++
            "<c r=\"A1\"><v>1</v></c><c r=\"B1\"><f>A1+1</f><v>999</v></c>" ++
            "<c r=\"C1\"><f>B1+1</f><v>999</v></c>" ++
            "</row></sheetData></worksheet>",
    });
    defer a.free(src);

    var wb = try Workbook.open(a, base, src);
    defer wb.deinit();

    // A deadline makes every poll read the clock, so the injected clock
    // is a poll counter; tripping it on the second read arms the token
    // for the poll the second cell performs.
    var flag: u8 = 0;
    const io = control.inject.wrap(base, .{ .trip_at = 2, .trip_flag = &flag });
    var run = fixed_run;
    run.deadline = .{ .nanoseconds = std.math.maxInt(i64) };
    run.cancel = .{ .flag = &flag };

    try testing.expectError(error.Cancelled, wb.recalculate(a, io, run, .{}));
    try testing.expectEqual(@as(usize, 0), wb.retained.items.len);
    try testing.expectEqualStrings("999", try cellCache(try wb.sheet(0), "B1"));
    try testing.expectEqualStrings("999", try cellCache(try wb.sheet(0), "C1"));
}

test "saveWithRecalc: saving over the source is the same transaction" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmpPath(a, io, &tmp);
    defer a.free(dir);
    const src = try writeFixture(a, io, dir, "in.xlsx", .{});
    defer a.free(src);

    {
        // §5.7.9's named case: the temp file lives in the destination's
        // directory and the rename replaces atomically, so the store may
        // publish over the very bytes it is still reading raw entries
        // from.
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        var r = try wb.saveWithRecalc(a, io, src, fixed_run, .{});
        defer r.deinit(a);
        try testing.expectEqual(@as(u32, 1), r.cells_written);
    }

    var reopened = try Workbook.open(a, io, src);
    defer reopened.deinit();
    try testing.expectEqualStrings("2", try cellCache(try reopened.sheet(0), "B1"));
}

// ─── done-when 7: the pre-M7 gate ────────────────────────────────

fn expectGateRefusal(f: Fixture) !void {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmpPath(a, io, &tmp);
    defer a.free(dir);
    const src = try writeFixture(a, io, dir, "in.xlsx", f);
    defer a.free(src);
    const out = try std.fs.path.join(a, &.{ dir, "out.xlsx" });
    defer a.free(out);

    var wb = try Workbook.open(a, io, src);
    defer wb.deinit();

    const before = (try wb.store.part(sheet_part)).?.bytes;

    try testing.expectError(
        error.FormulaSpillPersistUnsupported,
        wb.recalculate(a, io, fixed_run, .{}),
    );
    // Zero mutation: no generation, and the part is the same *slice*, not
    // merely equal bytes.
    try testing.expectEqual(@as(usize, 0), wb.retained.items.len);
    try testing.expectEqual(before.ptr, (try wb.store.part(sheet_part)).?.bytes.ptr);

    // And the file transaction refuses in the same place, having written
    // nothing.
    try testing.expectError(
        error.FormulaSpillPersistUnsupported,
        wb.saveWithRecalc(a, io, out, fixed_run, .{}),
    );
    try testing.expectError(error.FileNotFound, std.Io.Dir.cwd().access(io, out, .{}));
}

test "pre-M7 gate: a non-1x1 result refuses with zero mutation" {
    try expectGateRefusal(.{ .sheet = "<worksheet xmlns=\"" ++ ns_main ++ "\"><sheetData><row r=\"1\">" ++
        "<c r=\"A1\"><f>{1,2}</f><v>999</v></c>" ++
        "</row></sheetData></worksheet>" });
}

/// A stored CSE over `A1:A2` with both caches stale — the shape M7b1
/// opens: every mutation is `<v>` on a `<c>` that exists.
const sheet_cse =
    "<worksheet xmlns=\"" ++ ns_main ++ "\"><sheetData>" ++
    "<row r=\"1\"><c r=\"A1\"><f t=\"array\" ref=\"A1:A2\">SEQUENCE(2)</f><v>999</v></c></row>" ++
    "<row r=\"2\"><c r=\"A2\"><v>999</v></c></row>" ++
    "</sheetData></worksheet>";

test "M7b1: a legacy CSE persists — anchor and slave caches together, through the file transaction" {
    // The pre-M7 refusal this test replaces is now §5.8b's narrow: a
    // covered CSE is `<v>`+`t` on existing cells, no cm/vm byte
    // involved, and it commits.
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmpPath(a, io, &tmp);
    defer a.free(dir);
    const src = try writeFixture(a, io, dir, "in.xlsx", .{ .sheet = sheet_cse });
    defer a.free(src);
    const out = try std.fs.path.join(a, &.{ dir, "out.xlsx" });
    defer a.free(out);

    {
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        var r = try wb.saveWithRecalc(a, io, out, fixed_run, .{});
        defer r.deinit(a);
        try testing.expectEqual(@as(u32, 2), r.cells_written);
    }

    var reopened = try Workbook.open(a, io, out);
    defer reopened.deinit();
    try testing.expectEqualStrings("1", try cellCache(try reopened.sheet(0), "A1"));
    try testing.expectEqualStrings("2", try cellCache(try reopened.sheet(0), "A2"));
    // The declared range survives byte-identically — the patcher never
    // addresses a byte inside `<f>` except the ref value, and a CSE's
    // ref does not move.
    const bytes = (try reopened.store.part(sheet_part)).?.bytes;
    try testing.expect(std.mem.indexOf(u8, bytes, "<f t=\"array\" ref=\"A1:A2\">SEQUENCE(2)</f>") != null);
}

test "M7b1: an injected rename failure after multi-cell staging leaves prior bytes intact" {
    // DONE-WHEN 5, file side: the M5b2 zero-mutation gates extended to
    // the multi-cell staging M7b1 opened — anchor plus slave staged,
    // failure injected at the commit point, destination and memory
    // both untouched.
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const base = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmpPath(a, base, &tmp);
    defer a.free(dir);
    const src = try writeFixture(a, base, dir, "in.xlsx", .{ .sheet = sheet_cse });
    defer a.free(src);
    const dest = try writeFixture(a, base, dir, "dest.xlsx", .{ .sheet = sheet_no_formula });
    defer a.free(dest);

    const before = try readAll(a, base, dest);
    defer a.free(before);

    var wb = try Workbook.open(a, base, src);
    defer wb.deinit();

    const io = control.inject.wrap(base, .{ .fail_rename = true });
    try testing.expectError(error.AccessDenied, wb.saveWithRecalc(a, io, dest, fixed_run, .{}));

    try testing.expectEqual(@as(usize, 0), wb.retained.items.len);
    try testing.expectEqualStrings("999", try cellCache(try wb.sheet(0), "A1"));
    try testing.expectEqualStrings("999", try cellCache(try wb.sheet(0), "A2"));

    const after = try readAll(a, base, dest);
    defer a.free(after);
    try testing.expectEqualSlices(u8, before, after);
    try expectNoTempFiles(base, dir);
}

test "M7b1: an injected sync failure after multi-cell staging is pre-commit too" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const base = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmpPath(a, base, &tmp);
    defer a.free(dir);
    const src = try writeFixture(a, base, dir, "in.xlsx", .{ .sheet = sheet_cse });
    defer a.free(src);
    const dest = try std.fs.path.join(a, &.{ dir, "absent.xlsx" });
    defer a.free(dest);

    var wb = try Workbook.open(a, base, src);
    defer wb.deinit();

    const io = control.inject.wrap(base, .{ .fail_file_sync_at = 1 });
    try testing.expectError(error.InputOutput, wb.saveWithRecalc(a, io, dest, fixed_run, .{}));

    try testing.expectEqual(@as(usize, 0), wb.retained.items.len);
    try testing.expectEqualStrings("999", try cellCache(try wb.sheet(0), "A1"));
    try testing.expectEqualStrings("999", try cellCache(try wb.sheet(0), "A2"));
    try testing.expectError(error.FileNotFound, std.Io.Dir.cwd().access(base, dest, .{}));
    try expectNoTempFiles(base, dir);
}

test "pre-M7 gate: a dynamic-array anchor refuses with zero mutation" {
    try expectGateRefusal(.{
        .sheet = "<worksheet xmlns=\"" ++ ns_main ++ "\"><sheetData><row r=\"1\">" ++
            "<c r=\"A1\" cm=\"1\"><f t=\"array\" ref=\"A1\">1+1</f><v>999</v></c>" ++
            "</row></sheetData></worksheet>",
        .metadata = metadata_dynamic_array,
    });
}

test "pre-M7 gate: a spill the MODEL placed still refuses persistence (M7a regression)" {
    // M7a places tails in the model — this run's SEQUENCE(3) spills
    // A1:A3 before staging is reached — and §5.8b's approved mutation
    // set is still M7b1's: the persist half stays closed, with zero
    // mutation, exactly as before placement existed.
    try expectGateRefusal(.{
        .sheet = "<worksheet xmlns=\"" ++ ns_main ++ "\"><sheetData><row r=\"1\">" ++
            "<c r=\"A1\" cm=\"1\"><f>SEQUENCE(3)</f><v>0</v></c>" ++
            "</row></sheetData></worksheet>",
        .metadata = metadata_dynamic_array,
    });
}

// ─── the logical-view gate ───────────────────────────────────────

test "logical-view gate: a sheet with staged appends refuses before the model" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmpPath(a, io, &tmp);
    defer a.free(dir);
    const src = try writeFixture(a, io, dir, "in.xlsx", .{});
    defer a.free(src);

    var wb = try Workbook.open(a, io, src);
    defer wb.deinit();

    var ws = try wb.sheet(0);
    try ws.appendRows(&.{&.{.{ .number = 7 }}});

    try testing.expectError(
        error.SheetHasUnsavedAppends,
        wb.recalculate(a, io, fixed_run, .{}),
    );
    try testing.expectEqual(@as(usize, 0), wb.retained.items.len);
    try testing.expectEqualStrings("999", try cellCache(try wb.sheet(0), "B1"));
}

test "logical-view gate: a staged delta is modeled, not refused" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmpPath(a, io, &tmp);
    defer a.free(dir);
    const src = try writeFixture(a, io, dir, "in.xlsx", .{});
    defer a.free(src);

    var wb = try Workbook.open(a, io, src);
    defer wb.deinit();

    var ws = try wb.sheet(0);
    try ws.setCell("A1", .{ .number = 41 });

    var report = try wb.recalculate(a, io, fixed_run, .{});
    defer report.deinit(a);
    // The delta is the second layer of the logical view, so `B1=A1+1`
    // sees 41 rather than the 1 the part stores.
    try testing.expectEqualStrings("42", try cellCache(try wb.sheet(0), "B1"));
}

// ─── the embedding-staleness preflight ───────────────────────────

test "embedding preflight: a staged cell inside a coverage refuses" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmpPath(a, io, &tmp);
    defer a.free(dir);
    const src = try writeFixture(a, io, dir, "in.xlsx", .{ .embeddings = true });
    defer a.free(src);

    var wb = try Workbook.open(a, io, src);
    defer wb.deinit();

    const before = (try wb.store.part(sheet_part)).?.bytes;
    try testing.expectError(
        error.FormulaStaleEmbeddings,
        wb.recalculate(a, io, fixed_run, .{}),
    );
    try testing.expectEqual(@as(usize, 0), wb.retained.items.len);
    try testing.expectEqual(before.ptr, (try wb.store.part(sheet_part)).?.bytes.ptr);
}

test "embedding preflight: a recalc that changes no byte is not stale" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmpPath(a, io, &tmp);
    defer a.free(dir);
    // Same coverage, but the cache already agrees with the formula: the
    // patch produces no edit, so no hash can have gone stale. This is
    // what makes recalculating an embedded workbook twice legal.
    const src = try writeFixture(a, io, dir, "in.xlsx", .{
        .embeddings = true,
        .sheet = "<worksheet xmlns=\"" ++ ns_main ++ "\"><sheetData><row r=\"1\">" ++
            "<c r=\"A1\"><v>1</v></c><c r=\"B1\"><f>A1+1</f><v>2</v></c>" ++
            "</row></sheetData></worksheet>",
    });
    defer a.free(src);

    var wb = try Workbook.open(a, io, src);
    defer wb.deinit();
    var report = try wb.recalculate(a, io, fixed_run, .{});
    defer report.deinit(a);
    try testing.expectEqual(@as(u32, 0), report.cells_written);
}

// ─── §5.7.7's policy, through the pipeline ───────────────────────

test "refusal: an unsupported function refuses, and mark-only marks instead" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmpPath(a, io, &tmp);
    defer a.free(dir);
    const unregistered = "<worksheet xmlns=\"" ++ ns_main ++ "\"><sheetData><row r=\"1\">" ++
        "<c r=\"A1\"><v>1</v></c><c r=\"B1\"><f>NOSUCHFN(A1)</f><v>999</v></c>" ++
        "</row></sheetData></worksheet>";
    const src = try writeFixture(a, io, dir, "in.xlsx", .{ .sheet = unregistered });
    defer a.free(src);

    {
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        try testing.expectError(
            error.FormulaUnsupportedFunction,
            wb.recalculate(a, io, fixed_run, .{}),
        );
        try testing.expectEqual(@as(usize, 0), wb.retained.items.len);
    }
    {
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        var report = try wb.recalculate(a, io, fixed_run, .{
            .on_unsupported = .keep_stale_and_mark,
        });
        defer report.deinit(a);
        // §5.7.7's mark-only path: the caches stay as they were and the
        // file asks the next consumer to calculate them.
        try testing.expect(report.kept_stale);
        try testing.expectEqualStrings("999", try cellCache(try wb.sheet(0), "B1"));
        const part = (try wb.store.part("xl/workbook.xml")).?;
        try testing.expect(std.mem.indexOf(u8, part.bytes, "fullCalcOnLoad=\"1\"") != null);
    }
}

test "refusal_out: the census crosses the seam with the refusing cell" {
    const a = testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var tmp = testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmpPath(a, io, &tmp);
    defer a.free(dir);
    const unregistered = "<worksheet xmlns=\"" ++ ns_main ++ "\"><sheetData><row r=\"1\">" ++
        "<c r=\"A1\"><v>1</v></c><c r=\"B1\"><f>NOSUCHFN(A1)</f><v>999</v></c>" ++
        "</row></sheetData></worksheet>";
    const src = try writeFixture(a, io, dir, "in.xlsx", .{ .sheet = unregistered });
    defer a.free(src);

    var wb = try Workbook.open(a, io, src);
    defer wb.deinit();

    // The M9a2 seam (decision M9a1-4): the same refusal, but the caller
    // supplied a slot, so the refusing cell arrives instead of dying
    // inside the pipeline. B1 is row 1, col 1 (0-based).
    var refusal: recalc_txn.Refusal = undefined;
    try testing.expectError(
        error.FormulaUnsupportedFunction,
        wb.recalculate(a, io, fixed_run, .{ .refusal_out = &refusal }),
    );
    defer refusal.deinit(a);
    try testing.expectEqual(recalc_txn.Refusal.Reason.unsupported_construct, refusal.reason);
    try testing.expectEqual(@as(usize, 1), refusal.census.len);
    try testing.expectEqual(engine.decode.PlaneTwo.FormulaUnsupportedFunction, refusal.census[0].plane);
    try testing.expectEqual(@as(u32, 0), refusal.census[0].sheet);
    try testing.expectEqual(@as(u32, 1), refusal.census[0].row);
    try testing.expectEqual(@as(u32, 1), refusal.census[0].col);
    try testing.expect(!refusal.census_truncated);
    // And the workbook is untouched, refusal_out or not.
    try testing.expectEqual(@as(usize, 0), wb.retained.items.len);
    try testing.expectEqualStrings("999", try cellCache(try wb.sheet(0), "B1"));
}
