//! What a run is *given* — `RunInputs`, the workbook-derived
//! `CalcState`, and the §9 resource budget that enforces them
//! (`goal_formula.md` §5.5, §9).
//!
//! M3b of the tier-D1 ladder.
//!
//! The library reads no clocks and no entropy
//! ------------------------------------------
//! `now_utc_ms` and `rng_seed` are **required fields with no defaults**,
//! and nothing in `src/formula/` calls `std.time` or an entropy source.
//! That is what makes determinism checkable rather than aspirational:
//! equal `RunInputs` + equal bytes ⇒ equal output, and there is no third
//! input hiding in a `getrandom` call. The CLI, Python, and Spark
//! adapters resolve omitted values *once* and echo what they resolved
//! (§5.5's per-layer default table); Zig and C callers supply them.
//!
//! Two fields are deliberately outside the fingerprint
//! ---------------------------------------------------
//! `deadline` and `cancel` never enter `EffectiveRunInputs`. A run that
//! was cancelled has no output to be deterministic *about*, and keying a
//! cache on "how long the caller was willing to wait" would split
//! identical work across entries for no reason. Cancellation appears in
//! terminal status, never in identity. `test "effective inputs exclude
//! the two fields §5.5 excludes"` is the gate.
//!
//! Byte limits and work limits are different mechanisms
//! ----------------------------------------------------
//! §9 splits them on purpose. Bytes are counted by an allocator wrapper
//! — every allocation on the run arena passes through it, so nothing can
//! be charged to the wrong budget by forgetting a call. Work (cell
//! evaluations, SCC passes, comparisons) can burn CPU *without*
//! allocating, so it needs explicit counters, and each counter lands
//! with the code that can actually enforce it: `WorkLimits` below ships
//! the three the dependency graph charges (M5a1), and the iteration
//! engine's two follow at M5a2.
//!
//! An exhausted budget is a **refusal, not an allocation failure**.
//! `std.mem.Allocator` can only say `OutOfMemory`, so the counter
//! records *which* category tripped and the caller maps the failure onto
//! `FormulaLimitExceeded`. Without that record a genuine OOM and a
//! deliberate limit would be indistinguishable, and only one of them is
//! the caller's fault.

const std = @import("std");
const assert = std.debug.assert;

const coords = @import("zlsx_refs");
const value = @import("value.zig");

// ─── run inputs (§5.5) ───────────────────────────────────────────

/// v1 enum, deliberately CLOSED. `CHAR`/`CODE` resolve their code page
/// through it; a second profile is an M10+ addition that would join
/// every fingerprint when it lands.
pub const PlatformProfile = enum { windows_1252 };

/// §5.4d. Both versions are v1 scope; the flag is workbook-derived.
pub const CompatibilityVersion = enum { cv1, cv2 };

/// §5.4a. Workbook-derived (`workbookPr@date1904`), never caller-set.
pub const DateSystem = enum { d1900, d1904 };

/// §5.6c. Defaults are Excel's: iteration off, 100 passes, delta 0.001.
pub const Iteration = struct {
    enabled: bool = false,
    /// Clamped to Excel's documented maximum of 32 767 by the reader.
    count: u16 = 100,
    delta: f64 = 0.001,
};

/// **Workbook-derived, never caller-writable.** Separated from
/// `RunInputs` for exactly that reason: a caller who could set
/// `date_system` could silently reinterpret every date in a file.
pub const CalcState = struct {
    date_system: DateSystem = .d1900,
    iteration: Iteration = .{},
    /// `fullPrecision="0"` always refuses (§10), so a run that gets this
    /// far has it true.
    full_precision: bool = true,
    /// §5.4d: **absent compatibility metadata is CV1**, so that is the
    /// default here. M4f corrected it from `.cv2`, which was the right
    /// answer to a different question — CV2 is what Excel writes into
    /// NEW workbooks, not what a workbook that says nothing means. The
    /// files that say nothing are every pre-2024 workbook and every file
    /// zlsx's own Writer emits (`fresh_emit.zig` writes no metadata
    /// part), and under `.cv2` all of them would have counted an astral
    /// character once where Excel counts it twice.
    text_compat: CompatibilityVersion = .cv1,
};

/// Where a site-dependent construct is being evaluated. The typed
/// coordinates carry their own base, so the field names do not repeat it
/// the way §5.5's sketch does.
pub const EvalSite = struct {
    row: coords.Row,
    col: coords.Col,
};

/// Cooperative cancellation, as a no-alloc union over the two storage
/// kinds a caller can actually have: an atomic for multi-threaded hosts,
/// and a plain volatile flag for a signal handler in a single-threaded
/// CLI (writing an atomic from a signal handler is not async-signal-safe
/// in general). One `isTriggered` seam, so every poll site is identical.
///
/// **Defined in `pkg/control.zig` since M5d1, and re-exported here.** The
/// archive layer polls the same token the evaluator does — `Writer`'s
/// `saveToOwnedBufferControlled` takes a `control.Control` whose `cancel`
/// field has to be assignable from a `RunInputs.cancel` — and
/// `src/writer.zig` does not import this engine. Two structurally
/// identical unions would have needed a conversion function kept in sync
/// by hand, so the definition moved down to a std-only leaf both trees
/// sit above instead. Nothing in this module changed but the address.
pub const CancelToken = @import("zlsx_control").CancelToken;

/// M5d1's erased poll seam, re-exported for the same reason the token
/// is: M9c1's `WorkBudget` carries one so the engine's own poll points
/// (§5.5's work-unit bound) read the seam every other long operation
/// already reads, and a caller that has a `control.Watch` hands its
/// `poller()` straight across.
pub const Poller = @import("zlsx_control").Poller;

pub const utc_offset_min_min: i32 = -1440;
pub const utc_offset_min_max: i32 = 1440;

pub const RunInputsError = error{
    /// Outside [-1440, 1440]. Validated **pre-narrowing**, where the
    /// value is still an `i32` from the ABI.
    UtcOffsetOutOfRange,
    LimitOutOfRange,
};

/// §5.5. `now_utc_ms` and `rng_seed` have no defaults on purpose: a
/// library that defaulted them would read a clock or an entropy source,
/// and then "equal inputs ⇒ equal output" would be false in a way no
/// test could see.
pub const RunInputs = struct {
    now_utc_ms: i64,
    rng_seed: u64,
    limits: ResourceLimits,

    /// Fixed civil offset for `NOW()`/`TODAY()`. **UTC at every layer**:
    /// the Zig 0.16 stdlib has no portable local-timezone resolver and
    /// zlsx is stdlib-only, so a caller passes an offset or gets UTC.
    /// TZif resolution is M10+.
    utc_offset_min: i32 = 0,
    fidelity: value.Fidelity = .excel,
    platform_profile: PlatformProfile = .windows_1252,
    /// **Standalone eval only.** A recalc derives dialect per stored cell
    /// through `EvalEnv.dialectOf` and normalizes this field out of its
    /// fingerprint, so there is no phantom key (§5.3b).
    dialect: value.Dialect = .dynamic_array,

    /// Absolute monotonic deadline on the `.awake` clock. Excluded from
    /// `EffectiveRunInputs`.
    deadline: ?std.Io.Timestamp = null,
    /// Excluded from `EffectiveRunInputs`.
    cancel: ?CancelToken = null,

    pub fn validate(self: RunInputs) RunInputsError!void {
        if (self.utc_offset_min < utc_offset_min_min or self.utc_offset_min > utc_offset_min_max) {
            return error.UtcOffsetOutOfRange;
        }
        try self.limits.validate();
    }

    pub fn isCancelled(self: RunInputs) bool {
        const token = self.cancel orelse return false;
        return token.isTriggered();
    }

    /// The fingerprintable projection. `deadline` and `cancel` are
    /// absent by construction rather than by a filter someone can
    /// forget: a field that is not in this struct cannot leak into a
    /// cache key.
    pub fn effective(self: RunInputs, op: Operation) EffectiveRunInputs {
        return .{
            .now_utc_ms = self.now_utc_ms,
            .utc_offset_min = self.utc_offset_min,
            .rng_seed = self.rng_seed,
            .fidelity = self.fidelity,
            .platform_profile = self.platform_profile,
            // A recalc's dialect comes from the stored cells, so keying
            // on this field would split identical work across entries
            // that differ only in a value nobody read.
            .dialect = switch (op) {
                .standalone_eval => self.dialect,
                .recalc => null,
            },
            .limits = self.limits,
        };
    }
};

/// Which operation's projection to take. §5.5: the effective-input
/// projection differs per operation, and the difference is exactly the
/// dialect field.
pub const Operation = enum { standalone_eval, recalc };

pub const EffectiveRunInputs = struct {
    now_utc_ms: i64,
    utc_offset_min: i32,
    rng_seed: u64,
    fidelity: value.Fidelity,
    platform_profile: PlatformProfile,
    /// Null for a recalc, which derives dialect per stored cell.
    dialect: ?value.Dialect,
    limits: ResourceLimits,

    pub fn eql(a: EffectiveRunInputs, b: EffectiveRunInputs) bool {
        return std.meta.eql(a, b);
    }
};

// ─── §9 byte budget ──────────────────────────────────────────────

/// What a charge is charged *to*. Five categories because §9 lists five
/// aggregate limits, and one number could not tell a caller which of
/// them they hit.
pub const Category = enum {
    run_arena,
    matrix_cells,
    string_payload,
    retained_asts,
    diagnostics,

    pub fn unit(self: Category) []const u8 {
        return switch (self) {
            .matrix_cells => "cells",
            else => "bytes",
        };
    }
};

pub const category_count = @typeInfo(Category).@"enum".fields.len;

/// §9's aggregate byte/count limits. Caller-adjustable in Zig and C;
/// the CLI and Python fix them at defaults in v1.
///
/// `parser.Limits` (M2) is the other half of §9 — the *shape* of a
/// parse. They are separate structs because they are enforced by
/// separate mechanisms at separate times, and §9's work counters are a
/// third: `WorkLimits` below, whose graph three land at M5a1 and whose
/// iteration two (`max_scc_iterations`, `max_dynamic_passes`) land at
/// M5a2 with the engine that enforces them.
pub const ResourceLimits = struct {
    max_run_arena_bytes: u64 = default_run_arena_bytes,
    max_matrix_cells: u64 = default_matrix_cells,
    max_string_payload_bytes: u64 = default_string_payload_bytes,
    max_retained_ast_bytes: u64 = default_retained_ast_bytes,
    max_diagnostics_bytes: u64 = default_diagnostics_bytes,

    pub const default_run_arena_bytes: u64 = 1 << 30; // 1 GiB
    pub const default_matrix_cells: u64 = 8_000_000;
    pub const default_string_payload_bytes: u64 = 256 << 20;
    pub const default_retained_ast_bytes: u64 = 128 << 20;
    pub const default_diagnostics_bytes: u64 = 1 << 20;

    /// §9: hard maxima are 4× the defaults. A ceiling on the ceiling,
    /// so "caller-adjustable" cannot mean "caller-disabled".
    pub const hard_multiplier: u64 = 4;

    pub fn defaultFor(cat: Category) u64 {
        return switch (cat) {
            .run_arena => default_run_arena_bytes,
            .matrix_cells => default_matrix_cells,
            .string_payload => default_string_payload_bytes,
            .retained_asts => default_retained_ast_bytes,
            .diagnostics => default_diagnostics_bytes,
        };
    }

    pub fn hardMaxFor(cat: Category) u64 {
        return defaultFor(cat) * hard_multiplier;
    }

    pub fn get(self: ResourceLimits, cat: Category) u64 {
        return switch (cat) {
            .run_arena => self.max_run_arena_bytes,
            .matrix_cells => self.max_matrix_cells,
            .string_payload => self.max_string_payload_bytes,
            .retained_asts => self.max_retained_ast_bytes,
            .diagnostics => self.max_diagnostics_bytes,
        };
    }

    pub fn set(self: *ResourceLimits, cat: Category, v: u64) void {
        switch (cat) {
            .run_arena => self.max_run_arena_bytes = v,
            .matrix_cells => self.max_matrix_cells = v,
            .string_payload => self.max_string_payload_bytes = v,
            .retained_asts => self.max_retained_ast_bytes = v,
            .diagnostics => self.max_diagnostics_bytes = v,
        }
    }

    pub fn validate(self: ResourceLimits) RunInputsError!void {
        inline for (@typeInfo(Category).@"enum".fields) |f| {
            const cat: Category = @enumFromInt(f.value);
            const v = self.get(cat);
            if (v == 0 or v > hardMaxFor(cat)) return error.LimitOutOfRange;
        }
    }
};

pub const BudgetError = error{
    /// A category's limit was reached. The category is on the budget.
    LimitExceeded,
    OutOfMemory,
};

/// The counted allocator.
///
/// One wrapper, five categories, and a `tripped` field that survives the
/// failure — because `std.mem.Allocator` can only report `OutOfMemory`
/// and a refusal has to be distinguishable from a machine running out of
/// memory. The evaluator maps a trip onto `FormulaLimitExceeded`; a
/// genuine allocator failure stays `OutOfMemory`.
///
/// **Nothing is mutated on a refusal.** The charge is checked *before*
/// the backing allocator is called, so a rejected allocation leaves both
/// the counter and the heap exactly as they were.
pub const Budget = struct {
    backing: std.mem.Allocator,
    limits: ResourceLimits,
    used: [category_count]u64 = @splat(0),
    peak: [category_count]u64 = @splat(0),
    /// The first category to trip. First, not last: the later failures
    /// are consequences of unwinding the first.
    tripped: ?Category = null,
    handles: [category_count]Handle = undefined,

    const Handle = struct { owner: *Budget, cat: Category };

    pub fn init(backing: std.mem.Allocator, limits: ResourceLimits) Budget {
        return .{ .backing = backing, .limits = limits };
    }

    /// An allocator that charges `cat`. Must be called on a `Budget`
    /// whose address is stable — the handles point back at it.
    pub fn allocator(self: *Budget, cat: Category) std.mem.Allocator {
        const i = @intFromEnum(cat);
        self.handles[i] = .{ .owner = self, .cat = cat };
        return .{ .ptr = &self.handles[i], .vtable = &vtable };
    }

    pub fn usedBy(self: Budget, cat: Category) u64 {
        return self.used[@intFromEnum(cat)];
    }

    pub fn peakOf(self: Budget, cat: Category) u64 {
        return self.peak[@intFromEnum(cat)];
    }

    pub fn remaining(self: Budget, cat: Category) u64 {
        const limit = self.limits.get(cat);
        const u = self.usedBy(cat);
        return if (u >= limit) 0 else limit - u;
    }

    /// Charge a non-allocating quantity — `matrix_cells` is a count, not
    /// a byte total, and charging it through the allocator would make it
    /// depend on `@sizeOf(ScalarValue)`.
    pub fn charge(self: *Budget, cat: Category, n: u64) BudgetError!void {
        if (!self.tryCharge(cat, n)) return error.LimitExceeded;
    }

    pub fn release(self: *Budget, cat: Category, n: u64) void {
        const i = @intFromEnum(cat);
        assert(self.used[i] >= n);
        self.used[i] -= n;
    }

    fn tryCharge(self: *Budget, cat: Category, n: u64) bool {
        const i = @intFromEnum(cat);
        const limit = self.limits.get(cat);
        const next = std.math.add(u64, self.used[i], n) catch {
            if (self.tripped == null) self.tripped = cat;
            return false;
        };
        if (next > limit) {
            if (self.tripped == null) self.tripped = cat;
            return false;
        }
        self.used[i] = next;
        if (next > self.peak[i]) self.peak[i] = next;
        return true;
    }

    const vtable: std.mem.Allocator.VTable = .{
        .alloc = vtAlloc,
        .resize = vtResize,
        .remap = vtRemap,
        .free = vtFree,
    };

    fn handleOf(ctx: *anyopaque) *Handle {
        return @ptrCast(@alignCast(ctx));
    }

    fn vtAlloc(ctx: *anyopaque, len: usize, alignment: std.mem.Alignment, ret_addr: usize) ?[*]u8 {
        const h = handleOf(ctx);
        if (!h.owner.tryCharge(h.cat, len)) return null;
        const p = h.owner.backing.rawAlloc(len, alignment, ret_addr) orelse {
            // The backing allocator failed, not the budget. Give the
            // charge back so a later, smaller request can still succeed.
            h.owner.release(h.cat, len);
            return null;
        };
        return p;
    }

    fn vtResize(ctx: *anyopaque, memory: []u8, alignment: std.mem.Alignment, new_len: usize, ret_addr: usize) bool {
        const h = handleOf(ctx);
        if (new_len > memory.len) {
            const delta = new_len - memory.len;
            if (!h.owner.tryCharge(h.cat, delta)) return false;
            if (!h.owner.backing.rawResize(memory, alignment, new_len, ret_addr)) {
                h.owner.release(h.cat, delta);
                return false;
            }
            return true;
        }
        if (!h.owner.backing.rawResize(memory, alignment, new_len, ret_addr)) return false;
        h.owner.release(h.cat, memory.len - new_len);
        return true;
    }

    fn vtRemap(ctx: *anyopaque, memory: []u8, alignment: std.mem.Alignment, new_len: usize, ret_addr: usize) ?[*]u8 {
        const h = handleOf(ctx);
        if (new_len > memory.len) {
            const delta = new_len - memory.len;
            if (!h.owner.tryCharge(h.cat, delta)) return null;
            const p = h.owner.backing.rawRemap(memory, alignment, new_len, ret_addr) orelse {
                h.owner.release(h.cat, delta);
                return null;
            };
            return p;
        }
        const p = h.owner.backing.rawRemap(memory, alignment, new_len, ret_addr) orelse return null;
        h.owner.release(h.cat, memory.len - new_len);
        return p;
    }

    fn vtFree(ctx: *anyopaque, memory: []u8, alignment: std.mem.Alignment, ret_addr: usize) void {
        const h = handleOf(ctx);
        h.owner.backing.rawFree(memory, alignment, ret_addr);
        h.owner.release(h.cat, memory.len);
    }
};

// ─── §9 work counters ────────────────────────────────────────────

/// §9's third enforcement mechanism, and the reason it is a third:
/// `ResourceLimits` counts bytes an allocator hands out and
/// `parser.Limits` bounds the shape of a parse, but a dependency graph
/// can burn an unbounded amount of CPU while allocating almost nothing.
/// A whole-column reference on a sheet with a million stored cells
/// allocates one node and walks a million coordinates; only an explicit
/// counter catches that.
///
/// M5a1 landed the three counters the graph needs; M5a2 adds the two the
/// iteration engine needs, which is when they acquired an enforcement
/// site — a limit with none is a limit that lies.
///
/// The five are not the same *shape* of bound, and `kind` says so rather
/// than leaving a reader to infer it from which functions touch which
/// member. A total only grows, a depth unwinds, and a per-scope bound is
/// re-counted from zero in every scope it governs: §5.6c gives each SCC
/// its own pass counter, so charging passes into one running total would
/// make two components that each iterate legally refuse together.
pub const WorkCategory = enum {
    /// Edges admitted into the dependency graph.
    dependency_edges,
    /// Formula-cell evaluations a closure plan will perform.
    total_cell_evals,
    /// Cell-to-cell hops on the closure-discovery stack. A *depth*, so
    /// it is released as the walk unwinds; the other two only ever grow.
    eval_depth,
    /// §5.6c's resource ceiling on one SCC's iteration passes. The
    /// caller's, and never the workbook's own `calcPr@iterateCount`:
    /// reaching the semantic bound is success, reaching this one is a
    /// refusal, and conflating them would let a resource cap publish
    /// caches the workbook did not ask for.
    scc_iterations,
    /// §5.6e's ceiling on graph-rebuild passes.
    dynamic_passes,

    pub const Kind = enum {
        /// Monotone across a whole run.
        total,
        /// Charged on the way in, released on the way out.
        depth,
        /// Counted from zero inside each scope it governs, and therefore
        /// read by the engine rather than accumulated in `WorkCounters`.
        per_scope,
    };

    pub fn kind(self: WorkCategory) Kind {
        return switch (self) {
            .dependency_edges, .total_cell_evals => .total,
            .eval_depth => .depth,
            .scc_iterations, .dynamic_passes => .per_scope,
        };
    }

    pub fn unit(self: WorkCategory) []const u8 {
        return switch (self) {
            .dependency_edges => "edges",
            .total_cell_evals => "evaluations",
            .eval_depth => "cells",
            .scc_iterations => "passes",
            .dynamic_passes => "passes",
        };
    }
};

pub const work_category_count = @typeInfo(WorkCategory).@"enum".fields.len;

/// §9's work limits. Same contract as `ResourceLimits`: caller-adjustable
/// in Zig and C, fixed at defaults by the CLI and Python in v1, hard
/// maxima 4× the default, resolved values echoed and fingerprinted.
pub const WorkLimits = struct {
    max_dependency_edges: u64 = default_dependency_edges,
    max_total_cell_evals: u64 = default_total_cell_evals,
    max_eval_depth: u64 = default_eval_depth,
    max_scc_iterations: u64 = default_scc_iterations,
    max_dynamic_passes: u64 = default_dynamic_passes,

    pub const default_dependency_edges: u64 = 50_000_000;
    pub const default_total_cell_evals: u64 = 50_000_000;
    pub const default_eval_depth: u64 = 512;
    /// Excel's own maximum `iterateCount`. The resource ceiling defaults
    /// to it so that, out of the box, the caller's cap can never be the
    /// bound that fires first — a default that refused a workbook Excel
    /// would have calculated would be a default that changes answers.
    pub const default_scc_iterations: u64 = 32_767;
    pub const default_dynamic_passes: u64 = 3;

    pub const hard_multiplier: u64 = 4;

    /// The two §9 states as exceptions to `hard_multiplier`.
    ///
    /// `scc_iterations` cannot be raised at all: 32 767 is the largest
    /// `iterateCount` a workbook can carry, so a ceiling above it bounds
    /// nothing that the semantic clamp has not already bounded.
    /// `dynamic_passes` stops at 10 rather than at 12 because §5.6e
    /// names 10, and a limit whose hard maximum is derived where the
    /// spec states one is a limit that disagrees with its own spec.
    pub const hard_scc_iterations: u64 = 32_767;
    pub const hard_dynamic_passes: u64 = 10;

    pub fn defaultFor(cat: WorkCategory) u64 {
        return switch (cat) {
            .dependency_edges => default_dependency_edges,
            .total_cell_evals => default_total_cell_evals,
            .eval_depth => default_eval_depth,
            .scc_iterations => default_scc_iterations,
            .dynamic_passes => default_dynamic_passes,
        };
    }

    pub fn hardMaxFor(cat: WorkCategory) u64 {
        return switch (cat) {
            .scc_iterations => hard_scc_iterations,
            .dynamic_passes => hard_dynamic_passes,
            else => defaultFor(cat) * hard_multiplier,
        };
    }

    pub fn get(self: WorkLimits, cat: WorkCategory) u64 {
        return switch (cat) {
            .dependency_edges => self.max_dependency_edges,
            .total_cell_evals => self.max_total_cell_evals,
            .eval_depth => self.max_eval_depth,
            .scc_iterations => self.max_scc_iterations,
            .dynamic_passes => self.max_dynamic_passes,
        };
    }

    pub fn set(self: *WorkLimits, cat: WorkCategory, v: u64) void {
        switch (cat) {
            .dependency_edges => self.max_dependency_edges = v,
            .total_cell_evals => self.max_total_cell_evals = v,
            .eval_depth => self.max_eval_depth = v,
            .scc_iterations => self.max_scc_iterations = v,
            .dynamic_passes => self.max_dynamic_passes = v,
        }
    }

    pub fn validate(self: WorkLimits) RunInputsError!void {
        inline for (@typeInfo(WorkCategory).@"enum".fields) |f| {
            const cat: WorkCategory = @enumFromInt(f.value);
            const v = self.get(cat);
            if (v == 0 or v > hardMaxFor(cat)) return error.LimitOutOfRange;
        }
    }
};

/// The counters themselves.
///
/// **Charge and release sites, named per counter** (§9's own
/// requirement — a counter whose enforcement point is not written down
/// is a counter nobody can audit):
///
/// | counter | charged at | released at |
/// |---|---|---|
/// | `dependency_edges` | `graph.Builder.addEdge`, once per admitted edge | never — a built graph keeps its edges |
/// | `total_cell_evals` | `graph.Graph.plan`, once per cell admitted to the plan | never — the plan is exact, so charging at admission refuses *before* the first evaluation instead of halfway through one |
/// | `eval_depth` | `graph.Graph.plan`'s closure walk, on pushing a cell-like node | the same walk, on popping it |
/// | `scc_iterations` | nowhere — `iterate.Engine` re-counts it per component and refuses on its own | — |
/// | `dynamic_passes` | nowhere — `iterate.Engine` re-counts it per run of the outer loop | — |
///
/// The last two rows say "nowhere" and mean it: they are `.per_scope`
/// bounds, and `charge` rejects one at compile-nothing-runtime-assert.
/// Their *limits* still live in `WorkLimits` so a caller configures all
/// five in one place and one `validate` covers them, but a per-scope
/// bound accumulated across scopes would refuse a workbook whose every
/// component iterated legally.
///
/// Nothing is mutated on a refusal: like `Budget`, the check happens
/// before the counter moves, so a rejected charge leaves the counter
/// exactly where it was.
pub const WorkCounters = struct {
    limits: WorkLimits = .{},
    used: [work_category_count]u64 = @splat(0),
    peak: [work_category_count]u64 = @splat(0),
    /// The first category to trip, kept for the same reason `Budget`
    /// keeps its own: the caller's error type says "a limit was hit" and
    /// this says which.
    tripped: ?WorkCategory = null,

    pub fn usedBy(self: WorkCounters, cat: WorkCategory) u64 {
        return self.used[@intFromEnum(cat)];
    }

    pub fn peakOf(self: WorkCounters, cat: WorkCategory) u64 {
        return self.peak[@intFromEnum(cat)];
    }

    pub fn charge(self: *WorkCounters, cat: WorkCategory, n: u64) error{LimitExceeded}!void {
        assert(cat.kind() != .per_scope);
        const i = @intFromEnum(cat);
        const limit = self.limits.get(cat);
        const next = std.math.add(u64, self.used[i], n) catch {
            if (self.tripped == null) self.tripped = cat;
            return error.LimitExceeded;
        };
        if (next > limit) {
            if (self.tripped == null) self.tripped = cat;
            return error.LimitExceeded;
        }
        self.used[i] = next;
        if (next > self.peak[i]) self.peak[i] = next;
    }

    /// Only a `.depth` unwinds. The totals are monotone — releasing one
    /// would let a long run exceed its own limit by forgetting work it
    /// had already done — and a `.per_scope` bound never entered here.
    pub fn release(self: *WorkCounters, cat: WorkCategory, n: u64) void {
        assert(cat.kind() == .depth);
        const i = @intFromEnum(cat);
        assert(self.used[i] >= n);
        self.used[i] -= n;
    }
};

pub const WorkBudgetError = error{ LimitExceeded, Cancelled };

/// M9c1's shared work budget — the one meter the evaluator and every
/// solver draw on together (§5.5's work-unit bound, the ladder's
/// "combined exhaustion").
///
/// It is a *fourth* mechanism beside `Budget` and the two `WorkLimits`
/// kinds, and exists because none of the three says what a solver
/// costs: bytes are counted by the allocator wrapper, the graph
/// counters are charged at plan admission, and a Newton iteration
/// allocates nothing and is not a cell. The units are pinned here —
/// **an evaluated AST node is 1, a solver iteration is 4** — so
/// "evaluator work" and "solver work" are commensurable, and a nested
/// callback re-charges by construction: whatever an implementation
/// evaluates from inside a solve arrives back at `evalNode`, which
/// charges this same meter again.
///
/// Two fields are policy, not identity. `limit` changes what a run
/// answers (an exhausted budget is a refusal), so a caller that sets it
/// owns the consequences; `poller` is §5.5 cancellation and stays
/// outside identity for the same reason `RunInputs.cancel` does. The
/// stride guarantees the §5.5 bound — at least one poll per 65 536
/// work units — from *inside* the engine, which until this row polled
/// nowhere: the per-cell poll belongs to the recalc driver, and a
/// 128-iteration solve inside one cell is exactly the stretch that
/// driver cannot see into.
///
/// **Nothing is mutated on a refusal**: the limit is checked before the
/// counter moves, like `Budget` and `WorkCounters`.
pub const WorkBudget = struct {
    limit: u64 = default_limit,
    used: u64 = 0,
    /// Set on the charge the limit refused, never cleared: the caller's
    /// error type says a limit was hit, this says it was THIS meter.
    tripped: bool = false,
    /// §5.5's erased poll seam. `.none` costs nothing and polls nowhere,
    /// which keeps the uncontrolled path byte-identical to the
    /// controlled one under a disarmed control (M5d1's construction).
    poller: Poller = .none,
    since_poll: u64 = 0,

    /// One evaluated AST node.
    pub const node_units: u64 = 1;
    /// One solver iteration (§ M9c1: a residual + derivative evaluation
    /// and the step arithmetic — a few dozen flops, priced at four
    /// nodes so a full 128-iteration solve is 512 units, visibly more
    /// than the formula around it and still far under one poll stride).
    pub const solver_iteration_units: u64 = 4;
    /// §5.5: at least one poll per this many work units.
    pub const poll_stride: u64 = 65_536;
    /// 2^33 units. At the §9 default cell ceiling (50 M admitted
    /// evaluations) this is 171 units per cell — room for a formula of
    /// a hundred nodes AND a dozen solver iterations in every single
    /// cell of a maximal plan — so out of the box the other limits
    /// bind first and this one is the backstop for solver work in
    /// volume, which is what it exists to bound.
    pub const default_limit: u64 = 1 << 33;
    /// Same ceiling-on-the-ceiling rule as every §9 limit.
    pub const hard_multiplier: u64 = 4;

    pub fn remaining(self: WorkBudget) u64 {
        return if (self.used >= self.limit) 0 else self.limit - self.used;
    }

    /// Charge `n` units, then poll if a stride boundary was crossed.
    /// The poll sits *after* the charge: the units were spent by the
    /// work that just happened, and a cancellation is observed at the
    /// next poll point (§5.5's "observed before commit" — the commit
    /// point is the run's, not the counter's).
    pub fn charge(self: *WorkBudget, n: u64) WorkBudgetError!void {
        const next = std.math.add(u64, self.used, n) catch {
            self.tripped = true;
            return error.LimitExceeded;
        };
        if (next > self.limit) {
            self.tripped = true;
            return error.LimitExceeded;
        }
        self.used = next;
        self.since_poll += n;
        if (self.since_poll >= poll_stride) {
            self.since_poll = 0;
            try self.poll();
        }
    }

    /// An explicit poll point — what a solver loop calls once per
    /// iteration, before the iteration's work, so a token set mid-solve
    /// is observed within one iteration rather than within one stride.
    pub fn poll(self: *WorkBudget) error{Cancelled}!void {
        return self.poller.check();
    }
};

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

fn defaultInputs() RunInputs {
    return .{ .now_utc_ms = 0, .rng_seed = 0, .limits = .{} };
}

test "run inputs: the two required fields have no defaults" {
    // A default here would mean the library reads a clock or an entropy
    // source, and determinism would stop being checkable.
    const required = [_][]const u8{ "now_utc_ms", "rng_seed", "limits" };
    var found: usize = 0;
    inline for (required) |want| {
        inline for (@typeInfo(RunInputs).@"struct".fields) |f| {
            if (comptime std.mem.eql(u8, f.name, want)) {
                found += 1;
                try testing.expect(f.default_value_ptr == null);
            }
        }
    }
    try testing.expectEqual(required.len, found);

    // And everything else has one, so a caller supplies exactly three.
    inline for (@typeInfo(RunInputs).@"struct".fields) |f| {
        const is_required = comptime for (required) |r| {
            if (std.mem.eql(u8, f.name, r)) break true;
        } else false;
        if (!is_required) try testing.expect(f.default_value_ptr != null);
    }
}

test "run inputs: the UTC offset is validated pre-narrowing" {
    var ri = defaultInputs();
    try ri.validate();

    ri.utc_offset_min = 1440;
    try ri.validate();
    ri.utc_offset_min = -1440;
    try ri.validate();

    ri.utc_offset_min = 1441;
    try testing.expectError(error.UtcOffsetOutOfRange, ri.validate());
    ri.utc_offset_min = -1441;
    try testing.expectError(error.UtcOffsetOutOfRange, ri.validate());
}

test "effective inputs exclude the two fields §5.5 excludes" {
    // Absent by construction: `deadline` and `cancel` are not fields of
    // `EffectiveRunInputs`, so no filter can forget them.
    inline for (@typeInfo(EffectiveRunInputs).@"struct".fields) |f| {
        try testing.expect(!std.mem.eql(u8, f.name, "deadline"));
        try testing.expect(!std.mem.eql(u8, f.name, "cancel"));
    }

    var flag: u8 = 0;
    var patient = defaultInputs();
    var impatient = defaultInputs();
    impatient.deadline = .{ .nanoseconds = 1 };
    impatient.cancel = .{ .flag = &flag };

    // Two runs differing only in how long the caller would wait are the
    // same run as far as identity is concerned.
    try testing.expect(patient.effective(.standalone_eval).eql(impatient.effective(.standalone_eval)));

    // The cancel seam still works — it is just not part of identity.
    try testing.expect(!impatient.isCancelled());
    flag = 1;
    try testing.expect(impatient.isCancelled());
    var atomic = std.atomic.Value(bool).init(true);
    patient.cancel = .{ .atomic = &atomic };
    try testing.expect(patient.isCancelled());
}

test "effective inputs: recalc normalizes the dialect away, standalone eval keys on it" {
    var da = defaultInputs();
    da.dialect = .dynamic_array;
    var legacy = defaultInputs();
    legacy.dialect = .legacy;

    // Standalone eval takes the caller's dialect, so it is part of the key.
    try testing.expect(!da.effective(.standalone_eval).eql(legacy.effective(.standalone_eval)));
    // A recalc derives it per stored cell, so keying on it would be a
    // phantom: identical work, two cache entries.
    try testing.expect(da.effective(.recalc).eql(legacy.effective(.recalc)));
    try testing.expectEqual(@as(?value.Dialect, null), da.effective(.recalc).dialect);
}

test "limits: defaults are §9's, and the hard maximum is 4x" {
    const l: ResourceLimits = .{};
    try testing.expectEqual(@as(u64, 1 << 30), l.max_run_arena_bytes);
    try testing.expectEqual(@as(u64, 8_000_000), l.max_matrix_cells);
    try testing.expectEqual(@as(u64, 256 << 20), l.max_string_payload_bytes);
    try testing.expectEqual(@as(u64, 128 << 20), l.max_retained_ast_bytes);
    try testing.expectEqual(@as(u64, 1 << 20), l.max_diagnostics_bytes);
    try l.validate();

    inline for (@typeInfo(Category).@"enum".fields) |f| {
        const cat: Category = @enumFromInt(f.value);
        var raised = l;
        raised.set(cat, ResourceLimits.hardMaxFor(cat));
        try raised.validate();

        var over = l;
        over.set(cat, ResourceLimits.hardMaxFor(cat) + 1);
        try testing.expectError(error.LimitOutOfRange, over.validate());

        // Zero is not "unlimited"; it is a limit nothing can satisfy.
        var zero = l;
        zero.set(cat, 0);
        try testing.expectError(error.LimitOutOfRange, zero.validate());
    }
}

test "budget: every category refuses at its own boundary, below, at, and above" {
    inline for (@typeInfo(Category).@"enum".fields) |f| {
        const cat: Category = @enumFromInt(f.value);
        var limits: ResourceLimits = .{};
        limits.set(cat, 1024);

        // Below the limit.
        var b = Budget.init(testing.allocator, limits);
        try b.charge(cat, 1023);
        try testing.expectEqual(@as(?Category, null), b.tripped);
        // Exactly at it.
        try b.charge(cat, 1);
        try testing.expectEqual(@as(u64, 1024), b.usedBy(cat));
        try testing.expectEqual(@as(?Category, null), b.tripped);
        // One past it.
        try testing.expectError(error.LimitExceeded, b.charge(cat, 1));
        try testing.expectEqual(@as(?Category, cat), b.tripped);
        // Zero mutation on refusal: the counter is where it was.
        try testing.expectEqual(@as(u64, 1024), b.usedBy(cat));
        try testing.expectEqual(@as(u64, 0), b.remaining(cat));
    }
}

test "budget: an allocator charges its own category and nothing else" {
    var limits: ResourceLimits = .{};
    limits.set(.string_payload, 4096);
    var b = Budget.init(testing.allocator, limits);

    const strings = b.allocator(.string_payload);
    const arena = b.allocator(.run_arena);

    const s = try strings.alloc(u8, 1000);
    defer strings.free(s);
    try testing.expectEqual(@as(u64, 1000), b.usedBy(.string_payload));
    try testing.expectEqual(@as(u64, 0), b.usedBy(.run_arena));

    const a = try arena.alloc(u8, 64);
    defer arena.free(a);
    try testing.expectEqual(@as(u64, 64), b.usedBy(.run_arena));
    try testing.expectEqual(@as(u64, 1000), b.usedBy(.string_payload));

    // Past the string budget: OutOfMemory from the allocator, with the
    // category recorded so the caller can tell a refusal from a machine
    // that ran out of memory.
    try testing.expectError(error.OutOfMemory, strings.alloc(u8, 4000));
    try testing.expectEqual(@as(?Category, .string_payload), b.tripped);
    // And nothing leaked into the counter on the way out.
    try testing.expectEqual(@as(u64, 1000), b.usedBy(.string_payload));
}

test "budget: freeing releases the charge, and the peak remembers" {
    var b = Budget.init(testing.allocator, .{});
    const a = b.allocator(.run_arena);

    const first = try a.alloc(u8, 8192);
    try testing.expectEqual(@as(u64, 8192), b.usedBy(.run_arena));
    a.free(first);
    try testing.expectEqual(@as(u64, 0), b.usedBy(.run_arena));
    // Live usage falls; the high-water mark does not, because that is
    // the number a caller sizing a limit actually needs.
    try testing.expectEqual(@as(u64, 8192), b.peakOf(.run_arena));

    const grown = try a.alloc(u8, 16);
    defer a.free(grown);
    try testing.expectEqual(@as(u64, 16), b.usedBy(.run_arena));
}

test "budget: resize and remap charge only the delta" {
    var b = Budget.init(testing.allocator, .{});
    const a = b.allocator(.run_arena);

    var list: std.ArrayListUnmanaged(u8) = .empty;
    defer list.deinit(a);
    try list.appendSlice(a, "hello");
    const after_first = b.usedBy(.run_arena);
    try testing.expect(after_first >= 5);

    try list.appendSlice(a, "x" ** 4096);
    try testing.expect(b.usedBy(.run_arena) >= 4096);
    // Whatever the growth strategy did, the counter tracks live bytes
    // rather than the sum of every request ever made.
    try testing.expect(b.usedBy(.run_arena) <= list.capacity + 64);
}

test "budget: an arena over a budget still reports the category that tripped" {
    var limits: ResourceLimits = .{};
    limits.set(.run_arena, 4096);
    var b = Budget.init(testing.allocator, limits);

    var arena = std.heap.ArenaAllocator.init(b.allocator(.run_arena));
    defer arena.deinit();
    const a = arena.allocator();

    // Small allocations succeed; a request past the budget does not, and
    // the failure is attributable rather than an anonymous OOM.
    _ = try a.alloc(u8, 64);
    try testing.expectError(error.OutOfMemory, a.alloc(u8, 1 << 20));
    try testing.expectEqual(@as(?Category, .run_arena), b.tripped);
}

test "checkAllAllocationFailures: the budget leaks nothing under OOM" {
    const H = struct {
        fn run(allocator: std.mem.Allocator) !void {
            var b = Budget.init(allocator, .{});
            const a = b.allocator(.run_arena);
            var list: std.ArrayListUnmanaged(u32) = .empty;
            defer list.deinit(a);
            for (0..64) |i| try list.append(a, @intCast(i));
            const s = try a.alloc(u8, 128);
            a.free(s);
            // A failed run must still leave the counter self-consistent.
            if (b.usedBy(.run_arena) < 64 * @sizeOf(u32)) return error.Undercounted;
        }
    };
    try testing.checkAllAllocationFailures(testing.allocator, H.run, .{});
}

test "work limits: defaults are §9's, and the hard maximum is 4× unless stated" {
    const l: WorkLimits = .{};
    try testing.expectEqual(@as(u64, 50_000_000), l.max_dependency_edges);
    try testing.expectEqual(@as(u64, 50_000_000), l.max_total_cell_evals);
    try testing.expectEqual(@as(u64, 512), l.max_eval_depth);
    try testing.expectEqual(@as(u64, 32_767), l.max_scc_iterations);
    try testing.expectEqual(@as(u64, 3), l.max_dynamic_passes);

    inline for (@typeInfo(WorkCategory).@"enum".fields) |f| {
        const cat: WorkCategory = @enumFromInt(f.value);
        try testing.expectEqual(l.get(cat), WorkLimits.defaultFor(cat));
        // §9's rule is "4× unless stated", and both M5a2 states one.
        // Spelled as a switch rather than as an exception list so a
        // sixth category has to answer the question rather than inherit
        // an answer.
        const want: u64 = switch (cat) {
            .scc_iterations => 32_767,
            .dynamic_passes => 10,
            else => l.get(cat) * 4,
        };
        try testing.expectEqual(want, WorkLimits.hardMaxFor(cat));
    }

    // `scc_iterations` is the one category whose hard maximum IS its
    // default: 32 767 is the largest `iterateCount` a workbook can carry,
    // so there is nothing above it left to permit.
    try testing.expectEqual(
        WorkLimits.defaultFor(.scc_iterations),
        WorkLimits.hardMaxFor(.scc_iterations),
    );
}

test "work limits: zero and over-hard-max are both out of range" {
    inline for (@typeInfo(WorkCategory).@"enum".fields) |f| {
        const cat: WorkCategory = @enumFromInt(f.value);
        var l: WorkLimits = .{};
        try l.validate();

        l.set(cat, 0);
        try testing.expectError(error.LimitOutOfRange, l.validate());

        l.set(cat, WorkLimits.hardMaxFor(cat));
        try l.validate();

        l.set(cat, WorkLimits.hardMaxFor(cat) + 1);
        try testing.expectError(error.LimitOutOfRange, l.validate());
    }
}

test "work counters: below, at, and above — per accumulating category" {
    inline for (@typeInfo(WorkCategory).@"enum".fields) |f| {
        const cat: WorkCategory = @enumFromInt(f.value);
        // A per-scope bound is read by the engine that owns the scope,
        // never charged here. Its own below/at/above boundary tests live
        // with that engine, where the scope exists to be counted.
        if (comptime cat.kind() != .per_scope) {
            var c: WorkCounters = .{};
            c.limits.set(cat, 3);

            try c.charge(cat, 2); // below
            try testing.expectEqual(@as(?WorkCategory, null), c.tripped);
            try c.charge(cat, 1); // at
            try testing.expectEqual(@as(u64, 3), c.usedBy(cat));
            try testing.expectEqual(@as(?WorkCategory, null), c.tripped);

            // Above. The counter does not move, so a refusal is
            // observably non-mutating even in its own bookkeeping.
            try testing.expectError(error.LimitExceeded, c.charge(cat, 1));
            try testing.expectEqual(@as(u64, 3), c.usedBy(cat));
            try testing.expectEqual(@as(?WorkCategory, cat), c.tripped);
        }
    }
}

test "work counters: an overflowing charge trips rather than wrapping" {
    var c: WorkCounters = .{};
    c.limits.set(.dependency_edges, std.math.maxInt(u64));
    try c.charge(.dependency_edges, std.math.maxInt(u64) - 1);
    try testing.expectError(error.LimitExceeded, c.charge(.dependency_edges, 4));
    try testing.expectEqual(@as(u64, std.math.maxInt(u64) - 1), c.usedBy(.dependency_edges));
}

test "work limits: every category declares a kind, and the two M5a2 adds are per-scope" {
    // The classification is what keeps §5.6c's "each SCC gets its own
    // pass counter" from being a comment: a per-scope bound that could
    // be charged into `WorkCounters` would accumulate across components
    // and refuse a workbook whose every component iterated legally.
    var totals: usize = 0;
    var depths: usize = 0;
    var scoped: usize = 0;
    inline for (@typeInfo(WorkCategory).@"enum".fields) |f| {
        const cat: WorkCategory = @enumFromInt(f.value);
        switch (cat.kind()) {
            .total => totals += 1,
            .depth => depths += 1,
            .per_scope => scoped += 1,
        }
        // Every category names its unit, including the two that are not
        // counters — a §9 report prints all five.
        try testing.expect(cat.unit().len > 0);
    }
    try testing.expectEqual(@as(usize, 2), totals);
    try testing.expectEqual(@as(usize, 1), depths);
    try testing.expectEqual(@as(usize, 2), scoped);
    try testing.expectEqual(WorkCategory.Kind.per_scope, WorkCategory.scc_iterations.kind());
    try testing.expectEqual(WorkCategory.Kind.per_scope, WorkCategory.dynamic_passes.kind());
}

test "work counters: depth unwinds, peak does not" {
    var c: WorkCounters = .{};
    try c.charge(.eval_depth, 1);
    try c.charge(.eval_depth, 1);
    try testing.expectEqual(@as(u64, 2), c.usedBy(.eval_depth));
    c.release(.eval_depth, 2);
    try testing.expectEqual(@as(u64, 0), c.usedBy(.eval_depth));
    try testing.expectEqual(@as(u64, 2), c.peakOf(.eval_depth));
}

test "work budget: an exhausting charge refuses with nothing mutated" {
    var w: WorkBudget = .{ .limit = 10 };
    try w.charge(7);
    try testing.expectEqual(@as(u64, 7), w.used);
    try testing.expectEqual(@as(u64, 3), w.remaining());
    // The refused charge leaves the counter exactly where it was —
    // `Budget`'s and `WorkCounters`' rule, held here too.
    try testing.expectError(error.LimitExceeded, w.charge(4));
    try testing.expectEqual(@as(u64, 7), w.used);
    try testing.expect(w.tripped);
    // The meter still answers what fits.
    try w.charge(3);
    try testing.expectEqual(@as(u64, 0), w.remaining());
}

test "work budget: the poll stride fires at least once per 65 536 units" {
    const Counter = struct {
        n: usize = 0,
        fn check(ctx: ?*const anyopaque) error{Cancelled}!void {
            const self: *const @This() = @ptrCast(@alignCast(ctx.?));
            @constCast(self).n += 1;
        }
    };
    var counter: Counter = .{};
    var w: WorkBudget = .{ .poller = .{ .ctx = &counter, .check_fn = Counter.check } };

    // Three strides and a remainder, charged one node at a time: the
    // §5.5 bound is a property of accumulation, not of any one charge.
    var i: u64 = 0;
    while (i < 3 * WorkBudget.poll_stride + 7) : (i += 1) {
        try w.charge(WorkBudget.node_units);
    }
    try testing.expectEqual(@as(usize, 3), counter.n);

    // …and a single bulk charge that crosses a boundary polls too.
    try w.charge(WorkBudget.poll_stride);
    try testing.expectEqual(@as(usize, 4), counter.n);
}

test "work budget: a tripped poller surfaces Cancelled, distinct from exhaustion" {
    const Trip = struct {
        fn check(_: ?*const anyopaque) error{Cancelled}!void {
            return error.Cancelled;
        }
    };
    var w: WorkBudget = .{ .poller = .{ .ctx = null, .check_fn = Trip.check } };
    // An explicit poll point sees the token immediately…
    try testing.expectError(error.Cancelled, w.poll());
    // …and a stride-crossing charge sees it at the boundary, with the
    // units already spent — the work happened, the run did not commit.
    try testing.expectError(error.Cancelled, w.charge(WorkBudget.poll_stride));
    try testing.expectEqual(WorkBudget.poll_stride, w.used);
    try testing.expect(!w.tripped);
}
