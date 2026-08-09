//! The dependency graph: what a workbook's formulas depend on, in what
//! order that lets them be evaluated, and what a cycle's members start
//! from (`goal_formula.md` §5.6a–c, §5.6f, §5.6g, §9).
//!
//! M5a1 of the tier-D1 ladder. This file **builds and orders** the
//! graph; it does not run it. The iteration engine — multi-SCC
//! execution, convergence, the clamps, the callsite-keyed volatile
//! schedule, the dynamic-edge fixpoint — is M5a2's, and the seed table
//! here exists so that engine has something to start from.
//!
//! A missed edge is the failure mode
//! ---------------------------------
//! Every other kind of bug in this file announces itself. A missed edge
//! does not: the graph still builds, still orders, still passes every
//! performance test, and then writes a cache computed from a stale
//! input. That is why the gate for this row is a randomized differential
//! test against an independently written brute-force builder rather than
//! a fixture suite, and why the node model below is written down as a
//! closed enumeration instead of growing case by case.
//!
//! What is a node (§5.6b)
//! ----------------------
//! | kind | exists for | depends on |
//! |---|---|---|
//! | `cell` | every formula-bearing coordinate | what its body mentions |
//! | `spill_tail` | every non-anchor coordinate of a declared array shape | its anchor |
//! | `range` | every distinct area any body mentions | every producer coordinate inside it |
//! | `span` | every distinct 3D sheet-span reference | one node per member sheet (§5.6g) |
//! | `name` | every defined name | what its body mentions |
//! | `producer` | every table calculated column / totals row | what its body mentions |
//!
//! A reference to a coordinate that holds no formula contributes **no
//! edge**: a constant cannot be recalculated, so it cannot constrain an
//! order. That is the one rule the brute-force builder has to share, and
//! it is stated here rather than discovered from the code.
//!
//! Range nodes are the reason this is not O(F×R)
//! ---------------------------------------------
//! Two formulas reading `A1:A100` share one range node, so the area is
//! resolved against the producer index once rather than once per reader.
//! The index itself is sorted twice — row-major and column-major — and
//! each area is probed through whichever of the two has **fewer stored
//! coordinates in the band it would walk**, counted rather than guessed
//! from the area's extent (M5d4; the extent is right for `SUM(A:A)` and
//! wrong for `SUM(A5:A9)`, which is short in the dimension the extent
//! calls narrow and deep in the one it walks). That is what keeps
//! `SUM(A:A)` proportional to the stored cells in column A instead of to
//! 1 048 576 coordinates, and `stats.index_probes` is the instrument: a
//! counter, not a stopwatch.
//!
//! Purity
//! ------
//! Nothing here mutates a workbook. The builder reads decoded formula
//! text and coordinates, allocates into its own arena, and returns a
//! value. Refusals happen before any of it is handed back, which is what
//! makes "malformed `<v>` never seeds a zero" a statement about the
//! whole run rather than about one function's return.
//!
//! Lifetime
//! --------
//! A `Key` borrows its name, table and column spellings from the `Input`
//! it was built from, so the `Input` — and whatever owns the strings in
//! it — must outlive the `Graph`. Formula *bodies* need not: they are
//! parsed into a scratch arena that is released before the graph is
//! returned, because every reference a walk keeps is either a coordinate
//! or one of those borrowed spellings.

const std = @import("std");
const assert = std.debug.assert;

const coords = @import("zlsx_refs");
const env = @import("env.zig");
const eval = @import("eval.zig");
const name_rules = @import("names.zig");
const parser = @import("parser.zig");
const run_inputs = @import("run_inputs.zig");
const value = @import("value.zig");

pub const WorkLimits = run_inputs.WorkLimits;
pub const WorkCounters = run_inputs.WorkCounters;
pub const WorkCategory = run_inputs.WorkCategory;

/// Everything the builder and its resolver can fail with, separately
/// from the things they *refuse*. Same split the rest of the engine
/// keeps: an `Error` is about this machine, a `Refusal` is about the
/// workbook.
pub const Error = error{OutOfMemory};

// ─── the node model (§5.6b) ──────────────────────────────────────

pub const Kind = enum(u3) {
    cell = 0,
    spill_tail = 1,
    range = 2,
    span = 3,
    name = 4,
    producer = 5,
};

/// A node's identity, and the whole basis of the deterministic order.
///
/// Every field that participates in `order` is **content**: a
/// coordinate, a sheet index, a byte string. Nothing is an insertion
/// ordinal or a pointer, which is what makes "the same workbook yields
/// the same topological order across randomized insertion order" a
/// property of the type rather than of the builder's care.
pub const Key = union(Kind) {
    cell: env.CellRef,
    spill_tail: env.CellRef,
    range: env.RangeRef,
    span: Span,
    name: Name,
    producer: Producer,

    /// §5.6g's inclusive sheet span, plus the area it takes on each
    /// member sheet.
    pub const Span = struct {
        first: env.SheetIndex,
        last: env.SheetIndex,
        range: coords.Range,
    };

    /// `index` addresses `Input.names` and is deliberately **not** part
    /// of the order: a workbook cannot hold two names with the same
    /// scope and spelling, so (scope, identifier) already identifies
    /// one, and including the row number would let input order leak
    /// into the topological order.
    pub const Name = struct {
        scope: ?env.SheetIndex,
        identifier: []const u8,
        index: u32,
    };

    /// Same rule: (table, column, kind) identifies a producer, and
    /// `index` addresses `Input.producers` without entering the order.
    pub const Producer = struct {
        table: []const u8,
        column: []const u8,
        kind: name_rules.ProducerKind,
        index: u32,
    };

    pub fn kind(self: Key) Kind {
        return std.meta.activeTag(self);
    }

    /// §5.6b's tie-break, generalized to every node kind.
    ///
    /// Cells and spill tails order by (SheetIndex, Row1, Col0) exactly
    /// as §5.6b writes it. The other kinds need *an* order rather than a
    /// particular one, so they take the obvious content order and the
    /// kind tag decides between kinds. Kahn's algorithm picks the
    /// smallest ready component under this comparison, which makes the
    /// emitted order canonical — not merely reproducible.
    pub fn order(a: Key, b: Key) std.math.Order {
        const ka = @intFromEnum(a.kind());
        const kb = @intFromEnum(b.kind());
        if (ka != kb) return std.math.order(ka, kb);
        return switch (a) {
            .cell => |x| orderCell(x, b.cell),
            .spill_tail => |x| orderCell(x, b.spill_tail),
            .range => |x| orderArea(x, b.range),
            .span => |x| blk: {
                const y = b.span;
                const f = std.math.order(x.first.toInt(), y.first.toInt());
                if (f != .eq) break :blk f;
                const l = std.math.order(x.last.toInt(), y.last.toInt());
                if (l != .eq) break :blk l;
                break :blk orderRange(x.range, y.range);
            },
            .name => |x| blk: {
                const y = b.name;
                const s = std.math.order(scopeOrd(x.scope), scopeOrd(y.scope));
                if (s != .eq) break :blk s;
                break :blk std.mem.order(u8, x.identifier, y.identifier);
            },
            .producer => |x| blk: {
                const y = b.producer;
                const t = std.mem.order(u8, x.table, y.table);
                if (t != .eq) break :blk t;
                const c = std.mem.order(u8, x.column, y.column);
                if (c != .eq) break :blk c;
                break :blk std.math.order(@intFromEnum(x.kind), @intFromEnum(y.kind));
            },
        };
    }

    pub fn eql(a: Key, b: Key) bool {
        return a.order(b) == .eq;
    }

    /// The hash `eql` implies, for the callers that need a membership
    /// test rather than an order.
    ///
    /// It mirrors `order` field for field, which is what makes it a
    /// *correct* hash rather than merely a plausible one: `index` is
    /// absent from both, so two rows addressing the same name neither
    /// compare unequal nor land in different buckets. Every
    /// variable-length field is length-prefixed, so a `producer` cannot
    /// collide with one whose table and column split the same bytes at
    /// a different point.
    pub fn hash(self: Key, h: *std.hash.Wyhash) void {
        h.update(&[_]u8{@intFromEnum(self.kind())});
        switch (self) {
            .cell, .spill_tail => |c| hashCell(h, c),
            .range => |r| hashArea(h, r),
            .span => |x| {
                hashInt(h, x.first.toInt());
                hashInt(h, x.last.toInt());
                hashRange(h, x.range);
            },
            .name => |x| {
                hashInt(h, scopeOrd(x.scope));
                hashBytes(h, x.identifier);
            },
            .producer => |x| {
                hashBytes(h, x.table);
                hashBytes(h, x.column);
                h.update(&[_]u8{@intFromEnum(x.kind)});
            },
        }
    }

    /// `std.HashMapUnmanaged` context over `order`'s equality.
    pub const HashContext = struct {
        pub fn hash(_: HashContext, x: Key) u64 {
            var h: std.hash.Wyhash = .init(0);
            x.hash(&h);
            return h.final();
        }
        pub fn eql(_: HashContext, x: Key, y: Key) bool {
            return x.eql(y);
        }
    };

    /// Whether the node occupies a coordinate — the two kinds the
    /// producer index holds, and the two `max_eval_depth` counts.
    pub fn isCellLike(self: Key) bool {
        return switch (self) {
            .cell, .spill_tail => true,
            else => false,
        };
    }

    pub fn coordinate(self: Key) ?env.CellRef {
        return switch (self) {
            .cell, .spill_tail => |c| c,
            else => null,
        };
    }
};

/// Workbook scope sorts before every sheet scope. An arbitrary choice,
/// but a *stated* one — the alternative is a comparison that depends on
/// how `?SheetIndex` happens to be laid out.
fn scopeOrd(s: ?env.SheetIndex) u64 {
    return if (s) |x| @as(u64, x.toInt()) + 1 else 0;
}

fn orderCell(a: env.CellRef, b: env.CellRef) std.math.Order {
    const s = std.math.order(a.sheet.toInt(), b.sheet.toInt());
    if (s != .eq) return s;
    const r = std.math.order(a.row.oneBased(), b.row.oneBased());
    if (r != .eq) return r;
    return std.math.order(a.col.zeroBased(), b.col.zeroBased());
}

fn orderArea(a: env.RangeRef, b: env.RangeRef) std.math.Order {
    const s = std.math.order(a.sheet.toInt(), b.sheet.toInt());
    if (s != .eq) return s;
    return orderRange(a.range, b.range);
}

fn orderRange(a: coords.Range, b: coords.Range) std.math.Order {
    const fr = std.math.order(a.first.row.oneBased(), b.first.row.oneBased());
    if (fr != .eq) return fr;
    const fc = std.math.order(a.first.col.zeroBased(), b.first.col.zeroBased());
    if (fc != .eq) return fc;
    const lr = std.math.order(a.last.row.oneBased(), b.last.row.oneBased());
    if (lr != .eq) return lr;
    return std.math.order(a.last.col.zeroBased(), b.last.col.zeroBased());
}

fn hashInt(h: *std.hash.Wyhash, x: anytype) void {
    h.update(std.mem.asBytes(&@as(u64, x)));
}

fn hashBytes(h: *std.hash.Wyhash, b: []const u8) void {
    hashInt(h, b.len);
    h.update(b);
}

fn hashCell(h: *std.hash.Wyhash, c: env.CellRef) void {
    hashInt(h, c.sheet.toInt());
    hashInt(h, c.row.oneBased());
    hashInt(h, c.col.zeroBased());
}

fn hashArea(h: *std.hash.Wyhash, r: env.RangeRef) void {
    hashInt(h, r.sheet.toInt());
    hashRange(h, r.range);
}

fn hashRange(h: *std.hash.Wyhash, r: coords.Range) void {
    hashInt(h, r.first.row.oneBased());
    hashInt(h, r.first.col.zeroBased());
    hashInt(h, r.last.row.oneBased());
    hashInt(h, r.last.col.zeroBased());
}

// ─── the SCC seed table (§5.6c) ──────────────────────────────────

/// What a formula cell's `<v>` was, as the seed table sees it. The text
/// and boolean payloads are absent on purpose: every one of them seeds
/// zero, so carrying the value would only invite a reader to think it
/// mattered.
pub const CacheState = union(enum) {
    /// A numeric `<v>` that parsed.
    number: f64,
    text,
    boolean,
    err,
    /// No `<v>` element at all — an uncached formula.
    absent,
    /// A `<v>` that is present but unparseable for its declared type.
    /// The decode boundary already refuses these before a model exists
    /// (`decode.classifyCell`); the row is here so the seed table is
    /// complete and so a caller that builds an `Input` by hand cannot
    /// quietly turn one into a zero.
    malformed,
};

/// Whether a cell places an array, and whether its shape is knowable
/// without evaluating it.
pub const Anchor = union(enum) {
    /// Not an array anchor.
    none,
    /// A legacy CSE `ref`, or a dynamic-array anchor's metadata shape.
    shape: value.Shape,
    /// A dynamic-array anchor with no recoverable shape — a newly
    /// authored formula whose result has never been placed.
    unknown,
};

/// What an SCC member starts an iteration from.
pub const Seed = union(enum) {
    number: f64,
    /// A zero-filled array of exactly this shape.
    array_zeros: value.Shape,
    /// An array anchor with no recoverable shape. §5.6c's pre-iteration
    /// shape pass evaluates it once outside the cycle to fix the shape;
    /// running that pass is M5a2's, naming its outcome is this row's.
    shape_pass,
};

pub const SeedError = error{
    /// §5.6c: a malformed `<v>` is a pre-mutation typed refusal, never a
    /// zero seed. A malformed cache cannot be both "unparseable input"
    /// and "the number zero".
    MalformedCache,
};

/// §5.6c's seed table, all five rows, as one function.
///
/// The malformed check comes **first**, before the array-anchor row.
/// An anchor with a declared shape never reads its cache, so ordering it
/// the other way would let an anchor launder a malformed `<v>` into a
/// zero-filled array — the exact substitution the rule forbids.
pub fn seedFor(cache: CacheState, anchor: Anchor) SeedError!Seed {
    if (cache == .malformed) return error.MalformedCache;
    switch (anchor) {
        .shape => |s| return .{ .array_zeros = s },
        .unknown => return .shape_pass,
        .none => {},
    }
    return switch (cache) {
        .number => |n| .{ .number = n },
        // Text, booleans and error values all seed zero: iteration is
        // numeric, and a text cache carries no number to resume from.
        .text, .boolean, .err, .absent => .{ .number = 0 },
        .malformed => unreachable,
    };
}

// ─── the input model ─────────────────────────────────────────────

pub const CellInput = struct {
    cell: env.CellRef,
    /// Decoded `<f>` body (FORMULA carrier — XML entities only).
    formula: []const u8,
    cache: CacheState = .absent,
    /// The declared range of a legacy CSE array or the metadata shape of
    /// a dynamic-array anchor. `cell` must be its top-left; a range that
    /// says otherwise is a refusal, not a silent re-anchoring.
    array: ?coords.Range = null,
    /// The cell places an array but its shape is not recoverable from
    /// the file. Only meaningful with `array == null`: a declared range
    /// *is* a recovered shape. §5.6c sends these through the
    /// pre-iteration shape pass rather than seeding them.
    dynamic_anchor: bool = false,
};

pub const NameInput = struct {
    identifier: []const u8,
    /// Decoded body (FORMULA carrier).
    body: []const u8,
    /// Null for a workbook-scoped name.
    scope: ?env.SheetIndex = null,
};

pub const ProducerInput = struct {
    table: []const u8,
    column: []const u8,
    kind: name_rules.ProducerKind,
    body: []const u8,
    /// Where the producer materializes. Its sheet is also the sheet its
    /// body's unqualified references resolve against.
    span: env.RangeRef,
};

/// One workbook, as the graph builder sees it.
///
/// Plain data rather than a vtable, because unlike `EvalEnv` there is
/// nothing here to compute lazily — the builder reads all of it, once.
/// The one thing it cannot know from data is how a *spelling* binds, and
/// that is what `Resolver` is for.
pub const Input = struct {
    sheet_count: u32,
    cells: []const CellInput = &.{},
    names: []const NameInput = &.{},
    producers: []const ProducerInput = &.{},
};

// ─── spelling resolution ─────────────────────────────────────────

/// What a bare name in value position binds to, for edge purposes.
pub const NameBinding = union(enum) {
    /// A row of `Input.names`.
    name: u32,
    /// A spelling that denotes an area directly — a table, an `_xlnm.`
    /// builtin.
    area: env.RangeRef,
    /// Provably nowhere. `#NAME?` is a value outcome, and a value
    /// outcome has nothing to depend on.
    unresolved,
};

/// The three questions a static walk cannot answer from an AST alone.
///
/// A vtable for the same reason `EvalEnv` is one: the real
/// implementation lives in `pkg/`, which this file must not import, and
/// the tests here need a second one.
pub const Resolver = struct {
    ctx: *anyopaque,
    vtable: *const VTable,

    pub const VTable = struct {
        resolveSheet: *const fn (ctx: *anyopaque, name: []const u8) Error!?env.SheetIndex,
        /// `from` is the sheet the reference sits on, or null at
        /// workbook level. `spelling` is the source spelling, undecoded
        /// and unfolded — the implementation owns the comparator.
        resolveName: *const fn (
            ctx: *anyopaque,
            from: ?env.SheetIndex,
            spelling: []const u8,
        ) Error!NameBinding,
        /// A structured reference's area.
        ///
        /// The graph resolves these even though *evaluating* one is
        /// M7b's. The two are different obligations: an evaluator that
        /// refuses a construct returns an error value, while a graph
        /// that drops the edge writes a stale cache. Conservative wins.
        /// `owner` travels with it because the bare `[…]` form names
        /// *the owner's own* table — a calculated column's `[@Qty]` has
        /// no table spelling to look up — and because `#This Row` needs
        /// the owner's row.
        resolveStructured: *const fn (
            ctx: *anyopaque,
            from: ?env.SheetIndex,
            owner: Owner,
            table: ?[]const u8,
            items: parser.ItemSet,
            columns: parser.ColumnSelector,
        ) Error!?env.RangeRef,
    };

    pub fn resolveSheet(self: Resolver, name: []const u8) Error!?env.SheetIndex {
        return self.vtable.resolveSheet(self.ctx, name);
    }

    pub fn resolveName(
        self: Resolver,
        from: ?env.SheetIndex,
        spelling: []const u8,
    ) Error!NameBinding {
        return self.vtable.resolveName(self.ctx, from, spelling);
    }

    pub fn resolveStructured(
        self: Resolver,
        from: ?env.SheetIndex,
        owner: Owner,
        table: ?[]const u8,
        items: parser.ItemSet,
        columns: parser.ColumnSelector,
    ) Error!?env.RangeRef {
        return self.vtable.resolveStructured(self.ctx, from, owner, table, items, columns);
    }
};

/// A table's geometry, and the arithmetic that turns a structured
/// reference into an area.
///
/// Lives here rather than in the adapter so there is one implementation
/// of the row bands: the adapter supplies the numbers, the tests supply
/// the numbers, and both get the same answer.
pub const TableGeometry = struct {
    sheet: env.SheetIndex,
    /// The table's whole `ref`, headers and totals included.
    ref: coords.Range,
    header_rows: u32 = 0,
    totals_rows: u32 = 0,
};

/// The area a structured reference denotes, or null when the geometry
/// leaves it nothing (`#REF!` territory — a totals selector on a table
/// with no totals row, a data selector on a table with no data rows).
///
/// `first_col`/`last_col` are zero-based offsets **into the table**, not
/// sheet columns; the caller resolved the column names.
pub fn tableArea(
    g: TableGeometry,
    items: parser.ItemSet,
    first_col: u32,
    last_col: u32,
    site_row: ?coords.Row,
) ?env.RangeRef {
    const top = g.ref.first.row.oneBased();
    const bottom = g.ref.last.row.oneBased();
    if (bottom < top) return null;

    const data_first = top + g.header_rows;
    const data_last = bottom -| g.totals_rows;

    // No item specifier at all means the data rows — Excel's default,
    // and the reason `Table1[Col]` and `Table1[[#Data],[Col]]` are the
    // same reference.
    const sel: parser.ItemSet = if (items.count() == 0) .{ .data = true } else items;

    var lo: u32 = std.math.maxInt(u32);
    var hi: u32 = 0;
    var any = false;

    const Band = struct { on: bool, first: u32, last: u32 };
    const bands = [_]Band{
        .{ .on = sel.all, .first = top, .last = bottom },
        .{ .on = sel.headers, .first = top, .last = top + g.header_rows -| 1 },
        .{ .on = sel.data, .first = data_first, .last = data_last },
        .{ .on = sel.totals, .first = bottom + 1 -| g.totals_rows, .last = bottom },
        // `#This Row` with no site is a body that materializes into many
        // rows — a table producer. Its rows are, collectively, the data
        // band, so widening to it is the conservative answer rather than
        // a dropped edge.
        .{
            .on = sel.this_row,
            .first = if (site_row) |r| r.oneBased() else data_first,
            .last = if (site_row) |r| r.oneBased() else data_last,
        },
    };
    for (bands, 0..) |b, i| {
        if (!b.on) continue;
        // A band a table does not have contributes nothing rather than
        // an inverted range.
        if (i == 1 and g.header_rows == 0) continue;
        if (i == 3 and g.totals_rows == 0) continue;
        if (i == 2 and data_first > data_last) continue;
        if (b.first > b.last) continue;
        if (b.first < lo) lo = b.first;
        if (b.last > hi) hi = b.last;
        any = true;
    }
    if (!any) return null;
    // `#This Row` outside the table is `#VALUE!`, not an area.
    if (sel.this_row and site_row != null) {
        const r = site_row.?.oneBased();
        if (r < top or r > bottom) return null;
    }

    const base = g.ref.first.col.zeroBased();
    const lo_col = base + @min(first_col, last_col);
    const hi_col = base + @max(first_col, last_col);
    if (hi_col > g.ref.last.col.zeroBased()) return null;

    return .{
        .sheet = g.sheet,
        .range = (coords.Range{
            .first = .{
                .row = coords.Row.fromOneBased(lo) catch return null,
                .col = coords.Col.fromZeroBased(lo_col) catch return null,
            },
            .last = .{
                .row = coords.Row.fromOneBased(hi) catch return null,
                .col = coords.Col.fromZeroBased(hi_col) catch return null,
            },
        }).normalized(),
    };
}

// ─── refusals ────────────────────────────────────────────────────

pub const Refusal = struct {
    reason: Reason,
    /// The node whose body raised it, when there is one.
    at: ?Key = null,
    /// Set when `reason == .work_limit_exceeded`.
    limit: ?WorkCategory = null,

    pub const Reason = enum {
        /// A body that does not parse. The parser's own refusal names
        /// the construct; this only says which node carried it.
        formula_parse_failed,
        /// §5.6g's context legality — a 3D span inside an array or
        /// intersection context, or handed to an ineligible function.
        three_d_illegal_context,
        /// §5.6c's fourth row.
        malformed_cache_seed,
        /// A declared array range whose top-left is not the cell that
        /// declares it.
        array_anchor_not_top_left,
        /// §9.
        work_limit_exceeded,
        /// A cycle reached by a closure plan while iteration is off.
        /// §5.6c's `FormulaCycle`; M5a2 replaces it with an iteration
        /// schedule when `calcPr@iterate` says so.
        cycle,
    };

    /// §10's plane-2 vocabulary, so a graph refusal reaches a caller as
    /// the same kind of thing every other layer's refusal does.
    pub fn planeTwo(self: Refusal) parser.PlaneTwo {
        return switch (self.reason) {
            .formula_parse_failed, .malformed_cache_seed, .array_anchor_not_top_left => .FormulaMalformedInput,
            .three_d_illegal_context => .FormulaUnsupportedConstruct,
            .work_limit_exceeded => .FormulaLimitExceeded,
            .cycle => .FormulaCycle,
        };
    }
};

// ─── the built graph ─────────────────────────────────────────────

pub const Stats = struct {
    nodes: u32 = 0,
    edges: u64 = 0,
    /// Candidate coordinates examined while resolving areas against the
    /// producer index. §5.6a's sparseness assertion is written against
    /// this: a whole-column dependency must stay bounded by the stored
    /// cell count, and only a counter can say so without a stopwatch.
    index_probes: u64 = 0,
    components: u32 = 0,
    cyclic_components: u32 = 0,
};

pub const Graph = struct {
    arena: std.heap.ArenaAllocator,
    /// Canonical order. A node's id **is** its index here, so
    /// `keys[a].order(keys[b]) == order(a, b)` for every pair.
    keys: []const Key,
    /// `deps[u]` — the nodes `u` depends on, ascending and deduped.
    deps: []const []const u32,
    /// `component[u]` — which SCC `u` belongs to.
    component: []const u32,
    /// The condensation DAG in evaluation order: dependencies first,
    /// ties broken by each component's smallest node. Members are
    /// ascending.
    order: []const []const u32,
    /// Parallel to `order`'s index space (component id, not position).
    cyclic: []const bool,
    /// Seeds for cyclic members, indexed by node id. Null for every node
    /// in an acyclic component — an acyclic node is computed from its
    /// dependencies and has nothing to resume from.
    seeds: []const ?Seed,
    /// The producer index, kept rather than discarded with the build:
    /// a closure plan rooted at an area the workbook never mentions
    /// (`evaluate("SUM(B1:B2)")` against a workbook where nothing reads
    /// `B1:B2`) has no range node to start from and needs the same
    /// sparse lookup the builder used.
    index: Index,
    /// Per-node static walk log (M10b): for each owner node, the bounds
    /// of what the *walk* of its body noted — the pre-injection prefix
    /// of the dependency log, over the `refs` slice the arena already
    /// keeps for the graph's lifetime. Null for nodes that own no body.
    /// `iterate` probes it per runtime read to decide whether a rebuild
    /// could change anything, which is what lets a run skip the
    /// confirming rebuild instead of paying a full build to prove
    /// nothing did.
    walk_logs: []const ?WalkLog,
    stats: Stats,

    /// Bounds over an owner's captured refs, in `Capture.take`'s layout:
    /// cells first, then areas, then names, then spans — so the walk's
    /// cells are `refs[0..walk_cells]` and its areas
    /// `refs[cells .. cells + walk_areas]`. The injected runtime
    /// entries, when the graph is a rebuild's product, sit *after* each
    /// prefix, which is exactly why the prefix and not the whole list is
    /// the membership the gate needs (§5.6e's fold restates a runtime
    /// read every pass; a probe that saw injected entries would stop
    /// restating them and starve the fold).
    const WalkLog = struct {
        refs: []const Ref,
        cells: u32,
        walk_cells: u32,
        walk_areas: u32,
    };

    pub fn deinit(self: *Graph) void {
        self.arena.deinit();
        self.* = undefined;
    }

    pub fn nodeCount(self: Graph) u32 {
        return @intCast(self.keys.len);
    }

    /// The node id for a key, or null. Binary search — the keys are
    /// canonically sorted, which is the whole point of sorting them.
    pub fn find(self: Graph, key: Key) ?u32 {
        var lo: usize = 0;
        var hi: usize = self.keys.len;
        while (lo < hi) {
            const mid = lo + (hi - lo) / 2;
            switch (self.keys[mid].order(key)) {
                .eq => return @intCast(mid),
                .lt => lo = mid + 1,
                .gt => hi = mid,
            }
        }
        return null;
    }

    pub fn isCyclic(self: Graph, node: u32) bool {
        return self.cyclic[self.component[node]];
    }

    /// Whether the static walk of `owner`'s body noted exactly this
    /// target — the same membership `DependencyLog.noteCell` /
    /// `noteArea` would answer during a rebuild's injection, so "yes"
    /// means the edge would dedupe away there rather than change the
    /// graph. Exact `eql`, no containment: a cell inside a noted area
    /// is a different log entry, and the builder treats it as one.
    pub fn walkNoted(self: Graph, owner: u32, target: DynamicRef.Target) bool {
        const log = self.walk_logs[owner] orelse return false;
        switch (target) {
            .cell => |c| for (log.refs[0..log.walk_cells]) |r| {
                if (r.cell.eql(c)) return true;
            },
            .area => |x| for (log.refs[log.cells .. log.cells + log.walk_areas]) |r| {
                if (r.area.eql(x)) return true;
            },
        }
        return false;
    }

    /// The cell-like nodes inside an area, in area → sheet → row-major
    /// order (§5.6a). Sparse: the probe examines candidates, never
    /// coordinates.
    pub fn probeArea(self: Graph, area: env.RangeRef, probes: *u64) Index.Probe {
        return self.index.probe(area, probes);
    }

    /// Every node in evaluation order, flattened out of the
    /// condensation. Convenience for tests and for reporting; the
    /// component structure is what an engine actually wants.
    pub fn flatOrder(self: Graph, gpa: std.mem.Allocator) Error![]u32 {
        const out = try gpa.alloc(u32, self.keys.len);
        var i: usize = 0;
        for (self.order) |comp| {
            for (comp) |n| {
                out[i] = n;
                i += 1;
            }
        }
        assert(i == out.len);
        return out;
    }
};

// ─── building ────────────────────────────────────────────────────

pub const BuildResult = union(enum) {
    ok: Graph,
    refused: Refusal,
};

/// §5.6e's runtime-captured reference: what a body actually reached,
/// paired with the body that reached it.
///
/// A dynamic reference — `INDIRECT("A1")`, `OFFSET(A1,r,c)` — is
/// invisible to a walk over the text, so M5a2's outer loop evaluates,
/// records what was read, and builds again with these in hand. They
/// enter through the **same dependency log** the static walk fills,
/// which is what keeps a rebuilt graph identical in kind to a built one:
/// there is no second edge path, no second node model, and a captured
/// read the walk had already found dedupes away instead of doubling.
pub const DynamicRef = struct {
    /// Whose body produced it.
    owner: Key,
    target: Target,

    pub const Target = union(enum) {
        cell: env.CellRef,
        area: env.RangeRef,
    };

    pub fn eql(x: DynamicRef, y: DynamicRef) bool {
        if (!x.owner.eql(y.owner)) return false;
        if (@as(std.meta.Tag(Target), x.target) != @as(std.meta.Tag(Target), y.target)) return false;
        return switch (x.target) {
            .cell => |c| c.eql(y.target.cell),
            .area => |r| r.eql(y.target.area),
        };
    }

    pub fn hash(self: DynamicRef, h: *std.hash.Wyhash) void {
        self.owner.hash(h);
        h.update(&[_]u8{@intFromEnum(std.meta.activeTag(self.target))});
        switch (self.target) {
            .cell => |c| hashCell(h, c),
            .area => |r| hashArea(h, r),
        }
    }

    /// `std.HashMapUnmanaged` context, so a caller deduping edges gets
    /// the membership test the type already defines rather than
    /// inventing one over the same fields.
    pub const HashContext = struct {
        pub fn hash(_: HashContext, x: DynamicRef) u64 {
            var h: std.hash.Wyhash = .init(0);
            x.hash(&h);
            return h.final();
        }
        pub fn eql(_: HashContext, x: DynamicRef, y: DynamicRef) bool {
            return x.eql(y);
        }
    };
};

pub const Options = struct {
    limits: WorkLimits = .{},
    parse_limits: parser.Limits = .{},
    /// Where the §9 work counters accumulate. Optional so a caller that
    /// does not care need not thread one; when null the builder uses a
    /// local set and the resolved totals land in `Stats`.
    counters: ?*WorkCounters = null,
    /// §5.6e. Empty on a first build, by definition: nothing has run
    /// yet, so nothing has been read yet.
    dynamic_edges: []const DynamicRef = &.{},
};

pub fn build(
    gpa: std.mem.Allocator,
    input: Input,
    resolver: Resolver,
    opts: Options,
) Error!BuildResult {
    // A caller-supplied counter set carries its own limits — it is how
    // a run charges several builds against one budget — so `opts.limits`
    // configures only the local set.
    var local: WorkCounters = .{ .limits = opts.limits };
    const counters = opts.counters orelse &local;

    var b = Builder.init(gpa, input, resolver, opts, counters);
    defer b.deinit();
    return b.run();
}

/// One owner of a formula body, paired with the text to walk.
const BodyOwner = struct {
    owner: Owner,
    body: []const u8,
};

/// One captured reference, before node ids exist.
const Ref = union(enum) {
    cell: env.CellRef,
    area: env.RangeRef,
    name: Key.Name,
    span: Key.Span,
};

const Builder = struct {
    gpa: std.mem.Allocator,
    arena: std.heap.ArenaAllocator,
    input: Input,
    resolver: Resolver,
    opts: Options,
    counters: *WorkCounters,

    owners: std.ArrayListUnmanaged(BodyOwner) = .empty,
    /// `refs[i]` belongs to `owners[i]`.
    refs: std.ArrayListUnmanaged([]const Ref) = .empty,
    /// `logs[i]` bounds the walk's share of `refs[i]` (M10b): how many
    /// cells the log held in total, and how many cells/areas of those
    /// the walk itself noted before any §5.6e injection ran.
    logs: std.ArrayListUnmanaged(WalkBounds) = .empty,
    keys: std.ArrayListUnmanaged(Key) = .empty,
    /// `Input.cells` indices, sorted by coordinate. A linear scan per
    /// lookup would make the build quadratic in the cell count, which is
    /// the one asymptotic this row is not allowed to have.
    by_coord: std.ArrayListUnmanaged(u32) = .empty,
    /// Spill tails paired with the anchor that owns them, sorted by
    /// tail coordinate.
    tails: std.ArrayListUnmanaged(TailOwner) = .empty,
    stats: Stats = .{},

    const TailOwner = struct { tail: env.CellRef, anchor: env.CellRef };

    const WalkBounds = struct { cells: u32, walk_cells: u32, walk_areas: u32 };

    fn init(
        gpa: std.mem.Allocator,
        input: Input,
        resolver: Resolver,
        opts: Options,
        counters: *WorkCounters,
    ) Builder {
        return .{
            .gpa = gpa,
            .arena = std.heap.ArenaAllocator.init(gpa),
            .input = input,
            .resolver = resolver,
            .opts = opts,
            .counters = counters,
        };
    }

    fn deinit(self: *Builder) void {
        self.owners.deinit(self.gpa);
        self.refs.deinit(self.gpa);
        self.logs.deinit(self.gpa);
        self.keys.deinit(self.gpa);
        self.by_coord.deinit(self.gpa);
        self.tails.deinit(self.gpa);
        self.arena.deinit();
        self.* = undefined;
    }

    fn a(self: *Builder) std.mem.Allocator {
        return self.arena.allocator();
    }

    fn run(self: *Builder) Error!BuildResult {
        if (try self.collectOwners()) |r| return .{ .refused = r };
        if (try self.captureAll()) |r| return .{ .refused = r };
        try self.collectKeys();

        std.mem.sortUnstable(Key, self.keys.items, {}, keyLessThan);
        dedupeKeys(&self.keys);
        self.stats.nodes = @intCast(self.keys.items.len);

        const built = switch (try self.link()) {
            .ok => |g| g,
            .refused => |r| return .{ .refused = r },
        };
        return .{ .ok = built };
    }

    // ─── phase 1: who owns a formula body ────────────────────────

    fn collectOwners(self: *Builder) Error!?Refusal {
        try self.by_coord.ensureTotalCapacity(self.gpa, self.input.cells.len);
        for (0..self.input.cells.len) |i| {
            self.by_coord.appendAssumeCapacity(@intCast(i));
        }
        const cells = self.input.cells;
        std.mem.sortUnstable(u32, self.by_coord.items, cells, struct {
            fn less(cs: []const CellInput, x: u32, y: u32) bool {
                return orderCell(cs[x].cell, cs[y].cell) == .lt;
            }
        }.less);

        for (self.input.cells) |c| {
            const key: Key = .{ .cell = c.cell };
            if (c.array) |decl| {
                const norm = decl.normalized();
                if (norm.first.row.oneBased() != c.cell.row.oneBased() or
                    norm.first.col.zeroBased() != c.cell.col.zeroBased())
                {
                    return .{ .reason = .array_anchor_not_top_left, .at = key };
                }
            }
            try self.owners.append(self.gpa, .{
                .body = c.formula,
                .owner = .{
                    .key = key,
                    .sheet = c.cell.sheet,
                    .site = .{ .row = c.cell.row, .col = c.cell.col },
                    .array_formula = c.array != null,
                },
            });
        }

        for (self.input.names, 0..) |n, i| {
            try self.owners.append(self.gpa, .{
                .body = n.body,
                .owner = .{
                    .key = .{ .name = .{
                        .scope = n.scope,
                        .identifier = n.identifier,
                        .index = @intCast(i),
                    } },
                    .sheet = n.scope,
                },
            });
        }

        for (self.input.producers, 0..) |p, i| {
            try self.owners.append(self.gpa, .{
                .body = p.body,
                .owner = .{
                    .key = .{ .producer = .{
                        .table = p.table,
                        .column = p.column,
                        .kind = p.kind,
                        .index = @intCast(i),
                    } },
                    .sheet = p.span.sheet,
                    // A producer's body is materialized into every member
                    // of its span, so `#This Row` names a different row in
                    // each one. There is no single site, and the *union*
                    // of what `[@Col]` reaches is the column's whole data
                    // band — which is what `resolveStructured` returns for
                    // a null site, so declining here is conservative
                    // rather than lossy.
                },
            });
        }
        return null;
    }

    // ─── phase 2: what each body mentions ────────────────────────

    fn captureAll(self: *Builder) Error!?Refusal {
        try self.refs.ensureTotalCapacity(self.gpa, self.owners.items.len);

        // The ASTs die here. Every `Ref` a capture produces is either a
        // coordinate or a slice borrowed from `Input`, so nothing in the
        // graph points into a parse — which is what lets one scratch
        // arena cover all of them and be released before the graph is
        // handed back. §9's `max_retained_ast_bytes` is about the
        // evaluator retaining bodies, not about this.
        var scratch = std.heap.ArenaAllocator.init(self.gpa);
        defer scratch.deinit();

        // §5.6e's edges arrive as one flat list covering every owner, so
        // asking "which of these are mine?" inside the loop below is a
        // walk of the whole list per formula — quadratic in the
        // workbook, and the largest single cost in a rebuild once M5d4
        // removed the two above it. Grouped once, by owner, here.
        //
        // Indices rather than edges, and appended in input order, so
        // each owner replays exactly the subsequence the scan would have
        // handed it: the order reads enter the dependency log is the
        // order they enter the graph, and this is a change of cost, not
        // of graph.
        var by_owner: std.HashMapUnmanaged(Key, std.ArrayListUnmanaged(u32), Key.HashContext, 80) = .empty;
        defer {
            var vals = by_owner.valueIterator();
            while (vals.next()) |v| v.deinit(self.gpa);
            by_owner.deinit(self.gpa);
        }
        for (self.opts.dynamic_edges, 0..) |d, i| {
            const gop = try by_owner.getOrPut(self.gpa, d.owner);
            if (!gop.found_existing) gop.value_ptr.* = .empty;
            try gop.value_ptr.append(self.gpa, @intCast(i));
        }

        for (self.owners.items) |bo| {
            _ = scratch.reset(.retain_capacity);
            const s = scratch.allocator();

            const parsed = try parser.parse(s, bo.body, .{ .limits = self.opts.parse_limits });
            const ast = switch (parsed) {
                .ok => |x| x,
                .refused => return .{ .reason = .formula_parse_failed, .at = bo.owner.key },
            };

            // §5.6g's context legality, before any edge is drawn. Same
            // check the evaluator runs before it evaluates, for the same
            // reason: it is a statement about the text.
            if (name_rules.checkThreeD(ast, .{ .array_formula = bo.owner.array_formula }) != null) {
                return .{ .reason = .three_d_illegal_context, .at = bo.owner.key };
            }

            var cap: Capture = .{
                .gpa = self.gpa,
                .resolver = self.resolver,
                .names = self.input.names,
                .scratch = s,
                .ast = ast,
                .owner = bo.owner,
                .sheet = bo.owner.sheet,
                .deps = eval.DependencyLog.init(self.gpa),
                .names_seen = .empty,
                .spans = .empty,
            };
            defer cap.deinit();

            // A name body that is position-dependent has no fixed
            // meaning to record. Referencing one is refused at M4b3, so
            // no cache can be computed from an edge this declines to
            // draw; inventing a sheet for it would be the unsafe choice.
            const relative_name = bo.owner.key == .name and name_rules.bodyIsRelative(ast);
            if (!relative_name) try cap.walk(ast.root);

            // The gate's snapshot (M10b): everything in the log at this
            // instant came from the walk of the text, and everything
            // after it is injected runtime capture. `Graph.walkNoted`
            // answers membership against exactly this boundary.
            const walk_cells: u32 = @intCast(cap.deps.cells.items.len);
            const walk_areas: u32 = @intCast(cap.deps.areas.items.len);

            // §5.6e's runtime capture, into the same log the walk just
            // filled. A read the walk already found dedupes here rather
            // than doubling an edge, which is why a rebuild whose
            // dynamic references landed where the text said they would
            // produces the identical graph — the fixpoint's base case.
            if (by_owner.get(bo.owner.key)) |mine| {
                for (mine.items) |i| {
                    switch (self.opts.dynamic_edges[i].target) {
                        .cell => |c| try cap.deps.noteCell(c),
                        .area => |x| try cap.deps.noteArea(x),
                    }
                }
            }

            self.refs.appendAssumeCapacity(try cap.take(self.a()));
            try self.logs.append(self.gpa, .{
                .cells = @intCast(cap.deps.cells.items.len),
                .walk_cells = walk_cells,
                .walk_areas = walk_areas,
            });
        }
        return null;
    }

    // ─── phase 3: the node set ───────────────────────────────────

    fn collectKeys(self: *Builder) Error!void {
        for (self.owners.items) |bo| try self.keys.append(self.gpa, bo.owner.key);

        // Spill tails. A tail coordinate that itself holds a formula is
        // an obstruction (M7a's decision table); the cell node wins here
        // so the coordinate has exactly one node.
        for (self.input.cells) |c| {
            const decl = (c.array orelse continue).normalized();
            var r = decl.first.row.oneBased();
            while (r <= decl.last.row.oneBased()) : (r += 1) {
                var col = decl.first.col.zeroBased();
                while (col <= decl.last.col.zeroBased()) : (col += 1) {
                    if (r == c.cell.row.oneBased() and col == c.cell.col.zeroBased()) continue;
                    const tail: env.CellRef = .{
                        .sheet = c.cell.sheet,
                        .row = coords.Row.fromOneBased(r) catch continue,
                        .col = coords.Col.fromZeroBased(col) catch continue,
                    };
                    if (self.cellInput(tail) != null) continue;
                    try self.keys.append(self.gpa, .{ .spill_tail = tail });
                    try self.tails.append(self.gpa, .{ .tail = tail, .anchor = c.cell });
                }
            }
        }
        // Two anchors whose declared ranges overlap are an obstruction
        // (M7a's decision table). Until that row decides what happens,
        // the tail belongs to whichever anchor sorts first — a rule, so
        // that the node set cannot depend on the order the cells
        // arrived in.
        std.mem.sortUnstable(TailOwner, self.tails.items, {}, struct {
            fn less(_: void, x: TailOwner, y: TailOwner) bool {
                const t = orderCell(x.tail, y.tail);
                if (t != .eq) return t == .lt;
                return orderCell(x.anchor, y.anchor) == .lt;
            }
        }.less);
        var w: usize = 0;
        for (self.tails.items, 0..) |t, r| {
            if (r > 0 and orderCell(t.tail, self.tails.items[w - 1].tail) == .eq) continue;
            self.tails.items[w] = t;
            w += 1;
        }
        self.tails.shrinkRetainingCapacity(w);

        // Range and span nodes, plus the per-member ranges a span needs.
        for (self.refs.items) |list| {
            for (list) |ref| switch (ref) {
                .cell, .name => {},
                .area => |x| try self.keys.append(self.gpa, .{ .range = x }),
                .span => |s| {
                    try self.keys.append(self.gpa, .{ .span = s });
                    var sheet = s.first.toInt();
                    while (sheet <= s.last.toInt()) : (sheet += 1) {
                        const member: env.RangeRef = .{
                            .sheet = env.SheetIndex.fromInt(sheet),
                            .range = s.range,
                        };
                        if (member.isSingleCell()) continue;
                        try self.keys.append(self.gpa, .{ .range = member });
                    }
                },
            };
        }
    }

    /// The `Input.cells` row at a coordinate, through the sorted index.
    fn cellInput(self: *Builder, cell: env.CellRef) ?CellInput {
        const items = self.by_coord.items;
        const cells = self.input.cells;
        var lo: usize = 0;
        var hi: usize = items.len;
        while (lo < hi) {
            const mid = lo + (hi - lo) / 2;
            switch (orderCell(cells[items[mid]].cell, cell)) {
                .eq => return cells[items[mid]],
                .lt => lo = mid + 1,
                .gt => hi = mid,
            }
        }
        return null;
    }

    fn anchorOf(self: *Builder, tail: env.CellRef) ?env.CellRef {
        const items = self.tails.items;
        var lo: usize = 0;
        var hi: usize = items.len;
        while (lo < hi) {
            const mid = lo + (hi - lo) / 2;
            switch (orderCell(items[mid].tail, tail)) {
                .eq => return items[mid].anchor,
                .lt => lo = mid + 1,
                .gt => hi = mid,
            }
        }
        return null;
    }

    // ─── phase 4: edges, components, order ───────────────────────

    fn link(self: *Builder) Error!BuildResult {
        const n = self.keys.items.len;
        const keys = try self.a().dupe(Key, self.keys.items);

        // Scratch, not arena: the adjacency *lists* are temporary and the
        // arena outlives the build. An arena that kept every growth step
        // of every list would hold on to them for the graph's lifetime.
        var scratch = std.heap.ArenaAllocator.init(self.gpa);
        defer scratch.deinit();
        const sa = scratch.allocator();

        const sets = try sa.alloc(std.ArrayListUnmanaged(u32), n);
        for (sets) |*x| x.* = .empty;

        const idx = try Index.build(self.a(), keys);

        // Owner bodies. The walk logs ride along (M10b): `refs` already
        // lives in the arena the graph keeps, so retention is the
        // bounds and a pointer, not a copy.
        assert(self.logs.items.len == self.owners.items.len);
        const walk_logs = try self.a().alloc(?Graph.WalkLog, n);
        @memset(walk_logs, null);
        for (self.owners.items, self.refs.items, self.logs.items) |bo, list, rec| {
            const u = findIn(keys, bo.owner.key).?;
            walk_logs[u] = .{
                .refs = list,
                .cells = rec.cells,
                .walk_cells = rec.walk_cells,
                .walk_areas = rec.walk_areas,
            };
            for (list) |ref| {
                const target: ?u32 = switch (ref) {
                    .cell => |c| findIn(keys, .{ .cell = c }) orelse
                        findIn(keys, .{ .spill_tail = c }),
                    .area => |x| findIn(keys, .{ .range = x }),
                    .name => |x| findIn(keys, .{ .name = x }),
                    .span => |x| findIn(keys, .{ .span = x }),
                };
                if (target) |v| {
                    if (try self.addEdge(sa, sets, u, v)) |r| return .{ .refused = r };
                }
            }
        }

        // Range nodes reach their producers; span nodes reach their
        // members; spill tails reach their anchor.
        for (keys, 0..) |k, i| {
            const u: u32 = @intCast(i);
            switch (k) {
                .range => |area| {
                    var probe = idx.probe(area, &self.stats.index_probes);
                    while (probe.next()) |v| {
                        if (try self.addEdge(sa, sets, u, v)) |r| return .{ .refused = r };
                    }
                },
                .span => |sp| {
                    var sheet = sp.first.toInt();
                    while (sheet <= sp.last.toInt()) : (sheet += 1) {
                        const member: env.RangeRef = .{
                            .sheet = env.SheetIndex.fromInt(sheet),
                            .range = sp.range,
                        };
                        const v = if (member.isSingleCell())
                            findIn(keys, .{ .cell = member.topLeft() }) orelse
                                findIn(keys, .{ .spill_tail = member.topLeft() })
                        else
                            findIn(keys, .{ .range = member });
                        if (v) |w| {
                            if (try self.addEdge(sa, sets, u, w)) |r| return .{ .refused = r };
                        }
                    }
                },
                .spill_tail => |tail| {
                    const anchor = self.anchorOf(tail).?;
                    const v = findIn(keys, .{ .cell = anchor }).?;
                    if (try self.addEdge(sa, sets, u, v)) |r| return .{ .refused = r };
                },
                .cell, .name, .producer => {},
            }
        }

        const deps = try self.a().alloc([]const u32, n);
        for (sets, 0..) |*set, i| {
            std.mem.sortUnstable(u32, set.items, {}, std.sort.asc(u32));
            deps[i] = try self.a().dupe(u32, set.items);
        }

        const comp = try tarjan(self.gpa, self.a(), deps);
        const cyclic = try classifyComponents(self.gpa, self.a(), deps, comp.of, comp.count);
        const order = try condensationOrder(self.gpa, self.a(), deps, comp.of, comp.count);

        self.stats.components = comp.count;
        for (cyclic) |c| {
            if (c) self.stats.cyclic_components += 1;
        }

        const seeds = switch (try self.seedAll(keys, comp.of, cyclic)) {
            .ok => |s| s,
            .refused => |r| return .{ .refused = r },
        };

        const out: Graph = .{
            .arena = self.arena,
            .keys = keys,
            .deps = deps,
            .component = comp.of,
            .order = order,
            .cyclic = cyclic,
            .seeds = seeds,
            .index = idx,
            .walk_logs = walk_logs,
            .stats = self.stats,
        };
        // The arena moved into the result; the builder must not free it.
        self.arena = std.heap.ArenaAllocator.init(self.gpa);
        return .{ .ok = out };
    }

    /// §9's `dependency_edges` charge site. One charge per **admitted**
    /// edge — a duplicate is not new work and does not pay twice.
    fn addEdge(
        self: *Builder,
        sa: std.mem.Allocator,
        sets: []std.ArrayListUnmanaged(u32),
        u: u32,
        v: u32,
    ) Error!?Refusal {
        for (sets[u].items) |existing| {
            if (existing == v) return null;
        }
        self.counters.charge(.dependency_edges, 1) catch {
            return Refusal{ .reason = .work_limit_exceeded, .limit = .dependency_edges };
        };
        // Counted here as well as charged: the counter may be shared
        // across several builds, and `stats` is about *this* graph.
        self.stats.edges += 1;
        try sets[u].append(sa, v);
        return null;
    }

    const SeedResult = union(enum) { ok: []const ?Seed, refused: Refusal };

    fn seedAll(
        self: *Builder,
        keys: []const Key,
        comp_of: []const u32,
        cyclic: []const bool,
    ) Error!SeedResult {
        const seeds = try self.a().alloc(?Seed, keys.len);
        @memset(seeds, null);
        for (keys, 0..) |k, i| {
            if (!cyclic[comp_of[i]]) continue;
            const cell = switch (k) {
                .cell => |c| c,
                else => continue,
            };
            const in = self.cellInput(cell).?;
            seeds[i] = seedFor(in.cache, anchorOfInput(in)) catch |e| switch (e) {
                error.MalformedCache => return .{ .refused = .{
                    .reason = .malformed_cache_seed,
                    .at = k,
                } },
            };
        }
        return .{ .ok = seeds };
    }
};

fn anchorOfInput(in: CellInput) Anchor {
    if (in.array) |decl| {
        const norm = decl.normalized();
        return .{ .shape = .{ .rows = norm.rowCount(), .cols = norm.colCount() } };
    }
    return if (in.dynamic_anchor) .unknown else .none;
}

fn keyLessThan(_: void, x: Key, y: Key) bool {
    return x.order(y) == .lt;
}

fn dedupeKeys(list: *std.ArrayListUnmanaged(Key)) void {
    if (list.items.len == 0) return;
    var w: usize = 1;
    var r: usize = 1;
    while (r < list.items.len) : (r += 1) {
        if (list.items[r].order(list.items[w - 1]) == .eq) continue;
        list.items[w] = list.items[r];
        w += 1;
    }
    list.shrinkRetainingCapacity(w);
}

fn findIn(keys: []const Key, key: Key) ?u32 {
    var lo: usize = 0;
    var hi: usize = keys.len;
    while (lo < hi) {
        const mid = lo + (hi - lo) / 2;
        switch (keys[mid].order(key)) {
            .eq => return @intCast(mid),
            .lt => lo = mid + 1,
            .gt => hi = mid,
        }
    }
    return null;
}

// ─── the static walk (§5.3a) ─────────────────────────────────────

/// The graph's reference capture.
///
/// It is deliberately a **different** function from
/// `eval.staticDependencies`, which stays as the demonstrator of the
/// static-versus-runtime split, and it differs from it in one visible
/// way: `A1:B2` records the area alone, where the demonstrator also
/// records both endpoints. Both are correct dependency sets; only one is
/// minimal, and a graph that draws redundant edges makes the differential
/// test compare noise.
const RefTarget = union(enum) {
    plain: env.RangeRef,
    span: Key.Span,
};

fn bboxRanges(a: coords.Range, b: coords.Range) coords.Range {
    return .{
        .first = .{
            .row = coords.Row.fromOneBased(@min(a.first.row.oneBased(), b.first.row.oneBased())) catch unreachable,
            .col = coords.Col.fromZeroBased(@min(a.first.col.zeroBased(), b.first.col.zeroBased())) catch unreachable,
        },
        .last = .{
            .row = coords.Row.fromOneBased(@max(a.last.row.oneBased(), b.last.row.oneBased())) catch unreachable,
            .col = coords.Col.fromZeroBased(@max(a.last.col.zeroBased(), b.last.col.zeroBased())) catch unreachable,
        },
    };
}

const Capture = struct {
    gpa: std.mem.Allocator,
    /// Dies with the AST it was parsed into. Nothing a capture *keeps*
    /// may be allocated here.
    scratch: std.mem.Allocator,
    resolver: Resolver,
    /// For turning a `NameBinding.name` index back into a key.
    names: []const NameInput,
    ast: parser.Ast,
    owner: Owner,
    /// The sheet unqualified references resolve against right now.
    sheet: ?env.SheetIndex,
    deps: eval.DependencyLog,
    names_seen: std.ArrayListUnmanaged(Key.Name),
    spans: std.ArrayListUnmanaged(Key.Span),

    fn deinit(self: *Capture) void {
        self.deps.deinit();
        self.names_seen.deinit(self.gpa);
        self.spans.deinit(self.gpa);
        self.* = undefined;
    }

    fn take(self: *Capture, arena: std.mem.Allocator) Error![]const Ref {
        var out = try arena.alloc(
            Ref,
            self.deps.cells.items.len + self.deps.areas.items.len +
                self.names_seen.items.len + self.spans.items.len,
        );
        var i: usize = 0;
        for (self.deps.cells.items) |c| {
            out[i] = .{ .cell = c };
            i += 1;
        }
        for (self.deps.areas.items) |x| {
            out[i] = .{ .area = x };
            i += 1;
        }
        for (self.names_seen.items) |x| {
            out[i] = .{ .name = x };
            i += 1;
        }
        for (self.spans.items) |x| {
            out[i] = .{ .span = x };
            i += 1;
        }
        assert(i == out.len);
        return out;
    }

    /// A single-cell area is a cell, exactly as `Evaluator.refValue`
    /// decides it. Keeping the two normalizations identical is what lets
    /// the graph be described in terms of `DependencyLog`.
    fn note(self: *Capture, area: env.RangeRef) Error!void {
        if (area.isSingleCell()) {
            try self.deps.noteCell(area.topLeft());
        } else {
            try self.deps.noteArea(area);
        }
    }

    fn walk(self: *Capture, i: parser.Index) Error!void {
        switch (self.ast.node(i)) {
            .number, .string, .boolean, .error_lit, .missing_arg => {},
            .array => |n| for (self.ast.children(n.elems)) |c| try self.walk(c),
            .ref_cell, .ref_full_col, .ref_full_row => {
                if (try self.refTarget(i)) |t| try self.record(t);
            },
            .name => |n| {
                const binding = try self.resolver.resolveName(self.sheet, n.raw);
                switch (binding) {
                    .unresolved => {},
                    .area => |x| try self.note(x),
                    .name => |row| {
                        const nm = self.names[row];
                        try self.names_seen.append(self.gpa, .{
                            .scope = nm.scope,
                            .identifier = nm.identifier,
                            .index = row,
                        });
                    },
                }
            },
            .structured => |n| {
                const area = try self.resolver.resolveStructured(
                    self.sheet,
                    self.owner,
                    n.table,
                    n.items,
                    n.columns,
                );
                if (area) |x| try self.note(x);
            },
            .qualified => |n| {
                if (try self.refTarget(i)) |t| {
                    try self.record(t);
                    return;
                }
                // Not a reference subtree — `Sheet1!SomeName`, a
                // qualified call. Walk the target with the sheet
                // switched; a span over a non-reference is `#VALUE!`
                // and has nothing to depend on.
                if (name_rules.isSpan(n.sheet)) return;
                const nm = try self.unquoteSheetName(n.sheet);
                const idx = (try self.resolver.resolveSheet(nm)) orelse return;
                const saved = self.sheet;
                self.sheet = idx;
                defer self.sheet = saved;
                try self.walk(n.target);
            },
            .call => |n| {
                try self.walk(n.callee);
                // Every arm, including the ones evaluation will skip
                // (§5.3a). A cell only a dead branch reads is still an
                // edge and still triggers a recalc.
                for (self.ast.children(n.args)) |c| try self.walk(c);
            },
            .paren => |n| try self.walk(n.child),
            .unary => |n| try self.walk(n.child),
            .postfix => |n| try self.walk(n.child),
            .binary => |n| {
                if (n.op == .range) {
                    if (try self.refTarget(i)) |t| {
                        // Both endpoints are inside the area they bound,
                        // so recursing would only re-record them.
                        try self.record(t);
                        return;
                    }
                }
                try self.walk(n.lhs);
                try self.walk(n.rhs);
            },
        }
    }

    fn record(self: *Capture, t: RefTarget) Error!void {
        switch (t) {
            .plain => |x| try self.note(x),
            // §5.6g: one node per span, and one edge per member sheet.
            .span => |x| try self.spans.append(self.gpa, x),
        }
    }

    /// The area a reference subtree denotes, or null when the subtree is
    /// not a reference this walk can resolve.
    ///
    /// One function rather than a case per grammar shape, because
    /// `Sheet1:Sheet3!$B$1:$B$2` is **not** a single node: the tokenizer
    /// hands back `qualified(span, $B$1)` and the `:$B$2` stays outside
    /// it, so the span and the range operator have to be resolved
    /// together or the second endpoint lands on the wrong sheet. That is
    /// the bug this shape exists to prevent.
    fn refTarget(self: *Capture, i: parser.Index) Error!?RefTarget {
        switch (self.ast.node(i)) {
            .ref_cell => |n| {
                const sheet = self.sheet orelse return null;
                return .{ .plain = .{
                    .sheet = sheet,
                    .range = .{ .first = n.cell, .last = n.cell },
                } };
            },
            .ref_full_col => |n| {
                const sheet = self.sheet orelse return null;
                return .{ .plain = .{ .sheet = sheet, .range = eval.fullColRange(n.first, n.last) } };
            },
            .ref_full_row => |n| {
                const sheet = self.sheet orelse return null;
                return .{ .plain = .{ .sheet = sheet, .range = eval.fullRowRange(n.first, n.last) } };
            },
            .paren => |n| return self.refTarget(n.child),
            .qualified => |n| {
                const name = try self.unquoteSheetName(n.sheet);
                if (name_rules.isSpan(n.sheet)) {
                    const ends = name_rules.splitSpan(n.sheet, name) orelse return null;
                    const first = try self.resolver.resolveSheet(ends.first);
                    const last = try self.resolver.resolveSheet(ends.last);
                    const members = switch (name_rules.expandSpan(
                        if (first) |f| f.toInt() else null,
                        if (last) |l| l.toInt() else null,
                    )) {
                        // A deleted endpoint reads `#REF!` everywhere it
                        // reached, so there is nothing to depend on.
                        .ref_error => return null,
                        .members => |m| m,
                    };
                    const local = targetRange(self.ast, n.target) orelse return null;
                    return .{ .span = .{
                        .first = env.SheetIndex.fromInt(members.first),
                        .last = env.SheetIndex.fromInt(members.last),
                        .range = local,
                    } };
                }
                const idx = (try self.resolver.resolveSheet(name)) orelse return null;
                const saved = self.sheet;
                self.sheet = idx;
                defer self.sheet = saved;
                return self.refTarget(n.target);
            },
            .binary => |n| {
                if (n.op != .range) return null;
                // **The left endpoint's qualifier governs the whole
                // operator.** `Sheet3!$C$5:$C$6` arrives as
                // `qualified(Sheet3, C5) : C6` — the `:` and the second
                // endpoint sit *outside* the qualified node — so
                // resolving the right endpoint as an independent
                // reference would put it on the referencing sheet and
                // silently lose the area. Same shape, same reason, for a
                // 3D span: `Sheet1:Sheet3!$B$1:$B$2` is one span over
                // `B1:B2`, not a span over `B1` plus a stray `B2`.
                const l = (try self.refTarget(n.lhs)) orelse return null;
                const rr = targetRange(self.ast, n.rhs) orelse return null;
                return switch (l) {
                    .plain => |x| RefTarget{ .plain = .{
                        .sheet = x.sheet,
                        .range = bboxRanges(x.range, rr),
                    } },
                    .span => |x| RefTarget{ .span = .{
                        .first = x.first,
                        .last = x.last,
                        .range = bboxRanges(x.range, rr),
                    } },
                };
            },
            else => return null,
        }
    }

    fn unquoteSheetName(self: *Capture, spec: parser.SheetSpec) Error![]const u8 {
        if (!spec.quoted) return spec.first;
        const raw = spec.first;
        assert(raw.len >= 2 and raw[0] == '\'' and raw[raw.len - 1] == '\'');
        const body = raw[1 .. raw.len - 1];
        if (std.mem.indexOfScalar(u8, body, '\'') == null) return body;
        var out: std.ArrayListUnmanaged(u8) = .empty;
        const arena = self.scratch;
        try out.ensureTotalCapacity(arena, body.len);
        var k: usize = 0;
        while (k < body.len) : (k += 1) {
            out.appendAssumeCapacity(body[k]);
            if (body[k] == '\'') k += 1;
        }
        return out.items;
    }
};

/// The area a 3D qualifier's target denotes on one member sheet. By
/// grammar a 3D target is a reference; anything else is `#VALUE!` at
/// evaluation and no edge here.
fn targetRange(ast: parser.Ast, i: parser.Index) ?coords.Range {
    return switch (ast.node(i)) {
        .ref_cell => |n| .{ .first = n.cell, .last = n.cell },
        .ref_full_col => |n| eval.fullColRange(n.first, n.last),
        .ref_full_row => |n| eval.fullRowRange(n.first, n.last),
        .paren => |n| targetRange(ast, n.child),
        .binary => |n| if (n.op == .range) blk: {
            const x = rangeEndpoint(ast, n.lhs) orelse break :blk null;
            const y = rangeEndpoint(ast, n.rhs) orelse break :blk null;
            break :blk boundingBox(x, y);
        } else null,
        else => null,
    };
}

fn rangeEndpoint(ast: parser.Ast, i: parser.Index) ?coords.Cell {
    return switch (ast.node(i)) {
        .ref_cell => |n| n.cell,
        .paren => |n| rangeEndpoint(ast, n.child),
        else => null,
    };
}

fn boundingBox(x: coords.Cell, y: coords.Cell) coords.Range {
    return (coords.Range{ .first = x, .last = y }).normalized();
}

// ─── the producer index (§5.6a's sparseness) ─────────────────────

const Coord = struct {
    sheet: u32,
    row: u32,
    col: u32,
    node: u32,
};

/// Every coordinate a node occupies, sorted twice.
///
/// Two orders rather than one because the band an area asks for can be
/// narrow in either dimension: `A:A` is one column and a million rows,
/// `1:1` is one row and sixteen thousand columns. Probing through the
/// index whose leading key is the narrower band is what makes both
/// proportional to what is *stored*.
pub const Index = struct {
    by_rc: []Coord,
    by_cr: []Coord,

    fn build(arena: std.mem.Allocator, keys: []const Key) Error!Index {
        var n: usize = 0;
        for (keys) |k| {
            if (k.isCellLike()) n += 1;
        }
        const rc = try arena.alloc(Coord, n);
        var i: usize = 0;
        for (keys, 0..) |k, node| {
            const c = k.coordinate() orelse continue;
            rc[i] = .{
                .sheet = c.sheet.toInt(),
                .row = c.row.oneBased(),
                .col = c.col.zeroBased(),
                .node = @intCast(node),
            };
            i += 1;
        }
        const cr = try arena.dupe(Coord, rc);
        std.mem.sortUnstable(Coord, rc, {}, lessRc);
        std.mem.sortUnstable(Coord, cr, {}, lessCr);
        return .{ .by_rc = rc, .by_cr = cr };
    }

    fn lessRc(_: void, x: Coord, y: Coord) bool {
        if (x.sheet != y.sheet) return x.sheet < y.sheet;
        if (x.row != y.row) return x.row < y.row;
        return x.col < y.col;
    }

    fn lessCr(_: void, x: Coord, y: Coord) bool {
        if (x.sheet != y.sheet) return x.sheet < y.sheet;
        if (x.col != y.col) return x.col < y.col;
        return x.row < y.row;
    }

    /// Probe through whichever band holds fewer *stored* coordinates —
    /// which is what the walk actually pays for.
    ///
    /// Counted, not inferred from the area's extent. The extent is a
    /// proxy, and it misfires exactly where it costs most: `A5:A9` is
    /// one column and five rows, so the narrower band is the column —
    /// and the column band is every cell stored in column A, while the
    /// row band is five rows of a sparse sheet. A workbook whose every
    /// row reads a short window of one column therefore walked the whole
    /// column once per row — quadratic in the row count, and one of the
    /// terms M5d4 had to remove before the recalc pipeline scaled
    /// linearly. Four binary searches replace it.
    ///
    /// Whole-column and whole-row areas pick what they always picked:
    /// counting agrees with the extent wherever the extent was right.
    /// The two orders enumerate the same *set* — `next` filters on the
    /// whole rectangle either way — so this changes what a probe costs,
    /// not what it finds.
    fn probe(self: Index, area: env.RangeRef, probes: *u64) Probe {
        const sheet = area.sheet.toInt();
        const rows = band(
            self.by_rc,
            sheet,
            area.range.first.row.oneBased(),
            area.range.last.row.oneBased(),
            true,
        );
        const cols = band(
            self.by_cr,
            sheet,
            area.range.first.col.zeroBased(),
            area.range.last.col.zeroBased(),
            false,
        );
        const row_major = rows.len <= cols.len;
        return .{
            .items = if (row_major) self.by_rc else self.by_cr,
            .i = if (row_major) rows.start else cols.start,
            .area = area,
            .row_major = row_major,
            .probes = probes,
        };
    }

    const Band = struct { start: usize, len: usize };

    /// The half-open run of entries on `sheet` whose leading key lies in
    /// `[lo, hi]`.
    fn band(items: []const Coord, sheet: u32, lo: u32, hi: u32, row_major: bool) Band {
        const first = lowerBound(items, sheet, lo, row_major);
        const past = lowerBound(items, sheet, hi +| 1, row_major);
        assert(past >= first);
        return .{ .start = first, .len = past - first };
    }

    /// First index whose (sheet, leading key) is >= (sheet, lo).
    fn lowerBound(items: []const Coord, sheet: u32, lo: u32, row_major: bool) usize {
        var a: usize = 0;
        var b: usize = items.len;
        while (a < b) {
            const mid = a + (b - a) / 2;
            const c = items[mid];
            const lead = if (row_major) c.row else c.col;
            const before = c.sheet < sheet or (c.sheet == sheet and lead < lo);
            if (before) a = mid + 1 else b = mid;
        }
        return a;
    }

    pub const Probe = struct {
        items: []const Coord,
        i: usize,
        area: env.RangeRef,
        row_major: bool,
        probes: *u64,

        pub fn next(self: *Probe) ?u32 {
            const sheet = self.area.sheet.toInt();
            const r1 = self.area.range.first.row.oneBased();
            const r2 = self.area.range.last.row.oneBased();
            const c1 = self.area.range.first.col.zeroBased();
            const c2 = self.area.range.last.col.zeroBased();
            while (self.i < self.items.len) {
                const c = self.items[self.i];
                if (c.sheet != sheet) return null;
                const lead = if (self.row_major) c.row else c.col;
                const hi = if (self.row_major) r2 else c2;
                if (lead > hi) return null;
                self.i += 1;
                self.probes.* += 1;
                if (c.row >= r1 and c.row <= r2 and c.col >= c1 and c.col <= c2) {
                    return c.node;
                }
            }
            return null;
        }
    };
};

// ─── strongly connected components ───────────────────────────────

const Components = struct {
    of: []u32,
    count: u32,
};

/// Tarjan, iterative.
///
/// Iterative rather than recursive because the recursion depth is the
/// length of a dependency chain, which a workbook controls; a 200 000-cell
/// chain is a legal file and would be a stack overflow. §9's
/// `max_eval_depth` bounds the *closure walk*, not this — a graph is
/// built before anyone asks what to evaluate.
fn tarjan(gpa: std.mem.Allocator, arena: std.mem.Allocator, deps: []const []const u32) Error!Components {
    const n = deps.len;
    const unset = std.math.maxInt(u32);

    const disc = try gpa.alloc(u32, n);
    defer gpa.free(disc);
    const low = try gpa.alloc(u32, n);
    defer gpa.free(low);
    const on_stack = try gpa.alloc(bool, n);
    defer gpa.free(on_stack);
    const comp = try arena.alloc(u32, n);
    @memset(disc, unset);
    @memset(comp, unset);
    @memset(on_stack, false);

    const Frame = struct { v: u32, i: u32 };
    var frames: std.ArrayListUnmanaged(Frame) = .empty;
    defer frames.deinit(gpa);
    var stack: std.ArrayListUnmanaged(u32) = .empty;
    defer stack.deinit(gpa);

    var next: u32 = 0;
    var count: u32 = 0;

    for (0..n) |root_usize| {
        const root: u32 = @intCast(root_usize);
        if (disc[root] != unset) continue;

        disc[root] = next;
        low[root] = next;
        next += 1;
        try stack.append(gpa, root);
        on_stack[root] = true;
        try frames.append(gpa, .{ .v = root, .i = 0 });

        while (frames.items.len > 0) {
            const top = frames.items.len - 1;
            const v = frames.items[top].v;
            const i = frames.items[top].i;

            if (i < deps[v].len) {
                frames.items[top].i = i + 1;
                const w = deps[v][i];
                if (disc[w] == unset) {
                    disc[w] = next;
                    low[w] = next;
                    next += 1;
                    try stack.append(gpa, w);
                    on_stack[w] = true;
                    try frames.append(gpa, .{ .v = w, .i = 0 });
                } else if (on_stack[w]) {
                    low[v] = @min(low[v], disc[w]);
                }
                continue;
            }

            if (low[v] == disc[v]) {
                while (true) {
                    const w = stack.pop().?;
                    on_stack[w] = false;
                    comp[w] = count;
                    if (w == v) break;
                }
                count += 1;
            }
            _ = frames.pop();
            if (frames.items.len > 0) {
                const p = frames.items[frames.items.len - 1].v;
                low[p] = @min(low[p], low[v]);
            }
        }
    }

    for (comp) |c| assert(c != unset);
    return .{ .of = comp, .count = count };
}

/// A component is cyclic when it has more than one member, or exactly
/// one that depends on itself. `A1=A1+1` is a cycle; `A1=B1` is not,
/// however many components sit around it.
fn classifyComponents(
    gpa: std.mem.Allocator,
    arena: std.mem.Allocator,
    deps: []const []const u32,
    comp_of: []const u32,
    count: u32,
) Error![]bool {
    const size = try gpa.alloc(u32, count);
    defer gpa.free(size);
    @memset(size, 0);
    for (comp_of) |c| size[c] += 1;

    const out = try arena.alloc(bool, count);
    for (out, 0..) |*b, c| b.* = size[c] > 1;
    for (deps, 0..) |list, u| {
        for (list) |v| {
            if (v == u) out[comp_of[u]] = true;
        }
    }
    return out;
}

/// The condensation DAG, in the canonical topological order.
///
/// Kahn with a min-heap rather than the reverse of Tarjan's output. Both
/// are deterministic; only this one is **canonical** — it emits, at
/// every step, the ready component with the smallest node under
/// `Key.order`, so the result depends on the graph and on nothing else
/// (not on which node a depth-first search happened to start from).
fn condensationOrder(
    gpa: std.mem.Allocator,
    arena: std.mem.Allocator,
    deps: []const []const u32,
    comp_of: []const u32,
    count: u32,
) Error![]const []const u32 {
    // Members, ascending. Node ids are canonical, so "ascending" is
    // "in `Key.order`".
    const sizes = try gpa.alloc(u32, count);
    defer gpa.free(sizes);
    @memset(sizes, 0);
    for (comp_of) |c| sizes[c] += 1;

    const members = try arena.alloc([]u32, count);
    for (members, sizes) |*m, s| m.* = try arena.alloc(u32, s);
    const filled = try gpa.alloc(u32, count);
    defer gpa.free(filled);
    @memset(filled, 0);
    for (comp_of, 0..) |c, u| {
        members[c][filled[c]] = @intCast(u);
        filled[c] += 1;
    }

    // Condensation edges, deduped: `needs[c]` counts distinct components
    // c depends on, `feeds[c]` lists the components that depend on c.
    const feeds = try gpa.alloc(std.ArrayListUnmanaged(u32), count);
    for (feeds) |*f| f.* = .empty;
    defer {
        for (feeds) |*f| f.deinit(gpa);
        gpa.free(feeds);
    }
    const needs = try gpa.alloc(u32, count);
    defer gpa.free(needs);
    @memset(needs, 0);

    var seen: std.ArrayListUnmanaged(u32) = .empty;
    defer seen.deinit(gpa);
    for (0..count) |c| {
        seen.clearRetainingCapacity();
        for (members[c]) |u| {
            for (deps[u]) |v| {
                const d = comp_of[v];
                if (d == c) continue;
                var dup = false;
                for (seen.items) |s| {
                    if (s == d) {
                        dup = true;
                        break;
                    }
                }
                if (dup) continue;
                try seen.append(gpa, d);
                needs[c] += 1;
                try feeds[d].append(gpa, @intCast(c));
            }
        }
    }

    var heap: MinHeap = .{ .items = .empty, .keys = members };
    defer heap.items.deinit(gpa);
    for (0..count) |c| {
        if (needs[c] == 0) try heap.push(gpa, @intCast(c));
    }

    const out = try arena.alloc([]const u32, count);
    var emitted: usize = 0;
    while (heap.pop()) |c| {
        out[emitted] = members[c];
        emitted += 1;
        for (feeds[c].items) |d| {
            needs[d] -= 1;
            if (needs[d] == 0) try heap.push(gpa, d);
        }
    }
    // The condensation of any digraph is acyclic, so Kahn always drains.
    assert(emitted == count);
    return out;
}

/// A binary min-heap over component ids, ordered by each component's
/// smallest member. Hand-rolled and thirty lines, because the ordering
/// is the contract this row is tested on and it should not move when a
/// container's comparator signature does.
const MinHeap = struct {
    items: std.ArrayListUnmanaged(u32),
    keys: []const []const u32,

    fn less(self: MinHeap, x: u32, y: u32) bool {
        return self.keys[x][0] < self.keys[y][0];
    }

    fn push(self: *MinHeap, arena: std.mem.Allocator, c: u32) Error!void {
        try self.items.append(arena, c);
        var i = self.items.items.len - 1;
        while (i > 0) {
            const parent = (i - 1) / 2;
            if (!self.less(self.items.items[i], self.items.items[parent])) break;
            std.mem.swap(u32, &self.items.items[i], &self.items.items[parent]);
            i = parent;
        }
    }

    fn pop(self: *MinHeap) ?u32 {
        if (self.items.items.len == 0) return null;
        const top = self.items.items[0];
        const last = self.items.pop().?;
        if (self.items.items.len == 0) return top;
        self.items.items[0] = last;
        var i: usize = 0;
        while (true) {
            const l = 2 * i + 1;
            const r = l + 1;
            var m = i;
            if (l < self.items.items.len and self.less(self.items.items[l], self.items.items[m])) m = l;
            if (r < self.items.items.len and self.less(self.items.items[r], self.items.items[m])) m = r;
            if (m == i) break;
            std.mem.swap(u32, &self.items.items[i], &self.items.items[m]);
            i = m;
        }
        return top;
    }
};

// ─── closure planning (§5.6f) ────────────────────────────────────

/// What a closure evaluation will do, decided before it does any of it.
pub const Plan = struct {
    /// Formula cells to evaluate, in evaluation order.
    ///
    /// Cells only. A spill tail takes its value from its anchor and a
    /// name expands inline at its reference site, so neither is a unit
    /// of evaluation — they are units of *ordering*, which is a
    /// different thing and the reason they are nodes.
    cells: []const u32,
    /// The components the closure covers, in evaluation order.
    components: []const u32,
};

pub const PlanResult = union(enum) {
    ok: Plan,
    refused: Refusal,
};

pub const PlanOptions = struct {
    /// Whether the workbook's `calcPr` asks for iteration.
    ///
    /// M5a1 had no engine to hand a cycle to, so this was `false` in
    /// every line of code that could reach it and §5.6c's "with
    /// iteration off, a cycle is `FormulaCycle`" was the whole rule.
    /// M5a2 supplies the other half: with iteration on, a cyclic
    /// component is admitted to the plan and the iteration engine
    /// schedules it. Planning is the only thing that changes — the node
    /// model, the index and the order are what they were.
    iterating: bool = false,
    /// Whether admitting a cell charges `total_cell_evals`.
    ///
    /// True for a one-shot closure, where the plan IS the work and
    /// charging at admission refuses before the first evaluation rather
    /// than halfway through one. False for M5a2's engine, which re-plans
    /// once per §5.6e pass and charges per evaluation instead — the
    /// number §9 actually bounds for an iterating run is passes times
    /// members, and a plan charged per pass would count the same cell
    /// once for planning and once for running it.
    charge_evals: bool = true,
};

/// The transitive closure of `roots`, ordered.
///
/// §9 charge sites, both here: `eval_depth` on pushing a cell-like node
/// and released on popping it — the intermediary range, span, name and
/// producer nodes are how one cell *reaches* another and do not consume
/// a cell's worth of depth — and `total_cell_evals` once per cell
/// admitted to the plan, which refuses before the first evaluation
/// instead of halfway through one.
pub fn plan(
    g: Graph,
    arena: std.mem.Allocator,
    roots: []const Key,
    counters: *WorkCounters,
    opts: PlanOptions,
) Error!PlanResult {
    const n = g.keys.len;
    const reached = try arena.alloc(bool, n);
    @memset(reached, false);

    const Frame = struct { node: u32, i: u32, charged: bool };
    var stack: std.ArrayListUnmanaged(Frame) = .empty;
    defer stack.deinit(arena);

    // A root the graph has no node for is not an empty closure. An area
    // no stored formula mentions has no range node, and a coordinate an
    // anchor spilled into is a tail rather than a cell — both resolve to
    // the nodes they cover instead of being dropped.
    var starts: std.ArrayListUnmanaged(u32) = .empty;
    defer starts.deinit(arena);
    var probes: u64 = 0;
    for (roots) |root| {
        if (g.find(root)) |node| {
            try starts.append(arena, node);
            continue;
        }
        switch (root) {
            .cell => |c| {
                if (g.find(.{ .spill_tail = c })) |node| try starts.append(arena, node);
            },
            .range => |area| {
                var probe = g.probeArea(area, &probes);
                while (probe.next()) |node| try starts.append(arena, node);
            },
            .span => |sp| {
                var sheet = sp.first.toInt();
                while (sheet <= sp.last.toInt()) : (sheet += 1) {
                    var probe = g.probeArea(.{
                        .sheet = env.SheetIndex.fromInt(sheet),
                        .range = sp.range,
                    }, &probes);
                    while (probe.next()) |node| try starts.append(arena, node);
                }
            },
            else => {},
        }
    }

    for (starts.items) |start| {
        if (reached[start]) continue;
        reached[start] = true;
        const charged = g.keys[start].isCellLike();
        if (charged) {
            counters.charge(.eval_depth, 1) catch return .{ .refused = .{
                .reason = .work_limit_exceeded,
                .limit = .eval_depth,
                .at = g.keys[start],
            } };
        }
        try stack.append(arena, .{ .node = start, .i = 0, .charged = charged });

        while (stack.items.len > 0) {
            const top = stack.items.len - 1;
            const u = stack.items[top].node;
            const i = stack.items[top].i;
            if (i < g.deps[u].len) {
                stack.items[top].i = i + 1;
                const v = g.deps[u][i];
                if (reached[v]) continue;
                reached[v] = true;
                const c = g.keys[v].isCellLike();
                if (c) {
                    counters.charge(.eval_depth, 1) catch return .{ .refused = .{
                        .reason = .work_limit_exceeded,
                        .limit = .eval_depth,
                        .at = g.keys[v],
                    } };
                }
                try stack.append(arena, .{ .node = v, .i = 0, .charged = c });
                continue;
            }
            if (stack.items[top].charged) counters.release(.eval_depth, 1);
            _ = stack.pop();
        }
    }

    var components: std.ArrayListUnmanaged(u32) = .empty;
    errdefer components.deinit(arena);
    var cells: std.ArrayListUnmanaged(u32) = .empty;
    errdefer cells.deinit(arena);

    for (g.order) |comp| {
        // A component is strongly connected, so a reached member means a
        // reached component: checking the first is checking all of them.
        if (!reached[comp[0]]) continue;
        const cid = g.component[comp[0]];
        if (g.cyclic[cid] and !opts.iterating) {
            // §5.6c: with iteration off, a cycle is `FormulaCycle`.
            // With it on, the component is planned like any other and
            // `iterate.zig` gives it a pass counter instead.
            return .{ .refused = .{ .reason = .cycle, .at = g.keys[comp[0]] } };
        }
        try components.append(arena, cid);
        for (comp) |node| {
            if (g.keys[node].kind() != .cell) continue;
            if (opts.charge_evals) {
                counters.charge(.total_cell_evals, 1) catch return .{ .refused = .{
                    .reason = .work_limit_exceeded,
                    .limit = .total_cell_evals,
                    .at = g.keys[node],
                } };
            }
            try cells.append(arena, node);
        }
    }

    return .{ .ok = .{
        .cells = try cells.toOwnedSlice(arena),
        .components = try components.toOwnedSlice(arena),
    } };
}

/// The nodes a standalone formula reads.
///
/// The same walk the builder runs over a cell body, so a closure
/// evaluated for `=SUM(A1:B2)` covers exactly what the identical text in
/// a cell would — which is the property that makes §5.6f's "both
/// behaviours, no silent switch" checkable rather than asserted.
///
/// Takes a parsed tree rather than text: the caller already has to own
/// the parse refusal and §5.6g's context check, because those are the
/// same answers `Workbook.evaluate` gives for the same formula and the
/// two entry points must not disagree about what a refusal is.
///
/// A returned key need not be a node — `SUM(B1:B2)` mentions an area no
/// stored formula does — so `plan` resolves an unmatched area against
/// the producer index rather than treating it as an empty closure.
pub fn rootsOfAst(
    gpa: std.mem.Allocator,
    arena: std.mem.Allocator,
    input: Input,
    resolver: Resolver,
    owner: Owner,
    ast: parser.Ast,
) Error![]const Key {
    var scratch = std.heap.ArenaAllocator.init(gpa);
    defer scratch.deinit();

    var cap: Capture = .{
        .gpa = gpa,
        .resolver = resolver,
        .names = input.names,
        .scratch = scratch.allocator(),
        .ast = ast,
        .owner = owner,
        .sheet = owner.sheet,
        .deps = eval.DependencyLog.init(gpa),
        .names_seen = .empty,
        .spans = .empty,
    };
    defer cap.deinit();
    try cap.walk(ast.root);

    const refs = try cap.take(scratch.allocator());
    const out = try arena.alloc(Key, refs.len);
    for (refs, out) |r, *dst| {
        dst.* = switch (r) {
            .cell => |c| .{ .cell = c },
            .area => |x| .{ .range = x },
            .name => |x| .{ .name = x },
            .span => |x| .{ .span = x },
        };
    }
    return out;
}

/// Where a body sits, for the two questions a walk cannot answer from
/// the text: which sheet an unqualified reference means, and which row
/// `#This Row` names.
pub const Owner = struct {
    key: Key,
    /// Null at workbook level (a workbook-scoped name's own body).
    sheet: ?env.SheetIndex = null,
    site: ?coords.Cell = null,
    array_formula: bool = false,
};

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

/// A resolver over a hand-built world. `pkg/workbook.zig` has the other
/// implementation; the vtable exists because there are two.
const World = struct {
    sheets: []const []const u8,
    names: []const NameInput = &.{},
    tables: []const Table = &.{},

    const Table = struct {
        name: []const u8,
        geometry: TableGeometry,
        columns: []const []const u8,
    };

    fn resolver(self: *const World) Resolver {
        return .{ .ctx = @constCast(self), .vtable = &vtable };
    }

    const vtable: Resolver.VTable = .{
        .resolveSheet = vtSheet,
        .resolveName = vtName,
        .resolveStructured = vtStructured,
    };

    fn of(ctx: *anyopaque) *const World {
        return @ptrCast(@alignCast(ctx));
    }

    fn vtSheet(ctx: *anyopaque, name: []const u8) Error!?env.SheetIndex {
        for (of(ctx).sheets, 0..) |s, i| {
            if (std.ascii.eqlIgnoreCase(s, name)) return env.SheetIndex.fromInt(@intCast(i));
        }
        return null;
    }

    fn vtName(ctx: *anyopaque, from: ?env.SheetIndex, spelling: []const u8) Error!NameBinding {
        const w = of(ctx);
        // §5.9's order, minus the tiers a fake does not need: a
        // sheet-scoped name shadows a workbook one, and a table is the
        // tier after both.
        if (from) |sheet| {
            for (w.names, 0..) |n, i| {
                const sc = n.scope orelse continue;
                if (sc != sheet) continue;
                if (std.ascii.eqlIgnoreCase(n.identifier, spelling)) return .{ .name = @intCast(i) };
            }
        }
        for (w.names, 0..) |n, i| {
            if (n.scope != null) continue;
            if (std.ascii.eqlIgnoreCase(n.identifier, spelling)) return .{ .name = @intCast(i) };
        }
        for (w.tables) |t| {
            if (!std.ascii.eqlIgnoreCase(t.name, spelling)) continue;
            const a = tableArea(t.geometry, .{ .data = true }, 0, @intCast(t.columns.len - 1), null);
            return if (a) |x| .{ .area = x } else .unresolved;
        }
        return .unresolved;
    }

    fn vtStructured(
        ctx: *anyopaque,
        from: ?env.SheetIndex,
        owner: Owner,
        table: ?[]const u8,
        items: parser.ItemSet,
        columns: parser.ColumnSelector,
    ) Error!?env.RangeRef {
        _ = from;
        const w = of(ctx);
        // The bare `[…]` same-table form needs to know which table the
        // owner sits in. The fake declines rather than guessing.
        const want = table orelse return null;
        for (w.tables) |t| {
            if (!std.ascii.eqlIgnoreCase(t.name, want)) continue;
            const first: u32, const last: u32 = switch (columns) {
                .none => .{ 0, @intCast(t.columns.len - 1) },
                .one => |name| blk: {
                    const i = columnIndex(t, name) orelse return null;
                    break :blk .{ i, i };
                },
                .range => |r| blk: {
                    const a = columnIndex(t, r.first) orelse return null;
                    const b = columnIndex(t, r.last) orelse return null;
                    break :blk .{ a, b };
                },
            };
            return tableArea(t.geometry, items, first, last, if (owner.site) |s| s.row else null);
        }
        return null;
    }

    fn columnIndex(t: Table, name: []const u8) ?u32 {
        for (t.columns, 0..) |col, i| {
            if (std.ascii.eqlIgnoreCase(col, name)) return @intCast(i);
        }
        return null;
    }
};

fn cellAt(sheet: u32, a1: []const u8) env.CellRef {
    const p = coords.parseCell(a1, .{ .dollar = .accept }) catch unreachable;
    return .{ .sheet = env.SheetIndex.fromInt(sheet), .row = p.row, .col = p.col };
}

fn areaAt(sheet: u32, a1: []const u8) env.RangeRef {
    const r = coords.parseRange(a1, .{ .dollar = .accept }) catch unreachable;
    return .{ .sheet = env.SheetIndex.fromInt(sheet), .range = r.normalized() };
}

fn rangeOf(a1: []const u8) coords.Range {
    return (coords.parseRange(a1, .{ .dollar = .accept }) catch unreachable).normalized();
}

const two_sheets = World{ .sheets = &.{ "Sheet1", "Sheet2" } };

fn buildOk(gpa: std.mem.Allocator, input: Input, world: *const World, opts: Options) !Graph {
    return switch (try build(gpa, input, world.resolver(), opts)) {
        .ok => |g| g,
        .refused => |r| {
            std.debug.print("unexpected refusal: {t}\n", .{r.reason});
            return error.UnexpectedRefusal;
        },
    };
}

fn buildRefused(gpa: std.mem.Allocator, input: Input, world: *const World, opts: Options) !Refusal {
    return switch (try build(gpa, input, world.resolver(), opts)) {
        .ok => |g| {
            var g2 = g;
            g2.deinit();
            return error.UnexpectedSuccess;
        },
        .refused => |r| r,
    };
}

/// `a` depends on `b`.
fn hasEdge(g: Graph, a: Key, b: Key) bool {
    const u = g.find(a) orelse return false;
    const v = g.find(b) orelse return false;
    for (g.deps[u]) |w| {
        if (w == v) return true;
    }
    return false;
}

// ─── the node model ──────────────────────────────────────────────

test "nodes: a formula cell reaching another is one edge, and a constant is none" {
    const cells = [_]CellInput{
        .{ .cell = cellAt(0, "A1"), .formula = "B1+C1" },
        .{ .cell = cellAt(0, "B1"), .formula = "1" },
    };
    var g = try buildOk(testing.allocator, .{ .sheet_count = 2, .cells = &cells }, &two_sheets, .{});
    defer g.deinit();

    try testing.expect(hasEdge(g, .{ .cell = cellAt(0, "A1") }, .{ .cell = cellAt(0, "B1") }));
    // C1 holds no formula, so it is not a node and cannot be an edge.
    try testing.expectEqual(@as(?u32, null), g.find(.{ .cell = cellAt(0, "C1") }));
    try testing.expectEqual(@as(u32, 2), g.stats.nodes);
    try testing.expectEqual(@as(u64, 1), g.stats.edges);
}

test "nodes: an area is one range node, shared by every reader" {
    const cells = [_]CellInput{
        .{ .cell = cellAt(0, "D1"), .formula = "SUM(A1:B2)" },
        .{ .cell = cellAt(0, "D2"), .formula = "SUM(A1:B2)" },
        .{ .cell = cellAt(0, "A1"), .formula = "1" },
        .{ .cell = cellAt(0, "B2"), .formula = "2" },
    };
    var g = try buildOk(testing.allocator, .{ .sheet_count = 2, .cells = &cells }, &two_sheets, .{});
    defer g.deinit();

    const r: Key = .{ .range = areaAt(0, "A1:B2") };
    try testing.expect(g.find(r) != null);
    try testing.expect(hasEdge(g, .{ .cell = cellAt(0, "D1") }, r));
    try testing.expect(hasEdge(g, .{ .cell = cellAt(0, "D2") }, r));
    try testing.expect(hasEdge(g, r, .{ .cell = cellAt(0, "A1") }));
    try testing.expect(hasEdge(g, r, .{ .cell = cellAt(0, "B2") }));
    // 4 cells + 1 range. Two readers, one range node — that is the
    // whole point of §5.6b's range nodes.
    try testing.expectEqual(@as(u32, 5), g.stats.nodes);
}

test "nodes: a single-cell area is a cell reference, exactly as the evaluator decides it" {
    const cells = [_]CellInput{
        .{ .cell = cellAt(0, "D1"), .formula = "SUM(A1:A1)" },
        .{ .cell = cellAt(0, "A1"), .formula = "1" },
    };
    var g = try buildOk(testing.allocator, .{ .sheet_count = 2, .cells = &cells }, &two_sheets, .{});
    defer g.deinit();

    try testing.expectEqual(@as(?u32, null), g.find(.{ .range = areaAt(0, "A1:A1") }));
    try testing.expect(hasEdge(g, .{ .cell = cellAt(0, "D1") }, .{ .cell = cellAt(0, "A1") }));
}

test "nodes: a defined name is a node, and a cycle through two names is a cycle" {
    const names = [_]NameInput{
        .{ .identifier = "Alpha", .body = "Beta+1" },
        .{ .identifier = "Beta", .body = "Alpha+1" },
    };
    const world = World{ .sheets = &.{ "Sheet1", "Sheet2" }, .names = &names };
    const cells = [_]CellInput{
        .{ .cell = cellAt(0, "A1"), .formula = "Alpha" },
    };
    var g = try buildOk(
        testing.allocator,
        .{ .sheet_count = 2, .cells = &cells, .names = &names },
        &world,
        .{},
    );
    defer g.deinit();

    const alpha: Key = .{ .name = .{ .scope = null, .identifier = "Alpha", .index = 0 } };
    const beta: Key = .{ .name = .{ .scope = null, .identifier = "Beta", .index = 1 } };
    try testing.expect(hasEdge(g, .{ .cell = cellAt(0, "A1") }, alpha));
    try testing.expect(hasEdge(g, alpha, beta));
    try testing.expect(hasEdge(g, beta, alpha));
    try testing.expect(g.isCyclic(g.find(alpha).?));
    try testing.expect(!g.isCyclic(g.find(.{ .cell = cellAt(0, "A1") }).?));
}

test "nodes: a 3D span is one node with one edge per member sheet (§5.6g)" {
    const world = World{ .sheets = &.{ "Sheet1", "Sheet2", "Sheet3" } };
    const cells = [_]CellInput{
        .{ .cell = cellAt(0, "A1"), .formula = "SUM(Sheet1:Sheet3!$B$1:$B$2)" },
        .{ .cell = cellAt(1, "B1"), .formula = "1" },
        .{ .cell = cellAt(2, "B2"), .formula = "2" },
    };
    var g = try buildOk(testing.allocator, .{ .sheet_count = 3, .cells = &cells }, &world, .{});
    defer g.deinit();

    const span: Key = .{ .span = .{
        .first = env.SheetIndex.fromInt(0),
        .last = env.SheetIndex.fromInt(2),
        .range = rangeOf("B1:B2"),
    } };
    try testing.expect(hasEdge(g, .{ .cell = cellAt(0, "A1") }, span));
    for (0..3) |i| {
        try testing.expect(hasEdge(g, span, .{ .range = areaAt(@intCast(i), "B1:B2") }));
    }
    try testing.expect(hasEdge(g, .{ .range = areaAt(1, "B1:B2") }, .{ .cell = cellAt(1, "B1") }));
    try testing.expect(hasEdge(g, .{ .range = areaAt(2, "B1:B2") }, .{ .cell = cellAt(2, "B2") }));
}

test "nodes: a 3D span in an ineligible function refuses before an edge is drawn" {
    const world = World{ .sheets = &.{ "Sheet1", "Sheet2" } };
    const cells = [_]CellInput{
        .{ .cell = cellAt(0, "A1"), .formula = "ROUND(Sheet1:Sheet2!$B$1,0)" },
    };
    const r = try buildRefused(testing.allocator, .{ .sheet_count = 2, .cells = &cells }, &world, .{});
    try testing.expectEqual(Refusal.Reason.three_d_illegal_context, r.reason);
}

test "nodes: a spill tail is representable, and it depends on its anchor" {
    const cells = [_]CellInput{
        .{ .cell = cellAt(0, "A1"), .formula = "B1:B2", .array = rangeOf("A1:A2") },
        .{ .cell = cellAt(0, "D1"), .formula = "A2" },
    };
    var g = try buildOk(testing.allocator, .{ .sheet_count = 2, .cells = &cells }, &two_sheets, .{});
    defer g.deinit();

    const tail: Key = .{ .spill_tail = cellAt(0, "A2") };
    try testing.expect(g.find(tail) != null);
    try testing.expect(hasEdge(g, tail, .{ .cell = cellAt(0, "A1") }));
    // A reader of the tail coordinate reads the tail node, not the anchor.
    try testing.expect(hasEdge(g, .{ .cell = cellAt(0, "D1") }, tail));
}

test "nodes: resizing a declared array changes the tails and the edges" {
    const small = [_]CellInput{
        .{ .cell = cellAt(0, "A1"), .formula = "1", .array = rangeOf("A1:A2") },
    };
    const large = [_]CellInput{
        .{ .cell = cellAt(0, "A1"), .formula = "1", .array = rangeOf("A1:A3") },
    };
    var g1 = try buildOk(testing.allocator, .{ .sheet_count = 2, .cells = &small }, &two_sheets, .{});
    defer g1.deinit();
    var g2 = try buildOk(testing.allocator, .{ .sheet_count = 2, .cells = &large }, &two_sheets, .{});
    defer g2.deinit();

    try testing.expect(g1.find(.{ .spill_tail = cellAt(0, "A3") }) == null);
    try testing.expect(g2.find(.{ .spill_tail = cellAt(0, "A3") }) != null);
    try testing.expect(g1.stats.nodes < g2.stats.nodes);
    try testing.expect(g1.stats.edges < g2.stats.edges);
}

test "nodes: an array whose declared range does not start at its cell refuses" {
    const cells = [_]CellInput{
        .{ .cell = cellAt(0, "B2"), .formula = "1", .array = rangeOf("A1:C3") },
    };
    const r = try buildRefused(testing.allocator, .{ .sheet_count = 2, .cells = &cells }, &two_sheets, .{});
    try testing.expectEqual(Refusal.Reason.array_anchor_not_top_left, r.reason);
}

test "nodes: a table producer is a node with its own dependencies" {
    const tables = [_]World.Table{.{
        .name = "T",
        .geometry = .{
            .sheet = env.SheetIndex.fromInt(0),
            .ref = rangeOf("A1:B3"),
            .header_rows = 1,
        },
        .columns = &.{ "Qty", "Total" },
    }};
    const world = World{ .sheets = &.{ "Sheet1", "Sheet2" }, .tables = &tables };
    const producers = [_]ProducerInput{.{
        .table = "T",
        .column = "Total",
        .kind = .calculated_column,
        .body = "T[Qty]*2",
        .span = areaAt(0, "B2:B3"),
    }};
    const cells = [_]CellInput{
        .{ .cell = cellAt(0, "A2"), .formula = "1" },
        .{ .cell = cellAt(0, "A3"), .formula = "2" },
    };
    var g = try buildOk(
        testing.allocator,
        .{ .sheet_count = 2, .cells = &cells, .producers = &producers },
        &world,
        .{},
    );
    defer g.deinit();

    const p: Key = .{ .producer = .{
        .table = "T",
        .column = "Total",
        .kind = .calculated_column,
        .index = 0,
    } };
    const qty: Key = .{ .range = areaAt(0, "A2:A3") };
    try testing.expect(hasEdge(g, p, qty));
    try testing.expect(hasEdge(g, qty, .{ .cell = cellAt(0, "A2") }));
    try testing.expect(hasEdge(g, qty, .{ .cell = cellAt(0, "A3") }));
}

test "nodes: a dead branch is still an edge (§5.3a)" {
    const cells = [_]CellInput{
        .{ .cell = cellAt(0, "A1"), .formula = "IF(TRUE,B1,C1)" },
        .{ .cell = cellAt(0, "B1"), .formula = "1" },
        .{ .cell = cellAt(0, "C1"), .formula = "2" },
    };
    var g = try buildOk(testing.allocator, .{ .sheet_count = 2, .cells = &cells }, &two_sheets, .{});
    defer g.deinit();

    try testing.expect(hasEdge(g, .{ .cell = cellAt(0, "A1") }, .{ .cell = cellAt(0, "B1") }));
    try testing.expect(hasEdge(g, .{ .cell = cellAt(0, "A1") }, .{ .cell = cellAt(0, "C1") }));
}

// ─── order and determinism (§5.6b) ───────────────────────────────

fn describeKey(buf: []u8, k: Key) []const u8 {
    var cb: [16]u8 = undefined;
    return switch (k) {
        .cell => |c| std.fmt.bufPrint(buf, "cell S{d}!{s}", .{
            c.sheet.toInt(),
            coords.formatCell(&cb, .{ .row = c.row, .col = c.col }),
        }) catch buf[0..0],
        .spill_tail => |c| std.fmt.bufPrint(buf, "tail S{d}!{s}", .{
            c.sheet.toInt(),
            coords.formatCell(&cb, .{ .row = c.row, .col = c.col }),
        }) catch buf[0..0],
        .range => |x| std.fmt.bufPrint(buf, "range S{d}!R{d}C{d}:R{d}C{d}", .{
            x.sheet.toInt(),
            x.range.first.row.oneBased(),
            x.range.first.col.zeroBased(),
            x.range.last.row.oneBased(),
            x.range.last.col.zeroBased(),
        }) catch buf[0..0],
        .span => |x| std.fmt.bufPrint(buf, "span S{d}..S{d}!R{d}C{d}:R{d}C{d}", .{
            x.first.toInt(),
            x.last.toInt(),
            x.range.first.row.oneBased(),
            x.range.first.col.zeroBased(),
            x.range.last.row.oneBased(),
            x.range.last.col.zeroBased(),
        }) catch buf[0..0],
        .name => |x| std.fmt.bufPrint(buf, "name {s}", .{x.identifier}) catch buf[0..0],
        .producer => |x| std.fmt.bufPrint(buf, "producer {s}[{s}]", .{ x.table, x.column }) catch buf[0..0],
    };
}

fn orderedKeys(gpa: std.mem.Allocator, g: Graph) ![]Key {
    const flat = try g.flatOrder(gpa);
    defer gpa.free(flat);
    const out = try gpa.alloc(Key, flat.len);
    for (flat, 0..) |n, i| out[i] = g.keys[n];
    return out;
}

test "order: dependencies come first, and the tie-break is (sheet, row, col)" {
    const cells = [_]CellInput{
        .{ .cell = cellAt(0, "D1"), .formula = "A1" },
        .{ .cell = cellAt(0, "C1"), .formula = "B1" },
        .{ .cell = cellAt(0, "B1"), .formula = "1" },
        .{ .cell = cellAt(0, "A1"), .formula = "1" },
    };
    var g = try buildOk(testing.allocator, .{ .sheet_count = 2, .cells = &cells }, &two_sheets, .{});
    defer g.deinit();

    const got = try orderedKeys(testing.allocator, g);
    defer testing.allocator.free(got);

    // Both roots are ready at once; the smaller key goes first. Then
    // each dependent becomes ready and the same rule applies.
    const want = [_]Key{
        .{ .cell = cellAt(0, "A1") },
        .{ .cell = cellAt(0, "B1") },
        .{ .cell = cellAt(0, "C1") },
        .{ .cell = cellAt(0, "D1") },
    };
    try testing.expectEqual(want.len, got.len);
    for (want, got) |w, x| try testing.expect(w.eql(x));
}

test "order: a chain orders bottom-up whatever the input order" {
    var storage: [24]CellInput = undefined;
    // A1 = A2, A2 = A3, ... A23 = A24, A24 = 1.
    for (0..24) |i| {
        storage[i] = .{ .cell = cellAt(0, chainA1(i)), .formula = chainBody(i) };
    }
    var g = try buildOk(testing.allocator, .{ .sheet_count = 2, .cells = &storage }, &two_sheets, .{});
    defer g.deinit();

    const got = try orderedKeys(testing.allocator, g);
    defer testing.allocator.free(got);
    try testing.expectEqual(@as(usize, 24), got.len);
    // Deepest first: A24 has nothing to wait for, A1 waits for all of them.
    for (got, 0..) |k, i| {
        try testing.expectEqual(@as(u32, @intCast(24 - i)), k.cell.row.oneBased());
    }
}

const chain_a1 = blk: {
    var out: [24][:0]const u8 = undefined;
    for (0..24) |i| {
        out[i] = std.fmt.comptimePrint("A{d}", .{i + 1});
    }
    break :blk out;
};
const chain_bodies = blk: {
    var out: [24][:0]const u8 = undefined;
    for (0..24) |i| {
        out[i] = if (i == 23) "1" else std.fmt.comptimePrint("A{d}", .{i + 2});
    }
    break :blk out;
};

fn chainA1(i: usize) []const u8 {
    return chain_a1[i];
}

fn chainBody(i: usize) []const u8 {
    return chain_bodies[i];
}

test "determinism: repeated builds and randomized insertion order agree" {
    var cells = [_]CellInput{
        .{ .cell = cellAt(0, "A1"), .formula = "SUM(B1:B4)+Sheet2!$C$1" },
        .{ .cell = cellAt(0, "B1"), .formula = "1" },
        .{ .cell = cellAt(0, "B2"), .formula = "B1" },
        .{ .cell = cellAt(0, "B3"), .formula = "SUM(B1:B2)" },
        .{ .cell = cellAt(0, "B4"), .formula = "B3" },
        .{ .cell = cellAt(1, "C1"), .formula = "SUM(A:A)" },
        .{ .cell = cellAt(1, "A7"), .formula = "1" },
    };

    var first = try buildOk(testing.allocator, .{ .sheet_count = 2, .cells = &cells }, &two_sheets, .{});
    defer first.deinit();
    const want = try orderedKeys(testing.allocator, first);
    defer testing.allocator.free(want);

    var prng = std.Random.DefaultPrng.init(0x5EED_1234);
    for (0..32) |_| {
        prng.random().shuffle(CellInput, &cells);
        var g = try buildOk(testing.allocator, .{ .sheet_count = 2, .cells = &cells }, &two_sheets, .{});
        defer g.deinit();
        const got = try orderedKeys(testing.allocator, g);
        defer testing.allocator.free(got);
        try testing.expectEqual(want.len, got.len);
        for (want, got) |a, b| {
            if (!a.eql(b)) {
                var ba: [64]u8 = undefined;
                var bb: [64]u8 = undefined;
                std.debug.print("order diverged: {s} vs {s}\n", .{ describeKey(&ba, a), describeKey(&bb, b) });
                return error.OrderNotDeterministic;
            }
        }
    }
}

test "order: a range's dependencies are area → sheet → row-major, whatever the layer" {
    var cells = [_]CellInput{
        .{ .cell = cellAt(0, "Z1"), .formula = "SUM(A1:C3)" },
        .{ .cell = cellAt(0, "C3"), .formula = "1" },
        .{ .cell = cellAt(0, "A1"), .formula = "1" },
        .{ .cell = cellAt(0, "B2"), .formula = "1" },
        .{ .cell = cellAt(0, "A3"), .formula = "1" },
        .{ .cell = cellAt(0, "C1"), .formula = "1" },
    };
    var prng = std.Random.DefaultPrng.init(0xA11CE);
    prng.random().shuffle(CellInput, &cells);

    var g = try buildOk(testing.allocator, .{ .sheet_count = 2, .cells = &cells }, &two_sheets, .{});
    defer g.deinit();

    const r = g.find(.{ .range = areaAt(0, "A1:C3") }).?;
    const want = [_][]const u8{ "A1", "C1", "B2", "A3", "C3" };
    try testing.expectEqual(want.len, g.deps[r].len);
    for (want, g.deps[r]) |w, node| {
        try testing.expect(g.keys[node].eql(.{ .cell = cellAt(0, w) }));
    }
}

// ─── the seed table (§5.6c), one fixture per row ─────────────────

fn seedOfCycle(gpa: std.mem.Allocator, cell: CellInput) !Seed {
    const cells = [_]CellInput{cell};
    var g = try buildOk(gpa, .{ .sheet_count = 2, .cells = &cells }, &two_sheets, .{});
    defer g.deinit();
    const n = g.find(.{ .cell = cell.cell }).?;
    try testing.expect(g.isCyclic(n));
    return g.seeds[n].?;
}

test "seed table row 1: a numeric cache seeds its own value" {
    const seed = try seedOfCycle(testing.allocator, .{
        .cell = cellAt(0, "A1"),
        .formula = "A1+1",
        .cache = .{ .number = 7.5 },
    });
    try testing.expectEqual(@as(f64, 7.5), seed.number);
}

test "seed table row 2: text, boolean and error caches all seed zero" {
    for ([_]CacheState{ .text, .boolean, .err }) |cache| {
        const seed = try seedOfCycle(testing.allocator, .{
            .cell = cellAt(0, "A1"),
            .formula = "A1+1",
            .cache = cache,
        });
        try testing.expectEqual(@as(f64, 0), seed.number);
    }
}

test "seed table row 3: an absent <v> seeds zero" {
    const seed = try seedOfCycle(testing.allocator, .{
        .cell = cellAt(0, "A1"),
        .formula = "A1+1",
        .cache = .absent,
    });
    try testing.expectEqual(@as(f64, 0), seed.number);
}

test "seed table row 4: a malformed <v> is a typed refusal, and never a zero seed" {
    // The function itself: no seed comes back at all.
    try testing.expectError(error.MalformedCache, seedFor(.malformed, .none));
    try testing.expectError(error.MalformedCache, seedFor(.malformed, .{ .shape = .{ .rows = 2, .cols = 2 } }));
    try testing.expectError(error.MalformedCache, seedFor(.malformed, .unknown));

    // And the build: a refusal, with no graph — so there is no seed
    // table anywhere holding a zero for this cell.
    const cells = [_]CellInput{
        .{ .cell = cellAt(0, "A1"), .formula = "A1+1", .cache = .malformed },
    };
    const r = try buildRefused(testing.allocator, .{ .sheet_count = 2, .cells = &cells }, &two_sheets, .{});
    try testing.expectEqual(Refusal.Reason.malformed_cache_seed, r.reason);
    try testing.expect(r.at.?.eql(.{ .cell = cellAt(0, "A1") }));
}

test "seed table row 5: an anchor with a declared shape seeds a zero-filled array" {
    const seed = try seedOfCycle(testing.allocator, .{
        .cell = cellAt(0, "A1"),
        .formula = "A1+1",
        .cache = .{ .number = 7.5 },
        .array = rangeOf("A1:C2"),
    });
    // The declared shape wins over the cache: an array cell resumes as
    // an array, not as the one number its top-left happened to hold.
    try testing.expectEqual(@as(u32, 2), seed.array_zeros.rows);
    try testing.expectEqual(@as(u32, 3), seed.array_zeros.cols);
}

test "seed table row 5b: an anchor with no recoverable shape takes the shape pass" {
    const seed = try seedOfCycle(testing.allocator, .{
        .cell = cellAt(0, "A1"),
        .formula = "A1+1",
        .cache = .{ .number = 7.5 },
        .dynamic_anchor = true,
    });
    try testing.expectEqual(Seed.shape_pass, seed);
}

test "seeds: only cyclic members carry one" {
    const cells = [_]CellInput{
        .{ .cell = cellAt(0, "A1"), .formula = "A1+B1", .cache = .{ .number = 3 } },
        .{ .cell = cellAt(0, "B1"), .formula = "1", .cache = .{ .number = 9 } },
    };
    var g = try buildOk(testing.allocator, .{ .sheet_count = 2, .cells = &cells }, &two_sheets, .{});
    defer g.deinit();
    try testing.expect(g.seeds[g.find(.{ .cell = cellAt(0, "A1") }).?] != null);
    // B1 is computed from its dependencies. There is nothing to resume.
    try testing.expectEqual(@as(?Seed, null), g.seeds[g.find(.{ .cell = cellAt(0, "B1") }).?]);
}

// ─── the sparseness contract (§5.6a) ─────────────────────────────

/// A sheet with `n` stored formula cells spread over `cols` columns,
/// plus one cell that reads a whole column or a whole row.
fn scatterInput(arena: std.mem.Allocator, rows: u32, cols: u32, reader: []const u8) !Input {
    var list: std.ArrayListUnmanaged(CellInput) = .empty;
    var r: u32 = 1;
    while (r <= rows) : (r += 1) {
        var c: u32 = 0;
        while (c < cols) : (c += 1) {
            try list.append(arena, .{
                .cell = .{
                    .sheet = env.SheetIndex.fromInt(0),
                    .row = try coords.Row.fromOneBased(r),
                    .col = try coords.Col.fromZeroBased(c),
                },
                .formula = "1",
            });
        }
    }
    try list.append(arena, .{
        .cell = cellAt(0, "Z100"),
        .formula = reader,
    });
    return .{ .sheet_count = 2, .cells = try list.toOwnedSlice(arena) };
}

test "scaling: a whole-column dependency is bounded by stored cells, not by 1 048 576" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const input = try scatterInput(arena.allocator(), 40, 12, "SUM(A:A)");

    var g = try buildOk(testing.allocator, input, &two_sheets, .{});
    defer g.deinit();

    // The instrument is a counter, not a stopwatch. 40 rows × 12 columns
    // are stored; only column A can contribute an edge, and the
    // column-major index is what keeps the probe count near 40 rather
    // than near a million.
    try testing.expectEqual(@as(u64, 40), g.stats.index_probes);
    try testing.expect(g.stats.index_probes < coords.max_row);
    try testing.expect(g.stats.index_probes <= input.cells.len);
    try testing.expectEqual(@as(u64, 40), g.deps[g.find(.{ .range = areaAt(0, "A1:A1048576") }).?].len);
}

test "scaling: a whole-row dependency is bounded the same way" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const input = try scatterInput(arena.allocator(), 40, 12, "SUM(3:3)");

    var g = try buildOk(testing.allocator, input, &two_sheets, .{});
    defer g.deinit();

    try testing.expectEqual(@as(u64, 12), g.stats.index_probes);
    try testing.expect(g.stats.index_probes < coords.max_col_1based);
}

// ─── closure planning (§5.6f) ────────────────────────────────────

const Planned = struct {
    arena: std.heap.ArenaAllocator,
    graph: Graph,
    plan: Plan,
    counters: WorkCounters,

    fn deinit(self: *Planned) void {
        self.arena.deinit();
        self.graph.deinit();
    }
};

const PlanOutcome = union(enum) { ok: Planned, refused: Refusal };

fn planFrom(
    gpa: std.mem.Allocator,
    input: Input,
    world: *const World,
    roots: []const Key,
    limits: WorkLimits,
) !PlanOutcome {
    var g = try buildOk(gpa, input, world, .{});
    errdefer g.deinit();
    var arena = std.heap.ArenaAllocator.init(gpa);
    errdefer arena.deinit();
    var counters: WorkCounters = .{ .limits = limits };
    return switch (try plan(g, arena.allocator(), roots, &counters, .{})) {
        .ok => |p| .{ .ok = .{ .arena = arena, .graph = g, .plan = p, .counters = counters } },
        .refused => |r| {
            arena.deinit();
            g.deinit();
            return .{ .refused = r };
        },
    };
}

test "closure: the plan is the transitive closure, ordered, cells only" {
    const cells = [_]CellInput{
        .{ .cell = cellAt(0, "A1"), .formula = "SUM(B1:B2)" },
        .{ .cell = cellAt(0, "B1"), .formula = "C1" },
        .{ .cell = cellAt(0, "B2"), .formula = "1" },
        .{ .cell = cellAt(0, "C1"), .formula = "1" },
        // Reachable from nothing the root asks for.
        .{ .cell = cellAt(0, "Z9"), .formula = "1" },
    };
    var out = switch (try planFrom(
        testing.allocator,
        .{ .sheet_count = 2, .cells = &cells },
        &two_sheets,
        &.{.{ .cell = cellAt(0, "A1") }},
        .{},
    )) {
        .ok => |p| p,
        .refused => return error.UnexpectedRefusal,
    };
    defer out.deinit();

    // C1 and B2 are both ready at the start; C1 sorts first (row 1
    // before row 2), which frees B1 before B2's turn comes up.
    const want = [_][]const u8{ "C1", "B1", "B2", "A1" };
    try testing.expectEqual(want.len, out.plan.cells.len);
    for (want, out.plan.cells) |w, node| {
        try testing.expect(out.graph.keys[node].eql(.{ .cell = cellAt(0, w) }));
    }
    try testing.expectEqual(@as(u64, 4), out.counters.usedBy(.total_cell_evals));
    // Depth unwinds completely.
    try testing.expectEqual(@as(u64, 0), out.counters.usedBy(.eval_depth));
}

test "closure: a cycle refuses while iteration is off (§5.6c)" {
    const cells = [_]CellInput{
        .{ .cell = cellAt(0, "A1"), .formula = "B1" },
        .{ .cell = cellAt(0, "B1"), .formula = "A1" },
    };
    const outcome = try planFrom(
        testing.allocator,
        .{ .sheet_count = 2, .cells = &cells },
        &two_sheets,
        &.{.{ .cell = cellAt(0, "A1") }},
        .{},
    );
    switch (outcome) {
        .ok => |p| {
            var q = p;
            q.deinit();
            return error.UnexpectedSuccess;
        },
        .refused => |r| try testing.expectEqual(Refusal.Reason.cycle, r.reason),
    }
}

test "closure: a spill tail orders but does not evaluate" {
    const cells = [_]CellInput{
        .{ .cell = cellAt(0, "D1"), .formula = "A2" },
        .{ .cell = cellAt(0, "A1"), .formula = "1", .array = rangeOf("A1:A2") },
    };
    var out = switch (try planFrom(
        testing.allocator,
        .{ .sheet_count = 2, .cells = &cells },
        &two_sheets,
        &.{.{ .cell = cellAt(0, "D1") }},
        .{},
    )) {
        .ok => |p| p,
        .refused => return error.UnexpectedRefusal,
    };
    defer out.deinit();

    // A1 and D1 evaluate; the tail at A2 is in the closure and in the
    // order, and is not a unit of evaluation.
    try testing.expectEqual(@as(usize, 2), out.plan.cells.len);
    try testing.expect(out.graph.keys[out.plan.cells[0]].eql(.{ .cell = cellAt(0, "A1") }));
    try testing.expect(out.graph.keys[out.plan.cells[1]].eql(.{ .cell = cellAt(0, "D1") }));
}

// ─── §9 counters, below / at / above at their charge sites ───────

test "§9 dependency_edges: below, at, above — charged in addEdge" {
    const cells = [_]CellInput{
        .{ .cell = cellAt(0, "D1"), .formula = "SUM(A1:B2)" },
        .{ .cell = cellAt(0, "A1"), .formula = "1" },
        .{ .cell = cellAt(0, "B2"), .formula = "2" },
    };
    const input: Input = .{ .sheet_count = 2, .cells = &cells };

    var g = try buildOk(testing.allocator, input, &two_sheets, .{});
    const edges = g.stats.edges;
    g.deinit();
    // D1 → range, range → A1, range → B2.
    try testing.expectEqual(@as(u64, 3), edges);

    // Below the count: refused, and the category is named.
    const r = try buildRefused(testing.allocator, input, &two_sheets, .{
        .limits = .{ .max_dependency_edges = edges - 1 },
    });
    try testing.expectEqual(Refusal.Reason.work_limit_exceeded, r.reason);
    try testing.expectEqual(@as(?WorkCategory, .dependency_edges), r.limit);

    // Exactly the count, and one above: both build.
    for ([_]u64{ edges, edges + 1 }) |limit| {
        var ok = try buildOk(testing.allocator, input, &two_sheets, .{
            .limits = .{ .max_dependency_edges = limit },
        });
        defer ok.deinit();
        try testing.expectEqual(edges, ok.stats.edges);
    }
}

test "§9 dependency_edges: a duplicate edge is not charged twice" {
    // `A1+A1` mentions B1 twice; one edge, one charge.
    const cells = [_]CellInput{
        .{ .cell = cellAt(0, "A1"), .formula = "B1+B1" },
        .{ .cell = cellAt(0, "B1"), .formula = "1" },
    };
    var g = try buildOk(testing.allocator, .{ .sheet_count = 2, .cells = &cells }, &two_sheets, .{});
    defer g.deinit();
    try testing.expectEqual(@as(u64, 1), g.stats.edges);
}

test "§9 max_total_cell_evals: below, at, above — charged when the plan admits a cell" {
    var storage: [6]CellInput = undefined;
    for (0..6) |i| storage[i] = .{ .cell = cellAt(0, chainA1(i)), .formula = chainBody6(i) };
    const input: Input = .{ .sheet_count = 2, .cells = &storage };
    const root: []const Key = &.{.{ .cell = cellAt(0, "A1") }};

    for ([_]u64{ 6, 7 }) |limit| {
        var out = switch (try planFrom(testing.allocator, input, &two_sheets, root, .{
            .max_total_cell_evals = limit,
        })) {
            .ok => |p| p,
            .refused => return error.UnexpectedRefusal,
        };
        defer out.deinit();
        try testing.expectEqual(@as(usize, 6), out.plan.cells.len);
        try testing.expectEqual(@as(u64, 6), out.counters.usedBy(.total_cell_evals));
    }

    switch (try planFrom(testing.allocator, input, &two_sheets, root, .{ .max_total_cell_evals = 5 })) {
        .ok => |p| {
            var q = p;
            q.deinit();
            return error.UnexpectedSuccess;
        },
        .refused => |r| {
            try testing.expectEqual(Refusal.Reason.work_limit_exceeded, r.reason);
            try testing.expectEqual(@as(?WorkCategory, .total_cell_evals), r.limit);
        },
    }
}

const chain6_bodies = blk: {
    var out: [6][:0]const u8 = undefined;
    for (0..6) |i| out[i] = if (i == 5) "1" else std.fmt.comptimePrint("A{d}", .{i + 2});
    break :blk out;
};

fn chainBody6(i: usize) []const u8 {
    return chain6_bodies[i];
}

test "§9 max_eval_depth: below, at, above — charged on the closure walk" {
    var storage: [6]CellInput = undefined;
    for (0..6) |i| storage[i] = .{ .cell = cellAt(0, chainA1(i)), .formula = chainBody6(i) };
    const input: Input = .{ .sheet_count = 2, .cells = &storage };
    const root: []const Key = &.{.{ .cell = cellAt(0, "A1") }};

    for ([_]u64{ 6, 7 }) |limit| {
        var out = switch (try planFrom(testing.allocator, input, &two_sheets, root, .{
            .max_eval_depth = limit,
        })) {
            .ok => |p| p,
            .refused => return error.UnexpectedRefusal,
        };
        defer out.deinit();
        // Six cells on the stack at once, and every one released again.
        try testing.expectEqual(@as(u64, 6), out.counters.peakOf(.eval_depth));
        try testing.expectEqual(@as(u64, 0), out.counters.usedBy(.eval_depth));
    }

    switch (try planFrom(testing.allocator, input, &two_sheets, root, .{ .max_eval_depth = 5 })) {
        .ok => |p| {
            var q = p;
            q.deinit();
            return error.UnexpectedSuccess;
        },
        .refused => |r| {
            try testing.expectEqual(Refusal.Reason.work_limit_exceeded, r.reason);
            try testing.expectEqual(@as(?WorkCategory, .eval_depth), r.limit);
        },
    }
}

test "§9 max_eval_depth: range and name nodes do not consume a cell's depth" {
    // A1 → range(B1:B2) → B1 → C1. Three cells deep, four nodes.
    const cells = [_]CellInput{
        .{ .cell = cellAt(0, "A1"), .formula = "SUM(B1:B2)" },
        .{ .cell = cellAt(0, "B1"), .formula = "C1" },
        .{ .cell = cellAt(0, "B2"), .formula = "1" },
        .{ .cell = cellAt(0, "C1"), .formula = "1" },
    };
    var out = switch (try planFrom(
        testing.allocator,
        .{ .sheet_count = 2, .cells = &cells },
        &two_sheets,
        &.{.{ .cell = cellAt(0, "A1") }},
        .{ .max_eval_depth = 3 },
    )) {
        .ok => |p| p,
        .refused => return error.UnexpectedRefusal,
    };
    defer out.deinit();
    try testing.expectEqual(@as(u64, 3), out.counters.peakOf(.eval_depth));
}

// ─── the differential gate ───────────────────────────────────────
//
// A missed edge is invisible to every other test in this file, so the
// gate for this row is a second builder that shares no code with the
// first and a generator that provably reaches the five shapes an edge
// goes missing in.
//
// The generator emits both the formula **text** and the references it
// intended. The real builder parses the text; the brute-force builder
// reads the intentions and never sees an AST. Agreement over a thousand
// workbooks therefore pins two things at once: that the walk recovers
// exactly the references that were written, and that the sparse index
// finds exactly the producers a naive scan finds.

const SpecRef = union(enum) {
    cell: env.CellRef,
    area: env.RangeRef,
    name: u32,
    span: Key.Span,
};

const Spec = struct {
    input: Input,
    world: World,
    cell_refs: []const []const SpecRef,
    name_refs: []const []const SpecRef,
};

const Shapes = struct {
    overlapping_ranges: u32 = 0,
    full_row_or_col: u32 = 0,
    three_d_spans: u32 = 0,
    defined_names: u32 = 0,
    spill_resize: u32 = 0,
};

const gen_sheets = [_][]const u8{ "Sheet1", "Sheet2", "Sheet3" };
const gen_rows: u32 = 6;
const gen_cols: u32 = 4;

const Gen = struct {
    a: std.mem.Allocator,
    rand: std.Random,

    fn colName(self: *Gen, col: u32) []const u8 {
        var buf: [4]u8 = undefined;
        const n = coords.writeColLetters(&buf, coords.Col.fromZeroBased(col) catch unreachable);
        return self.a.dupe(u8, buf[0..n]) catch unreachable;
    }

    fn cellText(self: *Gen, row: u32, col: u32) []const u8 {
        return std.fmt.allocPrint(self.a, "${s}${d}", .{ self.colName(col), row }) catch unreachable;
    }

    fn ref(self: *Gen, sheet: u32, row: u32, col: u32) env.CellRef {
        _ = self;
        return .{
            .sheet = env.SheetIndex.fromInt(sheet),
            .row = coords.Row.fromOneBased(row) catch unreachable,
            .col = coords.Col.fromZeroBased(col) catch unreachable,
        };
    }

    fn area(self: *Gen, sheet: u32, r1: u32, c1: u32, r2: u32, c2: u32) env.RangeRef {
        _ = self;
        return .{
            .sheet = env.SheetIndex.fromInt(sheet),
            .range = (coords.Range{
                .first = .{
                    .row = coords.Row.fromOneBased(@min(r1, r2)) catch unreachable,
                    .col = coords.Col.fromZeroBased(@min(c1, c2)) catch unreachable,
                },
                .last = .{
                    .row = coords.Row.fromOneBased(@max(r1, r2)) catch unreachable,
                    .col = coords.Col.fromZeroBased(@max(c1, c2)) catch unreachable,
                },
            }),
        };
    }

    const Term = struct { text: []const u8, ref: SpecRef };

    /// One reference term. `qualified` forces the sheet-qualified,
    /// fully-absolute spellings a defined-name body is allowed to use
    /// (§5.9 refuses a relative body when it is referenced).
    fn term(self: *Gen, owner_sheet: u32, qualified: bool, allow_span: bool, name_count: u32) Term {
        // 0 cell · 1 area · 2 full column · 3 full row · 4 3D span · 5 name.
        var kind = self.rand.uintLessThan(u32, 6);
        if (kind == 4 and !allow_span) kind = 1;
        if (kind == 5 and name_count == 0) kind = 0;
        const sheet: u32 = if (qualified)
            self.rand.uintLessThan(u32, gen_sheets.len)
        else
            owner_sheet;
        const prefix: []const u8 = if (qualified)
            std.fmt.allocPrint(self.a, "{s}!", .{gen_sheets[sheet]}) catch unreachable
        else
            "";

        switch (kind) {
            0 => {
                const r = 1 + self.rand.uintLessThan(u32, gen_rows);
                const c = self.rand.uintLessThan(u32, gen_cols);
                return .{
                    .text = std.fmt.allocPrint(self.a, "{s}{s}", .{ prefix, self.cellText(r, c) }) catch unreachable,
                    .ref = .{ .cell = self.ref(sheet, r, c) },
                };
            },
            1 => {
                // A multi-cell area. Single-cell areas are cell
                // references by §5.6a's normalization, and the generator
                // does not emit the ambiguous spelling.
                const r1 = 1 + self.rand.uintLessThan(u32, gen_rows - 1);
                const r2 = r1 + 1 + self.rand.uintLessThan(u32, gen_rows - r1);
                const c1 = self.rand.uintLessThan(u32, gen_cols);
                const c2 = c1 + self.rand.uintLessThan(u32, gen_cols - c1);
                return .{
                    .text = std.fmt.allocPrint(self.a, "SUM({s}{s}:{s})", .{
                        prefix,
                        self.cellText(r1, c1),
                        self.cellText(r2, c2),
                    }) catch unreachable,
                    .ref = .{ .area = self.area(sheet, r1, c1, r2, c2) },
                };
            },
            2 => {
                const c1 = self.rand.uintLessThan(u32, gen_cols);
                const c2 = c1 + self.rand.uintLessThan(u32, gen_cols - c1);
                return .{
                    .text = std.fmt.allocPrint(self.a, "SUM({s}${s}:${s})", .{
                        prefix,
                        self.colName(c1),
                        self.colName(c2),
                    }) catch unreachable,
                    .ref = .{ .area = self.area(sheet, 1, c1, coords.max_row, c2) },
                };
            },
            3 => {
                const r1 = 1 + self.rand.uintLessThan(u32, gen_rows);
                const r2 = r1 + self.rand.uintLessThan(u32, gen_rows + 1 - r1);
                return .{
                    .text = std.fmt.allocPrint(self.a, "SUM({s}${d}:${d})", .{ prefix, r1, r2 }) catch unreachable,
                    .ref = .{ .area = self.area(sheet, r1, 0, r2, coords.max_col_1based - 1) },
                };
            },
            4 => {
                const first = self.rand.uintLessThan(u32, gen_sheets.len);
                const last = first + self.rand.uintLessThan(u32, @as(u32, gen_sheets.len) - first);
                const r1 = 1 + self.rand.uintLessThan(u32, gen_rows - 1);
                const r2 = r1 + 1 + self.rand.uintLessThan(u32, gen_rows - r1);
                const c1 = self.rand.uintLessThan(u32, gen_cols);
                const c2 = c1 + self.rand.uintLessThan(u32, gen_cols - c1);
                const a = self.area(0, r1, c1, r2, c2);
                return .{
                    .text = std.fmt.allocPrint(self.a, "SUM({s}:{s}!{s}:{s})", .{
                        gen_sheets[first],
                        gen_sheets[last],
                        self.cellText(@min(r1, r2), @min(c1, c2)),
                        self.cellText(@max(r1, r2), @max(c1, c2)),
                    }) catch unreachable,
                    .ref = .{ .span = .{
                        .first = env.SheetIndex.fromInt(first),
                        .last = env.SheetIndex.fromInt(last),
                        .range = a.range,
                    } },
                };
            },
            else => {
                const k = self.rand.uintLessThan(u32, name_count);
                return .{
                    // The underscore matters: `N1` is column N, row 1 —
                    // a cell reference — and a generator that used it
                    // would be testing the wrong grammar production.
                    .text = std.fmt.allocPrint(self.a, "Name_{d}", .{k}) catch unreachable,
                    .ref = .{ .name = k },
                };
            },
        }
    }

    const Body = struct { text: []const u8, refs: []const SpecRef };

    fn body(self: *Gen, owner_sheet: u32, qualified: bool, allow_span: bool, name_count: u32) Body {
        const n = self.rand.uintLessThan(u32, 4);
        if (n == 0) return .{ .text = "1", .refs = &.{} };
        var text: std.ArrayListUnmanaged(u8) = .empty;
        var refs: std.ArrayListUnmanaged(SpecRef) = .empty;
        for (0..n) |i| {
            if (i > 0) text.appendSlice(self.a, "+") catch unreachable;
            const t = self.term(owner_sheet, qualified, allow_span, name_count);
            text.appendSlice(self.a, t.text) catch unreachable;
            refs.append(self.a, t.ref) catch unreachable;
        }
        return .{ .text = text.items, .refs = refs.items };
    }

    fn workbook(self: *Gen, shapes: *Shapes) Spec {
        const name_count = self.rand.uintLessThan(u32, 3);
        var names: std.ArrayListUnmanaged(NameInput) = .empty;
        var name_refs: std.ArrayListUnmanaged([]const SpecRef) = .empty;
        for (0..name_count) |i| {
            // A name body sees no unqualified reference: §5.9 refuses a
            // relative body when it is referenced, so a generator that
            // wrote one would be testing a shape the graph declines to
            // model.
            const b = self.body(0, true, true, @intCast(i));
            names.append(self.a, .{
                .identifier = std.fmt.allocPrint(self.a, "Name_{d}", .{i}) catch unreachable,
                .body = b.text,
            }) catch unreachable;
            name_refs.append(self.a, b.refs) catch unreachable;
        }
        if (name_count > 0) shapes.defined_names += 1;

        const cell_count = 3 + self.rand.uintLessThan(u32, 8);
        var cells: std.ArrayListUnmanaged(CellInput) = .empty;
        var cell_refs: std.ArrayListUnmanaged([]const SpecRef) = .empty;
        var used: std.ArrayListUnmanaged(env.CellRef) = .empty;

        // At most one array anchor per workbook: two anchors whose
        // declared ranges overlap are an M7a obstruction, and the
        // generator stays inside what this row models.
        const anchor_at: ?u32 = if (self.rand.boolean()) self.rand.uintLessThan(u32, cell_count) else null;

        for (0..cell_count) |i| {
            const sheet = self.rand.uintLessThan(u32, gen_sheets.len);
            const row = 1 + self.rand.uintLessThan(u32, gen_rows);
            const col = self.rand.uintLessThan(u32, gen_cols);
            const c = self.ref(sheet, row, col);
            var dup = false;
            for (used.items) |u| {
                if (u.eql(c)) dup = true;
            }
            if (dup) continue;
            used.append(self.a, c) catch unreachable;

            const is_anchor = anchor_at != null and anchor_at.? == @as(u32, @intCast(i));
            const decl: ?coords.Range = if (is_anchor) blk: {
                const rows = 1 + self.rand.uintLessThan(u32, @min(3, gen_rows + 1 - row));
                const cols = 1 + self.rand.uintLessThan(u32, @min(2, gen_cols - col));
                break :blk self.area(sheet, row, col, row + rows - 1, col + cols - 1).range;
            } else null;

            // §5.6g forbids a 3D span inside an array formula.
            const b = self.body(sheet, self.rand.boolean(), !is_anchor, name_count);
            cells.append(self.a, .{ .cell = c, .formula = b.text, .array = decl }) catch unreachable;
            cell_refs.append(self.a, b.refs) catch unreachable;
        }

        var all: std.ArrayListUnmanaged(SpecRef) = .empty;
        for (cell_refs.items) |r| all.appendSlice(self.a, r) catch unreachable;
        for (name_refs.items) |r| all.appendSlice(self.a, r) catch unreachable;
        for (all.items) |r| switch (r) {
            .span => shapes.three_d_spans += 1,
            .area => |x| {
                if (x.range.rowCount() == coords.max_row or
                    x.range.colCount() == coords.max_col_1based)
                {
                    shapes.full_row_or_col += 1;
                }
            },
            else => {},
        };
        for (all.items, 0..) |x, i| {
            const xa = switch (x) {
                .area => |v| v,
                else => continue,
            };
            for (all.items[i + 1 ..]) |y| {
                const ya = switch (y) {
                    .area => |v| v,
                    else => continue,
                };
                if (xa.sheet != ya.sheet) continue;
                if (orderRange(xa.range, ya.range) == .eq) continue;
                if (xa.range.overlaps(ya.range)) shapes.overlapping_ranges += 1;
            }
        }

        return .{
            .input = .{
                .sheet_count = gen_sheets.len,
                .cells = cells.items,
                .names = names.items,
            },
            .world = .{ .sheets = &gen_sheets, .names = names.items },
            .cell_refs = cell_refs.items,
            .name_refs = name_refs.items,
        };
    }
};

// ─── the brute-force builder ─────────────────────────────────────

const Edge = struct { u: Key, v: Key };

fn edgeLess(_: void, x: Edge, y: Edge) bool {
    const o = x.u.order(y.u);
    if (o != .eq) return o == .lt;
    return x.v.order(y.v) == .lt;
}

const Brute = struct {
    keys: []Key,
    edges: []Edge,
};

/// Everything the real builder does, done the obvious way: linear
/// scans, no index, no Tarjan, and the references handed over rather
/// than parsed.
fn bruteForce(a: std.mem.Allocator, spec: Spec) !Brute {
    var keys: std.ArrayListUnmanaged(Key) = .empty;

    for (spec.input.cells) |c| try keys.append(a, .{ .cell = c.cell });
    for (spec.input.names, 0..) |n, i| try keys.append(a, .{ .name = .{
        .scope = n.scope,
        .identifier = n.identifier,
        .index = @intCast(i),
    } });

    // Tails: every coordinate a declared range covers that is not the
    // anchor and holds no formula of its own.
    var tails: std.ArrayListUnmanaged(struct { tail: env.CellRef, anchor: env.CellRef }) = .empty;
    for (spec.input.cells) |c| {
        const decl = (c.array orelse continue).normalized();
        var r = decl.first.row.oneBased();
        while (r <= decl.last.row.oneBased()) : (r += 1) {
            var col = decl.first.col.zeroBased();
            while (col <= decl.last.col.zeroBased()) : (col += 1) {
                const t: env.CellRef = .{
                    .sheet = c.cell.sheet,
                    .row = try coords.Row.fromOneBased(r),
                    .col = try coords.Col.fromZeroBased(col),
                };
                if (t.eql(c.cell)) continue;
                var occupied = false;
                for (spec.input.cells) |o| {
                    if (o.cell.eql(t)) occupied = true;
                }
                if (occupied) continue;
                try keys.append(a, .{ .spill_tail = t });
                try tails.append(a, .{ .tail = t, .anchor = c.cell });
            }
        }
    }

    const ref_lists = [_][]const []const SpecRef{ spec.cell_refs, spec.name_refs };
    for (ref_lists) |lists| {
        for (lists) |list| {
            for (list) |r| switch (r) {
                .area => |x| try keys.append(a, .{ .range = x }),
                .span => |sp| {
                    try keys.append(a, .{ .span = sp });
                    var sh = sp.first.toInt();
                    while (sh <= sp.last.toInt()) : (sh += 1) {
                        const member: env.RangeRef = .{
                            .sheet = env.SheetIndex.fromInt(sh),
                            .range = sp.range,
                        };
                        if (member.isSingleCell()) continue;
                        try keys.append(a, .{ .range = member });
                    }
                },
                else => {},
            };
        }
    }

    std.mem.sortUnstable(Key, keys.items, {}, keyLessThan);
    var deduped: std.ArrayListUnmanaged(Key) = .empty;
    for (keys.items, 0..) |k, i| {
        if (i > 0 and k.order(keys.items[i - 1]) == .eq) continue;
        try deduped.append(a, k);
    }

    const has = struct {
        fn f(list: []const Key, key: Key) bool {
            for (list) |k| {
                if (k.order(key) == .eq) return true;
            }
            return false;
        }
    }.f;
    const nodeAt = struct {
        fn f(list: []const Key, c: env.CellRef) ?Key {
            for (list) |k| {
                if (k == .cell and k.cell.eql(c)) return k;
            }
            for (list) |k| {
                if (k == .spill_tail and k.spill_tail.eql(c)) return k;
            }
            return null;
        }
    }.f;

    var edges: std.ArrayListUnmanaged(Edge) = .empty;
    const owners = [_]struct { keys: []Key, refs: []const []const SpecRef }{
        .{ .keys = try ownerKeysOfCells(a, spec), .refs = spec.cell_refs },
        .{ .keys = try ownerKeysOfNames(a, spec), .refs = spec.name_refs },
    };
    for (owners) |group| {
        for (group.keys, group.refs) |owner, list| {
            for (list) |r| {
                const target: ?Key = switch (r) {
                    .cell => |c| nodeAt(deduped.items, c),
                    .area => |x| Key{ .range = x },
                    .span => |sp| Key{ .span = sp },
                    .name => |k| Key{ .name = .{
                        .scope = spec.input.names[k].scope,
                        .identifier = spec.input.names[k].identifier,
                        .index = k,
                    } },
                };
                const t = target orelse continue;
                if (!has(deduped.items, t)) continue;
                try edges.append(a, .{ .u = owner, .v = t });
            }
        }
    }

    for (deduped.items) |k| switch (k) {
        .range => |x| {
            for (deduped.items) |other| {
                const c = other.coordinate() orelse continue;
                if (c.sheet != x.sheet) continue;
                if (!x.range.contains(.{ .row = c.row, .col = c.col })) continue;
                try edges.append(a, .{ .u = k, .v = other });
            }
        },
        .span => |sp| {
            var sh = sp.first.toInt();
            while (sh <= sp.last.toInt()) : (sh += 1) {
                const member: env.RangeRef = .{
                    .sheet = env.SheetIndex.fromInt(sh),
                    .range = sp.range,
                };
                const target: ?Key = if (member.isSingleCell())
                    nodeAt(deduped.items, member.topLeft())
                else
                    Key{ .range = member };
                const t = target orelse continue;
                if (!has(deduped.items, t)) continue;
                try edges.append(a, .{ .u = k, .v = t });
            }
        },
        .spill_tail => |t| {
            for (tails.items) |pair| {
                if (!pair.tail.eql(t)) continue;
                try edges.append(a, .{ .u = k, .v = .{ .cell = pair.anchor } });
                break;
            }
        },
        else => {},
    };

    std.mem.sortUnstable(Edge, edges.items, {}, edgeLess);
    var unique: std.ArrayListUnmanaged(Edge) = .empty;
    for (edges.items, 0..) |e, i| {
        if (i > 0 and !edgeLess({}, edges.items[i - 1], e)) continue;
        try unique.append(a, e);
    }
    return .{ .keys = deduped.items, .edges = unique.items };
}

fn ownerKeysOfCells(a: std.mem.Allocator, spec: Spec) ![]Key {
    const out = try a.alloc(Key, spec.input.cells.len);
    for (spec.input.cells, 0..) |c, i| out[i] = .{ .cell = c.cell };
    return out;
}

fn ownerKeysOfNames(a: std.mem.Allocator, spec: Spec) ![]Key {
    const out = try a.alloc(Key, spec.input.names.len);
    for (spec.input.names, 0..) |n, i| out[i] = .{ .name = .{
        .scope = n.scope,
        .identifier = n.identifier,
        .index = @intCast(i),
    } };
    return out;
}

fn realEdges(a: std.mem.Allocator, g: Graph) ![]Edge {
    var out: std.ArrayListUnmanaged(Edge) = .empty;
    for (g.deps, 0..) |list, u| {
        for (list) |v| try out.append(a, .{ .u = g.keys[u], .v = g.keys[v] });
    }
    std.mem.sortUnstable(Edge, out.items, {}, edgeLess);
    return out.items;
}

fn reportNodeDiff(want: []const Key, got: []const Key) void {
    var b: [64]u8 = undefined;
    std.debug.print("brute nodes:\n", .{});
    for (want) |k| std.debug.print("  {s}\n", .{describeKey(&b, k)});
    std.debug.print("graph nodes:\n", .{});
    for (got) |k| std.debug.print("  {s}\n", .{describeKey(&b, k)});
}

fn reportEdgeDiff(want: []const Edge, got: []const Edge) void {
    var bu: [64]u8 = undefined;
    var bv: [64]u8 = undefined;
    std.debug.print("\nbrute force ({d} edges):\n", .{want.len});
    for (want) |e| {
        std.debug.print("  {s} -> {s}\n", .{ describeKey(&bu, e.u), describeKey(&bv, e.v) });
    }
    std.debug.print("graph.zig ({d} edges):\n", .{got.len});
    for (got) |e| {
        std.debug.print("  {s} -> {s}\n", .{ describeKey(&bu, e.u), describeKey(&bv, e.v) });
    }
}

test "differential: graph.zig agrees with a brute-force builder over 1200 workbooks" {
    var shapes: Shapes = .{};
    var prng = std.Random.DefaultPrng.init(0x9E37_79B9_7F4A_7C15);

    var i: u32 = 0;
    while (i < 1200) : (i += 1) {
        var arena = std.heap.ArenaAllocator.init(testing.allocator);
        defer arena.deinit();
        const a = arena.allocator();

        var gen: Gen = .{ .a = a, .rand = prng.random() };
        const spec = gen.workbook(&shapes);

        var g = switch (try build(testing.allocator, spec.input, spec.world.resolver(), .{})) {
            .ok => |x| x,
            .refused => |r| {
                var bk: [64]u8 = undefined;
                std.debug.print("workbook {d} refused: {t} at {s}\n", .{
                    i,
                    r.reason,
                    if (r.at) |k| describeKey(&bk, k) else "?",
                });
                for (spec.input.cells) |c| {
                    var bc: [64]u8 = undefined;
                    std.debug.print("  cell {s} array={} = {s}\n", .{
                        describeKey(&bc, .{ .cell = c.cell }),
                        c.array != null,
                        c.formula,
                    });
                }
                for (spec.input.names) |n| std.debug.print("  name {s} = {s}\n", .{ n.identifier, n.body });
                return error.UnexpectedRefusal;
            },
        };
        defer g.deinit();

        const brute = try bruteForce(a, spec);

        if (brute.keys.len != g.keys.len) {
            std.debug.print("workbook {d}: {d} brute nodes vs {d} graph nodes\n", .{
                i,
                brute.keys.len,
                g.keys.len,
            });
            reportNodeDiff(brute.keys, g.keys);
            for (spec.input.cells) |c| {
                var bc: [64]u8 = undefined;
                std.debug.print("  cell {s} = {s}\n", .{ describeKey(&bc, .{ .cell = c.cell }), c.formula });
            }
            for (spec.input.names) |n| std.debug.print("  name {s} = {s}\n", .{ n.identifier, n.body });
            return error.NodeSetsDiffer;
        }
        for (brute.keys, g.keys) |x, y| {
            if (x.order(y) != .eq) {
                var bx: [64]u8 = undefined;
                var by: [64]u8 = undefined;
                std.debug.print("workbook {d}: node {s} vs {s}\n", .{
                    i,
                    describeKey(&bx, x),
                    describeKey(&by, y),
                });
                return error.NodeSetsDiffer;
            }
        }

        const got = try realEdges(a, g);
        if (got.len != brute.edges.len) {
            std.debug.print("workbook {d}: edge counts differ\n", .{i});
            reportEdgeDiff(brute.edges, got);
            return error.EdgeSetsDiffer;
        }
        for (brute.edges, got) |x, y| {
            if (x.u.order(y.u) == .eq and x.v.order(y.v) == .eq) continue;
            std.debug.print("workbook {d}: edge sets differ\n", .{i});
            reportEdgeDiff(brute.edges, got);
            return error.EdgeSetsDiffer;
        }

        // Spill resize / invalidation: the same workbook with one
        // anchor's declared shape changed must produce a different graph,
        // or nothing was invalidated.
        if (try resizedVariant(a, spec)) |resized| {
            var g2 = switch (try build(testing.allocator, resized, spec.world.resolver(), .{})) {
                .ok => |x| x,
                .refused => return error.UnexpectedRefusal,
            };
            defer g2.deinit();
            if (g2.keys.len != g.keys.len or g2.stats.edges != g.stats.edges) {
                shapes.spill_resize += 1;
            }
        }
    }

    // "Provably emits" — counters, not a hope about the seed. Silent on
    // success: a green run that prints is a green run someone has to
    // read to be sure it was green.
    const reached = [_]struct { name: []const u8, n: u32 }{
        .{ .name = "overlapping ranges", .n = shapes.overlapping_ranges },
        .{ .name = "full rows/columns", .n = shapes.full_row_or_col },
        .{ .name = "3D spans", .n = shapes.three_d_spans },
        .{ .name = "defined names", .n = shapes.defined_names },
        .{ .name = "spill resize/invalidation", .n = shapes.spill_resize },
    };
    for (reached) |r| {
        if (r.n > 0) continue;
        std.debug.print("generator never emitted: {s}\n", .{r.name});
        for (reached) |x| std.debug.print("  {s}: {d}\n", .{ x.name, x.n });
        return error.ShapeNeverGenerated;
    }
}

/// The same workbook with the first declared array one row taller, or
/// null when it has no anchor or no room to grow.
fn resizedVariant(a: std.mem.Allocator, spec: Spec) !?Input {
    for (spec.input.cells, 0..) |c, i| {
        const decl = (c.array orelse continue).normalized();
        const last = decl.last.row.oneBased();
        if (last >= gen_rows) return null;
        const cells = try a.dupe(CellInput, spec.input.cells);
        cells[i].array = .{
            .first = decl.first,
            .last = .{ .row = try coords.Row.fromOneBased(last + 1), .col = decl.last.col },
        };
        return Input{
            .sheet_count = spec.input.sheet_count,
            .cells = cells,
            .names = spec.input.names,
        };
    }
    return null;
}
