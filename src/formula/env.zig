//! `EvalEnv` — the merged logical view the evaluator reads through
//! (`goal_formula.md` §5.6a), and the in-memory fake M3a2 tests it with.
//!
//! M3a2 of the tier-D1 ladder. The evaluator never touches a workbook:
//! it reads cells, ranges, blank counts, and per-cell dialect through
//! this interface, and `pkg/workbook.zig` implements it at M4b1. That
//! split is not decoration — it is what keeps `src/formula/` free of any
//! `pkg/` import, and it is what makes the whole engine testable before
//! a single sheet has been parsed.
//!
//! What the interface promises
//! ---------------------------
//! **One ordered pass over a merged view.** `rangeIterator` and
//! `cellValue` read the same thing: stored cells, staged deltas, and
//! values computed earlier in the same run, with `computed > staged >
//! stored`. A stored-cells-only interface would make `C1=SUM(A1:B2)`
//! read stale inputs for a 2×2 computed at A1 in the same run — the bug
//! is invisible in a unit test and fatal in a workbook. Virtual spill
//! tails are the fourth layer and arrive with §5.8 at M7a; the `Layer`
//! enum leaves room for them rather than pretending three is the final
//! count.
//!
//! **Order is a property of the interface, not of the backing.**
//! Iteration is area → sheet → row-major regardless of insertion order
//! and regardless of which layer supplied a cell. `test "fake: iteration
//! order survives randomized insertion"` is the gate.
//!
//! **Sparse means proportional to occupancy.** Nothing here walks 1 048
//! 576 coordinates: iteration seeks by binary search and steps stored
//! entries, and `logicalBlankCount` subtracts an occupancy count from a
//! rectangle's area rather than testing cells one at a time. `A:A` is a
//! normal argument, not a pathological one.
//!
//! **Two blank classes, because Excel has two.** `.isblank_class` is
//! true blanks only; `.countblank_class` additionally counts a cell
//! whose value is `""`. `ISBLANK` is FALSE for `=""` while `COUNTBLANK`
//! counts it and `COUNTA` counts it too — one class would get one of the
//! three wrong.
//!
//! Allocation
//! ----------
//! Iteration allocates nothing. Both iterators carry their
//! implementation state inline in a fixed byte array whose size is a
//! compile-time bound (`RangeIterator.wrap` refuses to compile a state
//! that does not fit), and the N-way cursor takes its per-range scratch
//! from the caller. An evaluator that allocated per range would make
//! `SUMIFS` over a whole column an allocation storm.

const std = @import("std");
const assert = std.debug.assert;

const coords = @import("zlsx_refs");
const value = @import("value.zig");

// ─── typed coordinates (§5.5) ────────────────────────────────────

/// A sheet's position in the workbook's sheet order. Distinct type so it
/// cannot be swapped with a row, a column, or a raw index.
pub const SheetIndex = enum(u32) {
    _,

    pub fn fromInt(v: u32) SheetIndex {
        return @enumFromInt(v);
    }

    pub fn toInt(self: SheetIndex) u32 {
        return @intFromEnum(self);
    }
};

pub const CellRef = struct {
    sheet: SheetIndex,
    row: coords.Row,
    col: coords.Col,

    pub fn eql(a: CellRef, b: CellRef) bool {
        return a.sheet == b.sheet and a.row == b.row and a.col == b.col;
    }
};

/// A rectangular area on one sheet. `range` is always normalized.
pub const RangeRef = struct {
    sheet: SheetIndex,
    range: coords.Range,

    pub fn shape(self: RangeRef) value.Shape {
        return .{ .rows = self.range.rowCount(), .cols = self.range.colCount() };
    }

    pub fn cellCount(self: RangeRef) u64 {
        return @as(u64, self.range.rowCount()) * @as(u64, self.range.colCount());
    }

    pub fn isSingleCell(self: RangeRef) bool {
        return self.range.rowCount() == 1 and self.range.colCount() == 1;
    }

    pub fn topLeft(self: RangeRef) CellRef {
        return .{ .sheet = self.sheet, .row = self.range.first.row, .col = self.range.first.col };
    }

    /// The cell at a row-major offset into the area.
    pub fn cellAtOffset(self: RangeRef, offset: u64) CellRef {
        const cols = self.range.colCount();
        const r: u32 = @intCast(offset / cols);
        const c: u32 = @intCast(offset % cols);
        return .{
            .sheet = self.sheet,
            .row = coords.Row.fromOneBased(self.range.first.row.oneBased() + r) catch unreachable,
            .col = coords.Col.fromZeroBased(self.range.first.col.zeroBased() + c) catch unreachable,
        };
    }

    /// Row-major offset of a cell inside the area, or null if outside.
    pub fn offsetOf(self: RangeRef, row: coords.Row, col: coords.Col) ?u64 {
        if (row.oneBased() < self.range.first.row.oneBased()) return null;
        if (row.oneBased() > self.range.last.row.oneBased()) return null;
        if (col.zeroBased() < self.range.first.col.zeroBased()) return null;
        if (col.zeroBased() > self.range.last.col.zeroBased()) return null;
        const dr = row.oneBased() - self.range.first.row.oneBased();
        const dc = col.zeroBased() - self.range.first.col.zeroBased();
        return @as(u64, dr) * @as(u64, self.range.colCount()) + dc;
    }

    pub fn eql(a: RangeRef, b: RangeRef) bool {
        return a.sheet == b.sheet and
            a.range.first.eql(b.range.first) and
            a.range.last.eql(b.range.last);
    }
};

/// Where a value in the merged view came from. Precedence is the
/// declaration order: a later layer shadows every earlier one at the
/// same coordinate.
///
/// `.spill_tail` is declared but unused at M3a2 — cells materialized by
/// an already-evaluated dynamic-array anchor land there at M7a (§5.8a).
/// Naming it now keeps the precedence rule a single ordered enum rather
/// than a comparison someone has to re-derive.
pub const Layer = enum(u2) {
    stored = 0,
    staged = 1,
    spill_tail = 2,
    computed = 3,

    pub fn precedence(self: Layer) u2 {
        return @intFromEnum(self);
    }
};

/// Excel's own split, and the reason `logicalBlankCount` takes a class
/// rather than answering one question.
pub const BlankClass = enum {
    /// True blanks only. `ISBLANK` is FALSE for a cell holding `""`.
    isblank_class,
    /// True blanks **plus** cells whose value is the empty string.
    /// `COUNTBLANK` counts `=""`; `COUNTA` counts it too, as non-empty.
    countblank_class,
};

/// Everything an environment can fail with. Explicit and closed: an
/// adapter that needs a new failure adds it here, where every caller's
/// exhaustive switch will notice.
pub const Error = error{
    OutOfMemory,
    /// A `SheetIndex` with no sheet behind it.
    UnknownSheet,
    /// `.require_equal` alignment over ranges of differing dimensions.
    /// Callers turn this into `#VALUE!` (§5.6a).
    ShapeMismatch,
    /// A projected area would leave the grid (§5.6a criteria alignment).
    RefOutOfGrid,
    /// A `DialectResolver` refused the cell's `cm`/`vm` metadata (M4a).
    /// The typed refusal itself stays with the resolver, which owns the
    /// diagnostic; this interface only needs to know the read failed.
    MetadataRefused,
};

/// One occupied cell of the merged view.
pub const Entry = struct {
    row: coords.Row,
    col: coords.Col,
    value: value.ScalarValue,
    /// Which layer supplied it. Iteration order does not depend on this;
    /// it is exposed so a test can prove precedence was applied.
    layer: Layer,
};

// ─── iterators (allocation-free) ─────────────────────────────────

/// Inline storage for an iterator implementation's state. A real bound:
/// `wrap` refuses at compile time rather than silently heap-allocating.
pub const iterator_state_bytes: usize = 128;
pub const iterator_state_align: usize = 16;

/// An ordered sparse pass over one area. Yields only occupied cells, in
/// row-major order, one entry per coordinate (the highest-precedence
/// layer wins).
pub const RangeIterator = struct {
    state: [iterator_state_bytes]u8 align(iterator_state_align) = undefined,
    next_fn: *const fn (self: *RangeIterator) Error!?Entry,

    pub fn next(self: *RangeIterator) Error!?Entry {
        return self.next_fn(self);
    }

    /// Store `s` inline. The state is copied bytewise, so it must not
    /// contain a pointer to itself; pointers *into* the environment are
    /// fine and are what implementations actually keep.
    pub fn wrap(
        comptime S: type,
        s: S,
        next_fn: *const fn (self: *RangeIterator) Error!?Entry,
    ) RangeIterator {
        comptime assertFits(S, "RangeIterator");
        var it: RangeIterator = .{ .next_fn = next_fn };
        const slot: *S = @ptrCast(@alignCast(&it.state));
        slot.* = s;
        return it;
    }

    pub fn stateOf(self: *RangeIterator, comptime S: type) *S {
        comptime assertFits(S, "RangeIterator");
        return @ptrCast(@alignCast(&self.state));
    }
};

fn assertFits(comptime S: type, comptime who: []const u8) void {
    if (@sizeOf(S) > iterator_state_bytes) {
        @compileError(who ++ " state " ++ @typeName(S) ++ " exceeds iterator_state_bytes");
    }
    if (@alignOf(S) > iterator_state_align) {
        @compileError(who ++ " state " ++ @typeName(S) ++ " over-aligned for inline storage");
    }
}

/// How an N-way cursor lines its areas up (§5.6a criteria alignment).
pub const AlignMode = enum {
    /// `*IFS`: every area must have identical dimensions, else
    /// `ShapeMismatch` — which the caller reports as `#VALUE!`.
    require_equal,
    /// `SUMIF`/`AVERAGEIF`: areas after the first are **projected** from
    /// their own top-left using the first area's dimensions. Excel's
    /// documented rule; the areas need not be the same size as written.
    project_from_first,
};

/// One step of an N-way aligned pass. Blank positions arrive as *runs*
/// rather than one item each — that is what keeps `COUNTIFS(A:A,…)`
/// proportional to occupancy instead of to 1 048 576.
pub const AlignedItem = union(enum) {
    /// A position where at least one area is occupied. `values` has one
    /// slot per area, in the order the areas were given; an area with
    /// nothing there contributes `.blank`.
    cells: struct {
        row_offset: u32,
        col_offset: u32,
        values: []const value.ScalarValue,
    },
    /// `count` consecutive row-major positions, starting at the given
    /// offset, where **every** area is blank.
    blank_run: struct {
        row_offset: u32,
        col_offset: u32,
        count: u64,
    },
};

/// The N-way sparse cursor. One ordered pass over all areas at once —
/// deliberately not repeated pairwise zips, which would re-walk the
/// aggregation range once per criteria pair.
pub const AlignedIterator = struct {
    state: [iterator_state_bytes]u8 align(iterator_state_align) = undefined,
    /// Caller-owned scratch, one slot per area. Owned by the caller so
    /// the iterator itself never allocates.
    cursors: []usize,
    out: []value.ScalarValue,
    next_fn: *const fn (self: *AlignedIterator) Error!?AlignedItem,

    pub fn next(self: *AlignedIterator) Error!?AlignedItem {
        return self.next_fn(self);
    }

    pub fn wrap(
        comptime S: type,
        s: S,
        cursors: []usize,
        out: []value.ScalarValue,
        next_fn: *const fn (self: *AlignedIterator) Error!?AlignedItem,
    ) AlignedIterator {
        comptime assertFits(S, "AlignedIterator");
        var it: AlignedIterator = .{ .cursors = cursors, .out = out, .next_fn = next_fn };
        const slot: *S = @ptrCast(@alignCast(&it.state));
        slot.* = s;
        return it;
    }

    pub fn stateOf(self: *AlignedIterator, comptime S: type) *S {
        comptime assertFits(S, "AlignedIterator");
        return @ptrCast(@alignCast(&self.state));
    }
};

// ─── cm/vm-derived dialect (M4a) ─────────────────────────────────

/// How a stored cell's `cm`/`vm` metadata indexes become a dialect.
///
/// A function pointer rather than an import: `src/formula/metadata.zig`
/// owns the `xl/metadata.xml` reader and every refusal it can raise, and
/// routing the answer through this seam is what keeps `env.zig` a leaf —
/// the evaluator's window onto a workbook must not grow a dependency on
/// the part that happens to answer one of its questions.
pub const DialectResolver = struct {
    ctx: *anyopaque,
    /// `cm`/`vm` exactly as the cell carried them; `0` means "absent",
    /// which is `CT_Cell`'s own default. A cell whose metadata cannot be
    /// interpreted yields `error.MetadataRefused` — never a guessed
    /// dialect.
    resolve: *const fn (ctx: *anyopaque, cm: u32, vm: u32) Error!value.Dialect,

    pub fn dialectOf(self: DialectResolver, cm: u32, vm: u32) Error!value.Dialect {
        return self.resolve(self.ctx, cm, vm);
    }
};

// ─── the interface ───────────────────────────────────────────────

/// The evaluator's whole window onto a workbook. A vtable rather than a
/// generic parameter because two implementations ship (the fake here,
/// the `pkg/workbook.zig` adapter at M4b1) and the evaluator must not be
/// instantiated twice.
pub const EvalEnv = struct {
    ctx: *anyopaque,
    vtable: *const VTable,

    pub const VTable = struct {
        cellValue: *const fn (ctx: *anyopaque, cell: CellRef) Error!value.ScalarValue,
        rangeIterator: *const fn (ctx: *anyopaque, area: RangeRef) Error!RangeIterator,
        logicalBlankCount: *const fn (ctx: *anyopaque, area: RangeRef, class: BlankClass) Error!u64,
        alignedRangeIterator: *const fn (
            ctx: *anyopaque,
            areas: []const RangeRef,
            mode: AlignMode,
            cursors: []usize,
            out: []value.ScalarValue,
        ) Error!AlignedIterator,
        /// Dialect is a **stored-cell property** (§5.3b): recalc asks
        /// per cell, standalone eval passes its own and never calls this.
        dialectOf: *const fn (ctx: *anyopaque, cell: CellRef) Error!value.Dialect,
        /// The area a dynamic-array anchor occupies, or null if the cell
        /// is not an anchor — `A1#` against a non-anchor is `#REF!`.
        spillShape: *const fn (ctx: *anyopaque, cell: CellRef) Error!?value.Shape,
        resolveSheet: *const fn (ctx: *anyopaque, name: []const u8) Error!?SheetIndex,
    };

    pub fn cellValue(self: EvalEnv, cell: CellRef) Error!value.ScalarValue {
        return self.vtable.cellValue(self.ctx, cell);
    }

    pub fn rangeIterator(self: EvalEnv, area: RangeRef) Error!RangeIterator {
        return self.vtable.rangeIterator(self.ctx, area);
    }

    pub fn logicalBlankCount(self: EvalEnv, area: RangeRef, class: BlankClass) Error!u64 {
        return self.vtable.logicalBlankCount(self.ctx, area, class);
    }

    pub fn alignedRangeIterator(
        self: EvalEnv,
        areas: []const RangeRef,
        mode: AlignMode,
        cursors: []usize,
        out: []value.ScalarValue,
    ) Error!AlignedIterator {
        assert(cursors.len == areas.len and out.len == areas.len);
        return self.vtable.alignedRangeIterator(self.ctx, areas, mode, cursors, out);
    }

    pub fn dialectOf(self: EvalEnv, cell: CellRef) Error!value.Dialect {
        return self.vtable.dialectOf(self.ctx, cell);
    }

    pub fn spillShape(self: EvalEnv, cell: CellRef) Error!?value.Shape {
        return self.vtable.spillShape(self.ctx, cell);
    }

    pub fn resolveSheet(self: EvalEnv, name: []const u8) Error!?SheetIndex {
        return self.vtable.resolveSheet(self.ctx, name);
    }
};

/// The area a criteria/aggregation pair actually covers under `mode`.
/// Split out because both the iterator and its up-front validation need
/// exactly the same answer, and two copies of a projection rule is one
/// too many.
pub fn effectiveArea(area: RangeRef, dims: value.Shape, mode: AlignMode) Error!RangeRef {
    switch (mode) {
        .require_equal => {
            const s = area.shape();
            if (!s.eql(dims)) return error.ShapeMismatch;
            return area;
        },
        .project_from_first => {
            const first = area.range.first;
            const last_row = first.row.oneBased() + dims.rows - 1;
            const last_col = first.col.zeroBased() + dims.cols - 1;
            const row = coords.Row.fromOneBased(last_row) catch return error.RefOutOfGrid;
            const col = coords.Col.fromZeroBased(last_col) catch return error.RefOutOfGrid;
            return .{
                .sheet = area.sheet,
                .range = .{
                    .first = first,
                    .last = .{ .col = col, .row = row, .anchor = area.range.last.anchor },
                },
            };
        },
    }
}

// ─── in-memory fake (M3a2) ───────────────────────────────────────

/// The reference implementation of everything above.
///
/// It is a *fake*, not a mock: it implements the merge, the ordering,
/// and the sparseness for real, so a test that passes against it is
/// evidence about the contract rather than about the test double. M4b1's
/// adapter is checked against the same suite.
///
/// Storage is one array per sheet, kept sorted by (row, col, layer
/// descending) on insert. Sorting on insert rather than on read is what
/// makes ordered iteration independent of insertion order *by
/// construction* — there is no code path that could return cells in
/// backing order, because there is no backing order.
pub const Fake = struct {
    allocator: std.mem.Allocator,
    sheets: std.ArrayListUnmanaged(Sheet) = .empty,
    /// When set, `dialectOf` derives its answer from each cell's
    /// `cm`/`vm` instead of the stored `dialect` field — the M4a path a
    /// real workbook takes. Unset, the explicit field stands, so every
    /// test written before the metadata reader existed still means what
    /// it said.
    dialect_resolver: ?DialectResolver = null,

    pub const Cell = struct {
        row: coords.Row,
        col: coords.Col,
        /// Set by `put` from its own argument; the field exists so the
        /// sort key and the merged read can both see it.
        layer: Layer = .stored,
        v: value.ScalarValue,
        /// `.legacy` marks a cell authored before dynamic arrays;
        /// `dialectOf` reports it (§5.3b's stored-cell property).
        dialect: value.Dialect = .dynamic_array,
        /// `c@cm` / `c@vm`, one-based, `0` = absent. Read only when a
        /// `dialect_resolver` is attached.
        cm: u32 = 0,
        vm: u32 = 0,
        /// Set on a dynamic-array anchor; `A1#` resolves through it.
        spill: ?value.Shape = null,
    };

    pub const Sheet = struct {
        name: []const u8,
        cells: std.ArrayListUnmanaged(Cell) = .empty,
    };

    pub fn init(allocator: std.mem.Allocator) Fake {
        return .{ .allocator = allocator };
    }

    pub fn deinit(self: *Fake) void {
        for (self.sheets.items) |*s| s.cells.deinit(self.allocator);
        self.sheets.deinit(self.allocator);
        self.* = undefined;
    }

    pub fn addSheet(self: *Fake, name: []const u8) Error!SheetIndex {
        const idx: u32 = @intCast(self.sheets.items.len);
        try self.sheets.append(self.allocator, .{ .name = name });
        return SheetIndex.fromInt(idx);
    }

    /// Insert or replace one cell in one layer. A stored `.blank` is
    /// rejected: blank is the *absence* of a cell in the merged view, and
    /// letting it be stored would give two spellings for one state and
    /// break every blank count.
    pub fn put(self: *Fake, sheet: SheetIndex, layer: Layer, cell: Cell) Error!void {
        assert(cell.v != .blank);
        const s = try self.sheetMut(sheet);
        var c = cell;
        c.layer = layer;
        const at = lowerBound(s.cells.items, c.row, c.col, layer);
        if (at < s.cells.items.len) {
            const cur = s.cells.items[at];
            if (cur.row == c.row and cur.col == c.col and cur.layer == layer) {
                s.cells.items[at] = c;
                return;
            }
        }
        try s.cells.insert(self.allocator, at, c);
    }

    /// `put` with A1 text, for tests. Any coordinate the grid accepts.
    pub fn putA1(self: *Fake, sheet: SheetIndex, layer: Layer, a1: []const u8, v: value.ScalarValue) !void {
        const c = try coords.parseCell(a1, .{ .dollar = .accept });
        try self.put(sheet, layer, .{ .row = c.row, .col = c.col, .layer = layer, .v = v });
    }

    fn sheetMut(self: *Fake, idx: SheetIndex) Error!*Sheet {
        const i = idx.toInt();
        if (i >= self.sheets.items.len) return error.UnknownSheet;
        return &self.sheets.items[i];
    }

    fn sheetConst(self: *const Fake, idx: SheetIndex) Error!*const Sheet {
        const i = idx.toInt();
        if (i >= self.sheets.items.len) return error.UnknownSheet;
        return &self.sheets.items[i];
    }

    pub fn evalEnv(self: *Fake) EvalEnv {
        return .{ .ctx = self, .vtable = &vtable };
    }

    const vtable: EvalEnv.VTable = .{
        .cellValue = vtCellValue,
        .rangeIterator = vtRangeIterator,
        .logicalBlankCount = vtLogicalBlankCount,
        .alignedRangeIterator = vtAlignedRangeIterator,
        .dialectOf = vtDialectOf,
        .spillShape = vtSpillShape,
        .resolveSheet = vtResolveSheet,
    };

    fn selfOf(ctx: *anyopaque) *Fake {
        return @ptrCast(@alignCast(ctx));
    }

    /// The merged read: highest-precedence layer at a coordinate, or
    /// null when nothing occupies it. Every accessor goes through this,
    /// which is why `cellValue` and `rangeIterator` cannot disagree.
    fn merged(s: *const Sheet, row: coords.Row, col: coords.Col) ?Cell {
        // Entries are sorted layer-DESCENDING within a coordinate, so
        // the lower bound *is* the highest-precedence entry. Precedence
        // costs a comparison, not a scan.
        const lo = lowerBound(s.cells.items, row, col, .computed);
        if (lo >= s.cells.items.len) return null;
        const c = s.cells.items[lo];
        if (c.row != row or c.col != col) return null;
        return c;
    }

    fn vtCellValue(ctx: *anyopaque, cell: CellRef) Error!value.ScalarValue {
        const self = selfOf(ctx);
        const s = try self.sheetConst(cell.sheet);
        const m = merged(s, cell.row, cell.col) orelse return .blank;
        return m.v;
    }

    fn vtDialectOf(ctx: *anyopaque, cell: CellRef) Error!value.Dialect {
        const self = selfOf(ctx);
        const s = try self.sheetConst(cell.sheet);
        if (self.dialect_resolver) |r| {
            // An unoccupied coordinate carries no metadata, which is
            // exactly what `cm = 0` means — the resolver answers it the
            // same way it answers an unmarked stored cell.
            const m = merged(s, cell.row, cell.col) orelse
                return r.dialectOf(0, 0);
            return r.dialectOf(m.cm, m.vm);
        }
        const m = merged(s, cell.row, cell.col) orelse return .dynamic_array;
        return m.dialect;
    }

    fn vtSpillShape(ctx: *anyopaque, cell: CellRef) Error!?value.Shape {
        const self = selfOf(ctx);
        const s = try self.sheetConst(cell.sheet);
        const m = merged(s, cell.row, cell.col) orelse return null;
        return m.spill;
    }

    fn vtResolveSheet(ctx: *anyopaque, name: []const u8) Error!?SheetIndex {
        const self = selfOf(ctx);
        for (self.sheets.items, 0..) |s, i| {
            if (std.mem.eql(u8, s.name, name)) return SheetIndex.fromInt(@intCast(i));
        }
        return null;
    }

    const RangeState = struct {
        sheet: *const Sheet,
        area: RangeRef,
        idx: usize,
    };

    fn vtRangeIterator(ctx: *anyopaque, area: RangeRef) Error!RangeIterator {
        const self = selfOf(ctx);
        const s = try self.sheetConst(area.sheet);
        // Seek straight to the first entry that could be in the area.
        // Without this, `A1000000:A1000001` on a busy sheet would walk
        // every earlier row to find two cells.
        const start = lowerBound(s.cells.items, area.range.first.row, area.range.first.col, .computed);
        return RangeIterator.wrap(
            RangeState,
            .{ .sheet = s, .area = area, .idx = start },
            rangeNext,
        );
    }

    fn rangeNext(it: *RangeIterator) Error!?Entry {
        const st = it.stateOf(RangeState);
        const items = st.sheet.cells.items;
        while (st.idx < items.len) {
            const c = items[st.idx];
            if (c.row.oneBased() > st.area.range.last.row.oneBased()) return null;
            st.idx += 1;
            if (st.area.offsetOf(c.row, c.col) == null) continue;
            // Entries are layer-descending within a coordinate, so the
            // first one seen is the winner and the rest are shadowed.
            while (st.idx < items.len and
                items[st.idx].row == c.row and
                items[st.idx].col == c.col) : (st.idx += 1)
            {}
            return .{ .row = c.row, .col = c.col, .value = c.v, .layer = c.layer };
        }
        return null;
    }

    fn vtLogicalBlankCount(ctx: *anyopaque, area: RangeRef, class: BlankClass) Error!u64 {
        const self = selfOf(ctx);
        var it = try vtRangeIterator(self, area);
        var occupied: u64 = 0;
        while (try it.next()) |e| {
            const counts_as_blank = switch (class) {
                .isblank_class => false,
                .countblank_class => e.value == .text and e.value.text.len == 0,
            };
            if (!counts_as_blank) occupied += 1;
        }
        // Area minus occupancy: no per-coordinate test, so a whole
        // column costs what its stored cells cost.
        return area.cellCount() - occupied;
    }

    const AlignedState = struct {
        fake: *Fake,
        areas: []const RangeRef,
        mode: AlignMode,
        dims: value.Shape,
        offset: u64,
        total: u64,
    };

    fn vtAlignedRangeIterator(
        ctx: *anyopaque,
        areas: []const RangeRef,
        mode: AlignMode,
        cursors: []usize,
        out: []value.ScalarValue,
    ) Error!AlignedIterator {
        const self = selfOf(ctx);
        assert(areas.len > 0);
        assert(cursors.len == areas.len and out.len == areas.len);
        const dims = areas[0].shape();
        // Validate every area up front. A cursor that discovered a shape
        // mismatch halfway would have already reported rows the caller
        // must now un-count.
        for (areas) |a| {
            _ = try effectiveArea(a, dims, mode);
            _ = try self.sheetConst(a.sheet);
        }
        @memset(cursors, std.math.maxInt(usize));
        return AlignedIterator.wrap(
            AlignedState,
            .{
                .fake = self,
                .areas = areas,
                .mode = mode,
                .dims = dims,
                .offset = 0,
                .total = @as(u64, dims.rows) * @as(u64, dims.cols),
            },
            cursors,
            out,
            alignedNext,
        );
    }

    fn alignedNext(it: *AlignedIterator) Error!?AlignedItem {
        const st = it.stateOf(AlignedState);
        if (st.offset >= st.total) return null;

        // The next offset at which *any* area is occupied. Everything
        // strictly before it is a blank run for every area at once.
        var next_occupied: u64 = st.total;
        for (st.areas, it.cursors) |a, *cursor| {
            const eff = try effectiveArea(a, st.dims, st.mode);
            const s = try st.fake.sheetConst(eff.sheet);
            if (cursor.* == std.math.maxInt(usize)) {
                cursor.* = lowerBound(
                    s.cells.items,
                    eff.range.first.row,
                    eff.range.first.col,
                    .computed,
                );
            }
            const off = advance(s, eff, cursor, st.offset) orelse continue;
            next_occupied = @min(next_occupied, off);
        }

        const cols = st.dims.cols;
        const at = st.offset;
        if (next_occupied > at) {
            const run = next_occupied - at;
            st.offset = next_occupied;
            return .{ .blank_run = .{
                .row_offset = @intCast(at / cols),
                .col_offset = @intCast(at % cols),
                .count = run,
            } };
        }

        for (st.areas, it.out) |a, *slot| {
            const eff = try effectiveArea(a, st.dims, st.mode);
            const s = try st.fake.sheetConst(eff.sheet);
            const cell = eff.cellAtOffset(at);
            slot.* = if (merged(s, cell.row, cell.col)) |m| m.v else .blank;
        }
        st.offset = at + 1;
        return .{ .cells = .{
            .row_offset = @intCast(at / cols),
            .col_offset = @intCast(at % cols),
            .values = it.out,
        } };
    }

    /// Move one area's cursor to the first occupied offset ≥ `from`,
    /// returning it. Forward-only: offsets within an area increase with
    /// (row, col), which is exactly the order the backing is sorted in.
    fn advance(s: *const Sheet, eff: RangeRef, cursor: *usize, from: u64) ?u64 {
        const items = s.cells.items;
        while (cursor.* < items.len) {
            const c = items[cursor.*];
            if (c.row.oneBased() > eff.range.last.row.oneBased()) return null;
            const off = eff.offsetOf(c.row, c.col) orelse {
                cursor.* += 1;
                continue;
            };
            if (off < from) {
                cursor.* += 1;
                continue;
            }
            return off;
        }
        return null;
    }

    /// First index whose (row, col, layer-desc) key is ≥ the given one.
    fn lowerBound(items: []const Cell, row: coords.Row, col: coords.Col, layer: Layer) usize {
        var lo: usize = 0;
        var hi: usize = items.len;
        while (lo < hi) {
            const mid = lo + (hi - lo) / 2;
            if (lessThanKey(items[mid], row, col, layer)) lo = mid + 1 else hi = mid;
        }
        return lo;
    }

    fn lessThanKey(c: Cell, row: coords.Row, col: coords.Col, layer: Layer) bool {
        if (c.row.oneBased() != row.oneBased()) return c.row.oneBased() < row.oneBased();
        if (c.col.zeroBased() != col.zeroBased()) return c.col.zeroBased() < col.zeroBased();
        // Layer descending: the highest-precedence entry for a
        // coordinate sorts first, so a merged read is the first hit.
        return c.layer.precedence() > layer.precedence();
    }
};

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

fn cellOf(a1: []const u8) coords.Cell {
    return coords.parseCell(a1, .{ .dollar = .accept }) catch unreachable;
}

fn areaOf(sheet: SheetIndex, a1: []const u8) RangeRef {
    const r = coords.parseRange(a1, .{ .dollar = .accept }) catch unreachable;
    return .{ .sheet = sheet, .range = r.normalized() };
}

fn text(s: []const u8) value.ScalarValue {
    return .{ .text = s };
}

fn num(v: f64) value.ScalarValue {
    return value.ScalarValue.fromNumber(v);
}

test "fake: ordered sparse iteration yields only occupied cells, row-major" {
    var fake = Fake.init(testing.allocator);
    defer fake.deinit();
    const sh = try fake.addSheet("Sheet1");

    try fake.putA1(sh, .stored, "B2", num(2));
    try fake.putA1(sh, .stored, "A1", num(1));
    try fake.putA1(sh, .stored, "C1", num(3));

    var it = try fake.evalEnv().rangeIterator(areaOf(sh, "A1:C3"));
    var seen: [8][]const u8 = undefined;
    var n: usize = 0;
    var buf: [8][16]u8 = undefined;
    while (try it.next()) |e| {
        seen[n] = coords.formatCell(&buf[n], .{ .col = e.col, .row = e.row });
        n += 1;
    }
    try testing.expectEqual(@as(usize, 3), n);
    try testing.expectEqualStrings("A1", seen[0]);
    try testing.expectEqualStrings("C1", seen[1]);
    try testing.expectEqualStrings("B2", seen[2]);
}

test "fake: iteration order survives randomized insertion" {
    // Order is a property of the interface, not of the backing. The same
    // 40 cells, inserted in eight different orders across three layers,
    // must come back in one order every time — strictly ascending
    // row-major, one entry per coordinate.
    var prng = std.Random.DefaultPrng.init(0x5eed_5eed);
    const rnd = prng.random();

    var round: usize = 0;
    while (round < 8) : (round += 1) {
        var order: [40]u32 = undefined;
        for (&order, 0..) |*o, i| o.* = @intCast(i);
        rnd.shuffle(u32, &order);

        var fake = Fake.init(testing.allocator);
        defer fake.deinit();
        const sh = try fake.addSheet("S");
        for (order) |i| {
            const layer: Layer = switch (i % 3) {
                0 => .stored,
                1 => .staged,
                else => .computed,
            };
            try fake.put(sh, layer, .{
                .row = try coords.Row.fromOneBased((i / 5) + 1),
                .col = try coords.Col.fromZeroBased(i % 5),
                .v = num(@floatFromInt(i)),
            });
        }

        var it = try fake.evalEnv().rangeIterator(areaOf(sh, "A1:E8"));
        var n: usize = 0;
        var last_key: u64 = 0;
        while (try it.next()) |e| {
            const key = (@as(u64, e.row.oneBased()) << 32) | e.col.zeroBased();
            try testing.expect(key > last_key);
            last_key = key;
            // Row-major position must equal the insertion index, whatever
            // order the inserts arrived in.
            try testing.expect(value.ScalarValue.eql(num(@floatFromInt(n)), e.value));
            n += 1;
        }
        try testing.expectEqual(@as(usize, 40), n);
    }
}

test "fake: computed shadows staged shadows stored, in one merged view" {
    var fake = Fake.init(testing.allocator);
    defer fake.deinit();
    const sh = try fake.addSheet("S");
    const c = cellOf("A1");

    try fake.put(sh, .stored, .{ .row = c.row, .col = c.col, .layer = .stored, .v = num(1) });
    const e = fake.evalEnv();
    try testing.expect(value.ScalarValue.eql(num(1), try e.cellValue(.{ .sheet = sh, .row = c.row, .col = c.col })));

    try fake.put(sh, .staged, .{ .row = c.row, .col = c.col, .layer = .staged, .v = num(2) });
    try testing.expect(value.ScalarValue.eql(num(2), try e.cellValue(.{ .sheet = sh, .row = c.row, .col = c.col })));

    try fake.put(sh, .computed, .{ .row = c.row, .col = c.col, .layer = .computed, .v = num(3) });
    try testing.expect(value.ScalarValue.eql(num(3), try e.cellValue(.{ .sheet = sh, .row = c.row, .col = c.col })));

    // And the iterator agrees — one entry, the winning layer.
    var it = try e.rangeIterator(areaOf(sh, "A1:A1"));
    const first = (try it.next()).?;
    try testing.expectEqual(Layer.computed, first.layer);
    try testing.expect(value.ScalarValue.eql(num(3), first.value));
    try testing.expectEqual(@as(?Entry, null), try it.next());
}

test "fake: logicalBlankCount splits ISBLANK from COUNTBLANK" {
    var fake = Fake.init(testing.allocator);
    defer fake.deinit();
    const sh = try fake.addSheet("S");

    try fake.putA1(sh, .stored, "A1", num(0)); // a zero is not blank
    try fake.putA1(sh, .stored, "A2", text("")); // `=""` splits the classes
    try fake.putA1(sh, .stored, "A3", value.ScalarValue.errorOf(.na));
    // A4 left truly blank.

    const e = fake.evalEnv();
    const area = areaOf(sh, "A1:A4");
    // ISBLANK sees one blank: A4.
    try testing.expectEqual(@as(u64, 1), try e.logicalBlankCount(area, .isblank_class));
    // COUNTBLANK sees two: A4 and the `""` in A2.
    try testing.expectEqual(@as(u64, 2), try e.logicalBlankCount(area, .countblank_class));
}

test "fake: logicalBlankCount over a whole column stays proportional to occupancy" {
    var fake = Fake.init(testing.allocator);
    defer fake.deinit();
    const sh = try fake.addSheet("S");
    try fake.putA1(sh, .stored, "A7", num(7));

    const e = fake.evalEnv();
    const whole = RangeRef{ .sheet = sh, .range = .{
        .first = .{ .col = try coords.Col.fromZeroBased(0), .row = try coords.Row.fromOneBased(1) },
        .last = .{ .col = try coords.Col.fromZeroBased(0), .row = try coords.Row.fromOneBased(coords.max_row) },
    } };
    try testing.expectEqual(
        @as(u64, coords.max_row - 1),
        try e.logicalBlankCount(whole, .isblank_class),
    );
}

test "fake: aligned cursor over a single-cell range" {
    var fake = Fake.init(testing.allocator);
    defer fake.deinit();
    const sh = try fake.addSheet("S");
    try fake.putA1(sh, .stored, "A1", num(5));
    try fake.putA1(sh, .stored, "B1", num(6));

    const areas = [_]RangeRef{ areaOf(sh, "A1:A1"), areaOf(sh, "B1:B1") };
    var cursors: [2]usize = undefined;
    var out: [2]value.ScalarValue = undefined;
    var it = try fake.evalEnv().alignedRangeIterator(&areas, .require_equal, &cursors, &out);

    const item = (try it.next()).?;
    try testing.expect(item == .cells);
    try testing.expectEqual(@as(u32, 0), item.cells.row_offset);
    try testing.expect(value.ScalarValue.eql(num(5), item.cells.values[0]));
    try testing.expect(value.ScalarValue.eql(num(6), item.cells.values[1]));
    try testing.expectEqual(@as(?AlignedItem, null), try it.next());
}

test "fake: aligned cursor over an empty range is one blank run" {
    var fake = Fake.init(testing.allocator);
    defer fake.deinit();
    const sh = try fake.addSheet("S");

    const areas = [_]RangeRef{ areaOf(sh, "A1:B3"), areaOf(sh, "D1:E3") };
    var cursors: [2]usize = undefined;
    var out: [2]value.ScalarValue = undefined;
    var it = try fake.evalEnv().alignedRangeIterator(&areas, .require_equal, &cursors, &out);

    const item = (try it.next()).?;
    try testing.expect(item == .blank_run);
    // Six positions, delivered as one run rather than six items.
    try testing.expectEqual(@as(u64, 6), item.blank_run.count);
    try testing.expectEqual(@as(?AlignedItem, null), try it.next());
}

test "fake: aligned cursor interleaves runs and occupied positions" {
    var fake = Fake.init(testing.allocator);
    defer fake.deinit();
    const sh = try fake.addSheet("S");
    // A1:A4 criteria, B1:B4 values; only rows 2 and 4 are occupied.
    try fake.putA1(sh, .stored, "A2", text("x"));
    try fake.putA1(sh, .stored, "B4", num(9));

    const areas = [_]RangeRef{ areaOf(sh, "A1:A4"), areaOf(sh, "B1:B4") };
    var cursors: [2]usize = undefined;
    var out: [2]value.ScalarValue = undefined;
    var it = try fake.evalEnv().alignedRangeIterator(&areas, .require_equal, &cursors, &out);

    var runs: usize = 0;
    var occupied: usize = 0;
    var covered: u64 = 0;
    while (try it.next()) |item| switch (item) {
        .blank_run => |r| {
            runs += 1;
            covered += r.count;
        },
        .cells => {
            occupied += 1;
            covered += 1;
        },
    };
    try testing.expectEqual(@as(usize, 2), occupied);
    try testing.expectEqual(@as(usize, 2), runs);
    // Every position accounted for exactly once — the invariant that
    // makes a run-based cursor safe for counting aggregates.
    try testing.expectEqual(@as(u64, 4), covered);
}

test "fake: require_equal rejects mismatched dimensions, project_from_first rebases" {
    var fake = Fake.init(testing.allocator);
    defer fake.deinit();
    const sh = try fake.addSheet("S");
    try fake.putA1(sh, .stored, "C2", num(42));

    const areas = [_]RangeRef{ areaOf(sh, "A1:A2"), areaOf(sh, "C1:C1") };
    var cursors: [2]usize = undefined;
    var out: [2]value.ScalarValue = undefined;

    try testing.expectError(
        error.ShapeMismatch,
        fake.evalEnv().alignedRangeIterator(&areas, .require_equal, &cursors, &out),
    );

    // Projected, `C1:C1` covers C1:C2 — so C2's 42 lines up with A2.
    var it = try fake.evalEnv().alignedRangeIterator(&areas, .project_from_first, &cursors, &out);
    var found: ?f64 = null;
    while (try it.next()) |item| switch (item) {
        .cells => |c| {
            if (c.values[1] == .number) found = c.values[1].number;
            try testing.expectEqual(@as(u32, 1), c.row_offset);
        },
        .blank_run => {},
    };
    try testing.expectEqual(@as(?f64, 42), found);
}

test "fake: projection that would leave the grid is refused, not clamped" {
    var fake = Fake.init(testing.allocator);
    defer fake.deinit();
    const sh = try fake.addSheet("S");
    const last_row = try coords.Row.fromOneBased(coords.max_row);
    const col0 = try coords.Col.fromZeroBased(0);
    const col1 = try coords.Col.fromZeroBased(1);

    const areas = [_]RangeRef{
        .{ .sheet = sh, .range = .{
            .first = .{ .col = col0, .row = try coords.Row.fromOneBased(1) },
            .last = .{ .col = col0, .row = try coords.Row.fromOneBased(3) },
        } },
        .{ .sheet = sh, .range = .{
            .first = .{ .col = col1, .row = last_row },
            .last = .{ .col = col1, .row = last_row },
        } },
    };
    var cursors: [2]usize = undefined;
    var out: [2]value.ScalarValue = undefined;
    try testing.expectError(
        error.RefOutOfGrid,
        fake.evalEnv().alignedRangeIterator(&areas, .project_from_first, &cursors, &out),
    );
}

test "fake: dialect, spill shape, and sheet resolution are per-cell properties" {
    var fake = Fake.init(testing.allocator);
    defer fake.deinit();
    const sh = try fake.addSheet("Data");
    const a1 = cellOf("A1");
    const b1 = cellOf("B1");

    try fake.put(sh, .stored, .{
        .row = a1.row,
        .col = a1.col,
        .layer = .stored,
        .v = num(1),
        .dialect = .legacy,
    });
    try fake.put(sh, .stored, .{
        .row = b1.row,
        .col = b1.col,
        .layer = .stored,
        .v = num(2),
        .spill = .{ .rows = 3, .cols = 1 },
    });

    const e = fake.evalEnv();
    try testing.expectEqual(value.Dialect.legacy, try e.dialectOf(.{ .sheet = sh, .row = a1.row, .col = a1.col }));
    try testing.expectEqual(value.Dialect.dynamic_array, try e.dialectOf(.{ .sheet = sh, .row = b1.row, .col = b1.col }));
    try testing.expectEqual(@as(?value.Shape, null), try e.spillShape(.{ .sheet = sh, .row = a1.row, .col = a1.col }));
    try testing.expectEqual(@as(?value.Shape, .{ .rows = 3, .cols = 1 }), try e.spillShape(.{ .sheet = sh, .row = b1.row, .col = b1.col }));
    try testing.expectEqual(@as(?SheetIndex, sh), try e.resolveSheet("Data"));
    try testing.expectEqual(@as(?SheetIndex, null), try e.resolveSheet("Missing"));
}

test "fake: a dialect resolver takes over dialectOf, refusals included" {
    // The stub stands in for M4a's metadata reader: `cm = 1` is a
    // dynamic-array mark, `vm` of any kind refuses. The semantics are
    // proven against the real reader in `metadata.zig`; what this test
    // owns is the seam — that `dialectOf` asks the resolver at all, that
    // it passes the cell's own `cm`/`vm`, and that a refusal surfaces as
    // an error rather than as a guessed dialect.
    const Stub = struct {
        calls: u32 = 0,
        fn resolve(ctx: *anyopaque, cm: u32, vm: u32) Error!value.Dialect {
            const self: *@This() = @ptrCast(@alignCast(ctx));
            self.calls += 1;
            if (vm != 0) return error.MetadataRefused;
            return if (cm == 1) .dynamic_array else .legacy;
        }
    };

    var fake = Fake.init(testing.allocator);
    defer fake.deinit();
    const sh = try fake.addSheet("Data");
    const marked = cellOf("A1");
    const plain = cellOf("B1");
    const rich = cellOf("C1");

    // Every cell claims `.dynamic_array` in the stored field, so an
    // answer of `.legacy` can only have come from the resolver.
    try fake.put(sh, .stored, .{ .row = marked.row, .col = marked.col, .v = num(1), .cm = 1 });
    try fake.put(sh, .stored, .{ .row = plain.row, .col = plain.col, .v = num(2) });
    try fake.put(sh, .stored, .{ .row = rich.row, .col = rich.col, .v = num(3), .vm = 1 });

    var stub: Stub = .{};
    fake.dialect_resolver = .{ .ctx = &stub, .resolve = Stub.resolve };

    const e = fake.evalEnv();
    try testing.expectEqual(
        value.Dialect.dynamic_array,
        try e.dialectOf(.{ .sheet = sh, .row = marked.row, .col = marked.col }),
    );
    try testing.expectEqual(
        value.Dialect.legacy,
        try e.dialectOf(.{ .sheet = sh, .row = plain.row, .col = plain.col }),
    );
    try testing.expectError(
        error.MetadataRefused,
        e.dialectOf(.{ .sheet = sh, .row = rich.row, .col = rich.col }),
    );
    // An empty coordinate is asked about too — it carries no metadata,
    // which is a resolvable state, not an absent one.
    try testing.expectEqual(
        value.Dialect.legacy,
        try e.dialectOf(.{ .sheet = sh, .row = cellOf("Z9").row, .col = cellOf("Z9").col }),
    );
    try testing.expectEqual(@as(u32, 4), stub.calls);

    // Detaching restores the stored-field reading, unchanged.
    fake.dialect_resolver = null;
    try testing.expectEqual(
        value.Dialect.dynamic_array,
        try e.dialectOf(.{ .sheet = sh, .row = plain.row, .col = plain.col }),
    );
}

test "fake: an unknown sheet is an error, not a blank" {
    var fake = Fake.init(testing.allocator);
    defer fake.deinit();
    const e = fake.evalEnv();
    const ghost = SheetIndex.fromInt(7);
    try testing.expectError(error.UnknownSheet, e.cellValue(.{
        .sheet = ghost,
        .row = try coords.Row.fromOneBased(1),
        .col = try coords.Col.fromZeroBased(0),
    }));
}

test "iterator state fits its inline bound" {
    // The bound is real: if an implementation's state grows past it the
    // build fails rather than the iterator silently heap-allocating.
    try testing.expect(@sizeOf(Fake.RangeState) <= iterator_state_bytes);
    try testing.expect(@sizeOf(Fake.AlignedState) <= iterator_state_bytes);
    try testing.expect(@alignOf(Fake.RangeState) <= iterator_state_align);
    try testing.expect(@alignOf(Fake.AlignedState) <= iterator_state_align);
}

test "checkAllAllocationFailures: the fake is leak-safe under OOM" {
    const H = struct {
        fn build(allocator: std.mem.Allocator) !void {
            var fake = Fake.init(allocator);
            defer fake.deinit();
            const sh = try fake.addSheet("S");
            try fake.put(sh, .stored, .{
                .row = try coords.Row.fromOneBased(2),
                .col = try coords.Col.fromZeroBased(1),
                .layer = .stored,
                .v = num(2),
            });
            try fake.put(sh, .stored, .{
                .row = try coords.Row.fromOneBased(1),
                .col = try coords.Col.fromZeroBased(0),
                .layer = .stored,
                .v = num(1),
            });
            try fake.put(sh, .computed, .{
                .row = try coords.Row.fromOneBased(1),
                .col = try coords.Col.fromZeroBased(0),
                .layer = .computed,
                .v = num(3),
            });
            var it = try fake.evalEnv().rangeIterator(areaOf(sh, "A1:B2"));
            var n: usize = 0;
            while (try it.next()) |_| n += 1;
            if (n != 2) return error.WrongCount;
        }
    };
    try testing.checkAllAllocationFailures(testing.allocator, H.build, .{});
}
