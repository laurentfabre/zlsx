//! §5.8a — the spill decision table and the ownership protocol (M7a).
//!
//! One brain, several bodies: `decide` is the decision table as an
//! executable function, `place` is the clear-decide-place protocol, and
//! both run over a `Host` vtable so the workbook model, the closure
//! driver and the tests' fake grid cannot drift apart — the same reason
//! M4a's classification table IS the parser rather than documentation
//! of one.
//!
//! What this file deliberately does NOT decide:
//! - the `(indeterminate)` class is produced by §5.6c's shape-change
//!   rule inside `iterate.zig` (a shape that moves between iterations),
//!   never by the placement decision — a placement sees ONE shape;
//! - `> max_matrix_cells` is enforced by construction at `Matrix.init`
//!   (§9): a rectangle past the cap cannot exist as a matrix, so it can
//!   never reach a placement;
//! - persistence. M7a spills in the MODEL only; `resolved.gateOf` keeps
//!   refusing `FormulaSpillPersistUnsupported` for every non-scalar
//!   publication until M7b1.
//!
//! Decision order, pinned (fixtures below hold every pair): the
//! rectangle must FIT the grid before anything else is even
//! well-defined (a coordinate past `XFD1048576` cannot be probed), then
//! §5.8a's own listing order — foreign non-empty, table, merge. A
//! merged range whose covered cells also hold values answers
//! `obstruction`, because the value check runs first.

const std = @import("std");
const assert = std.debug.assert;

const coords = @import("zlsx_refs");
const env = @import("env.zig");
const value = @import("value.zig");

/// Why an anchor did not spill — the `#SPILL!` classes of §5.8a's
/// decision table. The cell's VALUE stays the bare `#SPILL!` (Excel's
/// own display); the class is the model's per-anchor record, which is
/// where fixtures and diagnostics distinguish the rows.
pub const Class = enum {
    /// A covered coordinate holds foreign content.
    obstruction,
    /// The rectangle leaves the grid.
    edge,
    /// The rectangle intersects a declared table.
    table,
    /// The rectangle intersects a merged range.
    merge,
    /// §5.6c: the shape did not survive iteration. Recorded by the
    /// driver from the iteration report, never decided here.
    indeterminate,
};

pub const Outcome = union(enum) {
    /// The anchor spilled; the shape is its current extent, 1×1
    /// included — a scalar-shaped dynamic result owns exactly its own
    /// cell and can never be blocked.
    spilled: value.Shape,
    /// The anchor holds `#SPILL!`, and this is why.
    blocked: Class,
};

/// The rectangle `anchor` + `shape` names, or null where it leaves the
/// grid — §5.8a's (edge). Arithmetic is u64 so a shape near the 4M cap
/// anchored near the corner cannot wrap.
pub fn extent(anchor: env.CellRef, shape: value.Shape) ?env.RangeRef {
    assert(shape.rows > 0 and shape.cols > 0);
    const last_row: u64 = @as(u64, anchor.row.oneBased()) + shape.rows - 1;
    const last_col: u64 = @as(u64, anchor.col.zeroBased()) + shape.cols; // 1-based
    if (last_row > coords.max_row) return null;
    if (last_col > coords.max_col_1based) return null;
    return .{ .sheet = anchor.sheet, .range = .{
        .first = .{ .row = anchor.row, .col = anchor.col },
        .last = .{
            .row = coords.Row.fromOneBased(@intCast(last_row)) catch unreachable,
            .col = coords.Col.fromZeroBased(anchor.col.zeroBased() + shape.cols - 1) catch unreachable,
        },
    } };
}

/// Rectangle-intersection over a host's declared ranges — the table and
/// merge rows ask exactly this, and both real hosts and the fake grid
/// answer through it so "intersects" cannot mean two things.
pub fn overlapsAny(ranges: []const coords.Range, area: env.RangeRef) bool {
    for (ranges) |r| {
        const n = r.normalized();
        if (n.first.row.oneBased() > area.range.last.row.oneBased()) continue;
        if (n.last.row.oneBased() < area.range.first.row.oneBased()) continue;
        if (n.first.col.zeroBased() > area.range.last.col.zeroBased()) continue;
        if (n.last.col.zeroBased() < area.range.first.col.zeroBased()) continue;
        return true;
    }
    return false;
}

/// The surface a spill lands on. One vtable, three implementors: the
/// workbook model, the closure driver's scratch model, and the fake
/// grid the fixtures and the obstruction fuzz drive.
pub const Host = struct {
    ctx: *anyopaque,
    vtable: *const VTable,

    pub const VTable = struct {
        /// True when `cell` holds content FOREIGN to `anchor`: a stored
        /// non-blank value, any formula, a staged or computed write, or
        /// another anchor's spill tail. The anchor's own cell is never
        /// asked about, and its own previous tails are cleared before
        /// the decision runs, so an implementation need not treat them.
        occupied: *const fn (ctx: *anyopaque, anchor: env.CellRef, cell: env.CellRef) env.Error!bool,
        /// True when the area intersects a declared table.
        overlapsTable: *const fn (ctx: *anyopaque, area: env.RangeRef) bool,
        /// True when the area intersects a merged range.
        overlapsMerge: *const fn (ctx: *anyopaque, area: env.RangeRef) bool,
        /// Remove every tail `anchor` previously placed. Infallible for
        /// the same reason `Fake.clear` is (M3a2): the rollback path
        /// cannot be allowed to fail.
        clearOwnTails: *const fn (ctx: *anyopaque, anchor: env.CellRef) void,
        /// Record one tail cell of `anchor` — §5.8a: spilled cells
        /// record their anchor.
        placeTail: *const fn (ctx: *anyopaque, anchor: env.CellRef, cell: env.CellRef, v: value.ScalarValue) env.Error!void,
    };
};

/// §5.8a's decision table over a grid the caller has already cleared of
/// the anchor's own tails. Total: every shape at every coordinate
/// answers a spill or a class, which is what the obstruction fuzz
/// asserts.
pub fn decide(host: Host, anchor: env.CellRef, shape: value.Shape) env.Error!Outcome {
    const area = extent(anchor, shape) orelse return .{ .blocked = .edge };
    // A 1×1 result occupies exactly the anchor's own cell; there is
    // nothing to obstruct, and a scalar inside a table or a merge is an
    // ordinary formula result.
    if (shape.isScalar()) return .{ .spilled = shape };

    var off: u64 = 1; // 0 is the anchor itself
    const n = area.cellCount();
    while (off < n) : (off += 1) {
        if (try host.vtable.occupied(host.ctx, anchor, area.cellAtOffset(off))) {
            return .{ .blocked = .obstruction };
        }
    }
    if (host.vtable.overlapsTable(host.ctx, area)) return .{ .blocked = .table };
    if (host.vtable.overlapsMerge(host.ctx, area)) return .{ .blocked = .merge };
    return .{ .spilled = shape };
}

/// The ownership protocol, host-neutral:
///   1. the anchor's PREVIOUS tails clear first — shrink clears own
///      tails, and a re-place never obstructs on its own leavings;
///   2. the decision runs against the cleared grid;
///   3. tails place only on a spill verdict — a blocked anchor placed
///      nothing, so a rollback has only the anchor's value to retract.
/// The anchor's own value is the CALLER's to write (hosts store it in
/// different layers): the matrix's top-left on `.spilled`, `#SPILL!`
/// on `.blocked`. A caller must retract (clear tails again) if a later
/// step of its own publication fails, which is what makes a mid-place
/// allocation failure leave no orphan.
pub fn place(host: Host, anchor: env.CellRef, m: value.Matrix) env.Error!Outcome {
    host.vtable.clearOwnTails(host.ctx, anchor);
    const outcome = try decide(host, anchor, m.shape());
    if (outcome == .spilled) {
        const area = extent(anchor, m.shape()).?;
        var off: u64 = 1;
        const n = area.cellCount();
        while (off < n) : (off += 1) {
            try host.vtable.placeTail(host.ctx, anchor, area.cellAtOffset(off), m.cells[@intCast(off)]);
        }
    }
    return outcome;
}

/// The per-anchor outcome record both real hosts embed. `shapeOf` is
/// what `EvalEnv.spillShape` answers (a blocked anchor has NO extent —
/// `A1#` against it is `#REF!`, M3a2's spec-pinned row); `classOf` is
/// what fixtures and diagnostics read to tell §5.8a's rows apart.
pub const Registry = struct {
    map: std.AutoHashMapUnmanaged(env.CellRef, Outcome) = .empty,

    pub fn deinit(self: *Registry, gpa: std.mem.Allocator) void {
        self.map.deinit(gpa);
        self.* = undefined;
    }

    pub fn note(self: *Registry, gpa: std.mem.Allocator, anchor: env.CellRef, outcome: Outcome) error{OutOfMemory}!void {
        try self.map.put(gpa, anchor, outcome);
    }

    /// Infallible: the rollback path.
    pub fn forget(self: *Registry, anchor: env.CellRef) void {
        _ = self.map.remove(anchor);
    }

    pub fn shapeOf(self: *const Registry, anchor: env.CellRef) ?value.Shape {
        const o = self.map.get(anchor) orelse return null;
        return switch (o) {
            .spilled => |s| s,
            .blocked => null,
        };
    }

    pub fn classOf(self: *const Registry, anchor: env.CellRef) ?Class {
        const o = self.map.get(anchor) orelse return null;
        return switch (o) {
            .spilled => null,
            .blocked => |c| c,
        };
    }
};

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

/// A dense little grid implementing the same vtable the workbook model
/// does, plus the mutations the fuzz target drives. Tail bookkeeping is
/// the real protocol's: a tail knows its anchor, `clearOwnTails` sweeps
/// by anchor, foreign tails obstruct.
const FakeGrid = struct {
    const rows = 16;
    const cols = 16;

    const Slot = union(enum) {
        empty,
        /// A stored value or formula a spill may not overwrite.
        content,
        /// A stored BLANK cell — a `<c>` with a style and no value.
        /// §5.8a's "empty": spillable.
        styled_blank,
        tail: struct { anchor: env.CellRef, v: value.ScalarValue },
    };

    slots: [rows * cols]Slot = @splat(.empty),
    tables: std.ArrayListUnmanaged(coords.Range) = .empty,
    merges: std.ArrayListUnmanaged(coords.Range) = .empty,
    /// The grid's formula cells. An anchor's own coordinate holds its
    /// formula, so a FOREIGN anchor is content — exactly what makes two
    /// anchors covering each other's cells resolve at the FIRST one's
    /// placement in the real model.
    anchors: []const env.CellRef = &.{},
    gpa: std.mem.Allocator,

    fn deinit(g: *FakeGrid) void {
        g.tables.deinit(g.gpa);
        g.merges.deinit(g.gpa);
    }

    fn at(g: *FakeGrid, cell: env.CellRef) *Slot {
        const r = cell.row.oneBased() - 1;
        const c = cell.col.zeroBased();
        assert(r < rows and c < cols);
        return &g.slots[r * cols + c];
    }

    fn inGrid(cell: env.CellRef) bool {
        return cell.row.oneBased() <= rows and cell.col.zeroBased() < cols;
    }

    fn vtOccupied(ctx: *anyopaque, anchor: env.CellRef, cell: env.CellRef) env.Error!bool {
        const g: *FakeGrid = @ptrCast(@alignCast(ctx));
        for (g.anchors) |a| {
            if (cell.eql(a) and !a.eql(anchor)) return true;
        }
        // The fuzz can push an in-cap shape past the 16×16 model; a
        // coordinate beyond the fake's storage is empty by definition
        // (the REAL grid's edge is `extent`'s, already decided).
        if (!inGrid(cell)) return false;
        return switch (g.at(cell).*) {
            .empty, .styled_blank => false,
            .content => true,
            .tail => |t| !t.anchor.eql(anchor),
        };
    }

    fn vtOverlapsTable(ctx: *anyopaque, area: env.RangeRef) bool {
        const g: *FakeGrid = @ptrCast(@alignCast(ctx));
        return overlapsAny(g.tables.items, area);
    }

    fn vtOverlapsMerge(ctx: *anyopaque, area: env.RangeRef) bool {
        const g: *FakeGrid = @ptrCast(@alignCast(ctx));
        return overlapsAny(g.merges.items, area);
    }

    fn vtClearOwnTails(ctx: *anyopaque, anchor: env.CellRef) void {
        const g: *FakeGrid = @ptrCast(@alignCast(ctx));
        for (&g.slots) |*s| {
            if (s.* == .tail and s.tail.anchor.eql(anchor)) s.* = .empty;
        }
    }

    fn vtPlaceTail(ctx: *anyopaque, anchor: env.CellRef, cell: env.CellRef, v: value.ScalarValue) env.Error!void {
        const g: *FakeGrid = @ptrCast(@alignCast(ctx));
        if (!inGrid(cell)) return;
        // The protocol probed before placing, so the slot is free.
        assert(g.at(cell).* == .empty or g.at(cell).* == .styled_blank);
        g.at(cell).* = .{ .tail = .{ .anchor = anchor, .v = v } };
    }

    const vtable: Host.VTable = .{
        .occupied = vtOccupied,
        .overlapsTable = vtOverlapsTable,
        .overlapsMerge = vtOverlapsMerge,
        .clearOwnTails = vtClearOwnTails,
        .placeTail = vtPlaceTail,
    };

    fn host(g: *FakeGrid) Host {
        return .{ .ctx = g, .vtable = &vtable };
    }

    fn tailCountOf(g: *FakeGrid, anchor: env.CellRef) usize {
        var n: usize = 0;
        for (g.slots) |s| {
            if (s == .tail and s.tail.anchor.eql(anchor)) n += 1;
        }
        return n;
    }
};

const sheet0 = env.SheetIndex.fromInt(0);

fn cellAt(row_1: u32, col_0: u32) env.CellRef {
    return .{
        .sheet = sheet0,
        .row = coords.Row.fromOneBased(row_1) catch unreachable,
        .col = coords.Col.fromZeroBased(col_0) catch unreachable,
    };
}

fn rangeOf(r1: u32, c1: u32, r2: u32, c2: u32) coords.Range {
    return .{
        .first = .{ .row = coords.Row.fromOneBased(r1) catch unreachable, .col = coords.Col.fromZeroBased(c1) catch unreachable },
        .last = .{ .row = coords.Row.fromOneBased(r2) catch unreachable, .col = coords.Col.fromZeroBased(c2) catch unreachable },
    };
}

fn matrixOf(a: std.mem.Allocator, rows: u32, cols: u32) !value.Matrix {
    const m = try value.Matrix.init(a, rows, cols);
    for (m.cells, 0..) |*c, i| c.* = value.ScalarValue.fromNumber(@floatFromInt(i + 1));
    return m;
}

test "§5.8a row 1: fits + owned/empty spills, tails record their anchor" {
    var g = FakeGrid{ .gpa = testing.allocator };
    defer g.deinit();
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();

    const anchor = cellAt(1, 0);
    // A styled blank inside the extent is §5.8a's "empty", not an
    // obstruction: a `<c>` with a style and no value spills over.
    g.at(cellAt(2, 0)).* = .styled_blank;
    const m = try matrixOf(arena.allocator(), 3, 2);
    const outcome = try place(g.host(), anchor, m);
    try testing.expect(outcome == .spilled);
    try testing.expectEqual(@as(u32, 3), outcome.spilled.rows);
    try testing.expectEqual(@as(usize, 5), g.tailCountOf(anchor));
    // Every tail knows whose it is.
    try testing.expect(g.at(cellAt(3, 1)).* == .tail);
    try testing.expect(g.at(cellAt(3, 1)).tail.anchor.eql(anchor));
    try testing.expectEqual(@as(f64, 6), g.at(cellAt(3, 1)).tail.v.number);
}

test "§5.8a row 2: foreign non-empty is (obstruction), and nothing places" {
    var g = FakeGrid{ .gpa = testing.allocator };
    defer g.deinit();
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();

    const anchor = cellAt(1, 0);
    g.at(cellAt(3, 1)).* = .content;
    const m = try matrixOf(arena.allocator(), 3, 2);
    const outcome = try place(g.host(), anchor, m);
    try testing.expectEqual(Class.obstruction, outcome.blocked);
    try testing.expectEqual(@as(usize, 0), g.tailCountOf(anchor));
}

test "§5.8a row 3: the grid's edge is (edge), from both axes" {
    var g = FakeGrid{ .gpa = testing.allocator };
    defer g.deinit();
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();

    // Row axis: anchored on the last grid row, needing two.
    const bottom = cellAt(coords.max_row, 0);
    const m2 = try matrixOf(arena.allocator(), 2, 1);
    try testing.expectEqual(Class.edge, (try place(g.host(), bottom, m2)).blocked);
    // Column axis: anchored on `XFD`, needing two.
    const right = cellAt(1, coords.max_col_1based - 1);
    const m1x2 = try matrixOf(arena.allocator(), 1, 2);
    try testing.expectEqual(Class.edge, (try place(g.host(), right, m1x2)).blocked);
    // The same anchors hold a 1×1 fine.
    const one = try matrixOf(arena.allocator(), 1, 1);
    try testing.expect((try place(g.host(), bottom, one)) == .spilled);
}

test "§5.8a rows 4 and 5: table and merge overlap block, in pinned order" {
    var g = FakeGrid{ .gpa = testing.allocator };
    defer g.deinit();
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();

    const anchor = cellAt(1, 0);
    const m = try matrixOf(arena.allocator(), 3, 1);

    try g.tables.append(testing.allocator, rangeOf(3, 0, 5, 2));
    try testing.expectEqual(Class.table, (try place(g.host(), anchor, m)).blocked);
    g.tables.clearRetainingCapacity();

    try g.merges.append(testing.allocator, rangeOf(2, 0, 2, 3));
    try testing.expectEqual(Class.merge, (try place(g.host(), anchor, m)).blocked);

    // Pinned pair order 1: a table AND a merge — the table answers,
    // §5.8a's listing order.
    try g.tables.append(testing.allocator, rangeOf(3, 0, 5, 2));
    try testing.expectEqual(Class.table, (try place(g.host(), anchor, m)).blocked);

    // Pinned pair order 2: a merge whose covered cell ALSO holds a
    // value — the value check runs first, so it is (obstruction).
    g.tables.clearRetainingCapacity();
    g.at(cellAt(2, 0)).* = .content;
    try testing.expectEqual(Class.obstruction, (try place(g.host(), anchor, m)).blocked);
}

test "ownership: shrink clears own tails, grow re-places, re-place never self-obstructs" {
    var g = FakeGrid{ .gpa = testing.allocator };
    defer g.deinit();
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();

    const anchor = cellAt(1, 0);
    // 4×1, then 2×1, then 5×1: the shrink's freed tails are gone, the
    // grow lands on them again.
    _ = try place(g.host(), anchor, try matrixOf(arena.allocator(), 4, 1));
    try testing.expectEqual(@as(usize, 3), g.tailCountOf(anchor));
    _ = try place(g.host(), anchor, try matrixOf(arena.allocator(), 2, 1));
    try testing.expectEqual(@as(usize, 1), g.tailCountOf(anchor));
    try testing.expect(g.at(cellAt(3, 0)).* == .empty);
    try testing.expect(g.at(cellAt(4, 0)).* == .empty);
    const grown = try place(g.host(), anchor, try matrixOf(arena.allocator(), 5, 1));
    try testing.expect(grown == .spilled);
    try testing.expectEqual(@as(usize, 4), g.tailCountOf(anchor));
}

test "ownership: competing anchors resolve in calc order — the later one blocks" {
    var g = FakeGrid{ .gpa = testing.allocator };
    defer g.deinit();
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();

    // A at B1 spills 2×2 (B1:C2); B at A2 wants 1×3 (A2:C2). The
    // extents overlap at B2:C2 without either covering the other's
    // formula cell — A got there first in calc order, so B meets A's
    // tails and blocks.
    const a = cellAt(1, 1);
    const b = cellAt(2, 0);
    g.anchors = &.{ a, b };
    _ = try place(g.host(), a, try matrixOf(arena.allocator(), 2, 2));
    const outcome = try place(g.host(), b, try matrixOf(arena.allocator(), 1, 3));
    try testing.expectEqual(Class.obstruction, outcome.blocked);
    // …and A's tails are untouched by B's failure.
    try testing.expectEqual(@as(usize, 3), g.tailCountOf(a));

    // A shrinking to 1×1 releases the overlap; B's next pass spills.
    _ = try place(g.host(), a, try matrixOf(arena.allocator(), 1, 1));
    try testing.expect((try place(g.host(), b, try matrixOf(arena.allocator(), 1, 3))) == .spilled);

    // The reverse calc order blocks A instead: the rule is the ORDER,
    // not the coordinates.
    var g2 = FakeGrid{ .gpa = testing.allocator };
    defer g2.deinit();
    g2.anchors = &.{ a, b };
    _ = try place(g2.host(), b, try matrixOf(arena.allocator(), 1, 3));
    try testing.expectEqual(Class.obstruction, (try place(g2.host(), a, try matrixOf(arena.allocator(), 2, 2))).blocked);

    // An anchor whose extent covers the OTHER'S FORMULA CELL blocks on
    // it directly — a formula is content wherever it is.
    var g3 = FakeGrid{ .gpa = testing.allocator };
    defer g3.deinit();
    const top = cellAt(1, 0);
    const below = cellAt(3, 0);
    g3.anchors = &.{ top, below };
    try testing.expectEqual(Class.obstruction, (try place(g3.host(), top, try matrixOf(arena.allocator(), 3, 1))).blocked);
}

test "registry: a blocked anchor has no extent, a spilled one has its shape" {
    var reg = Registry{};
    defer reg.deinit(testing.allocator);

    const a = cellAt(1, 0);
    const b = cellAt(9, 9);
    try reg.note(testing.allocator, a, .{ .spilled = .{ .rows = 3, .cols = 2 } });
    try reg.note(testing.allocator, b, .{ .blocked = .obstruction });

    try testing.expectEqual(@as(u32, 3), reg.shapeOf(a).?.rows);
    try testing.expect(reg.classOf(a) == null);
    // `A1#` against a blocked anchor is `#REF!` — spillShape answers
    // null, exactly like a non-anchor.
    try testing.expect(reg.shapeOf(b) == null);
    try testing.expectEqual(Class.obstruction, reg.classOf(b).?);

    try reg.note(testing.allocator, a, .{ .blocked = .edge });
    try testing.expect(reg.shapeOf(a) == null);
    reg.forget(a);
    try testing.expect(reg.classOf(a) == null);
}

// ─── the obstruction fuzz (§8.1, wired like the existing targets) ──

fn fuzzObstructionTarget(_: void, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();
    var buf: [512]u8 = undefined;
    const bytes = buf[0..smith.slice(&buf)];

    var g = FakeGrid{ .gpa = testing.allocator };
    defer g.deinit();
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();

    // Two competing anchors, fixed; the fuzz drives everything else.
    const anchors = [2]env.CellRef{ cellAt(1, 0), cellAt(4, 2) };
    g.anchors = &anchors;

    var i: usize = 0;
    while (i + 4 <= bytes.len) : (i += 4) {
        const op = bytes[i] % 6;
        const r: u32 = (bytes[i + 1] % FakeGrid.rows) + 1;
        const c: u32 = bytes[i + 2] % FakeGrid.cols;
        const k = bytes[i + 3];
        const anchor = anchors[k % 2];
        switch (op) {
            // Mutate the grid under whatever is spilled: content lands
            // anywhere EXCEPT on a live tail or an anchor (the model's
            // own writers go through the protocol; the fuzz mutates the
            // stored layer the way §5.8a's "tail mutations ride the
            // transaction" allows — by publication, not by fiat).
            0 => {
                const slot = g.at(cellAt(r, c));
                if (slot.* == .empty) slot.* = .content;
            },
            1 => {
                const slot = g.at(cellAt(r, c));
                if (slot.* == .content) slot.* = .empty;
            },
            2 => {
                if (g.merges.items.len < 4) {
                    try g.merges.append(testing.allocator, rangeOf(r, c, @min(r + 1, FakeGrid.rows), @min(c + 1, FakeGrid.cols - 1)));
                }
            },
            3 => {
                if (g.tables.items.len < 4) {
                    try g.tables.append(testing.allocator, rangeOf(r, c, @min(r + 2, FakeGrid.rows), @min(c + 2, FakeGrid.cols - 1)));
                }
            },
            // Re-place one of the two anchors at a fuzz-chosen shape.
            4 => {
                const shape_rows: u32 = (k % 7) + 1;
                const shape_cols: u32 = (bytes[i + 1] % 5) + 1;
                const m = try matrixOf(arena.allocator(), shape_rows, shape_cols);
                // The protocol's whole promise: whatever the grid holds,
                // the answer is a spill or a class — never a panic, and
                // never a partial placement.
                const outcome = try place(g.host(), anchor, m);
                switch (outcome) {
                    .spilled => |s| try testing.expectEqual(
                        @as(usize, @intCast(@as(u64, s.rows) * s.cols - 1)),
                        g.tailCountOf(anchor),
                    ),
                    .blocked => try testing.expectEqual(@as(usize, 0), g.tailCountOf(anchor)),
                }
            },
            // Retract an anchor entirely (the rollback path).
            5 => {
                FakeGrid.vtClearOwnTails(&g, anchor);
                try testing.expectEqual(@as(usize, 0), g.tailCountOf(anchor));
            },
            else => unreachable,
        }
        // Standing invariant: no tail survives without its anchor being
        // one of the two, and no slot is simultaneously two things —
        // the union enforces the latter, this loop the former.
        for (g.slots) |s| {
            if (s == .tail) {
                try testing.expect(s.tail.anchor.eql(anchors[0]) or s.tail.anchor.eql(anchors[1]));
            }
        }
    }
}

test "fuzz: grids mutate under a spilling anchor and always answer a class or a value" {
    try std.testing.fuzz({}, fuzzObstructionTarget, .{
        .corpus = &[_][]const u8{
            // One placement, then an obstruction, then a re-place.
            &.{ 4, 3, 1, 0, 0, 2, 0, 0, 4, 3, 1, 0 },
            // Merge, table, both anchors racing.
            &.{ 2, 2, 0, 0, 3, 5, 3, 1, 4, 4, 2, 0, 4, 4, 2, 1, 5, 0, 0, 0 },
        },
    });
}
