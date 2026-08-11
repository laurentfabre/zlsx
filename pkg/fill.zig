//! §9.1 M10q — arena **fill** accounting.
//!
//! The heap profiler (`tests/bench/bench_recalc.zig`) wraps the *backing*
//! allocator, so for an arena it records the chunk **request** and never
//! the bytes written into it. M10p proved the two diverge and by how
//! much: an eight-byte `Cell` cut moved RSS 1 572 864 B while moving peak
//! live by **+550**, because the saving came out of fill inside chunks
//! the arena had already mapped. Every budget figure in §9.1 from M10k
//! onward is denominated in peak-live bytes, so for an arena-held block
//! those budgets bound what the trace will *show*, not what RSS will
//! *give*.
//!
//! `Arena` is a drop-in for `std.heap.ArenaAllocator` that also counts
//! the bytes handed out.
//!
//! **It costs production nothing.** `sinks` is null in every build that
//! does not install one, and `allocator()` then returns the inner
//! arena's allocator *untouched* — there is no wrapper in the allocation
//! path at all. The branch is paid once per `allocator()` call, of which
//! there are a handful per run, not once per allocation. That is why
//! this is a runtime opt-in rather than a build option: the flag would
//! have bought the same zero and cost a module-graph edge in six places.

const std = @import("std");

/// Which arena. The set is deliberately small — these are the owners
/// §9.1's era trace already names, and an owner nobody prices does not
/// need a counter.
pub const Owner = enum {
    /// `WorkbookEnv.arena` — the model. Live era 12 → 23, and the block
    /// M10p's cut came out of.
    model,
    /// `prepare`'s run arena.
    run,
    /// The staging arena.
    stage,
    /// The per-evaluation scratch arena (M10a), reset `.retain_capacity`
    /// between cells — its *peak* fill is the interesting figure.
    scratch,
    /// `PartStore.arena` — decompressed part bytes and every borrowed
    /// slice the store exposes. §9.1's `id7`, and the **only site live in
    /// all 24 eras** (§9.1a): the one allocation whose removal could
    /// lower the envelope rather than expose the next era. It was the
    /// last big owner still uncounted, which is why M10t's first
    /// checkpoint run read `fill=92` through eleven eras that had already
    /// climbed to 48.9 MiB — the bytes were all here.
    parts,
};

/// Bytes handed out of one arena, as opposed to the bytes its chunks
/// requested from the backing allocator.
pub const Tally = struct {
    /// Currently handed out. Arenas rarely free, so this tracks `handed`
    /// closely — the difference is `resize` shrinking in place.
    live: usize = 0,
    /// High-water `live`. **This is the number to compare against a
    /// chunk request:** capacity is what the trace sees, this is what
    /// the pages are.
    peak: usize = 0,
    /// Cumulative, never decremented — fill and churn are different
    /// questions and M10a's lesson was that a site can be small in one
    /// and enormous in the other.
    handed: u64 = 0,
    /// Allocation count, so a mean size falls out.
    calls: u64 = 0,
    /// The arena's chunk capacity when it was torn down — what the heap
    /// profiler sees at the backing boundary. Arenas do not shrink, so
    /// for a monotonic owner this is also its capacity at peak.
    /// `capacity_end − peak` **is the gap M10p's cut fell into**.
    capacity_end: usize = 0,
};

/// Process-wide opt-in, one slot per `Owner`. Null in production; the
/// bench probe installs tallies before the run and reads them after.
/// Not thread-safe by construction: the recalc pipeline is
/// single-threaded and the probe runs one workload at a time.
pub var sinks: [std.meta.fields(Owner).len]?*Tally = @splat(null);

/// Install a tally for `owner`. Returns the previous one so a caller can
/// restore it; the bench does not, because it owns the whole process.
pub fn install(owner: Owner, t: ?*Tally) ?*Tally {
    const i = @intFromEnum(owner);
    const prev = sinks[i];
    sinks[i] = t;
    return prev;
}

/// Clear every sink. The bench calls this between workloads so one
/// run's fill cannot be read as another's.
pub fn clear() void {
    sinks = @splat(null);
}

/// A `std.heap.ArenaAllocator` that counts what it hands out.
///
/// Mirrors the methods this codebase actually calls on an arena —
/// `allocator`, `deinit`, `reset` — so an owner switches type and no
/// call site changes.
pub const Arena = struct {
    inner: std.heap.ArenaAllocator,
    tally: ?*Tally,
    /// What THIS arena instance is holding, as opposed to what the
    /// owner's tally holds. An `Owner` is a role, not an object, and
    /// `parts` proved the two are different: a recalc opens a second
    /// `PartStore` generation over the same archive, so two `parts`
    /// arenas are alive at once and a third has already died. Without a
    /// per-instance figure the tally could not give back what a dead
    /// arena had taken, and M10t's first run with `.parts` installed
    /// read `fill_peak=10 590 591` against `capacity_end=7 944 604` —
    /// a fill larger than the capacity that held it, which is not a
    /// quantity.
    own_fill: usize = 0,

    pub fn init(child: std.mem.Allocator, owner: Owner) Arena {
        return .{
            .inner = std.heap.ArenaAllocator.init(child),
            .tally = sinks[@intFromEnum(owner)],
        };
    }

    pub fn deinit(self: *Arena) void {
        // Queried once, here: `queryCapacity` walks the chunk list, so
        // sampling it per allocation would make the probe the thing
        // being measured.
        //
        // Accumulated, not assigned: with several generations of one
        // owner the last one to die is not the only one that asked its
        // backing for chunks, and `capacity_end` is read against
        // `fill_peak` — a sum against a high-water mark of a sum.
        if (self.tally) |t| {
            t.capacity_end += self.inner.queryCapacity();
            t.live -= self.own_fill;
        }
        self.own_fill = 0;
        self.inner.deinit();
    }

    pub fn reset(self: *Arena, mode: std.heap.ArenaAllocator.ResetMode) bool {
        // Fill is not reclaimed on the tally's `peak`: `scratch` resets
        // per cell, and a high-water mark that fell back to zero each
        // time would report the last cell's fill rather than the arena's
        // high water, which is the only figure a page-resident argument
        // can use. `live` gives back exactly what this instance held —
        // no more, so a sibling arena on the same owner keeps its own.
        if (self.tally) |t| t.live -= self.own_fill;
        self.own_fill = 0;
        return self.inner.reset(mode);
    }

    pub fn queryCapacity(self: *const Arena) usize {
        return self.inner.queryCapacity();
    }

    pub fn allocator(self: *Arena) std.mem.Allocator {
        // The whole cost of this module when it is off.
        if (self.tally == null) return self.inner.allocator();
        return .{ .ptr = self, .vtable = &counting_vtable };
    }

    const counting_vtable: std.mem.Allocator.VTable = .{
        .alloc = allocFn,
        .resize = resizeFn,
        .remap = remapFn,
        .free = freeFn,
    };

    fn note(self: *Arena, delta_add: usize, delta_sub: usize) void {
        const t = self.tally.?;
        self.own_fill = self.own_fill + delta_add - delta_sub;
        t.live = t.live + delta_add - delta_sub;
        if (t.live > t.peak) t.peak = t.live;
    }

    fn allocFn(ctx: *anyopaque, len: usize, a: std.mem.Alignment, ra: usize) ?[*]u8 {
        const self: *Arena = @ptrCast(@alignCast(ctx));
        const inner = self.inner.allocator();
        const p = inner.vtable.alloc(inner.ptr, len, a, ra) orelse return null;
        const t = self.tally.?;
        t.handed += len;
        t.calls += 1;
        self.note(len, 0);
        return p;
    }

    fn resizeFn(ctx: *anyopaque, buf: []u8, a: std.mem.Alignment, new_len: usize, ra: usize) bool {
        const self: *Arena = @ptrCast(@alignCast(ctx));
        const inner = self.inner.allocator();
        if (!inner.vtable.resize(inner.ptr, buf, a, new_len, ra)) return false;
        if (new_len >= buf.len) {
            self.tally.?.handed += new_len - buf.len;
            self.note(new_len - buf.len, 0);
        } else self.note(0, buf.len - new_len);
        return true;
    }

    fn remapFn(ctx: *anyopaque, buf: []u8, a: std.mem.Alignment, new_len: usize, ra: usize) ?[*]u8 {
        const self: *Arena = @ptrCast(@alignCast(ctx));
        const inner = self.inner.allocator();
        const p = inner.vtable.remap(inner.ptr, buf, a, new_len, ra) orelse return null;
        if (new_len >= buf.len) {
            self.tally.?.handed += new_len - buf.len;
            self.note(new_len - buf.len, 0);
        } else self.note(0, buf.len - new_len);
        return p;
    }

    fn freeFn(ctx: *anyopaque, buf: []u8, a: std.mem.Alignment, ra: usize) void {
        const self: *Arena = @ptrCast(@alignCast(ctx));
        const inner = self.inner.allocator();
        inner.vtable.free(inner.ptr, buf, a, ra);
        self.note(0, buf.len);
    }
};

test "fill.Arena: off by default, and then it is the inner allocator" {
    clear();
    var ar: Arena = .init(std.testing.allocator, .model);
    defer ar.deinit();
    try std.testing.expect(ar.tally == null);
    const a = ar.allocator();
    const p = try a.alloc(u8, 100);
    try std.testing.expectEqual(@as(usize, 100), p.len);
}

test "fill.Arena: an installed tally counts what the arena hands out" {
    clear();
    var t: Tally = .{};
    _ = install(.model, &t);
    defer clear();

    var ar: Arena = .init(std.testing.allocator, .model);
    defer ar.deinit();
    const a = ar.allocator();

    _ = try a.alloc(u8, 1000);
    _ = try a.alloc(u8, 2000);
    try std.testing.expectEqual(@as(u64, 3000), t.handed);
    try std.testing.expectEqual(@as(usize, 3000), t.peak);
    try std.testing.expectEqual(@as(u64, 2), t.calls);

    // The capacity the trace would see is the arena's chunk, which is
    // larger than the 3000 bytes actually handed out. That gap is the
    // whole reason this module exists.
    try std.testing.expect(ar.queryCapacity() >= t.peak);
}

test "fill.Arena: capacity_end records the gap fill cannot see" {
    clear();
    var t: Tally = .{};
    _ = install(.model, &t);
    defer clear();

    {
        var ar: Arena = .init(std.testing.allocator, .model);
        defer ar.deinit();
        _ = try ar.allocator().alloc(u8, 3000);
    }
    // The arena asked its backing for a whole chunk; 3000 bytes of it
    // were written. The difference is pages that are mapped and never
    // touched — invisible to the heap profiler, and exactly what an
    // RSS-priced cut comes out of.
    try std.testing.expectEqual(@as(usize, 3000), t.peak);
    try std.testing.expect(t.capacity_end >= t.peak);
}

test "fill.Arena: one owner, several generations — fill never exceeds capacity" {
    clear();
    var t: Tally = .{};
    _ = install(.parts, &t);
    defer clear();

    // The `parts` shape: a generation is opened, a second is opened over
    // it, the first dies. Before `own_fill` the dead one's bytes stayed
    // on `live` forever, so `peak` accumulated across generations and
    // could exceed the capacity that ever existed at one instant.
    {
        var gen1: Arena = .init(std.testing.allocator, .parts);
        _ = try gen1.allocator().alloc(u8, 4000);
        {
            var gen2: Arena = .init(std.testing.allocator, .parts);
            defer gen2.deinit();
            _ = try gen2.allocator().alloc(u8, 1000);
            // Both alive: the owner really is holding 5000 here.
            try std.testing.expectEqual(@as(usize, 5000), t.live);
            try std.testing.expectEqual(@as(usize, 5000), t.peak);
        }
        try std.testing.expectEqual(@as(usize, 4000), t.live);
        gen1.deinit();
    }
    try std.testing.expectEqual(@as(usize, 0), t.live);
    try std.testing.expectEqual(@as(usize, 5000), t.peak);
    try std.testing.expect(t.capacity_end >= t.peak);
}

test "fill.Arena: reset keeps the high water" {
    clear();
    var t: Tally = .{};
    _ = install(.scratch, &t);
    defer clear();

    var ar: Arena = .init(std.testing.allocator, .scratch);
    defer ar.deinit();
    _ = try ar.allocator().alloc(u8, 5000);
    _ = ar.reset(.retain_capacity);
    _ = try ar.allocator().alloc(u8, 10);

    try std.testing.expectEqual(@as(usize, 5000), t.peak);
    try std.testing.expectEqual(@as(usize, 10), t.live);
}
