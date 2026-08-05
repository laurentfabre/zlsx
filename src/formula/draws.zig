//! §5.6d's volatile draw schedule: which number a volatile call gets,
//! and why asking twice cannot change it.
//!
//! M5a2 of the tier-D1 ladder. A leaf on purpose — it imports nothing
//! from the engine, because both the evaluator (which knows where a draw
//! happens) and the iteration engine (which knows which pass it happens
//! in) have to reach it, and a file either of them owned would make the
//! other import a cycle.
//!
//! The key
//! -------
//! §5.6d keys a draw by **(invocation path, stable AST callsite ordinal,
//! SCC pass, element ordinal)**. Each term answers a question a cheaper
//! key gets wrong:
//!
//! | term | the question | what a missing term breaks |
//! |---|---|---|
//! | `path` | *whose* body is running | `A1=N+N` with `N=RAND()` would draw once and reuse it |
//! | `callsite` | *where* in that body | `RAND()+RAND()` would be one draw |
//! | `pass` | *which iteration* | an iterating cell would freeze its first pass's number forever |
//! | `element` | *which cell of an array* | `RANDARRAY` would fill with one value |
//!
//! There is deliberately no component term
//! ---------------------------------------
//! An earlier shape of this key carried the SCC's identity beside its
//! pass number, which reads as the obvious way to spell "SCC-pass". It
//! is wrong, and §5.6e is where it shows: when a dynamic reference
//! merges or splits a component, the affected SCCs "reset their pass
//! counters and re-seed" — and the same cell then re-runs pass 1 while
//! belonging to a *different* component. With a component term that is a
//! new key and a fresh draw, so the rebuild would change the answer;
//! §5.6e says in as many words that it must not. The path already names
//! the cell, and a cell is in exactly one component at a time, so the
//! term bought nothing and cost the one property it was next to.
//!
//! The invocation path is a rolling 64-bit fold rather than a list. It
//! starts at the owning cell and absorbs one segment per name or table
//! expansion — and the segment is the **AST node index of the reference
//! that expanded**, which is what makes two occurrences of one name in
//! one body two different paths. A list would say the same thing and
//! would have to be allocated, compared and hashed at every draw.
//!
//! Memoization is the contract, not an optimization
//! -----------------------------------------------
//! §5.6e rebuilds the graph when a dynamic reference moves, and §5.6c
//! runs a shape pass outside the cycle before iterating. Both re-evaluate
//! bodies that already drew. If those re-evaluations drew again, a
//! *discovery* pass would change a result — the graph would decide the
//! answer, and the graph is supposed to decide only the order. So the
//! memo is what makes "a discovery pass cannot perturb a result" true;
//! it is checked by KAT rather than by inspection.
//!
//! Counts are internal
//! -------------------
//! §5.6d's oracle policy: an external oracle verifies observable
//! properties (repeated-reference inequality, per-reference
//! re-execution, type and range) and **never** a draw count or a
//! sequence. Everything this file pins is therefore an internal KAT, and
//! saying so here is what keeps a later row from promoting one of these
//! numbers into a claim about Excel.

const std = @import("std");
const assert = std.debug.assert;

pub const Error = error{OutOfMemory};

/// The four terms of §5.6d's key, as one comparable value.
pub const Key = struct {
    /// The folded invocation path. `root` for a standalone formula.
    path: u64 = root,
    /// The AST node index of the call. Stable for a given body: the
    /// parser numbers nodes from the text, so the same body parsed twice
    /// numbers the same call the same way.
    callsite: u32 = 0,
    /// Which pass through the owning SCC. Zero outside one — an acyclic
    /// cell evaluates once, and so does §5.6c's pre-iteration shape pass
    /// — and **one-based inside**, so "pass 0" always means "not
    /// iterating" rather than "the first pass".
    pass: u32 = 0,
    /// Which element of an array-producing volatile (`RANDARRAY`, M7a).
    element: u32 = 0,

    /// The path a formula with no owner starts from. A standalone
    /// `evaluate` has no cell, and §5.6d says such roots use a constant
    /// path rather than an invented coordinate.
    pub const root: u64 = 0x9E3779B97F4A7C15;

    /// The path segment a name or table expansion adds.
    ///
    /// `node` is the index, in the CALLING body, of the reference being
    /// expanded — §5.6d's "reference-occurrence ordinal". `row` is the
    /// materialized row for a table producer and zero for everything
    /// else; a producer's body means a different thing in every row of
    /// its span, so two rows must not share a draw.
    pub fn descend(path: u64, node: u32, row: u32) u64 {
        return mix(mix(path, node), row);
    }

    /// The path a stored cell's body starts from.
    pub fn ofCell(sheet: u32, row: u32, col: u32) u64 {
        return mix(mix(mix(root, sheet), row), col);
    }

    pub fn eql(a: Key, b: Key) bool {
        return a.path == b.path and a.callsite == b.callsite and
            a.pass == b.pass and a.element == b.element;
    }
};

/// SplitMix64's finalizer, used as a mixing step rather than as a
/// generator. Avalanche is what is wanted here: two adjacent AST node
/// indices must not produce two adjacent paths, or a deep expansion
/// chain would start colliding with a shallow one.
fn mix(state: u64, v: u32) u64 {
    var z = state ^ (@as(u64, v) +% 0x9E3779B97F4A7C15);
    z = (z ^ (z >> 30)) *% 0xBF58476D1CE4E5B9;
    z = (z ^ (z >> 27)) *% 0x94D049BB133111EB;
    return z ^ (z >> 31);
}

/// The memo: every draw this run has already made, by key.
///
/// Bounded by the caller's byte budget like everything else in a run —
/// it lives in whatever allocator the run hands it, and a run that
/// cannot afford its own draw history is a run that refuses.
pub const Schedule = struct {
    memo: std.AutoHashMapUnmanaged(Key, f64) = .empty,
    /// Draws actually generated, as opposed to served from the memo.
    /// The instrument every rebuild-reuse KAT is written against: a
    /// discovery pass that perturbed a result would show up here as a
    /// second generation at a key that already had one.
    generated: u64 = 0,
    /// Lookups served from the memo. Together with `generated` this is
    /// the total number of times a volatile was *reached*, which is a
    /// different number and the one laziness fixtures care about.
    reused: u64 = 0,

    pub fn deinit(self: *Schedule, gpa: std.mem.Allocator) void {
        self.memo.deinit(gpa);
        self.* = undefined;
    }

    /// The number at `key`, drawn from `gen` exactly once per key.
    ///
    /// `gen` is a closure over whatever PRNG the run wired (`rng_v1`
    /// seeded from `RunInputs`, or a constant in a fixture). It is
    /// called only on a miss, which is the whole point.
    pub fn valueFor(
        self: *Schedule,
        gpa: std.mem.Allocator,
        key: Key,
        ctx: *anyopaque,
        gen: *const fn (ctx: *anyopaque) f64,
    ) Error!f64 {
        const slot = try self.memo.getOrPut(gpa, key);
        if (slot.found_existing) {
            self.reused += 1;
            return slot.value_ptr.*;
        }
        const v = gen(ctx);
        assert(std.math.isFinite(v));
        slot.value_ptr.* = v;
        self.generated += 1;
        return v;
    }

    /// Whether a key has already been decided. For KATs — nothing in the
    /// engine branches on it, because a schedule that behaved
    /// differently when asked would not be a memo.
    pub fn decided(self: Schedule, key: Key) bool {
        return self.memo.contains(key);
    }

    pub fn count(self: Schedule) u32 {
        return self.memo.count();
    }
};

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

/// A generator that hands out 1, 2, 3, … so a KAT can name which draw a
/// key received rather than only that two keys differ.
const Counting = struct {
    n: f64 = 0,

    fn gen(ctx: *anyopaque) f64 {
        const self: *Counting = @ptrCast(@alignCast(ctx));
        self.n += 1;
        return self.n;
    }
};

fn drawAt(s: *Schedule, c: *Counting, key: Key) !f64 {
    return s.valueFor(testing.allocator, key, c, Counting.gen);
}

test "M5a2 KAT: each of the four terms alone distinguishes a draw" {
    var s: Schedule = .{};
    defer s.deinit(testing.allocator);
    var c: Counting = .{};

    const base: Key = .{ .path = 7, .callsite = 3, .pass = 1, .element = 0 };
    try testing.expectEqual(@as(f64, 1), try drawAt(&s, &c, base));

    // Each variant changes exactly one term, so a key that ignored that
    // term would serve draw 1 again instead of a fresh one.
    var v = base;
    v.path = 8;
    try testing.expectEqual(@as(f64, 2), try drawAt(&s, &c, v));
    v = base;
    v.callsite = 4;
    try testing.expectEqual(@as(f64, 3), try drawAt(&s, &c, v));
    v = base;
    v.pass = 2;
    try testing.expectEqual(@as(f64, 4), try drawAt(&s, &c, v));
    v = base;
    v.element = 1;
    try testing.expectEqual(@as(f64, 5), try drawAt(&s, &c, v));

    // …and the original key still answers 1, which is the memo.
    try testing.expectEqual(@as(f64, 1), try drawAt(&s, &c, base));
    try testing.expectEqual(@as(u64, 5), s.generated);
    try testing.expectEqual(@as(u64, 1), s.reused);

    // The key has exactly the four terms §5.6d names. A fifth would be a
    // fifth thing a rebuild could accidentally change (see the header on
    // why the component was removed), so the count is asserted rather
    // than left to the field list.
    try testing.expectEqual(@as(usize, 4), @typeInfo(Key).@"struct".fields.len);
}

test "M5a2 KAT: two occurrences of one name are two paths" {
    // §5.6d's `A1=N+N` with `N=RAND()`. The two `N` references are
    // distinct AST nodes in the calling body, so the segment they add
    // differs — even though the name, the expanded body, and therefore
    // the callsite inside it are all identical.
    const owner = Key.ofCell(0, 1, 0);
    const first = Key.descend(owner, 11, 0);
    const second = Key.descend(owner, 12, 0);
    try testing.expect(first != second);

    var s: Schedule = .{};
    defer s.deinit(testing.allocator);
    var c: Counting = .{};

    // `callsite` is the same in both: it is the `RAND()` call inside
    // `N`'s body, which is one body parsed the same way twice.
    const a = try drawAt(&s, &c, .{ .path = first, .callsite = 1 });
    const b = try drawAt(&s, &c, .{ .path = second, .callsite = 1 });
    try testing.expect(a != b);
    try testing.expectEqual(@as(u64, 2), s.generated);
}

test "M5a2 KAT: a table producer's rows do not share a draw" {
    const owner = Key.ofCell(0, 1, 0);
    const row3 = Key.descend(owner, 5, 3);
    const row4 = Key.descend(owner, 5, 4);
    try testing.expect(row3 != row4);
    // …and the non-producer segment is neither of them, so "row 0" is a
    // third thing rather than an alias for the first row.
    try testing.expect(Key.descend(owner, 5, 0) != row3);
}

test "M5a2 KAT: re-evaluation at a decided key generates nothing" {
    // The property §5.6e rests on. A rebuild pass walks the same bodies
    // at the same keys; if this were not true the graph would decide the
    // answer instead of the order.
    var s: Schedule = .{};
    defer s.deinit(testing.allocator);
    var c: Counting = .{};

    const key: Key = .{ .path = Key.ofCell(0, 1, 0), .callsite = 2 };
    const first = try drawAt(&s, &c, key);
    try testing.expect(!s.decided(.{ .path = 99 }));
    try testing.expect(s.decided(key));

    var i: usize = 0;
    while (i < 16) : (i += 1) {
        try testing.expectEqual(first, try drawAt(&s, &c, key));
    }
    try testing.expectEqual(@as(u64, 1), s.generated);
    try testing.expectEqual(@as(u64, 16), s.reused);
    try testing.expectEqual(@as(u32, 1), s.count());
}

test "M5a2: pass 0 means not iterating, and the first pass is 1" {
    // The distinction exists so §5.6c's pre-iteration shape pass and the
    // first real pass of the same anchor cannot collide on a key. Both
    // would otherwise be `pass = 0`.
    const outside: Key = .{ .path = 1, .callsite = 0, .pass = 0 };
    const first_pass: Key = .{ .path = 1, .callsite = 0, .pass = 1 };
    try testing.expect(!Key.eql(outside, first_pass));
}

test "M5a2 KAT: a reset pass counter reuses its draws — §5.6e's whole promise" {
    // §5.6e resets a changed SCC's pass counter and re-seeds it. The
    // cell then runs pass 1 again, and the memo has to answer with what
    // pass 1 already decided — otherwise a *discovery* pass would change
    // a value, and the graph would be deciding answers instead of order.
    var s: Schedule = .{};
    defer s.deinit(testing.allocator);
    var c: Counting = .{};

    const cell = Key.ofCell(0, 1, 0);
    const first_run = try drawAt(&s, &c, .{ .path = cell, .callsite = 4, .pass = 1 });
    const second_run = try drawAt(&s, &c, .{ .path = cell, .callsite = 4, .pass = 1 });
    try testing.expectEqual(first_run, second_run);
    try testing.expectEqual(@as(u64, 1), s.generated);
}

test "M5a2: the standalone root path is a constant, not an invented cell" {
    // §5.6d: "standalone roots use a constant path". A1 is the cell a
    // guess would pick, so it is the collision worth naming.
    try testing.expect(Key.root != Key.ofCell(0, 1, 0));
    const k: Key = .{};
    try testing.expectEqual(Key.root, k.path);
}

test "M5a2: descending is order-sensitive and depth-sensitive" {
    // A fold that summed or xored would make `A→B` and `B→A` one path,
    // and a name that expanded to itself indistinguishable from one that
    // did not.
    const p = Key.root;
    try testing.expect(Key.descend(Key.descend(p, 1, 0), 2, 0) !=
        Key.descend(Key.descend(p, 2, 0), 1, 0));
    try testing.expect(Key.descend(p, 1, 0) != Key.descend(Key.descend(p, 1, 0), 1, 0));
}
