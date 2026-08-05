//! `rng_v1` — xoshiro256\*\* seeded from `RunInputs.rng_seed`
//! (`goal_formula.md` §5.5, §5.6d).
//!
//! M3b of the tier-D1 ladder.
//!
//! Seeded randomness is a declared divergence
//! ------------------------------------------
//! Excel's RNG is unspecified, so zlsx cannot match it and does not
//! pretend to. `RAND` and its family are **deterministic given
//! `RunInputs`**, in both fidelity modes, and that is an intentional
//! divergence recorded in §5.5. The consequence for testing is stated
//! there too: Excel-fidelity oracles check shape, type, range, and
//! same-context repeatability — never sequences. Sequences are pinned
//! here instead, by known-answer tests.
//!
//! Why not `std.Random.DefaultPrng`
//! --------------------------------
//! The plan names xoshiro256\*\*, and the stdlib's default is a
//! different member of the family. A generator whose outputs are frozen
//! in KATs has to *be* the named algorithm, not a neighbour of it — and
//! a stdlib default is free to change between Zig releases, which is
//! exactly what a versioned `rng_v1` must not do. Twenty lines of
//! explicit state is the cheaper guarantee.
//!
//! The KATs below were cross-checked against an independent
//! implementation rather than recorded from this one; a golden generated
//! by the code it guards proves only that the code is consistent with
//! itself. `splitmix64(0)`'s first output is the published
//! `0xE220A8397B1DCDAF`, which anchors the seeding half.
//!
//! What this module deliberately does not do
//! -----------------------------------------
//! The **callsite-keyed draw schedule** of §5.6d — draws keyed by
//! invocation path, AST callsite ordinal, SCC pass, and element ordinal
//! — is M5a2's, and needs a graph this row does not have. What lands
//! here is the generator and the seam it plugs into: a `DrawSource` the
//! evaluator already knows how to count.

const std = @import("std");
const assert = std.debug.assert;

const eval = @import("eval.zig");
const run_inputs = @import("run_inputs.zig");

pub const version = "rng_v1";

/// SplitMix64 — the reference seeding procedure for the xoshiro family.
///
/// Seeding a 256-bit state from a 64-bit seed by *copying* it would give
/// neighbouring seeds correlated streams; SplitMix64 is what the
/// algorithm's authors specify instead, and it is also what makes seed 0
/// a perfectly ordinary seed rather than a special case.
pub const SplitMix64 = struct {
    state: u64,

    pub fn init(seed: u64) SplitMix64 {
        return .{ .state = seed };
    }

    pub fn next(self: *SplitMix64) u64 {
        self.state +%= 0x9e3779b97f4a7c15;
        var z = self.state;
        z = (z ^ (z >> 30)) *% 0xbf58476d1ce4e5b9;
        z = (z ^ (z >> 27)) *% 0x94d049bb133111eb;
        return z ^ (z >> 31);
    }
};

/// xoshiro256\*\* 1.0 (Blackman & Vigna).
pub const Rng = struct {
    s: [4]u64,

    /// The generator a run uses, derived from `RunInputs` and nothing
    /// else.
    ///
    /// §5.5's contract is that equal inputs give equal output, and the
    /// draw sequence is the part of it most easily broken by reaching
    /// for a clock or an entropy source. Routing every run's generator
    /// through one named function makes "reproducible from `RunInputs`
    /// alone" a property of the seam rather than of each caller's
    /// discipline — a test can then take the argument list itself as the
    /// statement, which is what M4d's KAT does.
    pub fn fromRunInputs(inputs: run_inputs.RunInputs) Rng {
        return Rng.init(inputs.rng_seed);
    }

    /// The same, from the projection a cache key is built out of. Kept
    /// separate rather than folded into the above because the two
    /// structs are not interchangeable: `EffectiveRunInputs` is what
    /// survives into a fingerprint, and a run replayed from one must
    /// draw the same sequence as the run it came from.
    pub fn fromEffective(effective: run_inputs.EffectiveRunInputs) Rng {
        return Rng.init(effective.rng_seed);
    }

    pub fn init(seed: u64) Rng {
        var sm = SplitMix64.init(seed);
        const r: Rng = .{ .s = .{ sm.next(), sm.next(), sm.next(), sm.next() } };
        // An all-zero state is a fixed point of the generator. SplitMix64
        // cannot produce one from any seed, and asserting it says so
        // rather than leaving the reader to wonder.
        assert(r.s[0] | r.s[1] | r.s[2] | r.s[3] != 0);
        return r;
    }

    pub fn next(self: *Rng) u64 {
        const result = std.math.rotl(u64, self.s[1] *% 5, 7) *% 9;
        const t = self.s[1] << 17;
        self.s[2] ^= self.s[0];
        self.s[3] ^= self.s[1];
        self.s[1] ^= self.s[2];
        self.s[0] ^= self.s[3];
        self.s[2] ^= t;
        self.s[3] = std.math.rotl(u64, self.s[3], 45);
        return result;
    }

    /// A double in `[0, 1)`.
    ///
    /// The top 53 bits scaled by `2^-53`, which is the construction the
    /// algorithm's authors specify: it uses the bits that pass the
    /// generator's own statistical tests, and it cannot round to 1.0 the
    /// way dividing by `2^64 - 1` can.
    pub fn nextFloat(self: *Rng) f64 {
        const bits = self.next() >> 11;
        const v = @as(f64, @floatFromInt(bits)) * 0x1.0p-53;
        assert(v >= 0 and v < 1);
        return v;
    }

    /// A uniform integer in `[lo, hi]`, by rejection — the only method
    /// that is exactly uniform for every range. `RANDBETWEEN` is M4d's;
    /// the primitive belongs with the generator.
    pub fn nextIntInclusive(self: *Rng, lo: i64, hi: i64) i64 {
        assert(lo <= hi);
        const span: u64 = @intCast(@as(i128, hi) - @as(i128, lo));
        if (span == std.math.maxInt(u64)) return @bitCast(self.next());
        const n = span + 1;
        // Reject the tail that would bias the low values.
        const limit = std.math.maxInt(u64) - (std.math.maxInt(u64) % n);
        var x = self.next();
        while (x >= limit) x = self.next();
        return lo + @as(i64, @intCast(x % n));
    }

    /// Plug into the evaluator's counted volatile-draw seam. The counter
    /// lives in `DrawSource`, so a fixture asserting "no draw in the dead
    /// branch" is asserting something about the evaluator rather than
    /// about this generator.
    pub fn drawSource(self: *Rng) eval.DrawSource {
        return .{ .ctx = self, .draw_fn = drawFn };
    }

    fn drawFn(ctx: *anyopaque) f64 {
        const self: *Rng = @ptrCast(@alignCast(ctx));
        return self.nextFloat();
    }
};

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

test "KAT: splitmix64 reproduces the published anchor" {
    // `0xE220A8397B1DCDAF` is the published first output for seed 0 —
    // an anchor from outside this repository, which is the only kind
    // worth having for a seeding procedure.
    var sm = SplitMix64.init(0);
    try testing.expectEqual(@as(u64, 0xE220A8397B1DCDAF), sm.next());
    try testing.expectEqual(@as(u64, 0x6E789E6AA1B965F4), sm.next());
    try testing.expectEqual(@as(u64, 0x06C45D188009454F), sm.next());
    try testing.expectEqual(@as(u64, 0xF88BB8A8724C81EC), sm.next());
}

const Kat = struct { seed: u64, outputs: [8]u64 };

/// Cross-checked against an independent implementation of xoshiro256\*\*
/// + SplitMix64, not recorded from this file.
const kats = [_]Kat{
    .{ .seed = 0, .outputs = .{
        0x99EC5F36CB75F2B4,
        0xBF6E1F784956452A,
        0x1A5F849D4933E6E0,
        0x6AA594F1262D2D2C,
        0xBBA5AD4A1F842E59,
        0xFFEF8375D9EBCACA,
        0x6C160DEED2F54C98,
        0x8920AD648FC30A3F,
    } },
    .{ .seed = 1, .outputs = .{
        0xB3F2AF6D0FC710C5,
        0x853B559647364CEA,
        0x92F89756082A4514,
        0x642E1C7BC266A3A7,
        0xB27A48E29A233673,
        0x24C123126FFDA722,
        0x123004EF8DF510E6,
        0x61954DCC47B1E89D,
    } },
    .{ .seed = 0x123456789ABCDEF0, .outputs = .{
        0xE01D6FAFC557F1B9,
        0xBD627EBE4406B404,
        0x2C23132B578B57DB,
        0x2E8B319D4D1F276A,
        0x608D57ACF53888E4,
        0x9F44D4FE68BDC399,
        0x2BF98C082C7CD85A,
        0x42F3AA03D402664C,
    } },
};

test "KAT: rng_v1 reproduces its frozen streams for every pinned seed" {
    for (kats) |kat| {
        var r = Rng.init(kat.seed);
        for (kat.outputs, 0..) |want, i| {
            const got = r.next();
            if (got != want) {
                std.debug.print(
                    "rng_v1 seed 0x{X:0>16} output {d}: expected 0x{X:0>16}, got 0x{X:0>16}\n",
                    .{ kat.seed, i, want, got },
                );
                return error.KatMismatch;
            }
        }
    }
    try testing.expectEqual(@as(usize, 3), kats.len);
}

test "KAT: the float derivation is the top 53 bits, scaled" {
    for (kats) |kat| {
        var r = Rng.init(kat.seed);
        const got = r.nextFloat();
        const want = @as(f64, @floatFromInt(kat.outputs[0] >> 11)) * 0x1.0p-53;
        try testing.expectEqual(@as(u64, @bitCast(want)), @as(u64, @bitCast(got)));
    }
}

test "the same seed is the same stream, a different seed is not" {
    var a = Rng.init(12345);
    var b = Rng.init(12345);
    var c = Rng.init(12346);
    var same: usize = 0;
    for (0..64) |_| {
        const x = a.next();
        try testing.expectEqual(x, b.next());
        if (x == c.next()) same += 1;
    }
    // Two distinct streams agreeing on even one of 64 draws would be a
    // 2^-58 coincidence; agreeing on none is the assertion.
    try testing.expectEqual(@as(usize, 0), same);
}

test "floats stay in [0, 1) over a long run" {
    var r = Rng.init(0xDEADBEEF);
    var lo: f64 = 1;
    var hi: f64 = 0;
    for (0..100_000) |_| {
        const v = r.nextFloat();
        try testing.expect(v >= 0 and v < 1);
        lo = @min(lo, v);
        hi = @max(hi, v);
    }
    // Not a distribution test — just enough to catch a derivation that
    // collapsed onto a corner of the interval.
    try testing.expect(lo < 0.001);
    try testing.expect(hi > 0.999);
}

test "integers are uniform over their range, including the degenerate one" {
    var r = Rng.init(7);
    var seen = [_]usize{0} ** 6;
    for (0..60_000) |_| {
        const v = r.nextIntInclusive(1, 6);
        try testing.expect(v >= 1 and v <= 6);
        seen[@intCast(v - 1)] += 1;
    }
    // A rejection sampler that biased low values would show it here;
    // the bound is loose enough never to flake for a fixed seed.
    for (seen) |n| try testing.expect(n > 9_000 and n < 11_000);

    // A range of one always returns it, and must not spin.
    for (0..8) |_| try testing.expectEqual(@as(i64, 42), r.nextIntInclusive(42, 42));
    // Negative bounds work; the span is computed in i128 so it cannot
    // overflow on the way to a u64.
    for (0..64) |_| {
        const v = r.nextIntInclusive(-3, 3);
        try testing.expect(v >= -3 and v <= 3);
    }
    _ = r.nextIntInclusive(std.math.minInt(i64), std.math.maxInt(i64));
}

test "the draw source is the evaluator's seam, counted by it" {
    var r = Rng.init(99);
    var source = r.drawSource();
    try testing.expectEqual(@as(u64, 0), source.count);
    const first = try source.draw();
    const second = try source.draw();
    try testing.expectEqual(@as(u64, 2), source.count);
    try testing.expect(first != second);

    // The seam draws from the generator in order: the same seed through
    // the source and through the generator directly agree.
    var direct = Rng.init(99);
    try testing.expectEqual(direct.nextFloat(), first);
    try testing.expectEqual(direct.nextFloat(), second);
}

test "the run's generator is a function of RunInputs and of nothing else" {
    // M4d's draw KATs rest on this: the seam that turns a run into a
    // stream reads one field, so two runs agree exactly when their seeds
    // do — and the fields that are *not* the seed prove it by failing to
    // change anything.
    const base: run_inputs.RunInputs = .{ .now_utc_ms = 1, .rng_seed = 0xA5A5_1234, .limits = .{} };
    try base.validate();

    var a = Rng.fromRunInputs(base);
    var b = Rng.fromRunInputs(base);
    for (0..16) |_| try testing.expectEqual(a.next(), b.next());

    // A different clock, offset, fidelity, or dialect is the same
    // stream. That is not an accident of the implementation — a run
    // whose seed is pinned must replay whatever else moved.
    var moved = base;
    moved.now_utc_ms = 999_999;
    moved.utc_offset_min = 120;
    moved.fidelity = .ieee;
    moved.dialect = .legacy;
    var c = Rng.fromRunInputs(moved);
    var d = Rng.fromRunInputs(base);
    for (0..16) |_| try testing.expectEqual(d.next(), c.next());

    // A different seed is a different stream.
    var seeded = base;
    seeded.rng_seed = base.rng_seed +% 1;
    var e = Rng.fromRunInputs(seeded);
    var f = Rng.fromRunInputs(base);
    var same: usize = 0;
    for (0..16) |_| {
        if (e.next() == f.next()) same += 1;
    }
    try testing.expectEqual(@as(usize, 0), same);

    // The fingerprintable projection names the same stream, for both
    // operations — a replay from a cache key must not draw differently
    // from the run that wrote it.
    for ([_]run_inputs.Operation{ .standalone_eval, .recalc }) |op| {
        var from_full = Rng.fromRunInputs(base);
        var from_effective = Rng.fromEffective(base.effective(op));
        for (0..8) |_| try testing.expectEqual(from_full.next(), from_effective.next());
    }
}

test "state: no seed produces the all-zero fixed point" {
    // Sampled rather than exhaustive, but including the seeds most
    // likely to break a hand-rolled seeding step.
    const seeds = [_]u64{ 0, 1, std.math.maxInt(u64), 0x9e3779b97f4a7c15, 1 << 63 };
    for (seeds) |seed| {
        const r = Rng.init(seed);
        try testing.expect(r.s[0] | r.s[1] | r.s[2] | r.s[3] != 0);
    }
}

test "checkAllAllocationFailures: generation is allocation-free, and stays so" {
    // `Rng` holds four words of state and touches no allocator. The
    // runner allocates only to collect, which is what makes the check
    // non-vacuous: a generator that started allocating would have to
    // leak nothing, and a collection that fails half-way must not.
    const H = struct {
        fn run(allocator: std.mem.Allocator) !void {
            var out: std.ArrayListUnmanaged(u64) = .empty;
            defer out.deinit(allocator);
            for (kats) |kat| {
                var r = Rng.init(kat.seed);
                for (0..8) |_| try out.append(allocator, r.next());
                var source = r.drawSource();
                const v = try source.draw();
                if (!(v >= 0 and v < 1)) return error.OutOfUnitInterval;
            }
            if (out.items.len != kats.len * 8) return error.WrongCount;
        }
    };
    try testing.checkAllAllocationFailures(testing.allocator, H.run, .{});
}
