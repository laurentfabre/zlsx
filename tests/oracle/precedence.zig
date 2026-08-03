//! Fidelity-specific precedence and conflict recording
//! (M1b, `goal_formula.md` §8.2 / §8.3).
//!
//! When four oracles answer the same question and two of them disagree,
//! the tempting move is to split the difference. §8.2 forbids it:
//! **conflicts are recorded, never averaged**. An averaged golden
//! matches nothing that exists — not Excel, not LibreOffice, not the
//! specification — so a divergence gets replaced by a value that is
//! wrong everywhere and looks authoritative.
//!
//! This module therefore cannot average. `resolve` returns one of the
//! observations it was given, by pointer-identity of value, plus the
//! full list of disagreements. There is no arithmetic in this file.
//!
//! Precedence depends on which fidelity is being measured, because the
//! question is different in each:
//!
//!   `.excel` — "what does Excel do?" So Excel decides, the hand-derived
//!   spec suite breaks ties where Excel was not run, the corpus
//!   corroborates, and LibreOffice is the last word only when nothing
//!   else spoke.
//!
//!   `.ieee` — "what does IEEE-754 require?" Excel's documented
//!   departures from it are the very thing being measured, so Excel
//!   cannot decide its own exam. The hand-derived bit goldens lead and
//!   Excel is retained as a **witness**: its disagreement is recorded
//!   as evidence, never as an answer.

const std = @import("std");
const manifest = @import("manifest.zig");
const provenance = @import("provenance.zig");
const adapters = @import("adapters.zig");

pub const Adapter = provenance.Adapter;
pub const Fidelity = manifest.Fidelity;

/// What an adapter is allowed to do in a given fidelity mode.
pub const Role = enum {
    /// May be chosen as the value.
    deciding,
    /// Recorded, compared, and reported on disagreement — but never
    /// chosen. This is what "retained only as a recorded divergence
    /// witness" means in §8.2.
    witness,
};

pub const Standing = struct {
    role: Role,
    /// Lower decides first. Meaningless for witnesses.
    rank: u8,
};

/// The precedence tables. Written out per fidelity rather than derived,
/// because the two orders encode different questions and a shared
/// formula would invite "simplifying" them back together.
pub fn standing(fidelity: Fidelity, adapter: Adapter) Standing {
    return switch (fidelity) {
        // Excel > hand-spec > corpus > LO.
        .excel => switch (adapter) {
            .excel_mac => .{ .role = .deciding, .rank = 0 },
            .hand_spec => .{ .role = .deciding, .rank = 1 },
            .corpus => .{ .role = .witness, .rank = 2 },
            .libreoffice => .{ .role = .deciding, .rank = 3 },
        },
        // Hand-derived bit goldens lead; Excel is a witness only.
        .ieee => switch (adapter) {
            .hand_spec => .{ .role = .deciding, .rank = 0 },
            .excel_mac => .{ .role = .witness, .rank = 1 },
            .libreoffice => .{ .role = .witness, .rank = 2 },
            .corpus => .{ .role = .witness, .rank = 3 },
        },
    };
}

pub const Observation = struct {
    adapter: Adapter,
    entry: manifest.CellEntry,
};

pub const Conflict = struct {
    /// The observation that decided (or, when only witnesses spoke,
    /// the highest-ranked witness).
    reference: Observation,
    /// The observation that disagreed with it.
    dissenting: Observation,
};

pub const Resolution = struct {
    /// The chosen value — always one of the inputs, verbatim.
    chosen: ?Observation,
    /// Every disagreement, in adapter-rank order. Never empty when the
    /// oracles disagreed, never merged, never averaged.
    conflicts: []const Conflict,

    pub fn deinit(self: Resolution, allocator: std.mem.Allocator) void {
        allocator.free(self.conflicts);
    }

    pub fn agreed(self: Resolution) bool {
        return self.conflicts.len == 0;
    }
};

/// A tolerance is an exception to §8.3's bit-exact rule, so it has to
/// justify itself in the type. There is no anonymous epsilon: a
/// comparison either is bit-exact or names the case it is relaxed for
/// and why.
pub const Tolerance = struct {
    /// ULPs of slack allowed on a binary64 comparison.
    ulps: u32,
    /// Why this case cannot be bit-exact. Must be non-empty.
    reason: []const u8,
    /// Which cells it applies to; empty means the whole case.
    refs: []const []const u8 = &.{},

    pub fn appliesTo(self: Tolerance, ref: []const u8) bool {
        if (self.refs.len == 0) return true;
        for (self.refs) |r| {
            if (std.mem.eql(u8, r, ref)) return true;
        }
        return false;
    }
};

pub const Error = error{ToleranceMissingReason} || std.mem.Allocator.Error;

/// Resolve a set of observations of one cell.
///
/// `observations` may be in any order. The result's `chosen` is the
/// deciding observation of lowest rank; every other observation that
/// disagrees with it — including witnesses — becomes a conflict record.
pub fn resolve(
    allocator: std.mem.Allocator,
    fidelity: Fidelity,
    observations: []const Observation,
    tolerance: ?Tolerance,
) Error!Resolution {
    if (tolerance) |t| {
        if (t.reason.len == 0) return error.ToleranceMissingReason;
    }
    if (observations.len == 0) return .{ .chosen = null, .conflicts = &.{} };

    // Highest-ranked DECIDING observation wins. If none of the
    // observations may decide (e.g. `.ieee` with no hand-spec golden),
    // nothing is chosen — an unanswered question, recorded as such,
    // rather than a witness quietly promoted to authority.
    var chosen: ?Observation = null;
    var chosen_rank: u16 = std.math.maxInt(u16);
    for (observations) |o| {
        const s = standing(fidelity, o.adapter);
        if (s.role != .deciding) continue;
        if (s.rank < chosen_rank) {
            chosen = o;
            chosen_rank = s.rank;
        }
    }

    // The reference for conflict reporting: the chosen value, or the
    // best-ranked witness when nothing could decide, so a divergence
    // between witnesses is still visible.
    const reference = chosen orelse blk: {
        var best: ?Observation = null;
        var best_rank: u16 = std.math.maxInt(u16);
        for (observations) |o| {
            const s = standing(fidelity, o.adapter);
            if (s.rank < best_rank) {
                best = o;
                best_rank = s.rank;
            }
        }
        break :blk best.?;
    };

    var conflicts: std.ArrayListUnmanaged(Conflict) = .empty;
    errdefer conflicts.deinit(allocator);
    for (observations) |o| {
        if (o.adapter == reference.adapter) continue;
        if (equalValues(reference.entry, o.entry, tolerance)) continue;
        try conflicts.append(allocator, .{ .reference = reference, .dissenting = o });
    }

    // Deterministic order: by the dissenter's rank, so a report reads
    // the same on every run and diffs cleanly.
    std.mem.sort(Conflict, conflicts.items, fidelity, struct {
        fn lessThan(f: Fidelity, a: Conflict, b: Conflict) bool {
            return standing(f, a.dissenting.adapter).rank < standing(f, b.dissenting.adapter).rank;
        }
    }.lessThan);

    return .{ .chosen = chosen, .conflicts = try conflicts.toOwnedSlice(allocator) };
}

/// §8.3's comparison rule: bit-exact parsed binary64, decoded text and
/// bool, normalized-then-exact errors. Tolerance applies to numbers
/// only — there is no "nearly #VALUE!".
pub fn equalValues(a: manifest.CellEntry, b: manifest.CellEntry, tolerance: ?Tolerance) bool {
    // An excluded cell asserts nothing, so it cannot disagree with
    // anything. Two cells where either side is excluded are not a
    // conflict; they are an absence of evidence.
    if (a.excluded != null or b.excluded != null) return true;
    if (!std.mem.eql(u8, a.kind, b.kind)) return false;

    if (std.mem.eql(u8, a.kind, "number")) {
        const ab = a.numberBits() catch return false;
        const bb = b.numberBits() catch return false;
        if (ab == bb) return true;
        const t = tolerance orelse return false;
        if (!t.appliesTo(a.ref)) return false;
        return withinUlps(@bitCast(ab), @bitCast(bb), t.ulps);
    }
    if (std.mem.eql(u8, a.kind, "text")) {
        return std.mem.eql(u8, a.text orelse "", b.text orelse "");
    }
    if (std.mem.eql(u8, a.kind, "boolean")) {
        return (a.boolean orelse false) == (b.boolean orelse false);
    }
    if (std.mem.eql(u8, a.kind, "error")) {
        const ak = manifest.ErrorKind.normalize(a.error_spelling orelse "");
        const bk = manifest.ErrorKind.normalize(b.error_spelling orelse "");
        if (ak != bk) return false;
        // Two `unknown` spellings are only equal if they are the SAME
        // unknown: `#BLOCKED!` is not `#PYTHON!`.
        if (ak == .unknown) {
            return std.mem.eql(u8, a.error_spelling orelse "", b.error_spelling orelse "");
        }
        return true;
    }
    return std.mem.eql(u8, a.kind, "blank");
}

/// Distance in representable doubles.
///
/// Differing signs mean the values straddle zero, where ULP distance is
/// not meaningful, so no tolerance bridges them. That deliberately
/// includes `+0.0` vs `-0.0`: they are `==` in IEEE, and falling back to
/// `==` here would let any tolerance erase the signed-zero distinction
/// that `manifest.zig` exists to preserve. Identical bit patterns never
/// reach this function — `equalValues` answers those first.
fn withinUlps(a: f64, b: f64, ulps: u32) bool {
    if (std.math.isNan(a) or std.math.isNan(b)) return false;
    if (std.math.signbit(a) != std.math.signbit(b)) return false;
    const ai: i64 = @bitCast(a);
    const bi: i64 = @bitCast(b);
    const diff = if (ai > bi) ai - bi else bi - ai;
    return diff <= @as(i64, ulps);
}

/// Render a conflict set for the record. Caller frees.
pub fn report(allocator: std.mem.Allocator, conflicts: []const Conflict) ![]u8 {
    var out: std.Io.Writer.Allocating = .init(allocator);
    errdefer out.deinit();
    const w = &out.writer;
    for (conflicts) |c| {
        try w.print("DIVERGENCE {s}!{s}: {s} vs {s}\n", .{
            c.reference.entry.sheet,
            c.reference.entry.ref,
            @tagName(c.reference.adapter),
            @tagName(c.dissenting.adapter),
        });
        try w.print("  {s}: {s}\n", .{ @tagName(c.reference.adapter), describe(c.reference.entry) });
        try w.print("  {s}: {s}\n", .{ @tagName(c.dissenting.adapter), describe(c.dissenting.entry) });
    }
    return out.toOwnedSlice();
}

fn describe(e: manifest.CellEntry) []const u8 {
    if (e.excluded) |x| return x;
    if (e.bits) |b| return b;
    if (e.text) |t| return t;
    if (e.error_spelling) |s| return s;
    return e.kind;
}

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

fn num(ref: []const u8, bits: []const u8) manifest.CellEntry {
    return .{ .sheet = "S", .ref = ref, .kind = "number", .bits = bits };
}

fn bitsOf(x: f64) []const u8 {
    // Comptime-friendly hex for the handful of literals the tests use.
    return switch (@as(u64, @bitCast(x))) {
        @as(u64, @bitCast(@as(f64, 1.0))) => "0x3FF0000000000000",
        @as(u64, @bitCast(@as(f64, 2.0))) => "0x4000000000000000",
        else => unreachable,
    };
}

test "excel fidelity: Excel decides, then hand-spec, then LO" {
    try testing.expectEqual(Role.deciding, standing(.excel, .excel_mac).role);
    try testing.expect(standing(.excel, .excel_mac).rank < standing(.excel, .hand_spec).rank);
    try testing.expect(standing(.excel, .hand_spec).rank < standing(.excel, .corpus).rank);
    try testing.expect(standing(.excel, .corpus).rank < standing(.excel, .libreoffice).rank);
    // The corpus is a consistency signal and never decides, in either mode.
    try testing.expectEqual(Role.witness, standing(.excel, .corpus).role);
    try testing.expectEqual(Role.witness, standing(.ieee, .corpus).role);
}

test "ieee fidelity: Excel is a witness and cannot decide its own exam" {
    try testing.expectEqual(Role.deciding, standing(.ieee, .hand_spec).role);
    try testing.expectEqual(Role.witness, standing(.ieee, .excel_mac).role);
    try testing.expectEqual(Role.witness, standing(.ieee, .libreoffice).role);
    // The one adapter that decides in `.ieee` outranks everything.
    try testing.expectEqual(@as(u8, 0), standing(.ieee, .hand_spec).rank);
}

test "the standing table covers every adapter in every fidelity" {
    inline for (std.meta.fields(Fidelity)) |ff| {
        const f: Fidelity = @enumFromInt(ff.value);
        var ranks: [4]u8 = undefined;
        var n: usize = 0;
        inline for (std.meta.fields(Adapter)) |af| {
            const a: Adapter = @enumFromInt(af.value);
            ranks[n] = standing(f, a).rank;
            n += 1;
        }
        // Ranks must be a permutation — two adapters sharing a rank
        // would make resolution order depend on input order.
        std.mem.sort(u8, ranks[0..n], {}, std.sort.asc(u8));
        for (ranks[0..n], 0..) |r, i| try testing.expectEqual(@as(u8, @intCast(i)), r);
    }
}

test "a value never gets averaged: the winner is one of the inputs, verbatim" {
    // The rule §8.2 states outright. Excel says 1.0, LO says 2.0; the
    // answer must be 1.0 — never 1.5, and never a new bit pattern.
    const obs = [_]Observation{
        .{ .adapter = .libreoffice, .entry = num("A1", bitsOf(2.0)) },
        .{ .adapter = .excel_mac, .entry = num("A1", bitsOf(1.0)) },
    };
    const r = try resolve(testing.allocator, .excel, &obs, null);
    defer r.deinit(testing.allocator);

    try testing.expectEqual(Adapter.excel_mac, r.chosen.?.adapter);
    try testing.expectEqual(@as(f64, 1.0), try r.chosen.?.entry.value());
    // …and the disagreement is on the record, not smoothed away.
    try testing.expect(!r.agreed());
    try testing.expectEqual(@as(usize, 1), r.conflicts.len);
    try testing.expectEqual(Adapter.libreoffice, r.conflicts[0].dissenting.adapter);
    try testing.expectEqual(@as(f64, 2.0), try r.conflicts[0].dissenting.entry.value());
}

test "ieee fidelity: Excel's disagreement is recorded but never chosen" {
    const obs = [_]Observation{
        .{ .adapter = .excel_mac, .entry = num("A1", bitsOf(2.0)) },
        .{ .adapter = .hand_spec, .entry = num("A1", bitsOf(1.0)) },
    };
    const r = try resolve(testing.allocator, .ieee, &obs, null);
    defer r.deinit(testing.allocator);

    try testing.expectEqual(Adapter.hand_spec, r.chosen.?.adapter);
    try testing.expectEqual(@as(usize, 1), r.conflicts.len);
    try testing.expectEqual(Adapter.excel_mac, r.conflicts[0].dissenting.adapter);

    // The same two observations under `.excel` flip the answer — which
    // is the whole point of making precedence fidelity-specific.
    const r2 = try resolve(testing.allocator, .excel, &obs, null);
    defer r2.deinit(testing.allocator);
    try testing.expectEqual(Adapter.excel_mac, r2.chosen.?.adapter);
}

test "when only witnesses spoke, nothing is chosen" {
    // `.ieee` with no hand-derived golden: Excel and LO may both have
    // values, but promoting either would answer a question neither is
    // allowed to answer.
    const obs = [_]Observation{
        .{ .adapter = .excel_mac, .entry = num("A1", bitsOf(2.0)) },
        .{ .adapter = .libreoffice, .entry = num("A1", bitsOf(1.0)) },
    };
    const r = try resolve(testing.allocator, .ieee, &obs, null);
    defer r.deinit(testing.allocator);

    try testing.expect(r.chosen == null);
    // Their disagreement is still recorded.
    try testing.expectEqual(@as(usize, 1), r.conflicts.len);
}

test "agreement produces no conflicts" {
    const obs = [_]Observation{
        .{ .adapter = .excel_mac, .entry = num("A1", bitsOf(1.0)) },
        .{ .adapter = .libreoffice, .entry = num("A1", bitsOf(1.0)) },
        .{ .adapter = .hand_spec, .entry = num("A1", bitsOf(1.0)) },
    };
    const r = try resolve(testing.allocator, .excel, &obs, null);
    defer r.deinit(testing.allocator);
    try testing.expect(r.agreed());
    try testing.expectEqual(Adapter.excel_mac, r.chosen.?.adapter);
}

test "conflicts come back in a deterministic order" {
    const obs = [_]Observation{
        .{ .adapter = .libreoffice, .entry = num("A1", bitsOf(2.0)) },
        .{ .adapter = .corpus, .entry = num("A1", bitsOf(2.0)) },
        .{ .adapter = .hand_spec, .entry = num("A1", bitsOf(2.0)) },
        .{ .adapter = .excel_mac, .entry = num("A1", bitsOf(1.0)) },
    };
    const r = try resolve(testing.allocator, .excel, &obs, null);
    defer r.deinit(testing.allocator);
    try testing.expectEqual(@as(usize, 3), r.conflicts.len);
    // hand_spec(1) < corpus(2) < libreoffice(3)
    try testing.expectEqual(Adapter.hand_spec, r.conflicts[0].dissenting.adapter);
    try testing.expectEqual(Adapter.corpus, r.conflicts[1].dissenting.adapter);
    try testing.expectEqual(Adapter.libreoffice, r.conflicts[2].dissenting.adapter);
}

test "comparison is bit-exact by default" {
    // Adjacent doubles: numerically indistinguishable to any tolerance
    // anyone would type by accident, and a genuine divergence.
    const one = num("A1", "0x3FF0000000000000");
    const one_ulp_up = num("A1", "0x3FF0000000000001");
    try testing.expect(!equalValues(one, one_ulp_up, null));
    try testing.expect(equalValues(one, one, null));
}

test "signed zero is a real difference under the comparison rule" {
    const pos = num("A1", "0x0000000000000000");
    const neg = num("A1", "0x8000000000000000");
    try testing.expect(!equalValues(pos, neg, null));
    // …and no tolerance quietly erases it: they straddle zero, so ULP
    // distance is not meaningful and only true equality passes.
    const generous: Tolerance = .{ .ulps = 1_000_000, .reason = "test" };
    try testing.expect(!equalValues(pos, neg, generous));
}

test "a tolerance must justify itself" {
    const obs = [_]Observation{.{ .adapter = .excel_mac, .entry = num("A1", bitsOf(1.0)) }};
    try testing.expectError(error.ToleranceMissingReason, resolve(
        testing.allocator,
        .excel,
        &obs,
        .{ .ulps = 2, .reason = "" },
    ));
}

test "a documented tolerance applies only to the cells it names" {
    const a = num("A1", "0x3FF0000000000000");
    const b = num("A1", "0x3FF0000000000002");
    const scoped: Tolerance = .{
        .ulps = 4,
        .reason = "transcendental; Excel and the spec differ in the last two bits",
        .refs = &.{"A1"},
    };
    try testing.expect(equalValues(a, b, scoped));

    // A different cell gets no slack from a tolerance scoped elsewhere.
    const c = num("B2", "0x3FF0000000000000");
    const d = num("B2", "0x3FF0000000000002");
    try testing.expect(!equalValues(c, d, scoped));
}

test "text, boolean and error comparisons" {
    const t1: manifest.CellEntry = .{ .sheet = "S", .ref = "A1", .kind = "text", .text = "abc" };
    const t2: manifest.CellEntry = .{ .sheet = "S", .ref = "A1", .kind = "text", .text = "abd" };
    try testing.expect(equalValues(t1, t1, null));
    try testing.expect(!equalValues(t1, t2, null));

    const b1: manifest.CellEntry = .{ .sheet = "S", .ref = "A1", .kind = "boolean", .boolean = true };
    const b2: manifest.CellEntry = .{ .sheet = "S", .ref = "A1", .kind = "boolean", .boolean = false };
    try testing.expect(!equalValues(b1, b2, null));

    // Different kinds never compare equal, however similar they look.
    try testing.expect(!equalValues(t1, b1, null));
}

test "errors compare after normalization, and unknown spellings stay distinct" {
    const div: manifest.CellEntry = .{
        .sheet = "S",
        .ref = "A1",
        .kind = "error",
        .error_kind = "div0",
        .error_spelling = "#DIV/0!",
    };
    const div2: manifest.CellEntry = .{
        .sheet = "S",
        .ref = "A1",
        .kind = "error",
        .error_kind = "div0",
        .error_spelling = "#DIV/0!",
    };
    try testing.expect(equalValues(div, div2, null));

    const value: manifest.CellEntry = .{
        .sheet = "S",
        .ref = "A1",
        .kind = "error",
        .error_kind = "value",
        .error_spelling = "#VALUE!",
    };
    try testing.expect(!equalValues(div, value, null));

    // Two DIFFERENT unknown spellings both normalize to `.unknown`, but
    // they are not the same error — folding them together would hide a
    // real divergence between two applications' rich errors.
    const blocked: manifest.CellEntry = .{
        .sheet = "S",
        .ref = "A1",
        .kind = "error",
        .error_kind = "unknown",
        .error_spelling = "#BLOCKED!",
    };
    const python: manifest.CellEntry = .{
        .sheet = "S",
        .ref = "A1",
        .kind = "error",
        .error_kind = "unknown",
        .error_spelling = "#PYTHON!",
    };
    try testing.expect(!equalValues(blocked, python, null));
    try testing.expect(equalValues(blocked, blocked, null));
}

test "an excluded cell is an absence of evidence, not a disagreement" {
    const excluded: manifest.CellEntry = .{
        .sheet = "S",
        .ref = "A1",
        .kind = "number",
        .excluded = "volatile_formula",
    };
    const value = num("A1", bitsOf(1.0));
    try testing.expect(equalValues(excluded, value, null));

    const obs = [_]Observation{
        .{ .adapter = .excel_mac, .entry = excluded },
        .{ .adapter = .hand_spec, .entry = value },
    };
    const r = try resolve(testing.allocator, .excel, &obs, null);
    defer r.deinit(testing.allocator);
    try testing.expect(r.agreed());
}

test "the report names both sides and never merges them" {
    const obs = [_]Observation{
        .{ .adapter = .excel_mac, .entry = num("A1", bitsOf(1.0)) },
        .{ .adapter = .libreoffice, .entry = num("A1", bitsOf(2.0)) },
    };
    const r = try resolve(testing.allocator, .excel, &obs, null);
    defer r.deinit(testing.allocator);

    const text = try report(testing.allocator, r.conflicts);
    defer testing.allocator.free(text);
    try testing.expect(std.mem.indexOf(u8, text, "DIVERGENCE S!A1") != null);
    try testing.expect(std.mem.indexOf(u8, text, "excel_mac") != null);
    try testing.expect(std.mem.indexOf(u8, text, "libreoffice") != null);
    try testing.expect(std.mem.indexOf(u8, text, bitsOf(1.0)) != null);
    try testing.expect(std.mem.indexOf(u8, text, bitsOf(2.0)) != null);
}

test "no observations resolves to nothing rather than inventing a value" {
    const r = try resolve(testing.allocator, .excel, &.{}, null);
    defer r.deinit(testing.allocator);
    try testing.expect(r.chosen == null);
    try testing.expect(r.agreed());
}

test "an adapter the capability matrix bars from authority never decides" {
    // `adapters.zig` says the corpus can never be an authority; the
    // precedence table has to agree, in both fidelities.
    inline for (std.meta.fields(Fidelity)) |ff| {
        const f: Fidelity = @enumFromInt(ff.value);
        inline for (std.meta.fields(Adapter)) |af| {
            const a: Adapter = @enumFromInt(af.value);
            if (!adapters.get(a).can_be_authority) {
                try testing.expectEqual(Role.witness, standing(f, a).role);
            }
        }
    }
}
