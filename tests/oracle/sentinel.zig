//! Stale sentinels — the check that makes an oracle run trustworthy
//! (M1b, `goal_formula.md` §8.2).
//!
//! The failure this exists to prevent is quiet and total. Ask Excel to
//! recalculate, save, and read the result back: if the recalculation
//! silently did not happen — manual calc mode, a dialog swallowed the
//! event, the app decided nothing was dirty — the file still saves, the
//! extractor still reads it, and every cached value it contains is
//! whatever was there before. The oracle then records the OLD values as
//! ground truth and cheerfully confirms them forever. Nothing errors.
//!
//! So a run must prove it recalculated before its output is allowed to
//! count. Each oracle workbook carries cells whose cached values are
//! deliberately wrong; if they come back still wrong, the run is
//! rejected. §8.2 names three:
//!
//!  * **stale value** (Excel) — a formula whose cached `<v>` is a value
//!    it cannot possibly produce. Any calculation at all fixes it.
//!  * **stale dependency** (Excel) — a cell that is only wrong once the
//!    dependency graph is rebuilt. `CalculateFull` walks the recorded
//!    chain; `CalculateFullRebuild` rebuilds the edges first. This
//!    sentinel is what tells the two apart, which matters because
//!    rewritten and dynamic references are exactly where a stale edge
//!    survives a "full" calculation.
//!  * **volatile draw** (LibreOffice) — a volatile cell pinned to a
//!    fixed cached value. LO will happily save a document it merely
//!    loaded, echoing the caches straight back; a redrawn volatile is
//!    proof it actually calculated.
//!
//! Sentinels are checked against the raw extraction, never against a
//! manifest. A manifest excludes volatile cells by design (§8.2), so
//! checking the volatile sentinel there would look at a cell that had
//! been removed from consideration — the check would pass by being
//! absent, which is the exact failure mode it guards.

const std = @import("std");
const extractor = @import("extractor.zig");

pub const Kind = enum {
    stale_value,
    stale_dependency,
    volatile_draw,
};

pub const Sentinel = struct {
    kind: Kind,
    sheet: []const u8,
    ref: []const u8,
    /// The exact `<v>` text planted before the run. Coming back equal
    /// to this is the rejection condition.
    planted: []const u8,
    /// Human-readable statement of why the planted value is impossible.
    /// Carried so a rejection message explains itself without anyone
    /// having to reconstruct the intent from a bare cell reference.
    rationale: []const u8 = "",
};

pub const Reason = enum {
    /// Came back byte-for-byte (or bit-for-bit) as planted: no
    /// recalculation reached this cell.
    unchanged,
    /// The sentinel cell is not in the extraction at all. Either the
    /// workbook was replaced, or the app dropped the cell — both make
    /// the run unusable.
    missing,
    /// Present but carries no cached value, so it proves nothing.
    no_cached_value,
    /// Present but lost its formula, so it is no longer a sentinel.
    formula_lost,
};

pub const Failure = struct {
    sentinel: Sentinel,
    reason: Reason,
    /// What was actually found, for the diagnostic.
    observed: ?[]const u8,
};

pub const Verdict = union(enum) {
    /// Every sentinel moved; the run's values may be recorded.
    accepted,
    /// At least one sentinel did not. Nothing from this run is usable.
    rejected: []const Failure,

    pub fn isAccepted(self: Verdict) bool {
        return self == .accepted;
    }
};

/// Check every sentinel against a post-run extraction. The caller owns
/// the returned failures and frees them with `freeVerdict`.
///
/// The rule is all-or-nothing on purpose. A run where two of three
/// sentinels moved is not "mostly recalculated" — it is a run whose
/// behaviour we do not understand, and partial evidence from an oracle
/// is worse than none.
pub fn check(
    allocator: std.mem.Allocator,
    sentinels: []const Sentinel,
    wb: extractor.Workbook,
) std.mem.Allocator.Error!Verdict {
    var failures: std.ArrayListUnmanaged(Failure) = .empty;
    errdefer failures.deinit(allocator);

    for (sentinels) |s| {
        const cell = wb.cell(s.sheet, s.ref) orelse {
            try failures.append(allocator, .{ .sentinel = s, .reason = .missing, .observed = null });
            continue;
        };
        if (cell.formula == null) {
            try failures.append(allocator, .{
                .sentinel = s,
                .reason = .formula_lost,
                .observed = cell.value,
            });
            continue;
        }
        const observed = cell.value orelse {
            try failures.append(allocator, .{
                .sentinel = s,
                .reason = .no_cached_value,
                .observed = null,
            });
            continue;
        };
        if (sameValue(s.planted, observed)) {
            try failures.append(allocator, .{
                .sentinel = s,
                .reason = .unchanged,
                .observed = observed,
            });
        }
    }

    if (failures.items.len == 0) {
        failures.deinit(allocator);
        return .accepted;
    }
    return .{ .rejected = try failures.toOwnedSlice(allocator) };
}

pub fn freeVerdict(allocator: std.mem.Allocator, verdict: Verdict) void {
    switch (verdict) {
        .accepted => {},
        .rejected => |f| allocator.free(f),
    }
}

/// True when two cached values are the same value.
///
/// Numeric comparison goes through binary64 bits, not text: Excel and
/// LibreOffice round-trip the same double with different decimal
/// spellings (`0.5` vs `5.0000000000000000E-1`), and a text comparison
/// would call an untouched cell "changed" and wave through a run that
/// never recalculated. Non-numeric values compare byte-exact.
fn sameValue(a: []const u8, b: []const u8) bool {
    if (std.mem.eql(u8, a, b)) return true;
    const fa = std.fmt.parseFloat(f64, std.mem.trim(u8, a, " \t\r\n")) catch return false;
    const fb = std.fmt.parseFloat(f64, std.mem.trim(u8, b, " \t\r\n")) catch return false;
    return @as(u64, @bitCast(fa)) == @as(u64, @bitCast(fb));
}

/// Render a rejection for a human. Caller frees.
pub fn explain(allocator: std.mem.Allocator, failures: []const Failure) ![]u8 {
    var out: std.Io.Writer.Allocating = .init(allocator);
    errdefer out.deinit();
    const w = &out.writer;

    try w.print("oracle run REJECTED: {d} sentinel(s) did not prove a recalculation\n", .{failures.len});
    for (failures) |f| {
        try w.print("  [{s}] {s}!{s}: {s}\n", .{
            @tagName(f.sentinel.kind),
            f.sentinel.sheet,
            f.sentinel.ref,
            @tagName(f.reason),
        });
        try w.print("      planted:  {s}\n", .{f.sentinel.planted});
        try w.print("      observed: {s}\n", .{f.observed orelse "(none)"});
        if (f.sentinel.rationale.len > 0) {
            try w.print("      why it is impossible: {s}\n", .{f.sentinel.rationale});
        }
    }
    try w.writeAll(
        "  Nothing from this run may be recorded. Check that the app was driven with a\n" ++
            "  FULL recalculation (Excel: CalculateFullRebuild; LibreOffice: calculateAll)\n" ++
            "  and that no dialog intercepted it.\n",
    );
    return out.toOwnedSlice();
}

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

const test_sentinels = [_]Sentinel{
    .{
        .kind = .stale_value,
        .sheet = "Sentinels",
        .ref = "B1",
        .planted = "999",
        .rationale = "=1+1 cannot be 999",
    },
    .{
        .kind = .stale_dependency,
        .sheet = "Sentinels",
        .ref = "B2",
        .planted = "111",
        .rationale = "chained sum over a rewritten precedent cannot be 111",
    },
    .{
        .kind = .volatile_draw,
        .sheet = "Sentinels",
        .ref = "B3",
        .planted = "0.123456789",
        .rationale = "RAND() pinned to a fixed cache",
    },
};

/// Build a synthetic extraction with the three sentinel cells set to
/// the given cached values.
fn fakeWorkbook(a: std.mem.Allocator, b1: ?[]const u8, b2: ?[]const u8, b3: ?[]const u8) !extractor.Workbook {
    var cells: std.ArrayListUnmanaged(extractor.Cell) = .empty;
    const specs = [_]struct { ref: []const u8, v: ?[]const u8, f: []const u8 }{
        .{ .ref = "B1", .v = b1, .f = "1+1" },
        .{ .ref = "B2", .v = b2, .f = "SUM(A1:A2)" },
        .{ .ref = "B3", .v = b3, .f = "RAND()" },
    };
    for (specs) |s| {
        try cells.append(a, .{
            .sheet = "Sentinels",
            .ref = s.ref,
            .kind = .number,
            .value = s.v,
            .text = null,
            .formula = .{ .text = s.f, .kind = null, .ref = null, .si = null, .always_calc = false },
        });
    }
    return .{
        .arena = .init(testing.allocator),
        .sheets = try a.dupe([]const u8, &.{"Sentinels"}),
        .cells = try cells.toOwnedSlice(a),
        .calc = .{},
        .digest = "0".* ** 64,
    };
}

test "a run where every sentinel moved is accepted" {
    var arena: std.heap.ArenaAllocator = .init(testing.allocator);
    defer arena.deinit();
    var wb = try fakeWorkbook(arena.allocator(), "2", "42", "0.87654321");
    defer wb.deinit();

    const verdict = try check(testing.allocator, &test_sentinels, wb);
    defer freeVerdict(testing.allocator, verdict);
    try testing.expect(verdict.isAccepted());
}

test "PROOF: an unchanged stale-value sentinel rejects the run" {
    // The gate M1b is measured on. Plant the impossible cached value,
    // hand back a workbook that still carries it, and the run must be
    // refused rather than recorded.
    var arena: std.heap.ArenaAllocator = .init(testing.allocator);
    defer arena.deinit();
    var wb = try fakeWorkbook(arena.allocator(), "999", "42", "0.87654321");
    defer wb.deinit();

    const verdict = try check(testing.allocator, &test_sentinels, wb);
    defer freeVerdict(testing.allocator, verdict);

    try testing.expect(!verdict.isAccepted());
    const failures = verdict.rejected;
    try testing.expectEqual(@as(usize, 1), failures.len);
    try testing.expectEqual(Kind.stale_value, failures[0].sentinel.kind);
    try testing.expectEqual(Reason.unchanged, failures[0].reason);
    try testing.expectEqualStrings("999", failures[0].observed.?);
}

test "PROOF: an unchanged stale-dependency sentinel rejects the run" {
    // Distinct from the above: this is the one that survives a
    // `CalculateFull` without a dependency rebuild, so a run driven with
    // the wrong AppleScript verb fails HERE and nowhere else.
    var arena: std.heap.ArenaAllocator = .init(testing.allocator);
    defer arena.deinit();
    var wb = try fakeWorkbook(arena.allocator(), "2", "111", "0.87654321");
    defer wb.deinit();

    const verdict = try check(testing.allocator, &test_sentinels, wb);
    defer freeVerdict(testing.allocator, verdict);

    try testing.expect(!verdict.isAccepted());
    try testing.expectEqual(@as(usize, 1), verdict.rejected.len);
    try testing.expectEqual(Kind.stale_dependency, verdict.rejected[0].sentinel.kind);
    try testing.expectEqual(Reason.unchanged, verdict.rejected[0].reason);
}

test "PROOF: an unchanged volatile sentinel rejects the LibreOffice run" {
    // LO will re-save a document it merely opened, echoing every cached
    // value straight back. A redrawn volatile is the only proof it
    // calculated.
    var arena: std.heap.ArenaAllocator = .init(testing.allocator);
    defer arena.deinit();
    var wb = try fakeWorkbook(arena.allocator(), "2", "42", "0.123456789");
    defer wb.deinit();

    const verdict = try check(testing.allocator, &test_sentinels, wb);
    defer freeVerdict(testing.allocator, verdict);

    try testing.expect(!verdict.isAccepted());
    try testing.expectEqual(Kind.volatile_draw, verdict.rejected[0].sentinel.kind);
}

test "a run that recalculated nothing fails every sentinel at once" {
    var arena: std.heap.ArenaAllocator = .init(testing.allocator);
    defer arena.deinit();
    var wb = try fakeWorkbook(arena.allocator(), "999", "111", "0.123456789");
    defer wb.deinit();

    const verdict = try check(testing.allocator, &test_sentinels, wb);
    defer freeVerdict(testing.allocator, verdict);
    try testing.expectEqual(@as(usize, 3), verdict.rejected.len);
}

test "a differently-spelled but identical double still counts as unchanged" {
    // The subtle one. `0.123456789` re-written by another application as
    // `1.23456789E-1` is the SAME cached value; a text comparison would
    // call it changed and accept a run that never recalculated.
    var arena: std.heap.ArenaAllocator = .init(testing.allocator);
    defer arena.deinit();
    var wb = try fakeWorkbook(arena.allocator(), "2", "42", "1.23456789E-1");
    defer wb.deinit();

    const verdict = try check(testing.allocator, &test_sentinels, wb);
    defer freeVerdict(testing.allocator, verdict);
    try testing.expect(!verdict.isAccepted());
    try testing.expectEqual(Reason.unchanged, verdict.rejected[0].reason);

    // …and a genuinely different value is not caught by it.
    var wb2 = try fakeWorkbook(arena.allocator(), "2", "42", "1.23456788E-1");
    defer wb2.deinit();
    const ok = try check(testing.allocator, &test_sentinels, wb2);
    defer freeVerdict(testing.allocator, ok);
    try testing.expect(ok.isAccepted());
}

test "a missing sentinel cell rejects the run" {
    var arena: std.heap.ArenaAllocator = .init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    var wb: extractor.Workbook = .{
        .arena = .init(testing.allocator),
        .sheets = try a.dupe([]const u8, &.{"Sentinels"}),
        .cells = &.{},
        .calc = .{},
        .digest = "0".* ** 64,
    };
    defer wb.deinit();

    const verdict = try check(testing.allocator, &test_sentinels, wb);
    defer freeVerdict(testing.allocator, verdict);
    try testing.expectEqual(@as(usize, 3), verdict.rejected.len);
    for (verdict.rejected) |f| try testing.expectEqual(Reason.missing, f.reason);
}

test "a sentinel that lost its cached value or its formula rejects the run" {
    var arena: std.heap.ArenaAllocator = .init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    // No `<v>`: proves nothing either way.
    var no_value = try fakeWorkbook(a, null, "42", "0.9");
    defer no_value.deinit();
    const v1 = try check(testing.allocator, &test_sentinels, no_value);
    defer freeVerdict(testing.allocator, v1);
    try testing.expectEqual(Reason.no_cached_value, v1.rejected[0].reason);

    // Formula replaced by a literal: the cell is no longer a sentinel,
    // so its "changed" value means nothing.
    var cells: std.ArrayListUnmanaged(extractor.Cell) = .empty;
    try cells.append(a, .{
        .sheet = "Sentinels",
        .ref = "B1",
        .kind = .number,
        .value = "2",
        .text = null,
        .formula = null,
    });
    var literal: extractor.Workbook = .{
        .arena = .init(testing.allocator),
        .sheets = try a.dupe([]const u8, &.{"Sentinels"}),
        .cells = try cells.toOwnedSlice(a),
        .calc = .{},
        .digest = "0".* ** 64,
    };
    defer literal.deinit();
    const v2 = try check(testing.allocator, test_sentinels[0..1], literal);
    defer freeVerdict(testing.allocator, v2);
    try testing.expectEqual(Reason.formula_lost, v2.rejected[0].reason);
}

test "the rejection explains itself" {
    var arena: std.heap.ArenaAllocator = .init(testing.allocator);
    defer arena.deinit();
    var wb = try fakeWorkbook(arena.allocator(), "999", "42", "0.9");
    defer wb.deinit();

    const verdict = try check(testing.allocator, &test_sentinels, wb);
    defer freeVerdict(testing.allocator, verdict);

    const text = try explain(testing.allocator, verdict.rejected);
    defer testing.allocator.free(text);

    try testing.expect(std.mem.indexOf(u8, text, "REJECTED") != null);
    try testing.expect(std.mem.indexOf(u8, text, "Sentinels!B1") != null);
    try testing.expect(std.mem.indexOf(u8, text, "999") != null);
    try testing.expect(std.mem.indexOf(u8, text, "cannot be 999") != null);
    try testing.expect(std.mem.indexOf(u8, text, "CalculateFullRebuild") != null);
}

test "an empty sentinel set cannot vacuously accept a run" {
    // Guard against the degenerate configuration: a workbook recorded
    // with no sentinels at all would pass every check while proving
    // nothing. Callers get a hard answer here rather than a quiet
    // `accepted`.
    try testing.expect(!hasProof(&.{}));
    try testing.expect(hasProof(&test_sentinels));
    try testing.expect(!hasProof(test_sentinels[2..3])); // volatile alone: no value proof
    try testing.expect(hasProof(test_sentinels[0..1]));
}

/// True when a sentinel set can actually prove a recalculation
/// happened. An empty set proves nothing, and a set of volatile draws
/// alone proves only that a volatile was redrawn — Excel redraws those
/// on load without doing a full calculation, so at least one
/// value-class sentinel is required.
pub fn hasProof(set: []const Sentinel) bool {
    for (set) |s| {
        switch (s.kind) {
            .stale_value, .stale_dependency => return true,
            .volatile_draw => {},
        }
    }
    return false;
}
