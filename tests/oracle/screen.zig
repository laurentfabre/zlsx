//! Corpus screening (M1b, `goal_formula.md` §8.2).
//!
//! The corpus leg is a **consistency signal**, not an authority: real
//! workbooks whose cached values were computed by real applications, in
//! bulk. That only means anything if the workbooks whose caches cannot
//! be trusted are removed first — and §8.2 asks for the count as well
//! as the removal, because "412 workbooks agreed" and "412 of 500
//! agreed, and here is why the other 88 were dropped" are very
//! different claims.
//!
//! Six disqualifiers, each for a specific reason the cached values in
//! the file might not be what its own formulas produce:
//!
//!  * **manual calc mode** — the author could have edited inputs and
//!    saved without ever recalculating. The caches may predate the
//!    formulas.
//!  * **fullCalcOnLoad** — the file is asking to be recalculated on
//!    open, which is the author telling us the caches are stale.
//!  * **fullPrecision="0"** — precision-as-displayed rounds every
//!    intermediate to what the cell shows. `goal_formula.md` refuses
//!    this through all of v1, so its values answer a different question.
//!  * **external links** — values may come from a workbook we do not
//!    have, and cannot recompute.
//!  * **volatiles** — a cached `RAND()` is a draw, not a result.
//!  * **unknown provenance** — no record of which application wrote it,
//!    at what version, in what locale.

const std = @import("std");
const extractor = @import("extractor.zig");

pub const Reason = enum {
    manual_calc_mode,
    full_calc_on_load,
    precision_as_displayed,
    external_links,
    volatile_formulas,
    unknown_provenance,
};

pub const reason_count = std.meta.fields(Reason).len;

pub const Verdict = struct {
    admitted: bool,
    /// Every reason that applied, not just the first. A workbook can be
    /// disqualified four ways, and knowing that is more useful than
    /// knowing it failed.
    reasons: std.EnumSet(Reason),

    pub fn has(self: Verdict, r: Reason) bool {
        return self.reasons.contains(r);
    }
};

pub const Options = struct {
    /// False when nothing is known about which application wrote the
    /// file. Corpus workbooks fetched from third-party projects mostly
    /// fall here, which is why the count matters.
    provenance_known: bool = false,
};

pub fn screen(wb: extractor.Workbook, opts: Options) Verdict {
    var reasons: std.EnumSet(Reason) = .initEmpty();

    if (std.mem.eql(u8, wb.calc.calc_mode, "manual")) reasons.insert(.manual_calc_mode);
    if (wb.calc.full_calc_on_load) reasons.insert(.full_calc_on_load);
    if (!wb.calc.full_precision) reasons.insert(.precision_as_displayed);
    if (wb.calc.has_external_references) reasons.insert(.external_links);
    if (!opts.provenance_known) reasons.insert(.unknown_provenance);

    for (wb.cells) |c| {
        const f = c.formula orelse continue;
        if (f.always_calc or mentionsVolatile(f.text)) {
            reasons.insert(.volatile_formulas);
            break;
        }
    }

    return .{ .admitted = reasons.count() == 0, .reasons = reasons };
}

/// Volatile function names. `ca="1"` is the primary signal; this backs
/// it up because Excel does not set `ca` on every volatile, and one
/// unflagged `NOW()` is enough to make a whole workbook's caches
/// unreproducible.
fn mentionsVolatile(formula: []const u8) bool {
    const names = [_][]const u8{
        "RAND(",   "RANDBETWEEN(", "RANDARRAY(", "NOW(",  "TODAY(",
        "OFFSET(", "INDIRECT(",    "CELL(",      "INFO(", "AREAS(",
    };
    for (names) |n| {
        if (containsIgnoreCase(formula, n)) return true;
    }
    return false;
}

fn containsIgnoreCase(haystack: []const u8, needle: []const u8) bool {
    if (needle.len > haystack.len) return false;
    var i: usize = 0;
    while (i + needle.len <= haystack.len) : (i += 1) {
        if (std.ascii.eqlIgnoreCase(haystack[i .. i + needle.len], needle)) return true;
    }
    return false;
}

/// Running tally over a corpus run. §8.2 asks for "screen out + count";
/// this is the count.
pub const Tally = struct {
    total: usize = 0,
    admitted: usize = 0,
    /// How many workbooks each reason disqualified. A workbook counts
    /// under every reason that applied, so these sum to more than
    /// `total - admitted` — deliberately, since the question is "how
    /// common is each disqualifier", not "how many files were dropped".
    per_reason: [reason_count]usize = @splat(0),

    pub fn record(self: *Tally, v: Verdict) void {
        self.total += 1;
        if (v.admitted) {
            self.admitted += 1;
            return;
        }
        inline for (std.meta.fields(Reason)) |f| {
            const r: Reason = @enumFromInt(f.value);
            if (v.has(r)) self.per_reason[f.value] += 1;
        }
    }

    pub fn rejected(self: Tally) usize {
        return self.total - self.admitted;
    }

    pub fn count(self: Tally, r: Reason) usize {
        return self.per_reason[@intFromEnum(r)];
    }

    /// Render the tally. Caller frees. This text goes next to any claim
    /// the corpus leg makes, so the claim carries its own denominator.
    pub fn report(self: Tally, allocator: std.mem.Allocator) ![]u8 {
        var out: std.Io.Writer.Allocating = .init(allocator);
        errdefer out.deinit();
        const w = &out.writer;
        try w.print("corpus screen: {d}/{d} admitted, {d} screened out\n", .{
            self.admitted, self.total, self.rejected(),
        });
        inline for (std.meta.fields(Reason)) |f| {
            const n = self.per_reason[f.value];
            if (n > 0) try w.print("  {s}: {d}\n", .{ f.name, n });
        }
        return out.toOwnedSlice();
    }
};

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

fn wbWith(calc: extractor.CalcState, formula: ?[]const u8, always_calc: bool) extractor.Workbook {
    const cells: []extractor.Cell = if (formula) |f| blk: {
        const storage = struct {
            var cell: [1]extractor.Cell = undefined;
        };
        storage.cell[0] = .{
            .sheet = "S",
            .ref = "A1",
            .kind = .number,
            .value = "1",
            .text = null,
            .formula = .{ .text = f, .kind = null, .ref = null, .si = null, .always_calc = always_calc },
        };
        break :blk &storage.cell;
    } else &.{};

    return .{
        .arena = .init(testing.allocator),
        .sheets = &.{},
        .cells = cells,
        .calc = calc,
        .digest = "0".* ** 64,
    };
}

test "a clean workbook with known provenance is admitted" {
    var wb = wbWith(.{}, "1+1", false);
    defer wb.deinit();
    const v = screen(wb, .{ .provenance_known = true });
    try testing.expect(v.admitted);
    try testing.expectEqual(@as(usize, 0), v.reasons.count());
}

test "each disqualifier is detected on its own" {
    const Case = struct {
        calc: extractor.CalcState,
        formula: ?[]const u8,
        always_calc: bool,
        expect: Reason,
    };
    const cases = [_]Case{
        .{ .calc = .{ .calc_mode = "manual" }, .formula = "1+1", .always_calc = false, .expect = .manual_calc_mode },
        .{ .calc = .{ .full_calc_on_load = true }, .formula = "1+1", .always_calc = false, .expect = .full_calc_on_load },
        .{ .calc = .{ .full_precision = false }, .formula = "1+1", .always_calc = false, .expect = .precision_as_displayed },
        .{ .calc = .{ .has_external_references = true }, .formula = "1+1", .always_calc = false, .expect = .external_links },
        .{ .calc = .{}, .formula = "RAND()", .always_calc = false, .expect = .volatile_formulas },
        .{ .calc = .{}, .formula = "MYSTERY()", .always_calc = true, .expect = .volatile_formulas },
    };
    for (cases) |c| {
        var wb = wbWith(c.calc, c.formula, c.always_calc);
        defer wb.deinit();
        const v = screen(wb, .{ .provenance_known = true });
        try testing.expect(!v.admitted);
        try testing.expect(v.has(c.expect));
        try testing.expectEqual(@as(usize, 1), v.reasons.count());
    }
}

test "unknown provenance alone disqualifies" {
    var wb = wbWith(.{}, "1+1", false);
    defer wb.deinit();
    // The default: nothing is known about who wrote the file.
    const v = screen(wb, .{});
    try testing.expect(!v.admitted);
    try testing.expect(v.has(.unknown_provenance));
}

test "every applicable reason is recorded, not just the first" {
    var wb = wbWith(.{
        .calc_mode = "manual",
        .full_calc_on_load = true,
        .full_precision = false,
        .has_external_references = true,
    }, "NOW()", false);
    defer wb.deinit();

    const v = screen(wb, .{});
    try testing.expect(!v.admitted);
    // All six at once.
    try testing.expectEqual(@as(usize, reason_count), v.reasons.count());
    inline for (std.meta.fields(Reason)) |f| {
        try testing.expect(v.has(@enumFromInt(f.value)));
    }
}

test "volatile detection is case-insensitive and matches the call shape" {
    for ([_][]const u8{ "RAND()", "rand()", "1+Now()", "SUM(OFFSET(A1,1,1))", "INDIRECT(\"A1\")" }) |f| {
        var wb = wbWith(.{}, f, false);
        defer wb.deinit();
        try testing.expect(screen(wb, .{ .provenance_known = true }).has(.volatile_formulas));
    }
    // A name that merely CONTAINS a volatile's letters is not a call:
    // `RANDOM_SEED` is an ordinary defined name.
    for ([_][]const u8{ "SUM(A1:A2)", "RANDOM_SEED", "TODAYS_TOTAL", "A_RAND_B" }) |f| {
        var wb = wbWith(.{}, f, false);
        defer wb.deinit();
        try testing.expect(!screen(wb, .{ .provenance_known = true }).has(.volatile_formulas));
    }
}

test "the tally counts admissions, rejections and every reason" {
    var tally: Tally = .{};

    var clean = wbWith(.{}, "1+1", false);
    defer clean.deinit();
    tally.record(screen(clean, .{ .provenance_known = true }));

    var manual = wbWith(.{ .calc_mode = "manual" }, "1+1", false);
    defer manual.deinit();
    tally.record(screen(manual, .{ .provenance_known = true }));

    var both = wbWith(.{ .calc_mode = "manual" }, "RAND()", false);
    defer both.deinit();
    tally.record(screen(both, .{ .provenance_known = true }));

    try testing.expectEqual(@as(usize, 3), tally.total);
    try testing.expectEqual(@as(usize, 1), tally.admitted);
    try testing.expectEqual(@as(usize, 2), tally.rejected());
    try testing.expectEqual(@as(usize, 2), tally.count(.manual_calc_mode));
    try testing.expectEqual(@as(usize, 1), tally.count(.volatile_formulas));
    try testing.expectEqual(@as(usize, 0), tally.count(.external_links));

    const text = try tally.report(testing.allocator);
    defer testing.allocator.free(text);
    try testing.expect(std.mem.indexOf(u8, text, "1/3 admitted") != null);
    try testing.expect(std.mem.indexOf(u8, text, "manual_calc_mode: 2") != null);
    // A reason nothing triggered is omitted rather than printed as 0.
    try testing.expect(std.mem.indexOf(u8, text, "external_links") == null);
}

test "screens the real corpus and reports a denominator" {
    // The corpus leg's actual behaviour, on the actual corpus. The
    // assertion is not "most workbooks pass" — it is that the screen
    // runs over every one and produces a count, which is what makes the
    // consistency signal quotable.
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var dir = std.Io.Dir.cwd().openDir(io, "tests/corpus", .{ .iterate = true }) catch
        return error.SkipZigTest;
    defer dir.close(io);

    var tally: Tally = .{};
    var it = dir.iterate();
    while (try it.next(io)) |dirent| {
        if (dirent.kind != .file) continue;
        if (!std.mem.endsWith(u8, dirent.name, ".xlsx")) continue;
        const bytes = dir.readFileAlloc(io, dirent.name, testing.allocator, .limited(32 << 20)) catch
            continue;
        defer testing.allocator.free(bytes);
        var wb = extractor.extract(testing.allocator, bytes) catch continue;
        defer wb.deinit();
        tally.record(screen(wb, .{}));
    }

    if (tally.total == 0) return error.SkipZigTest;
    // Third-party corpus files carry no provenance record, so every one
    // is screened out on that basis alone. That is the correct answer
    // and the reason the corpus can never be an authority.
    try testing.expectEqual(@as(usize, 0), tally.admitted);
    try testing.expectEqual(tally.total, tally.count(.unknown_provenance));

    const text = try tally.report(testing.allocator);
    defer testing.allocator.free(text);
    try testing.expect(std.mem.indexOf(u8, text, "unknown_provenance") != null);
}
