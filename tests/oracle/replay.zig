//! Oracle replay — the CI half of the harness (M1b, `goal_formula.md` §8.2).
//!
//! Recording an oracle needs Excel, LibreOffice, and macOS. Replaying
//! one needs nothing: the manifests under `fixtures/` are committed, so
//! every gate below runs on a Linux CI box with no spreadsheet
//! application anywhere near it. That separation is the point —
//! evidence gathered once, on a machine with the applications, has to
//! keep gating the build everywhere else.
//!
//! This file is also the harness's test aggregator. Zig analyses lazily:
//! a module nothing imports has its tests silently skipped, and a whole
//! file of green-looking tests can run zero times. The `test { _ = … }`
//! block below is what stops that.

const std = @import("std");

const adapters = @import("adapters.zig");
const extractor = @import("extractor.zig");
const manifest = @import("manifest.zig");
const precedence = @import("precedence.zig");
const provenance = @import("provenance.zig");
const screen = @import("screen.zig");
const sentinel = @import("sentinel.zig");
const sentinel_set = @import("sentinel_set.zig");

// Pull every harness module into this compilation so its tests run.
// Without this, `zig build test` reports a cheerful pass for files it
// never analysed.
test {
    _ = @import("zip_reader.zig");
    _ = @import("xml_scan.zig");
    _ = extractor;
    _ = manifest;
    _ = provenance;
    _ = sentinel;
    _ = sentinel_set;
    _ = adapters;
    _ = precedence;
    _ = screen;
}

const testing = std.testing;
const fixtures_dir = "tests/oracle/fixtures";

fn readFixture(io: std.Io, allocator: std.mem.Allocator, name: []const u8) ![]u8 {
    var buf: [256]u8 = undefined;
    const path = try std.fmt.bufPrint(&buf, "{s}/{s}", .{ fixtures_dir, name });
    return std.Io.Dir.cwd().readFileAlloc(io, path, allocator, .limited(8 << 20));
}

/// Every committed manifest, parsed. Skips the test when the fixtures
/// directory is absent (a partial checkout), fails when it is present
/// but a manifest inside it is malformed.
fn loadAll(
    io: std.Io,
    allocator: std.mem.Allocator,
    out: *std.ArrayListUnmanaged(std.json.Parsed(manifest.Manifest)),
) !void {
    var dir = std.Io.Dir.cwd().openDir(io, fixtures_dir, .{ .iterate = true }) catch
        return error.SkipZigTest;
    defer dir.close(io);

    var it = dir.iterate();
    while (try it.next(io)) |dirent| {
        if (dirent.kind != .file) continue;
        if (!std.mem.endsWith(u8, dirent.name, ".json")) continue;
        if (std.mem.endsWith(u8, dirent.name, ".provenance.json")) continue;

        const bytes = try dir.readFileAlloc(io, dirent.name, allocator, .limited(8 << 20));
        defer allocator.free(bytes);
        const parsed = manifest.parse(allocator, bytes) catch |e| {
            std.debug.print("fixture {s} failed to parse/validate: {s}\n", .{ dirent.name, @errorName(e) });
            return e;
        };
        try out.append(allocator, parsed);
    }
}

test "every committed manifest parses, validates, and carries full provenance" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var loaded: std.ArrayListUnmanaged(std.json.Parsed(manifest.Manifest)) = .empty;
    defer {
        for (loaded.items) |p| p.deinit();
        loaded.deinit(testing.allocator);
    }
    try loadAll(io, testing.allocator, &loaded);
    if (loaded.items.len == 0) return error.SkipZigTest;

    for (loaded.items) |p| {
        const m = p.value;
        // `manifest.parse` already validated; assert the properties the
        // gate actually depends on, so a future loosening of `validate`
        // cannot quietly weaken this.
        try m.provenance.validate();
        try testing.expect(m.case.len > 0);
        try testing.expect(m.cells.len > 0);
        _ = try manifest.Fidelity.parse(m.fidelity);
        _ = try m.provenance.adapterEnum();
        try testing.expectEqualStrings(extractor.version, m.provenance.extractor_version);
    }
}

test "recorded values obey the manifest's own invariants" {
    // Re-checked here rather than trusted from record time: a manifest
    // is a file anyone can edit, and the properties below are the ones
    // every later comparison assumes.
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var loaded: std.ArrayListUnmanaged(std.json.Parsed(manifest.Manifest)) = .empty;
    defer {
        for (loaded.items) |p| p.deinit();
        loaded.deinit(testing.allocator);
    }
    try loadAll(io, testing.allocator, &loaded);
    if (loaded.items.len == 0) return error.SkipZigTest;

    var numbers: usize = 0;
    for (loaded.items) |p| {
        for (p.value.cells) |c| {
            if (c.excluded != null) continue;
            if (!std.mem.eql(u8, c.kind, "number")) continue;
            const f = try c.value();
            // NaN and infinity are hard errors (§8.2's semantic
            // manifests); `validate` enforces it, this proves it held
            // for every committed number rather than in principle.
            try testing.expect(!std.math.isNan(f));
            try testing.expect(!std.math.isInf(f));
            numbers += 1;
        }
    }
    try testing.expect(numbers > 0);
}

test "volatiles are excluded from every external-application manifest" {
    // §8.2, checked against what was actually recorded. A volatile that
    // slipped into a golden produces a test that fails on the next run
    // for no reproducible reason.
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var loaded: std.ArrayListUnmanaged(std.json.Parsed(manifest.Manifest)) = .empty;
    defer {
        for (loaded.items) |p| p.deinit();
        loaded.deinit(testing.allocator);
    }
    try loadAll(io, testing.allocator, &loaded);
    if (loaded.items.len == 0) return error.SkipZigTest;

    for (loaded.items) |p| {
        const m = p.value;
        const adapter = try m.provenance.adapterEnum();
        if (!adapters.get(adapter).excludes_volatiles) continue;
        for (m.cells) |c| {
            const formula = c.formula orelse continue;
            if (!containsIgnoreCase(formula, "RAND(") and !containsIgnoreCase(formula, "NOW(")) continue;
            if (c.excluded == null) {
                std.debug.print(
                    "{s}: volatile {s}!{s} ({s}) was recorded as a value\n",
                    .{ m.case, c.sheet, c.ref, formula },
                );
                return error.TestUnexpectedResult;
            }
            try testing.expectEqualStrings("volatile_formula", c.excluded.?);
        }
    }
}

test "PROOF: replay rejects a run whose sentinel came back unchanged" {
    // The gate M1b is measured on, exercised end-to-end over the real
    // committed sentinel set rather than a synthetic one. A workbook
    // still carrying its planted values is the exact artefact a failed
    // driver produces.
    var arena: std.heap.ArenaAllocator = .init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    var cells: std.ArrayListUnmanaged(extractor.Cell) = .empty;
    for (sentinel_set.all) |s| {
        try cells.append(a, .{
            .sheet = s.sheet,
            .ref = s.ref,
            .kind = .number,
            .value = s.planted, // never recalculated
            .text = null,
            .formula = .{ .text = "1+1", .kind = null, .ref = null, .si = null, .always_calc = false },
        });
    }
    var stale: extractor.Workbook = .{
        .arena = .init(testing.allocator),
        .sheets = try a.dupe([]const u8, &.{sentinel_set.sheet}),
        .cells = try cells.toOwnedSlice(a),
        .calc = .{},
        .digest = "0".* ** 64,
    };
    defer stale.deinit();

    for ([_]provenance.Adapter{ .excel_mac, .libreoffice }) |adapter| {
        const set = sentinel_set.forAdapter(adapter);
        try testing.expect(sentinel.hasProof(set));
        const verdict = try sentinel.check(testing.allocator, set, stale);
        defer sentinel.freeVerdict(testing.allocator, verdict);
        try testing.expect(!verdict.isAccepted());
        try testing.expectEqual(set.len, verdict.rejected.len);
        for (verdict.rejected) |f| try testing.expectEqual(sentinel.Reason.unchanged, f.reason);
    }
}

test "PROOF: the same workbook with recalculated sentinels is accepted" {
    // The other half of the proof: the check must not reject everything.
    var arena: std.heap.ArenaAllocator = .init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    var cells: std.ArrayListUnmanaged(extractor.Cell) = .empty;
    for (sentinel_set.all) |s| {
        try cells.append(a, .{
            .sheet = s.sheet,
            .ref = s.ref,
            .kind = .number,
            .value = "42.5", // any value that is not the planted one
            .text = null,
            .formula = .{ .text = "1+1", .kind = null, .ref = null, .si = null, .always_calc = false },
        });
    }
    var fresh: extractor.Workbook = .{
        .arena = .init(testing.allocator),
        .sheets = try a.dupe([]const u8, &.{sentinel_set.sheet}),
        .cells = try cells.toOwnedSlice(a),
        .calc = .{},
        .digest = "0".* ** 64,
    };
    defer fresh.deinit();

    const verdict = try sentinel.check(testing.allocator, sentinel_set.forAdapter(.excel_mac), fresh);
    defer sentinel.freeVerdict(testing.allocator, verdict);
    try testing.expect(verdict.isAccepted());
}

test "cross-adapter precedence resolves the recorded cases without averaging" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var loaded: std.ArrayListUnmanaged(std.json.Parsed(manifest.Manifest)) = .empty;
    defer {
        for (loaded.items) |p| p.deinit();
        loaded.deinit(testing.allocator);
    }
    try loadAll(io, testing.allocator, &loaded);
    if (loaded.items.len < 2) return error.SkipZigTest;

    // Group by case name, then resolve each cell across whichever
    // adapters recorded it.
    var resolved: usize = 0;
    var divergences: usize = 0;
    for (loaded.items) |first| {
        for (first.value.cells) |cell| {
            if (cell.excluded != null) continue;

            var obs: std.ArrayListUnmanaged(precedence.Observation) = .empty;
            defer obs.deinit(testing.allocator);
            for (loaded.items) |other| {
                if (!std.mem.eql(u8, other.value.case, first.value.case)) continue;
                if (!std.mem.eql(u8, other.value.fidelity, first.value.fidelity)) continue;
                const match = other.value.find(cell.sheet, cell.ref) orelse continue;
                try obs.append(testing.allocator, .{
                    .adapter = try other.value.provenance.adapterEnum(),
                    .entry = match,
                });
            }
            if (obs.items.len < 2) continue;

            const fidelity = try manifest.Fidelity.parse(first.value.fidelity);
            const r = try precedence.resolve(testing.allocator, fidelity, obs.items, null);
            defer r.deinit(testing.allocator);
            resolved += 1;
            divergences += r.conflicts.len;

            // The chosen value is always one of the inputs, verbatim —
            // never a blend. Proven by identity against the observations.
            if (r.chosen) |c| {
                var found = false;
                for (obs.items) |o| {
                    if (o.adapter == c.adapter and precedence.equalValues(o.entry, c.entry, null)) {
                        found = true;
                    }
                }
                try testing.expect(found);
            }
        }
    }
    try testing.expect(resolved > 0);
}

test "the corpus screen runs over the corpus and yields a denominator" {
    // The corpus leg is a consistency signal; §8.2 asks for the screen
    // AND the count, because a bare "they all agreed" hides how many
    // workbooks were dropped to get there.
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var dir = std.Io.Dir.cwd().openDir(io, "tests/corpus", .{ .iterate = true }) catch
        return error.SkipZigTest;
    defer dir.close(io);

    var tally: screen.Tally = .{};
    var it = dir.iterate();
    while (try it.next(io)) |dirent| {
        if (dirent.kind != .file) continue;
        if (!std.mem.endsWith(u8, dirent.name, ".xlsx")) continue;
        const bytes = dir.readFileAlloc(io, dirent.name, testing.allocator, .limited(32 << 20)) catch
            continue;
        defer testing.allocator.free(bytes);
        var wb = extractor.extract(testing.allocator, bytes) catch continue;
        defer wb.deinit();
        tally.record(screen.screen(wb, .{}));
    }
    if (tally.total == 0) return error.SkipZigTest;
    try testing.expectEqual(tally.total, tally.admitted + tally.rejected());
}

test "the hand-derived goldens pin the committed input workbook" {
    // The hand-spec leg runs nothing, so its provenance digest is the
    // only thing tying its values to a set of formulas. If the input
    // workbook is rebuilt and the goldens are not, that digest is the
    // one place the drift is visible — so check it rather than trust it.
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();

    const input = std.Io.Dir.cwd().readFileAlloc(
        io,
        "tests/oracle/inputs/oracle_suite.xlsx",
        testing.allocator,
        .limited(8 << 20),
    ) catch return error.SkipZigTest;
    defer testing.allocator.free(input);

    var digest: [32]u8 = undefined;
    std.crypto.hash.sha2.Sha256.hash(input, &digest, .{});
    var hex: [64]u8 = undefined;
    _ = try std.fmt.bufPrint(&hex, "{x}", .{&digest});

    var loaded: std.ArrayListUnmanaged(std.json.Parsed(manifest.Manifest)) = .empty;
    defer {
        for (loaded.items) |p| p.deinit();
        loaded.deinit(testing.allocator);
    }
    try loadAll(io, testing.allocator, &loaded);

    var checked: usize = 0;
    for (loaded.items) |p| {
        if (try p.value.provenance.adapterEnum() != .hand_spec) continue;
        if (!std.mem.eql(u8, &hex, p.value.provenance.workbook_digest)) {
            std.debug.print(
                "hand-spec manifest '{s}' was derived against a different input:\n" ++
                    "  manifest: {s}\n  on disk:  {s}\n" ++
                    "  Re-run scripts/oracle/regenerate.sh\n",
                .{ p.value.case, p.value.provenance.workbook_digest, &hex },
            );
            return error.TestExpectedEqual;
        }
        checked += 1;
    }
    try testing.expect(checked > 0);
}

test "recorded divergence: LibreOffice rounds where IEEE does not" {
    // Not a bug report — a pinned observation. §8.2 says conflicts are
    // recorded and never averaged, and this is what "recorded" means in
    // practice: the disagreement is a named test, so the day LO changes
    // its behaviour the harness says so instead of quietly agreeing.
    //
    // `.ieee` fidelity, where the hand-derived bit goldens DECIDE and
    // LibreOffice is only a witness.
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();

    const lo_bytes = readFixture(io, testing.allocator, "libreoffice_oracle_suite.json") catch
        return error.SkipZigTest;
    defer testing.allocator.free(lo_bytes);
    const ieee_bytes = readFixture(io, testing.allocator, "hand_spec_ieee.json") catch
        return error.SkipZigTest;
    defer testing.allocator.free(ieee_bytes);

    const lo = try manifest.parse(testing.allocator, lo_bytes);
    defer lo.deinit();
    const ieee = try manifest.parse(testing.allocator, ieee_bytes);
    defer ieee.deinit();

    // `0.1+0.2`: IEEE gives 0x…334, LibreOffice writes 0x…333 (exactly
    // the nearest double to 0.3). One ULP, and the reason `(0.1+0.2)=0.3`
    // comes back TRUE there.
    const lo_sum = lo.value.find("Spec", "A2").?;
    const ieee_sum = ieee.value.find("Spec", "A2").?;
    try testing.expectEqualStrings("0x3FD3333333333334", ieee_sum.bits.?);
    try testing.expectEqualStrings("0x3FD3333333333333", lo_sum.bits.?);
    try testing.expect(!precedence.equalValues(ieee_sum, lo_sum, null));

    // `1/3` diverges the same way.
    try testing.expectEqualStrings("0x3FD5555555555555", ieee.value.find("Spec", "A3").?.bits.?);
    try testing.expectEqualStrings("0x3FD555555555554F", lo.value.find("Spec", "A3").?.bits.?);

    // And the comparison that follows from it: FALSE under exact IEEE,
    // TRUE in LibreOffice.
    try testing.expect(!ieee.value.find("Spec", "A11").?.boolean.?);
    try testing.expect(lo.value.find("Spec", "A11").?.boolean.?);

    // In `.ieee` mode the hand-derived golden decides and LO is only a
    // witness — its disagreement is recorded, never chosen.
    const obs = [_]precedence.Observation{
        .{ .adapter = .libreoffice, .entry = lo_sum },
        .{ .adapter = .hand_spec, .entry = ieee_sum },
    };
    const r = try precedence.resolve(testing.allocator, .ieee, &obs, null);
    defer r.deinit(testing.allocator);
    try testing.expectEqual(provenance.Adapter.hand_spec, r.chosen.?.adapter);
    try testing.expectEqual(@as(usize, 1), r.conflicts.len);
}

test "recorded divergence: SQRT(-1) is #NUM! by the spec, #VALUE! in LibreOffice" {
    // A domain failure, not a type failure — so Excel documents #NUM!.
    // LibreOffice returns #VALUE!. Both are self-consistent; only an
    // independent expectation can say which answers the question.
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();

    const lo_bytes = readFixture(io, testing.allocator, "libreoffice_oracle_suite.json") catch
        return error.SkipZigTest;
    defer testing.allocator.free(lo_bytes);
    const spec_bytes = readFixture(io, testing.allocator, "hand_spec_excel.json") catch
        return error.SkipZigTest;
    defer testing.allocator.free(spec_bytes);

    const lo = try manifest.parse(testing.allocator, lo_bytes);
    defer lo.deinit();
    const spec = try manifest.parse(testing.allocator, spec_bytes);
    defer spec.deinit();

    const lo_sqrt = lo.value.find("Spec", "A7").?;
    const spec_sqrt = spec.value.find("Spec", "A7").?;
    try testing.expectEqualStrings("#VALUE!", lo_sqrt.error_spelling.?);
    try testing.expectEqualStrings("#NUM!", spec_sqrt.error_spelling.?);
    try testing.expect(!precedence.equalValues(spec_sqrt, lo_sqrt, null));

    // `.excel` fidelity with no Excel recording yet: the hand-derived
    // suite outranks LibreOffice, so it decides and LO's answer is kept
    // as the conflict. When the Excel leg lands it will outrank both.
    const obs = [_]precedence.Observation{
        .{ .adapter = .libreoffice, .entry = lo_sqrt },
        .{ .adapter = .hand_spec, .entry = spec_sqrt },
    };
    const r = try precedence.resolve(testing.allocator, .excel, &obs, null);
    defer r.deinit(testing.allocator);
    try testing.expectEqual(provenance.Adapter.hand_spec, r.chosen.?.adapter);
    try testing.expectEqual(@as(usize, 1), r.conflicts.len);
    try testing.expectEqual(provenance.Adapter.libreoffice, r.conflicts[0].dissenting.adapter);

    const text = try precedence.report(testing.allocator, r.conflicts);
    defer testing.allocator.free(text);
    try testing.expect(std.mem.indexOf(u8, text, "DIVERGENCE Spec!A7") != null);
}

test "agreement is recorded too: the cases where both legs match" {
    // A harness that only ever reports disagreement is not measuring
    // agreement — it is measuring nothing, and would look identical if
    // the comparison were broken.
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();

    const lo_bytes = readFixture(io, testing.allocator, "libreoffice_oracle_suite.json") catch
        return error.SkipZigTest;
    defer testing.allocator.free(lo_bytes);
    const spec_bytes = readFixture(io, testing.allocator, "hand_spec_excel.json") catch
        return error.SkipZigTest;
    defer testing.allocator.free(spec_bytes);

    const lo = try manifest.parse(testing.allocator, lo_bytes);
    defer lo.deinit();
    const spec = try manifest.parse(testing.allocator, spec_bytes);
    defer spec.deinit();

    // Every hand-derived `.excel` case except the one known divergence
    // must match what LibreOffice produced.
    var agreed: usize = 0;
    for (spec.value.cells) |expected| {
        const actual = lo.value.find(expected.sheet, expected.ref) orelse continue;
        if (std.mem.eql(u8, expected.ref, "A7")) continue; // the recorded divergence
        if (!precedence.equalValues(expected, actual, null)) {
            std.debug.print(
                "unexpected divergence at {s}!{s} ({s})\n",
                .{ expected.sheet, expected.ref, expected.formula orelse "?" },
            );
            return error.TestExpectedEqual;
        }
        agreed += 1;
    }
    // Operator precedence, error taxonomy, coercion, the percent
    // operator and overflow behaviour — all confirmed by an independent
    // engine against an independently-derived expectation.
    try testing.expect(agreed >= 10);
}

fn containsIgnoreCase(haystack: []const u8, needle: []const u8) bool {
    if (needle.len > haystack.len) return false;
    var i: usize = 0;
    while (i + needle.len <= haystack.len) : (i += 1) {
        if (std.ascii.eqlIgnoreCase(haystack[i .. i + needle.len], needle)) return true;
    }
    return false;
}
