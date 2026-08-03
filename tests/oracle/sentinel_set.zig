//! The sentinel cells every oracle input carries (M1b).
//!
//! This is the Zig half of a two-sided contract: `sentinel_set.zig`
//! must agree, cell for cell and value for value, with what
//! `scripts/oracle/build_inputs.py` plants. If they drift, the checker
//! looks at cells the builder never planted, finds them missing, and
//! rejects every run — loudly, which is the right direction for that
//! failure to fall.

const std = @import("std");
const sentinel = @import("sentinel.zig");

pub const sheet = "Sentinels";

/// Mirrors `PLANTED_*` in `scripts/oracle/build_inputs.py`.
pub const planted_stale_value = "999";
pub const planted_stale_dependency = "111";
pub const planted_volatile = "0.123456789";

pub const all = [_]sentinel.Sentinel{
    .{
        .kind = .stale_value,
        .sheet = sheet,
        .ref = "B1",
        .planted = planted_stale_value,
        .rationale = "=1+1 cannot be 999",
    },
    .{
        .kind = .stale_dependency,
        .sheet = sheet,
        .ref = "B2",
        .planted = planted_stale_dependency,
        .rationale = "=C2*2 over a three-deep chain with an inverted calcChain cannot be 111",
    },
    .{
        .kind = .volatile_draw,
        .sheet = sheet,
        .ref = "B3",
        .planted = planted_volatile,
        .rationale = "=RAND() pinned to a fixed cached value",
    },
};

/// Sentinels an adapter's runs must satisfy. Excel is driven with a
/// dependency rebuild, so it answers for both value-class sentinels;
/// LibreOffice's characteristic failure is re-saving an uncalculated
/// document, which the volatile draw catches.
pub fn forAdapter(adapter: @import("provenance.zig").Adapter) []const sentinel.Sentinel {
    return switch (adapter) {
        // Every sentinel: the Excel leg is the one that must prove the
        // dependency tree was rebuilt.
        .excel_mac => &all,
        .libreoffice => &all,
        // Nothing was run, so there is nothing to prove ran.
        .hand_spec, .corpus => &.{},
    };
}

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

test "the set covers all three sentinel kinds, on distinct cells" {
    var seen: std.EnumSet(sentinel.Kind) = .initEmpty();
    for (all) |s| {
        try testing.expect(!seen.contains(s.kind)); // no duplicate kinds
        seen.insert(s.kind);
        try testing.expectEqualStrings(sheet, s.sheet);
        try testing.expect(s.planted.len > 0);
        // Every sentinel must explain itself: a rejection message that
        // cannot say why the planted value was impossible sends whoever
        // reads it back to the source.
        try testing.expect(s.rationale.len > 0);
    }
    try testing.expectEqual(@as(usize, 3), seen.count());

    for (all, 0..) |a, i| {
        for (all[i + 1 ..]) |b| {
            try testing.expect(!std.mem.eql(u8, a.ref, b.ref));
        }
    }
}

test "every application adapter gets a set that can actually prove a recalc" {
    try testing.expect(sentinel.hasProof(forAdapter(.excel_mac)));
    try testing.expect(sentinel.hasProof(forAdapter(.libreoffice)));
    // The non-application adapters get an empty set, which correctly
    // proves nothing — they have nothing to prove.
    try testing.expectEqual(@as(usize, 0), forAdapter(.hand_spec).len);
    try testing.expectEqual(@as(usize, 0), forAdapter(.corpus).len);
}

test "planted values match the builder script" {
    // The two-sided contract, checked against the Python source rather
    // than trusted. A drift here silently rejects every future run.
    const src = @embedFile("build_inputs_py");
    var buf: [128]u8 = undefined;

    const pairs = [_]struct { name: []const u8, value: []const u8 }{
        .{ .name = "PLANTED_STALE_VALUE", .value = planted_stale_value },
        .{ .name = "PLANTED_STALE_DEPENDENCY", .value = planted_stale_dependency },
        .{ .name = "PLANTED_VOLATILE", .value = planted_volatile },
    };
    for (pairs) |p| {
        const needle = try std.fmt.bufPrint(&buf, "{s} = \"{s}\"", .{ p.name, p.value });
        if (std.mem.indexOf(u8, src, needle) == null) {
            std.debug.print(
                "sentinel drift: build_inputs.py does not plant {s} = \"{s}\"\n",
                .{ p.name, p.value },
            );
            return error.TestExpectedEqual;
        }
    }

    // The cell references have to match too.
    for (all) |s| {
        const needle = try std.fmt.bufPrint(&buf, "Cell(\"{s}\"", .{s.ref});
        try testing.expect(std.mem.indexOf(u8, src, needle) != null);
    }
    try testing.expect(std.mem.indexOf(u8, src, "SENTINEL_SHEET = \"Sentinels\"") != null);
}
