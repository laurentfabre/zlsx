//! The four oracle adapters and what each one is good for
//! (M1b, `goal_formula.md` §8.2).
//!
//! An adapter is not just "a source of values" — it is a source with
//! specific blind spots, and the blind spots are what this table
//! encodes. Excel-for-Mac cannot be trusted for `CHAR(200)` because its
//! code page is not CP-1252. No running application can be trusted for
//! `RAND()`, because a draw is not a contract. LibreOffice will re-save
//! a document it never calculated. Writing those facts down as data,
//! next to the adapter they belong to, is what stops a future batch
//! from recording a golden that was never valid.
//!
//! `provenance.Adapter` is the identity; this is the capability matrix.

const std = @import("std");
const provenance = @import("provenance.zig");
const sentinel = @import("sentinel.zig");

pub const Adapter = provenance.Adapter;

pub const Capabilities = struct {
    adapter: Adapter,
    /// One line on what this adapter is.
    role: []const u8,
    /// Runs a real spreadsheet application, so its output has to prove
    /// it recalculated before anything may be recorded.
    requires_sentinels: bool,
    /// Sentinel kinds this adapter's runs must satisfy.
    sentinel_kinds: []const sentinel.Kind,
    /// §8.2: volatiles are excluded from every EXTERNAL value oracle.
    /// A hand-derived value for `RAND()` is a statement about the
    /// function's contract, not a draw, so the spec suite keeps them.
    excludes_volatiles: bool,
    /// Mac Excel's CHAR/CODE above 127 follow the Mac code page rather
    /// than CP-1252, so those goldens must come from elsewhere.
    excludes_char_code_high: bool,
    /// Can this adapter ever be the deciding authority for a value, in
    /// any fidelity mode? The corpus never can — it is a consistency
    /// signal across many real workbooks, not a statement of what any
    /// one answer should be.
    can_be_authority: bool,
};

pub const all = [_]Capabilities{
    .{
        .adapter = .excel_mac,
        .role = "Excel for Mac over AppleScript; the reference implementation for .excel fidelity",
        .requires_sentinels = true,
        // Both value classes: one proves a calculation happened at all,
        // the other proves the dependency tree was rebuilt.
        .sentinel_kinds = &.{ .stale_value, .stale_dependency },
        .excludes_volatiles = true,
        .excludes_char_code_high = true,
        .can_be_authority = true,
    },
    .{
        .adapter = .libreoffice,
        .role = "LibreOffice Calc, pinned build + dedicated profile; a second independent engine",
        .requires_sentinels = true,
        // LO's specific failure is re-saving a document it only loaded,
        // which a redrawn volatile catches.
        .sentinel_kinds = &.{ .stale_value, .volatile_draw },
        .excludes_volatiles = true,
        .excludes_char_code_high = false,
        .can_be_authority = true,
    },
    .{
        .adapter = .hand_spec,
        .role = "hand-derived from the specification; anchors divergence points and leads .ieee fidelity",
        // Nothing was run, so there is nothing to prove ran.
        .requires_sentinels = false,
        .sentinel_kinds = &.{},
        .excludes_volatiles = false,
        .excludes_char_code_high = false,
        .can_be_authority = true,
    },
    .{
        .adapter = .corpus,
        .role = "screened real-world workbooks; a consistency signal, never a primary authority",
        .requires_sentinels = false,
        .sentinel_kinds = &.{},
        .excludes_volatiles = true,
        .excludes_char_code_high = false,
        .can_be_authority = false,
    },
};

pub fn get(adapter: Adapter) Capabilities {
    for (all) |c| {
        if (c.adapter == adapter) return c;
    }
    unreachable; // `all` covers the enum — the test below proves it
}

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

test "every adapter has exactly one capability row" {
    // An adapter added to the enum without a row here would hit
    // `unreachable` in `get` at runtime; this turns that into a test
    // failure at the point the enum changes.
    inline for (std.meta.fields(Adapter)) |field| {
        const adapter: Adapter = @enumFromInt(field.value);
        var seen: usize = 0;
        for (all) |c| {
            if (c.adapter == adapter) seen += 1;
        }
        try testing.expectEqual(@as(usize, 1), seen);
    }
    try testing.expectEqual(std.meta.fields(Adapter).len, all.len);
}

test "every adapter that runs an application requires sentinels that can prove it" {
    for (all) |c| {
        try testing.expectEqual(c.adapter.isExternalApp(), c.requires_sentinels);
        if (!c.requires_sentinels) {
            try testing.expectEqual(@as(usize, 0), c.sentinel_kinds.len);
            continue;
        }
        // A sentinel set that cannot prove a recalculation is worse than
        // none: it looks like a check and passes vacuously.
        var set: [4]sentinel.Sentinel = undefined;
        for (c.sentinel_kinds, 0..) |k, i| {
            set[i] = .{ .kind = k, .sheet = "S", .ref = "A1", .planted = "0" };
        }
        try testing.expect(sentinel.hasProof(set[0..c.sentinel_kinds.len]));
    }
}

test "volatile exclusion tracks the external-application distinction" {
    // §8.2 excludes volatiles from every external value oracle. The
    // corpus is not an application, but it IS a collection of values
    // somebody else's application produced, so it excludes them too.
    try testing.expect(get(.excel_mac).excludes_volatiles);
    try testing.expect(get(.libreoffice).excludes_volatiles);
    try testing.expect(get(.corpus).excludes_volatiles);
    // Only the hand-derived suite keeps them, because its "value" for a
    // volatile is a documented contract rather than an observed draw.
    try testing.expect(!get(.hand_spec).excludes_volatiles);
}

test "only the Mac Excel leg drops CHAR/CODE high bytes" {
    try testing.expect(get(.excel_mac).excludes_char_code_high);
    for ([_]Adapter{ .libreoffice, .hand_spec, .corpus }) |a| {
        try testing.expect(!get(a).excludes_char_code_high);
    }
}

test "the corpus can never be an authority" {
    try testing.expect(!get(.corpus).can_be_authority);
    for ([_]Adapter{ .excel_mac, .libreoffice, .hand_spec }) |a| {
        try testing.expect(get(a).can_be_authority);
    }
}
