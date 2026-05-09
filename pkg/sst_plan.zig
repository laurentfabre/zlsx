//! Shared-strings extension-plan substrate (B3 iter-wr-1).
//!
//! Storage + dedup types for the SST regeneration pipeline. Lives
//! outside `pkg/workbook.zig` so both Workbook (delta-on-existing-bytes
//! editor) and `xlsx.Writer` (fresh-emit producer) can stage strings
//! through the same shape without a circular module dependency
//! (workbook.zig imports `zlsx`, which contains writer.zig).
//!
//! Two axes:
//!
//! - **Plain entries** — `new_strings` is the dedup pool of unique
//!   string text. `new_strings_index` is the O(1) hash side-index.
//!   Plain entries are deduped against each other and (in Workbook's
//!   delta path) against the existing SST.
//! - **Rich entries** — `new_rich_strings` is the typed `RichRun[]`
//!   pool. Rich entries are NEVER deduped (matches `xlsx.Writer`'s
//!   iter33 policy: hashing the formatted form costs more than it
//!   saves at typical SST sizes).
//!
//! Indices: plain entries occupy `[base_index, base_index + new_strings.len)`;
//! rich entries follow at `[base_index + new_strings.len, ...)`. For a
//! freshly-emitted SST, `base_index` is 0; for a workbook with an
//! existing SST it is the existing entry count.
//!
//! Stdlib only. Zig 0.15.2.

const std = @import("std");

const Allocator = std.mem.Allocator;
const assert = std.debug.assert;

pub const Error = error{OutOfMemory};

/// Side-table record for plain-text deltas whose text matched an
/// existing SST entry. Not used by `xlsx.Writer`'s fresh-emit path
/// (no existing SST to match against), but kept on the plan for
/// Workbook's delta path.
pub const ExistingMatch = struct {
    text: []const u8,
    index: u32,
};

/// One run inside a rich-text SST entry. Mirrors `xlsx.RichTextRun`'s
/// public shape (text + bold + italic + size + color_argb + font_name)
/// so the writer-side wiring can stage runs without an extra
/// translation step. `color_argb` is encoded as the raw 8-hex-digit
/// ARGB value (e.g. "FF0000FF") rather than a `u32`; the SST emit path
/// passes the bytes through `<color rgb="…"/>` verbatim, so producers
/// do their own formatting if they want a different surface. Strike +
/// underline ride along (writer doesn't currently surface them, but
/// OOXML does — keeping the substrate complete avoids a follow-up plan
/// extension).
pub const RichRun = struct {
    text: []const u8,
    bold: bool = false,
    italic: bool = false,
    underline: bool = false,
    strike: bool = false,
    font_name: ?[]const u8 = null,
    font_size: ?f32 = null,
    color_argb: ?[]const u8 = null,
};

/// One rich-text SST entry. `runs` is borrowed by the registrar
/// (`SstExtensionPlan.registerRich` and `registerSharedRichString` dup
/// every run + every owned string field into the plan allocator at
/// registration time, so callers can free their staging buffers
/// immediately after the call returns).
pub const RichEntry = struct {
    runs: []const RichRun,
};

pub const SstExtensionPlan = struct {
    /// True when at least one new entry has been staged across either
    /// axis. Drives the "regenerate vs leave-alone" decision in
    /// Workbook's `applySstExtensionPlan`.
    has_new_strings: bool = false,
    /// Allocator owns: every entry of `new_strings` (duped on insert),
    /// the slice itself.
    new_strings: std.ArrayListUnmanaged([]const u8) = .empty,
    /// Side table: deltas whose text matched an existing SST entry.
    /// Allows `indexOf` to resolve those without rescanning the SST.
    /// Allocator owns each `text` slice.
    existing_matches: std.ArrayListUnmanaged(ExistingMatch) = .empty,
    /// Index of the FIRST new string within the regenerated SST. For
    /// an existing-SST workbook this is the existing entry count;
    /// for a freshly-created SST (Writer fresh-emit) it's 0.
    base_index: u32 = 0,
    /// Tracks whether the SST part already existed at plan-build time.
    /// Drives the `replacePart` vs `addPart + workbook.xml.rels splice`
    /// branch in `applySstExtensionPlan`. Always false for
    /// `xlsx.Writer`'s fresh-emit pipeline.
    sst_part_exists: bool = false,

    /// Rich-text new-entries axis. Each entry carries an array of
    /// typed `RichRun`s; emitters render one
    /// `<si><r><rPr/>…<t/></r>…</si>` block per entry, alongside the
    /// plain `new_strings` blocks. Indices in this axis follow the
    /// plain ones — the first rich entry sits at
    /// `base_index + new_strings.len`.
    new_rich_strings: std.ArrayListUnmanaged(RichEntry) = .empty,

    /// O(1) hash index over `new_strings` keyed by the (already-duped)
    /// owned text slice. Maps text → index in `new_strings` (NOT the
    /// SST index — add `base_index` to get the latter). Backs the fast
    /// path in `indexOf` and `registerNewPlain`; matters most under
    /// `xlsx.Writer`'s hot writeRow loop where the previous linear-scan
    /// dedup was O(n²) over thousands of string cells.
    new_strings_index: std.StringHashMapUnmanaged(u32) = .empty,

    pub fn deinit(self: *SstExtensionPlan, allocator: Allocator) void {
        for (self.new_strings.items) |s| allocator.free(s);
        self.new_strings.deinit(allocator);
        self.new_strings_index.deinit(allocator);
        for (self.existing_matches.items) |em| allocator.free(em.text);
        self.existing_matches.deinit(allocator);
        for (self.new_rich_strings.items) |entry| {
            for (entry.runs) |r| {
                allocator.free(r.text);
                if (r.font_name) |n| allocator.free(n);
                if (r.color_argb) |c| allocator.free(c);
            }
            allocator.free(entry.runs);
        }
        self.new_rich_strings.deinit(allocator);
        self.* = undefined;
    }

    /// Resolve the SST index for a (raw, unescaped) plain string.
    /// Returns null when `s` was never staged into this plan. Hits the
    /// O(1) hash on the new-strings axis first; falls back to the
    /// existing-matches side-table linear scan only on miss.
    pub fn indexOf(self: *const SstExtensionPlan, s: []const u8) ?u32 {
        if (self.new_strings_index.get(s)) |i| {
            return self.base_index + i;
        }
        for (self.existing_matches.items) |em| {
            if (std.mem.eql(u8, em.text, s)) return em.index;
        }
        return null;
    }

    /// Resolve the SST index for a rich-text entry by reference
    /// (pointer equality on the `RichEntry` slot in
    /// `new_rich_strings`). Rich entries are indexed AFTER plain new
    /// strings; the first rich entry lands at
    /// `base_index + new_strings.len`. Returns null when `entry`'s
    /// pointer is not one this plan handed out. Callers staging a
    /// rich entry typically retain the pointer they got back from
    /// `registerNewRich` (or `registerSharedRichString` on the
    /// Workbook side) and pass it straight through.
    pub fn indexOfRich(self: *const SstExtensionPlan, entry: *const RichEntry) ?u32 {
        for (self.new_rich_strings.items, 0..) |*staged, i| {
            if (staged == entry) {
                const rich_offset: u32 = @intCast(i);
                return self.base_index +
                    @as(u32, @intCast(self.new_strings.items.len)) +
                    rich_offset;
            }
        }
        return null;
    }

    /// Fresh-emit registration for a plain string. Used by
    /// `xlsx.Writer` (iter-wr-1) — there's no existing SST to dedup
    /// against, just the plan's own `new_strings`. Dedup is O(1) via
    /// the hash side-index. Returns the SST index assigned to `s`
    /// (existing match or freshly inserted slot). The plan owns the
    /// duped bytes; callers free their own staging copies.
    pub fn registerNewPlain(
        self: *SstExtensionPlan,
        allocator: Allocator,
        s: []const u8,
    ) Error!u32 {
        if (self.new_strings_index.get(s)) |i| {
            return self.base_index + i;
        }
        const owned = try allocator.dupe(u8, s);
        errdefer allocator.free(owned);
        const idx: u32 = @intCast(self.new_strings.items.len);
        try self.new_strings.append(allocator, owned);
        errdefer _ = self.new_strings.pop();
        try self.new_strings_index.put(allocator, owned, idx);
        self.has_new_strings = true;
        return self.base_index + idx;
    }

    /// Fresh-emit registration for a rich entry. Dups every owned byte
    /// (run text, font_name, color_argb) into `allocator` so callers
    /// can free their staging buffers immediately. Rich entries are
    /// NOT de-duplicated. Returns a pointer to the staged entry; the
    /// pointer is stable for the lifetime of the plan (entries are
    /// only ever appended, never reordered) and is the input to
    /// `indexOfRich`.
    ///
    /// Atomicity: every dupe runs BEFORE the entry is appended; an
    /// OOM mid-dupe walks back the partial state and leaves the plan
    /// untouched.
    pub fn registerNewRich(
        self: *SstExtensionPlan,
        allocator: Allocator,
        runs: []const RichRun,
    ) Error!*const RichEntry {
        const owned_runs = try allocator.alloc(RichRun, runs.len);
        var built: usize = 0;
        errdefer {
            for (owned_runs[0..built]) |r| {
                allocator.free(r.text);
                if (r.font_name) |n| allocator.free(n);
                if (r.color_argb) |c| allocator.free(c);
            }
            allocator.free(owned_runs);
        }
        for (runs, 0..) |r, i| {
            const owned_text = try allocator.dupe(u8, r.text);
            errdefer allocator.free(owned_text);
            const owned_fn: ?[]const u8 = if (r.font_name) |n| try allocator.dupe(u8, n) else null;
            errdefer if (owned_fn) |n| allocator.free(n);
            const owned_c: ?[]const u8 = if (r.color_argb) |c| try allocator.dupe(u8, c) else null;
            owned_runs[i] = .{
                .text = owned_text,
                .bold = r.bold,
                .italic = r.italic,
                .underline = r.underline,
                .strike = r.strike,
                .font_name = owned_fn,
                .font_size = r.font_size,
                .color_argb = owned_c,
            };
            built = i + 1;
        }

        try self.new_rich_strings.append(allocator, .{ .runs = owned_runs });
        self.has_new_strings = true;
        return &self.new_rich_strings.items[self.new_rich_strings.items.len - 1];
    }
};

// ─── Tests ────────────────────────────────────────────────────────────

test "SstExtensionPlan: plain register dedups via hash index" {
    const a = std.testing.allocator;
    var plan: SstExtensionPlan = .{};
    defer plan.deinit(a);

    const idx1 = try plan.registerNewPlain(a, "hello");
    const idx2 = try plan.registerNewPlain(a, "world");
    const idx3 = try plan.registerNewPlain(a, "hello"); // dedup

    try std.testing.expectEqual(@as(u32, 0), idx1);
    try std.testing.expectEqual(@as(u32, 1), idx2);
    try std.testing.expectEqual(@as(u32, 0), idx3);
    try std.testing.expectEqual(@as(usize, 2), plan.new_strings.items.len);
    try std.testing.expect(plan.has_new_strings);
}

test "SstExtensionPlan: rich register with no dedup, indexOfRich works" {
    const a = std.testing.allocator;
    var plan: SstExtensionPlan = .{};
    defer plan.deinit(a);

    const r1 = [_]RichRun{.{ .text = "alpha", .bold = true }};
    const r2 = [_]RichRun{.{ .text = "alpha", .bold = true }}; // identical content; still a fresh entry

    const e1 = try plan.registerNewRich(a, &r1);
    const e2 = try plan.registerNewRich(a, &r2);

    try std.testing.expectEqual(@as(usize, 2), plan.new_rich_strings.items.len);
    try std.testing.expectEqual(@as(?u32, 0), plan.indexOfRich(e1));
    try std.testing.expectEqual(@as(?u32, 1), plan.indexOfRich(e2));
}

test "SstExtensionPlan: rich follows plain in index space" {
    const a = std.testing.allocator;
    var plan: SstExtensionPlan = .{};
    defer plan.deinit(a);

    _ = try plan.registerNewPlain(a, "first");
    _ = try plan.registerNewPlain(a, "second");
    const r = [_]RichRun{.{ .text = "rich1" }};
    const e = try plan.registerNewRich(a, &r);

    // Plain count is 2, so first rich sits at index 2.
    try std.testing.expectEqual(@as(?u32, 2), plan.indexOfRich(e));
    try std.testing.expectEqual(@as(?u32, 0), plan.indexOf("first"));
    try std.testing.expectEqual(@as(?u32, 1), plan.indexOf("second"));
}

test "SstExtensionPlan: indexOf miss returns null" {
    const a = std.testing.allocator;
    var plan: SstExtensionPlan = .{};
    defer plan.deinit(a);

    try std.testing.expectEqual(@as(?u32, null), plan.indexOf("nope"));
}

test "SstExtensionPlan: base_index offsets plain + rich indices" {
    const a = std.testing.allocator;
    var plan: SstExtensionPlan = .{ .base_index = 100 };
    defer plan.deinit(a);

    const idx1 = try plan.registerNewPlain(a, "x");
    const r = [_]RichRun{.{ .text = "y" }};
    const e = try plan.registerNewRich(a, &r);

    try std.testing.expectEqual(@as(u32, 100), idx1);
    try std.testing.expectEqual(@as(?u32, 100), plan.indexOf("x"));
    try std.testing.expectEqual(@as(?u32, 101), plan.indexOfRich(e));
}
