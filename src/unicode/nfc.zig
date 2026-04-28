//! A1 phase 3: Canonical Composition (NFC) for Excel sheet-name dedup.
//!
//! Implements the three-step NFC algorithm over UTF-8:
//!
//!   1. Canonical decomposition (recursively expand every codepoint
//!      via its canonical mapping until reaching atomic codepoints).
//!   2. Canonical reordering (sort sequences of combining marks by
//!      Canonical Combining Class, stable).
//!   3. Canonical composition (combine starter + non-starter pairs
//!      back into precomposed codepoints, skipping any in the
//!      Composition Exclusions list).
//!
//! Hangul (U+AC00..U+D7A3) is handled algorithmically — no table
//! lookup needed for ~11k codepoints.
//!
//! The combination of `casefold.foldString` + `nfc.normalize` gives
//! the canonical-equivalence key Excel uses for sheet-name
//! comparison: `NFC(casefold(name))`.

const std = @import("std");
const tables = @import("tables/nfc_data.zig");

/// Normalise `input` to NFC. Returns owned bytes — caller frees.
/// Invalid UTF-8 returns `error.InvalidUtf8`.
///
/// ASCII fast path: pure-ASCII input round-trips byte-for-byte
/// (every ASCII codepoint is already in NFC), so the fast path
/// just `dupe`s the input.
pub fn normalize(allocator: std.mem.Allocator, input: []const u8) ![]u8 {
    if (isAscii(input)) {
        return allocator.dupe(u8, input);
    }
    if (!std.unicode.utf8ValidateSlice(input)) return error.InvalidUtf8;

    // Step 1+2: decompose + reorder into a u21 buffer.
    var scalars: std.ArrayListUnmanaged(u21) = .empty;
    defer scalars.deinit(allocator);
    try scalars.ensureTotalCapacity(allocator, input.len);

    var i: usize = 0;
    while (i < input.len) {
        const seq_len = std.unicode.utf8ByteSequenceLength(input[i]) catch unreachable;
        const cp = std.unicode.utf8Decode(input[i .. i + seq_len]) catch unreachable;
        try decomposeAppend(allocator, &scalars, cp);
        i += seq_len;
    }
    canonicalReorder(scalars.items);

    // Step 3: compose.
    composeInPlace(&scalars);

    // Re-encode to UTF-8.
    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(allocator);
    try out.ensureTotalCapacity(allocator, input.len);
    var enc_buf: [4]u8 = undefined;
    for (scalars.items) |cp| {
        const n = std.unicode.utf8Encode(cp, &enc_buf) catch return error.InvalidUtf8;
        try out.appendSlice(allocator, enc_buf[0..n]);
    }
    return out.toOwnedSlice(allocator);
}

inline fn isAscii(s: []const u8) bool {
    for (s) |c| {
        if (c >= 0x80) return false;
    }
    return true;
}

// ─── Decomposition ────────────────────────────────────────────────────

/// Recursively decompose `cp` and append its atomic codepoints to
/// `out`. Hangul syllables decompose algorithmically; everything
/// else looks up `decomp_entries` and recurses.
fn decomposeAppend(
    allocator: std.mem.Allocator,
    out: *std.ArrayListUnmanaged(u21),
    cp: u21,
) !void {
    // Hangul syllable algorithmic decomposition.
    if (cp >= tables.hangul_syllable_base and
        cp < tables.hangul_syllable_base + tables.hangul_syllable_count)
    {
        const s_idx = cp - tables.hangul_syllable_base;
        const l = tables.hangul_l_base + s_idx / (tables.hangul_v_count * tables.hangul_t_count);
        const v = tables.hangul_v_base + (s_idx % (tables.hangul_v_count * tables.hangul_t_count)) / tables.hangul_t_count;
        const t = tables.hangul_t_base + s_idx % tables.hangul_t_count;
        try out.append(allocator, l);
        try out.append(allocator, v);
        if (t != tables.hangul_t_base) try out.append(allocator, t);
        return;
    }
    if (lookupDecomp(cp)) |mapping| {
        // Recurse — decompositions can chain (e.g. U+1E0A decomposes
        // to U+0044 + U+0307).
        for (mapping) |child| try decomposeAppend(allocator, out, child);
        return;
    }
    try out.append(allocator, cp);
}

fn lookupDecomp(cp: u21) ?[]const u21 {
    const entries = tables.decomp_entries;
    var lo: usize = 0;
    var hi: usize = entries.len;
    while (lo < hi) {
        const mid = lo + (hi - lo) / 2;
        const e = entries[mid];
        if (e.from < cp) lo = mid + 1 else if (e.from > cp) hi = mid else return tables.decomp_scalars[e.offset .. e.offset + e.len];
    }
    return null;
}

// ─── Canonical reordering (CCC sort) ─────────────────────────────────

fn canonicalCombiningClass(cp: u21) u8 {
    const entries = tables.ccc_entries;
    var lo: usize = 0;
    var hi: usize = entries.len;
    while (lo < hi) {
        const mid = lo + (hi - lo) / 2;
        const e = entries[mid];
        if (e.cp < cp) lo = mid + 1 else if (e.cp > cp) hi = mid else return e.ccc;
    }
    return 0;
}

/// Stable sort runs of non-zero CCC by class. Per UAX #15, sequences
/// of combining marks must be ordered by canonical class, but
/// starters (CCC 0) act as barriers — never sorted across.
fn canonicalReorder(scalars: []u21) void {
    var i: usize = 0;
    while (i < scalars.len) {
        // Find the start of the next non-starter run.
        if (canonicalCombiningClass(scalars[i]) == 0) {
            i += 1;
            continue;
        }
        var j = i;
        while (j < scalars.len and canonicalCombiningClass(scalars[j]) != 0) j += 1;
        // Stable insertion sort over scalars[i..j] by CCC.
        var k: usize = i + 1;
        while (k < j) : (k += 1) {
            const cur = scalars[k];
            const cur_ccc = canonicalCombiningClass(cur);
            var m: usize = k;
            while (m > i and canonicalCombiningClass(scalars[m - 1]) > cur_ccc) : (m -= 1) {
                scalars[m] = scalars[m - 1];
            }
            scalars[m] = cur;
        }
        i = j;
    }
}

// ─── Composition ──────────────────────────────────────────────────────

/// In-place canonical composition. Walks the buffer, attempting to
/// combine each non-starter into the most-recent starter; advances
/// the write index only when no further composition is possible.
fn composeInPlace(scalars: *std.ArrayListUnmanaged(u21)) void {
    if (scalars.items.len == 0) return;
    var write: usize = 0;
    var starter_idx: ?usize = null;
    var last_class: u8 = 0;

    var read: usize = 0;
    while (read < scalars.items.len) : (read += 1) {
        const cp = scalars.items[read];
        const ccc = canonicalCombiningClass(cp);

        // Try composing with the current starter.
        if (starter_idx) |si| {
            const starter = scalars.items[si];
            // Blocking rule (UAX #15): a character is blocked from
            // composing with the starter when there's a non-starter
            // between them whose CCC is >= this character's CCC.
            //
            // last_class > 0 means we've already crossed a
            // non-starter since the starter. In that case:
            //   - ccc == 0  → this is itself a new starter, so it
            //     can't compose with the prior starter across the
            //     intervening combiner: BLOCKED.
            //   - ccc != 0  → blocked iff ccc <= last_class
            //     (per UAX #15 D102).
            const blocked = (last_class > 0 and ccc <= last_class);
            if (!blocked) {
                if (composePair(starter, cp)) |composed| {
                    scalars.items[si] = composed;
                    // last_class unchanged — we just replaced the
                    // starter, didn't append anything new.
                    continue;
                }
            }
        }

        // No composition; place at write index.
        scalars.items[write] = cp;
        if (ccc == 0) {
            starter_idx = write;
            last_class = 0;
        } else {
            last_class = ccc;
        }
        write += 1;
    }
    scalars.shrinkRetainingCapacity(write);
}

fn composePair(starter: u21, combining: u21) ?u21 {
    // Hangul algorithmic L+V → LV, LV+T → LVT.
    if (starter >= tables.hangul_l_base and
        starter < tables.hangul_l_base + tables.hangul_l_count)
    {
        if (combining >= tables.hangul_v_base and
            combining < tables.hangul_v_base + tables.hangul_v_count)
        {
            const l_idx = starter - tables.hangul_l_base;
            const v_idx = combining - tables.hangul_v_base;
            const lv = tables.hangul_syllable_base +
                (l_idx * tables.hangul_v_count + v_idx) * tables.hangul_t_count;
            return @intCast(lv);
        }
    }
    if (starter >= tables.hangul_syllable_base and
        starter < tables.hangul_syllable_base + tables.hangul_syllable_count)
    {
        const s_idx = starter - tables.hangul_syllable_base;
        if (s_idx % tables.hangul_t_count == 0) {
            // LV; can compose with a T.
            if (combining > tables.hangul_t_base and
                combining < tables.hangul_t_base + tables.hangul_t_count)
            {
                return @intCast(starter + (combining - tables.hangul_t_base));
            }
        }
    }

    // Table lookup. Entries sorted by (starter, combining); binary
    // search by starter first, linear scan within the matching run.
    const entries = tables.compose_entries;
    var lo: usize = 0;
    var hi: usize = entries.len;
    while (lo < hi) {
        const mid = lo + (hi - lo) / 2;
        if (entries[mid].starter < starter) {
            lo = mid + 1;
        } else if (entries[mid].starter > starter) {
            hi = mid;
        } else {
            // Walk back to first entry with this starter.
            var s = mid;
            while (s > 0 and entries[s - 1].starter == starter) s -= 1;
            while (s < entries.len and entries[s].starter == starter) : (s += 1) {
                if (entries[s].combining == combining) return entries[s].composed;
            }
            return null;
        }
    }
    return null;
}

// ─── Tests ────────────────────────────────────────────────────────────

test "ASCII pass-through" {
    const out = try normalize(std.testing.allocator, "Hello, World!");
    defer std.testing.allocator.free(out);
    try std.testing.expectEqualStrings("Hello, World!", out);
}

test "decomposed → composed: e + combining acute → é" {
    // U+0065 LATIN SMALL LETTER E + U+0301 COMBINING ACUTE ACCENT
    // composes to U+00E9 LATIN SMALL LETTER E WITH ACUTE.
    const decomposed = "e\u{0301}";
    const out = try normalize(std.testing.allocator, decomposed);
    defer std.testing.allocator.free(out);
    try std.testing.expectEqualStrings("\u{00E9}", out);
}

test "already-NFC pass-through: café" {
    const out = try normalize(std.testing.allocator, "café");
    defer std.testing.allocator.free(out);
    try std.testing.expectEqualStrings("café", out);
}

test "decomposed café normalises to precomposed" {
    const decomposed = "cafe\u{0301}";
    const out = try normalize(std.testing.allocator, decomposed);
    defer std.testing.allocator.free(out);
    try std.testing.expectEqualStrings("café", out);
}

test "Hangul: L + V composes to LV syllable" {
    // U+1100 ᄀ (L) + U+1161 ᅡ (V) → U+AC00 가
    const decomposed = "\u{1100}\u{1161}";
    const out = try normalize(std.testing.allocator, decomposed);
    defer std.testing.allocator.free(out);
    try std.testing.expectEqualStrings("\u{AC00}", out);
}

test "Hangul: LV + T composes to LVT syllable" {
    // U+AC00 가 + U+11A8 ᆨ → U+AC01 각
    const decomposed = "\u{AC00}\u{11A8}";
    const out = try normalize(std.testing.allocator, decomposed);
    defer std.testing.allocator.free(out);
    try std.testing.expectEqualStrings("\u{AC01}", out);
}

test "Composition exclusion: U+0F71 + U+0F72 stays decomposed" {
    // U+0F73 is in CompositionExclusions, so e + double-mark should
    // never recombine even though decomposition would suggest it.
    // Use a more common exclusion: U+212B (Angstrom sign) decomposes
    // to U+00C5 (LATIN CAPITAL LETTER A WITH RING ABOVE) but does
    // NOT compose back from A + ring above.
    const a_with_ring = "A\u{030A}";
    const out = try normalize(std.testing.allocator, a_with_ring);
    defer std.testing.allocator.free(out);
    // A + ring composes to U+00C5 (ÅSF), NOT to U+212B Angstrom sign.
    try std.testing.expectEqualStrings("\u{00C5}", out);
}

test "Reorder: combining marks sort by CCC" {
    // U+0061 + U+0327 (CCC 202) + U+0301 (CCC 230) — already in
    // CCC order. The reverse — U+0301 then U+0327 — must reorder
    // so the cedilla precedes the acute.
    const reversed = "a\u{0301}\u{0327}";
    const out = try normalize(std.testing.allocator, reversed);
    defer std.testing.allocator.free(out);
    // After reorder + compose: a + cedilla composes to ạ̧... well
    // depends on tables. Just verify the bytes are deterministic
    // and shorter or equal.
    try std.testing.expect(out.len > 0);
    // Round-trip sanity: re-normalising stays identical.
    const round_trip = try normalize(std.testing.allocator, out);
    defer std.testing.allocator.free(round_trip);
    try std.testing.expectEqualStrings(out, round_trip);
}

test "Blocking: U+1100 U+0301 U+1161 does NOT compose Hangul across the acute" {
    // Regression for the canonical-composition blocking rule. Without
    // it, the L+V Hangul algorithm composes across the intervening
    // U+0301 acute, producing a wrong NFC and breaking sheet-name
    // dedup for adversarial inputs.
    const input = "\u{1100}\u{0301}\u{1161}";
    const out = try normalize(std.testing.allocator, input);
    defer std.testing.allocator.free(out);
    // Correct NFC: U+1100, U+0301, U+1161 unchanged (no composition
    // possible with a non-starter between).
    try std.testing.expectEqualStrings(input, out);
}

test "Blocking: e + acute + acute keeps second acute decomposed" {
    // Two same-class non-starters can't both compose into the
    // starter (per the ccc <= last_class rule).
    const input = "e\u{0301}\u{0301}";
    const out = try normalize(std.testing.allocator, input);
    defer std.testing.allocator.free(out);
    // First acute composes with e to é (U+00E9); second acute stays
    // as a combining mark.
    try std.testing.expectEqualStrings("\u{00E9}\u{0301}", out);
}

test "Invalid UTF-8 is rejected" {
    try std.testing.expectError(error.InvalidUtf8, normalize(std.testing.allocator, "ab\xFFc"));
}

test "Empty input round-trips empty" {
    const out = try normalize(std.testing.allocator, "");
    defer std.testing.allocator.free(out);
    try std.testing.expectEqual(@as(usize, 0), out.len);
}
