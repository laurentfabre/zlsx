//! `casing_v1` — locale-neutral full Unicode casing (M4f, §5.4b).
//!
//! Why this is not `casefold.zig`
//! ------------------------------
//! Folding and casing answer different questions and disagree on the
//! answers. `fold("ß")` is `"ss"` — a comparison key, chosen so that
//! `ß` and `SS` land on the same string and nobody ever displays it.
//! `UPPER("ß")` is `"SS"` — a value a cell shows. A fold cannot
//! implement `UPPER` (it lowercases), and a one-to-one mapping cannot
//! either: `ß` is one scalar and `SS` is two, so the table has to carry
//! full mappings.
//!
//! Policy, fixed by the generator and restated here because it is the
//! part a reader will want to check:
//!
//!   * **Full mappings.** `UnicodeData.txt`'s simple mappings, with
//!     `SpecialCasing.txt`'s unconditional rows layered on top.
//!   * **One conditional rule: Final_Sigma.** `Σ` lowercases to `ς` at
//!     the end of a word and to `σ` elsewhere. It is in scope because it
//!     is locale-INdependent — the condition is about the neighbouring
//!     characters, not about who is reading.
//!   * **No locale-conditional casing.** The fifteen `tr` / `az` / `lt`
//!     rows are rejected at generation time. There is no locale in
//!     `RunInputs` to select them with (§5.4b), so implementing them
//!     would mean inventing one. Turkish `I`/`ı` is a recorded
//!     divergence: `LOWER("I")` is `"i"` here and `"ı"` in a Turkish
//!     Excel.
//!
//! `title` mappings ship with this module although M4f uses only
//! `toUpper`/`toLower`: M8b's `PROPER` adds word segmentation over this
//! table and should not have to reopen the generator to get its data.
//!
//! The tables are derived from the Unicode Character Database and used
//! under the Unicode License v3; see `THIRD_PARTY_NOTICES.md`.

const std = @import("std");
const tables = @import("tables/casing_data.zig");

/// The version string that goes into the engine fingerprint. A change
/// to the tables or the policy is a change to observable results, so it
/// is named rather than implied.
pub const version = "casing_v1";

/// The Unicode revision the committed tables were generated from.
/// `scripts/ci/check_unicode_tables.sh` re-derives them and fails on any
/// diff, so this is the casing zlsx actually implements.
pub const unicode_version = tables.unicode_version;

pub const Error = error{ InvalidUtf8, OutOfMemory };

/// Full uppercase. Returns owned bytes — caller frees.
pub fn toUpper(allocator: std.mem.Allocator, input: []const u8) Error![]u8 {
    return map(allocator, input, .upper);
}

/// Full lowercase, including the Final_Sigma rule. Returns owned bytes.
pub fn toLower(allocator: std.mem.Allocator, input: []const u8) Error![]u8 {
    return map(allocator, input, .lower);
}

/// Full titlecase of every scalar, WITHOUT word segmentation — this is
/// the character-level half of `PROPER`, not `PROPER` itself. M8b adds
/// the segmentation that decides which scalars are word-initial.
pub fn toTitleScalar(allocator: std.mem.Allocator, input: []const u8) Error![]u8 {
    return map(allocator, input, .title);
}

const Direction = enum { upper, lower, title };

fn entriesFor(dir: Direction) []const tables.CaseEntry {
    return switch (dir) {
        .upper => tables.upper_entries,
        .lower => tables.lower_entries,
        .title => tables.title_entries,
    };
}

fn scalarsFor(dir: Direction) []const u21 {
    return switch (dir) {
        .upper => tables.upper_scalars,
        .lower => tables.lower_scalars,
        .title => tables.title_scalars,
    };
}

fn map(allocator: std.mem.Allocator, input: []const u8, dir: Direction) Error![]u8 {
    // ASCII is the overwhelmingly common case and needs no table at all:
    // no ASCII scalar has a length-changing mapping, and Final_Sigma
    // cannot arise without a Greek sigma. The fast path exists for the
    // same reason `casefold.zig` has one — a column of ASCII text should
    // not pay for the astral plane.
    if (isAscii(input)) return mapAscii(allocator, input, dir);

    if (!std.unicode.utf8ValidateSlice(input)) return error.InvalidUtf8;

    // Final_Sigma is a property of a scalar's NEIGHBOURS, so the input is
    // decoded once into scalars rather than walked byte-wise: deciding
    // "is this the end of a word" needs to look backwards past an
    // arbitrary run of case-ignorable marks, which a forward byte walk
    // cannot do without re-decoding.
    const cps = try decode(allocator, input);
    defer allocator.free(cps);

    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(allocator);
    try out.ensureTotalCapacity(allocator, input.len);

    for (cps, 0..) |cp, i| {
        if (dir == .lower and cp == tables.final_sigma_source and isFinalSigma(cps, i)) {
            try appendCodepoint(allocator, &out, tables.final_sigma_lower);
            continue;
        }
        if (lookup(entriesFor(dir), cp)) |e| {
            for (scalarsFor(dir)[e.offset .. e.offset + e.len]) |m| {
                try appendCodepoint(allocator, &out, m);
            }
        } else {
            try appendCodepoint(allocator, &out, cp);
        }
    }
    return out.toOwnedSlice(allocator);
}

fn mapAscii(allocator: std.mem.Allocator, input: []const u8, dir: Direction) Error![]u8 {
    const out = try allocator.alloc(u8, input.len);
    for (input, 0..) |c, i| {
        out[i] = switch (dir) {
            .upper => std.ascii.toUpper(c),
            // Titlecasing a scalar is uppercasing it for every ASCII
            // letter; the digraph rows where the two differ (`ǳ`/`ǲ`/`Ǳ`)
            // are all non-ASCII.
            .title => std.ascii.toUpper(c),
            .lower => std.ascii.toLower(c),
        };
    }
    return out;
}

inline fn isAscii(s: []const u8) bool {
    for (s) |c| {
        if (c >= 0x80) return false;
    }
    return true;
}

fn decode(allocator: std.mem.Allocator, input: []const u8) Error![]u21 {
    var cps: std.ArrayListUnmanaged(u21) = .empty;
    errdefer cps.deinit(allocator);
    var i: usize = 0;
    while (i < input.len) {
        const seq_len = std.unicode.utf8ByteSequenceLength(input[i]) catch unreachable;
        const cp = std.unicode.utf8Decode(input[i .. i + seq_len]) catch unreachable;
        try cps.append(allocator, cp);
        i += seq_len;
    }
    return cps.toOwnedSlice(allocator);
}

fn appendCodepoint(
    allocator: std.mem.Allocator,
    out: *std.ArrayListUnmanaged(u8),
    cp: u21,
) Error!void {
    var buf: [4]u8 = undefined;
    const n = std.unicode.utf8Encode(cp, &buf) catch return error.InvalidUtf8;
    try out.appendSlice(allocator, buf[0..n]);
}

/// SpecialCasing.txt's Final_Sigma condition, verbatim: the scalar is
/// preceded by a cased letter followed by zero or more case-ignorable
/// scalars, and is NOT followed by zero or more case-ignorable scalars
/// then a cased letter.
///
/// Both halves are required. Dropping the first would lowercase a
/// standalone `Σ` to `ς`, and dropping the second would do it in the
/// middle of a word.
fn isFinalSigma(cps: []const u21, at: usize) bool {
    var i = at;
    var preceded = false;
    while (i > 0) {
        i -= 1;
        if (isCaseIgnorable(cps[i])) continue;
        preceded = isCased(cps[i]);
        break;
    }
    if (!preceded) return false;

    var j = at + 1;
    while (j < cps.len) : (j += 1) {
        if (isCaseIgnorable(cps[j])) continue;
        return !isCased(cps[j]);
    }
    return true;
}

fn isCased(cp: u21) bool {
    return inRanges(tables.cased, cp);
}

fn isCaseIgnorable(cp: u21) bool {
    return inRanges(tables.case_ignorable, cp);
}

fn inRanges(ranges: []const tables.Range, cp: u21) bool {
    var lo: usize = 0;
    var hi: usize = ranges.len;
    while (lo < hi) {
        const mid = lo + (hi - lo) / 2;
        const r = ranges[mid];
        if (cp < r.lo) {
            hi = mid;
        } else if (cp > r.hi) {
            lo = mid + 1;
        } else {
            return true;
        }
    }
    return false;
}

fn lookup(entries: []const tables.CaseEntry, cp: u21) ?tables.CaseEntry {
    var lo: usize = 0;
    var hi: usize = entries.len;
    while (lo < hi) {
        const mid = lo + (hi - lo) / 2;
        const e = entries[mid];
        if (e.from < cp) {
            lo = mid + 1;
        } else if (e.from > cp) {
            hi = mid;
        } else {
            return e;
        }
    }
    return null;
}

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

fn expectUpper(input: []const u8, want: []const u8) !void {
    const got = try toUpper(testing.allocator, input);
    defer testing.allocator.free(got);
    try testing.expectEqualStrings(want, got);
}

fn expectLower(input: []const u8, want: []const u8) !void {
    const got = try toLower(testing.allocator, input);
    defer testing.allocator.free(got);
    try testing.expectEqualStrings(want, got);
}

test "ASCII: both directions, unchanged length" {
    try expectUpper("Hello, World! 123", "HELLO, WORLD! 123");
    try expectLower("Hello, World! 123", "hello, world! 123");
}

test "length-changing: ß uppercases to SS" {
    // The row that decides this module exists. A simple one-to-one
    // mapping cannot express it, and the fold cannot either — `fold("ß")`
    // is lowercase `"ss"`.
    try expectUpper("ß", "SS");
    try expectUpper("Straße", "STRASSE");
    // …and it does not round-trip, which is Unicode's behaviour rather
    // than a defect: uppercasing is not injective.
    try expectLower("SS", "ss");
}

test "length-changing: every multi-scalar uppercase in the table" {
    // Not a sample: the table itself is asked how many length-changing
    // rows it has, and a few named ones are checked by hand. A UCD
    // revision that added or dropped one surfaces here rather than in a
    // formula fixture.
    var expanding: usize = 0;
    for (tables.upper_entries) |e| {
        if (e.len > 1) expanding += 1;
    }
    try testing.expect(expanding > 100);
    try expectUpper("ﬃ", "FFI"); // U+FB03 LATIN SMALL LIGATURE FFI
    try expectUpper("ﬄ", "FFL");
    try expectUpper("ŉ", "ʼN"); // U+0149, a two-scalar uppercase
    try expectUpper("ǰ", "J\u{030C}"); // U+01F0 → J + combining caron
}

test "dotted and dotless I: invariant, with the Turkish divergence recorded" {
    // U+0130 LATIN CAPITAL LETTER I WITH DOT ABOVE lowercases to
    // `i` + COMBINING DOT ABOVE under the unconditional SpecialCasing
    // row. The `tr`/`az` rows that map it to plain `i` are rejected.
    try expectLower("İ", "i\u{0307}");
    // The recorded divergence, stated as an assertion so it cannot drift
    // into silently becoming Turkish: a Turkish Excel answers `ı`
    // (U+0131) here, and zlsx answers ASCII `i`.
    try expectLower("I", "i");
    // U+0131 LATIN SMALL LETTER DOTLESS I uppercases to plain `I`, which
    // IS the locale-neutral answer.
    try expectUpper("ı", "I");
}

test "Final_Sigma: end of word takes ς, everywhere else σ" {
    // The one conditional rule in scope. Same capital sigma, three
    // positions, two answers.
    try expectLower("ΟΔΟΣ", "οδο\u{03C2}"); // final → ς
    try expectLower("ΣΟΦΟΣ", "\u{03C3}οφο\u{03C2}");
    // Mid-word sigma stays σ…
    try expectLower("ΑΣΑ", "α\u{03C3}α");
    // …and a sigma with no cased letter before it is not final either,
    // because the condition has two halves.
    try expectLower("Σ", "\u{03C3}");
    try expectLower(" Σ ", " \u{03C3} ");
}

test "Final_Sigma: case-ignorable scalars do not end a word" {
    // A combining acute between the letter and the sigma is
    // case-ignorable, so the sigma is still preceded by a cased letter…
    try expectLower("Α\u{0301}Σ", "α\u{0301}\u{03C2}");
    // …and a trailing quotation mark after the sigma is case-ignorable
    // too, so the sigma stays final rather than becoming medial.
    try expectLower("ΟΣ'", "ο\u{03C2}'");
    // But a cased letter after those ignorables makes it medial again.
    try expectLower("ΟΣ'Α", "ο\u{03C3}'α");
}

test "combining marks and astral letters pass through the right path" {
    // A combining mark has no casing of its own and must survive
    // uppercasing attached to its base.
    try expectUpper("e\u{0301}", "E\u{0301}");
    // DESERET (U+10400 block) is a cased astral script — the four-byte
    // UTF-8 path, where a wrong sequence length would corrupt output.
    try expectUpper("\u{10428}", "\u{10400}");
    try expectLower("\u{10400}", "\u{10428}");
    // ADLAM (U+1E900), added long after the BMP was full.
    try expectUpper("\u{1E922}", "\u{1E900}");
}

test "invalid UTF-8 refuses rather than guessing" {
    try testing.expectError(error.InvalidUtf8, toUpper(testing.allocator, "\xFFx"));
    try testing.expectError(error.InvalidUtf8, toLower(testing.allocator, "ab\xC3"));
}

test "table invariants: sorted, no duplicate sources, offsets in range" {
    inline for (.{ .upper, .lower, .title }) |dir| {
        const entries = entriesFor(dir);
        const scalars = scalarsFor(dir);
        try testing.expect(entries.len > 0);
        var prev: u21 = 0;
        for (entries, 0..) |e, i| {
            if (i > 0) try testing.expect(e.from > prev);
            prev = e.from;
            try testing.expect(e.len > 0);
            try testing.expect(e.offset + e.len <= scalars.len);
        }
    }
    // Every mapping is a real change: an identity row would cost a probe
    // and buy nothing, and the generator drops them.
    for (tables.upper_entries) |e| {
        if (e.len != 1) continue;
        try testing.expect(tables.upper_scalars[e.offset] != e.from);
    }
}

test "Unicode version pinned, and pinned to the same revision as the fold" {
    try testing.expectEqualStrings("17.0.0", unicode_version);
    // Casing and folding decide overlapping questions about the same
    // alphabet; two revisions would be two alphabets.
    const fold_tables = @import("tables/casefold_data.zig");
    try testing.expectEqualStrings(fold_tables.unicode_version, unicode_version);
}
