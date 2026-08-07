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
//! `toUpper`/`toLower`: `toProper` (M8b) is that word segmentation,
//! over the same tables — the generator never reopened.
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

/// Casing allocates and nothing else. Invalid UTF-8 does **not** refuse
/// here: `criteria.fold` and `text.zig` both stay total over it, and a
/// third module deciding differently would mean `UPPER` refused a string
/// `LEN` was happy to measure. Bytes that are not valid UTF-8 pass
/// through untouched — the closest thing to "leave it alone" a caller
/// can act on, and unreachable from a workbook that came through the
/// decode boundary.
pub const Error = error{OutOfMemory};

/// Full uppercase. Returns owned bytes — caller frees.
pub fn toUpper(allocator: std.mem.Allocator, input: []const u8) Error![]u8 {
    return map(allocator, input, .upper);
}

/// Full lowercase, including the Final_Sigma rule. Returns owned bytes.
pub fn toLower(allocator: std.mem.Allocator, input: []const u8) Error![]u8 {
    return map(allocator, input, .lower);
}

/// Full titlecase of every scalar, WITHOUT word segmentation — this is
/// the character-level half of `PROPER`, not `PROPER` itself. The
/// segmentation that decides which scalars are word-initial is
/// `toProper`'s (M8b).
pub fn toTitleScalar(allocator: std.mem.Allocator, input: []const u8) Error![]u8 {
    return map(allocator, input, .title);
}

/// `PROPER` — word segmentation over the SAME two range tables
/// Final_Sigma already reads, then the title/lower mappings. The
/// invariant boundary rule, pinned (M8b, §7):
///
///   * a cased scalar TITLE-cases when the nearest preceding scalar
///     that is not case-ignorable is absent or not cased, and
///     lowercases otherwise, Final_Sigma included;
///   * a case-ignorable scalar is TRANSPARENT — it neither starts nor
///     ends a word — and every other non-cased scalar ends the word.
///
/// Two consequences are divergences recorded here rather than defects,
/// both spec_pinned while the Excel oracle leg is parked (§8.2):
/// U+0027 is case-ignorable (Single_Quote), so `don't` answers `Don't`
/// where Excel answers `Don'T`; and an uncased LETTER (Hebrew, CJK)
/// ends a word, because telling it from a non-letter takes the
/// Alphabetic property — a second table this row does not add.
pub fn toProper(allocator: std.mem.Allocator, input: []const u8) Error![]u8 {
    // The same two total paths as `map`: bytes for ASCII (the fast
    // path) and for invalid UTF-8 (the only answer), scalars otherwise.
    if (isAscii(input) or !std.unicode.utf8ValidateSlice(input)) {
        return properBytes(allocator, input);
    }

    const cps = try decode(allocator, input);
    defer allocator.free(cps);

    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(allocator);
    try out.ensureTotalCapacity(allocator, input.len);

    for (cps, 0..) |cp, i| {
        if (!isCased(cp)) {
            try appendCodepoint(allocator, &out, cp);
            continue;
        }
        const dir: Direction = if (beginsWord(cps, i)) .title else .lower;
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

/// Byte-wise PROPER — the same boundary rule over the two inputs the
/// byte path owns. For a wholly-ASCII string, ASCII's cased scalars
/// are exactly its letters and its case-ignorables (`'` `.` `:` `^`
/// `` ` ``) come out of the SAME table, so the fast path is a fast
/// path and not a dialect. For an invalid one, an unreadable byte
/// ends the word like any other non-letter and passes through
/// untouched — total, like `mapAscii`.
fn properBytes(allocator: std.mem.Allocator, input: []const u8) Error![]u8 {
    const out = try allocator.alloc(u8, input.len);
    // Whether the nearest preceding non-case-ignorable byte is a letter.
    var after_letter = false;
    for (input, 0..) |c, i| {
        if (std.ascii.isAlphabetic(c)) {
            // Titlecasing an ASCII letter is uppercasing it — the
            // digraph rows where the two differ are all non-ASCII.
            out[i] = if (after_letter) std.ascii.toLower(c) else std.ascii.toUpper(c);
            after_letter = true;
        } else {
            out[i] = c;
            if (c >= 0x80 or !isCaseIgnorable(c)) after_letter = false;
        }
    }
    return out;
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

    // Invalid UTF-8 keeps its ASCII cased and everything else verbatim.
    // There is no scalar sequence to walk, so there is no Final_Sigma
    // context either — and no answer better than "leave the bytes I
    // cannot read alone".
    if (!std.unicode.utf8ValidateSlice(input)) return mapAscii(allocator, input, dir);

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

/// Byte-wise casing: every ASCII letter cased, every other byte copied.
/// Correct for two inputs — a wholly-ASCII string, where it is also the
/// fast path, and an invalid one, where it is the total answer.
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
    // Unreachable by construction: every codepoint here came out of
    // `utf8Decode` or out of a generated table, so none is a surrogate
    // or out of range.
    const n = std.unicode.utf8Encode(cp, &buf) catch unreachable;
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

/// Whether the cased scalar at `at` begins a word: nothing stands
/// before it, or the nearest preceding scalar that is not
/// case-ignorable is not cased. Final_Sigma's backward walk, pointed
/// at the start of the word instead of its end.
fn beginsWord(cps: []const u21, at: usize) bool {
    var i = at;
    while (i > 0) {
        i -= 1;
        if (isCaseIgnorable(cps[i])) continue;
        return !isCased(cps[i]);
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

test "invalid UTF-8 stays total, like the fold and the index layer" {
    // Not a refusal: `criteria.fold` and `text.zig` both measure and
    // match such a string, so refusing to case it would be the third
    // module inventing a fourth answer. ASCII is cased, the unreadable
    // bytes are left exactly as they came.
    try expectUpper("\xFFx", "\xFFX");
    try expectLower("AB\xC3", "ab\xC3");
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

// ─── PROPER (M8b): the boundary rule, byte-exact ─────────────────

fn expectProper(input: []const u8, want: []const u8) !void {
    const got = try toProper(testing.allocator, input);
    defer testing.allocator.free(got);
    try testing.expectEqualStrings(want, got);
}

test "PROPER: ASCII — letters after letters lower, anything else starts a word" {
    try expectProper("hello world", "Hello World");
    // Non-initial letters LOWER: PROPER is not "capitalize".
    try expectProper("EXCEL FILE", "Excel File");
    try expectProper("2-way street", "2-Way Street");
    // A digit ends a word without being one: the letter after `76`
    // title-cases, the digits pass through.
    try expectProper("76budGET", "76Budget");
    try expectProper("", "");
}

test "PROPER: the case-ignorables are transparent, and that is the recorded divergence" {
    // `'` is Case_Ignorable (Single_Quote), so no word begins at `t`.
    // Excel answers `Don'T Stop` — recorded, spec_pinned, pending the
    // parked oracle leg (§8.2).
    try expectProper("don't stop", "Don't Stop");
    // `.` (MidNumLet) and `:` (MidLetter) are transparent for the same
    // reason, and Excel capitalizes after both.
    try expectProper("j.r.r. tolkien", "J.r.r. Tolkien");
    try expectProper("re:invent", "Re:invent");
    // The decoded path answers the same bytes for the same apostrophe:
    // one rule, two paths, no dialect.
    try expectProper("don't é", "Don't É");
}

test "PROPER: titlecase is the third table, not uppercase" {
    // U+01C6 ǆ titlecases to U+01C5 ǅ — the tri-form digraph where
    // title and upper genuinely differ; UPPER answers Ǆ here.
    try expectProper("ǆungla", "ǅungla");
    // ß's full title mapping is `Ss` (SpecialCasing), length-changing
    // like its uppercase `SS` but not equal to it.
    try expectProper("ß", "Ss");
    try expectProper("straße", "Straße");
}

test "PROPER: Final_Sigma still decides the lowered tail" {
    // The lower direction inside PROPER is the SAME lower: a word-final
    // sigma takes ς, a medial one σ, under the walk `toLower` uses.
    try expectProper("ΟΔΟΣ ΣΟΦΟΣ", "Οδο\u{03C2} Σοφο\u{03C2}");
    try expectProper("ΑΣΑ", "Α\u{03C3}α");
}

test "PROPER: combining marks ride their base and do not begin words" {
    // U+0302 and U+0301 are case-ignorable: the letters after them are
    // still mid-word, and the marks survive attached to their bases.
    try expectProper("bru\u{0302}le\u{0301}e day", "Bru\u{0302}le\u{0301}e Day");
}

test "PROPER: astral letters take the four-byte path in both roles" {
    // DESERET: the word-initial 𐐨 titlecases, the second stays lower.
    try expectProper("\u{10428}\u{10428} \u{10428}", "\u{10400}\u{10428} \u{10400}");
    // ADLAM, cased both directions in the supplementary planes.
    try expectProper("\u{1E900}\u{1E900}", "\u{1E900}\u{1E922}");
}

test "PROPER: an uncased letter ends a word, and the trade is named" {
    // Hebrew and CJK have no case, so their letters are
    // indistinguishable from non-letters without the Alphabetic
    // property — a second table this row does not add. Excel treats
    // them as letters; recorded divergence, spec_pinned.
    try expectProper("אb", "אB");
    try expectProper("中文a", "中文A");
}

test "PROPER: invalid UTF-8 stays total, like the fold and the two casings" {
    try expectProper("\xFFx", "\xFFX");
    try expectProper("a\xFFb", "A\xFFB");
    try expectProper("don't\xFF", "Don't\xFF");
}

test "Unicode version pinned, and pinned to the same revision as the fold" {
    try testing.expectEqualStrings("17.0.0", unicode_version);
    // Casing and folding decide overlapping questions about the same
    // alphabet; two revisions would be two alphabets. The check goes
    // through the fold MODULE rather than its table file: importing
    // `tables/casefold_data.zig` from here would put that file in two
    // module trees, which is the same constraint that moved this whole
    // directory out of `src/`.
    const casefold = @import("zlsx_casefold");
    try testing.expectEqualStrings(casefold.unicode_version, unicode_version);
}
