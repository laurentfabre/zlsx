//! §5.4d's shared text layer: the index unit, and everything that
//! counts in it.
//!
//! The whole module exists because of one 2024 change. Excel's
//! compatibility version decides what a *character* is to five
//! functions — `LEN`, `MID`, `FIND`, `SEARCH`, `REPLACE`:
//!
//!   * **CV1** counts **UTF-16 code units**, so an astral character
//!     (emoji, ancient scripts, most CJK extensions) counts as **two**.
//!     That is not a rule anyone designed; it is Excel's internal string
//!     representation showing through.
//!   * **CV2** counts **code points**, so the same character counts as
//!     **one**. Not grapheme clusters: a variation selector or a
//!     combining mark is still its own unit, so `LEN` of a flag emoji or
//!     a skin-toned emoji is still more than one.
//!
//! Both ship in v1 (§5.4d). CV2 has been the default for newly created
//! Current-Channel workbooks since April 2026, so refusing it would
//! refuse ordinary files; CV1 is what every workbook written before that
//! is, **and what a file with no compatibility metadata at all is** —
//! including every file zlsx's own Writer emits today.
//!
//! What is NOT compatibility-dependent, and is easy to assume is:
//! wildcards (`?` consumes one code point in both versions), criteria,
//! `COUNTIF`, `MATCH`, `XMATCH`, and the collation order itself. The
//! version changes counting, not matching.

const std = @import("std");
const assert = std.debug.assert;

const run_inputs = @import("run_inputs.zig");

/// §5.4d's version selector. Workbook-derived: `CalcState` carries it,
/// a caller cannot set it, and its absence means CV1.
pub const Cv = run_inputs.CompatibilityVersion;

pub const Error = error{
    /// A CV1 index landed between the two UTF-16 code units of one
    /// astral character.
    ///
    /// Excel answers with a lone surrogate — half a character, which its
    /// UTF-16 strings can hold and UTF-8 cannot. zlsx will not fabricate
    /// one: `MID("😀",1,1)` under CV1 is a typed refusal, recorded as an
    /// intentional divergence rather than answered with U+FFFD or by
    /// silently rounding to a whole character. Under CV2 the same call
    /// is the whole emoji and no refusal arises, which is one more
    /// reason the two versions are worth telling apart.
    SplitSurrogate,
};

/// How many index units `s` occupies.
pub fn unitLen(cv: Cv, s: []const u8) usize {
    var it = Units.init(cv, s);
    var n: usize = 0;
    while (it.next()) |step| n += step.units;
    return n;
}

/// The byte offset at which unit `n` begins, clamped to `s.len` when `n`
/// is past the end (`LEFT("ab",9)` is `"ab"`, not an error).
///
/// `error.SplitSurrogate` when `n` falls between the halves of an astral
/// character under CV1.
pub fn byteOfUnit(cv: Cv, s: []const u8, n: usize) Error!usize {
    if (n == 0) return 0;
    var it = Units.init(cv, s);
    var seen: usize = 0;
    while (it.next()) |step| {
        if (seen == n) return step.byte;
        if (seen < n and n < seen + step.units) return error.SplitSurrogate;
        seen += step.units;
    }
    return s.len;
}

/// The unit index of the character that begins at byte offset `b`.
/// `b` must be a scalar boundary — every caller gets one from the
/// iterator, from a fold's positional map, or from a byte search that
/// only ever matches whole scalars.
pub fn unitOfByte(cv: Cv, s: []const u8, b: usize) usize {
    assert(b <= s.len);
    var it = Units.init(cv, s);
    var units: usize = 0;
    while (it.next()) |step| {
        if (step.byte >= b) break;
        units += step.units;
    }
    return units;
}

/// The unit index of code point number `cp_index` — the conversion
/// `SEARCH` needs, because the fold's positional map is keyed by code
/// point and the answer is reported in index units.
pub fn unitOfCodePoint(cv: Cv, s: []const u8, cp_index: usize) usize {
    var it = Units.init(cv, s);
    var units: usize = 0;
    var seen: usize = 0;
    while (it.next()) |step| {
        if (seen == cp_index) return units;
        units += step.units;
        seen += 1;
    }
    return units;
}

/// The code-point index at unit `start` — `unitOfCodePoint` inverted,
/// and the direction `SEARCH` needs to convert its caller's start
/// position into something the fold's positional map understands.
///
/// A start inside an astral character under CV1 rounds **down** to that
/// character rather than refusing: a search that begins mid-character
/// still has a well-defined first place to look, and Excel's own
/// `SEARCH("a","😀a",2)` finds the `a`. Only a *result* that would be
/// half a character is unrepresentable.
pub fn codePointOfUnit(cv: Cv, s: []const u8, start: usize) usize {
    var it = Units.init(cv, s);
    var units: usize = 0;
    var cps: usize = 0;
    while (it.next()) |step| {
        if (units + step.units > start) return cps;
        units += step.units;
        cps += 1;
    }
    return cps;
}

/// `count` units starting at unit `start`, clamped at both ends. The
/// slice borrows from `s`.
pub fn sliceUnits(cv: Cv, s: []const u8, start: usize, count: usize) Error![]const u8 {
    const from = try byteOfUnit(cv, s, start);
    const to = try byteOfUnit(cv, s, start +| count);
    return s[from..to];
}

/// One step of the walk: where the character starts, and how many index
/// units it is worth.
const Step = struct { byte: usize, len: usize, units: usize };

/// Walks `s` a character at a time, reporting each one's weight under
/// `cv`.
///
/// Invalid UTF-8 is walked byte by byte rather than refused, matching
/// `criteria.fold`: text reaching the evaluator has already been through
/// the decode boundary, so this is reachable only from a hand-built
/// value, and a total answer beats a refusal nobody can act on. The two
/// modules agree deliberately — one of them counting differently would
/// make `SEARCH` and `LEN` disagree about the same string.
const Units = struct {
    cv: Cv,
    s: []const u8,
    i: usize,

    fn init(cv: Cv, s: []const u8) Units {
        return .{ .cv = cv, .s = s, .i = 0 };
    }

    fn next(self: *Units) ?Step {
        if (self.i >= self.s.len) return null;
        const at = self.i;
        const seq_len = std.unicode.utf8ByteSequenceLength(self.s[at]) catch {
            self.i += 1;
            return .{ .byte = at, .len = 1, .units = 1 };
        };
        if (at + seq_len > self.s.len) {
            self.i += 1;
            return .{ .byte = at, .len = 1, .units = 1 };
        }
        const cp = std.unicode.utf8Decode(self.s[at .. at + seq_len]) catch {
            self.i += 1;
            return .{ .byte = at, .len = 1, .units = 1 };
        };
        self.i += seq_len;
        return .{ .byte = at, .len = seq_len, .units = unitsOf(self.cv, cp) };
    }
};

fn unitsOf(cv: Cv, cp: u21) usize {
    return switch (cv) {
        // The surrogate-pair rule, and the only place the two versions
        // differ: U+10000 and above take two UTF-16 code units.
        .cv1 => if (cp > 0xFFFF) 2 else 1,
        .cv2 => 1,
    };
}

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

// U+1F600 GRINNING FACE — one code point, two UTF-16 code units, four
// UTF-8 bytes. The character the two versions disagree about.
const emoji = "\u{1F600}";

test "unitLen: ASCII and BMP agree in both versions" {
    for ([_]Cv{ .cv1, .cv2 }) |cv| {
        try testing.expectEqual(@as(usize, 0), unitLen(cv, ""));
        try testing.expectEqual(@as(usize, 5), unitLen(cv, "hello"));
        // `é` is two UTF-8 bytes and one unit either way.
        try testing.expectEqual(@as(usize, 4), unitLen(cv, "café"));
        // Greek, Cyrillic, CJK: all BMP, all one unit per character.
        try testing.expectEqual(@as(usize, 3), unitLen(cv, "日本語"));
    }
}

test "unitLen: astral characters are where the versions part" {
    try testing.expectEqual(@as(usize, 2), unitLen(.cv1, emoji));
    try testing.expectEqual(@as(usize, 1), unitLen(.cv2, emoji));
    try testing.expectEqual(@as(usize, 4), unitLen(.cv1, "a" ++ emoji ++ "b"));
    try testing.expectEqual(@as(usize, 3), unitLen(.cv2, "a" ++ emoji ++ "b"));
}

test "unitLen: combining marks and variation selectors are separate in both" {
    // NOT grapheme clustering (§5.4d). `e` + combining acute is two
    // units under either version, however it renders.
    for ([_]Cv{ .cv1, .cv2 }) |cv| {
        try testing.expectEqual(@as(usize, 2), unitLen(cv, "e\u{0301}"));
    }
    // A variation selector is BMP, so it is one unit in both — but it is
    // still a unit, which is the half of this rule people expect to be
    // absorbed and is not.
    try testing.expectEqual(@as(usize, 2), unitLen(.cv2, "\u{2764}\u{FE0F}"));
}

test "sliceUnits: whole characters, clamped at both ends" {
    for ([_]Cv{ .cv1, .cv2 }) |cv| {
        try testing.expectEqualStrings("he", try sliceUnits(cv, "hello", 0, 2));
        try testing.expectEqualStrings("llo", try sliceUnits(cv, "hello", 2, 3));
        // Past the end clamps rather than refusing — `MID("ab",1,99)`.
        try testing.expectEqualStrings("llo", try sliceUnits(cv, "hello", 2, 99));
        try testing.expectEqualStrings("", try sliceUnits(cv, "hello", 9, 2));
        try testing.expectEqualStrings("", try sliceUnits(cv, "hello", 1, 0));
    }
}

test "sliceUnits: the astral character, taken and split" {
    const s = "a" ++ emoji ++ "b";
    // CV2: one unit, so unit 1 is the whole emoji.
    try testing.expectEqualStrings(emoji, try sliceUnits(.cv2, s, 1, 1));
    try testing.expectEqualStrings("a", try sliceUnits(.cv2, s, 0, 1));
    try testing.expectEqualStrings("b", try sliceUnits(.cv2, s, 2, 1));

    // CV1: two units, so taking both is the emoji…
    try testing.expectEqualStrings(emoji, try sliceUnits(.cv1, s, 1, 2));
    try testing.expectEqualStrings("b", try sliceUnits(.cv1, s, 3, 1));
    // …and taking one is half a surrogate pair, which UTF-8 cannot
    // hold. Excel hands back a lone surrogate; this refuses.
    try testing.expectError(error.SplitSurrogate, sliceUnits(.cv1, s, 1, 1));
    try testing.expectError(error.SplitSurrogate, sliceUnits(.cv1, s, 2, 1));
}

test "byteOfUnit: boundaries, the end, and past the end" {
    const s = "a" ++ emoji ++ "b"; // bytes: 0 | 1..4 | 5
    try testing.expectEqual(@as(usize, 0), try byteOfUnit(.cv1, s, 0));
    try testing.expectEqual(@as(usize, 1), try byteOfUnit(.cv1, s, 1));
    try testing.expectEqual(@as(usize, 5), try byteOfUnit(.cv1, s, 3));
    try testing.expectEqual(@as(usize, 6), try byteOfUnit(.cv1, s, 4));
    // Past the end clamps to the length, so a caller's `start + count`
    // needs no bound of its own.
    try testing.expectEqual(@as(usize, 6), try byteOfUnit(.cv1, s, 99));
    try testing.expectEqual(@as(usize, 5), try byteOfUnit(.cv2, s, 2));
    try testing.expectError(error.SplitSurrogate, byteOfUnit(.cv1, s, 2));
}

test "unitOfByte and unitOfCodePoint: the two conversions callers need" {
    const s = "a" ++ emoji ++ "b";
    // FIND matches at a byte offset and reports units.
    try testing.expectEqual(@as(usize, 0), unitOfByte(.cv1, s, 0));
    try testing.expectEqual(@as(usize, 1), unitOfByte(.cv1, s, 1));
    try testing.expectEqual(@as(usize, 3), unitOfByte(.cv1, s, 5));
    try testing.expectEqual(@as(usize, 2), unitOfByte(.cv2, s, 5));
    // SEARCH matches at a code-point index and reports units.
    try testing.expectEqual(@as(usize, 1), unitOfCodePoint(.cv1, s, 1));
    try testing.expectEqual(@as(usize, 3), unitOfCodePoint(.cv1, s, 2));
    try testing.expectEqual(@as(usize, 2), unitOfCodePoint(.cv2, s, 2));
    // Past the end is the length in units, which is what a match at the
    // very end of the string converts to.
    try testing.expectEqual(@as(usize, 4), unitOfCodePoint(.cv1, s, 9));
}

test "codePointOfUnit: inverts unitOfCodePoint, and rounds a split start down" {
    const s = "a" ++ emoji ++ "b";
    try testing.expectEqual(@as(usize, 0), codePointOfUnit(.cv1, s, 0));
    try testing.expectEqual(@as(usize, 1), codePointOfUnit(.cv1, s, 1));
    // Unit 2 is the emoji's second half under CV1: a search may start
    // there, and it starts AT the emoji.
    try testing.expectEqual(@as(usize, 1), codePointOfUnit(.cv1, s, 2));
    try testing.expectEqual(@as(usize, 2), codePointOfUnit(.cv1, s, 3));
    try testing.expectEqual(@as(usize, 3), codePointOfUnit(.cv1, s, 99));
    // CV2 has no split to round: one unit, one code point, all the way.
    for (0..4) |u| {
        try testing.expectEqual(@min(u, @as(usize, 3)), codePointOfUnit(.cv2, s, u));
    }
    // Round-trip on every boundary that is one.
    for ([_]Cv{ .cv1, .cv2 }) |cv| {
        for (0..3) |cp| {
            try testing.expectEqual(cp, codePointOfUnit(cv, s, unitOfCodePoint(cv, s, cp)));
        }
    }
}

test "invalid UTF-8 stays total rather than refusing" {
    // The same policy `criteria.fold` takes, and for the same reason:
    // text that reached here bypassed the decode boundary, and a
    // refusal at this depth is not something a caller can act on.
    const bad = "a\xFFb";
    try testing.expectEqual(@as(usize, 3), unitLen(.cv1, bad));
    try testing.expectEqual(@as(usize, 3), unitLen(.cv2, bad));
    try testing.expectEqualStrings("\xFF", try sliceUnits(.cv1, bad, 1, 1));
    // A truncated sequence is walked byte-wise too, rather than reading
    // past the end of the slice.
    try testing.expectEqual(@as(usize, 2), unitLen(.cv2, "a\xC3"));
}

test "the absent-metadata default is CV1" {
    // §5.4d: a workbook with no compatibility metadata — which is every
    // pre-2024 file, and every file zlsx's Writer emits — is CV1. The
    // default is asserted here rather than trusted, because it decides
    // what `LEN` answers for a file nobody configured.
    const state: run_inputs.CalcState = .{};
    try testing.expectEqual(Cv.cv1, state.text_compat);
    try testing.expectEqual(@as(usize, 2), unitLen(state.text_compat, emoji));
}
