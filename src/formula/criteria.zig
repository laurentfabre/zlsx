//! Criteria — the `criteria operand` column of §5.3b as executable
//! rules, plus the wildcard matcher §5.4b specifies
//! (`goal_formula.md` §5.3b, §5.4b, §5.6a).
//!
//! M3b of the tier-D1 ladder. `COUNTIF`, `SUMIF`, and the whole `*IFS`
//! family at M7b2 are thin wrappers over what is here; none of them
//! re-derives a matching rule and none of them re-derives an alignment.
//!
//! A criterion is not a comparison
//! -------------------------------
//! `">5"` looks like a comparison and is not one. Excel restricts a
//! criterion to cells **of the operand's own type**: `">5"` never
//! matches a text cell even though §5.3b's total preorder puts every
//! text above every number. Applying the preorder here would make
//! `COUNTIF(A:A,">5")` count words, which is both wrong and the kind of
//! wrong that looks right in a small fixture. The type restriction is
//! the rule, and it is spec-pinned — no committed manifest contains a
//! criteria row.
//!
//! Wildcards over an expanding fold
//! --------------------------------
//! §5.4b is precise about this and it is the subtle part. Matching is
//! case-insensitive under `collation_v1`, so `ß` and `ss` are the same
//! text — but `?` consumes **one code point of the original**, not one
//! unit of the fold. So `"?"` matches `ß` (one code point) while `"ss"`
//! *also* matches `ß` (two folded units), and `"?s"` matches neither.
//!
//! Getting that right needs a folded-unit → original-code-point map,
//! which `Folded` builds by folding one code point at a time. Full case
//! folding is defined per code point — there is no context-dependent
//! multi-character fold in the non-Turkic full fold — so folding
//! code-point-wise and folding the whole string agree, and only the
//! former yields the map. Literal runs additionally have to land on a
//! code-point boundary, which is what stops `"s"` from matching half of
//! a folded `ß`.
//!
//! Allocation
//! ----------
//! Matching one cell against one criterion costs one `Folded` per side.
//! Callers scanning a range fold the *criterion* once, outside the loop
//! (`Criterion.prepare`), and only the cell inside it.

const std = @import("std");
const assert = std.debug.assert;

const value = @import("value.zig");
const env = @import("env.zig");

pub const Error = error{
    OutOfMemory,
    /// Text that parses only under some locale reached a criterion
    /// operand — a plane-2 refusal, never a guessed number (§5.4b).
    LocaleSensitiveInput,
    /// The fold failed. `collation_v1`'s fold cannot fail on valid
    /// UTF-8; this covers an injected one that can.
    BadFold,
};

/// The comparison a criterion applies to the cells it accepts.
pub const Relation = enum {
    /// No operator prefix, or an explicit `=`. Wildcards are active.
    eq,
    /// `<>`. Wildcards are active, and the sense is inverted.
    ne,
    lt,
    le,
    gt,
    ge,

    /// Only equality and inequality honour `*`, `?`, and `~` — Excel
    /// treats them as literal characters under an ordering operator.
    pub fn wildcardsActive(self: Relation) bool {
        return self == .eq or self == .ne;
    }
};

/// What the criterion compares *against*, after the operator prefix is
/// stripped and the remainder classified.
pub const Operand = union(enum) {
    number: f64,
    text: []const u8,
    boolean: bool,
    err: value.ErrorValue,
    /// An empty criterion (`""`). Matches true blanks **and** cells
    /// holding the empty string — the `.countblank_class` population,
    /// which is why the two classes exist (§5.6a).
    empty,
};

pub const Criterion = struct {
    relation: Relation,
    operand: Operand,
    /// Whether the operand text needs the pattern matcher rather than a
    /// plain collation compare. True for an unescaped `*`/`?` **and** for
    /// a `~` on its own: `"~*"` has no active wildcard but still has an
    /// escape to strip before anything can match it literally.
    is_pattern: bool = false,
    /// The criterion's raw spelling, kept for diagnostics.
    source: []const u8 = "",
};

// ─── parsing ─────────────────────────────────────────────────────

/// Build a criterion from a scalar. A criterion arrives as a *value*,
/// not as text: `COUNTIF(A:A, B1)` passes whatever B1 holds, and only a
/// text criterion carries an operator prefix or wildcards.
pub fn parse(v: value.ScalarValue, fidelity: value.Fidelity) Error!Criterion {
    return switch (v) {
        .number => |n| .{ .relation = .eq, .operand = .{ .number = n } },
        .boolean => |b| .{ .relation = .eq, .operand = .{ .boolean = b } },
        .err => |e| .{ .relation = .eq, .operand = .{ .err = e } },
        // A blank criterion argument is Excel's "match blanks".
        .blank => .{ .relation = .eq, .operand = .empty },
        .text => |t| parseText(t, fidelity),
    };
}

fn parseText(text: []const u8, fidelity: value.Fidelity) Error!Criterion {
    // Longest operator first: `<=` and `<>` both start with `<`.
    const ops = [_]struct { spelling: []const u8, rel: Relation }{
        .{ .spelling = "<=", .rel = .le },
        .{ .spelling = ">=", .rel = .ge },
        .{ .spelling = "<>", .rel = .ne },
        .{ .spelling = "<", .rel = .lt },
        .{ .spelling = ">", .rel = .gt },
        .{ .spelling = "=", .rel = .eq },
    };
    var relation: Relation = .eq;
    var rest = text;
    for (ops) |op| {
        if (std.mem.startsWith(u8, text, op.spelling)) {
            relation = op.rel;
            rest = text[op.spelling.len..];
            break;
        }
    }

    if (rest.len == 0) {
        // `""` matches blanks; `"<>"` matches everything that is not
        // blank. Both are Excel's, and both fall out of one rule.
        return .{ .relation = relation, .operand = .empty, .source = text };
    }

    // An error spelling is an error operand: `COUNTIF(A:A,"#N/A")`
    // counts error cells, which §5.3c calls out for the COUNTIF family.
    if (value.KnownError.fromSpelling(rest)) |k| {
        return .{ .relation = relation, .operand = .{ .err = .{ .known = k } }, .source = text };
    }

    // Numeric text becomes a number, through the `.criteria` ingress —
    // §5.4's caller→ingress table names it, and it is the reason
    // `">1,5"` refuses rather than guessing.
    if (value.parseDecimal(fidelity, .criteria, rest)) |n| {
        return .{ .relation = relation, .operand = .{ .number = n }, .source = text };
    } else |e| switch (e) {
        error.LocaleSensitive => return error.LocaleSensitiveInput,
        error.NotNumeric => {},
    }

    if (asBooleanSpelling(rest)) |b| {
        return .{ .relation = relation, .operand = .{ .boolean = b }, .source = text };
    }

    return .{
        .relation = relation,
        .operand = .{ .text = rest },
        .is_pattern = relation.wildcardsActive() and
            (hasWildcards(rest) or std.mem.indexOfScalar(u8, rest, '~') != null),
        .source = text,
    };
}

fn asBooleanSpelling(s: []const u8) ?bool {
    if (std.ascii.eqlIgnoreCase(s, "TRUE")) return true;
    if (std.ascii.eqlIgnoreCase(s, "FALSE")) return false;
    return null;
}

/// Whether an unescaped `*` or `?` appears. `~` escapes the next
/// character, including another `~`.
pub fn hasWildcards(pattern: []const u8) bool {
    var i: usize = 0;
    while (i < pattern.len) : (i += 1) {
        switch (pattern[i]) {
            '~' => i += 1,
            '*', '?' => return true,
            else => {},
        }
    }
    return false;
}

// ─── folding with a positional map ───────────────────────────────

/// A string's full case fold, plus the map from folded byte offsets back
/// to original code-point boundaries. The map is what makes `?` mean
/// "one code point" even where the fold expanded one into several
/// (§5.4b).
pub const Folded = struct {
    /// The folded bytes.
    bytes: []const u8,
    /// `starts[i]` is the folded offset at which original code point `i`
    /// begins; `starts[len]` is `bytes.len`. Length is
    /// `code_points + 1`.
    starts: []const u32,

    pub fn codePoints(self: Folded) usize {
        return self.starts.len - 1;
    }

    /// The original code-point index whose fold begins at `offset`, or
    /// null when `offset` is inside an expansion. A literal run that
    /// ends mid-expansion has not matched a whole character.
    pub fn boundaryAt(self: Folded, offset: u32) ?usize {
        // `starts` is ascending, so this is a binary search rather than
        // a scan — criteria over a whole column call it per cell.
        var lo: usize = 0;
        var hi: usize = self.starts.len;
        while (lo < hi) {
            const mid = lo + (hi - lo) / 2;
            if (self.starts[mid] < offset) lo = mid + 1 else hi = mid;
        }
        if (lo < self.starts.len and self.starts[lo] == offset) return lo;
        return null;
    }
};

/// Fold `s` one code point at a time, recording where each landed.
///
/// Per-code-point folding is not an approximation of whole-string
/// folding: the non-Turkic full case fold is defined per code point, so
/// the two agree byte for byte. Only this form yields the positional
/// map, and §5.4b needs the map.
pub fn fold(
    allocator: std.mem.Allocator,
    fold_fn: value.FoldFn,
    s: []const u8,
) Error!Folded {
    var bytes: std.ArrayListUnmanaged(u8) = .empty;
    errdefer bytes.deinit(allocator);
    var starts: std.ArrayListUnmanaged(u32) = .empty;
    errdefer starts.deinit(allocator);

    var view = std.unicode.Utf8View.init(s) catch {
        // Invalid UTF-8 cannot be folded; treat each byte as its own
        // unit so matching stays total rather than refusing. The
        // tokenizer already rejects invalid UTF-8 upstream (M1a), so
        // this is reachable only through a hand-built value.
        for (s, 0..) |b, i| {
            try starts.append(allocator, @intCast(i));
            try bytes.append(allocator, b);
        }
        try starts.append(allocator, @intCast(s.len));
        const raw = try bytes.toOwnedSlice(allocator);
        errdefer allocator.free(raw);
        return .{ .bytes = raw, .starts = try starts.toOwnedSlice(allocator) };
    };

    var it = view.iterator();
    while (it.nextCodepointSlice()) |cp| {
        try starts.append(allocator, @intCast(bytes.items.len));
        const folded = fold_fn(allocator, cp) catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            else => return error.BadFold,
        };
        defer allocator.free(folded);
        try bytes.appendSlice(allocator, folded);
    }
    try starts.append(allocator, @intCast(bytes.items.len));
    const folded_bytes = try bytes.toOwnedSlice(allocator);
    errdefer allocator.free(folded_bytes);
    return .{ .bytes = folded_bytes, .starts = try starts.toOwnedSlice(allocator) };
}

// ─── wildcard matching ───────────────────────────────────────────

/// Match a folded target against a folded pattern.
///
/// Backtracking rather than a DFA because the patterns are short, the
/// alphabet is Unicode, and a wrong DFA is much harder to see than a
/// wrong backtracker. `*` uses the standard greedy-with-restart trick,
/// which keeps it linear in the common case and bounded in every case.
///
/// `pattern_escapes[i]` marks the pattern code points that `~` escaped,
/// so a literal `*` can be matched.
fn matchFolded(
    pattern: Folded,
    escapes: []const bool,
    target: Folded,
) bool {
    return matchFoldedFrom(pattern, escapes, target, 0, .whole);
}

/// Whether the pattern has to consume the target to the end.
///
/// A criterion matches a whole cell; `SEARCH` finds a substring. That is
/// the only difference between them, so it is a parameter rather than a
/// second matcher — a second one would be a second answer to "does `~*`
/// match a literal star", and §5.4b has one.
const MatchExtent = enum { whole, prefix };

fn matchFoldedFrom(
    pattern: Folded,
    escapes: []const bool,
    target: Folded,
    from: usize,
    extent: MatchExtent,
) bool {
    var p: usize = 0; // pattern code point
    var t: usize = from; // target code point
    var star_p: ?usize = null;
    var star_t: usize = from;

    while (t < target.codePoints() or p < pattern.codePoints()) {
        // A prefix match is done the moment the pattern runs out,
        // however much target is left. `matchRun` still had to land on a
        // code-point boundary to get here, so this cannot report a match
        // that ended halfway through an expansion.
        if (extent == .prefix and p >= pattern.codePoints()) return true;
        if (p < pattern.codePoints()) {
            const unit = patternUnit(pattern, p);
            const escaped = escapes[p];
            if (!escaped and unit.len == 1 and unit[0] == '*') {
                star_p = p;
                p += 1;
                star_t = t;
                continue;
            }
            if (!escaped and unit.len == 1 and unit[0] == '?') {
                if (t < target.codePoints()) {
                    // One code point of the ORIGINAL, however many
                    // folded units it turned into.
                    p += 1;
                    t += 1;
                    continue;
                }
            } else if (t < target.codePoints()) {
                if (std.mem.eql(u8, unit, targetUnit(target, t))) {
                    p += 1;
                    t += 1;
                    continue;
                }
                // A folded unit may span several target code points
                // (`ss` against `ß`) or several pattern ones. Try to
                // consume a run on either side that folds identically.
                if (matchRun(pattern, p, target, t)) |adv| {
                    p += adv.pattern;
                    t += adv.target;
                    continue;
                }
            }
        }
        // Mismatch: fall back to the last `*` and let it eat one more.
        if (star_p) |sp| {
            star_t += 1;
            if (star_t > target.codePoints()) return false;
            p = sp + 1;
            t = star_t;
            continue;
        }
        return false;
    }
    return true;
}

/// §5.4b's `SEARCH`: the first ORIGINAL code-point index at or after
/// `from` where `pattern` matches a prefix of `target`. Null when there
/// is none.
///
/// The index is the caller's, not the fold's — `Folded.starts` is keyed
/// by original code point, which is the whole reason the positional map
/// exists. Converting that to §5.4d's index unit is `text.zig`'s job,
/// because it is the compatibility version that decides whether an
/// astral character counts once or twice.
///
/// Both arguments are already folded, so this is case-insensitive by
/// construction. `FIND` does not come through here: it is `.raw`, and a
/// raw search over folded strings would quietly become case-insensitive.
pub fn searchFolded(
    pattern: Folded,
    escapes: []const bool,
    target: Folded,
    from: usize,
) ?usize {
    // An empty pattern matches at the start position, which is Excel's
    // answer for `SEARCH("",text)` — 1 — rather than "not found".
    var start = from;
    while (start <= target.codePoints()) : (start += 1) {
        if (matchFoldedFrom(pattern, escapes, target, start, .prefix)) return start;
    }
    return null;
}

fn patternUnit(f: Folded, i: usize) []const u8 {
    return f.bytes[f.starts[i]..f.starts[i + 1]];
}

fn targetUnit(f: Folded, i: usize) []const u8 {
    return f.bytes[f.starts[i]..f.starts[i + 1]];
}

/// Match the longest run of pattern code points against the longest run
/// of target code points whose folds are byte-identical. This is what
/// lets `"ss"` match `"ß"` while `"?s"` does not: both sides must end on
/// a code-point boundary.
fn matchRun(
    pattern: Folded,
    p0: usize,
    target: Folded,
    t0: usize,
) ?struct { pattern: usize, target: usize } {
    const p_start = pattern.starts[p0];
    const t_start = target.starts[t0];
    var pi = p0 + 1;
    while (pi <= pattern.codePoints()) : (pi += 1) {
        const p_end = pattern.starts[pi];
        const len = p_end - p_start;
        if (len == 0) continue;
        if (t_start + len > target.bytes.len) break;
        const t_end = t_start + len;
        const boundary = target.boundaryAt(t_end) orelse continue;
        if (std.mem.eql(u8, pattern.bytes[p_start..p_end], target.bytes[t_start..t_end])) {
            return .{ .pattern = pi - p0, .target = boundary - t0 };
        }
    }
    return null;
}

/// Which pattern code points are escaped by a preceding `~`, and the
/// pattern with its `~` characters removed. Both are needed: the fold
/// must not see the tildes, and the matcher must know which survivors
/// were literal.
fn stripEscapes(
    allocator: std.mem.Allocator,
    pattern: []const u8,
) Error!struct { text: []const u8, escaped: []const bool } {
    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(allocator);
    var flags: std.ArrayListUnmanaged(bool) = .empty;
    errdefer flags.deinit(allocator);

    var view = std.unicode.Utf8View.init(pattern) catch {
        for (pattern) |b| {
            try out.append(allocator, b);
            try flags.append(allocator, false);
        }
        const text = try out.toOwnedSlice(allocator);
        errdefer allocator.free(text);
        return .{ .text = text, .escaped = try flags.toOwnedSlice(allocator) };
    };
    var it = view.iterator();
    var pending_escape = false;
    while (it.nextCodepointSlice()) |cp| {
        if (!pending_escape and cp.len == 1 and cp[0] == '~') {
            pending_escape = true;
            continue;
        }
        try out.appendSlice(allocator, cp);
        try flags.append(allocator, pending_escape);
        pending_escape = false;
    }
    // A trailing `~` escapes nothing and is dropped, which is Excel's
    // behaviour and is spec-pinned here.
    const text = try out.toOwnedSlice(allocator);
    errdefer allocator.free(text);
    return .{ .text = text, .escaped = try flags.toOwnedSlice(allocator) };
}

// ─── matching a cell ─────────────────────────────────────────────

pub const Context = struct {
    allocator: std.mem.Allocator,
    collation: value.Collation,
    fidelity: value.Fidelity = .excel,
};

/// Does `cell` satisfy `criterion`?
///
/// Every path through this function is one row of §5.3b's
/// `criteria operand` column.
pub fn matches(ctx: Context, criterion: Criterion, cell: value.ScalarValue) Error!bool {
    const hit = try matchesPositive(ctx, criterion, cell);
    return if (criterion.relation == .ne) !hit else hit;
}

fn matchesPositive(ctx: Context, criterion: Criterion, cell: value.ScalarValue) Error!bool {
    switch (criterion.operand) {
        .empty => {
            // `""` and `"<>"`: the `.countblank_class` population — true
            // blanks plus cells holding the empty string.
            return cell == .blank or (cell == .text and cell.text.len == 0);
        },
        .err => |want| {
            // Errors match errors and nothing else. `COUNTIF` can match
            // them; `COUNT` cannot count them. Different questions.
            if (cell != .err) return false;
            return cell.err.eql(want);
        },
        .boolean => |want| {
            if (cell != .boolean) return false;
            return switch (criterion.relation) {
                .eq, .ne => cell.boolean == want,
                .lt => @intFromBool(cell.boolean) < @intFromBool(want),
                .le => @intFromBool(cell.boolean) <= @intFromBool(want),
                .gt => @intFromBool(cell.boolean) > @intFromBool(want),
                .ge => @intFromBool(cell.boolean) >= @intFromBool(want),
            };
        },
        .number => |want| {
            // Type-restricted: a text cell never satisfies a numeric
            // criterion, whatever §5.3b's cross-type order would say.
            if (cell != .number) return false;
            const rules = value.FpRules.of(ctx.fidelity);
            const d = value.applyZeroSnap(rules, cell.number - want, cell.number, want);
            return switch (criterion.relation) {
                .eq, .ne => d == 0,
                .lt => d < 0,
                .le => d <= 0,
                .gt => d > 0,
                .ge => d >= 0,
            };
        },
        .text => |want| {
            if (cell != .text) return false;
            if (criterion.is_pattern) {
                return matchWildcard(ctx, want, cell.text);
            }
            const order = ctx.collation.compare(ctx.allocator, cell.text, want) catch |e| {
                if (e == error.OutOfMemory) return error.OutOfMemory;
                return error.BadFold;
            };
            return switch (criterion.relation) {
                .eq, .ne => order == .eq,
                .lt => order == .lt,
                .le => order != .gt,
                .gt => order == .gt,
                .ge => order != .lt,
            };
        },
    }
}

fn matchWildcard(ctx: Context, pattern: []const u8, target: []const u8) Error!bool {
    const stripped = try stripEscapes(ctx.allocator, pattern);
    defer ctx.allocator.free(stripped.text);
    defer ctx.allocator.free(stripped.escaped);

    var p = try fold(ctx.allocator, ctx.collation.fold, stripped.text);
    defer freeFolded(ctx.allocator, &p);
    var t = try fold(ctx.allocator, ctx.collation.fold, target);
    defer freeFolded(ctx.allocator, &t);

    // `stripEscapes` and `fold` both count code points, and both count
    // the same ones — the flags line up with the folded units.
    assert(stripped.escaped.len == p.codePoints());
    return matchFolded(p, stripped.escaped, t);
}

/// `SEARCH`'s string half (M4f): the first ORIGINAL code-point index at
/// or after `from` where `pattern` occurs, folded and with wildcards
/// active. Null when it does not occur.
///
/// It lives here rather than in `text.zig` for the reason M4e gave for
/// putting lookup equality here: the escape rules, the fold and the
/// wildcard semantics are one thing, and a second copy of them in a
/// second module is how `SEARCH("~*",…)` and `COUNTIF(…,"~*")` end up
/// disagreeing about what a literal star is.
pub fn searchText(
    ctx: Context,
    pattern: []const u8,
    target: []const u8,
    from: usize,
) Error!?usize {
    const stripped = try stripEscapes(ctx.allocator, pattern);
    defer ctx.allocator.free(stripped.text);
    defer ctx.allocator.free(stripped.escaped);

    var p = try fold(ctx.allocator, ctx.collation.fold, stripped.text);
    defer freeFolded(ctx.allocator, &p);
    var t = try fold(ctx.allocator, ctx.collation.fold, target);
    defer freeFolded(ctx.allocator, &t);

    assert(stripped.escaped.len == p.codePoints());
    return searchFolded(p, stripped.escaped, t, from);
}

pub fn freeFolded(allocator: std.mem.Allocator, f: *Folded) void {
    allocator.free(f.bytes);
    allocator.free(f.starts);
    f.* = undefined;
}

// ─── aligned scanning (§5.6a) ────────────────────────────────────

/// One aligned position that satisfied every criterion.
pub const Hit = struct {
    row_offset: u32,
    col_offset: u32,
    /// The aggregation area's value there, or `.blank`.
    aggregate: value.ScalarValue,
};

/// Scan a criteria/aggregation pair the way §5.6a says: one N-way
/// ordered pass through `EvalEnv.alignedRangeIterator`, never a pairwise
/// zip per criteria pair, and never a per-coordinate walk.
///
/// `areas[0]` supplies the window dimensions. Areas `1..n` are the
/// remaining criteria ranges and, last, the aggregation range — the
/// caller decides which is which by ordering them, and `criteria` has
/// one entry per criteria area.
///
/// Blank runs are visited as runs. A criterion that a blank satisfies
/// (`""`, or `<>` against a non-empty operand) therefore has to account
/// for a run's `count` at once, which is exactly why the iterator
/// reports counts instead of positions.
pub fn scan(
    ctx: Context,
    environment: env.EvalEnv,
    areas: []const env.RangeRef,
    mode: env.AlignMode,
    criteria: []const Criterion,
    cursors: []usize,
    scratch: []value.ScalarValue,
    result: *ScanResult,
) (Error || env.Error)!void {
    assert(areas.len == cursors.len and areas.len == scratch.len);
    assert(criteria.len <= areas.len);
    // Reset the counters without discarding the caller's storage — a
    // blanket `result.* = .{}` would silently drop `hits`.
    const hits = result.hits;
    result.* = .{ .hits = hits };

    // Whether an all-blank position satisfies every criterion decides
    // what a blank run contributes, and it is the same answer for every
    // position in the run — so it is computed once, not per position.
    var blank_hits = true;
    for (criteria) |c| {
        if (!try matches(ctx, c, .blank)) {
            blank_hits = false;
            break;
        }
    }

    var it = try environment.alignedRangeIterator(areas, mode, cursors, scratch);
    while (try it.next()) |item| switch (item) {
        .blank_run => |run| {
            result.visited += run.count;
            if (blank_hits) result.matched += run.count;
        },
        .cells => |cells| {
            result.visited += 1;
            var all = true;
            for (criteria, 0..) |c, i| {
                if (!try matches(ctx, c, cells.values[i])) {
                    all = false;
                    break;
                }
            }
            if (!all) continue;
            result.matched += 1;
            if (result.hits.len > result.hit_count) {
                result.hits[result.hit_count] = .{
                    .row_offset = cells.row_offset,
                    .col_offset = cells.col_offset,
                    .aggregate = cells.values[areas.len - 1],
                };
            }
            result.hit_count += 1;
            if (cells.values[areas.len - 1] == .number) {
                result.numeric_total += cells.values[areas.len - 1].number;
                result.numeric_count += 1;
            }
        },
    };
}

pub const ScanResult = struct {
    /// Every aligned position the pass covered — the invariant that
    /// makes a run-based cursor safe to count with.
    visited: u64 = 0,
    matched: u64 = 0,
    /// Sum and count of the aggregation area's numeric values at
    /// matching positions. `SUMIF` and `AVERAGEIF` need nothing else.
    numeric_total: f64 = 0,
    numeric_count: u64 = 0,
    /// Optional caller-provided storage for the matching positions.
    hits: []Hit = &.{},
    hit_count: usize = 0,
};

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;
const coords = @import("zlsx_refs");
const casefold = @import("zlsx_casefold");

fn shippedFold(allocator: std.mem.Allocator, s: []const u8) anyerror![]u8 {
    return casefold.foldString(allocator, s);
}

fn ctxOf() Context {
    return .{ .allocator = testing.allocator, .collation = .{ .fold = shippedFold } };
}

fn crit(text: []const u8) !Criterion {
    return parse(.{ .text = text }, .excel);
}

fn matchText(criterion_text: []const u8, cell_text: []const u8) !bool {
    return matches(ctxOf(), try crit(criterion_text), .{ .text = cell_text });
}

test "parse: operator prefixes, longest first" {
    try testing.expectEqual(Relation.le, (try crit("<=5")).relation);
    try testing.expectEqual(Relation.ge, (try crit(">=5")).relation);
    try testing.expectEqual(Relation.ne, (try crit("<>5")).relation);
    try testing.expectEqual(Relation.lt, (try crit("<5")).relation);
    try testing.expectEqual(Relation.gt, (try crit(">5")).relation);
    try testing.expectEqual(Relation.eq, (try crit("=5")).relation);
    try testing.expectEqual(Relation.eq, (try crit("5")).relation);
    // `<=` must not be read as `<` followed by a literal `=`.
    try testing.expectEqual(@as(f64, 5), (try crit("<=5")).operand.number);
}

test "parse: the operand's type is what the criterion restricts to" {
    try testing.expectEqual(@as(f64, 5), (try crit(">5")).operand.number);
    try testing.expect((try crit(">TRUE")).operand == .boolean);
    try testing.expect((try crit("#N/A")).operand == .err);
    try testing.expect((try crit("apple")).operand == .text);
    try testing.expect((try crit("")).operand == .empty);
    try testing.expect((try crit("<>")).operand == .empty);

    // A non-text criterion value carries no operator and no wildcards.
    const from_number = try parse(value.ScalarValue.fromNumber(3), .excel);
    try testing.expectEqual(Relation.eq, from_number.relation);
    try testing.expectEqual(@as(f64, 3), from_number.operand.number);
    const from_blank = try parse(.blank, .excel);
    try testing.expect(from_blank.operand == .empty);
}

test "parse: locale-flavoured text refuses instead of guessing a number" {
    // The `.criteria` ingress, exactly as §5.4's caller→ingress table
    // says — so `">1,5"` is a refusal here for the same reason `A1+1`
    // is one in the evaluator.
    try testing.expectError(error.LocaleSensitiveInput, crit(">1,5"));
    try testing.expectError(error.LocaleSensitiveInput, crit("50%"));
    // And text that is not numeric under any locale is just text.
    try testing.expect((try crit(">abc")).operand == .text);
}

test "match: a numeric criterion is type-restricted" {
    const ctx = ctxOf();
    const c = try crit(">5");
    try testing.expect(try matches(ctx, c, value.ScalarValue.fromNumber(6)));
    try testing.expect(!try matches(ctx, c, value.ScalarValue.fromNumber(5)));
    // §5.3b's total preorder puts every text above every number. A
    // criterion does NOT: `">5"` counts numbers, not words.
    try testing.expect(!try matches(ctx, c, .{ .text = "zebra" }));
    try testing.expect(!try matches(ctx, c, .{ .boolean = true }));
    try testing.expect(!try matches(ctx, c, .blank));
    try testing.expect(!try matches(ctx, c, value.ScalarValue.errorOf(.na)));
}

test "match: a text criterion is type-restricted the same way" {
    const ctx = ctxOf();
    const c = try crit(">m");
    try testing.expect(try matches(ctx, c, .{ .text = "n" }));
    try testing.expect(!try matches(ctx, c, .{ .text = "a" }));
    try testing.expect(!try matches(ctx, c, value.ScalarValue.fromNumber(1e9)));
}

test "match: text equality is case-insensitive under collation_v1" {
    try testing.expect(try matchText("apple", "APPLE"));
    try testing.expect(try matchText("APPLE", "Apple"));
    // `ß` folds to `ss`, so these are the same text.
    try testing.expect(try matchText("straße", "STRASSE"));
    try testing.expect(try matchText("STRASSE", "straße"));
    try testing.expect(!try matchText("apple", "apples"));
}

test "match: `<>` inverts, including for blanks" {
    const ctx = ctxOf();
    try testing.expect(try matches(ctx, try crit("<>apple"), .{ .text = "pear" }));
    try testing.expect(!try matches(ctx, try crit("<>apple"), .{ .text = "APPLE" }));
    // `""` matches blanks and empty strings; `"<>"` matches everything
    // else, which is Excel's "non-blank" idiom.
    try testing.expect(try matches(ctx, try crit(""), .blank));
    try testing.expect(try matches(ctx, try crit(""), .{ .text = "" }));
    try testing.expect(!try matches(ctx, try crit(""), value.ScalarValue.fromNumber(0)));
    try testing.expect(!try matches(ctx, try crit("<>"), .blank));
    try testing.expect(!try matches(ctx, try crit("<>"), .{ .text = "" }));
    try testing.expect(try matches(ctx, try crit("<>"), value.ScalarValue.fromNumber(0)));
}

test "match: criteria can match error values, which COUNT cannot count" {
    const ctx = ctxOf();
    const c = try crit("#N/A");
    try testing.expect(try matches(ctx, c, value.ScalarValue.errorOf(.na)));
    try testing.expect(!try matches(ctx, c, value.ScalarValue.errorOf(.div0)));
    try testing.expect(!try matches(ctx, c, .{ .text = "#N/A" }));
    // An error criterion value, not just its spelling.
    const from_value = try parse(value.ScalarValue.errorOf(.div0), .excel);
    try testing.expect(try matches(ctx, from_value, value.ScalarValue.errorOf(.div0)));
}

test "match: booleans match booleans only" {
    const ctx = ctxOf();
    try testing.expect(try matches(ctx, try crit("TRUE"), .{ .boolean = true }));
    try testing.expect(!try matches(ctx, try crit("TRUE"), .{ .boolean = false }));
    // Cross-type is never equal (§5.3b), so the text `"TRUE"` does not
    // match the logical TRUE and vice versa.
    try testing.expect(!try matches(ctx, try crit("TRUE"), .{ .text = "TRUE" }));
    try testing.expect(!try matches(ctx, try crit("apple"), .{ .boolean = true }));
}

test "searchText: a substring, found case-insensitively, at a code-point index" {
    const ctx = ctxOf();
    try testing.expectEqual(@as(?usize, 0), try searchText(ctx, "ap", "apple", 0));
    try testing.expectEqual(@as(?usize, 1), try searchText(ctx, "PP", "apple", 0));
    try testing.expectEqual(@as(?usize, 4), try searchText(ctx, "e", "apple", 0));
    try testing.expectEqual(@as(?usize, null), try searchText(ctx, "z", "apple", 0));
    // Code points, not bytes: `é` is two bytes and one index, so `x`
    // is the fifth character and the fourth index.
    try testing.expectEqual(@as(?usize, 4), try searchText(ctx, "x", "caféx", 0));
}

test "searchText: `from` skips earlier hits without moving the answer" {
    const ctx = ctxOf();
    try testing.expectEqual(@as(?usize, 1), try searchText(ctx, "p", "apple", 0));
    try testing.expectEqual(@as(?usize, 2), try searchText(ctx, "p", "apple", 2));
    try testing.expectEqual(@as(?usize, null), try searchText(ctx, "p", "apple", 3));
    // Starting past the end finds nothing rather than wrapping.
    try testing.expectEqual(@as(?usize, null), try searchText(ctx, "a", "apple", 99));
}

test "searchText: an empty pattern matches where it starts looking" {
    // Excel's `SEARCH("",text)` is 1 — the empty string is everywhere,
    // and the first everywhere is the start position.
    const ctx = ctxOf();
    try testing.expectEqual(@as(?usize, 0), try searchText(ctx, "", "apple", 0));
    try testing.expectEqual(@as(?usize, 3), try searchText(ctx, "", "apple", 3));
    // Including at the very end, which is a position and not a
    // character — `SEARCH("","ab",3)` is 3 in Excel.
    try testing.expectEqual(@as(?usize, 5), try searchText(ctx, "", "apple", 5));
}

test "searchText: wildcards are active, and `~` still escapes them" {
    const ctx = ctxOf();
    try testing.expectEqual(@as(?usize, 1), try searchText(ctx, "p*e", "apple", 0));
    try testing.expectEqual(@as(?usize, 0), try searchText(ctx, "a?p", "apple", 0));
    try testing.expectEqual(@as(?usize, null), try searchText(ctx, "a?l", "apple", 0));
    // A literal star is findable, and only where it literally is.
    try testing.expectEqual(@as(?usize, 2), try searchText(ctx, "~*", "ab*c", 0));
    try testing.expectEqual(@as(?usize, null), try searchText(ctx, "~*", "abc", 0));
    // `*` matching nothing still matches at the earliest position.
    try testing.expectEqual(@as(?usize, 0), try searchText(ctx, "*", "apple", 0));
}

test "searchText: an expanding fold reports the original position" {
    // The reason the positional map exists. `ß` folds to `ss`, so a
    // search for `ss` matches it — and the answer must be where `ß` is
    // in the CALLER's string (index 1), not where `ss` is in the fold's
    // (offset 1 of 4 bytes, which would be index 1 only by accident).
    const ctx = ctxOf();
    try testing.expectEqual(@as(?usize, 1), try searchText(ctx, "ss", "aßb", 0));
    // …and `SS` finds it too, because the search is folded.
    try testing.expectEqual(@as(?usize, 1), try searchText(ctx, "SS", "aßb", 0));
    // But HALF an expansion is not a match: both sides must end on a
    // code-point boundary, which is the same rule that makes the
    // criterion `"?s"` fail against `"ß"` (M3b). A `SEARCH` that
    // answered 2 here would be claiming a position inside a character.
    try testing.expectEqual(@as(?usize, null), try searchText(ctx, "s", "aßb", 0));
    // `ﬃ` folds to three bytes from one code point; a match after it
    // reports index 2, not index 4.
    try testing.expectEqual(@as(?usize, 2), try searchText(ctx, "x", "aﬃx", 0));
}

test "searchText: `?` consumes one code point, however wide its fold" {
    // Version-INdependent by §5.4b: the compatibility version changes
    // what LEN counts, never what `?` consumes.
    const ctx = ctxOf();
    try testing.expectEqual(@as(?usize, 0), try searchText(ctx, "a?b", "aßb", 0));
    // An astral character is one code point, so one `?` takes it.
    try testing.expectEqual(@as(?usize, 0), try searchText(ctx, "a?b", "a\u{1F600}b", 0));
}

test "wildcards: `*` and `?` under the ordinary cases" {
    try testing.expect(try matchText("a*", "apple"));
    try testing.expect(try matchText("*e", "apple"));
    try testing.expect(try matchText("*pp*", "apple"));
    try testing.expect(try matchText("a?ple", "apple"));
    try testing.expect(!try matchText("a?le", "apple"));
    try testing.expect(try matchText("?????", "apple"));
    try testing.expect(!try matchText("????", "apple"));
    try testing.expect(try matchText("*", ""));
    try testing.expect(try matchText("*", "anything"));
    // Case-insensitive, like every other match.
    try testing.expect(try matchText("A*E", "apple"));
}

test "wildcards: `~` escapes, so a literal star is matchable" {
    try testing.expect(try matchText("~*", "*"));
    try testing.expect(!try matchText("~*", "x"));
    try testing.expect(try matchText("a~?c", "a?c"));
    try testing.expect(!try matchText("a~?c", "abc"));
    try testing.expect(try matchText("~~", "~"));
    // `~*` has no *active* wildcard, but the escape still has to be
    // stripped before anything can match literally — which is why
    // "is this a pattern" and "does it contain a wildcard" are two
    // different questions.
    try testing.expect(!hasWildcards("~*"));
    try testing.expect((try crit("~*")).is_pattern);
    try testing.expect(hasWildcards("a*b"));
    try testing.expect((try crit("a*b")).is_pattern);
}

test "wildcards: `?` consumes one code point of the ORIGINAL, not one folded unit" {
    // §5.4b's positional rule, which is the whole reason `Folded` keeps
    // a map. `ß` is one code point that folds to two units.
    try testing.expect(try matchText("?", "ß"));
    try testing.expect(!try matchText("??", "ß"));
    // …while a literal run still matches across the expansion.
    try testing.expect(try matchText("ss", "ß"));
    try testing.expect(try matchText("SS", "ß"));
    // …and must land on a code-point boundary, so half an expansion is
    // not a match.
    try testing.expect(!try matchText("?s", "ß"));
    try testing.expect(!try matchText("s?", "ß"));
    // Combined with `*`.
    try testing.expect(try matchText("*ss*", "straße"));
    try testing.expect(try matchText("stra?e", "straße"));
}

test "wildcards: astral code points count as one" {
    // A code point outside the BMP is four bytes and one `?`.
    try testing.expect(try matchText("?", "\u{1F600}"));
    try testing.expect(!try matchText("??", "\u{1F600}"));
    try testing.expect(try matchText("a?b", "a\u{1F600}b"));
    // Deseret, where the fold DOES change the code point but not its
    // count.
    try testing.expect(try matchText("?", "\u{10400}"));
    try testing.expect(try matchText("\u{10428}", "\u{10400}"));
}

test "wildcards: only equality honours them" {
    const ctx = ctxOf();
    // Under an ordering operator, `*` is an ordinary character.
    const c = try crit(">a*");
    try testing.expect(!c.is_pattern);
    try testing.expect(try matches(ctx, c, .{ .text = "b" }));
}

test "wildcards: backtracking terminates on the pathological shapes" {
    // The inputs that make a naive matcher exponential.
    try testing.expect(!try matchText("*a*a*a*a*a*a*b", "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"));
    try testing.expect(try matchText("*a*a*a*a*a*a*a", "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"));
    try testing.expect(try matchText("**********", "aaaa"));
}

test "fold: the positional map lines up with the folded bytes" {
    var f = try fold(testing.allocator, shippedFold, "aß\u{1F600}");
    defer freeFolded(testing.allocator, &f);
    try testing.expectEqual(@as(usize, 3), f.codePoints());
    try testing.expectEqualStrings("ass\u{1F600}", f.bytes);
    // Offsets: `a` at 0, `ß`→`ss` at 1, the emoji at 3.
    try testing.expectEqual(@as(?usize, 0), f.boundaryAt(0));
    try testing.expectEqual(@as(?usize, 1), f.boundaryAt(1));
    try testing.expectEqual(@as(?usize, 2), f.boundaryAt(3));
    // Inside the expansion there is no boundary — that is the check
    // that stops `"s"` from matching half a `ß`.
    try testing.expectEqual(@as(?usize, null), f.boundaryAt(2));

    // Per-code-point folding agrees with folding the whole string.
    const whole = try shippedFold(testing.allocator, "aß\u{1F600}");
    defer testing.allocator.free(whole);
    try testing.expectEqualStrings(whole, f.bytes);
}

test "fold: whole-string and per-code-point agree over a corpus" {
    const corpus = [_][]const u8{
        "",                  "a",        "ABC",
        "Straße",
        "ǅ",
        "ﬁ",
        "\u{10400}",
        "İ",
        "ΣΣΣ",
        "e\u{0301}",         "\u{FB03}",
        "µ",
        "ʼn",
        "a\u{0300}\u{0301}", "\u{1E9E}",
        "ß",
    };
    for (corpus) |s| {
        var f = try fold(testing.allocator, shippedFold, s);
        defer freeFolded(testing.allocator, &f);
        const whole = try shippedFold(testing.allocator, s);
        defer testing.allocator.free(whole);
        try testing.expectEqualStrings(whole, f.bytes);
        try testing.expectEqual(@as(u32, @intCast(f.bytes.len)), f.starts[f.starts.len - 1]);
    }
}

// ─── aligned scanning ────────────────────────────────────────────

fn areaOf(sheet: env.SheetIndex, a1: []const u8) env.RangeRef {
    const r = coords.parseRange(a1, .{ .dollar = .accept }) catch unreachable;
    return .{ .sheet = sheet, .range = r.normalized() };
}

test "scan: one ordered pass over criteria and aggregation ranges" {
    var fake = env.Fake.init(testing.allocator);
    defer fake.deinit();
    const sh = try fake.addSheet("S");
    // A = category, B = amount.
    try fake.putA1(sh, .stored, "A1", .{ .text = "apple" });
    try fake.putA1(sh, .stored, "B1", value.ScalarValue.fromNumber(10));
    try fake.putA1(sh, .stored, "A2", .{ .text = "pear" });
    try fake.putA1(sh, .stored, "B2", value.ScalarValue.fromNumber(20));
    try fake.putA1(sh, .stored, "A3", .{ .text = "APPLE" });
    try fake.putA1(sh, .stored, "B3", value.ScalarValue.fromNumber(30));
    // Row 4 is blank on both sides.

    const areas = [_]env.RangeRef{ areaOf(sh, "A1:A4"), areaOf(sh, "B1:B4") };
    var cursors: [2]usize = undefined;
    var scratch: [2]value.ScalarValue = undefined;
    var hits: [8]Hit = undefined;
    var out: ScanResult = .{ .hits = &hits };

    try scan(
        ctxOf(),
        fake.evalEnv(),
        &areas,
        .require_equal,
        &.{try crit("apple")},
        &cursors,
        &scratch,
        &out,
    );

    // Case-insensitive, so both apples matched.
    try testing.expectEqual(@as(u64, 2), out.matched);
    try testing.expectEqual(@as(f64, 40), out.numeric_total);
    try testing.expectEqual(@as(u64, 2), out.numeric_count);
    // Every position accounted for exactly once, blanks included.
    try testing.expectEqual(@as(u64, 4), out.visited);
    try testing.expectEqual(@as(usize, 2), out.hit_count);
    try testing.expectEqual(@as(u32, 0), hits[0].row_offset);
    try testing.expectEqual(@as(u32, 2), hits[1].row_offset);
}

test "scan: a blank run is counted as a run, not walked" {
    var fake = env.Fake.init(testing.allocator);
    defer fake.deinit();
    const sh = try fake.addSheet("S");
    try fake.putA1(sh, .stored, "A2", .{ .text = "x" });

    // A whole column: 1 048 576 positions, one of them occupied. A
    // per-coordinate scan would take minutes; this must be instant.
    const whole = env.RangeRef{ .sheet = sh, .range = .{
        .first = .{ .col = try coords.Col.fromZeroBased(0), .row = try coords.Row.fromOneBased(1) },
        .last = .{ .col = try coords.Col.fromZeroBased(0), .row = try coords.Row.fromOneBased(coords.max_row) },
    } };
    const values = env.RangeRef{ .sheet = sh, .range = .{
        .first = .{ .col = try coords.Col.fromZeroBased(1), .row = try coords.Row.fromOneBased(1) },
        .last = .{ .col = try coords.Col.fromZeroBased(1), .row = try coords.Row.fromOneBased(coords.max_row) },
    } };
    const areas = [_]env.RangeRef{ whole, values };
    var cursors: [2]usize = undefined;
    var scratch: [2]value.ScalarValue = undefined;
    var out: ScanResult = .{};

    // `""` matches blanks, so the run contributes its whole count in one
    // step — which is the only reason this test finishes.
    try scan(ctxOf(), fake.evalEnv(), &areas, .require_equal, &.{try crit("")}, &cursors, &scratch, &out);
    try testing.expectEqual(@as(u64, coords.max_row), out.visited);
    try testing.expectEqual(@as(u64, coords.max_row - 1), out.matched);
}

test "scan: SUMIF's projection lines a differently-sized range up" {
    var fake = env.Fake.init(testing.allocator);
    defer fake.deinit();
    const sh = try fake.addSheet("S");
    try fake.putA1(sh, .stored, "A1", .{ .text = "x" });
    try fake.putA1(sh, .stored, "A2", .{ .text = "y" });
    try fake.putA1(sh, .stored, "A3", .{ .text = "x" });
    try fake.putA1(sh, .stored, "C1", value.ScalarValue.fromNumber(1));
    try fake.putA1(sh, .stored, "C2", value.ScalarValue.fromNumber(2));
    try fake.putA1(sh, .stored, "C3", value.ScalarValue.fromNumber(4));

    // §5.6a: the sum range is written as one cell and PROJECTED to the
    // criteria range's dimensions. Excel's documented rule.
    const areas = [_]env.RangeRef{ areaOf(sh, "A1:A3"), areaOf(sh, "C1:C1") };
    var cursors: [2]usize = undefined;
    var scratch: [2]value.ScalarValue = undefined;
    var out: ScanResult = .{};
    try scan(ctxOf(), fake.evalEnv(), &areas, .project_from_first, &.{try crit("x")}, &cursors, &scratch, &out);
    try testing.expectEqual(@as(u64, 2), out.matched);
    try testing.expectEqual(@as(f64, 5), out.numeric_total);

    // Under `.require_equal` the same call is a shape mismatch, which
    // the caller reports as `#VALUE!`.
    try testing.expectError(
        error.ShapeMismatch,
        scan(ctxOf(), fake.evalEnv(), &areas, .require_equal, &.{try crit("x")}, &cursors, &scratch, &out),
    );
}

test "scan: several criteria over one N-way pass" {
    var fake = env.Fake.init(testing.allocator);
    defer fake.deinit();
    const sh = try fake.addSheet("S");
    inline for (.{ 1, 2, 3, 4 }) |i| {
        const row = std.fmt.comptimePrint("{d}", .{i});
        try fake.putA1(sh, .stored, "A" ++ row, .{ .text = if (i % 2 == 0) "even" else "odd" });
        try fake.putA1(sh, .stored, "B" ++ row, value.ScalarValue.fromNumber(i * 10));
        try fake.putA1(sh, .stored, "C" ++ row, value.ScalarValue.fromNumber(i));
    }

    const areas = [_]env.RangeRef{ areaOf(sh, "A1:A4"), areaOf(sh, "B1:B4"), areaOf(sh, "C1:C4") };
    var cursors: [3]usize = undefined;
    var scratch: [3]value.ScalarValue = undefined;
    var out: ScanResult = .{};
    try scan(
        ctxOf(),
        fake.evalEnv(),
        &areas,
        .require_equal,
        &.{ try crit("even"), try crit(">20") },
        &cursors,
        &scratch,
        &out,
    );
    // Rows 2 and 4 are "even"; of those, only row 4 has B > 20.
    try testing.expectEqual(@as(u64, 1), out.matched);
    try testing.expectEqual(@as(f64, 4), out.numeric_total);
}

test "checkAllAllocationFailures: criteria matching leaks nothing under OOM" {
    const H = struct {
        fn run(allocator: std.mem.Allocator) !void {
            const ctx = Context{ .allocator = allocator, .collation = .{ .fold = shippedFold } };
            const cases = [_][2][]const u8{
                .{ "a*e", "apple" },
                .{ "straße", "STRASSE" },
                .{ "?s", "ß" },
                .{ "~*", "*" },
                .{ "*\u{1F600}*", "a\u{1F600}b" },
            };
            for (cases) |c| {
                const criterion = try parse(.{ .text = c[0] }, .excel);
                _ = try matches(ctx, criterion, .{ .text = c[1] });
            }
        }
    };
    try testing.checkAllAllocationFailures(testing.allocator, H.run, .{});
}

// ─── fuzz (§8.1: criteria) ───────────────────────────────────────

fn fuzzCriteriaTarget(_: void, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();
    var pattern_buf: [128]u8 = undefined;
    var target_buf: [128]u8 = undefined;
    const pattern = pattern_buf[0..smith.slice(&pattern_buf)];
    const target = target_buf[0..smith.slice(&target_buf)];

    const ctx = Context{
        .allocator = std.testing.allocator,
        .collation = .{ .fold = shippedFold },
    };
    const criterion = parse(.{ .text = pattern }, .excel) catch |e| switch (e) {
        // A locale-flavoured operand is a refusal, not a match failure.
        error.LocaleSensitiveInput => return,
        else => return e,
    };

    const cells = [_]value.ScalarValue{
        .{ .text = target },
        .blank,
        .{ .text = "" },
        .{ .boolean = true },
        value.ScalarValue.fromNumber(5),
        value.ScalarValue.errorOf(.na),
    };
    for (cells) |cell| {
        const first = matches(ctx, criterion, cell) catch |e| switch (e) {
            error.OutOfMemory => return e,
            else => continue,
        };
        // Deterministic: the same criterion and the same cell must agree
        // with themselves. A matcher with a stale cursor or an
        // uninitialised flag fails here and nowhere else.
        const second = try matches(ctx, criterion, cell);
        try std.testing.expectEqual(first, second);

        // `<>` is exactly the negation of `=` over the same operand.
        var inverted = criterion;
        inverted.relation = switch (criterion.relation) {
            .eq => .ne,
            .ne => .eq,
            else => continue,
        };
        try std.testing.expectEqual(first, !try matches(ctx, inverted, cell));
    }
}

test "fuzz: no criterion can panic, leak, or match non-deterministically" {
    try std.testing.fuzz({}, fuzzCriteriaTarget, .{
        .corpus = &[_][]const u8{
            "*",        "?",         "~",     "~~",       "~*",
            "a*b",      "*a*a*a*b",  "?????",
            "ß",
            "ss",       ">5",        "<=5",   "<>",       "",
            "#N/A",     "TRUE",      ">1,5",  "\xFF\xFE", "*\x00*",
            "apple",    "\u{1F600}",
            "İ",
            "\u{FB03}", "*?*?*?*",   "~?",
        },
    });
}
