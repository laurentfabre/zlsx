//! Typed-overlay parser for `xl/sharedStrings.xml`.
//!
//! This is a *thin*, allocation-light view over the SST XML. It does
//! NOT decode XML entities. `StringEntry.plain`, `RichRun.text`,
//! `RichRun.font_name`, `RichRun.color_argb`, and `RichRun.underline`
//! are RAW byte slices borrowed directly from the input `xml` buffer.
//! Callers that need decoded bytes invoke `decodeText(allocator, raw)`
//! on demand. Rationale: the eager reader in `src/xlsx.zig` already
//! does decode-on-parse for the hot read path; this overlay is for
//! tooling that wants structural visibility without paying the decode
//! cost upfront, and it keeps `parse` linear and predictable.
//!
//! Lifetime: every borrowed slice in the returned tree must not
//! outlive the input `xml` slice. Only the spine (`entries` array,
//! per-entry `RichRun` arrays) is owned by the returned `SstXml`'s
//! arena and freed by `deinit`.
//!
//! Stdlib only. Zig 0.15.2.

const std = @import("std");
const assert = std.debug.assert;

// ─── Public types ─────────────────────────────────────────────────────

pub const Error = error{
    MalformedXml,
    UnknownEntity,
    BadNumericRef,
    UnterminatedEntity,
    OutOfMemory,
};

/// A single rich-text run inside a shared string. Slice fields borrow
/// from the source `xml` buffer; they are NOT entity-decoded. Callers
/// that need the decoded form pipe each slice through `decodeText`.
pub const RichRun = struct {
    /// Inner text of `<t>...</t>` for this run. RAW; not decoded.
    text: []const u8,
    /// `<rFont val="…"/>` (or `<name val="…"/>` for styles-form). RAW.
    font_name: ?[]const u8 = null,
    /// `<sz val="…"/>` parsed as f64 — null when absent or unparseable.
    size: ?f64 = null,
    /// `<b/>` or `<b val="1"/>`. `<b val="0"/>` and absent → false.
    bold: bool = false,
    /// `<i/>` mirror of `bold`.
    italic: bool = false,
    /// `<color rgb="AARRGGBB"/>` — RAW hex slice. Theme/indexed colors
    /// are NOT resolved here; that's the workbook's job.
    color_argb: ?[]const u8 = null,
    /// `<u/>` or `<u val="…"/>` — one of `single`, `double`,
    /// `singleAccounting`, `doubleAccounting`. Null when no `<u>` child
    /// is present. `<u val="none"/>` is treated as null too.
    underline: ?[]const u8 = null,
};

/// One `<si>` entry. The OOXML schema says an `<si>` is either a
/// single `<t>` (plain) or a sequence of `<r><rPr/><t/></r>` (rich).
/// Tooling that doesn't care about formatting can match on `.plain`
/// and treat `.rich` as a concatenation of run texts.
pub const StringEntry = union(enum) {
    /// RAW inner-text of the single `<t>`. Not entity-decoded.
    plain: []const u8,
    /// One run per `<r>` child. `len == 0` is reserved for an SST entry
    /// that contained `<r>` markers but no `<t>` children — degenerate
    /// but observed in the wild.
    rich: []RichRun,
};

/// Typed view over `xl/sharedStrings.xml`. The arena owns the spine
/// (`entries`, per-entry `[]RichRun`); leaf slices borrow from `xml`.
pub const SstXml = struct {
    entries: []StringEntry,
    /// `count="…"` attribute on `<sst>`, when present and parseable.
    /// Per OOXML this is the total number of `<si>` references across
    /// the workbook (NOT entries.len). Informational only.
    total_count: ?u32,
    /// `uniqueCount="…"` — should equal `entries.len` for well-formed
    /// inputs but generators lie, so we don't assert equality here.
    unique_count: ?u32,
    arena: ?std.heap.ArenaAllocator,

    pub fn deinit(self: *SstXml, allocator: std.mem.Allocator) void {
        _ = allocator; // arena owns everything; allocator only here for
        // API symmetry with sibling typed-overlays.
        if (self.arena) |*a| a.deinit();
        self.* = undefined;
    }
};

// ─── Public API ───────────────────────────────────────────────────────

/// Parse `xml` (which must be the bytes of `xl/sharedStrings.xml`)
/// into a typed overlay. The returned `SstXml` borrows leaf slices
/// from `xml`; the caller must keep `xml` alive for the lifetime of
/// the returned value.
///
/// Errors are precise: `MalformedXml` for structural breakage we can't
/// recover from, `OutOfMemory` from the arena. Entity-related errors
/// surface only when a caller invokes `decodeText` later — `parse`
/// never inspects `&` sequences.
pub fn parse(
    allocator: std.mem.Allocator,
    xml: []const u8,
) Error!SstXml {
    assert(xml.len < std.math.maxInt(usize) / 2); // sanity bound
    assert(@intFromPtr(allocator.vtable) != 0);

    var arena = std.heap.ArenaAllocator.init(allocator);
    errdefer arena.deinit();
    const a = arena.allocator();

    // Locate `<sst …>` and parse its count attributes.
    const counts = parseSstAttrs(xml);

    var entries: std.ArrayList(StringEntry) = .empty;
    errdefer entries.deinit(a);
    if (counts.unique_count) |hint| {
        // Cap against the smallest possible `<si/>` (5 bytes) so a
        // hostile uniqueCount can't force a giant upfront alloc.
        const safe_hint: usize = @min(hint, xml.len / 5 + 1);
        try entries.ensureTotalCapacity(a, safe_hint);
    } else {
        try entries.ensureTotalCapacity(a, 64);
    }

    var i: usize = 0;
    while (i < xml.len) {
        const i_prev = i;
        const lt = std.mem.indexOfScalarPos(u8, xml, i, '<') orelse break;

        // We only act on `<si` opens. Skip comments / PI / CDATA / any
        // other tag without false-positive matches.
        const skip = skipNonSi(xml, lt) orelse break;
        if (skip.is_si_open) {
            assert(skip.next_index > lt);
            // Process this `<si>`. `skip.next_index` is one past the
            // opening tag's `>`, so `next_index - 1` is the `>` itself.
            const si_open_gt = skip.next_index - 1;
            assert(xml[si_open_gt] == '>');

            // Self-closing `<si/>` → empty entry.
            if (si_open_gt > 0 and xml[si_open_gt - 1] == '/') {
                try entries.append(a, .{ .plain = "" });
                i = si_open_gt + 1;
                assert(i > i_prev);
                continue;
            }

            // Locate matching `</si>`.
            const si_close = std.mem.indexOfPos(u8, xml, si_open_gt + 1, "</si>") orelse {
                // Malformed — opening `<si>` with no close. Bail out
                // structurally rather than swallow.
                return error.MalformedXml;
            };
            const body = xml[si_open_gt + 1 .. si_close];
            const entry = try parseSiBody(a, body);
            try entries.append(a, entry);
            i = si_close + "</si>".len;
            assert(i > i_prev);
        } else {
            i = skip.next_index;
            assert(i > i_prev);
        }
    }

    return SstXml{
        .entries = try entries.toOwnedSlice(a),
        .total_count = counts.count,
        .unique_count = counts.unique_count,
        .arena = arena,
    };
}

/// Decode the five canonical XML entities (`&amp; &lt; &gt; &quot;
/// &apos;`) plus numeric character references (`&#N;`, `&#xN;`) in
/// `raw`. Allocates a fresh `[]u8` owned by `allocator`. Returns
/// `UnknownEntity`, `UnterminatedEntity`, or `BadNumericRef` on
/// malformed input — strict by design: a typed-overlay caller asking
/// for decoded bytes deserves a precise error, not silent passthrough.
///
/// Algorithm mirrors `pkg/store.zig`'s `decodeXmlEntities` but is
/// strict where store's variant is permissive (store falls back to
/// literal `&` on unknown refs because it was originally written for
/// path-name decoding where false positives are recoverable).
pub fn decodeText(
    allocator: std.mem.Allocator,
    raw: []const u8,
) Error![]u8 {
    assert(@intFromPtr(allocator.vtable) != 0);
    // Fast path: no `&` anywhere → straight dupe.
    if (std.mem.indexOfScalar(u8, raw, '&') == null) {
        return allocator.dupe(u8, raw);
    }

    var out: std.ArrayList(u8) = .empty;
    errdefer out.deinit(allocator);
    try out.ensureTotalCapacity(allocator, raw.len);

    var i: usize = 0;
    while (i < raw.len) {
        const i_prev = i;
        if (raw[i] != '&') {
            try out.append(allocator, raw[i]);
            i += 1;
            assert(i > i_prev);
            continue;
        }
        const remain = raw[i..];
        // Named refs first.
        if (std.mem.startsWith(u8, remain, "&amp;")) {
            try out.append(allocator, '&');
            i += 5;
        } else if (std.mem.startsWith(u8, remain, "&lt;")) {
            try out.append(allocator, '<');
            i += 4;
        } else if (std.mem.startsWith(u8, remain, "&gt;")) {
            try out.append(allocator, '>');
            i += 4;
        } else if (std.mem.startsWith(u8, remain, "&quot;")) {
            try out.append(allocator, '"');
            i += 6;
        } else if (std.mem.startsWith(u8, remain, "&apos;")) {
            try out.append(allocator, '\'');
            i += 6;
        } else if (std.mem.startsWith(u8, remain, "&#")) {
            const info = decodeNumericRef(remain) orelse return error.BadNumericRef;
            try out.appendSlice(allocator, info.utf8[0..info.utf8_len]);
            i += info.consumed;
        } else {
            // `&` not followed by a recognised entity prefix. Try to
            // distinguish "unknown named entity" (`&foo;`) from
            // "unterminated" (`&` followed by EOF or whitespace).
            const semi = std.mem.indexOfScalarPos(u8, raw, i + 1, ';') orelse
                return error.UnterminatedEntity;
            // Reject things like `&fo;` — strictly named-ref shape but
            // not one of the five we handle.
            _ = semi;
            return error.UnknownEntity;
        }
        assert(i > i_prev);
    }

    return out.toOwnedSlice(allocator);
}

// ─── Internals ────────────────────────────────────────────────────────

const SstAttrs = struct {
    count: ?u32 = null,
    unique_count: ?u32 = null,
};

/// Pull `count` and `uniqueCount` from `<sst …>`. Returns zeroed
/// struct when no `<sst` open is found. Quoted-attribute aware so a
/// stray `>` inside an attribute value doesn't truncate the tag.
fn parseSstAttrs(xml: []const u8) SstAttrs {
    var out: SstAttrs = .{};
    const open = std.mem.indexOf(u8, xml, "<sst") orelse return out;
    const tag_end = findTagEndQuoted(xml, open + 4) orelse return out;
    assert(tag_end >= open + 4);
    assert(tag_end < xml.len);
    const attrs = xml[open..tag_end];

    if (extractAttr(attrs, "count")) |raw| {
        out.count = std.fmt.parseInt(u32, raw, 10) catch null;
    }
    if (extractAttr(attrs, "uniqueCount")) |raw| {
        out.unique_count = std.fmt.parseInt(u32, raw, 10) catch null;
    }
    return out;
}

/// Find the closing `>` of a tag, respecting `"…"`-quoted attribute
/// values. Returns the index of the `>` itself.
fn findTagEndQuoted(xml: []const u8, from: usize) ?usize {
    assert(from <= xml.len);
    var i: usize = from;
    var in_quote: bool = false;
    var quote_ch: u8 = 0;
    while (i < xml.len) : (i += 1) {
        const c = xml[i];
        if (in_quote) {
            if (c == quote_ch) in_quote = false;
            continue;
        }
        if (c == '"' or c == '\'') {
            in_quote = true;
            quote_ch = c;
            continue;
        }
        if (c == '>') return i;
    }
    return null;
}

/// `name="value"` extraction inside an attribute blob. Returns the
/// raw (un-decoded) value slice. Borrows from `attrs`.
fn extractAttr(attrs: []const u8, name: []const u8) ?[]const u8 {
    assert(name.len > 0);
    assert(attrs.len < std.math.maxInt(usize));
    // Search for ` name="` so we don't false-match `xuniqueCount=`.
    var search_from: usize = 0;
    while (true) {
        const pos = std.mem.indexOfPos(u8, attrs, search_from, name) orelse return null;
        // Boundary check — char before must be whitespace or '<'.
        const left_ok = pos == 0 or attrs[pos - 1] == ' ' or
            attrs[pos - 1] == '\t' or attrs[pos - 1] == '\n' or
            attrs[pos - 1] == '\r' or attrs[pos - 1] == '<';
        const after = pos + name.len;
        if (after >= attrs.len) return null;
        if (left_ok and attrs[after] == '=') {
            const q_pos = after + 1;
            if (q_pos >= attrs.len) return null;
            const quote = attrs[q_pos];
            if (quote != '"' and quote != '\'') return null;
            const start = q_pos + 1;
            const end = std.mem.indexOfScalarPos(u8, attrs, start, quote) orelse return null;
            return attrs[start..end];
        }
        search_from = pos + 1;
    }
}

const SkipResult = struct {
    is_si_open: bool,
    /// One past the consumed region. For an `<si` we set this to the
    /// index AFTER the opening tag's `>`; for a non-si tag, past the
    /// closing `>`; for `<` followed by no recognisable structure,
    /// `lt + 1` so the outer loop strictly advances.
    next_index: usize,
};

/// Look at `xml[lt]` (must be `<`) and decide what to do. Comments,
/// CDATA, PI, and arbitrary tags are all skipped without false-
/// positive matches against `<si`.
fn skipNonSi(xml: []const u8, lt: usize) ?SkipResult {
    assert(lt < xml.len);
    assert(xml[lt] == '<');
    if (lt + 1 >= xml.len) return null;

    // `<!--` comment.
    if (lt + 4 <= xml.len and std.mem.startsWith(u8, xml[lt..], "<!--")) {
        const close = std.mem.indexOfPos(u8, xml, lt + 4, "-->") orelse return null;
        return .{ .is_si_open = false, .next_index = close + 3 };
    }
    // `<![CDATA[`.
    if (lt + 9 <= xml.len and std.mem.startsWith(u8, xml[lt..], "<![CDATA[")) {
        const close = std.mem.indexOfPos(u8, xml, lt + 9, "]]>") orelse return null;
        return .{ .is_si_open = false, .next_index = close + 3 };
    }
    // `<?…?>` processing instruction.
    if (xml[lt + 1] == '?') {
        const close = std.mem.indexOfPos(u8, xml, lt + 2, "?>") orelse return null;
        return .{ .is_si_open = false, .next_index = close + 2 };
    }

    // Tentative `<si` open — confirm the next byte isn't an identifier
    // continuation that would make this `<sst` or `<sheet` etc.
    if (lt + 3 <= xml.len and xml[lt + 1] == 's' and xml[lt + 2] == 'i') {
        if (lt + 3 == xml.len) return null;
        const c3 = xml[lt + 3];
        if (c3 == '>' or c3 == '/' or c3 == ' ' or c3 == '\t' or c3 == '\n' or c3 == '\r') {
            const gt = findTagEndQuoted(xml, lt + 3) orelse return null;
            return .{ .is_si_open = true, .next_index = gt + 1 };
        }
    }

    // Generic tag — find its `>` (quoted-aware) and step past.
    const gt = findTagEndQuoted(xml, lt + 1) orelse return null;
    return .{ .is_si_open = false, .next_index = gt + 1 };
}

/// Walk one `<si>...</si>` body and produce a `StringEntry`. Allocates
/// only when rich-text runs are present.
fn parseSiBody(a: std.mem.Allocator, body: []const u8) Error!StringEntry {
    assert(@intFromPtr(a.vtable) != 0);
    assert(body.len < std.math.maxInt(usize));

    // First scan: does the body contain any `<r` opener? If so, treat
    // as rich; otherwise, look for a single `<t>` and treat as plain.
    const has_r = containsRunOpener(body);

    if (!has_r) {
        // Plain — one `<t>...</t>` (or `<t/>` → empty). Tolerate
        // missing entirely (some generators emit `<si></si>`).
        return .{ .plain = extractFirstT(body) orelse "" };
    }

    var runs: std.ArrayList(RichRun) = .empty;
    errdefer runs.deinit(a);

    var i: usize = 0;
    while (i < body.len) {
        const i_prev = i;
        const lt = std.mem.indexOfScalarPos(u8, body, i, '<') orelse break;
        if (lt + 2 > body.len) break;

        // `<r>` or `<r ...>` (but NOT `<rPr` or `<rPh` at this layer).
        if (body[lt + 1] == 'r' and (lt + 2 == body.len or
            body[lt + 2] == '>' or body[lt + 2] == ' ' or
            body[lt + 2] == '\t' or body[lt + 2] == '\n'))
        {
            const r_open_gt = findTagEndQuoted(body, lt + 2) orelse return error.MalformedXml;
            const r_close = std.mem.indexOfPos(u8, body, r_open_gt + 1, "</r>") orelse
                return error.MalformedXml;
            const r_body = body[r_open_gt + 1 .. r_close];
            const run = parseRichRunBody(r_body);
            try runs.append(a, run);
            i = r_close + "</r>".len;
            assert(i > i_prev);
        } else {
            // Skip any other tag (rPh strays at top level, comments,
            // etc.). findTagEndQuoted guarantees monotonic progress.
            const gt = findTagEndQuoted(body, lt + 1) orelse break;
            i = gt + 1;
            assert(i > i_prev);
        }
    }

    return .{ .rich = try runs.toOwnedSlice(a) };
}

/// True iff `body` contains a `<r>` or `<r ...>` opener that is NOT
/// `<rPr`, `<rPh`, etc.
fn containsRunOpener(body: []const u8) bool {
    var search: usize = 0;
    while (std.mem.indexOfPos(u8, body, search, "<r")) |pos| {
        const after = pos + 2;
        if (after >= body.len) return false;
        const c = body[after];
        if (c == '>' or c == ' ' or c == '\t' or c == '\n' or c == '\r' or c == '/') {
            return true;
        }
        search = pos + 1;
    }
    return false;
}

/// Extract the inner text of the first `<t>` element in `body`.
/// Returns null when no `<t>` is found; returns "" for `<t/>` or
/// `<t></t>`.
fn extractFirstT(body: []const u8) ?[]const u8 {
    const t_open = std.mem.indexOf(u8, body, "<t") orelse return null;
    if (t_open + 2 >= body.len) return null;
    const c = body[t_open + 2];
    // Must be `<t>`, `<t ...>`, or `<t/>`. Reject `<title` etc.
    if (c != '>' and c != ' ' and c != '/' and c != '\t' and c != '\n') return null;
    const gt = findTagEndQuoted(body, t_open + 2) orelse return null;
    if (gt > 0 and body[gt - 1] == '/') return ""; // self-closing
    const close = std.mem.indexOfPos(u8, body, gt + 1, "</t>") orelse return null;
    return body[gt + 1 .. close];
}

/// Parse the body of `<r>...</r>` into a `RichRun`. `text` borrows
/// from `body` (and thus from the original xml). RAW; not decoded.
fn parseRichRunBody(body: []const u8) RichRun {
    assert(body.len < std.math.maxInt(usize));

    var run: RichRun = .{ .text = "" };
    if (extractFirstT(body)) |t| run.text = t;

    // `<rPr>...</rPr>` — formatting properties live here. Bounded
    // search; absent rPr is fine, run keeps defaults.
    const rpr_open = std.mem.indexOf(u8, body, "<rPr");
    if (rpr_open) |rp| {
        const rpr_close = std.mem.indexOfPos(u8, body, rp, "</rPr>") orelse return run;
        const rpr_end = rpr_close + "</rPr>".len;
        assert(rpr_end <= body.len);
        const rpr = body[rp..rpr_end];
        applyRprToRun(rpr, &run);
    }
    return run;
}

/// Mutate `run` from the contents of an `<rPr>...</rPr>` blob. RAW
/// borrows; no allocation.
fn applyRprToRun(rpr: []const u8, run: *RichRun) void {
    assert(rpr.len > 0);
    assert(@intFromPtr(run) != 0);

    run.bold = boolFlagPresent(rpr, "<b");
    run.italic = boolFlagPresent(rpr, "<i");

    if (findTagAttrValue(rpr, "<sz", "val")) |raw| {
        run.size = std.fmt.parseFloat(f64, raw) catch null;
    }
    // Try `<rFont val="…"/>` first (rich-text form), then `<name`.
    if (findTagAttrValue(rpr, "<rFont", "val")) |raw| {
        run.font_name = raw;
    } else if (findTagAttrValue(rpr, "<name", "val")) |raw| {
        run.font_name = raw;
    }
    if (findTagAttrValue(rpr, "<color", "rgb")) |raw| {
        run.color_argb = raw;
    }

    // Underline: `<u/>` (defaults to "single"), `<u val="…"/>`.
    // Treat val="none" as null.
    if (findElement(rpr, "<u")) |slice| {
        if (findTagAttrValue(slice, "<u", "val")) |val| {
            if (std.mem.eql(u8, val, "none")) {
                run.underline = null;
            } else {
                run.underline = val;
            }
        } else {
            // `<u/>` or `<u>` with no val — implicit "single".
            run.underline = "single";
        }
    }
}

/// True if a self-closing or attribute-bearing element `tag_open`
/// (e.g. `"<b"`) appears in `rpr` with either no `val` attribute or
/// `val` not in {`0`, `false`}. Mirrors OOXML's "missing val = true"
/// convention.
fn boolFlagPresent(rpr: []const u8, tag_open: []const u8) bool {
    assert(tag_open.len >= 2);
    const pos = std.mem.indexOf(u8, rpr, tag_open) orelse return false;
    if (pos + tag_open.len >= rpr.len) return false;
    const c = rpr[pos + tag_open.len];
    // Must be a tag boundary — otherwise `<b` could match `<bgColor`.
    if (c != '>' and c != '/' and c != ' ' and c != '\t' and c != '\n') return false;
    const gt = findTagEndQuoted(rpr, pos + tag_open.len) orelse return false;
    const tag = rpr[pos..gt];
    if (std.mem.indexOf(u8, tag, "val=\"0\"") != null) return false;
    if (std.mem.indexOf(u8, tag, "val=\"false\"") != null) return false;
    return true;
}

/// Look up `attr="…"` on the element opening with `tag_open` (e.g.
/// `"<sz"`). Returns the raw value slice, or null when the element or
/// attribute is missing. Element-boundary check prevents `<sz` from
/// matching `<szSomething`.
fn findTagAttrValue(rpr: []const u8, tag_open: []const u8, attr: []const u8) ?[]const u8 {
    assert(tag_open.len >= 2);
    assert(attr.len > 0);
    var search: usize = 0;
    while (std.mem.indexOfPos(u8, rpr, search, tag_open)) |pos| {
        const after = pos + tag_open.len;
        if (after >= rpr.len) return null;
        const c = rpr[after];
        if (c == '>' or c == '/' or c == ' ' or c == '\t' or c == '\n') {
            const gt = findTagEndQuoted(rpr, after) orelse return null;
            const tag = rpr[pos..gt];
            return extractAttr(tag, attr);
        }
        search = pos + 1;
    }
    return null;
}

/// Find the byte slice for the element starting with `tag_open` (up
/// to its `>`). Used to detect the *presence* of an element when
/// attribute-less variants matter (e.g. `<u/>`).
fn findElement(rpr: []const u8, tag_open: []const u8) ?[]const u8 {
    assert(tag_open.len >= 2);
    var search: usize = 0;
    while (std.mem.indexOfPos(u8, rpr, search, tag_open)) |pos| {
        const after = pos + tag_open.len;
        if (after >= rpr.len) return null;
        const c = rpr[after];
        if (c == '>' or c == '/' or c == ' ' or c == '\t' or c == '\n') {
            const gt = findTagEndQuoted(rpr, after) orelse return null;
            return rpr[pos .. gt + 1];
        }
        search = pos + 1;
    }
    return null;
}

// ─── Numeric character reference decoder ─────────────────────────────
// Algorithm cribbed from `pkg/store.zig:decodeNumericRef` (NOT
// imported — copied to keep this file self-contained per the build
// graph constraint that typed_parts/* won't import store.zig). Same
// validation: digit-class check, code-point cap, C0 control reject.

const NumericRef = struct {
    utf8: [4]u8,
    utf8_len: u3,
    consumed: usize,
};

fn decodeNumericRef(s: []const u8) ?NumericRef {
    assert(s.len > 0);
    assert(s[0] == '&');
    if (s.len < 4) return null; // need at least "&#0;"
    if (s[1] != '#') return null;
    var digit_start: usize = 2;
    var base: u8 = 10;
    if (s[2] == 'x' or s[2] == 'X') {
        digit_start = 3;
        base = 16;
    }
    const semi = std.mem.indexOfScalarPos(u8, s, digit_start, ';') orelse return null;
    if (semi == digit_start) return null; // empty digit run
    const digits = s[digit_start..semi];
    for (digits) |c| {
        const ok = if (base == 10)
            (c >= '0' and c <= '9')
        else
            ((c >= '0' and c <= '9') or (c >= 'a' and c <= 'f') or (c >= 'A' and c <= 'F'));
        if (!ok) return null;
    }
    const code = std.fmt.parseInt(u32, digits, base) catch return null;
    if (code > 0x10FFFF) return null;
    if (code < 0x20 and code != 0x09 and code != 0x0A and code != 0x0D) return null;
    var ref: NumericRef = .{ .utf8 = undefined, .utf8_len = 0, .consumed = semi + 1 };
    const len = std.unicode.utf8Encode(@intCast(code), &ref.utf8) catch return null;
    // `len` is in 1..=4; @intCast traps if not.
    ref.utf8_len = @intCast(len);
    return ref;
}

// ─── Tests ────────────────────────────────────────────────────────────

test "parse: minimal one-entry SST yields plain entry borrowing from xml" {
    const xml =
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>hello</t></si></sst>
    ;
    var sst = try parse(std.testing.allocator, xml);
    defer sst.deinit(std.testing.allocator);

    try std.testing.expectEqual(@as(usize, 1), sst.entries.len);
    try std.testing.expectEqual(@as(?u32, 1), sst.total_count);
    try std.testing.expectEqual(@as(?u32, 1), sst.unique_count);
    switch (sst.entries[0]) {
        .plain => |s| {
            try std.testing.expectEqualStrings("hello", s);
            // Confirm RAW borrow: pointer must lie inside `xml`.
            const base = @intFromPtr(xml.ptr);
            const ptr = @intFromPtr(s.ptr);
            try std.testing.expect(ptr >= base);
            try std.testing.expect(ptr < base + xml.len);
        },
        .rich => return error.TestUnexpectedResult,
    }
}

test "parse: plain and rich entries mixed, both borrow from xml" {
    const xml =
        \\<sst count="2" uniqueCount="2">
        \\  <si><t>plain text</t></si>
        \\  <si><r><rPr><b/></rPr><t>BOLD</t></r><r><t> tail</t></r></si>
        \\</sst>
    ;
    var sst = try parse(std.testing.allocator, xml);
    defer sst.deinit(std.testing.allocator);

    try std.testing.expectEqual(@as(usize, 2), sst.entries.len);
    switch (sst.entries[0]) {
        .plain => |s| try std.testing.expectEqualStrings("plain text", s),
        .rich => return error.TestUnexpectedResult,
    }
    switch (sst.entries[1]) {
        .plain => return error.TestUnexpectedResult,
        .rich => |runs| {
            try std.testing.expectEqual(@as(usize, 2), runs.len);
            try std.testing.expectEqualStrings("BOLD", runs[0].text);
            try std.testing.expect(runs[0].bold);
            try std.testing.expect(!runs[0].italic);
            try std.testing.expectEqualStrings(" tail", runs[1].text);
            try std.testing.expect(!runs[1].bold);
        },
    }
}

test "parse: rich-run with multiple format properties (bold, italic, size, font, color, underline)" {
    const xml =
        \\<sst count="1" uniqueCount="1"><si><r><rPr><b/><i/><sz val="14.5"/><rFont val="Calibri"/><color rgb="FFAABBCC"/><u val="double"/></rPr><t>fancy</t></r></si></sst>
    ;
    var sst = try parse(std.testing.allocator, xml);
    defer sst.deinit(std.testing.allocator);

    try std.testing.expectEqual(@as(usize, 1), sst.entries.len);
    const runs = switch (sst.entries[0]) {
        .plain => return error.TestUnexpectedResult,
        .rich => |r| r,
    };
    try std.testing.expectEqual(@as(usize, 1), runs.len);
    const r = runs[0];
    try std.testing.expectEqualStrings("fancy", r.text);
    try std.testing.expect(r.bold);
    try std.testing.expect(r.italic);
    try std.testing.expectEqual(@as(?f64, 14.5), r.size);
    try std.testing.expectEqualStrings("Calibri", r.font_name.?);
    try std.testing.expectEqualStrings("FFAABBCC", r.color_argb.?);
    try std.testing.expectEqualStrings("double", r.underline.?);
}

test "parse: count attributes parsed, missing → null" {
    const xml_with =
        \\<sst count="42" uniqueCount="7"><si><t>x</t></si></sst>
    ;
    var sst1 = try parse(std.testing.allocator, xml_with);
    defer sst1.deinit(std.testing.allocator);
    try std.testing.expectEqual(@as(?u32, 42), sst1.total_count);
    try std.testing.expectEqual(@as(?u32, 7), sst1.unique_count);

    const xml_without =
        \\<sst><si><t>x</t></si></sst>
    ;
    var sst2 = try parse(std.testing.allocator, xml_without);
    defer sst2.deinit(std.testing.allocator);
    try std.testing.expectEqual(@as(?u32, null), sst2.total_count);
    try std.testing.expectEqual(@as(?u32, null), sst2.unique_count);
}

test "parse: malformed input — unterminated <si> rejects" {
    const xml =
        \\<sst><si><t>oops
    ;
    try std.testing.expectError(error.MalformedXml, parse(std.testing.allocator, xml));
}

test "parse: comment / CDATA / PI tolerated and skipped" {
    const xml =
        \\<?xml version="1.0"?>
        \\<sst><!-- a comment with <si> inside --><si><t>real</t></si><![CDATA[<si>not really</si>]]></sst>
    ;
    var sst = try parse(std.testing.allocator, xml);
    defer sst.deinit(std.testing.allocator);

    try std.testing.expectEqual(@as(usize, 1), sst.entries.len);
    switch (sst.entries[0]) {
        .plain => |s| try std.testing.expectEqualStrings("real", s),
        .rich => return error.TestUnexpectedResult,
    }
}

test "parse: self-closing <si/> emits empty plain entry" {
    const xml =
        \\<sst count="2" uniqueCount="2"><si/><si><t>after</t></si></sst>
    ;
    var sst = try parse(std.testing.allocator, xml);
    defer sst.deinit(std.testing.allocator);

    try std.testing.expectEqual(@as(usize, 2), sst.entries.len);
    switch (sst.entries[0]) {
        .plain => |s| try std.testing.expectEqualStrings("", s),
        .rich => return error.TestUnexpectedResult,
    }
    switch (sst.entries[1]) {
        .plain => |s| try std.testing.expectEqualStrings("after", s),
        .rich => return error.TestUnexpectedResult,
    }
}

test "decodeText: happy path decodes named + numeric refs" {
    const a = std.testing.allocator;
    {
        const out = try decodeText(a, "no entities here");
        defer a.free(out);
        try std.testing.expectEqualStrings("no entities here", out);
    }
    {
        const out = try decodeText(a, "a &amp; b &lt;c&gt; &quot;q&quot; &apos;a&apos;");
        defer a.free(out);
        try std.testing.expectEqualStrings("a & b <c> \"q\" 'a'", out);
    }
    {
        // &#38; = '&', &#x26; = '&', &#x20AC; = €
        const out = try decodeText(a, "x&#38;y&#x26;z&#x20AC;w");
        defer a.free(out);
        try std.testing.expectEqualStrings("x&y&z\u{20AC}w", out);
    }
}

test "decodeText: malformed entities rejected with typed errors" {
    const a = std.testing.allocator;
    try std.testing.expectError(error.UnknownEntity, decodeText(a, "&unknown;"));
    try std.testing.expectError(error.UnterminatedEntity, decodeText(a, "abc & xyz"));
    try std.testing.expectError(error.BadNumericRef, decodeText(a, "&#xZ;"));
    try std.testing.expectError(error.BadNumericRef, decodeText(a, "&#;"));
    try std.testing.expectError(error.BadNumericRef, decodeText(a, "&#x"));
}
