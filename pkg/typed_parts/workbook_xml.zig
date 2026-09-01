//! Typed overlay parser for `xl/workbook.xml`.
//!
//! Mirrors what `src/xlsx.zig:parseWorkbookSheets` extracts (sheet
//! list) and extends it with the surface points the lower-level
//! reader currently doesn't capture: per-sheet visibility state,
//! workbook-scoped `<definedName>` entries, and `<calcPr>` properties.
//!
//! Stdlib-only. No third-party deps. Hand-rolled XML scanner — XML is
//! not regular, but the workbook.xml subset we accept is bounded
//! enough that a careful tag-by-tag walker is sound. The scanner
//! understands and skips:
//!
//!   - XML comments     `<!-- ... -->`
//!   - CDATA sections   `<![CDATA[ ... ]]>`
//!   - processing instr `<? ... ?>`
//!   - DOCTYPE          `<!DOCTYPE ...>`
//!
//! Attribute values may legally contain `>` when quoted; the
//! attribute-aware tag-end finder below honours that. Bare `>`
//! outside quoted attributes terminates a tag.
//!
//! ## Lifetime / borrowing
//!
//! Leaf string fields (`Sheet.name`, `Sheet.r_id`, `DefinedName.name`,
//! `DefinedName.formula`) are slices INTO the caller-provided `xml`
//! buffer. The caller MUST ensure `xml` outlives the returned
//! `WorkbookXml`. No entity decoding is performed on these slices —
//! consumers that need decoded text should run them through
//! `pkg/store.zig:decodeXmlEntities` after extraction. The rationale
//! is composability: callers that compare names against an already-
//! encoded reference (e.g. another workbook's raw XML) shouldn't pay
//! a decode round-trip.
//!
//! Backing storage for the `sheets` and `defined_names` slices lives
//! in an internal `ArenaAllocator` so partial-failure cleanup is a
//! single `arena.deinit()` call. `WorkbookXml.deinit` is the canonical
//! teardown path.

const std = @import("std");
const assert = std.debug.assert;

pub const SheetState = enum { visible, hidden, very_hidden };

pub const Sheet = struct {
    /// Raw `name=` attribute value. Borrowed from the caller's `xml`.
    /// Not entity-decoded.
    name: []const u8,
    /// Workbook-scoped sheet id from `sheetId=`. Excel uses these as
    /// stable identifiers across reorderings; the visible tab order is
    /// the order of `<sheet>` elements in the XML.
    sheet_id: u32,
    /// Relationship id (`r:id=`) pointing into the workbook rels.
    /// Borrowed from the caller's `xml`.
    r_id: []const u8,
    state: SheetState,
};

pub const DefinedName = struct {
    /// Raw `name=` attribute. Borrowed from the caller's `xml`.
    name: []const u8,
    /// Inner text of `<definedName>` — the formula. Borrowed from the
    /// caller's `xml` verbatim (still XML-escaped). Decode if you
    /// need to compare against a non-encoded reference.
    formula: []const u8,
    /// `localSheetId=` if present. Sheet-scoped name when set;
    /// workbook-scoped when null.
    local_sheet_id: ?u32,
    hidden: bool,
    /// Raw byte span of `formula` inside the xml slice handed to
    /// `parse` — `xml[body_start..body_end] == formula`. Recorded by
    /// the parser so a body-splicing editor patches exactly the
    /// element THIS entry was built from; reproducing the parser's
    /// indexing with a second scanner diverges on legal XML the
    /// parser tolerates (whitespace around `=`, commented-out
    /// elements, `>` inside quoted attribute values).
    body_start: usize,
    body_end: usize,
};

pub const CalcProperties = struct {
    calc_id: ?u32,
    full_calc_on_load: bool,
    iterate: bool,
    iterate_count: ?u32,
    iterate_delta: ?f64,
};

pub const Error = error{
    MalformedXml,
    MissingRoot,
    InvalidSheetId,
    InvalidLocalSheetId,
    InvalidCalcId,
    InvalidIterateCount,
    InvalidIterateDelta,
} || std.mem.Allocator.Error;

pub const WorkbookXml = struct {
    sheets: []Sheet,
    defined_names: []DefinedName,
    calc: CalcProperties,
    /// Internal arena backing `sheets` and `defined_names`. Always
    /// non-null on success. Callers must invoke `deinit` to free.
    arena: ?std.heap.ArenaAllocator,

    pub fn deinit(self: *WorkbookXml, allocator: std.mem.Allocator) void {
        // `allocator` is taken for symmetry with the rest of the
        // package layer (some sibling typed parts may not use an
        // arena and will free per-slice with the caller's allocator).
        // We currently route everything through the arena, so the
        // parameter is intentionally unused.
        _ = allocator;
        if (self.arena) |*a| a.deinit();
        self.* = undefined;
    }
};

/// Parse `xml` (raw bytes of `xl/workbook.xml`) into a typed view.
///
/// `allocator` is the long-lived allocator used to seed the internal
/// arena. The arena is reclaimed on the error path via `errdefer`.
///
/// Leaf string slices in the result borrow from `xml`; see the
/// module docstring for the full lifetime contract.
pub fn parse(allocator: std.mem.Allocator, xml: []const u8) Error!WorkbookXml {
    assert(@TypeOf(xml) == []const u8);
    var arena = std.heap.ArenaAllocator.init(allocator);
    errdefer arena.deinit();
    return parseWith(&arena, xml);
}

/// `parse` over a private copy of `xml`: the view borrows every name
/// and `r:id` from the bytes it was given, so a caller whose bytes are
/// transient (a patched `workbook.xml` not yet in the store — the
/// pre-write parse a transactional `addSheet` needs) parses this way
/// and the copy lives in the view's arena, freed by `deinit`.
pub fn parseOwning(allocator: std.mem.Allocator, xml: []const u8) Error!WorkbookXml {
    var arena = std.heap.ArenaAllocator.init(allocator);
    errdefer arena.deinit();
    const copy = try arena.allocator().dupe(u8, xml);
    return parseWith(&arena, copy);
}

/// The parse proper. `arena` is the caller's — on error the caller
/// frees it, on success the returned view takes it over by value.
fn parseWith(arena: *std.heap.ArenaAllocator, xml: []const u8) Error!WorkbookXml {
    // Empty input cannot describe a valid workbook root.
    if (xml.len == 0) return error.MissingRoot;
    const arena_alloc = arena.allocator();

    // Locate the `<workbook` root. We search for the prefix and
    // require the next character to be a tag-name boundary so that
    // hypothetical `<workbookFoo>` won't match.
    if (!hasTag(xml, "workbook")) return error.MissingRoot;

    var sheets: std.ArrayList(Sheet) = .empty;
    errdefer sheets.deinit(arena_alloc);

    var defined_names: std.ArrayList(DefinedName) = .empty;
    errdefer defined_names.deinit(arena_alloc);

    try collectSheets(arena_alloc, xml, &sheets);
    try collectDefinedNames(arena_alloc, xml, &defined_names);
    const calc = try parseCalcProperties(xml);

    const sheets_slice = try sheets.toOwnedSlice(arena_alloc);
    const defined_names_slice = try defined_names.toOwnedSlice(arena_alloc);

    // Postcondition: every sheet has a non-empty `r:id` (we rejected
    // entries lacking it during the scan, so this should hold).
    for (sheets_slice) |s| {
        assert(s.r_id.len > 0);
        assert(s.name.len > 0);
    }

    return .{
        .sheets = sheets_slice,
        .defined_names = defined_names_slice,
        .calc = calc,
        .arena = arena.*,
    };
}

// ─── Sheet collection ────────────────────────────────────────────────

fn collectSheets(
    allocator: std.mem.Allocator,
    xml: []const u8,
    out: *std.ArrayList(Sheet),
) Error!void {
    assert(xml.len > 0);
    assert(out.items.len == 0);

    var cursor: usize = 0;
    while (try findTagOpen(xml, cursor, "sheet")) |hit| {
        // `findTagOpen` returns positions for the opening `<` and the
        // first byte after `>` (or the self-close `/>`).
        cursor = hit.after_tag_close;

        const attrs = xml[hit.attrs_start..hit.attrs_end];
        const name = getAttr(attrs, "name") orelse continue;
        const r_id = getAttr(attrs, "r:id") orelse blk: {
            // Some workbook variants emit the unqualified `id=` —
            // accept it as a defensive fallback. Empty-string is
            // not a valid r:id.
            break :blk getAttr(attrs, "id") orelse continue;
        };
        if (name.len == 0) continue;
        if (r_id.len == 0) continue;

        const sheet_id_str = getAttr(attrs, "sheetId") orelse continue;
        const sheet_id = std.fmt.parseInt(u32, sheet_id_str, 10) catch
            return error.InvalidSheetId;

        const state: SheetState = if (getAttr(attrs, "state")) |s|
            parseSheetState(s) orelse .visible
        else
            .visible;

        try out.append(allocator, .{
            .name = name,
            .sheet_id = sheet_id,
            .r_id = r_id,
            .state = state,
        });
    }
}

fn parseSheetState(s: []const u8) ?SheetState {
    if (std.mem.eql(u8, s, "visible")) return .visible;
    if (std.mem.eql(u8, s, "hidden")) return .hidden;
    if (std.mem.eql(u8, s, "veryHidden")) return .very_hidden;
    return null;
}

// ─── Defined names ───────────────────────────────────────────────────

fn collectDefinedNames(
    allocator: std.mem.Allocator,
    xml: []const u8,
    out: *std.ArrayList(DefinedName),
) Error!void {
    assert(xml.len > 0);
    assert(out.items.len == 0);

    // Defined names live inside `<definedNames>...</definedNames>`.
    // We scan for the section to bound the work, then walk each
    // `<definedName ...>...</definedName>` element. A workbook
    // without defined names will simply leave `out` empty.
    const section_open = findSectionOpen(xml, "definedNames") orelse return;
    // `findSectionOpen` returns the byte after the opening tag's `>`.
    // Excel emits `<definedNames/>` (self-closing) when there are no
    // entries — in that shape there's no `</definedNames>` to find.
    // Detect by inspecting the two bytes preceding `section_open`.
    if (section_open >= 2 and xml[section_open - 2] == '/') return;
    const section_close = findSectionClose(xml, section_open, "definedNames") orelse
        return error.MalformedXml;
    assert(section_close >= section_open);

    const block = xml[section_open..section_close];
    var cursor: usize = 0;
    while (try findTagOpen(block, cursor, "definedName")) |hit| {
        // Self-closing `<definedName name="x"/>` has no formula;
        // skip those (defensive — Excel always emits a body).
        if (hit.self_closing) {
            cursor = hit.after_tag_close;
            continue;
        }
        const attrs = block[hit.attrs_start..hit.attrs_end];
        const name = getAttr(attrs, "name") orelse {
            cursor = hit.after_tag_close;
            continue;
        };
        if (name.len == 0) {
            cursor = hit.after_tag_close;
            continue;
        }

        // Locate the matching `</definedName>`, comment/CDATA-aware.
        // A raw `indexOfPos` stopped at a decoy close inside a
        // comment, which both truncated the recorded body AND left
        // the scan resuming mid-comment — a fake element in the same
        // comment then parsed as a real entry whose clean-looking
        // body was eligible for rewriting (Codex #188 r9).
        const body_start = hit.after_tag_close;
        const close_idx = (try findClosingTag(block, body_start, "</definedName>")) orelse
            return error.MalformedXml;
        const formula = block[body_start..close_idx];

        // Semantic scalars decode before interpretation: XML resolves
        // entity references in attribute values, so `hidden="&#49;"`
        // IS `hidden="1"` — comparing the raw spelling would read it
        // as false and the defined-names block re-emit would then
        // silently drop the flag.
        var lsid_buf: [32]u8 = undefined;
        const local_sheet_id: ?u32 = if (getAttr(attrs, "localSheetId")) |s| blk: {
            const dec = decodeScalarAttr(&lsid_buf, s) orelse return error.InvalidLocalSheetId;
            break :blk std.fmt.parseInt(u32, dec, 10) catch return error.InvalidLocalSheetId;
        } else null;

        var hid_buf: [32]u8 = undefined;
        const hidden = if (getAttr(attrs, "hidden")) |s|
            if (decodeScalarAttr(&hid_buf, s)) |dec| isXmlTrue(dec) else false
        else
            false;

        try out.append(allocator, .{
            .name = name,
            .formula = formula,
            .local_sheet_id = local_sheet_id,
            .hidden = hidden,
            // `body_start`/`close_idx` are block-relative; store the
            // absolute raw offsets.
            .body_start = section_open + body_start,
            .body_end = section_open + close_idx,
        });

        // Advance past `</definedName>`.
        cursor = close_idx + "</definedName>".len;
    }
}

/// Comment/CDATA/PI/DOCTYPE-aware search for a literal closing tag
/// (`</definedName>`, …). Returns the index of its `<`, or null when
/// the tag never occurs as real markup.
pub fn findClosingTag(xml: []const u8, from: usize, close_tag: []const u8) Error!?usize {
    assert(close_tag.len >= 3);
    assert(close_tag[0] == '<' and close_tag[1] == '/');
    var i = from;
    while (i < xml.len) {
        const lt = std.mem.indexOfScalarPos(u8, xml, i, '<') orelse return null;
        const skip_to = try skipNonElement(xml, lt);
        if (skip_to != lt) {
            i = skip_to;
            continue;
        }
        if (std.mem.startsWith(u8, xml[lt..], close_tag)) return lt;
        i = lt + 1;
    }
    return null;
}

fn isXmlTrue(s: []const u8) bool {
    // OOXML allows "1"/"true" interchangeably for boolean attrs.
    return std.mem.eql(u8, s, "1") or std.mem.eql(u8, s, "true");
}

/// Decode a SEMANTIC scalar attribute value (bool / integer) into
/// `buf`: the five named XML entities plus ASCII-range numeric
/// character references. Scalars are tiny and pure ASCII, so a
/// decoded byte outside 0..127 or an overflow of `buf` returns null
/// — the value was never a valid scalar. Unknown named entities pass
/// their `&` through verbatim, matching `store.decodeXmlEntities`'s
/// lenient contract.
pub fn decodeScalarAttr(buf: []u8, s: []const u8) ?[]const u8 {
    var n: usize = 0;
    var i: usize = 0;
    while (i < s.len) {
        var c: u8 = s[i];
        var consumed: usize = 1;
        if (s[i] == '&') {
            const rest = s[i..];
            if (std.mem.startsWith(u8, rest, "&amp;")) {
                c = '&';
                consumed = 5;
            } else if (std.mem.startsWith(u8, rest, "&lt;")) {
                c = '<';
                consumed = 4;
            } else if (std.mem.startsWith(u8, rest, "&gt;")) {
                c = '>';
                consumed = 4;
            } else if (std.mem.startsWith(u8, rest, "&quot;")) {
                c = '"';
                consumed = 6;
            } else if (std.mem.startsWith(u8, rest, "&apos;")) {
                c = '\'';
                consumed = 6;
            } else if (std.mem.startsWith(u8, rest, "&#")) {
                const semi = std.mem.indexOfScalarPos(u8, s, i + 2, ';') orelse return null;
                const digits = s[i + 2 .. semi];
                if (digits.len == 0) return null;
                // Validate the digit run before parseInt — parseInt
                // accepts `+` signs and `_` separators, which XML
                // numeric references forbid (`&#+49;` is malformed,
                // not "1"). The hex marker is lowercase `x` only
                // (Codex #215 r4 REL-405).
                const is_hex = digits[0] == 'x';
                const run = if (is_hex) digits[1..] else digits;
                if (run.len == 0) return null;
                for (run) |d| {
                    const ok = if (is_hex)
                        std.ascii.isHex(d)
                    else
                        d >= '0' and d <= '9';
                    if (!ok) return null;
                }
                const cp = std.fmt.parseInt(u32, run, if (is_hex) 16 else 10) catch return null;
                if (cp > 127) return null;
                c = @intCast(cp);
                consumed = semi - i + 1;
            }
        }
        if (n >= buf.len) return null;
        buf[n] = c;
        n += 1;
        i += consumed;
    }
    return buf[0..n];
}

// ─── Calc properties ─────────────────────────────────────────────────

fn parseCalcProperties(xml: []const u8) Error!CalcProperties {
    assert(xml.len > 0);

    var calc: CalcProperties = .{
        .calc_id = null,
        .full_calc_on_load = false,
        .iterate = false,
        .iterate_count = null,
        .iterate_delta = null,
    };

    const hit = (try findTagOpen(xml, 0, "calcPr")) orelse return calc;
    const attrs = xml[hit.attrs_start..hit.attrs_end];

    if (getAttr(attrs, "calcId")) |s| {
        calc.calc_id = std.fmt.parseInt(u32, s, 10) catch return error.InvalidCalcId;
    }
    if (getAttr(attrs, "fullCalcOnLoad")) |s| {
        calc.full_calc_on_load = isXmlTrue(s);
    }
    if (getAttr(attrs, "iterate")) |s| {
        calc.iterate = isXmlTrue(s);
    }
    if (getAttr(attrs, "iterateCount")) |s| {
        calc.iterate_count = std.fmt.parseInt(u32, s, 10) catch
            return error.InvalidIterateCount;
    }
    if (getAttr(attrs, "iterateDelta")) |s| {
        calc.iterate_delta = std.fmt.parseFloat(f64, s) catch
            return error.InvalidIterateDelta;
    }

    return calc;
}

// ─── XML scanner primitives ──────────────────────────────────────────

pub const TagHit = struct {
    /// Byte index of the opening `<`.
    open_lt: usize,
    /// Index of the first byte of the attributes region (just past the
    /// tag name).
    attrs_start: usize,
    /// Index one past the last attribute byte (i.e. the `>` or the
    /// `/` of a self-close).
    attrs_end: usize,
    /// First byte after the tag-closing `>` (regardless of whether
    /// the tag was self-closing).
    after_tag_close: usize,
    self_closing: bool,
};

/// Locate the next `<tag` whose name matches `tag` exactly, starting
/// from `from`. Skips XML comments, CDATA, processing instructions,
/// and DOCTYPE blocks. Returns `null` when no further match exists.
pub fn findTagOpen(xml: []const u8, from: usize, tag: []const u8) Error!?TagHit {
    assert(tag.len > 0);

    var i: usize = from;
    while (i < xml.len) {
        const lt = std.mem.indexOfScalarPos(u8, xml, i, '<') orelse return null;

        // Skip comments / CDATA / PIs / DOCTYPE before re-trying.
        const skip_to = try skipNonElement(xml, lt);
        if (skip_to != lt) {
            i = skip_to;
            continue;
        }

        // Bounds: need room for `<tag` plus the boundary char.
        if (lt + 1 + tag.len >= xml.len) return null;

        if (!std.mem.eql(u8, xml[lt + 1 .. lt + 1 + tag.len], tag)) {
            i = lt + 1;
            continue;
        }
        const boundary = xml[lt + 1 + tag.len];
        if (!isTagNameBoundary(boundary)) {
            i = lt + 1;
            continue;
        }

        // Found `<tag<boundary>`. Walk attributes to find the closing
        // `>` while respecting quoted attribute values.
        const attrs_start = lt + 1 + tag.len;
        const close = try findTagEnd(xml, attrs_start) orelse return error.MalformedXml;
        const self_closing = close.self_closing;

        return .{
            .open_lt = lt,
            .attrs_start = attrs_start,
            .attrs_end = close.attrs_end,
            .after_tag_close = close.after_gt,
            .self_closing = self_closing,
        };
    }
    return null;
}

/// Quick existence check — used by `parse` to validate a `<workbook`
/// root is present without committing to a full TagHit.
fn hasTag(xml: []const u8, tag: []const u8) bool {
    assert(tag.len > 0);
    assert(xml.len > 0);
    var i: usize = 0;
    while (std.mem.indexOfScalarPos(u8, xml, i, '<')) |lt| {
        const skip_to = skipNonElement(xml, lt) catch return false;
        if (skip_to != lt) {
            i = skip_to;
            continue;
        }
        if (lt + 1 + tag.len >= xml.len) return false;
        if (std.mem.eql(u8, xml[lt + 1 .. lt + 1 + tag.len], tag) and
            isTagNameBoundary(xml[lt + 1 + tag.len]))
        {
            return true;
        }
        i = lt + 1;
    }
    return false;
}

fn isTagNameBoundary(c: u8) bool {
    return c == ' ' or c == '\t' or c == '\n' or c == '\r' or c == '/' or c == '>';
}

/// If `xml[at]..` opens a comment / CDATA / PI / DOCTYPE, return the
/// index just past the construct. Otherwise return `at` unchanged.
pub fn skipNonElement(xml: []const u8, at: usize) Error!usize {
    assert(at < xml.len);
    assert(xml[at] == '<');

    if (at + 4 <= xml.len and std.mem.eql(u8, xml[at .. at + 4], "<!--")) {
        const end = std.mem.indexOfPos(u8, xml, at + 4, "-->") orelse
            return error.MalformedXml;
        return end + 3;
    }
    if (at + 9 <= xml.len and std.mem.eql(u8, xml[at .. at + 9], "<![CDATA[")) {
        const end = std.mem.indexOfPos(u8, xml, at + 9, "]]>") orelse
            return error.MalformedXml;
        return end + 3;
    }
    if (at + 2 <= xml.len and xml[at + 1] == '?') {
        const end = std.mem.indexOfPos(u8, xml, at + 2, "?>") orelse
            return error.MalformedXml;
        return end + 2;
    }
    if (at + 2 <= xml.len and xml[at + 1] == '!') {
        // `<!DOCTYPE …>` and friends. We consume up to the next bare
        // `>` outside any quoted region. DOCTYPE is rare in xlsx but
        // handling it keeps the scanner well-defined.
        const close = try findBareGt(xml, at + 2);
        return close + 1;
    }
    return at;
}

/// Walk forward from `attrs_start` to the tag-terminating `>` while
/// honouring quoted attribute values. Returns the position of the
/// terminator and whether the tag was self-closing.
const TagEnd = struct { attrs_end: usize, after_gt: usize, self_closing: bool };

fn findTagEnd(xml: []const u8, attrs_start: usize) Error!?TagEnd {
    var i: usize = attrs_start;
    while (i < xml.len) {
        const c = xml[i];
        if (c == '"' or c == '\'') {
            const close = std.mem.indexOfScalarPos(u8, xml, i + 1, c) orelse
                return error.MalformedXml;
            i = close + 1;
            continue;
        }
        if (c == '>') {
            const self_closing = i > attrs_start and xml[i - 1] == '/';
            const attrs_end = if (self_closing) i - 1 else i;
            return .{
                .attrs_end = attrs_end,
                .after_gt = i + 1,
                .self_closing = self_closing,
            };
        }
        i += 1;
    }
    return null;
}

fn findBareGt(xml: []const u8, from: usize) Error!usize {
    var i: usize = from;
    while (i < xml.len) {
        const c = xml[i];
        if (c == '"' or c == '\'') {
            const close = std.mem.indexOfScalarPos(u8, xml, i + 1, c) orelse
                return error.MalformedXml;
            i = close + 1;
            continue;
        }
        if (c == '>') return i;
        i += 1;
    }
    return error.MalformedXml;
}

/// Find the byte index just past `<section>` (the opening element of
/// a wrapper section like `<definedNames>` or `<sheets>`). Returns
/// `null` when the section is absent.
fn findSectionOpen(xml: []const u8, section: []const u8) ?usize {
    assert(section.len > 0);
    var i: usize = 0;
    while (std.mem.indexOfScalarPos(u8, xml, i, '<')) |lt| {
        const skip_to = skipNonElement(xml, lt) catch return null;
        if (skip_to != lt) {
            i = skip_to;
            continue;
        }
        if (lt + 1 + section.len >= xml.len) return null;
        if (std.mem.eql(u8, xml[lt + 1 .. lt + 1 + section.len], section) and
            isTagNameBoundary(xml[lt + 1 + section.len]))
        {
            const gt = std.mem.indexOfScalarPos(u8, xml, lt, '>') orelse return null;
            return gt + 1;
        }
        i = lt + 1;
    }
    return null;
}

/// Find the start of `</section>` at or after `from`. Used to bound a
/// wrapper section's contents. Comment/CDATA/PI-aware like the rest
/// of the scanner: a raw search truncated the section at a decoy
/// close inside a comment, handing the aware entry walk a block that
/// ends MID-COMMENT — which it then rejected as malformed, refusing
/// a valid file (Codex #188 r10).
fn findSectionClose(xml: []const u8, from: usize, section: []const u8) ?usize {
    assert(section.len > 0);
    // Allocate a small stack buffer for `</section>`. Section names in
    // workbook.xml are short (`definedNames` is the longest we care
    // about at 12 bytes) so a 64-byte cap is generous.
    var buf: [64]u8 = undefined;
    if (section.len + 3 > buf.len) return null;
    buf[0] = '<';
    buf[1] = '/';
    @memcpy(buf[2 .. 2 + section.len], section);
    buf[2 + section.len] = '>';
    const needle = buf[0 .. 3 + section.len];
    return findClosingTag(xml, from, needle) catch null;
}

// ─── Attribute extraction ────────────────────────────────────────────

/// Pull a quoted attribute value out of an attributes region. Mirrors
/// the helper in `src/xlsx.zig` (kept private here so this file is
/// self-contained per the project's per-typed-part isolation rule).
/// Values are returned verbatim — no entity decoding.
pub fn getAttr(attrs: []const u8, name: []const u8) ?[]const u8 {
    assert(name.len > 0);
    var i: usize = 0;
    while (i < attrs.len) {
        while (i < attrs.len and std.ascii.isWhitespace(attrs[i])) i += 1;
        if (i >= attrs.len) break;

        const name_start = i;
        while (i < attrs.len and attrs[i] != '=' and !std.ascii.isWhitespace(attrs[i])) i += 1;
        const attr_name = attrs[name_start..i];

        while (i < attrs.len and (attrs[i] == '=' or std.ascii.isWhitespace(attrs[i]))) i += 1;
        if (i >= attrs.len) break;
        if (attrs[i] != '"' and attrs[i] != '\'') break;

        const quote = attrs[i];
        i += 1;
        const val_start = i;
        while (i < attrs.len and attrs[i] != quote) i += 1;
        const val = attrs[val_start..i];
        if (i < attrs.len) i += 1;

        if (std.mem.eql(u8, attr_name, name)) return val;
    }
    return null;
}

// ─── Tests ───────────────────────────────────────────────────────────

test "parse: minimal one-sheet workbook" {
    const xml =
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        \\  <sheets>
        \\    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
        \\  </sheets>
        \\</workbook>
    ;

    var wb = try parse(std.testing.allocator, xml);
    defer wb.deinit(std.testing.allocator);

    try std.testing.expectEqual(@as(usize, 1), wb.sheets.len);
    try std.testing.expectEqualStrings("Sheet1", wb.sheets[0].name);
    try std.testing.expectEqual(@as(u32, 1), wb.sheets[0].sheet_id);
    try std.testing.expectEqualStrings("rId1", wb.sheets[0].r_id);
    try std.testing.expectEqual(SheetState.visible, wb.sheets[0].state);
    try std.testing.expectEqual(@as(usize, 0), wb.defined_names.len);
}

test "parse: multi-sheet with defined names + calcPr" {
    const xml =
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        \\  <sheets>
        \\    <sheet name="Data" sheetId="1" r:id="rId1"/>
        \\    <sheet name="Summary" sheetId="2" r:id="rId2"/>
        \\    <sheet name="Notes" sheetId="3" r:id="rId3"/>
        \\  </sheets>
        \\  <definedNames>
        \\    <definedName name="Range1">Data!$A$1:$B$10</definedName>
        \\    <definedName name="Local" localSheetId="1" hidden="1">Summary!$C$5</definedName>
        \\  </definedNames>
        \\  <calcPr calcId="191029" iterate="true" iterateCount="50" iterateDelta="0.0001" fullCalcOnLoad="1"/>
        \\</workbook>
    ;

    var wb = try parse(std.testing.allocator, xml);
    defer wb.deinit(std.testing.allocator);

    try std.testing.expectEqual(@as(usize, 3), wb.sheets.len);
    try std.testing.expectEqualStrings("Data", wb.sheets[0].name);
    try std.testing.expectEqualStrings("Summary", wb.sheets[1].name);
    try std.testing.expectEqualStrings("Notes", wb.sheets[2].name);
    try std.testing.expectEqual(@as(u32, 1), wb.sheets[0].sheet_id);
    try std.testing.expectEqual(@as(u32, 2), wb.sheets[1].sheet_id);
    try std.testing.expectEqual(@as(u32, 3), wb.sheets[2].sheet_id);

    try std.testing.expectEqual(@as(usize, 2), wb.defined_names.len);
    try std.testing.expectEqualStrings("Range1", wb.defined_names[0].name);
    try std.testing.expectEqualStrings("Data!$A$1:$B$10", wb.defined_names[0].formula);
    try std.testing.expectEqual(@as(?u32, null), wb.defined_names[0].local_sheet_id);
    try std.testing.expectEqual(false, wb.defined_names[0].hidden);

    try std.testing.expectEqualStrings("Local", wb.defined_names[1].name);
    try std.testing.expectEqualStrings("Summary!$C$5", wb.defined_names[1].formula);
    try std.testing.expectEqual(@as(?u32, 1), wb.defined_names[1].local_sheet_id);
    try std.testing.expectEqual(true, wb.defined_names[1].hidden);

    try std.testing.expectEqual(@as(?u32, 191029), wb.calc.calc_id);
    try std.testing.expectEqual(true, wb.calc.iterate);
    try std.testing.expectEqual(@as(?u32, 50), wb.calc.iterate_count);
    try std.testing.expect(wb.calc.iterate_delta != null);
    try std.testing.expect(@abs(wb.calc.iterate_delta.? - 0.0001) < 1e-12);
    try std.testing.expectEqual(true, wb.calc.full_calc_on_load);
}

test "parse: hidden and very-hidden sheet states" {
    const xml =
        \\<?xml version="1.0"?>
        \\<workbook>
        \\  <sheets>
        \\    <sheet name="A" sheetId="1" state="visible" r:id="rId1"/>
        \\    <sheet name="B" sheetId="2" state="hidden" r:id="rId2"/>
        \\    <sheet name="C" sheetId="3" state="veryHidden" r:id="rId3"/>
        \\    <sheet name="D" sheetId="4" state="bogus" r:id="rId4"/>
        \\  </sheets>
        \\</workbook>
    ;

    var wb = try parse(std.testing.allocator, xml);
    defer wb.deinit(std.testing.allocator);

    try std.testing.expectEqual(@as(usize, 4), wb.sheets.len);
    try std.testing.expectEqual(SheetState.visible, wb.sheets[0].state);
    try std.testing.expectEqual(SheetState.hidden, wb.sheets[1].state);
    try std.testing.expectEqual(SheetState.very_hidden, wb.sheets[2].state);
    // Unknown state value falls back to visible (defensive default).
    try std.testing.expectEqual(SheetState.visible, wb.sheets[3].state);
}

test "parse: rejects empty input and missing root" {
    try std.testing.expectError(error.MissingRoot, parse(std.testing.allocator, ""));
    try std.testing.expectError(
        error.MissingRoot,
        parse(std.testing.allocator, "<notWorkbook/>"),
    );
}

test "decodeScalarAttr: decodes entities, rejects malformed numeric refs" {
    var buf: [32]u8 = undefined;
    // Named + numeric forms of "1".
    try std.testing.expectEqualStrings("1", decodeScalarAttr(&buf, "1").?);
    try std.testing.expectEqualStrings("1", decodeScalarAttr(&buf, "&#49;").?);
    try std.testing.expectEqualStrings("1", decodeScalarAttr(&buf, "&#x31;").?);
    try std.testing.expectEqualStrings("true", decodeScalarAttr(&buf, "true").?);
    // parseInt would accept these; XML numeric references forbid
    // signs and separators (Codex #188 r5 finding 3).
    try std.testing.expectEqual(@as(?[]const u8, null), decodeScalarAttr(&buf, "&#+49;"));
    try std.testing.expectEqual(@as(?[]const u8, null), decodeScalarAttr(&buf, "&#4_9;"));
    try std.testing.expectEqual(@as(?[]const u8, null), decodeScalarAttr(&buf, "&#x+31;"));
    // Non-ASCII code points are never valid scalars.
    try std.testing.expectEqual(@as(?[]const u8, null), decodeScalarAttr(&buf, "&#955;"));
}

test "parse: rejects invalid sheetId" {
    const xml =
        \\<workbook>
        \\  <sheets>
        \\    <sheet name="A" sheetId="not-a-number" r:id="rId1"/>
        \\  </sheets>
        \\</workbook>
    ;
    try std.testing.expectError(error.InvalidSheetId, parse(std.testing.allocator, xml));
}

test "parse: skips comments, CDATA, and processing instructions" {
    const xml =
        \\<?xml version="1.0"?>
        \\<!-- a stray comment with <sheet name="GHOST" sheetId="99" r:id="rIdX"/> inside -->
        \\<workbook>
        \\  <sheets>
        \\    <!-- comment between sheets -->
        \\    <sheet name="Real" sheetId="1" r:id="rId1"/>
        \\  </sheets>
        \\  <definedNames>
        \\    <definedName name="WithCdata"><![CDATA[NOT_A_FORMULA]]></definedName>
        \\  </definedNames>
        \\</workbook>
    ;

    var wb = try parse(std.testing.allocator, xml);
    defer wb.deinit(std.testing.allocator);

    try std.testing.expectEqual(@as(usize, 1), wb.sheets.len);
    try std.testing.expectEqualStrings("Real", wb.sheets[0].name);
    try std.testing.expectEqual(@as(usize, 1), wb.defined_names.len);
    try std.testing.expectEqualStrings("WithCdata", wb.defined_names[0].name);
}

test "parse: handles attribute values containing '>'" {
    // Quoted attribute values may contain `>` legally. The scanner
    // must not stop at the bare `>` inside the quotes.
    const xml =
        \\<workbook>
        \\  <sheets>
        \\    <sheet name="Has &gt; Inside" sheetId="1" r:id="rId1"/>
        \\  </sheets>
        \\</workbook>
    ;

    var wb = try parse(std.testing.allocator, xml);
    defer wb.deinit(std.testing.allocator);

    try std.testing.expectEqual(@as(usize, 1), wb.sheets.len);
    try std.testing.expectEqualStrings("Has &gt; Inside", wb.sheets[0].name);
}

test "parse: omitted calcPr leaves all fields default" {
    const xml =
        \\<workbook>
        \\  <sheets>
        \\    <sheet name="A" sheetId="1" r:id="rId1"/>
        \\  </sheets>
        \\</workbook>
    ;

    var wb = try parse(std.testing.allocator, xml);
    defer wb.deinit(std.testing.allocator);

    try std.testing.expectEqual(@as(?u32, null), wb.calc.calc_id);
    try std.testing.expectEqual(false, wb.calc.full_calc_on_load);
    try std.testing.expectEqual(false, wb.calc.iterate);
    try std.testing.expectEqual(@as(?u32, null), wb.calc.iterate_count);
    try std.testing.expectEqual(@as(?f64, null), wb.calc.iterate_delta);
}

test "parse: tolerates `<workbookPr>`-style siblings without confusing them with `<workbook>`" {
    // Regression guard: scanner must require the tag-name boundary
    // before accepting a match — otherwise `<workbook…>` would also
    // match against `workbookPr` etc.
    const xml =
        \\<workbook>
        \\  <workbookPr date1904="1"/>
        \\  <workbookProtection lockStructure="1"/>
        \\  <sheets>
        \\    <sheet name="One" sheetId="42" r:id="rId7"/>
        \\  </sheets>
        \\</workbook>
    ;

    var wb = try parse(std.testing.allocator, xml);
    defer wb.deinit(std.testing.allocator);

    try std.testing.expectEqual(@as(usize, 1), wb.sheets.len);
    try std.testing.expectEqualStrings("One", wb.sheets[0].name);
    try std.testing.expectEqual(@as(u32, 42), wb.sheets[0].sheet_id);
}
