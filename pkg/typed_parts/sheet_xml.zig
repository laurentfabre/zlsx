//! Typed-overlay parser for one `xl/worksheets/sheet*.xml` part.
//!
//! Lifetime contract: every `[]const u8` field on a returned struct
//! borrows from the input `xml` slice handed to `parse`. The caller
//! owns `xml` and must keep it alive for as long as the `SheetXml` is
//! used. The arena attached to `SheetXml.arena` only owns the spine
//! allocations (the `rows`, `cells`, `merges`, `hyperlinks`,
//! `validations`, `conditional_formats` slices), never their text
//! content.
//!
//! Defensive XML scanning here mirrors `pkg/store.zig`: comments,
//! CDATA, and processing-instruction blocks are skipped before any
//! tag match, and attribute-value scanning honours both `"` and `'`
//! quoting so we don't trip on a literal `>` inside an attribute.

const std = @import("std");
const assert = std.debug.assert;
const wbxml = @import("workbook_xml.zig");

// ─── Public types ────────────────────────────────────────────────────

pub const Dimension = struct {
    /// Borrows from the input xml. Examples: "A1", "A1:Z100".
    ref: []const u8,
};

pub const CellType = enum {
    number,
    shared_string,
    boolean,
    formula_string, // OOXML `t="str"` — formula returns a string.
    inline_string, // OOXML `t="inlineStr"` — `<is><t>…</t></is>` body.
    error_value,
    date,
};

pub const Cell = struct {
    /// Borrows. The `<c r="…">` reference, e.g. "A1", "B12".
    ref: []const u8,
    /// `<c s="…">` style index into the workbook's cellXfs table.
    style_idx: ?u32,
    /// `s` was written but is not a number (`s="bogus"`, `s=""`):
    /// `style_idx` is null, and the cell is not unstyled. A reader
    /// that keys on the style (the S7b-4 rebuild's date check) refuses
    /// rather than read it as General (Codex #205 r5 REL-504).
    style_invalid: bool = false,
    cell_type: CellType,
    /// `t` was written but names no type this view knows: `cell_type`
    /// is `.number` by the schema's default, and the cell is not one
    /// the S7b-4 rebuild reads (Codex #205 r8 REL-801).
    cell_type_invalid: bool = false,
    /// Raw inner text of `<v>` (or `<is><t>` for inline strings).
    /// Borrows. No XML-entity decoding — caller decodes if needed.
    raw_value: ?[]const u8,
    /// Raw inner text of `<f>`. Borrows. Not rewritten.
    formula: ?[]const u8,
    /// An inline string's whole `<is>` body, raw — its runs, where
    /// `raw_value` is the first `<t>` alone. Borrows. Null for every
    /// other cell.
    inline_body: ?[]const u8 = null,
    /// The `<is>` was not the cell's one direct child — nested under
    /// another element, or spelt twice: a reader that carries the
    /// string refuses rather than take a decoy (Codex #206 r23
    /// REL-2302).
    inline_invalid: bool = false,
};

pub const Row = struct {
    row_idx: u32, // 1-based, matches OOXML `<row r="…">`.
    cells: []Cell,
    height: ?f64,
    custom_height: bool,
    hidden: bool,
    /// `<c>` elements of this row without an `r` — not in `cells`,
    /// since the view has no coordinate for them. A reader that needs
    /// the row whole (the S7b-4 rebuild) refuses when this is not 0.
    unaddressed_cells: u32 = 0,
};

pub const MergeRange = struct {
    /// Borrows. e.g. "A1:B2".
    ref: []const u8,
};

pub const Hyperlink = struct {
    ref: []const u8,
    r_id: ?[]const u8,
    location: ?[]const u8,
    display: ?[]const u8,
    tooltip: ?[]const u8,
};

pub const DataValidation = struct {
    sqref: []const u8,
    type: []const u8,
    formula1: ?[]const u8,
    formula2: ?[]const u8,
    operator: ?[]const u8,
    allow_blank: bool,
    show_dropdown: bool,
};

pub const ConditionalFormat = struct {
    sqref: []const u8,
    type: []const u8,
    formula: ?[]const u8,
    /// CT_CfRule allows up to three `<formula>` children — a `cellIs`
    /// `between` rule carries two bounds. `formula` is the first body
    /// (the one the DV/CF formula rewriter locates and splices);
    /// these are the second and third in document order, null when
    /// the rule spells fewer.
    formula2: ?[]const u8 = null,
    formula3: ?[]const u8 = null,
    dxf_id: ?u32,
    priority: ?u32,
};

pub const FreezePane = struct {
    x_split: u32,
    y_split: u32,
    /// Borrows. The `topLeftCell` attribute, e.g. "B2".
    top_left_cell: []const u8,
};

/// One contiguity run of the sanitized buffer: `len` bytes starting
/// at `san_start` were copied VERBATIM from the source part at
/// `src_start`. Runs break wherever the sanitizer elides (comments,
/// PIs) or re-encodes (CDATA contents) — those sanitized bytes belong
/// to no run.
pub const SrcRun = struct { san_start: u32, src_start: u32, len: u32 };

pub const SheetXml = struct {
    dimension: ?Dimension,
    rows: []Row,
    merges: []MergeRange,
    hyperlinks: []Hyperlink,
    validations: []DataValidation,
    conditional_formats: []ConditionalFormat,
    freeze: ?FreezePane,
    /// The sanitized buffer every view slice borrows from, plus the
    /// sanitized→source contiguity runs. The DV/CF formula rewriter
    /// locates its splice targets by MAPPING a view slice's offset
    /// back to the source part (`sourceSpanOf`) — four Codex #215
    /// rounds (r1 REL-103, r2 REL-204, r4 REL-401, r5 REL-501) showed
    /// that re-scanning the raw source "in lockstep" cannot reproduce
    /// the sanitizer's lexical state; the map ends the class.
    sanitized: []const u8 = "",
    src_runs: []const SrcRun = &.{},
    /// `<row>` elements without a usable `r` — not in `rows`, since the
    /// view has no coordinate for them (the Editor's own scanner numbers
    /// them implicitly). A reader that needs the grid whole refuses
    /// when this is not 0.
    unaddressed_rows: u32 = 0,
    /// Owns the spine slices. `null` only when no allocations were
    /// performed — `parse` always sets it. Kept optional so a
    /// caller-constructed empty `SheetXml` (e.g. in tests) can free
    /// trivially via `deinit`.
    arena: ?std.heap.ArenaAllocator,

    pub fn deinit(self: *SheetXml, allocator: std.mem.Allocator) void {
        _ = allocator; // arena owns everything; allocator parity for
        // sibling deinits in the typed_parts namespace.
        if (self.arena) |*a| {
            a.deinit();
            self.arena = null;
        }
        self.rows = &.{};
        self.merges = &.{};
        self.hyperlinks = &.{};
        self.validations = &.{};
        self.conditional_formats = &.{};
        self.sanitized = "";
        self.src_runs = &.{};
        self.dimension = null;
        self.freeze = null;
    }

    /// Map a slice BORROWED FROM THIS VIEW back to its span in the
    /// source part the view was parsed from. Returns null when the
    /// slice does not lie in the sanitized buffer, or does not sit
    /// wholly inside one verbatim contiguity run (its bytes crossed a
    /// stripped construct or were re-encoded from CDATA) — such a
    /// span has no byte-identical home in the source and cannot be
    /// spliced.
    pub fn sourceSpanOf(self: *const SheetXml, slice: []const u8) ?[2]usize {
        const base = @intFromPtr(self.sanitized.ptr);
        const p = @intFromPtr(slice.ptr);
        if (p < base or p - base + slice.len > self.sanitized.len) return null;
        const off = p - base;
        // Binary search: the last run with san_start <= off.
        var lo: usize = 0;
        var hi: usize = self.src_runs.len;
        while (lo < hi) {
            const mid = lo + (hi - lo) / 2;
            if (self.src_runs[mid].san_start <= off) lo = mid + 1 else hi = mid;
        }
        if (lo == 0) return null;
        const run = self.src_runs[lo - 1];
        if (off + slice.len > @as(usize, run.san_start) + run.len) return null;
        const src_lo = @as(usize, run.src_start) + (off - run.san_start);
        return .{ src_lo, src_lo + slice.len };
    }
};

pub const ParseError = error{
    MalformedXml,
    UnexpectedEof,
    OutOfMemory,
};

// ─── Public entry point ──────────────────────────────────────────────

/// Parse one `xl/worksheets/sheet*.xml` part into a typed overlay.
///
/// The returned `SheetXml` borrows every textual field from an
/// arena-owned, comment/CDATA/PI-stripped copy of `xml`. The caller
/// does NOT need to keep the original `xml` alive after `parse`
/// returns. The `SheetXml` owns one arena (the sanitized buffer +
/// slice spines); call `deinit` on success or error to reclaim it.
///
/// The one-time sanitizer pass is a perf bound: per-tag in-comment
/// checks were O(N²) on large worksheets (1+ GiB sheet bodies on
/// real-world fixtures hung CI for >1h). The sanitizer is O(N), and
/// the rest of the parser sees no comment markers.
pub fn parse(allocator: std.mem.Allocator, xml: []const u8) ParseError!SheetXml {
    assert(xml.len < (1 << 31)); // OOXML sheet parts are bounded; refuse 2 GiB+.
    assert(@TypeOf(allocator) == std.mem.Allocator);

    var arena = std.heap.ArenaAllocator.init(allocator);
    errdefer arena.deinit();
    const a = arena.allocator();

    // Strip comments / CDATA / PI once, up front. Every downstream
    // scan operates on the sanitized buffer; no per-tag comment check
    // is needed (and `insideComment` was the O(N²) hot-path on real
    // worksheets — see doc-comment above). The contiguity runs come
    // along so a view slice can be mapped back to its source span.
    var src_runs: std.ArrayListUnmanaged(SrcRun) = .empty;
    const sanitized = try sanitizeXml(a, xml, &src_runs);

    // Quick well-formedness gate: a worksheet part must contain a
    // `<worksheet` root tag. Anything else is rejected up front so
    // callers don't get an empty struct from a styles or sst part
    // wired into the wrong slot.
    if (std.mem.indexOf(u8, sanitized, "<worksheet") == null) return error.MalformedXml;

    const dim = parseDimension(sanitized);
    const freeze = parseFreezePane(sanitized);

    var unaddressed_rows: u32 = 0;
    const rows = try parseRows(a, sanitized, &unaddressed_rows);
    const merges = try parseMerges(a, sanitized);
    const hyperlinks = try parseHyperlinks(a, sanitized);
    const validations = try parseValidations(a, sanitized);
    const cfs = try parseConditionalFormats(a, sanitized);

    return .{
        .dimension = dim,
        .rows = rows,
        .merges = merges,
        .hyperlinks = hyperlinks,
        .validations = validations,
        .conditional_formats = cfs,
        .freeze = freeze,
        .unaddressed_rows = unaddressed_rows,
        .sanitized = sanitized,
        .src_runs = src_runs.items,
        .arena = arena,
    };
}

/// Produces a copy of `xml` with comments (`<!-- ... -->`), CDATA
/// (`<![CDATA[ ... ]]>`), and processing instructions (`<? ... ?>`)
/// elided. CDATA contents are CHARACTER DATA and are entity-escaped
/// into the copy so a `<![CDATA[<row/>]]>` payload can never be
/// misread as a real element — the verbatim copy this doc always
/// promised to prevent actually allowed it (Codex #188 r7). Returns
/// a fresh allocator-owned slice; caller owns the memory.
///
/// This is the perf-critical pre-pass: it converts every downstream
/// scan from "comment-aware" to "comment-free", eliminating the
/// O(N²) repeated `lastIndexOf("<!--")` over each tag-match.
/// `runs`, when non-null, receives the sanitized→source contiguity
/// map (`SrcRun`) — one entry per maximal verbatim-copied region.
/// CDATA contents are re-encoded, not copied, so they belong to no
/// run; a copy that resumes after any elision starts a new run.
fn sanitizeXml(
    allocator: std.mem.Allocator,
    xml: []const u8,
    runs: ?*std.ArrayListUnmanaged(SrcRun),
) ParseError![]const u8 {
    assert(xml.len < (1 << 31));

    var out: std.ArrayList(u8) = .empty;
    errdefer out.deinit(allocator);
    try out.ensureTotalCapacity(allocator, xml.len);

    var i: usize = 0;
    while (i < xml.len) {
        const c = xml[i];
        if (c != '<') {
            try mapRun(allocator, runs, out.items.len, i, 1);
            try out.append(allocator, c);
            i += 1;
            continue;
        }
        if (i + 4 <= xml.len and std.mem.eql(u8, xml[i .. i + 4], "<!--")) {
            const close = std.mem.indexOfPos(u8, xml, i + 4, "-->") orelse return error.MalformedXml;
            i = close + 3;
            continue;
        }
        if (i + 9 <= xml.len and std.mem.eql(u8, xml[i .. i + 9], "<![CDATA[")) {
            const close = std.mem.indexOfPos(u8, xml, i + 9, "]]>") orelse return error.MalformedXml;
            for (xml[i + 9 .. close]) |b| switch (b) {
                '&' => try out.appendSlice(allocator, "&amp;"),
                '<' => try out.appendSlice(allocator, "&lt;"),
                '>' => try out.appendSlice(allocator, "&gt;"),
                else => try out.append(allocator, b),
            };
            i = close + 3;
            continue;
        }
        if (i + 2 <= xml.len and xml[i + 1] == '?') {
            const close = std.mem.indexOfPos(u8, xml, i + 2, "?>") orelse return error.MalformedXml;
            i = close + 2;
            continue;
        }
        // Plain tag — copy through to the matching `>`, respecting
        // quoted attribute values (`>` inside `attr="..."` is data).
        const end = findTagEnd(xml, i) orelse return error.MalformedXml;
        try mapRun(allocator, runs, out.items.len, i, end + 1 - i);
        try out.appendSlice(allocator, xml[i .. end + 1]);
        i = end + 1;
    }

    return try out.toOwnedSlice(allocator);
}

/// Record `len` verbatim bytes copied from source offset `src` to
/// sanitized offset `san` — extending the last run when both sides
/// are contiguous, else opening a new one.
fn mapRun(
    allocator: std.mem.Allocator,
    runs: ?*std.ArrayListUnmanaged(SrcRun),
    san: usize,
    src: usize,
    len: usize,
) ParseError!void {
    const list = runs orelse return;
    if (list.items.len > 0) {
        const last = &list.items[list.items.len - 1];
        if (@as(usize, last.san_start) + last.len == san and
            @as(usize, last.src_start) + last.len == src)
        {
            last.len += @intCast(len);
            return;
        }
    }
    try list.append(allocator, .{
        .san_start = @intCast(san),
        .src_start = @intCast(src),
        .len = @intCast(len),
    });
}

// ─── XML scanner primitives ──────────────────────────────────────────

/// Find the next byte position after any leading comment / CDATA /
/// processing-instruction starting at `pos`. Returns the same `pos`
/// when nothing skippable is at the cursor. Returns `null` only when
/// an opener was seen but its terminator is missing — that's a
/// hard error which callers translate into `error.MalformedXml`.
fn skipBoilerplate(xml: []const u8, pos: usize) ?usize {
    assert(pos <= xml.len);
    var i = pos;
    while (i + 4 <= xml.len) {
        const remain = xml[i..];
        if (std.mem.startsWith(u8, remain, "<!--")) {
            const end = std.mem.indexOfPos(u8, xml, i + 4, "-->") orelse return null;
            i = end + 3;
            continue;
        }
        if (std.mem.startsWith(u8, remain, "<![CDATA[")) {
            const end = std.mem.indexOfPos(u8, xml, i + 9, "]]>") orelse return null;
            i = end + 3;
            continue;
        }
        if (std.mem.startsWith(u8, remain, "<?")) {
            const end = std.mem.indexOfPos(u8, xml, i + 2, "?>") orelse return null;
            i = end + 2;
            continue;
        }
        break;
    }
    return i;
}

/// Find the closing `>` of a tag whose opener begins at `tag_open`
/// (inclusive, the byte equal to `<`). Honours quoted attributes —
/// a literal `>` inside `attr="..."` or `attr='...'` does NOT
/// terminate the tag. Returns the index of the closing `>`.
fn findTagEnd(xml: []const u8, tag_open: usize) ?usize {
    assert(tag_open < xml.len);
    assert(xml[tag_open] == '<');
    var i = tag_open + 1;
    var quote: ?u8 = null;
    while (i < xml.len) : (i += 1) {
        const c = xml[i];
        if (quote) |q| {
            if (c == q) quote = null;
        } else if (c == '"' or c == '\'') {
            quote = c;
        } else if (c == '>') {
            return i;
        }
    }
    return null;
}

/// Locate the substring `needle` inside `hay`, starting at `from`,
/// Returns the next position of `needle` in `hay` at or after `from`.
/// Operates on sanitized input (comments / CDATA / PI already stripped
/// by `sanitizeXml`), so no per-tag in-comment check is needed —
/// previously the bottleneck on large real-world sheets.
fn indexOfTag(hay: []const u8, from: usize, needle: []const u8) ?usize {
    assert(needle.len > 0);
    assert(from <= hay.len);
    return std.mem.indexOfPos(u8, hay, from, needle);
}

/// Match `key="value"` or `key='value'` inside an attribute slice.
/// Returns the value (no quote-stripping issues, no entity decode).
fn attrAt(attrs: []const u8, key: []const u8) ?[]const u8 {
    return attrAtQuote(attrs, key, '"') orelse attrAtQuote(attrs, key, '\'');
}

fn attrAtQuote(attrs: []const u8, key: []const u8, quote: u8) ?[]const u8 {
    assert(key.len > 0);
    assert(key.len < 64);
    var buf: [80]u8 = undefined;
    if (key.len + 2 > buf.len) return null;
    @memcpy(buf[0..key.len], key);
    buf[key.len] = '=';
    buf[key.len + 1] = quote;
    const needle = buf[0 .. key.len + 2];

    // Scan with a left-boundary check so `key` doesn't match the
    // tail of another attribute name (e.g. searching `id` shouldn't
    // hit `r:id` or `xr:id`). Boundary chars are space/tab/lf/cr or
    // start-of-slice.
    var probe: usize = 0;
    while (std.mem.indexOfPos(u8, attrs, probe, needle)) |hit| {
        const left_ok = hit == 0 or isAttrBoundary(attrs[hit - 1]);
        if (!left_ok) {
            probe = hit + 1;
            continue;
        }
        const start = hit + needle.len;
        const close = std.mem.indexOfScalarPos(u8, attrs, start, quote) orelse return null;
        return attrs[start..close];
    }
    return null;
}

fn isAttrBoundary(c: u8) bool {
    return c == ' ' or c == '\t' or c == '\n' or c == '\r';
}

// ─── <dimension> ─────────────────────────────────────────────────────

fn parseDimension(xml: []const u8) ?Dimension {
    const open = indexOfTag(xml, 0, "<dimension") orelse return null;
    const end = findTagEnd(xml, open) orelse return null;
    const attrs = xml[open..end];
    const ref = attrAt(attrs, "ref") orelse return null;
    if (ref.len == 0) return null;
    return .{ .ref = ref };
}

// ─── <sheetView><pane …/></sheetView> ────────────────────────────────

fn parseFreezePane(xml: []const u8) ?FreezePane {
    // The sheet may carry multiple `<sheetView>` blocks (one per
    // window pane in older OOXML); freeze info lives on `<pane …
    // state="frozen">`. We accept any pane with `state="frozen"` or
    // `state="frozenSplit"`; non-frozen split panes carry no freeze
    // semantics so we ignore them.
    var probe: usize = 0;
    while (indexOfTag(xml, probe, "<pane")) |open| {
        const end = findTagEnd(xml, open) orelse return null;
        // Reject `<paneN>`-style hits (the tag must end at byte 5
        // with whitespace or `/` or `>`).
        const after = open + "<pane".len;
        assert(after <= xml.len);
        if (after >= xml.len) return null;
        const sep = xml[after];
        if (sep != ' ' and sep != '\t' and sep != '\n' and sep != '\r' and sep != '/' and sep != '>') {
            probe = after;
            continue;
        }
        const attrs = xml[open..end];

        const state = attrAt(attrs, "state") orelse {
            probe = end + 1;
            continue;
        };
        const is_frozen = std.mem.eql(u8, state, "frozen") or
            std.mem.eql(u8, state, "frozenSplit");
        if (!is_frozen) {
            probe = end + 1;
            continue;
        }

        const x_split = parseU32Attr(attrs, "xSplit") orelse 0;
        const y_split = parseU32Attr(attrs, "ySplit") orelse 0;
        const top_left = attrAt(attrs, "topLeftCell") orelse "";

        return .{
            .x_split = x_split,
            .y_split = y_split,
            .top_left_cell = top_left,
        };
    }
    return null;
}

/// An attribute's text with its character references resolved, when
/// it has any: the value the schema types, not its spelling
/// (`r="B&#50;"` is `B2` — Codex #205 r10 REL-1002). A reference that
/// does not resolve to a short ASCII scalar leaves the text as written,
/// for its parser to refuse. Arena-owned when decoded.
fn scalarAttr(a: std.mem.Allocator, attrs: []const u8, key: []const u8) ParseError!?[]const u8 {
    const raw = attrAt(attrs, key) orelse return null;
    if (std.mem.indexOfScalar(u8, raw, '&') == null) return raw;
    var buf: [64]u8 = undefined;
    const decoded = wbxml.decodeScalarAttr(&buf, raw) orelse return raw;
    return try a.dupe(u8, decoded);
}

fn parseU32Attr(attrs: []const u8, key: []const u8) ?u32 {
    const raw = attrAt(attrs, key) orelse return null;
    if (raw.len == 0) return null;
    return std.fmt.parseInt(u32, raw, 10) catch null;
}

fn parseF64Attr(attrs: []const u8, key: []const u8) ?f64 {
    const raw = attrAt(attrs, key) orelse return null;
    if (raw.len == 0) return null;
    return std.fmt.parseFloat(f64, raw) catch null;
}

fn parseBoolAttr(attrs: []const u8, key: []const u8) bool {
    const raw = attrAt(attrs, key) orelse return false;
    // OOXML `xsd:boolean`: "true" / "1" → true, "false" / "0" → false.
    if (raw.len == 0) return false;
    if (std.mem.eql(u8, raw, "1")) return true;
    if (std.mem.eql(u8, raw, "true")) return true;
    return false;
}

// ─── <sheetData><row><c>… ────────────────────────────────────────────

fn parseRows(a: std.mem.Allocator, xml: []const u8, unaddressed: *u32) ParseError![]Row {
    assert(xml.len > 0);

    const sd_start = indexOfTag(xml, 0, "<sheetData") orelse return &.{};
    const sd_open_end = findTagEnd(xml, sd_start) orelse return error.MalformedXml;
    // Self-closing <sheetData/> has no body.
    if (sd_open_end > 0 and xml[sd_open_end - 1] == '/') return &.{};
    const sd_close = std.mem.indexOfPos(u8, xml, sd_open_end, "</sheetData>") orelse
        return error.MalformedXml;
    const body = xml[sd_open_end + 1 .. sd_close];

    var rows: std.ArrayListUnmanaged(Row) = .empty;
    errdefer rows.deinit(a);

    var probe: usize = 0;
    while (indexOfTag(body, probe, "<row")) |row_open| {
        const after = row_open + "<row".len;
        if (after >= body.len) return error.UnexpectedEof;
        const sep = body[after];
        if (sep != ' ' and sep != '\t' and sep != '\n' and sep != '\r' and sep != '/' and sep != '>') {
            probe = after;
            continue;
        }
        const row_open_end = findTagEnd(body, row_open) orelse return error.MalformedXml;
        const row_attrs = body[row_open + "<row".len .. row_open_end];

        const row_idx_raw = (try scalarAttr(a, row_attrs, "r")) orelse {
            unaddressed.* +|= 1;
            probe = row_open_end + 1;
            continue;
        };
        const row_idx = std.fmt.parseInt(u32, row_idx_raw, 10) catch {
            unaddressed.* +|= 1;
            probe = row_open_end + 1;
            continue;
        };
        if (row_idx == 0) {
            unaddressed.* +|= 1;
            probe = row_open_end + 1;
            continue;
        }

        const height = parseF64Attr(row_attrs, "ht");
        const custom_height = parseBoolAttr(row_attrs, "customHeight");
        const hidden = parseBoolAttr(row_attrs, "hidden");

        // Self-closing `<row r="N"/>` — empty row.
        const self_closing = row_open_end > row_open and body[row_open_end - 1] == '/';
        var cells: []Cell = &.{};
        var unaddressed_cells: u32 = 0;
        if (!self_closing) {
            const row_close = std.mem.indexOfPos(u8, body, row_open_end, "</row>") orelse
                return error.MalformedXml;
            const row_body = body[row_open_end + 1 .. row_close];
            cells = try parseCells(a, row_body, &unaddressed_cells);
            probe = row_close + "</row>".len;
        } else {
            probe = row_open_end + 1;
        }

        try rows.append(a, .{
            .row_idx = row_idx,
            .cells = cells,
            .height = height,
            .custom_height = custom_height,
            .hidden = hidden,
            .unaddressed_cells = unaddressed_cells,
        });
    }

    return try rows.toOwnedSlice(a);
}

fn parseCells(a: std.mem.Allocator, row_body: []const u8, unaddressed: *u32) ParseError![]Cell {
    assert(row_body.len < (1 << 31));

    var cells: std.ArrayListUnmanaged(Cell) = .empty;
    errdefer cells.deinit(a);

    var probe: usize = 0;
    while (indexOfTag(row_body, probe, "<c")) |c_open| {
        const after = c_open + "<c".len;
        if (after >= row_body.len) return error.UnexpectedEof;
        const sep = row_body[after];
        if (sep != ' ' and sep != '\t' and sep != '\n' and sep != '\r' and sep != '/' and sep != '>') {
            probe = after;
            continue;
        }
        const c_open_end = findTagEnd(row_body, c_open) orelse return error.MalformedXml;
        const c_attrs = row_body[c_open + "<c".len .. c_open_end];

        const ref = (try scalarAttr(a, c_attrs, "r")) orelse {
            unaddressed.* +|= 1;
            probe = c_open_end + 1;
            continue;
        };
        const style_raw = try scalarAttr(a, c_attrs, "s");
        const style_idx: ?u32 = if (style_raw) |v| (if (v.len == 0) null else std.fmt.parseInt(u32, v, 10) catch null) else null;
        const style_invalid = style_raw != null and style_idx == null;
        const type_raw = try scalarAttr(a, c_attrs, "t");
        const cell_type_known = parseCellType(type_raw);
        const cell_type = cell_type_known orelse .number;
        const cell_type_invalid = type_raw != null and cell_type_known == null;

        const self_closing = c_open_end > c_open and row_body[c_open_end - 1] == '/';
        var raw_value: ?[]const u8 = null;
        var formula: ?[]const u8 = null;
        var inline_body: ?[]const u8 = null;
        var inline_invalid = false;
        if (!self_closing) {
            const c_close = std.mem.indexOfPos(u8, row_body, c_open_end, "</c>") orelse
                return error.MalformedXml;
            const c_body = row_body[c_open_end + 1 .. c_close];
            raw_value = extractCellValue(c_body, cell_type);
            formula = extractInner(c_body, "<f", "</f>");
            if (cell_type == .inline_string) {
                const found = extractInlineBodyChecked(c_body);
                inline_body = found.body;
                inline_invalid = found.invalid;
            }
            probe = c_close + "</c>".len;
        } else {
            probe = c_open_end + 1;
        }

        try cells.append(a, .{
            .ref = ref,
            .style_idx = style_idx,
            .style_invalid = style_invalid,
            .cell_type = cell_type,
            .cell_type_invalid = cell_type_invalid,
            .raw_value = raw_value,
            .inline_body = inline_body,
            .inline_invalid = inline_invalid,
            .formula = formula,
        });
    }

    return try cells.toOwnedSlice(a);
}

/// The schema's default (`.number`) for an absent `t`; null for a `t`
/// this view does not know — the caller keeps that provenance.
fn parseCellType(raw: ?[]const u8) ?CellType {
    const t = raw orelse return .number;
    if (std.mem.eql(u8, t, "n")) return .number;
    if (std.mem.eql(u8, t, "s")) return .shared_string;
    if (std.mem.eql(u8, t, "b")) return .boolean;
    if (std.mem.eql(u8, t, "str")) return .formula_string;
    if (std.mem.eql(u8, t, "inlineStr")) return .inline_string;
    if (std.mem.eql(u8, t, "e")) return .error_value;
    if (std.mem.eql(u8, t, "d")) return .date;
    return null;
}

/// `</name>` at `at`, its `>` padded by whitespace or not: one past
/// the `>`, or null.
fn closeTagAt(xml: []const u8, at: usize, name: []const u8) ?usize {
    if (at + 2 + name.len > xml.len) return null;
    if (xml[at] != '<' or xml[at + 1] != '/') return null;
    if (!std.mem.eql(u8, xml[at + 2 .. at + 2 + name.len], name)) return null;
    var i = at + 2 + name.len;
    while (i < xml.len and std.ascii.isWhitespace(xml[i])) i += 1;
    if (i >= xml.len or xml[i] != '>') return null;
    return i + 1;
}

const InlineBody = struct { body: ?[]const u8, invalid: bool };

/// `extractInlineBody` with its structure judged: the `<is>` must be
/// the cell's one direct child. An `<is>` nested under another
/// element (a decoy the shallow read would take), or a second one,
/// is `invalid` (Codex #206 r23 REL-2302).
fn extractInlineBodyChecked(c_body: []const u8) InlineBody {
    var depth: usize = 0;
    var direct: usize = 0;
    var nested = false;
    var pos: usize = 0;
    while (std.mem.indexOfScalarPos(u8, c_body, pos, '<')) |lt| {
        if (lt + 1 >= c_body.len) return .{ .body = null, .invalid = true };
        const c = c_body[lt + 1];
        if (c == '!' or c == '?') {
            pos = wbxml.skipNonElement(c_body, lt) catch return .{ .body = null, .invalid = true };
            continue;
        }
        const end = findTagEnd(c_body, lt) orelse return .{ .body = null, .invalid = true };
        if (c == '/') {
            if (depth == 0) return .{ .body = null, .invalid = true };
            depth -= 1;
            pos = end + 1;
            continue;
        }
        var j = lt + 1;
        while (j < c_body.len and !std.ascii.isWhitespace(c_body[j]) and c_body[j] != '/' and c_body[j] != '>') j += 1;
        const name = c_body[lt + 1 .. j];
        const self_closing = c_body[end - 1] == '/';
        if (std.mem.eql(u8, name, "is")) {
            if (depth == 0) direct += 1 else nested = true;
        }
        if (!self_closing) depth += 1;
        pos = end + 1;
    }
    if (nested or direct > 1 or depth != 0) return .{ .body = null, .invalid = true };
    return .{ .body = extractInlineBody(c_body), .invalid = false };
}

/// The inner bytes of `<is>…</is>`, when the cell has one — the
/// element, not a spelling of it inside a comment, CDATA section or
/// processing instruction (Codex #206 r7 REL-702).
fn extractInlineBody(c_body: []const u8) ?[]const u8 {
    var pos: usize = 0;
    while (std.mem.indexOfScalarPos(u8, c_body, pos, '<')) |lt| {
        if (lt + 1 >= c_body.len) return null;
        const c = c_body[lt + 1];
        if (c == '!' or c == '?') {
            pos = wbxml.skipNonElement(c_body, lt) catch return null;
            continue;
        }
        const is_open_end = findTagEnd(c_body, lt) orelse return null;
        const name_end = lt + 1 + "is".len;
        const is_is = name_end <= c_body.len and std.mem.eql(u8, c_body[lt + 1 .. name_end], "is") and
            (name_end == is_open_end or std.ascii.isWhitespace(c_body[name_end]) or c_body[name_end] == '/');
        if (!is_is) {
            pos = is_open_end + 1;
            continue;
        }
        if (c_body[is_open_end - 1] == '/') return c_body[is_open_end..is_open_end];
        // The close by the walk, past any comment inside.
        var inner = is_open_end + 1;
        while (std.mem.indexOfScalarPos(u8, c_body, inner, '<')) |ilt| {
            if (ilt + 1 >= c_body.len) return null;
            const ic = c_body[ilt + 1];
            if (ic == '!' or ic == '?') {
                inner = wbxml.skipNonElement(c_body, ilt) catch return null;
                continue;
            }
            if (closeTagAt(c_body, ilt, "is") != null) return c_body[is_open_end + 1 .. ilt];
            inner = (findTagEnd(c_body, ilt) orelse return null) + 1;
        }
        return null;
    }
    return null;
}

fn extractCellValue(c_body: []const u8, kind: CellType) ?[]const u8 {
    if (kind == .inline_string) {
        // `<is><t>text</t></is>` — also accept attribute-bearing
        // `<t xml:space="preserve">`. We use the prefix-match form
        // so we don't lose whitespace-significant strings.
        // The body as `extractInlineBody` finds it — by the walk, its
        // close padded or not (Codex #206 r16 REL-1601).
        const is_body = extractInlineBody(c_body) orelse return null;
        return extractInner(is_body, "<t", "</t>");
    }
    return extractInner(c_body, "<v", "</v>");
}

/// Extract the inner text of an element whose opener begins with
/// `open_prefix` (e.g. "<v") and whose closer is `close_tag` (e.g.
/// "</v>"). Tolerates attributes on the opener (e.g.
/// `<t xml:space="preserve">`). Returns `null` when the element is
/// absent. Borrows from `body`.
fn extractInner(body: []const u8, open_prefix: []const u8, close_tag: []const u8) ?[]const u8 {
    var rest = body;
    return extractInnerAdvance(&rest, open_prefix, close_tag);
}

/// `extractInner`, advancing `rest` past the consumed element so a
/// caller can collect repeated children (`<cfRule>`'s up-to-three
/// `<formula>` bodies). First-match semantics are `extractInner`'s
/// exactly — the scan anchors on the literal prefix and bails (null)
/// when the first hit is a longer-named element, the same shape the
/// DV/CF rewriter's `findInnerSpan` bails on, so the two keep
/// agreeing on which body is "the formula".
fn extractInnerAdvance(rest: *[]const u8, open_prefix: []const u8, close_tag: []const u8) ?[]const u8 {
    const body = rest.*;
    const o = std.mem.indexOf(u8, body, open_prefix) orelse return null;
    const after = o + open_prefix.len;
    if (after >= body.len) return null;
    const sep = body[after];
    if (sep != ' ' and sep != '\t' and sep != '\n' and sep != '\r' and sep != '/' and sep != '>') return null;
    const open_end = findTagEnd(body, o) orelse return null;
    if (open_end > o and body[open_end - 1] == '/') {
        // Self-closing <v/> or <f/> — empty body.
        rest.* = body[open_end + 1 ..];
        return body[open_end..open_end];
    }
    const close = std.mem.indexOfPos(u8, body, open_end, close_tag) orelse return null;
    rest.* = body[close + close_tag.len ..];
    return body[open_end + 1 .. close];
}

// ─── <mergeCells><mergeCell …/></mergeCells> ─────────────────────────

fn parseMerges(a: std.mem.Allocator, xml: []const u8) ParseError![]MergeRange {
    assert(xml.len < (1 << 31));

    const mc_open = indexOfTag(xml, 0, "<mergeCells") orelse return &.{};
    const mc_open_end = findTagEnd(xml, mc_open) orelse return error.MalformedXml;
    if (mc_open_end > 0 and xml[mc_open_end - 1] == '/') return &.{};
    const mc_close = std.mem.indexOfPos(u8, xml, mc_open_end, "</mergeCells>") orelse
        return error.MalformedXml;
    const block = xml[mc_open_end + 1 .. mc_close];

    var ranges: std.ArrayListUnmanaged(MergeRange) = .empty;
    errdefer ranges.deinit(a);

    var probe: usize = 0;
    while (indexOfTag(block, probe, "<mergeCell")) |open| {
        const after = open + "<mergeCell".len;
        if (after >= block.len) return error.UnexpectedEof;
        const sep = block[after];
        if (sep != ' ' and sep != '\t' and sep != '\n' and sep != '\r' and sep != '/' and sep != '>') {
            probe = after;
            continue;
        }
        const open_end = findTagEnd(block, open) orelse return error.MalformedXml;
        const attrs = block[open + "<mergeCell".len .. open_end];
        if (attrAt(attrs, "ref")) |ref| {
            try ranges.append(a, .{ .ref = ref });
        }
        probe = open_end + 1;
    }
    return try ranges.toOwnedSlice(a);
}

// ─── <hyperlinks><hyperlink …/></hyperlinks> ─────────────────────────

fn parseHyperlinks(a: std.mem.Allocator, xml: []const u8) ParseError![]Hyperlink {
    assert(xml.len < (1 << 31));

    const hl_open = indexOfTag(xml, 0, "<hyperlinks") orelse return &.{};
    const hl_open_end = findTagEnd(xml, hl_open) orelse return error.MalformedXml;
    if (hl_open_end > 0 and xml[hl_open_end - 1] == '/') return &.{};
    const hl_close = std.mem.indexOfPos(u8, xml, hl_open_end, "</hyperlinks>") orelse
        return error.MalformedXml;
    const block = xml[hl_open_end + 1 .. hl_close];

    var out: std.ArrayListUnmanaged(Hyperlink) = .empty;
    errdefer out.deinit(a);

    var probe: usize = 0;
    while (indexOfTag(block, probe, "<hyperlink")) |open| {
        const after = open + "<hyperlink".len;
        if (after >= block.len) return error.UnexpectedEof;
        const sep = block[after];
        if (sep != ' ' and sep != '\t' and sep != '\n' and sep != '\r' and sep != '/' and sep != '>') {
            probe = after;
            continue;
        }
        const open_end = findTagEnd(block, open) orelse return error.MalformedXml;
        const attrs = block[open + "<hyperlink".len .. open_end];

        const ref = attrAt(attrs, "ref") orelse {
            probe = open_end + 1;
            continue;
        };
        try out.append(a, .{
            .ref = ref,
            .r_id = attrAt(attrs, "r:id"),
            .location = attrAt(attrs, "location"),
            .display = attrAt(attrs, "display"),
            .tooltip = attrAt(attrs, "tooltip"),
        });
        probe = open_end + 1;
    }

    return try out.toOwnedSlice(a);
}

// ─── <dataValidations><dataValidation …>…</dataValidation> ───────────

fn parseValidations(a: std.mem.Allocator, xml: []const u8) ParseError![]DataValidation {
    assert(xml.len < (1 << 31));

    const dv_open = indexOfTag(xml, 0, "<dataValidations") orelse return &.{};
    const dv_open_end = findTagEnd(xml, dv_open) orelse return error.MalformedXml;
    if (dv_open_end > 0 and xml[dv_open_end - 1] == '/') return &.{};
    const dv_close = std.mem.indexOfPos(u8, xml, dv_open_end, "</dataValidations>") orelse
        return error.MalformedXml;
    const block = xml[dv_open_end + 1 .. dv_close];

    var out: std.ArrayListUnmanaged(DataValidation) = .empty;
    errdefer out.deinit(a);

    var probe: usize = 0;
    while (indexOfTag(block, probe, "<dataValidation")) |open| {
        // Must be `<dataValidation` (not `<dataValidations`).
        const after = open + "<dataValidation".len;
        if (after >= block.len) return error.UnexpectedEof;
        const sep = block[after];
        if (sep != ' ' and sep != '\t' and sep != '\n' and sep != '\r' and sep != '/' and sep != '>') {
            probe = after;
            continue;
        }
        const open_end = findTagEnd(block, open) orelse return error.MalformedXml;
        const attrs = block[open + "<dataValidation".len .. open_end];

        const sqref = attrAt(attrs, "sqref") orelse {
            probe = open_end + 1;
            continue;
        };
        // `type` defaults to "list" when absent — Excel writes a bare
        // dropdown without an explicit type attr. `[Community]` and
        // matches src/xlsx.zig:3037.
        const dv_type = attrAt(attrs, "type") orelse "list";
        const operator = attrAt(attrs, "operator");
        const allow_blank = parseBoolAttr(attrs, "allowBlank");
        const show_dropdown = parseBoolAttr(attrs, "showDropDown");

        var formula1: ?[]const u8 = null;
        var formula2: ?[]const u8 = null;

        const self_closing = open_end > open and block[open_end - 1] == '/';
        if (!self_closing) {
            const close = std.mem.indexOfPos(u8, block, open_end, "</dataValidation>") orelse
                return error.MalformedXml;
            const body = block[open_end + 1 .. close];
            formula1 = extractInner(body, "<formula1", "</formula1>");
            formula2 = extractInner(body, "<formula2", "</formula2>");
            probe = close + "</dataValidation>".len;
        } else {
            probe = open_end + 1;
        }

        try out.append(a, .{
            .sqref = sqref,
            .type = dv_type,
            .formula1 = formula1,
            .formula2 = formula2,
            .operator = operator,
            .allow_blank = allow_blank,
            .show_dropdown = show_dropdown,
        });
    }

    return try out.toOwnedSlice(a);
}

// ─── <conditionalFormatting>…<cfRule …>…</cfRule>…/> ─────────────────

fn parseConditionalFormats(a: std.mem.Allocator, xml: []const u8) ParseError![]ConditionalFormat {
    assert(xml.len < (1 << 31));

    var out: std.ArrayListUnmanaged(ConditionalFormat) = .empty;
    errdefer out.deinit(a);

    var probe: usize = 0;
    while (indexOfTag(xml, probe, "<conditionalFormatting")) |cf_open| {
        const after = cf_open + "<conditionalFormatting".len;
        if (after >= xml.len) return error.UnexpectedEof;
        const sep = xml[after];
        if (sep != ' ' and sep != '\t' and sep != '\n' and sep != '\r' and sep != '/' and sep != '>') {
            probe = after;
            continue;
        }
        const cf_open_end = findTagEnd(xml, cf_open) orelse return error.MalformedXml;
        const cf_attrs = xml[cf_open + "<conditionalFormatting".len .. cf_open_end];
        const sqref = attrAt(cf_attrs, "sqref") orelse "";

        const self_closing = cf_open_end > cf_open and xml[cf_open_end - 1] == '/';
        if (self_closing) {
            probe = cf_open_end + 1;
            continue;
        }
        const cf_close = std.mem.indexOfPos(u8, xml, cf_open_end, "</conditionalFormatting>") orelse
            return error.MalformedXml;
        const cf_body = xml[cf_open_end + 1 .. cf_close];
        probe = cf_close + "</conditionalFormatting>".len;

        // Walk every `<cfRule …>` inside this conditionalFormatting.
        var rule_probe: usize = 0;
        while (indexOfTag(cf_body, rule_probe, "<cfRule")) |r_open| {
            const r_after = r_open + "<cfRule".len;
            if (r_after >= cf_body.len) return error.UnexpectedEof;
            const r_sep = cf_body[r_after];
            if (r_sep != ' ' and r_sep != '\t' and r_sep != '\n' and r_sep != '\r' and r_sep != '/' and r_sep != '>') {
                rule_probe = r_after;
                continue;
            }
            const r_open_end = findTagEnd(cf_body, r_open) orelse return error.MalformedXml;
            const r_attrs = cf_body[r_open + "<cfRule".len .. r_open_end];

            const rule_type = attrAt(r_attrs, "type") orelse "";
            const dxf_id = parseU32Attr(r_attrs, "dxfId");
            const priority = parseU32Attr(r_attrs, "priority");

            var formula: ?[]const u8 = null;
            var formula2: ?[]const u8 = null;
            var formula3: ?[]const u8 = null;
            const r_self_closing = r_open_end > r_open and cf_body[r_open_end - 1] == '/';
            if (!r_self_closing) {
                const r_close = std.mem.indexOfPos(u8, cf_body, r_open_end, "</cfRule>") orelse
                    return error.MalformedXml;
                var f_rest = cf_body[r_open_end + 1 .. r_close];
                formula = extractInnerAdvance(&f_rest, "<formula", "</formula>");
                if (formula != null) {
                    formula2 = extractInnerAdvance(&f_rest, "<formula", "</formula>");
                    if (formula2 != null) {
                        formula3 = extractInnerAdvance(&f_rest, "<formula", "</formula>");
                    }
                }
                rule_probe = r_close + "</cfRule>".len;
            } else {
                rule_probe = r_open_end + 1;
            }

            try out.append(a, .{
                .sqref = sqref,
                .type = rule_type,
                .formula = formula,
                .formula2 = formula2,
                .formula3 = formula3,
                .dxf_id = dxf_id,
                .priority = priority,
            });
        }
    }

    return try out.toOwnedSlice(a);
}

// ─── Tests ───────────────────────────────────────────────────────────

const testing = std.testing;

const minimal_sheet_xml =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    \\  <dimension ref="A1"/>
    \\  <sheetData>
    \\    <row r="1"><c r="A1" t="n"><v>42</v></c></row>
    \\  </sheetData>
    \\</worksheet>
;

test "parse: minimal sheet, one row, one numeric cell" {
    var sx = try parse(testing.allocator, minimal_sheet_xml);
    defer sx.deinit(testing.allocator);

    try testing.expect(sx.dimension != null);
    try testing.expectEqualStrings("A1", sx.dimension.?.ref);

    try testing.expectEqual(@as(usize, 1), sx.rows.len);
    try testing.expectEqual(@as(u32, 1), sx.rows[0].row_idx);
    try testing.expectEqual(@as(usize, 1), sx.rows[0].cells.len);

    const c = sx.rows[0].cells[0];
    try testing.expectEqualStrings("A1", c.ref);
    try testing.expectEqual(CellType.number, c.cell_type);
    try testing.expect(c.raw_value != null);
    try testing.expectEqualStrings("42", c.raw_value.?);
    try testing.expect(c.formula == null);

    try testing.expectEqual(@as(usize, 0), sx.merges.len);
    try testing.expectEqual(@as(usize, 0), sx.hyperlinks.len);
    try testing.expectEqual(@as(usize, 0), sx.validations.len);
    try testing.expectEqual(@as(usize, 0), sx.conditional_formats.len);
    try testing.expect(sx.freeze == null);
}

test "parse: shared-string, inline, number, formula cells" {
    const xml =
        \\<worksheet>
        \\  <sheetData>
        \\    <row r="1">
        \\      <c r="A1" t="s"><v>0</v></c>
        \\      <c r="B1" t="inlineStr"><is><t>hello</t></is></c>
        \\      <c r="C1" t="n"><v>3.14</v></c>
        \\      <c r="D1"><f>SUM(A1:C1)</f><v>5</v></c>
        \\    </row>
        \\  </sheetData>
        \\</worksheet>
    ;

    var sx = try parse(testing.allocator, xml);
    defer sx.deinit(testing.allocator);

    try testing.expectEqual(@as(usize, 1), sx.rows.len);
    const cells = sx.rows[0].cells;
    try testing.expectEqual(@as(usize, 4), cells.len);

    try testing.expectEqual(CellType.shared_string, cells[0].cell_type);
    try testing.expectEqualStrings("0", cells[0].raw_value.?);

    try testing.expectEqual(CellType.inline_string, cells[1].cell_type);
    try testing.expectEqualStrings("hello", cells[1].raw_value.?);

    try testing.expectEqual(CellType.number, cells[2].cell_type);
    try testing.expectEqualStrings("3.14", cells[2].raw_value.?);

    try testing.expectEqual(CellType.number, cells[3].cell_type);
    try testing.expect(cells[3].formula != null);
    try testing.expectEqualStrings("SUM(A1:C1)", cells[3].formula.?);
    try testing.expectEqualStrings("5", cells[3].raw_value.?);
}

test "parse: merge ranges and hyperlinks" {
    const xml =
        \\<worksheet>
        \\  <sheetData>
        \\    <row r="1"><c r="A1"><v>1</v></c></row>
        \\  </sheetData>
        \\  <mergeCells count="2">
        \\    <mergeCell ref="A1:B2"/>
        \\    <mergeCell ref="C3:D4"/>
        \\  </mergeCells>
        \\  <hyperlinks>
        \\    <hyperlink ref="E5" r:id="rId1" display="Example" tooltip="link"/>
        \\    <hyperlink ref="F6" location="Sheet2!A1"/>
        \\  </hyperlinks>
        \\</worksheet>
    ;

    var sx = try parse(testing.allocator, xml);
    defer sx.deinit(testing.allocator);

    try testing.expectEqual(@as(usize, 2), sx.merges.len);
    try testing.expectEqualStrings("A1:B2", sx.merges[0].ref);
    try testing.expectEqualStrings("C3:D4", sx.merges[1].ref);

    try testing.expectEqual(@as(usize, 2), sx.hyperlinks.len);
    try testing.expectEqualStrings("E5", sx.hyperlinks[0].ref);
    try testing.expect(sx.hyperlinks[0].r_id != null);
    try testing.expectEqualStrings("rId1", sx.hyperlinks[0].r_id.?);
    try testing.expectEqualStrings("Example", sx.hyperlinks[0].display.?);
    try testing.expectEqualStrings("link", sx.hyperlinks[0].tooltip.?);

    try testing.expectEqualStrings("F6", sx.hyperlinks[1].ref);
    try testing.expect(sx.hyperlinks[1].r_id == null);
    try testing.expectEqualStrings("Sheet2!A1", sx.hyperlinks[1].location.?);
}

test "parse: data validations" {
    const xml =
        \\<worksheet>
        \\  <sheetData/>
        \\  <dataValidations count="2">
        \\    <dataValidation type="list" allowBlank="1" showDropDown="0" sqref="A1:A10">
        \\      <formula1>"yes,no,maybe"</formula1>
        \\    </dataValidation>
        \\    <dataValidation type="whole" operator="between" sqref="B1:B10">
        \\      <formula1>1</formula1>
        \\      <formula2>100</formula2>
        \\    </dataValidation>
        \\  </dataValidations>
        \\</worksheet>
    ;

    var sx = try parse(testing.allocator, xml);
    defer sx.deinit(testing.allocator);

    try testing.expectEqual(@as(usize, 2), sx.validations.len);

    const v0 = sx.validations[0];
    try testing.expectEqualStrings("A1:A10", v0.sqref);
    try testing.expectEqualStrings("list", v0.type);
    try testing.expect(v0.allow_blank);
    try testing.expect(!v0.show_dropdown);
    try testing.expect(v0.formula1 != null);
    try testing.expectEqualStrings("\"yes,no,maybe\"", v0.formula1.?);

    const v1 = sx.validations[1];
    try testing.expectEqualStrings("B1:B10", v1.sqref);
    try testing.expectEqualStrings("whole", v1.type);
    try testing.expectEqualStrings("between", v1.operator.?);
    try testing.expectEqualStrings("1", v1.formula1.?);
    try testing.expectEqualStrings("100", v1.formula2.?);
}

test "parse: conditional formats with cfRule" {
    const xml =
        \\<worksheet>
        \\  <sheetData/>
        \\  <conditionalFormatting sqref="A1:A100">
        \\    <cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan">
        \\      <formula>10</formula>
        \\    </cfRule>
        \\    <cfRule type="expression" dxfId="2" priority="3">
        \\      <formula>$B1=TRUE</formula>
        \\    </cfRule>
        \\  </conditionalFormatting>
        \\</worksheet>
    ;

    var sx = try parse(testing.allocator, xml);
    defer sx.deinit(testing.allocator);

    try testing.expectEqual(@as(usize, 2), sx.conditional_formats.len);

    const cf0 = sx.conditional_formats[0];
    try testing.expectEqualStrings("A1:A100", cf0.sqref);
    try testing.expectEqualStrings("cellIs", cf0.type);
    try testing.expectEqual(@as(?u32, 0), cf0.dxf_id);
    try testing.expectEqual(@as(?u32, 1), cf0.priority);
    try testing.expectEqualStrings("10", cf0.formula.?);

    const cf1 = sx.conditional_formats[1];
    try testing.expectEqualStrings("expression", cf1.type);
    try testing.expectEqualStrings("$B1=TRUE", cf1.formula.?);
    try testing.expectEqual(@as(?u32, 2), cf1.dxf_id);
    try testing.expectEqual(@as(?u32, 3), cf1.priority);
    // One-formula rules leave the second and third slots empty.
    try testing.expect(cf0.formula2 == null and cf0.formula3 == null);
    try testing.expect(cf1.formula2 == null and cf1.formula3 == null);
}

test "parse: cfRule collects up to three formula bodies in document order" {
    const xml =
        \\<worksheet>
        \\  <sheetData/>
        \\  <conditionalFormatting sqref="A1:A10">
        \\    <cfRule type="cellIs" dxfId="0" priority="1" operator="between">
        \\      <formula>2</formula>
        \\      <formula>4</formula>
        \\    </cfRule>
        \\    <cfRule type="cellIs" priority="2" operator="equal">
        \\      <formula>1</formula>
        \\      <formula>2</formula>
        \\      <formula>3</formula>
        \\      <formula>ignored past the schema's three</formula>
        \\    </cfRule>
        \\    <cfRule type="containsBlanks" priority="3"/>
        \\  </conditionalFormatting>
        \\</worksheet>
    ;

    var sx = try parse(testing.allocator, xml);
    defer sx.deinit(testing.allocator);

    try testing.expectEqual(@as(usize, 3), sx.conditional_formats.len);

    const between = sx.conditional_formats[0];
    try testing.expectEqualStrings("2", between.formula.?);
    try testing.expectEqualStrings("4", between.formula2.?);
    try testing.expect(between.formula3 == null);

    const three = sx.conditional_formats[1];
    try testing.expectEqualStrings("1", three.formula.?);
    try testing.expectEqualStrings("2", three.formula2.?);
    try testing.expectEqualStrings("3", three.formula3.?);

    const bodiless = sx.conditional_formats[2];
    try testing.expect(bodiless.formula == null);
    try testing.expect(bodiless.formula2 == null);
    try testing.expect(bodiless.formula3 == null);
}

test "sourceSpanOf: view slices map to their source bytes; re-encoded CDATA has no home" {
    // A comment before the block shifts sanitized offsets away from
    // source offsets; the contiguity runs carry the mapping across.
    const xml =
        "<worksheet><!-- shift --><conditionalFormatting sqref=\"A1\">" ++
        "<cfRule type=\"expression\" priority=\"1\"><formula>AB1+C2</formula></cfRule>" ++
        "</conditionalFormatting></worksheet>";
    var sx = try parse(testing.allocator, xml);
    defer sx.deinit(testing.allocator);
    const f = sx.conditional_formats[0].formula.?;
    const span = sx.sourceSpanOf(f) orelse return error.TestUnexpectedResult;
    try testing.expectEqualStrings("AB1+C2", xml[span[0]..span[1]]);

    // CDATA contents are re-encoded, not copied — a slice drawn from
    // them maps to nothing and a splice must refuse rather than guess.
    const xml2 =
        "<worksheet><conditionalFormatting sqref=\"A1\">" ++
        "<cfRule type=\"expression\" priority=\"1\"><formula><![CDATA[X1]]></formula></cfRule>" ++
        "</conditionalFormatting></worksheet>";
    var sx2 = try parse(testing.allocator, xml2);
    defer sx2.deinit(testing.allocator);
    const f2 = sx2.conditional_formats[0].formula.?;
    try testing.expectEqualStrings("X1", f2);
    try testing.expect(sx2.sourceSpanOf(f2) == null);

    // A slice from some other buffer maps to nothing either.
    try testing.expect(sx.sourceSpanOf("AB1+C2") == null);
}

test "parse: cfRule formula scan keeps extractInner's bail-on-decoy shape" {
    // A longer-named element at the first `<formula` hit bails the
    // whole scan (formula == null), exactly as `extractInner` did —
    // the DV/CF rewriter's `findInnerSpan` bails on the same shape,
    // and the two must agree on which body is "the formula". A decoy
    // AFTER a real formula likewise ends the collection there.
    const xml =
        \\<worksheet>
        \\  <sheetData/>
        \\  <conditionalFormatting sqref="A1">
        \\    <cfRule type="cellIs" priority="1" operator="equal">
        \\      <formulaX>9</formulaX>
        \\      <formula>1</formula>
        \\    </cfRule>
        \\    <cfRule type="cellIs" priority="2" operator="between">
        \\      <formula>1</formula>
        \\      <formulaX>9</formulaX>
        \\      <formula>2</formula>
        \\    </cfRule>
        \\    <cfRule type="cellIs" priority="3" operator="equal">
        \\      <formula/>
        \\      <formula>7</formula>
        \\    </cfRule>
        \\  </conditionalFormatting>
        \\</worksheet>
    ;

    var sx = try parse(testing.allocator, xml);
    defer sx.deinit(testing.allocator);

    try testing.expectEqual(@as(usize, 3), sx.conditional_formats.len);
    try testing.expect(sx.conditional_formats[0].formula == null);
    try testing.expectEqualStrings("1", sx.conditional_formats[1].formula.?);
    try testing.expect(sx.conditional_formats[1].formula2 == null);
    // A self-closing <formula/> is an empty first body; the scan
    // continues past it to the second.
    try testing.expectEqualStrings("", sx.conditional_formats[2].formula.?);
    try testing.expectEqualStrings("7", sx.conditional_formats[2].formula2.?);
}

test "parse: freeze pane" {
    const xml =
        \\<worksheet>
        \\  <sheetViews>
        \\    <sheetView workbookViewId="0">
        \\      <pane xSplit="2" ySplit="3" topLeftCell="C4" state="frozen"/>
        \\    </sheetView>
        \\  </sheetViews>
        \\  <sheetData/>
        \\</worksheet>
    ;

    var sx = try parse(testing.allocator, xml);
    defer sx.deinit(testing.allocator);

    try testing.expect(sx.freeze != null);
    try testing.expectEqual(@as(u32, 2), sx.freeze.?.x_split);
    try testing.expectEqual(@as(u32, 3), sx.freeze.?.y_split);
    try testing.expectEqualStrings("C4", sx.freeze.?.top_left_cell);
}

test "parse: row attributes — height, customHeight, hidden" {
    const xml =
        \\<worksheet>
        \\  <sheetData>
        \\    <row r="1" ht="22.5" customHeight="1"><c r="A1"><v>1</v></c></row>
        \\    <row r="2" hidden="true"><c r="A2"><v>2</v></c></row>
        \\    <row r="3"/>
        \\  </sheetData>
        \\</worksheet>
    ;

    var sx = try parse(testing.allocator, xml);
    defer sx.deinit(testing.allocator);

    try testing.expectEqual(@as(usize, 3), sx.rows.len);

    try testing.expect(sx.rows[0].height != null);
    try testing.expect(sx.rows[0].height.? > 22.0 and sx.rows[0].height.? < 23.0);
    try testing.expect(sx.rows[0].custom_height);
    try testing.expect(!sx.rows[0].hidden);

    try testing.expect(sx.rows[1].height == null);
    try testing.expect(sx.rows[1].hidden);

    try testing.expectEqual(@as(usize, 0), sx.rows[2].cells.len);
}

test "parse: malformed sheet without <worksheet> root rejected" {
    const xml = "<?xml version=\"1.0\"?><styleSheet><cellXfs/></styleSheet>";
    try testing.expectError(error.MalformedXml, parse(testing.allocator, xml));
}

test "parse: comment-wrapped tags don't poison detection" {
    // Comments containing tag-shaped substrings must not register as
    // real elements. The scanner skips over `<!-- ... -->` blocks.
    const xml =
        \\<worksheet>
        \\  <!-- <mergeCell ref="X9:X10"/> this is a comment, not a real merge -->
        \\  <sheetData>
        \\    <row r="1"><c r="A1"><v>1</v></c></row>
        \\  </sheetData>
        \\</worksheet>
    ;

    var sx = try parse(testing.allocator, xml);
    defer sx.deinit(testing.allocator);
    try testing.expectEqual(@as(usize, 0), sx.merges.len);
    try testing.expectEqual(@as(usize, 1), sx.rows.len);
}

test "parse: attribute with literal '>' inside quotes does not terminate tag" {
    // Defensive: a `>` inside a quoted attribute must not be read
    // as the tag-closing `>`. We pick an attribute that's unlikely
    // in real workbooks but exercises the scanner's quote tracker.
    const xml =
        \\<worksheet>
        \\  <sheetData>
        \\    <row r="1"><c r="A1" t="inlineStr"><is><t>a&gt;b</t></is></c></row>
        \\  </sheetData>
        \\</worksheet>
    ;

    var sx = try parse(testing.allocator, xml);
    defer sx.deinit(testing.allocator);
    try testing.expectEqual(@as(usize, 1), sx.rows.len);
    try testing.expectEqual(@as(usize, 1), sx.rows[0].cells.len);
    try testing.expectEqual(CellType.inline_string, sx.rows[0].cells[0].cell_type);
    // Raw value is borrowed from xml, with entities still escaped —
    // the contract of `raw_value` is "no decode".
    try testing.expectEqualStrings("a&gt;b", sx.rows[0].cells[0].raw_value.?);
}
