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
    cell_type: CellType,
    /// Raw inner text of `<v>` (or `<is><t>` for inline strings).
    /// Borrows. No XML-entity decoding — caller decodes if needed.
    raw_value: ?[]const u8,
    /// Raw inner text of `<f>`. Borrows. Not rewritten.
    formula: ?[]const u8,
};

pub const Row = struct {
    row_idx: u32, // 1-based, matches OOXML `<row r="…">`.
    cells: []Cell,
    height: ?f64,
    custom_height: bool,
    hidden: bool,
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
    dxf_id: ?u32,
    priority: ?u32,
};

pub const FreezePane = struct {
    x_split: u32,
    y_split: u32,
    /// Borrows. The `topLeftCell` attribute, e.g. "B2".
    top_left_cell: []const u8,
};

pub const SheetXml = struct {
    dimension: ?Dimension,
    rows: []Row,
    merges: []MergeRange,
    hyperlinks: []Hyperlink,
    validations: []DataValidation,
    conditional_formats: []ConditionalFormat,
    freeze: ?FreezePane,
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
        self.dimension = null;
        self.freeze = null;
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
    // worksheets — see doc-comment above).
    const sanitized = try sanitizeXml(a, xml);

    // Quick well-formedness gate: a worksheet part must contain a
    // `<worksheet` root tag. Anything else is rejected up front so
    // callers don't get an empty struct from a styles or sst part
    // wired into the wrong slot.
    if (std.mem.indexOf(u8, sanitized, "<worksheet") == null) return error.MalformedXml;

    const dim = parseDimension(sanitized);
    const freeze = parseFreezePane(sanitized);

    const rows = try parseRows(a, sanitized);
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
fn sanitizeXml(allocator: std.mem.Allocator, xml: []const u8) ParseError![]const u8 {
    assert(xml.len < (1 << 31));

    var out: std.ArrayList(u8) = .empty;
    errdefer out.deinit(allocator);
    try out.ensureTotalCapacity(allocator, xml.len);

    var i: usize = 0;
    while (i < xml.len) {
        const c = xml[i];
        if (c != '<') {
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
        try out.appendSlice(allocator, xml[i .. end + 1]);
        i = end + 1;
    }

    return try out.toOwnedSlice(allocator);
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

fn parseRows(a: std.mem.Allocator, xml: []const u8) ParseError![]Row {
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

        const row_idx_raw = attrAt(row_attrs, "r") orelse {
            probe = row_open_end + 1;
            continue;
        };
        const row_idx = std.fmt.parseInt(u32, row_idx_raw, 10) catch {
            probe = row_open_end + 1;
            continue;
        };
        if (row_idx == 0) {
            probe = row_open_end + 1;
            continue;
        }

        const height = parseF64Attr(row_attrs, "ht");
        const custom_height = parseBoolAttr(row_attrs, "customHeight");
        const hidden = parseBoolAttr(row_attrs, "hidden");

        // Self-closing `<row r="N"/>` — empty row.
        const self_closing = row_open_end > row_open and body[row_open_end - 1] == '/';
        var cells: []Cell = &.{};
        if (!self_closing) {
            const row_close = std.mem.indexOfPos(u8, body, row_open_end, "</row>") orelse
                return error.MalformedXml;
            const row_body = body[row_open_end + 1 .. row_close];
            cells = try parseCells(a, row_body);
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
        });
    }

    return try rows.toOwnedSlice(a);
}

fn parseCells(a: std.mem.Allocator, row_body: []const u8) ParseError![]Cell {
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

        const ref = attrAt(c_attrs, "r") orelse {
            probe = c_open_end + 1;
            continue;
        };
        const style_idx: ?u32 = parseU32Attr(c_attrs, "s");
        const cell_type = parseCellType(attrAt(c_attrs, "t"));

        const self_closing = c_open_end > c_open and row_body[c_open_end - 1] == '/';
        var raw_value: ?[]const u8 = null;
        var formula: ?[]const u8 = null;
        if (!self_closing) {
            const c_close = std.mem.indexOfPos(u8, row_body, c_open_end, "</c>") orelse
                return error.MalformedXml;
            const c_body = row_body[c_open_end + 1 .. c_close];
            raw_value = extractCellValue(c_body, cell_type);
            formula = extractInner(c_body, "<f", "</f>");
            probe = c_close + "</c>".len;
        } else {
            probe = c_open_end + 1;
        }

        try cells.append(a, .{
            .ref = ref,
            .style_idx = style_idx,
            .cell_type = cell_type,
            .raw_value = raw_value,
            .formula = formula,
        });
    }

    return try cells.toOwnedSlice(a);
}

fn parseCellType(raw: ?[]const u8) CellType {
    const t = raw orelse return .number;
    if (std.mem.eql(u8, t, "n")) return .number;
    if (std.mem.eql(u8, t, "s")) return .shared_string;
    if (std.mem.eql(u8, t, "b")) return .boolean;
    if (std.mem.eql(u8, t, "str")) return .formula_string;
    if (std.mem.eql(u8, t, "inlineStr")) return .inline_string;
    if (std.mem.eql(u8, t, "e")) return .error_value;
    if (std.mem.eql(u8, t, "d")) return .date;
    return .number;
}

fn extractCellValue(c_body: []const u8, kind: CellType) ?[]const u8 {
    if (kind == .inline_string) {
        // `<is><t>text</t></is>` — also accept attribute-bearing
        // `<t xml:space="preserve">`. We use the prefix-match form
        // so we don't lose whitespace-significant strings.
        const is_open = std.mem.indexOf(u8, c_body, "<is") orelse return null;
        const is_open_end = findTagEnd(c_body, is_open) orelse return null;
        const is_close = std.mem.indexOfPos(u8, c_body, is_open_end, "</is>") orelse return null;
        const is_body = c_body[is_open_end + 1 .. is_close];
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
    const o = std.mem.indexOf(u8, body, open_prefix) orelse return null;
    const after = o + open_prefix.len;
    if (after >= body.len) return null;
    const sep = body[after];
    if (sep != ' ' and sep != '\t' and sep != '\n' and sep != '\r' and sep != '/' and sep != '>') return null;
    const open_end = findTagEnd(body, o) orelse return null;
    if (open_end > o and body[open_end - 1] == '/') {
        // Self-closing <v/> or <f/> — empty body.
        return body[open_end..open_end];
    }
    const close = std.mem.indexOfPos(u8, body, open_end, close_tag) orelse return null;
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
            const r_self_closing = r_open_end > r_open and cf_body[r_open_end - 1] == '/';
            if (!r_self_closing) {
                const r_close = std.mem.indexOfPos(u8, cf_body, r_open_end, "</cfRule>") orelse
                    return error.MalformedXml;
                const r_body = cf_body[r_open_end + 1 .. r_close];
                formula = extractInner(r_body, "<formula", "</formula>");
                rule_probe = r_close + "</cfRule>".len;
            } else {
                rule_probe = r_open_end + 1;
            }

            try out.append(a, .{
                .sqref = sqref,
                .type = rule_type,
                .formula = formula,
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
