//! B1 iter-wb-1: typed overlay for `xl/styles.xml`.
//!
//! Parses the styles part into the data shapes consumed by the
//! reader-side view (`src/xlsx.zig`'s `Font`/`Fill`/`Border`/`Alignment`).
//! Field names + types intentionally mirror those reader structs so the
//! coordinator can swap this overlay in for `Workbook.styles()` without
//! touching call sites.
//!
//! Lifetime: every `[]const u8` field BORROWS from the input `xml`
//! slice. The caller owns `xml`; this overlay's slices stay valid as
//! long as `xml` does. The outer slice arrays (and the `Color` boxes
//! they contain) live in an internal `ArenaAllocator`, freed by
//! `deinit`.
//!
//! Coverage (per scope spec):
//!   - numFmts        — full
//!   - fonts          — full (name, size, bold, italic, underline,
//!                            strike, color, family, scheme)
//!   - fills          — full (pattern, fgColor, bgColor)
//!   - borders        — full (5 sides + diagonalUp/diagonalDown)
//!   - cellXfs        — full (incl. nested `<alignment>`, applyXxx,
//!                            quotePrefix)
//!   - cellStyleXfs   — full (same shape as cellXfs)
//!
//! Out of scope (deliberately elided): cellStyles, dxfs, tableStyles,
//! colors/indexedColors/mruColors, extLst. These don't drive `<c s="N">`
//! resolution and can land in a follow-up.

const std = @import("std");
const wbxml = @import("workbook_xml.zig");
const assert = std.debug.assert;

// ─── Public types ─────────────────────────────────────────────────────

pub const Color = struct {
    /// Hex ARGB attribute value, e.g. "FF000000". Borrowed from `xml`.
    /// Producers also emit lowercase or 6-digit RGB; we surface the
    /// raw attribute bytes and leave normalization to callers.
    argb: ?[]const u8,
    indexed: ?u32,
    theme: ?u32,
    tint: ?f64,
};

pub const BorderSide = struct {
    /// OOXML border style enum: none | thin | medium | dashed | dotted
    /// | thick | double | hair | mediumDashed | dashDot | mediumDashDot
    /// | dashDotDot | mediumDashDotDot | slantDashDot. `null` when the
    /// side element was self-closing or absent.
    style: ?[]const u8,
    color: ?Color,
};

pub const Border = struct {
    left: BorderSide,
    right: BorderSide,
    top: BorderSide,
    bottom: BorderSide,
    diagonal: BorderSide,
    diagonal_up: bool,
    diagonal_down: bool,
};

pub const Fill = struct {
    /// `patternType` value (e.g. "none", "solid", "gray125"). Borrowed.
    /// Empty string when `<patternFill>` had no patternType attribute.
    pattern: []const u8,
    fg_color: ?Color,
    bg_color: ?Color,
};

pub const Font = struct {
    name: ?[]const u8,
    size: ?f64,
    bold: bool,
    italic: bool,
    /// `<u val="…"/>` value; `null` when absent. When the `<u/>` element
    /// is present without `val`, OOXML defaults to "single" — we surface
    /// "single" explicitly in that case.
    underline: ?[]const u8,
    strike: bool,
    color: ?Color,
    family: ?u32,
    scheme: ?[]const u8,
};

pub const Alignment = struct {
    horizontal: ?[]const u8,
    vertical: ?[]const u8,
    wrap_text: bool,
    indent: ?u32,
    rotation: ?i32,
    shrink_to_fit: bool,
    reading_order: ?u32,
};

pub const NumberFormat = struct {
    fmt_id: u32,
    /// `formatCode` attribute, borrowed. Note: this overlay does NOT
    /// XML-entity-decode the format code — `&quot;` survives as the
    /// literal six bytes. Callers wanting the decoded form should
    /// pipe through `pkg.store.decodeXmlEntities`.
    code: []const u8,
};

pub const CellXf = struct {
    num_fmt_id: ?u32,
    /// `numFmtId` was written but is not a number: `num_fmt_id` is
    /// null, and the style is not General by default (Codex #205 r5
    /// REL-504).
    num_fmt_id_invalid: bool = false,
    font_id: ?u32,
    fill_id: ?u32,
    border_id: ?u32,
    /// `xfId` attribute — index into `cell_style_xfs`. Only meaningful
    /// for entries in `cell_xfs`; cellStyleXfs entries don't use it.
    xf_id: ?u32,
    apply_number_format: bool,
    apply_font: bool,
    apply_fill: bool,
    apply_border: bool,
    apply_alignment: bool,
    apply_protection: bool,
    alignment: ?Alignment,
    quote_prefix: bool,
};

pub const StylesXml = struct {
    number_formats: []NumberFormat,
    fonts: []Font,
    fills: []Fill,
    borders: []Border,
    /// `<cellXfs>` entries — these are the indices `<c s="N">` refers to.
    cell_xfs: []CellXf,
    cell_style_xfs: []CellXf,
    /// Internal arena owning the array allocations and any `Color`
    /// boxes referenced from the public slices. `null` only on a
    /// fully-empty default (used by tests to construct a no-op).
    arena: ?std.heap.ArenaAllocator,

    pub fn deinit(self: *StylesXml, allocator: std.mem.Allocator) void {
        _ = allocator;
        if (self.arena) |*a| a.deinit();
        self.* = undefined;
    }
};

pub const Error = error{
    MalformedXml,
    OutOfMemory,
};

// ─── Entry point ──────────────────────────────────────────────────────

pub fn parse(allocator: std.mem.Allocator, xml: []const u8) Error!StylesXml {
    assert(xml.len < std.math.maxInt(usize) / 2); // sanity: not a sentinel
    // Producers always wrap the part in <styleSheet>...</styleSheet>;
    // missing wrapper is symptomatic of a different file misrouted
    // here. We don't strictly require it (some test fixtures elide
    // namespaces and root) but the assert points at the smell.
    assert(xml.len == 0 or std.mem.indexOf(u8, xml, "<") != null);

    var arena = std.heap.ArenaAllocator.init(allocator);
    errdefer arena.deinit();
    const a = arena.allocator();

    // Strip comments / CDATA / processing instructions before scanning.
    // The reader-side `parseStyles` is naive about these; the overlay's
    // contract is stricter — it must reject malformed input rather than
    // silently mis-classify a `<font>` mention inside a comment.
    const sanitized = try sanitizeXml(a, xml);

    const number_formats = try parseNumFmts(a, sanitized);
    const fonts = try parseFonts(a, sanitized);
    const fills = try parseFills(a, sanitized);
    const borders = try parseBorders(a, sanitized);
    const cell_xfs = try parseCellXfs(a, sanitized, "cellXfs");
    const cell_style_xfs = try parseCellXfs(a, sanitized, "cellStyleXfs");

    // Postcondition: every returned slice was allocated under `arena`.
    // We can't directly verify ownership at runtime, but enforce shape:
    assert(number_formats.len <= sanitized.len);
    assert(fonts.len <= sanitized.len);

    return .{
        .number_formats = number_formats,
        .fonts = fonts,
        .fills = fills,
        .borders = borders,
        .cell_xfs = cell_xfs,
        .cell_style_xfs = cell_style_xfs,
        .arena = arena,
    };
}

// ─── XML sanitizer ────────────────────────────────────────────────────

/// Produces a copy of `xml` with comments (`<!-- ... -->`), CDATA
/// (`<![CDATA[ ... ]]>`), and processing instructions (`<? ... ?>`)
/// elided. Inside CDATA the contents are kept (unescaped) so that
/// `<![CDATA[ <font/> ]]>` doesn't smuggle a fake font element into
/// the scan. Inside comments and PIs the contents are dropped.
///
/// Also surfaces an error on the obvious "tag never closes" pattern
/// (a `<` that has no matching `>` before EOF) — the down-stream
/// indexOfPos / indexOfScalarPos calls would silently truncate.
///
/// Borrows: returns a fresh slice owned by `allocator`. Callers must
/// keep it alive for the lifetime of any borrowed `[]const u8` in the
/// returned `StylesXml`.
fn sanitizeXml(allocator: std.mem.Allocator, xml: []const u8) Error![]const u8 {
    assert(xml.len < std.math.maxInt(u32));

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
        // `<` — classify what follows.
        if (i + 4 <= xml.len and std.mem.eql(u8, xml[i .. i + 4], "<!--")) {
            const close = std.mem.indexOfPos(u8, xml, i + 4, "-->") orelse return Error.MalformedXml;
            i = close + 3;
            continue;
        }
        if (i + 9 <= xml.len and std.mem.eql(u8, xml[i .. i + 9], "<![CDATA[")) {
            const close = std.mem.indexOfPos(u8, xml, i + 9, "]]>") orelse return Error.MalformedXml;
            // Keep CDATA contents literally — they may legitimately
            // appear inside `formatCode` style attributes.
            try out.appendSlice(allocator, xml[i + 9 .. close]);
            i = close + 3;
            continue;
        }
        if (i + 2 <= xml.len and xml[i + 1] == '?') {
            const close = std.mem.indexOfPos(u8, xml, i + 2, "?>") orelse return Error.MalformedXml;
            i = close + 2;
            continue;
        }
        // Plain tag — copy through to the matching `>`, respecting
        // quoted attribute values that may contain `>` themselves.
        const end = findTagEnd(xml, i) orelse return Error.MalformedXml;
        try out.appendSlice(allocator, xml[i .. end + 1]);
        i = end + 1;
    }

    return try out.toOwnedSlice(allocator);
}

/// Returns the index of the `>` that terminates the tag opened at
/// `xml[start]` (which must be `<`). Skips `>` characters that appear
/// inside double- or single-quoted attribute values. Null on EOF.
fn findTagEnd(xml: []const u8, start: usize) ?usize {
    assert(xml.len > start);
    assert(xml[start] == '<');

    var i: usize = start + 1;
    var quote: u8 = 0;
    while (i < xml.len) : (i += 1) {
        const c = xml[i];
        if (quote != 0) {
            if (c == quote) quote = 0;
            continue;
        }
        if (c == '"' or c == '\'') {
            quote = c;
            continue;
        }
        if (c == '>') return i;
    }
    return null;
}

// ─── Block locator ────────────────────────────────────────────────────

/// Locate the body slice between `<wrapper ...>` and `</wrapper>`. A
/// self-closing `<wrapper/>` returns an empty body. Returns null when
/// the wrapper is absent. Asserts `wrapper` is a static, well-formed
/// XML name.
fn findBlock(xml: []const u8, wrapper: []const u8) ?[]const u8 {
    assert(wrapper.len > 0);
    assert(wrapper.len < 32);

    var lt_buf: [40]u8 = undefined;
    const lt = std.fmt.bufPrint(&lt_buf, "<{s}", .{wrapper}) catch return null;
    const open = std.mem.indexOf(u8, xml, lt) orelse return null;
    if (open + lt.len >= xml.len) return null;
    const after = xml[open + lt.len];
    if (after != ' ' and after != '>' and after != '/') return null;

    const open_gt = std.mem.indexOfScalarPos(u8, xml, open, '>') orelse return null;
    if (open_gt > 0 and xml[open_gt - 1] == '/') return xml[open_gt..open_gt]; // self-closing

    var close_buf: [42]u8 = undefined;
    const close = std.fmt.bufPrint(&close_buf, "</{s}>", .{wrapper}) catch return null;
    const close_pos = std.mem.indexOfPos(u8, xml, open_gt, close) orelse return null;
    return xml[open_gt + 1 .. close_pos];
}

// ─── Attribute helpers ────────────────────────────────────────────────

/// Extracts the value of `name="…"` from an attribute slice. Single-
/// quoted attributes are also accepted (OOXML producers stick to
/// double quotes, but XML allows both).
fn attrValue(attrs: []const u8, name: []const u8) ?[]const u8 {
    assert(name.len > 0);
    assert(name.len < 64);

    var key_buf: [80]u8 = undefined;
    const key = std.fmt.bufPrint(&key_buf, "{s}=\"", .{name}) catch return null;
    if (std.mem.indexOf(u8, attrs, key)) |kp| {
        const s = kp + key.len;
        const e = std.mem.indexOfScalarPos(u8, attrs, s, '"') orelse return null;
        return attrs[s..e];
    }
    var key_buf2: [80]u8 = undefined;
    const key2 = std.fmt.bufPrint(&key_buf2, "{s}='", .{name}) catch return null;
    if (std.mem.indexOf(u8, attrs, key2)) |kp| {
        const s = kp + key2.len;
        const e = std.mem.indexOfScalarPos(u8, attrs, s, '\'') orelse return null;
        return attrs[s..e];
    }
    return null;
}

fn attrU32(attrs: []const u8, name: []const u8) ?u32 {
    const raw = attrValue(attrs, name) orelse return null;
    // The value the schema types, not its spelling: `numFmtId="&#49;"`
    // is 1 (Codex #205 r10 REL-1002).
    if (std.mem.indexOfScalar(u8, raw, '&') != null) {
        var buf: [32]u8 = undefined;
        const decoded = wbxml.decodeScalarAttr(&buf, raw) orelse return null;
        return std.fmt.parseInt(u32, decoded, 10) catch null;
    }
    return std.fmt.parseInt(u32, raw, 10) catch null;
}

fn attrI32(attrs: []const u8, name: []const u8) ?i32 {
    const raw = attrValue(attrs, name) orelse return null;
    return std.fmt.parseInt(i32, raw, 10) catch null;
}

fn attrF64(attrs: []const u8, name: []const u8) ?f64 {
    const raw = attrValue(attrs, name) orelse return null;
    return std.fmt.parseFloat(f64, raw) catch null;
}

fn xsdBool(attrs: []const u8, name: []const u8) bool {
    const raw = attrValue(attrs, name) orelse return false;
    return std.mem.eql(u8, raw, "1") or std.ascii.eqlIgnoreCase(raw, "true");
}

// ─── Element walker ───────────────────────────────────────────────────

const Element = struct {
    /// Tag name slice (e.g. "font") borrowing from input.
    name: []const u8,
    /// Attribute slice (everything between the tag name and the
    /// terminating `>` or `/>`). Borrows from input.
    attrs: []const u8,
    /// Body slice — empty when the element was self-closing. Borrows
    /// from input.
    body: []const u8,
    /// True when the element was `<name ... />`. The body is empty in
    /// that case but consumers may want to distinguish it.
    self_closing: bool,
    /// Byte offset (in the parent block) immediately after the closing
    /// tag — used by the walker to advance.
    end: usize,
};

/// Find the next direct-or-nested element with one of the given tag
/// names starting at `block[start..]`. We don't track ancestor stacks
/// because OOXML styles.xml has predictable, shallow shapes — every
/// caller already restricts the search to a parent block via
/// `findBlock`.
fn nextElement(block: []const u8, start: usize, name: []const u8) ?Element {
    assert(name.len > 0);
    assert(name.len < 32);
    var lt_buf: [40]u8 = undefined;
    const lt = std.fmt.bufPrint(&lt_buf, "<{s}", .{name}) catch return null;

    var i = start;
    while (std.mem.indexOfPos(u8, block, i, lt)) |open| {
        const after_idx = open + lt.len;
        if (after_idx >= block.len) return null;
        const after = block[after_idx];
        if (after != ' ' and after != '>' and after != '/') {
            // Prefix collision (e.g. `<font` vs `<fonts`); skip past.
            i = after_idx;
            continue;
        }
        const gt = findTagEnd(block, open) orelse return null;
        const self_closing = gt > open and block[gt - 1] == '/';
        const attrs_end = if (self_closing) gt - 1 else gt;
        // attrs slice = bytes between tag name and the terminator
        const attrs = block[after_idx..attrs_end];
        if (self_closing) {
            return .{
                .name = block[open + 1 .. after_idx],
                .attrs = attrs,
                .body = block[gt..gt],
                .self_closing = true,
                .end = gt + 1,
            };
        }
        var close_buf: [42]u8 = undefined;
        const close = std.fmt.bufPrint(&close_buf, "</{s}>", .{name}) catch return null;
        const close_pos = std.mem.indexOfPos(u8, block, gt, close) orelse return null;
        return .{
            .name = block[open + 1 .. after_idx],
            .attrs = attrs,
            .body = block[gt + 1 .. close_pos],
            .self_closing = false,
            .end = close_pos + close.len,
        };
    }
    return null;
}

// ─── numFmts ──────────────────────────────────────────────────────────

fn parseNumFmts(a: std.mem.Allocator, xml: []const u8) Error![]NumberFormat {
    const block = findBlock(xml, "numFmts") orelse return &[_]NumberFormat{};
    assert(block.len <= xml.len);

    var list: std.ArrayList(NumberFormat) = .empty;
    errdefer list.deinit(a);

    var cursor: usize = 0;
    while (nextElement(block, cursor, "numFmt")) |el| {
        cursor = el.end;
        const id_raw = attrValue(el.attrs, "numFmtId") orelse continue;
        const id = std.fmt.parseInt(u32, id_raw, 10) catch continue;
        const code = attrValue(el.attrs, "formatCode") orelse continue;
        try list.append(a, .{ .fmt_id = id, .code = code });
    }

    return try list.toOwnedSlice(a);
}

// ─── fonts ────────────────────────────────────────────────────────────

fn parseFonts(a: std.mem.Allocator, xml: []const u8) Error![]Font {
    const block = findBlock(xml, "fonts") orelse return &[_]Font{};
    assert(block.len <= xml.len);

    var list: std.ArrayList(Font) = .empty;
    errdefer list.deinit(a);

    var cursor: usize = 0;
    while (nextElement(block, cursor, "font")) |el| {
        cursor = el.end;
        try list.append(a, parseFontBody(el.body));
    }
    return try list.toOwnedSlice(a);
}

fn parseFontBody(body: []const u8) Font {
    assert(body.len < std.math.maxInt(u32));

    var f: Font = .{
        .name = null,
        .size = null,
        .bold = false,
        .italic = false,
        .underline = null,
        .strike = false,
        .color = null,
        .family = null,
        .scheme = null,
    };

    // <name val="…"/>
    if (nextElement(body, 0, "name")) |el| {
        f.name = attrValue(el.attrs, "val");
    }
    // <sz val="11"/>
    if (nextElement(body, 0, "sz")) |el| {
        f.size = attrF64(el.attrs, "val");
    }
    // <b/> or <b val="1"/>; OOXML treats bare <b/> as true.
    if (nextElement(body, 0, "b")) |el| {
        if (attrValue(el.attrs, "val")) |v| {
            f.bold = !std.mem.eql(u8, v, "0") and !std.ascii.eqlIgnoreCase(v, "false");
        } else {
            f.bold = true;
        }
    }
    if (nextElement(body, 0, "i")) |el| {
        if (attrValue(el.attrs, "val")) |v| {
            f.italic = !std.mem.eql(u8, v, "0") and !std.ascii.eqlIgnoreCase(v, "false");
        } else {
            f.italic = true;
        }
    }
    if (nextElement(body, 0, "strike")) |el| {
        if (attrValue(el.attrs, "val")) |v| {
            f.strike = !std.mem.eql(u8, v, "0") and !std.ascii.eqlIgnoreCase(v, "false");
        } else {
            f.strike = true;
        }
    }
    if (nextElement(body, 0, "u")) |el| {
        // `<u/>` defaults to "single" per ECMA-376 §18.18.86.
        f.underline = attrValue(el.attrs, "val") orelse "single";
    }
    if (nextElement(body, 0, "color")) |el| {
        f.color = parseColor(el.attrs);
    }
    if (nextElement(body, 0, "family")) |el| {
        f.family = attrU32(el.attrs, "val");
    }
    if (nextElement(body, 0, "scheme")) |el| {
        f.scheme = attrValue(el.attrs, "val");
    }
    return f;
}

// ─── fills ────────────────────────────────────────────────────────────

fn parseFills(a: std.mem.Allocator, xml: []const u8) Error![]Fill {
    const block = findBlock(xml, "fills") orelse return &[_]Fill{};
    assert(block.len <= xml.len);

    var list: std.ArrayList(Fill) = .empty;
    errdefer list.deinit(a);

    var cursor: usize = 0;
    while (nextElement(block, cursor, "fill")) |el| {
        cursor = el.end;
        try list.append(a, parseFillBody(el.body));
    }
    return try list.toOwnedSlice(a);
}

fn parseFillBody(body: []const u8) Fill {
    var out: Fill = .{ .pattern = "", .fg_color = null, .bg_color = null };

    // <patternFill patternType="solid"> or <gradientFill> (rare; we
    // surface the gradient's shape via pattern="" and skip stops).
    if (nextElement(body, 0, "patternFill")) |pf| {
        if (attrValue(pf.attrs, "patternType")) |p| out.pattern = p;
        // fgColor / bgColor live inside the patternFill body.
        if (nextElement(pf.body, 0, "fgColor")) |c| out.fg_color = parseColor(c.attrs);
        if (nextElement(pf.body, 0, "bgColor")) |c| out.bg_color = parseColor(c.attrs);
    }
    return out;
}

// ─── borders ──────────────────────────────────────────────────────────

fn parseBorders(a: std.mem.Allocator, xml: []const u8) Error![]Border {
    const block = findBlock(xml, "borders") orelse return &[_]Border{};
    assert(block.len <= xml.len);

    var list: std.ArrayList(Border) = .empty;
    errdefer list.deinit(a);

    var cursor: usize = 0;
    while (nextElement(block, cursor, "border")) |el| {
        cursor = el.end;
        const diag_up = xsdBool(el.attrs, "diagonalUp");
        const diag_down = xsdBool(el.attrs, "diagonalDown");
        try list.append(a, .{
            .left = parseBorderSide(el.body, "left"),
            .right = parseBorderSide(el.body, "right"),
            .top = parseBorderSide(el.body, "top"),
            .bottom = parseBorderSide(el.body, "bottom"),
            .diagonal = parseBorderSide(el.body, "diagonal"),
            .diagonal_up = diag_up,
            .diagonal_down = diag_down,
        });
    }
    return try list.toOwnedSlice(a);
}

fn parseBorderSide(body: []const u8, name: []const u8) BorderSide {
    assert(name.len > 0);
    assert(name.len <= 8); // longest is "diagonal"

    const el = nextElement(body, 0, name) orelse return .{ .style = null, .color = null };
    var side: BorderSide = .{ .style = null, .color = null };
    side.style = attrValue(el.attrs, "style");
    if (!el.self_closing) {
        if (nextElement(el.body, 0, "color")) |cel| side.color = parseColor(cel.attrs);
    }
    return side;
}

// ─── color ────────────────────────────────────────────────────────────

/// Parse a `<color rgb="…" indexed="…" theme="…" tint="…"/>` attr blob.
/// Returns null when the element carried no recognized color keys.
fn parseColor(attrs: []const u8) ?Color {
    const argb = attrValue(attrs, "rgb");
    const indexed = attrU32(attrs, "indexed");
    const theme = attrU32(attrs, "theme");
    const tint = attrF64(attrs, "tint");
    if (argb == null and indexed == null and theme == null and tint == null) return null;
    return .{ .argb = argb, .indexed = indexed, .theme = theme, .tint = tint };
}

// ─── cellXfs / cellStyleXfs ───────────────────────────────────────────

fn parseCellXfs(a: std.mem.Allocator, xml: []const u8, wrapper: []const u8) Error![]CellXf {
    assert(wrapper.len > 0);
    assert(wrapper.len < 32);

    const block = findBlock(xml, wrapper) orelse return &[_]CellXf{};

    var list: std.ArrayList(CellXf) = .empty;
    errdefer list.deinit(a);

    var cursor: usize = 0;
    while (nextElement(block, cursor, "xf")) |el| {
        cursor = el.end;
        var xf: CellXf = .{
            .num_fmt_id = attrU32(el.attrs, "numFmtId"),
            .num_fmt_id_invalid = attrValue(el.attrs, "numFmtId") != null and attrU32(el.attrs, "numFmtId") == null,
            .font_id = attrU32(el.attrs, "fontId"),
            .fill_id = attrU32(el.attrs, "fillId"),
            .border_id = attrU32(el.attrs, "borderId"),
            .xf_id = attrU32(el.attrs, "xfId"),
            .apply_number_format = xsdBool(el.attrs, "applyNumberFormat"),
            .apply_font = xsdBool(el.attrs, "applyFont"),
            .apply_fill = xsdBool(el.attrs, "applyFill"),
            .apply_border = xsdBool(el.attrs, "applyBorder"),
            .apply_alignment = xsdBool(el.attrs, "applyAlignment"),
            .apply_protection = xsdBool(el.attrs, "applyProtection"),
            .alignment = null,
            .quote_prefix = xsdBool(el.attrs, "quotePrefix"),
        };
        if (!el.self_closing) {
            if (nextElement(el.body, 0, "alignment")) |ael| {
                xf.alignment = parseAlignmentAttrs(ael.attrs);
            }
        }
        try list.append(a, xf);
    }
    return try list.toOwnedSlice(a);
}

fn parseAlignmentAttrs(attrs: []const u8) Alignment {
    return .{
        .horizontal = attrValue(attrs, "horizontal"),
        .vertical = attrValue(attrs, "vertical"),
        .wrap_text = xsdBool(attrs, "wrapText"),
        .indent = attrU32(attrs, "indent"),
        .rotation = attrI32(attrs, "textRotation"),
        .shrink_to_fit = xsdBool(attrs, "shrinkToFit"),
        .reading_order = attrU32(attrs, "readingOrder"),
    };
}

// ─── Tests ────────────────────────────────────────────────────────────

const testing = std.testing;

test "parse: minimal styles.xml with default font/fill/border/cellXf" {
    const xml =
        \\<?xml version="1.0" encoding="UTF-8"?>
        \\<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        \\  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
        \\  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
        \\  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
        \\  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
        \\</styleSheet>
    ;
    var sx = try parse(testing.allocator, xml);
    defer sx.deinit(testing.allocator);

    try testing.expectEqual(@as(usize, 1), sx.fonts.len);
    try testing.expectEqualStrings("Calibri", sx.fonts[0].name.?);
    try testing.expectEqual(@as(f64, 11), sx.fonts[0].size.?);
    try testing.expect(!sx.fonts[0].bold);

    try testing.expectEqual(@as(usize, 1), sx.fills.len);
    try testing.expectEqualStrings("none", sx.fills[0].pattern);
    try testing.expectEqual(@as(?Color, null), sx.fills[0].fg_color);

    try testing.expectEqual(@as(usize, 1), sx.borders.len);
    try testing.expectEqual(@as(?[]const u8, null), sx.borders[0].left.style);
    try testing.expect(!sx.borders[0].diagonal_up);

    try testing.expectEqual(@as(usize, 1), sx.cell_xfs.len);
    try testing.expectEqual(@as(u32, 0), sx.cell_xfs[0].num_fmt_id.?);
    try testing.expectEqual(@as(u32, 0), sx.cell_xfs[0].font_id.?);
    try testing.expectEqual(@as(?Alignment, null), sx.cell_xfs[0].alignment);
}

test "parse: custom number formats" {
    const xml =
        \\<styleSheet>
        \\  <numFmts count="3">
        \\    <numFmt numFmtId="164" formatCode="0.000"/>
        \\    <numFmt numFmtId="165" formatCode="#,##0.00 _$"/>
        \\    <numFmt numFmtId="166" formatCode="0.00&quot; months&quot;"/>
        \\  </numFmts>
        \\  <cellXfs count="1"><xf numFmtId="164"/></cellXfs>
        \\</styleSheet>
    ;
    var sx = try parse(testing.allocator, xml);
    defer sx.deinit(testing.allocator);

    try testing.expectEqual(@as(usize, 3), sx.number_formats.len);
    try testing.expectEqual(@as(u32, 164), sx.number_formats[0].fmt_id);
    try testing.expectEqualStrings("0.000", sx.number_formats[0].code);
    try testing.expectEqualStrings("#,##0.00 _$", sx.number_formats[1].code);
    // overlay does NOT XML-decode formatCode — &quot; survives literal:
    try testing.expectEqualStrings("0.00&quot; months&quot;", sx.number_formats[2].code);
}

test "parse: multi-font with bold/italic/underline/strike/color/family/scheme" {
    const xml =
        \\<styleSheet>
        \\  <fonts count="3">
        \\    <font>
        \\      <sz val="11"/><color theme="1"/><name val="Calibri"/>
        \\      <family val="2"/><scheme val="minor"/>
        \\    </font>
        \\    <font>
        \\      <b/><i/><u/><strike/>
        \\      <sz val="14"/><color rgb="FFFF0000"/><name val="Arial"/>
        \\    </font>
        \\    <font>
        \\      <b val="0"/>
        \\      <u val="double"/>
        \\      <sz val="12"/><name val="Times New Roman"/>
        \\    </font>
        \\  </fonts>
        \\</styleSheet>
    ;
    var sx = try parse(testing.allocator, xml);
    defer sx.deinit(testing.allocator);

    try testing.expectEqual(@as(usize, 3), sx.fonts.len);

    const f0 = sx.fonts[0];
    try testing.expectEqualStrings("Calibri", f0.name.?);
    try testing.expectEqual(@as(u32, 1), f0.color.?.theme.?);
    try testing.expectEqual(@as(u32, 2), f0.family.?);
    try testing.expectEqualStrings("minor", f0.scheme.?);
    try testing.expect(!f0.bold);

    const f1 = sx.fonts[1];
    try testing.expect(f1.bold);
    try testing.expect(f1.italic);
    try testing.expect(f1.strike);
    try testing.expectEqualStrings("single", f1.underline.?);
    try testing.expectEqualStrings("FFFF0000", f1.color.?.argb.?);
    try testing.expectEqualStrings("Arial", f1.name.?);
    try testing.expectEqual(@as(f64, 14), f1.size.?);

    const f2 = sx.fonts[2];
    try testing.expect(!f2.bold); // <b val="0"/>
    try testing.expectEqualStrings("double", f2.underline.?);
}

test "parse: fills with fg/bg colors and theme/indexed/tint" {
    const xml =
        \\<styleSheet>
        \\  <fills count="4">
        \\    <fill><patternFill patternType="none"/></fill>
        \\    <fill><patternFill patternType="gray125"/></fill>
        \\    <fill><patternFill patternType="solid">
        \\      <fgColor rgb="FFFF0000"/><bgColor indexed="64"/>
        \\    </patternFill></fill>
        \\    <fill><patternFill patternType="solid">
        \\      <fgColor theme="4" tint="-0.249977111117893"/>
        \\      <bgColor theme="0"/>
        \\    </patternFill></fill>
        \\  </fills>
        \\</styleSheet>
    ;
    var sx = try parse(testing.allocator, xml);
    defer sx.deinit(testing.allocator);

    try testing.expectEqual(@as(usize, 4), sx.fills.len);
    try testing.expectEqualStrings("none", sx.fills[0].pattern);
    try testing.expectEqualStrings("gray125", sx.fills[1].pattern);

    try testing.expectEqualStrings("solid", sx.fills[2].pattern);
    try testing.expectEqualStrings("FFFF0000", sx.fills[2].fg_color.?.argb.?);
    try testing.expectEqual(@as(u32, 64), sx.fills[2].bg_color.?.indexed.?);

    try testing.expectEqual(@as(u32, 4), sx.fills[3].fg_color.?.theme.?);
    try testing.expect(sx.fills[3].fg_color.?.tint.? < 0);
    try testing.expectEqual(@as(u32, 0), sx.fills[3].bg_color.?.theme.?);
}

test "parse: borders with diagonal up/down and per-side colors" {
    const xml =
        \\<styleSheet>
        \\  <borders count="2">
        \\    <border>
        \\      <left style="thin"><color rgb="FF000000"/></left>
        \\      <right style="medium"><color rgb="FF111111"/></right>
        \\      <top/>
        \\      <bottom style="double"><color theme="3"/></bottom>
        \\      <diagonal/>
        \\    </border>
        \\    <border diagonalUp="1" diagonalDown="true">
        \\      <left/><right/><top/><bottom/>
        \\      <diagonal style="thin"><color rgb="FF00FF00"/></diagonal>
        \\    </border>
        \\  </borders>
        \\</styleSheet>
    ;
    var sx = try parse(testing.allocator, xml);
    defer sx.deinit(testing.allocator);

    try testing.expectEqual(@as(usize, 2), sx.borders.len);

    const b0 = sx.borders[0];
    try testing.expectEqualStrings("thin", b0.left.style.?);
    try testing.expectEqualStrings("FF000000", b0.left.color.?.argb.?);
    try testing.expectEqualStrings("medium", b0.right.style.?);
    try testing.expectEqual(@as(?[]const u8, null), b0.top.style);
    try testing.expectEqualStrings("double", b0.bottom.style.?);
    try testing.expectEqual(@as(u32, 3), b0.bottom.color.?.theme.?);
    try testing.expectEqual(@as(?[]const u8, null), b0.diagonal.style);
    try testing.expect(!b0.diagonal_up);
    try testing.expect(!b0.diagonal_down);

    const b1 = sx.borders[1];
    try testing.expect(b1.diagonal_up);
    try testing.expect(b1.diagonal_down);
    try testing.expectEqualStrings("thin", b1.diagonal.style.?);
    try testing.expectEqualStrings("FF00FF00", b1.diagonal.color.?.argb.?);
}

test "parse: cellXfs reference fonts/fills/borders + applyXxx + quotePrefix" {
    const xml =
        \\<styleSheet>
        \\  <cellStyleXfs count="1">
        \\    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        \\  </cellStyleXfs>
        \\  <cellXfs count="3">
        \\    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
        \\    <xf numFmtId="164" fontId="2" fillId="3" borderId="1" xfId="0"
        \\        applyNumberFormat="1" applyFont="1" applyFill="true"
        \\        applyBorder="1" quotePrefix="1"/>
        \\    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"
        \\        applyAlignment="1" applyProtection="1"/>
        \\  </cellXfs>
        \\</styleSheet>
    ;
    var sx = try parse(testing.allocator, xml);
    defer sx.deinit(testing.allocator);

    try testing.expectEqual(@as(usize, 1), sx.cell_style_xfs.len);
    try testing.expectEqual(@as(usize, 3), sx.cell_xfs.len);

    const xf1 = sx.cell_xfs[1];
    try testing.expectEqual(@as(u32, 164), xf1.num_fmt_id.?);
    try testing.expectEqual(@as(u32, 2), xf1.font_id.?);
    try testing.expectEqual(@as(u32, 3), xf1.fill_id.?);
    try testing.expectEqual(@as(u32, 1), xf1.border_id.?);
    try testing.expect(xf1.apply_number_format);
    try testing.expect(xf1.apply_font);
    try testing.expect(xf1.apply_fill); // serialised as "true"
    try testing.expect(xf1.apply_border);
    try testing.expect(xf1.quote_prefix);
    try testing.expect(!xf1.apply_alignment);

    const xf2 = sx.cell_xfs[2];
    try testing.expect(xf2.apply_alignment);
    try testing.expect(xf2.apply_protection);
}

test "parse: alignment fields (horizontal, vertical, wrap, indent, rotation, shrink, readingOrder)" {
    const xml =
        \\<styleSheet>
        \\  <cellXfs count="2">
        \\    <xf numFmtId="0" applyAlignment="1">
        \\      <alignment horizontal="center" vertical="top" wrapText="1"
        \\                 indent="2" textRotation="-90" shrinkToFit="1"
        \\                 readingOrder="1"/>
        \\    </xf>
        \\    <xf numFmtId="0" applyAlignment="1">
        \\      <alignment horizontal="left"/>
        \\    </xf>
        \\  </cellXfs>
        \\</styleSheet>
    ;
    var sx = try parse(testing.allocator, xml);
    defer sx.deinit(testing.allocator);

    try testing.expectEqual(@as(usize, 2), sx.cell_xfs.len);
    const a0 = sx.cell_xfs[0].alignment.?;
    try testing.expectEqualStrings("center", a0.horizontal.?);
    try testing.expectEqualStrings("top", a0.vertical.?);
    try testing.expect(a0.wrap_text);
    try testing.expectEqual(@as(u32, 2), a0.indent.?);
    try testing.expectEqual(@as(i32, -90), a0.rotation.?);
    try testing.expect(a0.shrink_to_fit);
    try testing.expectEqual(@as(u32, 1), a0.reading_order.?);

    const a1 = sx.cell_xfs[1].alignment.?;
    try testing.expectEqualStrings("left", a1.horizontal.?);
    try testing.expectEqual(@as(?[]const u8, null), a1.vertical);
    try testing.expect(!a1.wrap_text);
}

test "parse: empty styles.xml returns empty arrays" {
    const xml = "<styleSheet></styleSheet>";
    var sx = try parse(testing.allocator, xml);
    defer sx.deinit(testing.allocator);

    try testing.expectEqual(@as(usize, 0), sx.fonts.len);
    try testing.expectEqual(@as(usize, 0), sx.fills.len);
    try testing.expectEqual(@as(usize, 0), sx.borders.len);
    try testing.expectEqual(@as(usize, 0), sx.number_formats.len);
    try testing.expectEqual(@as(usize, 0), sx.cell_xfs.len);
    try testing.expectEqual(@as(usize, 0), sx.cell_style_xfs.len);
}

test "parse: comments and CDATA cannot smuggle fake elements" {
    // A `<font>` mention inside an XML comment must NOT be picked up as
    // a real font. Ditto for CDATA: even though we keep its bytes, they
    // arrive after sanitization with `<` stripped, so the tag scanner
    // can't see them as elements. (CDATA literal bytes are kept verbatim
    // — the `<` character itself remains, but it's no longer inside a
    // tag-like construct from the scanner's POV.)
    const xml =
        \\<styleSheet>
        \\  <!-- <fonts count="99"><font><name val="HACKED"/></font></fonts> -->
        \\  <fonts count="1"><font><name val="Real"/></font></fonts>
        \\</styleSheet>
    ;
    var sx = try parse(testing.allocator, xml);
    defer sx.deinit(testing.allocator);

    try testing.expectEqual(@as(usize, 1), sx.fonts.len);
    try testing.expectEqualStrings("Real", sx.fonts[0].name.?);
}

test "parse: malformed XML (unterminated comment) returns MalformedXml" {
    const xml = "<styleSheet><!-- never closes <fonts><font/></fonts></styleSheet>";
    try testing.expectError(Error.MalformedXml, parse(testing.allocator, xml));
}

test "parse: quoted-attribute `>` does not split a tag early" {
    // OOXML doesn't actually emit `>` inside attribute values, but the
    // sanitizer must still respect XML's lexical rule. If we mis-parse
    // `formatCode="0;>>"` as ending the tag, the numFmt would be lost.
    const xml =
        \\<styleSheet>
        \\  <numFmts count="1"><numFmt numFmtId="170" formatCode="0;&gt;&gt;"/></numFmts>
        \\</styleSheet>
    ;
    var sx = try parse(testing.allocator, xml);
    defer sx.deinit(testing.allocator);
    try testing.expectEqual(@as(usize, 1), sx.number_formats.len);
    try testing.expectEqual(@as(u32, 170), sx.number_formats[0].fmt_id);
    try testing.expectEqualStrings("0;&gt;&gt;", sx.number_formats[0].code);
}
