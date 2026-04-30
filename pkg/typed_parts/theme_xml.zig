//! Typed-overlay parser for `xl/theme/theme1.xml`.
//!
//! Mirrors what `parseTheme` in `src/xlsx.zig` extracts: the 12-entry
//! `<a:clrScheme>` palette that `<color theme="N"/>` indexes into.
//! Each entry is either `<a:srgbClr val="HEXHEX"/>` or
//! `<a:sysClr val="..." lastClr="HEXHEX"/>`.
//!
//! All string fields borrow from `xml`. Caller guarantees the source
//! buffer outlives the returned `ThemeXml`. The `arena` field is held
//! for parity with sibling typed-overlay modules (workbook/sheet/sst/
//! styles) — this parser does not currently allocate, so it is null.

const std = @import("std");
const assert = std.debug.assert;

pub const ParseError = error{
    MalformedXml,
    OutOfMemory,
};

/// One entry in the `<a:clrScheme>` palette. At most one of `srgb` or
/// `sys_color_value` is non-null; both null means "no color set" — the
/// theme XML omitted this slot. `sys_color_kind` is the `val=` attribute
/// on `<a:sysClr>` (e.g. "windowText", "window") which `parseTheme`
/// surfaces only as the lastClr fallback.
pub const Color = struct {
    srgb: ?[]const u8 = null,
    sys_color_value: ?[]const u8 = null,
    sys_color_kind: ?[]const u8 = null,
};

pub const ColorScheme = struct {
    name: ?[]const u8 = null,
    lt1: Color = .{},
    dk1: Color = .{},
    lt2: Color = .{},
    dk2: Color = .{},
    accent1: Color = .{},
    accent2: Color = .{},
    accent3: Color = .{},
    accent4: Color = .{},
    accent5: Color = .{},
    accent6: Color = .{},
    hlink: Color = .{},
    fol_hlink: Color = .{},
};

pub const ThemeXml = struct {
    color_scheme: ColorScheme,
    arena: ?std.heap.ArenaAllocator,

    pub fn deinit(self: *ThemeXml, allocator: std.mem.Allocator) void {
        _ = allocator;
        if (self.arena) |*a| a.deinit();
        self.arena = null;
    }
};

/// Parse `xml` into a `ThemeXml`. All `[]const u8` fields borrow from
/// `xml`. On `MalformedXml` no partial state is leaked. The defensive
/// scan is comment / CDATA / processing-instruction aware so attribute
/// `>` characters inside quoted strings or stray `<!-- > -->` comments
/// don't terminate elements early.
pub fn parse(allocator: std.mem.Allocator, xml: []const u8) ParseError!ThemeXml {
    _ = allocator; // reserved for future allocations
    assert(xml.len <= std.math.maxInt(u32));

    // Reject obvious junk early — empty input is malformed (a real
    // theme1.xml carries an `<?xml` prolog plus a `<a:theme>` root).
    if (xml.len == 0) return error.MalformedXml;
    assert(xml.len > 0);

    // The clrScheme element is optional in theory but always present in
    // well-formed Office-emitted theme1.xml. When absent, return all
    // defaults rather than failing — matches `parseTheme`'s behavior.
    const scheme_open_pos = findElementOpen(xml, "a:clrScheme") orelse {
        return .{ .color_scheme = .{}, .arena = null };
    };
    const scheme_close_pos = std.mem.indexOfPos(u8, xml, scheme_open_pos, "</a:clrScheme>") orelse
        return error.MalformedXml;
    assert(scheme_close_pos > scheme_open_pos);

    // Slice covering `<a:clrScheme ...>...` up to (but not including)
    // the closing tag — child element scans are bounded by this window.
    const scheme = xml[scheme_open_pos..scheme_close_pos];
    assert(scheme.len > 0);

    var cs: ColorScheme = .{};
    cs.name = readNameAttr(scheme);

    // Schema-order children of <a:clrScheme>. Index correspondence is
    // documented in the parseTheme comment in src/xlsx.zig:3217.
    const Slot = struct { tag: []const u8, field: []const u8 };
    const slots = [_]Slot{
        .{ .tag = "a:dk1", .field = "dk1" },
        .{ .tag = "a:lt1", .field = "lt1" },
        .{ .tag = "a:dk2", .field = "dk2" },
        .{ .tag = "a:lt2", .field = "lt2" },
        .{ .tag = "a:accent1", .field = "accent1" },
        .{ .tag = "a:accent2", .field = "accent2" },
        .{ .tag = "a:accent3", .field = "accent3" },
        .{ .tag = "a:accent4", .field = "accent4" },
        .{ .tag = "a:accent5", .field = "accent5" },
        .{ .tag = "a:accent6", .field = "accent6" },
        .{ .tag = "a:hlink", .field = "hlink" },
        .{ .tag = "a:folHlink", .field = "fol_hlink" },
    };

    inline for (slots) |s| {
        if (try findChildBody(scheme, s.tag)) |body| {
            const c = try extractColor(body);
            @field(cs, s.field) = c;
        }
    }

    return .{ .color_scheme = cs, .arena = null };
}

// ─── Internals ───────────────────────────────────────────────────────

/// Find the byte offset of `<tag` such that the next char is one of
/// ` `, `\t`, `\n`, `\r`, `>`, `/` — i.e. the actual element open, not
/// a tag whose name is a prefix of another (e.g. `a:dk1` vs nothing,
/// but `a:accent1` vs `a:accent10` — paranoid future-proofing).
fn findElementOpen(xml: []const u8, tag: []const u8) ?usize {
    assert(tag.len > 0);
    assert(xml.len <= std.math.maxInt(u32));
    var search_buf: [64]u8 = undefined;
    if (tag.len + 1 > search_buf.len) return null;
    search_buf[0] = '<';
    @memcpy(search_buf[1 .. 1 + tag.len], tag);
    const needle = search_buf[0 .. 1 + tag.len];
    var i: usize = 0;
    while (std.mem.indexOfPos(u8, xml, i, needle)) |pos| {
        const after = pos + needle.len;
        if (after >= xml.len) return null;
        const c = xml[after];
        if (c == ' ' or c == '\t' or c == '\n' or c == '\r' or c == '>' or c == '/') {
            return pos;
        }
        i = after;
    }
    return null;
}

/// Locate `<tag ...>...</tag>` or `<tag .../>` within `xml`, return the
/// inner body (between `>` and `</tag>`) or null when absent. For
/// self-closing forms the body is empty (`""`).
fn findChildBody(xml: []const u8, tag: []const u8) ParseError!?[]const u8 {
    assert(tag.len > 0);
    assert(tag.len < 32);
    const open_pos = findElementOpen(xml, tag) orelse return null;
    // Walk forward past the open tag, honoring quoted attribute values
    // so a stray `>` inside `name="foo>bar"` doesn't fool us.
    const tag_end = scanPastTag(xml, open_pos) orelse return error.MalformedXml;
    assert(tag_end > open_pos);
    assert(tag_end <= xml.len);
    // Self-closing? `<tag ... />` — last byte before `>` is `/`.
    if (tag_end >= 2 and xml[tag_end - 2] == '/') return xml[tag_end..tag_end];

    // Build `</tag>` and find the close. Use a fixed buffer — tag is
    // bounded by the assertion above.
    var close_buf: [40]u8 = undefined;
    close_buf[0] = '<';
    close_buf[1] = '/';
    @memcpy(close_buf[2 .. 2 + tag.len], tag);
    close_buf[2 + tag.len] = '>';
    const close_needle = close_buf[0 .. 3 + tag.len];
    const close_pos = std.mem.indexOfPos(u8, xml, tag_end, close_needle) orelse
        return error.MalformedXml;
    assert(close_pos >= tag_end);
    return xml[tag_end..close_pos];
}

/// Walk past a `<...>` opening tag whose first byte is at `start`,
/// honoring quoted attribute values. Returns the index just past `>`.
/// Also skips `<!-- ... -->`, `<![CDATA[ ... ]]>`, and `<? ... ?>`
/// constructs if those happen to begin at `start` — but the caller
/// only invokes this on element opens located by `findElementOpen`, so
/// those branches are defense-in-depth.
fn scanPastTag(xml: []const u8, start: usize) ?usize {
    assert(start < xml.len);
    assert(xml[start] == '<');
    // Comment.
    if (std.mem.startsWith(u8, xml[start..], "<!--")) {
        const end = std.mem.indexOfPos(u8, xml, start + 4, "-->") orelse return null;
        return end + 3;
    }
    // CDATA.
    if (std.mem.startsWith(u8, xml[start..], "<![CDATA[")) {
        const end = std.mem.indexOfPos(u8, xml, start + 9, "]]>") orelse return null;
        return end + 3;
    }
    // Processing instruction.
    if (start + 1 < xml.len and xml[start + 1] == '?') {
        const end = std.mem.indexOfPos(u8, xml, start + 2, "?>") orelse return null;
        return end + 2;
    }
    // Element open with attribute-aware quote handling.
    var i: usize = start + 1;
    while (i < xml.len) {
        const c = xml[i];
        if (c == '"' or c == '\'') {
            const q_end = std.mem.indexOfScalarPos(u8, xml, i + 1, c) orelse return null;
            i = q_end + 1;
            continue;
        }
        if (c == '>') return i + 1;
        i += 1;
    }
    return null;
}

/// Read `name="..."` (or `name='...'`) from the `<a:clrScheme ...>`
/// open tag. Borrows from `xml`. Returns null when the attribute is
/// absent or the tag close can't be located.
fn readNameAttr(xml: []const u8) ?[]const u8 {
    assert(xml.len > 0);
    const tag_end = scanPastTag(xml, 0) orelse return null;
    assert(tag_end > 0);
    const tag = xml[0..tag_end];
    return attr(tag, "name");
}

/// Find `key="..."` or `key='...'` in `attrs`. Borrows. Returns null
/// when the key isn't present or the value is unterminated.
fn attr(attrs: []const u8, key: []const u8) ?[]const u8 {
    assert(key.len > 0);
    assert(key.len < 32);
    var search_buf: [40]u8 = undefined;
    @memcpy(search_buf[0..key.len], key);
    search_buf[key.len] = '=';
    // Try double-quote first — overwhelmingly common in OOXML.
    const quotes = [_]u8{ '"', '\'' };
    for (quotes) |q| {
        search_buf[key.len + 1] = q;
        const needle = search_buf[0 .. key.len + 2];
        if (std.mem.indexOf(u8, attrs, needle)) |pos| {
            const start = pos + needle.len;
            if (std.mem.indexOfScalarPos(u8, attrs, start, q)) |close| {
                return attrs[start..close];
            }
        }
    }
    return null;
}

/// Extract a `Color` from the body of a clrScheme child element. The
/// body wraps either `<a:srgbClr val="HEX"/>` or
/// `<a:sysClr val="kind" lastClr="HEX"/>`. When neither is present
/// returns the default `Color{}` ("no color set").
fn extractColor(body: []const u8) ParseError!Color {
    var c: Color = .{};

    if (findElementOpen(body, "a:srgbClr")) |pos| {
        const tag_end = scanPastTag(body, pos) orelse return error.MalformedXml;
        assert(tag_end > pos);
        const tag = body[pos..tag_end];
        c.srgb = attr(tag, "val");
        return c;
    }

    if (findElementOpen(body, "a:sysClr")) |pos| {
        const tag_end = scanPastTag(body, pos) orelse return error.MalformedXml;
        assert(tag_end > pos);
        const tag = body[pos..tag_end];
        c.sys_color_kind = attr(tag, "val");
        c.sys_color_value = attr(tag, "lastClr");
        return c;
    }

    return c;
}

// ─── Tests ───────────────────────────────────────────────────────────

test "parse: minimal theme with srgbClr scheme" {
    const xml =
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
        \\  <a:themeElements>
        \\    <a:clrScheme name="Office">
        \\      <a:dk1><a:srgbClr val="000000"/></a:dk1>
        \\      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
        \\      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
        \\      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
        \\      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
        \\      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
        \\      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
        \\      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
        \\      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
        \\      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
        \\      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
        \\      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
        \\    </a:clrScheme>
        \\  </a:themeElements>
        \\</a:theme>
    ;

    var t = try parse(std.testing.allocator, xml);
    defer t.deinit(std.testing.allocator);

    try std.testing.expectEqualStrings("Office", t.color_scheme.name.?);
    try std.testing.expectEqualStrings("000000", t.color_scheme.dk1.srgb.?);
    try std.testing.expectEqualStrings("FFFFFF", t.color_scheme.lt1.srgb.?);
    try std.testing.expectEqualStrings("4472C4", t.color_scheme.accent1.srgb.?);
    try std.testing.expectEqualStrings("70AD47", t.color_scheme.accent6.srgb.?);
    try std.testing.expectEqualStrings("0563C1", t.color_scheme.hlink.srgb.?);
    try std.testing.expectEqualStrings("954F72", t.color_scheme.fol_hlink.srgb.?);
    try std.testing.expect(t.color_scheme.dk1.sys_color_value == null);
    try std.testing.expect(t.color_scheme.dk1.sys_color_kind == null);
}

test "parse: scheme using sysClr (windowText / window)" {
    const xml =
        \\<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        \\  <a:clrScheme name="Office">
        \\    <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
        \\    <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
        \\    <a:dk2><a:srgbClr val="44546A"/></a:dk2>
        \\    <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
        \\  </a:clrScheme>
        \\</a:theme>
    ;

    var t = try parse(std.testing.allocator, xml);
    defer t.deinit(std.testing.allocator);

    try std.testing.expectEqualStrings("windowText", t.color_scheme.dk1.sys_color_kind.?);
    try std.testing.expectEqualStrings("000000", t.color_scheme.dk1.sys_color_value.?);
    try std.testing.expect(t.color_scheme.dk1.srgb == null);

    try std.testing.expectEqualStrings("window", t.color_scheme.lt1.sys_color_kind.?);
    try std.testing.expectEqualStrings("FFFFFF", t.color_scheme.lt1.sys_color_value.?);

    try std.testing.expectEqualStrings("44546A", t.color_scheme.dk2.srgb.?);
    try std.testing.expect(t.color_scheme.dk2.sys_color_kind == null);
}

test "parse: missing entries default to no color set" {
    // No clrScheme at all → all-defaults ColorScheme.
    const xml_no_scheme =
        \\<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        \\  <a:themeElements/>
        \\</a:theme>
    ;
    var t1 = try parse(std.testing.allocator, xml_no_scheme);
    defer t1.deinit(std.testing.allocator);
    try std.testing.expect(t1.color_scheme.name == null);
    try std.testing.expect(t1.color_scheme.dk1.srgb == null);
    try std.testing.expect(t1.color_scheme.dk1.sys_color_value == null);
    try std.testing.expect(t1.color_scheme.accent6.srgb == null);

    // Partial scheme — only dk1 + lt1 set; the other ten slots remain
    // at their defaults (both color members null).
    const xml_partial =
        \\<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        \\  <a:clrScheme name="Partial">
        \\    <a:dk1><a:srgbClr val="111111"/></a:dk1>
        \\    <a:lt1><a:srgbClr val="EEEEEE"/></a:lt1>
        \\  </a:clrScheme>
        \\</a:theme>
    ;
    var t2 = try parse(std.testing.allocator, xml_partial);
    defer t2.deinit(std.testing.allocator);
    try std.testing.expectEqualStrings("Partial", t2.color_scheme.name.?);
    try std.testing.expectEqualStrings("111111", t2.color_scheme.dk1.srgb.?);
    try std.testing.expectEqualStrings("EEEEEE", t2.color_scheme.lt1.srgb.?);
    try std.testing.expect(t2.color_scheme.dk2.srgb == null);
    try std.testing.expect(t2.color_scheme.accent1.srgb == null);
    try std.testing.expect(t2.color_scheme.fol_hlink.srgb == null);
}

test "parse: malformed input rejected" {
    // Empty buffer is malformed — a real theme1.xml has at minimum a
    // prolog plus the a:theme root.
    try std.testing.expectError(error.MalformedXml, parse(std.testing.allocator, ""));

    // clrScheme open without a closing tag.
    const xml_unterminated_scheme =
        \\<a:theme><a:clrScheme name="Bad">
        \\  <a:dk1><a:srgbClr val="000000"/></a:dk1>
    ;
    try std.testing.expectError(
        error.MalformedXml,
        parse(std.testing.allocator, xml_unterminated_scheme),
    );

    // srgbClr open with no terminator inside the scheme body.
    const xml_unterminated_color =
        \\<a:theme><a:clrScheme name="Bad">
        \\  <a:dk1><a:srgbClr val="000000"
        \\</a:clrScheme></a:theme>
    ;
    try std.testing.expectError(
        error.MalformedXml,
        parse(std.testing.allocator, xml_unterminated_color),
    );
}

test "parse: comment / CDATA / quoted-attribute defenses" {
    // A comment containing `>` inside the scheme must not terminate
    // any element. A quoted `>` in an attribute likewise must not.
    const xml =
        \\<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        \\  <a:clrScheme name="A>B">
        \\    <!-- a stray > inside a comment -->
        \\    <a:dk1><a:srgbClr val="010203"/></a:dk1>
        \\  </a:clrScheme>
        \\</a:theme>
    ;
    var t = try parse(std.testing.allocator, xml);
    defer t.deinit(std.testing.allocator);

    try std.testing.expectEqualStrings("A>B", t.color_scheme.name.?);
    try std.testing.expectEqualStrings("010203", t.color_scheme.dk1.srgb.?);
}
