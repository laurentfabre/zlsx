//! Minimal pull XML scanner for the oracle extractor (M1b).
//!
//! Independent of `pkg/sheet_xml.zig` / `pkg/sst_xml.zig` for the same
//! reason `zip_reader.zig` is independent of `pkg/zip.zig`: an oracle
//! that shares a decoder with the implementation under test cannot
//! detect a bug in that decoder. See `zip_reader.zig`'s header.
//!
//! Deliberately not a general XML parser. It handles exactly what
//! SpreadsheetML parts contain — elements, attributes, text, comments,
//! processing instructions and CDATA — and it does NOT validate
//! well-formedness, resolve namespaces, expand entities declared in a
//! DTD, or follow external references. Anything it cannot classify it
//! reports, rather than guessing.

const std = @import("std");

pub const Error = error{
    XmlUnterminatedTag,
    XmlUnterminatedComment,
    XmlUnterminatedCdata,
    XmlBadEntity,
    XmlBadXstring,
    OutOfMemory,
};

pub const Element = struct {
    /// Element name exactly as written, prefix included (`x:worksheet`).
    qname: []const u8,
    /// Raw attribute region, between the name and the closing `>`/`/>`.
    attrs: []const u8,

    /// Name with any namespace prefix stripped. The oracle matches on
    /// local names and records the prefix separately: what Excel and
    /// LibreOffice actually write is data we want, not a parse error.
    pub fn local(self: Element) []const u8 {
        if (std.mem.indexOfScalar(u8, self.qname, ':')) |c| return self.qname[c + 1 ..];
        return self.qname;
    }

    pub fn prefix(self: Element) ?[]const u8 {
        if (std.mem.indexOfScalar(u8, self.qname, ':')) |c| return self.qname[0..c];
        return null;
    }

    /// Raw (still entity-encoded) value of the named attribute, matched
    /// on the local name so `r:id` and `id` both answer to "id".
    pub fn attr(self: Element, name: []const u8) ?[]const u8 {
        var i: usize = 0;
        while (i < self.attrs.len) {
            while (i < self.attrs.len and isSpace(self.attrs[i])) : (i += 1) {}
            if (i >= self.attrs.len) return null;

            const name_start = i;
            while (i < self.attrs.len and self.attrs[i] != '=' and !isSpace(self.attrs[i])) : (i += 1) {}
            const raw_name = self.attrs[name_start..i];

            while (i < self.attrs.len and isSpace(self.attrs[i])) : (i += 1) {}
            if (i >= self.attrs.len or self.attrs[i] != '=') return null;
            i += 1;
            while (i < self.attrs.len and isSpace(self.attrs[i])) : (i += 1) {}
            if (i >= self.attrs.len) return null;

            const quote = self.attrs[i];
            if (quote != '"' and quote != '\'') return null;
            i += 1;
            const value_start = i;
            while (i < self.attrs.len and self.attrs[i] != quote) : (i += 1) {}
            const value = self.attrs[value_start..i];
            if (i < self.attrs.len) i += 1;

            const local_name = if (std.mem.indexOfScalar(u8, raw_name, ':')) |c|
                raw_name[c + 1 ..]
            else
                raw_name;
            if (std.mem.eql(u8, local_name, name)) return value;
        }
        return null;
    }
};

pub const Event = union(enum) {
    open: Element,
    self_closing: Element,
    close: []const u8,
    /// Character data between tags, still entity-encoded. CDATA
    /// sections arrive here already unwrapped and must NOT be
    /// entity-decoded — but SpreadsheetML does not use CDATA, so the
    /// scanner reports it and the extractor refuses it rather than
    /// silently conflating the two.
    text: []const u8,
    cdata: []const u8,
};

pub const Scanner = struct {
    xml: []const u8,
    i: usize = 0,

    pub fn init(xml: []const u8) Scanner {
        // A UTF-8 BOM ahead of the declaration is legal and common.
        const start: usize = if (std.mem.startsWith(u8, xml, "\xEF\xBB\xBF")) 3 else 0;
        return .{ .xml = xml, .i = start };
    }

    pub fn next(self: *Scanner) Error!?Event {
        if (self.i >= self.xml.len) return null;

        if (self.xml[self.i] != '<') {
            const start = self.i;
            while (self.i < self.xml.len and self.xml[self.i] != '<') : (self.i += 1) {}
            return .{ .text = self.xml[start..self.i] };
        }

        // `<!-- … -->`, `<![CDATA[ … ]]>`, `<!DOCTYPE …>`
        if (std.mem.startsWith(u8, self.xml[self.i..], "<!--")) {
            const end = std.mem.indexOfPos(u8, self.xml, self.i + 4, "-->") orelse
                return error.XmlUnterminatedComment;
            self.i = end + 3;
            return self.next();
        }
        if (std.mem.startsWith(u8, self.xml[self.i..], "<![CDATA[")) {
            const body = self.i + 9;
            const end = std.mem.indexOfPos(u8, self.xml, body, "]]>") orelse
                return error.XmlUnterminatedCdata;
            const out = self.xml[body..end];
            self.i = end + 3;
            return .{ .cdata = out };
        }
        if (std.mem.startsWith(u8, self.xml[self.i..], "<?") or
            std.mem.startsWith(u8, self.xml[self.i..], "<!"))
        {
            const end = std.mem.indexOfScalarPos(u8, self.xml, self.i, '>') orelse
                return error.XmlUnterminatedTag;
            self.i = end + 1;
            return self.next();
        }

        const tag_end = indexOfTagEnd(self.xml, self.i) orelse return error.XmlUnterminatedTag;
        const inner = self.xml[self.i + 1 .. tag_end];
        self.i = tag_end + 1;

        if (inner.len > 0 and inner[0] == '/') {
            return .{ .close = trimSpace(inner[1..]) };
        }

        const self_closing = inner.len > 0 and inner[inner.len - 1] == '/';
        const body = if (self_closing) inner[0 .. inner.len - 1] else inner;

        var n: usize = 0;
        while (n < body.len and !isSpace(body[n])) : (n += 1) {}
        const element: Element = .{ .qname = body[0..n], .attrs = body[n..] };

        return if (self_closing) .{ .self_closing = element } else .{ .open = element };
    }
};

/// Find the `>` closing the tag that starts at `start`, skipping any
/// that appear inside quoted attribute values. `<c r="A>1"/>` is
/// pathological but legal, and a naive `indexOfScalar` truncates it.
fn indexOfTagEnd(xml: []const u8, start: usize) ?usize {
    var i = start + 1;
    var quote: ?u8 = null;
    while (i < xml.len) : (i += 1) {
        const c = xml[i];
        if (quote) |q| {
            if (c == q) quote = null;
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

inline fn isSpace(c: u8) bool {
    return c == ' ' or c == '\t' or c == '\n' or c == '\r';
}

fn trimSpace(s: []const u8) []const u8 {
    var a: usize = 0;
    var b: usize = s.len;
    while (a < b and isSpace(s[a])) : (a += 1) {}
    while (b > a and isSpace(s[b - 1])) : (b -= 1) {}
    return s[a..b];
}

/// Decode the five predefined XML entities plus numeric character
/// references. This is the ONLY decoding a formula carrier gets — see
/// `goal_formula.md` §5 M4b1: a literal `_x0041_` inside an `<f>` body
/// survives byte-exact, so applying the ST_Xstring pass here would
/// corrupt formula text.
pub fn decodeEntities(allocator: std.mem.Allocator, raw: []const u8) Error![]u8 {
    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(allocator);
    try out.ensureTotalCapacity(allocator, raw.len);

    var i: usize = 0;
    while (i < raw.len) {
        if (raw[i] != '&') {
            try out.append(allocator, raw[i]);
            i += 1;
            continue;
        }
        const semi = std.mem.indexOfScalarPos(u8, raw, i, ';') orelse return error.XmlBadEntity;
        const body = raw[i + 1 .. semi];
        if (body.len == 0) return error.XmlBadEntity;

        if (std.mem.eql(u8, body, "amp")) {
            try out.append(allocator, '&');
        } else if (std.mem.eql(u8, body, "lt")) {
            try out.append(allocator, '<');
        } else if (std.mem.eql(u8, body, "gt")) {
            try out.append(allocator, '>');
        } else if (std.mem.eql(u8, body, "quot")) {
            try out.append(allocator, '"');
        } else if (std.mem.eql(u8, body, "apos")) {
            try out.append(allocator, '\'');
        } else if (body[0] == '#') {
            const cp = parseCharRef(body[1..]) orelse return error.XmlBadEntity;
            var buf: [4]u8 = undefined;
            const n = std.unicode.utf8Encode(cp, &buf) catch return error.XmlBadEntity;
            try out.appendSlice(allocator, buf[0..n]);
        } else {
            // An undeclared entity. Refusing beats passing it through:
            // the oracle would otherwise record `&nbsp;` as literal text
            // and pin a wrong expectation.
            return error.XmlBadEntity;
        }
        i = semi + 1;
    }
    return out.toOwnedSlice(allocator);
}

fn parseCharRef(digits: []const u8) ?u21 {
    if (digits.len == 0) return null;
    const value = if (digits[0] == 'x' or digits[0] == 'X')
        std.fmt.parseInt(u32, digits[1..], 16) catch return null
    else
        std.fmt.parseInt(u32, digits, 10) catch return null;
    if (value > 0x10FFFF) return null;
    if (value >= 0xD800 and value <= 0xDFFF) return null;
    return @intCast(value);
}

/// Decode ST_Xstring's `_xHHHH_` escapes. Applied ONLY to string
/// carriers (SST, inline strings, `t="str"` values) — never to formula
/// text. `_x005F_` is the escape for a literal underscore, so
/// `_x005F_x0041_` decodes to the literal text `_x0041_`.
pub fn decodeXstring(allocator: std.mem.Allocator, s: []const u8) Error![]u8 {
    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(allocator);
    try out.ensureTotalCapacity(allocator, s.len);

    var i: usize = 0;
    while (i < s.len) {
        if (s[i] == '_' and i + 7 <= s.len and s[i + 6] == '_' and isHex4(s[i + 1 .. i + 5]) and
            (s[i + 5] == '_' or isHexDigit(s[i + 5])))
        {
            // Shape is `_xHHHH_`: positions 1..5 are `x` + 4 hex digits.
            if (s[i + 1] == 'x' or s[i + 1] == 'X') {
                const cp_val = std.fmt.parseInt(u32, s[i + 2 .. i + 6], 16) catch {
                    try out.append(allocator, s[i]);
                    i += 1;
                    continue;
                };
                if (cp_val >= 0xD800 and cp_val <= 0xDFFF) return error.XmlBadXstring;
                var buf: [4]u8 = undefined;
                const n = std.unicode.utf8Encode(@intCast(cp_val), &buf) catch
                    return error.XmlBadXstring;
                try out.appendSlice(allocator, buf[0..n]);
                i += 7;
                continue;
            }
        }
        try out.append(allocator, s[i]);
        i += 1;
    }
    return out.toOwnedSlice(allocator);
}

fn isHex4(s: []const u8) bool {
    if (s.len < 4) return false;
    if (s[0] != 'x' and s[0] != 'X') return false;
    for (s[1..4]) |c| {
        if (!isHexDigit(c)) return false;
    }
    return true;
}

inline fn isHexDigit(c: u8) bool {
    return (c >= '0' and c <= '9') or (c >= 'a' and c <= 'f') or (c >= 'A' and c <= 'F');
}

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

test "scans elements, attributes and text" {
    var s: Scanner = .init(
        \\<?xml version="1.0"?><worksheet xmlns="ns"><sheetData><row r="1"><c r="A1" t="n"><v>42</v></c></row></sheetData></worksheet>
    );
    var opens: usize = 0;
    var text_seen: ?[]const u8 = null;
    while (try s.next()) |ev| {
        switch (ev) {
            .open => |e| {
                opens += 1;
                if (std.mem.eql(u8, e.local(), "c")) {
                    try testing.expectEqualStrings("A1", e.attr("r").?);
                    try testing.expectEqualStrings("n", e.attr("t").?);
                    try testing.expect(e.attr("s") == null);
                }
            },
            .text => |t| text_seen = t,
            else => {},
        }
    }
    try testing.expectEqual(@as(usize, 5), opens);
    try testing.expectEqualStrings("42", text_seen.?);
}

test "self-closing elements are distinct from open" {
    var s: Scanner = .init("<a><b/><c d='1' /></a>");
    var self_closing: usize = 0;
    var open: usize = 0;
    var close: usize = 0;
    while (try s.next()) |ev| switch (ev) {
        .self_closing => |e| {
            self_closing += 1;
            if (std.mem.eql(u8, e.local(), "c")) try testing.expectEqualStrings("1", e.attr("d").?);
        },
        .open => open += 1,
        .close => close += 1,
        else => {},
    };
    try testing.expectEqual(@as(usize, 2), self_closing);
    try testing.expectEqual(@as(usize, 1), open);
    try testing.expectEqual(@as(usize, 1), close);
}

test "namespace prefixes are reported, not resolved away" {
    var s: Scanner = .init("<x:worksheet><x:c r=\"A1\"/></x:worksheet>");
    var checked = false;
    while (try s.next()) |ev| switch (ev) {
        .self_closing => |e| {
            try testing.expectEqualStrings("x", e.prefix().?);
            try testing.expectEqualStrings("c", e.local());
            try testing.expectEqualStrings("x:c", e.qname);
            checked = true;
        },
        else => {},
    };
    try testing.expect(checked);
}

test "a `>` inside a quoted attribute does not end the tag" {
    var s: Scanner = .init("<c r=\"A>1\" t='a>b'><v>1</v></c>");
    const ev = (try s.next()).?;
    try testing.expectEqualStrings("A>1", ev.open.attr("r").?);
    try testing.expectEqualStrings("a>b", ev.open.attr("t").?);
}

test "comments, declarations and doctypes are skipped" {
    var s: Scanner = .init("<?xml v='1'?><!DOCTYPE x><!-- <fake/> --><a/>");
    const ev = (try s.next()).?;
    try testing.expectEqualStrings("a", ev.self_closing.qname);
    try testing.expect(try s.next() == null);
}

test "CDATA is reported as its own event, never as text" {
    var s: Scanner = .init("<a><![CDATA[<not a tag>]]></a>");
    _ = try s.next(); // <a>
    const ev = (try s.next()).?;
    try testing.expectEqualStrings("<not a tag>", ev.cdata);
}

test "unterminated constructs are typed errors" {
    var a: Scanner = .init("<a");
    try testing.expectError(error.XmlUnterminatedTag, a.next());
    var b: Scanner = .init("<!-- oops");
    try testing.expectError(error.XmlUnterminatedComment, b.next());
    var c: Scanner = .init("<![CDATA[oops");
    try testing.expectError(error.XmlUnterminatedCdata, c.next());
}

test "decodeEntities handles the predefined five and numeric refs" {
    const cases = [_]struct { in: []const u8, out: []const u8 }{
        .{ .in = "plain", .out = "plain" },
        .{ .in = "a&amp;b", .out = "a&b" },
        .{ .in = "&lt;tag&gt;", .out = "<tag>" },
        .{ .in = "&quot;q&quot;", .out = "\"q\"" },
        .{ .in = "&apos;a&apos;", .out = "'a'" },
        .{ .in = "&#65;&#66;", .out = "AB" },
        .{ .in = "&#x41;&#X42;", .out = "AB" },
        .{ .in = "&#8364;", .out = "€" },
        .{ .in = "&amp;amp;", .out = "&amp;" },
    };
    for (cases) |c| {
        const got = try decodeEntities(testing.allocator, c.in);
        defer testing.allocator.free(got);
        try testing.expectEqualStrings(c.out, got);
    }
}

test "decodeEntities refuses what it cannot decode" {
    // Passing an unknown entity through as literal text would pin a
    // wrong expected value in a manifest.
    for ([_][]const u8{ "&nbsp;", "&", "a&b", "&;", "&#;", "&#xZZ;", "&#1114112;", "&#xD800;" }) |bad| {
        try testing.expectError(error.XmlBadEntity, decodeEntities(testing.allocator, bad));
    }
}

test "decodeXstring handles escapes, including the escaped underscore" {
    const cases = [_]struct { in: []const u8, out: []const u8 }{
        .{ .in = "plain", .out = "plain" },
        .{ .in = "_x0041_", .out = "A" },
        .{ .in = "a_x0042_c", .out = "aBc" },
        .{ .in = "_x000A_", .out = "\n" },
        .{ .in = "_x20AC_", .out = "€" },
        // Not an escape: too short, wrong shape, or no trailing `_`.
        .{ .in = "_x41_", .out = "_x41_" },
        .{ .in = "_y0041_", .out = "_y0041_" },
        .{ .in = "_x0041", .out = "_x0041" },
        .{ .in = "under_score", .out = "under_score" },
    };
    for (cases) |c| {
        const got = try decodeXstring(testing.allocator, c.in);
        defer testing.allocator.free(got);
        try testing.expectEqualStrings(c.out, got);
    }
}

test "decodeXstring: an escaped underscore protects the text after it" {
    // `_x005F_` is the escape for `_`, so `_x005F_x0041_` is the LITERAL
    // text `_x0041_`, not the letter A. Getting this backwards is how a
    // round-trip silently rewrites user data.
    const got = try decodeXstring(testing.allocator, "_x005F_x0041_");
    defer testing.allocator.free(got);
    try testing.expectEqualStrings("_x0041_", got);
}

test "decodeXstring refuses a surrogate escape" {
    try testing.expectError(error.XmlBadXstring, decodeXstring(testing.allocator, "_xD800_"));
}
