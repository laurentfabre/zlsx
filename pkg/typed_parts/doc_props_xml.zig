//! Typed-overlay parser + scrubber for `docProps/core.xml` and
//! `docProps/app.xml`.
//!
//! These two parts carry the workbook's authorship metadata, and zlsx
//! was previously blind to them: they round-tripped byte-for-byte
//! through every edit, carrying `dc:creator`, `cp:lastModifiedBy`,
//! `Company` and friends along invisibly. For a pseudonymisation
//! pipeline that is a silent PII leak — the cells get masked and the
//! author's name rides out untouched in the archive.
//!
//! Two capabilities live here:
//!
//!   * **Read** — `parseCore` / `parseApp` fill a `DocProps` view.
//!   * **Scrub** — `scrubCore` / `scrubApp` emit a copy of the part
//!     with masked elements *removed entirely*, every other byte
//!     preserved. Removal beats blanking: an empty `<dc:creator/>`
//!     still says "this document had an author", and some tools
//!     round-trip the empty element back into a populated one.
//!
//! All `DocProps` string fields borrow from the source `xml`, matching
//! the sibling typed-overlay modules. The caller guarantees the source
//! buffer outlives the view. Scrub output is allocator-owned.
//!
//! Deliberately tolerant: these parts are optional in OOXML, frequently
//! hand-rolled by third-party writers, and never load-bearing for cell
//! data. A malformed core.xml must never fail a workbook open — every
//! field is optional and an unparseable element yields null rather than
//! an error.

const std = @import("std");
const assert = std.debug.assert;

pub const Error = error{
    OutOfMemory,
};

/// Canonical part names. Exposed so callers don't restate the strings.
pub const core_part_name = "docProps/core.xml";
pub const app_part_name = "docProps/app.xml";
pub const custom_part_name = "docProps/custom.xml";

/// Typed view over the workbook's document properties.
///
/// Every field is optional because every underlying element is
/// optional. `null` means "absent from the source", which is distinct
/// from present-but-empty (`""`).
pub const DocProps = struct {
    // ── docProps/core.xml ──────────────────────────────────────────
    /// `dc:creator` — original author. Prime PII field.
    creator: ?[]const u8 = null,
    /// `cp:lastModifiedBy` — last editor. Prime PII field.
    last_modified_by: ?[]const u8 = null,
    title: ?[]const u8 = null,
    subject: ?[]const u8 = null,
    /// `dc:description` (Excel surfaces this as "Comments").
    description: ?[]const u8 = null,
    keywords: ?[]const u8 = null,
    category: ?[]const u8 = null,
    /// `dcterms:created` — W3CDTF timestamp, kept verbatim.
    created: ?[]const u8 = null,
    /// `dcterms:modified` — W3CDTF timestamp, kept verbatim.
    modified: ?[]const u8 = null,
    revision: ?[]const u8 = null,

    // ── docProps/app.xml ───────────────────────────────────────────
    company: ?[]const u8 = null,
    manager: ?[]const u8 = null,
    /// Producing application, e.g. "Microsoft Excel" or
    /// "LibreOffice/4.3.3.2$Linux_X86_64…". Fingerprints the toolchain
    /// that touched the file, which is why the scrub mask can drop it.
    application: ?[]const u8 = null,
    hyperlink_base: ?[]const u8 = null,

    // ── docProps/custom.xml ────────────────────────────────────────
    /// True when the archive carries a custom-properties part. Its
    /// contents are arbitrary user-defined key/value pairs, so zlsx
    /// does not model them — it only reports presence and can drop the
    /// part wholesale.
    has_custom_properties: bool = false,

    /// True when nothing at all was found. Lets callers distinguish
    /// "no docProps parts" from "parts present but empty".
    pub fn isEmpty(self: DocProps) bool {
        return self.creator == null and self.last_modified_by == null and
            self.title == null and self.subject == null and
            self.description == null and self.keywords == null and
            self.category == null and self.created == null and
            self.modified == null and self.revision == null and
            self.company == null and self.manager == null and
            self.application == null and self.hyperlink_base == null and
            !self.has_custom_properties;
    }

    /// True when any field the default scrub mask targets is populated.
    /// Callers use this to decide whether a scrub would change anything.
    pub fn hasIdentifyingFields(self: DocProps) bool {
        return self.creator != null or self.last_modified_by != null or
            self.company != null or self.manager != null or
            self.title != null or self.subject != null or
            self.description != null or self.keywords != null or
            self.has_custom_properties;
    }
};

/// Which fields `scrubCore` / `scrubApp` remove.
///
/// Defaults target everything that identifies a person, an organisation
/// or a machine. Timestamps (`created` / `modified`) and `revision` are
/// deliberately NOT scrubbed by default: they are rarely identifying on
/// their own, and stripping them visibly breaks Excel's document-info
/// pane. Callers who want them gone can flip the flags.
pub const Mask = struct {
    creator: bool = true,
    last_modified_by: bool = true,
    title: bool = true,
    subject: bool = true,
    description: bool = true,
    keywords: bool = true,
    category: bool = true,
    company: bool = true,
    manager: bool = true,
    application: bool = false,
    hyperlink_base: bool = true,
    /// Drop `docProps/custom.xml` wholesale, along with its
    /// `[Content_Types].xml` override and its relationship entry.
    custom_properties: bool = true,

    /// Timestamps and revision counter. Off by default — see above.
    created: bool = false,
    modified: bool = false,
    revision: bool = false,

    /// Everything, including timestamps. For callers who want the
    /// archive to carry no document metadata at all.
    pub const all: Mask = .{
        .application = true,
        .created = true,
        .modified = true,
        .revision = true,
    };

    /// Nothing. Useful as a base to enable single fields from.
    pub const none: Mask = .{
        .creator = false,
        .last_modified_by = false,
        .title = false,
        .subject = false,
        .description = false,
        .keywords = false,
        .category = false,
        .company = false,
        .manager = false,
        .application = false,
        .hyperlink_base = false,
        .custom_properties = false,
    };
};

// ─── Read ────────────────────────────────────────────────────────────

/// Parse `docProps/core.xml` into the core-side fields of `out`.
/// Never fails on malformed input — absent or unreadable elements stay
/// null, because document properties must not be able to break an open.
pub fn parseCore(xml: []const u8, out: *DocProps) void {
    out.creator = elementText(xml, "dc:creator");
    out.last_modified_by = elementText(xml, "cp:lastModifiedBy");
    out.title = elementText(xml, "dc:title");
    out.subject = elementText(xml, "dc:subject");
    out.description = elementText(xml, "dc:description");
    out.keywords = elementText(xml, "cp:keywords");
    out.category = elementText(xml, "cp:category");
    out.created = elementText(xml, "dcterms:created");
    out.modified = elementText(xml, "dcterms:modified");
    out.revision = elementText(xml, "cp:revision");
}

/// Parse `docProps/app.xml` into the app-side fields of `out`.
pub fn parseApp(xml: []const u8, out: *DocProps) void {
    out.company = elementText(xml, "Company");
    out.manager = elementText(xml, "Manager");
    out.application = elementText(xml, "Application");
    out.hyperlink_base = elementText(xml, "HyperlinkBase");
}

/// Convenience: build a full view from whichever parts exist.
pub fn parse(
    core_xml: ?[]const u8,
    app_xml: ?[]const u8,
    has_custom: bool,
) DocProps {
    var out: DocProps = .{ .has_custom_properties = has_custom };
    if (core_xml) |x| parseCore(x, &out);
    if (app_xml) |x| parseApp(x, &out);
    return out;
}

// ─── Scrub ───────────────────────────────────────────────────────────

/// Emit `xml` with every element the mask targets removed. Returns
/// allocator-owned bytes; caller frees.
///
/// Byte-preserving for everything untouched: the output is the input
/// with whole `<tag>…</tag>` (or `<tag/>`) spans excised, so namespace
/// declarations, attribute order, and any element zlsx does not model
/// survive verbatim.
pub fn scrubCore(
    allocator: std.mem.Allocator,
    xml: []const u8,
    mask: Mask,
) Error![]u8 {
    var tags: std.ArrayListUnmanaged([]const u8) = .empty;
    defer tags.deinit(allocator);
    if (mask.creator) try tags.append(allocator, "dc:creator");
    if (mask.last_modified_by) try tags.append(allocator, "cp:lastModifiedBy");
    if (mask.title) try tags.append(allocator, "dc:title");
    if (mask.subject) try tags.append(allocator, "dc:subject");
    if (mask.description) try tags.append(allocator, "dc:description");
    if (mask.keywords) try tags.append(allocator, "cp:keywords");
    if (mask.category) try tags.append(allocator, "cp:category");
    if (mask.created) try tags.append(allocator, "dcterms:created");
    if (mask.modified) try tags.append(allocator, "dcterms:modified");
    if (mask.revision) try tags.append(allocator, "cp:revision");
    return removeElements(allocator, xml, tags.items);
}

/// Emit `docProps/app.xml` with masked elements removed.
pub fn scrubApp(
    allocator: std.mem.Allocator,
    xml: []const u8,
    mask: Mask,
) Error![]u8 {
    var tags: std.ArrayListUnmanaged([]const u8) = .empty;
    defer tags.deinit(allocator);
    if (mask.company) try tags.append(allocator, "Company");
    if (mask.manager) try tags.append(allocator, "Manager");
    if (mask.application) try tags.append(allocator, "Application");
    if (mask.hyperlink_base) try tags.append(allocator, "HyperlinkBase");
    return removeElements(allocator, xml, tags.items);
}

// ─── Internals ───────────────────────────────────────────────────────

/// Text content of the first `<tag …>text</tag>`, or null when the
/// element is absent, self-closing, or unterminated. Borrows from `xml`.
///
/// Returns an empty slice (not null) for `<tag></tag>` so callers can
/// tell "present but blank" from "absent" — that distinction decides
/// whether a scrub actually has work to do.
fn elementText(xml: []const u8, tag: []const u8) ?[]const u8 {
    const span = findElement(xml, tag) orelse return null;
    if (span.self_closing) return null;
    return xml[span.body_start..span.body_end];
}

const ElementSpan = struct {
    /// Index of the opening `<`.
    start: usize,
    /// One past the closing `>`.
    end: usize,
    body_start: usize,
    body_end: usize,
    self_closing: bool,
};

/// Locate `<tag …>` … `</tag>` (or `<tag …/>`).
///
/// Matches the tag name exactly, requiring the next byte after the name
/// to be whitespace, `>` or `/` — otherwise `dc:title` would match
/// `dc:titleAlt`.
fn findElement(xml: []const u8, tag: []const u8) ?ElementSpan {
    var i: usize = 0;
    while (i < xml.len) {
        const lt = std.mem.indexOfScalarPos(u8, xml, i, '<') orelse return null;
        const name_start = lt + 1;
        if (name_start + tag.len > xml.len) return null;

        if (std.mem.eql(u8, xml[name_start .. name_start + tag.len], tag)) {
            const after = name_start + tag.len;
            if (after < xml.len) {
                const c = xml[after];
                if (c == '>' or c == '/' or c == ' ' or c == '\t' or c == '\n' or c == '\r') {
                    const gt = std.mem.indexOfScalarPos(u8, xml, after, '>') orelse return null;
                    // `<tag …/>` — no body.
                    if (xml[gt - 1] == '/') {
                        return .{
                            .start = lt,
                            .end = gt + 1,
                            .body_start = gt + 1,
                            .body_end = gt + 1,
                            .self_closing = true,
                        };
                    }
                    // Find the matching close tag.
                    const body_start = gt + 1;
                    const close_needle = "</";
                    var search = body_start;
                    while (search < xml.len) {
                        const close = std.mem.indexOfPos(u8, xml, search, close_needle) orelse return null;
                        const cname = close + close_needle.len;
                        if (cname + tag.len <= xml.len and
                            std.mem.eql(u8, xml[cname .. cname + tag.len], tag))
                        {
                            const cgt = std.mem.indexOfScalarPos(u8, xml, cname, '>') orelse return null;
                            return .{
                                .start = lt,
                                .end = cgt + 1,
                                .body_start = body_start,
                                .body_end = close,
                                .self_closing = false,
                            };
                        }
                        search = close + close_needle.len;
                    }
                    return null;
                }
            }
        }
        i = lt + 1;
    }
    return null;
}

/// Copy `xml`, dropping every span named in `tags`. Repeats per tag
/// until no further match, so duplicated elements are all removed.
fn removeElements(
    allocator: std.mem.Allocator,
    xml: []const u8,
    tags: []const []const u8,
) Error![]u8 {
    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(allocator);
    try out.appendSlice(allocator, xml);

    for (tags) |tag| {
        // Loop: an element may legally appear more than once, and a
        // scrub that removed only the first would leak the rest.
        while (findElement(out.items, tag)) |span| {
            out.replaceRange(allocator, span.start, span.end - span.start, &.{}) catch |err| switch (err) {
                error.OutOfMemory => return error.OutOfMemory,
            };
        }
    }
    return out.toOwnedSlice(allocator);
}

// ─── Tests ───────────────────────────────────────────────────────────

const core_sample =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"><dcterms:created xsi:type="dcterms:W3CDTF">2015-06-29T06:57:34Z</dcterms:created><dc:creator>Jane Q. Fixture</dc:creator><dc:title>Quarterly</dc:title><cp:lastModifiedBy>Jane Q. Fixture</cp:lastModifiedBy><dcterms:modified xsi:type="dcterms:W3CDTF">2015-06-29T06:59:03Z</dcterms:modified><cp:revision>1</cp:revision></cp:coreProperties>
;

const app_sample =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><TotalTime>89</TotalTime><Application>LibreOffice/4.3.3.2</Application><Company>AcmeCorp</Company><Manager>Bob Boss</Manager></Properties>
;

test "parseCore: extracts identifying fields" {
    var props: DocProps = .{};
    parseCore(core_sample, &props);

    try std.testing.expectEqualStrings("Jane Q. Fixture", props.creator.?);
    try std.testing.expectEqualStrings("Jane Q. Fixture", props.last_modified_by.?);
    try std.testing.expectEqualStrings("Quarterly", props.title.?);
    try std.testing.expectEqualStrings("2015-06-29T06:57:34Z", props.created.?);
    try std.testing.expectEqualStrings("1", props.revision.?);
    // Absent elements stay null rather than empty.
    try std.testing.expect(props.subject == null);
    try std.testing.expect(props.keywords == null);
}

test "parseApp: extracts company and manager" {
    var props: DocProps = .{};
    parseApp(app_sample, &props);

    try std.testing.expectEqualStrings("AcmeCorp", props.company.?);
    try std.testing.expectEqualStrings("Bob Boss", props.manager.?);
    try std.testing.expectEqualStrings("LibreOffice/4.3.3.2", props.application.?);
    try std.testing.expect(props.hyperlink_base == null);
}

test "scrubCore: removes identifying fields, keeps timestamps" {
    const out = try scrubCore(std.testing.allocator, core_sample, .{});
    defer std.testing.allocator.free(out);

    try std.testing.expect(std.mem.indexOf(u8, out, "Jane Q. Fixture") == null);
    try std.testing.expect(std.mem.indexOf(u8, out, "Quarterly") == null);
    // Not masked by default — timestamps and revision survive.
    try std.testing.expect(std.mem.indexOf(u8, out, "2015-06-29T06:57:34Z") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "<cp:revision>1</cp:revision>") != null);
    // The document element and its namespaces are untouched.
    try std.testing.expect(std.mem.indexOf(u8, out, "cp:coreProperties") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "xmlns:dcterms") != null);

    // And the result still parses, with the scrubbed fields now absent.
    var after: DocProps = .{};
    parseCore(out, &after);
    try std.testing.expect(after.creator == null);
    try std.testing.expect(after.last_modified_by == null);
    try std.testing.expect(after.title == null);
    try std.testing.expectEqualStrings("1", after.revision.?);
}

test "scrubCore: Mask.all also removes timestamps" {
    const out = try scrubCore(std.testing.allocator, core_sample, Mask.all);
    defer std.testing.allocator.free(out);

    try std.testing.expect(std.mem.indexOf(u8, out, "2015-06-29T06:57:34Z") == null);
    try std.testing.expect(std.mem.indexOf(u8, out, "cp:revision") == null);
    try std.testing.expect(std.mem.indexOf(u8, out, "cp:coreProperties") != null);
}

test "scrubCore: Mask.none is a byte-identical copy" {
    const out = try scrubCore(std.testing.allocator, core_sample, Mask.none);
    defer std.testing.allocator.free(out);
    try std.testing.expectEqualStrings(core_sample, out);
}

test "scrubApp: removes company and manager, keeps TotalTime" {
    const out = try scrubApp(std.testing.allocator, app_sample, .{});
    defer std.testing.allocator.free(out);

    try std.testing.expect(std.mem.indexOf(u8, out, "AcmeCorp") == null);
    try std.testing.expect(std.mem.indexOf(u8, out, "Bob Boss") == null);
    // Application is off by default; TotalTime is not modelled at all.
    try std.testing.expect(std.mem.indexOf(u8, out, "LibreOffice/4.3.3.2") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "<TotalTime>89</TotalTime>") != null);
}

test "removeElements: strips every occurrence, not just the first" {
    const dup = "<r><dc:creator>A</dc:creator><x/><dc:creator>B</dc:creator></r>";
    const out = try scrubCore(std.testing.allocator, dup, .{});
    defer std.testing.allocator.free(out);
    try std.testing.expectEqualStrings("<r><x/></r>", out);
}

test "findElement: does not match a longer tag with the same prefix" {
    const xml = "<r><dc:titleAlt>no</dc:titleAlt><dc:title>yes</dc:title></r>";
    var props: DocProps = .{};
    parseCore(xml, &props);
    try std.testing.expectEqualStrings("yes", props.title.?);
}

test "parse: tolerates malformed and absent parts" {
    // Unterminated element must not crash or hang — it yields null.
    const broken = "<cp:coreProperties><dc:creator>oops";
    var props = parse(broken, null, false);
    try std.testing.expect(props.creator == null);
    try std.testing.expect(props.isEmpty());

    // Absent everything.
    props = parse(null, null, false);
    try std.testing.expect(props.isEmpty());
    try std.testing.expect(!props.hasIdentifyingFields());

    // custom.xml presence alone counts as identifying.
    props = parse(null, null, true);
    try std.testing.expect(!props.isEmpty());
    try std.testing.expect(props.hasIdentifyingFields());
}

test "self-closing element reads as null, and scrubs away" {
    const xml = "<r><dc:creator/><dc:title>T</dc:title></r>";
    var props: DocProps = .{};
    parseCore(xml, &props);
    try std.testing.expect(props.creator == null);

    const out = try scrubCore(std.testing.allocator, xml, .{});
    defer std.testing.allocator.free(out);
    try std.testing.expectEqualStrings("<r></r>", out);
}

test "empty element body reads as empty, not absent" {
    const xml = "<r><dc:creator></dc:creator></r>";
    var props: DocProps = .{};
    parseCore(xml, &props);
    try std.testing.expect(props.creator != null);
    try std.testing.expectEqualStrings("", props.creator.?);
}
