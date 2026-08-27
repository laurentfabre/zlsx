//! Typed-overlay parsers for the two pivot parts that carry structure
//! (S6 of `goal_sigmoid.md`):
//!
//!   - `parseCacheDefinition` — `xl/pivotCache/pivotCacheDefinitionN.xml`
//!     (`CT_PivotCacheDefinition`: the cache source, its fields, the
//!     refresh facts and the `r:id` edge to the records part).
//!   - `parseTableDefinition` — `xl/pivotTables/pivotTableN.xml`
//!     (`CT_pivotTableDefinition`: the output `location`, the field
//!     roles, the four axes, the data fields and the style).
//!
//! `pivotCacheRecordsN.xml` has no parser: the records are the cached
//! *values*, which the typed read reports by count (`recordCount`) and
//! otherwise leaves byte-preserved.
//!
//! **Conservative by construction.** Only the elements a later edit
//! lift (S7a: `location@ref`; S7b: `worksheetSource@ref`; S7c: the
//! cache-field schema) or a reader of `zlsx pivots` needs are typed.
//! Formats, conditional formats, chart formats, hierarchies, KPIs,
//! calculated members, `extLst` and every OLAP-only element stay in the
//! raw part bytes the caller already holds. The two `ref` attributes a
//! rewriter will splice are exported as **byte spans of the input**
//! (`Location.ref_span`, `WorksheetSource.ref_span`) so the future
//! editor patches exactly the attribute this parser read — reproducing
//! a parser's indexing with a second scanner is how a decoy becomes a
//! corruption (the `<tableParts>` lesson, Codex #190 r1).
//!
//! **One tree, read by its direct children.** `preflight` checks the
//! part is one balanced element tree (UTF-8, every tag closed, closes
//! matching opens, attributes well formed, nothing after the root), and
//! every parser walks the DIRECT children of the element it is reading
//! (`Children`) — a `<location>` nested inside some extension is not
//! the pivot's location, and a second `<cacheSource>` is a refusal. An
//! attribute that is present but unreadable (`cacheId="abc"`,
//! `rowGrandTotals="maybe"`) is a refusal rather than an absence or a
//! default; scalar attributes are entity-decoded first, so
//! `recordCount="&#51;"` is 3 and `"1_0"` is nothing.
//!
//! Lifetime contract: every `[]const u8` field borrows from the `xml`
//! handed to the parser and is **not decoded** — attribute values are
//! still XML-escaped and, for the ST_Xstring-typed names, still
//! `_xHHHH_`-encoded. Decoding is the caller's, by site
//! (`pkg/pivots.zig` does it with the engine's `decodeAt`). The spine
//! slices (`fields`, `row_fields`, …) come from the allocator handed
//! to the parser and are released by `deinit(allocator)` — no arena is
//! embedded, so a caller that allocates from its own arena simply
//! never calls `deinit`. (An `ArenaAllocator` carries its state by
//! value; an arena nested inside a struct that is later moved keeps an
//! allocator pointing at the old location — Codex #199 r1 REL-001.)
//!
//! Namespace prefixes: the SpreadsheetML main namespace may be bound to
//! a prefix (`<x:pivotCacheDefinition xmlns:x="…">`), and Strict OOXML
//! binds a different URI to it. The parser reads the root element's
//! prefix once and matches every child under it; it never matches on
//! the URI, so Strict and Transitional parts parse alike. A root whose
//! own binding is declared to a foreign URI is not a pivot part, and a
//! root that binds the main namespace twice (a prefix *and* the
//! default, or two prefixes) is refused rather than half-read under one
//! of them. The relationships namespace is resolved from its `xmlns`
//! declaration on the root (or the element itself) so `r:id` cannot be
//! shadowed by a foreign `vendor:id`, nor claimed by an element that
//! rebinds `r`. Comments, CDATA and processing instructions are skipped
//! through the shared decoy-aware scanner in `workbook_xml.zig`.

const std = @import("std");
const assert = std.debug.assert;
const Allocator = std.mem.Allocator;
const wbxml = @import("workbook_xml.zig");

pub const Error = error{
    OutOfMemory,
    /// The part is not the pivot part it was handed as, a required
    /// element or attribute is missing or unreadable, an element does
    /// not close, or the markup is not one well-formed tree.
    MalformedXml,
};

/// Absolute byte span inside the `xml` handed to the parser:
/// `xml[start..end]` is the attribute value (no quotes).
pub const Span = struct {
    start: usize,
    end: usize,
};

/// The two spellings of the SpreadsheetML main namespace.
pub const main_ns_uris = [_][]const u8{
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "http://purl.oclc.org/ooxml/spreadsheetml/main",
};
/// The two spellings of the relationships namespace (`r:id`).
pub const rel_ns_uris = [_][]const u8{
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://purl.oclc.org/ooxml/officeDocument/relationships",
};

// ─── Cache definition ────────────────────────────────────────────────

/// `CT_CacheSource@type` (ST_SourceType).
pub const SourceType = enum {
    worksheet,
    external,
    consolidation,
    scenario,
    /// A `type` this reader does not know. `CacheSource.type_raw`
    /// carries the spelling.
    unknown,
};

/// `CT_WorksheetSource` — also the shape of a consolidation
/// `<rangeSet>`, which carries the same four locators.
pub const WorksheetSource = struct {
    /// `ref` (ST_Ref) — the A1 rectangle, when the source is a range.
    ref: ?[]const u8 = null,
    /// Where `ref`'s value sits in the input. S7b's splice target.
    ref_span: ?Span = null,
    /// `name` — a table or defined name the range is spelled through.
    name: ?[]const u8 = null,
    /// `sheet` — the sheet `ref` is on. Absent when `name` carries a
    /// workbook-scoped spelling.
    sheet: ?[]const u8 = null,
    /// `r:id` — the relationship to an external workbook the source
    /// lives in. Absent for a local source.
    r_id: ?[]const u8 = null,
};

pub const CacheSource = struct {
    type: SourceType = .unknown,
    /// The raw `type` attribute when `type == .unknown`.
    type_raw: ?[]const u8 = null,
    /// `connectionId` — external sources.
    connection_id: ?u32 = null,
    /// `<worksheetSource>` — present for `type == .worksheet`.
    worksheet: ?WorksheetSource = null,
    /// `<consolidation><rangeSets><rangeSet …/>` — one per set, in
    /// document order; empty for every other type.
    range_sets: []WorksheetSource = &.{},
};

/// `CT_SharedItems` — the field's value inventory. Defaults are the
/// schema's, so a `<sharedItems/>` with no attributes reads as "strings,
/// semi-mixed, non-date", which is exactly what Excel means by it.
pub const SharedItems = struct {
    count: ?u32 = null,
    contains_semi_mixed_types: bool = true,
    contains_non_date: bool = true,
    contains_date: bool = false,
    contains_string: bool = true,
    contains_blank: bool = false,
    contains_mixed_types: bool = false,
    contains_number: bool = false,
    contains_integer: bool = false,
    long_text: bool = false,
    /// Raw lexical values — the number or date spelling as written.
    min_value: ?[]const u8 = null,
    max_value: ?[]const u8 = null,
    min_date: ?[]const u8 = null,
    max_date: ?[]const u8 = null,
};

pub const CacheField = struct {
    /// `name` (ST_Xstring, required). Raw.
    name: []const u8,
    caption: ?[]const u8 = null,
    num_fmt_id: ?u32 = null,
    /// `formula` — set for a calculated field. Raw (FORMULA carrier).
    formula: ?[]const u8 = null,
    /// `databaseField` — false for a calculated or group field.
    database_field: bool = true,
    shared_items: ?SharedItems = null,
};

pub const CacheDefinition = struct {
    /// `r:id` on the root — the relationship to `pivotCacheRecordsN.xml`.
    /// Raw (entity-encoded as written).
    r_id: ?[]const u8,
    record_count: ?u32,
    refreshed_by: ?[]const u8,
    /// `refreshedDate` — an Excel serial as written (some producers
    /// write ISO 8601 here; the value is not interpreted).
    refreshed_date: ?[]const u8,
    refreshed_date_iso: ?[]const u8,
    refresh_on_load: bool,
    save_data: bool,
    enable_refresh: bool,
    invalid: bool,
    background_query: bool,
    created_version: ?u32,
    refreshed_version: ?u32,
    min_refreshable_version: ?u32,
    missing_items_limit: ?u32,
    source: CacheSource,
    /// `<cacheFields>` children in document order. Ordinals are what
    /// `pivotField` / `dataField@fld` index.
    fields: []CacheField,
    /// `<cacheFields count>` as written, for a consistency check
    /// against `fields.len`.
    fields_count_attr: ?u32,

    /// Release the spines. Pass the allocator `parseCacheDefinition`
    /// was given; a caller that parsed into an arena never calls this.
    pub fn deinit(self: *CacheDefinition, allocator: Allocator) void {
        allocator.free(self.fields);
        allocator.free(self.source.range_sets);
        self.* = undefined;
    }
};

/// Parse one `pivotCacheDefinitionN.xml` part.
pub fn parseCacheDefinition(allocator: Allocator, xml: []const u8) Error!CacheDefinition {
    assert(xml.len < (1 << 31));
    const root = try scanRoot(xml, "pivotCacheDefinition");
    const attrs = xml[root.hit.attrs_start..root.hit.attrs_end];

    var def: CacheDefinition = .{
        .r_id = nsAttr(attrs, root.rel_prefix, "id"),
        .record_count = try u32Attr(attrs, "recordCount"),
        .refreshed_by = wbxml.getAttr(attrs, "refreshedBy"),
        .refreshed_date = wbxml.getAttr(attrs, "refreshedDate"),
        .refreshed_date_iso = wbxml.getAttr(attrs, "refreshedDateIso"),
        .refresh_on_load = try boolAttr(attrs, "refreshOnLoad", false),
        .save_data = try boolAttr(attrs, "saveData", true),
        .enable_refresh = try boolAttr(attrs, "enableRefresh", true),
        .invalid = try boolAttr(attrs, "invalid", false),
        .background_query = try boolAttr(attrs, "backgroundQuery", false),
        .created_version = try u32Attr(attrs, "createdVersion"),
        .refreshed_version = try u32Attr(attrs, "refreshedVersion"),
        .min_refreshable_version = try u32Attr(attrs, "minRefreshableVersion"),
        .missing_items_limit = try u32Attr(attrs, "missingItemsLimit"),
        .source = .{},
        .fields = &.{},
        .fields_count_attr = null,
    };
    errdefer def.deinit(allocator);

    var have_source = false;
    var have_fields = false;
    var kids = Children.init(xml, root.hit, root.body_end, root.prefix);
    while (try kids.next()) |k| {
        if (std.mem.eql(u8, k.local, "cacheSource")) {
            // Required, and one: a definition without a source describes
            // no data, and one with two describes two.
            if (have_source) return error.MalformedXml;
            have_source = true;
            def.source = try parseCacheSource(allocator, xml, k, root);
        } else if (std.mem.eql(u8, k.local, "cacheFields")) {
            if (have_fields) return error.MalformedXml;
            have_fields = true;
            def.fields_count_attr = try u32Attr(k.attrs(xml), "count");
            def.fields = try parseCacheFields(allocator, xml, k, root.prefix);
        }
    }
    if (!have_source) return error.MalformedXml;
    return def;
}

fn parseCacheSource(allocator: Allocator, xml: []const u8, el: Child, root: Root) Error!CacheSource {
    const attrs = el.attrs(xml);
    var src: CacheSource = .{
        .connection_id = try u32Attr(attrs, "connectionId"),
    };
    if (wbxml.getAttr(attrs, "type")) |t| {
        var buf: [32]u8 = undefined;
        src.type = if (wbxml.decodeScalarAttr(&buf, t)) |d| std.meta.stringToEnum(SourceType, d) orelse .unknown else .unknown;
        if (src.type == .unknown) src.type_raw = t;
    }
    var kids = Children.init(xml, el.hit, el.end, root.prefix);
    while (try kids.next()) |k| {
        if (std.mem.eql(u8, k.local, "worksheetSource")) {
            if (src.worksheet != null) return error.MalformedXml;
            src.worksheet = parseWorksheetSource(xml, k, root.rel_prefix);
        } else if (std.mem.eql(u8, k.local, "consolidation")) {
            if (src.range_sets.len != 0) return error.MalformedXml;
            var cons = Children.init(xml, k.hit, k.end, root.prefix);
            while (try cons.next()) |c| {
                if (!std.mem.eql(u8, c.local, "rangeSets")) continue;
                var sets: std.ArrayListUnmanaged(WorksheetSource) = .empty;
                errdefer sets.deinit(allocator);
                var rs = Children.init(xml, c.hit, c.end, root.prefix);
                while (try rs.next()) |set| {
                    if (!std.mem.eql(u8, set.local, "rangeSet")) continue;
                    try sets.append(allocator, parseWorksheetSource(xml, set, root.rel_prefix));
                }
                allocator.free(src.range_sets);
                src.range_sets = try sets.toOwnedSlice(allocator);
            }
        }
    }
    return src;
}

fn parseWorksheetSource(xml: []const u8, el: Child, rel_prefix: ?[]const u8) WorksheetSource {
    const attrs = el.attrs(xml);
    var ws: WorksheetSource = .{
        .name = wbxml.getAttr(attrs, "name"),
        .sheet = wbxml.getAttr(attrs, "sheet"),
        .r_id = nsAttr(attrs, rel_prefix, "id"),
    };
    if (wbxml.getAttr(attrs, "ref")) |r| {
        ws.ref = r;
        ws.ref_span = spanOf(xml, r);
    }
    return ws;
}

fn parseCacheFields(allocator: Allocator, xml: []const u8, el: Child, p: []const u8) Error![]CacheField {
    var out: std.ArrayListUnmanaged(CacheField) = .empty;
    errdefer out.deinit(allocator);
    var kids = Children.init(xml, el.hit, el.end, p);
    while (try kids.next()) |k| {
        if (!std.mem.eql(u8, k.local, "cacheField")) continue;
        const attrs = k.attrs(xml);
        var field: CacheField = .{
            .name = wbxml.getAttr(attrs, "name") orelse return error.MalformedXml,
            .caption = wbxml.getAttr(attrs, "caption"),
            .num_fmt_id = try u32Attr(attrs, "numFmtId"),
            .formula = wbxml.getAttr(attrs, "formula"),
            .database_field = try boolAttr(attrs, "databaseField", true),
        };
        var inner = Children.init(xml, k.hit, k.end, p);
        while (try inner.next()) |c| {
            if (!std.mem.eql(u8, c.local, "sharedItems")) continue;
            if (field.shared_items != null) return error.MalformedXml;
            field.shared_items = try parseSharedItems(c.attrs(xml));
        }
        try out.append(allocator, field);
    }
    return out.toOwnedSlice(allocator);
}

fn parseSharedItems(attrs: []const u8) Error!SharedItems {
    return .{
        .count = try u32Attr(attrs, "count"),
        .contains_semi_mixed_types = try boolAttr(attrs, "containsSemiMixedTypes", true),
        .contains_non_date = try boolAttr(attrs, "containsNonDate", true),
        .contains_date = try boolAttr(attrs, "containsDate", false),
        .contains_string = try boolAttr(attrs, "containsString", true),
        .contains_blank = try boolAttr(attrs, "containsBlank", false),
        .contains_mixed_types = try boolAttr(attrs, "containsMixedTypes", false),
        .contains_number = try boolAttr(attrs, "containsNumber", false),
        .contains_integer = try boolAttr(attrs, "containsInteger", false),
        .long_text = try boolAttr(attrs, "longText", false),
        .min_value = wbxml.getAttr(attrs, "minValue"),
        .max_value = wbxml.getAttr(attrs, "maxValue"),
        .min_date = wbxml.getAttr(attrs, "minDate"),
        .max_date = wbxml.getAttr(attrs, "maxDate"),
    };
}

// ─── Table definition ────────────────────────────────────────────────

/// `CT_Location` — where the pivot renders on its host sheet.
pub const Location = struct {
    /// `ref` (ST_Ref, required) — the output rectangle. Raw.
    ref: []const u8,
    /// Where `ref`'s value sits in the input. S7a's splice target.
    ref_span: Span,
    first_header_row: ?u32,
    first_data_row: ?u32,
    first_data_col: ?u32,
    row_page_count: ?u32,
    col_page_count: ?u32,
};

/// `CT_PivotField@axis` (ST_Axis).
pub const Axis = enum {
    row,
    col,
    page,
    values,
};

pub const PivotField = struct {
    /// `name` — a caption override; the field's real name is the
    /// cache field at the same ordinal.
    name: ?[]const u8 = null,
    /// `axis` — the axis the field is placed on, or null when it is
    /// unplaced (hidden in the field list, or a data-only field).
    axis: ?Axis = null,
    /// The raw `axis` spelling when it is not one of the four.
    axis_raw: ?[]const u8 = null,
    /// `dataField` — the field feeds at least one data field.
    data_field: bool = false,
    show_all: bool = true,
    default_subtotal: bool = true,
    num_fmt_id: ?u32 = null,
    subtotal_caption: ?[]const u8 = null,
    /// Number of `<item>` children of `<items>` — the field's distinct
    /// items as the pivot last saw them.
    item_count: u32 = 0,
};

/// A pivot-field ordinal, or the values axis — spelled `x="-2"` on
/// `<field>` and `fld="-2"` on `<pageField>`. Any other negative
/// ordinal names no field and is refused by the parser.
pub const AxisField = union(enum) {
    field: u32,
    values,
};

pub const PageField = struct {
    fld: AxisField,
    item: ?u32 = null,
    hier: ?i32 = null,
    name: ?[]const u8 = null,
    cap: ?[]const u8 = null,
};

/// ST_DataConsolidateFunction — `dataField@subtotal`.
pub const ConsolidateFunction = enum {
    average,
    count,
    count_nums,
    max,
    min,
    product,
    std_dev,
    std_dev_p,
    sum,
    variance,
    variance_p,
    unknown,

    fn fromXml(s: []const u8) ConsolidateFunction {
        const map = std.StaticStringMap(ConsolidateFunction).initComptime(.{
            .{ "average", .average },
            .{ "count", .count },
            .{ "countNums", .count_nums },
            .{ "max", .max },
            .{ "min", .min },
            .{ "product", .product },
            .{ "stdDev", .std_dev },
            .{ "stdDevp", .std_dev_p },
            .{ "sum", .sum },
            .{ "var", .variance },
            .{ "varp", .variance_p },
        });
        return map.get(s) orelse .unknown;
    }

    /// The canonical spelling a reader emits — the XML token for the
    /// known values, so a consumer can round-trip it.
    pub fn xmlName(self: ConsolidateFunction) []const u8 {
        return switch (self) {
            .average => "average",
            .count => "count",
            .count_nums => "countNums",
            .max => "max",
            .min => "min",
            .product => "product",
            .std_dev => "stdDev",
            .std_dev_p => "stdDevp",
            .sum => "sum",
            .variance => "var",
            .variance_p => "varp",
            .unknown => "unknown",
        };
    }
};

pub const DataField = struct {
    /// `name` — the caption ("Sum of Qty"). Raw.
    name: ?[]const u8 = null,
    /// `fld` (required) — the pivot/cache field ordinal.
    fld: u32,
    subtotal: ConsolidateFunction = .sum,
    /// The raw `subtotal` spelling when `subtotal == .unknown`.
    subtotal_raw: ?[]const u8 = null,
    /// `showDataAs` (ST_ShowDataAs) — raw; `null` means `normal`.
    show_data_as: ?[]const u8 = null,
    base_field: ?i32 = null,
    base_item: ?u32 = null,
    num_fmt_id: ?u32 = null,
};

pub const StyleInfo = struct {
    name: ?[]const u8 = null,
    show_row_headers: bool = false,
    show_col_headers: bool = false,
    show_row_stripes: bool = false,
    show_col_stripes: bool = false,
    show_last_column: bool = false,
};

pub const TableDefinition = struct {
    /// `name` (ST_Xstring, required). Raw.
    name: []const u8,
    /// `cacheId` (required) — matches a `<pivotCache cacheId>` in
    /// `xl/workbook.xml`.
    cache_id: u32,
    /// `dataCaption` — what the values-axis header shows ("Values").
    data_caption: ?[]const u8,
    grand_total_caption: ?[]const u8,
    data_on_rows: bool,
    data_position: ?u32,
    row_grand_totals: bool,
    col_grand_totals: bool,
    compact: bool,
    outline: bool,
    compact_data: bool,
    outline_data: bool,
    created_version: ?u32,
    updated_version: ?u32,
    min_refreshable_version: ?u32,
    location: Location,
    /// `<pivotFields>` children, parallel to the cache's fields.
    fields: []PivotField,
    row_fields: []AxisField,
    col_fields: []AxisField,
    page_fields: []PageField,
    data_fields: []DataField,
    style: ?StyleInfo,

    /// Release the spines. Pass the allocator `parseTableDefinition`
    /// was given; a caller that parsed into an arena never calls this.
    pub fn deinit(self: *TableDefinition, allocator: Allocator) void {
        allocator.free(self.fields);
        allocator.free(self.row_fields);
        allocator.free(self.col_fields);
        allocator.free(self.page_fields);
        allocator.free(self.data_fields);
        self.* = undefined;
    }
};

/// Parse one `pivotTableN.xml` part.
pub fn parseTableDefinition(allocator: Allocator, xml: []const u8) Error!TableDefinition {
    assert(xml.len < (1 << 31));
    const root = try scanRoot(xml, "pivotTableDefinition");
    const attrs = xml[root.hit.attrs_start..root.hit.attrs_end];
    const p = root.prefix;

    var def: TableDefinition = .{
        .name = wbxml.getAttr(attrs, "name") orelse return error.MalformedXml,
        .cache_id = (try u32Attr(attrs, "cacheId")) orelse return error.MalformedXml,
        .data_caption = wbxml.getAttr(attrs, "dataCaption"),
        .grand_total_caption = wbxml.getAttr(attrs, "grandTotalCaption"),
        .data_on_rows = try boolAttr(attrs, "dataOnRows", false),
        .data_position = try u32Attr(attrs, "dataPosition"),
        .row_grand_totals = try boolAttr(attrs, "rowGrandTotals", true),
        .col_grand_totals = try boolAttr(attrs, "colGrandTotals", true),
        .compact = try boolAttr(attrs, "compact", true),
        .outline = try boolAttr(attrs, "outline", false),
        .compact_data = try boolAttr(attrs, "compactData", true),
        .outline_data = try boolAttr(attrs, "outlineData", false),
        .created_version = try u32Attr(attrs, "createdVersion"),
        .updated_version = try u32Attr(attrs, "updatedVersion"),
        .min_refreshable_version = try u32Attr(attrs, "minRefreshableVersion"),
        .location = undefined,
        .fields = &.{},
        .row_fields = &.{},
        .col_fields = &.{},
        .page_fields = &.{},
        .data_fields = &.{},
        .style = null,
    };
    errdefer def.deinit(allocator);

    var have_location = false;
    var seen = std.enums.EnumSet(enum { fields, rows, cols, pages, data, style }).initEmpty();
    var kids = Children.init(xml, root.hit, root.body_end, p);
    while (try kids.next()) |k| {
        if (std.mem.eql(u8, k.local, "location")) {
            if (have_location) return error.MalformedXml;
            have_location = true;
            const la = k.attrs(xml);
            const ref = wbxml.getAttr(la, "ref") orelse return error.MalformedXml;
            if (ref.len == 0) return error.MalformedXml;
            def.location = .{
                .ref = ref,
                .ref_span = spanOf(xml, ref),
                .first_header_row = try u32Attr(la, "firstHeaderRow"),
                .first_data_row = try u32Attr(la, "firstDataRow"),
                .first_data_col = try u32Attr(la, "firstDataCol"),
                .row_page_count = try u32Attr(la, "rowPageCount"),
                .col_page_count = try u32Attr(la, "colPageCount"),
            };
        } else if (std.mem.eql(u8, k.local, "pivotFields")) {
            if (seen.contains(.fields)) return error.MalformedXml;
            seen.insert(.fields);
            def.fields = try parsePivotFields(allocator, xml, k, p);
        } else if (std.mem.eql(u8, k.local, "rowFields")) {
            if (seen.contains(.rows)) return error.MalformedXml;
            seen.insert(.rows);
            def.row_fields = try parseAxisFields(allocator, xml, k, p);
        } else if (std.mem.eql(u8, k.local, "colFields")) {
            if (seen.contains(.cols)) return error.MalformedXml;
            seen.insert(.cols);
            def.col_fields = try parseAxisFields(allocator, xml, k, p);
        } else if (std.mem.eql(u8, k.local, "pageFields")) {
            if (seen.contains(.pages)) return error.MalformedXml;
            seen.insert(.pages);
            def.page_fields = try parsePageFields(allocator, xml, k, p);
        } else if (std.mem.eql(u8, k.local, "dataFields")) {
            if (seen.contains(.data)) return error.MalformedXml;
            seen.insert(.data);
            def.data_fields = try parseDataFields(allocator, xml, k, p);
        } else if (std.mem.eql(u8, k.local, "pivotTableStyleInfo")) {
            if (seen.contains(.style)) return error.MalformedXml;
            seen.insert(.style);
            const s = k.attrs(xml);
            def.style = .{
                .name = wbxml.getAttr(s, "name"),
                .show_row_headers = try boolAttr(s, "showRowHeaders", false),
                .show_col_headers = try boolAttr(s, "showColHeaders", false),
                .show_row_stripes = try boolAttr(s, "showRowStripes", false),
                .show_col_stripes = try boolAttr(s, "showColStripes", false),
                .show_last_column = try boolAttr(s, "showLastColumn", false),
            };
        }
    }
    // `<location>` is required: without it the pivot has no output
    // rectangle, and the output rectangle is the one thing every
    // consumer of this part — the S7a lift included — exists for.
    if (!have_location) return error.MalformedXml;
    return def;
}

fn parsePivotFields(allocator: Allocator, xml: []const u8, el: Child, p: []const u8) Error![]PivotField {
    var out: std.ArrayListUnmanaged(PivotField) = .empty;
    errdefer out.deinit(allocator);
    var kids = Children.init(xml, el.hit, el.end, p);
    while (try kids.next()) |k| {
        if (!std.mem.eql(u8, k.local, "pivotField")) continue;
        const attrs = k.attrs(xml);
        var f: PivotField = .{
            .name = wbxml.getAttr(attrs, "name"),
            .data_field = try boolAttr(attrs, "dataField", false),
            .show_all = try boolAttr(attrs, "showAll", true),
            .default_subtotal = try boolAttr(attrs, "defaultSubtotal", true),
            .num_fmt_id = try u32Attr(attrs, "numFmtId"),
            .subtotal_caption = wbxml.getAttr(attrs, "subtotalCaption"),
        };
        if (wbxml.getAttr(attrs, "axis")) |ax| {
            var buf: [32]u8 = undefined;
            f.axis = if (wbxml.decodeScalarAttr(&buf, ax)) |d| axisFromXml(d) else null;
            if (f.axis == null) f.axis_raw = ax;
        }
        var inner = Children.init(xml, k.hit, k.end, p);
        while (try inner.next()) |c| {
            if (!std.mem.eql(u8, c.local, "items")) continue;
            var items = Children.init(xml, c.hit, c.end, p);
            while (try items.next()) |it| {
                if (std.mem.eql(u8, it.local, "item")) f.item_count += 1;
            }
        }
        try out.append(allocator, f);
    }
    return out.toOwnedSlice(allocator);
}

fn axisFromXml(s: []const u8) ?Axis {
    if (std.mem.eql(u8, s, "axisRow")) return .row;
    if (std.mem.eql(u8, s, "axisCol")) return .col;
    if (std.mem.eql(u8, s, "axisPage")) return .page;
    if (std.mem.eql(u8, s, "axisValues")) return .values;
    return null;
}

fn parseAxisFields(allocator: Allocator, xml: []const u8, el: Child, p: []const u8) Error![]AxisField {
    var out: std.ArrayListUnmanaged(AxisField) = .empty;
    errdefer out.deinit(allocator);
    var kids = Children.init(xml, el.hit, el.end, p);
    while (try kids.next()) |k| {
        if (!std.mem.eql(u8, k.local, "field")) continue;
        const x = (try i32Attr(k.attrs(xml), "x")) orelse return error.MalformedXml;
        try out.append(allocator, try ordinalOrValues(x));
    }
    return out.toOwnedSlice(allocator);
}

/// `-2` is the values axis; any other negative ordinal names no field
/// and cannot be read.
fn ordinalOrValues(x: i32) Error!AxisField {
    if (x == -2) return .values;
    if (x >= 0) return .{ .field = @intCast(x) };
    return error.MalformedXml;
}

fn parsePageFields(allocator: Allocator, xml: []const u8, el: Child, p: []const u8) Error![]PageField {
    var out: std.ArrayListUnmanaged(PageField) = .empty;
    errdefer out.deinit(allocator);
    var kids = Children.init(xml, el.hit, el.end, p);
    while (try kids.next()) |k| {
        if (!std.mem.eql(u8, k.local, "pageField")) continue;
        const attrs = k.attrs(xml);
        const fld = (try i32Attr(attrs, "fld")) orelse return error.MalformedXml;
        try out.append(allocator, .{
            .fld = try ordinalOrValues(fld),
            .item = try u32Attr(attrs, "item"),
            .hier = try i32Attr(attrs, "hier"),
            .name = wbxml.getAttr(attrs, "name"),
            .cap = wbxml.getAttr(attrs, "cap"),
        });
    }
    return out.toOwnedSlice(allocator);
}

fn parseDataFields(allocator: Allocator, xml: []const u8, el: Child, p: []const u8) Error![]DataField {
    var out: std.ArrayListUnmanaged(DataField) = .empty;
    errdefer out.deinit(allocator);
    var kids = Children.init(xml, el.hit, el.end, p);
    while (try kids.next()) |k| {
        if (!std.mem.eql(u8, k.local, "dataField")) continue;
        const attrs = k.attrs(xml);
        var df: DataField = .{
            .name = wbxml.getAttr(attrs, "name"),
            .fld = (try u32Attr(attrs, "fld")) orelse return error.MalformedXml,
            .show_data_as = wbxml.getAttr(attrs, "showDataAs"),
            .base_field = try i32Attr(attrs, "baseField"),
            .base_item = try u32Attr(attrs, "baseItem"),
            .num_fmt_id = try u32Attr(attrs, "numFmtId"),
        };
        if (wbxml.getAttr(attrs, "subtotal")) |s| {
            var buf: [32]u8 = undefined;
            df.subtotal = if (wbxml.decodeScalarAttr(&buf, s)) |d| ConsolidateFunction.fromXml(d) else .unknown;
            if (df.subtotal == .unknown) df.subtotal_raw = s;
        }
        try out.append(allocator, df);
    }
    return out.toOwnedSlice(allocator);
}

// ─── Scanner glue ────────────────────────────────────────────────────
//
// Public below the parsers because `pkg/pivots.zig` reads
// `<pivotCaches>` out of `xl/workbook.xml` with the same rules —
// the root's prefix, the root's relationships binding, direct
// children only — and a second scanner with slightly different rules
// is the class of bug this file exists to avoid.

/// Longest namespace prefix this parser will bind. Prefixes in real
/// parts are one to three characters; the bound exists so the tag-name
/// scratch buffers are fixed-size, and a part that exceeds it is
/// refused rather than truncated into a wrong match.
pub const max_prefix_len = 32;
const max_local_len = 40;
/// Deepest element nesting the preflight follows. Pivot parts nest
/// about eight deep; a part past this is refused rather than tracked
/// on the heap.
const max_depth = 64;

pub const Root = struct {
    hit: wbxml.TagHit,
    /// The root element's namespace prefix, without the colon; empty
    /// when the main namespace is the default.
    prefix: []const u8,
    /// The prefix bound to the relationships namespace on the root,
    /// without the colon, or null when the root declares none.
    rel_prefix: ?[]const u8,
    /// Index of the `<` of the root's closing tag — the bound every
    /// child walk runs under. For a self-closing root, the end of the
    /// root tag itself.
    body_end: usize,
};

/// Check the part is one well-formed tree, locate the root element,
/// check its local name, and learn its prefixes. Anything before the
/// root (XML declaration, comments) is skipped.
pub fn scanRoot(xml: []const u8, local: []const u8) Error!Root {
    try preflight(xml);
    var i: usize = 0;
    while (i < xml.len) {
        const lt = std.mem.indexOfScalarPos(u8, xml, i, '<') orelse return error.MalformedXml;
        const skip_to = wbxml.skipNonElement(xml, lt) catch |e| return narrow(e);
        if (skip_to != lt) {
            i = skip_to;
            continue;
        }
        // The first real element is the root.
        var j = lt + 1;
        while (j < xml.len and !isNameBoundary(xml[j])) j += 1;
        const qname = xml[lt + 1 .. j];
        const colon = std.mem.indexOfScalar(u8, qname, ':');
        const prefix = if (colon) |c| qname[0..c] else "";
        const name = if (colon) |c| qname[c + 1 ..] else qname;
        if (prefix.len > max_prefix_len) return error.MalformedXml;
        if (!std.mem.eql(u8, name, local)) return error.MalformedXml;

        const hit = (wbxml.findTagOpen(xml, lt, qname) catch |e| return narrow(e)) orelse
            return error.MalformedXml;
        const attrs = xml[hit.attrs_start..hit.attrs_end];
        // One binding of the main namespace per part. Two (a prefix and
        // the default, or two prefixes) would let children hide under
        // the one this parser is not matching.
        if (countBindings(attrs, &main_ns_uris) > 1) return error.MalformedXml;
        // And the root's own binding must be it: a `<v:pivotCacheDefinition
        // xmlns:v="urn:vendor">` is a vendor's element, whatever else the
        // root declares. Undeclared is tolerated (a hand-rolled part).
        if (declaredBinding(attrs, prefix)) |uri| {
            if (!isOneOf(uri, &main_ns_uris)) return error.MalformedXml;
        }
        const rel_prefix = bindingPrefix(attrs, &rel_ns_uris);
        if (rel_prefix) |rp| if (rp.len > max_prefix_len) return error.MalformedXml;
        const body_end = try endOfQ(xml, hit, qname);
        return .{ .hit = hit, .prefix = prefix, .rel_prefix = rel_prefix, .body_end = body_end };
    }
    return error.MalformedXml;
}

/// One decoy-aware pass over the markup that refuses what would let a
/// tag search see an element that is not one (Codex #199 r1 REL-002,
/// r2 REL-009, r3 REL-017): bytes that are not UTF-8; a `<` that opens
/// neither a tag nor a construct the scanner skips; a tag whose
/// attributes are not `name="value"` pairs, or that never closes; a
/// raw `<` anywhere inside a tag; a close tag that does not match the
/// innermost open one (so `<a><b></a></b>` and a child closing after
/// its root are refused); a second element once the root — explicit
/// or self-closing — has closed; an element still open at the end.
fn preflight(xml: []const u8) Error!void {
    if (!std.unicode.utf8ValidateSlice(xml)) return error.MalformedXml;
    var stack: [max_depth][]const u8 = undefined;
    var depth: usize = 0;
    var root_closed = false;
    var i: usize = 0;
    while (i < xml.len) {
        const lt = std.mem.indexOfScalarPos(u8, xml, i, '<') orelse break;
        const skip_to = wbxml.skipNonElement(xml, lt) catch |e| return narrow(e);
        if (skip_to != lt) {
            i = skip_to;
            continue;
        }
        if (lt + 1 >= xml.len) return error.MalformedXml;
        if (xml[lt + 1] == '/') {
            var j = lt + 2;
            while (j < xml.len and !isNameBoundary(xml[j])) j += 1;
            const qname = xml[lt + 2 .. j];
            while (j < xml.len and std.ascii.isWhitespace(xml[j])) j += 1;
            if (qname.len == 0 or j >= xml.len or xml[j] != '>') return error.MalformedXml;
            if (depth == 0 or !std.mem.eql(u8, stack[depth - 1], qname)) return error.MalformedXml;
            depth -= 1;
            if (depth == 0) root_closed = true;
            i = j + 1;
            continue;
        }
        if (!isNameStart(xml[lt + 1])) return error.MalformedXml;
        if (root_closed) return error.MalformedXml;
        var j = lt + 1;
        while (j < xml.len and !isNameBoundary(xml[j])) j += 1;
        const qname = xml[lt + 1 .. j];
        const tag_end = try attributesEnd(xml, j);
        if (!tag_end.self_closing) {
            if (depth >= max_depth) return error.MalformedXml;
            stack[depth] = qname;
            depth += 1;
        } else if (depth == 0) {
            root_closed = true;
        }
        i = tag_end.after_gt;
    }
    if (depth != 0) return error.MalformedXml;
}

const TagEnd = struct { after_gt: usize, self_closing: bool };

/// Walk a start tag's attribute region from just past its name to its
/// `>`, requiring `name = "value"` pairs (either quote, `=` with
/// optional whitespace on both sides) and nothing else. A raw `<`
/// anywhere is refused. Returns where the tag ends.
fn attributesEnd(xml: []const u8, from: usize) Error!TagEnd {
    var j = from;
    while (true) {
        while (j < xml.len and std.ascii.isWhitespace(xml[j])) j += 1;
        if (j >= xml.len) return error.MalformedXml;
        if (xml[j] == '>') return .{ .after_gt = j + 1, .self_closing = false };
        if (xml[j] == '/') {
            if (j + 1 < xml.len and xml[j + 1] == '>') return .{ .after_gt = j + 2, .self_closing = true };
            return error.MalformedXml;
        }
        if (!isNameStart(xml[j])) return error.MalformedXml;
        while (j < xml.len and xml[j] != '=' and !std.ascii.isWhitespace(xml[j]) and xml[j] != '<' and xml[j] != '>' and xml[j] != '/') j += 1;
        while (j < xml.len and std.ascii.isWhitespace(xml[j])) j += 1;
        if (j >= xml.len or xml[j] != '=') return error.MalformedXml;
        j += 1;
        while (j < xml.len and std.ascii.isWhitespace(xml[j])) j += 1;
        if (j >= xml.len or (xml[j] != '"' and xml[j] != '\'')) return error.MalformedXml;
        const quote = xml[j];
        j += 1;
        while (j < xml.len and xml[j] != quote) : (j += 1) {
            if (xml[j] == '<') return error.MalformedXml;
        }
        if (j >= xml.len) return error.MalformedXml;
        j += 1;
    }
}

fn isNameStart(c: u8) bool {
    return std.ascii.isAlphabetic(c) or c == '_' or c == ':';
}

fn isNameBoundary(c: u8) bool {
    return c == ' ' or c == '\t' or c == '\n' or c == '\r' or c == '/' or c == '>';
}

/// One direct child of an element, as `Children` yields it.
pub const Child = struct {
    /// The child's local name under the parent's prefix.
    local: []const u8,
    hit: wbxml.TagHit,
    /// One past the child: the byte after its `/>`, or the `<` of its
    /// close tag.
    end: usize,

    pub fn attrs(self: Child, xml: []const u8) []const u8 {
        return xml[self.hit.attrs_start..self.hit.attrs_end];
    }
};

/// The direct children of an element, in order — never a descendant,
/// and only the ones in the parent's namespace prefix: a child under
/// another prefix (an `x14:` extension body) is stepped over whole.
/// Runs on preflighted markup, where every open has its close.
pub const Children = struct {
    xml: []const u8,
    prefix: []const u8,
    cursor: usize,
    end: usize,

    pub fn init(xml: []const u8, parent: wbxml.TagHit, parent_end: usize, prefix: []const u8) Children {
        return .{
            .xml = xml,
            .prefix = prefix,
            .cursor = if (parent.self_closing) parent_end else parent.after_tag_close,
            .end = parent_end,
        };
    }

    pub fn next(self: *Children) Error!?Child {
        const xml = self.xml;
        while (self.cursor < self.end) {
            const lt = std.mem.indexOfScalarPos(u8, xml, self.cursor, '<') orelse return null;
            if (lt >= self.end) return null;
            const skip_to = wbxml.skipNonElement(xml, lt) catch |e| return narrow(e);
            if (skip_to != lt) {
                self.cursor = skip_to;
                continue;
            }
            // The parent's own close tag ends the walk.
            if (lt + 1 >= xml.len or xml[lt + 1] == '/') return null;
            var j = lt + 1;
            while (j < xml.len and !isNameBoundary(xml[j])) j += 1;
            const qname = xml[lt + 1 .. j];
            const hit = (wbxml.findTagOpen(xml, lt, qname) catch |e| return narrow(e)) orelse
                return error.MalformedXml;
            const child_end = try endOfQ(xml, hit, qname);
            self.cursor = if (hit.self_closing) child_end else child_end + "</".len + qname.len + ">".len;
            if (localUnder(qname, self.prefix)) |local| {
                return .{ .local = local, .hit = hit, .end = child_end };
            }
        }
        return null;
    }
};

/// `qname`'s local name when its prefix is `prefix` (empty = none).
fn localUnder(qname: []const u8, prefix: []const u8) ?[]const u8 {
    if (std.mem.indexOfScalar(u8, qname, ':')) |c| {
        if (prefix.len == 0) return null;
        if (!std.mem.eql(u8, qname[0..c], prefix)) return null;
        return qname[c + 1 ..];
    }
    return if (prefix.len == 0) qname else null;
}

/// One past the element opened at `hit`: the byte after its `/>`, or
/// the `<` of its matching close tag — which must exist.
pub fn endOf(xml: []const u8, hit: wbxml.TagHit, prefix: []const u8, local: []const u8) Error!usize {
    var buf: [max_prefix_len + 1 + max_local_len]u8 = undefined;
    return endOfQ(xml, hit, qualify(&buf, prefix, local));
}

fn endOfQ(xml: []const u8, hit: wbxml.TagHit, qname: []const u8) Error!usize {
    if (hit.self_closing) return hit.after_tag_close;
    return closeOfQ(xml, hit, qname);
}

/// The `<` of the close tag matching the open tag at `hit`, by depth:
/// a same-name element opened inside is closed before the outer one
/// is. None of these schemas nest an element in itself, but a scanner
/// that took the first `</…>` would pair the wrong tags on a part that
/// does, and read the outer element's children as the inner one's.
fn closeOfQ(xml: []const u8, hit: wbxml.TagHit, qname: []const u8) Error!usize {
    assert(!hit.self_closing);
    if (qname.len > max_prefix_len + 1 + max_local_len) return error.MalformedXml;
    var close_buf: [2 + max_prefix_len + 1 + max_local_len + 1]u8 = undefined;
    close_buf[0] = '<';
    close_buf[1] = '/';
    @memcpy(close_buf[2 .. 2 + qname.len], qname);
    close_buf[2 + qname.len] = '>';
    const close_tag = close_buf[0 .. 2 + qname.len + 1];

    var depth: usize = 1;
    var cursor = hit.after_tag_close;
    while (true) {
        const close = (wbxml.findClosingTag(xml, cursor, close_tag) catch |e| return narrow(e)) orelse
            return error.MalformedXml;
        var scan = cursor;
        while (wbxml.findTagOpen(xml[0..close], scan, qname) catch |e| return narrow(e)) |inner| {
            if (!inner.self_closing) depth += 1;
            scan = inner.after_tag_close;
        }
        depth -= 1;
        if (depth == 0) return close;
        cursor = close + close_tag.len;
    }
}

/// The shared scanner declares `workbook_xml.zig`'s whole error set;
/// only two of its members can arise from a tag search.
fn narrow(err: wbxml.Error) Error {
    return switch (err) {
        error.OutOfMemory => error.OutOfMemory,
        else => error.MalformedXml,
    };
}

fn qualify(buf: []u8, prefix: []const u8, local: []const u8) []const u8 {
    if (prefix.len == 0) {
        @memcpy(buf[0..local.len], local);
        return buf[0..local.len];
    }
    @memcpy(buf[0..prefix.len], prefix);
    buf[prefix.len] = ':';
    @memcpy(buf[prefix.len + 1 .. prefix.len + 1 + local.len], local);
    return buf[0 .. prefix.len + 1 + local.len];
}

/// Walks `name="value"` pairs of a preflighted attributes region.
const AttrIter = struct {
    attrs: []const u8,
    i: usize = 0,

    const Pair = struct { name: []const u8, value: []const u8 };

    fn next(self: *AttrIter) ?Pair {
        const attrs = self.attrs;
        var i = self.i;
        while (i < attrs.len and std.ascii.isWhitespace(attrs[i])) i += 1;
        if (i >= attrs.len) return null;
        const name_start = i;
        while (i < attrs.len and attrs[i] != '=' and !std.ascii.isWhitespace(attrs[i])) i += 1;
        const name = attrs[name_start..i];
        while (i < attrs.len and (attrs[i] == '=' or std.ascii.isWhitespace(attrs[i]))) i += 1;
        if (i >= attrs.len) return null;
        if (attrs[i] != '"' and attrs[i] != '\'') return null;
        const quote = attrs[i];
        i += 1;
        const val_start = i;
        while (i < attrs.len and attrs[i] != quote) i += 1;
        const value = attrs[val_start..i];
        if (i < attrs.len) i += 1;
        self.i = i;
        return .{ .name = name, .value = value };
    }
};

/// How many distinct `xmlns` bindings on this element name one of
/// `uris` — the default namespace counts as one.
fn countBindings(attrs: []const u8, uris: []const []const u8) usize {
    var n: usize = 0;
    var it: AttrIter = .{ .attrs = attrs };
    while (it.next()) |a| {
        if (!isXmlnsDecl(a.name)) continue;
        if (isOneOf(a.value, uris)) n += 1;
    }
    return n;
}

/// The prefix `xmlns:PREFIX` binds to one of `uris` on this element,
/// or null. A default binding (`xmlns="…"`) does not apply to
/// attributes and is not returned.
fn bindingPrefix(attrs: []const u8, uris: []const []const u8) ?[]const u8 {
    var it: AttrIter = .{ .attrs = attrs };
    while (it.next()) |a| {
        if (!std.mem.startsWith(u8, a.name, "xmlns:")) continue;
        if (isOneOf(a.value, uris)) return a.name["xmlns:".len..];
    }
    return null;
}

fn isXmlnsDecl(name: []const u8) bool {
    return std.mem.eql(u8, name, "xmlns") or std.mem.startsWith(u8, name, "xmlns:");
}

/// The URI this element binds `prefix` to (`xmlns:PREFIX`, or `xmlns`
/// for the empty prefix), or null when it declares no such binding.
fn declaredBinding(attrs: []const u8, prefix: []const u8) ?[]const u8 {
    var it: AttrIter = .{ .attrs = attrs };
    while (it.next()) |a| {
        const declared: []const u8 = if (std.mem.eql(u8, a.name, "xmlns"))
            ""
        else if (std.mem.startsWith(u8, a.name, "xmlns:"))
            a.name["xmlns:".len..]
        else
            continue;
        if (std.mem.eql(u8, declared, prefix)) return a.value;
    }
    return null;
}

fn isOneOf(uri: []const u8, uris: []const []const u8) bool {
    for (uris) |u| if (std.mem.eql(u8, uri, u)) return true;
    return false;
}

/// A relationships-namespace attribute (`r:id`) on an element, raw. A
/// prefix the element itself declares counts only if it declares it
/// to the relationships namespace — `xmlns:r="urn:vendor" r:id="…"`
/// is a vendor's attribute even under the root's `r`. Otherwise the
/// prefix must be the one the element or the root binds to that
/// namespace (`Root.rel_prefix`); with none declared anywhere the
/// attribute is matched under any prefix, which tolerates a producer
/// that left the declaration out. `xmlns:*` declarations never match.
pub fn nsAttr(attrs: []const u8, root_rel_prefix: ?[]const u8, local: []const u8) ?[]const u8 {
    const bound = bindingPrefix(attrs, &rel_ns_uris) orelse root_rel_prefix;
    var it: AttrIter = .{ .attrs = attrs };
    while (it.next()) |a| {
        const c = std.mem.indexOfScalar(u8, a.name, ':') orelse continue;
        const prefix = a.name[0..c];
        if (prefix.len == 0 or std.mem.eql(u8, prefix, "xmlns")) continue;
        if (!std.mem.eql(u8, a.name[c + 1 ..], local)) continue;
        if (declaredBinding(attrs, prefix)) |uri| {
            if (isOneOf(uri, &rel_ns_uris)) return a.value;
            continue;
        }
        if (bound) |b| {
            if (std.mem.eql(u8, prefix, b)) return a.value;
        } else {
            return a.value;
        }
    }
    return null;
}

/// Longest scalar attribute value this parser decodes. Every scalar
/// here is a small integer or a boolean token.
const max_scalar_len = 32;

/// An optional unsigned attribute: null when absent, `MalformedXml`
/// when present but not an xsd decimal — entity-decoded first, so
/// `&#51;` is 3, and digits only, so `1_0` and `+1` are nothing. A
/// number that cannot be read is not a number that is missing.
pub fn u32Attr(attrs: []const u8, name: []const u8) Error!?u32 {
    const raw = wbxml.getAttr(attrs, name) orelse return null;
    var buf: [max_scalar_len]u8 = undefined;
    const v = wbxml.decodeScalarAttr(&buf, raw) orelse return error.MalformedXml;
    if (v.len == 0) return error.MalformedXml;
    for (v) |d| if (!std.ascii.isDigit(d)) return error.MalformedXml;
    return std.fmt.parseInt(u32, v, 10) catch error.MalformedXml;
}

fn i32Attr(attrs: []const u8, name: []const u8) Error!?i32 {
    const raw = wbxml.getAttr(attrs, name) orelse return null;
    var buf: [max_scalar_len]u8 = undefined;
    const v = wbxml.decodeScalarAttr(&buf, raw) orelse return error.MalformedXml;
    const digits = if (v.len > 0 and v[0] == '-') v[1..] else v;
    if (digits.len == 0) return error.MalformedXml;
    for (digits) |d| if (!std.ascii.isDigit(d)) return error.MalformedXml;
    return std.fmt.parseInt(i32, v, 10) catch error.MalformedXml;
}

/// xsd:boolean: `1`/`true` and `0`/`false` (entity-decoded); absent →
/// the schema default; anything else is refused rather than defaulted.
fn boolAttr(attrs: []const u8, name: []const u8, default: bool) Error!bool {
    const raw = wbxml.getAttr(attrs, name) orelse return default;
    var buf: [max_scalar_len]u8 = undefined;
    const v = wbxml.decodeScalarAttr(&buf, raw) orelse return error.MalformedXml;
    if (std.mem.eql(u8, v, "1") or std.mem.eql(u8, v, "true")) return true;
    if (std.mem.eql(u8, v, "0") or std.mem.eql(u8, v, "false")) return false;
    return error.MalformedXml;
}

/// The absolute span of a slice that borrows from `xml`.
fn spanOf(xml: []const u8, sub: []const u8) Span {
    const base = @intFromPtr(xml.ptr);
    const start = @intFromPtr(sub.ptr) - base;
    assert(start + sub.len <= xml.len);
    return .{ .start = start, .end = start + sub.len };
}

// ─── Tests ───────────────────────────────────────────────────────────

const testing = std.testing;

const cache_def_worksheet =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" refreshedBy="Alex" refreshedDate="42722.61" createdVersion="5" refreshedVersion="5" minRefreshableVersion="3" recordCount="29" refreshOnLoad="1">
    \\<cacheSource type="worksheet"><worksheetSource sheet="Data" ref="A1:C4"/></cacheSource>
    \\<cacheFields count="3">
    \\<cacheField name="Region" numFmtId="0"><sharedItems count="2"><s v="East"/><s v="West"/></sharedItems></cacheField>
    \\<cacheField name="Qty" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="1" maxValue="9"/></cacheField>
    \\<cacheField name="Margin" numFmtId="0" databaseField="0" formula="Qty*2"/>
    \\</cacheFields>
    \\<extLst><ext uri="{725AE2AE-9491-48be-B2B4-4EB974FC3084}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:pivotCacheDefinition pivotCacheId="2"/></ext></extLst>
    \\</pivotCacheDefinition>
;

test "parseCacheDefinition: worksheet source, fields, shared-item flags, spans" {
    var def = try parseCacheDefinition(testing.allocator, cache_def_worksheet);
    defer def.deinit(testing.allocator);
    try testing.expectEqualStrings("rId1", def.r_id.?);
    try testing.expectEqual(@as(?u32, 29), def.record_count);
    try testing.expectEqualStrings("Alex", def.refreshed_by.?);
    try testing.expect(def.refresh_on_load);
    try testing.expect(def.save_data);
    try testing.expectEqual(SourceType.worksheet, def.source.type);
    const ws = def.source.worksheet.?;
    try testing.expectEqualStrings("Data", ws.sheet.?);
    try testing.expectEqualStrings("A1:C4", ws.ref.?);
    try testing.expect(ws.name == null);
    // The span points at exactly the attribute value.
    const span = ws.ref_span.?;
    try testing.expectEqualStrings("A1:C4", cache_def_worksheet[span.start..span.end]);
    try testing.expectEqual(@as(u8, '"'), cache_def_worksheet[span.start - 1]);

    try testing.expectEqual(@as(usize, 3), def.fields.len);
    try testing.expectEqual(@as(?u32, 3), def.fields_count_attr);
    try testing.expectEqualStrings("Region", def.fields[0].name);
    const region = def.fields[0].shared_items.?;
    try testing.expectEqual(@as(?u32, 2), region.count);
    try testing.expect(region.contains_string);
    try testing.expect(!region.contains_number);
    const qty = def.fields[1].shared_items.?;
    try testing.expect(!qty.contains_string);
    try testing.expect(qty.contains_number);
    try testing.expect(qty.contains_integer);
    try testing.expectEqualStrings("1", qty.min_value.?);
    try testing.expectEqualStrings("9", qty.max_value.?);
    try testing.expectEqualStrings("Qty*2", def.fields[2].formula.?);
    try testing.expect(!def.fields[2].database_field);
    try testing.expect(def.fields[2].shared_items == null);
}

test "parseCacheDefinition: table-name source carries no sheet, no ref" {
    const xml =
        \\<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" recordCount="50"><cacheSource type="worksheet"><worksheetSource name="Table2"/></cacheSource><cacheFields count="0"/></pivotCacheDefinition>
    ;
    var def = try parseCacheDefinition(testing.allocator, xml);
    defer def.deinit(testing.allocator);
    const ws = def.source.worksheet.?;
    try testing.expectEqualStrings("Table2", ws.name.?);
    try testing.expect(ws.sheet == null);
    try testing.expect(ws.ref == null);
    try testing.expect(ws.ref_span == null);
    try testing.expectEqual(@as(usize, 0), def.fields.len);
}

test "parseCacheDefinition: external and consolidation sources" {
    const ext =
        \\<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="external" connectionId="3"/><cacheFields/></pivotCacheDefinition>
    ;
    var e = try parseCacheDefinition(testing.allocator, ext);
    defer e.deinit(testing.allocator);
    try testing.expectEqual(SourceType.external, e.source.type);
    try testing.expectEqual(@as(?u32, 3), e.source.connection_id);
    try testing.expect(e.source.worksheet == null);

    const cons =
        \\<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="consolidation"><consolidation autoPage="0"><rangeSets count="2"><rangeSet i1="0" sheet="Q1" ref="A1:B9"/><rangeSet i1="1" name="Q2Data"/></rangeSets></consolidation></cacheSource></pivotCacheDefinition>
    ;
    var c = try parseCacheDefinition(testing.allocator, cons);
    defer c.deinit(testing.allocator);
    try testing.expectEqual(SourceType.consolidation, c.source.type);
    try testing.expectEqual(@as(usize, 2), c.source.range_sets.len);
    try testing.expectEqualStrings("Q1", c.source.range_sets[0].sheet.?);
    try testing.expectEqualStrings("A1:B9", c.source.range_sets[0].ref.?);
    try testing.expectEqualStrings("Q2Data", c.source.range_sets[1].name.?);
}

test "parseCacheDefinition: external-workbook worksheet source keeps its r:id" {
    const xml =
        \\<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><cacheSource type="worksheet"><worksheetSource r:id="rId2" sheet="Sheet1" ref="A1:B2"/></cacheSource></pivotCacheDefinition>
    ;
    var def = try parseCacheDefinition(testing.allocator, xml);
    defer def.deinit(testing.allocator);
    try testing.expectEqualStrings("rId2", def.source.worksheet.?.r_id.?);
}

test "parseCacheDefinition: prefixed main namespace (Strict-style binding)" {
    const xml =
        \\<?xml version="1.0"?>
        \\<x:pivotCacheDefinition xmlns:x="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:rel="http://purl.oclc.org/ooxml/officeDocument/relationships" rel:id="rId9" recordCount="2"><x:cacheSource type="worksheet"><x:worksheetSource sheet="S" ref="A1:A3"/></x:cacheSource><x:cacheFields count="1"><x:cacheField name="A" numFmtId="0"><x:sharedItems/></x:cacheField></x:cacheFields></x:pivotCacheDefinition>
    ;
    var def = try parseCacheDefinition(testing.allocator, xml);
    defer def.deinit(testing.allocator);
    try testing.expectEqualStrings("rId9", def.r_id.?);
    try testing.expectEqualStrings("S", def.source.worksheet.?.sheet.?);
    try testing.expectEqual(@as(usize, 1), def.fields.len);
    try testing.expectEqualStrings("A", def.fields[0].name);
}

test "parseCacheDefinition: r:id is bound to the relationships namespace, not any prefix" {
    // A foreign `vendor:id` before the real one does not shadow it …
    const shadowed =
        \\<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:vendor="urn:vendor" vendor:id="nope" r:id="rId1"><cacheSource type="worksheet"><worksheetSource vendor:id="nope" r:id="rId2" sheet="S" ref="A1"/></cacheSource></pivotCacheDefinition>
    ;
    var s = try parseCacheDefinition(testing.allocator, shadowed);
    defer s.deinit(testing.allocator);
    try testing.expectEqualStrings("rId1", s.r_id.?);
    try testing.expectEqualStrings("rId2", s.source.worksheet.?.r_id.?);
    // … a binding on the element itself wins over the root's …
    const local =
        \\<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><cacheSource type="worksheet"><worksheetSource xmlns:q="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="wrong" q:id="rId7" sheet="S" ref="A1"/></cacheSource></pivotCacheDefinition>
    ;
    var l = try parseCacheDefinition(testing.allocator, local);
    defer l.deinit(testing.allocator);
    try testing.expectEqualStrings("rId7", l.source.worksheet.?.r_id.?);
    // … and with no declaration anywhere, any prefix is tolerated.
    const undeclared =
        \\<pivotCacheDefinition r:id="rId3"><cacheSource type="worksheet"/></pivotCacheDefinition>
    ;
    var u = try parseCacheDefinition(testing.allocator, undeclared);
    defer u.deinit(testing.allocator);
    try testing.expectEqualStrings("rId3", u.r_id.?);
}

test "parseCacheDefinition: decoys in comments and CDATA are not elements" {
    const xml =
        \\<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><!-- <cacheSource type="external"/> --><cacheSource type="worksheet"><!-- <worksheetSource sheet="Decoy" ref="Z1"/> --><worksheetSource sheet="Real" ref="A1:B2"/></cacheSource><cacheFields count="1"><![CDATA[<cacheField name="Decoy"/>]]><cacheField name="Real"/></cacheFields></pivotCacheDefinition>
    ;
    var def = try parseCacheDefinition(testing.allocator, xml);
    defer def.deinit(testing.allocator);
    try testing.expectEqual(SourceType.worksheet, def.source.type);
    try testing.expectEqualStrings("Real", def.source.worksheet.?.sheet.?);
    try testing.expectEqual(@as(usize, 1), def.fields.len);
    try testing.expectEqualStrings("Real", def.fields[0].name);
}

test "parseCacheDefinition: only direct children count — a nested or extension-held source is not the source" {
    // A `<worksheetSource>` inside a same-namespace extension body is a
    // descendant, not the cache's source; the direct child wins, and a
    // second direct child is a refusal.
    const nested =
        \\<pivotCacheDefinition><cacheSource type="worksheet"><extLst><ext><worksheetSource sheet="Decoy" ref="Z1"/></ext></extLst><worksheetSource sheet="Real" ref="A1"/></cacheSource><extLst><ext><cacheFields count="1"><cacheField name="Decoy"/></cacheFields></ext></extLst></pivotCacheDefinition>
    ;
    var def = try parseCacheDefinition(testing.allocator, nested);
    defer def.deinit(testing.allocator);
    try testing.expectEqualStrings("Real", def.source.worksheet.?.sheet.?);
    try testing.expectEqual(@as(usize, 0), def.fields.len);
    try testing.expectError(error.MalformedXml, parseCacheDefinition(testing.allocator,
        \\<pivotCacheDefinition><cacheSource type="worksheet"/><cacheSource type="external"/></pivotCacheDefinition>
    ));
    try testing.expectError(error.MalformedXml, parseCacheDefinition(testing.allocator,
        \\<pivotCacheDefinition><cacheSource type="worksheet"><worksheetSource sheet="A" ref="A1"/><worksheetSource sheet="B" ref="A1"/></cacheSource></pivotCacheDefinition>
    ));
}

test "parseCacheDefinition: refuses the wrong part, a missing cacheSource, an unclosed root" {
    try testing.expectError(error.MalformedXml, parseCacheDefinition(testing.allocator,
        \\<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>
    ));
    try testing.expectError(error.MalformedXml, parseCacheDefinition(testing.allocator,
        \\<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheFields/></pivotCacheDefinition>
    ));
    try testing.expectError(error.MalformedXml, parseCacheDefinition(testing.allocator,
        \\<pivotCacheDefinition><cacheSource type="worksheet"><worksheetSource sheet="S" ref="A1"/></cacheSource>
    ));
    try testing.expectError(error.MalformedXml, parseCacheDefinition(testing.allocator,
        \\<pivotCacheDefinition><cacheSource type="worksheet"/><cacheFields><cacheField numFmtId="0"/></cacheFields></pivotCacheDefinition>
    ));
    try testing.expectError(error.MalformedXml, parseCacheDefinition(testing.allocator, ""));
}

test "parseCacheDefinition: an element that does not close is refused" {
    try testing.expectError(error.MalformedXml, parseCacheDefinition(testing.allocator,
        \\<pivotCacheDefinition><cacheSource type="worksheet"><worksheetSource sheet="S" ref="A1"></cacheSource></pivotCacheDefinition>
    ));
    try testing.expectError(error.MalformedXml, parseCacheDefinition(testing.allocator,
        \\<pivotCacheDefinition><cacheSource type="worksheet"/><cacheFields><cacheField name="A"><sharedItems count="1"></cacheField></cacheFields></pivotCacheDefinition>
    ));
    try testing.expectError(error.MalformedXml, parseCacheDefinition(testing.allocator,
        \\<pivotCacheDefinition><cacheSource type="worksheet"><worksheetSource sheet="S" ref="A1"/></pivotCacheDefinition>
    ));
}

test "parseCacheDefinition: scalars are xsd lexical values, entity-decoded" {
    try testing.expectError(error.MalformedXml, parseCacheDefinition(testing.allocator,
        \\<pivotCacheDefinition recordCount="lots"><cacheSource type="worksheet"/></pivotCacheDefinition>
    ));
    try testing.expectError(error.MalformedXml, parseCacheDefinition(testing.allocator,
        \\<pivotCacheDefinition recordCount="1_0"><cacheSource type="worksheet"/></pivotCacheDefinition>
    ));
    try testing.expectError(error.MalformedXml, parseCacheDefinition(testing.allocator,
        \\<pivotCacheDefinition recordCount="+3"><cacheSource type="worksheet"/></pivotCacheDefinition>
    ));
    try testing.expectError(error.MalformedXml, parseCacheDefinition(testing.allocator,
        \\<pivotCacheDefinition refreshOnLoad="maybe"><cacheSource type="worksheet"/></pivotCacheDefinition>
    ));
    try testing.expectError(error.MalformedXml, parseCacheDefinition(testing.allocator,
        \\<pivotCacheDefinition><cacheSource type="worksheet"/><cacheFields count="two"/></pivotCacheDefinition>
    ));
    try testing.expectError(error.MalformedXml, parseCacheDefinition(testing.allocator,
        \\<pivotCacheDefinition><cacheSource type="worksheet"/><cacheFields><cacheField name="A"><sharedItems containsNumber="yes"/></cacheField></cacheFields></pivotCacheDefinition>
    ));
    // Character references are the same value.
    var ok = try parseCacheDefinition(testing.allocator,
        \\<pivotCacheDefinition recordCount="&#51;" refreshOnLoad="&#49;"><cacheSource type="work&#115;heet"/></pivotCacheDefinition>
    );
    defer ok.deinit(testing.allocator);
    try testing.expectEqual(@as(?u32, 3), ok.record_count);
    try testing.expect(ok.refresh_on_load);
    try testing.expectEqual(SourceType.worksheet, ok.source.type);
}

test "preflight: a raw `<` inside an attribute value cannot spoof an element" {
    try testing.expectError(error.MalformedXml, parseCacheDefinition(testing.allocator,
        \\<pivotCacheDefinition><cacheSource type="worksheet" bad="<worksheetSource sheet='Decoy' ref='Z1'/>"><worksheetSource sheet="Real" ref="A1"/></cacheSource></pivotCacheDefinition>
    ));
    // A `>` inside a quoted value is data, not a tag end.
    var ok = try parseCacheDefinition(testing.allocator,
        \\<pivotCacheDefinition refreshedBy="a > b"><cacheSource type="worksheet"><worksheetSource sheet="S" ref="A1"/></cacheSource></pivotCacheDefinition>
    );
    defer ok.deinit(testing.allocator);
    try testing.expectEqualStrings("a > b", ok.refreshed_by.?);
    try testing.expectEqualStrings("S", ok.source.worksheet.?.sheet.?);
}

test "preflight: crossed and out-of-root closures, a raw `<` in a tag, bad UTF-8, bad attributes are refused" {
    // A child that closes after its root.
    try testing.expectError(error.MalformedXml, parseTableDefinition(testing.allocator,
        \\<pivotTableDefinition name="P" cacheId="0"><location ref="A1"></pivotTableDefinition></location>
    ));
    // Crossed: `<cacheSource><worksheetSource></cacheSource></worksheetSource>`.
    try testing.expectError(error.MalformedXml, parseCacheDefinition(testing.allocator,
        \\<pivotCacheDefinition><cacheSource type="worksheet"><worksheetSource sheet="S" ref="A1"></cacheSource></worksheetSource></pivotCacheDefinition>
    ));
    // Markup after the root has closed — explicitly, or self-closing.
    try testing.expectError(error.MalformedXml, parseCacheDefinition(testing.allocator,
        \\<pivotCacheDefinition><cacheSource type="worksheet"/></pivotCacheDefinition><cacheSource type="external"/>
    ));
    try testing.expectError(error.MalformedXml, parseCacheDefinition(testing.allocator,
        \\<pivotCacheDefinition/><pivotCacheDefinition><cacheSource type="worksheet"/></pivotCacheDefinition>
    ));
    // A raw `<` inside a tag, outside quotes.
    try testing.expectError(error.MalformedXml, parseCacheDefinition(testing.allocator,
        \\<pivotCacheDefinition bad=<x><cacheSource type="worksheet"/></pivotCacheDefinition>
    ));
    // A bare attribute (no `=`), which would otherwise hide the bindings after it.
    try testing.expectError(error.MalformedXml, parseCacheDefinition(testing.allocator,
        \\<pivotCacheDefinition bare xmlns="urn:vendor"><cacheSource type="worksheet"/></pivotCacheDefinition>
    ));
    // Not UTF-8.
    try testing.expectError(error.MalformedXml, parseCacheDefinition(testing.allocator, "<pivotCacheDefinition refreshedBy=\"z\xff\"><cacheSource type=\"worksheet\"/></pivotCacheDefinition>"));
    // A comment after the root, `>` in text, an attribute spanning
    // lines, whitespace around `=`, and a PI are all fine.
    var ok = try parseCacheDefinition(testing.allocator,
        \\<?xml version="1.0"?><pivotCacheDefinition refreshedBy = "two
        \\lines"><cacheSource type="worksheet"><worksheetSource sheet="S" ref="A1"/></cacheSource>a > b</pivotCacheDefinition><!-- trailing -->
    );
    defer ok.deinit(testing.allocator);
    try testing.expectEqualStrings("S", ok.source.worksheet.?.sheet.?);
    try testing.expectEqualStrings("two\nlines", ok.refreshed_by.?);
}

test "scanRoot: a root bound to a foreign namespace is not a pivot part" {
    try testing.expectError(error.MalformedXml, parseCacheDefinition(testing.allocator,
        \\<v:pivotCacheDefinition xmlns:v="urn:vendor" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><v:cacheSource type="worksheet"/></v:pivotCacheDefinition>
    ));
    try testing.expectError(error.MalformedXml, parseTableDefinition(testing.allocator,
        \\<pivotTableDefinition xmlns="urn:vendor" name="P" cacheId="0"><location ref="A1"/></pivotTableDefinition>
    ));
}

test "scanRoot: two bindings of the main namespace on the root are refused" {
    try testing.expectError(error.MalformedXml, parseCacheDefinition(testing.allocator,
        \\<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="worksheet"/></pivotCacheDefinition>
    ));
}

const table_def_xml =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable3" cacheId="1" dataCaption="Values" updatedVersion="5" minRefreshableVersion="3" createdVersion="5" outline="1" outlineData="1" multipleFieldFilters="0" chartFormat="1"><location ref="A1:D5" firstHeaderRow="0" firstDataRow="1" firstDataCol="1"/><pivotFields count="4"><pivotField dataField="1" showAll="0"/><pivotField axis="axisRow" showAll="0"><items count="4"><item x="1"/><item x="0"/><item x="2"/><item t="default"/></items></pivotField><pivotField axis="axisPage" showAll="0"><items count="1"><item t="default"/></items></pivotField><pivotField showAll="0"/></pivotFields><rowFields count="1"><field x="1"/></rowFields><rowItems count="2"><i><x/></i><i t="grand"><x/></i></rowItems><colFields count="1"><field x="-2"/></colFields><pageFields count="1"><pageField fld="2" hier="-1"/></pageFields><dataFields count="2"><dataField name="Average of mpg" fld="0" subtotal="average" baseField="1" baseItem="0"/><dataField name="Sum of wt" fld="3" baseField="1" baseItem="0"/></dataFields><chartFormats count="1"><chartFormat chart="0" format="0" series="1"><pivotArea type="data" outline="0" fieldPosition="0"><references count="1"><reference field="4294967294" count="1" selected="0"><x v="0"/></reference></references></pivotArea></chartFormat></chartFormats><pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/><extLst><ext uri="{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:pivotTableDefinition hideValuesRow="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"/></ext></extLst></pivotTableDefinition>
;

test "parseTableDefinition: name, cache, location, roles, axes, data fields, style" {
    var def = try parseTableDefinition(testing.allocator, table_def_xml);
    defer def.deinit(testing.allocator);
    try testing.expectEqualStrings("PivotTable3", def.name);
    try testing.expectEqual(@as(u32, 1), def.cache_id);
    try testing.expectEqualStrings("Values", def.data_caption.?);
    try testing.expect(def.outline);
    try testing.expect(def.outline_data);
    try testing.expect(def.row_grand_totals);

    try testing.expectEqualStrings("A1:D5", def.location.ref);
    try testing.expectEqualStrings("A1:D5", table_def_xml[def.location.ref_span.start..def.location.ref_span.end]);
    try testing.expectEqual(@as(?u32, 0), def.location.first_header_row);
    try testing.expectEqual(@as(?u32, 1), def.location.first_data_row);
    try testing.expectEqual(@as(?u32, 1), def.location.first_data_col);

    try testing.expectEqual(@as(usize, 4), def.fields.len);
    try testing.expect(def.fields[0].data_field);
    try testing.expect(def.fields[0].axis == null);
    try testing.expectEqual(Axis.row, def.fields[1].axis.?);
    try testing.expectEqual(@as(u32, 4), def.fields[1].item_count);
    try testing.expectEqual(Axis.page, def.fields[2].axis.?);
    try testing.expectEqual(@as(u32, 1), def.fields[2].item_count);
    try testing.expect(!def.fields[3].show_all);

    try testing.expectEqual(@as(usize, 1), def.row_fields.len);
    try testing.expectEqual(@as(u32, 1), def.row_fields[0].field);
    try testing.expectEqual(@as(usize, 1), def.col_fields.len);
    try testing.expect(def.col_fields[0] == .values);
    try testing.expectEqual(@as(usize, 1), def.page_fields.len);
    try testing.expectEqual(@as(u32, 2), def.page_fields[0].fld.field);
    try testing.expectEqual(@as(?i32, -1), def.page_fields[0].hier);

    try testing.expectEqual(@as(usize, 2), def.data_fields.len);
    try testing.expectEqualStrings("Average of mpg", def.data_fields[0].name.?);
    try testing.expectEqual(@as(u32, 0), def.data_fields[0].fld);
    try testing.expectEqual(ConsolidateFunction.average, def.data_fields[0].subtotal);
    try testing.expectEqual(ConsolidateFunction.sum, def.data_fields[1].subtotal);
    try testing.expectEqual(@as(u32, 3), def.data_fields[1].fld);

    const style = def.style.?;
    try testing.expectEqualStrings("PivotStyleLight16", style.name.?);
    try testing.expect(style.show_row_headers);
    try testing.expect(!style.show_row_stripes);
    try testing.expect(style.show_last_column);
}

test "parseTableDefinition: unknown subtotal / axis spellings are carried raw; the values page field" {
    const xml =
        \\<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="P" cacheId="0"><location ref="A1" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="1"><pivotField axis="axisFuture"/></pivotFields><pageFields count="1"><pageField fld="-2"/></pageFields><dataFields count="1"><dataField fld="0" subtotal="median"/></dataFields></pivotTableDefinition>
    ;
    var def = try parseTableDefinition(testing.allocator, xml);
    defer def.deinit(testing.allocator);
    try testing.expect(def.fields[0].axis == null);
    try testing.expectEqualStrings("axisFuture", def.fields[0].axis_raw.?);
    try testing.expect(def.page_fields[0].fld == .values);
    try testing.expectEqual(ConsolidateFunction.unknown, def.data_fields[0].subtotal);
    try testing.expectEqualStrings("median", def.data_fields[0].subtotal_raw.?);
    try testing.expectEqualStrings("unknown", ConsolidateFunction.unknown.xmlName());
    try testing.expectEqualStrings("countNums", ConsolidateFunction.count_nums.xmlName());
}

test "parseTableDefinition: refuses a missing location, name or cacheId, a bad ordinal, a doubled block" {
    try testing.expectError(error.MalformedXml, parseTableDefinition(testing.allocator,
        \\<pivotTableDefinition name="P" cacheId="0"><pivotFields count="0"/></pivotTableDefinition>
    ));
    try testing.expectError(error.MalformedXml, parseTableDefinition(testing.allocator,
        \\<pivotTableDefinition cacheId="0"><location ref="A1"/></pivotTableDefinition>
    ));
    try testing.expectError(error.MalformedXml, parseTableDefinition(testing.allocator,
        \\<pivotTableDefinition name="P"><location ref="A1"/></pivotTableDefinition>
    ));
    try testing.expectError(error.MalformedXml, parseTableDefinition(testing.allocator,
        \\<pivotTableDefinition name="P" cacheId="abc"><location ref="A1"/></pivotTableDefinition>
    ));
    try testing.expectError(error.MalformedXml, parseTableDefinition(testing.allocator,
        \\<pivotTableDefinition name="P" cacheId="0"><location ref="A1"/><rowFields count="1"><field x="-7"/></rowFields></pivotTableDefinition>
    ));
    try testing.expectError(error.MalformedXml, parseTableDefinition(testing.allocator,
        \\<pivotTableDefinition name="P" cacheId="0"><location ref="A1"/><pageFields count="1"><pageField fld="-7"/></pageFields></pivotTableDefinition>
    ));
    try testing.expectError(error.MalformedXml, parseTableDefinition(testing.allocator,
        \\<pivotTableDefinition name="P" cacheId="0" rowGrandTotals="maybe"><location ref="A1"/></pivotTableDefinition>
    ));
    try testing.expectError(error.MalformedXml, parseTableDefinition(testing.allocator,
        \\<pivotTableDefinition name="P" cacheId="0"><location ref="A1"/><location ref="B2"/></pivotTableDefinition>
    ));
    try testing.expectError(error.MalformedXml, parseTableDefinition(testing.allocator,
        \\<pivotTableDefinition name="P" cacheId="0"><location ref="A1"/><dataFields/><dataFields/></pivotTableDefinition>
    ));
    try testing.expectError(error.MalformedXml, parseTableDefinition(testing.allocator,
        \\<pivotCacheDefinition><cacheSource type="worksheet"/></pivotCacheDefinition>
    ));
}

test "parseTableDefinition: an unclosed location is not a location, nor is a nested one" {
    try testing.expectError(error.MalformedXml, parseTableDefinition(testing.allocator,
        \\<pivotTableDefinition name="P" cacheId="0"><location ref="A1"></pivotTableDefinition>
    ));
    try testing.expectError(error.MalformedXml, parseTableDefinition(testing.allocator,
        \\<pivotTableDefinition name="P" cacheId="0" bad="<location ref='Z9'/>"></pivotTableDefinition>
    ));
    // A `<location>` inside an extension is a descendant, not the location.
    try testing.expectError(error.MalformedXml, parseTableDefinition(testing.allocator,
        \\<pivotTableDefinition name="P" cacheId="0"><extLst><ext><location ref="Z9"/></ext></extLst></pivotTableDefinition>
    ));
}

test "parseTableDefinition: a location ref inside a comment is not the location" {
    const xml =
        \\<pivotTableDefinition name="P" cacheId="0"><!-- <location ref="Z9"/> --><location ref="B2:C3"/></pivotTableDefinition>
    ;
    var def = try parseTableDefinition(testing.allocator, xml);
    defer def.deinit(testing.allocator);
    try testing.expectEqualStrings("B2:C3", def.location.ref);
    try testing.expectEqualStrings("B2:C3", xml[def.location.ref_span.start..def.location.ref_span.end]);
}

test "Children: a same-name element nested inside is a descendant; closes pair by depth" {
    // Items are counted under `<items>` only; an `<items>` nested in an
    // `<item>` is a descendant of that item, not a second list.
    const xml =
        \\<pivotTableDefinition name="P" cacheId="0"><location ref="A1"/><pivotFields count="1"><pivotField axis="axisRow"><items count="2"><item x="0"><items><item x="9"/></items></item><item x="1"/></items></pivotField></pivotFields></pivotTableDefinition>
    ;
    var def = try parseTableDefinition(testing.allocator, xml);
    defer def.deinit(testing.allocator);
    try testing.expectEqual(@as(usize, 1), def.fields.len);
    try testing.expectEqual(@as(u32, 2), def.fields[0].item_count);

    const root = try scanRoot(xml, "pivotTableDefinition");
    var kids = Children.init(xml, root.hit, root.body_end, "");
    const loc = (try kids.next()).?;
    try testing.expectEqualStrings("location", loc.local);
    const pf = (try kids.next()).?;
    try testing.expectEqualStrings("pivotFields", pf.local);
    try testing.expectEqualStrings("</pivotFields>", xml[pf.end .. pf.end + "</pivotFields>".len]);
    try testing.expect((try kids.next()) == null);
}

test "nsAttr: bound prefix wins, xmlns declarations never match, bare name never matches, a rebound prefix is foreign" {
    const attrs =
        \\ xmlns:r="urn:r" id="bare" xmlns:id="urn:id" rel:id="rId7"
    ;
    try testing.expectEqualStrings("rId7", nsAttr(attrs, null, "id").?);
    try testing.expect(nsAttr(attrs, "r", "id") == null);
    try testing.expect(nsAttr(" id=\"bare\" xmlns:id=\"u\"", null, "id") == null);
    try testing.expect(nsAttr(" xmlns:r=\"urn:vendor\" r:id=\"bad\"", "r", "id") == null);
}

fn parseCacheForFailures(allocator: Allocator, xml: []const u8) !void {
    var def = try parseCacheDefinition(allocator, xml);
    def.deinit(allocator);
}

fn parseTableForFailures(allocator: Allocator, xml: []const u8) !void {
    var def = try parseTableDefinition(allocator, xml);
    def.deinit(allocator);
}

test "allocation failure at every point leaves nothing behind" {
    const cons =
        \\<pivotCacheDefinition><cacheSource type="consolidation"><consolidation><rangeSets><rangeSet sheet="Q1" ref="A1:B9"/><rangeSet name="Q2"/></rangeSets></consolidation></cacheSource><cacheFields count="2"><cacheField name="A"><sharedItems/></cacheField><cacheField name="B"/></cacheFields></pivotCacheDefinition>
    ;
    try testing.checkAllAllocationFailures(testing.allocator, parseCacheForFailures, .{cache_def_worksheet});
    try testing.checkAllAllocationFailures(testing.allocator, parseCacheForFailures, .{cons});
    try testing.checkAllAllocationFailures(testing.allocator, parseTableForFailures, .{table_def_xml});
}
