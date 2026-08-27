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
//! Lifetime contract: every `[]const u8` field borrows from the `xml`
//! handed to the parser and is **not decoded** — attribute values are
//! still XML-escaped and, for the ST_Xstring-typed names, still
//! `_xHHHH_`-encoded. Decoding is the caller's, by site
//! (`pkg/pivots.zig` does it with the engine's `decodeAt`). Each parsed
//! struct owns an arena for its spines (`fields`, `row_fields`, …).
//!
//! Namespace prefixes: the SpreadsheetML main namespace may be bound to
//! a prefix (`<x:pivotCacheDefinition xmlns:x="…">`), and Strict OOXML
//! binds a different URI to it. The parser reads the root element's
//! prefix once and matches every child under it; it never matches on
//! the URI, so Strict and Transitional parts parse alike. Comments,
//! CDATA and processing instructions are skipped through the shared
//! decoy-aware scanner in `workbook_xml.zig`.

const std = @import("std");
const assert = std.debug.assert;
const wbxml = @import("workbook_xml.zig");

pub const Error = error{
    OutOfMemory,
    /// The part is not the pivot part it was handed as, a required
    /// element or attribute is missing, or the markup is unbalanced.
    MalformedXml,
};

/// Absolute byte span inside the `xml` handed to the parser:
/// `xml[start..end]` is the attribute value (no quotes).
pub const Span = struct {
    start: usize,
    end: usize,
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
    arena: std.heap.ArenaAllocator,

    /// `r:id` on the root — the relationship to `pivotCacheRecordsN.xml`.
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

    pub fn deinit(self: *CacheDefinition) void {
        self.arena.deinit();
        self.* = undefined;
    }
};

/// Parse one `pivotCacheDefinitionN.xml` part.
pub fn parseCacheDefinition(allocator: std.mem.Allocator, xml: []const u8) Error!CacheDefinition {
    assert(xml.len < (1 << 31));
    const root = try findRoot(xml, "pivotCacheDefinition");
    const attrs = xml[root.hit.attrs_start..root.hit.attrs_end];
    const p = root.prefix;

    // The arena lives inside `def` from the start: an `ArenaAllocator`
    // carries its state by value, so an allocator taken from a local
    // copy would grow a different arena than the one `deinit` frees.
    var def: CacheDefinition = .{
        .arena = std.heap.ArenaAllocator.init(allocator),
        .r_id = nsAttr(attrs, "id"),
        .record_count = u32Attr(attrs, "recordCount"),
        .refreshed_by = wbxml.getAttr(attrs, "refreshedBy"),
        .refreshed_date = wbxml.getAttr(attrs, "refreshedDate"),
        .refreshed_date_iso = wbxml.getAttr(attrs, "refreshedDateIso"),
        .refresh_on_load = boolAttr(attrs, "refreshOnLoad", false),
        .save_data = boolAttr(attrs, "saveData", true),
        .enable_refresh = boolAttr(attrs, "enableRefresh", true),
        .invalid = boolAttr(attrs, "invalid", false),
        .background_query = boolAttr(attrs, "backgroundQuery", false),
        .created_version = u32Attr(attrs, "createdVersion"),
        .refreshed_version = u32Attr(attrs, "refreshedVersion"),
        .min_refreshable_version = u32Attr(attrs, "minRefreshableVersion"),
        .missing_items_limit = u32Attr(attrs, "missingItemsLimit"),
        .source = .{},
        .fields = &.{},
        .fields_count_attr = null,
    };
    errdefer def.arena.deinit();
    const a = def.arena.allocator();
    const body_end = root.body_end;

    // `<cacheSource>` is required by the schema; a definition without
    // one describes no data and cannot be a pivot's cache.
    const src_hit = (try findTag(xml[0..body_end], root.hit.after_tag_close, p, "cacheSource")) orelse
        return error.MalformedXml;
    def.source = try parseCacheSource(a, xml, src_hit, p);

    if (try findTag(xml[0..body_end], root.hit.after_tag_close, p, "cacheFields")) |cf_hit| {
        const cf_attrs = xml[cf_hit.attrs_start..cf_hit.attrs_end];
        def.fields_count_attr = u32Attr(cf_attrs, "count");
        if (!cf_hit.self_closing) {
            const cf_end = try closeOf(xml, cf_hit, p, "cacheFields");
            def.fields = try parseCacheFields(a, xml, cf_hit.after_tag_close, cf_end, p);
        }
    }

    return def;
}

fn parseCacheSource(
    a: std.mem.Allocator,
    xml: []const u8,
    hit: wbxml.TagHit,
    p: []const u8,
) Error!CacheSource {
    const attrs = xml[hit.attrs_start..hit.attrs_end];
    var src: CacheSource = .{
        .connection_id = u32Attr(attrs, "connectionId"),
    };
    if (wbxml.getAttr(attrs, "type")) |t| {
        src.type = std.meta.stringToEnum(SourceType, t) orelse .unknown;
        if (src.type == .unknown) src.type_raw = t;
    }
    if (hit.self_closing) return src;
    const end = try closeOf(xml, hit, p, "cacheSource");
    const region = xml[0..end];

    if (try findTag(region, hit.after_tag_close, p, "worksheetSource")) |ws| {
        src.worksheet = parseWorksheetSource(xml, ws);
    }
    if (try findTag(region, hit.after_tag_close, p, "rangeSets")) |rs| {
        if (!rs.self_closing) {
            const rs_end = try closeOf(xml, rs, p, "rangeSets");
            var sets: std.ArrayListUnmanaged(WorksheetSource) = .empty;
            var cursor = rs.after_tag_close;
            while (try findTag(xml[0..rs_end], cursor, p, "rangeSet")) |set_hit| {
                try sets.append(a, parseWorksheetSource(xml, set_hit));
                cursor = set_hit.after_tag_close;
            }
            src.range_sets = try sets.toOwnedSlice(a);
        }
    }
    return src;
}

fn parseWorksheetSource(xml: []const u8, hit: wbxml.TagHit) WorksheetSource {
    const attrs = xml[hit.attrs_start..hit.attrs_end];
    var ws: WorksheetSource = .{
        .name = wbxml.getAttr(attrs, "name"),
        .sheet = wbxml.getAttr(attrs, "sheet"),
        .r_id = nsAttr(attrs, "id"),
    };
    if (wbxml.getAttr(attrs, "ref")) |r| {
        ws.ref = r;
        ws.ref_span = spanOf(xml, r);
    }
    return ws;
}

fn parseCacheFields(
    a: std.mem.Allocator,
    xml: []const u8,
    from: usize,
    end: usize,
    p: []const u8,
) Error![]CacheField {
    var out: std.ArrayListUnmanaged(CacheField) = .empty;
    var cursor = from;
    while (try findTag(xml[0..end], cursor, p, "cacheField")) |hit| {
        const attrs = xml[hit.attrs_start..hit.attrs_end];
        var field: CacheField = .{
            .name = wbxml.getAttr(attrs, "name") orelse return error.MalformedXml,
            .caption = wbxml.getAttr(attrs, "caption"),
            .num_fmt_id = u32Attr(attrs, "numFmtId"),
            .formula = wbxml.getAttr(attrs, "formula"),
            .database_field = boolAttr(attrs, "databaseField", true),
        };
        cursor = hit.after_tag_close;
        if (!hit.self_closing) {
            const field_end = try closeOf(xml, hit, p, "cacheField");
            if (try findTag(xml[0..field_end], hit.after_tag_close, p, "sharedItems")) |si| {
                field.shared_items = parseSharedItems(xml[si.attrs_start..si.attrs_end]);
            }
            cursor = field_end;
        }
        try out.append(a, field);
    }
    return out.toOwnedSlice(a);
}

fn parseSharedItems(attrs: []const u8) SharedItems {
    return .{
        .count = u32Attr(attrs, "count"),
        .contains_semi_mixed_types = boolAttr(attrs, "containsSemiMixedTypes", true),
        .contains_non_date = boolAttr(attrs, "containsNonDate", true),
        .contains_date = boolAttr(attrs, "containsDate", false),
        .contains_string = boolAttr(attrs, "containsString", true),
        .contains_blank = boolAttr(attrs, "containsBlank", false),
        .contains_mixed_types = boolAttr(attrs, "containsMixedTypes", false),
        .contains_number = boolAttr(attrs, "containsNumber", false),
        .contains_integer = boolAttr(attrs, "containsInteger", false),
        .long_text = boolAttr(attrs, "longText", false),
        .min_value = wbxml.getAttr(attrs, "minValue"),
        .max_value = wbxml.getAttr(attrs, "maxValue"),
        .min_date = wbxml.getAttr(attrs, "minDate"),
        .max_date = wbxml.getAttr(attrs, "maxDate"),
    };
}

// ─── Table definition ────────────────────────────────────────────────

/// `CT_Location` — where the pivot renders on its host sheet.
pub const Location = struct {
    /// `ref` (ST_Ref, required) — the output rectangle.
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
    /// Number of `<item>` children — the field's distinct items as the
    /// pivot last saw them.
    item_count: u32 = 0,
};

/// One entry of `<rowFields>` / `<colFields>`: `<field x="N"/>`. The
/// values axis is spelled `x="-2"`.
pub const AxisField = union(enum) {
    field: u32,
    values,
};

pub const PageField = struct {
    /// `fld` — the pivot field ordinal, or `-2` for the values axis.
    fld: i32,
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
    arena: std.heap.ArenaAllocator,

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

    pub fn deinit(self: *TableDefinition) void {
        self.arena.deinit();
        self.* = undefined;
    }
};

/// Parse one `pivotTableN.xml` part.
pub fn parseTableDefinition(allocator: std.mem.Allocator, xml: []const u8) Error!TableDefinition {
    assert(xml.len < (1 << 31));
    const root = try findRoot(xml, "pivotTableDefinition");
    const attrs = xml[root.hit.attrs_start..root.hit.attrs_end];
    const p = root.prefix;
    const body_end = root.body_end;
    const from = root.hit.after_tag_close;

    // `<location>` is required: without it the pivot has no output
    // rectangle, and the output rectangle is the one thing every
    // consumer of this part — the S7a lift included — exists for.
    const loc_hit = (try findTag(xml[0..body_end], from, p, "location")) orelse
        return error.MalformedXml;
    const loc_attrs = xml[loc_hit.attrs_start..loc_hit.attrs_end];
    const loc_ref = wbxml.getAttr(loc_attrs, "ref") orelse return error.MalformedXml;
    if (loc_ref.len == 0) return error.MalformedXml;

    // Same arena-by-value rule as `parseCacheDefinition`.
    var def: TableDefinition = .{
        .arena = std.heap.ArenaAllocator.init(allocator),
        .name = wbxml.getAttr(attrs, "name") orelse return error.MalformedXml,
        .cache_id = u32Attr(attrs, "cacheId") orelse return error.MalformedXml,
        .data_caption = wbxml.getAttr(attrs, "dataCaption"),
        .grand_total_caption = wbxml.getAttr(attrs, "grandTotalCaption"),
        .data_on_rows = boolAttr(attrs, "dataOnRows", false),
        .data_position = u32Attr(attrs, "dataPosition"),
        .row_grand_totals = boolAttr(attrs, "rowGrandTotals", true),
        .col_grand_totals = boolAttr(attrs, "colGrandTotals", true),
        .compact = boolAttr(attrs, "compact", true),
        .outline = boolAttr(attrs, "outline", false),
        .compact_data = boolAttr(attrs, "compactData", true),
        .outline_data = boolAttr(attrs, "outlineData", false),
        .created_version = u32Attr(attrs, "createdVersion"),
        .updated_version = u32Attr(attrs, "updatedVersion"),
        .min_refreshable_version = u32Attr(attrs, "minRefreshableVersion"),
        .location = .{
            .ref = loc_ref,
            .ref_span = spanOf(xml, loc_ref),
            .first_header_row = u32Attr(loc_attrs, "firstHeaderRow"),
            .first_data_row = u32Attr(loc_attrs, "firstDataRow"),
            .first_data_col = u32Attr(loc_attrs, "firstDataCol"),
            .row_page_count = u32Attr(loc_attrs, "rowPageCount"),
            .col_page_count = u32Attr(loc_attrs, "colPageCount"),
        },
        .fields = &.{},
        .row_fields = &.{},
        .col_fields = &.{},
        .page_fields = &.{},
        .data_fields = &.{},
        .style = null,
    };
    errdefer def.arena.deinit();
    const a = def.arena.allocator();

    if (try findTag(xml[0..body_end], from, p, "pivotFields")) |hit| {
        if (!hit.self_closing) {
            const end = try closeOf(xml, hit, p, "pivotFields");
            def.fields = try parsePivotFields(a, xml, hit.after_tag_close, end, p);
        }
    }
    def.row_fields = try parseAxisFields(a, xml, body_end, from, p, "rowFields");
    def.col_fields = try parseAxisFields(a, xml, body_end, from, p, "colFields");
    if (try findTag(xml[0..body_end], from, p, "pageFields")) |hit| {
        if (!hit.self_closing) {
            const end = try closeOf(xml, hit, p, "pageFields");
            def.page_fields = try parsePageFields(a, xml, hit.after_tag_close, end, p);
        }
    }
    if (try findTag(xml[0..body_end], from, p, "dataFields")) |hit| {
        if (!hit.self_closing) {
            const end = try closeOf(xml, hit, p, "dataFields");
            def.data_fields = try parseDataFields(a, xml, hit.after_tag_close, end, p);
        }
    }
    if (try findTag(xml[0..body_end], from, p, "pivotTableStyleInfo")) |hit| {
        const s = xml[hit.attrs_start..hit.attrs_end];
        def.style = .{
            .name = wbxml.getAttr(s, "name"),
            .show_row_headers = boolAttr(s, "showRowHeaders", false),
            .show_col_headers = boolAttr(s, "showColHeaders", false),
            .show_row_stripes = boolAttr(s, "showRowStripes", false),
            .show_col_stripes = boolAttr(s, "showColStripes", false),
            .show_last_column = boolAttr(s, "showLastColumn", false),
        };
    }
    return def;
}

fn parsePivotFields(
    a: std.mem.Allocator,
    xml: []const u8,
    from: usize,
    end: usize,
    p: []const u8,
) Error![]PivotField {
    var out: std.ArrayListUnmanaged(PivotField) = .empty;
    var cursor = from;
    while (try findTag(xml[0..end], cursor, p, "pivotField")) |hit| {
        const attrs = xml[hit.attrs_start..hit.attrs_end];
        var f: PivotField = .{
            .name = wbxml.getAttr(attrs, "name"),
            .data_field = boolAttr(attrs, "dataField", false),
            .show_all = boolAttr(attrs, "showAll", true),
            .default_subtotal = boolAttr(attrs, "defaultSubtotal", true),
            .num_fmt_id = u32Attr(attrs, "numFmtId"),
            .subtotal_caption = wbxml.getAttr(attrs, "subtotalCaption"),
        };
        if (wbxml.getAttr(attrs, "axis")) |ax| {
            f.axis = axisFromXml(ax);
            if (f.axis == null) f.axis_raw = ax;
        }
        cursor = hit.after_tag_close;
        if (!hit.self_closing) {
            const field_end = try closeOf(xml, hit, p, "pivotField");
            var item_cursor = hit.after_tag_close;
            while (try findTag(xml[0..field_end], item_cursor, p, "item")) |item| {
                f.item_count += 1;
                item_cursor = item.after_tag_close;
            }
            cursor = field_end;
        }
        try out.append(a, f);
    }
    return out.toOwnedSlice(a);
}

fn axisFromXml(s: []const u8) ?Axis {
    if (std.mem.eql(u8, s, "axisRow")) return .row;
    if (std.mem.eql(u8, s, "axisCol")) return .col;
    if (std.mem.eql(u8, s, "axisPage")) return .page;
    if (std.mem.eql(u8, s, "axisValues")) return .values;
    return null;
}

fn parseAxisFields(
    a: std.mem.Allocator,
    xml: []const u8,
    body_end: usize,
    from: usize,
    p: []const u8,
    block: []const u8,
) Error![]AxisField {
    const hit = (try findTag(xml[0..body_end], from, p, block)) orelse return &.{};
    if (hit.self_closing) return &.{};
    const end = try closeOf(xml, hit, p, block);
    var out: std.ArrayListUnmanaged(AxisField) = .empty;
    var cursor = hit.after_tag_close;
    while (try findTag(xml[0..end], cursor, p, "field")) |f| {
        const attrs = xml[f.attrs_start..f.attrs_end];
        const x = i32Attr(attrs, "x") orelse return error.MalformedXml;
        // `-2` is the values axis; any other negative ordinal names no
        // field and cannot be read.
        if (x == -2) {
            try out.append(a, .values);
        } else if (x >= 0) {
            try out.append(a, .{ .field = @intCast(x) });
        } else {
            return error.MalformedXml;
        }
        cursor = f.after_tag_close;
    }
    return out.toOwnedSlice(a);
}

fn parsePageFields(
    a: std.mem.Allocator,
    xml: []const u8,
    from: usize,
    end: usize,
    p: []const u8,
) Error![]PageField {
    var out: std.ArrayListUnmanaged(PageField) = .empty;
    var cursor = from;
    while (try findTag(xml[0..end], cursor, p, "pageField")) |hit| {
        const attrs = xml[hit.attrs_start..hit.attrs_end];
        try out.append(a, .{
            .fld = i32Attr(attrs, "fld") orelse return error.MalformedXml,
            .item = u32Attr(attrs, "item"),
            .hier = i32Attr(attrs, "hier"),
            .name = wbxml.getAttr(attrs, "name"),
            .cap = wbxml.getAttr(attrs, "cap"),
        });
        cursor = if (hit.self_closing) hit.after_tag_close else try closeOf(xml, hit, p, "pageField");
    }
    return out.toOwnedSlice(a);
}

fn parseDataFields(
    a: std.mem.Allocator,
    xml: []const u8,
    from: usize,
    end: usize,
    p: []const u8,
) Error![]DataField {
    var out: std.ArrayListUnmanaged(DataField) = .empty;
    var cursor = from;
    while (try findTag(xml[0..end], cursor, p, "dataField")) |hit| {
        const attrs = xml[hit.attrs_start..hit.attrs_end];
        var df: DataField = .{
            .name = wbxml.getAttr(attrs, "name"),
            .fld = u32Attr(attrs, "fld") orelse return error.MalformedXml,
            .show_data_as = wbxml.getAttr(attrs, "showDataAs"),
            .base_field = i32Attr(attrs, "baseField"),
            .base_item = u32Attr(attrs, "baseItem"),
            .num_fmt_id = u32Attr(attrs, "numFmtId"),
        };
        if (wbxml.getAttr(attrs, "subtotal")) |s| {
            df.subtotal = ConsolidateFunction.fromXml(s);
            if (df.subtotal == .unknown) df.subtotal_raw = s;
        }
        try out.append(a, df);
        cursor = if (hit.self_closing) hit.after_tag_close else try closeOf(xml, hit, p, "dataField");
    }
    return out.toOwnedSlice(a);
}

// ─── Scanner glue ────────────────────────────────────────────────────

/// Longest namespace prefix this parser will bind. Prefixes in real
/// parts are one to three characters; the bound exists so the tag-name
/// scratch buffers are fixed-size, and a part that exceeds it is
/// refused rather than truncated into a wrong match.
const max_prefix_len = 32;
const max_local_len = 40;

const Root = struct {
    hit: wbxml.TagHit,
    /// The root element's namespace prefix, without the colon; empty
    /// when the main namespace is the default.
    prefix: []const u8,
    /// Index of the `<` of the root's closing tag — the bound every
    /// child search runs under. For a self-closing root, the end of
    /// the root tag itself.
    body_end: usize,
};

/// Locate the root element, check its local name, and learn its
/// prefix. Anything before it (XML declaration, comments) is skipped.
fn findRoot(xml: []const u8, local: []const u8) Error!Root {
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
        const body_end = if (hit.self_closing)
            hit.after_tag_close
        else
            try closeOf(xml, hit, prefix, local);
        return .{ .hit = hit, .prefix = prefix, .body_end = body_end };
    }
    return error.MalformedXml;
}

fn isNameBoundary(c: u8) bool {
    return c == ' ' or c == '\t' or c == '\n' or c == '\r' or c == '/' or c == '>';
}

/// `<prefix:local` search bounded by `region.len`, decoy-aware.
fn findTag(region: []const u8, from: usize, prefix: []const u8, local: []const u8) Error!?wbxml.TagHit {
    assert(local.len <= max_local_len);
    if (from >= region.len) return null;
    var buf: [max_prefix_len + 1 + max_local_len]u8 = undefined;
    const tag = qualify(&buf, prefix, local);
    return wbxml.findTagOpen(region, from, tag) catch |e| return narrow(e);
}

/// The `<` of `</prefix:local>` matching an open tag at `hit`, or
/// `MalformedXml` when the element never closes. The search starts
/// after the open tag; nesting of the same element name does not occur
/// in these parts (`cacheField` never contains a `cacheField`), so the
/// first decoy-free close is the right one.
fn closeOf(xml: []const u8, hit: wbxml.TagHit, prefix: []const u8, local: []const u8) Error!usize {
    assert(!hit.self_closing);
    var buf: [2 + max_prefix_len + 1 + max_local_len + 1]u8 = undefined;
    buf[0] = '<';
    buf[1] = '/';
    const q = qualify(buf[2..], prefix, local);
    buf[2 + q.len] = '>';
    const close_tag = buf[0 .. 2 + q.len + 1];
    const close = wbxml.findClosingTag(xml, hit.after_tag_close, close_tag) catch |e| return narrow(e);
    return close orelse error.MalformedXml;
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

/// An attribute in a *foreign* namespace, matched on its local name
/// under any prefix (`r:id`, `rel:id`, …). `xmlns:*` declarations are
/// not attributes in this sense and never match.
pub fn nsAttr(attrs: []const u8, local: []const u8) ?[]const u8 {
    var i: usize = 0;
    while (i < attrs.len) {
        while (i < attrs.len and std.ascii.isWhitespace(attrs[i])) i += 1;
        if (i >= attrs.len) break;
        const name_start = i;
        while (i < attrs.len and attrs[i] != '=' and !std.ascii.isWhitespace(attrs[i])) i += 1;
        const name = attrs[name_start..i];
        while (i < attrs.len and (attrs[i] == '=' or std.ascii.isWhitespace(attrs[i]))) i += 1;
        if (i >= attrs.len) break;
        if (attrs[i] != '"' and attrs[i] != '\'') break;
        const quote = attrs[i];
        i += 1;
        const val_start = i;
        while (i < attrs.len and attrs[i] != quote) i += 1;
        const val = attrs[val_start..i];
        if (i < attrs.len) i += 1;

        if (std.mem.indexOfScalar(u8, name, ':')) |c| {
            const prefix = name[0..c];
            if (prefix.len > 0 and !std.mem.eql(u8, prefix, "xmlns") and
                std.mem.eql(u8, name[c + 1 ..], local))
            {
                return val;
            }
        }
    }
    return null;
}

fn u32Attr(attrs: []const u8, name: []const u8) ?u32 {
    const v = wbxml.getAttr(attrs, name) orelse return null;
    return std.fmt.parseInt(u32, v, 10) catch null;
}

fn i32Attr(attrs: []const u8, name: []const u8) ?i32 {
    const v = wbxml.getAttr(attrs, name) orelse return null;
    return std.fmt.parseInt(i32, v, 10) catch null;
}

/// xsd:boolean: `1`/`true` and `0`/`false`; anything else keeps the
/// schema default rather than guessing.
fn boolAttr(attrs: []const u8, name: []const u8, default: bool) bool {
    const v = wbxml.getAttr(attrs, name) orelse return default;
    if (std.mem.eql(u8, v, "1") or std.mem.eql(u8, v, "true")) return true;
    if (std.mem.eql(u8, v, "0") or std.mem.eql(u8, v, "false")) return false;
    return default;
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
    defer def.deinit();
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
    defer def.deinit();
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
    defer e.deinit();
    try testing.expectEqual(SourceType.external, e.source.type);
    try testing.expectEqual(@as(?u32, 3), e.source.connection_id);
    try testing.expect(e.source.worksheet == null);

    const cons =
        \\<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="consolidation"><consolidation autoPage="0"><rangeSets count="2"><rangeSet i1="0" sheet="Q1" ref="A1:B9"/><rangeSet i1="1" name="Q2Data"/></rangeSets></consolidation></cacheSource></pivotCacheDefinition>
    ;
    var c = try parseCacheDefinition(testing.allocator, cons);
    defer c.deinit();
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
    defer def.deinit();
    try testing.expectEqualStrings("rId2", def.source.worksheet.?.r_id.?);
}

test "parseCacheDefinition: prefixed main namespace (Strict-style binding)" {
    const xml =
        \\<?xml version="1.0"?>
        \\<x:pivotCacheDefinition xmlns:x="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:rel="http://purl.oclc.org/ooxml/officeDocument/relationships" rel:id="rId9" recordCount="2"><x:cacheSource type="worksheet"><x:worksheetSource sheet="S" ref="A1:A3"/></x:cacheSource><x:cacheFields count="1"><x:cacheField name="A" numFmtId="0"><x:sharedItems/></x:cacheField></x:cacheFields></x:pivotCacheDefinition>
    ;
    var def = try parseCacheDefinition(testing.allocator, xml);
    defer def.deinit();
    try testing.expectEqualStrings("rId9", def.r_id.?);
    try testing.expectEqualStrings("S", def.source.worksheet.?.sheet.?);
    try testing.expectEqual(@as(usize, 1), def.fields.len);
    try testing.expectEqualStrings("A", def.fields[0].name);
}

test "parseCacheDefinition: decoys in comments and CDATA are not elements" {
    const xml =
        \\<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><!-- <cacheSource type="external"/> --><cacheSource type="worksheet"><!-- <worksheetSource sheet="Decoy" ref="Z1"/> --><worksheetSource sheet="Real" ref="A1:B2"/></cacheSource><cacheFields count="1"><![CDATA[<cacheField name="Decoy"/>]]><cacheField name="Real"/></cacheFields></pivotCacheDefinition>
    ;
    var def = try parseCacheDefinition(testing.allocator, xml);
    defer def.deinit();
    try testing.expectEqual(SourceType.worksheet, def.source.type);
    try testing.expectEqualStrings("Real", def.source.worksheet.?.sheet.?);
    try testing.expectEqual(@as(usize, 1), def.fields.len);
    try testing.expectEqualStrings("Real", def.fields[0].name);
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

const table_def_xml =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable3" cacheId="1" dataCaption="Values" updatedVersion="5" minRefreshableVersion="3" createdVersion="5" outline="1" outlineData="1" multipleFieldFilters="0" chartFormat="1"><location ref="A1:D5" firstHeaderRow="0" firstDataRow="1" firstDataCol="1"/><pivotFields count="4"><pivotField dataField="1" showAll="0"/><pivotField axis="axisRow" showAll="0"><items count="4"><item x="1"/><item x="0"/><item x="2"/><item t="default"/></items></pivotField><pivotField axis="axisPage" showAll="0"><items count="1"><item t="default"/></items></pivotField><pivotField showAll="0"/></pivotFields><rowFields count="1"><field x="1"/></rowFields><rowItems count="2"><i><x/></i><i t="grand"><x/></i></rowItems><colFields count="1"><field x="-2"/></colFields><pageFields count="1"><pageField fld="2" hier="-1"/></pageFields><dataFields count="2"><dataField name="Average of mpg" fld="0" subtotal="average" baseField="1" baseItem="0"/><dataField name="Sum of wt" fld="3" baseField="1" baseItem="0"/></dataFields><chartFormats count="1"><chartFormat chart="0" format="0" series="1"><pivotArea type="data" outline="0" fieldPosition="0"><references count="1"><reference field="4294967294" count="1" selected="0"><x v="0"/></reference></references></pivotArea></chartFormat></chartFormats><pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/><extLst><ext uri="{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:pivotTableDefinition hideValuesRow="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"/></ext></extLst></pivotTableDefinition>
;

test "parseTableDefinition: name, cache, location, roles, axes, data fields, style" {
    var def = try parseTableDefinition(testing.allocator, table_def_xml);
    defer def.deinit();
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
    try testing.expectEqual(@as(i32, 2), def.page_fields[0].fld);
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

test "parseTableDefinition: unknown subtotal / axis spellings are carried raw" {
    const xml =
        \\<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="P" cacheId="0"><location ref="A1" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="1"><pivotField axis="axisFuture"/></pivotFields><dataFields count="1"><dataField fld="0" subtotal="median"/></dataFields></pivotTableDefinition>
    ;
    var def = try parseTableDefinition(testing.allocator, xml);
    defer def.deinit();
    try testing.expect(def.fields[0].axis == null);
    try testing.expectEqualStrings("axisFuture", def.fields[0].axis_raw.?);
    try testing.expectEqual(ConsolidateFunction.unknown, def.data_fields[0].subtotal);
    try testing.expectEqualStrings("median", def.data_fields[0].subtotal_raw.?);
    try testing.expectEqualStrings("unknown", ConsolidateFunction.unknown.xmlName());
    try testing.expectEqualStrings("countNums", ConsolidateFunction.count_nums.xmlName());
}

test "parseTableDefinition: refuses a missing location, name or cacheId, and a bad axis ordinal" {
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
        \\<pivotTableDefinition name="P" cacheId="0"><location ref="A1"/><rowFields count="1"><field x="-7"/></rowFields></pivotTableDefinition>
    ));
    try testing.expectError(error.MalformedXml, parseTableDefinition(testing.allocator,
        \\<pivotCacheDefinition><cacheSource type="worksheet"/></pivotCacheDefinition>
    ));
}

test "parseTableDefinition: a location ref inside a comment is not the location" {
    const xml =
        \\<pivotTableDefinition name="P" cacheId="0"><!-- <location ref="Z9"/> --><location ref="B2:C3"/></pivotTableDefinition>
    ;
    var def = try parseTableDefinition(testing.allocator, xml);
    defer def.deinit();
    try testing.expectEqualStrings("B2:C3", def.location.ref);
    try testing.expectEqualStrings("B2:C3", xml[def.location.ref_span.start..def.location.ref_span.end]);
}

test "nsAttr: matches a foreign-namespace local name, never xmlns declarations or the bare name" {
    const attrs =
        \\ xmlns:r="urn:r" id="bare" xmlns:id="urn:id" rel:id="rId7"
    ;
    try testing.expectEqualStrings("rId7", nsAttr(attrs, "id").?);
    try testing.expect(nsAttr(" id=\"bare\" xmlns:id=\"u\"", "id") == null);
}
