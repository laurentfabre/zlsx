//! S6 (`goal_sigmoid.md`): the pivot graph, read-only.
//!
//! A pivot is four parts held together by relationships, none of which
//! the worksheet body names:
//!
//!     xl/workbook.xml  <pivotCaches><pivotCache cacheId r:id>   ─┐
//!                                                                 ▼
//!     xl/pivotCache/pivotCacheDefinitionN.xml  ── r:id ──▶  pivotCacheRecordsN.xml
//!            │  <cacheSource><worksheetSource sheet= ref= | name= | r:id=>
//!            ▼
//!     (the SOURCE sheet, a table on it, a defined name, or another workbook)
//!
//!     xl/worksheets/sheetK.xml  ── rels(pivotTable) ──▶  xl/pivotTables/pivotTableM.xml
//!            (the HOST sheet)                              │  cacheId=, rels(pivotCacheDefinition)
//!                                                          ▼  <location ref=>  on the host
//!
//! `collect` walks that graph from both roots — the workbook's cache
//! list and every sheet's relationships — and hands back one typed
//! `Pivots` view: every pivot table with its host sheet, output
//! rectangle, field roles and axes; every cache with its source, its
//! field schema and, when the source is local, **which sheet it reads
//! from** and how that was established (`SourceResolution`). The raw
//! part bytes ride along untouched; nothing here writes.
//!
//! Source resolution is the audit half of S6 (surface-matrix footnote
//! ¹⁰): the `#139` guard refuses row/col edits on sheets that *host* a
//! pivot because the host's relationships name the pivot part; a sheet
//! a pivot only *reads from* has no such edge, so the guard cannot see
//! it. The only way to see it is to read `worksheetSource` and resolve
//! its spelling the way Excel does — `sheet` is a sheet name (case-
//! folded, like every sheet qualifier), `name` is a table or a defined
//! name (one namespace, folded), `r:id` is another workbook. That
//! resolution is what `Pivots.readsFromSheet` answers, and what S7b
//! will guard on.
//!
//! Relationship types are matched on their trailing segment, so the
//! Strict (`purl.oclc.org`) and Transitional (`schemas.openxmlformats.org`)
//! URIs both resolve; the part parsers bind the main namespace by prefix
//! rather than by URI for the same reason.

const std = @import("std");
const assert = std.debug.assert;
const Allocator = std.mem.Allocator;
const zlsx = @import("zlsx");
const formula = @import("zlsx_formula");
const store_mod = @import("store.zig");
const PartStore = store_mod.PartStore;
const wbxml = @import("typed_parts/workbook_xml.zig");
const pivot_xml = @import("typed_parts/pivot_xml.zig");
const table_edit = @import("table_edit.zig");
const sheet_edit = @import("sheet_edit.zig");
const coords = @import("zlsx_refs");
const workbook_mod = @import("workbook.zig");
const recalc_run = @import("recalc_run.zig");

pub const Error = store_mod.Error || error{
    /// The pivot graph cannot be read whole: a part a recognised
    /// relationship names is missing or unreadable (wrong root, a
    /// required element or attribute missing or unreadable, unbalanced
    /// markup, a name that fails its carrier's decode), a `<pivotCache>`
    /// entry lacks its `cacheId` or `r:id`, a cache's identity disagrees
    /// between `xl/workbook.xml` and the pivot that reads it, two caches
    /// claim one id, or a records part is named but absent. The whole
    /// read refuses rather than reporting the pivots it could parse — a
    /// partial inventory is the shape of every guard hole this row
    /// exists to close.
    MalformedPivotXml,
};

/// How a local source sheet was established.
pub const ResolvedVia = enum {
    /// `worksheetSource@sheet` named it directly.
    sheet_attr,
    /// `worksheetSource@name` is a table; the sheet is the table's host.
    table,
    /// `worksheetSource@name` is a defined name whose body is a
    /// single sheet-qualified reference.
    defined_name,
};

/// The area a local source reads — what a static spelling proves.
/// `null` on `LocalSheet.bounds` for a spelling that names a sheet and
/// nothing on it (a `sheet` attribute alone, a `ref` no rectangle
/// parser accepts).
pub const Bounds = union(enum) {
    /// A finite rectangle: a direct `ref`, a table's `ref`, a name body
    /// such as `Data!$A$1:$C$4` or `Data!$A$1`.
    rect: edit.Rect,
    /// Whole columns, 1-based and inclusive: `Data!$A:$C`.
    whole_columns: struct { first_col: u32, last_col: u32 },
    /// Whole rows, 1-based and inclusive: `Data!$1:$4`.
    whole_rows: struct { first_row: u32, last_row: u32 },

    /// The A1 spelling — `A1:C4`, `A:C`, `1:4`. `buf` holds any of them
    /// at `format_buf_len`. Null for a value outside the grid (a
    /// column or row of 0, or past `XFD` / 1048576): the parsers here
    /// never produce one, but the type is public (Codex #202 r4 F4).
    pub fn formatA1(self: Bounds, buf: *[format_buf_len]u8) ?[]const u8 {
        if (!self.inGrid()) return null;
        return switch (self) {
            .rect => |r| blk: {
                var tl: [coords.max_col_letters]u8 = undefined;
                var br: [coords.max_col_letters]u8 = undefined;
                const tl_len = coords.writeColNumberLetters(&tl, r.tl_col) catch return null;
                const br_len = coords.writeColNumberLetters(&br, r.br_col) catch return null;
                break :blk std.fmt.bufPrint(buf, "{s}{d}:{s}{d}", .{ tl[0..tl_len], r.tl_row, br[0..br_len], r.br_row }) catch return null;
            },
            .whole_columns => |c| blk: {
                var f: [coords.max_col_letters]u8 = undefined;
                var l: [coords.max_col_letters]u8 = undefined;
                const f_len = coords.writeColNumberLetters(&f, c.first_col) catch return null;
                const l_len = coords.writeColNumberLetters(&l, c.last_col) catch return null;
                break :blk std.fmt.bufPrint(buf, "{s}:{s}", .{ f[0..f_len], l[0..l_len] }) catch return null;
            },
            .whole_rows => |r| std.fmt.bufPrint(buf, "{d}:{d}", .{ r.first_row, r.last_row }) catch return null,
        };
    }

    pub fn eql(a: Bounds, b: Bounds) bool {
        return switch (a) {
            .rect => |r| b == .rect and r.eql(b.rect),
            .whole_columns => |c| b == .whole_columns and c.first_col == b.whole_columns.first_col and c.last_col == b.whole_columns.last_col,
            .whole_rows => |r| b == .whole_rows and r.first_row == b.whole_rows.first_row and r.last_row == b.whole_rows.last_row,
        };
    }

    /// 1-based and inside the grid, corners ordered.
    pub fn inGrid(self: Bounds) bool {
        return switch (self) {
            .rect => |r| r.tl_col >= 1 and r.tl_row >= 1 and r.br_col <= zlsx.max_col_1based and r.br_row <= zlsx.max_row and r.tl_col <= r.br_col and r.tl_row <= r.br_row,
            .whole_columns => |c| c.first_col >= 1 and c.last_col <= zlsx.max_col_1based and c.first_col <= c.last_col,
            .whole_rows => |r| r.first_row >= 1 and r.last_row <= zlsx.max_row and r.first_row <= r.last_row,
        };
    }

    pub const format_buf_len = 32;
};

/// A defined name a source reads through, as `<definedName>` spells
/// it: the decoded identifier and the `localSheetId` scope. What the
/// S7b guard dry-runs the name sweep on — the sweep moves such a body
/// with the grid, and where it would spell `#REF!` (a deleted endpoint
/// or anchor) the edit refuses instead of leaving the source unresolved.
pub const NameKey = struct {
    identifier: []const u8,
    scope: ?u32,
};

/// Why a worksheet-type source could not be placed, and what it does
/// prove — the S7b guard's evidence for a source it cannot bound.
pub const Unresolved = struct {
    why: Why,
    /// Sheets of this workbook the spelling or its name closure names —
    /// ascending, without duplicates: the local `sheet` beside an
    /// unplaceable `r:id`; every sheet qualifier (a 3D span: each sheet
    /// between its ends, in tab order) and every table's host reachable
    /// from a name body through the names it references. Empty when
    /// nothing local is proven.
    sheets: []const u32,
    /// The defined names the spelling reads through — the source name
    /// and every name its body references, transitively (the same walk
    /// as `sheets`). Empty unless `why` is `unbounded_body`.
    names: []const NameKey = &.{},

    pub const Why = enum {
        /// `sheet` names no sheet of this workbook.
        dangling_sheet,
        /// `name` is neither a defined name nor a table.
        dangling_name,
        /// `name` is a defined name whose body is not one static
        /// sheet-qualified area — dynamic, 3D, a union, a bare range,
        /// or reaching its sheet only through another name.
        unbounded_body,
        /// `r:id` names no External-mode external-link relationship.
        unplaceable_rid,
        /// `ref` without `sheet` or `name`: a range on no sheet.
        sheetless_ref,
        /// Neither `sheet`, `name` nor `ref`, or a `worksheet`-type
        /// source without its `<worksheetSource>` child.
        no_locator,
    };
};

pub const SourceResolution = union(enum) {
    /// A sheet of this workbook.
    sheet: LocalSheet,
    /// Another workbook: the relationship target, as written.
    external: []const u8,
    /// A worksheet-type source whose spelling names nothing this
    /// workbook has — a dangling sheet name, a name with a dynamic or
    /// 3D body, a missing relationship. Excel would fail the refresh.
    /// The payload says why, and which sheets the spelling still proves.
    unresolved: Unresolved,
    /// Not a worksheet-type source (external connection, scenario,
    /// unknown), so there is no sheet to resolve.
    none,

    pub const LocalSheet = struct {
        sheet_idx: u32,
        /// Decoded sheet name.
        sheet_name: []const u8,
        part_name: []const u8,
        via: ResolvedVia,
        /// The area read, when the spelling proves one (S7b's rectangle).
        bounds: ?Bounds,
        /// What `bounds` came from, and so what a row edit moves.
        /// `.none` exactly when `bounds` is null.
        carrier: SourceCarrier = .none,
        /// The defined names the spelling reads through (`NameKey`):
        /// the source name and its closure for a defined-name source,
        /// or a name beside a `sheet` attribute; empty for a direct
        /// `ref` and for a table.
        names: []const NameKey = &.{},
        /// Rows at the top of `bounds` that are field names rather
        /// than data: 1 for a direct `ref` and a name body (Excel
        /// reads a range source's first row as its headers), the
        /// table's `headerRowCount` for a table — 0 for a headerless
        /// one, whose field names come from `<tableColumns>`. What
        /// the S7b-4 rebuild splits the rectangle by.
        header_rows: u32 = 1,
        /// Rows at the bottom of `bounds` that are a table's totals
        /// (`totalsRowCount`), not data; 0 for every other carrier
        /// (Codex #205 r4 REL-401).
        totals_rows: u32 = 0,
        /// A table's `<tableColumn>` names, decoded, in order — a
        /// headerless table's field names (Codex #205 r4 REL-402);
        /// empty for every other carrier.
        columns: []const []const u8 = &.{},
        /// The table part behind a `.table` carrier — where S7c reads
        /// the column name a table's own column insert will synthesize;
        /// null for every other carrier.
        table_part_name: ?[]const u8 = null,
    };
};

/// Does one resolution depend on this sheet — resolved to it, or
/// unresolved with the sheet among what its spelling proves? The per-
/// resolution half of `Pivots.dependsOnSheet`.
pub fn resolutionDependsOn(r: SourceResolution, sheet_idx: u32) bool {
    return switch (r) {
        .sheet => |s| s.sheet_idx == sheet_idx,
        .unresolved => |u| std.mem.indexOfScalar(u32, u.sheets, sheet_idx) != null,
        .external, .none => false,
    };
}

/// The defined names a resolution reads through, whichever arm holds
/// them — empty for the arms that read through none.
pub fn namesOf(r: SourceResolution) []const NameKey {
    return switch (r) {
        .sheet => |s| s.names,
        .unresolved => |u| u.names,
        .external, .none => &.{},
    };
}

/// What carries a local source's coordinate — the thing a row edit
/// moves. Distinct from `ResolvedVia`, which says what named the
/// SHEET: `<worksheetSource sheet="Data" name="SalesTbl"/>` is placed
/// by its `sheet` attribute and bounded by a table part (Codex #203
/// r1 REL-102).
pub const SourceCarrier = enum {
    /// The spelling's own `ref` — spliced at `ref_span`.
    ref,
    /// A table part's `ref` — `table_edit` moves it, with the table's
    /// own header knowledge.
    table,
    /// A defined name's body — the name sweep moves it.
    defined_name,
    /// Nothing bounds the source: `sheet` alone, a name the reader
    /// could not place on that sheet.
    none,
};

/// The decoded spellings of one `worksheetSource` / `rangeSet` — what
/// a reader shows next to the resolution the spellings led to.
pub const SourceSpelling = struct {
    /// Decoded `sheet`.
    sheet: ?[]const u8 = null,
    /// Decoded `name` (a table or defined-name spelling).
    name: ?[]const u8 = null,
    /// `ref`, lexical.
    ref: ?[]const u8 = null,
};

pub const PivotCache = struct {
    /// `<pivotCache cacheId>` from `xl/workbook.xml`. Null when the part
    /// is reachable only through a pivot table's own relationship —
    /// Excel would not load such a cache, but the read still names it.
    cache_id: ?u32,
    part_name: []const u8,
    /// The `pivotCacheRecordsN.xml` part the definition's `r:id` names,
    /// when it exists.
    records_part_name: ?[]const u8,
    /// Spines live in the `Pivots` arena — never `deinit` this.
    definition: pivot_xml.CacheDefinition,
    /// Decoded (plain-text) field names, parallel to `definition.fields`.
    field_names: []const []const u8,
    /// Decoded calculated-field formulas (FORMULA carrier: entities
    /// only), parallel to `definition.fields`; null for a plain field.
    field_formulas: []const ?[]const u8,
    /// The decoded `worksheetSource` spellings, for
    /// `source.type == .worksheet`.
    source: SourceSpelling,
    /// Where the data comes from, for `source.type == .worksheet`.
    resolution: SourceResolution,
    /// For a consolidation source, one per `definition.source.range_sets`.
    range_set_sources: []const SourceSpelling,
    range_set_resolutions: []const SourceResolution,
    /// How many pivot tables read this cache.
    consumer_count: u32,
    /// The definition part's bytes, borrowed from the store.
    raw_xml: []const u8,
};

pub const PivotTable = struct {
    /// Decoded `name`.
    name: []const u8,
    part_name: []const u8,
    /// The host sheet — the one whose relationships name this part and
    /// whose grid `definition.location.ref` addresses.
    sheet_idx: u32,
    /// Decoded host sheet name.
    sheet_name: []const u8,
    sheet_part_name: []const u8,
    /// Spines live in the `Pivots` arena — never `deinit` this.
    definition: pivot_xml.TableDefinition,
    /// `definition.location.ref`, entity-decoded — the lexical value a
    /// reader shows; the raw slice and its span stay for the splice.
    location_ref: []const u8,
    /// Decoded captions.
    data_caption: ?[]const u8,
    grand_total_caption: ?[]const u8,
    /// Decoded `pivotTableStyleInfo@name`.
    style_name: ?[]const u8,
    /// Decoded `dataField@name`, parallel to `definition.data_fields`.
    data_field_names: []const ?[]const u8,
    /// Index into `Pivots.caches`, or null when neither the part's
    /// relationship nor its `cacheId` reaches a cache.
    cache: ?usize,
    /// The definition part's bytes, borrowed from the store.
    raw_xml: []const u8,
};

pub const Pivots = struct {
    arena: std.heap.ArenaAllocator,
    /// In host-sheet order, then relationship order within a sheet.
    tables: []const PivotTable,
    /// In `<pivotCaches>` order, then discovery order for caches only a
    /// pivot table's relationship reaches.
    caches: []const PivotCache,
    /// Every sheet's decoded name, in workbook order — what `sheet_idx`
    /// indexes, so a consumer can name sheets without a second decode.
    sheet_names: []const []const u8,
    /// Every sheet's part name, parallel to `sheet_names` — resolved
    /// and checked by the walk, so an editor can go from the part it
    /// is about to rewrite to the index every resolution carries.
    sheet_parts: []const []const u8,

    pub fn deinit(self: *Pivots) void {
        self.arena.deinit();
        self.* = undefined;
    }

    pub fn cacheOf(self: *const Pivots, table: PivotTable) ?*const PivotCache {
        const i = table.cache orelse return null;
        return &self.caches[i];
    }

    /// The decoded name of a pivot field by ordinal — the cache field
    /// at the same position, which is what every `fld` / `x` indexes.
    pub fn fieldName(self: *const Pivots, table: PivotTable, ordinal: u32) ?[]const u8 {
        const c = self.cacheOf(table) orelse return null;
        if (ordinal >= c.field_names.len) return null;
        return c.field_names[ordinal];
    }

    /// Does a pivot table render on this sheet? The `#139` guard's
    /// question, answered from the graph rather than the rels scan.
    pub fn hostsPivot(self: *const Pivots, sheet_idx: u32) bool {
        for (self.tables) |t| if (t.sheet_idx == sheet_idx) return true;
        return false;
    }

    /// The index of the sheet a part name belongs to, or null when no
    /// sheet of the workbook is that part.
    pub fn sheetIndexOfPart(self: *const Pivots, part_name: []const u8) ?u32 {
        for (self.sheet_parts, 0..) |p, i| {
            if (std.mem.eql(u8, p, part_name)) return @intCast(i);
        }
        return null;
    }

    /// Does any cache read its data from this sheet? The footnote-¹⁰
    /// question: true for a sheet a pivot only *reads from*, which the
    /// `#139` guard does not see.
    pub fn readsFromSheet(self: *const Pivots, sheet_idx: u32) bool {
        for (self.caches) |c| {
            if (resolvesTo(c.resolution, sheet_idx)) return true;
            for (c.range_set_resolutions) |r| if (resolvesTo(r, sheet_idx)) return true;
        }
        return false;
    }

    /// The S7b guard's selection: does any cache *provably* depend on
    /// this sheet — a source resolved to it, or an unresolved one whose
    /// evidence names it (the local `sheet` beside an unplaceable
    /// `r:id`; a name body reaching the sheet through the names it
    /// references)? Between `readsFromSheet` (resolved only) and
    /// `mayReadFromSheet` (every sheet, once a source proves nothing).
    pub fn dependsOnSheet(self: *const Pivots, sheet_idx: u32) bool {
        for (self.caches) |c| {
            if (dependsOn(c.resolution, sheet_idx)) return true;
            for (c.range_set_resolutions) |r| if (dependsOn(r, sheet_idx)) return true;
        }
        return false;
    }

    /// The S7a guard's question, which is wider than `readsFromSheet`:
    /// could any cache read this sheet? True when a source resolves to
    /// it, and — conservatively — when a source is `unresolved` (a
    /// defined name with a dynamic body such as `OFFSET(Report!$D$1,…)`
    /// is a source Excel accepts and this reader cannot place; a
    /// dangling spelling is one it cannot rule out either) or when a
    /// cache's source type is one the reader does not know. "Not
    /// proven local" is not "proven elsewhere" (Codex #200 r1 REL-033).
    pub fn mayReadFromSheet(self: *const Pivots, sheet_idx: u32) bool {
        for (self.caches) |c| {
            if (c.definition.source.type == .unknown) return true;
            if (mayResolveTo(c.resolution, sheet_idx)) return true;
            for (c.range_set_resolutions) |r| if (mayResolveTo(r, sheet_idx)) return true;
        }
        return false;
    }

    fn resolvesTo(r: SourceResolution, sheet_idx: u32) bool {
        return switch (r) {
            .sheet => |s| s.sheet_idx == sheet_idx,
            else => false,
        };
    }

    fn dependsOn(r: SourceResolution, sheet_idx: u32) bool {
        return resolutionDependsOn(r, sheet_idx);
    }

    fn mayResolveTo(r: SourceResolution, sheet_idx: u32) bool {
        return switch (r) {
            .sheet => |s| s.sheet_idx == sheet_idx,
            .unresolved => true,
            .external, .none => false,
        };
    }
};

/// Walk the graph. `wb` is the parsed `xl/workbook.xml` the caller
/// already holds (a `Workbook` has one); the walk reads the part bytes
/// again only for `<pivotCaches>`, which the typed view does not carry.
///
/// Raw slices in the result (`raw_xml`, part names, the definitions'
/// fields) borrow from the store's arena and live as long as the store;
/// decoded strings (names, captions, spellings) live in the result's
/// own arena and end with `Pivots.deinit`.
pub fn collect(allocator: Allocator, store: *PartStore, wb: *const wbxml.WorkbookXml) Error!Pivots {
    var out: Pivots = .{
        .arena = std.heap.ArenaAllocator.init(allocator),
        .tables = &.{},
        .caches = &.{},
        .sheet_names = &.{},
        .sheet_parts = &.{},
    };
    errdefer out.arena.deinit();
    const a = out.arena.allocator();

    const sheet_names = try a.alloc([]const u8, wb.sheets.len);
    for (wb.sheets, 0..) |s, i| sheet_names[i] = try decode(a, .sheet_name, s.name);
    out.sheet_names = sheet_names;

    const wb_part = (try store.part("xl/workbook.xml")) orelse return out;
    const wb_rels = store.rels("xl/workbook.xml");

    // Every sheet root must reach its part: a workbook-listed sheet whose
    // relationship dangles or whose part is absent is a broken workbook,
    // and a pivot it hosts or feeds would otherwise vanish from the
    // inventory rather than refuse it (Codex #199 r3 REL-016).
    const sheet_parts = try a.alloc([]const u8, wb.sheets.len);
    for (wb.sheets, 0..) |s, i| {
        const rel = try requiredRel(wb_rels, s.r_id, &sheet_rel_leaves);
        sheet_parts[i] = try requiredTarget(store, "xl/workbook.xml", rel.target);
        _ = try requiredPart(store, sheet_parts[i]);
    }
    out.sheet_parts = sheet_parts;

    var caches: std.ArrayListUnmanaged(CacheSlot) = .empty;
    var tables: std.ArrayListUnmanaged(PivotTable) = .empty;

    // Root 1: the workbook's cache list. Its order is the cache order
    // Excel shows, and `cacheId` is the only place it lives.
    try collectWorkbookCaches(a, store, wb_part.bytes, wb_rels, &caches);

    // Root 2: every sheet's relationships. A pivot table part not linked
    // from a sheet is not a pivot Excel renders; it is not listed.
    for (wb.sheets, 0..) |s, i| {
        _ = s;
        const sheet_idx: u32 = @intCast(i);
        const sheet_part = sheet_parts[i];
        for (store.rels(sheet_part)) |rel| {
            if (!relLeafIs(rel.type, "pivotTable")) continue;
            // Once the relationship type says "pivot table", the edge is
            // part of the graph: an External target, a target that does
            // not resolve or a part that is not there is a broken
            // workbook, not a pivot to leave out of the inventory.
            if (rel.target_mode == .external) return error.MalformedPivotXml;
            const pt_name = try requiredTarget(store, sheet_part, rel.target);
            const pt_part = try requiredPart(store, pt_name);
            const def = pivot_xml.parseTableDefinition(a, pt_part.bytes) catch |e| return mapParse(e);

            // The cache, from both edges Excel has: the part's own
            // relationship, and `cacheId` against the workbook list. When
            // both exist they must name the same cache; a relationship to
            // a cache the workbook lists under another id, or lists not at
            // all while the id names a different one, is a pivot Excel
            // could not refresh (Codex #199 r2 REL-007).
            var rel_cache: ?usize = null;
            for (store.rels(pt_name)) |prel| {
                if (!relLeafIs(prel.type, "pivotCacheDefinition")) continue;
                if (prel.target_mode == .external) return error.MalformedPivotXml;
                // One cache per pivot: a second edge of the type is a
                // pivot that reads two caches, which is not a pivot.
                if (rel_cache != null) return error.MalformedPivotXml;
                const cd_name = try requiredTarget(store, pt_name, prel.target);
                _ = try requiredPart(store, cd_name);
                rel_cache = try findOrAddCache(a, &caches, cd_name, null);
            }
            var id_cache: ?usize = null;
            for (caches.items, 0..) |c, ci| {
                if (c.cache_id != null and c.cache_id.? == def.cache_id) {
                    id_cache = ci;
                    break;
                }
            }
            const cache_idx: ?usize = blk: {
                const rc = rel_cache orelse break :blk id_cache;
                if (id_cache) |ic| {
                    if (ic != rc) return error.MalformedPivotXml;
                } else if (caches.items[rc].cache_id != null) {
                    return error.MalformedPivotXml;
                }
                break :blk rc;
            };

            const data_field_names = try a.alloc(?[]const u8, def.data_fields.len);
            for (def.data_fields, 0..) |df, k| {
                data_field_names[k] = if (df.name) |n| try decode(a, .pivot_field_name, n) else null;
            }

            try tables.append(a, .{
                .name = try decode(a, .pivot_table_name, def.name),
                .part_name = pt_name,
                .location_ref = try decodeLexical(a, def.location.ref),
                .sheet_idx = sheet_idx,
                .sheet_name = sheet_names[i],
                .sheet_part_name = sheet_part,
                .definition = def,
                .data_caption = if (def.data_caption) |c| try decode(a, .pivot_table_name, c) else null,
                .grand_total_caption = if (def.grand_total_caption) |c| try decode(a, .pivot_table_name, c) else null,
                .style_name = if (def.style) |st| (if (st.name) |n| try decode(a, .pivot_table_name, n) else null) else null,
                .data_field_names = data_field_names,
                .cache = cache_idx,
                .raw_xml = pt_part.bytes,
            });
            if (cache_idx) |ci| caches.items[ci].consumer_count += 1;
        }
    }

    // Every cache: parse, find the records part, resolve the source.
    var resolver: Resolver = .{
        .gpa = allocator,
        .arena = a,
        .store = store,
        .wb = wb,
        .wb_xml = wb_part.bytes,
        .sheet_names = sheet_names,
        .sheet_parts = sheet_parts,
    };
    defer resolver.deinit();

    const finished = try a.alloc(PivotCache, caches.items.len);
    for (caches.items, 0..) |slot, ci| {
        const part = try requiredPart(store, slot.part_name);
        const def = pivot_xml.parseCacheDefinition(a, part.bytes) catch |e| return mapParse(e);

        // The records part: the definition's `r:id` names it, and only
        // that — a definition without one (`saveData="0"`) has none,
        // whatever relationships the part happens to carry. Named and
        // absent is a refusal: the count says there are records, and
        // there is nothing to hold them.
        const cache_rels = store.rels(slot.part_name);
        var records: ?[]const u8 = null;
        if (def.r_id) |rid| {
            const rel = try requiredRel(cache_rels, rid, &.{"pivotCacheRecords"});
            const r = try requiredTarget(store, slot.part_name, rel.target);
            _ = try requiredPart(store, r);
            // One records part per definition: two naming the same one
            // would each rebuild it in their own image, and the one
            // installed last would hold the other's records under its
            // own inventories (Codex #205 r3 REL-301).
            for (finished[0..ci]) |prev| {
                if (prev.records_part_name) |taken| {
                    if (std.mem.eql(u8, taken, r)) return error.MalformedPivotXml;
                }
            }
            records = r;
        }

        const field_names = try a.alloc([]const u8, def.fields.len);
        const field_formulas = try a.alloc(?[]const u8, def.fields.len);
        for (def.fields, 0..) |f, k| {
            field_names[k] = try decode(a, .pivot_field_name, f.name);
            field_formulas[k] = if (f.formula) |raw| try decode(a, .cell_formula_body, raw) else null;
        }

        var source_spelling: SourceSpelling = .{};
        var resolution: SourceResolution = .none;
        var set_spellings: []SourceSpelling = &.{};
        var set_resolutions: []SourceResolution = &.{};
        switch (def.source.type) {
            .worksheet => {
                if (def.source.worksheet) |ws| {
                    source_spelling = try spell(a, ws);
                    resolution = try resolver.resolve(ws, cache_rels);
                } else {
                    resolution = .{ .unresolved = .{ .why = .no_locator, .sheets = &.{} } };
                }
            },
            .consolidation => {
                set_spellings = try a.alloc(SourceSpelling, def.source.range_sets.len);
                set_resolutions = try a.alloc(SourceResolution, def.source.range_sets.len);
                for (def.source.range_sets, 0..) |rs, k| {
                    set_spellings[k] = try spell(a, rs);
                    set_resolutions[k] = try resolver.resolve(rs, cache_rels);
                }
            },
            // A locator carried under a `type` this reader does not know
            // is authoritative (S7b gate, Q5): whatever the type means,
            // a `<worksheetSource>` or `<rangeSet>` names what it names,
            // and an edit of that sheet must move or refuse it. Without
            // one the cache names no sheet — `.none`, as before.
            .unknown => {
                if (def.source.worksheet) |ws| {
                    source_spelling = try spell(a, ws);
                    resolution = try resolver.resolve(ws, cache_rels);
                }
                if (def.source.range_sets.len > 0) {
                    set_spellings = try a.alloc(SourceSpelling, def.source.range_sets.len);
                    set_resolutions = try a.alloc(SourceResolution, def.source.range_sets.len);
                    for (def.source.range_sets, 0..) |rs, k| {
                        set_spellings[k] = try spell(a, rs);
                        set_resolutions[k] = try resolver.resolve(rs, cache_rels);
                    }
                }
            },
            .external, .scenario => {},
        }

        finished[ci] = .{
            .cache_id = slot.cache_id,
            .part_name = slot.part_name,
            .records_part_name = records,
            .definition = def,
            .field_names = field_names,
            .field_formulas = field_formulas,
            .source = source_spelling,
            .resolution = resolution,
            .range_set_sources = set_spellings,
            .range_set_resolutions = set_resolutions,
            .consumer_count = slot.consumer_count,
            .raw_xml = part.bytes,
        };
    }

    out.caches = finished;
    out.tables = try tables.toOwnedSlice(a);
    return out;
}

const CacheSlot = struct {
    cache_id: ?u32,
    part_name: []const u8,
    consumer_count: u32 = 0,
};

/// `<pivotCaches><pivotCache cacheId="N" r:id="…"/>` in `xl/workbook.xml`.
/// Read with the part parsers' own scanner — the workbook root's
/// prefix, its relationships binding, the container's close — so only
/// the children of the main-namespace `<pivotCaches>` count; a
/// `<vendor:pivotCache>` inside `<extLst>` is not a cache. Both
/// attributes are required by the schema and here: an entry that
/// cannot say which cache it is, or where, breaks the list.
fn collectWorkbookCaches(
    a: Allocator,
    store: *PartStore,
    wb_xml: []const u8,
    wb_rels: []const store_mod.Relationship,
    caches: *std.ArrayListUnmanaged(CacheSlot),
) Error!void {
    const root = pivot_xml.scanRoot(wb_xml, "workbook") catch |e| return mapParse(e);
    var kids = pivot_xml.Children.init(wb_xml, root.hit, root.body_end, root.prefix, root.env);
    var seen_container = false;
    while (kids.next() catch |e| return mapParse(e)) |k| {
        if (!std.mem.eql(u8, k.local, "pivotCaches")) continue;
        if (seen_container) return error.MalformedPivotXml;
        seen_container = true;
        var entries = pivot_xml.Children.init(wb_xml, k.hit, k.end, root.prefix, k.env);
        while (entries.next() catch |e| return mapParse(e)) |c| {
            if (!std.mem.eql(u8, c.local, "pivotCache")) continue;
            const attrs = c.attrs(wb_xml);
            const rid = pivot_xml.nsAttr(attrs, c.env, "id") orelse return error.MalformedPivotXml;
            const cache_id = (pivot_xml.u32Attr(attrs, "cacheId") catch |e| return mapParse(e)) orelse
                return error.MalformedPivotXml;
            const rel = try requiredRel(wb_rels, rid, &.{"pivotCacheDefinition"});
            const part_name = try requiredTarget(store, "xl/workbook.xml", rel.target);
            _ = try requiredPart(store, part_name);
            _ = try findOrAddCache(a, caches, part_name, cache_id);
        }
    }
}

/// The relationship types a `<sheet r:id>` may carry.
const sheet_rel_leaves = [_][]const u8{ "worksheet", "chartsheet", "dialogsheet", "xlMacrosheet", "xlIntlMacrosheet" };

/// The one way a raw `r:id` becomes a relationship: entity-decoded
/// (the store decodes ids the same way), found by id, of one of the
/// expected types, and internal. Anything else on a recognised edge
/// is a broken graph, not an absent one (Codex #199 r3 REL-015).
fn requiredRel(
    rels: []const store_mod.Relationship,
    raw_rid: []const u8,
    leaves: []const []const u8,
) Error!store_mod.Relationship {
    const rel = relById(rels, raw_rid) orelse return error.MalformedPivotXml;
    var typed = false;
    for (leaves) |leaf| typed = typed or relLeafIs(rel.type, leaf);
    if (!typed) return error.MalformedPivotXml;
    if (rel.target_mode == .external) return error.MalformedPivotXml;
    return rel;
}

/// The relationship a raw `r:id` names, by decoded id, or null.
fn relById(rels: []const store_mod.Relationship, raw_rid: []const u8) ?store_mod.Relationship {
    var buf: [128]u8 = undefined;
    const rid = wbxml.decodeScalarAttr(&buf, raw_rid) orelse return null;
    if (rid.len == 0) return null;
    for (rels) |rel| {
        if (std.mem.eql(u8, rel.id, rid)) return rel;
    }
    return null;
}

/// A relationship target that must resolve to a part name.
fn requiredTarget(store: *PartStore, owner: []const u8, target: []const u8) Error![]const u8 {
    return (try store.resolve(owner, target)) orelse error.MalformedPivotXml;
}

/// A part that must exist, under a name a reader can show. Materialises it.
fn requiredPart(store: *PartStore, name: []const u8) Error!store_mod.Part {
    if (!std.unicode.utf8ValidateSlice(name)) return error.MalformedPivotXml;
    return (try store.part(name)) orelse error.MalformedPivotXml;
}

/// One slot per cache part, one id per slot: a part listed under two
/// ids, or an id given to two parts, is an identity Excel cannot
/// resolve and is refused.
fn findOrAddCache(
    a: Allocator,
    caches: *std.ArrayListUnmanaged(CacheSlot),
    part_name: []const u8,
    cache_id: ?u32,
) Error!usize {
    for (caches.items, 0..) |c, i| {
        if (std.mem.eql(u8, c.part_name, part_name)) {
            if (cache_id) |id| {
                if (c.cache_id) |have| {
                    if (have != id) return error.MalformedPivotXml;
                } else {
                    caches.items[i].cache_id = id;
                }
            }
            return i;
        }
        if (cache_id != null and c.cache_id != null and c.cache_id.? == cache_id.?) {
            return error.MalformedPivotXml;
        }
    }
    try caches.append(a, .{ .cache_id = cache_id, .part_name = part_name });
    return caches.items.len - 1;
}

/// Does `xl/workbook.xml` list a cache — a main-namespace
/// `<pivotCaches>` under the root, by the same scanner `collect` reads
/// it with? The S7b guard's cheap gate, beside the relationship types:
/// a `<pivotCaches>` entry whose relationship is absent or mistyped is
/// a graph `collect` refuses, and the refusal must be reachable from a
/// sheet that hosts nothing (Codex #203 r1 REL-101). A root the
/// scanner cannot read answers true for the same reason — the walk,
/// not this gate, decides what it means.
pub fn workbookListsCaches(wb_xml: []const u8) bool {
    const root = pivot_xml.scanRoot(wb_xml, "workbook") catch return true;
    var kids = pivot_xml.Children.init(wb_xml, root.hit, root.body_end, root.prefix, root.env);
    while (kids.next() catch return true) |k| {
        if (std.mem.eql(u8, k.local, "pivotCaches")) return true;
    }
    return false;
}

/// Trailing-segment match, case-insensitive: the one thing the Strict
/// and Transitional relationship URIs share. Same rule as
/// `Workbook.preflightPivotEditsForSheet`.
pub fn relLeafIs(rel_type: []const u8, leaf: []const u8) bool {
    const l = if (std.mem.lastIndexOfScalar(u8, rel_type, '/')) |i| rel_type[i + 1 ..] else rel_type;
    return std.ascii.eqlIgnoreCase(l, leaf);
}

/// The extension namespaces a slicer / timeline cache definition
/// lives in — x14 and x15. The attachment reader admits exactly
/// these; a part in any other spelling refuses the read, and the
/// caller refuses with it.
const slicer_ns_uris = [_][]const u8{"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"};
const timeline_ns_uris = [_][]const u8{"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"};

pub const AttachmentKind = enum { slicer, timeline };

/// The decoded pivot-table names a slicer or timeline cache is
/// attached to — `<pivotTables><pivotTable name="…"/></pivotTables>`
/// under the part's own extension namespace, read with the shared
/// one-prefix scanner (S7c's gate: a schema edit refuses when one of
/// them names a consumer of the edited cache — its `sourceName` is a
/// field this reader does not type, and a removed field would leave a
/// slicer no refresh repairs). Everything else in the part is
/// tolerated unread; a `pivotTable` without a `name`, or a part the
/// scanner refuses, is the caller's refusal.
pub fn attachedPivotNames(arena: Allocator, xml: []const u8, kind: AttachmentKind) Error![]const []const u8 {
    const local: []const u8 = switch (kind) {
        .slicer => "slicerCacheDefinition",
        .timeline => "timelineCacheDefinition",
    };
    const uris: []const []const u8 = switch (kind) {
        .slicer => &slicer_ns_uris,
        .timeline => &timeline_ns_uris,
    };
    const root = pivot_xml.scanRootIn(xml, local, uris, .require_family) catch return error.MalformedPivotXml;
    // A second attachment list can ride the part's `extLst`
    // (`x15:slicerCachePivotTables`, the data-model spelling) — an
    // attachment this reader would not see. Its token anywhere in the
    // part is an attachment list it cannot read, and the caller's
    // refusal is the safe answer (in-house review S7C-R3; a decoy
    // spelling only costs a refusal).
    if (std.mem.indexOf(u8, xml, "slicerCachePivotTables") != null) return error.MalformedPivotXml;
    var out: std.ArrayListUnmanaged([]const u8) = .empty;
    var kids = pivot_xml.Children.init(xml, root.hit, root.body_end, root.prefix, root.env);
    while (kids.next() catch return error.MalformedPivotXml) |k| {
        if (!std.mem.eql(u8, k.local, "pivotTables")) continue;
        var inner = pivot_xml.Children.init(xml, k.hit, k.end, root.prefix, k.env);
        while (inner.next() catch return error.MalformedPivotXml) |pt| {
            if (!std.mem.eql(u8, pt.local, "pivotTable")) continue;
            const raw = wbxml.getAttr(pt.attrs(xml), "name") orelse return error.MalformedPivotXml;
            try out.append(arena, try decode(arena, .pivot_table_name, raw));
        }
    }
    return out.toOwnedSlice(arena);
}

/// A decoded pivot-table name under the collation fold Excel compares
/// them with (unique case-insensitively) — folded ONCE by the caller
/// and matched with `std.mem.eql`, so a gate over N attachment names
/// and M consumers folds N + M times, not N × M (Codex #208 r3
/// PERF-301). Null when the fold cannot spell the name: the caller
/// treats it as matching everything, the refusing direction.
pub fn foldedPivotName(a: Allocator, name: []const u8) error{OutOfMemory}!?[]const u8 {
    return fold(a, name) catch |e| switch (e) {
        error.OutOfMemory => return error.OutOfMemory,
        else => return null,
    };
}

fn spell(a: Allocator, ws: pivot_xml.WorksheetSource) Error!SourceSpelling {
    return .{
        .sheet = if (ws.sheet) |raw| try decode(a, .pivot_source_sheet_name, raw) else null,
        .name = if (ws.name) |raw| try decode(a, .pivot_source_name, raw) else null,
        .ref = if (ws.ref) |raw| try decodeLexical(a, raw) else null,
    };
}

/// Entity decode for a lexical token (ST_Ref): `A1&#58;C4` is `A1:C4`.
/// The raw slice and its `ref_span` stay as written, for the splice.
fn decodeLexical(a: Allocator, raw: []const u8) Error![]const u8 {
    return formula.decode.decodeCarrier(a, .lexical, raw) catch |e| switch (e) {
        error.OutOfMemory => error.OutOfMemory,
        else => error.MalformedPivotXml,
    };
}

fn mapParse(e: pivot_xml.Error) Error {
    return switch (e) {
        error.OutOfMemory => error.OutOfMemory,
        error.MalformedXml => error.MalformedPivotXml,
    };
}

/// STRING-carrier decode (entities + ST_Xstring) at a named site — the
/// codec every name attribute in these parts is written with.
fn decode(a: Allocator, site: formula.decode.Site, raw: []const u8) Error![]const u8 {
    return formula.decode.decodeAt(a, site, raw) catch |e| switch (e) {
        error.OutOfMemory => error.OutOfMemory,
        else => error.MalformedPivotXml,
    };
}

// ─── Source resolution ───────────────────────────────────────────────

/// Resolves `worksheetSource` spellings against the workbook's symbols.
///
/// Two inventories, each built lazily and only when a spelling needs
/// it. The sheet list — every sheet's name, folded by the engine's
/// collation — serves `sheet` spellings and the qualifiers inside a
/// defined name's body; a workbook's sheets are its own and this
/// inventory cannot refuse. The engine's symbol table (sheets + defined
/// names) serves `name` spellings, and it CAN refuse — duplicate
/// symbols, a name inventory the engine will not read — in which case
/// a name-based source is a refusal of the read, not "names nothing":
/// `.unresolved` means the lookup ran and found no such symbol
/// (Codex #199 r2 REL-011).
const Resolver = struct {
    gpa: Allocator,
    arena: Allocator,
    store: *PartStore,
    wb: *const wbxml.WorkbookXml,
    wb_xml: []const u8,
    sheet_names: []const []const u8,
    /// Every sheet's part, resolved and checked by `collect`.
    sheet_parts: []const []const u8,

    sheet_folds: ?[]const []const u8 = null,
    tables: ?[]const TableEntry = null,
    symbols: ?formula.SymbolTable = null,
    symbols_refused: bool = false,

    const TableEntry = struct {
        folded: []const u8,
        sheet_idx: u32,
        /// The table part's `ref`, when it parses as a rectangle.
        bounds: ?Bounds,
        /// The table's `headerRowCount` (default 1).
        header_rows: u32,
        /// The table's `totalsRowCount` (default 0).
        totals_rows: u32,
        /// The table's `<tableColumn>` names, decoded, in order.
        columns: []const []const u8,
        /// The table part's name — S7c reads the part again to plan a
        /// schema edit beside the table's own rewrite.
        part_name: []const u8,

        fn shape(self: TableEntry) TableShape {
            return .{ .header_rows = self.header_rows, .totals_rows = self.totals_rows, .columns = self.columns, .table_part_name = self.part_name };
        }
    };

    fn deinit(self: *Resolver) void {
        if (self.symbols) |*s| s.deinit();
    }

    fn resolve(
        self: *Resolver,
        ws: pivot_xml.WorksheetSource,
        cache_rels: []const store_mod.Relationship,
    ) Error!SourceResolution {
        // Another workbook: the relationship says so, whatever the
        // sheet and range spell — and only an External-mode
        // relationship of the external-link type says so. An internal
        // or differently typed target under `r:id` is not a workbook
        // this reader can name; what such a spelling still proves is
        // the local `sheet` beside it, if any.
        if (ws.r_id) |rid| {
            if (relById(cache_rels, rid)) |rel| {
                if (rel.target_mode == .external and relLeafIs(rel.type, "externalLinkPath")) {
                    // The target is handed to readers as text; bytes that
                    // are not text are refused at this boundary, not
                    // emitted.
                    if (!std.unicode.utf8ValidateSlice(rel.target)) return error.MalformedPivotXml;
                    return .{ .external = rel.target };
                }
            }
            return unresolved(.unplaceable_rid, try self.sheetsOfAttr(ws.sheet), &.{});
        }
        if (ws.sheet) |raw_sheet| {
            const name = try decode(self.arena, .pivot_source_sheet_name, raw_sheet);
            const idx = (try self.sheetIndexOf(name)) orelse return unresolved(.dangling_sheet, &.{}, &.{});
            // `sheet` names the sheet; a `ref` bounds it, else a `name`
            // beside it may (a table or a static name on that same
            // sheet — Codex #202 r1 F4). The sheet wins on identity.
            if (ws.ref != null) return try self.local(idx, .sheet_attr, try self.boundsOfRef(ws.ref), .ref, &.{}, .{});
            if (ws.name) |raw_name| {
                const carrier = try self.carrierBounds(idx, raw_name);
                return try self.local(idx, .sheet_attr, carrier.bounds, carrier.kind, carrier.names, carrier.shape);
            }
            return try self.local(idx, .sheet_attr, null, .none, &.{}, .{});
        }
        if (ws.name) |raw_name| {
            const name = try decode(self.arena, .pivot_source_name, raw_name);
            // The symbol inventory first: a refused inventory refuses
            // every name-based source, table-spelled or not, and a name
            // the engine refuses to reference is a refusal here too.
            // Tables and defined names share one namespace in Excel, so
            // a table is consulted only for a spelling no name has.
            const symbols = try self.ensureSymbols();
            switch (try symbols.resolveName(self.gpa, null, name)) {
                .name => |n| {
                    const cl = try self.closure(n);
                    if (try self.areaOfBody(n.body)) |area| return try self.local(area.sheet_idx, .defined_name, area.bounds, .defined_name, cl.names, .{});
                    return unresolved(.unbounded_body, cl.sheets, cl.names);
                },
                .refused => return error.MalformedPivotXml,
                .table, .not_found => {},
            }
            const folded = (try fold(self.arena, name)) orelse return unresolved(.dangling_name, &.{}, &.{});
            for (try self.ensureTables()) |t| {
                if (std.mem.eql(u8, t.folded, folded)) return try self.local(t.sheet_idx, .table, t.bounds, .table, &.{}, t.shape());
            }
            return unresolved(.dangling_name, &.{}, &.{});
        }
        return unresolved(if (ws.ref != null) .sheetless_ref else .no_locator, &.{}, &.{});
    }

    fn unresolved(why: Unresolved.Why, sheets: []const u32, names: []const NameKey) SourceResolution {
        return .{ .unresolved = .{ .why = why, .sheets = sheets, .names = names } };
    }

    /// A table's shape beyond its bounds: what the rectangle's rows
    /// are, and what its columns are called.
    const TableShape = struct {
        header_rows: u32 = 1,
        totals_rows: u32 = 0,
        columns: []const []const u8 = &.{},
        table_part_name: ?[]const u8 = null,
    };

    fn local(self: *Resolver, sheet_idx: u32, via: ResolvedVia, bounds: ?Bounds, carrier: SourceCarrier, names: []const NameKey, shape: TableShape) Error!SourceResolution {
        if (sheet_idx >= self.sheet_parts.len) return unresolved(.dangling_sheet, &.{}, &.{});
        return .{ .sheet = .{
            .sheet_idx = sheet_idx,
            .sheet_name = self.sheet_names[sheet_idx],
            .part_name = self.sheet_parts[sheet_idx],
            .via = via,
            .bounds = bounds,
            .carrier = if (bounds == null) .none else carrier,
            .names = names,
            .header_rows = shape.header_rows,
            .totals_rows = shape.totals_rows,
            .columns = shape.columns,
            .table_part_name = shape.table_part_name,
        } };
    }

    const Carrier = struct { bounds: ?Bounds, kind: SourceCarrier, names: []const NameKey, shape: TableShape = .{} };

    /// The bounds a `name` beside a `sheet` lends the source, when the
    /// carrier is on that sheet: a static defined-name body, else a
    /// table. Nothing refuses here — the sheet is already placed, and
    /// the name is evidence for the area only. A defined-name carrier
    /// also hands back its closure, bounded or not: its body moves with
    /// the grid, and the S7b guard dry-runs that move.
    fn carrierBounds(self: *Resolver, sheet_idx: u32, raw_name: []const u8) Error!Carrier {
        const none: Carrier = .{ .bounds = null, .kind = .none, .names = &.{} };
        const name = try decode(self.arena, .pivot_source_name, raw_name);
        const symbols = self.ensureSymbols() catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            else => return none,
        };
        // Looked up FROM the stated sheet: its own scoped name shadows a
        // workbook one of the same spelling there (Codex #202 r2 F3).
        switch (try symbols.resolveName(self.gpa, formula.SheetIndex.fromInt(sheet_idx), name)) {
            .name => |n| {
                const cl = try self.closure(n);
                const area = (try self.areaOfBody(n.body)) orelse return .{ .bounds = null, .kind = .none, .names = cl.names };
                const on_sheet = area.sheet_idx == sheet_idx;
                return .{ .bounds = if (on_sheet) area.bounds else null, .kind = if (on_sheet) .defined_name else .none, .names = cl.names };
            },
            .refused => return none,
            .table, .not_found => {},
        }
        const folded = (try fold(self.arena, name)) orelse return none;
        const tables = self.ensureTables() catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            else => return none,
        };
        for (tables) |t| {
            if (!std.mem.eql(u8, t.folded, folded)) continue;
            const on_sheet = t.sheet_idx == sheet_idx;
            return .{ .bounds = if (on_sheet) t.bounds else null, .kind = if (on_sheet) .table else .none, .names = &.{}, .shape = t.shape() };
        }
        return none;
    }

    /// The one sheet a `sheet` attribute names, as evidence: `[idx]`,
    /// or nothing when it names no sheet here.
    fn sheetsOfAttr(self: *Resolver, raw_sheet: ?[]const u8) Error![]const u32 {
        const raw = raw_sheet orelse return &.{};
        const name = try decode(self.arena, .pivot_source_sheet_name, raw);
        const idx = (try self.sheetIndexOf(name)) orelse return &.{};
        const one = try self.arena.alloc(u32, 1);
        one[0] = idx;
        return one;
    }

    /// A direct `ref`, entity-decoded, as bounds — or null when there is
    /// none or it is not a rectangle, whole columns or whole rows.
    fn boundsOfRef(self: *Resolver, raw_ref: ?[]const u8) Error!?Bounds {
        const raw = raw_ref orelse return null;
        const decoded = try decodeLexical(self.arena, raw);
        return parseBounds(decoded);
    }

    /// Sheet lookup by decoded name, case-folded the way every sheet
    /// qualifier is.
    fn sheetIndexOf(self: *Resolver, name: []const u8) Error!?u32 {
        const folds = try self.ensureSheetFolds();
        const folded = (try fold(self.arena, name)) orelse return null;
        for (folds, 0..) |f, i| {
            if (std.mem.eql(u8, f, folded)) return @intCast(i);
        }
        return null;
    }

    fn ensureSheetFolds(self: *Resolver) Error![]const []const u8 {
        if (self.sheet_folds) |f| return f;
        const folds = try self.arena.alloc([]const u8, self.sheet_names.len);
        for (self.sheet_names, 0..) |n, i| {
            // A name the fold rejects (it never does for a name Excel
            // wrote) gets a spelling no query can fold to.
            folds[i] = (try fold(self.arena, n)) orelse "\x00";
        }
        self.sheet_folds = folds;
        return folds;
    }

    /// The table index: every table a sheet ATTACHES through its
    /// `<tableParts>` block, keyed by the folded display name a
    /// `worksheetSource@name` spells. A table relationship no
    /// `<tablePart>` names is not a table of the sheet; a `<tablePart>`
    /// whose relationship is missing, mistyped, external or whose part
    /// is absent is a broken attachment and refuses the read (Codex
    /// #199 r3 REL-018). The block is read by the scanner the table
    /// editor reads it with — `<tableParts>` in the default main
    /// namespace, `r:id` — so a producer that binds a prefix to the
    /// main namespace attaches nothing here, exactly as it attaches
    /// nothing to `renameTableColumn`; the source then resolves to
    /// nothing rather than the wrong sheet (Codex r4 REL-024, declined:
    /// one scanner for one block).
    fn ensureTables(self: *Resolver) Error![]const TableEntry {
        if (self.tables) |t| return t;
        var entries: std.ArrayListUnmanaged(TableEntry) = .empty;
        for (self.sheet_parts, 0..) |sheet_part, i| {
            const sheet = try requiredPart(self.store, sheet_part);
            const rels = self.store.rels(sheet_part);
            var rids = workbook_mod.TablePartRidIterator.init(sheet.bytes);
            while (rids.next()) |rid| {
                const rel = try requiredRel(rels, rid, &.{"table"});
                const table_part_name = try requiredTarget(self.store, sheet_part, rel.target);
                const table_part = try requiredPart(self.store, table_part_name);
                const raw = table_edit.tableDisplayNameRaw(table_part.bytes) orelse return error.MalformedPivotXml;
                const name = try decode(self.arena, .table_name, raw);
                const folded = (try fold(self.arena, name)) orelse continue;
                // The table's `ref` is the source's rectangle — the same
                // scanner and attribute the table editor reads; a `ref`
                // that is not a rectangle leaves the bounds unproven,
                // it does not refuse the read.
                var bounds: ?Bounds = null;
                if (tableRefRaw(table_part.bytes)) |ref_raw| {
                    bounds = parseBounds(try decodeLexical(self.arena, ref_raw));
                }
                const header_rows = table_edit.tableHeaderRowCount(table_part.bytes) orelse return error.MalformedPivotXml;
                const totals_rows = table_edit.tableTotalsRowCount(table_part.bytes) orelse return error.MalformedPivotXml;
                // Only a headerless table's names are its schema; a
                // headered one's are read off its header row.
                const columns: []const []const u8 = if (header_rows == 0) try self.tableColumnNames(table_part.bytes) else &.{};
                try entries.append(self.arena, .{ .folded = folded, .sheet_idx = @intCast(i), .bounds = bounds, .header_rows = header_rows, .totals_rows = totals_rows, .columns = columns, .part_name = try self.arena.dupe(u8, table_part_name) });
            }
        }
        self.tables = try entries.toOwnedSlice(self.arena);
        return self.tables.?;
    }

    /// A table's `<tableColumn>` names in order, decoded — read through
    /// the pivot scanner (one preflighted tree, direct children only,
    /// decoy-aware, the close where the scanner finds it) rather than
    /// a lexical search a comment could feed (Codex #205 r5 REL-503).
    /// Arena-owned.
    fn tableColumnNames(self: *Resolver, src: []const u8) Error![]const []const u8 {
        const root = pivot_xml.scanRoot(src, "table") catch |e| return mapParse(e);
        var kids = pivot_xml.Children.init(src, root.hit, root.body_end, root.prefix, root.env);
        var names: std.ArrayListUnmanaged([]const u8) = .empty;
        var seen = false;
        while (kids.next() catch |e| return mapParse(e)) |k| {
            if (!std.mem.eql(u8, k.local, "tableColumns")) continue;
            if (seen) return error.MalformedPivotXml;
            seen = true;
            var cols = pivot_xml.Children.init(src, k.hit, k.end, root.prefix, k.env);
            while (cols.next() catch |e| return mapParse(e)) |c| {
                if (!std.mem.eql(u8, c.local, "tableColumn")) continue;
                const raw = wbxml.getAttr(c.attrs(src), "name") orelse return error.MalformedPivotXml;
                try names.append(self.arena, try decode(self.arena, .table_column_name, raw));
            }
        }
        if (!seen) return error.MalformedPivotXml;
        return names.toOwnedSlice(self.arena);
    }

    /// The engine's symbol table, or `MalformedPivotXml` when it
    /// refuses: with the name inventory unreadable, no name-based
    /// source can be resolved either way.
    fn ensureSymbols(self: *Resolver) Error!*const formula.SymbolTable {
        if (self.symbols) |*t| return t;
        if (self.symbols_refused) return error.MalformedPivotXml;

        var builder = formula.Builder.init(self.gpa, recalc_run.collation_v1);
        defer builder.deinit();
        for (self.wb.sheets) |s| try builder.addSheet(s.name);
        // The names come from the part, as the evaluator's do: the
        // typed view drops the attribute region that says whether a
        // name is a macro entry point rather than a range.
        switch (try formula.names.scanDefinedNames(self.gpa, self.wb_xml)) {
            .ok => |d| {
                var defined = d;
                defer defined.deinit();
                for (defined.rows) |dn| {
                    try builder.addName(dn.raw_identifier, dn.raw_body, .{
                        .scope = if (dn.local_sheet_id) |id| formula.env.SheetIndex.fromInt(id) else null,
                        .hidden = dn.hidden,
                        .attr_refusal = dn.refusal_when_referenced,
                    });
                }
            },
            .refused => {
                self.symbols_refused = true;
                return error.MalformedPivotXml;
            },
        }
        switch (try builder.finish()) {
            .ok => |t| self.symbols = t,
            .refused => {
                self.symbols_refused = true;
                return error.MalformedPivotXml;
            },
        }
        return &self.symbols.?;
    }

    const Area = struct { sheet_idx: u32, bounds: ?Bounds };

    /// The area a defined name's body denotes, when the body is exactly
    /// one static sheet-qualified area (`Data!$A$1:$C$4`, `Data!$A$1`,
    /// `'My Data'!A1:B2`, `Data!$A:$C`). A 3D span, a dynamic body
    /// (`OFFSET(…)`, `Data!A1:INDEX(…)`), a union or a bare range
    /// resolves to null: none names one area a pivot could read. A
    /// range whose ends are static references of different kinds
    /// (`Data!$A$1:$C:$C`) names the sheet — as it did before S7b-1 —
    /// with no bounds (Codex #202 r1 F1).
    fn areaOfBody(self: *Resolver, body: []const u8) Error!?Area {
        var parsed = try formula.parser.parse(self.gpa, body, .{});
        defer parsed.deinit(self.gpa);
        const ast = switch (parsed) {
            .ok => |t| t,
            .refused => return null,
        };
        // `Data!A1` is one qualified node; `Data!$A$1:$C$4` is a range
        // operator whose LEFT operand carries the qualifier. The right
        // operand must be a static reference too, on the same sheet if
        // it is qualified at all — compared by resolved index, so
        // `Data!A1:data!C4` is one sheet. Parentheses around the whole
        // body keep it a reference (Codex #202 r4 F2).
        var root = ast.root;
        while (true) {
            switch (ast.node(root)) {
                .paren => |pn| root = pn.child,
                else => break,
            }
        }
        switch (ast.node(root)) {
            .qualified => |q| {
                const bounds = staticBounds(ast.node(q.target)) orelse return null;
                const idx = (try self.sheetOfSpec(q.sheet)) orelse return null;
                return .{ .sheet_idx = idx, .bounds = bounds };
            },
            .binary => |b| {
                if (b.op != .range) return null;
                const lhs = switch (ast.node(b.lhs)) {
                    .qualified => |q| q,
                    else => return null,
                };
                const idx = (try self.sheetOfSpec(lhs.sheet)) orelse return null;
                const rhs_node = ast.node(b.rhs);
                const rhs_target: formula.parser.Node = switch (rhs_node) {
                    .qualified => |q| blk: {
                        const r = (try self.sheetOfSpec(q.sheet)) orelse return null;
                        if (r != idx) return null;
                        break :blk ast.node(q.target);
                    },
                    else => rhs_node,
                };
                const lhs_target = ast.node(lhs.target);
                if (staticBounds(lhs_target) == null or staticBounds(rhs_target) == null) return null;
                return .{ .sheet_idx = idx, .bounds = rangeBounds(lhs_target, rhs_target) };
            },
            else => return null,
        }
    }

    /// Every sheet a name body depends on, through the names it
    /// references: sheet qualifiers (a 3D span contributes each sheet
    /// between its ends, in tab order, by the engine's own expansion —
    /// a reversed or dangling span is `#REF!` and contributes nothing),
    /// tables (their host sheet), and the bodies of referenced names —
    /// each name resolved in the scope the reference sits in (a
    /// qualifier's sheet, else the enclosing scope), so a sheet-scoped
    /// name that shadows a workbook one is its own dependency (Codex
    /// #202 r1 F2). The bodies are walked from a worklist, not by
    /// recursion: every (name, scope) pair is walked exactly once, so a
    /// cycle terminates, a chain of any length costs one visit per name,
    /// and no depth cap can cut a body short (Codex #202 r3 F1). A body
    /// the parser refuses and a name the inventory refuses or lacks
    /// contribute nothing — evidence, not refusal. Ascending,
    /// deduplicated, arena-owned.
    const Closure = struct { sheets: []const u32, names: []const NameKey };

    /// The sheets as documented above, plus every name the walk visited
    /// — the root and each name reachable from it, once each whatever
    /// the scopes it was invoked from, keyed as `<definedName>` spells
    /// it (the name's own `localSheetId`, not the invoking scope).
    /// Arena-owned.
    fn closure(self: *Resolver, root: *const formula.Name) Error!Closure {
        var sheets: std.ArrayListUnmanaged(u32) = .empty;
        defer sheets.deinit(self.gpa);
        var walk: Walk = .{};
        defer walk.deinit(self.gpa);
        try walk.enqueue(self.gpa, .{ .name = root, .scope = scopeOf(root) });
        while (walk.pending.pop()) |v| try self.walkBody(v.name.body, v.scope, &sheets, &walk);
        std.mem.sort(u32, sheets.items, {}, std.sort.asc(u32));

        var names: std.ArrayListUnmanaged(NameKey) = .empty;
        defer names.deinit(self.gpa);
        for (walk.visited.items, 0..) |v, i| {
            var dup = false;
            for (walk.visited.items[0..i]) |earlier| dup = dup or earlier.name == v.name;
            if (dup) continue;
            try names.append(self.gpa, .{
                .identifier = try self.arena.dupe(u8, v.name.identifier),
                .scope = scopeOf(v.name),
            });
        }
        return .{
            .sheets = try self.arena.dupe(u32, sheets.items),
            .names = try self.arena.dupe(NameKey, names.items),
        };
    }

    fn scopeOf(n: *const formula.Name) ?u32 {
        return if (n.scope) |sc| sc.toInt() else null;
    }

    /// A name is walked once per scope it is invoked from: a
    /// sheet-scoped name has one scope, its own; a workbook-scoped body
    /// resolves its unqualified names from the sheet that invoked it,
    /// so the same body under two invoking sheets is two walks (Codex
    /// #202 r2 F1).
    const Visit = struct { name: *const formula.Name, scope: ?u32 };

    const Walk = struct {
        visited: std.ArrayListUnmanaged(Visit) = .empty,
        pending: std.ArrayListUnmanaged(Visit) = .empty,

        fn deinit(self: *Walk, gpa: Allocator) void {
            self.visited.deinit(gpa);
            self.pending.deinit(gpa);
        }

        /// Queue a visit not yet made.
        fn enqueue(self: *Walk, gpa: Allocator, v: Visit) Error!void {
            for (self.visited.items) |seen| if (seen.name == v.name and seen.scope == v.scope) return;
            try self.visited.append(gpa, v);
            try self.pending.append(gpa, v);
        }
    };

    fn walkBody(
        self: *Resolver,
        body: []const u8,
        scope: ?u32,
        sheets: *std.ArrayListUnmanaged(u32),
        walk: *Walk,
    ) Error!void {
        var parsed = try formula.parser.parse(self.gpa, body, .{});
        defer parsed.deinit(self.gpa);
        const ast = switch (parsed) {
            .ok => |t| t,
            // A body the parser refuses — an external workbook reference
            // beside a local one, say — still names what it names: the
            // sheets, names and tables are read off the text (Codex #202
            // r4 F1, r5 F1).
            .refused => return self.scanBodyText(body, scope, sheets, walk),
        };
        try self.walkNode(ast, ast.root, scope, sheets, walk);
    }

    /// The refusal-tolerant fallback: every sheet of this workbook whose
    /// name, bare or quoted (`'` doubled inside), is followed by `!` in
    /// the body text; every defined name and every table whose
    /// identifier appears in it as a whole word — the names queued for
    /// their own walk, the tables for their host. Evidence only, so a
    /// match inside a string literal over-marks rather than under-marks.
    fn scanBodyText(self: *Resolver, body: []const u8, scope: ?u32, sheets: *std.ArrayListUnmanaged(u32), walk: *Walk) Error!void {
        for (self.sheet_names, 0..) |name, idx| {
            if (name.len == 0) continue;
            if (try self.textNamesSheet(body, name)) try addSheet(self.gpa, sheets, @intCast(idx));
        }
        const folded_body = (try fold(self.arena, body)) orelse return;
        // 3D spans, `First:Last!` bare or `'First:Last'!` quoted, in tab
        // order — every member between the ends (Codex #202 r6 F1).
        const folds = try self.ensureSheetFolds();
        for (folds, 0..) |first, i| {
            for (folds[i..], i..) |last, j| {
                const bare = try std.mem.concat(self.arena, u8, &.{ first, ":", last, "!" });
                const quoted = try std.mem.concat(self.arena, u8, &.{ "'", first, ":", last, "'!" });
                if (std.mem.indexOf(u8, folded_body, bare) != null or std.mem.indexOf(u8, folded_body, quoted) != null) {
                    var k: u32 = @intCast(i);
                    while (k <= j) : (k += 1) try addSheet(self.gpa, sheets, k);
                }
            }
        }
        if (self.ensureSymbols()) |symbols| {
            for (symbols.names) |*n| {
                if (textHasWord(folded_body, n.folded)) try walk.enqueue(self.gpa, .{ .name = n, .scope = scopeOf(n) orelse scope });
            }
        } else |e| if (e == error.OutOfMemory) return error.OutOfMemory;
        if (self.ensureTables()) |tables| {
            for (tables) |t| if (textHasWord(folded_body, t.folded)) try addSheet(self.gpa, sheets, t.sheet_idx);
        } else |e| if (e == error.OutOfMemory) return error.OutOfMemory;
    }

    /// `word` in `text` with no identifier byte on either side.
    fn textHasWord(text: []const u8, word: []const u8) bool {
        if (word.len == 0) return false;
        var from: usize = 0;
        while (std.mem.indexOfPos(u8, text, from, word)) |at| : (from = at + 1) {
            const end = at + word.len;
            const before_ok = at == 0 or !isIdentByte(text[at - 1]);
            const after_ok = end == text.len or !isIdentByte(text[end]);
            if (before_ok and after_ok) return true;
        }
        return false;
    }

    fn isIdentByte(c: u8) bool {
        return (c >= 'a' and c <= 'z') or (c >= 'A' and c <= 'Z') or (c >= '0' and c <= '9') or c == '_' or c == '.' or c >= 0x80;
    }

    fn textNamesSheet(self: *Resolver, body: []const u8, name: []const u8) Error!bool {
        // Bare: `Name!`, matched case-insensitively by the sheet fold.
        const folded_body = (try fold(self.arena, body)) orelse return false;
        const folded_name = (try fold(self.arena, name)) orelse return false;
        var from: usize = 0;
        while (std.mem.indexOfPos(u8, folded_body, from, folded_name)) |at| : (from = at + 1) {
            const end = at + folded_name.len;
            if (end < folded_body.len and folded_body[end] == '!') return true;
            if (end + 1 < folded_body.len and folded_body[end] == '\'' and folded_body[end + 1] == '!' and at > 0 and folded_body[at - 1] == '\'') return true;
        }
        // Quoted with an apostrophe inside: `'It''s'!`.
        if (std.mem.indexOfScalar(u8, name, '\'') != null) {
            const doubled = try std.mem.replaceOwned(u8, self.arena, folded_name, "'", "''");
            const needle = try std.mem.concat(self.arena, u8, &.{ "'", doubled, "'!" });
            return std.mem.indexOf(u8, folded_body, needle) != null;
        }
        return false;
    }

    fn walkNode(
        self: *Resolver,
        ast: formula.parser.Ast,
        i: formula.parser.Index,
        scope: ?u32,
        sheets: *std.ArrayListUnmanaged(u32),
        walk: *Walk,
    ) Error!void {
        switch (ast.node(i)) {
            .qualified => |q| {
                const members = try self.addSpecSheets(q.sheet, sheets);
                // `Report!N` is the name `N` looked up from `Report`: the
                // qualifier is the scope of the lookup, and it counts as
                // a dependency like any other qualifier. A span qualifier
                // (`'Data:Report'!N`) looks `N` up from every member
                // (Codex #202 r2 F2); a qualifier that names nothing is
                // `#REF!`, and the name behind it is not looked up from
                // anywhere else (Codex #202 r3 F2).
                switch (ast.node(q.target)) {
                    .name => |n| if (members) |m| {
                        var k = m.first;
                        while (k <= m.last) : (k += 1) try self.enqueueName(n.raw, k, sheets, walk);
                    },
                    else => try self.walkNode(ast, q.target, scope, sheets, walk),
                }
            },
            // The raw spelling: a value-position name resolves as written,
            // so `_xlfn.Anchor` is not `Anchor` (Codex #202 r1 F5).
            .name => |n| try self.enqueueName(n.raw, scope, sheets, walk),
            .structured => |st| if (st.table) |t| try self.addTableSheet(t, sheets),
            // The callee is a function, not a name — except that
            // `INDIRECT("Report!$A$1")` reads what its literal spells: the
            // evaluator resolves the text, so the walk does too, by
            // walking the literal as a body (Codex #202 r5 F2).
            .call => |c| {
                const args = ast.children(c.args);
                if (args.len > 0 and isIndirect(ast, c.callee)) {
                    // Through any parentheses around the literal (Codex
                    // #202 r6 F2).
                    var arg = args[0];
                    while (true) {
                        switch (ast.node(arg)) {
                            .paren => |pn| arg = pn.child,
                            else => break,
                        }
                    }
                    switch (ast.node(arg)) {
                        .string => |lit| {
                            const text = try unquoteStringLiteral(self.arena, lit.text);
                            try self.walkBody(text, scope, sheets, walk);
                        },
                        else => {},
                    }
                }
                for (args) |k| try self.walkNode(ast, k, scope, sheets, walk);
            },
            .array => |a| for (ast.children(a.elems)) |k| try self.walkNode(ast, k, scope, sheets, walk),
            .paren => |pn| try self.walkNode(ast, pn.child, scope, sheets, walk),
            .unary => |u| try self.walkNode(ast, u.child, scope, sheets, walk),
            .postfix => |pf| try self.walkNode(ast, pf.child, scope, sheets, walk),
            .binary => |b| {
                try self.walkNode(ast, b.lhs, scope, sheets, walk);
                try self.walkNode(ast, b.rhs, scope, sheets, walk);
            },
            .number, .string, .boolean, .error_lit, .missing_arg, .ref_cell, .ref_full_col, .ref_full_row => {},
        }
    }

    fn isIndirect(ast: formula.parser.Ast, callee: formula.parser.Index) bool {
        return switch (ast.node(callee)) {
            .name => |n| std.ascii.eqlIgnoreCase(n.bare, "INDIRECT"),
            else => false,
        };
    }

    /// `"Report!$A$1"` → `Report!$A$1`, a doubled `""` → `"`.
    fn unquoteStringLiteral(arena: Allocator, lit: []const u8) Error![]const u8 {
        if (lit.len < 2 or lit[0] != '"' or lit[lit.len - 1] != '"') return lit;
        const inner = lit[1 .. lit.len - 1];
        if (std.mem.indexOf(u8, inner, "\"\"") == null) return inner;
        return std.mem.replaceOwned(u8, arena, inner, "\"\"", "\"");
    }

    /// Resolve a name from `scope` and queue its body — or, for a table
    /// spelled bare, take the host sheet now.
    fn enqueueName(
        self: *Resolver,
        raw: []const u8,
        scope: ?u32,
        sheets: *std.ArrayListUnmanaged(u32),
        walk: *Walk,
    ) Error!void {
        const symbols = self.ensureSymbols() catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            else => return,
        };
        const from: ?formula.SheetIndex = if (scope) |sc| formula.SheetIndex.fromInt(sc) else null;
        switch (try symbols.resolveName(self.gpa, from, raw)) {
            .name => |n| try walk.enqueue(self.gpa, .{ .name = n, .scope = scopeOf(n) orelse scope }),
            .table, .not_found => {
                // A table spelled bare, or nothing: the table index has
                // the host if there is one.
                try self.addTableSheet(raw, sheets);
            },
            // A name the engine refuses to reference evaluates to nothing
            // — it does not fall through to a table of the same spelling
            // (Codex #202 r2 F5).
            .refused => {},
        }
    }

    const Members = struct { first: u32, last: u32 };

    /// The sheets a qualifier names — one, or every member of a 3D span
    /// (quoted `'Data:Report'!` and unquoted `Data:Report!` alike, by
    /// the engine's own split and expansion; a reversed or dangling span
    /// is `#REF!` and names nothing — Codex #202 r1 F3). Returns the
    /// members, for a name looked up through the qualifier.
    fn addSpecSheets(self: *Resolver, spec: formula.parser.SheetSpec, sheets: *std.ArrayListUnmanaged(u32)) Error!?Members {
        const unquoted = try self.unquoteSpec(spec);
        if (!formula.names.isSpan(spec)) {
            const idx = (try self.sheetIndexOf(unquoted)) orelse return null;
            try addSheet(self.gpa, sheets, idx);
            return .{ .first = idx, .last = idx };
        }
        const ends = formula.names.splitSpan(spec, unquoted) orelse return null;
        const symbols = self.ensureSymbols() catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            else => return null,
        };
        switch (try symbols.resolveSheetSpan(self.gpa, ends.first, ends.last)) {
            .members => |m| {
                var k = m.first;
                while (k <= m.last) : (k += 1) try addSheet(self.gpa, sheets, k);
                return .{ .first = m.first, .last = m.last };
            },
            .ref_error => return null,
        }
    }

    fn addTableSheet(self: *Resolver, raw_table: []const u8, sheets: *std.ArrayListUnmanaged(u32)) Error!void {
        const folded = (try fold(self.arena, raw_table)) orelse return;
        const tables = self.ensureTables() catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            else => return,
        };
        for (tables) |t| {
            if (std.mem.eql(u8, t.folded, folded)) return addSheet(self.gpa, sheets, t.sheet_idx);
        }
    }

    fn addSheet(gpa: Allocator, sheets: *std.ArrayListUnmanaged(u32), idx: u32) Error!void {
        if (std.mem.indexOfScalar(u32, sheets.items, idx) != null) return;
        try sheets.append(gpa, idx);
    }

    fn sheetOfSpec(self: *Resolver, spec: formula.parser.SheetSpec) Error!?u32 {
        if (spec.last != null) return null;
        const name = try self.unquoteSpec(spec);
        return self.sheetIndexOf(name);
    }

    /// `unquoteSheetSpec` sized by the token, not a fixed buffer: a
    /// sheet name the inventories accepted is one the walk must be able
    /// to look up, whatever its length (Codex #202 r4 F3).
    fn unquoteSpec(self: *Resolver, spec: formula.parser.SheetSpec) Error![]const u8 {
        if (!spec.quoted) return spec.first;
        const buf = try self.arena.alloc(u8, spec.first.len);
        return unquoteSheetSpec(buf, spec) orelse spec.first;
    }
};

/// A cell, a whole-column span or a whole-row span — the reference
/// nodes that denote a fixed area — as bounds.
fn staticBounds(n: formula.parser.Node) ?Bounds {
    return switch (n) {
        .ref_cell => |c| .{ .rect = .{
            .tl_col = c.cell.col.oneBased(),
            .tl_row = c.cell.row.oneBased(),
            .br_col = c.cell.col.oneBased(),
            .br_row = c.cell.row.oneBased(),
        } },
        .ref_full_col => |fc| .{ .whole_columns = .{
            .first_col = @min(fc.first.col.oneBased(), fc.last.col.oneBased()),
            .last_col = @max(fc.first.col.oneBased(), fc.last.col.oneBased()),
        } },
        .ref_full_row => |fr| .{ .whole_rows = .{
            .first_row = @min(fr.first.row.oneBased(), fr.last.row.oneBased()),
            .last_row = @max(fr.first.row.oneBased(), fr.last.row.oneBased()),
        } },
        else => null,
    };
}

/// `lhs:rhs` over two static references of one kind — two cells make a
/// rectangle (corners normalised), two column spans or two row spans
/// merge. Mixed kinds denote an area Excel accepts but this reader
/// does not bound.
fn rangeBounds(lhs: formula.parser.Node, rhs: formula.parser.Node) ?Bounds {
    const a = staticBounds(lhs) orelse return null;
    const b = staticBounds(rhs) orelse return null;
    if (std.meta.activeTag(a) != std.meta.activeTag(b)) return null;
    return switch (a) {
        .rect => |ra| .{ .rect = .{
            .tl_col = @min(ra.tl_col, b.rect.tl_col),
            .tl_row = @min(ra.tl_row, b.rect.tl_row),
            .br_col = @max(ra.br_col, b.rect.br_col),
            .br_row = @max(ra.br_row, b.rect.br_row),
        } },
        .whole_columns => |ca| .{ .whole_columns = .{
            .first_col = @min(ca.first_col, b.whole_columns.first_col),
            .last_col = @max(ca.last_col, b.whole_columns.last_col),
        } },
        .whole_rows => |ra| .{ .whole_rows = .{
            .first_row = @min(ra.first_row, b.whole_rows.first_row),
            .last_row = @max(ra.last_row, b.whole_rows.last_row),
        } },
    };
}

/// A decoded `ref` as bounds: `A1:C4` / `A1` (corners in either order —
/// the bounds are the normalised rectangle; the S7a splice parser stays
/// strict), `A:C` (letters on both sides), `1:4` (digits on both sides).
/// Plain `ST_Ref` spellings only: uppercase, no `$`, no leading zero.
/// Null on anything else — bounds are evidence, and an unparseable
/// spelling is none.
pub fn parseBounds(ref: []const u8) ?Bounds {
    const cell_opts: coords.CellParseOptions = .{ .case = .upper_only };
    const colon = std.mem.indexOfScalar(u8, ref, ':') orelse {
        const c = coords.parseCell(ref, cell_opts) catch return null;
        return .{ .rect = .{ .tl_col = c.col.oneBased(), .tl_row = c.row.oneBased(), .br_col = c.col.oneBased(), .br_row = c.row.oneBased() } };
    };
    const lhs = ref[0..colon];
    const rhs = ref[colon + 1 ..];
    if (lhs.len == 0 or rhs.len == 0) return null;
    if (coords.parseCell(lhs, cell_opts)) |a| {
        const b = coords.parseCell(rhs, cell_opts) catch return null;
        return .{ .rect = .{
            .tl_col = @min(a.col.oneBased(), b.col.oneBased()),
            .tl_row = @min(a.row.oneBased(), b.row.oneBased()),
            .br_col = @max(a.col.oneBased(), b.col.oneBased()),
            .br_row = @max(a.row.oneBased(), b.row.oneBased()),
        } };
    } else |_| {}
    if (allLetters(lhs) and allLetters(rhs)) {
        // Uppercase only and inside the grid — the spelling `ST_Ref` and
        // the S7a rectangle parser accept.
        const a = coords.parseColNumber(lhs, .{ .case = .upper_only }) catch return null;
        const b = coords.parseColNumber(rhs, .{ .case = .upper_only }) catch return null;
        return .{ .whole_columns = .{ .first_col = @min(a, b), .last_col = @max(a, b) } };
    }
    if (allDigits(lhs) and allDigits(rhs)) {
        if (lhs[0] == '0' or rhs[0] == '0') return null;
        const a = std.fmt.parseInt(u32, lhs, 10) catch return null;
        const b = std.fmt.parseInt(u32, rhs, 10) catch return null;
        if (a == 0 or b == 0 or a > zlsx.max_row or b > zlsx.max_row) return null;
        return .{ .whole_rows = .{ .first_row = @min(a, b), .last_row = @max(a, b) } };
    }
    return null;
}

fn allLetters(s: []const u8) bool {
    for (s) |ch| if (ch < 'A' or ch > 'Z') return false;
    return true;
}

fn allDigits(s: []const u8) bool {
    for (s) |ch| if (ch < '0' or ch > '9') return false;
    return true;
}

/// The table part's `ref`, raw — the first real `<table` open tag's
/// attribute, read by the scanner the table editor reads it with.
fn tableRefRaw(src: []const u8) ?[]const u8 {
    const hit = (wbxml.findTagOpen(src, 0, "table") catch return null) orelse return null;
    return wbxml.getAttr(src[hit.attrs_start..hit.attrs_end], "ref");
}

/// The engine's shipped fold, arena-owned. Null when the fold refuses
/// the text (malformed UTF-8), which no decoded name is.
fn fold(a: Allocator, s: []const u8) Error!?[]const u8 {
    return recalc_run.collation_v1.fold(a, s) catch |e| {
        if (e == error.OutOfMemory) return error.OutOfMemory;
        return null;
    };
}

/// `'It''s'` → `It's`; an unquoted spelling is itself. Null when the
/// unescaped name does not fit `buf` — Excel caps sheet names at 31
/// characters, so it always does for a name Excel wrote.
fn unquoteSheetSpec(buf: []u8, spec: formula.parser.SheetSpec) ?[]const u8 {
    if (!spec.quoted) return spec.first;
    const raw = spec.first;
    if (raw.len < 2 or raw[0] != '\'' or raw[raw.len - 1] != '\'') return null;
    const inner = raw[1 .. raw.len - 1];
    var n: usize = 0;
    var k: usize = 0;
    while (k < inner.len) : (k += 1) {
        if (n >= buf.len) return null;
        buf[n] = inner[k];
        n += 1;
        if (inner[k] == '\'') k += 1;
    }
    return buf[0..n];
}

// ─── S7a: the output-location lift ───────────────────────────────────

/// S7a (`goal_sigmoid.md`): shift `pivotTableDefinition/location@ref`
/// in step with a row / col edit on the pivot's HOST sheet.
///
/// The output rectangle is the one absolute coordinate a pivot-table
/// definition carries: `firstHeaderRow` / `firstDataRow` /
/// `firstDataCol` are offsets inside it, formats and selections are
/// field-addressed (`pivotArea`), and the cache's `worksheetSource`
/// addresses the SOURCE sheet (S7b). So a whole-row / whole-column edit
/// that leaves the rectangle intact — an insert at or above its top, a
/// delete above it — moves the rectangle and nothing else; an edit that
/// lands inside it would change the pivot's own layout, which Excel
/// itself refuses ("We can't make this change for the selected cells
/// because it will affect a PivotTable"), so it refuses here too.
///
/// **The footprint is wider than `ref` when the pivot has report
/// filters.** Excel renders page fields ABOVE the rectangle —
/// `rowPageCount` rows of label + value pairs, a blank separator row,
/// then the body at `ref`'s top — and, for the over-then-down layout,
/// across `colPageCount` pairs with a blank column between. Those cells
/// are the pivot's too, but no attribute names them, so the lift treats
/// a conservative superset as inside: `rowPageCount + 1` rows above the
/// top, `3 · colPageCount` columns from the left edge (against Excel's
/// `3 · colPageCount − 1`). The oracle question the row parks — how
/// Excel re-lays a pivot whose rectangle moved — is what would let a
/// later row narrow it; until then over-refusing costs an edit,
/// under-refusing costs a workbook.
///
/// The splice lands exactly where `parseTableDefinition` read the
/// attribute (`Location.ref_span`): the parser is decoy-aware and reads
/// one tree, so there is no second scanner to disagree with it. Every
/// other byte of the part is preserved; a no-op edit (the rectangle is
/// below / right of the edit) returns the input unchanged, entities and
/// all, so the caller's byte compare skips the part.
pub const edit = struct {
    pub const Axis = enum { row, col };
    pub const Kind = enum { insert, delete };

    pub const EditError = error{
        /// The edit lands inside the pivot's footprint — its output
        /// rectangle, or the report-filter band above it.
        PivotLocationEditUnsafe,
        /// S7b: the edit lands on a source in a way no range semantics
        /// admit — a delete of its header row (the field names) or of
        /// its only row, a column edit inside it (the cache's field
        /// schema, S7c's), an unplaceable `r:id` or a `sheet`-only
        /// spelling naming the edited sheet, a `ref` no rectangle
        /// parser accepts (`docs/plans/s7b-cache-policy.md` §2.2, §7 Q4).
        PivotSourceEditUnsafe,
        /// The shift would push the rectangle past `XFD` / `1048576`.
        PivotCoordinateOverflow,
        /// S7b-4: the edit changed a source's content and the cache
        /// cannot be rebuilt from it — a shape the engine's first
        /// slice does not evaluate (`engine.RebuildError` names them).
        PivotShapeUnsupported,
        /// The part is not one readable `pivotTableDefinition`, or its
        /// `location@ref` is not an A1 rectangle.
        MalformedPivotXml,
        OutOfMemory,
    };

    /// 1-based, inclusive, top-left ≤ bottom-right.
    pub const Rect = struct {
        tl_col: u32,
        tl_row: u32,
        br_col: u32,
        br_row: u32,

        pub fn eql(a: Rect, b: Rect) bool {
            return a.tl_col == b.tl_col and a.tl_row == b.tl_row and
                a.br_col == b.br_col and a.br_row == b.br_row;
        }
    };

    /// The rows and columns the pivot occupies on its host: the output
    /// rectangle plus the report-filter band (see the namespace note).
    /// `first_row ≤ rect.tl_row`, `last_col ≥ rect.br_col`; the left
    /// edge and the bottom are the rectangle's own.
    pub const Footprint = struct {
        rect: Rect,
        first_row: u32,
        last_col: u32,
    };

    /// Rewrite one `pivotTableN.xml` part for a row / col edit on its
    /// host sheet. `idx_1based` is the inserted row / column position or
    /// the deleted one, in the sheet's pre-edit coordinates. Returns a
    /// fresh buffer the caller owns — byte-equal to `src` when the edit
    /// does not move the rectangle.
    pub fn applyToTableDefinition(
        allocator: Allocator,
        src: []const u8,
        axis: Axis,
        idx_1based: u32,
        kind: Kind,
    ) EditError![]u8 {
        if (idx_1based == 0) return error.MalformedPivotXml;
        var def = pivot_xml.parseTableDefinition(allocator, src) catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            error.MalformedXml => return error.MalformedPivotXml,
        };
        defer def.deinit(allocator);

        const fp = try footprintOf(allocator, def);
        const shifted = try shiftRect(fp, axis, idx_1based, kind);
        if (shifted.eql(fp.rect)) return allocator.dupe(u8, src);

        var ref_buf: [Bounds.format_buf_len]u8 = undefined;
        const new_ref = formatRect(&ref_buf, shifted) catch return error.PivotCoordinateOverflow;
        const span = def.location.ref_span;
        assert(span.start <= span.end and span.end <= src.len);
        return std.mem.concat(allocator, u8, &.{ src[0..span.start], new_ref, src[span.end..] });
    }

    /// The footprint a part's bytes declare — a parse plus
    /// `footprintOf`, for a caller holding only the bytes (the
    /// sweep's pre-schema host clear, S7c-2). The parse lives in a
    /// local arena and the result is all values, so any allocator
    /// serves (in-house review S7C2-A5).
    pub fn footprintOfBytes(allocator: Allocator, table_xml: []const u8) EditError!Footprint {
        var arena_state = std.heap.ArenaAllocator.init(allocator);
        defer arena_state.deinit();
        const arena = arena_state.allocator();
        const def = pivot_xml.parseTableDefinition(arena, table_xml) catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            error.MalformedXml => return error.MalformedPivotXml,
        };
        return footprintOf(arena, def);
    }

    /// The footprint of a parsed definition. Refuses (as malformed) a
    /// `ref` that is not an A1 rectangle — ST_Ref has no `$`, no sheet
    /// qualifier, no whitespace.
    pub fn footprintOf(allocator: Allocator, def: pivot_xml.TableDefinition) EditError!Footprint {
        const decoded = formula.decode.decodeCarrier(allocator, .lexical, def.location.ref) catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            else => return error.MalformedPivotXml,
        };
        defer allocator.free(decoded);
        const rect = parseRect(decoded) orelse return error.MalformedPivotXml;

        // Report filters: `rowPageCount` / `colPageCount` when Excel
        // wrote them, the page-field count as a superset for BOTH when
        // a producer left them out — an over-then-down layout puts every
        // field on one row, so the count bounds the columns as much as
        // the rows (Codex #200 r3 REL-041). `max(1, …)` because a band
        // with rows has at least one pair per row.
        const pages: u32 = @intCast(@min(def.page_fields.len, std.math.maxInt(u32)));
        const page_rows: u32 = @max(def.location.row_page_count orelse 0, pages);
        var first_row = rect.tl_row;
        var last_col = rect.br_col;
        if (page_rows > 0) {
            // Subtraction only: `rowPageCount` is any u32 the part
            // declares, and `page_rows + 1` would wrap at the maximum
            // (Codex #200 r1 REL-037). `rect.tl_row ≥ 1`.
            first_row = if (page_rows >= rect.tl_row - 1) 1 else rect.tl_row - page_rows - 1;
            const page_cols: u32 = @max(@max(def.location.col_page_count orelse 0, pages), 1);
            const band_width = std.math.mul(u32, page_cols, 3) catch return error.PivotCoordinateOverflow;
            const band_right = std.math.add(u32, rect.tl_col, band_width - 1) catch return error.PivotCoordinateOverflow;
            last_col = @max(last_col, band_right);
        }
        return .{ .rect = rect, .first_row = first_row, .last_col = last_col };
    }

    /// Interval semantics on the footprint. An insert AT the footprint's
    /// first row / column pushes the whole pivot (Excel's "insert above"),
    /// so the shift zone is `idx ≤ first` for inserts and `idx < first`
    /// for deletes; `first < idx ≤ last` (delete: `first ≤ idx ≤ last`)
    /// is inside and refuses; beyond `last` nothing moves.
    pub fn shiftRect(fp: Footprint, axis: Axis, idx_1based: u32, kind: Kind) EditError!Rect {
        // Public: a caller reaching this without `applyToTableDefinition`
        // gets the same answer for a position that does not exist,
        // never a rectangle at row or column 0 (Codex #200 r4 REL-044).
        if (idx_1based == 0) return error.MalformedPivotXml;
        var r = fp.rect;
        switch (axis) {
            .row => switch (kind) {
                .insert => {
                    if (idx_1based <= fp.first_row) {
                        if (r.br_row >= zlsx.max_row) return error.PivotCoordinateOverflow;
                        r.tl_row += 1;
                        r.br_row += 1;
                    } else if (idx_1based <= r.br_row) {
                        return error.PivotLocationEditUnsafe;
                    }
                },
                .delete => {
                    if (idx_1based < fp.first_row) {
                        r.tl_row -= 1;
                        r.br_row -= 1;
                    } else if (idx_1based <= r.br_row) {
                        return error.PivotLocationEditUnsafe;
                    }
                },
            },
            .col => switch (kind) {
                .insert => {
                    if (idx_1based <= r.tl_col) {
                        if (r.br_col >= zlsx.max_col_1based) return error.PivotCoordinateOverflow;
                        r.tl_col += 1;
                        r.br_col += 1;
                    } else if (idx_1based <= fp.last_col) {
                        return error.PivotLocationEditUnsafe;
                    }
                },
                .delete => {
                    if (idx_1based < r.tl_col) {
                        r.tl_col -= 1;
                        r.br_col -= 1;
                    } else if (idx_1based <= fp.last_col) {
                        return error.PivotLocationEditUnsafe;
                    }
                },
            },
        }
        return r;
    }

    // ─── S7b: the source range ─────────────────────────────────────
    //
    // A pivot's source is a range reference, and a row edit on its
    // sheet treats it as one (`docs/plans/s7b-cache-policy.md` §2.2):
    // an edit above shifts it, an insert inside grows it, a delete
    // inside shrinks it — where the host rectangle above refuses. What
    // refuses is what Excel's own refresh would fail on: deleting the
    // header row (the field names come from it) or the only row. A
    // column edit inside the range is the cache's field schema, which
    // is S7c's row — refused here; outside, the range shifts.

    /// The range semantics on one rectangle. `idx_1based` in the
    /// sheet's pre-edit coordinates.
    pub fn shiftSourceRect(r: Rect, axis: Axis, idx_1based: u32, kind: Kind) EditError!Rect {
        if (idx_1based == 0) return error.MalformedPivotXml;
        var out = r;
        switch (axis) {
            .row => switch (kind) {
                .insert => {
                    // At or above the top row: a pure shift. Inside:
                    // the bottom edge alone moves (one blank record).
                    if (idx_1based <= r.br_row) {
                        if (r.br_row >= zlsx.max_row) return error.PivotCoordinateOverflow;
                        out.br_row += 1;
                        if (idx_1based <= r.tl_row) out.tl_row += 1;
                    }
                },
                .delete => {
                    // The top row is the header; a range of one row
                    // collapses. Both refuse (the table rewriter's
                    // `TableHeaderRowDeleteUnsafe` / `TableCollapseUnsafe`).
                    if (idx_1based == r.tl_row) return error.PivotSourceEditUnsafe;
                    if (idx_1based < r.tl_row) {
                        out.tl_row -= 1;
                        out.br_row -= 1;
                    } else if (idx_1based <= r.br_row) {
                        out.br_row -= 1;
                    }
                },
            },
            .col => switch (kind) {
                .insert => {
                    if (idx_1based <= r.tl_col) {
                        if (r.br_col >= zlsx.max_col_1based) return error.PivotCoordinateOverflow;
                        out.tl_col += 1;
                        out.br_col += 1;
                    } else if (idx_1based <= r.br_col) {
                        return error.PivotSourceEditUnsafe;
                    }
                },
                .delete => {
                    if (idx_1based < r.tl_col) {
                        out.tl_col -= 1;
                        out.br_col -= 1;
                    } else if (idx_1based <= r.br_col) {
                        return error.PivotSourceEditUnsafe;
                    }
                },
            },
        }
        return out;
    }

    /// The same semantics on any bounds: the moved bounds, or null when
    /// the spelling is unchanged. A whole-column source moves on the
    /// column axis as a rectangle's columns do (shift left of it, a
    /// column inside is the schema) and never on the row axis — but
    /// row 1 is its header (an insert there blanks it, a delete
    /// promotes data to field names) and refuses. A whole-row source
    /// moves on the row axis as a rectangle's rows do, and every
    /// column is inside it (Codex #203 r2 REL-201).
    pub fn shiftSourceBounds(b: Bounds, axis: Axis, idx_1based: u32, kind: Kind) EditError!?Bounds {
        if (idx_1based == 0) return error.MalformedPivotXml;
        const moved: Bounds = switch (b) {
            .rect => |r| .{ .rect = try shiftSourceRect(r, axis, idx_1based, kind) },
            .whole_columns => |c| blk: {
                if (axis == .row) {
                    if (idx_1based == 1) return error.PivotSourceEditUnsafe;
                    return null;
                }
                const r = try shiftSourceRect(.{ .tl_col = c.first_col, .tl_row = 1, .br_col = c.last_col, .br_row = 1 }, axis, idx_1based, kind);
                break :blk .{ .whole_columns = .{ .first_col = r.tl_col, .last_col = r.br_col } };
            },
            .whole_rows => |r| blk: {
                if (axis == .col) return error.PivotSourceEditUnsafe;
                const m = try shiftSourceRect(.{ .tl_col = 1, .tl_row = r.first_row, .br_col = 1, .br_row = r.last_row }, axis, idx_1based, kind);
                break :blk .{ .whole_rows = .{ .first_row = m.tl_row, .last_row = m.br_row } };
            },
        };
        return if (moved.eql(b)) null else moved;
    }

    // ─── S7b-3: the refresh marker ─────────────────────────────────
    //
    // The cache is a snapshot of the source range at `refreshedDate`.
    // An edit that changes what the range HOLDS — not merely where it
    // sits — leaves that snapshot describing rows the range no longer
    // has, and zlsx is headless: nobody clicks *Refresh* after a
    // scripted edit. So a cache whose source content an edit may have
    // changed is marked to refresh at open with Excel's own option,
    // `refreshOnLoad="1"` (`docs/plans/s7b-cache-policy.md` §5, A1),
    // under ONE predicate wherever it appears — *the edit may have
    // changed the source's content, and is not a proven pure shift*:
    //
    //   · a source with a finite rectangle marks when the edited row
    //     is inside it (an insert adds a blank record, a delete drops
    //     one); an edit above it is a proven shift, below it a no-op,
    //     and neither marks — the part stays byte-faithful to what
    //     Excel writes after the same edit;
    //   · a whole-column source marks on every admitted row edit (row 1
    //     refuses) and on no column edit — inside refuses (S7c),
    //     outside is a shift; `Data!$A:$C` is byte-identical under
    //     every row edit, its content is not;
    //   · an unbounded name body marks on any row or column edit of a
    //     sheet its closure references, because no shift can be
    //     proven for it;
    //   · a cell write (`setCell`, an appended row) marks under the
    //     same predicate at save time (§7 Q3): inside a rectangle, in
    //     a whole-column source's columns or a whole-row source's
    //     rows, anywhere on the sheet a `sheet`-only spelling claims,
    //     anywhere on a referenced sheet of an unbounded source —
    //     `cellWriteChangesSource`, applied by `Workbook`'s save.
    //
    // A2 — `invalid="1"`, the spec's "needs refreshing" state flag —
    // is the same write to another attribute; whether Excel acts on
    // it at open is oracle-pending (§6, oracle 3), so A1 ships and
    // `marker_attr` is the one place the answer lands. Best-effort by
    // the doc's own terms: inert under `enableRefresh="0"` (left as
    // the user set it) and under a programmatic open; refreshes every
    // consumer of a shared cache.

    /// The root attribute the marker sets — A1 (`refreshOnLoad`) until
    /// the `invalid` oracle says A2.
    pub const marker_attr = "refreshOnLoad";
    /// The attribute as inserted on a root that lacks it.
    pub const marker_insert = " " ++ marker_attr ++ "=\"1\"";

    /// Does the definition already carry the marker — nothing to write?
    /// The one read that pairs with `marker_attr`.
    pub fn markerSet(def: *const pivot_xml.CacheDefinition) bool {
        return def.refresh_on_load;
    }

    /// Does a row / column edit change what one bounds HOLDS, rather
    /// than where it is? Decided for an edit the range semantics
    /// admit — a refusal never reaches the marker — so a delete at the
    /// top row counts (the headerless-table case `table_edit` admits;
    /// every other carrier refused it before asking).
    pub fn editChangesContent(b: Bounds, axis: Axis, idx_1based: u32, kind: Kind) bool {
        return switch (b) {
            .rect => |r| switch (axis) {
                .row => rowEditInside(r.tl_row, r.br_row, idx_1based, kind),
                // Inside is the field schema and refuses (S7c);
                // outside is a shift.
                .col => false,
            },
            // Every row of the sheet is inside whole columns (row 1,
            // the header, refuses); a column inside refuses, outside
            // shifts.
            .whole_columns => axis == .row,
            .whole_rows => |r| switch (axis) {
                .row => rowEditInside(r.first_row, r.last_row, idx_1based, kind),
                // Every column is inside: refused before this.
                .col => false,
            },
        };
    }

    /// §2.2's two *content changed* rows: an insert strictly inside
    /// (`r1 < i ≤ r2`; at `r1` it is a shift), a delete anywhere in
    /// the span (`r1 ≤ i ≤ r2`; at `r1` it is the admitted headerless
    /// case).
    fn rowEditInside(r1: u32, r2: u32, idx_1based: u32, kind: Kind) bool {
        return switch (kind) {
            .insert => idx_1based > r1 and idx_1based <= r2,
            .delete => idx_1based >= r1 and idx_1based <= r2,
        };
    }

    /// The predicate for a cell write at (`row`, `col`), 1-based, on
    /// sheet `sheet_idx`: inside a source's finite rectangle, or
    /// anywhere on a sheet an unbounded name body's closure references.
    /// A resolved spelling that bounds nothing (`sheet` alone) claims
    /// the whole sheet. A spelling that proves no local range — external,
    /// dangling, an `r:id` the reader could not place — never marks: a
    /// mark would ask Excel to refresh at open a source it may not
    /// have, where leaving the snapshot is the state every workbook is
    /// in after a cell edit (§3).
    pub fn cellWriteChangesSource(res: SourceResolution, sheet_idx: u32, row: u32, col: u32) bool {
        switch (res) {
            .external, .none => return false,
            .unresolved => |u| return u.why == .unbounded_body and std.mem.indexOfScalar(u32, u.sheets, sheet_idx) != null,
            .sheet => |s| {
                if (s.sheet_idx != sheet_idx) return false;
                const b = s.bounds orelse return true;
                return switch (b) {
                    .rect => |r| row >= r.tl_row and row <= r.br_row and col >= r.tl_col and col <= r.br_col,
                    .whole_columns => |c| col >= c.first_col and col <= c.last_col,
                    .whole_rows => |r| row >= r.first_row and row <= r.last_row,
                };
            },
        }
    }

    /// Mark one definition for refresh — the save-time half of the
    /// predicate, for a cell write `cellWriteChangesSource` admitted.
    /// A fresh buffer the caller owns, or null when the marker is
    /// already set (the part is then byte-preserved).
    pub fn markForRefresh(allocator: Allocator, cache: *const PivotCache) EditError!?[]u8 {
        var one = [_]Splice{markerSplice(cache.raw_xml, &cache.definition) orelse return null};
        return try spliceAll(allocator, cache.raw_xml, &one);
    }

    /// The marker as a splice on `src`, the bytes `def` was parsed
    /// from: a present attribute has its value replaced (`0` → `1`),
    /// an absent one is inserted before the root's `>` — the shared
    /// attribute writer substitutes values it meets and has no
    /// insertion path, and neither corpus definition carries the
    /// attribute. Null when the definition already carries it.
    pub fn markerSplice(src: []const u8, def: *const pivot_xml.CacheDefinition) ?Splice {
        if (markerSet(def)) return null;
        if (def.rootAttrValueSpan(src, marker_attr)) |span| return .{ .span = span, .text = "1" };
        return .{ .span = .{ .start = def.root_attrs.end, .end = def.root_attrs.end }, .text = marker_insert };
    }

    /// What one row / col edit does to one `pivotCacheDefinitionN.xml`
    /// part: the splices that move its coordinates and set its marker,
    /// and what the S7b-4 rebuild needs to know — whether the edit
    /// changed some source's content, and the rectangle to rebuild
    /// from when it did. `planCacheEdit` computes it; `applyPlan`
    /// renders it. The splice texts live in the arena the plan was
    /// built in.
    pub const Plan = struct {
        splices: std.ArrayListUnmanaged(Splice) = .empty,
        /// The S7b-3 predicate over every source on the edited sheet:
        /// the edit changed what some source HOLDS (an insert inside a
        /// rectangle, a delete inside it, any edit of a sheet an
        /// unbounded body references) — the marker is set, and the
        /// snapshot is no longer the source.
        changed: bool = false,
        /// The rectangle the cache can be rebuilt from when `changed`:
        /// a `worksheet`-type source with a finite rectangle on the
        /// edited sheet, in pre-edit coordinates. Null when `changed`
        /// and no such rectangle exists — whole columns or rows, an
        /// unbounded name body, a consolidation set, a locator under
        /// an unknown `type` — which the rebuild refuses
        /// (`docs/plans/s7b-cache-policy.md` §9, S7b-4).
        rebuild: ?RebuildSource = null,
        /// S7c: the edit is a column edit strictly inside the source
        /// rectangle — the field schema changes, and the rebuild adds
        /// or removes a cache field with it.
        schema: ?SchemaEdit = null,
    };

    /// S7c's schema change, in 0-based field ordinals: a source-column
    /// delete removes the field at `remove` (K3 — admitted only when
    /// no consumer references it); a source-column insert adds one at
    /// `insert.at` (K2 — admitted only for a headerless table, whose
    /// own rewrite names the column).
    pub const SchemaEdit = union(enum) {
        remove: u32,
        insert: Insert,

        pub const Insert = struct {
            at: u32,
            /// The new field's decoded name — the column the table's
            /// own rewrite will synthesize. Empty until the sweep
            /// resolves it from the table part; the engine refuses an
            /// empty one.
            name: []const u8 = "",
        };
    };

    /// Where a rebuild reads: the source rectangle as the sheet is
    /// before the edit, and how many of its top rows are field names.
    pub const RebuildSource = struct {
        rect: Rect,
        header_rows: u32,
        /// A table's totals rows at the bottom of `rect` — not data.
        totals_rows: u32 = 0,
        /// A table's column names — a headerless table's field names.
        columns: []const []const u8 = &.{},
    };

    /// The rectangle a cache is rebuilt from with no edit to apply —
    /// a save with staged cell writes (S7b-5, §7 Q3): a
    /// `worksheet`-type source resolved to a finite rectangle. Null
    /// for every other source, which a save marks and leaves.
    pub fn rebuildSourceOf(cache: *const PivotCache) ?RebuildSource {
        if (cache.definition.source.type != .worksheet) return null;
        const local = switch (cache.resolution) {
            .sheet => |l| l,
            else => return null,
        };
        const b = local.bounds orelse return null;
        if (b != .rect) return null;
        return .{ .rect = b.rect, .header_rows = local.header_rows, .totals_rows = local.totals_rows, .columns = local.columns };
    }

    /// Plan one `pivotCacheDefinitionN.xml` part's rewrite for a row /
    /// col edit on sheet `sheet_idx` — the S7b splice. Every source the
    /// cache reads on that sheet, `worksheetSource` and each `rangeSet`
    /// alike, passes the range semantics above; the ones that carry
    /// their own coordinate (`sheet` + `ref`) are respelled at the
    /// parser's `ref_span` — each set at its own, the part rebuilt from
    /// its raw bytes in span order — the rest move with their carrier — a table
    /// part under `table_edit`, a defined name's body under the name
    /// sweep — and a source on another sheet, in another workbook, or
    /// placed nowhere is left alone. Refuses per §7 Q4: an unplaceable
    /// `r:id` whose `sheet` is this sheet, a spelling that claims the
    /// sheet and bounds nothing. When any source's content changed
    /// under the edit (the S7b-3 predicate above), the root gains the
    /// refresh marker in the same rebuild and the plan says so, with
    /// the rectangle the S7b-4 engine rebuilds the snapshot from.
    pub fn planCacheEdit(
        arena: Allocator,
        cache: *const PivotCache,
        sheet_idx: u32,
        axis: Axis,
        idx_1based: u32,
        kind: Kind,
    ) EditError!Plan {
        if (idx_1based == 0) return error.MalformedPivotXml;
        var plan: Plan = .{};
        const def = &cache.definition;
        if (def.source.worksheet) |ws| {
            const r = try sourceSplice(arena, &plan.splices, ws, cache.resolution, sheet_idx, axis, idx_1based, kind, def.source.type == .worksheet);
            plan.schema = r.schema;
            if (r.changed) {
                plan.changed = true;
                // Only a worksheet-type source is a rectangle the
                // engine reads as records; a carried locator under an
                // unknown `type` moves (Q5) but is not rebuilt.
                if (def.source.type == .worksheet) {
                    if (r.local) |local| {
                        if (local.bounds) |b| {
                            if (b == .rect) plan.rebuild = .{ .rect = b.rect, .header_rows = local.header_rows, .totals_rows = local.totals_rows, .columns = local.columns };
                        }
                    }
                }
            }
        }
        // The walk resolved every set it parsed; a definition that
        // disagrees with its own resolutions is not one this row read.
        if (def.source.range_sets.len != cache.range_set_resolutions.len) return error.MalformedPivotXml;
        for (def.source.range_sets, cache.range_set_resolutions) |rs, res| {
            const r = try sourceSplice(arena, &plan.splices, rs, res, sheet_idx, axis, idx_1based, kind, false);
            plan.changed = plan.changed or r.changed;
        }
        if (plan.changed) {
            if (markerSplice(cache.raw_xml, def)) |sp| try plan.splices.append(arena, sp);
        }
        return plan;
    }

    /// `planCacheEdit` rendered: the part with every planned splice
    /// applied, as a fresh buffer the caller owns; null when the plan
    /// holds none (the part is then byte-preserved).
    pub fn applyPlan(allocator: Allocator, cache: *const PivotCache, plan: *const Plan) EditError!?[]u8 {
        if (plan.splices.items.len == 0) return null;
        return try spliceAll(allocator, cache.raw_xml, plan.splices.items);
    }

    /// The coordinate move and the marker alone — `planCacheEdit` +
    /// `applyPlan`, without the S7b-4 rebuild. Returns a fresh buffer
    /// the caller owns, or null when nothing in the part changes.
    pub fn applyToCacheDefinition(
        allocator: Allocator,
        cache: *const PivotCache,
        sheet_idx: u32,
        axis: Axis,
        idx_1based: u32,
        kind: Kind,
    ) EditError!?[]u8 {
        var arena_state = std.heap.ArenaAllocator.init(allocator);
        defer arena_state.deinit();
        const plan = try planCacheEdit(arena_state.allocator(), cache, sheet_idx, axis, idx_1based, kind);
        // A schema edit needs the engine (S7c): this seam moves the
        // coordinate and the marker alone, and refuses what it cannot
        // install whole.
        if (plan.schema != null) return error.PivotSourceEditUnsafe;
        return try applyPlan(allocator, cache, &plan);
    }

    /// One consumer part under a schema edit (S7c): the removed
    /// ordinal's `<pivotField>` taken out whole (K3) or a bare
    /// `<pivotField showAll="0"/>` inserted at the new ordinal (K2),
    /// `<pivotFields count>` adjusted, and every ordinal carrier the
    /// admitted form holds moved — `<field x>` on the row axis,
    /// `dataField@fld` and `@baseField`. A removed field one of ≥ 2
    /// data fields read is K4a, lifted in S7c-2
    /// (`docs/plans/s7c-column-edits.md` §4 Q2): each `<dataField>`
    /// reading it leaves whole, `<colItems>` re-enumerates the
    /// survivors, `location@ref` narrows by the vanished values
    /// column(s), and a drop to a single survivor collapses the
    /// values axis off the columns — Excel's own one-data-field
    /// spelling, the form every single-data-field fixture attests.
    /// Still refused (K4b): the row axis, the only data field, an
    /// ordinal a surviving `baseField` names. A form with carriers
    /// this rewrite has no spans for — a page field, a real field on
    /// the columns axis, a chart format not proven values-only —
    /// refuses here, as the layout refuses it after.
    pub fn applyConsumerSchemaEdit(arena: Allocator, table_xml: []const u8, schema: SchemaEdit) EditError![]u8 {
        const def = pivot_xml.parseTableDefinition(arena, table_xml) catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            error.MalformedXml => return error.MalformedPivotXml,
        };
        if (def.page_fields.len != 0 or def.chart_formats == .other) return error.PivotShapeUnsupported;
        for (def.col_fields) |cf| if (cf == .field) return error.PivotShapeUnsupported;
        if (def.row_fields.len != def.row_field_x_spans.len) return error.MalformedPivotXml;
        // A wrapper whose `count` disagrees with its children refuses
        // every row edit at the layout; the count splice below would
        // heal the disagreement before the layout re-parses, so it is
        // refused here, where the evidence still exists (in-house
        // review S7C-MUT-1).
        if (def.axes_count_mismatch) return error.MalformedPivotXml;
        // Every ordinal carrier must name a field the part enumerates,
        // checked before any arithmetic on it: a `fld="4294967295"`
        // would otherwise trap in the increment instead of folding
        // into the typed refusal (Codex #208 r1 SEC-104).
        for (def.row_fields) |rf| {
            if (rf == .field and rf.field >= def.fields.len) return error.MalformedPivotXml;
        }
        for (def.data_fields) |df| {
            if (df.fld >= def.fields.len) return error.MalformedPivotXml;
            if (df.base_field) |bf| {
                if (bf >= 0 and @as(u64, @intCast(bf)) >= def.fields.len) return error.MalformedPivotXml;
            }
        }
        // The two directions must AGREE for EVERY field, not only the
        // edited one: the positional `pivotField@dataField` and the
        // `dataField@fld` set name the same ordinals (several data
        // fields on one ordinal are one entry) — a part that
        // disagrees with itself is not one this rewrite re-lays
        // (Codex #208 r4 REL-401; the ruled predicate's own sentence).
        {
            const referenced = try arena.alloc(bool, def.fields.len);
            @memset(referenced, false);
            for (def.data_fields) |df| referenced[df.fld] = true;
            for (def.fields, referenced) |f, r| {
                if (f.data_field != r) return error.MalformedPivotXml;
            }
        }
        // S7c is the first slice that MOVES ordinals, so content this
        // reader tolerates unread must be proven not to carry any. A
        // surviving pivotField's body beyond its `<items>` (an
        // `autoSortScope`, its own `extLst`) can name a field by
        // ordinal — refused; the removed field's body leaves whole
        // with it. The root `<extLst>` is probed for the ordinal-
        // carrier tokens — presence refuses, and the corpus'
        // attribute-only `x14:pivotTableDefinition` carries none
        // (in-house review S7C-R1).
        for (def.fields, 0..) |f, m| {
            if (schema == .remove and schema.remove == m) continue;
            if (f.has_other_children) return error.PivotShapeUnsupported;
        }
        if (def.ext_lst_start) |at| {
            if (extRegionNamesOrdinal(table_xml[at..])) return error.PivotShapeUnsupported;
        }

        var splices: std.ArrayListUnmanaged(Splice) = .empty;
        switch (schema) {
            .remove => |k| {
                if (k >= def.fields.len) return error.MalformedPivotXml;
                const pf = def.fields[k];
                // K4b: an axis field. A data field is K4a — the
                // cross-check above already ties `pf.data_field` to
                // the `dataField@fld` set, so the drop count below is
                // its one source of truth.
                if (pf.axis != null or pf.axis_raw != null) return error.PivotSourceEditUnsafe;
                if (pf.span.start == 0 and pf.span.end == 0) return error.MalformedPivotXml;
                try splices.append(arena, .{ .span = pf.span, .text = "" });
                for (def.row_fields, def.row_field_x_spans) |rf, span| {
                    if (rf != .field) continue;
                    if (rf.field == k) return error.PivotSourceEditUnsafe;
                    if (rf.field > k) try splices.append(arena, .{ .span = span, .text = try std.fmt.allocPrint(arena, "{d}", .{rf.field - 1}) });
                }
                var dropped: u32 = 0;
                for (def.data_fields) |df| {
                    if (df.fld == k) dropped += 1;
                }
                // K4b: the only data field (`wt` twice is two entries
                // and one ordinal — all of them go, or none). The
                // layout has no form without a values column.
                if (dropped > 0 and dropped == def.data_fields.len) return error.PivotSourceEditUnsafe;
                for (def.data_fields) |df| {
                    if (df.fld == k) {
                        // K4a: the data field leaves whole, its own
                        // `baseField` and `baseItem` with it.
                        if (df.span.start == 0 and df.span.end == 0) return error.MalformedPivotXml;
                        try splices.append(arena, .{ .span = df.span, .text = "" });
                        continue;
                    }
                    if (df.fld > k) try splices.append(arena, .{ .span = df.fld_span, .text = try std.fmt.allocPrint(arena, "{d}", .{df.fld - 1}) });
                    if (df.base_field) |bf| {
                        if (bf >= 0) {
                            const b: u32 = @intCast(bf);
                            // K4b: a surviving data field baselined on
                            // the removed ordinal has no defined
                            // successor.
                            if (b == k) return error.PivotSourceEditUnsafe;
                            if (b > k) try splices.append(arena, .{ .span = df.base_field_span.?, .text = try std.fmt.allocPrint(arena, "{d}", .{b - 1}) });
                        }
                    }
                }
                if (dropped > 0) {
                    try narrowValuesAxis(arena, &splices, &def, dropped);
                    try rewriteValuesChartFormats(arena, &splices, &def, k);
                }
                if (def.pivot_fields_count_span) |span| {
                    try splices.append(arena, .{ .span = span, .text = try std.fmt.allocPrint(arena, "{d}", .{def.fields.len - 1}) });
                }
            },
            .insert => |ins| {
                const j = ins.at;
                if (j == 0 or j >= def.fields.len) return error.MalformedPivotXml;
                const anchor = def.fields[j].span.start;
                if (anchor == 0) return error.MalformedPivotXml;
                const p = if (def.prefix.len == 0) "" else try std.mem.concat(arena, u8, &.{ def.prefix, ":" });
                const fresh = try std.mem.concat(arena, u8, &.{ "<", p, "pivotField showAll=\"0\"/>" });
                try splices.append(arena, .{ .span = .{ .start = anchor, .end = anchor }, .text = fresh });
                for (def.row_fields, def.row_field_x_spans) |rf, span| {
                    if (rf != .field or rf.field < j) continue;
                    try splices.append(arena, .{ .span = span, .text = try std.fmt.allocPrint(arena, "{d}", .{rf.field + 1}) });
                }
                for (def.data_fields) |df| {
                    if (df.fld >= j) try splices.append(arena, .{ .span = df.fld_span, .text = try std.fmt.allocPrint(arena, "{d}", .{df.fld + 1}) });
                    if (df.base_field) |bf| {
                        if (bf >= 0 and @as(u32, @intCast(bf)) >= j) {
                            try splices.append(arena, .{ .span = df.base_field_span.?, .text = try std.fmt.allocPrint(arena, "{d}", .{@as(u32, @intCast(bf)) + 1}) });
                        }
                    }
                }
                if (def.pivot_fields_count_span) |span| {
                    try splices.append(arena, .{ .span = span, .text = try std.fmt.allocPrint(arena, "{d}", .{def.fields.len + 1}) });
                }
            },
        }
        // Every span this plan emits is disjoint: a nested one (a
        // `fld` inside a removed `<dataField>`) is a bug HERE, not a
        // malformed part — assert the producer's invariant where it
        // is produced, so it cannot present as an input refusal
        // (in-house review S7C2-A3).
        if (std.debug.runtime_safety) {
            std.mem.sort(Splice, splices.items, {}, Splice.before);
            var i: usize = 1;
            while (i < splices.items.len) : (i += 1) {
                assert(splices.items[i - 1].span.end <= splices.items[i].span.start);
            }
        }
        return try spliceAll(arena, table_xml, splices.items);
    }

    /// K4a's values-axis narrow (S7c-2): with `dropped` of the part's
    /// data fields gone, `<colItems>`, the `<dataFields count>` and
    /// `location@ref` must follow, and a drop to one survivor takes
    /// `<colFields>` (the `x="-2"` values axis) out whole — Excel's
    /// own one-data-field spelling, the form every single-data-field
    /// fixture attests. The wrappers are regenerated, so the canonical
    /// form the layout admits is required FIRST and anything else
    /// refuses: a regenerated wrapper would erase the evidence the
    /// layout's own checks refuse on (the S7C-MUT-1 rule). Host-cell
    /// styles carry by POSITION (S7b-5's rule), so a survivor shifted
    /// left wears the vanished column's format until Excel's
    /// open-refresh re-lays it — the marker covers it, and the corpus
    /// cannot tell (every mtCars data cell shares one style).
    fn narrowValuesAxis(
        arena: Allocator,
        splices: *std.ArrayListUnmanaged(Splice),
        def: *const pivot_xml.TableDefinition,
        dropped: u32,
    ) EditError!void {
        assert(dropped > 0);
        assert(dropped < def.data_fields.len);
        const survivors: usize = def.data_fields.len - dropped;
        // An axis wrapper carrying content the parser did not classify
        // — a stray attribute, a child that is not its field, markup
        // between children — is evidence the layout refuses on, and
        // the collapse arm below removes `<colFields>` whole: healed
        // evidence, the S7C-MUT-1 rule (Codex #210 r1 REL-101). Refuse
        // while it still exists; the K3 path never reads this gate.
        if (def.axes_other) return error.PivotShapeUnsupported;
        // ≥ 2 data fields spell the values axis across; the single-
        // data-field form (no colFields) has no second entry for a
        // drop to leave behind.
        if (def.col_fields.len != 1 or def.col_fields[0] != .values) return error.PivotShapeUnsupported;
        const ci = def.col_items orelse return error.PivotShapeUnsupported;
        if (ci.other_attrs) return error.PivotShapeUnsupported;
        if (ci.count) |n| if (n != ci.items.len) return error.MalformedPivotXml;
        if (ci.items.len != def.data_fields.len) return error.PivotShapeUnsupported;
        for (ci.items, 0..) |it, j| {
            if (it.t != null or it.r != null or it.other_attrs or it.has_other_children) return error.PivotShapeUnsupported;
            if (it.xs.len != 1 or it.xs[0] != j or (it.i orelse 0) != j) return error.PivotShapeUnsupported;
        }
        const p = if (def.prefix.len == 0) "" else try std.mem.concat(arena, u8, &.{ def.prefix, ":" });
        if (survivors == 1) {
            const cf_span = def.col_fields_span orelse return error.MalformedPivotXml;
            try splices.append(arena, .{ .span = cf_span, .text = "" });
            try splices.append(arena, .{ .span = ci.span, .text = try std.mem.concat(arena, u8, &.{ "<", p, "colItems count=\"1\"><", p, "i/></", p, "colItems>" }) });
        } else {
            var out: std.ArrayListUnmanaged(u8) = .empty;
            try out.append(arena, '<');
            try out.appendSlice(arena, p);
            try out.appendSlice(arena, "colItems count=\"");
            try out.appendSlice(arena, try std.fmt.allocPrint(arena, "{d}", .{survivors}));
            try out.appendSlice(arena, "\">");
            for (0..survivors) |j| {
                try out.append(arena, '<');
                try out.appendSlice(arena, p);
                if (j == 0) {
                    try out.appendSlice(arena, "i><");
                    try out.appendSlice(arena, p);
                    try out.appendSlice(arena, "x/></");
                } else {
                    try out.appendSlice(arena, "i i=\"");
                    try out.appendSlice(arena, try std.fmt.allocPrint(arena, "{d}", .{j}));
                    try out.appendSlice(arena, "\"><");
                    try out.appendSlice(arena, p);
                    try out.appendSlice(arena, "x v=\"");
                    try out.appendSlice(arena, try std.fmt.allocPrint(arena, "{d}", .{j}));
                    try out.appendSlice(arena, "\"/></");
                }
                try out.appendSlice(arena, p);
                try out.appendSlice(arena, "i>");
            }
            try out.appendSlice(arena, "</");
            try out.appendSlice(arena, p);
            try out.appendSlice(arena, "colItems>");
            try splices.append(arena, .{ .span = ci.span, .text = out.items });
        }
        if (def.data_fields_count_span) |span| {
            try splices.append(arena, .{ .span = span, .text = try std.fmt.allocPrint(arena, "{d}", .{survivors}) });
        }
        // The location narrows with the vanished values column(s). The
        // width invariant is the layout's own — a part whose rectangle
        // is not one label column plus its data fields is not one this
        // rewrite narrows.
        const fp = try footprintOf(arena, def.*);
        const r = fp.rect;
        if (@as(u64, r.br_col) - r.tl_col + 1 != @as(u64, def.data_fields.len) + 1) return error.MalformedPivotXml;
        var narrowed = r;
        narrowed.br_col -= dropped;
        var ref_buf: [Bounds.format_buf_len]u8 = undefined;
        const new_ref = formatRect(&ref_buf, narrowed) catch return error.PivotCoordinateOverflow;
        try splices.append(arena, .{ .span = def.location.ref_span, .text = try arena.dupe(u8, new_ref) });
    }

    /// K4a's chart-format move (S7c-2): a values-only `<chartFormat>`
    /// selects a data field BY INDEX (`<x v>` under its
    /// `field="4294967294"` reference) — the one admitted carrier of
    /// data-field indices S7b-5 left as written, sound only while no
    /// edit moved them (the corpus' mtCars pivot chart rides three
    /// such blocks; in-house review S7C2-B8). A block naming a
    /// dropped index leaves whole; a survivor's index decrements in
    /// place; `<chartFormats count>` follows; a rewrite that empties
    /// the list takes the element out whole (the chart part itself
    /// stays as written — Excel's open-refresh re-lays a pivot
    /// chart's series from the pivot, the S7b-5 safety-net rule).
    /// A shape the collector could not read as one-block-one-index —
    /// a second `<chartFormats>`, a non-canonical block, an index
    /// past the data fields, a lying `count` — refuses rather than
    /// heals (the S7C-MUT-1 rule); the K3 path never reads any of
    /// this.
    fn rewriteValuesChartFormats(
        arena: Allocator,
        splices: *std.ArrayListUnmanaged(Splice),
        def: *const pivot_xml.TableDefinition,
        k: u32,
    ) EditError!void {
        // Multiplicity first: a trailing element could have folded the
        // classification down while the first element's blocks are the
        // ones going stale (Codex #210 r2 REL-202).
        if (def.chart_formats_multi) return error.PivotShapeUnsupported;
        if (def.chart_formats != .values_only) return;
        const refs = def.chart_format_values_refs;
        if (def.chart_formats_count) |n| {
            if (n != refs.len) return error.MalformedPivotXml;
        }
        // The survivor map: data-field position -> its index after the
        // drop, null for a dropped position.
        const map = try arena.alloc(?u32, def.data_fields.len);
        var next: u32 = 0;
        for (def.data_fields, map) |df, *m| {
            m.* = if (df.fld == k) null else blk: {
                const v = next;
                next += 1;
                break :blk v;
            };
        }
        var kept: usize = 0;
        for (refs) |r| {
            if (!r.canonical) return error.PivotShapeUnsupported;
            if (r.index >= def.data_fields.len) return error.MalformedPivotXml;
            if (map[r.index] != null) kept += 1;
        }
        if (kept == 0 and refs.len > 0) {
            const span = def.chart_formats_span orelse return error.MalformedPivotXml;
            try splices.append(arena, .{ .span = span, .text = "" });
            return;
        }
        for (refs) |r| {
            const new = map[r.index] orelse {
                try splices.append(arena, .{ .span = r.span, .text = "" });
                continue;
            };
            if (new != r.index) {
                // An absent `v` is index 0, which no drop moves; a
                // moved index was spelled.
                const span = r.v_span orelse return error.MalformedPivotXml;
                try splices.append(arena, .{ .span = span, .text = try std.fmt.allocPrint(arena, "{d}", .{new}) });
            }
        }
        if (kept != refs.len) {
            if (def.chart_formats_count_span) |span| {
                try splices.append(arena, .{ .span = span, .text = try std.fmt.allocPrint(arena, "{d}", .{kept}) });
            }
        }
    }

    /// Does the extension region spell an attribute named `field` or
    /// `fld` — an ordinal carrier this reader tolerates unread? The
    /// probe matches the attribute NAME at a name boundary with XML
    /// whitespace admitted around the `=` (Codex #208 r1 SEC-101:
    /// `field = '2'` is the same attribute as `field='2'`; attribute
    /// names are never entity-spelled), so `fieldPosition=` and
    /// `sourceField=` stay outside it. Presence anywhere in the
    /// region — a comment included — refuses: over-matching costs a
    /// refusal.
    fn extRegionNamesOrdinal(tail: []const u8) bool {
        for ([_][]const u8{ "field", "fld" }) |name| {
            var from: usize = 0;
            while (std.mem.indexOfPos(u8, tail, from, name)) |hit| {
                from = hit + 1;
                if (hit == 0) continue;
                const before = tail[hit - 1];
                if (before != ' ' and before != '\t' and before != '\n' and before != '\r' and before != ':') continue;
                var j = hit + name.len;
                while (j < tail.len and (tail[j] == ' ' or tail[j] == '\t' or tail[j] == '\n' or tail[j] == '\r')) j += 1;
                if (j < tail.len and tail[j] == '=') return true;
            }
        }
        return false;
    }

    /// `src` rebuilt from its own bytes with each splice swapped in
    /// place, in span order, so no span moves under another. An
    /// insertion is a splice with an empty span; two insertions at one
    /// position keep their order (the sort is stable). A span that is
    /// reversed, overlaps an earlier one or reaches past `src` is
    /// `MalformedPivotXml` — a public seam answers, it does not assert
    /// (Codex #205 r7 REL-701).
    pub fn spliceAll(allocator: Allocator, src: []const u8, splices: []Splice) EditError![]u8 {
        std.mem.sort(Splice, splices, {}, Splice.before);
        var out: std.ArrayListUnmanaged(u8) = .empty;
        errdefer out.deinit(allocator);
        var pos: usize = 0;
        for (splices) |sp| {
            if (pos > sp.span.start or sp.span.start > sp.span.end or sp.span.end > src.len) return error.MalformedPivotXml;
            try out.appendSlice(allocator, src[pos..sp.span.start]);
            try out.appendSlice(allocator, sp.text);
            pos = sp.span.end;
        }
        try out.appendSlice(allocator, src[pos..]);
        return try out.toOwnedSlice(allocator);
    }

    /// One replacement on a part: `text` in place of `span` (empty
    /// span = insertion). `text` is borrowed — a literal, or bytes in
    /// the arena the plan was built in.
    pub const Splice = struct {
        span: pivot_xml.Span,
        text: []const u8,

        fn before(_: void, a: Splice, b: Splice) bool {
            return a.span.start < b.span.start;
        }
    };

    const SourceOutcome = struct {
        /// The edit changed what this source holds — the marker's input.
        changed: bool,
        /// The local resolution the outcome was judged on, when it was
        /// one on the edited sheet.
        local: ?SourceResolution.LocalSheet = null,
        /// S7c: the edit changes the field schema (a column edit
        /// strictly inside the source rectangle).
        schema: ?SchemaEdit = null,
    };

    /// A column edit strictly inside a finite source rectangle is the
    /// field schema (S7c). A delete removes the field at that column —
    /// unless it would collapse the rectangle (K5, the row-collapse
    /// twin). An insert adds a field there — admitted only where the
    /// new field has a name the engine can prove, a headerless table's
    /// synthesized column (K2); every other insert refuses (K1: the
    /// new header cell is blank, and Excel's own refresh fails on it
    /// with *"The PivotTable field name is not valid"*). Outside the
    /// rectangle is a shift or a no-op, not this function's.
    fn schemaEditFor(s: SourceResolution.LocalSheet, r: Rect, idx_1based: u32, kind: Kind) EditError!?SchemaEdit {
        switch (kind) {
            .insert => {
                if (idx_1based <= r.tl_col or idx_1based > r.br_col) return null;
                if (s.header_rows != 0) return error.PivotSourceEditUnsafe;
                return .{ .insert = .{ .at = idx_1based - r.tl_col } };
            },
            .delete => {
                if (idx_1based < r.tl_col or idx_1based > r.br_col) return null;
                if (r.tl_col == r.br_col) return error.PivotSourceEditUnsafe;
                return .{ .remove = idx_1based - r.tl_col };
            },
        }
    }

    /// One source under the edit: its `ref` splice appended when it has
    /// one, its refusal raised, and whether the edit changed what it
    /// holds — the marker's input. `allow_schema` opens the S7c arm —
    /// the worksheet source of a `worksheet`-type cache; a consolidation
    /// `rangeSet` or an unknown-`type` locator keeps the refusal (the
    /// engine rebuilds neither).
    fn sourceSplice(
        arena: Allocator,
        splices: *std.ArrayListUnmanaged(Splice),
        ws: pivot_xml.WorksheetSource,
        res: SourceResolution,
        sheet_idx: u32,
        axis: Axis,
        idx_1based: u32,
        kind: Kind,
        allow_schema: bool,
    ) EditError!SourceOutcome {
        switch (res) {
            .external, .none => return .{ .changed = false },
            .unresolved => |u| {
                if (std.mem.indexOfScalar(u32, u.sheets, sheet_idx) == null) return .{ .changed = false };
                switch (u.why) {
                    // The `sheet` beside an `r:id` the reader could not
                    // place: it may be this sheet, and its `ref` cannot
                    // be moved without knowing (Q4 i).
                    .unplaceable_rid => return error.PivotSourceEditUnsafe,
                    // A name body reaching this sheet without bounding
                    // it: the body moves under the name sweep, whose
                    // dry-run is the workbook's (Q4 ii). Nothing to
                    // move here — and no shift to prove, so the
                    // content may have changed.
                    .unbounded_body => return .{ .changed = true },
                    // These prove no sheet; `sheets` is empty for them.
                    .dangling_sheet, .dangling_name, .sheetless_ref, .no_locator => return .{ .changed = false },
                }
            },
            .sheet => |s| {
                if (s.sheet_idx != sheet_idx) return .{ .changed = false };
                // Claims the sheet, bounds nothing: `sheet` alone, a
                // `ref` the bounds parser rejects, a name the reader
                // could not place on that sheet (Q4 iv).
                const bounds = s.bounds orelse return error.PivotSourceEditUnsafe;
                // S7c first: a column edit strictly inside the finite
                // rectangle changes the field schema. The coordinate:
                // a direct `ref` respells the shrunk / grown rectangle;
                // a table or a name body moves with its own carrier.
                if (allow_schema and axis == .col and bounds == .rect) {
                    if (try schemaEditFor(s, bounds.rect, idx_1based, kind)) |se| {
                        if (s.carrier == .ref) {
                            const span = ws.ref_span orelse return error.MalformedPivotXml;
                            var moved = bounds.rect;
                            switch (kind) {
                                .insert => {
                                    if (moved.br_col >= zlsx.max_col_1based) return error.PivotCoordinateOverflow;
                                    moved.br_col += 1;
                                },
                                .delete => moved.br_col -= 1,
                            }
                            var buf: [Bounds.format_buf_len]u8 = undefined;
                            const text = formatRect(&buf, moved) catch return error.PivotCoordinateOverflow;
                            try splices.append(arena, .{ .span = span, .text = try arena.dupe(u8, text) });
                        }
                        return .{ .changed = true, .local = s, .schema = se };
                    }
                }
                switch (s.carrier) {
                    .none => unreachable, // bounds and carrier are set together
                    // A table carries its own row rules — `table_edit`
                    // refuses a header-row or collapsing delete and
                    // admits the top row of a `headerRowCount="0"` table
                    // — and moves its own `ref`; only the column axis,
                    // the field schema, is judged here. Whether the
                    // sheet came from the table or from a `sheet`
                    // attribute beside its name makes no difference
                    // (Codex #203 r1 REL-102).
                    .table => {
                        if (axis == .col) _ = try shiftSourceBounds(bounds, axis, idx_1based, kind);
                        return .{ .changed = editChangesContent(bounds, axis, idx_1based, kind), .local = s };
                    },
                    .ref, .defined_name => {},
                }
                // Refusals first: a content change is judged only on an
                // admitted edit.
                const shifted = try shiftSourceBounds(bounds, axis, idx_1based, kind);
                const changed = editChangesContent(bounds, axis, idx_1based, kind);
                const outcome: SourceOutcome = .{ .changed = changed, .local = s };
                const moved = shifted orelse return outcome;
                // Only a spelling with its own `ref` is respelled; a
                // name-spelled area moves with the name's body.
                if (s.carrier != .ref) return outcome;
                const span = ws.ref_span orelse return error.MalformedPivotXml;
                // A rectangle keeps the single-cell spelling `A1`;
                // whole columns / rows spell `A:C` / `1:4` as read.
                var buf: [Bounds.format_buf_len]u8 = undefined;
                const text = switch (moved) {
                    .rect => |r| formatRect(&buf, r) catch return error.PivotCoordinateOverflow,
                    else => moved.formatA1(&buf) orelse return error.PivotCoordinateOverflow,
                };
                try splices.append(arena, .{ .span = span, .text = try arena.dupe(u8, text) });
                return outcome;
            },
        }
    }

    /// `A1` or `A1:C4`, plain A1 only. Null on anything else.
    pub fn parseRect(ref: []const u8) ?Rect {
        const colon = std.mem.indexOfScalar(u8, ref, ':');
        const tl = parseCell(if (colon) |c| ref[0..c] else ref) orelse return null;
        const br = if (colon) |c| (parseCell(ref[c + 1 ..]) orelse return null) else tl;
        if (br.col < tl.col or br.row < tl.row) return null;
        return .{ .tl_col = tl.col, .tl_row = tl.row, .br_col = br.col, .br_row = br.row };
    }

    const Cell = struct { col: u32, row: u32 };

    fn parseCell(s: []const u8) ?Cell {
        var letters: usize = 0;
        while (letters < s.len and s[letters] >= 'A' and s[letters] <= 'Z') letters += 1;
        if (letters == 0 or letters > 3 or letters == s.len) return null;
        const digits = s[letters..];
        // `A01` is not a cell Excel writes; a leading zero would round-
        // trip to a different spelling, so it is refused with the rest.
        if (digits[0] == '0') return null;
        for (digits) |d| if (d < '0' or d > '9') return null;
        const col = sheet_edit.parseColLetters(s[0..letters]) orelse return null;
        const row = std.fmt.parseInt(u32, digits, 10) catch return null;
        if (col == 0 or col > zlsx.max_col_1based or row == 0 or row > zlsx.max_row) return null;
        return .{ .col = col, .row = row };
    }

    pub fn formatRect(buf: *[Bounds.format_buf_len]u8, r: Rect) ![]const u8 {
        var tl_buf: [8]u8 = undefined;
        var br_buf: [8]u8 = undefined;
        const tl_len = try coords.writeColNumberLetters(&tl_buf, r.tl_col);
        const br_len = try coords.writeColNumberLetters(&br_buf, r.br_col);
        if (r.tl_col == r.br_col and r.tl_row == r.br_row) {
            return std.fmt.bufPrint(buf, "{s}{d}", .{ tl_buf[0..tl_len], r.tl_row });
        }
        return std.fmt.bufPrint(buf, "{s}{d}:{s}{d}", .{ tl_buf[0..tl_len], r.tl_row, br_buf[0..br_len], r.br_row });
    }
};

// ─── S7b-4: the cache rebuild — the engine's first slice ─────────────
//
// The owner's cache policy is B (`docs/plans/s7b-cache-policy.md`
// §8): zlsx PERFORMS the refresh. This slice rebuilds the CACHE from
// the source cells — the records, every field's inventory,
// `recordCount`, `refreshedDate` — for the shapes it can read and
// write exactly; the consumers' items, layout and output cells are
// the next slice, and until they land the refresh marker (S7b-3)
// stays on a rebuilt cache, so Excel's own refresh at open lays the
// consumers out over a snapshot that is already the source.
//
// Two invariants make a rebuilt cache one its consumers still index:
// an inventory keeps every item it had, in its order — a
// `pivotField/items/item@x`, a `rowItems` position, a chart's
// selection keep naming the same value; Excel itself retains items no
// record references, up to `missingItemsLimit` — and a value the
// inventory lacks is appended after them, in first-appearance order.
// A field whose records were inline (`<n>` in the records part — a
// data-only numeric field) stays inline unless it now holds a string;
// a field whose inventory was enumerated stays enumerated. Items match
// as Excel groups them: numbers by value, strings case-insensitively
// under the workbook's collation, spelled by their first occurrence.
//
// Everything the slice does not evaluate REFUSES the edit rather than
// write a partial rebuild (§8 Q1): calculated and group fields, OLAP
// and consolidation shapes, a source without a finite rectangle, a
// records part carrying anything but records, and cells the oracle
// matrix has not covered — dates (a `t="d"` cell, a number under a
// date format, an inventory that held dates), booleans, errors, an
// uncomputed formula, an inline string. A pure shift never reaches
// the engine: the S7b-3 predicate gates it, so a part Excel would
// leave byte-identical still is.

pub const engine = struct {
    pub const RebuildError = error{
        /// A shape this slice does not rebuild — see the namespace
        /// note. The edit refuses; nothing is written.
        PivotShapeUnsupported,
        /// The definition or records part disagrees with itself (an
        /// item without its value, a `count` that is not the number
        /// of items, a row that is not one value per field).
        MalformedPivotXml,
        OutOfMemory,
    };

    /// One source cell as the rebuild reads it — what the workbook
    /// resolves from the sheet's typed view, the shared strings and
    /// the styles before the engine sees it.
    pub const Value = union(enum) {
        blank,
        /// The cell's `<v>` as written: an xsd:double lexical the
        /// producer wrote and Excel read, kept verbatim so a value
        /// round-trips to the byte (`4.4000000000000004` stays so).
        /// Finite — the reader checks, the engine re-checks.
        number: []const u8,
        /// The cell's text, decoded (entities and ST_Xstring resolved).
        string: []const u8,
    };

    /// One data row of the source rectangle, one value per field.
    pub const Row = []const Value;

    /// The most cells — rectangle width × height — one rebuild reads:
    /// 16 Mi, a million-row source of sixteen fields. A finite
    /// rectangle past it refuses before any read (a hand-written
    /// `A1:XFD1048576` is not a rectangle this slice reads — Codex #205
    /// r3 PERF-301); a streaming read is a later slice's.
    pub const max_rebuild_cells: usize = 1 << 24;

    /// The data rows after a row edit inside the rectangle, in source
    /// order: an insert at `idx_1based` puts a blank row there, a
    /// delete drops the row there. `first_data_row` is the sheet row
    /// of `rows[0]` — the rectangle's top plus its header rows. The
    /// edit is one `edit.editChangesContent` admitted, so it lands on
    /// a data row.
    pub fn rowsAfterEdit(arena: Allocator, rows: []const Row, width: usize, first_data_row: u32, idx_1based: u32, kind: edit.Kind) RebuildError![]const Row {
        if (idx_1based < first_data_row) return error.MalformedPivotXml;
        const k: usize = idx_1based - first_data_row;
        switch (kind) {
            .insert => {
                if (k > rows.len) return error.MalformedPivotXml;
                const out = try arena.alloc(Row, rows.len + 1);
                @memcpy(out[0..k], rows[0..k]);
                const blank = try arena.alloc(Value, width);
                @memset(blank, .blank);
                out[k] = blank;
                @memcpy(out[k + 1 ..], rows[k..]);
                return out;
            },
            .delete => {
                if (k >= rows.len) return error.MalformedPivotXml;
                const out = try arena.alloc(Row, rows.len - 1);
                @memcpy(out[0..k], rows[0..k]);
                @memcpy(out[k..], rows[k + 1 ..]);
                return out;
            },
        }
    }

    /// The data rows after a column edit inside the rectangle (S7c),
    /// each at the effective width: a delete drops every row's value
    /// at the removed 0-based ordinal, an insert puts a blank there.
    /// `width` is the pre-edit field count.
    pub fn rowsAfterColEdit(arena: Allocator, rows: []const Row, width: usize, ordinal: u32, kind: edit.Kind) RebuildError![]const Row {
        const k: usize = ordinal;
        const out = try arena.alloc(Row, rows.len);
        switch (kind) {
            .insert => {
                if (k == 0 or k > width) return error.MalformedPivotXml;
                for (rows, 0..) |r, i| {
                    if (r.len != width) return error.MalformedPivotXml;
                    const nr = try arena.alloc(Value, width + 1);
                    @memcpy(nr[0..k], r[0..k]);
                    nr[k] = .blank;
                    @memcpy(nr[k + 1 ..], r[k..]);
                    out[i] = nr;
                }
            },
            .delete => {
                if (k >= width) return error.MalformedPivotXml;
                for (rows, 0..) |r, i| {
                    if (r.len != width) return error.MalformedPivotXml;
                    const nr = try arena.alloc(Value, width - 1);
                    @memcpy(nr[0..k], r[0..k]);
                    @memcpy(nr[k..], r[k + 1 ..]);
                    out[i] = nr;
                }
            },
        }
        return out;
    }

    /// The refresh instant: an Excel serial under the workbook's date
    /// system for `refreshedDate`, and the same instant as ISO 8601
    /// for a part that also spells `refreshedDateIso` (Codex #205 r9
    /// REL-902).
    pub const Refreshed = struct {
        serial: f64,
        iso: []const u8,
    };

    /// The rebuilt parts.
    pub const Rebuild = struct {
        /// Splices on the definition: every field's inventory element
        /// replaced whole, `recordCount` and `refreshedDate` on the
        /// root. Texts live in the arena the rebuild ran in.
        splices: []edit.Splice,
        /// The records part rebuilt whole — its root tag as written
        /// with `count` set, one `<r>` per data row — or null when the
        /// cache names no records part.
        records: ?[]u8,
        record_count: u32,
        /// The rebuilt fields — each one's final item list and, per
        /// data row, the item it indexes — and the rows they were
        /// built from: what the consumers' slice lays out (S7b-5).
        fields: []const Field,
        rows: []const Row,
    };

    /// The definition's fields under a schema edit: the old list less
    /// the removed one, or with the new field spliced in at its
    /// ordinal — the view every consumer-side check indexes after the
    /// edit. The inserted view entry carries the new field's decoded
    /// name and no inventory; null schema hands the old list back.
    pub fn effectiveCacheFields(arena: Allocator, old: []const pivot_xml.CacheField, schema: ?edit.SchemaEdit) RebuildError![]const pivot_xml.CacheField {
        const se = schema orelse return old;
        switch (se) {
            .remove => |k| {
                if (k >= old.len) return error.MalformedPivotXml;
                const out = try arena.alloc(pivot_xml.CacheField, old.len - 1);
                @memcpy(out[0..k], old[0..k]);
                @memcpy(out[k..], old[k + 1 ..]);
                return out;
            },
            .insert => |ins| {
                if (ins.at == 0 or ins.at > old.len or ins.name.len == 0) return error.MalformedPivotXml;
                const out = try arena.alloc(pivot_xml.CacheField, old.len + 1);
                @memcpy(out[0..ins.at], old[0..ins.at]);
                out[ins.at] = .{ .name = ins.name };
                @memcpy(out[ins.at + 1 ..], old[ins.at..]);
                return out;
            },
        }
    }

    /// Rebuild `cache` from `rows`, the data rows of its source
    /// rectangle after the edit (`rowsAfterEdit`), each one value per
    /// field. `records_xml` is the cache's records part as stored,
    /// when it has one. `refreshed` is the refresh instant, or null
    /// when the caller has no clock — `refreshedDate` (and a
    /// `refreshedDateIso` beside it) is then removed rather than left
    /// describing a refresh that did not happen.
    pub fn rebuild(arena: Allocator, cache: *const PivotCache, rows: []const Row, records_xml: ?[]const u8, refreshed: ?Refreshed) RebuildError!Rebuild {
        return rebuildWith(arena, cache, rows, records_xml, refreshed, null);
    }

    /// `rebuild` with S7c's schema edit, when the edit is one: `rows`
    /// are then already the effective width, and the definition gains
    /// the field splice — the removed `<cacheField>` taken out whole,
    /// the inserted one rendered fresh — beside `<cacheFields count>`.
    pub fn rebuildWith(arena: Allocator, cache: *const PivotCache, rows: []const Row, records_xml: ?[]const u8, refreshed: ?Refreshed, schema: ?edit.SchemaEdit) RebuildError!Rebuild {
        const def = &cache.definition;
        try checkShape(def);
        if (schema) |se| {
            // S7c moves cache-field ordinals, so extension content
            // this reader tolerates unread must be proven not to
            // carry any — the cache-side twin of the consumer probe
            // (Codex #208 r2 REL-203). A SURVIVING field's own
            // `extLst` refuses; the removed field's leaves whole with
            // it. The root extension region is probed for the
            // ordinal-carrier attribute names — the corpus'
            // `pivotCacheId` extension carries none.
            if (def.ext_lst_start) |at| {
                // A caller-assembled offset answers, it does not trap
                // — the parser always points at a `<` inside the part
                // (Codex #208 r5 REL-503).
                if (at >= cache.raw_xml.len or cache.raw_xml[at] != '<') return error.MalformedPivotXml;
                if (edit.extRegionNamesOrdinal(cache.raw_xml[at..])) return error.PivotShapeUnsupported;
            }
            for (def.fields, 0..) |f, k| {
                switch (se) {
                    .remove => |r| if (r == k) continue,
                    .insert => {},
                }
                if (f.has_ext_lst) return error.PivotShapeUnsupported;
            }
        }
        const eff = try effectiveCacheFields(arena, def.fields, schema);
        if (rows.len == 0 or rows.len > std.math.maxInt(u32)) return error.PivotShapeUnsupported;
        // The rectangle's width is the field schema: a disagreement
        // is a shape the engine does not read.
        if (rows[0].len != eff.len) return error.PivotShapeUnsupported;
        for (rows) |r| if (r.len != eff.len) return error.MalformedPivotXml;

        const p = try qualified(arena, def.prefix);
        // The records as written say how each field spelt its values —
        // `<x>` into the inventory, or inline — which is what "stays
        // inline" is measured against (Codex #205 r10 REL-1001). Read
        // at the written arity, then mapped to the effective schema —
        // a removed field's spelling leaves with it, an inserted one
        // has none to read.
        const spelt_old = if (records_xml) |xml| try inspectRecords(arena, xml, def.fields.len) else try arena.alloc(?bool, def.fields.len);
        if (records_xml == null) @memset(spelt_old, null);
        const spelt = try effectiveSpelt(arena, spelt_old, schema);
        const fields = try arena.alloc(Field, eff.len);
        for (eff, 0..) |f, k| {
            const si = f.shared_items orelse blk: {
                // Only the inserted field has no inventory element;
                // `checkShape` required one of every written field.
                if (!insertedAt(schema, k)) return error.MalformedPivotXml;
                break :blk pivot_xml.SharedItems{};
            };
            fields[k] = try Field.build(arena, si, rows, k, p, spelt[k]);
        }

        var splices: std.ArrayListUnmanaged(edit.Splice) = .empty;
        for (eff, fields, 0..) |f, built, k| {
            if (insertedAt(schema, k)) continue;
            try splices.append(arena, .{ .span = f.shared_items.?.span, .text = built.xml });
        }
        if (schema) |se| try appendSchemaSplices(arena, &splices, def, se, fields, p);

        // Root attributes: replaced where present, inserted before the
        // root's `>` where absent — one insertion for both, so their
        // order is this writer's, not a sort's.
        const count_text = try std.fmt.allocPrint(arena, "{d}", .{rows.len});
        var insert: std.ArrayListUnmanaged(u8) = .empty;
        if (def.rootAttrValueSpan(cache.raw_xml, "recordCount")) |span| {
            try splices.append(arena, .{ .span = span, .text = count_text });
        } else {
            try insert.appendSlice(arena, " recordCount=\"");
            try insert.appendSlice(arena, count_text);
            try insert.append(arena, '"');
        }
        if (refreshed) |now| {
            const date_text = try std.fmt.allocPrint(arena, "{d}", .{now.serial});
            if (def.rootAttrValueSpan(cache.raw_xml, "refreshedDate")) |span| {
                try splices.append(arena, .{ .span = span, .text = date_text });
            } else {
                try insert.appendSlice(arena, " refreshedDate=\"");
                try insert.appendSlice(arena, date_text);
                try insert.append(arena, '"');
            }
            // The ISO spelling, where the part carries it, is the same
            // instant — never a stale one beside a fresh serial (Codex
            // #205 r9 REL-902); a part without it does not gain one.
            if (def.rootAttrValueSpan(cache.raw_xml, "refreshedDateIso")) |span| {
                try splices.append(arena, .{ .span = span, .text = now.iso });
            }
        } else {
            if (rootAttrSpan(cache.raw_xml, def, "refreshedDate")) |span| try splices.append(arena, .{ .span = span, .text = "" });
            if (rootAttrSpan(cache.raw_xml, def, "refreshedDateIso")) |span| try splices.append(arena, .{ .span = span, .text = "" });
        }
        if (insert.items.len > 0) {
            try splices.append(arena, .{ .span = .{ .start = def.root_attrs.end, .end = def.root_attrs.end }, .text = insert.items });
        }

        const records = if (records_xml) |xml| try renderRecords(arena, xml, fields, rows) else null;
        return .{ .splices = try splices.toOwnedSlice(arena), .records = records, .record_count = @intCast(rows.len), .fields = fields, .rows = rows };
    }

    /// Is effective ordinal `k` the field a schema insert adds?
    fn insertedAt(schema: ?edit.SchemaEdit, k: usize) bool {
        const se = schema orelse return false;
        return switch (se) {
            .remove => false,
            .insert => |ins| ins.at == k,
        };
    }

    /// The written per-field record spelling mapped to the effective
    /// schema: a removed field's spelling leaves with it; an inserted
    /// field had no records to spell it.
    fn effectiveSpelt(arena: Allocator, old: []const ?bool, schema: ?edit.SchemaEdit) RebuildError![]const ?bool {
        const se = schema orelse return old;
        switch (se) {
            .remove => |k| {
                if (k >= old.len) return error.MalformedPivotXml;
                const out = try arena.alloc(?bool, old.len - 1);
                @memcpy(out[0..k], old[0..k]);
                @memcpy(out[k..], old[k + 1 ..]);
                return out;
            },
            .insert => |ins| {
                if (ins.at > old.len) return error.MalformedPivotXml;
                const out = try arena.alloc(?bool, old.len + 1);
                @memcpy(out[0..ins.at], old[0..ins.at]);
                out[ins.at] = null;
                @memcpy(out[ins.at + 1 ..], old[ins.at..]);
                return out;
            },
        }
    }

    /// The definition's field-list splices for a schema edit: the
    /// removed `<cacheField>` taken out whole, or the new one rendered
    /// before the field at its ordinal — `name` encoded, `numFmtId="0"`
    /// as Excel spells a fresh field, the built inventory inside — and
    /// `<cacheFields count>` set to the effective count where the
    /// wrapper spells one.
    fn appendSchemaSplices(
        arena: Allocator,
        splices: *std.ArrayListUnmanaged(edit.Splice),
        def: *const pivot_xml.CacheDefinition,
        se: edit.SchemaEdit,
        fields: []const Field,
        p: []const u8,
    ) RebuildError!void {
        switch (se) {
            .remove => |k| {
                const f = def.fields[k];
                if (f.span.start == 0 and f.span.end == 0) return error.MalformedPivotXml;
                try splices.append(arena, .{ .span = f.span, .text = "" });
            },
            .insert => |ins| {
                if (ins.at >= def.fields.len) return error.MalformedPivotXml;
                const anchor = def.fields[ins.at].span.start;
                // A caller-assembled field without a span would turn
                // the insertion into a prepend before the declaration
                // — answered, as the remove arm answers (in-house
                // review S7C-MUT-2).
                if (anchor == 0) return error.MalformedPivotXml;
                var out: std.ArrayListUnmanaged(u8) = .empty;
                try out.append(arena, '<');
                try out.appendSlice(arena, p);
                try out.appendSlice(arena, "cacheField name=\"");
                try out.appendSlice(arena, try formula.decode.encodeAuthoredString(arena, ins.name));
                try out.appendSlice(arena, "\" numFmtId=\"0\">");
                try out.appendSlice(arena, fields[ins.at].xml);
                try out.appendSlice(arena, "</");
                try out.appendSlice(arena, p);
                try out.appendSlice(arena, "cacheField>");
                try splices.append(arena, .{ .span = .{ .start = anchor, .end = anchor }, .text = out.items });
            },
        }
        if (def.fields_count_span) |span| {
            const n: usize = switch (se) {
                .remove => def.fields.len - 1,
                .insert => def.fields.len + 1,
            };
            try splices.append(arena, .{ .span = span, .text = try std.fmt.allocPrint(arena, "{d}", .{n}) });
        }
    }

    /// The definition shapes this slice rebuilds — everything else
    /// refuses. A worksheet source with plain database fields, each
    /// with an inventory of simple string / number / blank items and
    /// no date; no calculated or group field, no OLAP element.
    pub fn checkShape(def: *const pivot_xml.CacheDefinition) RebuildError!void {
        if (def.source.type != .worksheet or def.source.has_consolidation or def.source.range_sets.len != 0) return error.PivotShapeUnsupported;
        if (def.has_other_children) return error.PivotShapeUnsupported;
        if (def.fields_count_attr) |n| {
            if (n != def.fields.len) return error.MalformedPivotXml;
        }
        for (def.fields) |f| {
            if (f.formula != null or !f.database_field or f.has_other_children) return error.PivotShapeUnsupported;
            const si = f.shared_items orelse return error.PivotShapeUnsupported;
            if (si.has_other_children) return error.PivotShapeUnsupported;
            if (si.contains_date or si.min_date != null or si.max_date != null) return error.PivotShapeUnsupported;
            if (si.count) |n| {
                if (n != si.items.len) return error.MalformedPivotXml;
            }
            for (si.items) |it| {
                if (!it.simple) return error.PivotShapeUnsupported;
                switch (it.kind) {
                    .s, .n, .m => {},
                    .b, .d, .e, .other => return error.PivotShapeUnsupported,
                }
            }
        }
    }

    /// The whole ` name="value"` of a root attribute, leading
    /// whitespace included — what removing it takes out. Null when the
    /// root does not carry it, or when the bytes around the value are
    /// not the `name = "value"` the parser read it from.
    fn rootAttrSpan(src: []const u8, def: *const pivot_xml.CacheDefinition, name: []const u8) ?pivot_xml.Span {
        const value = def.rootAttrValueSpan(src, name) orelse return null;
        const lo = def.root_attrs.start;
        if (value.start <= lo + 1 or value.end >= src.len) return null;
        var i = value.start - 1;
        if (src[i] != '"' and src[i] != '\'') return null;
        i -= 1;
        while (i > lo and isXmlSpace(src[i])) i -= 1;
        if (src[i] != '=') return null;
        i -= 1;
        while (i > lo and isXmlSpace(src[i])) i -= 1;
        if (i + 1 < lo + name.len) return null;
        const name_start = i + 1 - name.len;
        if (!std.mem.eql(u8, src[name_start .. i + 1], name)) return null;
        var start = name_start;
        while (start > lo and isXmlSpace(src[start - 1])) start -= 1;
        return .{ .start = start, .end = value.end + 1 };
    }

    fn isXmlSpace(c: u8) bool {
        return c == ' ' or c == '\t' or c == '\n' or c == '\r';
    }

    /// `prefix:` for a prefixed root, empty otherwise.
    fn qualified(arena: Allocator, prefix: []const u8) RebuildError![]const u8 {
        if (prefix.len == 0) return "";
        return std.mem.concat(arena, u8, &.{ prefix, ":" });
    }

    /// One field's rebuilt inventory and how its records spell it.
    pub const Field = struct {
        /// The rebuilt `<sharedItems …>` element.
        xml: []const u8,
        /// Records spell this field as `<x v>` into the inventory;
        /// otherwise inline (`<n>` / `<m>`).
        indexed: bool,
        /// Per data row, the item index — meaningful when `indexed`.
        index_of_row: []const u32,
        /// The final inventory: the items as written, in their order,
        /// then the appended ones. Empty for an inline field.
        items: []const Item,

        pub const Item = struct {
            /// As written, for an item the inventory already had.
            raw: ?[]const u8,
            /// The value: a number's `f64` and lexical (`<n>`), a
            /// string decoded (`<s>`) — retained and appended alike, so
            /// the element can describe the items it holds.
            num: f64 = 0,
            lex: []const u8 = "",
            text: []const u8 = "",
            kind: pivot_xml.SharedItems.Item.Kind,
        };

        /// What a `sharedItems` says of its contents: the `contains*`
        /// flags, the integer and long-text hints, the extrema. Fed
        /// the items the element will hold for an enumerated field —
        /// retained and appended, since a flag that ignored a retained
        /// `<s>` would deny the child it sits beside (Codex #205 r3
        /// REL-303) — and the rows for an inline one.
        const Description = struct {
            has_string: bool = false,
            has_number: bool = false,
            has_blank: bool = false,
            all_int: bool = true,
            long_text: bool = false,
            min: f64 = 0,
            max: f64 = 0,
            min_lex: []const u8 = "",
            max_lex: []const u8 = "",

            fn blank(self: *Description) void {
                self.has_blank = true;
            }

            fn number(self: *Description, x: f64, lex: []const u8) void {
                if (!self.has_number or x < self.min) {
                    self.min = x;
                    self.min_lex = lex;
                }
                if (!self.has_number or x > self.max) {
                    self.max = x;
                    self.max_lex = lex;
                }
                self.has_number = true;
                if (!isInteger(x)) self.all_int = false;
            }

            fn string(self: *Description, s: []const u8) RebuildError!void {
                self.has_string = true;
                const cps = std.unicode.utf8CountCodepoints(s) catch return error.PivotShapeUnsupported;
                if (cps > 255) self.long_text = true;
            }
        };

        /// `spelt_indexed` is how the records as written spelt this
        /// field — `<x>` into the inventory, or a value inline — when
        /// there was a record to read it from; without one, an
        /// inventory with items is the enumerated one. An explicit
        /// `count` alone says nothing (Codex #205 r10 REL-1001).
        fn build(arena: Allocator, si: pivot_xml.SharedItems, rows: []const Row, k: usize, p: []const u8, spelt_indexed: ?bool) RebuildError!Field {
            var items: std.ArrayListUnmanaged(Item) = .empty;
            var by_string: std.StringHashMapUnmanaged(u32) = .empty;
            var by_number: std.AutoHashMapUnmanaged(u64, u32) = .empty;
            var blank_at: ?u32 = null;

            // The inventory as written, in its order.
            for (si.items) |it| {
                const idx: u32 = @intCast(items.items.len);
                var item: Item = .{ .raw = it.raw, .kind = it.kind };
                switch (it.kind) {
                    .m => {
                        if (blank_at == null) blank_at = idx;
                    },
                    .n => {
                        // The value as written may spell itself with
                        // references (`&#49;`); it is matched by what
                        // it is (Codex #205 r8 REL-802).
                        const lex = decodeLexical(arena, it.v orelse return error.MalformedPivotXml) catch |e| switch (e) {
                            error.OutOfMemory => return error.OutOfMemory,
                            else => return error.MalformedPivotXml,
                        };
                        const x = parseNumber(lex) orelse return error.MalformedPivotXml;
                        const gop = try by_number.getOrPut(arena, numberKey(x));
                        if (!gop.found_existing) gop.value_ptr.* = idx;
                        item.num = x;
                        item.lex = lex;
                    },
                    .s => {
                        const text = try decodeItem(arena, it.v orelse return error.MalformedPivotXml);
                        const gop = try by_string.getOrPut(arena, try foldOrRefuse(arena, text));
                        if (!gop.found_existing) gop.value_ptr.* = idx;
                        item.text = text;
                    },
                    else => unreachable, // checkShape
                }
                try items.append(arena, item);
            }
            // Records spelt inline beside an inventory that holds items
            // are not one shape: an inline rebuild would emit the
            // attribute-only element and drop every item a consumer
            // indexes (Codex #205 r11 REL-1101).
            if (spelt_indexed == false and si.items.len > 0) return error.PivotShapeUnsupported;
            const was_indexed = spelt_indexed orelse (si.items.len > 0);

            // Pass one: what the column holds — whether the records can
            // stay inline.
            var rows_string = false;
            var rows_number = false;
            for (rows) |r| switch (r[k]) {
                .blank => {},
                .number => |lex| {
                    if (parseNumber(lex) == null) return error.PivotShapeUnsupported;
                    rows_number = true;
                },
                .string => rows_string = true,
            };
            const indexed = was_indexed or rows_string or !rows_number;

            // Pass two: index every row, appending what the inventory
            // lacks in first-appearance order.
            const index_of_row = try arena.alloc(u32, rows.len);
            @memset(index_of_row, 0);
            if (indexed) {
                for (rows, 0..) |r, i| {
                    const next: u32 = @intCast(items.items.len);
                    switch (r[k]) {
                        .blank => {
                            if (blank_at) |at| {
                                index_of_row[i] = at;
                            } else {
                                blank_at = next;
                                index_of_row[i] = next;
                                try items.append(arena, .{ .raw = null, .text = "", .kind = .m });
                            }
                        },
                        .number => |lex| {
                            const x = parseNumber(lex).?;
                            const gop = try by_number.getOrPut(arena, numberKey(x));
                            if (!gop.found_existing) {
                                gop.value_ptr.* = next;
                                try items.append(arena, .{ .raw = null, .num = x, .lex = lex, .kind = .n });
                            }
                            index_of_row[i] = gop.value_ptr.*;
                        },
                        .string => |s| {
                            const gop = try by_string.getOrPut(arena, try foldOrRefuse(arena, s));
                            if (!gop.found_existing) {
                                gop.value_ptr.* = next;
                                try items.append(arena, .{ .raw = null, .text = s, .kind = .s });
                            }
                            index_of_row[i] = gop.value_ptr.*;
                        },
                    }
                }
            }
            if (items.items.len > std.math.maxInt(u32)) return error.PivotShapeUnsupported;

            // What the element says of itself — from its items when it
            // enumerates them, from the rows when the records are inline.
            var d: Description = .{};
            if (indexed) {
                for (items.items) |it| switch (it.kind) {
                    .m => d.blank(),
                    .n => d.number(it.num, it.lex),
                    .s => try d.string(it.text),
                    else => unreachable,
                };
            } else {
                for (rows) |r| switch (r[k]) {
                    .blank => d.blank(),
                    .number => |lex| d.number(parseNumber(lex).?, lex),
                    .string => unreachable, // a string enumerates the field
                };
            }

            // The element, attributes in the schema's order, defaults
            // omitted — the spelling Excel writes.
            var out: std.ArrayListUnmanaged(u8) = .empty;
            try out.append(arena, '<');
            try out.appendSlice(arena, p);
            try out.appendSlice(arena, "sharedItems");
            if (d.has_number and !d.has_string and !d.has_blank) try out.appendSlice(arena, " containsSemiMixedTypes=\"0\"");
            if (!d.has_number and !d.has_string) try out.appendSlice(arena, " containsNonDate=\"0\"");
            if (!d.has_string) try out.appendSlice(arena, " containsString=\"0\"");
            if (d.has_blank) try out.appendSlice(arena, " containsBlank=\"1\"");
            if (d.has_string and d.has_number) try out.appendSlice(arena, " containsMixedTypes=\"1\"");
            if (d.has_number) {
                try out.appendSlice(arena, " containsNumber=\"1\"");
                if (d.all_int) try out.appendSlice(arena, " containsInteger=\"1\"");
                try out.appendSlice(arena, " minValue=\"");
                try out.appendSlice(arena, d.min_lex);
                try out.appendSlice(arena, "\" maxValue=\"");
                try out.appendSlice(arena, d.max_lex);
                try out.append(arena, '"');
            }
            if (indexed) {
                try out.appendSlice(arena, " count=\"");
                try out.appendSlice(arena, try std.fmt.allocPrint(arena, "{d}", .{items.items.len}));
                try out.append(arena, '"');
            }
            if (d.long_text) try out.appendSlice(arena, " longText=\"1\"");
            if (indexed and items.items.len > 0) {
                try out.append(arena, '>');
                for (items.items) |it| {
                    if (it.raw) |raw| {
                        try out.appendSlice(arena, raw);
                        continue;
                    }
                    try out.append(arena, '<');
                    try out.appendSlice(arena, p);
                    switch (it.kind) {
                        .m => try out.appendSlice(arena, "m/>"),
                        .n => {
                            try out.appendSlice(arena, "n v=\"");
                            try out.appendSlice(arena, it.lex);
                            try out.appendSlice(arena, "\"/>");
                        },
                        .s => {
                            try out.appendSlice(arena, "s v=\"");
                            try out.appendSlice(arena, try formula.decode.encodeAuthoredString(arena, it.text));
                            try out.appendSlice(arena, "\"/>");
                        },
                        else => unreachable,
                    }
                }
                try out.appendSlice(arena, "</");
                try out.appendSlice(arena, p);
                try out.appendSlice(arena, "sharedItems>");
            } else {
                try out.appendSlice(arena, "/>");
            }
            return .{ .xml = out.items, .indexed = indexed, .index_of_row = index_of_row, .items = items.items };
        }
    };

    /// The records part as written, checked to be a shape the rebuild
    /// carries — one `<r>` per record, each holding exactly one value
    /// element per field, childless, under the part's prefix, with no
    /// qualified attribute, nothing else anywhere (Codex #205 r3
    /// REL-304, r4 REL-403, r6 REL-602) — and read for how each field
    /// spelt its values: `<x>` into the inventory (true), inline
    /// (false), or unknown where no record spoke (null). A field spelt
    /// both ways is not one shape (r10 REL-1001).
    fn inspectRecords(arena: Allocator, xml: []const u8, field_count: usize) RebuildError![]?bool {
        const spelt = try arena.alloc(?bool, field_count);
        @memset(spelt, null);
        const root = pivot_xml.scanRoot(xml, "pivotCacheRecords") catch |e| return mapRecordsParse(e);
        var kids = pivot_xml.Children.init(xml, root.hit, root.body_end, root.prefix, root.env);
        while (kids.next() catch |e| return mapRecordsParse(e)) |k| {
            if (!std.mem.eql(u8, k.local, "r")) return error.PivotShapeUnsupported;
            if (pivot_xml.hasAnyAttr(k.attrs(xml))) return error.PivotShapeUnsupported;
            var vals = pivot_xml.Children.init(xml, k.hit, k.end, root.prefix, k.env);
            var j: usize = 0;
            while (vals.next() catch |e| return mapRecordsParse(e)) |v| {
                if (v.local.len != 1) return error.PivotShapeUnsupported;
                switch (v.local[0]) {
                    'x', 'n', 's', 'm', 'b', 'd', 'e' => {},
                    else => return error.PivotShapeUnsupported,
                }
                if (pivot_xml.hasQualifiedAttr(v.attrs(xml))) return error.PivotShapeUnsupported;
                if (!pivot_xml.isBlank(xml[v.hit.after_tag_close..v.end])) return error.PivotShapeUnsupported;
                if (j >= field_count) return error.PivotShapeUnsupported;
                const is_x = v.local[0] == 'x';
                if (spelt[j]) |was| {
                    if (was != is_x) return error.PivotShapeUnsupported;
                } else spelt[j] = is_x;
                j += 1;
            }
            if (j != field_count) return error.PivotShapeUnsupported;
            if (vals.skipped > 0 or vals.other) return error.PivotShapeUnsupported;
        }
        if (kids.skipped > 0 or kids.other) return error.PivotShapeUnsupported;
        return spelt;
    }

    /// The records part read back as rows — what `rebuild` would read
    /// from the source if the snapshot were current. `<x v>` resolves
    /// through the inventory as written (`<s>` its text, `<n>` its
    /// value, `<m>` a blank); an inline value stands as spelt. Refuses
    /// what `inspectRecords` refuses, and a value kind the cache slice
    /// does not carry (`<b>`, `<d>`, `<e>`).
    pub fn rowsFromRecords(arena: Allocator, cache: *const PivotCache, xml: []const u8) RebuildError![]const Row {
        const def = &cache.definition;
        try checkShape(def);
        _ = try inspectRecords(arena, xml, def.fields.len);
        const root = pivot_xml.scanRoot(xml, "pivotCacheRecords") catch |e| return mapRecordsParse(e);
        var out: std.ArrayListUnmanaged(Row) = .empty;
        var kids = pivot_xml.Children.init(xml, root.hit, root.body_end, root.prefix, root.env);
        while (kids.next() catch |e| return mapRecordsParse(e)) |k| {
            const row = try arena.alloc(Value, def.fields.len);
            var vals = pivot_xml.Children.init(xml, k.hit, k.end, root.prefix, k.env);
            var j: usize = 0;
            while (vals.next() catch |e| return mapRecordsParse(e)) |v| : (j += 1) {
                const attrs = v.attrs(xml);
                switch (v.local[0]) {
                    'x' => {
                        const idx = (pivot_xml.u32Attr(attrs, "v") catch |e| return mapRecordsParse(e)) orelse 0;
                        const si = def.fields[j].shared_items.?;
                        if (idx >= si.items.len) return error.MalformedPivotXml;
                        const it = si.items[idx];
                        row[j] = switch (it.kind) {
                            .s => .{ .string = try decodeItem(arena, it.v orelse return error.MalformedPivotXml) },
                            .n => .{ .number = try lexicalOf(arena, it.v orelse return error.MalformedPivotXml) },
                            .m => .blank,
                            else => unreachable, // checkShape
                        };
                    },
                    'n' => row[j] = .{ .number = try lexicalOf(arena, wbxml.getAttr(attrs, "v") orelse return error.MalformedPivotXml) },
                    's' => row[j] = .{ .string = try decodeItem(arena, wbxml.getAttr(attrs, "v") orelse return error.MalformedPivotXml) },
                    'm' => row[j] = .blank,
                    else => return error.PivotShapeUnsupported,
                }
            }
            try out.append(arena, row);
        }
        return out.items;
    }

    /// A numeric attribute's value as a number's lexical.
    fn lexicalOf(arena: Allocator, raw: []const u8) RebuildError![]const u8 {
        const lex = decodeLexical(arena, raw) catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            else => return error.MalformedPivotXml,
        };
        if (parseNumber(lex) == null) return error.MalformedPivotXml;
        return lex;
    }

    /// The records part: the root tag as written with `count` set,
    /// one `<r>` per row — `<x v>` for an indexed field, `<n v>` /
    /// `<m/>` inline for the rest — the close, and whatever followed
    /// the root. Runs after `inspectRecords` read the part.
    fn renderRecords(arena: Allocator, xml: []const u8, fields: []const Field, rows: []const Row) RebuildError![]u8 {
        const root = pivot_xml.scanRoot(xml, "pivotCacheRecords") catch |e| return mapRecordsParse(e);
        const p = try qualified(arena, root.prefix);
        var out: std.ArrayListUnmanaged(u8) = .empty;
        try out.appendSlice(arena, xml[0..root.hit.attrs_start]);
        const attrs = xml[root.hit.attrs_start..root.hit.attrs_end];
        const count_text = try std.fmt.allocPrint(arena, "{d}", .{rows.len});
        if (wbxml.getAttr(attrs, "count")) |v| {
            const off = @intFromPtr(v.ptr) - @intFromPtr(attrs.ptr);
            try out.appendSlice(arena, attrs[0..off]);
            try out.appendSlice(arena, count_text);
            try out.appendSlice(arena, attrs[off + v.len ..]);
        } else {
            try out.appendSlice(arena, attrs);
            try out.appendSlice(arena, " count=\"");
            try out.appendSlice(arena, count_text);
            try out.append(arena, '"');
        }
        try out.append(arena, '>');
        for (rows, 0..) |r, i| {
            try out.append(arena, '<');
            try out.appendSlice(arena, p);
            try out.appendSlice(arena, "r>");
            for (fields, 0..) |f, k| {
                try out.append(arena, '<');
                try out.appendSlice(arena, p);
                if (f.indexed) {
                    try out.appendSlice(arena, "x v=\"");
                    try out.appendSlice(arena, try std.fmt.allocPrint(arena, "{d}", .{f.index_of_row[i]}));
                    try out.appendSlice(arena, "\"/>");
                    continue;
                }
                switch (r[k]) {
                    .blank => try out.appendSlice(arena, "m/>"),
                    .number => |lex| {
                        try out.appendSlice(arena, "n v=\"");
                        try out.appendSlice(arena, lex);
                        try out.appendSlice(arena, "\"/>");
                    },
                    // A string forces the field indexed.
                    .string => unreachable,
                }
            }
            try out.appendSlice(arena, "</");
            try out.appendSlice(arena, p);
            try out.appendSlice(arena, "r>");
        }
        try out.appendSlice(arena, "</");
        try out.appendSlice(arena, p);
        try out.appendSlice(arena, "pivotCacheRecords>");
        // What followed the root — a comment, a processing instruction,
        // whitespace — as written (Codex #205 r9 REL-903).
        try out.appendSlice(arena, xml[root.after..]);
        return out.items;
    }

    /// An xsd:double lexical — `[+-]? (D+ ('.' D*)? | '.' D+)
    /// ([eE] [+-]? D+)?` — as a finite value; null otherwise. Zig's
    /// parser reads more (`0x1p0`, `1_0`, `inf`), none of which a cache
    /// may spell, so the grammar is checked first (Codex #205 r9
    /// REL-901).
    // ─── S7b-5: the consumers — items, layout, output cells ─────────
    //
    // The second slice lays every consumer of a rebuilt cache out
    // again: the row field's `<items>` (the written order kept, a new
    // inventory item appended after it — `sortType="manual"`, the
    // schema's default — or every item re-sorted under `ascending`),
    // `<rowItems>` (one per item with a record, then the grand total),
    // `location@ref` (the rectangle the rows now fill) and the cells
    // of the host rectangle: the header row (the row-labels caption
    // the host already spells, then each data field's caption), one
    // row per item with its aggregates, the grand total. What it lays
    // out is the one report form the corpus and the fixture carry —
    // compact, one row field, the values axis across (or one data
    // field), no page, column or hidden item, plain aggregates (sum,
    // count, countNums, average, min, max, product) — and every other
    // form refuses (`PivotShapeUnsupported`), as the cache slice does.
    //
    // Two of Excel's spellings the layout takes on faith, oracle
    // pending: a group with no value to aggregate (an inserted blank
    // row's `(blank)` item) is an empty cell, not 0; a manual-sort
    // field appends its new item after the written ones.

    /// The captions the host rectangle already spells — Excel writes
    /// them in its UI language, so a rebuild reuses what is there
    /// rather than the English default. Null when the host has no
    /// such cell (a first layout, a caption that is not text).
    pub const Captions = struct {
        row_labels: ?[]const u8 = null,
        grand_total: ?[]const u8 = null,
        blank: ?[]const u8 = null,
    };

    /// One cell of the laid-out rectangle, in the coordinates of the
    /// definition's `location@ref`.
    pub const OutCell = struct {
        row: u32,
        col: u32,
        value: union(enum) {
            blank,
            /// An xsd:double lexical — an aggregate spelt shortest
            /// round-trip, or a numeric item's spelling as inventoried.
            number: []const u8,
            /// Decoded text.
            string: []const u8,
        },
        kind: RowKind,
    };

    pub const RowKind = enum { header, item, grand };

    /// A consumer re-laid: the part and the cells.
    pub const Layout = struct {
        /// The part with the row field's `<pivotField>`, `<rowItems>`
        /// and `location@ref` regenerated; every other byte as given.
        table_xml: []u8,
        /// The rectangle the cells fill.
        rect: edit.Rect,
        /// The rectangle the part named before — the cells to clear
        /// where the new one no longer covers them.
        old_rect: edit.Rect,
        /// Row-major, header first.
        cells: []const OutCell,
        /// The old rectangle's row kinds, top to bottom, for a caller
        /// carrying styles from the old cells to the new by kind.
        old_kinds: []const RowKind,
    };

    /// Lay one consumer of `cache` out over `rb`, its rebuilt cache.
    /// `table_xml` is the consumer part as it is to be laid out — for
    /// a host on the edited sheet, after the S7a `location@ref` move,
    /// so the rectangle is spelt in post-edit coordinates. Refuses a
    /// form the slice does not lay out; a part that disagrees with
    /// itself (an item naming no inventory entry, a `location` a row
    /// short of its `rowItems`) is malformed.
    pub fn layout(arena: Allocator, table_xml: []const u8, cache: *const PivotCache, rb: *const Rebuild, captions: Captions) RebuildError!Layout {
        const def = pivot_xml.parseTableDefinition(arena, table_xml) catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            error.MalformedXml => return error.MalformedPivotXml,
        };
        try validateRebuild(arena, rb);
        try checkTableShape(arena, &def, cache, rb);

        const rf: u32 = def.row_fields[0].field;
        const pf = def.fields[rf];
        const field = rb.fields[rf];

        // The item list: as written, then what the inventory gained,
        // then re-sorted when the field asks; `<item t="default"/>`
        // stays last where the part had it.
        var order: std.ArrayListUnmanaged(u32) = .empty;
        const referenced = try arena.alloc(bool, field.items.len);
        @memset(referenced, false);
        var default_item = false;
        for (pf.items) |it| {
            if (it.t) |t| {
                // One subtotal item, last; an item that is both a
                // cache item and a derived one is neither.
                if (!attrIs(t, "default") or default_item or it.x != null) return error.PivotShapeUnsupported;
                default_item = true;
                continue;
            }
            if (default_item) return error.PivotShapeUnsupported;
            const x = it.x orelse return error.MalformedPivotXml;
            if (x >= field.items.len or referenced[x]) return error.MalformedPivotXml;
            referenced[x] = true;
            try order.append(arena, x);
        }
        for (field.items, 0..) |_, k| {
            if (!referenced[k]) try order.append(arena, @intCast(k));
        }
        if (pf.sort_type == .ascending) try sortItems(arena, field, order.items);

        // Records per item.
        const counts = try arena.alloc(u32, field.items.len);
        @memset(counts, 0);
        for (field.index_of_row) |k| counts[k] += 1;

        // The rows shown: items with a record, in item-list order
        // (`showAll="0"`, the one setting the slice lays out).
        var shown: std.ArrayListUnmanaged(u32) = .empty; // positions in `order`
        for (order.items, 0..) |k, pos| {
            if (counts[k] > 0) try shown.append(arena, @intCast(pos));
        }

        // The rectangle.
        const old = edit.footprintOf(arena, def) catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            else => return error.MalformedPivotXml,
        };
        const old_rect = old.rect;
        const old_items = def.row_items.?;
        const old_kinds = try arena.alloc(RowKind, old_items.items.len + 1);
        old_kinds[0] = .header;
        for (old_items.items, 1..) |it, i| old_kinds[i] = if (it.t == null) .item else .grand;
        const old_height: usize = old_rect.br_row - old_rect.tl_row + 1;
        if (old_height != old_kinds.len) return error.MalformedPivotXml;
        const width: usize = 1 + def.data_fields.len;
        if (old_rect.br_col - old_rect.tl_col + 1 != width) return error.MalformedPivotXml;
        const grand = def.row_grand_totals;
        const height: usize = 1 + shown.items.len + @intFromBool(grand);
        if (height > zlsx.max_row or @as(u64, old_rect.tl_row) + height - 1 > zlsx.max_row) return error.PivotShapeUnsupported;
        const rect: edit.Rect = .{
            .tl_col = old_rect.tl_col,
            .tl_row = old_rect.tl_row,
            .br_col = old_rect.br_col,
            .br_row = @intCast(@as(u64, old_rect.tl_row) + height - 1),
        };

        // The cells.
        var cells: std.ArrayListUnmanaged(OutCell) = .empty;
        // The header's caption: the definition's own where it spells
        // one (`rowHeaderCaption`), else the host's, else the default.
        const row_labels: []const u8 = if (def.row_header_caption) |c| try decodeItem(arena, c) else captions.row_labels orelse "Row Labels";
        try cells.append(arena, .{ .row = rect.tl_row, .col = rect.tl_col, .value = .{ .string = row_labels }, .kind = .header });
        for (def.data_fields, 0..) |df, j| {
            const caption = if (df.name) |n| try decodeItem(arena, n) else try defaultCaption(arena, df.subtotal, cache.field_names[df.fld]);
            try cells.append(arena, .{ .row = rect.tl_row, .col = rect.tl_col + 1 + @as(u32, @intCast(j)), .value = .{ .string = caption }, .kind = .header });
        }
        const groups = try arena.alloc(std.ArrayListUnmanaged(u32), field.items.len);
        @memset(groups, .empty);
        for (field.index_of_row, 0..) |k, i| try groups[k].append(arena, @intCast(i));
        const folded = try arena.alloc(Agg, def.data_fields.len);
        @memset(folded, .{});
        var r: u32 = rect.tl_row + 1;
        for (shown.items) |pos| {
            const k = order.items[pos];
            const it = field.items[k];
            const label: OutCell = .{
                .row = r,
                .col = rect.tl_col,
                .kind = .item,
                .value = switch (it.kind) {
                    .n => .{ .number = it.lex },
                    .s => .{ .string = it.text },
                    .m => .{ .string = captions.blank orelse "(blank)" },
                    else => unreachable, // checkShape
                },
            };
            try cells.append(arena, label);
            for (def.data_fields, 0..) |df, j| {
                const agg = aggregateGroup(rb, df.fld, groups[k].items);
                folded[j].fold(agg);
                try cells.append(arena, .{ .row = r, .col = rect.tl_col + 1 + @as(u32, @intCast(j)), .value = try renderAgg(arena, df, agg), .kind = .item });
            }
            r += 1;
        }
        if (grand) {
            const caption = if (def.grand_total_caption) |c| try decodeItem(arena, c) else captions.grand_total orelse "Grand Total";
            try cells.append(arena, .{ .row = r, .col = rect.tl_col, .value = .{ .string = caption }, .kind = .grand });
            // Every record once, in record order — the rows of the
            // shown items, which under `showAll="0"` is every record.
            var all: std.ArrayListUnmanaged(u32) = .empty;
            for (rb.fields[rf].index_of_row, 0..) |k, i| {
                if (counts[k] > 0) try all.append(arena, @intCast(i));
            }
            for (def.data_fields, 0..) |df, j| {
                var total = aggregateGroup(rb, df.fld, all.items);
                if (df.subtotal == .sum) total.sum = folded[j].sum;
                if (df.subtotal == .product) total.prod = folded[j].prod;
                try cells.append(arena, .{ .row = r, .col = rect.tl_col + 1 + @as(u32, @intCast(j)), .value = try renderAgg(arena, df, total), .kind = .grand });
            }
        }

        // The part: three splices, in span order.
        const p = try qualified(arena, def.prefix);
        var splices: [3]edit.Splice = undefined;
        splices[0] = .{ .span = pf.span, .text = try renderPivotField(arena, p, pf, order.items, counts, default_item) };
        splices[1] = .{ .span = old_items.span, .text = try renderRowItems(arena, p, shown.items, grand) };
        var ref_buf: [Bounds.format_buf_len]u8 = undefined;
        const new_ref = edit.formatRect(&ref_buf, rect) catch return error.PivotShapeUnsupported;
        splices[2] = .{ .span = def.location.ref_span, .text = try arena.dupe(u8, new_ref) };
        const table = edit.spliceAll(arena, table_xml, &splices) catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            else => return error.MalformedPivotXml,
        };
        return .{ .table_xml = table, .rect = rect, .old_rect = old_rect, .cells = cells.items, .old_kinds = old_kinds };
    }

    /// A `Rebuild` is public: one `rebuild` made holds these by
    /// construction, one a caller assembled is checked before anything
    /// indexes by it (Codex #206 r7 REL-701) — the rows count the
    /// records and are one value per field; an indexed field names
    /// one inventory item per row, of a kind the slice lays out; an
    /// inline field's rows are numbers or blanks.
    fn validateRebuild(arena: Allocator, rb: *const Rebuild) RebuildError!void {
        // `rebuild` never makes an empty one (Codex #206 r14 REL-1401).
        if (rb.rows.len == 0 or rb.rows.len != rb.record_count) return error.MalformedPivotXml;
        for (rb.rows) |r| if (r.len != rb.fields.len) return error.MalformedPivotXml;
        for (rb.fields, 0..) |f, k| {
            if (f.index_of_row.len != rb.rows.len) return error.MalformedPivotXml;
            for (f.items) |it| switch (it.kind) {
                .s, .m => {},
                // A number's spelling and its value are one thing
                // (Codex #206 r8 REL-802).
                .n => {
                    const x = parseNumber(it.lex) orelse return error.MalformedPivotXml;
                    if (!std.math.isFinite(it.num) or x != it.num) return error.MalformedPivotXml;
                },
                else => return error.MalformedPivotXml,
            };
            if (f.indexed) {
                // Each row's value is the item its index names (Codex
                // #206 r10 REL-1003).
                for (f.index_of_row, rb.rows) |i, r| {
                    if (i >= f.items.len) return error.MalformedPivotXml;
                    const it = f.items[i];
                    switch (r[k]) {
                        .blank => if (it.kind != .m) return error.MalformedPivotXml,
                        .number => |lex| {
                            const x = parseNumber(lex) orelse return error.MalformedPivotXml;
                            if (it.kind != .n or numberKey(x) != numberKey(it.num)) return error.MalformedPivotXml;
                        },
                        .string => |txt| {
                            if (it.kind != .s) return error.MalformedPivotXml;
                            // An allocation failure is its own error,
                            // never "malformed" (Codex #206 r11
                            // REL-1101).
                            const ka = try foldForValidation(arena, txt);
                            const kb = try foldForValidation(arena, it.text);
                            if (!std.mem.eql(u8, ka, kb)) return error.MalformedPivotXml;
                        },
                    }
                }
            } else {
                if (f.items.len != 0) return error.MalformedPivotXml;
                for (rb.rows) |r| switch (r[k]) {
                    .blank => {},
                    .number => |lex| if (parseNumber(lex) == null) return error.MalformedPivotXml,
                    .string => return error.MalformedPivotXml,
                };
            }
        }
    }

    fn foldForValidation(arena: Allocator, s: []const u8) RebuildError![]const u8 {
        const folded = fold(arena, s) catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            else => return error.MalformedPivotXml,
        };
        return folded orelse error.MalformedPivotXml;
    }

    /// The consumer forms this slice lays out; everything else refuses
    /// before a cell is computed. Malformed where the part disagrees
    /// with itself or with the cache it reads.
    fn checkTableShape(arena: Allocator, def: *const pivot_xml.TableDefinition, cache: *const PivotCache, rb: *const Rebuild) RebuildError!void {
        if (def.has_other_children or def.chart_formats == .other or def.axes_other) return error.PivotShapeUnsupported;
        if (def.axes_count_mismatch) return error.MalformedPivotXml;
        if (def.page_fields.len != 0 or def.data_on_rows or !def.compact) return error.PivotShapeUnsupported;
        try checkRootDisplayAttrs(def.root_attrs);
        // The one form has its grand total; a layout without one, or
        // with the column totals off, is not oracled (Codex #206 r8
        // REL-801).
        if (!def.row_grand_totals or !def.col_grand_totals) return error.PivotShapeUnsupported;
        if (def.fields.len != cache.definition.fields.len or rb.fields.len != def.fields.len) return error.MalformedPivotXml;
        // A cache is public too: its decoded spines are parallel to
        // its definition's fields by construction, and checked here
        // before anything indexes by them (Codex #206 r12 REL-1203).
        if (cache.field_names.len != def.fields.len or cache.field_formulas.len != def.fields.len) return error.MalformedPivotXml;
        if (def.row_fields.len != 1 or def.row_fields[0] != .field) return error.PivotShapeUnsupported;
        const rf = def.row_fields[0].field;
        if (rf >= def.fields.len) return error.MalformedPivotXml;
        if (def.data_fields.len == 0) return error.PivotShapeUnsupported;
        // The values axis: across, or nowhere for a single data field.
        switch (def.col_fields.len) {
            0 => if (def.data_fields.len != 1) return error.PivotShapeUnsupported,
            1 => if (def.col_fields[0] != .values) return error.PivotShapeUnsupported,
            else => return error.PivotShapeUnsupported,
        }
        // `<colItems>` as Excel lays the values axis out: one `<i>`
        // per data field, `<x v>` its index, `i` its index past the
        // first — left as written, since no row edit changes it.
        const ci = def.col_items orelse return error.PivotShapeUnsupported;
        if (ci.items.len != def.data_fields.len or ci.other_attrs) return error.PivotShapeUnsupported;
        if (ci.count) |n| if (n != ci.items.len) return error.MalformedPivotXml;
        for (ci.items, 0..) |it, j| {
            if (it.t != null or it.r != null or it.other_attrs or it.has_other_children) return error.PivotShapeUnsupported;
            if (def.col_fields.len == 0) {
                if (it.xs.len != 0 or it.i != null) return error.PivotShapeUnsupported;
            } else {
                if (it.xs.len != 1 or it.xs[0] != j) return error.PivotShapeUnsupported;
                if ((it.i orelse 0) != j) return error.PivotShapeUnsupported;
            }
        }
        // `<rowItems>` as written: data items then the grand total.
        const ri = def.row_items orelse return error.PivotShapeUnsupported;
        if (ri.other_attrs) return error.PivotShapeUnsupported;
        if (ri.count) |n| if (n != ri.items.len) return error.MalformedPivotXml;
        var seen_grand = false;
        for (ri.items) |it| {
            if (it.other_attrs or it.has_other_children or it.r != null or it.i != null) return error.PivotShapeUnsupported;
            if (it.t) |t| {
                if (!attrIs(t, "grand") or seen_grand) return error.PivotShapeUnsupported;
                seen_grand = true;
                if (it.xs.len != 1) return error.PivotShapeUnsupported;
            } else {
                if (seen_grand or it.xs.len != 1) return error.PivotShapeUnsupported;
            }
        }
        if (seen_grand != def.row_grand_totals) return error.PivotShapeUnsupported;
        // The location's counts, as the form has them.
        // The schema requires all three offsets; the form has its
        // data at (1, 1) and its header row at 0 or 1 — the two the
        // corpus and the fixture spell (Codex #206 r15 REL-1505).
        const loc = def.location;
        const fhr = loc.first_header_row orelse return error.MalformedPivotXml;
        const fdr = loc.first_data_row orelse return error.MalformedPivotXml;
        const fdc = loc.first_data_col orelse return error.MalformedPivotXml;
        if (fdr != 1 or fdc != 1 or fhr > 1) return error.PivotShapeUnsupported;
        if (loc.row_page_count != null or loc.col_page_count != null) return error.PivotShapeUnsupported;
        // The row field.
        const pf = def.fields[rf];
        if (pf.axis != .row or !pf.has_items or pf.has_other_children or pf.show_all or pf.items_other_attrs) return error.PivotShapeUnsupported;
        if (pf.items_count) |n| if (n != pf.items.len) return error.MalformedPivotXml;
        if (pf.sort_type != .manual and pf.sort_type != .ascending) return error.PivotShapeUnsupported;
        try checkRowFieldAttrs(pf.attrs);
        for (pf.items) |it| if (it.other_attrs) return error.PivotShapeUnsupported;
        if (!rb.fields[rf].indexed) return error.PivotShapeUnsupported;
        // Each `<rowItems>` data item names one position of the row
        // field's list — a cache item, once; the grand total names
        // position 0 (Codex #206 r2 REL-202).
        const named = try arena.alloc(bool, pf.items.len);
        @memset(named, false);
        for (ri.items) |it| {
            const pos = it.xs[0];
            if (it.t != null) {
                if (pos != 0) return error.PivotShapeUnsupported;
                continue;
            }
            if (pos >= pf.items.len or pf.items[pos].t != null or named[pos]) return error.PivotShapeUnsupported;
            named[pos] = true;
        }
        // Every other field: unplaced, or a data field.
        for (def.fields, 0..) |f, k| {
            if (k == rf) continue;
            if (f.axis != null or f.axis_raw != null) return error.PivotShapeUnsupported;
        }
        // The data fields.
        for (def.data_fields) |df| {
            if (df.fld >= def.fields.len) return error.MalformedPivotXml;
            if (df.show_data_as) |s| if (!attrIs(s, "normal")) return error.PivotShapeUnsupported;
            // Text under a numeric aggregate is skipped, as SUM skips
            // it on the grid; `count` counts it. The dispersions need
            // Excel's own summation order to agree to the bit.
            switch (df.subtotal) {
                .sum, .average, .min, .max, .product, .count, .count_nums => {},
                .std_dev, .std_dev_p, .variance, .variance_p, .unknown => return error.PivotShapeUnsupported,
            }
        }
    }

    /// An enum-valued attribute kept as written says `want` by its
    /// decoded value — `def&#x61;ult` is `default` (Codex #206 r31
    /// REL-3101); one that does not decode says nothing.
    pub fn attrIs(raw: []const u8, want: []const u8) bool {
        var buf: [32]u8 = undefined;
        const v = wbxml.decodeScalarAttr(&buf, raw) orelse return false;
        return std.mem.eql(u8, v, want);
    }

    /// Root display options that change what the cells hold, which
    /// the layout neither models nor oracles (Codex #206 r16
    /// REL-1603): headers hidden, a missing-value caption, merged
    /// labels, empty rows / columns shown, an error caption,
    /// asterisked or non-visual totals, hidden items subtotalled.
    fn checkRootDisplayAttrs(attrs: []const u8) RebuildError!void {
        var it: pivot_xml.AttrIter = .{ .attrs = attrs };
        while (it.next()) |a| {
            const n = a.name;
            if (std.mem.eql(u8, n, "missingCaption") or std.mem.eql(u8, n, "errorCaption")) return error.PivotShapeUnsupported;
            const must_hold = std.mem.eql(u8, n, "showHeaders") or std.mem.eql(u8, n, "showMissing") or std.mem.eql(u8, n, "visualTotals");
            const must_not = std.mem.eql(u8, n, "mergeItem") or std.mem.eql(u8, n, "showEmptyRow") or std.mem.eql(u8, n, "showEmptyCol") or
                std.mem.eql(u8, n, "showError") or std.mem.eql(u8, n, "asteriskTotals") or std.mem.eql(u8, n, "subtotalHiddenItems");
            if (!must_hold and !must_not) continue;
            var buf: [32]u8 = undefined;
            const v = wbxml.decodeScalarAttr(&buf, a.value) orelse return error.PivotShapeUnsupported;
            const is_false = std.mem.eql(u8, v, "0") or std.mem.eql(u8, v, "false");
            const is_true = std.mem.eql(u8, v, "1") or std.mem.eql(u8, v, "true");
            if (must_hold and !is_true) return error.PivotShapeUnsupported;
            if (must_not and !is_false) return error.PivotShapeUnsupported;
        }
        if (it.malformed) return error.PivotShapeUnsupported;
    }

    /// Attributes of the row field that would change what the layout
    /// shows: a top-N filter (`autoShow`), a custom subtotal set, a
    /// multi-select filter, a blank row after each item — and the
    /// field's own `compact="0"`, which lays its labels out under
    /// the field's name rather than `Row Labels` (Codex #206 r5
    /// REL-501).
    fn checkRowFieldAttrs(attrs: []const u8) RebuildError!void {
        var it: pivot_xml.AttrIter = .{ .attrs = attrs };
        while (it.next()) |a| {
            const n = a.name;
            const flagged = std.mem.eql(u8, n, "autoShow") or std.mem.eql(u8, n, "multipleItemSelectionAllowed") or
                std.mem.eql(u8, n, "hideNewItems") or std.mem.eql(u8, n, "insertBlankRow") or
                (std.mem.endsWith(u8, n, "Subtotal") and !std.mem.eql(u8, n, "defaultSubtotal"));
            const must_hold = std.mem.eql(u8, n, "compact");
            if (!flagged and !must_hold) continue;
            var buf: [32]u8 = undefined;
            const v = wbxml.decodeScalarAttr(&buf, a.value) orelse return error.PivotShapeUnsupported;
            const is_false = std.mem.eql(u8, v, "0") or std.mem.eql(u8, v, "false");
            const is_true = std.mem.eql(u8, v, "1") or std.mem.eql(u8, v, "true");
            if (flagged and !is_false) return error.PivotShapeUnsupported;
            if (must_hold and !is_true) return error.PivotShapeUnsupported;
        }
        if (it.malformed) return error.PivotShapeUnsupported;
    }

    /// `ascending`: numbers by value, then text under the workbook's
    /// collation (folded, then by spelling), the blank last; ties keep
    /// inventory order.
    fn sortItems(arena: Allocator, field: Field, order: []u32) RebuildError!void {
        const keys = try arena.alloc([]const u8, field.items.len);
        for (field.items, keys) |it, *k| k.* = if (it.kind == .s) try foldOrRefuse(arena, it.text) else "";
        const Ctx = struct {
            items: []const Field.Item,
            keys: []const []const u8,
            fn rank(kind: pivot_xml.SharedItems.Item.Kind) u8 {
                return switch (kind) {
                    .n => 0,
                    .s => 1,
                    .m => 2,
                    else => unreachable,
                };
            }
            fn less(ctx: @This(), a: u32, b: u32) bool {
                const ia = ctx.items[a];
                const ib = ctx.items[b];
                const ra = rank(ia.kind);
                const rbk = rank(ib.kind);
                if (ra != rbk) return ra < rbk;
                switch (ia.kind) {
                    .n => if (ia.num != ib.num) return ia.num < ib.num,
                    .s => {
                        switch (std.mem.order(u8, ctx.keys[a], ctx.keys[b])) {
                            .lt => return true,
                            .gt => return false,
                            .eq => switch (std.mem.order(u8, ia.text, ib.text)) {
                                .lt => return true,
                                .gt => return false,
                                .eq => {},
                            },
                        }
                    },
                    else => {},
                }
                return a < b;
            }
        };
        std.sort.pdq(u32, order, Ctx{ .items = field.items, .keys = keys }, Ctx.less);
    }

    /// One value of field `k` in data row `i`, as the aggregates see it.
    const Scalar = union(enum) { blank, number: f64, text };

    fn scalarAt(rb: *const Rebuild, k: u32, i: usize) Scalar {
        const f = rb.fields[k];
        if (f.indexed) {
            const it = f.items[f.index_of_row[i]];
            return switch (it.kind) {
                .m => .blank,
                .n => .{ .number = it.num },
                .s => .text,
                else => unreachable,
            };
        }
        return switch (rb.rows[i][k]) {
            .blank => .blank,
            .number => |lex| .{ .number = parseNumber(lex).? },
            .string => unreachable, // a string enumerates the field
        };
    }

    /// One data field's running aggregate over a group of rows, in
    /// record order. The grand total is two things at once, and the
    /// corpus fixes both to the last bit: a SUM is the fold of the
    /// subtotals in item order (`294.80000000000007` is three group
    /// sums added, not fifty records), while an AVERAGE is one running
    /// sum over every record in record order, divided (mtcars:
    /// `20.210344827586205`, which no fold of the three subtotals
    /// gives). Count, min and max are order-blind; a product folds
    /// like a sum (unverified either way: the corpus carries none).
    const Agg = struct {
        sum: f64 = 0,
        prod: f64 = 1,
        min: f64 = 0,
        max: f64 = 0,
        numbers: u64 = 0,
        filled: u64 = 0,

        fn number(self: *Agg, x: f64) void {
            if (self.numbers == 0 or x < self.min) self.min = x;
            if (self.numbers == 0 or x > self.max) self.max = x;
            self.numbers += 1;
            self.filled += 1;
            self.sum += x;
            self.prod *= x;
        }

        fn fold(self: *Agg, g: Agg) void {
            if (g.numbers > 0) {
                if (self.numbers == 0 or g.min < self.min) self.min = g.min;
                if (self.numbers == 0 or g.max > self.max) self.max = g.max;
                self.sum += g.sum;
                self.prod *= g.prod;
            }
            self.numbers += g.numbers;
            self.filled += g.filled;
        }
    };

    fn aggregateGroup(rb: *const Rebuild, fld: u32, group: []const u32) Agg {
        var agg: Agg = .{};
        for (group) |i| switch (scalarAt(rb, fld, i)) {
            .blank => {},
            .text => agg.filled += 1,
            .number => |x| agg.number(x),
        };
        return agg;
    }

    /// The cell an aggregate renders to: shortest round-trip, or an
    /// empty cell where there was nothing to aggregate.
    fn renderAgg(arena: Allocator, df: pivot_xml.DataField, agg: Agg) RebuildError!@FieldType(OutCell, "value") {
        const n = agg.numbers;
        const value: ?f64 = switch (df.subtotal) {
            .sum => if (n > 0) agg.sum else null,
            .count => if (agg.filled > 0) @as(f64, @floatFromInt(agg.filled)) else null,
            .count_nums => if (n > 0) @as(f64, @floatFromInt(n)) else null,
            .average => if (n > 0) agg.sum / @as(f64, @floatFromInt(n)) else null,
            .min => if (n > 0) agg.min else null,
            .max => if (n > 0) agg.max else null,
            .product => if (n > 0) agg.prod else null,
            else => unreachable, // checkTableShape
        };
        const x = value orelse return .blank;
        if (!std.math.isFinite(x)) return error.PivotShapeUnsupported;
        return .{ .number = try std.fmt.allocPrint(arena, "{d}", .{x}) };
    }

    /// `dataField@name` is what Excel always writes; a producer that
    /// left it out gets Excel's default caption.
    fn defaultCaption(arena: Allocator, f: pivot_xml.ConsolidateFunction, field_name: []const u8) RebuildError![]const u8 {
        const verb: []const u8 = switch (f) {
            .sum => "Sum",
            .count, .count_nums => "Count",
            .average => "Average",
            .min => "Min",
            .max => "Max",
            .product => "Product",
            else => unreachable,
        };
        return std.mem.concat(arena, u8, &.{ verb, " of ", field_name });
    }

    /// `<pivotField ATTRS><items count><item x/>…<item t="default"/></items></pivotField>`:
    /// the attributes verbatim, an item missing from the records
    /// marked `m="1"` as Excel marks a retained one.
    fn renderPivotField(arena: Allocator, p: []const u8, pf: pivot_xml.PivotField, order: []const u32, counts: []const u32, default_item: bool) RebuildError![]const u8 {
        var out: std.ArrayListUnmanaged(u8) = .empty;
        try out.append(arena, '<');
        try out.appendSlice(arena, p);
        try out.appendSlice(arena, "pivotField");
        if (pf.attrs.len > 0 and !isXmlSpace(pf.attrs[0])) try out.append(arena, ' ');
        try out.appendSlice(arena, pf.attrs);
        try out.appendSlice(arena, "><");
        try out.appendSlice(arena, p);
        try out.appendSlice(arena, "items count=\"");
        try out.appendSlice(arena, try std.fmt.allocPrint(arena, "{d}", .{order.len + @intFromBool(default_item)}));
        try out.appendSlice(arena, "\">");
        for (order) |k| {
            try out.append(arena, '<');
            try out.appendSlice(arena, p);
            try out.appendSlice(arena, "item x=\"");
            try out.appendSlice(arena, try std.fmt.allocPrint(arena, "{d}", .{k}));
            try out.appendSlice(arena, if (counts[k] == 0) "\" m=\"1\"/>" else "\"/>");
        }
        if (default_item) {
            try out.append(arena, '<');
            try out.appendSlice(arena, p);
            try out.appendSlice(arena, "item t=\"default\"/>");
        }
        try out.appendSlice(arena, "</");
        try out.appendSlice(arena, p);
        try out.appendSlice(arena, "items></");
        try out.appendSlice(arena, p);
        try out.appendSlice(arena, "pivotField>");
        return out.items;
    }

    /// `<rowItems count><i><x/></i><i><x v="1"/></i>…<i t="grand"><x/></i></rowItems>`
    /// — `v` is the position in the item list, omitted at 0 as Excel
    /// omits it.
    fn renderRowItems(arena: Allocator, p: []const u8, shown: []const u32, grand: bool) RebuildError![]const u8 {
        var out: std.ArrayListUnmanaged(u8) = .empty;
        try out.append(arena, '<');
        try out.appendSlice(arena, p);
        try out.appendSlice(arena, "rowItems count=\"");
        try out.appendSlice(arena, try std.fmt.allocPrint(arena, "{d}", .{shown.len + @intFromBool(grand)}));
        try out.appendSlice(arena, "\">");
        for (shown) |pos| {
            try out.append(arena, '<');
            try out.appendSlice(arena, p);
            try out.appendSlice(arena, "i><");
            try out.appendSlice(arena, p);
            if (pos == 0) {
                try out.appendSlice(arena, "x/></");
            } else {
                try out.appendSlice(arena, "x v=\"");
                try out.appendSlice(arena, try std.fmt.allocPrint(arena, "{d}", .{pos}));
                try out.appendSlice(arena, "\"/></");
            }
            try out.appendSlice(arena, p);
            try out.appendSlice(arena, "i>");
        }
        if (grand) {
            try out.append(arena, '<');
            try out.appendSlice(arena, p);
            try out.appendSlice(arena, "i t=\"grand\"><");
            try out.appendSlice(arena, p);
            try out.appendSlice(arena, "x/></");
            try out.appendSlice(arena, p);
            try out.appendSlice(arena, "i>");
        }
        try out.appendSlice(arena, "</");
        try out.appendSlice(arena, p);
        try out.appendSlice(arena, "rowItems>");
        return out.items;
    }

    pub fn parseNumber(lex: []const u8) ?f64 {
        if (!isXsdDoubleLexical(lex)) return null;
        const x = std.fmt.parseFloat(f64, lex) catch return null;
        if (!std.math.isFinite(x)) return null;
        return x;
    }

    fn isXsdDoubleLexical(s: []const u8) bool {
        var i: usize = 0;
        if (i < s.len and (s[i] == '+' or s[i] == '-')) i += 1;
        var int_digits: usize = 0;
        while (i < s.len and std.ascii.isDigit(s[i])) : (i += 1) int_digits += 1;
        var frac_digits: usize = 0;
        if (i < s.len and s[i] == '.') {
            i += 1;
            while (i < s.len and std.ascii.isDigit(s[i])) : (i += 1) frac_digits += 1;
        }
        if (int_digits == 0 and frac_digits == 0) return false;
        if (i < s.len and (s[i] == 'e' or s[i] == 'E')) {
            i += 1;
            if (i < s.len and (s[i] == '+' or s[i] == '-')) i += 1;
            var exp_digits: usize = 0;
            while (i < s.len and std.ascii.isDigit(s[i])) : (i += 1) exp_digits += 1;
            if (exp_digits == 0) return false;
        }
        return i == s.len;
    }

    /// `containsInteger`'s reading: integral and within a 32-bit int —
    /// a hint Excel sets for such fields, safe to leave unset.
    fn isInteger(x: f64) bool {
        return @trunc(x) == x and x >= -2147483648.0 and x <= 2147483647.0;
    }

    /// One key per numeric value: `-0` and `0` are one item.
    fn numberKey(x: f64) u64 {
        return @bitCast(if (x == 0) @as(f64, 0) else x);
    }

    fn foldOrRefuse(arena: Allocator, s: []const u8) RebuildError![]const u8 {
        const folded = fold(arena, s) catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            else => return error.MalformedPivotXml,
        };
        return folded orelse error.PivotShapeUnsupported;
    }

    /// An inventory string as written (`<s v>` is an ST_Xstring
    /// attribute, like a field name).
    fn decodeItem(arena: Allocator, raw: []const u8) RebuildError![]const u8 {
        return formula.decode.decodeAt(arena, .pivot_field_name, raw) catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            else => return error.MalformedPivotXml,
        };
    }

    fn mapRecordsParse(e: pivot_xml.Error) RebuildError {
        return switch (e) {
            error.OutOfMemory => error.OutOfMemory,
            error.MalformedXml => error.MalformedPivotXml,
        };
    }
};

// ─── Test fixture ────────────────────────────────────────────────────

/// A synthetic pivot workbook, one per source spelling, for the tests
/// that pin the graph walk, the audit and the CLI shape without a
/// fetched fixture. Public so `pkg/editor.zig` and `src/cli.zig` build
/// the same workbook this file's tests do.
///
/// Sheet 0 `Data` holds `A1:C4` (`Region`, `Qty`, `Price` + three rows);
/// sheet 1 `Report` hosts `PivotTable1` at `A3:B6` (rows: Region;
/// values: Sum of Qty). `cacheId` is 7, chosen to differ from every
/// ordinal so a test that confuses the two fails.
/// The `pivots` NDJSON records — the S6 shape, written once for the
/// CLI and the C ABI (S3a). Lives in its own file; rooted here so its
/// tests run under `test-pivots` and the full suite alike.
pub const ndjson = @import("pivot_ndjson.zig");

test {
    _ = ndjson;
}

pub const fixture = struct {
    pub const SourceKind = enum {
        /// `<worksheetSource sheet="Data" ref="A1:C4"/>`
        sheet_ref,
        /// `<worksheetSource name="SalesTbl"/>` — a table on `Data`.
        table_name,
        /// `<worksheetSource name="PivotSrc"/>` — a defined name
        /// whose body is `Data!$A$1:$C$4`.
        defined_name,
        /// `<worksheetSource r:id="rIdExt" sheet="Sheet1" ref="A1:C4"/>`
        /// — another workbook.
        external,
        /// `<worksheetSource sheet="Nope" ref="A1:C4"/>` — a sheet
        /// the workbook does not have.
        dangling,
        /// `<cacheSource type="consolidation">` with two range sets:
        /// `<rangeSet sheet="Data" ref="A1:C4"/>` and
        /// `<rangeSet name="PivotSrc"/>`, the name's body
        /// `Report!$A$1:$B$2` — so the second set reads the HOST sheet.
        consolidation,
    };

    pub const cache_id: u32 = 7;

    pub fn write(allocator: Allocator, io: std.Io, path: []const u8, kind: SourceKind) !void {
        {
            var w = zlsx.Writer.init(allocator);
            defer w.deinit();
            var data = try w.addSheet("Data");
            try data.writeRow(&.{ .{ .string = "Region" }, .{ .string = "Qty" }, .{ .string = "Price" } });
            try data.writeRow(&.{ .{ .string = "East" }, .{ .integer = 3 }, .{ .number = 1.5 } });
            try data.writeRow(&.{ .{ .string = "West" }, .{ .integer = 4 }, .{ .number = 2.5 } });
            try data.writeRow(&.{ .{ .string = "East" }, .{ .integer = 5 }, .{ .number = 3.5 } });
            var report = try w.addSheet("Report");
            try report.writeRow(&.{.{ .string = "pivot host" }});
            try w.save(io, path);
        }

        // Injected through a real save/reopen: a genuine pivot workbook
        // arrives from disk, and the reopen is what refreshes every
        // relationship cache the walk reads.
        var store = try PartStore.open(allocator, io, path);
        defer store.deinit();

        const source: []const u8 = switch (kind) {
            .sheet_ref => "<cacheSource type=\"worksheet\"><worksheetSource sheet=\"Data\" ref=\"A1:C4\"/></cacheSource>",
            .table_name => "<cacheSource type=\"worksheet\"><worksheetSource name=\"SalesTbl\"/></cacheSource>",
            .defined_name => "<cacheSource type=\"worksheet\"><worksheetSource name=\"PivotSrc\"/></cacheSource>",
            .external => "<cacheSource type=\"worksheet\"><worksheetSource r:id=\"rIdExt\" sheet=\"Sheet1\" ref=\"A1:C4\"/></cacheSource>",
            .dangling => "<cacheSource type=\"worksheet\"><worksheetSource sheet=\"Nope\" ref=\"A1:C4\"/></cacheSource>",
            .consolidation => "<cacheSource type=\"consolidation\"><consolidation autoPage=\"0\"><rangeSets count=\"2\"><rangeSet i1=\"0\" sheet=\"Data\" ref=\"A1:C4\"/><rangeSet i1=\"1\" name=\"PivotSrc\"/></rangeSets></consolidation></cacheSource>",
        };
        const cache_def = try std.fmt.allocPrint(allocator,
            \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            \\<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" refreshedBy="zlsx" refreshedDate="45000.5" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="3">{s}<cacheFields count="3"><cacheField name="Region" numFmtId="0"><sharedItems count="2"><s v="East"/><s v="West"/></sharedItems></cacheField><cacheField name="Qty" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="3" maxValue="5"/></cacheField><cacheField name="Price" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" minValue="1.5" maxValue="3.5"/></cacheField></cacheFields></pivotCacheDefinition>
        , .{source});
        defer allocator.free(cache_def);
        try store.addPart(
            "xl/pivotCache/pivotCacheDefinition1.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml",
            cache_def,
        );
        try store.addPart(
            "xl/pivotCache/pivotCacheRecords1.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml",
            \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            \\<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" count="3"><r><x v="0"/><n v="3"/><n v="1.5"/></r><r><x v="1"/><n v="4"/><n v="2.5"/></r><r><x v="0"/><n v="5"/><n v="3.5"/></r></pivotCacheRecords>
            ,
        );
        const external_rel: []const u8 = if (kind == .external)
            "<Relationship Id=\"rIdExt\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath\" Target=\"file:///C:/data/other.xlsx\" TargetMode=\"External\"/>"
        else
            "";
        const cache_rels = try std.fmt.allocPrint(allocator,
            \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            \\<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>{s}</Relationships>
        , .{external_rel});
        defer allocator.free(cache_rels);
        try store.addPart(
            "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
            "application/vnd.openxmlformats-package.relationships+xml",
            cache_rels,
        );
        try store.addPart(
            "xl/pivotTables/pivotTable1.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml",
            \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            \\<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="7" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0"><location ref="A3:B6" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="3"><pivotField axis="axisRow" showAll="0"><items count="3"><item x="0"/><item x="1"/><item t="default"/></items></pivotField><pivotField dataField="1" showAll="0"/><pivotField showAll="0"/></pivotFields><rowFields count="1"><field x="0"/></rowFields><rowItems count="3"><i><x/></i><i><x v="1"/></i><i t="grand"><x/></i></rowItems><colItems count="1"><i/></colItems><dataFields count="1"><dataField name="Sum of Qty" fld="1" baseField="0" baseItem="0"/></dataFields><pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>
            ,
        );
        try store.addPart(
            "xl/pivotTables/_rels/pivotTable1.xml.rels",
            "application/vnd.openxmlformats-package.relationships+xml",
            \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            \\<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>
            ,
        );
        try upsertRels(&store, "xl/worksheets/_rels/sheet2.xml.rels",
            \\<Relationship Id="rIdPT1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
        );

        if (kind == .table_name) {
            try store.addPart(
                "xl/tables/table1.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
                \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                \\<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="SalesTbl" displayName="SalesTbl" ref="A1:C4" totalsRowShown="0"><autoFilter ref="A1:C4"/><tableColumns count="3"><tableColumn id="1" name="Region"/><tableColumn id="2" name="Qty"/><tableColumn id="3" name="Price"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>
                ,
            );
            try upsertRels(&store, "xl/worksheets/_rels/sheet1.xml.rels",
                \\<Relationship Id="rIdT1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
            );
            try spliceBefore(allocator, &store, "xl/worksheets/sheet1.xml", "</worksheet>",
                \\<tableParts count="1"><tablePart r:id="rIdT1"/></tableParts>
            );
        }

        // Workbook: the cache list, and for the defined-name kind the
        // name itself. `<definedNames>` precedes `<pivotCaches>` in the
        // schema, and both precede `</workbook>`.
        const wb_tail: []const u8 = switch (kind) {
            .defined_name => "<definedNames><definedName name=\"PivotSrc\">Data!$A$1:$C$4</definedName></definedNames><pivotCaches><pivotCache cacheId=\"7\" r:id=\"rIdPC1\"/></pivotCaches>",
            .consolidation => "<definedNames><definedName name=\"PivotSrc\">Report!$A$1:$B$2</definedName></definedNames><pivotCaches><pivotCache cacheId=\"7\" r:id=\"rIdPC1\"/></pivotCaches>",
            else => "<pivotCaches><pivotCache cacheId=\"7\" r:id=\"rIdPC1\"/></pivotCaches>",
        };
        try spliceBefore(allocator, &store, "xl/workbook.xml", "</workbook>", wb_tail);
        try spliceBefore(allocator, &store, "xl/_rels/workbook.xml.rels", "</Relationships>",
            \\<Relationship Id="rIdPC1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
        );

        try store.save(io, path);
    }

    /// Add an empty third sheet, `Third` (index 2), to a written fixture
    /// — for the cases where two sheets cannot tell a right answer from
    /// a wrong one (a name scoped to one sheet reading a third).
    pub fn addThirdSheet(allocator: Allocator, io: std.Io, path: []const u8) !void {
        var store = try PartStore.open(allocator, io, path);
        defer store.deinit();
        try store.addPart(
            "xl/worksheets/sheet3.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            \\<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/></worksheet>
            ,
        );
        try spliceBefore(allocator, &store, "xl/workbook.xml", "</sheets>",
            \\<sheet name="Third" sheetId="3" r:id="rIdS3"/>
        );
        try spliceBefore(allocator, &store, "xl/_rels/workbook.xml.rels", "</Relationships>",
            \\<Relationship Id="rIdS3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
        );
        try store.save(io, path);
    }

    /// `write`, plus a second cache (`cacheId` 8) no pivot table reads —
    /// the shape `zlsx pivots` reports as a `pivot_cache` record.
    pub fn writeWithOrphanCache(allocator: Allocator, io: std.Io, path: []const u8, kind: SourceKind) !void {
        try write(allocator, io, path, kind);
        var store = try PartStore.open(allocator, io, path);
        defer store.deinit();
        try store.addPart(
            "xl/pivotCache/pivotCacheDefinition2.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml",
            \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            \\<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" recordCount="0" saveData="0"><cacheSource type="worksheet"><worksheetSource sheet="Report" ref="A1:A1"/></cacheSource><cacheFields count="1"><cacheField name="Note" numFmtId="0"><sharedItems/></cacheField></cacheFields></pivotCacheDefinition>
            ,
        );
        try store.addPart(
            "xl/pivotCache/pivotCacheRecords2.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml",
            \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            \\<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0"/>
            ,
        );
        try store.addPart(
            "xl/pivotCache/_rels/pivotCacheDefinition2.xml.rels",
            "application/vnd.openxmlformats-package.relationships+xml",
            \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            \\<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords2.xml"/></Relationships>
            ,
        );
        try spliceBefore(allocator, &store, "xl/workbook.xml", "</pivotCaches>",
            \\<pivotCache cacheId="8" r:id="rIdPC2"/>
        );
        try spliceBefore(allocator, &store, "xl/_rels/workbook.xml.rels", "</Relationships>",
            \\<Relationship Id="rIdPC2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition2.xml"/>
        );
        try store.save(io, path);
    }

    /// Byte-replace the first `old` in one part of a saved workbook and
    /// save it back — how the tests here, in `pkg/editor.zig` and in
    /// `src/cli.zig` make a fixture wrong in exactly one place.
    pub fn patchPart(allocator: Allocator, io: std.Io, path: []const u8, part: []const u8, old: []const u8, new: []const u8) !void {
        var store = try PartStore.open(allocator, io, path);
        defer store.deinit();
        const p = (try store.part(part)) orelse return error.PartNotFound;
        const at = std.mem.indexOf(u8, p.bytes, old) orelse return error.PatchAnchorNotFound;
        const patched = try std.mem.concat(allocator, u8, &.{ p.bytes[0..at], new, p.bytes[at + old.len ..] });
        defer allocator.free(patched);
        try store.replacePart(part, patched);
        try store.save(io, path);
    }

    fn upsertRels(store: *PartStore, name: []const u8, rel: []const u8) !void {
        if ((try store.part(name)) != null) {
            try spliceBefore(store.allocator, store, name, "</Relationships>", rel);
            return;
        }
        const bytes = try std.fmt.allocPrint(store.allocator,
            \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            \\<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{s}</Relationships>
        , .{rel});
        defer store.allocator.free(bytes);
        try store.addPart(name, "application/vnd.openxmlformats-package.relationships+xml", bytes);
    }

    fn spliceBefore(allocator: Allocator, store: *PartStore, name: []const u8, marker: []const u8, insert: []const u8) !void {
        const part = (try store.part(name)) orelse return error.PartNotFound;
        const at = std.mem.lastIndexOf(u8, part.bytes, marker) orelse return error.MalformedXml;
        const out = try std.mem.concat(allocator, u8, &.{ part.bytes[0..at], insert, part.bytes[at..] });
        defer allocator.free(out);
        try store.replacePart(name, out);
    }
};

// ─── Tests ───────────────────────────────────────────────────────────

const testing = std.testing;

const TestTmp = struct {
    dir: std.testing.TmpDir,
    fn init() TestTmp {
        return .{ .dir = std.testing.tmpDir(.{}) };
    }
    fn deinit(self: *TestTmp) void {
        self.dir.cleanup();
    }
    fn path(self: *TestTmp, alloc: Allocator, io: std.Io, name: []const u8) ![:0]u8 {
        const d = try self.dir.dir.realPathFileAlloc(io, ".", alloc);
        defer alloc.free(d);
        return std.fs.path.joinZ(alloc, &.{ d, name });
    }
};

/// Open the fixture, walk it, and return what the walk says. The
/// workbook view is parsed the way `Workbook.open` parses it.
const Opened = struct {
    store: PartStore,
    wb: wbxml.WorkbookXml,
    pivots: Pivots,

    fn open(alloc: Allocator, io: std.Io, path: []const u8) !Opened {
        var store = try PartStore.open(alloc, io, path);
        errdefer store.deinit();
        const wb_part = (try store.part("xl/workbook.xml")) orelse return error.MissingWorkbookPart;
        var wb = try wbxml.parse(alloc, wb_part.bytes);
        errdefer wb.deinit(alloc);
        const pivots = try collect(alloc, &store, &wb);
        return .{ .store = store, .wb = wb, .pivots = pivots };
    }

    fn deinit(self: *Opened, alloc: Allocator) void {
        self.pivots.deinit();
        self.wb.deinit(alloc);
        self.store.deinit();
    }
};

fn expectFixtureShape(o: *const Opened) !void {
    const p = &o.pivots;
    try testing.expectEqual(@as(usize, 1), p.tables.len);
    try testing.expectEqual(@as(usize, 1), p.caches.len);
    const pt = p.tables[0];
    try testing.expectEqualStrings("PivotTable1", pt.name);
    try testing.expectEqualStrings("xl/pivotTables/pivotTable1.xml", pt.part_name);
    try testing.expectEqual(@as(u32, 1), pt.sheet_idx);
    try testing.expectEqualStrings("Report", pt.sheet_name);
    try testing.expectEqualStrings("xl/worksheets/sheet2.xml", pt.sheet_part_name);
    try testing.expectEqual(@as(usize, 2), p.sheet_names.len);
    try testing.expectEqualStrings("Data", p.sheet_names[0]);
    try testing.expectEqualStrings("A3:B6", pt.definition.location.ref);
    try testing.expectEqualStrings("A3:B6", pt.location_ref);
    try testing.expectEqualStrings("Values", pt.data_caption.?);
    try testing.expectEqual(@as(?usize, 0), pt.cache);
    try testing.expectEqualStrings("Sum of Qty", pt.data_field_names[0].?);
    try testing.expectEqualStrings("Region", p.fieldName(pt, 0).?);
    try testing.expectEqualStrings("Qty", p.fieldName(pt, 1).?);
    try testing.expect(p.fieldName(pt, 3) == null);

    const c = p.caches[0];
    try testing.expectEqual(@as(?u32, fixture.cache_id), c.cache_id);
    try testing.expectEqualStrings("xl/pivotCache/pivotCacheDefinition1.xml", c.part_name);
    try testing.expectEqualStrings("xl/pivotCache/pivotCacheRecords1.xml", c.records_part_name.?);
    try testing.expectEqual(@as(?u32, 3), c.definition.record_count);
    try testing.expectEqual(@as(u32, 1), c.consumer_count);
    try testing.expectEqual(@as(usize, 3), c.field_names.len);
    try testing.expectEqualStrings("Price", c.field_names[2]);
    try testing.expect(c.field_formulas[2] == null);
    try testing.expect(p.hostsPivot(1));
    try testing.expect(!p.hostsPivot(0));
}

test "collect: sheet+ref source resolves to the data sheet via the sheet attribute" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "pivot_sheet_ref.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .sheet_ref);

    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    try expectFixtureShape(&o);
    const r = o.pivots.caches[0].resolution.sheet;
    try testing.expectEqual(@as(u32, 0), r.sheet_idx);
    try testing.expectEqualStrings("Data", r.sheet_name);
    try testing.expectEqualStrings("xl/worksheets/sheet1.xml", r.part_name);
    try testing.expectEqual(ResolvedVia.sheet_attr, r.via);
    try testing.expectEqualStrings("Data", o.pivots.caches[0].source.sheet.?);
    try testing.expectEqualStrings("A1:C4", o.pivots.caches[0].source.ref.?);
    try testing.expect(o.pivots.caches[0].source.name == null);
    try testing.expect(o.pivots.readsFromSheet(0));
    try testing.expect(!o.pivots.readsFromSheet(1));
    // The span the S7b rewriter will splice points at the live bytes.
    const ws = o.pivots.caches[0].definition.source.worksheet.?;
    const span = ws.ref_span.?;
    try testing.expectEqualStrings("A1:C4", o.pivots.caches[0].raw_xml[span.start..span.end]);
}

test "collect: table-name source resolves through the table's host sheet" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "pivot_table_name.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .table_name);

    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    try expectFixtureShape(&o);
    const r = o.pivots.caches[0].resolution.sheet;
    try testing.expectEqual(@as(u32, 0), r.sheet_idx);
    try testing.expectEqual(ResolvedVia.table, r.via);
    try testing.expectEqualStrings("SalesTbl", o.pivots.caches[0].source.name.?);
    try testing.expect(o.pivots.readsFromSheet(0));
    try testing.expect(o.pivots.caches[0].definition.source.worksheet.?.ref == null);
}

test "collect: defined-name source resolves through the name's sheet-qualified body" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "pivot_defined_name.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .defined_name);

    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    try expectFixtureShape(&o);
    const r = o.pivots.caches[0].resolution.sheet;
    try testing.expectEqual(@as(u32, 0), r.sheet_idx);
    try testing.expectEqual(ResolvedVia.defined_name, r.via);
}

test "collect: external and dangling sources never resolve to a local sheet" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();

    const ext = try tt.path(testing.allocator, io, "pivot_external.xlsx");
    defer testing.allocator.free(ext);
    try fixture.write(testing.allocator, io, ext, .external);
    var oe = try Opened.open(testing.allocator, io, ext);
    defer oe.deinit(testing.allocator);
    try expectFixtureShape(&oe);
    try testing.expectEqualStrings("file:///C:/data/other.xlsx", oe.pivots.caches[0].resolution.external);
    try testing.expect(!oe.pivots.readsFromSheet(0));

    const dangling = try tt.path(testing.allocator, io, "pivot_dangling.xlsx");
    defer testing.allocator.free(dangling);
    try fixture.write(testing.allocator, io, dangling, .dangling);
    var od = try Opened.open(testing.allocator, io, dangling);
    defer od.deinit(testing.allocator);
    try expectFixtureShape(&od);
    try testing.expect(od.pivots.caches[0].resolution == .unresolved);
    try testing.expect(!od.pivots.readsFromSheet(0));
}

test "collect: a workbook without pivots yields an empty view" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "no_pivots.xlsx");
    defer testing.allocator.free(path);
    {
        var w = zlsx.Writer.init(testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, path);
    }
    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    try testing.expectEqual(@as(usize, 0), o.pivots.tables.len);
    try testing.expectEqual(@as(usize, 0), o.pivots.caches.len);
    try testing.expect(!o.pivots.hostsPivot(0));
    try testing.expect(!o.pivots.readsFromSheet(0));
}

test "collect: a malformed pivot part refuses the whole read" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "pivot_malformed.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .sheet_ref);
    {
        var store = try PartStore.open(testing.allocator, io, path);
        defer store.deinit();
        // A pivot table with no `<location>` has no output rectangle.
        try store.replacePart("xl/pivotTables/pivotTable1.xml",
            \\<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="P" cacheId="7"/>
        );
        try store.save(io, path);
    }
    try testing.expectError(error.MalformedPivotXml, Opened.open(testing.allocator, io, path));
}

test "collect: the corpus fixture — two pivots, table-named sources, one source-only sheet" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const path = "tests/corpus/openxlsx_loadExample.xlsx";
    std.Io.Dir.cwd().access(io, path, .{}) catch return error.SkipZigTest;

    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    const p = &o.pivots;
    try testing.expectEqual(@as(usize, 2), p.tables.len);
    try testing.expectEqual(@as(usize, 2), p.caches.len);

    // PivotTable1 on `IrisSample` (sheet 0) reads `Table2`, a table on
    // the same sheet; PivotTable3 on `mtCars Pivot` (sheet 3) reads
    // `Table3`, a table on `mtcars` (sheet 2) — which hosts no pivot.
    try testing.expectEqualStrings("PivotTable1", p.tables[0].name);
    try testing.expectEqual(@as(u32, 0), p.tables[0].sheet_idx);
    try testing.expectEqualStrings("G2:K6", p.tables[0].definition.location.ref);
    try testing.expectEqualStrings("PivotTable3", p.tables[1].name);
    try testing.expectEqual(@as(u32, 3), p.tables[1].sheet_idx);
    try testing.expectEqualStrings("A1:D5", p.tables[1].definition.location.ref);

    const c0 = p.cacheOf(p.tables[0]).?;
    try testing.expectEqual(@as(?u32, 0), c0.cache_id);
    try testing.expectEqualStrings("Table2", c0.definition.source.worksheet.?.name.?);
    try testing.expectEqual(@as(u32, 0), c0.resolution.sheet.sheet_idx);
    try testing.expectEqual(ResolvedVia.table, c0.resolution.sheet.via);
    try testing.expectEqual(@as(usize, 5), c0.field_names.len);
    try testing.expectEqualStrings("Sepal Length", c0.field_names[0]);
    try testing.expectEqualStrings("xl/pivotCache/pivotCacheRecords1.xml", c0.records_part_name.?);

    const c1 = p.cacheOf(p.tables[1]).?;
    try testing.expectEqual(@as(?u32, 1), c1.cache_id);
    try testing.expectEqualStrings("Table3", c1.definition.source.worksheet.?.name.?);
    try testing.expectEqual(@as(u32, 2), c1.resolution.sheet.sheet_idx);
    try testing.expectEqualStrings("mtcars", c1.resolution.sheet.sheet_name);
    try testing.expectEqualStrings("mtCars Pivot", p.tables[1].sheet_name);
    try testing.expectEqual(ResolvedVia.table, c1.resolution.sheet.via);
    try testing.expectEqual(@as(usize, 11), c1.field_names.len);
    try testing.expectEqualStrings("Average of mpg", p.tables[1].data_field_names[0].?);
    try testing.expectEqual(pivot_xml.ConsolidateFunction.average, p.tables[1].definition.data_fields[0].subtotal);

    // The audit's answer in one line each: hosts are {0, 3}; sources
    // are {0, 2}; sheet 2 is source-only.
    try testing.expect(p.hostsPivot(0) and p.hostsPivot(3));
    try testing.expect(!p.hostsPivot(1) and !p.hostsPivot(2));
    try testing.expect(p.readsFromSheet(0) and p.readsFromSheet(2));
    try testing.expect(!p.readsFromSheet(1) and !p.readsFromSheet(3));
}

test "relLeafIs: Strict and Transitional URIs, case-insensitive leaf, no partial match" {
    try testing.expect(relLeafIs("http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable", "pivotTable"));
    try testing.expect(relLeafIs("http://purl.oclc.org/ooxml/officeDocument/relationships/pivotTable", "pivotTable"));
    try testing.expect(relLeafIs("x/PIVOTTABLE", "pivotTable"));
    try testing.expect(!relLeafIs("x/pivotTables", "pivotTable"));
    try testing.expect(!relLeafIs("x/table", "pivotTable"));
}

test "unquoteSheetSpec: quoted, doubled apostrophe, bare" {
    var buf: [64]u8 = undefined;
    try testing.expectEqualStrings("Data", unquoteSheetSpec(&buf, .{ .first = "Data" }).?);
    try testing.expectEqualStrings("My Data", unquoteSheetSpec(&buf, .{ .first = "'My Data'", .quoted = true }).?);
    try testing.expectEqualStrings("It's", unquoteSheetSpec(&buf, .{ .first = "'It''s'", .quoted = true }).?);
    try testing.expect(unquoteSheetSpec(&buf, .{ .first = "'broken", .quoted = true }) == null);
}

fn patchPart(io: std.Io, path: []const u8, part: []const u8, old: []const u8, new: []const u8) !void {
    return fixture.patchPart(testing.allocator, io, path, part, old, new);
}

test "collect: an orphan cache is listed with no consumer; a vendor pivotCache in extLst is not a cache" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "pivot_orphan.xlsx");
    defer testing.allocator.free(path);
    try fixture.writeWithOrphanCache(testing.allocator, io, path, .sheet_ref);
    try patchPart(io, path, "xl/workbook.xml", "</workbook>",
        \\<extLst><ext uri="{decoy}" xmlns:v="urn:vendor"><v:pivotCaches><v:pivotCache cacheId="99" r:id="rIdNope"/></v:pivotCaches></ext></extLst></workbook>
    );

    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    try testing.expectEqual(@as(usize, 1), o.pivots.tables.len);
    try testing.expectEqual(@as(?usize, 0), o.pivots.tables[0].cache);
    try testing.expectEqual(@as(usize, 2), o.pivots.caches.len);
    try testing.expectEqual(@as(u32, 1), o.pivots.caches[0].consumer_count);
    const orphan = o.pivots.caches[1];
    try testing.expectEqual(@as(?u32, 8), orphan.cache_id);
    try testing.expectEqual(@as(u32, 0), orphan.consumer_count);
    try testing.expectEqualStrings("xl/pivotCache/pivotCacheRecords2.xml", orphan.records_part_name.?);
    try testing.expectEqual(@as(u32, 1), orphan.resolution.sheet.sheet_idx);
    // The orphan reads `Report`, so both sheets are now sources.
    try testing.expect(o.pivots.readsFromSheet(1));
}

test "collect: a recognised edge that leads nowhere refuses the read" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();

    const Case = struct { name: []const u8, part: []const u8, old: []const u8, new: []const u8 };
    const cases = [_]Case{
        // The host sheet's relationship names a pivot part that is not there.
        .{ .name = "missing_pivot.xlsx", .part = "xl/worksheets/_rels/sheet2.xml.rels", .old = "pivotTables/pivotTable1.xml", .new = "pivotTables/pivotTable9.xml" },
        // The pivot's relationship names a cache part that is not there.
        .{ .name = "missing_cache.xlsx", .part = "xl/pivotTables/_rels/pivotTable1.xml.rels", .old = "pivotCacheDefinition1.xml", .new = "pivotCacheDefinition9.xml" },
        // The definition's r:id names a records part that is not there.
        .{ .name = "missing_records.xlsx", .part = "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels", .old = "pivotCacheRecords1.xml", .new = "pivotCacheRecords9.xml" },
        // The pivot's cacheId disagrees with the cache its relationship names.
        .{ .name = "cache_mismatch.xlsx", .part = "xl/pivotTables/pivotTable1.xml", .old = "cacheId=\"7\"", .new = "cacheId=\"9\"" },
        // The workbook's cache entry cannot say which cache it is.
        .{ .name = "cacheid_bad.xlsx", .part = "xl/workbook.xml", .old = "cacheId=\"7\"", .new = "cacheId=\"abc\"" },
        // The workbook's cache entry has no relationship.
        .{ .name = "cache_rid_missing.xlsx", .part = "xl/workbook.xml", .old = "r:id=\"rIdPC1\"", .new = "r:id=\"rIdGone\"" },
    };
    for (cases) |case| {
        const path = try tt.path(testing.allocator, io, case.name);
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path, .sheet_ref);
        try patchPart(io, path, case.part, case.old, case.new);
        try testing.expectError(error.MalformedPivotXml, Opened.open(testing.allocator, io, path));
    }

    // Two caches under one id.
    const dup = try tt.path(testing.allocator, io, "cache_dup.xlsx");
    defer testing.allocator.free(dup);
    try fixture.writeWithOrphanCache(testing.allocator, io, dup, .sheet_ref);
    try patchPart(io, dup, "xl/workbook.xml", "cacheId=\"8\"", "cacheId=\"7\"");
    try testing.expectError(error.MalformedPivotXml, Opened.open(testing.allocator, io, dup));
}

test "collect: a definition without r:id has no records part; an internal r:id source is unresolved" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();

    const none = try tt.path(testing.allocator, io, "no_records.xlsx");
    defer testing.allocator.free(none);
    try fixture.write(testing.allocator, io, none, .sheet_ref);
    // The records relationship stays in the part's rels: without an
    // `r:id` naming it, it is not the cache's records.
    try patchPart(io, none, "xl/pivotCache/pivotCacheDefinition1.xml", " r:id=\"rId1\"", " saveData=\"0\"");
    var on = try Opened.open(testing.allocator, io, none);
    defer on.deinit(testing.allocator);
    try testing.expect(on.pivots.caches[0].records_part_name == null);
    try testing.expect(!on.pivots.caches[0].definition.save_data);

    const internal = try tt.path(testing.allocator, io, "internal_rid.xlsx");
    defer testing.allocator.free(internal);
    try fixture.write(testing.allocator, io, internal, .external);
    try patchPart(io, internal, "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels", " TargetMode=\"External\"", "");
    var oi = try Opened.open(testing.allocator, io, internal);
    defer oi.deinit(testing.allocator);
    try testing.expect(oi.pivots.caches[0].resolution == .unresolved);
}

test "collect: a relationship to an unlisted cache while cacheId names a listed one is refused; external pivot edges are refused" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();

    // Cache B (definition2) exists but is not listed; the pivot's
    // relationship names B while its cacheId="7" names listed cache A.
    const unlisted = try tt.path(testing.allocator, io, "unlisted_cache.xlsx");
    defer testing.allocator.free(unlisted);
    try fixture.writeWithOrphanCache(testing.allocator, io, unlisted, .sheet_ref);
    try patchPart(io, unlisted, "xl/workbook.xml", "<pivotCache cacheId=\"8\" r:id=\"rIdPC2\"/>", "");
    try patchPart(io, unlisted, "xl/pivotTables/_rels/pivotTable1.xml.rels", "pivotCacheDefinition1.xml", "pivotCacheDefinition2.xml");
    try testing.expectError(error.MalformedPivotXml, Opened.open(testing.allocator, io, unlisted));

    // The same relationship with no workbook entry claiming the id at all
    // is a rel-only cache: listed with a null id, not refused.
    const relonly = try tt.path(testing.allocator, io, "relonly_cache.xlsx");
    defer testing.allocator.free(relonly);
    try fixture.writeWithOrphanCache(testing.allocator, io, relonly, .sheet_ref);
    try patchPart(io, relonly, "xl/workbook.xml", "<pivotCache cacheId=\"7\" r:id=\"rIdPC1\"/><pivotCache cacheId=\"8\" r:id=\"rIdPC2\"/>", "");
    var o = try Opened.open(testing.allocator, io, relonly);
    defer o.deinit(testing.allocator);
    try testing.expectEqual(@as(usize, 1), o.pivots.caches.len);
    try testing.expect(o.pivots.caches[0].cache_id == null);
    try testing.expectEqual(@as(?usize, 0), o.pivots.tables[0].cache);

    // A pivot-table edge, then a pivot→cache edge, marked External.
    const ext_host = try tt.path(testing.allocator, io, "ext_host_edge.xlsx");
    defer testing.allocator.free(ext_host);
    try fixture.write(testing.allocator, io, ext_host, .sheet_ref);
    try patchPart(io, ext_host, "xl/worksheets/_rels/sheet2.xml.rels", "Target=\"../pivotTables/pivotTable1.xml\"", "Target=\"../pivotTables/pivotTable1.xml\" TargetMode=\"External\"");
    try testing.expectError(error.MalformedPivotXml, Opened.open(testing.allocator, io, ext_host));
    const ext_cache = try tt.path(testing.allocator, io, "ext_cache_edge.xlsx");
    defer testing.allocator.free(ext_cache);
    try fixture.write(testing.allocator, io, ext_cache, .sheet_ref);
    try patchPart(io, ext_cache, "xl/pivotTables/_rels/pivotTable1.xml.rels", "Target=\"../pivotCache/pivotCacheDefinition1.xml\"", "Target=\"../pivotCache/pivotCacheDefinition1.xml\" TargetMode=\"External\"");
    try testing.expectError(error.MalformedPivotXml, Opened.open(testing.allocator, io, ext_cache));
}

test "collect: spellings decode (entity-encoded ref); a refused name inventory refuses name-based sources only" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();

    const enc = try tt.path(testing.allocator, io, "encoded_ref.xlsx");
    defer testing.allocator.free(enc);
    try fixture.write(testing.allocator, io, enc, .sheet_ref);
    try patchPart(io, enc, "xl/pivotCache/pivotCacheDefinition1.xml", "ref=\"A1:C4\"", "ref=\"A1&#58;C4\"");
    var oe = try Opened.open(testing.allocator, io, enc);
    defer oe.deinit(testing.allocator);
    try testing.expectEqualStrings("A1:C4", oe.pivots.caches[0].source.ref.?);
    // The raw slice and its span are what the part says, for the splice.
    try testing.expectEqualStrings("A1&#58;C4", oe.pivots.caches[0].definition.source.worksheet.?.ref.?);
    try testing.expectEqual(@as(u32, 0), oe.pivots.caches[0].resolution.sheet.sheet_idx);

    // Two defined names spelled the same: the engine refuses the
    // inventory. A name-based source cannot be resolved either way …
    const dup_name = try tt.path(testing.allocator, io, "dup_name.xlsx");
    defer testing.allocator.free(dup_name);
    try fixture.write(testing.allocator, io, dup_name, .defined_name);
    try patchPart(io, dup_name, "xl/workbook.xml", "</definedNames>", "<definedName name=\"PivotSrc\">Data!$A$1:$B$2</definedName></definedNames>");
    try testing.expectError(error.MalformedPivotXml, Opened.open(testing.allocator, io, dup_name));
    // … while a sheet+ref source in the same workbook never needs it.
    const dup_sheet = try tt.path(testing.allocator, io, "dup_name_sheet_src.xlsx");
    defer testing.allocator.free(dup_sheet);
    try fixture.write(testing.allocator, io, dup_sheet, .sheet_ref);
    try patchPart(io, dup_sheet, "xl/workbook.xml", "<pivotCaches>", "<definedNames><definedName name=\"X\">Data!$A$1</definedName><definedName name=\"X\">Data!$A$2</definedName></definedNames><pivotCaches>");
    var od = try Opened.open(testing.allocator, io, dup_sheet);
    defer od.deinit(testing.allocator);
    try testing.expectEqual(@as(u32, 0), od.pivots.caches[0].resolution.sheet.sheet_idx);
}

test "collect: a defined-name body resolves only as one static sheet-qualified area" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();

    const Case = struct { name: []const u8, body: []const u8, resolves: bool };
    const cases = [_]Case{
        .{ .name = "nb_dynamic.xlsx", .body = "Data!$A$1:INDEX(Data!$A:$A,4)", .resolves = false },
        .{ .name = "nb_two_sheets.xlsx", .body = "Data!$A$1:Report!$C$4", .resolves = false },
        .{ .name = "nb_case.xlsx", .body = "data!$A$1:DATA!$C$4", .resolves = true },
        .{ .name = "nb_quoted.xlsx", .body = "'Data'!$A$1:$C$4", .resolves = true },
        .{ .name = "nb_offset.xlsx", .body = "OFFSET(Data!$A$1,0,0,4,3)", .resolves = false },
        .{ .name = "nb_bare.xlsx", .body = "$A$1:$C$4", .resolves = false },
        .{ .name = "nb_cell.xlsx", .body = "Data!$A$1", .resolves = true },
        .{ .name = "nb_cols.xlsx", .body = "Data!$A:$C", .resolves = true },
    };
    for (cases) |case| {
        const path = try tt.path(testing.allocator, io, case.name);
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path, .defined_name);
        try patchPart(io, path, "xl/workbook.xml", "Data!$A$1:$C$4", case.body);
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        const r = o.pivots.caches[0].resolution;
        if (case.resolves) {
            try testing.expectEqual(@as(u32, 0), r.sheet.sheet_idx);
            try testing.expectEqual(ResolvedVia.defined_name, r.sheet.via);
        } else {
            try testing.expect(r == .unresolved);
        }
    }
}

test "collect: relationships are typed, decoded, and singular; sheet roots must reach their parts" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();

    const Case = struct { name: []const u8, part: []const u8, old: []const u8, new: []const u8 };
    const refused = [_]Case{
        // The workbook's cache entry names a relationship of the wrong type.
        .{ .name = "wrong_type.xlsx", .part = "xl/_rels/workbook.xml.rels", .old = "relationships/pivotCacheDefinition\" Target=\"pivotCache/pivotCacheDefinition1.xml\"", .new = "relationships/worksheet\" Target=\"pivotCache/pivotCacheDefinition1.xml\"" },
        // A second cache edge on the pivot part.
        .{ .name = "two_cache_edges.xlsx", .part = "xl/pivotTables/_rels/pivotTable1.xml.rels", .old = "</Relationships>", .new = "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition\" Target=\"../pivotCache/pivotCacheDefinition1.xml\"/></Relationships>" },
        // A sheet root whose relationship dangles.
        .{ .name = "dangling_sheet.xlsx", .part = "xl/workbook.xml", .old = "<sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/>", .new = "<sheet name=\"Data\" sheetId=\"1\" r:id=\"rIdGone\"/>" },
        // A sheet root whose part is absent.
        .{ .name = "missing_sheet_part.xlsx", .part = "xl/_rels/workbook.xml.rels", .old = "Target=\"worksheets/sheet1.xml\"", .new = "Target=\"worksheets/sheet9.xml\"" },
        // A `<tablePart>` whose relationship is missing.
        .{ .name = "dangling_table_part.xlsx", .part = "xl/worksheets/sheet1.xml", .old = "<tablePart r:id=\"rIdT1\"/>", .new = "<tablePart r:id=\"rIdGone\"/>" },
    };
    for (refused) |case| {
        const path = try tt.path(testing.allocator, io, case.name);
        defer testing.allocator.free(path);
        const kind: fixture.SourceKind = if (std.mem.eql(u8, case.name, "dangling_table_part.xlsx")) .table_name else .sheet_ref;
        try fixture.write(testing.allocator, io, path, kind);
        try patchPart(io, path, case.part, case.old, case.new);
        try testing.expectError(error.MalformedPivotXml, Opened.open(testing.allocator, io, path));
    }

    // An entity-encoded r:id is the same id — at the workbook cache
    // entry, at the records edge, and at the external source.
    const enc = try tt.path(testing.allocator, io, "encoded_rid.xlsx");
    defer testing.allocator.free(enc);
    try fixture.write(testing.allocator, io, enc, .external);
    try patchPart(io, enc, "xl/workbook.xml", "r:id=\"rIdPC1\"", "r:id=\"rIdPC&#49;\"");
    try patchPart(io, enc, "xl/pivotCache/pivotCacheDefinition1.xml", " r:id=\"rId1\" refreshedBy", " r:id=\"rId&#49;\" refreshedBy");
    try patchPart(io, enc, "xl/pivotCache/pivotCacheDefinition1.xml", "r:id=\"rIdExt\"", "r:id=\"rId&#69;xt\"");
    var oe = try Opened.open(testing.allocator, io, enc);
    defer oe.deinit(testing.allocator);
    try testing.expectEqual(@as(usize, 1), oe.pivots.caches.len);
    try testing.expectEqualStrings("xl/pivotCache/pivotCacheRecords1.xml", oe.pivots.caches[0].records_part_name.?);
    try testing.expectEqualStrings("file:///C:/data/other.xlsx", oe.pivots.caches[0].resolution.external);

    // A relationships prefix declared on `<cacheSource>` and used on the
    // `<worksheetSource>` below it is the same relationship — the source
    // is the other workbook, not the local `sheet` it also spells.
    const scoped = try tt.path(testing.allocator, io, "scoped_rid.xlsx");
    defer testing.allocator.free(scoped);
    try fixture.write(testing.allocator, io, scoped, .external);
    try patchPart(io, scoped, "xl/pivotCache/pivotCacheDefinition1.xml", "<cacheSource type=\"worksheet\"><worksheetSource r:id=\"rIdExt\"", "<cacheSource type=\"worksheet\" xmlns:q=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><worksheetSource q:id=\"rIdExt\"");
    var osc = try Opened.open(testing.allocator, io, scoped);
    defer osc.deinit(testing.allocator);
    try testing.expectEqualStrings("file:///C:/data/other.xlsx", osc.pivots.caches[0].resolution.external);

    // An external-source relationship of the wrong type is not a workbook.
    const wrong_ext = try tt.path(testing.allocator, io, "wrong_ext_type.xlsx");
    defer testing.allocator.free(wrong_ext);
    try fixture.write(testing.allocator, io, wrong_ext, .external);
    try patchPart(io, wrong_ext, "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels", "relationships/externalLinkPath", "relationships/hyperlink");
    var ow = try Opened.open(testing.allocator, io, wrong_ext);
    defer ow.deinit(testing.allocator);
    try testing.expect(ow.pivots.caches[0].resolution == .unresolved);

    // A table relationship no `<tablePart>` attaches is not a table of
    // the sheet: the name resolves to nothing.
    const orphan_tbl = try tt.path(testing.allocator, io, "orphan_table_rel.xlsx");
    defer testing.allocator.free(orphan_tbl);
    try fixture.write(testing.allocator, io, orphan_tbl, .table_name);
    try patchPart(io, orphan_tbl, "xl/worksheets/sheet1.xml", "<tableParts count=\"1\"><tablePart r:id=\"rIdT1\"/></tableParts>", "");
    var ot = try Opened.open(testing.allocator, io, orphan_tbl);
    defer ot.deinit(testing.allocator);
    try testing.expect(ot.pivots.caches[0].resolution == .unresolved);

    // `location@ref` is entity-decoded for the reader, raw for the splice.
    const loc = try tt.path(testing.allocator, io, "encoded_location.xlsx");
    defer testing.allocator.free(loc);
    try fixture.write(testing.allocator, io, loc, .sheet_ref);
    try patchPart(io, loc, "xl/pivotTables/pivotTable1.xml", "<location ref=\"A3:B6\"", "<location ref=\"A3&#58;B6\"");
    var ol = try Opened.open(testing.allocator, io, loc);
    defer ol.deinit(testing.allocator);
    try testing.expectEqualStrings("A3:B6", ol.pivots.tables[0].location_ref);
    try testing.expectEqualStrings("A3&#58;B6", ol.pivots.tables[0].definition.location.ref);
    const span = ol.pivots.tables[0].definition.location.ref_span;
    try testing.expectEqualStrings("A3&#58;B6", ol.pivots.tables[0].raw_xml[span.start..span.end]);

    // A same-prefix `<pivotCache>` that is not a direct child of
    // `<pivotCaches>` is not a cache entry.
    const deep = try tt.path(testing.allocator, io, "deep_pivotcache.xlsx");
    defer testing.allocator.free(deep);
    try fixture.write(testing.allocator, io, deep, .sheet_ref);
    try patchPart(io, deep, "xl/workbook.xml", "</pivotCaches>", "<extLst><ext><pivotCache cacheId=\"99\" r:id=\"rIdNope\"/></ext></extLst></pivotCaches>");
    var od = try Opened.open(testing.allocator, io, deep);
    defer od.deinit(testing.allocator);
    try testing.expectEqual(@as(usize, 1), od.pivots.caches.len);
}

test "collect: macro sheets are sheet roots; style names decode; a non-text external target refuses" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();

    // An Excel 4.0 macro sheet listed under `<sheets>` is a legitimate
    // root, pivots or not.
    const macro = try tt.path(testing.allocator, io, "macro_sheet.xlsx");
    defer testing.allocator.free(macro);
    try fixture.write(testing.allocator, io, macro, .sheet_ref);
    try patchPart(io, macro, "xl/_rels/workbook.xml.rels", "relationships/worksheet\" Target=\"worksheets/sheet1.xml\"", "relationships/xlMacrosheet\" Target=\"worksheets/sheet1.xml\"");
    var om = try Opened.open(testing.allocator, io, macro);
    defer om.deinit(testing.allocator);
    try testing.expectEqual(@as(usize, 1), om.pivots.tables.len);

    const styled = try tt.path(testing.allocator, io, "style_name.xlsx");
    defer testing.allocator.free(styled);
    try fixture.write(testing.allocator, io, styled, .sheet_ref);
    try patchPart(io, styled, "xl/pivotTables/pivotTable1.xml", "name=\"PivotStyleLight16\"", "name=\"Finance &amp; Ops_x0021_\"");
    var os = try Opened.open(testing.allocator, io, styled);
    defer os.deinit(testing.allocator);
    try testing.expectEqualStrings("Finance & Ops!", os.pivots.tables[0].style_name.?);

    const bad = try tt.path(testing.allocator, io, "bad_target.xlsx");
    defer testing.allocator.free(bad);
    try fixture.write(testing.allocator, io, bad, .external);
    try patchPart(io, bad, "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels", "file:///C:/data/other.xlsx", "file:///C:/data/\xff.xlsx");
    try testing.expectError(error.MalformedPivotXml, Opened.open(testing.allocator, io, bad));
}

test "collect: a table-spelled source under a refused name inventory is refused too" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "table_dup_names.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .table_name);
    try patchPart(io, path, "xl/workbook.xml", "<pivotCaches>", "<definedNames><definedName name=\"X\">Data!$A$1</definedName><definedName name=\"X\">Data!$A$2</definedName></definedNames><pivotCaches>");
    try testing.expectError(error.MalformedPivotXml, Opened.open(testing.allocator, io, path));
}

fn collectForFailures(allocator: Allocator, store: *PartStore, wb: *const wbxml.WorkbookXml) !void {
    var p = try collect(allocator, store, wb);
    p.deinit();
}

test "collect: allocation failure at every point leaves nothing behind" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // The defined-name kind exercises the sheet folds, the symbol table
    // and the body parser; the table-name kind adds the table index.
    // The consolidation kind adds a second resolution per cache; the
    // last iteration walks a name closure (a dynamic body through a
    // second name and a structured reference).
    const kinds = [_]fixture.SourceKind{ .defined_name, .table_name, .consolidation, .defined_name };
    for (kinds, 0..) |kind, k| {
        var name_buf: [32]u8 = undefined;
        const name = try std.fmt.bufPrint(&name_buf, "alloc_fail_{d}.xlsx", .{k});
        const path = try tt.path(testing.allocator, io, name);
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path, kind);
        if (k == kinds.len - 1) {
            try patchPart(io, path, "xl/workbook.xml", "<definedName name=\"PivotSrc\">Data!$A$1:$C$4</definedName>", "<definedName name=\"Anchor\">Report!$D$1</definedName><definedName name=\"PivotSrc\">OFFSET(Anchor,0,0,COUNTA(Data!$A:$A),3)</definedName>");
        }
        var store = try PartStore.open(testing.allocator, io, path);
        defer store.deinit();
        const wb_part = (try store.part("xl/workbook.xml")).?;
        var wb = try wbxml.parse(testing.allocator, wb_part.bytes);
        defer wb.deinit(testing.allocator);
        // Materialise every part first: the store allocates through its
        // own allocator on first access, and that is not what is under test.
        for (try store.partNames()) |n| _ = try store.part(n);
        try testing.checkAllAllocationFailures(testing.allocator, collectForFailures, .{ &store, &wb });
    }
}

test "collect: one pivot part linked from two sheets is two pivots reading one cache" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "pivot_two_hosts.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .sheet_ref);
    {
        var store = try PartStore.open(testing.allocator, io, path);
        defer store.deinit();
        try store.addPart("xl/worksheets/_rels/sheet1.xml.rels", "application/vnd.openxmlformats-package.relationships+xml",
            \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            \\<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPT1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>
        );
        try store.save(io, path);
    }
    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    try testing.expectEqual(@as(usize, 2), o.pivots.tables.len);
    try testing.expectEqual(@as(u32, 0), o.pivots.tables[0].sheet_idx);
    try testing.expectEqual(@as(u32, 1), o.pivots.tables[1].sheet_idx);
    try testing.expectEqual(@as(usize, 1), o.pivots.caches.len);
    try testing.expectEqual(@as(u32, 2), o.pivots.caches[0].consumer_count);
    try testing.expect(o.pivots.hostsPivot(0) and o.pivots.hostsPivot(1));
}

// ─── S7a: the output-location lift ───────────────────────────────────

const pt_head =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="7">
;
const pt_tail =
    \\<pivotFields count="1"><pivotField axis="axisRow"/></pivotFields><rowFields count="1"><field x="0"/></rowFields></pivotTableDefinition>
;

fn ptWith(middle: []const u8) ![]u8 {
    return std.mem.concat(testing.allocator, u8, &.{ pt_head, middle, pt_tail });
}

fn expectMove(middle: []const u8, axis: edit.Axis, idx: u32, kind: edit.Kind, want_middle: []const u8) !void {
    const src = try ptWith(middle);
    defer testing.allocator.free(src);
    const out = try edit.applyToTableDefinition(testing.allocator, src, axis, idx, kind);
    defer testing.allocator.free(out);
    const want = try ptWith(want_middle);
    defer testing.allocator.free(want);
    try testing.expectEqualStrings(want, out);
}

fn expectRefusal(middle: []const u8, axis: edit.Axis, idx: u32, kind: edit.Kind, err: anyerror) !void {
    const src = try ptWith(middle);
    defer testing.allocator.free(src);
    try testing.expectError(err, edit.applyToTableDefinition(testing.allocator, src, axis, idx, kind));
}

test "edit: an insert at or above the rectangle moves it; below, the bytes are untouched" {
    const loc = "<location ref=\"C3:D6\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>";
    // Rows: the offsets inside the rectangle do not move with it.
    try expectMove(loc, .row, 1, .insert, "<location ref=\"C4:D7\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>");
    try expectMove(loc, .row, 3, .insert, "<location ref=\"C4:D7\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>");
    try expectMove(loc, .row, 7, .insert, loc);
    try expectMove(loc, .row, 2, .delete, "<location ref=\"C2:D5\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>");
    try expectMove(loc, .row, 7, .delete, loc);
    // Columns.
    try expectMove(loc, .col, 1, .insert, "<location ref=\"D3:E6\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>");
    try expectMove(loc, .col, 3, .insert, "<location ref=\"D3:E6\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>");
    try expectMove(loc, .col, 5, .insert, loc);
    try expectMove(loc, .col, 2, .delete, "<location ref=\"B3:C6\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>");
    try expectMove(loc, .col, 5, .delete, loc);
}

test "edit: an edit inside the rectangle refuses — Excel refuses it too" {
    const loc = "<location ref=\"C3:D6\"/>";
    try expectRefusal(loc, .row, 4, .insert, error.PivotLocationEditUnsafe);
    try expectRefusal(loc, .row, 6, .insert, error.PivotLocationEditUnsafe);
    try expectRefusal(loc, .row, 3, .delete, error.PivotLocationEditUnsafe);
    try expectRefusal(loc, .row, 6, .delete, error.PivotLocationEditUnsafe);
    try expectRefusal(loc, .col, 4, .insert, error.PivotLocationEditUnsafe);
    try expectRefusal(loc, .col, 3, .delete, error.PivotLocationEditUnsafe);
    try expectRefusal(loc, .col, 4, .delete, error.PivotLocationEditUnsafe);
}

test "edit: report filters widen the footprint above the rectangle and to its right" {
    // One filter: label + value at A1:B1, blank row 2, body at A3.
    // The band is `rowPageCount + 1` rows above and `3 · colPageCount`
    // columns wide (A..C) — a superset of the cells Excel draws.
    const loc = "<location ref=\"A3:B6\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\" rowPageCount=\"1\" colPageCount=\"1\"/>";
    try expectMove(loc, .row, 1, .insert, "<location ref=\"A4:B7\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\" rowPageCount=\"1\" colPageCount=\"1\"/>");
    try expectRefusal(loc, .row, 2, .insert, error.PivotLocationEditUnsafe);
    try expectRefusal(loc, .row, 1, .delete, error.PivotLocationEditUnsafe);
    try expectRefusal(loc, .row, 2, .delete, error.PivotLocationEditUnsafe);
    try expectRefusal(loc, .col, 3, .insert, error.PivotLocationEditUnsafe);
    try expectRefusal(loc, .col, 3, .delete, error.PivotLocationEditUnsafe);
    try expectMove(loc, .col, 4, .insert, loc);
    try expectMove(loc, .col, 1, .insert, "<location ref=\"B3:C6\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\" rowPageCount=\"1\" colPageCount=\"1\"/>");

    // A producer that wrote `<pageFields>` but no counts: the field
    // count stands in for both.
    const bare = "<location ref=\"A3:B6\"/><pageFields count=\"1\"><pageField fld=\"0\" hier=\"-1\"/></pageFields>";
    try expectRefusal(bare, .row, 2, .insert, error.PivotLocationEditUnsafe);
    try expectRefusal(bare, .col, 3, .insert, error.PivotLocationEditUnsafe);
    try expectMove(bare, .row, 1, .insert, "<location ref=\"A4:B7\"/><pageFields count=\"1\"><pageField fld=\"0\" hier=\"-1\"/></pageFields>");
    // Two fields, no counts: over-then-down would put the second pair
    // at D:E, so the count bounds the columns too — A..F is inside, G
    // is not (Codex #200 r3 REL-041).
    const two = "<location ref=\"A4:B6\"/><pageFields count=\"2\"><pageField fld=\"0\" hier=\"-1\"/><pageField fld=\"2\" hier=\"-1\"/></pageFields>";
    try expectRefusal(two, .col, 4, .insert, error.PivotLocationEditUnsafe);
    try expectRefusal(two, .col, 4, .delete, error.PivotLocationEditUnsafe);
    try expectRefusal(two, .col, 6, .insert, error.PivotLocationEditUnsafe);
    try expectMove(two, .col, 7, .insert, two);
    try expectRefusal(two, .row, 3, .insert, error.PivotLocationEditUnsafe);
    try expectMove(two, .row, 1, .insert, "<location ref=\"A5:B7\"/><pageFields count=\"2\"><pageField fld=\"0\" hier=\"-1\"/><pageField fld=\"2\" hier=\"-1\"/></pageFields>");

    // A rectangle too close to the top for its band: the band is
    // clamped to row 1, never wraps — including at the largest count a
    // part can declare (Codex #200 r1 REL-037).
    try expectMove("<location ref=\"A2:B4\" rowPageCount=\"3\"/>", .row, 1, .insert, "<location ref=\"A3:B5\" rowPageCount=\"3\"/>");
    try expectMove("<location ref=\"A3:B6\" rowPageCount=\"4294967295\"/>", .row, 1, .insert, "<location ref=\"A4:B7\" rowPageCount=\"4294967295\"/>");
    try expectRefusal("<location ref=\"A3:B6\" rowPageCount=\"4294967295\"/>", .row, 1, .delete, error.PivotLocationEditUnsafe);
    try expectRefusal("<location ref=\"A3:B6\" rowPageCount=\"4294967295\"/>", .row, 2, .insert, error.PivotLocationEditUnsafe);
    // The band's width is checked arithmetic too.
    try expectRefusal("<location ref=\"A3:B6\" rowPageCount=\"1\" colPageCount=\"4294967295\"/>", .col, 9, .insert, error.PivotCoordinateOverflow);
}

test "mayReadFromSheet: an unresolved source may be any sheet; external and non-worksheet sources are not" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "pivot_may_read.xlsx");
    defer testing.allocator.free(path);

    // A dynamic defined name: Excel accepts `OFFSET(...)` as a pivot
    // source; the reader cannot place it, so every sheet may be read.
    try fixture.write(testing.allocator, io, path, .defined_name);
    try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "Data!$A$1:$C$4", "OFFSET(Report!$D$1,0,0,4,3)");
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try testing.expect(o.pivots.caches[0].resolution == .unresolved);
        try testing.expect(!o.pivots.readsFromSheet(1));
        try testing.expect(o.pivots.mayReadFromSheet(0) and o.pivots.mayReadFromSheet(1));
    }
    // An external workbook is proven elsewhere.
    try fixture.write(testing.allocator, io, path, .external);
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try testing.expect(o.pivots.caches[0].resolution == .external);
        try testing.expect(!o.pivots.mayReadFromSheet(0) and !o.pivots.mayReadFromSheet(1));
    }
    // A resolved local source: the one sheet, and no other.
    try fixture.write(testing.allocator, io, path, .sheet_ref);
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try testing.expect(o.pivots.mayReadFromSheet(0) and !o.pivots.mayReadFromSheet(1));
    }
    // A source type the reader does not know may be anything.
    try fixture.patchPart(testing.allocator, io, path, "xl/pivotCache/pivotCacheDefinition1.xml", "type=\"worksheet\"", "type=\"lakehouse\"");
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try testing.expect(o.pivots.caches[0].definition.source.type == .unknown);
        try testing.expect(o.pivots.mayReadFromSheet(1));
    }
}

fn expectRect(b: ?Bounds, tl_col: u32, tl_row: u32, br_col: u32, br_row: u32) !void {
    const r = (b orelse return error.TestExpectedBounds).rect;
    try testing.expectEqual(edit.Rect{ .tl_col = tl_col, .tl_row = tl_row, .br_col = br_col, .br_row = br_row }, r);
}

fn expectA1(b: ?Bounds, want: []const u8) !void {
    var buf: [Bounds.format_buf_len]u8 = undefined;
    try testing.expectEqualStrings(want, (b orelse return error.TestExpectedBounds).formatA1(&buf) orelse return error.TestExpectedBounds);
}

test "bounds: every spelling that proves an area carries it, and one that does not carries none" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "pivot_bounds.xlsx");
    defer testing.allocator.free(path);

    // A direct `ref`.
    try fixture.write(testing.allocator, io, path, .sheet_ref);
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        const r = o.pivots.caches[0].resolution.sheet;
        try expectRect(r.bounds, 1, 1, 3, 4);
        try expectA1(r.bounds, "A1:C4");
    }
    // A table's `ref`.
    try fixture.write(testing.allocator, io, path, .table_name);
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        const r = o.pivots.caches[0].resolution.sheet;
        try testing.expectEqual(ResolvedVia.table, r.via);
        try expectRect(r.bounds, 1, 1, 3, 4);
    }
    // A static name body, in each of its shapes.
    const Case = struct { body: []const u8, a1: []const u8 };
    const cases = [_]Case{
        .{ .body = "Data!$A$1:$C$4", .a1 = "A1:C4" },
        .{ .body = "Data!$C$4:$A$1", .a1 = "A1:C4" },
        .{ .body = "Data!$A$1", .a1 = "A1:A1" },
        .{ .body = "Data!$A:$C", .a1 = "A:C" },
        .{ .body = "Data!$C:$A", .a1 = "A:C" },
        .{ .body = "Data!$1:$4", .a1 = "1:4" },
        // Absolute on purpose: a name body with a relative reference is
        // one the engine refuses to reference (`relative_reference_name`),
        // and S6 made that refuse the read.
        .{ .body = "'Data'!$B$2:data!$C$3", .a1 = "B2:C3" },
        // Parentheses keep a reference a reference (Codex r4 F2).
        .{ .body = "(Data!$A$1:$C$4)", .a1 = "A1:C4" },
        .{ .body = "((Data!$A$1))", .a1 = "A1:A1" },
    };
    for (cases) |case| {
        try fixture.write(testing.allocator, io, path, .defined_name);
        try patchPart(io, path, "xl/workbook.xml", "Data!$A$1:$C$4", case.body);
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        const r = o.pivots.caches[0].resolution.sheet;
        try testing.expectEqual(@as(u32, 0), r.sheet_idx);
        try testing.expectEqual(ResolvedVia.defined_name, r.via);
        try expectA1(r.bounds, case.a1);
    }
    // A range whose ends are static references of different kinds
    // names the sheet, as before S7b-1, with no bounds (Codex r1 F1).
    for ([_][]const u8{ "Data!$A$1:$C:$C", "Data!$A:$A:$B$2", "Data!$1:$1:$C$4" }) |body| {
        try fixture.write(testing.allocator, io, path, .defined_name);
        try patchPart(io, path, "xl/workbook.xml", "Data!$A$1:$C$4", body);
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        const r = o.pivots.caches[0].resolution.sheet;
        try testing.expectEqual(@as(u32, 0), r.sheet_idx);
        try testing.expectEqual(ResolvedVia.defined_name, r.via);
        try testing.expect(r.bounds == null);
        try testing.expect(o.pivots.readsFromSheet(0) and !o.pivots.mayReadFromSheet(1));
    }
    // A reversed direct `ref` is the same rectangle (Codex r1 F6).
    try fixture.write(testing.allocator, io, path, .sheet_ref);
    try patchPart(io, path, "xl/pivotCache/pivotCacheDefinition1.xml", "ref=\"A1:C4\"", "ref=\"C4:A1\"");
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectA1(o.pivots.caches[0].resolution.sheet.bounds, "A1:C4");
        try testing.expectEqualStrings("C4:A1", o.pivots.caches[0].source.ref.?);
    }
    // `sheet` + `name`, no `ref`: the sheet is the identity, the carrier
    // lends its bounds when it is on that sheet (Codex r1 F4).
    try fixture.write(testing.allocator, io, path, .table_name);
    try patchPart(io, path, "xl/pivotCache/pivotCacheDefinition1.xml", "<worksheetSource name=\"SalesTbl\"/>", "<worksheetSource sheet=\"Data\" name=\"SalesTbl\"/>");
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        const r = o.pivots.caches[0].resolution.sheet;
        try testing.expectEqual(ResolvedVia.sheet_attr, r.via);
        try expectA1(r.bounds, "A1:C4");
    }
    // The carrier is looked up from the stated sheet: its scoped name
    // shadows the workbook one there (Codex r2 F3).
    try fixture.write(testing.allocator, io, path, .defined_name);
    try patchPart(io, path, "xl/workbook.xml", "<definedName name=\"PivotSrc\">Data!$A$1:$C$4</definedName>", "<definedName name=\"Rate\">Report!$A$1:$A$2</definedName><definedName name=\"Rate\" localSheetId=\"1\">Report!$C$3:$D$4</definedName>");
    try patchPart(io, path, "xl/pivotCache/pivotCacheDefinition1.xml", "<worksheetSource name=\"PivotSrc\"/>", "<worksheetSource sheet=\"Report\" name=\"Rate\"/>");
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        const r = o.pivots.caches[0].resolution.sheet;
        try testing.expectEqual(@as(u32, 1), r.sheet_idx);
        try testing.expectEqual(ResolvedVia.sheet_attr, r.via);
        try expectA1(r.bounds, "C3:D4");
    }
    try fixture.write(testing.allocator, io, path, .defined_name);
    try patchPart(io, path, "xl/pivotCache/pivotCacheDefinition1.xml", "<worksheetSource name=\"PivotSrc\"/>", "<worksheetSource sheet=\"Report\" name=\"PivotSrc\"/>");
    {
        // The name's area is on `Data`; the source says `Report` — the
        // sheet wins, the bounds are not lent across sheets.
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        const r = o.pivots.caches[0].resolution.sheet;
        try testing.expectEqual(@as(u32, 1), r.sheet_idx);
        try testing.expectEqual(ResolvedVia.sheet_attr, r.via);
        try testing.expect(r.bounds == null);
    }
    // `Data!$A:$C` is whole columns, not a rectangle; `Data!$1:$4` whole rows.
    {
        try fixture.write(testing.allocator, io, path, .defined_name);
        try patchPart(io, path, "xl/workbook.xml", "Data!$A$1:$C$4", "Data!$A:$C");
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        const b = o.pivots.caches[0].resolution.sheet.bounds.?;
        try testing.expectEqual(@as(u32, 1), b.whole_columns.first_col);
        try testing.expectEqual(@as(u32, 3), b.whole_columns.last_col);
    }
    // A `sheet` attribute alone names a sheet and proves no area; so
    // does a `ref` the rectangle parser rejects.
    for ([_][]const u8{ "<worksheetSource sheet=\"Data\"/>", "<worksheetSource sheet=\"Data\" ref=\"A1:C\"/>" }) |spelling| {
        try fixture.write(testing.allocator, io, path, .sheet_ref);
        try patchPart(io, path, "xl/pivotCache/pivotCacheDefinition1.xml", "<worksheetSource sheet=\"Data\" ref=\"A1:C4\"/>", spelling);
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        const r = o.pivots.caches[0].resolution.sheet;
        try testing.expectEqual(@as(u32, 0), r.sheet_idx);
        try testing.expect(r.bounds == null);
        try testing.expect(o.pivots.dependsOnSheet(0));
    }
    // A table whose `ref` is not a rectangle is a table part the table
    // editor cannot read, and the index refuses the read (S6) — so a
    // table-spelled source that resolves always carries its bounds.
    try fixture.write(testing.allocator, io, path, .table_name);
    try patchPart(io, path, "xl/tables/table1.xml", "ref=\"A1:C4\" totalsRowShown", "ref=\"A1:C\" totalsRowShown");
    try testing.expectError(error.MalformedPivotXml, Opened.open(testing.allocator, io, path));
}

test "formatA1: a value outside the grid formats as nothing, not a panic (Codex r4 F4)" {
    var buf: [Bounds.format_buf_len]u8 = undefined;
    const bad = [_]Bounds{
        .{ .whole_columns = .{ .first_col = 0, .last_col = 1 } },
        .{ .whole_rows = .{ .first_row = 0, .last_row = 1 } },
        .{ .rect = .{ .tl_col = 1, .tl_row = 1, .br_col = 0, .br_row = 1 } },
        .{ .rect = .{ .tl_col = 1, .tl_row = 1, .br_col = 16385, .br_row = 1 } },
        .{ .whole_rows = .{ .first_row = 5, .last_row = 4 } },
    };
    for (bad) |b| try testing.expect(b.formatA1(&buf) == null);
    try testing.expectEqualStrings("A1:A1", (Bounds{ .rect = .{ .tl_col = 1, .tl_row = 1, .br_col = 1, .br_row = 1 } }).formatA1(&buf).?);
}

test "parseBounds: rectangles, whole columns, whole rows, and nothing else" {
    try expectA1(parseBounds("A1:C4"), "A1:C4");
    try expectA1(parseBounds("C4:A1"), "A1:C4");
    try expectA1(parseBounds("A4:C1"), "A1:C4");
    try expectA1(parseBounds("B7"), "B7:B7");
    try expectA1(parseBounds("XFD1048576:A1"), "A1:XFD1048576");
    try expectA1(parseBounds("A:C"), "A:C");
    try expectA1(parseBounds("C:A"), "A:C");
    try expectA1(parseBounds("XFD:XFD"), "XFD:XFD");
    try expectA1(parseBounds("3:9"), "3:9");
    try expectA1(parseBounds("9:3"), "3:9");
    try expectA1(parseBounds("1:1048576"), "1:1048576");
    for ([_][]const u8{ "", ":", "A:", ":C", "A1:C", "A:C4", "a:c", "a1:c4", "XFE:XFE", "XFE1:XFE1", "A0:C4", "A01:C4", "0:4", "01:4", "1:04", "1:1048577", "A:C:D", "$A:$C", "$A$1:$C$4", "A1:C4 " }) |bad| {
        try testing.expect(parseBounds(bad) == null);
    }
}

fn expectUnresolved(r: SourceResolution, why: Unresolved.Why, sheets: []const u32) !void {
    if (r != .unresolved) return error.TestExpectedUnresolved;
    try testing.expectEqual(why, r.unresolved.why);
    try testing.expectEqualSlices(u32, sheets, r.unresolved.sheets);
}

test "unresolved: the provenance says why, and which sheets the spelling still proves" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "pivot_provenance.xlsx");
    defer testing.allocator.free(path);
    const def_part = "xl/pivotCache/pivotCacheDefinition1.xml";

    // A dangling sheet proves nothing.
    try fixture.write(testing.allocator, io, path, .dangling);
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectUnresolved(o.pivots.caches[0].resolution, .dangling_sheet, &.{});
        try testing.expect(!o.pivots.dependsOnSheet(0) and !o.pivots.dependsOnSheet(1));
    }
    // A dynamic body: the sheet it reads is proven, not bounded.
    try fixture.write(testing.allocator, io, path, .defined_name);
    try patchPart(io, path, "xl/workbook.xml", "Data!$A$1:$C$4", "OFFSET(Report!$D$1,0,0,4,3)");
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectUnresolved(o.pivots.caches[0].resolution, .unbounded_body, &.{1});
        try testing.expect(o.pivots.dependsOnSheet(1) and !o.pivots.dependsOnSheet(0));
        // The S7a question is unchanged: an unresolved source may be anywhere.
        try testing.expect(o.pivots.mayReadFromSheet(0) and o.pivots.mayReadFromSheet(1));
        try testing.expect(!o.pivots.readsFromSheet(1));
    }
    // Through another name: `PivotSrc = OFFSET(Anchor,…)`, `Anchor = Report!$D$1`.
    try fixture.write(testing.allocator, io, path, .defined_name);
    try patchPart(io, path, "xl/workbook.xml", "<definedName name=\"PivotSrc\">Data!$A$1:$C$4</definedName>", "<definedName name=\"Anchor\">Report!$D$1</definedName><definedName name=\"PivotSrc\">OFFSET(anchor,0,0,COUNTA(Data!$A:$A),3)</definedName>");
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectUnresolved(o.pivots.caches[0].resolution, .unbounded_body, &.{ 0, 1 });
    }
    // A 3D span proves every sheet between its ends — quoted or not —
    // and a reversed or dangling one is `#REF!`, proving nothing
    // (Codex r1 F3); a union proves both sides.
    for ([_][]const u8{ "SUM(Data:Report!$A$1)", "SUM('Data:Report'!$A$1)", "SUM(data:REPORT!$A$1)" }) |body| {
        try fixture.write(testing.allocator, io, path, .defined_name);
        try patchPart(io, path, "xl/workbook.xml", "Data!$A$1:$C$4", body);
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectUnresolved(o.pivots.caches[0].resolution, .unbounded_body, &.{ 0, 1 });
    }
    for ([_][]const u8{ "SUM(Report:Data!$A$1)", "SUM(Data:Nope!$A$1)", "SUM('Nope:Report'!$A$1)" }) |body| {
        try fixture.write(testing.allocator, io, path, .defined_name);
        try patchPart(io, path, "xl/workbook.xml", "Data!$A$1:$C$4", body);
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectUnresolved(o.pivots.caches[0].resolution, .unbounded_body, &.{});
    }
    // A sheet-scoped name shadowing a workbook one is its own
    // dependency: `N` from the workbook reads `Report`, `Report!N` is
    // the Report-scoped `N` reading `Data` (Codex r1 F2). Visited by
    // identity, not spelling, both are walked.
    try fixture.write(testing.allocator, io, path, .defined_name);
    try patchPart(io, path, "xl/workbook.xml", "<definedName name=\"PivotSrc\">Data!$A$1:$C$4</definedName>", "<definedName name=\"N\">Report!$A$1</definedName><definedName name=\"N\" localSheetId=\"1\">Data!$C$3</definedName><definedName name=\"PivotSrc\">OFFSET(N,0,0,1,1)+OFFSET(Report!N,0,0,1,1)</definedName>");
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectUnresolved(o.pivots.caches[0].resolution, .unbounded_body, &.{ 0, 1 });
    }
    try fixture.write(testing.allocator, io, path, .defined_name);
    try patchPart(io, path, "xl/workbook.xml", "<definedName name=\"PivotSrc\">Data!$A$1:$C$4</definedName>", "<definedName name=\"N\">Report!$A$1</definedName><definedName name=\"N\" localSheetId=\"1\">Data!$C$3</definedName><definedName name=\"PivotSrc\">OFFSET(N,0,0,1,1)</definedName>");
    {
        // Unqualified from the workbook scope: the workbook `N` only.
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectUnresolved(o.pivots.caches[0].resolution, .unbounded_body, &.{1});
    }
    // A workbook-scoped body resolves its unqualified names from the
    // sheet that invoked it: `Report!W`, `W = N`, and the Report-scoped
    // `N` reads `Third` (Codex r2 F1).
    try fixture.write(testing.allocator, io, path, .defined_name);
    try fixture.addThirdSheet(testing.allocator, io, path);
    try patchPart(io, path, "xl/workbook.xml", "<definedName name=\"PivotSrc\">Data!$A$1:$C$4</definedName>", "<definedName name=\"N\" localSheetId=\"1\">Third!$A$1</definedName><definedName name=\"PivotSrc\">OFFSET(Report!W,0,0,1,1)</definedName><definedName name=\"W\">N</definedName>");
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try testing.expectEqual(@as(usize, 3), o.pivots.sheet_names.len);
        try expectUnresolved(o.pivots.caches[0].resolution, .unbounded_body, &.{ 1, 2 });
    }
    // A span qualifier looks the name up from every member: the
    // workbook `N` reads `Data`, the Report-scoped `N` reads `Third`
    // (Codex r2 F2).
    try fixture.write(testing.allocator, io, path, .defined_name);
    try fixture.addThirdSheet(testing.allocator, io, path);
    try patchPart(io, path, "xl/workbook.xml", "<definedName name=\"PivotSrc\">Data!$A$1:$C$4</definedName>", "<definedName name=\"N\">Data!$A$1</definedName><definedName name=\"N\" localSheetId=\"1\">Third!$D$1</definedName><definedName name=\"PivotSrc\">SUM('Data:Report'!N)</definedName>");
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectUnresolved(o.pivots.caches[0].resolution, .unbounded_body, &.{ 0, 1, 2 });
    }
    // A name the engine refuses to reference contributes nothing — it
    // does not fall through to a table of the same spelling (Codex r2 F5).
    try fixture.write(testing.allocator, io, path, .table_name);
    try patchPart(io, path, "xl/workbook.xml", "<pivotCaches>", "<definedNames><definedName name=\"PivotSrc\">OFFSET(SalesTbl,0,0,1,1)</definedName><definedName name=\"SalesTbl\" function=\"1\">SalesTbl</definedName></definedNames><pivotCaches>");
    try patchPart(io, path, def_part, "name=\"SalesTbl\"", "name=\"PivotSrc\"");
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectUnresolved(o.pivots.caches[0].resolution, .unbounded_body, &.{});
    }
    // A chain of any length reaches its end: `PivotSrc = _D0 + _W`,
    // `_D0 = _D1`, …, `_D99 = _W`, `_W = _Target`, `_Target = Report!$A$1`
    // — a depth cap once cut `_Target` off on the deep path and the
    // shallow `_W` was then skipped as visited (Codex r3 F1).
    try fixture.write(testing.allocator, io, path, .defined_name);
    {
        var names: std.ArrayListUnmanaged(u8) = .empty;
        defer names.deinit(testing.allocator);
        try names.appendSlice(testing.allocator, "<definedName name=\"PivotSrc\">_D0+_W</definedName>");
        var k: u32 = 0;
        while (k < 100) : (k += 1) {
            const line = if (k == 99)
                try std.fmt.allocPrint(testing.allocator, "<definedName name=\"_D{d}\">_W</definedName>", .{k})
            else
                try std.fmt.allocPrint(testing.allocator, "<definedName name=\"_D{d}\">_D{d}</definedName>", .{ k, k + 1 });
            defer testing.allocator.free(line);
            try names.appendSlice(testing.allocator, line);
        }
        try names.appendSlice(testing.allocator, "<definedName name=\"_Target\">Report!$A$1</definedName><definedName name=\"_W\">_Target</definedName>");
        try patchPart(io, path, "xl/workbook.xml", "<definedName name=\"PivotSrc\">Data!$A$1:$C$4</definedName>", names.items);
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectUnresolved(o.pivots.caches[0].resolution, .unbounded_body, &.{1});
    }
    // A qualifier that names nothing is `#REF!`: the name behind it is
    // not looked up from the enclosing scope instead (Codex r3 F2).
    for ([_][]const u8{ "OFFSET(Nope!_N,0,0,1,1)", "OFFSET(Report:Data!_N,0,0,1,1)" }) |body| {
        try fixture.write(testing.allocator, io, path, .defined_name);
        try patchPart(io, path, "xl/workbook.xml", "<definedName name=\"PivotSrc\">Data!$A$1:$C$4</definedName>", "<definedName name=\"_N\">Report!$A$1</definedName><definedName name=\"PivotSrc\">PLACEHOLDER</definedName>");
        try patchPart(io, path, "xl/workbook.xml", "PLACEHOLDER", body);
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectUnresolved(o.pivots.caches[0].resolution, .unbounded_body, &.{});
    }
    // A body the parser refuses (an external workbook reference beside
    // a local one) still proves the local sheets it names (Codex r4 F1).
    for ([_][]const u8{
        "OFFSET(Data!$A$1,0,0,MAX(1,'[Book.xlsx]Sheet1'!$A$1),1)",
        "OFFSET('Data'!$A$1,0,0,[Book.xlsx]Sheet1!$A$1,1)",
        "OFFSET(data!$A$1,0,0,[Book.xlsx]Sheet1!$A$1,1)",
    }) |body| {
        try fixture.write(testing.allocator, io, path, .defined_name);
        try patchPart(io, path, "xl/workbook.xml", "Data!$A$1:$C$4", body);
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectUnresolved(o.pivots.caches[0].resolution, .unbounded_body, &.{0});
    }
    // A refused body still proves a 3D span's members (Codex r6 F1).
    for ([_][]const u8{
        "OFFSET(Report!$A$1,0,0,MAX(1,SUM(Data:Report!$A$1),'[Book.xlsx]Sheet1'!$A$1),1)",
        "OFFSET(Report!$A$1,0,0,MAX(1,SUM('Data:Report'!$A$1),'[Book.xlsx]Sheet1'!$A$1),1)",
    }) |body| {
        try fixture.write(testing.allocator, io, path, .defined_name);
        try patchPart(io, path, "xl/workbook.xml", "Data!$A$1:$C$4", body);
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectUnresolved(o.pivots.caches[0].resolution, .unbounded_body, &.{ 0, 1 });
    }
    // A refused body still queues the names and tables it spells
    // (Codex r5 F1): `Anchor` reads `Report`; `SalesTbl` is on `Data`.
    try fixture.write(testing.allocator, io, path, .defined_name);
    try patchPart(io, path, "xl/workbook.xml", "<definedName name=\"PivotSrc\">Data!$A$1:$C$4</definedName>", "<definedName name=\"Anchor\">Report!$A$1</definedName><definedName name=\"PivotSrc\">OFFSET(Anchor,0,0,MAX(1,'[Book.xlsx]Sheet1'!$A$1),1)</definedName>");
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectUnresolved(o.pivots.caches[0].resolution, .unbounded_body, &.{1});
    }
    try fixture.write(testing.allocator, io, path, .table_name);
    try patchPart(io, path, "xl/workbook.xml", "<pivotCaches>", "<definedNames><definedName name=\"PivotSrc\">OFFSET(SalesTbl[Qty],0,0,[Book.xlsx]Sheet1!$A$1,1)</definedName></definedNames><pivotCaches>");
    try patchPart(io, path, def_part, "name=\"SalesTbl\"", "name=\"PivotSrc\"");
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectUnresolved(o.pivots.caches[0].resolution, .unbounded_body, &.{0});
    }
    // `INDIRECT` with a literal reads what the literal spells: a sheet
    // reference, or a name; an unqualified cell proves nothing
    // (Codex r5 F2).
    const Indirect = struct { body: []const u8, sheets: []const u32 };
    for ([_]Indirect{
        .{ .body = "OFFSET(INDIRECT(\"Report!$A$1\"),0,0,2,2)", .sheets = &.{1} },
        .{ .body = "OFFSET(INDIRECT(\"'Report'!A1\"),0,0,2,2)", .sheets = &.{1} },
        .{ .body = "OFFSET(indirect(\"Anchor\"),0,0,2,2)", .sheets = &.{1} },
        .{ .body = "OFFSET(INDIRECT((\"Report!$A$1\")),0,0,2,2)", .sheets = &.{1} },
        .{ .body = "OFFSET(_xlfn.INDIRECT(\"Report!$A$1\"),0,0,2,2)", .sheets = &.{1} },
        .{ .body = "OFFSET(INDIRECT(\"A1\"),0,0,2,2)", .sheets = &.{} },
        .{ .body = "OFFSET(INDIRECT(\"Nope!A1\"),0,0,2,2)", .sheets = &.{} },
    }) |case| {
        try fixture.write(testing.allocator, io, path, .defined_name);
        try patchPart(io, path, "xl/workbook.xml", "<definedName name=\"PivotSrc\">Data!$A$1:$C$4</definedName>", "<definedName name=\"Anchor\">Report!$B$2</definedName><definedName name=\"PivotSrc\">PLACEHOLDER</definedName>");
        try patchPart(io, path, "xl/workbook.xml", "PLACEHOLDER", case.body);
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectUnresolved(o.pivots.caches[0].resolution, .unbounded_body, case.sheets);
    }
    // A sheet name longer than any fixed buffer is still looked up
    // (Codex r4 F3): the workbook accepts it, so must the walk.
    {
        const long = "L" ** 300;
        try fixture.write(testing.allocator, io, path, .defined_name);
        try patchPart(io, path, "xl/workbook.xml", "<sheet name=\"Data\"", "<sheet name=\"" ++ long ++ "\"");
        try patchPart(io, path, "xl/workbook.xml", "Data!$A$1:$C$4", "OFFSET('" ++ long ++ "'!$A$1,0,0,1,1)");
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectUnresolved(o.pivots.caches[0].resolution, .unbounded_body, &.{0});
    }
    // A value-position name resolves as written: `_xlfn.Anchor` is not
    // `Anchor` (Codex r1 F5).
    try fixture.write(testing.allocator, io, path, .defined_name);
    try patchPart(io, path, "xl/workbook.xml", "<definedName name=\"PivotSrc\">Data!$A$1:$C$4</definedName>", "<definedName name=\"Anchor\">Report!$D$1</definedName><definedName name=\"PivotSrc\">OFFSET(_xlfn.Anchor,0,0,1,1)</definedName>");
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectUnresolved(o.pivots.caches[0].resolution, .unbounded_body, &.{});
    }
    try fixture.write(testing.allocator, io, path, .defined_name);
    try patchPart(io, path, "xl/workbook.xml", "Data!$A$1:$C$4", "(Report!$A$1:$B$2,Data!$A$1:$C$4)");
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectUnresolved(o.pivots.caches[0].resolution, .unbounded_body, &.{ 0, 1 });
    }
    // A cycle terminates and proves nothing.
    try fixture.write(testing.allocator, io, path, .defined_name);
    try patchPart(io, path, "xl/workbook.xml", "<definedName name=\"PivotSrc\">Data!$A$1:$C$4</definedName>", "<definedName name=\"Loop\">PivotSrc*2</definedName><definedName name=\"PivotSrc\">Loop+1</definedName>");
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectUnresolved(o.pivots.caches[0].resolution, .unbounded_body, &.{});
    }
    // A structured reference proves the table's host.
    try fixture.write(testing.allocator, io, path, .table_name);
    try patchPart(io, path, "xl/workbook.xml", "<pivotCaches>", "<definedNames><definedName name=\"PivotSrc\">SUM(SalesTbl[Qty])</definedName></definedNames><pivotCaches>");
    try patchPart(io, path, def_part, "name=\"SalesTbl\"", "name=\"PivotSrc\"");
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectUnresolved(o.pivots.caches[0].resolution, .unbounded_body, &.{0});
    }
    // A name nobody defined.
    try fixture.write(testing.allocator, io, path, .defined_name);
    try patchPart(io, path, def_part, "name=\"PivotSrc\"", "name=\"Nope\"");
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectUnresolved(o.pivots.caches[0].resolution, .dangling_name, &.{});
    }
    // An `r:id` the relationships cannot place: the `sheet` beside it
    // is the evidence — a local one proves that sheet, a dangling one
    // proves nothing.
    try fixture.write(testing.allocator, io, path, .sheet_ref);
    try patchPart(io, path, def_part, "<worksheetSource sheet=\"Data\"", "<worksheetSource r:id=\"rIdNope\" sheet=\"Data\"");
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectUnresolved(o.pivots.caches[0].resolution, .unplaceable_rid, &.{0});
        try testing.expect(o.pivots.dependsOnSheet(0) and !o.pivots.dependsOnSheet(1));
    }
    try fixture.write(testing.allocator, io, path, .dangling);
    try patchPart(io, path, def_part, "<worksheetSource sheet=\"Nope\"", "<worksheetSource r:id=\"rIdNope\" sheet=\"Nope\"");
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectUnresolved(o.pivots.caches[0].resolution, .unplaceable_rid, &.{});
    }
    // A `ref` on no sheet; no locator at all.
    try fixture.write(testing.allocator, io, path, .sheet_ref);
    try patchPart(io, path, def_part, "<worksheetSource sheet=\"Data\" ref=\"A1:C4\"/>", "<worksheetSource ref=\"A1:C4\"/>");
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectUnresolved(o.pivots.caches[0].resolution, .sheetless_ref, &.{});
    }
    try fixture.write(testing.allocator, io, path, .sheet_ref);
    try patchPart(io, path, def_part, "<worksheetSource sheet=\"Data\" ref=\"A1:C4\"/>", "<worksheetSource/>");
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectUnresolved(o.pivots.caches[0].resolution, .no_locator, &.{});
    }
    // An external source proves nothing local: `dependsOnSheet` is false everywhere.
    try fixture.write(testing.allocator, io, path, .external);
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try testing.expect(o.pivots.caches[0].resolution == .external);
        try testing.expect(!o.pivots.dependsOnSheet(0) and !o.pivots.dependsOnSheet(1));
    }
}

test "collect: the consolidation fixture — one resolution per range set, each bounded, both sheets depended on" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "pivot_consolidation.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .consolidation);

    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    try expectFixtureShape(&o);
    const c = o.pivots.caches[0];
    try testing.expect(c.definition.source.type == .consolidation);
    try testing.expect(c.resolution == .none);
    try testing.expectEqual(@as(usize, 2), c.range_set_resolutions.len);
    try testing.expectEqualStrings("Data", c.range_set_sources[0].sheet.?);
    try testing.expectEqualStrings("A1:C4", c.range_set_sources[0].ref.?);
    try testing.expectEqualStrings("PivotSrc", c.range_set_sources[1].name.?);
    const s0 = c.range_set_resolutions[0].sheet;
    try testing.expectEqual(@as(u32, 0), s0.sheet_idx);
    try testing.expectEqual(ResolvedVia.sheet_attr, s0.via);
    try expectA1(s0.bounds, "A1:C4");
    const s1 = c.range_set_resolutions[1].sheet;
    try testing.expectEqual(@as(u32, 1), s1.sheet_idx);
    try testing.expectEqual(ResolvedVia.defined_name, s1.via);
    try expectA1(s1.bounds, "A1:B2");
    // Each set's `ref` span points at the live bytes, for the splice.
    const span = c.definition.source.range_sets[0].ref_span.?;
    try testing.expectEqualStrings("A1:C4", c.raw_xml[span.start..span.end]);
    try testing.expect(c.definition.source.range_sets[1].ref_span == null);
    try testing.expect(o.pivots.dependsOnSheet(0) and o.pivots.dependsOnSheet(1));
    try testing.expect(o.pivots.readsFromSheet(0) and o.pivots.readsFromSheet(1));
    try testing.expect(o.pivots.hostsPivot(1) and !o.pivots.hostsPivot(0));
}

test "edit: the splice lands on the parser's span — past a decoy, through an entity" {
    const decoy = "<!-- <location ref=\"Z9\"/> --><location ref=\"A3&#58;B6\"/>";
    try expectMove(decoy, .row, 1, .insert, "<!-- <location ref=\"Z9\"/> --><location ref=\"A4:B7\"/>");
    // A no-op keeps the entity spelling: the part is byte-preserved.
    try expectMove(decoy, .row, 9, .insert, decoy);
}

test "edit: a single-cell rectangle and the grid edges" {
    try expectMove("<location ref=\"A3\"/>", .row, 1, .insert, "<location ref=\"A4\"/>");
    try expectMove("<location ref=\"A3\"/>", .col, 1, .insert, "<location ref=\"B3\"/>");
    try expectMove("<location ref=\"A1048576\"/>", .row, 1, .delete, "<location ref=\"A1048575\"/>");
    try expectRefusal("<location ref=\"A1048576\"/>", .row, 1, .insert, error.PivotCoordinateOverflow);
    try expectRefusal("<location ref=\"A1:XFD1\"/>", .col, 1, .insert, error.PivotCoordinateOverflow);
    try expectMove("<location ref=\"XFD1\"/>", .col, 1, .delete, "<location ref=\"XFC1\"/>");
}

test "edit: a ref that is not an A1 rectangle refuses as malformed" {
    for ([_][]const u8{ "$A$3:B6", "A3:", ":B6", "B6:A3", "Sheet!A3:B6", "A03", "a3", "A3:B6:C7", "A0", "XFE1", " A3" }) |bad| {
        const middle = try std.fmt.allocPrint(testing.allocator, "<location ref=\"{s}\"/>", .{bad});
        defer testing.allocator.free(middle);
        try expectRefusal(middle, .row, 1, .insert, error.MalformedPivotXml);
    }
    // No `<location>` at all — the parser's refusal, surfaced whole.
    try expectRefusal("", .row, 1, .insert, error.MalformedPivotXml);
    // Row / column 0 is not a position — on the splice and on the
    // public rectangle helper alike (Codex #200 r4 REL-044).
    try expectRefusal("<location ref=\"A3\"/>", .row, 0, .insert, error.MalformedPivotXml);
    const fp: edit.Footprint = .{ .rect = .{ .tl_col = 1, .tl_row = 1, .br_col = 1, .br_row = 1 }, .first_row = 1, .last_col = 1 };
    try testing.expectError(error.MalformedPivotXml, edit.shiftRect(fp, .row, 0, .delete));
    try testing.expectError(error.MalformedPivotXml, edit.shiftRect(fp, .col, 0, .delete));
    try testing.expectError(error.MalformedPivotXml, edit.shiftRect(fp, .row, 0, .insert));
    try testing.expectEqual(@as(u32, 2), (try edit.shiftRect(fp, .row, 1, .insert)).tl_row);
}

fn editForFailures(allocator: Allocator, src: []const u8) !void {
    const out = try edit.applyToTableDefinition(allocator, src, .row, 1, .insert);
    allocator.free(out);
}

test "edit: allocation failure at every point leaves nothing behind" {
    const src = try ptWith("<location ref=\"A3:B6\" rowPageCount=\"1\"/><pageFields count=\"1\"><pageField fld=\"0\"/></pageFields>");
    defer testing.allocator.free(src);
    try testing.checkAllAllocationFailures(testing.allocator, editForFailures, .{src});
}

// ─── S7b: the source range ───────────────────────────────────────────

fn testRect(tl_col: u32, tl_row: u32, br_col: u32, br_row: u32) edit.Rect {
    return .{ .tl_col = tl_col, .tl_row = tl_row, .br_col = br_col, .br_row = br_row };
}

fn expectSourceMove(r: edit.Rect, axis: edit.Axis, idx: u32, kind: edit.Kind, want: edit.Rect) !void {
    const got = try edit.shiftSourceRect(r, axis, idx, kind);
    if (!got.eql(want)) {
        std.debug.print("want {any}, got {any}\n", .{ want, got });
        return error.TestExpectedEqual;
    }
}

test "edit: a source rectangle follows range semantics on the row axis" {
    const src = testRect(1, 1, 3, 4); // A1:C4 — a header row and three records.
    // Above or at the top row: a pure shift.
    try expectSourceMove(src, .row, 1, .insert, testRect(1, 2, 3, 5));
    // Inside: the bottom edge alone (one blank record, a content change).
    try expectSourceMove(src, .row, 2, .insert, testRect(1, 1, 3, 5));
    try expectSourceMove(src, .row, 4, .insert, testRect(1, 1, 3, 5));
    // Below: nothing.
    try expectSourceMove(src, .row, 5, .insert, src);
    // The header row feeds the field names — refused, as the table
    // rewriter refuses it for a headered table.
    try testing.expectError(error.PivotSourceEditUnsafe, edit.shiftSourceRect(src, .row, 1, .delete));
    try expectSourceMove(src, .row, 2, .delete, testRect(1, 1, 3, 3));
    try expectSourceMove(src, .row, 4, .delete, testRect(1, 1, 3, 3));
    try expectSourceMove(src, .row, 5, .delete, src);
    // A range that starts lower shifts whole under an edit above it.
    const low = testRect(2, 3, 4, 6); // B3:D6
    try expectSourceMove(low, .row, 1, .delete, testRect(2, 2, 4, 5));
    try expectSourceMove(low, .row, 3, .insert, testRect(2, 4, 4, 7));
    try expectSourceMove(low, .row, 2, .delete, testRect(2, 2, 4, 5));
    // A one-row source collapses under its only delete — the same
    // row is also its header, and both refuse.
    try testing.expectError(error.PivotSourceEditUnsafe, edit.shiftSourceRect(testRect(1, 1, 3, 1), .row, 1, .delete));
    // Overflow: an insert inside or above cannot grow past the grid.
    const bottom = testRect(1, 1, 3, zlsx.max_row);
    try testing.expectError(error.PivotCoordinateOverflow, edit.shiftSourceRect(bottom, .row, 1, .insert));
    try testing.expectError(error.PivotCoordinateOverflow, edit.shiftSourceRect(bottom, .row, 5, .insert));
    try expectSourceMove(bottom, .row, 2, .delete, testRect(1, 1, 3, zlsx.max_row - 1));
    // Position 0 is not a position.
    try testing.expectError(error.MalformedPivotXml, edit.shiftSourceRect(src, .row, 0, .insert));
    try testing.expectError(error.MalformedPivotXml, edit.shiftSourceBounds(.{ .rect = src }, .col, 0, .delete));
}

test "edit: a column edit inside a source is the field schema (S7c) — refused; outside, the range shifts" {
    const src = testRect(2, 1, 4, 4); // B1:D4
    try expectSourceMove(src, .col, 1, .insert, testRect(3, 1, 5, 4));
    try expectSourceMove(src, .col, 2, .insert, testRect(3, 1, 5, 4));
    try testing.expectError(error.PivotSourceEditUnsafe, edit.shiftSourceRect(src, .col, 3, .insert));
    try testing.expectError(error.PivotSourceEditUnsafe, edit.shiftSourceRect(src, .col, 4, .insert));
    try expectSourceMove(src, .col, 5, .insert, src);
    try expectSourceMove(src, .col, 1, .delete, testRect(1, 1, 3, 4));
    try testing.expectError(error.PivotSourceEditUnsafe, edit.shiftSourceRect(src, .col, 2, .delete));
    try testing.expectError(error.PivotSourceEditUnsafe, edit.shiftSourceRect(src, .col, 4, .delete));
    try expectSourceMove(src, .col, 5, .delete, src);
    const right = testRect(zlsx.max_col_1based - 1, 1, zlsx.max_col_1based, 4);
    try testing.expectError(error.PivotCoordinateOverflow, edit.shiftSourceRect(right, .col, 1, .insert));
    try expectSourceMove(right, .col, 1, .delete, testRect(zlsx.max_col_1based - 2, 1, zlsx.max_col_1based - 1, 4));
}

test "edit: whole columns and whole rows have no coordinate to move, and refuse what a rectangle would" {
    const cols: Bounds = .{ .whole_columns = .{ .first_col = 1, .last_col = 3 } };
    // Row 1 is the header of a whole-column source: an insert blanks
    // it, a delete promotes a record to the field names.
    try testing.expectError(error.PivotSourceEditUnsafe, edit.shiftSourceBounds(cols, .row, 1, .insert));
    try testing.expectError(error.PivotSourceEditUnsafe, edit.shiftSourceBounds(cols, .row, 1, .delete));
    try testing.expect((try edit.shiftSourceBounds(cols, .row, 2, .insert)) == null);
    try testing.expect((try edit.shiftSourceBounds(cols, .row, 9, .delete)) == null);
    // On the column axis the columns move as a rectangle's would.
    try expectBoundsA1(try edit.shiftSourceBounds(cols, .col, 1, .insert), "B:D");
    try testing.expectError(error.PivotSourceEditUnsafe, edit.shiftSourceBounds(cols, .col, 2, .insert));
    try testing.expectError(error.PivotSourceEditUnsafe, edit.shiftSourceBounds(cols, .col, 1, .delete));
    try testing.expectError(error.PivotSourceEditUnsafe, edit.shiftSourceBounds(cols, .col, 3, .delete));
    try testing.expect((try edit.shiftSourceBounds(cols, .col, 4, .insert)) == null);
    try testing.expect((try edit.shiftSourceBounds(cols, .col, 4, .delete)) == null);
    const right: Bounds = .{ .whole_columns = .{ .first_col = 2, .last_col = zlsx.max_col_1based } };
    try testing.expectError(error.PivotCoordinateOverflow, edit.shiftSourceBounds(right, .col, 1, .insert));
    try expectBoundsA1(try edit.shiftSourceBounds(right, .col, 1, .delete), "A:XFC");

    const rows: Bounds = .{ .whole_rows = .{ .first_row = 2, .last_row = 4 } };
    try testing.expectError(error.PivotSourceEditUnsafe, edit.shiftSourceBounds(rows, .row, 2, .delete));
    try expectBoundsA1(try edit.shiftSourceBounds(rows, .row, 1, .insert), "3:5");
    try expectBoundsA1(try edit.shiftSourceBounds(rows, .row, 2, .insert), "3:5");
    try expectBoundsA1(try edit.shiftSourceBounds(rows, .row, 3, .insert), "2:5");
    try expectBoundsA1(try edit.shiftSourceBounds(rows, .row, 1, .delete), "1:3");
    try expectBoundsA1(try edit.shiftSourceBounds(rows, .row, 3, .delete), "2:3");
    try testing.expect((try edit.shiftSourceBounds(rows, .row, 5, .insert)) == null);
    // Every column is inside a whole-row source.
    try testing.expectError(error.PivotSourceEditUnsafe, edit.shiftSourceBounds(rows, .col, 1, .insert));
    try testing.expectError(error.PivotSourceEditUnsafe, edit.shiftSourceBounds(rows, .col, 500, .delete));

    const r: Bounds = .{ .rect = testRect(1, 1, 3, 4) };
    try expectBoundsA1(try edit.shiftSourceBounds(r, .row, 2, .insert), "A1:C5");
    try testing.expect((try edit.shiftSourceBounds(r, .row, 5, .insert)) == null);
}

fn expectBoundsA1(b: ?Bounds, want: []const u8) !void {
    var buf: [Bounds.format_buf_len]u8 = undefined;
    const got = (b orelse return error.TestExpectedBounds).formatA1(&buf) orelse return error.TestExpectedBounds;
    try testing.expectEqualStrings(want, got);
}

test "edit: a direct whole-row or whole-column ref is respelled like its rectangle counterpart (Codex r2 REL-201)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const def_part = "xl/pivotCache/pivotCacheDefinition1.xml";
    const rows = try tt.path(testing.allocator, io, "s7b_edit_whole_rows.xlsx");
    defer testing.allocator.free(rows);
    try fixture.write(testing.allocator, io, rows, .sheet_ref);
    try fixture.patchPart(testing.allocator, io, rows, def_part, "ref=\"A1:C4\"", "ref=\"1:4\"");
    {
        var o = try Opened.open(testing.allocator, io, rows);
        defer o.deinit(testing.allocator);
        try expectA1(o.pivots.caches[0].resolution.sheet.bounds, "1:4");
        try expectCacheSource(testing.allocator, &o, 0, .row, 1, .insert, "<worksheetSource sheet=\"Data\" ref=\"2:5\"/>", false);
        try expectCacheSource(testing.allocator, &o, 0, .row, 3, .insert, "<worksheetSource sheet=\"Data\" ref=\"1:5\"/>", true);
        try expectCacheSource(testing.allocator, &o, 0, .row, 4, .delete, "<worksheetSource sheet=\"Data\" ref=\"1:3\"/>", true);
        try testing.expectError(error.PivotSourceEditUnsafe, cacheEdit(testing.allocator, &o, 0, .row, 1, .delete));
        try testing.expectError(error.PivotSourceEditUnsafe, cacheEdit(testing.allocator, &o, 0, .col, 1, .insert));
        try testing.expect((try cacheEdit(testing.allocator, &o, 0, .row, 5, .insert)) == null);
    }
    const cols = try tt.path(testing.allocator, io, "s7b_edit_whole_cols.xlsx");
    defer testing.allocator.free(cols);
    try fixture.write(testing.allocator, io, cols, .sheet_ref);
    try fixture.patchPart(testing.allocator, io, cols, def_part, "ref=\"A1:C4\"", "ref=\"B:D\"");
    var o = try Opened.open(testing.allocator, io, cols);
    defer o.deinit(testing.allocator);
    try expectA1(o.pivots.caches[0].resolution.sheet.bounds, "B:D");
    try expectCacheSource(testing.allocator, &o, 0, .col, 1, .insert, "<worksheetSource sheet=\"Data\" ref=\"C:E\"/>", false);
    try expectCacheSource(testing.allocator, &o, 0, .col, 2, .insert, "<worksheetSource sheet=\"Data\" ref=\"C:E\"/>", false);
    try expectCacheSource(testing.allocator, &o, 0, .col, 1, .delete, "<worksheetSource sheet=\"Data\" ref=\"A:C\"/>", false);
    try testing.expectError(error.PivotSourceEditUnsafe, cacheEdit(testing.allocator, &o, 0, .col, 3, .insert));
    try testing.expectError(error.PivotSourceEditUnsafe, cacheEdit(testing.allocator, &o, 0, .col, 4, .delete));
    try testing.expect((try cacheEdit(testing.allocator, &o, 0, .col, 5, .insert)) == null);
    // Row 1 is the header; every other row is inside the columns —
    // nothing in the spelling moves, and the content changed: marked.
    try testing.expectError(error.PivotSourceEditUnsafe, cacheEdit(testing.allocator, &o, 0, .row, 1, .insert));
    try testing.expectError(error.PivotSourceEditUnsafe, cacheEdit(testing.allocator, &o, 0, .row, 1, .delete));
    try expectMarkedOnly(testing.allocator, &o, 0, .row, 2, .delete);
    try expectMarkedOnly(testing.allocator, &o, 0, .row, 9, .insert);
}

/// The fixture's cache definition under one edit on one sheet: the
/// rewritten part, or null when nothing in it moves.
fn cacheEdit(alloc: Allocator, o: *const Opened, sheet_idx: u32, axis: edit.Axis, idx: u32, kind: edit.Kind) !?[]u8 {
    return edit.applyToCacheDefinition(alloc, &o.pivots.caches[0], sheet_idx, axis, idx, kind);
}

/// The rewritten part holds `want_source` where the source element was,
/// and its root is the original's — plus the refresh marker, inserted
/// before the root's `>`, exactly when `marked`. Nothing else differs.
fn expectCacheSource(alloc: Allocator, o: *const Opened, sheet_idx: u32, axis: edit.Axis, idx: u32, kind: edit.Kind, want_source: []const u8, marked: bool) !void {
    const out = (try cacheEdit(alloc, o, sheet_idx, axis, idx, kind)) orelse return error.TestExpectedMove;
    defer alloc.free(out);
    const src = o.pivots.caches[0].raw_xml;
    const at = std.mem.indexOf(u8, out, want_source) orelse {
        std.debug.print("want {s}\nin   {s}\n", .{ want_source, out });
        return error.TestExpectedMove;
    };
    const old_open = std.mem.indexOf(u8, src, "<cacheSource").?;
    const want_head = if (marked)
        try markedPart(alloc, src[0..old_open])
    else
        try alloc.dupe(u8, src[0..old_open]);
    defer alloc.free(want_head);
    try testing.expectEqualStrings(want_head, out[0..want_head.len]);
    const tail_from = std.mem.indexOf(u8, src, "<cacheFields").?;
    try testing.expectEqualStrings(src[tail_from..], out[at + want_source.len ..][std.mem.indexOf(u8, out[at + want_source.len ..], "<cacheFields").?..]);
}

/// The edit changes nothing in the part but the marker: the rewrite is
/// the original with ` refreshOnLoad="1"` inserted on the root and not
/// one other byte — a table-named or defined-name source under an edit
/// inside its rectangle, an unbounded source under any edit of a sheet
/// it references.
fn expectMarkedOnly(alloc: Allocator, o: *const Opened, sheet_idx: u32, axis: edit.Axis, idx: u32, kind: edit.Kind) !void {
    const out = (try cacheEdit(alloc, o, sheet_idx, axis, idx, kind)) orelse return error.TestExpectedMark;
    defer alloc.free(out);
    const want = try markedPart(alloc, o.pivots.caches[0].raw_xml);
    defer alloc.free(want);
    try testing.expectEqualStrings(want, out);
}

/// `bytes` (a definition, or its prefix up to `<cacheSource`) with
/// ` refreshOnLoad="1"` inserted before the root open tag's `>` — what
/// a marked part must read, spelled independently of the splice under
/// test. The fixture's root carries no `>` inside an attribute value.
fn markedPart(alloc: Allocator, bytes: []const u8) ![]u8 {
    const root = std.mem.indexOf(u8, bytes, "<pivotCacheDefinition") orelse return error.TestExpectedMark;
    const gt = std.mem.indexOfScalarPos(u8, bytes, root, '>') orelse return error.TestExpectedMark;
    return std.mem.concat(alloc, u8, &.{ bytes[0..gt], edit.marker_insert, bytes[gt..] });
}

test "edit: a sheet+ref source is respelled at the parser's span; another sheet's edit leaves the part alone" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "s7b_edit_sheet_ref.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .sheet_ref);
    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    try testing.expectEqualStrings("xl/worksheets/sheet1.xml", o.pivots.sheet_parts[0]);
    try testing.expectEqual(@as(?u32, 0), o.pivots.sheetIndexOfPart("xl/worksheets/sheet1.xml"));
    try testing.expectEqual(@as(?u32, null), o.pivots.sheetIndexOfPart("xl/worksheets/sheet9.xml"));
    try testing.expectEqual(SourceCarrier.ref, o.pivots.caches[0].resolution.sheet.carrier);

    // `Data` (0) feeds `A1:C4`.
    try expectCacheSource(testing.allocator, &o, 0, .row, 2, .insert, "<worksheetSource sheet=\"Data\" ref=\"A1:C5\"/>", true);
    try expectCacheSource(testing.allocator, &o, 0, .row, 1, .insert, "<worksheetSource sheet=\"Data\" ref=\"A2:C5\"/>", false);
    try expectCacheSource(testing.allocator, &o, 0, .row, 3, .delete, "<worksheetSource sheet=\"Data\" ref=\"A1:C3\"/>", true);
    try expectCacheSource(testing.allocator, &o, 0, .col, 1, .insert, "<worksheetSource sheet=\"Data\" ref=\"B1:D4\"/>", false);
    try testing.expect((try cacheEdit(testing.allocator, &o, 0, .row, 5, .insert)) == null);
    try testing.expect((try cacheEdit(testing.allocator, &o, 0, .col, 4, .delete)) == null);
    try testing.expectError(error.PivotSourceEditUnsafe, cacheEdit(testing.allocator, &o, 0, .row, 1, .delete));
    try testing.expectError(error.PivotSourceEditUnsafe, cacheEdit(testing.allocator, &o, 0, .col, 2, .insert));
    try testing.expectError(error.MalformedPivotXml, cacheEdit(testing.allocator, &o, 0, .row, 0, .insert));
    // `Report` (1) hosts and does not feed: nothing to move.
    try testing.expect((try cacheEdit(testing.allocator, &o, 1, .row, 1, .insert)) == null);
    try testing.expect((try cacheEdit(testing.allocator, &o, 1, .row, 1, .delete)) == null);
}

test "edit: a consolidation definition splices each set at its own span" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "s7b_edit_consolidation.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .consolidation);
    {
        // The fixture's second set is name-spelled; two direct sets on
        // one sheet are what the ordering is for.
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        // Set 1 reads `Report!A1:B2` through `PivotSrc`: a `Report` row
        // edit judges its rectangle (the header row refuses) and
        // moves nothing here — the name's body is the name sweep's;
        // a delete inside drops a record and marks the definition.
        try testing.expectError(error.PivotSourceEditUnsafe, cacheEdit(testing.allocator, &o, 1, .row, 1, .delete));
        try testing.expect((try cacheEdit(testing.allocator, &o, 1, .row, 1, .insert)) == null);
        try expectMarkedOnly(testing.allocator, &o, 1, .row, 2, .delete);
        // Set 0 is `Data!A1:C4` by `sheet` + `ref`.
        try expectCacheSource(testing.allocator, &o, 0, .row, 1, .insert, "<rangeSet i1=\"0\" sheet=\"Data\" ref=\"A2:C5\"/><rangeSet i1=\"1\" name=\"PivotSrc\"/>", false);
    }
    try fixture.patchPart(testing.allocator, io, path, "xl/pivotCache/pivotCacheDefinition1.xml", "<rangeSet i1=\"1\" name=\"PivotSrc\"/>", "<rangeSet i1=\"1\" sheet=\"Data\" ref=\"A6:C9\"/>");
    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    try expectCacheSource(testing.allocator, &o, 0, .row, 1, .insert, "<rangeSet i1=\"0\" sheet=\"Data\" ref=\"A2:C5\"/><rangeSet i1=\"1\" sheet=\"Data\" ref=\"A7:C10\"/>", false);
    try expectCacheSource(testing.allocator, &o, 0, .row, 5, .insert, "<rangeSet i1=\"0\" sheet=\"Data\" ref=\"A1:C4\"/><rangeSet i1=\"1\" sheet=\"Data\" ref=\"A7:C10\"/>", false);
    try expectCacheSource(testing.allocator, &o, 0, .row, 3, .delete, "<rangeSet i1=\"0\" sheet=\"Data\" ref=\"A1:C3\"/><rangeSet i1=\"1\" sheet=\"Data\" ref=\"A5:C8\"/>", true);
    // One set's refusal is the definition's.
    try testing.expectError(error.PivotSourceEditUnsafe, cacheEdit(testing.allocator, &o, 0, .row, 6, .delete));
    try testing.expect((try cacheEdit(testing.allocator, &o, 0, .row, 10, .insert)) == null);
    try testing.expect(!o.pivots.dependsOnSheet(1));
}

test "edit: a table-named source moves with its table; a defined-name source with its body — both judged on their rectangle" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const table = try tt.path(testing.allocator, io, "s7b_edit_table.xlsx");
    defer testing.allocator.free(table);
    try fixture.write(testing.allocator, io, table, .table_name);
    {
        var o = try Opened.open(testing.allocator, io, table);
        defer o.deinit(testing.allocator);
        // The row axis is `table_edit`'s: header and collapse refuse
        // there, with the table's own `headerRowCount` knowledge — so
        // the top-row delete reaching here is the admitted headerless
        // case, a content change. Inside marks; at or above the top
        // row an insert is the table's shift and leaves the part alone.
        try expectMarkedOnly(testing.allocator, &o, 0, .row, 1, .delete);
        try expectMarkedOnly(testing.allocator, &o, 0, .row, 2, .insert);
        try testing.expect((try cacheEdit(testing.allocator, &o, 0, .row, 1, .insert)) == null);
        try testing.expect((try cacheEdit(testing.allocator, &o, 0, .row, 5, .insert)) == null);
        // The column axis inside the table is the field schema.
        try testing.expect((try cacheEdit(testing.allocator, &o, 0, .col, 1, .insert)) == null);
        try testing.expectError(error.PivotSourceEditUnsafe, cacheEdit(testing.allocator, &o, 0, .col, 2, .insert));
        try testing.expectError(error.PivotSourceEditUnsafe, cacheEdit(testing.allocator, &o, 0, .col, 3, .delete));
        try testing.expect((try cacheEdit(testing.allocator, &o, 0, .col, 4, .delete)) == null);
        try testing.expectEqual(@as(usize, 0), namesOf(o.pivots.caches[0].resolution).len);
        try testing.expectEqual(SourceCarrier.table, o.pivots.caches[0].resolution.sheet.carrier);
    }
    // `sheet` beside the table's name: placed by the attribute, bounded
    // by the table — and judged as a table (Codex #203 r1 REL-102).
    const beside = try tt.path(testing.allocator, io, "s7b_edit_sheet_table.xlsx");
    defer testing.allocator.free(beside);
    try fixture.write(testing.allocator, io, beside, .table_name);
    try fixture.patchPart(testing.allocator, io, beside, "xl/pivotCache/pivotCacheDefinition1.xml", "<worksheetSource name=\"SalesTbl\"/>", "<worksheetSource sheet=\"Data\" name=\"SalesTbl\"/>");
    {
        var o = try Opened.open(testing.allocator, io, beside);
        defer o.deinit(testing.allocator);
        const s = o.pivots.caches[0].resolution.sheet;
        try testing.expectEqual(ResolvedVia.sheet_attr, s.via);
        try testing.expectEqual(SourceCarrier.table, s.carrier);
        try expectA1(s.bounds, "A1:C4");
        try expectMarkedOnly(testing.allocator, &o, 0, .row, 1, .delete);
        try testing.expectError(error.PivotSourceEditUnsafe, cacheEdit(testing.allocator, &o, 0, .col, 2, .insert));
    }
    // The same spelling naming a table on ANOTHER sheet bounds nothing.
    try fixture.patchPart(testing.allocator, io, beside, "xl/pivotCache/pivotCacheDefinition1.xml", "sheet=\"Data\" name=", "sheet=\"Report\" name=");
    {
        var o = try Opened.open(testing.allocator, io, beside);
        defer o.deinit(testing.allocator);
        const s = o.pivots.caches[0].resolution.sheet;
        try testing.expectEqual(@as(u32, 1), s.sheet_idx);
        try testing.expectEqual(SourceCarrier.none, s.carrier);
        try testing.expect(s.bounds == null);
        try testing.expectError(error.PivotSourceEditUnsafe, cacheEdit(testing.allocator, &o, 1, .row, 9, .insert));
        try testing.expect((try cacheEdit(testing.allocator, &o, 0, .row, 1, .delete)) == null);
    }
    const name = try tt.path(testing.allocator, io, "s7b_edit_name.xlsx");
    defer testing.allocator.free(name);
    try fixture.write(testing.allocator, io, name, .defined_name);
    var o = try Opened.open(testing.allocator, io, name);
    defer o.deinit(testing.allocator);
    try testing.expectError(error.PivotSourceEditUnsafe, cacheEdit(testing.allocator, &o, 0, .row, 1, .delete));
    try testing.expectError(error.PivotSourceEditUnsafe, cacheEdit(testing.allocator, &o, 0, .col, 2, .insert));
    // Inside `Data!$A$1:$C$4`: the body is the name sweep's, the
    // content change is the marker's. Above: a shift, nothing.
    try expectMarkedOnly(testing.allocator, &o, 0, .row, 2, .insert);
    try expectMarkedOnly(testing.allocator, &o, 0, .row, 4, .delete);
    try testing.expect((try cacheEdit(testing.allocator, &o, 0, .row, 1, .insert)) == null);
    try testing.expect((try cacheEdit(testing.allocator, &o, 0, .col, 1, .insert)) == null);
    try testing.expectEqual(SourceCarrier.defined_name, o.pivots.caches[0].resolution.sheet.carrier);
    // The names the sweep's dry-run judges: the source name itself.
    const keys = namesOf(o.pivots.caches[0].resolution);
    try testing.expectEqual(@as(usize, 1), keys.len);
    try testing.expectEqualStrings("PivotSrc", keys[0].identifier);
    try testing.expectEqual(@as(?u32, null), keys[0].scope);
}

test "collect: a name-spelled source carries its closure's names — the root and every name it reaches, once each" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "s7b_closure_names.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .defined_name);
    // `PivotSrc` reaches `Report` through `Anchor`, which cites
    // `PivotSrc` back — a cycle. (A sheet-scoped `Anchor` would be
    // invisible from the workbook-scoped root — Codex #202 r2 F1 — and
    // the closure rightly empty; the scope rules have their own tests.)
    try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "<definedName name=\"PivotSrc\">Data!$A$1:$C$4</definedName>", "<definedName name=\"Anchor\">Report!$D$1+ROWS(PivotSrc)</definedName><definedName name=\"PivotSrc\">OFFSET(Anchor,0,0,4,3)</definedName>");
    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    const u = o.pivots.caches[0].resolution.unresolved;
    try testing.expectEqual(Unresolved.Why.unbounded_body, u.why);
    try testing.expectEqualSlices(u32, &.{1}, u.sheets);
    try testing.expectEqual(@as(usize, 2), u.names.len);
    try testing.expectEqualStrings("PivotSrc", u.names[0].identifier);
    try testing.expectEqual(@as(?u32, null), u.names[0].scope);
    try testing.expectEqualStrings("Anchor", u.names[1].identifier);
    try testing.expectEqual(@as(?u32, null), u.names[1].scope);
    // Unbounded on `Report`: nothing to move, nothing refused here —
    // the body is the name sweep's, whose dry-run is the workbook's —
    // and no shift to prove, so any edit of `Report` marks, on either
    // axis. `Data` is not in the closure: untouched.
    try expectMarkedOnly(testing.allocator, &o, 1, .row, 1, .delete);
    try expectMarkedOnly(testing.allocator, &o, 1, .row, 9, .insert);
    try expectMarkedOnly(testing.allocator, &o, 1, .col, 1, .insert);
    try testing.expect((try cacheEdit(testing.allocator, &o, 0, .row, 1, .delete)) == null);
}

test "edit: the marker is an upsert — inserted on a root that lacks it, replaced in place on one that spells it 0, left alone once set" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const def_part = "xl/pivotCache/pivotCacheDefinition1.xml";
    const path = try tt.path(testing.allocator, io, "s7b_marker_upsert.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .sheet_ref);
    {
        // Absent (both corpus definitions): inserted before the root's
        // `>`, after the last attribute, beside the moved `ref`.
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try testing.expect(!edit.markerSet(&o.pivots.caches[0].definition));
        const out = (try cacheEdit(testing.allocator, &o, 0, .row, 2, .insert)) orelse return error.TestExpectedMark;
        defer testing.allocator.free(out);
        try testing.expect(std.mem.indexOf(u8, out, "recordCount=\"3\" refreshOnLoad=\"1\"><cacheSource") != null);
        try testing.expect(std.mem.indexOf(u8, out, "ref=\"A1:C5\"") != null);
        try testing.expectEqual(@as(usize, 1), std.mem.count(u8, out, "refreshOnLoad"));
        // The same part reparses, and the reader sees the option set.
        var def = try pivot_xml.parseCacheDefinition(testing.allocator, out);
        defer def.deinit(testing.allocator);
        try testing.expect(def.refresh_on_load);
        // The save-time write on an unmarked part: the marker alone.
        const marked = (try edit.markForRefresh(testing.allocator, &o.pivots.caches[0])) orelse return error.TestExpectedMark;
        defer testing.allocator.free(marked);
        const want = try markedPart(testing.allocator, o.pivots.caches[0].raw_xml);
        defer testing.allocator.free(want);
        try testing.expectEqualStrings(want, marked);
    }
    // Present and off, single-quoted, with whitespace around `=`: the
    // value is replaced where it sits and nothing is inserted.
    try fixture.patchPart(testing.allocator, io, path, def_part, " recordCount=\"3\">", " refreshOnLoad = '0' recordCount=\"3\" >");
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try testing.expect(!edit.markerSet(&o.pivots.caches[0].definition));
        const out = (try cacheEdit(testing.allocator, &o, 0, .row, 2, .insert)) orelse return error.TestExpectedMark;
        defer testing.allocator.free(out);
        try testing.expect(std.mem.indexOf(u8, out, " refreshOnLoad = '1' recordCount=\"3\" ><cacheSource") != null);
        try testing.expectEqual(@as(usize, 1), std.mem.count(u8, out, "refreshOnLoad"));
        // A pure shift replaces nothing: `0` stays as the user set it.
        const shifted = (try cacheEdit(testing.allocator, &o, 0, .row, 1, .insert)) orelse return error.TestExpectedMove;
        defer testing.allocator.free(shifted);
        try testing.expect(std.mem.indexOf(u8, shifted, " refreshOnLoad = '0' ") != null);
    }
    // Present and on (`true` is the schema's other spelling): nothing to
    // write — the `ref` still moves, the root is byte-identical, and
    // the save-time write is a no-op.
    try fixture.patchPart(testing.allocator, io, path, def_part, " refreshOnLoad = '0' ", " refreshOnLoad=\"true\" ");
    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    try testing.expect(edit.markerSet(&o.pivots.caches[0].definition));
    try expectCacheSource(testing.allocator, &o, 0, .row, 2, .insert, "<worksheetSource sheet=\"Data\" ref=\"A1:C5\"/>", false);
    try testing.expect((try edit.markForRefresh(testing.allocator, &o.pivots.caches[0])) == null);
}

test "edit: the cell-write predicate — inside a rectangle, on a whole-column source's columns, anywhere on an unbounded source's sheet, nowhere for what proves no local range" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const def_part = "xl/pivotCache/pivotCacheDefinition1.xml";
    const path = try tt.path(testing.allocator, io, "s7b_marker_cells.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .sheet_ref);
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        const r = o.pivots.caches[0].resolution;
        // `Data!A1:C4`: the four corners are in, one past each edge is out.
        try testing.expect(edit.cellWriteChangesSource(r, 0, 1, 1));
        try testing.expect(edit.cellWriteChangesSource(r, 0, 4, 3));
        try testing.expect(edit.cellWriteChangesSource(r, 0, 2, 2));
        try testing.expect(!edit.cellWriteChangesSource(r, 0, 5, 1));
        try testing.expect(!edit.cellWriteChangesSource(r, 0, 1, 4));
        try testing.expect(!edit.cellWriteChangesSource(r, 1, 2, 2));
    }
    try fixture.patchPart(testing.allocator, io, path, def_part, "ref=\"A1:C4\"", "ref=\"B:C\"");
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        const r = o.pivots.caches[0].resolution;
        try testing.expect(edit.cellWriteChangesSource(r, 0, 100000, 2));
        try testing.expect(!edit.cellWriteChangesSource(r, 0, 1, 1));
        try testing.expect(!edit.cellWriteChangesSource(r, 0, 1, 4));
    }
    // `sheet` alone claims the whole sheet.
    try fixture.patchPart(testing.allocator, io, path, def_part, " ref=\"B:C\"", "");
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        const r = o.pivots.caches[0].resolution;
        try testing.expect(r.sheet.bounds == null);
        try testing.expect(edit.cellWriteChangesSource(r, 0, 7, 7));
        try testing.expect(!edit.cellWriteChangesSource(r, 1, 7, 7));
    }
    const ext = try tt.path(testing.allocator, io, "s7b_marker_cells_ext.xlsx");
    defer testing.allocator.free(ext);
    try fixture.write(testing.allocator, io, ext, .external);
    {
        var o = try Opened.open(testing.allocator, io, ext);
        defer o.deinit(testing.allocator);
        try testing.expect(!edit.cellWriteChangesSource(o.pivots.caches[0].resolution, 0, 1, 1));
        try testing.expect(!edit.cellWriteChangesSource(o.pivots.caches[0].resolution, 1, 1, 1));
    }
    // An unbounded body reaching `Report` (1) through `Anchor`: every
    // cell of `Report`, no cell of `Data`. An `r:id` the reader cannot
    // place proves no local range and never marks, though the sweep
    // refuses its `sheet` (Q4 i) — a mark would ask Excel to refresh
    // at open a source it may not have.
    const name = try tt.path(testing.allocator, io, "s7b_marker_cells_name.xlsx");
    defer testing.allocator.free(name);
    try fixture.write(testing.allocator, io, name, .defined_name);
    try fixture.patchPart(testing.allocator, io, name, "xl/workbook.xml", "<definedName name=\"PivotSrc\">Data!$A$1:$C$4</definedName>", "<definedName name=\"Anchor\">Report!$D$1</definedName><definedName name=\"PivotSrc\">OFFSET(Anchor,0,0,4,3)</definedName>");
    {
        var o = try Opened.open(testing.allocator, io, name);
        defer o.deinit(testing.allocator);
        const r = o.pivots.caches[0].resolution;
        try testing.expectEqual(Unresolved.Why.unbounded_body, r.unresolved.why);
        try testing.expect(edit.cellWriteChangesSource(r, 1, 1000, 50));
        try testing.expect(!edit.cellWriteChangesSource(r, 0, 1, 1));
    }
    try fixture.patchPart(testing.allocator, io, name, def_part, "<worksheetSource name=\"PivotSrc\"/>", "<worksheetSource r:id=\"rIdNone\" sheet=\"Data\" ref=\"A1:C4\"/>");
    var o = try Opened.open(testing.allocator, io, name);
    defer o.deinit(testing.allocator);
    const r = o.pivots.caches[0].resolution;
    try testing.expectEqual(Unresolved.Why.unplaceable_rid, r.unresolved.why);
    try testing.expectEqualSlices(u32, &.{0}, r.unresolved.sheets);
    try testing.expect(!edit.cellWriteChangesSource(r, 0, 2, 2));
    try testing.expectError(error.PivotSourceEditUnsafe, cacheEdit(testing.allocator, &o, 0, .row, 2, .insert));
}

test "collect: a locator under a source type the reader does not know is authoritative (S7b gate, Q5)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "s7b_unknown_type.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .sheet_ref);
    try fixture.patchPart(testing.allocator, io, path, "xl/pivotCache/pivotCacheDefinition1.xml", "<cacheSource type=\"worksheet\">", "<cacheSource type=\"zlsxFuture\">");
    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    const c = o.pivots.caches[0];
    try testing.expect(c.definition.source.type == .unknown);
    try testing.expectEqualStrings("Data", c.source.sheet.?);
    try testing.expectEqual(@as(u32, 0), c.resolution.sheet.sheet_idx);
    try expectA1(c.resolution.sheet.bounds, "A1:C4");
    try testing.expect(o.pivots.dependsOnSheet(0) and !o.pivots.dependsOnSheet(1));
    try expectCacheSource(testing.allocator, &o, 0, .row, 2, .insert, "<worksheetSource sheet=\"Data\" ref=\"A1:C5\"/>", true);
    // Without a locator the type names no sheet, as before.
    try fixture.patchPart(testing.allocator, io, path, "xl/pivotCache/pivotCacheDefinition1.xml", "<worksheetSource sheet=\"Data\" ref=\"A1:C4\"/>", "");
    var o2 = try Opened.open(testing.allocator, io, path);
    defer o2.deinit(testing.allocator);
    try testing.expect(o2.pivots.caches[0].resolution == .none);
    try testing.expect(!o2.pivots.dependsOnSheet(0) and !o2.pivots.dependsOnSheet(1));
}

test "edit: spellings that claim the edited sheet without a range refuse; the ones that claim nothing are left alone" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const def_part = "xl/pivotCache/pivotCacheDefinition1.xml";

    // `sheet` alone (Q4 iv): it claims `Data` and gives no rectangle.
    const only = try tt.path(testing.allocator, io, "s7b_sheet_only.xlsx");
    defer testing.allocator.free(only);
    try fixture.write(testing.allocator, io, only, .sheet_ref);
    try fixture.patchPart(testing.allocator, io, only, def_part, "<worksheetSource sheet=\"Data\" ref=\"A1:C4\"/>", "<worksheetSource sheet=\"Data\"/>");
    {
        var o = try Opened.open(testing.allocator, io, only);
        defer o.deinit(testing.allocator);
        try testing.expect(o.pivots.caches[0].resolution.sheet.bounds == null);
        try testing.expectError(error.PivotSourceEditUnsafe, cacheEdit(testing.allocator, &o, 0, .row, 9, .insert));
        try testing.expect((try cacheEdit(testing.allocator, &o, 1, .row, 1, .insert)) == null);
    }
    // A `ref` the bounds parser rejects is the same shape.
    const bad = try tt.path(testing.allocator, io, "s7b_bad_ref.xlsx");
    defer testing.allocator.free(bad);
    try fixture.write(testing.allocator, io, bad, .sheet_ref);
    try fixture.patchPart(testing.allocator, io, bad, def_part, "ref=\"A1:C4\"", "ref=\"A1:C4:E6\"");
    {
        var o = try Opened.open(testing.allocator, io, bad);
        defer o.deinit(testing.allocator);
        try testing.expect(o.pivots.caches[0].resolution.sheet.bounds == null);
        try testing.expectError(error.PivotSourceEditUnsafe, cacheEdit(testing.allocator, &o, 0, .row, 9, .insert));
    }
    // `ref` alone claims no sheet; a dangling sheet claims none either.
    const sheetless = try tt.path(testing.allocator, io, "s7b_sheetless.xlsx");
    defer testing.allocator.free(sheetless);
    try fixture.write(testing.allocator, io, sheetless, .sheet_ref);
    try fixture.patchPart(testing.allocator, io, sheetless, def_part, "<worksheetSource sheet=\"Data\" ref=\"A1:C4\"/>", "<worksheetSource ref=\"A1:C4\"/>");
    {
        var o = try Opened.open(testing.allocator, io, sheetless);
        defer o.deinit(testing.allocator);
        try testing.expectEqual(Unresolved.Why.sheetless_ref, o.pivots.caches[0].resolution.unresolved.why);
        try testing.expect((try cacheEdit(testing.allocator, &o, 0, .row, 1, .delete)) == null);
        try testing.expect((try cacheEdit(testing.allocator, &o, 1, .row, 1, .delete)) == null);
    }
    const dangling = try tt.path(testing.allocator, io, "s7b_dangling.xlsx");
    defer testing.allocator.free(dangling);
    try fixture.write(testing.allocator, io, dangling, .dangling);
    {
        var o = try Opened.open(testing.allocator, io, dangling);
        defer o.deinit(testing.allocator);
        try testing.expect((try cacheEdit(testing.allocator, &o, 0, .row, 1, .delete)) == null);
        try testing.expect((try cacheEdit(testing.allocator, &o, 1, .row, 1, .delete)) == null);
    }
    // An `r:id` the reader could not place beside a local `sheet`
    // (Q4 i): it may be this sheet, and the `ref` cannot be moved.
    // Another workbook's sheet of the same name is not local at all.
    const rid = try tt.path(testing.allocator, io, "s7b_unplaceable.xlsx");
    defer testing.allocator.free(rid);
    try fixture.write(testing.allocator, io, rid, .external);
    {
        var o = try Opened.open(testing.allocator, io, rid);
        defer o.deinit(testing.allocator);
        try testing.expect(o.pivots.caches[0].resolution == .external);
        try testing.expect((try cacheEdit(testing.allocator, &o, 0, .row, 1, .delete)) == null);
    }
    try fixture.patchPart(testing.allocator, io, rid, def_part, "sheet=\"Sheet1\"", "sheet=\"Data\"");
    try fixture.patchPart(testing.allocator, io, rid, "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels", "Id=\"rIdExt\"", "Id=\"rIdGone\"");
    var o = try Opened.open(testing.allocator, io, rid);
    defer o.deinit(testing.allocator);
    try testing.expectEqual(Unresolved.Why.unplaceable_rid, o.pivots.caches[0].resolution.unresolved.why);
    try testing.expectError(error.PivotSourceEditUnsafe, cacheEdit(testing.allocator, &o, 0, .row, 9, .insert));
    try testing.expect((try cacheEdit(testing.allocator, &o, 1, .row, 1, .insert)) == null);
}

test "workbookListsCaches: a main-namespace <pivotCaches> under the root, and nothing else" {
    try testing.expect(workbookListsCaches("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheets/><pivotCaches><pivotCache cacheId=\"1\" r:id=\"rId9\"/></pivotCaches></workbook>"));
    try testing.expect(workbookListsCaches("<x:workbook xmlns:x=\"http://purl.oclc.org/ooxml/spreadsheetml/main\"><x:pivotCaches/></x:workbook>"));
    try testing.expect(!workbookListsCaches("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheets/></workbook>"));
    // A vendor element of the same local name is not the list; a
    // comment that spells it is not either.
    try testing.expect(!workbookListsCaches("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:v=\"urn:v\"><extLst><ext><v:pivotCaches/></ext></extLst><!-- <pivotCaches/> --></workbook>"));
    // A root the scanner cannot read is the walk's to refuse.
    try testing.expect(workbookListsCaches("<workbook"));
}

fn cacheEditForFailures(allocator: Allocator, o: *const Opened) !void {
    const out = (try cacheEdit(allocator, o, 0, .row, 1, .insert)) orelse return error.TestExpectedMove;
    allocator.free(out);
}

test "edit: allocation failure in the cache splice leaves nothing behind" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "s7b_edit_failures.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .consolidation);
    try fixture.patchPart(testing.allocator, io, path, "xl/pivotCache/pivotCacheDefinition1.xml", "<rangeSet i1=\"1\" name=\"PivotSrc\"/>", "<rangeSet i1=\"1\" sheet=\"Data\" ref=\"A6:C9\"/>");
    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    try testing.checkAllAllocationFailures(testing.allocator, cacheEditForFailures, .{&o});
}

// ─── S7b-4: the engine ───────────────────────────────────────────────

const fixture_records_head =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" count="
;

/// The fixture's three data rows (`Data!A2:C4`) as the reader hands
/// them to the engine.
const fixture_rows = [_]engine.Row{
    &.{ .{ .string = "East" }, .{ .number = "3" }, .{ .number = "1.5" } },
    &.{ .{ .string = "West" }, .{ .number = "4" }, .{ .number = "2.5" } },
    &.{ .{ .string = "East" }, .{ .number = "5" }, .{ .number = "3.5" } },
};

const Rebuilt = struct { definition: []u8, records: []u8 };

/// The workbook's composition, on the fixture: plan the edit on sheet
/// 0, rebuild from `rows` (the data rows AFTER the edit), render both
/// parts.
fn rebuildFixture(alloc: Allocator, arena: Allocator, o: *const Opened, axis: edit.Axis, idx: u32, kind: edit.Kind, rows: []const engine.Row, refreshed: ?engine.Refreshed) !Rebuilt {
    const cache = &o.pivots.caches[0];
    var plan = try edit.planCacheEdit(arena, cache, 0, axis, idx, kind);
    try testing.expect(plan.changed);
    try testing.expect(plan.rebuild != null);
    const records_xml = (try o.store.part(cache.records_part_name.?)).?.bytes;
    const rb = try engine.rebuild(arena, cache, rows, records_xml, refreshed);
    try plan.splices.appendSlice(arena, rb.splices);
    const definition = (try edit.applyPlan(alloc, cache, &plan)).?;
    errdefer alloc.free(definition);
    return .{ .definition = definition, .records = try alloc.dupe(u8, rb.records.?) };
}

/// `src` with exactly one `old` replaced by `new`.
fn replacedOnce(alloc: Allocator, src: []const u8, old: []const u8, new: []const u8) ![]u8 {
    try testing.expectEqual(@as(usize, 1), std.mem.count(u8, src, old));
    return std.mem.replaceOwned(u8, alloc, src, old, new);
}

fn expectRecords(alloc: Allocator, got: []const u8, count: []const u8, body: []const u8) !void {
    const want = try std.mem.concat(alloc, u8, &.{ fixture_records_head, count, "\">", body, "</pivotCacheRecords>" });
    defer alloc.free(want);
    try testing.expectEqualStrings(want, got);
}

test "engine: an insert inside adds one blank record — every inventory keeps its order, the blank appends, counts and extremes follow, the root is redated" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "s7b4_insert.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .sheet_ref);
    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();

    // Row 2 is the first data row: the blank lands first.
    const rows = try engine.rowsAfterEdit(arena, &fixture_rows, 3, 2, 2, .insert);
    try testing.expectEqual(@as(usize, 4), rows.len);
    try testing.expect(rows[0][0] == .blank and rows[0][2] == .blank);
    try testing.expectEqualStrings("East", rows[1][0].string);

    const rb = try rebuildFixture(testing.allocator, arena, &o, .row, 2, .insert, rows, .{ .serial = 46000.5, .iso = "2025-12-13T12:00:00" });
    defer testing.allocator.free(rb.definition);
    defer testing.allocator.free(rb.records);

    const src = o.pivots.caches[0].raw_xml;
    var want = try replacedOnce(arena, src, "refreshedDate=\"45000.5\"", "refreshedDate=\"46000.5\"");
    want = try replacedOnce(arena, want, "recordCount=\"3\">", "recordCount=\"4\" refreshOnLoad=\"1\">");
    want = try replacedOnce(arena, want, "ref=\"A1:C4\"", "ref=\"A1:C5\"");
    want = try replacedOnce(arena, want, "<sharedItems count=\"2\"><s v=\"East\"/><s v=\"West\"/></sharedItems>", "<sharedItems containsBlank=\"1\" count=\"3\"><s v=\"East\"/><s v=\"West\"/><m/></sharedItems>");
    want = try replacedOnce(arena, want, "<sharedItems containsSemiMixedTypes=\"0\" containsString=\"0\" containsNumber=\"1\" containsInteger=\"1\" minValue=\"3\" maxValue=\"5\"/>", "<sharedItems containsString=\"0\" containsBlank=\"1\" containsNumber=\"1\" containsInteger=\"1\" minValue=\"3\" maxValue=\"5\"/>");
    want = try replacedOnce(arena, want, "<sharedItems containsSemiMixedTypes=\"0\" containsString=\"0\" containsNumber=\"1\" minValue=\"1.5\" maxValue=\"3.5\"/>", "<sharedItems containsString=\"0\" containsBlank=\"1\" containsNumber=\"1\" minValue=\"1.5\" maxValue=\"3.5\"/>");
    try testing.expectEqualStrings(want, rb.definition);
    try expectRecords(testing.allocator, rb.records, "4", "<r><x v=\"2\"/><m/><m/></r><r><x v=\"0\"/><n v=\"3\"/><n v=\"1.5\"/></r><r><x v=\"1\"/><n v=\"4\"/><n v=\"2.5\"/></r><r><x v=\"0\"/><n v=\"5\"/><n v=\"3.5\"/></r>");
}

test "engine: a delete inside drops one record and keeps the item it alone referenced — a consumer's index still names it" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "s7b4_delete.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .sheet_ref);
    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();

    // Row 3 is `West`: gone from the records, kept in the inventory.
    const rows = try engine.rowsAfterEdit(arena, &fixture_rows, 3, 2, 3, .delete);
    try testing.expectEqual(@as(usize, 2), rows.len);
    try testing.expectEqualStrings("5", rows[1][1].number);
    const rb = try rebuildFixture(testing.allocator, arena, &o, .row, 3, .delete, rows, .{ .serial = 46000.5, .iso = "2025-12-13T12:00:00" });
    defer testing.allocator.free(rb.definition);
    defer testing.allocator.free(rb.records);

    const src = o.pivots.caches[0].raw_xml;
    var want = try replacedOnce(arena, src, "refreshedDate=\"45000.5\"", "refreshedDate=\"46000.5\"");
    want = try replacedOnce(arena, want, "recordCount=\"3\">", "recordCount=\"2\" refreshOnLoad=\"1\">");
    want = try replacedOnce(arena, want, "ref=\"A1:C4\"", "ref=\"A1:C3\"");
    // The inventories are byte-identical: nothing new, nothing dropped,
    // the extremes unchanged (3..5 and 1.5..3.5 survive).
    try testing.expectEqualStrings(want, rb.definition);
    try expectRecords(testing.allocator, rb.records, "2", "<r><x v=\"0\"/><n v=\"3\"/><n v=\"1.5\"/></r><r><x v=\"0\"/><n v=\"5\"/><n v=\"3.5\"/></r>");
}

test "engine: a stale snapshot meets the cells — a value the inventory lacks appends, a string matches case-insensitively, a string in a numeric field enumerates it, mixed types and blanks are flagged, markup is escaped" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "s7b4_stale.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .sheet_ref);
    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();

    const long = try arena.alloc(u8, 256);
    @memset(long, 'x');
    const rows = [_]engine.Row{
        &.{ .{ .string = "east" }, .{ .number = "2.25" }, .{ .number = "10" } },
        &.{ .{ .string = "A&B<\"c\">" }, .{ .string = "n/a" }, .blank },
        &.{ .{ .number = "7" }, .{ .number = "3" }, .{ .number = "1E+15" } },
        &.{ .{ .string = long }, .blank, .{ .number = "-0" } },
    };
    const rb = try rebuildFixture(testing.allocator, arena, &o, .row, 3, .delete, &rows, .{ .serial = 46000.5, .iso = "2025-12-13T12:00:00" });
    defer testing.allocator.free(rb.definition);
    defer testing.allocator.free(rb.records);

    // Region: `east` is item 0; the markup string, the number and the
    // 256-character string append after the inventory, in order.
    const want_region = try std.fmt.allocPrint(arena, "<sharedItems containsMixedTypes=\"1\" containsNumber=\"1\" containsInteger=\"1\" minValue=\"7\" maxValue=\"7\" count=\"5\" longText=\"1\"><s v=\"East\"/><s v=\"West\"/><s v=\"A&amp;B&lt;&quot;c&quot;&gt;\"/><n v=\"7\"/><s v=\"{s}\"/></sharedItems>", .{long});
    try testing.expect(std.mem.indexOf(u8, rb.definition, want_region) != null);
    // Qty was inline (numbers only); a string enumerates it, in
    // first-appearance order, and the blank follows.
    try testing.expect(std.mem.indexOf(u8, rb.definition, "<sharedItems containsBlank=\"1\" containsMixedTypes=\"1\" containsNumber=\"1\" minValue=\"2.25\" maxValue=\"3\" count=\"4\"><n v=\"2.25\"/><s v=\"n/a\"/><n v=\"3\"/><m/></sharedItems>") != null);
    // Price stays inline: numbers and a blank; `1E+15` is integral but
    // past the 32-bit hint, `-0` is a number like any other.
    try testing.expect(std.mem.indexOf(u8, rb.definition, "<sharedItems containsString=\"0\" containsBlank=\"1\" containsNumber=\"1\" minValue=\"-0\" maxValue=\"1E+15\"/>") != null);
    try expectRecords(testing.allocator, rb.records, "4", "<r><x v=\"0\"/><x v=\"0\"/><n v=\"10\"/></r><r><x v=\"2\"/><x v=\"1\"/><m/></r><r><x v=\"3\"/><x v=\"2\"/><n v=\"1E+15\"/></r><r><x v=\"4\"/><x v=\"3\"/><n v=\"-0\"/></r>");
}

test "engine: an all-blank column, no clock, and a root without the counted attributes" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "s7b4_blank.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .sheet_ref);
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    const rows = [_]engine.Row{
        &.{ .{ .string = "East" }, .blank, .{ .number = "1.5" } },
    };
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        const rb = try rebuildFixture(testing.allocator, arena, &o, .row, 3, .delete, &rows, null);
        defer testing.allocator.free(rb.definition);
        defer testing.allocator.free(rb.records);
        // No clock: `refreshedDate` is removed, whitespace and all.
        try testing.expect(std.mem.indexOf(u8, rb.definition, "refreshedDate") == null);
        try testing.expect(std.mem.indexOf(u8, rb.definition, "refreshedBy=\"zlsx\" createdVersion=\"6\"") != null);
        try testing.expect(std.mem.indexOf(u8, rb.definition, "recordCount=\"1\" refreshOnLoad=\"1\">") != null);
        // An all-blank field is enumerated with its one blank item.
        try testing.expect(std.mem.indexOf(u8, rb.definition, "<cacheField name=\"Qty\" numFmtId=\"0\"><sharedItems containsNonDate=\"0\" containsString=\"0\" containsBlank=\"1\" count=\"1\"><m/></sharedItems></cacheField>") != null);
        try expectRecords(testing.allocator, rb.records, "1", "<r><x v=\"0\"/><x v=\"0\"/><n v=\"1.5\"/></r>");
    }
    // A root that carries neither `recordCount` nor `refreshedDate`
    // gains both in one insertion, before the marker's.
    try fixture.patchPart(testing.allocator, io, path, "xl/pivotCache/pivotCacheDefinition1.xml", " refreshedDate=\"45000.5\"", "");
    try fixture.patchPart(testing.allocator, io, path, "xl/pivotCache/pivotCacheDefinition1.xml", " recordCount=\"3\"", "");
    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    const rb = try rebuildFixture(testing.allocator, arena, &o, .row, 3, .delete, &rows, .{ .serial = 46001, .iso = "2025-12-14T00:00:00" });
    defer testing.allocator.free(rb.definition);
    defer testing.allocator.free(rb.records);
    try testing.expect(std.mem.indexOf(u8, rb.definition, "minRefreshableVersion=\"3\" refreshOnLoad=\"1\" recordCount=\"1\" refreshedDate=\"46001\">") != null);
}

test "engine: rowsAfterEdit lands on the data row the edit named, header rows or none" {
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    // Header at row 1, data at 2..4: an insert at 4 lands before the
    // last row; a delete at 4 drops it.
    const ins = try engine.rowsAfterEdit(arena, &fixture_rows, 3, 2, 4, .insert);
    try testing.expectEqual(@as(usize, 4), ins.len);
    try testing.expect(ins[2][0] == .blank);
    try testing.expectEqualStrings("East", ins[3][0].string);
    const del = try engine.rowsAfterEdit(arena, &fixture_rows, 3, 2, 4, .delete);
    try testing.expectEqual(@as(usize, 2), del.len);
    try testing.expectEqualStrings("West", del[1][0].string);
    // Headerless: row 1 is data index 0.
    const del0 = try engine.rowsAfterEdit(arena, &fixture_rows, 3, 1, 1, .delete);
    try testing.expectEqualStrings("West", del0[0][0].string);
    // Off the data: not an edit the predicate admitted.
    try testing.expectError(error.MalformedPivotXml, engine.rowsAfterEdit(arena, &fixture_rows, 3, 2, 1, .delete));
    try testing.expectError(error.MalformedPivotXml, engine.rowsAfterEdit(arena, &fixture_rows, 3, 2, 5, .delete));
    try testing.expectError(error.MalformedPivotXml, engine.rowsAfterEdit(arena, &fixture_rows, 3, 2, 6, .insert));
}

fn expectShape(alloc: Allocator, src: []const u8, old: []const u8, new: []const u8, want: anyerror) !void {
    const xml = try replacedOnce(alloc, src, old, new);
    defer alloc.free(xml);
    var def = try pivot_xml.parseCacheDefinition(alloc, xml);
    defer def.deinit(alloc);
    try testing.expectError(want, engine.checkShape(&def));
}

test "engine: the shapes the slice refuses — calculated, group and OLAP elements, date inventories, items with children, a missing inventory; and the shapes it reads" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "s7b4_shapes.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .sheet_ref);
    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    const alloc = testing.allocator;
    const src = o.pivots.caches[0].raw_xml;
    try engine.checkShape(&o.pivots.caches[0].definition);

    const price_field = "<cacheField name=\"Price\" numFmtId=\"0\">";
    try expectShape(alloc, src, price_field, "<cacheField name=\"Price\" numFmtId=\"0\" formula=\"Qty*2\" databaseField=\"0\">", error.PivotShapeUnsupported);
    try expectShape(alloc, src, price_field, "<cacheField name=\"Price\" numFmtId=\"0\" databaseField=\"0\">", error.PivotShapeUnsupported);
    try expectShape(alloc, src, price_field, price_field ++ "<fieldGroup base=\"1\"/>", error.PivotShapeUnsupported);
    try expectShape(alloc, src, "</cacheFields>", "</cacheFields><cacheHierarchies count=\"0\"/>", error.PivotShapeUnsupported);
    try expectShape(alloc, src, "</cacheFields>", "</cacheFields><kpis count=\"0\"/>", error.PivotShapeUnsupported);
    try expectShape(alloc, src, "<sharedItems count=\"2\"><s v=\"East\"/><s v=\"West\"/></sharedItems>", "<sharedItems containsSemiMixedTypes=\"0\" containsNonDate=\"0\" containsDate=\"1\" containsString=\"0\" minDate=\"2024-01-01T00:00:00\" maxDate=\"2024-01-02T00:00:00\" count=\"1\"><d v=\"2024-01-01T00:00:00\"/></sharedItems>", error.PivotShapeUnsupported);
    try expectShape(alloc, src, "<s v=\"West\"/>", "<b v=\"1\"/>", error.PivotShapeUnsupported);
    try expectShape(alloc, src, "<s v=\"West\"/>", "<e v=\"#N/A\"/>", error.PivotShapeUnsupported);
    try expectShape(alloc, src, "<s v=\"West\"/>", "<s v=\"West\"><tpls c=\"1\"><tpl fld=\"0\" item=\"1\"/></tpls></s>", error.PivotShapeUnsupported);
    try expectShape(alloc, src, "<sharedItems containsSemiMixedTypes=\"0\" containsString=\"0\" containsNumber=\"1\" minValue=\"1.5\" maxValue=\"3.5\"/>", "", error.PivotShapeUnsupported);
    try expectShape(alloc, src, "<sharedItems count=\"2\">", "<sharedItems count=\"3\">", error.MalformedPivotXml);
    try expectShape(alloc, src, "<cacheFields count=\"3\">", "<cacheFields count=\"4\">", error.MalformedPivotXml);
    // A consolidation source, an external one: not a rectangle.
    try expectShape(alloc, src, "<cacheSource type=\"worksheet\"><worksheetSource sheet=\"Data\" ref=\"A1:C4\"/></cacheSource>", "<cacheSource type=\"consolidation\"><consolidation><rangeSets count=\"1\"><rangeSet sheet=\"Data\" ref=\"A1:C4\"/></rangeSets></consolidation></cacheSource>", error.PivotShapeUnsupported);
    try expectShape(alloc, src, "type=\"worksheet\"", "type=\"external\" connectionId=\"1\"", error.PivotShapeUnsupported);
    // The corpus' own shapes read: an `extLst` on the root and an
    // inventory attribute Excel writes (`containsInteger`) are fine.
    {
        const xml = try replacedOnce(alloc, src, "</cacheFields>", "</cacheFields><extLst><ext uri=\"{725AE2AE-9491-48be-B2B4-4EB974FC3084}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:pivotCacheDefinition pivotCacheId=\"2\"/></ext></extLst>");
        defer alloc.free(xml);
        var def = try pivot_xml.parseCacheDefinition(alloc, xml);
        defer def.deinit(alloc);
        try engine.checkShape(&def);
    }
}

test "engine: the rectangle must be the schema — a wider or narrower one, no rows, or a records part carrying more than records refuse" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "s7b4_width.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .sheet_ref);
    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    const cache = &o.pivots.caches[0];
    const records_xml = (try o.store.part(cache.records_part_name.?)).?.bytes;

    const narrow = [_]engine.Row{&.{ .{ .string = "East" }, .{ .number = "3" } }};
    try testing.expectError(error.PivotShapeUnsupported, engine.rebuild(arena, cache, &narrow, records_xml, null));
    const wide = [_]engine.Row{&.{ .{ .string = "East" }, .{ .number = "3" }, .{ .number = "1" }, .blank }};
    try testing.expectError(error.PivotShapeUnsupported, engine.rebuild(arena, cache, &wide, records_xml, null));
    try testing.expectError(error.PivotShapeUnsupported, engine.rebuild(arena, cache, &.{}, records_xml, null));
    const ragged = [_]engine.Row{ fixture_rows[0], &.{ .{ .string = "East" }, .{ .number = "3" } } };
    try testing.expectError(error.MalformedPivotXml, engine.rebuild(arena, cache, &ragged, records_xml, null));
    // A number lexical the reader would never pass.
    const nan = [_]engine.Row{&.{ .{ .string = "East" }, .{ .number = "nan" }, .{ .number = "1" } }};
    try testing.expectError(error.PivotShapeUnsupported, engine.rebuild(arena, cache, &nan, records_xml, null));
    // Records with an extension list: not carried over.
    const ext = try replacedOnce(arena, records_xml, "</pivotCacheRecords>", "<extLst/></pivotCacheRecords>");
    try testing.expectError(error.PivotShapeUnsupported, engine.rebuild(arena, cache, &fixture_rows, ext, null));
    // No records part at all: the definition alone.
    const rb = try engine.rebuild(arena, cache, &fixture_rows, null, null);
    try testing.expect(rb.records == null);
    try testing.expectEqual(@as(u32, 3), rb.record_count);
}

test "engine: a Strict-prefixed part is rebuilt under its prefix — items, records and the root's close" {
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    const def_xml =
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<x:pivotCacheDefinition xmlns:x="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" r:id="rId1" recordCount="1"><x:cacheSource type="worksheet"><x:worksheetSource sheet="Data" ref="A1:B2"/></x:cacheSource><x:cacheFields count="2"><x:cacheField name="K" numFmtId="0"><x:sharedItems count="1"><x:s v="a"/></x:sharedItems></x:cacheField><x:cacheField name="V" numFmtId="0"><x:sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="1" maxValue="1"/></x:cacheField></x:cacheFields></x:pivotCacheDefinition>
    ;
    const rec_xml =
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<x:pivotCacheRecords xmlns:x="http://purl.oclc.org/ooxml/spreadsheetml/main" count="1"><x:r><x:x v="0"/><x:n v="1"/></x:r></x:pivotCacheRecords>
    ;
    const def = try pivot_xml.parseCacheDefinition(arena, def_xml);
    try testing.expectEqualStrings("x", def.prefix);
    try testing.expectEqual(@as(usize, 1), def.fields[0].shared_items.?.items.len);
    try testing.expectEqualStrings("<x:sharedItems count=\"1\"><x:s v=\"a\"/></x:sharedItems>", def_xml[def.fields[0].shared_items.?.span.start..def.fields[0].shared_items.?.span.end]);
    const cache: PivotCache = .{
        .cache_id = null,
        .part_name = "xl/pivotCache/pivotCacheDefinition1.xml",
        .records_part_name = "xl/pivotCache/pivotCacheRecords1.xml",
        .definition = def,
        .field_names = &.{ "K", "V" },
        .field_formulas = &.{ null, null },
        .source = .{},
        .resolution = .none,
        .range_set_sources = &.{},
        .range_set_resolutions = &.{},
        .consumer_count = 0,
        .raw_xml = def_xml,
    };
    const rows = [_]engine.Row{
        &.{ .{ .string = "a" }, .{ .number = "1" } },
        &.{ .{ .string = "b" }, .blank },
    };
    const rb = try engine.rebuild(arena, &cache, &rows, rec_xml, .{ .serial = 46002.25, .iso = "2025-12-15T06:00:00" });
    const out = try edit.spliceAll(arena, def_xml, rb.splices);
    try testing.expect(std.mem.indexOf(u8, out, "<x:sharedItems count=\"2\"><x:s v=\"a\"/><x:s v=\"b\"/></x:sharedItems>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<x:sharedItems containsString=\"0\" containsBlank=\"1\" containsNumber=\"1\" containsInteger=\"1\" minValue=\"1\" maxValue=\"1\"/>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "r:id=\"rId1\" recordCount=\"2\" refreshedDate=\"46002.25\">") != null);
    try testing.expectEqualStrings(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<x:pivotCacheRecords xmlns:x=\"http://purl.oclc.org/ooxml/spreadsheetml/main\" count=\"2\"><x:r><x:x v=\"0\"/><x:n v=\"1\"/></x:r><x:r><x:x v=\"1\"/><x:m/></x:r></x:pivotCacheRecords>",
        rb.records.?,
    );
}

test "engine: the plan names the rectangle for the carriers the slice reads and none for the rest" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();

    const kinds = [_]fixture.SourceKind{ .sheet_ref, .table_name, .defined_name };
    for (kinds, 0..) |kind, i| {
        const name = try std.fmt.allocPrint(arena, "s7b4_plan_{d}.xlsx", .{i});
        const path = try tt.path(testing.allocator, io, name);
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path, kind);
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        const inside = try edit.planCacheEdit(arena, &o.pivots.caches[0], 0, .row, 3, .insert);
        try testing.expect(inside.changed);
        try testing.expect(inside.rebuild.?.rect.eql(.{ .tl_col = 1, .tl_row = 1, .br_col = 3, .br_row = 4 }));
        try testing.expectEqual(@as(u32, 1), inside.rebuild.?.header_rows);
        // A pure shift plans no rebuild and no marker.
        const above = try edit.planCacheEdit(arena, &o.pivots.caches[0], 0, .row, 1, .insert);
        try testing.expect(!above.changed and above.rebuild == null);
        // Another sheet: nothing.
        const other = try edit.planCacheEdit(arena, &o.pivots.caches[0], 1, .row, 2, .insert);
        try testing.expect(!other.changed and other.splices.items.len == 0);
    }
    // A headerless table: the rectangle's first row is data.
    {
        const path = try tt.path(testing.allocator, io, "s7b4_plan_headerless.xlsx");
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path, .table_name);
        try fixture.patchPart(testing.allocator, io, path, "xl/tables/table1.xml", "ref=\"A1:C4\" totalsRowShown=\"0\"", "ref=\"A1:C4\" headerRowCount=\"0\" totalsRowShown=\"0\"");
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try testing.expectEqual(@as(u32, 0), o.pivots.caches[0].resolution.sheet.header_rows);
        const top = try edit.planCacheEdit(arena, &o.pivots.caches[0], 0, .row, 1, .delete);
        try testing.expect(top.changed);
        try testing.expectEqual(@as(u32, 0), top.rebuild.?.header_rows);
    }
    // Whole columns, an unbounded body, a consolidation set: changed,
    // and nothing to rebuild from.
    {
        const path = try tt.path(testing.allocator, io, "s7b4_plan_cols.xlsx");
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path, .defined_name);
        try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "Data!$A$1:$C$4", "Data!$A:$C");
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        const plan = try edit.planCacheEdit(arena, &o.pivots.caches[0], 0, .row, 2, .insert);
        try testing.expect(plan.changed and plan.rebuild == null);
    }
    {
        const path = try tt.path(testing.allocator, io, "s7b4_plan_offset.xlsx");
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path, .defined_name);
        try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "Data!$A$1:$C$4", "OFFSET(Data!$A$1,0,0,4,3)");
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        const plan = try edit.planCacheEdit(arena, &o.pivots.caches[0], 0, .row, 9, .insert);
        try testing.expect(plan.changed and plan.rebuild == null);
    }
    {
        const path = try tt.path(testing.allocator, io, "s7b4_plan_consolidation.xlsx");
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path, .consolidation);
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        const plan = try edit.planCacheEdit(arena, &o.pivots.caches[0], 0, .row, 2, .insert);
        try testing.expect(plan.changed and plan.rebuild == null);
    }
}

// ─── Codex #205 round 1 ──────────────────────────────────────────────

test "engine: a close tag padded with whitespace bounds the inventory span — the splice takes the whole element, on either prefix (Codex #205 r1 REL-101)" {
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    const main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const plain = "<?xml version=\"1.0\"?><pivotCacheDefinition xmlns=\"" ++ main_ns ++ "\" recordCount=\"1\"><cacheSource type=\"worksheet\"><worksheetSource sheet=\"Data\" ref=\"A1:A2\"/></cacheSource><cacheFields count=\"1\"><cacheField name=\"K\" numFmtId=\"0\"><sharedItems count=\"1\"><s v=\"a\"/></sharedItems \t></cacheField></cacheFields></pivotCacheDefinition>";
    const strict = "<x:pivotCacheDefinition xmlns:x=\"http://purl.oclc.org/ooxml/spreadsheetml/main\"><x:cacheSource type=\"worksheet\"><x:worksheetSource sheet=\"Data\" ref=\"A1:A2\"/></x:cacheSource><x:cacheFields count=\"1\"><x:cacheField name=\"K\" numFmtId=\"0\"><x:sharedItems count=\"1\"><x:s v=\"a\"/></x:sharedItems\n></x:cacheField></x:cacheFields></x:pivotCacheDefinition>";
    const cases = [_]struct { xml: []const u8, want: []const u8 }{
        .{ .xml = plain, .want = "<sharedItems count=\"1\"><s v=\"a\"/></sharedItems \t>" },
        .{ .xml = strict, .want = "<x:sharedItems count=\"1\"><x:s v=\"a\"/></x:sharedItems\n>" },
    };
    for (cases) |c| {
        const def = try pivot_xml.parseCacheDefinition(arena, c.xml);
        const span = def.fields[0].shared_items.?.span;
        try testing.expectEqualStrings(c.want, c.xml[span.start..span.end]);
    }
    // Rebuilt: the padded close is gone with the element, the part
    // still one readable tree.
    const cache: PivotCache = .{
        .cache_id = null,
        .part_name = "xl/pivotCache/pivotCacheDefinition1.xml",
        .records_part_name = null,
        .definition = try pivot_xml.parseCacheDefinition(arena, plain),
        .field_names = &.{"K"},
        .field_formulas = &.{null},
        .source = .{},
        .resolution = .none,
        .range_set_sources = &.{},
        .range_set_resolutions = &.{},
        .consumer_count = 0,
        .raw_xml = plain,
    };
    const rows = [_]engine.Row{&.{.{ .string = "b" }}};
    const rb = try engine.rebuild(arena, &cache, &rows, null, null);
    const out = try edit.spliceAll(arena, plain, rb.splices);
    try testing.expect(std.mem.indexOf(u8, out, "<sharedItems count=\"2\"><s v=\"a\"/><s v=\"b\"/></sharedItems></cacheField>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "\t>") == null);
    _ = try pivot_xml.parseCacheDefinition(arena, out);
}

test "engine: a namespace alias or rebinding below the root refuses the part; a foreign direct child reads but refuses the rebuild (Codex #205 r1 REL-102)" {
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    const main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const strict_ns = "http://purl.oclc.org/ooxml/spreadsheetml/main";
    const head = "<pivotCacheDefinition xmlns=\"" ++ main_ns ++ "\" recordCount=\"1\"><cacheSource type=\"worksheet\"><worksheetSource sheet=\"Data\" ref=\"A1:A2\"/></cacheSource><cacheFields count=\"1\"><cacheField name=\"K\" numFmtId=\"0\">";
    const tail = "</cacheField></cacheFields></pivotCacheDefinition>";
    const inventory = "<sharedItems count=\"1\"><s v=\"a\"/></sharedItems>";

    // Refused whole: the main namespace under another prefix on a
    // descendant (either spelling of the URI), the default namespace
    // rebound, the root's prefix rebound, and a root whose one main
    // binding is not its own prefix.
    const refused = [_][]const u8{
        head ++ "<sharedItems count=\"1\" xmlns:y=\"" ++ main_ns ++ "\"><y:s v=\"a\"/></sharedItems>" ++ tail,
        head ++ inventory ++ "<y:fieldGroup xmlns:y=\"" ++ strict_ns ++ "\" base=\"0\"/>" ++ tail,
        head ++ inventory ++ "<fieldGroup xmlns=\"urn:vendor\" base=\"0\"/>" ++ tail,
        "<x:pivotCacheDefinition xmlns:x=\"" ++ main_ns ++ "\"><x:cacheSource type=\"worksheet\"><x:worksheetSource sheet=\"Data\" ref=\"A1:A2\"/></x:cacheSource><x:cacheFields count=\"1\"><x:cacheField name=\"K\" numFmtId=\"0\" xmlns:x=\"urn:vendor\"><x:sharedItems count=\"1\"><x:s v=\"a\"/></x:sharedItems></x:cacheField></x:cacheFields></x:pivotCacheDefinition>",
        "<pivotCacheDefinition xmlns:y=\"" ++ main_ns ++ "\"><cacheSource type=\"worksheet\"><worksheetSource sheet=\"Data\" ref=\"A1:A2\"/></cacheSource></pivotCacheDefinition>",
    };
    for (refused) |xml| try testing.expectError(error.MalformedXml, pivot_xml.parseCacheDefinition(arena, xml));

    // A redundant redeclaration of the root's own binding changes
    // nothing; a vendor prefix bound to a vendor URI is nobody's.
    {
        const def = try pivot_xml.parseCacheDefinition(arena, head ++ "<sharedItems count=\"1\" xmlns=\"" ++ main_ns ++ "\" xmlns:z=\"urn:vendor\"><s v=\"a\"/></sharedItems>" ++ tail);
        try engine.checkShape(&def);
    }

    // A foreign direct child where the schema names only its own:
    // read and flagged; the rebuild, which would regenerate around it,
    // refuses.
    const foreign = [_][]const u8{
        head ++ "<sharedItems count=\"1\"><s v=\"a\"/><z:s xmlns:z=\"urn:vendor\" v=\"b\"/></sharedItems>" ++ tail,
        head ++ inventory ++ "<z:fieldGroup xmlns:z=\"urn:vendor\" base=\"0\"/>" ++ tail,
        head ++ inventory ++ "</cacheField></cacheFields><z:kpis xmlns:z=\"urn:vendor\"/></pivotCacheDefinition>",
    };
    for (foreign) |xml| {
        const def = try pivot_xml.parseCacheDefinition(arena, xml);
        try testing.expectError(error.PivotShapeUnsupported, engine.checkShape(&def));
    }

    // The records part: an aliased record refuses the read, a foreign
    // direct child refuses the rebuild — neither is dropped.
    const def_xml = head ++ inventory ++ tail;
    const cache: PivotCache = .{
        .cache_id = null,
        .part_name = "xl/pivotCache/pivotCacheDefinition1.xml",
        .records_part_name = "xl/pivotCache/pivotCacheRecords1.xml",
        .definition = try pivot_xml.parseCacheDefinition(arena, def_xml),
        .field_names = &.{"K"},
        .field_formulas = &.{null},
        .source = .{},
        .resolution = .none,
        .range_set_sources = &.{},
        .range_set_resolutions = &.{},
        .consumer_count = 0,
        .raw_xml = def_xml,
    };
    const rows = [_]engine.Row{&.{.{ .string = "a" }}};
    const aliased = "<pivotCacheRecords xmlns=\"" ++ main_ns ++ "\" count=\"1\"><y:r xmlns:y=\"" ++ main_ns ++ "\"><y:x v=\"0\"/></y:r></pivotCacheRecords>";
    try testing.expectError(error.MalformedPivotXml, engine.rebuild(arena, &cache, &rows, aliased, null));
    const vendor = "<pivotCacheRecords xmlns=\"" ++ main_ns ++ "\" count=\"1\"><r><x v=\"0\"/></r><z:r xmlns:z=\"urn:vendor\"/></pivotCacheRecords>";
    try testing.expectError(error.PivotShapeUnsupported, engine.rebuild(arena, &cache, &rows, vendor, null));
    const clean = "<pivotCacheRecords xmlns=\"" ++ main_ns ++ "\" count=\"1\"><r><x v=\"0\"/></r></pivotCacheRecords>";
    const rb = try engine.rebuild(arena, &cache, &rows, clean, null);
    try testing.expectEqualStrings(clean, rb.records.?);
}

test "engine: containsInteger's 32-bit interval is the signed one, both ends in (Codex #205 r1 REL-105)" {
    try testing.expect(engine.isInteger(-2147483648));
    try testing.expect(engine.isInteger(2147483647));
    try testing.expect(engine.isInteger(0));
    try testing.expect(!engine.isInteger(-2147483649));
    try testing.expect(!engine.isInteger(2147483648));
    try testing.expect(!engine.isInteger(1.5));
}

test "engine: an explicitly closed empty item is the self-closing one — kept as written on either prefix; a child still refuses (Codex #205 r2 REL-201)" {
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    const main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const head = "<pivotCacheDefinition xmlns=\"" ++ main_ns ++ "\" recordCount=\"1\"><cacheSource type=\"worksheet\"><worksheetSource sheet=\"Data\" ref=\"A1:A2\"/></cacheSource><cacheFields count=\"1\"><cacheField name=\"K\" numFmtId=\"0\">";
    const tail = "</cacheField></cacheFields></pivotCacheDefinition>";
    const xml = head ++ "<sharedItems containsMixedTypes=\"1\" containsNumber=\"1\" containsInteger=\"1\" minValue=\"1\" maxValue=\"1\" containsBlank=\"1\" count=\"3\"><s v=\"a\"></s><n v=\"1\"></n><m>\n</m></sharedItems>" ++ tail;
    const def = try pivot_xml.parseCacheDefinition(arena, xml);
    const items = def.fields[0].shared_items.?.items;
    try testing.expectEqual(@as(usize, 3), items.len);
    for (items) |it| try testing.expect(it.simple);
    try testing.expectEqualStrings("<s v=\"a\"></s>", items[0].raw);
    try engine.checkShape(&def);
    const cache: PivotCache = .{
        .cache_id = null,
        .part_name = "xl/pivotCache/pivotCacheDefinition1.xml",
        .records_part_name = null,
        .definition = def,
        .field_names = &.{"K"},
        .field_formulas = &.{null},
        .source = .{},
        .resolution = .none,
        .range_set_sources = &.{},
        .range_set_resolutions = &.{},
        .consumer_count = 0,
        .raw_xml = xml,
    };
    const rows = [_]engine.Row{&.{.{ .string = "b" }}};
    const rb = try engine.rebuild(arena, &cache, &rows, null, null);
    const out = try edit.spliceAll(arena, xml, rb.splices);
    try testing.expect(std.mem.indexOf(u8, out, "count=\"4\"><s v=\"a\"></s><n v=\"1\"></n><m>\n</m><s v=\"b\"/></sharedItems>") != null);
    _ = try pivot_xml.parseCacheDefinition(arena, out);

    // Strict, explicit close.
    const strict = "<x:pivotCacheDefinition xmlns:x=\"http://purl.oclc.org/ooxml/spreadsheetml/main\"><x:cacheSource type=\"worksheet\"><x:worksheetSource sheet=\"Data\" ref=\"A1:A2\"/></x:cacheSource><x:cacheFields count=\"1\"><x:cacheField name=\"K\" numFmtId=\"0\"><x:sharedItems count=\"1\"><x:s v=\"a\"></x:s></x:sharedItems></x:cacheField></x:cacheFields></x:pivotCacheDefinition>";
    const sdef = try pivot_xml.parseCacheDefinition(arena, strict);
    try testing.expect(sdef.fields[0].shared_items.?.items[0].simple);
    try engine.checkShape(&sdef);

    // A child is a child.
    const nested = try pivot_xml.parseCacheDefinition(arena, head ++ "<sharedItems count=\"1\"><s v=\"a\"><tpls c=\"1\"><tpl fld=\"0\" item=\"0\"/></tpls></s></sharedItems>" ++ tail);
    try testing.expect(!nested.fields[0].shared_items.?.items[0].simple);
    try testing.expectError(error.PivotShapeUnsupported, engine.checkShape(&nested));
}

test "engine: a rebuilt inventory describes the items it holds — a retained string, number, blank or long text sets the flags the rows alone would not (Codex #205 r3 REL-303)" {
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    const main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const head = "<pivotCacheDefinition xmlns=\"" ++ main_ns ++ "\" recordCount=\"1\"><cacheSource type=\"worksheet\"><worksheetSource sheet=\"Data\" ref=\"A1:A2\"/></cacheSource><cacheFields count=\"1\"><cacheField name=\"K\" numFmtId=\"0\">";
    const tail = "</cacheField></cacheFields></pivotCacheDefinition>";
    const long = try arena.alloc(u8, 256);
    @memset(long, 'y');
    const Case = struct { inventory: []const u8, row: engine.Value, want: []const u8 };
    const cases = [_]Case{
        // Retained string + number, rows all integral numbers: mixed, not
        // integer (1.5 is retained), extrema over the items.
        .{ .inventory = "<sharedItems count=\"2\"><s v=\"East\"/><n v=\"1.5\"/></sharedItems>", .row = .{ .number = "7" }, .want = "<sharedItems containsMixedTypes=\"1\" containsNumber=\"1\" minValue=\"1.5\" maxValue=\"7\" count=\"3\"><s v=\"East\"/><n v=\"1.5\"/><n v=\"7\"/></sharedItems>" },
        // Retained blank, rows a string.
        .{ .inventory = "<sharedItems containsBlank=\"1\" count=\"1\"><m/></sharedItems>", .row = .{ .string = "a" }, .want = "<sharedItems containsBlank=\"1\" count=\"2\"><m/><s v=\"a\"/></sharedItems>" },
        // Retained long text, rows a short one.
        .{ .inventory = try std.fmt.allocPrint(arena, "<sharedItems count=\"1\" longText=\"1\"><s v=\"{s}\"/></sharedItems>", .{long}), .row = .{ .string = "a" }, .want = try std.fmt.allocPrint(arena, "<sharedItems count=\"2\" longText=\"1\"><s v=\"{s}\"/><s v=\"a\"/></sharedItems>", .{long}) },
        // Retained numbers only, rows a number: still semi-mixed 0 /
        // string 0, the integer hint from every item.
        .{ .inventory = "<sharedItems containsSemiMixedTypes=\"0\" containsString=\"0\" containsNumber=\"1\" containsInteger=\"1\" minValue=\"2\" maxValue=\"2\" count=\"1\"><n v=\"2\"/></sharedItems>", .row = .{ .number = "9" }, .want = "<sharedItems containsSemiMixedTypes=\"0\" containsString=\"0\" containsNumber=\"1\" containsInteger=\"1\" minValue=\"2\" maxValue=\"9\" count=\"2\"><n v=\"2\"/><n v=\"9\"/></sharedItems>" },
    };
    for (cases) |c| {
        const xml = try std.mem.concat(arena, u8, &.{ head, c.inventory, tail });
        const cache: PivotCache = .{
            .cache_id = null,
            .part_name = "xl/pivotCache/pivotCacheDefinition1.xml",
            .records_part_name = null,
            .definition = try pivot_xml.parseCacheDefinition(arena, xml),
            .field_names = &.{"K"},
            .field_formulas = &.{null},
            .source = .{},
            .resolution = .none,
            .range_set_sources = &.{},
            .range_set_resolutions = &.{},
            .consumer_count = 0,
            .raw_xml = xml,
        };
        const rows = [_]engine.Row{&.{c.row}};
        const rb = try engine.rebuild(arena, &cache, &rows, null, null);
        const out = try edit.spliceAll(arena, xml, rb.splices);
        try testing.expect(std.mem.indexOf(u8, out, c.want) != null);
    }
}

test "engine: a main-namespace binding first introduced below the root refuses; a record holding anything but childless value elements refuses (Codex #205 r3 REL-302, REL-304)" {
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    const main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    // The root never bound the main namespace: a descendant cannot be
    // the one to, on either spelling of the root.
    const late = [_][]const u8{
        "<pivotCacheDefinition recordCount=\"1\"><cacheSource type=\"worksheet\"><worksheetSource sheet=\"Data\" ref=\"A1:A2\"/></cacheSource><cacheFields count=\"1\"><cacheField name=\"K\" numFmtId=\"0\"><sharedItems count=\"1\" xmlns=\"" ++ main_ns ++ "\"><s v=\"a\"/></sharedItems></cacheField></cacheFields></pivotCacheDefinition>",
        "<x:pivotCacheDefinition recordCount=\"1\"><x:cacheSource type=\"worksheet\"><x:worksheetSource sheet=\"Data\" ref=\"A1:A2\"/></x:cacheSource><x:cacheFields count=\"1\"><x:cacheField name=\"K\" numFmtId=\"0\"><x:sharedItems count=\"1\" xmlns:x=\"" ++ main_ns ++ "\"><x:s v=\"a\"/></x:sharedItems></x:cacheField></x:cacheFields></x:pivotCacheDefinition>",
    };
    for (late) |xml| try testing.expectError(error.MalformedXml, pivot_xml.parseCacheDefinition(arena, xml));

    const def_xml = "<pivotCacheDefinition xmlns=\"" ++ main_ns ++ "\" recordCount=\"1\"><cacheSource type=\"worksheet\"><worksheetSource sheet=\"Data\" ref=\"A1:A2\"/></cacheSource><cacheFields count=\"1\"><cacheField name=\"K\" numFmtId=\"0\"><sharedItems count=\"1\"><s v=\"a\"/></sharedItems></cacheField></cacheFields></pivotCacheDefinition>";
    const cache: PivotCache = .{
        .cache_id = null,
        .part_name = "xl/pivotCache/pivotCacheDefinition1.xml",
        .records_part_name = "xl/pivotCache/pivotCacheRecords1.xml",
        .definition = try pivot_xml.parseCacheDefinition(arena, def_xml),
        .field_names = &.{"K"},
        .field_formulas = &.{null},
        .source = .{},
        .resolution = .none,
        .range_set_sources = &.{},
        .range_set_resolutions = &.{},
        .consumer_count = 0,
        .raw_xml = def_xml,
    };
    const rows = [_]engine.Row{&.{.{ .string = "a" }}};
    const rec_head = "<pivotCacheRecords xmlns=\"" ++ main_ns ++ "\" count=\"1\">";
    const unsupported = [_][]const u8{
        rec_head ++ "<r><x v=\"0\"/><extLst/></r></pivotCacheRecords>",
        rec_head ++ "<r><x v=\"0\"/><z:q xmlns:z=\"urn:vendor\"/></r></pivotCacheRecords>",
        rec_head ++ "<r><x v=\"0\"><t/></x></r></pivotCacheRecords>",
        rec_head ++ "<r><x v=\"0\"/><y v=\"1\"/></r></pivotCacheRecords>",
    };
    for (unsupported) |xml| try testing.expectError(error.PivotShapeUnsupported, engine.rebuild(arena, &cache, &rows, xml, null));
    // A binding introduced on a record refuses the read.
    try testing.expectError(error.MalformedPivotXml, engine.rebuild(arena, &cache, &rows, "<pivotCacheRecords count=\"1\"><r xmlns=\"" ++ main_ns ++ "\"><x v=\"0\"/></r></pivotCacheRecords>", null));
    // Every value kind, explicitly closed around nothing, reads — one
    // per record, a record being one value per field: `<x>` against
    // the stocked inventory, the inline kinds against an item-less one
    // (inline records beside items are not one shape — r11 REL-1101).
    {
        const rb = try engine.rebuild(arena, &cache, &rows, rec_head ++ "<r><x v=\"0\"></x></r></pivotCacheRecords>", null);
        try testing.expectEqualStrings(rec_head ++ "<r><x v=\"0\"/></r></pivotCacheRecords>", rb.records.?);
    }
    const inline_xml = "<pivotCacheDefinition xmlns=\"" ++ main_ns ++ "\" recordCount=\"1\"><cacheSource type=\"worksheet\"><worksheetSource sheet=\"Data\" ref=\"A1:A2\"/></cacheSource><cacheFields count=\"1\"><cacheField name=\"V\" numFmtId=\"0\"><sharedItems containsSemiMixedTypes=\"0\" containsString=\"0\" containsNumber=\"1\" containsInteger=\"1\" minValue=\"1\" maxValue=\"1\"/></cacheField></cacheFields></pivotCacheDefinition>";
    const inline_cache: PivotCache = .{
        .cache_id = null,
        .part_name = "xl/pivotCache/pivotCacheDefinition1.xml",
        .records_part_name = "xl/pivotCache/pivotCacheRecords1.xml",
        .definition = try pivot_xml.parseCacheDefinition(arena, inline_xml),
        .field_names = &.{"V"},
        .field_formulas = &.{null},
        .source = .{},
        .resolution = .none,
        .range_set_sources = &.{},
        .range_set_resolutions = &.{},
        .consumer_count = 0,
        .raw_xml = inline_xml,
    };
    const number_rows = [_]engine.Row{&.{.{ .number = "1" }}};
    for ([_][]const u8{ "<n v=\"1\"></n>", "<s v=\"q\"></s>", "<m></m>", "<b v=\"1\"></b>", "<d v=\"2024-01-01T00:00:00\"></d>", "<e v=\"#N/A\"></e>" }) |one| {
        const part = try std.mem.concat(arena, u8, &.{ rec_head, "<r>", one, "</r></pivotCacheRecords>" });
        const rb = try engine.rebuild(arena, &inline_cache, &number_rows, part, null);
        try testing.expectEqualStrings(rec_head ++ "<r><n v=\"1\"/></r></pivotCacheRecords>", rb.records.?);
    }
}

test "engine: a records part shared by two definitions is a graph that refuses (Codex #205 r3 REL-301)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "s7b4_r3_shared_records.xlsx");
    defer testing.allocator.free(path);
    try fixture.writeWithOrphanCache(testing.allocator, io, path, .sheet_ref);
    {
        var o = try Opened.open(testing.allocator, io, path);
        o.deinit(testing.allocator);
    }
    try fixture.patchPart(testing.allocator, io, path, "xl/pivotCache/_rels/pivotCacheDefinition2.xml.rels", "Target=\"pivotCacheRecords2.xml\"", "Target=\"pivotCacheRecords1.xml\"");
    try testing.expectError(error.MalformedPivotXml, Opened.open(testing.allocator, io, path));
}

test "engine: character data or a comment between the children of a records part, a record or an inventory refuses; a name twice on one tag refuses the part (Codex #205 r4 REL-403, REL-405)" {
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    const main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const head = "<pivotCacheDefinition xmlns=\"" ++ main_ns ++ "\" recordCount=\"1\"><cacheSource type=\"worksheet\"><worksheetSource sheet=\"Data\" ref=\"A1:A2\"/></cacheSource><cacheFields count=\"1\"><cacheField name=\"K\" numFmtId=\"0\">";
    const tail = "</cacheField></cacheFields></pivotCacheDefinition>";
    const def_xml = head ++ "<sharedItems count=\"1\"><s v=\"a\"/></sharedItems>" ++ tail;
    const cache: PivotCache = .{
        .cache_id = null,
        .part_name = "xl/pivotCache/pivotCacheDefinition1.xml",
        .records_part_name = "xl/pivotCache/pivotCacheRecords1.xml",
        .definition = try pivot_xml.parseCacheDefinition(arena, def_xml),
        .field_names = &.{"K"},
        .field_formulas = &.{null},
        .source = .{},
        .resolution = .none,
        .range_set_sources = &.{},
        .range_set_resolutions = &.{},
        .consumer_count = 0,
        .raw_xml = def_xml,
    };
    const rows = [_]engine.Row{&.{.{ .string = "a" }}};
    const rec_head = "<pivotCacheRecords xmlns=\"" ++ main_ns ++ "\" count=\"1\">";
    const lossy = [_][]const u8{
        rec_head ++ "opaque<r><x v=\"0\"/></r></pivotCacheRecords>",
        rec_head ++ "<r>head<x v=\"0\"/></r></pivotCacheRecords>",
        rec_head ++ "<r><x v=\"0\"/>tail</r></pivotCacheRecords>",
        rec_head ++ "<r><x v=\"0\"/></r>trailer</pivotCacheRecords>",
        rec_head ++ "<r><x v=\"0\"/></r><!-- lost --></pivotCacheRecords>",
        rec_head ++ "<r><x v=\"0\"/><![CDATA[lost]]></r></pivotCacheRecords>",
    };
    for (lossy) |xml| try testing.expectError(error.PivotShapeUnsupported, engine.rebuild(arena, &cache, &rows, xml, null));
    // Whitespace between and around is nothing.
    const spaced = "<pivotCacheRecords xmlns=\"" ++ main_ns ++ "\" count=\"1\">\n  <r>\n    <x v=\"0\"/>\n  </r>\n</pivotCacheRecords>";
    const rb = try engine.rebuild(arena, &cache, &rows, spaced, null);
    try testing.expectEqualStrings(rec_head ++ "<r><x v=\"0\"/></r></pivotCacheRecords>", rb.records.?);
    // An inventory with text between its items.
    const chatty = try pivot_xml.parseCacheDefinition(arena, head ++ "<sharedItems count=\"1\"><s v=\"a\"/>junk</sharedItems>" ++ tail);
    try testing.expectError(error.PivotShapeUnsupported, engine.checkShape(&chatty));
    // A duplicated attribute, on either root.
    try testing.expectError(error.MalformedPivotXml, engine.rebuild(arena, &cache, &rows, "<pivotCacheRecords xmlns=\"" ++ main_ns ++ "\" count=\"1\" count=\"2\"><r><x v=\"0\"/></r></pivotCacheRecords>", null));
    try testing.expectError(error.MalformedXml, pivot_xml.parseCacheDefinition(arena, "<pivotCacheDefinition xmlns=\"" ++ main_ns ++ "\" recordCount=\"1\" recordCount=\"2\"><cacheSource type=\"worksheet\"><worksheetSource sheet=\"Data\" ref=\"A1:A2\"/></cacheSource><cacheFields count=\"1\"><cacheField name=\"K\" numFmtId=\"0\"><sharedItems count=\"1\"><s v=\"a\"/></sharedItems></cacheField></cacheFields></pivotCacheDefinition>"));
    try testing.expectError(error.MalformedXml, pivot_xml.parseCacheDefinition(arena, head ++ "<sharedItems count=\"1\" count=\"1\"><s v=\"a\"/></sharedItems>" ++ tail));
}

test "engine: an item with a prefixed attribute refuses — the declaration it may hang on lives on the element the rebuild replaces (Codex #205 r5 REL-501)" {
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    const main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const head = "<pivotCacheDefinition xmlns=\"" ++ main_ns ++ "\" xmlns:z=\"urn:vendor\" recordCount=\"1\"><cacheSource type=\"worksheet\"><worksheetSource sheet=\"Data\" ref=\"A1:A2\"/></cacheSource><cacheFields count=\"1\"><cacheField name=\"K\" numFmtId=\"0\">";
    const tail = "</cacheField></cacheFields></pivotCacheDefinition>";
    // Declared on the inventory element, and declared on the root: the
    // slice keeps neither item.
    const on_element = try pivot_xml.parseCacheDefinition(arena, head ++ "<sharedItems xmlns:y=\"urn:other\" count=\"1\"><s y:meta=\"x\" v=\"a\"/></sharedItems>" ++ tail);
    try testing.expectError(error.PivotShapeUnsupported, engine.checkShape(&on_element));
    const on_root = try pivot_xml.parseCacheDefinition(arena, head ++ "<sharedItems count=\"1\"><s z:meta=\"x\" v=\"a\"/></sharedItems>" ++ tail);
    try testing.expectError(error.PivotShapeUnsupported, engine.checkShape(&on_root));
    // A declaration nothing uses is nothing.
    const unused = try pivot_xml.parseCacheDefinition(arena, head ++ "<sharedItems xmlns:y=\"urn:other\" count=\"1\"><s v=\"a\"/></sharedItems>" ++ tail);
    try engine.checkShape(&unused);
}

test "engine: a qualified attribute on the inventory element, a record or a value element refuses; an unqualified one on a value element reads; empty consolidation markup refuses; a tag past the attribute ceiling refuses the part (Codex #205 r6 REL-602, REL-604, PERF-601)" {
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    const main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const head = "<pivotCacheDefinition xmlns=\"" ++ main_ns ++ "\" xmlns:z=\"urn:vendor\" recordCount=\"1\"><cacheSource type=\"worksheet\"><worksheetSource sheet=\"Data\" ref=\"A1:A2\"/></cacheSource><cacheFields count=\"1\"><cacheField name=\"K\" numFmtId=\"0\">";
    const tail = "</cacheField></cacheFields></pivotCacheDefinition>";
    const on_element = try pivot_xml.parseCacheDefinition(arena, head ++ "<sharedItems z:meta=\"x\" count=\"1\"><s v=\"a\"/></sharedItems>" ++ tail);
    try testing.expectError(error.PivotShapeUnsupported, engine.checkShape(&on_element));

    const def_xml = head ++ "<sharedItems count=\"1\"><s v=\"a\"/></sharedItems>" ++ tail;
    const cache: PivotCache = .{
        .cache_id = null,
        .part_name = "xl/pivotCache/pivotCacheDefinition1.xml",
        .records_part_name = "xl/pivotCache/pivotCacheRecords1.xml",
        .definition = try pivot_xml.parseCacheDefinition(arena, def_xml),
        .field_names = &.{"K"},
        .field_formulas = &.{null},
        .source = .{},
        .resolution = .none,
        .range_set_sources = &.{},
        .range_set_resolutions = &.{},
        .consumer_count = 0,
        .raw_xml = def_xml,
    };
    const rows = [_]engine.Row{&.{.{ .string = "a" }}};
    const rec_head = "<pivotCacheRecords xmlns=\"" ++ main_ns ++ "\" xmlns:z=\"urn:vendor\" count=\"1\">";
    const lossy = [_][]const u8{
        rec_head ++ "<r z:meta=\"x\"><x v=\"0\"/></r></pivotCacheRecords>",
        rec_head ++ "<r u=\"1\"><x v=\"0\"/></r></pivotCacheRecords>",
        rec_head ++ "<r><x z:meta=\"x\" v=\"0\"/></r></pivotCacheRecords>",
        "<x:pivotCacheRecords xmlns:x=\"" ++ main_ns ++ "\" count=\"1\"><x:r><x:x x:meta=\"1\" v=\"0\"/></x:r></x:pivotCacheRecords>",
    };
    for (lossy) |xml| try testing.expectError(error.PivotShapeUnsupported, engine.rebuild(arena, &cache, &rows, xml, null));
    const rb = try engine.rebuild(arena, &cache, &rows, rec_head ++ "<r><x v=\"0\" u=\"1\"/></r></pivotCacheRecords>", null);
    try testing.expectEqualStrings(rec_head ++ "<r><x v=\"0\"/></r></pivotCacheRecords>", rb.records.?);

    // Consolidation markup beside a worksheet locator, sets or none.
    const cons = [_][]const u8{
        "<pivotCacheDefinition xmlns=\"" ++ main_ns ++ "\" recordCount=\"1\"><cacheSource type=\"worksheet\"><worksheetSource sheet=\"Data\" ref=\"A1:A2\"/><consolidation/></cacheSource><cacheFields count=\"1\"><cacheField name=\"K\" numFmtId=\"0\"><sharedItems count=\"1\"><s v=\"a\"/></sharedItems></cacheField></cacheFields></pivotCacheDefinition>",
        "<pivotCacheDefinition xmlns=\"" ++ main_ns ++ "\" recordCount=\"1\"><cacheSource type=\"worksheet\"><worksheetSource sheet=\"Data\" ref=\"A1:A2\"/><consolidation><rangeSets count=\"0\"/></consolidation></cacheSource><cacheFields count=\"1\"><cacheField name=\"K\" numFmtId=\"0\"><sharedItems count=\"1\"><s v=\"a\"/></sharedItems></cacheField></cacheFields></pivotCacheDefinition>",
    };
    for (cons) |xml| {
        const def = try pivot_xml.parseCacheDefinition(arena, xml);
        try testing.expect(def.source.has_consolidation);
        try testing.expectError(error.PivotShapeUnsupported, engine.checkShape(&def));
    }

    // The attribute ceiling: 256 unknown attributes read, 257 refuse.
    for ([_]usize{ 254, 255 }) |extra| {
        var xml: std.ArrayListUnmanaged(u8) = .empty;
        try xml.appendSlice(arena, "<pivotCacheDefinition xmlns=\"" ++ main_ns ++ "\" recordCount=\"1\"");
        for (0..extra) |i| try xml.appendSlice(arena, try std.fmt.allocPrint(arena, " a{d}=\"\"", .{i}));
        try xml.appendSlice(arena, "><cacheSource type=\"worksheet\"><worksheetSource sheet=\"Data\" ref=\"A1:A2\"/></cacheSource></pivotCacheDefinition>");
        // xmlns + recordCount + extra: 256 passes, 257 does not.
        if (extra == 254) {
            _ = try pivot_xml.parseCacheDefinition(arena, xml.items);
        } else {
            try testing.expectError(error.MalformedXml, pivot_xml.parseCacheDefinition(arena, xml.items));
        }
    }
}

test "edit: spliceAll answers a reversed, overlapping or out-of-range span with MalformedPivotXml, never an assertion (Codex #205 r7 REL-701)" {
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    const src = "abcdef";
    var reversed = [_]edit.Splice{.{ .span = .{ .start = 3, .end = 2 }, .text = "" }};
    try testing.expectError(error.MalformedPivotXml, edit.spliceAll(arena, src, &reversed));
    var past = [_]edit.Splice{.{ .span = .{ .start = 0, .end = src.len + 1 }, .text = "" }};
    try testing.expectError(error.MalformedPivotXml, edit.spliceAll(arena, src, &past));
    var overlapping = [_]edit.Splice{ .{ .span = .{ .start = 0, .end = 2 }, .text = "X" }, .{ .span = .{ .start = 1, .end = 3 }, .text = "Y" } };
    try testing.expectError(error.MalformedPivotXml, edit.spliceAll(arena, src, &overlapping));
    var fine = [_]edit.Splice{ .{ .span = .{ .start = 4, .end = 6 }, .text = "Z" }, .{ .span = .{ .start = 1, .end = 1 }, .text = "-" }, .{ .span = .{ .start = 1, .end = 3 }, .text = "" } };
    try testing.expectEqualStrings("a-dZ", try edit.spliceAll(arena, src, &fine));
}

test "engine: a retained number spelt with a character reference matches by value; an `xmlnsfoo:` prefix is a prefix (Codex #205 r8 REL-802, REL-804)" {
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    const main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const head = "<pivotCacheDefinition xmlns=\"" ++ main_ns ++ "\" recordCount=\"1\"><cacheSource type=\"worksheet\"><worksheetSource sheet=\"Data\" ref=\"A1:A2\"/></cacheSource><cacheFields count=\"1\"><cacheField name=\"K\" numFmtId=\"0\">";
    const tail = "</cacheField></cacheFields></pivotCacheDefinition>";
    const xml = head ++ "<sharedItems containsSemiMixedTypes=\"0\" containsString=\"0\" containsNumber=\"1\" containsInteger=\"1\" minValue=\"1\" maxValue=\"1\" count=\"1\"><n v=\"&#49;\"/></sharedItems>" ++ tail;
    const cache: PivotCache = .{
        .cache_id = null,
        .part_name = "xl/pivotCache/pivotCacheDefinition1.xml",
        .records_part_name = null,
        .definition = try pivot_xml.parseCacheDefinition(arena, xml),
        .field_names = &.{"K"},
        .field_formulas = &.{null},
        .source = .{},
        .resolution = .none,
        .range_set_sources = &.{},
        .range_set_resolutions = &.{},
        .consumer_count = 0,
        .raw_xml = xml,
    };
    const rows = [_]engine.Row{&.{.{ .number = "1" }}};
    const rb = try engine.rebuild(arena, &cache, &rows, null, null);
    const out = try edit.spliceAll(arena, xml, rb.splices);
    // Matched, not appended: one item, kept as written, the extrema
    // spelt by the item's decoded lexical.
    try testing.expect(std.mem.indexOf(u8, out, "minValue=\"1\" maxValue=\"1\" count=\"1\"><n v=\"&#49;\"/></sharedItems>") != null);

    const sneaky = [_][]const u8{
        head ++ "<sharedItems xmlns:xmlnsfoo=\"urn:vendor\" xmlnsfoo:meta=\"x\" count=\"1\"><s v=\"a\"/></sharedItems>" ++ tail,
        head ++ "<sharedItems xmlns:xmlnsfoo=\"urn:vendor\" count=\"1\"><s xmlnsfoo:meta=\"x\" v=\"a\"/></sharedItems>" ++ tail,
    };
    for (sneaky) |s| {
        const def = try pivot_xml.parseCacheDefinition(arena, s);
        try testing.expectError(error.PivotShapeUnsupported, engine.checkShape(&def));
    }
}

test "engine: parseNumber reads the xsd:double grammar and nothing wider; a rebuild redates or removes refreshedDateIso with its sibling; the records root's suffix is kept (Codex #205 r9 REL-901, REL-902, REL-903)" {
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    for ([_][]const u8{ "1_0", "0x1p0", "1e", "1e+", " 1", "1 ", "1..", ".", "+", "", "inf", "nan", "1.5f", "0x10" }) |bad| try testing.expect(engine.parseNumber(bad) == null);
    for ([_][]const u8{ "+1", ".5", "1.", "1E-3", "-0", "4.4000000000000004", "1e10", "-.5E+2" }) |good| try testing.expect(engine.parseNumber(good) != null);

    const main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const Case = struct { attrs: []const u8, with_clock: []const u8, without: []const u8 };
    const cases = [_]Case{
        .{ .attrs = "refreshedDate=\"45000.5\" refreshedDateIso=\"2023-03-15T12:00:00\" recordCount=\"1\"", .with_clock = "refreshedDate=\"46000.5\" refreshedDateIso=\"2025-12-13T12:00:00\" recordCount=\"1\"", .without = " recordCount=\"1\"" },
        .{ .attrs = "refreshedDateIso=\"2023-03-15T12:00:00\" recordCount=\"1\"", .with_clock = "refreshedDateIso=\"2025-12-13T12:00:00\" recordCount=\"1\" refreshedDate=\"46000.5\"", .without = " recordCount=\"1\"" },
    };
    for (cases) |c| {
        const xml = try std.mem.concat(arena, u8, &.{ "<pivotCacheDefinition xmlns=\"", main_ns, "\" ", c.attrs, "><cacheSource type=\"worksheet\"><worksheetSource sheet=\"Data\" ref=\"A1:A2\"/></cacheSource><cacheFields count=\"1\"><cacheField name=\"K\" numFmtId=\"0\"><sharedItems count=\"1\"><s v=\"a\"/></sharedItems></cacheField></cacheFields></pivotCacheDefinition>" });
        const cache: PivotCache = .{
            .cache_id = null,
            .part_name = "xl/pivotCache/pivotCacheDefinition1.xml",
            .records_part_name = "xl/pivotCache/pivotCacheRecords1.xml",
            .definition = try pivot_xml.parseCacheDefinition(arena, xml),
            .field_names = &.{"K"},
            .field_formulas = &.{null},
            .source = .{},
            .resolution = .none,
            .range_set_sources = &.{},
            .range_set_resolutions = &.{},
            .consumer_count = 0,
            .raw_xml = xml,
        };
        const rows = [_]engine.Row{&.{.{ .string = "a" }}};
        const rec = "<pivotCacheRecords xmlns=\"" ++ main_ns ++ "\" count=\"1\"><r><x v=\"0\"/></r></pivotCacheRecords><!-- keep -->\n<?zlsx trailing?>\n";
        const dated = try engine.rebuild(arena, &cache, &rows, rec, .{ .serial = 46000.5, .iso = "2025-12-13T12:00:00" });
        const dated_out = try edit.spliceAll(arena, xml, dated.splices);
        try testing.expect(std.mem.indexOf(u8, dated_out, c.with_clock) != null);
        try testing.expectEqualStrings(rec, dated.records.?);
        const undated = try engine.rebuild(arena, &cache, &rows, rec, null);
        const undated_out = try edit.spliceAll(arena, xml, undated.splices);
        try testing.expect(std.mem.indexOf(u8, undated_out, "refreshedDate") == null);
        try testing.expect(std.mem.indexOf(u8, undated_out, c.without) != null);
    }
    // A self-closing records root followed by a comment.
    {
        const xml = "<pivotCacheDefinition xmlns=\"" ++ main_ns ++ "\" recordCount=\"0\"><cacheSource type=\"worksheet\"><worksheetSource sheet=\"Data\" ref=\"A1:A2\"/></cacheSource><cacheFields count=\"1\"><cacheField name=\"K\" numFmtId=\"0\"><sharedItems count=\"1\"><s v=\"a\"/></sharedItems></cacheField></cacheFields></pivotCacheDefinition>";
        const cache: PivotCache = .{
            .cache_id = null,
            .part_name = "xl/pivotCache/pivotCacheDefinition1.xml",
            .records_part_name = "xl/pivotCache/pivotCacheRecords1.xml",
            .definition = try pivot_xml.parseCacheDefinition(arena, xml),
            .field_names = &.{"K"},
            .field_formulas = &.{null},
            .source = .{},
            .resolution = .none,
            .range_set_sources = &.{},
            .range_set_resolutions = &.{},
            .consumer_count = 0,
            .raw_xml = xml,
        };
        const rows = [_]engine.Row{&.{.{ .string = "a" }}};
        const rb = try engine.rebuild(arena, &cache, &rows, "<pivotCacheRecords xmlns=\"" ++ main_ns ++ "\" count=\"0\"/><!-- keep -->", null);
        try testing.expectEqualStrings("<pivotCacheRecords xmlns=\"" ++ main_ns ++ "\" count=\"1\"><r><x v=\"0\"/></r></pivotCacheRecords><!-- keep -->", rb.records.?);
    }
}

test "engine: the records say what was indexed — an explicit count=\"0\" keeps an inline field inline; a field spelt both ways, or a record of the wrong arity, refuses (Codex #205 r10 REL-1001)" {
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    const main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const xml = "<pivotCacheDefinition xmlns=\"" ++ main_ns ++ "\" recordCount=\"1\"><cacheSource type=\"worksheet\"><worksheetSource sheet=\"Data\" ref=\"A1:B2\"/></cacheSource><cacheFields count=\"2\"><cacheField name=\"K\" numFmtId=\"0\"><sharedItems count=\"1\"><s v=\"a\"/></sharedItems></cacheField><cacheField name=\"V\" numFmtId=\"0\"><sharedItems containsSemiMixedTypes=\"0\" containsString=\"0\" containsNumber=\"1\" containsInteger=\"1\" minValue=\"1\" maxValue=\"1\" count=\"0\"/></cacheField></cacheFields></pivotCacheDefinition>";
    const cache: PivotCache = .{
        .cache_id = null,
        .part_name = "xl/pivotCache/pivotCacheDefinition1.xml",
        .records_part_name = "xl/pivotCache/pivotCacheRecords1.xml",
        .definition = try pivot_xml.parseCacheDefinition(arena, xml),
        .field_names = &.{ "K", "V" },
        .field_formulas = &.{ null, null },
        .source = .{},
        .resolution = .none,
        .range_set_sources = &.{},
        .range_set_resolutions = &.{},
        .consumer_count = 0,
        .raw_xml = xml,
    };
    const rows = [_]engine.Row{ &.{ .{ .string = "a" }, .{ .number = "1" } }, &.{ .{ .string = "b" }, .{ .number = "2" } } };
    const rec_head = "<pivotCacheRecords xmlns=\"" ++ main_ns ++ "\" count=\"1\">";
    // Inline as written: stays inline, the `count="0"` gone with the
    // attribute-only form.
    const rb = try engine.rebuild(arena, &cache, &rows, rec_head ++ "<r><x v=\"0\"/><n v=\"1\"/></r></pivotCacheRecords>", null);
    const out = try edit.spliceAll(arena, xml, rb.splices);
    try testing.expect(std.mem.indexOf(u8, out, "<cacheField name=\"V\" numFmtId=\"0\"><sharedItems containsSemiMixedTypes=\"0\" containsString=\"0\" containsNumber=\"1\" containsInteger=\"1\" minValue=\"1\" maxValue=\"2\"/></cacheField>") != null);
    try testing.expectEqualStrings("<pivotCacheRecords xmlns=\"" ++ main_ns ++ "\" count=\"2\"><r><x v=\"0\"/><n v=\"1\"/></r><r><x v=\"1\"/><n v=\"2\"/></r></pivotCacheRecords>", rb.records.?);
    // Indexed as written (an inventory the definition does not show
    // is not this slice's to invent): refused by the inventory, which
    // has no item for index 0 — the records are read first.
    const indexed = try engine.rebuild(arena, &cache, &rows, rec_head ++ "<r><x v=\"0\"/><x v=\"0\"/></r></pivotCacheRecords>", null);
    const indexed_out = try edit.spliceAll(arena, xml, indexed.splices);
    try testing.expect(std.mem.indexOf(u8, indexed_out, "<cacheField name=\"V\" numFmtId=\"0\"><sharedItems containsSemiMixedTypes=\"0\" containsString=\"0\" containsNumber=\"1\" containsInteger=\"1\" minValue=\"1\" maxValue=\"2\" count=\"2\"><n v=\"1\"/><n v=\"2\"/></sharedItems></cacheField>") != null);
    // Spelt both ways, or the wrong arity: not one shape.
    try testing.expectError(error.PivotShapeUnsupported, engine.rebuild(arena, &cache, &rows, rec_head ++ "<r><x v=\"0\"/><n v=\"1\"/></r><r><x v=\"0\"/><x v=\"0\"/></r></pivotCacheRecords>", null));
    try testing.expectError(error.PivotShapeUnsupported, engine.rebuild(arena, &cache, &rows, rec_head ++ "<r><x v=\"0\"/></r></pivotCacheRecords>", null));
    try testing.expectError(error.PivotShapeUnsupported, engine.rebuild(arena, &cache, &rows, rec_head ++ "<r><x v=\"0\"/><n v=\"1\"/><n v=\"1\"/></r></pivotCacheRecords>", null));
    // No records part: an inventory with items is the enumerated one.
    const none = try engine.rebuild(arena, &cache, &rows, null, null);
    const none_out = try edit.spliceAll(arena, xml, none.splices);
    try testing.expect(std.mem.indexOf(u8, none_out, "minValue=\"1\" maxValue=\"2\"/></cacheField>") != null);
}

test "engine: inline records beside an inventory that holds items refuse — an inline rebuild would drop what a consumer indexes; the same inventory with `<x>` records keeps it (Codex #205 r11 REL-1101)" {
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    const main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const xml = "<pivotCacheDefinition xmlns=\"" ++ main_ns ++ "\" recordCount=\"1\"><cacheSource type=\"worksheet\"><worksheetSource sheet=\"Data\" ref=\"A1:A2\"/></cacheSource><cacheFields count=\"1\"><cacheField name=\"V\" numFmtId=\"0\"><sharedItems containsSemiMixedTypes=\"0\" containsString=\"0\" containsNumber=\"1\" containsInteger=\"1\" minValue=\"99\" maxValue=\"99\" count=\"1\"><n v=\"99\"/></sharedItems></cacheField></cacheFields></pivotCacheDefinition>";
    const cache: PivotCache = .{
        .cache_id = null,
        .part_name = "xl/pivotCache/pivotCacheDefinition1.xml",
        .records_part_name = "xl/pivotCache/pivotCacheRecords1.xml",
        .definition = try pivot_xml.parseCacheDefinition(arena, xml),
        .field_names = &.{"V"},
        .field_formulas = &.{null},
        .source = .{},
        .resolution = .none,
        .range_set_sources = &.{},
        .range_set_resolutions = &.{},
        .consumer_count = 0,
        .raw_xml = xml,
    };
    const rows = [_]engine.Row{&.{.{ .number = "1" }}};
    const rec_head = "<pivotCacheRecords xmlns=\"" ++ main_ns ++ "\" count=\"1\">";
    try testing.expectError(error.PivotShapeUnsupported, engine.rebuild(arena, &cache, &rows, rec_head ++ "<r><n v=\"1\"/></r></pivotCacheRecords>", null));
    const rb = try engine.rebuild(arena, &cache, &rows, rec_head ++ "<r><x v=\"0\"/></r></pivotCacheRecords>", null);
    const out = try edit.spliceAll(arena, xml, rb.splices);
    try testing.expect(std.mem.indexOf(u8, out, "minValue=\"1\" maxValue=\"99\" count=\"2\"><n v=\"99\"/><n v=\"1\"/></sharedItems>") != null);
    try testing.expectEqualStrings(rec_head ++ "<r><x v=\"1\"/></r></pivotCacheRecords>", rb.records.?);
}

// ─── S7b-5: the consumers — the engine's second slice ────────────────

const sheet_xml_tp = @import("typed_parts/sheet_xml.zig");
const sst_xml_tp = @import("typed_parts/sst_xml.zig");

/// The cell at (row, col) in a parsed sheet, by its `r`.
fn cellAt(view: *const sheet_xml_tp.SheetXml, row: u32, col: u32) ?sheet_xml_tp.Cell {
    for (view.rows) |r| {
        if (r.row_idx != row) continue;
        for (r.cells) |c| {
            const parsed = coords.parseCell(c.ref, .{ .case = .upper_only }) catch continue;
            if (parsed.col.oneBased() == col) return c;
        }
    }
    return null;
}

/// Every cell the layout computed against the cell the host carries:
/// text by the shared string it names, a number by value (the
/// spelling may differ — Excel writes some of its own with seventeen
/// digits — the double may not), a blank by the absence of a value.
fn expectLayoutMatchesHost(alloc: Allocator, lay: engine.Layout, sheet_bytes: []const u8, sst: *const sst_xml_tp.SstXml) !usize {
    var view = try sheet_xml_tp.parse(alloc, sheet_bytes);
    defer view.deinit(alloc);
    var checked: usize = 0;
    for (lay.cells) |c| {
        const cell = cellAt(&view, c.row, c.col) orelse return error.TestExpectedCell;
        switch (c.value) {
            .string => |s| {
                try testing.expectEqual(sheet_xml_tp.CellType.shared_string, cell.cell_type);
                const idx = try std.fmt.parseInt(usize, cell.raw_value.?, 10);
                const text = try sst_xml_tp.decodeText(alloc, sst.entries[idx].plain);
                defer alloc.free(text);
                try testing.expectEqualStrings(s, text);
            },
            .number => |lex| {
                try testing.expectEqual(sheet_xml_tp.CellType.number, cell.cell_type);
                const want = try std.fmt.parseFloat(f64, cell.raw_value.?);
                const got = try std.fmt.parseFloat(f64, lex);
                try testing.expectEqual(want, got);
            },
            .blank => try testing.expect(cell.raw_value == null),
        }
        checked += 1;
    }
    return checked;
}

test "S7b-5 oracle: the corpus pivots re-laid from their own records reproduce every cell Excel wrote, and the parts byte for byte" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const path = "tests/corpus/openxlsx_loadExample.xlsx";
    std.Io.Dir.cwd().access(io, path, .{}) catch return error.SkipZigTest;

    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    const sst_part = (try o.store.part("xl/sharedStrings.xml")).?;
    var sst = try sst_xml_tp.parse(testing.allocator, sst_part.bytes);
    defer sst.deinit(testing.allocator);

    // PivotTable1: four sums over Species (50 records); PivotTable3: an
    // average, a min and a max over cyl (29 records). Both compact,
    // one row field, the values axis across, a grand total.
    var checked: usize = 0;
    for (o.pivots.tables) |t| {
        const cache = o.pivots.cacheOf(t).?;
        const rec = (try o.store.part(cache.records_part_name.?)).?.bytes;
        const rows = try engine.rowsFromRecords(arena, cache, rec);
        const rb = try engine.rebuild(arena, cache, rows, rec, null);
        const lay = try engine.layout(arena, t.raw_xml, cache, &rb, .{});
        try testing.expect(lay.rect.eql(lay.old_rect));
        try testing.expectEqualStrings(t.raw_xml, lay.table_xml);
        const sheet_bytes = (try o.store.part(t.sheet_part_name)).?.bytes;
        checked += try expectLayoutMatchesHost(testing.allocator, lay, sheet_bytes, &sst);
    }
    // 5 × 5 cells on `IrisSample`, 5 × 4 on `mtCars Pivot`.
    try testing.expectEqual(@as(usize, 45), checked);
}

test "S7b-5: a caller-assembled Rebuild that disagrees with itself is malformed, never indexed (Codex #206 r7 REL-701)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "s7b5_r7_rebuild.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .sheet_ref);
    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    const cache = &o.pivots.caches[0];
    const t = o.pivots.tables[0];
    const rec = (try o.store.part(cache.records_part_name.?)).?.bytes;
    const good = try engine.rebuild(arena, cache, &fixture_rows, rec, null);
    _ = try engine.layout(arena, t.raw_xml, cache, &good, .{});

    // A row index past the inventory.
    {
        var rb = good;
        const fields = try arena.dupe(engine.Field, good.fields);
        const idx = try arena.dupe(u32, good.fields[0].index_of_row);
        idx[1] = 99;
        fields[0].index_of_row = idx;
        rb.fields = fields;
        try testing.expectError(error.MalformedPivotXml, engine.layout(arena, t.raw_xml, cache, &rb, .{}));
    }
    // Fewer indices than rows.
    {
        var rb = good;
        const fields = try arena.dupe(engine.Field, good.fields);
        fields[0].index_of_row = good.fields[0].index_of_row[0..1];
        rb.fields = fields;
        try testing.expectError(error.MalformedPivotXml, engine.layout(arena, t.raw_xml, cache, &rb, .{}));
    }
    // A short row, and a count that is not the rows'.
    {
        var rb = good;
        const rows = try arena.dupe(engine.Row, good.rows);
        rows[2] = rows[2][0..2];
        rb.rows = rows;
        try testing.expectError(error.MalformedPivotXml, engine.layout(arena, t.raw_xml, cache, &rb, .{}));
        var counted = good;
        counted.record_count = 7;
        try testing.expectError(error.MalformedPivotXml, engine.layout(arena, t.raw_xml, cache, &counted, .{}));
    }
    // An inline field's row that is not a number.
    {
        var rb = good;
        const rows = try arena.dupe(engine.Row, good.rows);
        const row = try arena.dupe(engine.Value, rows[0]);
        row[1] = .{ .number = "0x1p0" };
        rows[0] = row;
        rb.rows = rows;
        try testing.expectError(error.MalformedPivotXml, engine.layout(arena, t.raw_xml, cache, &rb, .{}));
    }
    // No records at all (Codex #206 r14 REL-1401).
    {
        var rb = good;
        rb.rows = &.{};
        rb.record_count = 0;
        const fields = try arena.dupe(engine.Field, good.fields);
        for (fields) |*f| f.index_of_row = &.{};
        rb.fields = fields;
        try testing.expectError(error.MalformedPivotXml, engine.layout(arena, t.raw_xml, cache, &rb, .{}));
    }
    // A cache whose decoded names are not parallel to its fields, and
    // a data field with no caption to fall back on (Codex #206 r12
    // REL-1203).
    {
        var short = cache.*;
        short.field_names = cache.field_names[0..1];
        const nameless = try std.mem.replaceOwned(u8, arena, t.raw_xml, "<dataField name=\"Sum of Qty\" ", "<dataField ");
        try testing.expectError(error.MalformedPivotXml, engine.layout(arena, nameless, &short, &good, .{}));
        try testing.expectError(error.MalformedPivotXml, engine.layout(arena, t.raw_xml, &short, &good, .{}));
    }
    // A row whose index names another item than its value, an inline
    // field with an inventory (Codex #206 r10 REL-1003).
    {
        var rb = good;
        const fields = try arena.dupe(engine.Field, good.fields);
        const idx = try arena.dupe(u32, good.fields[0].index_of_row);
        idx[1] = 0; // `West` indexed as `East`
        fields[0].index_of_row = idx;
        rb.fields = fields;
        try testing.expectError(error.MalformedPivotXml, engine.layout(arena, t.raw_xml, cache, &rb, .{}));
        var inline_items = good;
        const f2 = try arena.dupe(engine.Field, good.fields);
        f2[1].items = good.fields[0].items;
        inline_items.fields = f2;
        try testing.expectError(error.MalformedPivotXml, engine.layout(arena, t.raw_xml, cache, &inline_items, .{}));
    }
    // An indexed numeric item whose spelling is not the grammar's, or
    // not its value (Codex #206 r8 REL-802).
    {
        const bad = [_]engine.Field.Item{
            .{ .raw = null, .kind = .n, .lex = "0x1p0", .num = 1 },
            .{ .raw = null, .kind = .n, .lex = "1", .num = 2 },
            .{ .raw = null, .kind = .n, .lex = "1", .num = std.math.inf(f64) },
        };
        for (bad) |item| {
            var rb = good;
            const fields = try arena.dupe(engine.Field, good.fields);
            const items = try arena.dupe(engine.Field.Item, good.fields[0].items);
            items[0] = item;
            fields[0].items = items;
            rb.fields = fields;
            try testing.expectError(error.MalformedPivotXml, engine.layout(arena, t.raw_xml, cache, &rb, .{}));
        }
    }
}

fn layoutForFailures(allocator: Allocator, table_xml: []const u8, cache: *const PivotCache, rb: *const engine.Rebuild) !void {
    var arena_state = std.heap.ArenaAllocator.init(allocator);
    defer arena_state.deinit();
    _ = try engine.layout(arena_state.allocator(), table_xml, cache, rb, .{});
}

test "S7b-5: an allocation failure anywhere in a layout is OutOfMemory, never a refusal (Codex #206 r11 REL-1101)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "s7b5_r11_oom.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .sheet_ref);
    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    const cache = &o.pivots.caches[0];
    const t = o.pivots.tables[0];
    const rec = (try o.store.part(cache.records_part_name.?)).?.bytes;
    const rb = try engine.rebuild(arena, cache, &fixture_rows, rec, null);
    try testing.checkAllAllocationFailures(testing.allocator, layoutForFailures, .{ t.raw_xml, cache, &rb });
}

// ─── S7c: column edits inside a source — the schema edits ────────────

test "S7c edit: a column edit strictly inside a finite source is the plan's schema edit — the ref respells, K1/K5 refuse, outside still shifts" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "s7c_plan.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .sheet_ref);
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        const cache = &o.pivots.caches[0];
        // Delete inside: the plan carries the ordinal, the shrunk ref
        // and the rebuild rectangle.
        var plan = try edit.planCacheEdit(arena, cache, 0, .col, 3, .delete);
        try testing.expect(plan.changed);
        try testing.expect(plan.rebuild != null);
        try testing.expectEqual(@as(u32, 2), plan.schema.?.remove);
        var found_ref = false;
        for (plan.splices.items) |sp| {
            if (std.mem.eql(u8, sp.text, "A1:B4")) found_ref = true;
        }
        try testing.expect(found_ref);
        // The first column is ordinal 0.
        plan = try edit.planCacheEdit(arena, cache, 0, .col, 1, .delete);
        try testing.expectEqual(@as(u32, 0), plan.schema.?.remove);
        // K1: an insert inside a headered source refuses — the new
        // header cell is blank, and Excel's own refresh fails on it.
        try testing.expectError(error.PivotSourceEditUnsafe, edit.planCacheEdit(arena, cache, 0, .col, 2, .insert));
        try testing.expectError(error.PivotSourceEditUnsafe, edit.planCacheEdit(arena, cache, 0, .col, 3, .insert));
        // Outside: right of the rectangle is a no-op; at the left edge
        // an insert is a pure shift, as before.
        try testing.expect((try cacheEdit(testing.allocator, &o, 0, .col, 4, .delete)) == null);
        const shifted = (try cacheEdit(testing.allocator, &o, 0, .col, 1, .insert)).?;
        defer testing.allocator.free(shifted);
        try testing.expect(std.mem.indexOf(u8, shifted, "ref=\"B1:D4\"") != null);
        // The marker-only seam cannot rebuild a schema: it refuses.
        try testing.expectError(error.PivotSourceEditUnsafe, cacheEdit(testing.allocator, &o, 0, .col, 3, .delete));
    }
    // K5: a delete that would collapse the rectangle refuses.
    try fixture.patchPart(testing.allocator, io, path, "xl/pivotCache/pivotCacheDefinition1.xml", "sheet=\"Data\" ref=\"A1:C4\"", "sheet=\"Data\" ref=\"C1:C4\"");
    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    try testing.expectError(error.PivotSourceEditUnsafe, edit.planCacheEdit(arena, &o.pivots.caches[0], 0, .col, 3, .delete));
}

test "S7c engine: a remove rebuild takes the cacheField out whole — the count follows, the records narrow, every other inventory holds" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "s7c_remove.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .sheet_ref);
    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    const cache = &o.pivots.caches[0];
    var plan = try edit.planCacheEdit(arena, cache, 0, .col, 3, .delete);
    const rows = [_]engine.Row{
        &.{ .{ .string = "East" }, .{ .number = "3" } },
        &.{ .{ .string = "West" }, .{ .number = "4" } },
        &.{ .{ .string = "East" }, .{ .number = "5" } },
    };
    const rec = (try o.store.part(cache.records_part_name.?)).?.bytes;
    const rb = try engine.rebuildWith(arena, cache, &rows, rec, .{ .serial = 46001, .iso = "2025-12-14T00:00:00" }, plan.schema);
    try plan.splices.appendSlice(arena, rb.splices);
    const def = (try edit.applyPlan(testing.allocator, cache, &plan)).?;
    defer testing.allocator.free(def);
    try testing.expect(std.mem.indexOf(u8, def, "Price") == null);
    try testing.expect(std.mem.indexOf(u8, def, "<cacheFields count=\"2\">") != null);
    try testing.expect(std.mem.indexOf(u8, def, "ref=\"A1:B4\"") != null);
    try testing.expect(std.mem.indexOf(u8, def, "refreshOnLoad=\"1\"") != null);
    try testing.expect(std.mem.indexOf(u8, def, "<cacheField name=\"Region\" numFmtId=\"0\">") != null);
    try expectRecords(testing.allocator, rb.records.?, "3", "<r><x v=\"0\"/><n v=\"3\"/></r><r><x v=\"1\"/><n v=\"4\"/></r><r><x v=\"0\"/><n v=\"5\"/></r>");
    // The effective width is the contract: pre-edit rows refuse.
    try testing.expectError(error.PivotShapeUnsupported, engine.rebuildWith(arena, cache, &fixture_rows, rec, null, plan.schema));
}

test "S7c engine: an insert rebuild lands the new field named and blank-inventoried; an empty or misplaced name refuses" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "s7c_insert.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .sheet_ref);
    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    const cache = &o.pivots.caches[0];
    const rows = [_]engine.Row{
        &.{ .{ .string = "East" }, .blank, .{ .number = "3" }, .{ .number = "1.5" } },
        &.{ .{ .string = "West" }, .blank, .{ .number = "4" }, .{ .number = "2.5" } },
        &.{ .{ .string = "East" }, .blank, .{ .number = "5" }, .{ .number = "3.5" } },
    };
    const rec = (try o.store.part(cache.records_part_name.?)).?.bytes;
    const rb = try engine.rebuildWith(arena, cache, &rows, rec, null, .{ .insert = .{ .at = 1, .name = "Column4" } });
    const def = try edit.spliceAll(arena, cache.raw_xml, rb.splices);
    try testing.expect(std.mem.indexOf(u8, def, "<cacheFields count=\"4\">") != null);
    try testing.expect(std.mem.indexOf(u8, def, "<cacheField name=\"Column4\" numFmtId=\"0\"><sharedItems containsNonDate=\"0\" containsString=\"0\" containsBlank=\"1\" count=\"1\"><m/></sharedItems></cacheField><cacheField name=\"Qty\"") != null);
    try expectRecords(testing.allocator, rb.records.?, "3", "<r><x v=\"0\"/><x v=\"0\"/><n v=\"3\"/><n v=\"1.5\"/></r><r><x v=\"1\"/><x v=\"0\"/><n v=\"4\"/><n v=\"2.5\"/></r><r><x v=\"0\"/><x v=\"0\"/><n v=\"5\"/><n v=\"3.5\"/></r>");
    // The name is the table's to give: empty refuses, as does an
    // ordinal outside the written list.
    try testing.expectError(error.MalformedPivotXml, engine.rebuildWith(arena, cache, &rows, rec, null, .{ .insert = .{ .at = 1 } }));
    try testing.expectError(error.MalformedPivotXml, engine.rebuildWith(arena, cache, &rows, rec, null, .{ .insert = .{ .at = 9, .name = "X" } }));
}

test "S7c edit: applyConsumerSchemaEdit — an unreferenced field leaves whole, every referenced form refuses, the inserted one lands bare" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "s7c_consumer.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .sheet_ref);
    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    const t_xml = o.pivots.tables[0].raw_xml;
    // Remove the unreferenced Price (ordinal 2): its pivotField leaves,
    // the count follows, the row and data ordinals hold still.
    const out = try edit.applyConsumerSchemaEdit(arena, t_xml, .{ .remove = 2 });
    try testing.expect(std.mem.indexOf(u8, out, "<pivotFields count=\"2\">") != null);
    try testing.expectEqual(@as(usize, 2), std.mem.count(u8, out, "<pivotField "));
    try testing.expect(std.mem.indexOf(u8, out, "<field x=\"0\"/>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "fld=\"1\"") != null);
    // K4b: the fixture's one data field (dropping it would leave no
    // values column), the row field, and a surviving baseField all
    // refuse; K4a (one of >= 2 data fields) lifts in its own test.
    try testing.expectError(error.PivotSourceEditUnsafe, edit.applyConsumerSchemaEdit(arena, t_xml, .{ .remove = 1 }));
    try testing.expectError(error.PivotSourceEditUnsafe, edit.applyConsumerSchemaEdit(arena, t_xml, .{ .remove = 0 }));
    const based = try replacedOnce(arena, t_xml, "baseField=\"0\"", "baseField=\"2\"");
    try testing.expectError(error.PivotSourceEditUnsafe, edit.applyConsumerSchemaEdit(arena, based, .{ .remove = 2 }));
    // Insert at 1: a bare pivotField between the row field and the
    // data field; `fld` moves up, `x` and `baseField` hold.
    const ins = try edit.applyConsumerSchemaEdit(arena, t_xml, .{ .insert = .{ .at = 1, .name = "Column4" } });
    try testing.expect(std.mem.indexOf(u8, ins, "<pivotFields count=\"4\">") != null);
    try testing.expect(std.mem.indexOf(u8, ins, "</pivotField><pivotField showAll=\"0\"/><pivotField dataField=\"1\" showAll=\"0\"/>") != null);
    try testing.expect(std.mem.indexOf(u8, ins, "fld=\"2\"") != null);
    try testing.expect(std.mem.indexOf(u8, ins, "baseField=\"0\"") != null);
    try testing.expect(std.mem.indexOf(u8, ins, "<field x=\"0\"/>") != null);
    // A baseField right of the insertion moves with it.
    const based_up = try replacedOnce(arena, t_xml, "baseField=\"0\"", "baseField=\"2\"");
    const ins2 = try edit.applyConsumerSchemaEdit(arena, based_up, .{ .insert = .{ .at = 1, .name = "Column4" } });
    try testing.expect(std.mem.indexOf(u8, ins2, "baseField=\"3\"") != null);
}

test "S7c engine: rowsAfterColEdit and effectiveCacheFields hold their bounds" {
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    const dropped = try engine.rowsAfterColEdit(arena, &fixture_rows, 3, 1, .delete);
    try testing.expectEqual(@as(usize, 3), dropped.len);
    try testing.expectEqual(@as(usize, 2), dropped[0].len);
    try testing.expectEqualStrings("1.5", dropped[0][1].number);
    const grown = try engine.rowsAfterColEdit(arena, &fixture_rows, 3, 2, .insert);
    try testing.expectEqual(@as(usize, 4), grown[1].len);
    try testing.expect(grown[1][2] == .blank);
    try testing.expectEqualStrings("2.5", grown[1][3].number);
    try testing.expectError(error.MalformedPivotXml, engine.rowsAfterColEdit(arena, &fixture_rows, 3, 3, .delete));
    try testing.expectError(error.MalformedPivotXml, engine.rowsAfterColEdit(arena, &fixture_rows, 3, 0, .insert));
    const two = [_]pivot_xml.CacheField{ .{ .name = "A" }, .{ .name = "B" } };
    const one = try engine.effectiveCacheFields(arena, &two, .{ .remove = 1 });
    try testing.expectEqual(@as(usize, 1), one.len);
    try testing.expectEqualStrings("A", one[0].name);
    try testing.expectError(error.MalformedPivotXml, engine.effectiveCacheFields(arena, &two, .{ .remove = 2 }));
    try testing.expectError(error.MalformedPivotXml, engine.effectiveCacheFields(arena, &two, .{ .insert = .{ .at = 0, .name = "X" } }));
    const three = try engine.effectiveCacheFields(arena, &two, .{ .insert = .{ .at = 1, .name = "X" } });
    try testing.expectEqualStrings("X", three[1].name);
    try testing.expectEqualStrings("B", three[2].name);
}

const s7c_row_axis_high =
    \\<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="P" cacheId="7"><location ref="A3:B6" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="3"><pivotField showAll="0"/><pivotField dataField="1" showAll="0"/><pivotField axis="axisRow" showAll="0"><items count="3"><item x="0"/><item x="1"/><item t="default"/></items></pivotField></pivotFields><rowFields count="1"><field x="2"/></rowFields><rowItems count="3"><i><x/></i><i><x v="1"/></i><i t="grand"><x/></i></rowItems><colItems count="1"><i/></colItems><dataFields count="1"><dataField name="Sum of Qty" fld="1" baseField="2" baseItem="0"/></dataFields></pivotTableDefinition>
;

test "S7c edit: every ordinal carrier above the edit moves — the row axis and baseField included; tolerated-unread content refuses (in-house S7C-R1, S7C-MUT-1)" {
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    // Remove the unreferenced ordinal 0: the row axis, the data field
    // and its baseField all sit above and decrement.
    const removed = try edit.applyConsumerSchemaEdit(arena, s7c_row_axis_high, .{ .remove = 0 });
    try testing.expect(std.mem.indexOf(u8, removed, "<rowFields count=\"1\"><field x=\"1\"/></rowFields>") != null);
    try testing.expect(std.mem.indexOf(u8, removed, "fld=\"0\" baseField=\"1\"") != null);
    try testing.expect(std.mem.indexOf(u8, removed, "<pivotFields count=\"2\">") != null);
    // Insert at 1: the same carriers increment.
    const grown = try edit.applyConsumerSchemaEdit(arena, s7c_row_axis_high, .{ .insert = .{ .at = 1, .name = "N" } });
    try testing.expect(std.mem.indexOf(u8, grown, "<rowFields count=\"1\"><field x=\"3\"/></rowFields>") != null);
    try testing.expect(std.mem.indexOf(u8, grown, "fld=\"2\" baseField=\"3\"") != null);
    try testing.expect(std.mem.indexOf(u8, grown, "<pivotFields count=\"4\">") != null);
    // A wrapper count that disagrees with its children refuses where
    // the splice would have healed the evidence.
    const lying = try replacedOnce(arena, s7c_row_axis_high, "<pivotFields count=\"3\">", "<pivotFields count=\"9\">");
    try testing.expectError(error.MalformedPivotXml, edit.applyConsumerSchemaEdit(arena, lying, .{ .remove = 0 }));
    // A SURVIVING field's body beyond its items may name ordinals this
    // rewrite cannot see — refused; the removed field's body leaves
    // whole with it.
    const bodied = try replacedOnce(arena, s7c_row_axis_high, "<pivotField showAll=\"0\"/>", "<pivotField showAll=\"0\"><autoSortScope><pivotArea type=\"normal\"/></autoSortScope></pivotField>");
    try testing.expectError(error.PivotShapeUnsupported, edit.applyConsumerSchemaEdit(arena, bodied, .{ .insert = .{ .at = 1, .name = "N" } }));
    _ = try edit.applyConsumerSchemaEdit(arena, bodied, .{ .remove = 0 });
    // The root extLst: attribute-only extension content admits (the
    // corpus' x14:pivotTableDefinition), an ordinal-carrier token
    // refuses.
    const ext_ok = try replacedOnce(arena, s7c_row_axis_high, "</pivotTableDefinition>", "<extLst><ext uri=\"{X}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:pivotTableDefinition hideValuesRow=\"1\"/></ext></extLst></pivotTableDefinition>");
    _ = try edit.applyConsumerSchemaEdit(arena, ext_ok, .{ .remove = 0 });
    const ext_bad = try replacedOnce(arena, ext_ok, "hideValuesRow=\"1\"", "hideValuesRow=\"1\" field=\"2\"");
    try testing.expectError(error.PivotShapeUnsupported, edit.applyConsumerSchemaEdit(arena, ext_bad, .{ .remove = 0 }));
    // The probe matches the attribute NAME, whitespace around `=`
    // included; `fieldPosition=` is another name (Codex #208 r1
    // SEC-101).
    const ext_ws = try replacedOnce(arena, ext_ok, "hideValuesRow=\"1\"", "hideValuesRow=\"1\" field = '2'");
    try testing.expectError(error.PivotShapeUnsupported, edit.applyConsumerSchemaEdit(arena, ext_ws, .{ .remove = 0 }));
    const ext_pos = try replacedOnce(arena, ext_ok, "hideValuesRow=\"1\"", "hideValuesRow=\"1\" fieldPosition=\"0\"");
    _ = try edit.applyConsumerSchemaEdit(arena, ext_pos, .{ .remove = 0 });
    // An ordinal past the field list refuses before any arithmetic on
    // it (Codex #208 r1 SEC-104).
    const wild = try replacedOnce(arena, s7c_row_axis_high, "fld=\"1\"", "fld=\"4294967295\"");
    try testing.expectError(error.MalformedPivotXml, edit.applyConsumerSchemaEdit(arena, wild, .{ .insert = .{ .at = 1, .name = "N" } }));
    try testing.expectError(error.MalformedPivotXml, edit.applyConsumerSchemaEdit(arena, wild, .{ .remove = 0 }));
    // The two directions must agree on every SURVIVING field too
    // (Codex #208 r4 REL-401): a data field whose pivotField lacks the
    // flag, and a flagged pivotField no data field reads, both refuse
    // whichever ordinal the edit touches.
    const flagless = try replacedOnce(arena, s7c_row_axis_high, "<pivotField dataField=\"1\" showAll=\"0\"/>", "<pivotField showAll=\"0\"/>");
    try testing.expectError(error.MalformedPivotXml, edit.applyConsumerSchemaEdit(arena, flagless, .{ .remove = 0 }));
    const readless = try replacedOnce(arena, s7c_row_axis_high, "<pivotField showAll=\"0\"/>", "<pivotField dataField=\"1\" showAll=\"0\"/>");
    try testing.expectError(error.MalformedPivotXml, edit.applyConsumerSchemaEdit(arena, readless, .{ .insert = .{ .at = 1, .name = "N" } }));
}

const s7c_multi_data =
    \\<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="P" cacheId="7"><location ref="B2:E7" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="4"><pivotField dataField="1" showAll="0"/><pivotField axis="axisRow" showAll="0"><items count="3"><item x="0"/><item x="1"/><item t="default"/></items></pivotField><pivotField dataField="1" showAll="0"/><pivotField showAll="0"/></pivotFields><rowFields count="1"><field x="1"/></rowFields><rowItems count="3"><i><x/></i><i><x v="1"/></i><i t="grand"><x/></i></rowItems><colFields count="1"><field x="-2"/></colFields><colItems count="3"><i><x/></i><i i="1"><x v="1"/></i><i i="2"><x v="2"/></i></colItems><dataFields count="3"><dataField name="Sum of A" fld="0" baseField="1" baseItem="0"/><dataField name="Min of C" fld="2" subtotal="min" baseField="1" baseItem="0"/><dataField name="Max of C" fld="2" subtotal="max" baseField="1" baseItem="0"/></dataFields></pivotTableDefinition>
;

test "S7c-2 edit: K4a — a data-field delete drops the dataField, re-enumerates the values axis and narrows the location; one survivor collapses the axis; K4b keeps refusing" {
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    // Remove ordinal 0 (`Sum of A`): two data fields survive — the
    // values axis stays across, `<colItems>` re-enumerates, the
    // location loses one column, every carrier above 0 decrements.
    const two = try edit.applyConsumerSchemaEdit(arena, s7c_multi_data, .{ .remove = 0 });
    try testing.expect(std.mem.indexOf(u8, two, "<pivotFields count=\"3\">") != null);
    try testing.expect(std.mem.indexOf(u8, two, "Sum of A") == null);
    try testing.expect(std.mem.indexOf(u8, two, "<rowFields count=\"1\"><field x=\"0\"/></rowFields>") != null);
    try testing.expect(std.mem.indexOf(u8, two, "<dataFields count=\"2\"><dataField name=\"Min of C\" fld=\"1\" subtotal=\"min\" baseField=\"0\" baseItem=\"0\"/><dataField name=\"Max of C\" fld=\"1\" subtotal=\"max\" baseField=\"0\" baseItem=\"0\"/></dataFields>") != null);
    try testing.expect(std.mem.indexOf(u8, two, "<colFields count=\"1\"><field x=\"-2\"/></colFields>") != null);
    try testing.expect(std.mem.indexOf(u8, two, "<colItems count=\"2\"><i><x/></i><i i=\"1\"><x v=\"1\"/></i></colItems>") != null);
    try testing.expect(std.mem.indexOf(u8, two, "<location ref=\"B2:D7\"") != null);
    // Remove ordinal 2 (`C`, two data fields at once): one survivor —
    // the values axis leaves the columns whole (Excel's own
    // one-data-field spelling), the single bare item stands, the
    // location loses two.
    const one = try edit.applyConsumerSchemaEdit(arena, s7c_multi_data, .{ .remove = 2 });
    try testing.expect(std.mem.indexOf(u8, one, "<pivotFields count=\"3\">") != null);
    try testing.expect(std.mem.indexOf(u8, one, "<colFields") == null);
    try testing.expect(std.mem.indexOf(u8, one, "<colItems count=\"1\"><i/></colItems>") != null);
    try testing.expect(std.mem.indexOf(u8, one, "<dataFields count=\"1\"><dataField name=\"Sum of A\" fld=\"0\" baseField=\"1\" baseItem=\"0\"/></dataFields>") != null);
    try testing.expect(std.mem.indexOf(u8, one, "<location ref=\"B2:C7\"") != null);
    try testing.expect(std.mem.indexOf(u8, one, "<rowFields count=\"1\"><field x=\"1\"/></rowFields>") != null);
    // K4b: the row axis; a surviving baseField naming the ordinal.
    try testing.expectError(error.PivotSourceEditUnsafe, edit.applyConsumerSchemaEdit(arena, s7c_multi_data, .{ .remove = 1 }));
    const based = try replacedOnce(arena, s7c_multi_data, "fld=\"0\" baseField=\"1\"", "fld=\"0\" baseField=\"2\"");
    try testing.expectError(error.PivotSourceEditUnsafe, edit.applyConsumerSchemaEdit(arena, based, .{ .remove = 2 }));
    // The K4a gates refuse rather than heal what the layout's own
    // checks would refuse on (the S7C-MUT-1 rule) — and only the K4a
    // path reads them: the unreferenced ordinal 3 narrows no axis and
    // still lifts, the layout refusing the part later on its own.
    const lying_count = try replacedOnce(arena, s7c_multi_data, "<colItems count=\"3\">", "<colItems count=\"4\">");
    try testing.expectError(error.MalformedPivotXml, edit.applyConsumerSchemaEdit(arena, lying_count, .{ .remove = 0 }));
    _ = try edit.applyConsumerSchemaEdit(arena, lying_count, .{ .remove = 3 });
    const attributed = try replacedOnce(arena, s7c_multi_data, "<colItems count=\"3\">", "<colItems count=\"3\" grandTotalCaption=\"x\">");
    try testing.expectError(error.PivotShapeUnsupported, edit.applyConsumerSchemaEdit(arena, attributed, .{ .remove = 0 }));
    const shuffled = try replacedOnce(arena, s7c_multi_data, "<i i=\"1\"><x v=\"1\"/></i>", "<i i=\"1\"><x v=\"2\"/></i>");
    try testing.expectError(error.PivotShapeUnsupported, edit.applyConsumerSchemaEdit(arena, shuffled, .{ .remove = 0 }));
    const wide = try replacedOnce(arena, s7c_multi_data, "ref=\"B2:E7\"", "ref=\"B2:F7\"");
    try testing.expectError(error.MalformedPivotXml, edit.applyConsumerSchemaEdit(arena, wide, .{ .remove = 0 }));
    _ = try edit.applyConsumerSchemaEdit(arena, wide, .{ .remove = 3 });
    // A stray attribute or child on an axis wrapper is evidence the
    // layout refuses on — the collapse would remove `<colFields>`
    // whole and heal it (Codex #210 r1 REL-101): refused on the K4a
    // path, unread on the K3 path.
    const cf_attr = try replacedOnce(arena, s7c_multi_data, "<colFields count=\"1\">", "<colFields count=\"1\" foo=\"x\">");
    try testing.expectError(error.PivotShapeUnsupported, edit.applyConsumerSchemaEdit(arena, cf_attr, .{ .remove = 2 }));
    try testing.expectError(error.PivotShapeUnsupported, edit.applyConsumerSchemaEdit(arena, cf_attr, .{ .remove = 0 }));
    _ = try edit.applyConsumerSchemaEdit(arena, cf_attr, .{ .remove = 3 });
    const cf_child = try replacedOnce(arena, s7c_multi_data, "<field x=\"-2\"/></colFields>", "<field x=\"-2\"/><!-- c --></colFields>");
    try testing.expectError(error.PivotShapeUnsupported, edit.applyConsumerSchemaEdit(arena, cf_child, .{ .remove = 2 }));
    _ = try edit.applyConsumerSchemaEdit(arena, cf_child, .{ .remove = 3 });
    // A `dataFields` wrapper spelling no `count` narrows without one.
    const countless = try replacedOnce(arena, s7c_multi_data, "<dataFields count=\"3\">", "<dataFields>");
    const nc = try edit.applyConsumerSchemaEdit(arena, countless, .{ .remove = 0 });
    try testing.expect(std.mem.indexOf(u8, nc, "<dataFields><dataField name=\"Min of C\"") != null);
    try testing.expect(std.mem.indexOf(u8, nc, "<location ref=\"B2:D7\"") != null);
}

test "S7c-2 edit: a strict-prefixed part regenerates its values axis under its own prefix" {
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    const prefixed = try std.mem.concat(arena, u8, &.{
        "<x:pivotTableDefinition xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" name=\"P\" cacheId=\"7\"><x:location ref=\"B2:E7\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/><x:pivotFields count=\"4\"><x:pivotField dataField=\"1\" showAll=\"0\"/><x:pivotField axis=\"axisRow\" showAll=\"0\"><x:items count=\"3\"><x:item x=\"0\"/><x:item x=\"1\"/><x:item t=\"default\"/></x:items></x:pivotField><x:pivotField dataField=\"1\" showAll=\"0\"/><x:pivotField showAll=\"0\"/></x:pivotFields><x:rowFields count=\"1\"><x:field x=\"1\"/></x:rowFields><x:rowItems count=\"3\"><x:i><x:x/></x:i><x:i><x:x v=\"1\"/></x:i><x:i t=\"grand\"><x:x/></x:i></x:rowItems><x:colFields count=\"1\"><x:field x=\"-2\"/></x:colFields><x:colItems count=\"3\"><x:i><x:x/></x:i><x:i i=\"1\"><x:x v=\"1\"/></x:i><x:i i=\"2\"><x:x v=\"2\"/></x:i></x:colItems><x:dataFields count=\"3\"><x:dataField name=\"Sum of A\" fld=\"0\" baseField=\"1\" baseItem=\"0\"/><x:dataField name=\"Min of C\" fld=\"2\" subtotal=\"min\" baseField=\"1\" baseItem=\"0\"/><x:dataField name=\"Max of C\" fld=\"2\" subtotal=\"max\" baseField=\"1\" baseItem=\"0\"/></x:dataFields></x:pivotTableDefinition>",
    });
    const two = try edit.applyConsumerSchemaEdit(arena, prefixed, .{ .remove = 0 });
    try testing.expect(std.mem.indexOf(u8, two, "<x:colItems count=\"2\"><x:i><x:x/></x:i><x:i i=\"1\"><x:x v=\"1\"/></x:i></x:colItems>") != null);
    try testing.expect(std.mem.indexOf(u8, two, "<x:colFields count=\"1\"><x:field x=\"-2\"/></x:colFields>") != null);
    try testing.expect(std.mem.indexOf(u8, two, "ref=\"B2:D7\"") != null);
    const one = try edit.applyConsumerSchemaEdit(arena, prefixed, .{ .remove = 2 });
    try testing.expect(std.mem.indexOf(u8, one, "<x:colFields") == null);
    try testing.expect(std.mem.indexOf(u8, one, "<x:colItems count=\"1\"><x:i/></x:colItems>") != null);
    try testing.expect(std.mem.indexOf(u8, one, "ref=\"B2:C7\"") != null);
}

const s7c_charts_tail =
    \\<chartFormats count="3"><chartFormat chart="0" format="0" series="1"><pivotArea type="data" outline="0" fieldPosition="0"><references count="1"><reference field="4294967294" count="1" selected="0"><x/></reference></references></pivotArea></chartFormat><chartFormat chart="0" format="2" series="1"><pivotArea type="data" outline="0" fieldPosition="0"><references count="1"><reference field="4294967294" count="1" selected="0"><x v="1"/></reference></references></pivotArea></chartFormat><chartFormat chart="0" format="3" series="1"><pivotArea type="data" outline="0" fieldPosition="0"><references count="1"><reference field="4294967294" count="1" selected="0"><x v="2"/></reference></references></pivotArea></chartFormat></chartFormats>
;
const s7c_multi_data_charts = s7c_multi_data[0 .. s7c_multi_data.len - "</pivotTableDefinition>".len] ++ s7c_charts_tail ++ "</pivotTableDefinition>";

test "S7c-2 edit: values-only chartFormats move with the data-field drop — removed, renumbered, emptied whole; non-canonical shapes refuse; K3 leaves them (in-house S7C2-B8)" {
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    // Remove ordinal 0: the block selecting data field 0 leaves whole,
    // the survivors renumber 1 -> 0 and 2 -> 1, the count follows.
    const two = try edit.applyConsumerSchemaEdit(arena, s7c_multi_data_charts, .{ .remove = 0 });
    try testing.expect(std.mem.indexOf(u8, two, "<chartFormats count=\"2\">") != null);
    try testing.expectEqual(@as(usize, 2), std.mem.count(u8, two, "<chartFormat "));
    try testing.expect(std.mem.indexOf(u8, two, "format=\"0\"") == null);
    try testing.expect(std.mem.indexOf(u8, two, "selected=\"0\"><x v=\"0\"/>") != null);
    try testing.expect(std.mem.indexOf(u8, two, "<x v=\"1\"/>") != null);
    try testing.expect(std.mem.indexOf(u8, two, "<x v=\"2\"/>") == null);
    // Remove ordinal 2 (two data fields at once): both their blocks
    // leave, the bare-`<x/>` block (index 0) stands as written.
    const one = try edit.applyConsumerSchemaEdit(arena, s7c_multi_data_charts, .{ .remove = 2 });
    try testing.expect(std.mem.indexOf(u8, one, "<chartFormats count=\"1\">") != null);
    try testing.expectEqual(@as(usize, 1), std.mem.count(u8, one, "<chartFormat "));
    try testing.expect(std.mem.indexOf(u8, one, "format=\"0\"") != null);
    try testing.expect(std.mem.indexOf(u8, one, "selected=\"0\"><x/>") != null);
    // Every block naming a dropped index empties the element whole.
    const only_dropped = try replacedOnce(arena, s7c_multi_data_charts, "<chartFormat chart=\"0\" format=\"0\" series=\"1\"><pivotArea type=\"data\" outline=\"0\" fieldPosition=\"0\"><references count=\"1\"><reference field=\"4294967294\" count=\"1\" selected=\"0\"><x/></reference></references></pivotArea></chartFormat>", "");
    const emptied_src = try replacedOnce(arena, only_dropped, "<chartFormats count=\"3\">", "<chartFormats count=\"2\">");
    const emptied = try edit.applyConsumerSchemaEdit(arena, emptied_src, .{ .remove = 2 });
    try testing.expect(std.mem.indexOf(u8, emptied, "<chartFormats") == null);
    // Non-canonical shapes refuse rather than heal: a second `<x>`,
    // an index past the data fields, a lying count.
    const two_x = try replacedOnce(arena, s7c_multi_data_charts, "<x v=\"1\"/></reference>", "<x v=\"1\"/><x v=\"1\"/></reference>");
    try testing.expectError(error.PivotShapeUnsupported, edit.applyConsumerSchemaEdit(arena, two_x, .{ .remove = 0 }));
    const dangling = try replacedOnce(arena, s7c_multi_data_charts, "<x v=\"2\"/></reference>", "<x v=\"9\"/></reference>");
    try testing.expectError(error.MalformedPivotXml, edit.applyConsumerSchemaEdit(arena, dangling, .{ .remove = 0 }));
    const lying = try replacedOnce(arena, s7c_multi_data_charts, "<chartFormats count=\"3\">", "<chartFormats count=\"4\">");
    try testing.expectError(error.MalformedPivotXml, edit.applyConsumerSchemaEdit(arena, lying, .{ .remove = 0 }));
    // A second `<chartFormats>` refuses — even an empty trailing one,
    // which must not fold the classification down while the first
    // element's blocks go stale (Codex #210 r2 REL-202).
    const trailing = try replacedOnce(arena, s7c_multi_data_charts, "</chartFormats>", "</chartFormats><chartFormats/>");
    try testing.expectError(error.PivotShapeUnsupported, edit.applyConsumerSchemaEdit(arena, trailing, .{ .remove = 0 }));
    try testing.expectError(error.PivotShapeUnsupported, edit.applyConsumerSchemaEdit(arena, trailing, .{ .remove = 2 }));
    // An extra `<references>` wrapper — even an empty one — is not the
    // one-block-one-index shape (Codex #210 r2 REL-203).
    const two_wrappers = try replacedOnce(arena, s7c_multi_data_charts, "fieldPosition=\"0\"><references count=\"1\"><reference field=\"4294967294\" count=\"1\" selected=\"0\"><x v=\"1\"/>", "fieldPosition=\"0\"><references count=\"0\"/><references count=\"1\"><reference field=\"4294967294\" count=\"1\" selected=\"0\"><x v=\"1\"/>");
    try testing.expectError(error.PivotShapeUnsupported, edit.applyConsumerSchemaEdit(arena, two_wrappers, .{ .remove = 0 }));
    // The K3 path reads none of it — the same shapes lift untouched.
    for ([_][]const u8{ s7c_multi_data_charts, two_x, dangling, lying, trailing, two_wrappers }) |src| {
        const k3 = try edit.applyConsumerSchemaEdit(arena, src, .{ .remove = 3 });
        try testing.expect(std.mem.indexOf(u8, k3, "<x v=\"2\"/>") != null or std.mem.indexOf(u8, k3, "<x v=\"9\"/>") != null);
    }
    // A wrapper count that lies about its children is evidence, kept:
    // `colFields` and `dataFields` set the same flag the layout
    // refuses on, before any splice could heal it (Codex #210 r2
    // MNT-201) — on EVERY schema edit, the S7c-1 rule.
    const cf_lying = try replacedOnce(arena, s7c_multi_data, "<colFields count=\"1\">", "<colFields count=\"2\">");
    try testing.expectError(error.MalformedPivotXml, edit.applyConsumerSchemaEdit(arena, cf_lying, .{ .remove = 0 }));
    try testing.expectError(error.MalformedPivotXml, edit.applyConsumerSchemaEdit(arena, cf_lying, .{ .remove = 3 }));
    const df_lying = try replacedOnce(arena, s7c_multi_data, "<dataFields count=\"3\">", "<dataFields count=\"4\">");
    try testing.expectError(error.MalformedPivotXml, edit.applyConsumerSchemaEdit(arena, df_lying, .{ .remove = 0 }));
    try testing.expectError(error.MalformedPivotXml, edit.applyConsumerSchemaEdit(arena, df_lying, .{ .remove = 3 }));
}

test "S7c-2 edit: the colItems canonical gates each refuse — and three survivors re-enumerate; a data-field-less consumer still lifts K3 (in-house S7C2-B1..B6)" {
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    // A typed `<i>`, a second `<x>`, an `i` disagreeing with its
    // position, an item list shorter than the data fields, a missing
    // values axis: each refuses the drop and leaves K3 alone.
    const typed = try replacedOnce(arena, s7c_multi_data, "<i i=\"2\"><x v=\"2\"/></i>", "<i i=\"2\" t=\"grand\"><x v=\"2\"/></i>");
    try testing.expectError(error.PivotShapeUnsupported, edit.applyConsumerSchemaEdit(arena, typed, .{ .remove = 0 }));
    _ = try edit.applyConsumerSchemaEdit(arena, typed, .{ .remove = 3 });
    const double_x = try replacedOnce(arena, s7c_multi_data, "<i i=\"1\"><x v=\"1\"/></i>", "<i i=\"1\"><x v=\"1\"/><x v=\"1\"/></i>");
    try testing.expectError(error.PivotShapeUnsupported, edit.applyConsumerSchemaEdit(arena, double_x, .{ .remove = 0 }));
    const misnumbered = try replacedOnce(arena, s7c_multi_data, "<i i=\"1\"><x v=\"1\"/></i>", "<i i=\"2\"><x v=\"1\"/></i>");
    try testing.expectError(error.PivotShapeUnsupported, edit.applyConsumerSchemaEdit(arena, misnumbered, .{ .remove = 0 }));
    const short = try replacedOnce(arena, s7c_multi_data, "<colItems count=\"3\"><i><x/></i><i i=\"1\"><x v=\"1\"/></i><i i=\"2\"><x v=\"2\"/></i></colItems>", "<colItems count=\"2\"><i><x/></i><i i=\"1\"><x v=\"1\"/></i></colItems>");
    try testing.expectError(error.PivotShapeUnsupported, edit.applyConsumerSchemaEdit(arena, short, .{ .remove = 0 }));
    _ = try edit.applyConsumerSchemaEdit(arena, short, .{ .remove = 3 });
    const axisless = try replacedOnce(arena, s7c_multi_data, "<colFields count=\"1\"><field x=\"-2\"/></colFields>", "");
    try testing.expectError(error.PivotShapeUnsupported, edit.applyConsumerSchemaEdit(arena, axisless, .{ .remove = 0 }));
    _ = try edit.applyConsumerSchemaEdit(arena, axisless, .{ .remove = 3 });
    // Four data fields: three survive, the re-enumeration is not a
    // hardcoded pair.
    var four = try replacedOnce(arena, s7c_multi_data, "<pivotField showAll=\"0\"/></pivotFields>", "<pivotField dataField=\"1\" showAll=\"0\"/></pivotFields>");
    four = try replacedOnce(arena, four, "<dataFields count=\"3\">", "<dataFields count=\"4\">");
    four = try replacedOnce(arena, four, "</dataFields>", "<dataField name=\"Sum of D\" fld=\"3\" baseField=\"1\" baseItem=\"0\"/></dataFields>");
    four = try replacedOnce(arena, four, "<i i=\"2\"><x v=\"2\"/></i></colItems>", "<i i=\"2\"><x v=\"2\"/></i><i i=\"3\"><x v=\"3\"/></i></colItems>");
    four = try replacedOnce(arena, four, "<colItems count=\"3\">", "<colItems count=\"4\">");
    four = try replacedOnce(arena, four, "ref=\"B2:E7\"", "ref=\"B2:F7\"");
    const three = try edit.applyConsumerSchemaEdit(arena, four, .{ .remove = 0 });
    try testing.expect(std.mem.indexOf(u8, three, "<colItems count=\"3\"><i><x/></i><i i=\"1\"><x v=\"1\"/></i><i i=\"2\"><x v=\"2\"/></i></colItems>") != null);
    try testing.expect(std.mem.indexOf(u8, three, "<dataFields count=\"3\">") != null);
    try testing.expect(std.mem.indexOf(u8, three, "<location ref=\"B2:E7\"") != null);
    // A consumer with no data fields at all is not K4b — its
    // unreferenced columns keep lifting (the `dropped > 0` clause).
    var bare = try replacedOnce(arena, s7c_row_axis_high, "<pivotField dataField=\"1\" showAll=\"0\"/>", "<pivotField showAll=\"0\"/>");
    bare = try replacedOnce(arena, bare, "<dataFields count=\"1\"><dataField name=\"Sum of Qty\" fld=\"1\" baseField=\"2\" baseItem=\"0\"/></dataFields>", "");
    _ = try edit.applyConsumerSchemaEdit(arena, bare, .{ .remove = 0 });
}

test "S7c: an attachment list the reader cannot see refuses; the plain list decodes (in-house S7C-R3b)" {
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    const x15 =
        \\<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="S" sourceName="F"><extLst><ext uri="{Y}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:slicerCachePivotTables><x15:pivotTable tabId="2" name="P"/></x15:slicerCachePivotTables></ext></extLst></slicerCacheDefinition>
    ;
    try testing.expectError(error.MalformedPivotXml, attachedPivotNames(arena, x15, .slicer));
    const plain =
        \\<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="S" sourceName="F"><pivotTables><pivotTable tabId="2" name="P&amp;L"/></pivotTables><data><tabular pivotCacheId="2"/></data></slicerCacheDefinition>
    ;
    const names = try attachedPivotNames(arena, plain, .slicer);
    try testing.expectEqual(@as(usize, 1), names.len);
    try testing.expectEqualStrings("P&L", names[0]);
    // A descendant alias of the part's own family would hide the list
    // under a prefix the walk steps over — the family-parameterized
    // preflight refuses it (Codex #208 r1 SEC-102).
    const aliased =
        \\<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="S" sourceName="F"><data xmlns:y="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><y:pivotTables><y:pivotTable tabId="2" name="P"/></y:pivotTables></data></slicerCacheDefinition>
    ;
    try testing.expectError(error.MalformedPivotXml, attachedPivotNames(arena, aliased, .slicer));
    // An entity-spelled alias binds the same namespace as its plain
    // spelling (Codex #208 r2 SEC-201): decoded before the family
    // comparison, it refuses the same way.
    const ent_alias =
        \\<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="S" sourceName="F"><data xmlns:y="http://schemas.microsoft.com/office/spreadsheetml/2009/9/mai&#x6e;"><y:pivotTables><y:pivotTable tabId="2" name="P"/></y:pivotTables></data></slicerCacheDefinition>
    ;
    try testing.expectError(error.MalformedPivotXml, attachedPivotNames(arena, ent_alias, .slicer));
    // A root NOT bound to the family is not a slicer cache this reader
    // can walk — undeclared or vendor-namespaced roots refuse rather
    // than skip a foreign-prefixed attachment list (Codex #208 r5
    // REL-502).
    const unbound =
        \\<slicerCacheDefinition name="S" sourceName="F"><pivotTables><pivotTable tabId="2" name="P"/></pivotTables></slicerCacheDefinition>
    ;
    try testing.expectError(error.MalformedPivotXml, attachedPivotNames(arena, unbound, .slicer));
    const vendor =
        \\<v:slicerCacheDefinition xmlns:v="urn:vendor" name="S"><v:pivotTables><v:pivotTable tabId="2" name="P"/></v:pivotTables></v:slicerCacheDefinition>
    ;
    try testing.expectError(error.MalformedPivotXml, attachedPivotNames(arena, vendor, .slicer));
}

test "S7c engine: cache-side tolerated-unread content refuses a schema edit — a surviving field's extLst, an ordinal token in the root extension (Codex #208 r2 REL-203)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "s7c_cache_ext.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .sheet_ref);
    // The corpus shape: an attribute-only root extension with no
    // ordinal name admits the rebuild.
    try fixture.patchPart(testing.allocator, io, path, "xl/pivotCache/pivotCacheDefinition1.xml", "</pivotCacheDefinition>", "<extLst><ext uri=\"{X}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:pivotCacheDefinition pivotCacheId=\"2\"/></ext></extLst></pivotCacheDefinition>");
    var arena_state = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena_state.deinit();
    const arena = arena_state.allocator();
    const rows = [_]engine.Row{
        &.{ .{ .string = "East" }, .{ .number = "3" } },
        &.{ .{ .string = "West" }, .{ .number = "4" } },
        &.{ .{ .string = "East" }, .{ .number = "5" } },
    };
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        const cache = &o.pivots.caches[0];
        const rec = (try o.store.part(cache.records_part_name.?)).?.bytes;
        _ = try engine.rebuildWith(arena, cache, &rows, rec, null, .{ .remove = 2 });
    }
    // An ordinal-carrier attribute anywhere in the extension region
    // refuses the schema edit; the same cache still rebuilds under a
    // ROW edit (no ordinal moves).
    try fixture.patchPart(testing.allocator, io, path, "xl/pivotCache/pivotCacheDefinition1.xml", "pivotCacheId=\"2\"", "pivotCacheId=\"2\" fld=\"1\"");
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        const cache = &o.pivots.caches[0];
        const rec = (try o.store.part(cache.records_part_name.?)).?.bytes;
        try testing.expectError(error.PivotShapeUnsupported, engine.rebuildWith(arena, cache, &rows, rec, null, .{ .remove = 2 }));
        _ = try engine.rebuild(arena, cache, &fixture_rows, rec, null);
    }
    // A SURVIVING field's own extLst refuses; the removed field's
    // leaves whole with it.
    const path2 = try tt.path(testing.allocator, io, "s7c_field_ext.xlsx");
    defer testing.allocator.free(path2);
    try fixture.write(testing.allocator, io, path2, .sheet_ref);
    try fixture.patchPart(testing.allocator, io, path2, "xl/pivotCache/pivotCacheDefinition1.xml", "maxValue=\"5\"/></cacheField>", "maxValue=\"5\"/><extLst/></cacheField>");
    var o = try Opened.open(testing.allocator, io, path2);
    defer o.deinit(testing.allocator);
    const cache = &o.pivots.caches[0];
    const rec = (try o.store.part(cache.records_part_name.?)).?.bytes;
    try testing.expectError(error.PivotShapeUnsupported, engine.rebuildWith(arena, cache, &rows, rec, null, .{ .remove = 2 }));
    const keep = [_]engine.Row{
        &.{ .{ .string = "East" }, .{ .number = "1.5" } },
        &.{ .{ .string = "West" }, .{ .number = "2.5" } },
        &.{ .{ .string = "East" }, .{ .number = "3.5" } },
    };
    _ = try engine.rebuildWith(arena, cache, &keep, rec, null, .{ .remove = 1 });
    // A caller-assembled extension offset answers, it does not trap
    // (Codex #208 r5 REL-503).
    var forged = cache.*;
    forged.definition.ext_lst_start = forged.raw_xml.len + 1;
    try testing.expectError(error.MalformedPivotXml, engine.rebuildWith(arena, &forged, &rows, rec, null, .{ .remove = 2 }));
}

fn rebuildWithForFailures(allocator: Allocator, cache: *const PivotCache, rows: []const engine.Row, rec: []const u8, schema: edit.SchemaEdit) !void {
    var arena_state = std.heap.ArenaAllocator.init(allocator);
    defer arena_state.deinit();
    _ = try engine.rebuildWith(arena_state.allocator(), cache, rows, rec, null, schema);
}

fn consumerSchemaForFailures(allocator: Allocator, table_xml: []const u8, schema: edit.SchemaEdit) !void {
    var arena_state = std.heap.ArenaAllocator.init(allocator);
    defer arena_state.deinit();
    _ = try edit.applyConsumerSchemaEdit(arena_state.allocator(), table_xml, schema);
}

test "S7c: an allocation failure anywhere in a schema rebuild or the consumer rewrite is OutOfMemory, never a refusal (in-house S7C-MUT-4)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "s7c_oom.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .sheet_ref);
    var o = try Opened.open(testing.allocator, io, path);
    defer o.deinit(testing.allocator);
    const cache = &o.pivots.caches[0];
    const rec = (try o.store.part(cache.records_part_name.?)).?.bytes;
    const rows = [_]engine.Row{
        &.{ .{ .string = "East" }, .{ .number = "3" } },
        &.{ .{ .string = "West" }, .{ .number = "4" } },
        &.{ .{ .string = "East" }, .{ .number = "5" } },
    };
    try testing.checkAllAllocationFailures(testing.allocator, rebuildWithForFailures, .{ cache, &rows, rec, .{ .remove = 2 } });
    try testing.checkAllAllocationFailures(testing.allocator, consumerSchemaForFailures, .{ o.pivots.tables[0].raw_xml, .{ .remove = 2 } });
    // The K2 insert allocates down its own paths — the fresh cache
    // field, the blank inventory, the bare pivotField, the inserted
    // spines (Codex #208 r3 MNT-302).
    const grown = [_]engine.Row{
        &.{ .{ .string = "East" }, .blank, .{ .number = "3" }, .{ .number = "1.5" } },
        &.{ .{ .string = "West" }, .blank, .{ .number = "4" }, .{ .number = "2.5" } },
        &.{ .{ .string = "East" }, .blank, .{ .number = "5" }, .{ .number = "3.5" } },
    };
    try testing.checkAllAllocationFailures(testing.allocator, rebuildWithForFailures, .{ cache, &grown, rec, .{ .insert = .{ .at = 1, .name = "Column4" } } });
    try testing.checkAllAllocationFailures(testing.allocator, consumerSchemaForFailures, .{ o.pivots.tables[0].raw_xml, .{ .insert = .{ .at = 1, .name = "Column4" } } });
    // The K4a drop allocates down its own paths — the re-enumerated
    // values axis, the narrowed location, the collapse, the
    // chart-format move (S7c-2).
    try testing.checkAllAllocationFailures(testing.allocator, consumerSchemaForFailures, .{ s7c_multi_data, .{ .remove = 0 } });
    try testing.checkAllAllocationFailures(testing.allocator, consumerSchemaForFailures, .{ s7c_multi_data, .{ .remove = 2 } });
    try testing.checkAllAllocationFailures(testing.allocator, consumerSchemaForFailures, .{ s7c_multi_data_charts, .{ .remove = 0 } });
}
