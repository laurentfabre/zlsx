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
const engine = @import("zlsx_formula");
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
    /// at `format_buf_len`. Every column here is in-grid (the parsers
    /// below refuse anything past `XFD`), so the letter writer cannot
    /// fail.
    pub fn formatA1(self: Bounds, buf: *[format_buf_len]u8) []const u8 {
        return switch (self) {
            .rect => |r| blk: {
                var tl: [coords.max_col_letters]u8 = undefined;
                var br: [coords.max_col_letters]u8 = undefined;
                const tl_len = coords.writeColNumberLetters(&tl, r.tl_col) catch unreachable;
                const br_len = coords.writeColNumberLetters(&br, r.br_col) catch unreachable;
                break :blk std.fmt.bufPrint(buf, "{s}{d}:{s}{d}", .{ tl[0..tl_len], r.tl_row, br[0..br_len], r.br_row }) catch unreachable;
            },
            .whole_columns => |c| blk: {
                var f: [coords.max_col_letters]u8 = undefined;
                var l: [coords.max_col_letters]u8 = undefined;
                const f_len = coords.writeColNumberLetters(&f, c.first_col) catch unreachable;
                const l_len = coords.writeColNumberLetters(&l, c.last_col) catch unreachable;
                break :blk std.fmt.bufPrint(buf, "{s}:{s}", .{ f[0..f_len], l[0..l_len] }) catch unreachable;
            },
            .whole_rows => |r| std.fmt.bufPrint(buf, "{d}:{d}", .{ r.first_row, r.last_row }) catch unreachable,
        };
    }

    pub const format_buf_len = 32;
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
    };
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
        return switch (r) {
            .sheet => |s| s.sheet_idx == sheet_idx,
            .unresolved => |u| std.mem.indexOfScalar(u32, u.sheets, sheet_idx) != null,
            .external, .none => false,
        };
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
            .external, .scenario, .unknown => {},
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

/// Trailing-segment match, case-insensitive: the one thing the Strict
/// and Transitional relationship URIs share. Same rule as
/// `Workbook.preflightPivotEditsForSheet`.
pub fn relLeafIs(rel_type: []const u8, leaf: []const u8) bool {
    const l = if (std.mem.lastIndexOfScalar(u8, rel_type, '/')) |i| rel_type[i + 1 ..] else rel_type;
    return std.ascii.eqlIgnoreCase(l, leaf);
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
    return engine.decode.decodeCarrier(a, .lexical, raw) catch |e| switch (e) {
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
fn decode(a: Allocator, site: engine.decode.Site, raw: []const u8) Error![]const u8 {
    return engine.decode.decodeAt(a, site, raw) catch |e| switch (e) {
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
    symbols: ?engine.SymbolTable = null,
    symbols_refused: bool = false,

    const TableEntry = struct {
        folded: []const u8,
        sheet_idx: u32,
        /// The table part's `ref`, when it parses as a rectangle.
        bounds: ?Bounds,
    };

    /// Names walked per closure before the walk stops — deeper nests are
    /// not workbooks Excel wrote, and a bound keeps a hostile chain
    /// finite even without the visited set.
    const max_closure_depth: u32 = 64;

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
            return unresolved(.unplaceable_rid, try self.sheetsOfAttr(ws.sheet));
        }
        if (ws.sheet) |raw_sheet| {
            const name = try decode(self.arena, .pivot_source_sheet_name, raw_sheet);
            const idx = (try self.sheetIndexOf(name)) orelse return unresolved(.dangling_sheet, &.{});
            return try self.local(idx, .sheet_attr, try self.boundsOfRef(ws.ref));
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
                    if (try self.areaOfBody(n.body)) |area| return try self.local(area.sheet_idx, .defined_name, area.bounds);
                    return unresolved(.unbounded_body, try self.closureSheets(n.body, n.folded));
                },
                .refused => return error.MalformedPivotXml,
                .table, .not_found => {},
            }
            const folded = (try fold(self.arena, name)) orelse return unresolved(.dangling_name, &.{});
            for (try self.ensureTables()) |t| {
                if (std.mem.eql(u8, t.folded, folded)) return try self.local(t.sheet_idx, .table, t.bounds);
            }
            return unresolved(.dangling_name, &.{});
        }
        return unresolved(if (ws.ref != null) .sheetless_ref else .no_locator, &.{});
    }

    fn unresolved(why: Unresolved.Why, sheets: []const u32) SourceResolution {
        return .{ .unresolved = .{ .why = why, .sheets = sheets } };
    }

    fn local(self: *Resolver, sheet_idx: u32, via: ResolvedVia, bounds: ?Bounds) Error!SourceResolution {
        if (sheet_idx >= self.sheet_parts.len) return unresolved(.dangling_sheet, &.{});
        return .{ .sheet = .{
            .sheet_idx = sheet_idx,
            .sheet_name = self.sheet_names[sheet_idx],
            .part_name = self.sheet_parts[sheet_idx],
            .via = via,
            .bounds = bounds,
        } };
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
                try entries.append(self.arena, .{ .folded = folded, .sheet_idx = @intCast(i), .bounds = bounds });
            }
        }
        self.tables = try entries.toOwnedSlice(self.arena);
        return self.tables.?;
    }

    /// The engine's symbol table, or `MalformedPivotXml` when it
    /// refuses: with the name inventory unreadable, no name-based
    /// source can be resolved either way.
    fn ensureSymbols(self: *Resolver) Error!*const engine.SymbolTable {
        if (self.symbols) |*t| return t;
        if (self.symbols_refused) return error.MalformedPivotXml;

        var builder = engine.Builder.init(self.gpa, recalc_run.collation_v1);
        defer builder.deinit();
        for (self.wb.sheets) |s| try builder.addSheet(s.name);
        // The names come from the part, as the evaluator's do: the
        // typed view drops the attribute region that says whether a
        // name is a macro entry point rather than a range.
        switch (try engine.names.scanDefinedNames(self.gpa, self.wb_xml)) {
            .ok => |d| {
                var defined = d;
                defer defined.deinit();
                for (defined.rows) |dn| {
                    try builder.addName(dn.raw_identifier, dn.raw_body, .{
                        .scope = if (dn.local_sheet_id) |id| engine.env.SheetIndex.fromInt(id) else null,
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

    const Area = struct { sheet_idx: u32, bounds: Bounds };

    /// The area a defined name's body denotes, when the body is exactly
    /// one static sheet-qualified area (`Data!$A$1:$C$4`, `Data!$A$1`,
    /// `'My Data'!A1:B2`, `Data!$A:$C`). A 3D span, a dynamic body
    /// (`OFFSET(…)`, `Data!A1:INDEX(…)`), a union, a bare range, or a
    /// range whose ends are not the same kind of reference resolves to
    /// null: none names one area a pivot could read.
    fn areaOfBody(self: *Resolver, body: []const u8) Error!?Area {
        var parsed = try engine.parser.parse(self.gpa, body, .{});
        defer parsed.deinit(self.gpa);
        const ast = switch (parsed) {
            .ok => |t| t,
            .refused => return null,
        };
        // `Data!A1` is one qualified node; `Data!$A$1:$C$4` is a range
        // operator whose LEFT operand carries the qualifier. The right
        // operand must be a static reference too, on the same sheet if
        // it is qualified at all — compared by resolved index, so
        // `Data!A1:data!C4` is one sheet.
        switch (ast.node(ast.root)) {
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
                const rhs_target: engine.parser.Node = switch (rhs_node) {
                    .qualified => |q| blk: {
                        const r = (try self.sheetOfSpec(q.sheet)) orelse return null;
                        if (r != idx) return null;
                        break :blk ast.node(q.target);
                    },
                    else => rhs_node,
                };
                const bounds = rangeBounds(ast.node(lhs.target), rhs_target) orelse return null;
                return .{ .sheet_idx = idx, .bounds = bounds };
            },
            else => return null,
        }
    }

    /// Every sheet a name body depends on, through the names it
    /// references: sheet qualifiers (a 3D span contributes each sheet
    /// between its ends, in tab order), tables (their host sheet), and
    /// the bodies of referenced names, recursively. A name is walked
    /// once, so a cycle terminates; a body the parser refuses, a name
    /// the inventory refuses or lacks, and a nest deeper than
    /// `max_closure_depth` contribute nothing — evidence, not refusal.
    /// Ascending, deduplicated, arena-owned.
    fn closureSheets(self: *Resolver, body: []const u8, folded_self: []const u8) Error![]const u32 {
        var sheets: std.ArrayListUnmanaged(u32) = .empty;
        defer sheets.deinit(self.gpa);
        var visited: std.ArrayListUnmanaged([]const u8) = .empty;
        defer visited.deinit(self.gpa);
        try visited.append(self.gpa, folded_self);
        try self.walkBody(body, &sheets, &visited, 0);
        std.mem.sort(u32, sheets.items, {}, std.sort.asc(u32));
        return try self.arena.dupe(u32, sheets.items);
    }

    fn walkBody(
        self: *Resolver,
        body: []const u8,
        sheets: *std.ArrayListUnmanaged(u32),
        visited: *std.ArrayListUnmanaged([]const u8),
        depth: u32,
    ) Error!void {
        if (depth > max_closure_depth) return;
        var parsed = try engine.parser.parse(self.gpa, body, .{});
        defer parsed.deinit(self.gpa);
        const ast = switch (parsed) {
            .ok => |t| t,
            .refused => return,
        };
        try self.walkNode(ast, ast.root, sheets, visited, depth);
    }

    fn walkNode(
        self: *Resolver,
        ast: engine.parser.Ast,
        i: engine.parser.Index,
        sheets: *std.ArrayListUnmanaged(u32),
        visited: *std.ArrayListUnmanaged([]const u8),
        depth: u32,
    ) Error!void {
        switch (ast.node(i)) {
            .qualified => |q| {
                try self.addSpecSheets(q.sheet, sheets);
                try self.walkNode(ast, q.target, sheets, visited, depth);
            },
            .name => |n| try self.walkName(n.bare, sheets, visited, depth),
            .structured => |st| if (st.table) |t| try self.addTableSheet(t, sheets),
            // The callee is a function, not a name.
            .call => |c| for (ast.children(c.args)) |k| try self.walkNode(ast, k, sheets, visited, depth),
            .array => |a| for (ast.children(a.elems)) |k| try self.walkNode(ast, k, sheets, visited, depth),
            .paren => |pn| try self.walkNode(ast, pn.child, sheets, visited, depth),
            .unary => |u| try self.walkNode(ast, u.child, sheets, visited, depth),
            .postfix => |pf| try self.walkNode(ast, pf.child, sheets, visited, depth),
            .binary => |b| {
                try self.walkNode(ast, b.lhs, sheets, visited, depth);
                try self.walkNode(ast, b.rhs, sheets, visited, depth);
            },
            .number, .string, .boolean, .error_lit, .missing_arg, .ref_cell, .ref_full_col, .ref_full_row => {},
        }
    }

    fn walkName(
        self: *Resolver,
        bare: []const u8,
        sheets: *std.ArrayListUnmanaged(u32),
        visited: *std.ArrayListUnmanaged([]const u8),
        depth: u32,
    ) Error!void {
        const folded = (try fold(self.arena, bare)) orelse return;
        for (visited.items) |v| if (std.mem.eql(u8, v, folded)) return;
        try visited.append(self.gpa, folded);
        const symbols = self.ensureSymbols() catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            else => return,
        };
        switch (try symbols.resolveName(self.gpa, null, bare)) {
            .name => |n| try self.walkBody(n.body, sheets, visited, depth + 1),
            .table, .not_found, .refused => {
                // A table spelled bare, or nothing: the table index has
                // the host if there is one.
                try self.addTableSheet(bare, sheets);
            },
        }
    }

    fn addSpecSheets(self: *Resolver, spec: engine.parser.SheetSpec, sheets: *std.ArrayListUnmanaged(u32)) Error!void {
        var buf: [256]u8 = undefined;
        const first_name = unquoteSheetSpec(&buf, spec) orelse return;
        const first = (try self.sheetIndexOf(first_name)) orelse return;
        var last = first;
        if (spec.last) |raw_last| {
            var last_buf: [256]u8 = undefined;
            const last_name = unquoteSheetSpec(&last_buf, .{ .first = raw_last, .quoted = spec.quoted }) orelse return;
            last = (try self.sheetIndexOf(last_name)) orelse return;
        }
        var k = @min(first, last);
        while (k <= @max(first, last)) : (k += 1) try addSheet(self.gpa, sheets, k);
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

    fn sheetOfSpec(self: *Resolver, spec: engine.parser.SheetSpec) Error!?u32 {
        if (spec.last != null) return null;
        var buf: [256]u8 = undefined;
        const name = unquoteSheetSpec(&buf, spec) orelse return null;
        return self.sheetIndexOf(name);
    }
};

/// A cell, a whole-column span or a whole-row span — the reference
/// nodes that denote a fixed area — as bounds.
fn staticBounds(n: engine.parser.Node) ?Bounds {
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
fn rangeBounds(lhs: engine.parser.Node, rhs: engine.parser.Node) ?Bounds {
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

/// A decoded `ref` as bounds: `A1:C4` / `A1` (the S7a rectangle parser),
/// `A:C` (letters on both sides), `1:4` (digits on both sides). Null on
/// anything else — bounds are evidence, and an unparseable spelling is
/// none.
pub fn parseBounds(ref: []const u8) ?Bounds {
    if (edit.parseRect(ref)) |r| return .{ .rect = r };
    const colon = std.mem.indexOfScalar(u8, ref, ':') orelse return null;
    const lhs = ref[0..colon];
    const rhs = ref[colon + 1 ..];
    if (lhs.len == 0 or rhs.len == 0) return null;
    if (allLetters(lhs) and allLetters(rhs)) {
        // Uppercase only and inside the grid — the spelling `ST_Ref` and
        // the S7a rectangle parser accept.
        const a = coords.parseColNumber(lhs, .{ .case = .upper_only }) catch return null;
        const b = coords.parseColNumber(rhs, .{ .case = .upper_only }) catch return null;
        return .{ .whole_columns = .{ .first_col = @min(a, b), .last_col = @max(a, b) } };
    }
    if (allDigits(lhs) and allDigits(rhs)) {
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
fn unquoteSheetSpec(buf: []u8, spec: engine.parser.SheetSpec) ?[]const u8 {
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
        /// The shift would push the rectangle past `XFD` / `1048576`.
        PivotCoordinateOverflow,
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

        var ref_buf: [48]u8 = undefined;
        const new_ref = formatRect(&ref_buf, shifted) catch return error.PivotCoordinateOverflow;
        const span = def.location.ref_span;
        assert(span.start <= span.end and span.end <= src.len);
        return std.mem.concat(allocator, u8, &.{ src[0..span.start], new_ref, src[span.end..] });
    }

    /// The footprint of a parsed definition. Refuses (as malformed) a
    /// `ref` that is not an A1 rectangle — ST_Ref has no `$`, no sheet
    /// qualifier, no whitespace.
    pub fn footprintOf(allocator: Allocator, def: pivot_xml.TableDefinition) EditError!Footprint {
        const decoded = engine.decode.decodeCarrier(allocator, .lexical, def.location.ref) catch |e| switch (e) {
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

    fn formatRect(buf: *[48]u8, r: Rect) ![]const u8 {
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
    try testing.expectEqualStrings(want, (b orelse return error.TestExpectedBounds).formatA1(&buf));
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

test "parseBounds: rectangles, whole columns, whole rows, and nothing else" {
    try expectA1(parseBounds("A1:C4"), "A1:C4");
    try expectA1(parseBounds("B7"), "B7:B7");
    try expectA1(parseBounds("A:C"), "A:C");
    try expectA1(parseBounds("C:A"), "A:C");
    try expectA1(parseBounds("XFD:XFD"), "XFD:XFD");
    try expectA1(parseBounds("3:9"), "3:9");
    try expectA1(parseBounds("9:3"), "3:9");
    try expectA1(parseBounds("1:1048576"), "1:1048576");
    for ([_][]const u8{ "", ":", "A:", ":C", "A1:C", "a:c", "XFE:XFE", "0:4", "1:1048577", "A:C:D", "$A:$C", "A1:C4 " }) |bad| {
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
    // A 3D span proves every sheet between its ends; a union both sides.
    try fixture.write(testing.allocator, io, path, .defined_name);
    try patchPart(io, path, "xl/workbook.xml", "Data!$A$1:$C$4", "SUM(Data:Report!$A$1)");
    {
        var o = try Opened.open(testing.allocator, io, path);
        defer o.deinit(testing.allocator);
        try expectUnresolved(o.pivots.caches[0].resolution, .unbounded_body, &.{ 0, 1 });
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
