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

pub const SourceResolution = union(enum) {
    /// A sheet of this workbook.
    sheet: LocalSheet,
    /// Another workbook: the relationship target, as written.
    external: []const u8,
    /// A worksheet-type source whose spelling names nothing this
    /// workbook has — a dangling sheet name, a name with a dynamic or
    /// 3D body, a missing relationship. Excel would fail the refresh.
    unresolved,
    /// Not a worksheet-type source (external connection, scenario,
    /// unknown), so there is no sheet to resolve.
    none,

    pub const LocalSheet = struct {
        sheet_idx: u32,
        /// Decoded sheet name.
        sheet_name: []const u8,
        part_name: []const u8,
        via: ResolvedVia,
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
    /// Decoded captions.
    data_caption: ?[]const u8,
    grand_total_caption: ?[]const u8,
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

    fn resolvesTo(r: SourceResolution, sheet_idx: u32) bool {
        return switch (r) {
            .sheet => |s| s.sheet_idx == sheet_idx,
            else => false,
        };
    }
};

/// Walk the graph. `wb` is the parsed `xl/workbook.xml` the caller
/// already holds (a `Workbook` has one); the walk reads the part bytes
/// again only for `<pivotCaches>`, which the typed view does not carry.
///
/// Leaf strings in the result either borrow from the store's arena or
/// live in the result's own arena; both outlive the `Pivots` for as
/// long as the store does.
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

    var caches: std.ArrayListUnmanaged(CacheSlot) = .empty;
    var tables: std.ArrayListUnmanaged(PivotTable) = .empty;

    // Root 1: the workbook's cache list. Its order is the cache order
    // Excel shows, and `cacheId` is the only place it lives.
    try collectWorkbookCaches(a, store, wb_part.bytes, wb_rels, &caches);

    // Root 2: every sheet's relationships. A pivot table part not linked
    // from a sheet is not a pivot Excel renders; it is not listed.
    for (wb.sheets, 0..) |s, i| {
        const sheet_idx: u32 = @intCast(i);
        const sheet_part = (try relTarget(store, "xl/workbook.xml", wb_rels, s.r_id)) orelse continue;
        for (store.rels(sheet_part)) |rel| {
            if (rel.target_mode == .external) continue;
            if (!relLeafIs(rel.type, "pivotTable")) continue;
            // Once the relationship type says "pivot table", the edge is
            // part of the graph: a target that does not resolve or a
            // part that is not there is a broken workbook, not a pivot
            // to leave out of the inventory.
            const pt_name = try requiredTarget(store, sheet_part, rel.target);
            const pt_part = try requiredPart(store, pt_name);
            const def = pivot_xml.parseTableDefinition(a, pt_part.bytes) catch |e| return mapParse(e);

            // The cache: the part's own relationship first (that is the
            // edge Excel follows), then `cacheId` against the workbook
            // list when the relationship is missing. Both must agree
            // when both are present — a pivot whose relationship names
            // one cache and whose `cacheId` names another is not one
            // Excel could refresh.
            var cache_idx: ?usize = null;
            for (store.rels(pt_name)) |prel| {
                if (prel.target_mode == .external) continue;
                if (!relLeafIs(prel.type, "pivotCacheDefinition")) continue;
                const cd_name = try requiredTarget(store, pt_name, prel.target);
                _ = try requiredPart(store, cd_name);
                cache_idx = try findOrAddCache(a, &caches, cd_name, null);
                break;
            }
            if (cache_idx) |ci| {
                if (caches.items[ci].cache_id) |id| {
                    if (id != def.cache_id) return error.MalformedPivotXml;
                }
            } else {
                for (caches.items, 0..) |c, ci| {
                    if (c.cache_id != null and c.cache_id.? == def.cache_id) {
                        cache_idx = ci;
                        break;
                    }
                }
            }

            const data_field_names = try a.alloc(?[]const u8, def.data_fields.len);
            for (def.data_fields, 0..) |df, k| {
                data_field_names[k] = if (df.name) |n| try decode(a, .pivot_field_name, n) else null;
            }

            try tables.append(a, .{
                .name = try decode(a, .pivot_table_name, def.name),
                .part_name = pt_name,
                .sheet_idx = sheet_idx,
                .sheet_name = sheet_names[i],
                .sheet_part_name = sheet_part,
                .definition = def,
                .data_caption = if (def.data_caption) |c| try decode(a, .pivot_table_name, c) else null,
                .grand_total_caption = if (def.grand_total_caption) |c| try decode(a, .pivot_table_name, c) else null,
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
        .wb_rels = wb_rels,
        .sheet_names = sheet_names,
    };
    defer resolver.deinit();

    const finished = try a.alloc(PivotCache, caches.items.len);
    for (caches.items, 0..) |slot, ci| {
        const part = try requiredPart(store, slot.part_name);
        const def = pivot_xml.parseCacheDefinition(a, part.bytes) catch |e| return mapParse(e);

        // The records part: the definition's `r:id` names it; a
        // definition without one (`saveData="0"`) has none, and a
        // relationship of the records type is taken in its place. Named
        // and absent is a refusal — the count says there are records,
        // and there is nothing to hold them.
        const cache_rels = store.rels(slot.part_name);
        var records: ?[]const u8 = null;
        if (def.r_id) |rid| {
            records = (try relTarget(store, slot.part_name, cache_rels, rid)) orelse
                return error.MalformedPivotXml;
        } else {
            for (cache_rels) |rel| {
                if (rel.target_mode == .external) continue;
                if (!relLeafIs(rel.type, "pivotCacheRecords")) continue;
                records = try requiredTarget(store, slot.part_name, rel.target);
                break;
            }
        }
        if (records) |r| _ = try requiredPart(store, r);

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
                    resolution = .unresolved;
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
    const container = (pivot_xml.findTag(wb_xml[0..root.body_end], root.hit.after_tag_close, root.prefix, "pivotCaches") catch |e| return mapParse(e)) orelse
        return;
    const end = pivot_xml.endOf(wb_xml, container, root.prefix, "pivotCaches") catch |e| return mapParse(e);
    if (container.self_closing) return;
    var cursor = container.after_tag_close;
    while (pivot_xml.findTag(wb_xml[0..end], cursor, root.prefix, "pivotCache") catch |e| return mapParse(e)) |hit| {
        cursor = pivot_xml.endOf(wb_xml, hit, root.prefix, "pivotCache") catch |e| return mapParse(e);
        const attrs = wb_xml[hit.attrs_start..hit.attrs_end];
        const rid = pivot_xml.nsAttr(attrs, root.rel_prefix, "id") orelse return error.MalformedPivotXml;
        const cache_id = (pivot_xml.u32Attr(attrs, "cacheId") catch |e| return mapParse(e)) orelse
            return error.MalformedPivotXml;
        const part_name = (try relTarget(store, "xl/workbook.xml", wb_rels, rid)) orelse
            return error.MalformedPivotXml;
        _ = try requiredPart(store, part_name);
        _ = try findOrAddCache(a, caches, part_name, cache_id);
    }
}

/// A relationship target that must resolve to a part name.
fn requiredTarget(store: *PartStore, owner: []const u8, target: []const u8) Error![]const u8 {
    return (try store.resolve(owner, target)) orelse error.MalformedPivotXml;
}

/// A part that must exist. Materialises it.
fn requiredPart(store: *PartStore, name: []const u8) Error!store_mod.Part {
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

/// Resolve a relationship id against an owner's relationships to a
/// part name. External targets and dangling ids resolve to null.
fn relTarget(
    store: *PartStore,
    owner: []const u8,
    rels: []const store_mod.Relationship,
    rid: []const u8,
) Error!?[]const u8 {
    if (rid.len == 0) return null;
    for (rels) |rel| {
        if (!std.mem.eql(u8, rel.id, rid)) continue;
        if (rel.target_mode == .external) return null;
        return try store.resolve(owner, rel.target);
    }
    return null;
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
        .ref = ws.ref,
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
/// Built lazily: a workbook whose caches are all external never pays
/// for a symbol table.
const Resolver = struct {
    gpa: Allocator,
    arena: Allocator,
    store: *PartStore,
    wb: *const wbxml.WorkbookXml,
    wb_xml: []const u8,
    wb_rels: []const store_mod.Relationship,
    sheet_names: []const []const u8,

    symbols: ?engine.SymbolTable = null,
    /// Null until built; empty when the build refused (duplicate sheet
    /// names, …), in which case nothing local can resolve.
    built: bool = false,
    tables: []const TableEntry = &.{},

    const TableEntry = struct {
        folded: []const u8,
        sheet_idx: u32,
        part_name: []const u8,
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
        // relationship says so. An internal target under `r:id` is not
        // a workbook this reader can name.
        if (ws.r_id) |rid| {
            for (cache_rels) |rel| {
                if (!std.mem.eql(u8, rel.id, rid)) continue;
                if (rel.target_mode != .external) return .unresolved;
                return .{ .external = rel.target };
            }
            return .unresolved;
        }
        try self.ensureBuilt();
        const symbols = &(self.symbols orelse return .unresolved);

        if (ws.sheet) |raw_sheet| {
            const name = try decode(self.arena, .pivot_source_sheet_name, raw_sheet);
            const idx = (try symbols.resolveSheet(self.gpa, name)) orelse return .unresolved;
            return try self.local(idx.toInt(), .sheet_attr);
        }
        if (ws.name) |raw_name| {
            const name = try decode(self.arena, .pivot_source_name, raw_name);
            // Tables and defined names share one namespace in Excel, so
            // the order only matters for a workbook no Excel produced.
            const folded = symbols.fold(self.arena, name) catch |e| {
                if (e == error.OutOfMemory) return error.OutOfMemory;
                return .unresolved;
            };
            for (self.tables) |t| {
                if (std.mem.eql(u8, t.folded, folded)) return try self.local(t.sheet_idx, .table);
            }
            switch (try symbols.resolveName(self.gpa, null, name)) {
                .name => |n| {
                    const idx = (try sheetOfBody(self.gpa, symbols, n.body)) orelse return .unresolved;
                    return try self.local(idx, .defined_name);
                },
                .table, .not_found, .refused => return .unresolved,
            }
        }
        return .unresolved;
    }

    fn local(self: *Resolver, sheet_idx: u32, via: ResolvedVia) Error!SourceResolution {
        if (sheet_idx >= self.wb.sheets.len) return .unresolved;
        const part = (try relTarget(self.store, "xl/workbook.xml", self.wb_rels, self.wb.sheets[sheet_idx].r_id)) orelse
            return .unresolved;
        return .{ .sheet = .{
            .sheet_idx = sheet_idx,
            .sheet_name = self.sheet_names[sheet_idx],
            .part_name = part,
            .via = via,
        } };
    }

    fn ensureBuilt(self: *Resolver) Error!void {
        if (self.built) return;
        self.built = true;

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
            .refused => {},
        }
        switch (try builder.finish()) {
            .ok => |t| self.symbols = t,
            .refused => return,
        }

        // The table index: every `<tableParts>` edge of every sheet,
        // keyed by the folded display name a `worksheetSource@name`
        // spells.
        var entries: std.ArrayListUnmanaged(TableEntry) = .empty;
        for (self.wb.sheets, 0..) |s, i| {
            const sheet_part = (try relTarget(self.store, "xl/workbook.xml", self.wb_rels, s.r_id)) orelse continue;
            for (self.store.rels(sheet_part)) |rel| {
                if (rel.target_mode == .external) continue;
                if (!relLeafIs(rel.type, "table")) continue;
                const table_part_name = (try self.store.resolve(sheet_part, rel.target)) orelse continue;
                const table_part = (try self.store.part(table_part_name)) orelse continue;
                const raw = table_edit.tableDisplayNameRaw(table_part.bytes) orelse continue;
                const name = try decode(self.arena, .table_name, raw);
                const folded = self.symbols.?.fold(self.arena, name) catch |e| {
                    if (e == error.OutOfMemory) return error.OutOfMemory;
                    continue;
                };
                try entries.append(self.arena, .{
                    .folded = folded,
                    .sheet_idx = @intCast(i),
                    .part_name = table_part_name,
                });
            }
        }
        self.tables = try entries.toOwnedSlice(self.arena);
    }
};

/// The sheet a defined name's body denotes, when the body is exactly one
/// sheet-qualified reference (`Data!$A$1:$C$4`, `'My Data'!A1:B2`).
/// A 3D span, a dynamic body (`OFFSET(…)`), a union or a bare range
/// resolves to null: none names one sheet a pivot could read.
fn sheetOfBody(gpa: Allocator, symbols: *const engine.SymbolTable, body: []const u8) Error!?u32 {
    var parsed = try engine.parser.parse(gpa, body, .{});
    defer parsed.deinit(gpa);
    const ast = switch (parsed) {
        .ok => |t| t,
        .refused => return null,
    };
    // `Data!A1` is one qualified node; `Data!$A$1:$C$4` is a range
    // operator whose LEFT operand carries the qualifier. Both denote one
    // sheet — unless the right operand names another one.
    const spec: engine.parser.SheetSpec = switch (ast.node(ast.root)) {
        .qualified => |q| q.sheet,
        .binary => |b| blk: {
            if (b.op != .range) return null;
            const lhs = switch (ast.node(b.lhs)) {
                .qualified => |q| q.sheet,
                else => return null,
            };
            switch (ast.node(b.rhs)) {
                .qualified => |q| if (!q.sheet.eql(lhs)) return null,
                else => {},
            }
            break :blk lhs;
        },
        else => return null,
    };
    if (spec.last != null) return null;
    var buf: [256]u8 = undefined;
    const name = unquoteSheetSpec(&buf, spec) orelse return null;
    const idx = (try symbols.resolveSheet(gpa, name)) orelse return null;
    return idx.toInt();
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
            .sheet_ref => "<worksheetSource sheet=\"Data\" ref=\"A1:C4\"/>",
            .table_name => "<worksheetSource name=\"SalesTbl\"/>",
            .defined_name => "<worksheetSource name=\"PivotSrc\"/>",
            .external => "<worksheetSource r:id=\"rIdExt\" sheet=\"Sheet1\" ref=\"A1:C4\"/>",
            .dangling => "<worksheetSource sheet=\"Nope\" ref=\"A1:C4\"/>",
        };
        const cache_def = try std.fmt.allocPrint(allocator,
            \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            \\<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" refreshedBy="zlsx" refreshedDate="45000.5" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="3"><cacheSource type="worksheet">{s}</cacheSource><cacheFields count="3"><cacheField name="Region" numFmtId="0"><sharedItems count="2"><s v="East"/><s v="West"/></sharedItems></cacheField><cacheField name="Qty" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="3" maxValue="5"/></cacheField><cacheField name="Price" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" minValue="1.5" maxValue="3.5"/></cacheField></cacheFields></pivotCacheDefinition>
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
        const wb_tail: []const u8 = if (kind == .defined_name)
            "<definedNames><definedName name=\"PivotSrc\">Data!$A$1:$C$4</definedName></definedNames><pivotCaches><pivotCache cacheId=\"7\" r:id=\"rIdPC1\"/></pivotCaches>"
        else
            "<pivotCaches><pivotCache cacheId=\"7\" r:id=\"rIdPC1\"/></pivotCaches>";
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

/// Patch one part of a saved fixture (byte replace, first occurrence)
/// and save it back.
fn patchPart(io: std.Io, path: []const u8, part: []const u8, old: []const u8, new: []const u8) !void {
    var store = try PartStore.open(testing.allocator, io, path);
    defer store.deinit();
    const p = (try store.part(part)) orelse return error.PartNotFound;
    const at = std.mem.indexOf(u8, p.bytes, old) orelse return error.TestUnexpectedResult;
    const patched = try std.mem.concat(testing.allocator, u8, &.{ p.bytes[0..at], new, p.bytes[at + old.len ..] });
    defer testing.allocator.free(patched);
    try store.replacePart(part, patched);
    try store.save(io, path);
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
    try patchPart(io, none, "xl/pivotCache/pivotCacheDefinition1.xml", " r:id=\"rId1\"", " saveData=\"0\"");
    try patchPart(io, none, "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels", "pivotCacheRecords\" ", "notRecords\" ");
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
