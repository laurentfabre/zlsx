//! The `pivots` NDJSON records — S6's frozen shape (`docs/cli.md`,
//! "pivots"), written once for every surface.
//!
//! `zlsx pivots` emits these through its own selection and pagination;
//! `zlsx_editor_pivots_ndjson` (S3a) hands the same bytes to the C ABI
//! and py-zlsx parses them line by line. One writer, so the record a
//! Python caller sees is byte-for-byte the one the CLI prints — a shape
//! frozen at the S6 gate cannot drift between surfaces if only one
//! piece of code knows how to spell it.

const std = @import("std");
const pivots_mod = @import("pivots.zig");
const pivot_xml = @import("typed_parts/pivot_xml.zig");
const json = @import("json_text.zig");

const Pivots = pivots_mod.Pivots;
const PivotTable = pivots_mod.PivotTable;
const PivotCache = pivots_mod.PivotCache;

/// Whether a pivot record carries the `sheet` / `sheet_idx` envelope.
/// `compact` is the CLI's `--output compact-ndjson`, where a sheet
/// prologue record names the sheet once for every record after it.
pub const Envelope = enum { full, compact };

/// Every pivot table in host-sheet order, then every cache no table
/// reads (`{"kind":"pivot_cache",…}`) — the unselected `zlsx pivots`
/// stream. An empty graph writes nothing.
pub fn writeAll(out: *std.Io.Writer, pivots: *const Pivots) !void {
    for (pivots.tables) |pt| try writeTable(out, pivots, pt, .full);
    for (pivots.caches) |c| {
        if (c.consumer_count != 0) continue;
        try writeCacheRecord(out, pivots, c);
    }
}

/// One `{"kind":"pivot",…}` line.
pub fn writeTable(
    out: *std.Io.Writer,
    pivots: *const Pivots,
    pt: PivotTable,
    envelope: Envelope,
) !void {
    const def = &pt.definition;
    try out.writeAll("{\"kind\":\"pivot\"");
    if (envelope == .full) {
        try out.writeAll(",\"sheet\":");
        try json.writeString(out, pt.sheet_name);
        try out.print(",\"sheet_idx\":{d}", .{pt.sheet_idx});
    }
    try out.writeAll(",\"name\":");
    try json.writeString(out, pt.name);
    try out.writeAll(",\"part\":");
    try json.writeString(out, pt.part_name);
    try out.writeAll(",\"location\":{\"ref\":");
    try json.writeString(out, pt.location_ref);
    try out.writeAll(",\"first_header_row\":");
    try json.writeOptU32(out, def.location.first_header_row);
    try out.writeAll(",\"first_data_row\":");
    try json.writeOptU32(out, def.location.first_data_row);
    try out.writeAll(",\"first_data_col\":");
    try json.writeOptU32(out, def.location.first_data_col);
    try out.writeAll("},\"rows\":");
    try writeAxis(out, pivots, pt, def.row_fields);
    try out.writeAll(",\"cols\":");
    try writeAxis(out, pivots, pt, def.col_fields);
    try out.writeAll(",\"pages\":[");
    for (def.page_fields, 0..) |pf, i| {
        if (i > 0) try out.writeByte(',');
        switch (pf.fld) {
            .values => try out.writeAll("{\"values\":true}"),
            .field => |ordinal| {
                try out.writeAll("{\"field\":");
                try json.writeOptString(out, pivots.fieldName(pt, ordinal));
                try out.print(",\"idx\":{d}}}", .{ordinal});
            },
        }
    }
    try out.writeAll("],\"values\":[");
    for (def.data_fields, 0..) |df, i| {
        if (i > 0) try out.writeByte(',');
        try out.writeAll("{\"name\":");
        try json.writeOptString(out, pt.data_field_names[i]);
        try out.writeAll(",\"field\":");
        try json.writeOptString(out, pivots.fieldName(pt, df.fld));
        try out.print(",\"idx\":{d},\"subtotal\":", .{df.fld});
        try json.writeString(out, if (df.subtotal == .unknown) (df.subtotal_raw orelse "unknown") else df.subtotal.xmlName());
        try out.writeAll(",\"show_data_as\":");
        try json.writeOptString(out, df.show_data_as);
        try out.writeAll(",\"num_fmt_id\":");
        try json.writeOptU32(out, df.num_fmt_id);
        try out.writeByte('}');
    }
    try out.writeAll("],\"data_caption\":");
    try json.writeOptString(out, pt.data_caption);
    try out.print(",\"grand_totals\":{{\"rows\":{},\"cols\":{}}},\"style\":", .{ def.row_grand_totals, def.col_grand_totals });
    try json.writeOptString(out, pt.style_name);
    try out.writeAll(",\"cache\":");
    if (pivots.cacheOf(pt)) |c| try writeCacheObject(out, pivots, c.*) else try out.writeAll("null");
    try out.writeAll("}\n");
}

/// One `{"kind":"pivot_cache","cache":{…}}` line — a cache no pivot
/// table reads.
pub fn writeCacheRecord(out: *std.Io.Writer, pivots: *const Pivots, c: PivotCache) !void {
    try out.writeAll("{\"kind\":\"pivot_cache\",\"cache\":");
    try writeCacheObject(out, pivots, c);
    try out.writeAll("}\n");
}

/// `[{"field":"Region","idx":0},{"values":true}]` — a row or column
/// axis: pivot-field ordinals resolved to their cache-field names, the
/// values axis (`x="-2"`) as its own marker.
fn writeAxis(
    out: *std.Io.Writer,
    pivots: *const Pivots,
    pt: PivotTable,
    axis: []const pivot_xml.AxisField,
) !void {
    try out.writeByte('[');
    for (axis, 0..) |af, i| {
        if (i > 0) try out.writeByte(',');
        switch (af) {
            .values => try out.writeAll("{\"values\":true}"),
            .field => |ordinal| {
                try out.writeAll("{\"field\":");
                try json.writeOptString(out, pivots.fieldName(pt, ordinal));
                try out.print(",\"idx\":{d}}}", .{ordinal});
            },
        }
    }
    try out.writeByte(']');
}

pub fn writeCacheObject(out: *std.Io.Writer, pivots: *const Pivots, c: PivotCache) !void {
    const def = &c.definition;
    try out.writeAll("{\"id\":");
    try json.writeOptU32(out, c.cache_id);
    try out.writeAll(",\"part\":");
    try json.writeString(out, c.part_name);
    try out.writeAll(",\"records_part\":");
    try json.writeOptString(out, c.records_part_name);
    try out.writeAll(",\"record_count\":");
    try json.writeOptU32(out, def.record_count);
    try out.writeAll(",\"refreshed_by\":");
    try json.writeOptString(out, def.refreshed_by);
    // `refreshedDate` is the serial Excel writes; a producer that wrote
    // only the ISO form (`refreshedDateIso`) still gets a date here.
    try out.writeAll(",\"refreshed_date\":");
    try json.writeOptString(out, def.refreshed_date orelse def.refreshed_date_iso);
    try out.print(",\"refresh_on_load\":{},\"save_data\":{},\"source\":", .{ def.refresh_on_load, def.save_data });
    try writeSource(out, pivots, c);
    try out.writeAll(",\"fields\":[");
    for (def.fields, 0..) |f, i| {
        if (i > 0) try out.writeByte(',');
        try out.writeAll("{\"name\":");
        try json.writeString(out, c.field_names[i]);
        try out.writeAll(",\"num_fmt_id\":");
        try json.writeOptU32(out, f.num_fmt_id);
        try out.writeAll(",\"formula\":");
        try json.writeOptString(out, c.field_formulas[i]);
        try out.writeAll(",\"items\":");
        try json.writeOptU32(out, if (f.shared_items) |si| si.count else null);
        try out.writeAll(",\"types\":");
        if (f.shared_items) |si| try writeSharedItemTypes(out, si) else try out.writeAll("null");
        try out.writeAll(",\"min\":");
        try json.writeOptString(out, if (f.shared_items) |si| (si.min_value orelse si.min_date) else null);
        try out.writeAll(",\"max\":");
        try json.writeOptString(out, if (f.shared_items) |si| (si.max_value orelse si.max_date) else null);
        try out.writeByte('}');
    }
    try out.writeAll("]}");
}

/// The `containsX` inventory as a list of the kinds present, in a fixed
/// order: `string`, `number`, `integer`, `blank`, `date`, `mixed`.
fn writeSharedItemTypes(out: *std.Io.Writer, si: pivot_xml.SharedItems) !void {
    try out.writeByte('[');
    var first = true;
    const kinds = [_]struct { on: bool, name: []const u8 }{
        .{ .on = si.contains_string, .name = "string" },
        .{ .on = si.contains_number, .name = "number" },
        .{ .on = si.contains_integer, .name = "integer" },
        .{ .on = si.contains_blank, .name = "blank" },
        .{ .on = si.contains_date, .name = "date" },
        .{ .on = si.contains_mixed_types, .name = "mixed" },
    };
    for (kinds) |k| {
        if (!k.on) continue;
        if (!first) try out.writeByte(',');
        first = false;
        try out.print("\"{s}\"", .{k.name});
    }
    try out.writeByte(']');
}

fn writeSource(out: *std.Io.Writer, pivots: *const Pivots, c: PivotCache) !void {
    const src = &c.definition.source;
    switch (src.type) {
        .worksheet => {
            try out.writeAll("{\"type\":\"worksheet\",");
            try writeSourceSpelling(out, pivots, c.source, c.resolution);
            try out.writeByte('}');
        },
        .consolidation => {
            try out.writeAll("{\"type\":\"consolidation\",\"range_sets\":[");
            for (c.range_set_sources, 0..) |sp, i| {
                if (i > 0) try out.writeByte(',');
                try out.writeByte('{');
                try writeSourceSpelling(out, pivots, sp, c.range_set_resolutions[i]);
                try out.writeByte('}');
            }
            try out.writeAll("]}");
        },
        .external => {
            try out.writeAll("{\"type\":\"external\",\"connection_id\":");
            try json.writeOptU32(out, src.connection_id);
            try out.writeByte('}');
        },
        .scenario => try out.writeAll("{\"type\":\"scenario\"}"),
        .unknown => {
            try out.writeAll("{\"type\":\"unknown\",\"raw\":");
            try json.writeOptString(out, src.type_raw);
            try out.writeByte('}');
        },
    }
}

/// `"sheet":…,"ref":…,"name":…,"resolved":…,"unresolved":…` (no braces,
/// no leading comma) — the spellings as written and what they led to: a
/// local sheet (`{"sheet":"Data","sheet_idx":0,"via":"sheet_attr",
/// "bounds":"A1:C4"}` — `bounds` the A1 area the spelling proves, or
/// `null`), another workbook (`{"external":"…"}`), or `null` when the
/// spelling names nothing this workbook has — in which case
/// `unresolved` says why and which sheets it still proves (S7b-1).
fn writeSourceSpelling(
    out: *std.Io.Writer,
    pivots: *const Pivots,
    sp: pivots_mod.SourceSpelling,
    res: pivots_mod.SourceResolution,
) !void {
    try out.writeAll("\"sheet\":");
    try json.writeOptString(out, sp.sheet);
    try out.writeAll(",\"ref\":");
    try json.writeOptString(out, sp.ref);
    try out.writeAll(",\"name\":");
    try json.writeOptString(out, sp.name);
    try out.writeAll(",\"resolved\":");
    switch (res) {
        .sheet => |s| {
            try out.writeAll("{\"sheet\":");
            try json.writeString(out, s.sheet_name);
            try out.print(",\"sheet_idx\":{d},\"via\":\"{s}\",\"bounds\":", .{ s.sheet_idx, @tagName(s.via) });
            var buf: [pivots_mod.Bounds.format_buf_len]u8 = undefined;
            try json.writeOptString(out, if (s.bounds) |b| b.formatA1(&buf) else null);
            try out.writeByte('}');
        },
        .external => |target| {
            try out.writeAll("{\"external\":");
            try json.writeString(out, target);
            try out.writeByte('}');
        },
        .unresolved, .none => try out.writeAll("null"),
    }
    try out.writeAll(",\"unresolved\":");
    switch (res) {
        .unresolved => |u| {
            try out.print("{{\"why\":\"{s}\",\"sheets\":[", .{@tagName(u.why)});
            for (u.sheets, 0..) |idx, k| {
                if (k > 0) try out.writeByte(',');
                try out.writeAll("{\"sheet\":");
                try json.writeOptString(out, if (idx < pivots.sheet_names.len) pivots.sheet_names[idx] else null);
                try out.print(",\"sheet_idx\":{d}}}", .{idx});
            }
            try out.writeAll("]}");
        },
        .sheet, .external, .none => try out.writeAll("null"),
    }
}

// ─── Tests ───────────────────────────────────────────────────────────

const workbook_mod = @import("workbook.zig");

/// The record the S6 gate froze for `pivots.fixture.write(.sheet_ref)`:
/// the same literal `src/cli.zig`'s `runPivotsCommand` test pins, so a
/// drift between the two writers fails here and there.
pub const fixture_sheet_ref_record =
    "{\"kind\":\"pivot\",\"sheet\":\"Report\",\"sheet_idx\":1,\"name\":\"PivotTable1\"," ++
    "\"part\":\"xl/pivotTables/pivotTable1.xml\"," ++
    "\"location\":{\"ref\":\"A3:B6\",\"first_header_row\":1,\"first_data_row\":1,\"first_data_col\":1}," ++
    "\"rows\":[{\"field\":\"Region\",\"idx\":0}],\"cols\":[],\"pages\":[]," ++
    "\"values\":[{\"name\":\"Sum of Qty\",\"field\":\"Qty\",\"idx\":1,\"subtotal\":\"sum\",\"show_data_as\":null,\"num_fmt_id\":null}]," ++
    "\"data_caption\":\"Values\",\"grand_totals\":{\"rows\":true,\"cols\":true},\"style\":\"PivotStyleLight16\"," ++
    "\"cache\":{\"id\":7,\"part\":\"xl/pivotCache/pivotCacheDefinition1.xml\"," ++
    "\"records_part\":\"xl/pivotCache/pivotCacheRecords1.xml\",\"record_count\":3," ++
    "\"refreshed_by\":\"zlsx\",\"refreshed_date\":\"45000.5\",\"refresh_on_load\":false,\"save_data\":true," ++
    "\"source\":{\"type\":\"worksheet\",\"sheet\":\"Data\",\"ref\":\"A1:C4\",\"name\":null," ++
    "\"resolved\":{\"sheet\":\"Data\",\"sheet_idx\":0,\"via\":\"sheet_attr\",\"bounds\":\"A1:C4\"},\"unresolved\":null}," ++
    "\"fields\":[" ++
    "{\"name\":\"Region\",\"num_fmt_id\":0,\"formula\":null,\"items\":2,\"types\":[\"string\"],\"min\":null,\"max\":null}," ++
    "{\"name\":\"Qty\",\"num_fmt_id\":0,\"formula\":null,\"items\":null,\"types\":[\"number\",\"integer\"],\"min\":\"3\",\"max\":\"5\"}," ++
    "{\"name\":\"Price\",\"num_fmt_id\":0,\"formula\":null,\"items\":null,\"types\":[\"number\"],\"min\":\"1.5\",\"max\":\"3.5\"}" ++
    "]}}\n";

fn tmpFixturePath(alloc: std.mem.Allocator, io: std.Io, dir: *std.testing.TmpDir, name: []const u8) ![:0]u8 {
    const d = try dir.dir.realPathFileAlloc(io, ".", alloc);
    defer alloc.free(d);
    return std.fs.path.joinZ(alloc, &.{ d, name });
}

test "writeAll: the frozen record for the sheet_ref fixture, then nothing (no orphan)" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    const path = try tmpFixturePath(alloc, io, &tmp, "ndjson_sheet_ref.xlsx");
    defer alloc.free(path);
    try pivots_mod.fixture.write(alloc, io, path, .sheet_ref);

    var wb = try workbook_mod.Workbook.open(alloc, io, path);
    defer wb.deinit();
    var pv = try wb.pivotTables();
    defer pv.deinit();

    var scratch: [8192]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try writeAll(&w, &pv);
    try std.testing.expectEqualStrings(fixture_sheet_ref_record, w.buffered());

    // The compact envelope drops exactly the two sheet keys.
    w = std.Io.Writer.fixed(&scratch);
    try writeTable(&w, &pv, pv.tables[0], .compact);
    const compact = w.buffered();
    try std.testing.expect(std.mem.startsWith(u8, compact, "{\"kind\":\"pivot\",\"name\":\"PivotTable1\""));
    try std.testing.expect(std.mem.indexOf(u8, compact, "\"sheet_idx\":1") == null);
}

test "writeAll: an orphan cache follows the tables as its own record" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    const path = try tmpFixturePath(alloc, io, &tmp, "ndjson_orphan.xlsx");
    defer alloc.free(path);
    try pivots_mod.fixture.writeWithOrphanCache(alloc, io, path, .sheet_ref);

    var wb = try workbook_mod.Workbook.open(alloc, io, path);
    defer wb.deinit();
    var pv = try wb.pivotTables();
    defer pv.deinit();

    var scratch: [16384]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try writeAll(&w, &pv);
    const got = w.buffered();
    try std.testing.expect(std.mem.startsWith(u8, got, fixture_sheet_ref_record));
    const rest = got[fixture_sheet_ref_record.len..];
    try std.testing.expect(std.mem.startsWith(u8, rest, "{\"kind\":\"pivot_cache\",\"cache\":{\"id\":"));
    try std.testing.expectEqual(@as(usize, 2), std.mem.count(u8, got, "\n"));
}
