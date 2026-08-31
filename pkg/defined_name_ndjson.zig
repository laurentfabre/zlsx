//! The `defined-names` NDJSON records — S3b's typed workbook-name read
//! (`docs/cli.md`, "defined-names"), written once for every surface.
//!
//! `zlsx defined-names` emits these through its own selection and
//! pagination; the C and Python legs of row S3b hand over the same
//! bytes when they land. One writer, so the record a Python caller
//! will see is byte-for-byte the one the CLI prints — the
//! `pivot_ndjson.zig` precedent.
//!
//! The read is a decode of what `xl/workbook.xml` already parsed:
//! `name` by its string carrier (entities + ST_Xstring — the codec
//! every name attribute is written with), the body by its formula
//! carrier (entities only). Nothing is resolved or rewritten; the
//! body is the formula text as authored.

const std = @import("std");
const formula = @import("zlsx_formula");
const workbook_xml = @import("typed_parts/root.zig").workbook_xml;
const json = @import("json_text.zig");

const Allocator = std.mem.Allocator;

pub const Error = error{ MalformedWorkbookXml, OutOfMemory };

/// One `<definedName>`, decoded and ready to emit. `scope_sheet_idx`
/// is the raw zero-based `localSheetId` as written — kept even when it
/// names no sheet of the workbook, because the attribute is what the
/// producer wrote and dropping it would merge two distinct shapes into
/// the workbook-scope spelling. `scope_sheet` is the decoded name of
/// that sheet, or null when the scope is the workbook or the id is
/// past the sheet list.
pub const DecodedName = struct {
    name: []const u8,
    body: []const u8,
    scope_sheet_idx: ?u32,
    scope_sheet: ?[]const u8,
    hidden: bool,
};

/// Every defined name of a parsed workbook part, decoded, in document
/// order — the order Excel's name manager and the CLI stream share.
/// Owns its decoded strings; `deinit` frees them.
pub const DefinedNames = struct {
    arena: std.heap.ArenaAllocator,
    /// Decoded sheet names, parallel to `WorkbookXml.sheets` — the
    /// inventory `--sheet-glob` and the `sheet` field read from.
    sheet_names: []const []const u8,
    names: []const DecodedName,

    pub fn deinit(self: *DefinedNames) void {
        self.arena.deinit();
        self.* = undefined;
    }
};

/// Decode every defined name of `wb`. A carrier that does not decode
/// (a bad entity, an ill-formed ST_Xstring escape) refuses the whole
/// read — a partial name inventory is the shape of a guard hole, as
/// the pivot read established.
pub fn collect(gpa: Allocator, wb: *const workbook_xml.WorkbookXml) Error!DefinedNames {
    var arena = std.heap.ArenaAllocator.init(gpa);
    errdefer arena.deinit();
    const a = arena.allocator();

    const sheet_names = try a.alloc([]const u8, wb.sheets.len);
    for (wb.sheets, 0..) |s, i| sheet_names[i] = try decode(a, .sheet_name, s.name);

    const names = try a.alloc(DecodedName, wb.defined_names.len);
    for (wb.defined_names, 0..) |dn, i| {
        names[i] = .{
            .name = try decode(a, .defined_name_identifier, dn.name),
            .body = try decode(a, .defined_name_body, dn.formula),
            .scope_sheet_idx = dn.local_sheet_id,
            .scope_sheet = if (dn.local_sheet_id) |sid|
                (if (sid < sheet_names.len) sheet_names[sid] else null)
            else
                null,
            .hidden = dn.hidden,
        };
    }
    return .{ .arena = arena, .sheet_names = sheet_names, .names = names };
}

/// One `{"kind":"defined_name",…}` line. The field order is the
/// docs/cli.md contract; a change here is a wire-format change on
/// every surface at once.
pub fn writeName(out: *std.Io.Writer, d: DecodedName) !void {
    try out.writeAll("{\"kind\":\"defined_name\",\"name\":");
    try json.writeString(out, d.name);
    try out.writeAll(",\"scope\":");
    if (d.scope_sheet_idx == null) {
        try out.writeAll("\"workbook\",\"sheet\":null,\"sheet_idx\":null");
    } else {
        try out.writeAll("\"sheet\",\"sheet\":");
        try json.writeOptString(out, d.scope_sheet);
        try out.writeAll(",\"sheet_idx\":");
        try json.writeOptU32(out, d.scope_sheet_idx);
    }
    try out.writeAll(",\"body\":");
    try json.writeString(out, d.body);
    try out.print(",\"hidden\":{}}}\n", .{d.hidden});
}

/// The unselected stream — every name, document order. The C leg's
/// entry point when it lands.
pub fn writeAll(out: *std.Io.Writer, view: *const DefinedNames) !void {
    for (view.names) |d| try writeName(out, d);
}

fn decode(a: Allocator, site: formula.decode.Site, raw: []const u8) Error![]u8 {
    return formula.decode.decodeAt(a, site, raw) catch |e| switch (e) {
        error.OutOfMemory => error.OutOfMemory,
        else => error.MalformedWorkbookXml,
    };
}

// ─── Tests ───────────────────────────────────────────────────────────

const testing = std.testing;

fn parseWb(xml: []const u8) !workbook_xml.WorkbookXml {
    return workbook_xml.parse(testing.allocator, xml);
}

test "collect + writeName: workbook scope, sheet scope, hidden, document order" {
    const xml =
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" ++
        "<sheets>" ++
        "<sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/>" ++
        "<sheet name=\"R&amp;D\" sheetId=\"2\" r:id=\"rId2\"/>" ++
        "</sheets>" ++
        "<definedNames>" ++
        "<definedName name=\"Prices\">Data!$A$1:$C$4</definedName>" ++
        "<definedName name=\"_xlnm.Print_Area\" localSheetId=\"1\">'R&amp;D'!$A$1:$B$9</definedName>" ++
        "<definedName name=\"Secret\" hidden=\"1\">Data!$Z$1</definedName>" ++
        "</definedNames>" ++
        "</workbook>";
    var wb = try parseWb(xml);
    defer wb.deinit(testing.allocator);

    var view = try collect(testing.allocator, &wb);
    defer view.deinit();

    try testing.expectEqual(@as(usize, 3), view.names.len);

    var buf: [1024]u8 = undefined;
    var w = std.Io.Writer.fixed(&buf);
    try writeAll(&w, &view);
    const expected =
        "{\"kind\":\"defined_name\",\"name\":\"Prices\",\"scope\":\"workbook\",\"sheet\":null,\"sheet_idx\":null,\"body\":\"Data!$A$1:$C$4\",\"hidden\":false}\n" ++
        "{\"kind\":\"defined_name\",\"name\":\"_xlnm.Print_Area\",\"scope\":\"sheet\",\"sheet\":\"R&D\",\"sheet_idx\":1,\"body\":\"'R&D'!$A$1:$B$9\",\"hidden\":false}\n" ++
        "{\"kind\":\"defined_name\",\"name\":\"Secret\",\"scope\":\"workbook\",\"sheet\":null,\"sheet_idx\":null,\"body\":\"Data!$Z$1\",\"hidden\":true}\n";
    try testing.expectEqualStrings(expected, w.buffered());
}

test "collect: a localSheetId past the sheet list keeps the index, sheet null" {
    const xml =
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" ++
        "<sheets>" ++
        "<sheet name=\"Only\" sheetId=\"1\" r:id=\"rId1\"/>" ++
        "</sheets>" ++
        "<definedNames>" ++
        "<definedName name=\"Dangling\" localSheetId=\"7\">Only!$A$1</definedName>" ++
        "</definedNames>" ++
        "</workbook>";
    var wb = try parseWb(xml);
    defer wb.deinit(testing.allocator);

    var view = try collect(testing.allocator, &wb);
    defer view.deinit();

    var buf: [512]u8 = undefined;
    var w = std.Io.Writer.fixed(&buf);
    try writeName(&w, view.names[0]);
    try testing.expectEqualStrings(
        "{\"kind\":\"defined_name\",\"name\":\"Dangling\",\"scope\":\"sheet\",\"sheet\":null,\"sheet_idx\":7,\"body\":\"Only!$A$1\",\"hidden\":false}\n",
        w.buffered(),
    );
}

test "collect: a name body with a bad entity refuses the whole read" {
    const xml =
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" ++
        "<sheets>" ++
        "<sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/>" ++
        "</sheets>" ++
        "<definedNames>" ++
        "<definedName name=\"Ok\">Data!$A$1</definedName>" ++
        "<definedName name=\"Bad\">Data!$A$1&bogus;</definedName>" ++
        "</definedNames>" ++
        "</workbook>";
    var wb = try parseWb(xml);
    defer wb.deinit(testing.allocator);

    try testing.expectError(error.MalformedWorkbookXml, collect(testing.allocator, &wb));
}

test "collect: a workbook without defined names is an empty view" {
    const xml =
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" ++
        "<sheets>" ++
        "<sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/>" ++
        "</sheets>" ++
        "</workbook>";
    var wb = try parseWb(xml);
    defer wb.deinit(testing.allocator);

    var view = try collect(testing.allocator, &wb);
    defer view.deinit();
    try testing.expectEqual(@as(usize, 0), view.names.len);
    try testing.expectEqual(@as(usize, 1), view.sheet_names.len);
}
