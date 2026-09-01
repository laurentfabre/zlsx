//! The `conditional-formats` NDJSON records — S3b's typed
//! conditional-format read (`docs/cli.md`, "conditional-formats"),
//! written once for every surface.
//!
//! `zlsx conditional-formats` emits these through its own selection
//! and pagination; the C and Python legs of row S3b hand over the same
//! bytes when they land — the `pivot_ndjson.zig` /
//! `defined_name_ndjson.zig` / `anchor_ndjson.zig` precedent.
//!
//! The view is `Worksheet.conditionalFormats` attributed to workbook
//! sheets: one record per `<cfRule>`, sheets in workbook order, rules
//! in sheet-document order, each carrying its parent
//! `<conditionalFormatting>` block's `sqref`. The record reports the
//! rule envelope — where it applies, what kind it is, its formula
//! bodies, its differential style id and priority — not the visual
//! payload (`<colorScale>` / `<dataBar>` / `<iconSet>` children stay
//! in the part, byte-preserved, for callers that need them raw).
//!
//! Decode discipline: `sqref` and the rule type are plain attribute
//! carriers, the formulas element-text formula carriers — entities
//! resolve, nothing else (no ST_Xstring layer on either, the C1
//! ruling). A carrier that does not decode, decodes to non-UTF-8, or
//! carries embedded markup refuses the whole read — a partial or
//! wrong rule inventory is the shape of a guard hole, as the pivot,
//! defined-name and anchor reads established.

const std = @import("std");
const formula_mod = @import("zlsx_formula");
const workbook_mod = @import("workbook.zig");
const store_mod = @import("store.zig");
const json = @import("json_text.zig");

const Allocator = std.mem.Allocator;
const PartStore = store_mod.PartStore;

/// Whether a record carries the `sheet` / `sheet_idx` envelope.
/// `compact` is the CLI's `--output compact-ndjson`, where a sheet
/// prologue record names the sheet once for every record after it.
pub const Envelope = enum { full, compact };

/// One `<cfRule>`, attributed and decoded. `sqref` is the parent
/// `<conditionalFormatting>` block's target list as authored
/// (space-separated A1 areas); `formulas` holds the rule's up-to-three
/// `<formula>` bodies in document order (a `cellIs` `between` carries
/// two). `sqref` and `rule_type` are the empty string when the source
/// omits the attribute — the lenient sheet view models absent and
/// empty as one shape, and this read reports the view.
pub const Record = struct {
    sheet: []const u8,
    sheet_idx: u32,
    sqref: []const u8,
    rule_type: []const u8,
    formulas: []const []const u8,
    dxf_id: ?u32,
    priority: ?u32,
};

/// Every conditional-format rule of a workbook, attributed and in
/// emission order: sheets in workbook order, rules in sheet-document
/// order. Owns its decoded strings; `deinit` frees them.
pub const ConditionalFormats = struct {
    arena: std.heap.ArenaAllocator,
    /// Decoded sheet names, parallel to `WorkbookXml.sheets` — the
    /// inventory the CLI's selectors and the `sheet` field read from.
    sheet_names: []const []const u8,
    records: []const Record,

    pub fn deinit(self: *ConditionalFormats) void {
        self.arena.deinit();
        self.* = undefined;
    }
};

pub const Error = error{
    /// A sheet-name carrier that does not decode or decodes to
    /// non-UTF-8 — the NDJSON must stay parseable.
    MalformedWorkbookXml,
    /// A rule field the NDJSON cannot carry faithfully: a `sqref`,
    /// rule type or formula whose carrier does not decode, is not
    /// UTF-8, or carries embedded markup.
    MalformedSheetXml,
    OutOfMemory,
};

/// Collect every rule of every workbook sheet. A read this module
/// cannot serve faithfully refuses whole; a sheet the workbook itself
/// cannot read (a dangling relationship, a part that does not parse)
/// passes through as `Workbook`'s own error, exactly as the
/// `Worksheet.conditionalFormats` caller would see it.
pub fn collect(gpa: Allocator, wb: *workbook_mod.Workbook) !ConditionalFormats {
    var arena = std.heap.ArenaAllocator.init(gpa);
    errdefer arena.deinit();
    const a = arena.allocator();

    const sheet_names = try a.alloc([]const u8, wb.workbook.sheets.len);
    for (wb.workbook.sheets, 0..) |s, i| sheet_names[i] = try decodeSheetName(a, s.name);

    var records: std.ArrayListUnmanaged(Record) = .empty;
    for (0..wb.workbook.sheets.len) |idx| {
        const ws = try wb.sheet(@intCast(idx));
        const cfs = try ws.conditionalFormats();
        for (cfs) |cf| {
            var formulas: std.ArrayListUnmanaged([]const u8) = .empty;
            for ([_]?[]const u8{ cf.formula, cf.formula2, cf.formula3 }) |maybe| {
                const raw = maybe orelse break;
                try formulas.append(a, try decodeRuleText(a, raw));
            }
            try records.append(a, .{
                .sheet = sheet_names[idx],
                .sheet_idx = @intCast(idx),
                .sqref = try decodeRuleText(a, cf.sqref),
                .rule_type = try decodeRuleText(a, cf.type),
                .formulas = try formulas.toOwnedSlice(a),
                .dxf_id = cf.dxf_id,
                .priority = cf.priority,
            });
        }
    }

    return .{
        .arena = arena,
        .sheet_names = sheet_names,
        .records = try records.toOwnedSlice(a),
    };
}

/// One `{"kind":"conditional_format",…}` line. The field order is the
/// docs/cli.md contract; a change here is a wire-format change on
/// every surface at once.
pub fn writeRecord(out: *std.Io.Writer, r: Record, envelope: Envelope) !void {
    try out.writeAll("{\"kind\":\"conditional_format\"");
    if (envelope == .full) {
        try out.writeAll(",\"sheet\":");
        try json.writeString(out, r.sheet);
        try out.print(",\"sheet_idx\":{d}", .{r.sheet_idx});
    }
    try out.writeAll(",\"sqref\":");
    try json.writeString(out, r.sqref);
    try out.writeAll(",\"rule_type\":");
    try json.writeString(out, r.rule_type);
    try out.writeAll(",\"formulas\":[");
    for (r.formulas, 0..) |f, i| {
        if (i > 0) try out.writeByte(',');
        try json.writeString(out, f);
    }
    try out.writeAll("],\"dxf_id\":");
    try json.writeOptU32(out, r.dxf_id);
    try out.writeAll(",\"priority\":");
    try json.writeOptU32(out, r.priority);
    try out.writeAll("}\n");
}

/// The unselected stream — every record, emission order. The future C
/// leg's entry point.
pub fn writeAll(out: *std.Io.Writer, view: *const ConditionalFormats) !void {
    for (view.records) |r| try writeRecord(out, r, .full);
}

fn decodeSheetName(a: Allocator, raw: []const u8) Error![]u8 {
    const decoded = formula_mod.decode.decodeAt(a, .sheet_name, raw) catch |e| switch (e) {
        error.OutOfMemory => return error.OutOfMemory,
        else => return error.MalformedWorkbookXml,
    };
    if (!std.unicode.utf8ValidateSlice(decoded)) return error.MalformedWorkbookXml;
    return decoded;
}

/// `sqref` / `type` are attribute values, the formulas element text —
/// all three are entities-only carriers (the formula carrier's
/// decode; no ST_Xstring layer). A raw `<` in any of them can only be
/// markup the lenient sheet view sliced through — an ill-formed
/// attribute, or an element nested where the schema puts formula text
/// — content no consumer of the rule reads through, so it refuses,
/// as do a carrier that does not decode and a decoded value that is
/// not UTF-8 (the JSON writer passes bytes through verbatim).
fn decodeRuleText(a: Allocator, raw: []const u8) Error![]u8 {
    if (std.mem.indexOfScalar(u8, raw, '<') != null) return error.MalformedSheetXml;
    const decoded = formula_mod.decode.decodeAt(a, .cell_formula_body, raw) catch |e| switch (e) {
        error.OutOfMemory => return error.OutOfMemory,
        else => return error.MalformedSheetXml,
    };
    if (!std.unicode.utf8ValidateSlice(decoded)) return error.MalformedSheetXml;
    return decoded;
}

// ─── Test fixture ────────────────────────────────────────────────────

/// Writes a real two-sheet workbook with conditional formats on BOTH
/// sheets through the public Writer, so the read exercises exactly
/// the bytes the write path produces: `Data` carries all four rule
/// families the writer can author (a `between` cellIs — two formula
/// bodies — an expression whose formula needs entity decoding, a
/// colorScale and a dataBar, priorities 1..4 in document order),
/// `Report` one expression rule whose formula carries a literal `&`.
/// `src/cli.zig` and the tests below share it.
pub const fixture = struct {
    pub fn write(allocator: Allocator, io: std.Io, path: []const u8) !void {
        const zlsx = @import("zlsx");
        var w = zlsx.Writer.init(allocator);
        defer w.deinit();
        const dxf = try w.addDxf(.{ .font_bold = true, .fill_fg_argb = 0xFFFFC7CE });
        var data = try w.addSheet("Data");
        try data.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 5 }, .{ .integer = 9 }, .{ .integer = 3 } });
        try data.writeRow(&.{ .{ .integer = 2 }, .{ .integer = 6 }, .{ .integer = 10 }, .{ .integer = 4 } });
        try data.addConditionalFormatCellIs("A1:A4", .between, "2", "4", dxf);
        try data.addConditionalFormatExpression("B1:B4", "B1>3", dxf);
        try data.addConditionalFormatColorScale("C1:C4", 0xFFF8696B, null, 0xFF63BE7B);
        try data.addConditionalFormatDataBar("D1:D4", 0xFF638EC6);
        var report = try w.addSheet("Report");
        try report.writeRow(&.{.{ .string = "R&D" }});
        try report.addConditionalFormatExpression("A1:A2", "$A1=\"R&D\"", dxf);
        try w.save(io, path);
    }

    /// Byte-replace the first `old` in one part of a saved workbook
    /// and save it back — how the refusal tests make a fixture wrong
    /// in exactly one place (the `pivots.fixture` / `anchor_ndjson`
    /// helper, duplicated so this module keeps its imports narrow).
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

test "writeRecord: full and compact envelopes, exact bytes" {
    const rec: Record = .{
        .sheet = "R\"D",
        .sheet_idx = 1,
        .sqref = "A1:A4 C2",
        .rule_type = "cellIs",
        .formulas = &.{ "2", "SUM(A1,\"x\")" },
        .dxf_id = 0,
        .priority = 1,
    };
    var buf: [512]u8 = undefined;
    var w = std.Io.Writer.fixed(&buf);
    try writeRecord(&w, rec, .full);
    try testing.expectEqualStrings(
        "{\"kind\":\"conditional_format\",\"sheet\":\"R\\\"D\",\"sheet_idx\":1," ++
            "\"sqref\":\"A1:A4 C2\",\"rule_type\":\"cellIs\"," ++
            "\"formulas\":[\"2\",\"SUM(A1,\\\"x\\\")\"],\"dxf_id\":0,\"priority\":1}\n",
        w.buffered(),
    );
    var w2 = std.Io.Writer.fixed(&buf);
    try writeRecord(&w2, rec, .compact);
    try testing.expectEqualStrings(
        "{\"kind\":\"conditional_format\",\"sqref\":\"A1:A4 C2\",\"rule_type\":\"cellIs\"," ++
            "\"formulas\":[\"2\",\"SUM(A1,\\\"x\\\")\"],\"dxf_id\":0,\"priority\":1}\n",
        w2.buffered(),
    );
}

test "writeRecord: empty formulas, absent dxf and priority, empty sqref and type" {
    const rec: Record = .{
        .sheet = "S",
        .sheet_idx = 0,
        .sqref = "",
        .rule_type = "",
        .formulas = &.{},
        .dxf_id = null,
        .priority = null,
    };
    var buf: [256]u8 = undefined;
    var w = std.Io.Writer.fixed(&buf);
    try writeRecord(&w, rec, .full);
    try testing.expectEqualStrings(
        "{\"kind\":\"conditional_format\",\"sheet\":\"S\",\"sheet_idx\":0," ++
            "\"sqref\":\"\",\"rule_type\":\"\",\"formulas\":[],\"dxf_id\":null,\"priority\":null}\n",
        w.buffered(),
    );
}

test "collect: fixture rules attributed in sheet then document order, entities decoded" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "cf_fixture.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);

    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    var view = try collect(testing.allocator, &wb);
    defer view.deinit();

    try testing.expectEqual(@as(usize, 2), view.sheet_names.len);
    try testing.expectEqual(@as(usize, 5), view.records.len);

    const between = view.records[0];
    try testing.expectEqualStrings("Data", between.sheet);
    try testing.expectEqual(@as(u32, 0), between.sheet_idx);
    try testing.expectEqualStrings("A1:A4", between.sqref);
    try testing.expectEqualStrings("cellIs", between.rule_type);
    try testing.expectEqual(@as(usize, 2), between.formulas.len);
    try testing.expectEqualStrings("2", between.formulas[0]);
    try testing.expectEqualStrings("4", between.formulas[1]);
    try testing.expectEqual(@as(?u32, 0), between.dxf_id);
    try testing.expectEqual(@as(?u32, 1), between.priority);

    const expr = view.records[1];
    try testing.expectEqualStrings("expression", expr.rule_type);
    // The writer stored `B1&gt;3`; the read hands back what was
    // authored.
    try testing.expectEqual(@as(usize, 1), expr.formulas.len);
    try testing.expectEqualStrings("B1>3", expr.formulas[0]);
    try testing.expectEqual(@as(?u32, 2), expr.priority);

    const scale = view.records[2];
    try testing.expectEqualStrings("colorScale", scale.rule_type);
    try testing.expectEqual(@as(usize, 0), scale.formulas.len);
    try testing.expectEqual(@as(?u32, null), scale.dxf_id);
    try testing.expectEqual(@as(?u32, 3), scale.priority);

    const bar = view.records[3];
    try testing.expectEqualStrings("dataBar", bar.rule_type);
    try testing.expectEqualStrings("D1:D4", bar.sqref);
    try testing.expectEqual(@as(usize, 0), bar.formulas.len);

    const report = view.records[4];
    try testing.expectEqualStrings("Report", report.sheet);
    try testing.expectEqual(@as(u32, 1), report.sheet_idx);
    try testing.expectEqualStrings("$A1=\"R&D\"", report.formulas[0]);
    try testing.expectEqual(@as(?u32, 1), report.priority);
}

test "collect: a rule field the stream cannot carry faithfully refuses whole" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();

    const sheet1 = "xl/worksheets/sheet1.xml";
    const cases = [_]struct { name: []const u8, old: []const u8, new: []const u8 }{
        // A formula carrier that does not decode (a bad entity).
        .{ .name = "r1.xlsx", .old = "B1&gt;3", .new = "B1&bogus;3" },
        // Embedded markup where the schema puts formula text.
        .{ .name = "r2.xlsx", .old = "<formula>2</formula>", .new = "<formula>2<x/></formula>" },
        // A formula that decodes to non-UTF-8.
        .{ .name = "r3.xlsx", .old = "B1&gt;3", .new = "B1\xff3" },
        // A sqref carrier that does not decode.
        .{ .name = "r4.xlsx", .old = "sqref=\"A1:A4\"", .new = "sqref=\"A1&bogus;A4\"" },
        // A rule type that is not UTF-8.
        .{ .name = "r5.xlsx", .old = "type=\"dataBar\"", .new = "type=\"data\xffBar\"" },
    };
    for (cases) |case| {
        const path = try tt.path(testing.allocator, io, case.name);
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path);
        try fixture.patchPart(testing.allocator, io, path, sheet1, case.old, case.new);
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        try testing.expectError(error.MalformedSheetXml, collect(testing.allocator, &wb));
    }
}

test "collect: a sheet-name carrier that does not decode refuses whole" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "cf_bad_sheet_name.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);
    try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "name=\"Report\"", "name=\"Rep&bogus;\"");
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    try testing.expectError(error.MalformedWorkbookXml, collect(testing.allocator, &wb));
}

test "collect: absent sqref or type reads as the empty string — the view's merge, pinned" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "cf_absent_attrs.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);
    try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet1.xml", " sqref=\"D1:D4\"", "");
    try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet1.xml", "type=\"dataBar\" ", "");
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    var view = try collect(testing.allocator, &wb);
    defer view.deinit();
    try testing.expectEqual(@as(usize, 5), view.records.len);
    try testing.expectEqualStrings("", view.records[3].sqref);
    try testing.expectEqualStrings("", view.records[3].rule_type);
}

test "collect: a bodiless self-closing conditionalFormatting block emits nothing" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "cf_selfclosing_block.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);
    try fixture.patchPart(
        testing.allocator,
        io,
        path,
        "xl/worksheets/sheet2.xml",
        "<conditionalFormatting sqref=\"A1:A2\">",
        "<conditionalFormatting sqref=\"Z9\"/><conditionalFormatting sqref=\"A1:A2\">",
    );
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    var view = try collect(testing.allocator, &wb);
    defer view.deinit();
    // Still five records — the empty block contributes none.
    try testing.expectEqual(@as(usize, 5), view.records.len);
    try testing.expectEqualStrings("A1:A2", view.records[4].sqref);
}

test "collect: a workbook without conditional formats is an empty view" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "cf_none.xlsx");
    defer testing.allocator.free(path);
    {
        const zlsx = @import("zlsx");
        var w = zlsx.Writer.init(testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Only");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, path);
    }
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    var view = try collect(testing.allocator, &wb);
    defer view.deinit();
    try testing.expectEqual(@as(usize, 0), view.records.len);
    try testing.expectEqual(@as(usize, 1), view.sheet_names.len);
}
