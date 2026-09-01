//! The `anchors` NDJSON records — S3b's typed drawing-anchor read
//! (`docs/cli.md`, "anchors"), written once for every surface.
//!
//! `zlsx anchors` emits these through its own selection and pagination;
//! the C and Python legs of row S3b hand over the same bytes when they
//! land — the `pivot_ndjson.zig` / `defined_name_ndjson.zig` precedent.
//!
//! The view is `drawings.imageAnchors` + `drawings.chartAnchors`
//! attributed to workbook sheets: the walkers key anchors by worksheet
//! *part* name, this module resolves each part back to the `<sheets>`
//! entry that owns it so every record carries the same `sheet` /
//! `sheet_idx` envelope as the rest of the read family. Image bytes and
//! chart XML stay off the wire — the record reports the anchor geometry
//! and where the payload lives (`part`), not the payload itself.

const std = @import("std");
const formula = @import("zlsx_formula");
const drawings = @import("drawings.zig");
const store_mod = @import("store.zig");
const workbook_xml = @import("typed_parts/root.zig").workbook_xml;
const json = @import("json_text.zig");

const Allocator = std.mem.Allocator;
const PartStore = store_mod.PartStore;

/// Whether a record carries the `sheet` / `sheet_idx` envelope.
/// `compact` is the CLI's `--output compact-ndjson`, where a sheet
/// prologue record names the sheet once for every record after it.
pub const Envelope = enum { full, compact };

/// One anchored image, attributed. `from` / `to` are the cell-grid
/// anchors exactly as `drawings.ImageAnchor` models them (0-based
/// col/row + EMU offsets); `from` is null — not the walker's zero
/// sentinel — when the source used `<xdr:absoluteAnchor>`.
pub const ImageRecord = struct {
    sheet: []const u8,
    sheet_idx: u32,
    /// Archive name of the image part, e.g. `xl/media/image1.png`.
    part: []const u8,
    from: ?drawings.CellAnchor,
    to: ?drawings.CellAnchor,
    absolute: ?drawings.AbsoluteAnchor,
    /// Decompressed size of the image part in bytes — the payload is
    /// not on the wire.
    byte_count: u64,
};

/// One anchored chart, attributed. Same anchor contract as
/// `ImageRecord`; `series_refs` are the chart's `<c:f>` formula refs,
/// entity-decoded, in document order.
pub const ChartRecord = struct {
    sheet: []const u8,
    sheet_idx: u32,
    /// Archive name of the chart part, e.g. `xl/charts/chart1.xml`.
    part: []const u8,
    from: ?drawings.CellAnchor,
    to: ?drawings.CellAnchor,
    absolute: ?drawings.AbsoluteAnchor,
    chart_type: drawings.ChartType,
    series_refs: []const []const u8,
};

pub const Record = union(enum) {
    image: ImageRecord,
    chart: ChartRecord,

    pub fn sheetIdx(self: Record) u32 {
        return switch (self) {
            .image => |r| r.sheet_idx,
            .chart => |r| r.sheet_idx,
        };
    }

    pub fn sheetName(self: Record) []const u8 {
        return switch (self) {
            .image => |r| r.sheet,
            .chart => |r| r.sheet,
        };
    }
};

/// Every anchored image and chart of a workbook, attributed and in
/// emission order: sheets in workbook order, a sheet's images before
/// its charts, each class in drawing-document order. Owns its decoded
/// strings; `deinit` frees them.
pub const Anchors = struct {
    arena: std.heap.ArenaAllocator,
    /// Decoded sheet names, parallel to `WorkbookXml.sheets` — the
    /// inventory the CLI's selectors and the `sheet` field read from.
    sheet_names: []const []const u8,
    records: []const Record,

    pub fn deinit(self: *Anchors) void {
        self.arena.deinit();
        self.* = undefined;
    }
};

pub const Error = error{
    /// A sheet-name carrier that does not decode, or decodes to
    /// non-UTF-8 — the NDJSON must stay parseable.
    MalformedWorkbookXml,
    /// An anchor on a worksheet part `xl/workbook.xml` does not list:
    /// the record could not carry a truthful `sheet` / `sheet_idx`,
    /// and dropping it would be a partial inventory.
    DrawingOnUnlistedSheet,
    /// A part name or series ref the NDJSON cannot carry faithfully:
    /// not UTF-8, a ref whose carrier does not decode, or a ref with
    /// embedded markup.
    MalformedDrawingXml,
    OutOfMemory,
};

/// Collect and attribute every anchor. A read this module cannot serve
/// faithfully refuses whole — a partial anchor inventory is the shape
/// of a guard hole, as the pivot and defined-name reads established.
/// Store-level failures (a part that does not decompress) pass through
/// as their own errors.
pub fn collect(
    gpa: Allocator,
    store: *PartStore,
    wb: *const workbook_xml.WorkbookXml,
) !Anchors {
    var arena = std.heap.ArenaAllocator.init(gpa);
    errdefer arena.deinit();
    const a = arena.allocator();

    const sheet_names = try a.alloc([]const u8, wb.sheets.len);
    for (wb.sheets, 0..) |s, i| sheet_names[i] = try decodeSheetName(a, s.name);

    // Resolve each workbook sheet to its part, best-effort: a sheet
    // whose relationship dangles simply owns no part name here. That
    // sheet cannot host anchors (the walkers never see it), so the
    // strictness lands where it matters — below, on an anchor whose
    // part no listed sheet owns.
    const wb_rels = store.rels("xl/workbook.xml");
    const sheet_parts = try a.alloc(?[]const u8, wb.sheets.len);
    for (wb.sheets, 0..) |s, i| {
        sheet_parts[i] = null;
        const rel = relById(wb_rels, s.r_id) orelse continue;
        if (rel.target_mode == .external) continue;
        sheet_parts[i] = try store.resolve("xl/workbook.xml", rel.target);
    }

    // The walkers allocate their result slices; handing them the view
    // arena parks that scaffolding with everything else this view owns.
    const images = try drawings.imageAnchors(store, a);
    const charts = try drawings.chartAnchors(store, a);

    const image_seen = try a.alloc(bool, images.len);
    @memset(image_seen, false);
    const chart_seen = try a.alloc(bool, charts.len);
    @memset(chart_seen, false);

    var records: std.ArrayListUnmanaged(Record) = .empty;
    for (0..wb.sheets.len) |idx| {
        const part_name = sheet_parts[idx] orelse continue;
        for (images, 0..) |img, i| {
            if (!std.mem.eql(u8, img.sheet_part_name, part_name)) continue;
            image_seen[i] = true;
            try records.append(a, .{ .image = .{
                .sheet = sheet_names[idx],
                .sheet_idx = @intCast(idx),
                .part = try retainPartName(a, img.image_part_name),
                .from = if (img.absolute == null) img.from else null,
                .to = img.to,
                .absolute = img.absolute,
                .byte_count = img.bytes.len,
            } });
        }
        for (charts, 0..) |ch, i| {
            if (!std.mem.eql(u8, ch.sheet_part_name, part_name)) continue;
            chart_seen[i] = true;
            const refs = try a.alloc([]const u8, ch.series_refs.len);
            for (ch.series_refs, 0..) |raw, ri| refs[ri] = try decodeSeriesRef(a, raw);
            try records.append(a, .{ .chart = .{
                .sheet = sheet_names[idx],
                .sheet_idx = @intCast(idx),
                .part = try retainPartName(a, ch.chart_part_name),
                .from = if (ch.absolute == null) ch.from else null,
                .to = ch.to,
                .absolute = ch.absolute,
                .chart_type = ch.chart_type,
                .series_refs = refs,
            } });
        }
    }
    for (image_seen) |seen| if (!seen) return error.DrawingOnUnlistedSheet;
    for (chart_seen) |seen| if (!seen) return error.DrawingOnUnlistedSheet;

    return .{
        .arena = arena,
        .sheet_names = sheet_names,
        .records = try records.toOwnedSlice(a),
    };
}

/// One `{"kind":"image_anchor",…}` or `{"kind":"chart_anchor",…}`
/// line. The field order is the docs/cli.md contract; a change here is
/// a wire-format change on every surface at once.
pub fn writeRecord(out: *std.Io.Writer, r: Record, envelope: Envelope) !void {
    switch (r) {
        .image => |img| try writeImage(out, img, envelope),
        .chart => |ch| try writeChart(out, ch, envelope),
    }
}

/// The unselected stream — every record, emission order. The future C
/// leg's entry point.
pub fn writeAll(out: *std.Io.Writer, view: *const Anchors) !void {
    for (view.records) |r| try writeRecord(out, r, .full);
}

pub fn writeImage(out: *std.Io.Writer, r: ImageRecord, envelope: Envelope) !void {
    try out.writeAll("{\"kind\":\"image_anchor\"");
    try writeCommon(out, r.sheet, r.sheet_idx, r.part, r.from, r.to, r.absolute, envelope);
    try out.print(",\"bytes\":{d}}}\n", .{r.byte_count});
}

pub fn writeChart(out: *std.Io.Writer, r: ChartRecord, envelope: Envelope) !void {
    try out.writeAll("{\"kind\":\"chart_anchor\"");
    try writeCommon(out, r.sheet, r.sheet_idx, r.part, r.from, r.to, r.absolute, envelope);
    try out.print(",\"chart_type\":\"{s}\",\"series_refs\":[", .{@tagName(r.chart_type)});
    for (r.series_refs, 0..) |ref, i| {
        if (i > 0) try out.writeByte(',');
        try json.writeString(out, ref);
    }
    try out.writeAll("]}\n");
}

fn writeCommon(
    out: *std.Io.Writer,
    sheet: []const u8,
    sheet_idx: u32,
    part: []const u8,
    from: ?drawings.CellAnchor,
    to: ?drawings.CellAnchor,
    absolute: ?drawings.AbsoluteAnchor,
    envelope: Envelope,
) !void {
    if (envelope == .full) {
        try out.writeAll(",\"sheet\":");
        try json.writeString(out, sheet);
        try out.print(",\"sheet_idx\":{d}", .{sheet_idx});
    }
    try out.writeAll(",\"part\":");
    try json.writeString(out, part);
    const anchor_kind: []const u8 = if (absolute != null)
        "absolute"
    else if (to != null)
        "two_cell"
    else
        "one_cell";
    try out.print(",\"anchor\":\"{s}\",\"from\":", .{anchor_kind});
    try writeOptCellAnchor(out, from);
    try out.writeAll(",\"to\":");
    try writeOptCellAnchor(out, to);
    try out.writeAll(",\"absolute\":");
    if (absolute) |abs| {
        try out.print(
            "{{\"x\":{d},\"y\":{d},\"cx\":{d},\"cy\":{d}}}",
            .{ abs.x, abs.y, abs.cx, abs.cy },
        );
    } else {
        try out.writeAll("null");
    }
}

fn writeOptCellAnchor(out: *std.Io.Writer, anchor: ?drawings.CellAnchor) !void {
    const c = anchor orelse {
        try out.writeAll("null");
        return;
    };
    // The drawing XML is 0-based on both axes; the wire is 1-based
    // like the `cells` / `merges` envelopes. Widen before the +1 so a
    // hostile 0xFFFFFFFF index cannot wrap.
    try out.print(
        "{{\"row\":{d},\"col\":{d},\"row_off\":{d},\"col_off\":{d}}}",
        .{ @as(u64, c.row) + 1, @as(u64, c.col) + 1, c.row_off, c.col_off },
    );
}

fn decodeSheetName(a: Allocator, raw: []const u8) Error![]u8 {
    const decoded = formula.decode.decodeAt(a, .sheet_name, raw) catch |e| switch (e) {
        error.OutOfMemory => return error.OutOfMemory,
        else => return error.MalformedWorkbookXml,
    };
    if (!std.unicode.utf8ValidateSlice(decoded)) return error.MalformedWorkbookXml;
    return decoded;
}

/// A `<c:f>` body is a formula carrier: entities resolve, everything
/// else passes through. A raw `<` in the element text can only open
/// markup (CDATA, a comment) the chart's own consumers do not read
/// through, and a ref that does not decode or is not UTF-8 would make
/// the stream lie or stop parsing — all three refuse.
fn decodeSeriesRef(a: Allocator, raw: []const u8) Error![]u8 {
    if (std.mem.indexOfScalar(u8, raw, '<') != null) return error.MalformedDrawingXml;
    const decoded = formula.decode.decodeAt(a, .cell_formula_body, raw) catch |e| switch (e) {
        error.OutOfMemory => return error.OutOfMemory,
        else => return error.MalformedDrawingXml,
    };
    if (!std.unicode.utf8ValidateSlice(decoded)) return error.MalformedDrawingXml;
    return decoded;
}

/// Part names come from the archive directory; the store admits any
/// bytes there, the JSON writer passes bytes through verbatim, so the
/// UTF-8 floor lands here. Duped so the view owns every string it
/// serves.
fn retainPartName(a: Allocator, name: []const u8) Error![]u8 {
    if (!std.unicode.utf8ValidateSlice(name)) return error.MalformedDrawingXml;
    return a.dupe(u8, name);
}

fn relById(rels: []const store_mod.Relationship, raw_rid: []const u8) ?store_mod.Relationship {
    var buf: [128]u8 = undefined;
    const rid = workbook_xml.decodeScalarAttr(&buf, raw_rid) orelse return null;
    if (rid.len == 0) return null;
    for (rels) |rel| {
        if (std.mem.eql(u8, rel.id, rid)) return rel;
    }
    return null;
}

// ─── Test fixture ────────────────────────────────────────────────────

/// Writes a real two-sheet workbook whose second sheet (`Report`,
/// index 1) carries a drawing: a twoCellAnchor image, a oneCellAnchor
/// chart with three series refs, and — for `.with_absolute` — an
/// absoluteAnchor image. Injected through a real save/reopen, the
/// `pivots.fixture` pattern, so the walk reads genuine parts and
/// refreshed relationship caches. `src/cli.zig` and the tests below
/// share it.
pub const fixture = struct {
    pub const Kind = enum { image_and_chart, with_absolute };

    pub const png_bytes = "\x89PNG\r\n\x1a\n01234567";

    pub fn write(allocator: Allocator, io: std.Io, path: []const u8, kind: Kind) !void {
        {
            const zlsx = @import("zlsx");
            var w = zlsx.Writer.init(allocator);
            defer w.deinit();
            var data = try w.addSheet("Data");
            try data.writeRow(&.{ .{ .string = "Region" }, .{ .string = "Qty" } });
            try data.writeRow(&.{ .{ .string = "East" }, .{ .integer = 3 } });
            try data.writeRow(&.{ .{ .string = "West" }, .{ .integer = 4 } });
            try data.writeRow(&.{ .{ .string = "East" }, .{ .integer = 5 } });
            var report = try w.addSheet("Report");
            try report.writeRow(&.{.{ .string = "drawing host" }});
            try w.save(io, path);
        }

        var store = try PartStore.open(allocator, io, path);
        defer store.deinit();

        try store.addPart("xl/media/image1.png", "image/png", png_bytes);
        try store.addPart(
            "xl/charts/chart1.xml",
            "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
            \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            \\<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Data!$B$1</c:f></c:strRef></c:tx><c:cat><c:strRef><c:f>Data!$A$2:$A$4</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>Data!$B$2:$B$4</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>
            ,
        );
        const absolute_block: []const u8 = switch (kind) {
            .image_and_chart => "",
            .with_absolute => "<xdr:absoluteAnchor><xdr:pos x=\"1000\" y=\"2000\"/><xdr:ext cx=\"914400\" cy=\"457200\"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"4\" name=\"Picture 2\"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed=\"rIdI1\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:absoluteAnchor>",
        };
        const drawing = try std.fmt.allocPrint(allocator,
            \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            \\<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>1</xdr:col><xdr:colOff>9525</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>7</xdr:row><xdr:rowOff>19050</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 1"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdI1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>{s}<xdr:oneCellAnchor><xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="3048000" cy="2286000"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="3" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdC1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>
        , .{absolute_block});
        defer allocator.free(drawing);
        try store.addPart(
            "xl/drawings/drawing1.xml",
            "application/vnd.openxmlformats-officedocument.drawing+xml",
            drawing,
        );
        try store.addPart(
            "xl/drawings/_rels/drawing1.xml.rels",
            "application/vnd.openxmlformats-package.relationships+xml",
            \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            \\<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdI1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/><Relationship Id="rIdC1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>
            ,
        );
        try upsertRels(&store, "xl/worksheets/_rels/sheet2.xml.rels",
            \\<Relationship Id="rIdD1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
        );
        try spliceBefore(allocator, &store, "xl/worksheets/sheet2.xml", "</worksheet>",
            \\<drawing r:id="rIdD1"/>
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

test "writeImage: two_cell full and compact, exact bytes" {
    const rec: ImageRecord = .{
        .sheet = "R\"D",
        .sheet_idx = 1,
        .part = "xl/media/image1.png",
        .from = .{ .col = 1, .col_off = 9525, .row = 2, .row_off = 0 },
        .to = .{ .col = 4, .col_off = 0, .row = 7, .row_off = 19050 },
        .absolute = null,
        .byte_count = 16,
    };
    var buf: [512]u8 = undefined;
    var w = std.Io.Writer.fixed(&buf);
    try writeImage(&w, rec, .full);
    try testing.expectEqualStrings(
        "{\"kind\":\"image_anchor\",\"sheet\":\"R\\\"D\",\"sheet_idx\":1,\"part\":\"xl/media/image1.png\"," ++
            "\"anchor\":\"two_cell\",\"from\":{\"row\":3,\"col\":2,\"row_off\":0,\"col_off\":9525}," ++
            "\"to\":{\"row\":8,\"col\":5,\"row_off\":19050,\"col_off\":0},\"absolute\":null,\"bytes\":16}\n",
        w.buffered(),
    );
    var w2 = std.Io.Writer.fixed(&buf);
    try writeImage(&w2, rec, .compact);
    try testing.expectEqualStrings(
        "{\"kind\":\"image_anchor\",\"part\":\"xl/media/image1.png\"," ++
            "\"anchor\":\"two_cell\",\"from\":{\"row\":3,\"col\":2,\"row_off\":0,\"col_off\":9525}," ++
            "\"to\":{\"row\":8,\"col\":5,\"row_off\":19050,\"col_off\":0},\"absolute\":null,\"bytes\":16}\n",
        w2.buffered(),
    );
}

test "writeImage: absolute anchor — from and to null, geometry verbatim EMUs" {
    const rec: ImageRecord = .{
        .sheet = "S",
        .sheet_idx = 0,
        .part = "xl/media/image2.png",
        .from = null,
        .to = null,
        .absolute = .{ .x = 1000, .y = 2000, .cx = 914400, .cy = 457200 },
        .byte_count = 5,
    };
    var buf: [512]u8 = undefined;
    var w = std.Io.Writer.fixed(&buf);
    try writeImage(&w, rec, .full);
    try testing.expectEqualStrings(
        "{\"kind\":\"image_anchor\",\"sheet\":\"S\",\"sheet_idx\":0,\"part\":\"xl/media/image2.png\"," ++
            "\"anchor\":\"absolute\",\"from\":null,\"to\":null," ++
            "\"absolute\":{\"x\":1000,\"y\":2000,\"cx\":914400,\"cy\":457200},\"bytes\":5}\n",
        w.buffered(),
    );
}

test "writeChart: one_cell with refs, and an empty refs list" {
    const rec: ChartRecord = .{
        .sheet = "Report",
        .sheet_idx = 1,
        .part = "xl/charts/chart1.xml",
        .from = .{ .col = 5, .col_off = 0, .row = 1, .row_off = 0 },
        .to = null,
        .absolute = null,
        .chart_type = .bar,
        .series_refs = &.{ "Data!$B$1", "'R&D'!$A$2:$A$4" },
    };
    var buf: [512]u8 = undefined;
    var w = std.Io.Writer.fixed(&buf);
    try writeChart(&w, rec, .full);
    try testing.expectEqualStrings(
        "{\"kind\":\"chart_anchor\",\"sheet\":\"Report\",\"sheet_idx\":1,\"part\":\"xl/charts/chart1.xml\"," ++
            "\"anchor\":\"one_cell\",\"from\":{\"row\":2,\"col\":6,\"row_off\":0,\"col_off\":0},\"to\":null," ++
            "\"absolute\":null,\"chart_type\":\"bar\",\"series_refs\":[\"Data!$B$1\",\"'R&D'!$A$2:$A$4\"]}\n",
        w.buffered(),
    );
    var empty = rec;
    empty.series_refs = &.{};
    empty.chart_type = .other;
    var w2 = std.Io.Writer.fixed(&buf);
    try writeChart(&w2, empty, .compact);
    try testing.expectEqualStrings(
        "{\"kind\":\"chart_anchor\",\"part\":\"xl/charts/chart1.xml\"," ++
            "\"anchor\":\"one_cell\",\"from\":{\"row\":2,\"col\":6,\"row_off\":0,\"col_off\":0},\"to\":null," ++
            "\"absolute\":null,\"chart_type\":\"other\",\"series_refs\":[]}\n",
        w2.buffered(),
    );
}

test "writeOptCellAnchor: a u32-max index widens instead of wrapping" {
    var buf: [128]u8 = undefined;
    var w = std.Io.Writer.fixed(&buf);
    try writeOptCellAnchor(&w, .{ .col = std.math.maxInt(u32), .col_off = 0, .row = std.math.maxInt(u32), .row_off = 0 });
    try testing.expectEqualStrings(
        "{\"row\":4294967296,\"col\":4294967296,\"row_off\":0,\"col_off\":0}",
        w.buffered(),
    );
}

test "collect: fixture anchors attributed to Report, images before charts, refs decoded" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "anchors_fixture.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .with_absolute);

    const workbook_mod = @import("workbook.zig");
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();

    var view = try collect(testing.allocator, &wb.store, &wb.workbook);
    defer view.deinit();

    try testing.expectEqual(@as(usize, 2), view.sheet_names.len);
    try testing.expectEqual(@as(usize, 3), view.records.len);
    // Images (document order: two_cell, then absolute) before charts.
    try testing.expectEqualStrings("xl/media/image1.png", view.records[0].image.part);
    try testing.expect(view.records[0].image.to != null);
    try testing.expect(view.records[1].image.absolute != null);
    try testing.expect(view.records[1].image.from == null);
    try testing.expectEqual(@as(u64, fixture.png_bytes.len), view.records[0].image.byte_count);
    const chart = view.records[2].chart;
    try testing.expectEqual(@as(u32, 1), chart.sheet_idx);
    try testing.expectEqualStrings("Report", chart.sheet);
    try testing.expectEqual(drawings.ChartType.bar, chart.chart_type);
    try testing.expectEqual(@as(usize, 3), chart.series_refs.len);
    try testing.expectEqualStrings("Data!$B$1", chart.series_refs[0]);
    for (view.records) |r| try testing.expectEqual(@as(u32, 1), r.sheetIdx());
}

test "collect: a workbook without drawings is an empty view" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "anchors_none.xlsx");
    defer testing.allocator.free(path);
    {
        const zlsx = @import("zlsx");
        var w = zlsx.Writer.init(testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Only");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, path);
    }
    const workbook_mod = @import("workbook.zig");
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    var view = try collect(testing.allocator, &wb.store, &wb.workbook);
    defer view.deinit();
    try testing.expectEqual(@as(usize, 0), view.records.len);
    try testing.expectEqual(@as(usize, 1), view.sheet_names.len);
}

test "collect: an anchor on a worksheet part xl/workbook.xml does not list refuses whole" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "anchors_orphan.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .image_and_chart);
    // Re-point the drawing at an orphan worksheet part: same bytes as
    // sheet2 (drawing reference and rels ride along under the copied
    // part name), but no <sheet> entry reaches it. The walkers still
    // key anchors by the part, and the view cannot attribute them.
    {
        var store = try PartStore.open(testing.allocator, io, path);
        defer store.deinit();
        const sheet2 = (try store.part("xl/worksheets/sheet2.xml")) orelse return error.PartNotFound;
        try store.addPart(
            "xl/worksheets/sheet9.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            sheet2.bytes,
        );
        const rels = (try store.part("xl/worksheets/_rels/sheet2.xml.rels")) orelse return error.PartNotFound;
        try store.addPart(
            "xl/worksheets/_rels/sheet9.xml.rels",
            "application/vnd.openxmlformats-package.relationships+xml",
            rels.bytes,
        );
        try store.save(io, path);
    }
    const workbook_mod = @import("workbook.zig");
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    try testing.expectError(
        error.DrawingOnUnlistedSheet,
        collect(testing.allocator, &wb.store, &wb.workbook),
    );
}
