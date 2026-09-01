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

/// One anchored image, attributed. `kind` is the anchor wrapper as
/// spelled in the drawing XML — the wire's `anchor` field; `from` /
/// `to` are the cell-grid anchors exactly as `drawings.ImageAnchor`
/// models them (0-based col/row + EMU offsets); `from` is null — not
/// the walker's zero sentinel — when the source used
/// `<xdr:absoluteAnchor>`.
pub const ImageRecord = struct {
    sheet: []const u8,
    sheet_idx: u32,
    /// Archive name of the image part, e.g. `xl/media/image1.png`.
    part: []const u8,
    kind: drawings.AnchorKind,
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
    kind: drawings.AnchorKind,
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
    /// A sheet-name carrier that does not decode or decodes to
    /// non-UTF-8 (the NDJSON must stay parseable), or a sheet list
    /// the anchors cannot be attributed against: a sheet whose
    /// relationship dangles, is mistyped or external, whose part the
    /// archive does not hold, or two `<sheet>` entries reaching one
    /// part (Codex #214 r1 REL-101/REL-102).
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

    // Resolve each workbook sheet to its part, strictly — the pivots
    // walk's rule: the relationship must exist under the sheet's
    // `r:id`, be a sheet-family type, be internal, and reach a part
    // the archive holds. A listed sheet the walk cannot place makes
    // the whole inventory unprovable (its drawing, had the part
    // existed, is unknown), so it refuses rather than emitting under
    // it (Codex #214 r1 REL-101). Two `<sheet>` entries reaching one
    // part would emit the same anchor twice under two identities —
    // refused too (REL-102).
    const wb_rels = store.rels("xl/workbook.xml");
    const sheet_parts = try a.alloc([]const u8, wb.sheets.len);
    for (wb.sheets, 0..) |s, i| {
        const rel = relById(wb_rels, s.r_id) orelse return error.MalformedWorkbookXml;
        var typed = false;
        for (sheet_rel_leaves) |leaf| typed = typed or relLeafIs(rel.type, leaf);
        if (!typed) return error.MalformedWorkbookXml;
        if (rel.target_mode == .external) return error.MalformedWorkbookXml;
        const name = (try store.resolve("xl/workbook.xml", rel.target)) orelse
            return error.MalformedWorkbookXml;
        if ((try store.part(name)) == null) return error.MalformedWorkbookXml;
        for (sheet_parts[0..i]) |prev| {
            if (std.mem.eql(u8, prev, name)) return error.MalformedWorkbookXml;
        }
        sheet_parts[i] = name;
    }

    // Strict walk: a drawing structure the walkers recognise but
    // cannot read whole refuses here rather than thinning the
    // inventory. The walkers allocate their result slices; handing
    // them the view arena parks that scaffolding with everything else
    // this view owns.
    const images = try drawings.imageAnchorsIn(store, a, .strict);
    const charts = try drawings.chartAnchorsIn(store, a, .strict);

    const image_seen = try a.alloc(bool, images.len);
    @memset(image_seen, false);
    const chart_seen = try a.alloc(bool, charts.len);
    @memset(chart_seen, false);

    var records: std.ArrayListUnmanaged(Record) = .empty;
    var scratch: std.ArrayListUnmanaged(usize) = .empty;
    for (0..wb.sheets.len) |idx| {
        const part_name = sheet_parts[idx];
        // Restore document order within the sheet: the mixed-prefix
        // replay appends alternate-prefixed anchors after the primary
        // scan, so walker slice order is not document order — the
        // anchors' doc_offset is (Codex #214 r1 REL-103).
        scratch.clearRetainingCapacity();
        for (images, 0..) |img, i| {
            if (std.mem.eql(u8, img.sheet_part_name, part_name)) try scratch.append(a, i);
        }
        std.sort.insertion(usize, scratch.items, images, imageOffsetLess);
        for (scratch.items) |i| {
            const img = images[i];
            image_seen[i] = true;
            try records.append(a, .{ .image = .{
                .sheet = sheet_names[idx],
                .sheet_idx = @intCast(idx),
                .part = try retainPartName(a, img.image_part_name),
                .kind = img.kind,
                .from = if (img.kind == .absolute) null else img.from,
                .to = img.to,
                .absolute = img.absolute,
                .byte_count = img.bytes.len,
            } });
        }
        scratch.clearRetainingCapacity();
        for (charts, 0..) |ch, i| {
            if (std.mem.eql(u8, ch.sheet_part_name, part_name)) try scratch.append(a, i);
        }
        std.sort.insertion(usize, scratch.items, charts, chartOffsetLess);
        for (scratch.items) |i| {
            const ch = charts[i];
            chart_seen[i] = true;
            const refs = try a.alloc([]const u8, ch.series_refs.len);
            for (ch.series_refs, 0..) |raw, ri| refs[ri] = try decodeSeriesRef(a, raw);
            try records.append(a, .{ .chart = .{
                .sheet = sheet_names[idx],
                .sheet_idx = @intCast(idx),
                .part = try retainPartName(a, ch.chart_part_name),
                .kind = ch.kind,
                .from = if (ch.kind == .absolute) null else ch.from,
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
    try writeCommon(out, r.sheet, r.sheet_idx, r.part, r.kind, r.from, r.to, r.absolute, envelope);
    try out.print(",\"bytes\":{d}}}\n", .{r.byte_count});
}

pub fn writeChart(out: *std.Io.Writer, r: ChartRecord, envelope: Envelope) !void {
    try out.writeAll("{\"kind\":\"chart_anchor\"");
    try writeCommon(out, r.sheet, r.sheet_idx, r.part, r.kind, r.from, r.to, r.absolute, envelope);
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
    kind: drawings.AnchorKind,
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
    // `anchor` is the wrapper as spelled in the source, never derived
    // from which optional fields happen to be present — the strict
    // walk refuses a two-cell anchor whose `<to>` did not parse, so
    // the spelling and the fields cannot disagree here (REL-101).
    try out.print(",\"anchor\":\"{s}\",\"from\":", .{@tagName(kind)});
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

/// The sheet-family relationship types a `<sheet r:id>` may carry —
/// the pivots walk's list (`pkg/pivots.zig`), duplicated here because
/// this module keeps its imports to the walkers it fronts.
const sheet_rel_leaves = [_][]const u8{ "worksheet", "chartsheet", "dialogsheet", "xlMacrosheet", "xlIntlMacrosheet" };

fn relLeafIs(rel_type: []const u8, leaf: []const u8) bool {
    const l = if (std.mem.lastIndexOfScalar(u8, rel_type, '/')) |i| rel_type[i + 1 ..] else rel_type;
    return std.ascii.eqlIgnoreCase(l, leaf);
}

fn imageOffsetLess(images: []const drawings.ImageAnchor, lhs: usize, rhs: usize) bool {
    return images[lhs].doc_offset < images[rhs].doc_offset;
}

fn chartOffsetLess(charts: []const drawings.ChartAnchor, lhs: usize, rhs: usize) bool {
    return charts[lhs].doc_offset < charts[rhs].doc_offset;
}

// ─── Test fixture ────────────────────────────────────────────────────

/// Writes a real two-sheet workbook with anchors on BOTH sheets:
/// `Data` (index 0) carries one twoCellAnchor image; `Report` (index
/// 1) carries a drawing whose DOCUMENT order is chart first — a
/// oneCellAnchor chart with three series refs, then a twoCellAnchor
/// image, then (for `.with_absolute`) an absoluteAnchor image — so
/// the view's images-before-charts regrouping is exercised, not
/// mirrored (Codex #214 r1 MNT-101). Injected through a real
/// save/reopen, the `pivots.fixture` pattern, so the walk reads
/// genuine parts and refreshed relationship caches. `src/cli.zig`
/// and the tests below share it.
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
        // Report's drawing: chart FIRST in document order.
        const drawing = try std.fmt.allocPrint(allocator,
            \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            \\<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:oneCellAnchor><xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="3048000" cy="2286000"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="3" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdC1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor><xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>1</xdr:col><xdr:colOff>9525</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>7</xdr:row><xdr:rowOff>19050</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 1"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdI1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>{s}</xdr:wsDr>
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

        // Data's drawing: one twoCellAnchor image, so the stream
        // crosses a sheet boundary.
        try store.addPart(
            "xl/drawings/drawing2.xml",
            "application/vnd.openxmlformats-officedocument.drawing+xml",
            \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            \\<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Logo"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdI1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>
            ,
        );
        try store.addPart(
            "xl/drawings/_rels/drawing2.xml.rels",
            "application/vnd.openxmlformats-package.relationships+xml",
            \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            \\<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdI1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>
            ,
        );
        try upsertRels(&store, "xl/worksheets/_rels/sheet1.xml.rels",
            \\<Relationship Id="rIdD2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing2.xml"/>
        );
        try spliceBefore(allocator, &store, "xl/worksheets/sheet1.xml", "</worksheet>",
            \\<drawing r:id="rIdD2"/>
        );

        try store.save(io, path);
    }

    /// Byte-replace the first `old` in one part of a saved workbook
    /// and save it back — how the refusal tests make a fixture wrong
    /// in exactly one place (the `pivots.fixture` helper, duplicated
    /// so this module keeps its imports to the walkers it fronts).
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

test "writeImage: two_cell full and compact, exact bytes" {
    const rec: ImageRecord = .{
        .sheet = "R\"D",
        .sheet_idx = 1,
        .part = "xl/media/image1.png",
        .kind = .two_cell,
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
        .kind = .absolute,
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
        .kind = .one_cell,
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
    try testing.expectEqual(@as(usize, 4), view.records.len);
    // Data's image leads (sheets in workbook order) …
    try testing.expectEqual(@as(u32, 0), view.records[0].sheetIdx());
    try testing.expectEqualStrings("Data", view.records[0].image.sheet);
    try testing.expectEqual(drawings.AnchorKind.two_cell, view.records[0].image.kind);
    // … then Report's images BEFORE its chart, though the chart is
    // FIRST in the drawing's document order — the regrouping, not an
    // echo of the source (MNT-101).
    try testing.expectEqual(drawings.AnchorKind.two_cell, view.records[1].image.kind);
    try testing.expect(view.records[1].image.to != null);
    try testing.expectEqual(drawings.AnchorKind.absolute, view.records[2].image.kind);
    try testing.expect(view.records[2].image.absolute != null);
    try testing.expect(view.records[2].image.from == null);
    try testing.expectEqual(@as(u64, fixture.png_bytes.len), view.records[1].image.byte_count);
    const chart = view.records[3].chart;
    try testing.expectEqual(@as(u32, 1), chart.sheet_idx);
    try testing.expectEqualStrings("Report", chart.sheet);
    try testing.expectEqual(drawings.AnchorKind.one_cell, chart.kind);
    try testing.expectEqual(drawings.ChartType.bar, chart.chart_type);
    try testing.expectEqual(@as(usize, 3), chart.series_refs.len);
    try testing.expectEqualStrings("Data!$B$1", chart.series_refs[0]);
}

test "collect: mixed-prefix anchors come back in document order (REL-103)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "anchors_mixed_prefix.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path, .image_and_chart);
    // Replace Report's drawing with one that binds TWO prefixes to the
    // spreadsheetDrawing URI and puts the alternate-prefixed anchor
    // FIRST. The walkers scan per prefix and replay alternates after
    // the primary pass, so slice order alone would reverse these.
    {
        var store = try PartStore.open(testing.allocator, io, path);
        defer store.deinit();
        try store.replacePart("xl/drawings/drawing1.xml",
            \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            \\<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:x2="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><x2:oneCellAnchor><x2:from><x2:col>7</x2:col><x2:colOff>0</x2:colOff><x2:row>7</x2:row><x2:rowOff>0</x2:rowOff></x2:from><x2:ext cx="1" cy="1"/><x2:pic><x2:nvPicPr><x2:cNvPr id="5" name="P"/><x2:cNvPicPr/></x2:nvPicPr><x2:blipFill><a:blip r:embed="rIdI1"/></x2:blipFill><x2:spPr/></x2:pic><x2:clientData/></x2:oneCellAnchor><xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="1" cy="1"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="6" name="Q"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdI1"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>
        );
        try store.save(io, path);
    }
    const workbook_mod = @import("workbook.zig");
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    var view = try collect(testing.allocator, &wb.store, &wb.workbook);
    defer view.deinit();
    // Data's image, then Report's two: the x2-prefixed anchor (col 7)
    // FIRST — document order restored via doc_offset.
    try testing.expectEqual(@as(usize, 3), view.records.len);
    try testing.expectEqual(@as(u32, 7), view.records[1].image.from.?.col);
    try testing.expectEqual(@as(u32, 1), view.records[2].image.from.?.col);
}

test "collect: a drawing graph edge the walk cannot follow refuses whole (REL-101)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const workbook_mod = @import("workbook.zig");

    const cases = [_]struct { name: []const u8, part: []const u8, old: []const u8, new: []const u8 }{
        // The sheet's drawing relationship dangles.
        .{ .name = "a1.xlsx", .part = "xl/worksheets/_rels/sheet2.xml.rels", .old = "../drawings/drawing1.xml", .new = "../drawings/nope.xml" },
        // A pic's blip relationship names nothing.
        .{ .name = "a2.xlsx", .part = "xl/drawings/drawing1.xml", .old = "r:embed=\"rIdI1\"", .new = "r:embed=\"rIdXX\"" },
        // A two-cell anchor whose <to> does not parse must refuse, not
        // ride out as one_cell.
        .{ .name = "a3.xlsx", .part = "xl/drawings/drawing1.xml", .old = "<xdr:to>", .new = "<xdr:zz>" },
    };
    for (cases) |case| {
        const path = try tt.path(testing.allocator, io, case.name);
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path, .image_and_chart);
        try fixture.patchPart(testing.allocator, io, path, case.part, case.old, case.new);
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        try testing.expectError(
            error.MalformedDrawingXml,
            collect(testing.allocator, &wb.store, &wb.workbook),
        );
    }
}

test "collect: a sheet list the anchors cannot be attributed against refuses whole (REL-101/102)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const workbook_mod = @import("workbook.zig");

    const cases = [_]struct { name: []const u8, old: []const u8, new: []const u8 }{
        // A listed sheet whose part is absent: the inventory under it
        // is unprovable.
        .{ .name = "b1.xlsx", .old = "worksheets/sheet2.xml", .new = "worksheets/nope.xml" },
        // Two <sheet> entries reaching one part: the same anchor would
        // ride out twice under two identities.
        .{ .name = "b2.xlsx", .old = "worksheets/sheet1.xml", .new = "worksheets/sheet2.xml" },
    };
    for (cases) |case| {
        const path = try tt.path(testing.allocator, io, case.name);
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path, .image_and_chart);
        try fixture.patchPart(testing.allocator, io, path, "xl/_rels/workbook.xml.rels", case.old, case.new);
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        try testing.expectError(
            error.MalformedWorkbookXml,
            collect(testing.allocator, &wb.store, &wb.workbook),
        );
    }
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
