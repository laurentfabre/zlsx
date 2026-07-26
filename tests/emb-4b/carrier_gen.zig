//! emb-4B carrier fixture generator.
//!
//! Writes ONE .xlsx carrying the same recovery marker in six different
//! places (see `carriers.zig` for the catalogue and the reasoning), so
//! a single open → save round-trip through a consumer measures every
//! carrier at once.
//!
//! Usage:
//!   zig build emb4b-fixture -- <out.xlsx>
//!
//! Built in two passes. Pass 1 uses the ordinary typed writer surface
//! (sheets, cells, defined names, `setEmbeddings`) and saves. Pass 2
//! reopens and injects the raw parts through `PartStore`. The split is
//! deliberate: `Workbook.save` splices staged plans (defined names,
//! SST, per-sheet deltas) into the existing part bytes, so injecting
//! raw XML in the same pass would race the splice for `xl/workbook.xml`.
//! On the reopen there are no staged plans, and save is a pure repack.

const std = @import("std");
const pkg = @import("zlsx_pkg");
const zlsx = @import("zlsx");
const carriers = @import("emb4b_carriers");

const Cell = zlsx.Cell;
const Carrier = carriers.Carrier;

pub fn main(init: std.process.Init) !u8 {
    const allocator = init.gpa;
    const io = init.io;

    const args = try init.minimal.args.toSlice(init.arena.allocator());

    var err_buf: [256]u8 = undefined;
    var err_w = std.Io.File.stderr().writer(io, &err_buf);
    const err = &err_w.interface;
    defer err.flush() catch {};

    if (args.len < 2) {
        try err.print("usage: {s} <out.xlsx>\n", .{args[0]});
        return 2;
    }
    const out_path = args[1];

    // Pass 1 lands here; pass 2 reads it back and writes `out_path`.
    const stage_path = try std.fmt.allocPrint(allocator, "{s}.stage", .{out_path});
    defer allocator.free(stage_path);
    defer std.Io.Dir.cwd().deleteFile(io, stage_path) catch {};

    try passOne(allocator, io, stage_path);
    try passTwo(allocator, io, stage_path, out_path);

    var stdout_buf: [1024]u8 = undefined;
    var stdout_w = std.Io.File.stdout().writer(io, &stdout_buf);
    const out = &stdout_w.interface;
    defer out.flush() catch {};
    try out.print("wrote emb-4B carrier fixture: {s}\n", .{out_path});
    for (carriers.ALL) |c| {
        var mbuf: [carriers.MARKER_MAX]u8 = undefined;
        try out.print("  {s:<12} {s}\n", .{ carriers.marker(&mbuf, c), c.location() });
    }
    return 0;
}

/// Typed-writer carriers: cell data, defined name, and the emb-1a
/// custom OPC part via `setEmbeddings`.
fn passOne(allocator: std.mem.Allocator, io: std.Io, stage_path: []const u8) !void {
    var wb = try pkg.Workbook.empty(allocator, io);
    defer wb.deinit();

    // A plausible data sheet, so consumers see a normal workbook and
    // take their normal save path rather than an empty-file shortcut.
    const items = try wb.addSheet("Items");
    try items.appendRows(&[_][]const Cell{
        &.{ .{ .string = "Title" }, .{ .string = "Body" } },
        &.{ .{ .string = "Alpha" }, .{ .string = "First entry, body text alpha." } },
        &.{ .{ .string = "Beta" }, .{ .string = "Second entry, body text beta." } },
        &.{ .{ .string = "Gamma" }, .{ .string = "Third entry, body text gamma." } },
        &.{ .{ .string = "Delta" }, .{ .string = "Fourth entry, body text delta." } },
    });

    // Carrier: cell_data. Its own sheet so the marker cannot be
    // confused with the data grid, and so pass 2 can hide that sheet
    // by name.
    var cell_marker: [carriers.MARKER_MAX]u8 = undefined;
    const cell_m = carriers.marker(&cell_marker, .cell_data);
    const rec = try wb.addSheet(carriers.CELL_SHEET);
    try rec.appendRows(&[_][]const Cell{&.{.{ .string = cell_m }}});

    // Carrier: defined_name. A string-literal formula — the only
    // defined-name shape that stores opaque payload rather than a
    // reference that a consumer would try to fix up on load.
    var dn_marker: [carriers.MARKER_MAX]u8 = undefined;
    const dn_m = carriers.marker(&dn_marker, .defined_name);
    const dn_formula = try std.fmt.allocPrint(allocator, "\"{s}\"", .{dn_m});
    defer allocator.free(dn_formula);
    try wb.addDefinedName(carriers.DEFINED_NAME, dn_formula, .{});

    // Carrier: opc_part — the emb-4 control. The marker rides in the
    // model name, which `Workbook.embeddings()` reads back verbatim.
    var opc_marker: [carriers.MARKER_MAX]u8 = undefined;
    const opc_m = carriers.marker(&opc_marker, .opc_part);
    const dim: u32 = 4;
    var vec_body: [4 * (4 + 4)]u8 = undefined;
    encodeInt8SymRows(&vec_body, 0.5, [_]i8{ 10, -10, 5, -5 });
    const hashes = [_]u64{ 0xA000_0001, 0xA000_0002, 0xA000_0003, 0xA000_0004 };
    try wb.setEmbeddings(opc_m, dim, .int8_sym_per_vec, &[_]pkg.EmbeddingCoverageInput{.{
        .id = "title",
        .worksheet_target = "worksheets/sheet1.xml",
        .range = "A2:A5",
        .column = "A",
        .include_formulas = false,
        .vec_body = &vec_body,
        .hashes = &hashes,
    }});

    try wb.save(io, stage_path);
}

/// Raw-part carriers: customXml, docProps/custom.xml, the workbook
/// `<extLst>`, and the `state="hidden"` flag on the marker sheet.
fn passTwo(
    allocator: std.mem.Allocator,
    io: std.Io,
    stage_path: []const u8,
    out_path: []const u8,
) !void {
    var wb = try pkg.Workbook.open(allocator, io, stage_path);
    defer wb.deinit();

    var mbuf: [carriers.MARKER_MAX]u8 = undefined;

    // ---- Carrier: custom_xml ------------------------------------
    // The realistic four-piece shape, not a bare part: item + props +
    // item rels + a workbook rel. A consumer that prunes on rel
    // reachability must see the same graph Office would emit,
    // otherwise we would be measuring a malformed carrier.
    const cx_m = carriers.marker(&mbuf, .custom_xml);
    const cx_item = try std.fmt.allocPrint(allocator,
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<zlsxRecovery xmlns="http://schemas.laurentfabre.dev/zlsx/2026/recovery"><record>{s}</record></zlsxRecovery>
    , .{cx_m});
    defer allocator.free(cx_item);
    try wb.store.addPart("customXml/item1.xml", "application/xml", cx_item);

    try wb.store.addPart(
        "customXml/itemProps1.xml",
        "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="{7A1C4E90-6B2D-4F31-9E5A-2C1D8F3A0002}"><ds:schemaRefs><ds:schemaRef ds:uri="http://schemas.laurentfabre.dev/zlsx/2026/recovery"/></ds:schemaRefs></ds:datastoreItem>
        ,
    );

    try wb.store.addPart(
        "customXml/_rels/item1.xml.rels",
        "application/vnd.openxmlformats-package.relationships+xml",
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>
        ,
    );

    // Target is `../customXml/item1.xml`: relationship targets resolve
    // against the source part's directory, which for
    // xl/_rels/workbook.xml.rels is `xl/`.
    try spliceRels(
        allocator,
        &wb,
        "xl/_rels/workbook.xml.rels",
        \\<Relationship Id="rIdZlsxE4BCustomXml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>
        ,
    );

    // ---- Carrier: doc_props -------------------------------------
    // fmtid is the fixed OOXML custom-properties GUID; pid must start
    // at 2 (0 and 1 are reserved by the spec).
    const dp_m = carriers.marker(&mbuf, .doc_props);
    const dp = try std.fmt.allocPrint(allocator,
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}" pid="2" name="ZlsxE4BRecovery"><vt:lpwstr>{s}</vt:lpwstr></property></Properties>
    , .{dp_m});
    defer allocator.free(dp);
    // May already exist if a future `empty()` starts emitting it.
    if (try wb.store.part("docProps/custom.xml") != null) {
        try wb.store.replacePart("docProps/custom.xml", dp);
    } else {
        try wb.store.addPart(
            "docProps/custom.xml",
            "application/vnd.openxmlformats-officedocument.custom-properties+xml",
            dp,
        );
        try spliceRels(
            allocator,
            &wb,
            "_rels/.rels",
            \\<Relationship Id="rIdZlsxE4BCustomProps" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>
            ,
        );
    }

    // ---- Carriers: ext_lst + the hidden flag on the marker sheet --
    // Both live in xl/workbook.xml, so patch it once.
    const wbx = (try wb.store.part("xl/workbook.xml")) orelse return error.MissingWorkbookPart;
    const ext_m = carriers.marker(&mbuf, .ext_lst);
    const ext = try std.fmt.allocPrint(allocator,
        \\<extLst><ext uri="{s}" xmlns:zlsxe4b="http://schemas.laurentfabre.dev/zlsx/2026/recovery"><zlsxe4b:recovery>{s}</zlsxe4b:recovery></ext></extLst>
    , .{ carriers.EXT_URI, ext_m });
    defer allocator.free(ext);

    var patched = try hideSheet(allocator, wbx.bytes, carriers.CELL_SHEET);
    defer allocator.free(patched);

    // `<extLst>` is the last child of `<workbook>` in the CT_Workbook
    // sequence, so it goes immediately before the closing tag.
    const close = "</workbook>";
    const at = std.mem.lastIndexOf(u8, patched, close) orelse return error.MalformedWorkbookXml;
    const with_ext = try std.fmt.allocPrint(allocator, "{s}{s}{s}", .{
        patched[0..at],
        ext,
        patched[at..],
    });
    defer allocator.free(with_ext);
    try wb.store.replacePart("xl/workbook.xml", with_ext);

    try wb.save(io, out_path);
}

/// Insert `rel` immediately before `</Relationships>` in the named
/// rels part.
fn spliceRels(
    allocator: std.mem.Allocator,
    wb: *pkg.Workbook,
    part_name: []const u8,
    rel: []const u8,
) !void {
    const p = (try wb.store.part(part_name)) orelse return error.MissingRelsPart;
    const close = "</Relationships>";
    const at = std.mem.lastIndexOf(u8, p.bytes, close) orelse return error.MalformedRels;
    const out = try std.fmt.allocPrint(allocator, "{s}{s}{s}", .{
        p.bytes[0..at],
        rel,
        p.bytes[at..],
    });
    defer allocator.free(out);
    try wb.store.replacePart(part_name, out);
}

/// Add `state="hidden"` to the `<sheet>` element named `name`.
///
/// Returns an owned copy either way: if the sheet already carries a
/// `state` attribute the input is passed through unchanged, so callers
/// get uniform ownership.
fn hideSheet(
    allocator: std.mem.Allocator,
    workbook_xml: []const u8,
    name: []const u8,
) ![]u8 {
    const needle = try std.fmt.allocPrint(allocator, "<sheet name=\"{s}\"", .{name});
    defer allocator.free(needle);

    const at = std.mem.indexOf(u8, workbook_xml, needle) orelse
        return allocator.dupe(u8, workbook_xml);

    // End of this `<sheet .../>` element.
    const rest = workbook_xml[at..];
    const end_rel = std.mem.indexOfScalar(u8, rest, '>') orelse return error.MalformedWorkbookXml;
    const elem = rest[0..end_rel];
    if (std.mem.indexOf(u8, elem, "state=") != null) return allocator.dupe(u8, workbook_xml);

    // Insert before the self-closing slash if there is one, so the
    // result stays well-formed for both `<sheet .../>` and `<sheet ...>`.
    var insert_at = at + end_rel;
    if (insert_at > 0 and workbook_xml[insert_at - 1] == '/') insert_at -= 1;

    return std.fmt.allocPrint(allocator, "{s} state=\"hidden\"{s}", .{
        workbook_xml[0..insert_at],
        workbook_xml[insert_at..],
    });
}

/// Same encoding as the emb-4 fixture: little-endian f32 scale
/// followed by `dim` i8 values, repeated per row.
fn encodeInt8SymRows(out: *[4 * (4 + 4)]u8, scale: f32, q: [4]i8) void {
    var i: usize = 0;
    while (i < 4) : (i += 1) {
        const off = i * (4 + 4);
        std.mem.writeInt(u32, out[off..][0..4], @bitCast(scale), .little);
        out[off + 4] = @bitCast(q[0]);
        out[off + 5] = @bitCast(q[1]);
        out[off + 6] = @bitCast(q[2]);
        out[off + 7] = @bitCast(q[3]);
    }
}

test "hideSheet adds state on a self-closing element" {
    const a = std.testing.allocator;
    const src =
        \\<workbook><sheets><sheet name="Items" sheetId="1" r:id="rId1"/><sheet name="zlsxE4B" sheetId="2" r:id="rId2"/></sheets></workbook>
    ;
    const got = try hideSheet(a, src, "zlsxE4B");
    defer a.free(got);
    try std.testing.expect(std.mem.indexOf(u8, got, "<sheet name=\"zlsxE4B\" sheetId=\"2\" r:id=\"rId2\" state=\"hidden\"/>") != null);
    // The other sheet is untouched.
    try std.testing.expect(std.mem.indexOf(u8, got, "<sheet name=\"Items\" sheetId=\"1\" r:id=\"rId1\"/>") != null);
}

test "hideSheet is idempotent and tolerates a missing sheet" {
    const a = std.testing.allocator;
    const already =
        \\<workbook><sheets><sheet name="zlsxE4B" sheetId="1" state="hidden" r:id="rId1"/></sheets></workbook>
    ;
    const got = try hideSheet(a, already, "zlsxE4B");
    defer a.free(got);
    try std.testing.expectEqualStrings(already, got);

    const absent = "<workbook><sheets/></workbook>";
    const got2 = try hideSheet(a, absent, "zlsxE4B");
    defer a.free(got2);
    try std.testing.expectEqualStrings(absent, got2);
}
