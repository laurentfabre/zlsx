//! Fresh-archive emit substrate (B3 iter-wr-7).
//!
//! Lifts the entire Writer.save archive orchestration into a std-only
//! module both `xlsx.Writer.save` and `pkg.Workbook.saveFreshEmit`
//! consume. Inputs are the per-subsystem plans + per-sheet emit state;
//! output is the full OOXML archive byte image (Content_Types.xml,
//! _rels/.rels, xl/workbook.xml, xl/_rels/workbook.xml.rels, per-sheet
//! sheet/rels/comments/vml, xl/sharedStrings.xml, optional xl/styles.xml,
//! ZIP CD + EOCD).
//!
//! Byte-stability: this module is the canonical home for every
//! catalogued §2 invariant in `docs/plans/writer-rebase.md`. The
//! pre-iter-wr-7 emit branch in `src/writer.zig::save` is preserved
//! verbatim — only its location moves.
//!
//! Stdlib only. Zig 0.15.2.

const std = @import("std");

const Allocator = std.mem.Allocator;

const sst_plan_mod = @import("zlsx_sst_plan");
const styles_plan_mod = @import("zlsx_styles_plan");
const workbook_xml_plan_mod = @import("zlsx_workbook_xml_plan");
const sheet_plan = @import("zlsx_sheet_plan");
const zip = @import("zlsx_zip");

pub const Error = error{
    NoSheets,
    OutOfMemory,
    EntryTooLarge,
    NameTooLong,
    TooManyZipEntries,
    ZipArchiveTooLarge,
};

// ─── OOXML skeleton blobs ────────────────────────────────────────────

const CONTENT_TYPES_DEFAULTS: []const u8 =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    \\<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    \\<Default Extension="xml" ContentType="application/xml"/>
;
const CONTENT_TYPES_FIXED_OVERRIDES: []const u8 =
    \\<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    \\<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
;
const CONTENT_TYPES_TAIL: []const u8 = "</Types>";

const ROOT_RELS: []const u8 =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    \\<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
    \\</Relationships>
;

const WORKBOOK_RELS_HEAD: []const u8 =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
;
const WORKBOOK_RELS_TAIL: []const u8 = "</Relationships>";

const SST_HEAD_FMT: []const u8 =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{d}" uniqueCount="{d}">
;
const SST_TAIL: []const u8 = "</sst>";

// ─── Inputs ──────────────────────────────────────────────────────────

/// One sheet's emit inputs. `name` is the sheet's display name (used
/// in workbook.xml and CT.xml ordering — NOT the part name). `body` is
/// the pre-built `<row>...</row>` payload, SST indices baked in. State
/// owns the per-sheet registries.
pub const SheetInput = struct {
    name: []const u8,
    body: []const u8,
    state: *const sheet_plan.SheetState,
};

/// All inputs for `emitArchiveBytes`.
pub const ArchiveInputs = struct {
    sheets: []const SheetInput,
    sst_plan: *const sst_plan_mod.SstExtensionPlan,
    /// Total number of string-typed cells written (the OOXML
    /// `<sst count=>` attribute). Distinct from
    /// `sst_plan.new_strings.items.len` which is uniqueCount.
    sst_count: u64,
    styles_plan: *const styles_plan_mod.StylesPlan,
    workbook_xml_plan: *const workbook_xml_plan_mod.WorkbookXmlPlan,
    /// §9's `max_output_archive_bytes`. Carried on the inputs rather
    /// than passed alongside them so **every** caller of the substrate
    /// observes the same bound — the path save and the buffer save are
    /// one emitter, and a limit only one of them consulted would make
    /// "identical typed outcome at every layer" a coincidence.
    max_archive_bytes: u64 = max_output_archive_bytes,
};

/// §9's cap on a serialized output archive: 2³²−1 bytes exactly,
/// matching the ZIP32 sentinel bounds `pkg/zip.zig` enforces and
/// `PartStore.save` preflights. Re-exported here because this is the
/// substrate a producer reaches it through.
pub const max_output_archive_bytes: u64 = zip.default_max_archive_bytes;

// ─── deflate hookup (broken into a function pointer to keep this
// module stdlib-only — see `pkg/zip.zig` for the same pattern) ──────

pub const DeflateFn = zip.DeflateFn;

/// M5d1's poll seam, re-exported for the same reason `DeflateFn` is: a
/// producer reaches the archive substrate through this module and should
/// not have to name a second one to pass a control through.
pub const Poller = zip.Poller;

// ─── Public API ──────────────────────────────────────────────────────

/// Build a complete OOXML archive byte image into `out`. Caller owns
/// `out` and its lifetime. Mirrors the pre-iter-wr-7 `Writer.save`
/// archive orchestration: byte format, attribute orders, optional-block
/// elision, rId numbering, and content-types layering are all locked
/// by the existing parity test set in `src/writer.zig` and
/// `pkg/workbook.zig`.
pub fn emitArchiveBytes(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    inputs: ArchiveInputs,
    deflate: DeflateFn,
    poller: Poller,
) !void {
    if (inputs.sheets.len == 0) return error.NoSheets;

    var arc = zip.Archive.initControlled(allocator, out, inputs.max_archive_bytes, poller);
    defer arc.deinit();

    const have_styles = !inputs.styles_plan.isEmpty();

    // 1. [Content_Types].xml
    {
        var ct: std.ArrayListUnmanaged(u8) = .empty;
        defer ct.deinit(allocator);
        try ct.appendSlice(allocator, CONTENT_TYPES_DEFAULTS);
        var any_comments = false;
        for (inputs.sheets) |s| {
            if (s.state.comments.items.len > 0) {
                any_comments = true;
                break;
            }
        }
        if (any_comments) {
            try ct.appendSlice(
                allocator,
                "<Default Extension=\"vml\" ContentType=\"application/vnd.openxmlformats-officedocument.vmlDrawing\"/>",
            );
        }
        try ct.appendSlice(allocator, CONTENT_TYPES_FIXED_OVERRIDES);
        for (inputs.sheets, 0..) |_, i| {
            try ct.print(
                allocator,
                "<Override PartName=\"/xl/worksheets/sheet{d}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>",
                .{i + 1},
            );
        }
        if (have_styles) {
            try ct.appendSlice(
                allocator,
                "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>",
            );
        }
        if (any_comments) {
            for (inputs.sheets, 0..) |s, i| {
                if (s.state.comments.items.len == 0) continue;
                try ct.print(
                    allocator,
                    "<Override PartName=\"/xl/comments{d}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml\"/>",
                    .{i + 1},
                );
            }
        }
        try ct.appendSlice(allocator, CONTENT_TYPES_TAIL);
        try arc.addEntry("[Content_Types].xml", ct.items, deflate);
    }

    // 2. _rels/.rels (static)
    try arc.addEntry("_rels/.rels", ROOT_RELS, deflate);

    // 3. xl/workbook.xml — via the workbook_xml plan.
    {
        const sheet_entries = try allocator.alloc(workbook_xml_plan_mod.SheetEntry, inputs.sheets.len);
        defer allocator.free(sheet_entries);
        for (inputs.sheets, 0..) |s, i| {
            sheet_entries[i] = .{ .name = s.name };
        }
        const wb_bytes = try workbook_xml_plan_mod.emitWorkbookXml(
            allocator,
            sheet_entries,
            inputs.workbook_xml_plan,
        );
        defer allocator.free(wb_bytes);
        try arc.addEntry("xl/workbook.xml", wb_bytes, deflate);
    }

    // 4. xl/_rels/workbook.xml.rels
    {
        var rels: std.ArrayListUnmanaged(u8) = .empty;
        defer rels.deinit(allocator);
        try rels.appendSlice(allocator, WORKBOOK_RELS_HEAD);
        for (inputs.sheets, 0..) |_, i| {
            try rels.print(
                allocator,
                "<Relationship Id=\"rId{d}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{d}.xml\"/>",
                .{ i + 1, i + 1 },
            );
        }
        try rels.print(
            allocator,
            "<Relationship Id=\"rId{d}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>",
            .{inputs.sheets.len + 1},
        );
        if (have_styles) {
            try rels.print(
                allocator,
                "<Relationship Id=\"rId{d}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>",
                .{inputs.sheets.len + 2},
            );
        }
        try rels.appendSlice(allocator, WORKBOOK_RELS_TAIL);
        try arc.addEntry("xl/_rels/workbook.xml.rels", rels.items, deflate);
    }

    // 5. xl/worksheets/sheetN.xml — per-sheet via sheet_plan emitter.
    for (inputs.sheets, 0..) |s, i| {
        var full: std.ArrayListUnmanaged(u8) = .empty;
        defer full.deinit(allocator);

        const st = s.state;

        // Project state's owned-slice registries onto plan const-slice views.
        const hyperlinks_view = try allocator.alloc(sheet_plan.Hyperlink, st.hyperlinks.items.len);
        defer allocator.free(hyperlinks_view);
        for (st.hyperlinks.items, 0..) |h, k| {
            hyperlinks_view[k] = .{ .range = h.range, .url = h.url };
        }

        const internal_hl_view = try allocator.alloc(sheet_plan.InternalHyperlink, st.internal_hyperlinks.items.len);
        defer allocator.free(internal_hl_view);
        for (st.internal_hyperlinks.items, 0..) |h, k| {
            internal_hl_view[k] = .{ .range = h.range, .location = h.location };
        }

        const merges_view = try allocator.alloc([]const u8, st.merged_cells.items.len);
        defer allocator.free(merges_view);
        for (st.merged_cells.items, 0..) |range, k| {
            merges_view[k] = range;
        }

        const cw_view = try allocator.alloc(sheet_plan.ColumnWidth, st.column_widths.items.len);
        defer allocator.free(cw_view);
        for (st.column_widths.items, 0..) |cw, k| {
            cw_view[k] = .{ .col_min = cw.col_min, .col_max = cw.col_max, .width = cw.width };
        }

        const cf_view = try allocator.alloc(sheet_plan.ConditionalFormat, st.conditional_formats.items.len);
        defer allocator.free(cf_view);
        for (st.conditional_formats.items, 0..) |cf, k| {
            cf_view[k] = .{
                .range = cf.range,
                .rule = switch (cf.rule) {
                    .cell_is => |r| .{ .cell_is = .{
                        .operator = r.operator,
                        .formula1 = r.formula1,
                        .formula2 = r.formula2,
                        .dxf_id = r.dxf_id,
                    } },
                    .expression => |r| .{ .expression = .{
                        .formula = r.formula,
                        .dxf_id = r.dxf_id,
                    } },
                    .color_scale => |r| .{ .color_scale = .{
                        .low_color_argb = r.low_color_argb,
                        .mid_color_argb = r.mid_color_argb,
                        .high_color_argb = r.high_color_argb,
                    } },
                    .data_bar => |r| .{ .data_bar = .{ .color_argb = r.color_argb } },
                },
            };
        }

        const dvl_view = try allocator.alloc(sheet_plan.DataValidationList, st.data_validations.items.len);
        defer allocator.free(dvl_view);
        for (st.data_validations.items, 0..) |dv, k| {
            dvl_view[k] = .{ .range = dv.range, .values = @ptrCast(dv.values) };
        }

        const dvr_view = try allocator.alloc(sheet_plan.DataValidationRange, st.data_validation_ranges.items.len);
        defer allocator.free(dvr_view);
        for (st.data_validation_ranges.items, 0..) |dv, k| {
            dvr_view[k] = .{
                .range = dv.range,
                .kind_name = dv.kind_name,
                .op_name = dv.op_name,
                .formula1 = dv.formula1,
                .formula2 = dv.formula2,
            };
        }

        try sheet_plan.emitWorksheetXml(allocator, &full, .{
            .body = s.body,
            .freeze_rows = st.freeze_rows,
            .freeze_cols = st.freeze_cols,
            .column_widths = cw_view,
            .auto_filter_range = st.auto_filter_range,
            .merged_cells = merges_view,
            .conditional_formats = cf_view,
            .data_validations = dvl_view,
            .data_validation_ranges = dvr_view,
            .hyperlinks = hyperlinks_view,
            .internal_hyperlinks = internal_hl_view,
            .comment_count = st.comments.items.len,
        });

        var name_buf: [64]u8 = undefined;
        const entry_name = try std.fmt.bufPrint(&name_buf, "xl/worksheets/sheet{d}.xml", .{i + 1});
        try arc.addEntry(entry_name, full.items, deflate);
    }

    // 5a. Per-sheet rels (hyperlinks + comments).
    for (inputs.sheets, 0..) |s, i| {
        const st = s.state;
        const hyperlinks_view = try allocator.alloc(sheet_plan.Hyperlink, st.hyperlinks.items.len);
        defer allocator.free(hyperlinks_view);
        for (st.hyperlinks.items, 0..) |h, k| {
            hyperlinks_view[k] = .{ .range = h.range, .url = h.url };
        }

        var rels: std.ArrayListUnmanaged(u8) = .empty;
        defer rels.deinit(allocator);
        const wrote = try sheet_plan.emitSheetRels(
            allocator,
            &rels,
            i,
            hyperlinks_view,
            st.comments.items.len,
        );
        if (!wrote) continue;

        var rels_name_buf: [64]u8 = undefined;
        const rels_name = try std.fmt.bufPrint(&rels_name_buf, "xl/worksheets/_rels/sheet{d}.xml.rels", .{i + 1});
        try arc.addEntry(rels_name, rels.items, deflate);
    }

    // 5b. Per-sheet commentsN.xml + vmlDrawingN.vml.
    for (inputs.sheets, 0..) |s, i| {
        const st = s.state;
        if (st.comments.items.len == 0) continue;

        const comments_view = try allocator.alloc(sheet_plan.Comment, st.comments.items.len);
        defer allocator.free(comments_view);
        for (st.comments.items, 0..) |c, k| {
            comments_view[k] = .{ .ref = c.ref, .author = c.author, .text = c.text };
        }

        var cx: std.ArrayListUnmanaged(u8) = .empty;
        defer cx.deinit(allocator);
        try sheet_plan.emitCommentsXml(allocator, &cx, comments_view);
        var cn_buf: [64]u8 = undefined;
        const cn = try std.fmt.bufPrint(&cn_buf, "xl/comments{d}.xml", .{i + 1});
        try arc.addEntry(cn, cx.items, deflate);

        var vml: std.ArrayListUnmanaged(u8) = .empty;
        defer vml.deinit(allocator);
        try sheet_plan.emitVmlDrawingXml(allocator, &vml, comments_view);
        var vml_buf: [64]u8 = undefined;
        const vml_name = try std.fmt.bufPrint(&vml_buf, "xl/drawings/vmlDrawing{d}.vml", .{i + 1});
        try arc.addEntry(vml_name, vml.items, deflate);
    }

    // 6. xl/sharedStrings.xml
    {
        var sst: std.ArrayListUnmanaged(u8) = .empty;
        defer sst.deinit(allocator);
        const unique_count = inputs.sst_plan.new_strings.items.len +
            inputs.sst_plan.new_rich_strings.items.len;
        try sst.print(allocator, SST_HEAD_FMT, .{ inputs.sst_count, unique_count });
        for (inputs.sst_plan.new_strings.items) |s| {
            try sst.appendSlice(allocator, "<si><t xml:space=\"preserve\">");
            try appendXmlEscaped(allocator, &sst, s);
            try sst.appendSlice(allocator, "</t></si>");
        }
        for (inputs.sst_plan.new_rich_strings.items) |entry| {
            try sst.appendSlice(allocator, "<si>");
            for (entry.runs) |r| {
                try sst.appendSlice(allocator, "<r>");
                const has_props = r.bold or r.italic or r.color_argb != null or
                    r.font_size != null or r.font_name != null;
                if (has_props) {
                    try sst.appendSlice(allocator, "<rPr>");
                    if (r.bold) try sst.appendSlice(allocator, "<b/>");
                    if (r.italic) try sst.appendSlice(allocator, "<i/>");
                    if (r.font_size) |sz| try sst.print(allocator, "<sz val=\"{d}\"/>", .{sz});
                    if (r.color_argb) |c| {
                        try sst.appendSlice(allocator, "<color rgb=\"");
                        try sst.appendSlice(allocator, c);
                        try sst.appendSlice(allocator, "\"/>");
                    }
                    if (r.font_name) |n| {
                        try sst.appendSlice(allocator, "<rFont val=\"");
                        try appendXmlEscaped(allocator, &sst, n);
                        try sst.appendSlice(allocator, "\"/>");
                    }
                    try sst.appendSlice(allocator, "</rPr>");
                }
                try sst.appendSlice(allocator, "<t xml:space=\"preserve\">");
                try appendXmlEscaped(allocator, &sst, r.text);
                try sst.appendSlice(allocator, "</t></r>");
            }
            try sst.appendSlice(allocator, "</si>");
        }
        try sst.appendSlice(allocator, SST_TAIL);
        try arc.addEntry("xl/sharedStrings.xml", sst.items, deflate);
    }

    // 7. xl/styles.xml — only when caller registered any styles.
    if (have_styles) {
        var styles_buf: std.ArrayListUnmanaged(u8) = .empty;
        defer styles_buf.deinit(allocator);
        // StylesPlan.emit may return error.OutOfMemory or its broader
        // validation set; bubble through.
        try inputs.styles_plan.emit(allocator, &styles_buf);
        try arc.addEntry("xl/styles.xml", styles_buf.items, deflate);
    }

    try arc.finalize();
}

/// Convenience: build the archive into a fresh ArrayList and write it
/// to `path`, truncating any existing file.
pub fn saveArchiveToPath(
    allocator: Allocator,
    io: std.Io,
    path: []const u8,
    inputs: ArchiveInputs,
    deflate: DeflateFn,
    poller: Poller,
) !void {
    var zip_buf: std.ArrayListUnmanaged(u8) = .empty;
    defer zip_buf.deinit(allocator);

    try emitArchiveBytes(allocator, &zip_buf, inputs, deflate, poller);

    try std.Io.Dir.cwd().writeFile(io, .{
        .sub_path = path,
        .data = zip_buf.items,
        .flags = .{ .truncate = true },
    });
}

// ─── Helpers ─────────────────────────────────────────────────────────

/// XML-attribute / text escape covering the five XML-significant
/// characters. Lifted verbatim from `src/writer.zig`'s
/// `appendXmlEscaped`.
fn appendXmlEscaped(
    alloc: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    s: []const u8,
) !void {
    for (s) |c| switch (c) {
        '&' => try out.appendSlice(alloc, "&amp;"),
        '<' => try out.appendSlice(alloc, "&lt;"),
        '>' => try out.appendSlice(alloc, "&gt;"),
        '"' => try out.appendSlice(alloc, "&quot;"),
        '\'' => try out.appendSlice(alloc, "&apos;"),
        else => try out.append(alloc, c),
    };
}

// ─── Tests ───────────────────────────────────────────────────────────

const testing = std.testing;

fn stubDeflate(
    alloc: Allocator,
    input: []const u8,
    out: *std.ArrayListUnmanaged(u8),
    poller: Poller,
) anyerror!void {
    var it = poller.chunks(input);
    while (try it.next()) |chunk| try out.appendSlice(alloc, chunk);
    try out.append(alloc, 0);
}

test "fresh_emit: empty single-sheet archive — no styles, no SST entries" {
    const a = testing.allocator;

    var st = sheet_plan.SheetState{};
    defer st.deinit(a);

    const sheets = [_]SheetInput{
        .{ .name = "Sheet1", .body = "", .state = &st },
    };

    var sst_plan: sst_plan_mod.SstExtensionPlan = .{};
    defer sst_plan.deinit(a);

    var styles_plan: styles_plan_mod.StylesPlan = .{};
    defer styles_plan.deinit(a);

    var workbook_xml_plan: workbook_xml_plan_mod.WorkbookXmlPlan = .{};
    defer workbook_xml_plan.deinit(a);

    var out: std.ArrayListUnmanaged(u8) = .empty;
    defer out.deinit(a);

    try emitArchiveBytes(a, &out, .{
        .sheets = &sheets,
        .sst_plan = &sst_plan,
        .sst_count = 0,
        .styles_plan = &styles_plan,
        .workbook_xml_plan = &workbook_xml_plan,
    }, stubDeflate, .none);

    // Final output should be a parseable ZIP archive.
    try testing.expect(out.items.len > 22);
    try testing.expectEqualSlices(u8, &std.zip.end_record_sig, out.items[out.items.len - 22 .. out.items.len - 22 + 4]);
}

test "fresh_emit: a fresh archive carries no metadata part, so it reads as CV1" {
    // §5.4d (M4f): a workbook with no compatibility metadata IS
    // compatibility version 1, and every file this emitter writes is
    // such a workbook. The engine's default depends on that being true
    // of the emitter, so it is asserted here rather than inferred from
    // the part list in a comment: if a later row starts writing
    // `xl/metadata.xml`, `LEN` of an astral character changes meaning
    // for every fresh file and this is where that surfaces.
    const a = testing.allocator;

    var st = sheet_plan.SheetState{};
    defer st.deinit(a);

    const sheets = [_]SheetInput{
        .{ .name = "Sheet1", .body = "", .state = &st },
    };

    var sst_plan: sst_plan_mod.SstExtensionPlan = .{};
    defer sst_plan.deinit(a);

    var styles_plan: styles_plan_mod.StylesPlan = .{};
    defer styles_plan.deinit(a);

    var workbook_xml_plan: workbook_xml_plan_mod.WorkbookXmlPlan = .{};
    defer workbook_xml_plan.deinit(a);

    var out: std.ArrayListUnmanaged(u8) = .empty;
    defer out.deinit(a);

    try emitArchiveBytes(a, &out, .{
        .sheets = &sheets,
        .sst_plan = &sst_plan,
        .sst_count = 0,
        .styles_plan = &styles_plan,
        .workbook_xml_plan = &workbook_xml_plan,
    }, stubDeflate, .none);

    // The stub deflate stores rather than compresses, so the part names
    // and the content-type overrides are both in the bytes verbatim.
    try testing.expect(std.mem.indexOf(u8, out.items, "metadata.xml") == null);
    try testing.expect(std.mem.indexOf(u8, out.items, "sheetMetadata") == null);
    // …and the archive is otherwise the one the test above built, so a
    // vacuous pass is not available: the parts that SHOULD be there are.
    try testing.expect(std.mem.indexOf(u8, out.items, "xl/workbook.xml") != null);
}

test "fresh_emit: NoSheets refusal" {
    const a = testing.allocator;

    var sst_plan: sst_plan_mod.SstExtensionPlan = .{};
    defer sst_plan.deinit(a);

    var styles_plan: styles_plan_mod.StylesPlan = .{};
    defer styles_plan.deinit(a);

    var workbook_xml_plan: workbook_xml_plan_mod.WorkbookXmlPlan = .{};
    defer workbook_xml_plan.deinit(a);

    var out: std.ArrayListUnmanaged(u8) = .empty;
    defer out.deinit(a);

    const result = emitArchiveBytes(a, &out, .{
        .sheets = &.{},
        .sst_plan = &sst_plan,
        .sst_count = 0,
        .styles_plan = &styles_plan,
        .workbook_xml_plan = &workbook_xml_plan,
    }, stubDeflate, .none);
    try testing.expectError(error.NoSheets, result);
}
