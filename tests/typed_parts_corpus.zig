//! Corpus parity sweep for B1 iter-wb-1 typed-overlay parsers.
//!
//! For every committed-or-fetched .xlsx in tests/corpus/, opens the
//! package store, locates each well-known part (workbook.xml,
//! sharedStrings.xml, styles.xml, theme*.xml, the first sheet), and
//! runs it through the matching typed_parts parser. The contract is
//! "every part that exists parses cleanly without leaks." Missing
//! parts are skipped (a workbook may legitimately have no SST or no
//! styles).
//!
//! Per-shape semantic assertions live inline in the parser files.
//! This sweep is the wide robustness check — does the parser survive
//! every real-world OOXML emitter the corpus has?

const std = @import("std");
const pkg = @import("zlsx_pkg");

const corpus_dir = "tests/corpus/";

const fixtures = [_][]const u8{
    "frictionless_2sheets.xlsx",
    "openpyxl_guess_types.xlsx",
    "phpoi_test1.xlsx",
    "calamine_empty_s_attribute.xlsx",
    "calamine_empty_shared_string.xlsx",
    "calamine_encoded_entities.xlsx",
    "calamine_non_monotonic_si.xlsx",
    "openxlsx_loadExample.xlsx",
    "phpsheet_3654c.xlsx",
    "poi_57893_many_merges.xlsx",
    "poi_58325_db.xlsx",
    "poi_excel_with_trash_item.xlsx",
    "poi_poc_shared_strings.xlsx",
    "ecdc_covid.xlsx",
    "ons_cpi_detailed.xlsx",
    "wdi_excel.xlsx",
    "worldbank_catalog.xlsx",
};

test "typed_parts corpus sweep — every well-known part parses without leak" {
    const alloc = std.testing.allocator;
    var any_seen: bool = false;
    var workbooks_parsed: u32 = 0;
    var sheets_parsed: u32 = 0;
    var ssts_parsed: u32 = 0;
    var styles_parsed: u32 = 0;
    var themes_parsed: u32 = 0;

    for (fixtures) |name| {
        var path_buf: [256]u8 = undefined;
        const path = try std.fmt.bufPrint(&path_buf, "{s}{s}", .{ corpus_dir, name });
        std.fs.cwd().access(path, .{}) catch continue;
        any_seen = true;

        var store = pkg.PartStore.open(alloc, path) catch |err| {
            std.debug.print("\n  [open failed] {s} -> {s}\n", .{ name, @errorName(err) });
            return err;
        };
        defer store.deinit();

        // workbook.xml — every well-formed OOXML package has exactly one.
        if (store.part("xl/workbook.xml")) |p| {
            var view = pkg.typed_parts.workbook_xml.parse(alloc, p.bytes) catch |err| {
                std.debug.print("\n  [workbook.xml parse failed] {s} -> {s}\n", .{ name, @errorName(err) });
                return err;
            };
            defer view.deinit(alloc);
            try std.testing.expect(view.sheets.len > 0);
            workbooks_parsed += 1;
        }

        // sharedStrings.xml — optional.
        if (store.part("xl/sharedStrings.xml")) |p| {
            var view = pkg.typed_parts.sst_xml.parse(alloc, p.bytes) catch |err| {
                std.debug.print("\n  [sharedStrings.xml parse failed] {s} -> {s}\n", .{ name, @errorName(err) });
                return err;
            };
            defer view.deinit(alloc);
            ssts_parsed += 1;
        }

        // styles.xml — optional but present on every fixture this corpus carries.
        if (store.part("xl/styles.xml")) |p| {
            var view = pkg.typed_parts.styles_xml.parse(alloc, p.bytes) catch |err| {
                std.debug.print("\n  [styles.xml parse failed] {s} -> {s}\n", .{ name, @errorName(err) });
                return err;
            };
            defer view.deinit(alloc);
            styles_parsed += 1;
        }

        // theme*.xml — find the first one. Excel-emitted workbooks
        // typically ship xl/theme/theme1.xml; some emitters omit it.
        const part_names = try store.partNames();
        for (part_names) |pn| {
            if (std.mem.startsWith(u8, pn, "xl/theme/") and std.mem.endsWith(u8, pn, ".xml")) {
                if (store.part(pn)) |p| {
                    var view = pkg.typed_parts.theme_xml.parse(alloc, p.bytes) catch |err| {
                        std.debug.print("\n  [{s} parse failed] {s} -> {s}\n", .{ pn, name, @errorName(err) });
                        return err;
                    };
                    defer view.deinit(alloc);
                    themes_parsed += 1;
                }
                break;
            }
        }

        // First sheet found by part name — sufficient for parity sweep.
        for (part_names) |pn| {
            if (std.mem.startsWith(u8, pn, "xl/worksheets/sheet") and std.mem.endsWith(u8, pn, ".xml")) {
                if (store.part(pn)) |p| {
                    var view = pkg.typed_parts.sheet_xml.parse(alloc, p.bytes) catch |err| {
                        std.debug.print("\n  [{s} parse failed] {s} -> {s}\n", .{ pn, name, @errorName(err) });
                        return err;
                    };
                    defer view.deinit(alloc);
                    sheets_parsed += 1;
                }
                break;
            }
        }
    }

    if (!any_seen) {
        std.debug.print("\n  [no fixtures present — run scripts/fetch_test_corpus.sh]\n", .{});
        return error.SkipZigTest;
    }

    // Sanity: every fixture has at least workbook + first sheet.
    try std.testing.expect(workbooks_parsed > 0);
    try std.testing.expect(sheets_parsed > 0);
    // SST / styles / theme are optional, but the committed corpus
    // covers all three at least once.
    try std.testing.expect(ssts_parsed > 0);
    try std.testing.expect(styles_parsed > 0);
    try std.testing.expect(themes_parsed > 0);
}
