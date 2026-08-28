//! Corpus parity sweep for B1 iter-wb-2 — Workbook + Worksheet
//! composition over every fixture in tests/corpus/.
//!
//! Opens each fixture as a `Workbook`, asserts non-zero sheet count,
//! materialises every Worksheet (forcing SheetXml parse), and reads
//! rows / merges / hyperlinks / validations / conditional formats /
//! freeze pane through the typed accessors. The contract is "no
//! crash, no leak, every sheet's typed view materialises".
//!
//! Per-shape semantic assertions live inline in `pkg/workbook.zig`.
//! This sweep is the wide robustness check across all real-world
//! emitters.

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

test "Workbook corpus sweep — open + materialise every sheet without leak" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const alloc = std.testing.allocator;
    var any_seen: bool = false;
    var sheets_total: u32 = 0;
    var rows_total: u64 = 0;
    var pivots_total: usize = 0;

    for (fixtures) |name| {
        var path_buf: [256]u8 = undefined;
        const path = try std.fmt.bufPrint(&path_buf, "{s}{s}", .{ corpus_dir, name });
        std.Io.Dir.cwd().access(io, path, .{}) catch continue;
        any_seen = true;

        var wb = pkg.Workbook.open(alloc, io, path) catch |err| {
            std.debug.print("\n  [Workbook.open failed] {s} -> {s}\n", .{ name, @errorName(err) });
            return err;
        };
        defer wb.deinit();

        const count = wb.sheetCount();
        try std.testing.expect(count > 0);
        sheets_total += count;

        var i: u32 = 0;
        while (i < count) : (i += 1) {
            const ws = wb.sheet(i) catch |err| {
                std.debug.print("\n  [{s} sheet({d})] -> {s}\n", .{ name, i, @errorName(err) });
                return err;
            };
            const rows = ws.rows() catch |err| {
                std.debug.print("\n  [{s} sheet({d}).rows] -> {s}\n", .{ name, i, @errorName(err) });
                return err;
            };
            rows_total += rows.len;

            // Touch every other accessor to confirm none crashes
            // through the composition. Ignore values; the contract
            // is "no error" over the whole corpus.
            _ = try ws.merges();
            _ = try ws.hyperlinks();
            _ = try ws.validations();
            _ = try ws.conditionalFormats();
            _ = try ws.dimension();
            _ = try ws.freezePane();
        }

        // Workbook-scope views are optional but should never error
        // when the part is present.
        _ = try wb.sst();
        _ = try wb.styles();
        _ = wb.definedNames();
        _ = wb.calcProperties();

        // S6: the pivot graph walks relationships every fixture has and
        // must never error on a workbook without pivots. The one fixture
        // with pivots is pinned in `pkg/pivots.zig`'s own corpus test.
        var pivots = wb.pivotTables() catch |err| {
            std.debug.print("\n  [{s} pivotTables] -> {s}\n", .{ name, @errorName(err) });
            return err;
        };
        defer pivots.deinit();
        pivots_total += pivots.tables.len;
    }

    if (!any_seen) {
        std.debug.print("\n  [no fixtures present — run scripts/fetch_test_corpus.sh]\n", .{});
        return error.SkipZigTest;
    }

    try std.testing.expect(sheets_total > 0);
    try std.testing.expect(rows_total > 0);
    // `openxlsx_loadExample.xlsx` carries two; the count is a floor so a
    // fixture added later with pivots of its own does not fail this.
    var has_pivot_fixture = false;
    for (fixtures) |name| {
        if (std.mem.eql(u8, name, "openxlsx_loadExample.xlsx")) {
            var path_buf: [256]u8 = undefined;
            const path = try std.fmt.bufPrint(&path_buf, "{s}{s}", .{ corpus_dir, name });
            std.Io.Dir.cwd().access(io, path, .{}) catch continue;
            has_pivot_fixture = true;
        }
    }
    if (has_pivot_fixture) try std.testing.expect(pivots_total >= 2);
}
