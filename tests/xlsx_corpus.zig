//! Integration tests for `zlsx` against a curated set of public xlsx
//! files. Run with `zig build test-corpus`.
//!
//! The corpus is materialized by `scripts/fetch_test_corpus.sh` into
//! `tests/corpus/`. If a file is missing these tests skip with
//! `error.SkipZigTest` rather than fail — so CI can run them after the
//! fetch step without hard-coupling them to the fast `zig build test`
//! default.
//!
//! See `docs/xlsx_test_corpus.md` for provenance and what each file
//! exercises.

const std = @import("std");
const xlsx = @import("zlsx");

const corpus_dir = "tests/corpus/";

fn openOrSkip(alloc: std.mem.Allocator, filename: []const u8) !xlsx.Book {
    var path_buf: [256]u8 = undefined;
    const path = try std.fmt.bufPrint(&path_buf, "{s}{s}", .{ corpus_dir, filename });
    return xlsx.Book.open(alloc, path) catch |err| switch (err) {
        error.FileNotFound => {
            std.debug.print("\n  [skip] {s} not in corpus — run scripts/fetch_test_corpus.sh\n", .{filename});
            return error.SkipZigTest;
        },
        else => return err,
    };
}

fn rowCount(book: *const xlsx.Book, sheet: xlsx.Sheet, alloc: std.mem.Allocator) !usize {
    var rows = try book.rows(sheet, alloc);
    defer rows.deinit();
    var n: usize = 0;
    while (try rows.next()) |_| : (n += 1) {}
    return n;
}

fn firstRowCells(
    book: *const xlsx.Book,
    sheet: xlsx.Sheet,
    alloc: std.mem.Allocator,
    out: []xlsx.Cell,
) !usize {
    var rows = try book.rows(sheet, alloc);
    defer rows.deinit();
    const first = (try rows.next()) orelse return 0;
    // Must deep-copy the string bytes — the owned slices in rows.owned
    // live only for this iteration.
    const n = @min(first.len, out.len);
    for (first[0..n], 0..) |cell, i| {
        out[i] = switch (cell) {
            .string => |s| .{ .string = try alloc.dupe(u8, s) },
            else => cell,
        };
    }
    return n;
}

test "frictionless sample-2-sheets — small SST, multi-sheet" {
    const alloc = std.testing.allocator;
    var book = try openOrSkip(alloc, "frictionless_2sheets.xlsx");
    defer book.deinit();

    try std.testing.expectEqual(@as(usize, 2), book.sheets.len);
    try std.testing.expectEqualStrings("Sheet1", book.sheets[0].name);
    try std.testing.expectEqualStrings("Sheet2", book.sheets[1].name);
    try std.testing.expectEqual(@as(usize, 18), book.shared_strings.len);

    const sheet = book.sheetByName("Sheet1") orelse return error.SheetMissing;

    // Header row: "header1" "header2" "header3".
    var cells: [3]xlsx.Cell = undefined;
    const n = try firstRowCells(&book, sheet, alloc, &cells);
    defer for (cells[0..n]) |c| switch (c) {
        .string => |s| alloc.free(s),
        else => {},
    };
    try std.testing.expectEqual(@as(usize, 3), n);
    try std.testing.expectEqualStrings("header1", cells[0].string);
    try std.testing.expectEqualStrings("header2", cells[1].string);
    try std.testing.expectEqualStrings("header3", cells[2].string);

    try std.testing.expectEqual(@as(usize, 3), try rowCount(&book, sheet, alloc));
}

test "openpyxl guess_types — mixed cell types in a genuine fixture" {
    const alloc = std.testing.allocator;
    var book = try openOrSkip(alloc, "openpyxl_guess_types.xlsx");
    defer book.deinit();

    try std.testing.expectEqual(@as(usize, 1), book.sheets.len);
    try std.testing.expectEqualStrings("Sheet1", book.sheets[0].name);

    // Sheet has 2 rows; every cell type (number, date-as-string, scientific-notation)
    // should decode without error. We don't assert exact values because this
    // fixture exists to exercise type-guessing and we only need to round-trip.
    const sheet = book.sheets[0];
    const n = try rowCount(&book, sheet, alloc);
    try std.testing.expect(n >= 2);
}

test "ph-poi test1 — 3 sheets, sparse diagonal + embedded newline" {
    const alloc = std.testing.allocator;
    var book = try openOrSkip(alloc, "phpoi_test1.xlsx");
    defer book.deinit();

    try std.testing.expectEqual(@as(usize, 3), book.sheets.len);
    try std.testing.expectEqualStrings("Sheet1", book.sheets[0].name);
    try std.testing.expectEqualStrings("Sheet2", book.sheets[1].name);
    try std.testing.expectEqualStrings("Sheet3", book.sheets[2].name);

    // Sheet1 has a diagonal layout: A1, B2, C3 (embedded newline), … across rows.
    const sheet = book.sheets[0];
    const n = try rowCount(&book, sheet, alloc);
    try std.testing.expect(n >= 3);
}

test "World Bank Data Catalog — heavy SST (1144 entries, 143 KB)" {
    const alloc = std.testing.allocator;
    var book = try openOrSkip(alloc, "worldbank_catalog.xlsx");
    defer book.deinit();

    try std.testing.expectEqual(@as(usize, 2), book.sheets.len);
    try std.testing.expectEqualStrings("World Bank Data Catalog", book.sheets[0].name);
    // The SST carries the 1,144 unique text values. Exact count ties this
    // test to the file version currently committed — if WB ships an update
    // and the count drifts, this is the signal to re-pin. Pin a lower
    // bound instead of equality to tolerate small catalog updates.
    try std.testing.expect(book.shared_strings.len >= 1000);

    const sheet = book.sheetByName("World Bank Data Catalog") orelse return error.SheetMissing;

    // Header row first-cell must be the well-known column name. This
    // exercises: SST index parsing (t="s" cells), lookup into
    // shared_strings, and string return to the caller.
    var cells: [30]xlsx.Cell = undefined;
    const n = try firstRowCells(&book, sheet, alloc, &cells);
    defer for (cells[0..n]) |c| switch (c) {
        .string => |s| alloc.free(s),
        else => {},
    };
    try std.testing.expect(n >= 26);
    try std.testing.expectEqualStrings("DataCatalog_id", cells[0].string);
    try std.testing.expectEqualStrings("Name", cells[1].string);

    // Full iteration must produce the expected row count (pinned — small
    // tolerance for catalog updates).
    const total = try rowCount(&book, sheet, alloc);
    try std.testing.expect(total >= 100 and total <= 500);
}
