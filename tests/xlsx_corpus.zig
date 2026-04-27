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

fn rowCount(book: *xlsx.Book, sheet: xlsx.Sheet, alloc: std.mem.Allocator) !usize {
    var rows = try book.rows(sheet, alloc);
    defer rows.deinit();
    var n: usize = 0;
    while (try rows.next()) |_| : (n += 1) {}
    return n;
}

fn firstRowCells(
    book: *xlsx.Book,
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
    try std.testing.expectEqual(@as(usize, 18), book.sharedStringsCount());

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
    try std.testing.expect(book.sharedStringsCount() >= 1000);

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

// ─── Large fixtures (group 2) — fetched, not committed ────────────
// Tests skip cleanly when the file is missing. Each one targets a
// specific stress dimension the small base corpus can't reach.

test "WDI Excel — 401k rows × 6 sheets, r-less <c> cells" {
    // The World Bank's WDI exporter emits `<row><c t="s"><v>N</v></c>…`
    // with no `r=` on either tag. This is spec-legal but unusual; before
    // the implicit-column fix, every cell tripped MalformedXml. Pin the
    // row count + sheet count so a regression in the fallback path
    // re-fails the suite.
    const alloc = std.testing.allocator;
    var book = try openOrSkip(alloc, "wdi_excel.xlsx");
    defer book.deinit();

    try std.testing.expect(book.sheets.len >= 6);
    try std.testing.expect(book.sharedStringsCount() >= 200_000);

    const sheet = book.sheets[0];
    const n = try rowCount(&book, sheet, alloc);
    // Sheet1 declares dimension A1:BQ401395; tolerate updates.
    try std.testing.expect(n >= 100_000);
}

test "ECDC COVID — 49k single-sheet rows, modest SST" {
    const alloc = std.testing.allocator;
    var book = try openOrSkip(alloc, "ecdc_covid.xlsx");
    defer book.deinit();

    try std.testing.expectEqual(@as(usize, 1), book.sheets.len);
    const sheet = book.sheets[0];
    const n = try rowCount(&book, sheet, alloc);
    try std.testing.expect(n >= 40_000);
}

test "ONS CPI detailed — 41-sheet workbook" {
    const alloc = std.testing.allocator;
    var book = try openOrSkip(alloc, "ons_cpi_detailed.xlsx");
    defer book.deinit();
    // 1 contents + 40 numbered tables; tolerate ±2 around the publication.
    try std.testing.expect(book.sheets.len >= 35 and book.sheets.len <= 45);
}

test "POI 57893 many-merges — 50k mergeCells" {
    const alloc = std.testing.allocator;
    var book = try openOrSkip(alloc, "poi_57893_many_merges.xlsx");
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 1), book.sheets.len);
    const merged = book.mergedRanges(book.sheets[0]);
    try std.testing.expect(merged.len >= 50_000);
}

test "POI 58325_db — self-closing <row/> elements" {
    // Sheet1 has 4 rows, three of which are `<row r="N" ht="..."/>`
    // (style/height-only, no cells). Before the self-closing fix the
    // row iterator went off the rails on the first such row.
    const alloc = std.testing.allocator;
    var book = try openOrSkip(alloc, "poi_58325_db.xlsx");
    defer book.deinit();
    const n = try rowCount(&book, book.sheets[0], alloc);
    try std.testing.expectEqual(@as(usize, 4), n);
}

test "openxlsx loadExample — 4 sheets including pivot/slicer" {
    const alloc = std.testing.allocator;
    var book = try openOrSkip(alloc, "openxlsx_loadExample.xlsx");
    defer book.deinit();
    try std.testing.expect(book.sheets.len >= 4);
}

// ─── Adversarial fixtures (group 3) — fetched + locally derived ───
// Each test captures the *current* zlsx behavior. "Cleanly errored"
// = typed error returned, no panic, no hang. "Permissively read"
// = open succeeds, downstream API returns sane values (no OOM, no
// false-positive parse).

test "adversarial: truncated ZIPs error cleanly" {
    const alloc = std.testing.allocator;
    const cases = [_][]const u8{
        "derived_truncated_pre_eocd.xlsx",
        "derived_truncated_mid_payload.xlsx",
        "derived_truncated_signature.xlsx",
        "poi_crash_274d6342.xlsx",
        "poi_crash_9bf3cd4b.xlsx",
        "poi_xlsx_corrupted.xlsx", // permissive open ok
    };
    for (cases) |name| {
        var path_buf: [256]u8 = undefined;
        const path = try std.fmt.bufPrint(&path_buf, "{s}{s}", .{ corpus_dir, name });
        if (std.fs.cwd().access(path, .{})) |_| {} else |_| continue;
        // Whether open returns an error or succeeds is fixture-dependent;
        // the contract is "no crash / no hang / no UB". We only assert
        // we get *some* result back without panicking.
        if (xlsx.Book.open(alloc, path)) |book_const| {
            var book = book_const;
            defer book.deinit();
        } else |_| {}
    }
}

test "adversarial: bare ZIPs (not xlsx) error with a typed reason" {
    const alloc = std.testing.allocator;
    const cases = [_][]const u8{
        "ziprs_invalid_offset.zip",
        "ziprs_invalid_cde_files_greater.zip",
        "ziprs_aes_archive.zip",
        "ziprs_data_descriptor.zip",
        "ziprs_comment_garbage.zip",
        "ziprs_extended_timestamp_bad.zip",
        "ziprs_misaligned_comment.zip",
        "wdi_excel.zip",
    };
    for (cases) |name| {
        var path_buf: [256]u8 = undefined;
        const path = try std.fmt.bufPrint(&path_buf, "{s}{s}", .{ corpus_dir, name });
        if (std.fs.cwd().access(path, .{})) |_| {} else |_| continue;
        // Bare ZIPs must NOT parse as xlsx. They either fail at the ZIP
        // layer (BadZip) or at the missing-workbook step
        // (MissingWorkbook). Anything else means the reader silently
        // accepted a non-xlsx archive, which is a real defect.
        const result = xlsx.Book.open(alloc, path);
        if (result) |book_const| {
            var book = book_const;
            defer book.deinit();
            std.debug.print("\n  [unexpected] {s} opened as xlsx — expected error\n", .{name});
            return error.TestUnexpectedResult;
        } else |_| {}
    }
}

test "adversarial: MalformedSSTCount — declared count clamped to actual" {
    // The SST is declared as `count="8876876876876"` but contains 8
    // entries. zlsx must NOT trust the attribute (would over-allocate
    // and crash); it must walk the actual <si> elements. Real count = 8.
    const alloc = std.testing.allocator;
    var book = try openOrSkip(alloc, "poi_MalformedSSTCount.xlsx");
    defer book.deinit();
    try std.testing.expect(book.sharedStringsCount() < 100);
}

test "adversarial: shared-strings amplification PoC opens without OOM" {
    // POI's poc-shared-strings.xlsx is a billion-laughs-style PoC with
    // a single huge <si> blob. Opening must complete in bounded memory
    // (no exponential expansion) and the row iterator must not fault.
    const alloc = std.testing.allocator;
    var book = try openOrSkip(alloc, "poi_poc_shared_strings.xlsx");
    defer book.deinit();
    // 1 SST entry (the giant one); 4000 rows referencing it.
    try std.testing.expect(book.sharedStringsCount() < 1000);
    const n = try rowCount(&book, book.sheets[0], alloc);
    try std.testing.expect(n >= 1000);
}

test "adversarial: encrypted xlsx — opens permissively, no crash" {
    // POI's workbookProtection-workbook_password-2013.xlsx is encrypted
    // at the OLE-CFB layer. zlsx isn't an OLE reader, so it'll read
    // whatever ZIP happens to be inside (typically a small placeholder
    // workbook). Goal: no panic, no UB. v1 doesn't surface a typed
    // "encrypted" error — that's a follow-up.
    const alloc = std.testing.allocator;
    var book = try openOrSkip(alloc, "poi_workbook_password_2013.xlsx");
    defer book.deinit();
}

test "adversarial: trash entry inside otherwise-valid xlsx" {
    // POI's Excel_file_with_trash_item.xlsx has an extra ZIP entry
    // unrelated to the OOXML structure. zlsx ignores unknown entries
    // and reads the workbook normally.
    const alloc = std.testing.allocator;
    var book = try openOrSkip(alloc, "poi_excel_with_trash_item.xlsx");
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 1), book.sheets.len);
    const n = try rowCount(&book, book.sheets[0], alloc);
    try std.testing.expect(n >= 100);
}

test "adversarial: XXE in [Content_Types].xml — no external fetch" {
    // POI's xxe_in_schema.xlsx attempts XML external-entity injection.
    // zlsx must parse with no external resolution (we never call out
    // to a network or filesystem path). Goal: no panic, no hang.
    const alloc = std.testing.allocator;
    var book = try openOrSkip(alloc, "poi_xxe_in_schema.xlsx");
    defer book.deinit();
}

test "adversarial: clusterfuzz minimised XSSF input" {
    const alloc = std.testing.allocator;
    var book = try openOrSkip(alloc, "poi_clusterfuzz_xssf.xlsx");
    defer book.deinit();
    // Whatever the minimised input parses to, opening must not panic.
}

test "adversarial: calamine fixtures (encoded entities, empty SI, etc.)" {
    const alloc = std.testing.allocator;
    const cases = [_]struct { name: []const u8, min_rows: usize }{
        .{ .name = "calamine_encoded_entities.xlsx", .min_rows = 1 },
        .{ .name = "calamine_empty_shared_string.xlsx", .min_rows = 1 },
        .{ .name = "calamine_empty_s_attribute.xlsx", .min_rows = 1 },
        .{ .name = "calamine_non_monotonic_si.xlsx", .min_rows = 1 },
    };
    for (cases) |c| {
        var path_buf: [256]u8 = undefined;
        const path = try std.fmt.bufPrint(&path_buf, "{s}{s}", .{ corpus_dir, c.name });
        if (std.fs.cwd().access(path, .{})) |_| {} else |_| continue;
        var book = try xlsx.Book.open(alloc, path);
        defer book.deinit();
        try std.testing.expect(book.sheets.len >= 1);
        const n = try rowCount(&book, book.sheets[0], alloc);
        try std.testing.expect(n >= c.min_rows);
    }
}

test "Editor.scanWorksheet (iter-cm-1): every Book.rows cell has a matching span" {
    // Phase 3d foundation. Walk worldbank_catalog through the new
    // span scanner; every non-empty cell that Book.rows surfaces
    // must have a matching CellSpan at the same (row, col).
    const alloc = std.testing.allocator;
    const path = corpus_dir ++ "worldbank_catalog.xlsx";
    var ed = try xlsx.Editor.open(alloc, path);
    defer ed.deinit();
    var spans = try ed.scanWorksheet(0);
    defer spans.deinit();

    var book = try xlsx.Book.open(alloc, path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], alloc);
    defer rows.deinit();

    var row_idx: u32 = 1;
    var checked: usize = 0;
    while (try rows.next()) |row| : (row_idx += 1) {
        for (row, 0..) |cell, col| {
            if (cell == .empty) continue;
            const found = spans.find(row_idx, @intCast(col));
            try std.testing.expect(found != null);
            checked += 1;
        }
    }
    // The fixture has 161 rows × 26 cols with sparse blanks; we
    // should be checking thousands of cells. If this drops to a
    // suspiciously small number, the row iterator changed shape.
    try std.testing.expect(checked >= 2000);
}

test "corpus surface: iter28-34 reader APIs round-trip on real fixtures" {
    // The per-cell style / font / fill / border / numFmt / rich-runs /
    // comments APIs were added in iter28-34 but their tests use
    // synthesised fixtures. Exercise them against real-world xlsx files
    // so a regression in the styles.xml parser (or equivalent) fails
    // CI on the fixtures we already ship — not just on inputs I
    // hand-crafted to match the parser's expectations.
    const alloc = std.testing.allocator;

    // openpyxl_guess_types has an xl/styles.xml part (date + numeric
    // number formats). Every row must produce a styles slice of the
    // same length as cells; numberFormat() + isDateFormat() must not
    // crash on any style index the sheet references.
    {
        var book = try openOrSkip(alloc, "openpyxl_guess_types.xlsx");
        defer book.deinit();
        const sheet = book.sheets[0];
        var rows = try book.rows(sheet, alloc);
        defer rows.deinit();
        while (try rows.next()) |cells| {
            const styles = rows.styleIndices();
            try std.testing.expectEqual(cells.len, styles.len);
            for (styles) |maybe_idx| {
                const s = maybe_idx orelse continue;
                // numberFormat may return null for a corrupt / sparse
                // styles.xml, but it must NEVER crash. isDateFormat
                // piggybacks on it and returns false on absence.
                _ = book.numberFormat(s);
                _ = book.isDateFormat(s);
                _ = book.cellFont(s);
                _ = book.cellFill(s);
                _ = book.cellBorder(s);
            }
        }
    }

    // Worldbank catalog has 1,144 SST entries. richRuns(i) must return
    // null for every plain entry (the common case — no regression
    // into false-positive rich-runs on files without <r> wrappers).
    {
        var book = try openOrSkip(alloc, "worldbank_catalog.xlsx");
        defer book.deinit();
        var rich_count: usize = 0;
        for (0..book.sharedStringsCount()) |i| {
            if (book.richRuns(i) != null) rich_count += 1;
        }
        // Plain-text SST — no runs should surface. If any do, it's a
        // parser false-positive.
        try std.testing.expectEqual(@as(usize, 0), rich_count);
    }

    // None of the corpus files are authored with cell comments today,
    // but comments() must still return an empty slice (never crash /
    // leak) for every sheet.
    {
        var book = try openOrSkip(alloc, "worldbank_catalog.xlsx");
        defer book.deinit();
        for (book.sheets) |sheet| {
            try std.testing.expectEqual(@as(usize, 0), book.comments(sheet).len);
        }
    }
}
