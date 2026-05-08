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
const zlsx_pkg = @import("zlsx_pkg");

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
    var ed = try zlsx_pkg.Editor.open(alloc, path);
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

// ─── iter-er-7 task A: corpus parity sweep across editor mutation axes ───
//
// `Editor.save` was rewritten in PR #84 (`refactor/editor-er-6-thin-shim`)
// to delegate to `Workbook.save` / `PartStore.save`. The unit tests cover
// representative inputs; this sweep walks every committed corpus fixture
// through every supported mutation axis and asserts reader-shape parity
// (NOT byte-identity — the new layout differs from the legacy
// raw-rebuild path).
//
// Axes:
//   1. setCell           (one cell mutation per sheet)
//   2. appendRows        (append a single row of mixed cell types)
//   3. addSheet          (add a sheet with a unique name + 1 row)
//   4. deleteSheet       (delete the LAST sheet only — never sheet 0,
//                         never on a single-sheet fixture)
//   5. renameSheet       (rename sheet 0 to a unique-suffixed name)
//   6. insertRow         (insert at row 1)
//   7. deleteRow         (delete row 1)
//   8. insertColumn      (insert at col 1)
//   9. deleteColumn      (delete col 1)
//
// Refusal tolerance — these come from `docs/plans/refusal-audit.md` and
// are model-level invariants, not regressions. Each axis logs+skips on:
//   - error.RowEditUnsafeForSheet, error.ColEditUnsafeForSheet
//     (sheet has drawings, autoFilter, tableParts, frozen panes …)
//   - error.RowEditExceedsMaxRow, error.ColEditExceedsMaxCol
//   - error.RowEditRequiresCleanSheet
//   - error.SheetDeleteWithDefinedNamesNotSupported
//   - error.SheetDeleteRequiresCleanState
//   - error.CannotDeleteLastSheet, error.LastSheetUndeletable
//   - error.SheetHasUnsavedAppends, error.SheetHasUnsavedMutations
//   - error.InvalidCellRef (insertRow/deleteRow on completely empty sheets)
// Anything else is a real failure.

const corpus_fixtures = [_][]const u8{
    "frictionless_2sheets.xlsx",
    "openpyxl_guess_types.xlsx",
    "phpoi_test1.xlsx",
    "worldbank_catalog.xlsx",
};

const TmpFs = struct {
    dir: std.testing.TmpDir,
    pub fn init() TmpFs {
        return .{ .dir = std.testing.tmpDir(.{}) };
    }
    pub fn deinit(self: *TmpFs) void {
        self.dir.cleanup();
    }
    pub fn path(self: *TmpFs, alloc: std.mem.Allocator, name: []const u8) ![:0]u8 {
        const d = try self.dir.dir.realpathAlloc(alloc, ".");
        defer alloc.free(d);
        return std.fs.path.joinZ(alloc, &.{ d, name });
    }
};

fn corpusPath(alloc: std.mem.Allocator, fixture: []const u8) ![]u8 {
    return std.fmt.allocPrint(alloc, "{s}{s}", .{ corpus_dir, fixture });
}

fn fixtureExists(fixture: []const u8) bool {
    var path_buf: [256]u8 = undefined;
    const path = std.fmt.bufPrint(&path_buf, "{s}{s}", .{ corpus_dir, fixture }) catch return false;
    std.fs.cwd().access(path, .{}) catch return false;
    return true;
}

/// True iff `err` is a documented refusal we should log+skip for the
/// fixture/axis pair. Anything else propagates so the test fails.
fn isRefusalSkip(err: anyerror) bool {
    return switch (err) {
        error.RowEditUnsafeForSheet,
        error.ColEditUnsafeForSheet,
        error.RowEditRequiresCleanSheet,
        error.RowEditExceedsMaxRow,
        error.ColEditExceedsMaxCol,
        error.SheetDeleteWithDefinedNamesNotSupported,
        error.SheetDeleteRequiresCleanState,
        error.CannotDeleteLastSheet,
        error.LastSheetUndeletable,
        error.SheetHasUnsavedAppends,
        error.SheetHasUnsavedMutations,
        error.DuplicateSheetName,
        error.InvalidCellRef,
        => true,
        else => false,
    };
}

/// True iff `err` is an `Editor.open`-time ZIP-shape refusal. These
/// fixtures can't be loaded by `Editor` at all (separate contract from
/// the per-axis refusals tracked in `isRefusalSkip`). Skip the fixture
/// for ALL axes when this fires.
fn isOpenSkip(err: anyerror) bool {
    return switch (err) {
        error.ZipDataDescriptorNotSupported,
        error.ZipEncryptedNotSupported,
        error.Zip64NotSupported,
        error.ZipSplitNotSupported,
        error.ZipTooLarge,
        => true,
        else => false,
    };
}

/// Walk every sheet of `book` row-by-row; surfaces malformed-XML /
/// reader-shape regressions that the matrix path would mask.
fn walkAllSheets(book: *xlsx.Book, alloc: std.mem.Allocator) !void {
    for (book.sheets) |sheet| {
        var rows = try book.rows(sheet, alloc);
        defer rows.deinit();
        while (try rows.next()) |_| {}
    }
}

/// Count rows of `sheet` — used to assert pre/post deltas where the
/// axis predicts an exact change.
fn countRowsOf(book: *xlsx.Book, sheet: xlsx.Sheet, alloc: std.mem.Allocator) !usize {
    var rows = try book.rows(sheet, alloc);
    defer rows.deinit();
    var n: usize = 0;
    while (try rows.next()) |_| : (n += 1) {}
    return n;
}

/// Search every row of `sheet` for a cell whose `.string` matches
/// `needle`. Returns true on first hit. Used by setCell / appendRows /
/// addSheet / renameSheet to confirm the mutation reached the wire.
fn sheetContainsString(
    book: *xlsx.Book,
    sheet: xlsx.Sheet,
    alloc: std.mem.Allocator,
    needle: []const u8,
) !bool {
    var rows = try book.rows(sheet, alloc);
    defer rows.deinit();
    while (try rows.next()) |row| {
        for (row) |cell| switch (cell) {
            .string => |s| if (std.mem.eql(u8, s, needle)) return true,
            else => {},
        };
    }
    return false;
}

test "corpus parity: setCell on every fixture round-trips through reader" {
    const alloc = std.testing.allocator;
    var any_run: usize = 0;
    for (corpus_fixtures) |fixture| {
        if (!fixtureExists(fixture)) {
            std.debug.print("\n  [skip] {s} not in corpus\n", .{fixture});
            continue;
        }
        var tt = TmpFs.init();
        defer tt.deinit();
        const src = try corpusPath(alloc, fixture);
        defer alloc.free(src);
        const dst = try tt.path(alloc, "out.xlsx");
        defer alloc.free(dst);

        // setCell on row 1 col 0 of EVERY sheet — uses a unique sentinel
        // string per sheet so we can grep it back through the reader.
        const sentinel = "ZLSX_SETCELL_E7A_SENTINEL";
        var skip_fixture = false;
        open_block: {
            var ed = zlsx_pkg.Editor.open(alloc, src) catch |err| {
                if (isOpenSkip(err)) {
                    std.debug.print(
                        "\n  [skip open] {s}: {s}\n",
                        .{ fixture, @errorName(err) },
                    );
                    skip_fixture = true;
                    break :open_block;
                }
                return err;
            };
            defer ed.deinit();
            const sheet_count: u32 = @intCast(ed.workbook.sheetCount());
            var i: u32 = 0;
            while (i < sheet_count) : (i += 1) {
                ed.setCell(i, 1, 0, .{ .string = sentinel }) catch |err| {
                    if (isRefusalSkip(err)) {
                        std.debug.print(
                            "\n  [refuse-skip setCell] {s} sheet#{d}: {s}\n",
                            .{ fixture, i, @errorName(err) },
                        );
                        continue;
                    }
                    return err;
                };
            }
            try ed.save(dst);
        }
        if (skip_fixture) continue;

        var book = try xlsx.Book.open(alloc, dst);
        defer book.deinit();
        try walkAllSheets(&book, alloc);
        // Every sheet that accepted the setCell must surface the
        // sentinel; sheets that refused (above, via continue) won't.
        // Detecting "at least one hit" is enough for parity — we
        // already gated per-sheet errors above.
        var any_hit = false;
        for (book.sheets) |sheet| {
            if (try sheetContainsString(&book, sheet, alloc, sentinel)) {
                any_hit = true;
                break;
            }
        }
        try std.testing.expect(any_hit);
        any_run += 1;
    }
    if (any_run == 0) return error.SkipZigTest;
}

test "corpus parity: appendRows on every fixture round-trips through reader" {
    const alloc = std.testing.allocator;
    var any_run: usize = 0;
    for (corpus_fixtures) |fixture| {
        if (!fixtureExists(fixture)) {
            std.debug.print("\n  [skip] {s} not in corpus\n", .{fixture});
            continue;
        }
        var tt = TmpFs.init();
        defer tt.deinit();
        const src = try corpusPath(alloc, fixture);
        defer alloc.free(src);
        const dst = try tt.path(alloc, "out.xlsx");
        defer alloc.free(dst);

        const sentinel = "ZLSX_APPEND_E7A_SENTINEL";
        const append_row = [_]xlsx.Cell{
            .{ .string = sentinel },
            .{ .integer = 12345 },
            .{ .number = 6.5 },
            .{ .boolean = true },
        };
        const rows_to_append = [_][]const xlsx.Cell{&append_row};

        var per_sheet_ok = std.AutoHashMap(u32, void).init(alloc);
        defer per_sheet_ok.deinit();

        var skip_fixture = false;
        open_block: {
            var ed = zlsx_pkg.Editor.open(alloc, src) catch |err| {
                if (isOpenSkip(err)) {
                    std.debug.print(
                        "\n  [skip open] {s}: {s}\n",
                        .{ fixture, @errorName(err) },
                    );
                    skip_fixture = true;
                    break :open_block;
                }
                return err;
            };
            defer ed.deinit();
            const sheet_count: u32 = @intCast(ed.workbook.sheetCount());
            var i: u32 = 0;
            while (i < sheet_count) : (i += 1) {
                ed.appendRows(i, &rows_to_append) catch |err| {
                    if (isRefusalSkip(err)) {
                        std.debug.print(
                            "\n  [refuse-skip appendRows] {s} sheet#{d}: {s}\n",
                            .{ fixture, i, @errorName(err) },
                        );
                        continue;
                    }
                    return err;
                };
                try per_sheet_ok.put(i, {});
            }
            try ed.save(dst);
        }
        if (skip_fixture) continue;

        var book = try xlsx.Book.open(alloc, dst);
        defer book.deinit();
        try walkAllSheets(&book, alloc);

        // For each sheet that accepted the append, the sentinel must
        // be present somewhere in that sheet's body.
        var it = per_sheet_ok.keyIterator();
        while (it.next()) |idx_ptr| {
            const idx = idx_ptr.*;
            try std.testing.expect(idx < book.sheets.len);
            const sheet = book.sheets[idx];
            try std.testing.expect(try sheetContainsString(&book, sheet, alloc, sentinel));
        }
        any_run += 1;
    }
    if (any_run == 0) return error.SkipZigTest;
}

test "corpus parity: addSheet on every fixture round-trips through reader" {
    const alloc = std.testing.allocator;
    var any_run: usize = 0;
    for (corpus_fixtures) |fixture| {
        if (!fixtureExists(fixture)) {
            std.debug.print("\n  [skip] {s} not in corpus\n", .{fixture});
            continue;
        }
        var tt = TmpFs.init();
        defer tt.deinit();
        const src = try corpusPath(alloc, fixture);
        defer alloc.free(src);
        const dst = try tt.path(alloc, "out.xlsx");
        defer alloc.free(dst);

        const new_name = "ZlsxE7A_Added";
        const sentinel = "ZLSX_ADDSHEET_E7A_SENTINEL";
        const original_sheet_count = blk: {
            var b = try xlsx.Book.open(alloc, src);
            defer b.deinit();
            break :blk b.sheets.len;
        };

        var skip_fixture = false;
        open_block: {
            var ed = zlsx_pkg.Editor.open(alloc, src) catch |err| {
                if (isOpenSkip(err)) {
                    std.debug.print(
                        "\n  [skip open] {s}: {s}\n",
                        .{ fixture, @errorName(err) },
                    );
                    skip_fixture = true;
                    break :open_block;
                }
                return err;
            };
            defer ed.deinit();
            const new_idx = ed.addSheet(new_name) catch |err| {
                if (isRefusalSkip(err)) {
                    std.debug.print(
                        "\n  [refuse-skip addSheet] {s}: {s}\n",
                        .{ fixture, @errorName(err) },
                    );
                    continue;
                }
                return err;
            };
            // Write one row through the workbook fast-path. This goes
            // through `Worksheet.appendRows` for added sheets too
            // (post iter-er-4) — no separate code path.
            const append_row = [_]xlsx.Cell{
                .{ .string = sentinel },
                .{ .integer = 1 },
            };
            const rows_to_append = [_][]const xlsx.Cell{&append_row};
            try ed.appendRows(new_idx, &rows_to_append);
            try ed.save(dst);
        }
        if (skip_fixture) continue;

        var book = try xlsx.Book.open(alloc, dst);
        defer book.deinit();
        try walkAllSheets(&book, alloc);
        try std.testing.expectEqual(original_sheet_count + 1, book.sheets.len);
        const added = book.sheetByName(new_name) orelse return error.AddedSheetMissing;
        try std.testing.expect(try sheetContainsString(&book, added, alloc, sentinel));
        any_run += 1;
    }
    if (any_run == 0) return error.SkipZigTest;
}

test "corpus parity: deleteSheet (last sheet) on multi-sheet fixtures" {
    const alloc = std.testing.allocator;
    var any_run: usize = 0;
    var any_attempted: usize = 0;
    for (corpus_fixtures) |fixture| {
        if (!fixtureExists(fixture)) {
            std.debug.print("\n  [skip] {s} not in corpus\n", .{fixture});
            continue;
        }
        var tt = TmpFs.init();
        defer tt.deinit();
        const src = try corpusPath(alloc, fixture);
        defer alloc.free(src);
        const dst = try tt.path(alloc, "out.xlsx");
        defer alloc.free(dst);

        // Probe sheet count + capture last sheet's name BEFORE the
        // delete so we can verify it's gone post-save.
        const probe = blk: {
            var b = try xlsx.Book.open(alloc, src);
            defer b.deinit();
            const last_idx = b.sheets.len - 1;
            const last_name = try alloc.dupe(u8, b.sheets[last_idx].name);
            break :blk .{
                .count = b.sheets.len,
                .last_idx = @as(u32, @intCast(last_idx)),
                .last_name = last_name,
            };
        };
        defer alloc.free(probe.last_name);

        if (probe.count <= 1) {
            // Single-sheet fixtures — never delete the only sheet
            // (per task constraints + LastSheetUndeletable invariant).
            std.debug.print(
                "\n  [skip deleteSheet] {s}: single-sheet fixture\n",
                .{fixture},
            );
            continue;
        }

        any_attempted += 1;
        var skip_fixture = false;
        open_block: {
            var ed = zlsx_pkg.Editor.open(alloc, src) catch |err| {
                if (isOpenSkip(err)) {
                    std.debug.print(
                        "\n  [skip open] {s}: {s}\n",
                        .{ fixture, @errorName(err) },
                    );
                    skip_fixture = true;
                    break :open_block;
                }
                return err;
            };
            defer ed.deinit();
            ed.deleteSheet(probe.last_idx) catch |err| {
                if (isRefusalSkip(err)) {
                    std.debug.print(
                        "\n  [refuse-skip deleteSheet] {s} idx={d}: {s}\n",
                        .{ fixture, probe.last_idx, @errorName(err) },
                    );
                    continue;
                }
                return err;
            };
            try ed.save(dst);
        }
        if (skip_fixture) continue;

        var book = try xlsx.Book.open(alloc, dst);
        defer book.deinit();
        try walkAllSheets(&book, alloc);
        try std.testing.expectEqual(probe.count - 1, book.sheets.len);
        // The deleted sheet's name must no longer resolve.
        try std.testing.expect(book.sheetByName(probe.last_name) == null);
        any_run += 1;
    }
    if (any_attempted == 0) return error.SkipZigTest;
}

test "corpus parity: renameSheet on every fixture round-trips through reader" {
    const alloc = std.testing.allocator;
    var any_run: usize = 0;
    for (corpus_fixtures) |fixture| {
        if (!fixtureExists(fixture)) {
            std.debug.print("\n  [skip] {s} not in corpus\n", .{fixture});
            continue;
        }
        var tt = TmpFs.init();
        defer tt.deinit();
        const src = try corpusPath(alloc, fixture);
        defer alloc.free(src);
        const dst = try tt.path(alloc, "out.xlsx");
        defer alloc.free(dst);

        // Capture sheet 0's original name + its first row's first
        // string-or-numeric cell (the latter is just "are we still
        // reading the same data shape?").
        const original_name = blk: {
            var b = try xlsx.Book.open(alloc, src);
            defer b.deinit();
            break :blk try alloc.dupe(u8, b.sheets[0].name);
        };
        defer alloc.free(original_name);

        const new_name = "ZlsxE7A_Renamed_Sheet0";
        var skip_fixture = false;
        open_block: {
            var ed = zlsx_pkg.Editor.open(alloc, src) catch |err| {
                if (isOpenSkip(err)) {
                    std.debug.print(
                        "\n  [skip open] {s}: {s}\n",
                        .{ fixture, @errorName(err) },
                    );
                    skip_fixture = true;
                    break :open_block;
                }
                return err;
            };
            defer ed.deinit();
            ed.renameSheet(0, new_name) catch |err| {
                if (isRefusalSkip(err)) {
                    std.debug.print(
                        "\n  [refuse-skip renameSheet] {s}: {s}\n",
                        .{ fixture, @errorName(err) },
                    );
                    continue;
                }
                return err;
            };
            try ed.save(dst);
        }
        if (skip_fixture) continue;

        var book = try xlsx.Book.open(alloc, dst);
        defer book.deinit();
        try walkAllSheets(&book, alloc);
        try std.testing.expectEqualStrings(new_name, book.sheets[0].name);
        // The old name must no longer resolve unless another sheet
        // shared it (none of the corpus fixtures do, but keep the
        // assertion soft if equality drifts).
        if (!std.mem.eql(u8, original_name, new_name)) {
            try std.testing.expect(book.sheetByName(original_name) == null);
        }
        any_run += 1;
    }
    if (any_run == 0) return error.SkipZigTest;
}

test "corpus parity: insertRow on every fixture round-trips through reader" {
    const alloc = std.testing.allocator;
    var any_attempted: usize = 0;
    for (corpus_fixtures) |fixture| {
        if (!fixtureExists(fixture)) {
            std.debug.print("\n  [skip] {s} not in corpus\n", .{fixture});
            continue;
        }
        var tt = TmpFs.init();
        defer tt.deinit();
        const src = try corpusPath(alloc, fixture);
        defer alloc.free(src);
        const dst = try tt.path(alloc, "out.xlsx");
        defer alloc.free(dst);

        const pre_count = blk: {
            var b = try xlsx.Book.open(alloc, src);
            defer b.deinit();
            // Sheet 0 row count for the post-save assertion.
            break :blk try countRowsOf(&b, b.sheets[0], alloc);
        };

        any_attempted += 1;
        var skip_fixture = false;
        open_block: {
            var ed = zlsx_pkg.Editor.open(alloc, src) catch |err| {
                if (isOpenSkip(err)) {
                    std.debug.print(
                        "\n  [skip open] {s}: {s}\n",
                        .{ fixture, @errorName(err) },
                    );
                    skip_fixture = true;
                    break :open_block;
                }
                return err;
            };
            defer ed.deinit();
            ed.insertRow(0, 1) catch |err| {
                if (isRefusalSkip(err)) {
                    std.debug.print(
                        "\n  [refuse-skip insertRow] {s}: {s}\n",
                        .{ fixture, @errorName(err) },
                    );
                    continue;
                }
                return err;
            };
            try ed.save(dst);
        }
        if (skip_fixture) continue;

        var book = try xlsx.Book.open(alloc, dst);
        defer book.deinit();
        try walkAllSheets(&book, alloc);
        // Inserted row at row 1 has no <row> element — readable row
        // count is unchanged in the typical case. Allow ±1 slack so
        // fixtures whose insert path materialises a placeholder
        // <row> (today: none in the corpus, but allow drift) still
        // pass shape-parity.
        const post_count = try countRowsOf(&book, book.sheets[0], alloc);
        try std.testing.expect(post_count >= pre_count and post_count <= pre_count + 1);
    }
    if (any_attempted == 0) return error.SkipZigTest;
}

test "corpus parity: deleteRow on every fixture round-trips through reader" {
    const alloc = std.testing.allocator;
    var any_attempted: usize = 0;
    for (corpus_fixtures) |fixture| {
        if (!fixtureExists(fixture)) {
            std.debug.print("\n  [skip] {s} not in corpus\n", .{fixture});
            continue;
        }
        var tt = TmpFs.init();
        defer tt.deinit();
        const src = try corpusPath(alloc, fixture);
        defer alloc.free(src);
        const dst = try tt.path(alloc, "out.xlsx");
        defer alloc.free(dst);

        const pre_count = blk: {
            var b = try xlsx.Book.open(alloc, src);
            defer b.deinit();
            break :blk try countRowsOf(&b, b.sheets[0], alloc);
        };

        any_attempted += 1;
        var accepted = false;
        var skip_fixture = false;
        open_block: {
            var ed = zlsx_pkg.Editor.open(alloc, src) catch |err| {
                if (isOpenSkip(err)) {
                    std.debug.print(
                        "\n  [skip open] {s}: {s}\n",
                        .{ fixture, @errorName(err) },
                    );
                    skip_fixture = true;
                    break :open_block;
                }
                return err;
            };
            defer ed.deinit();
            ed.deleteRow(0, 1) catch |err| {
                if (isRefusalSkip(err)) {
                    std.debug.print(
                        "\n  [refuse-skip deleteRow] {s}: {s}\n",
                        .{ fixture, @errorName(err) },
                    );
                    continue;
                }
                return err;
            };
            try ed.save(dst);
            accepted = true;
        }
        if (skip_fixture) continue;
        if (!accepted) continue;

        var book = try xlsx.Book.open(alloc, dst);
        defer book.deinit();
        try walkAllSheets(&book, alloc);
        const post_count = try countRowsOf(&book, book.sheets[0], alloc);
        // Deleting a populated row drops one entry from the iterator
        // (the row has cells; not just a style-only `<row r="N"/>`).
        // Allow ±1 slack to absorb fixtures whose row 1 happens to be
        // empty (no <row> element) — deleteRow on those is a no-op
        // shape-wise but still legal.
        try std.testing.expect(post_count == pre_count or post_count + 1 == pre_count);
    }
    if (any_attempted == 0) return error.SkipZigTest;
}

test "corpus parity: insertColumn on every fixture round-trips through reader" {
    const alloc = std.testing.allocator;
    var any_attempted: usize = 0;
    for (corpus_fixtures) |fixture| {
        if (!fixtureExists(fixture)) {
            std.debug.print("\n  [skip] {s} not in corpus\n", .{fixture});
            continue;
        }
        var tt = TmpFs.init();
        defer tt.deinit();
        const src = try corpusPath(alloc, fixture);
        defer alloc.free(src);
        const dst = try tt.path(alloc, "out.xlsx");
        defer alloc.free(dst);

        any_attempted += 1;
        var accepted = false;
        var skip_fixture = false;
        open_block: {
            var ed = zlsx_pkg.Editor.open(alloc, src) catch |err| {
                if (isOpenSkip(err)) {
                    std.debug.print(
                        "\n  [skip open] {s}: {s}\n",
                        .{ fixture, @errorName(err) },
                    );
                    skip_fixture = true;
                    break :open_block;
                }
                return err;
            };
            defer ed.deinit();
            ed.insertColumn(0, 1) catch |err| {
                if (isRefusalSkip(err)) {
                    std.debug.print(
                        "\n  [refuse-skip insertColumn] {s}: {s}\n",
                        .{ fixture, @errorName(err) },
                    );
                    continue;
                }
                return err;
            };
            try ed.save(dst);
            accepted = true;
        }
        if (skip_fixture) continue;
        if (!accepted) continue;

        var book = try xlsx.Book.open(alloc, dst);
        defer book.deinit();
        try walkAllSheets(&book, alloc);
    }
    if (any_attempted == 0) return error.SkipZigTest;
}

test "corpus parity: deleteColumn on every fixture round-trips through reader" {
    const alloc = std.testing.allocator;
    var any_attempted: usize = 0;
    for (corpus_fixtures) |fixture| {
        if (!fixtureExists(fixture)) {
            std.debug.print("\n  [skip] {s} not in corpus\n", .{fixture});
            continue;
        }
        var tt = TmpFs.init();
        defer tt.deinit();
        const src = try corpusPath(alloc, fixture);
        defer alloc.free(src);
        const dst = try tt.path(alloc, "out.xlsx");
        defer alloc.free(dst);

        any_attempted += 1;
        var accepted = false;
        var skip_fixture = false;
        open_block: {
            var ed = zlsx_pkg.Editor.open(alloc, src) catch |err| {
                if (isOpenSkip(err)) {
                    std.debug.print(
                        "\n  [skip open] {s}: {s}\n",
                        .{ fixture, @errorName(err) },
                    );
                    skip_fixture = true;
                    break :open_block;
                }
                return err;
            };
            defer ed.deinit();
            ed.deleteColumn(0, 1) catch |err| {
                if (isRefusalSkip(err)) {
                    std.debug.print(
                        "\n  [refuse-skip deleteColumn] {s}: {s}\n",
                        .{ fixture, @errorName(err) },
                    );
                    continue;
                }
                return err;
            };
            try ed.save(dst);
            accepted = true;
        }
        if (skip_fixture) continue;
        if (!accepted) continue;

        var book = try xlsx.Book.open(alloc, dst);
        defer book.deinit();
        try walkAllSheets(&book, alloc);
    }
    if (any_attempted == 0) return error.SkipZigTest;
}
