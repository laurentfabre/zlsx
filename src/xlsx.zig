//! Read-only `.xlsx` parser — just enough to walk rows of a sheet.
//!
//! Motivated by Alfred's `classify_pdfs` port: we need to read 1,000+ rows
//! of a single worksheet from a Python-generated xlsx (via openpyxl). Not
//! a full Office Open XML implementation — no styles, no formulas, no
//! writes, no multi-sheet streaming, no charts, no comments.
//!
//! Public surface:
//!   * `Book`    — opens a file, decompresses the few XML parts we care
//!                 about, resolves shared strings + sheet targets.
//!   * `Sheet`   — (name, path) pair exposed on `Book.sheets`.
//!   * `Rows`    — iterator over a sheet's rows; yields dense `[]Cell`.
//!   * `Cell`    — tagged union: `.empty | .string | .integer | .number | .boolean`.
//!
//! Ownership: every string slice returned from `Rows.next()` is borrowed
//! from the `Book`'s internal buffers. The `Book` must outlive any cells
//! the caller is still reading.
//!
//! Tested against `data/output/alfred_bdr_prospect_list_1000.xlsx`
//! (openpyxl 3.x, single sheet, 1,007 rows × 35 cols, inline strings +
//! shared strings + integer/float/boolean cells).

const std = @import("std");
const Allocator = std.mem.Allocator;

// ─── Public types ────────────────────────────────────────────────────

/// A date/time decoded from an Excel serial-date number. Columns in
/// xlsx files don't carry enough metadata on the cell value alone to
/// distinguish a plain number from a date (the date-ness lives in the
/// number-format applied via styling, which `Rows.next()` intentionally
/// doesn't surface), so callers who know a column holds dates convert
/// explicitly via `fromExcelSerial(cell.number)`.
///
/// Fractional seconds are rounded to the nearest second.
pub const DateTime = struct {
    year: u16, // 1900..=9999
    month: u8, // 1..=12
    day: u8, // 1..=31
    hour: u8, // 0..=23
    minute: u8, // 0..=59
    second: u8, // 0..=59
};

/// Days from 1970-01-01 → Gregorian `{year, month, day}` using
/// Howard Hinnant's civil_from_days algorithm (proleptic Gregorian,
/// valid through year 9999). Epoch-shift to 0000-03-01 (adding 719468)
/// keeps the arithmetic positive for all in-range Excel dates.
fn daysSinceUnixEpochToYMD(days_since_unix: i32) struct { year: u16, month: u8, day: u8 } {
    const z: i32 = days_since_unix + 719468;
    const era: i32 = if (z >= 0) @divFloor(z, 146097) else @divFloor(z - 146096, 146097);
    const doe: u32 = @intCast(z - era * 146097);
    const yoe: u32 = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    const y: i32 = @as(i32, @intCast(yoe)) + era * 400;
    const doy: u32 = doe - (365 * yoe + yoe / 4 - yoe / 100);
    const mp: u32 = (5 * doy + 2) / 153;
    const d: u32 = doy - (153 * mp + 2) / 5 + 1;
    const m: u32 = if (mp < 10) mp + 3 else mp - 9;
    const year: i32 = if (m <= 2) y + 1 else y;
    return .{
        .year = @intCast(year),
        .month = @intCast(m),
        .day = @intCast(d),
    };
}

/// Convert an Excel serial-date number (as produced by `cell.number`
/// when the source cell was styled as a date) into a calendar
/// `DateTime`. Returns `null` on:
///   - non-finite input (`NaN` / `±inf`)
///   - serials < 61 — those fall inside Excel's 1900 leap-year bug
///     window where serials 1..=60 displayably include the fictitious
///     1900-02-29, and the correct decoded date is ambiguous without
///     an opt-in flag
///   - serials > 2958465 (past 9999-12-31)
///
/// Serials 61..=2958465 cover 1900-03-01 through 9999-12-31 — the
/// overwhelming majority of real-world dates — and decode against
/// the proleptic Gregorian calendar.
pub fn fromExcelSerial(serial: f64) ?DateTime {
    if (!std.math.isFinite(serial)) return null;
    if (serial < 61.0 or serial >= 2958466.0) return null;

    const floored = @floor(serial);
    // Excel serial 1 = 1900-01-01, but the 1900 leap-year bug shifts
    // serials ≥ 60 forward by one day vs the real Gregorian calendar.
    // Serial 25569 corresponds to 1970-01-01; subtracting it yields
    // days-since-Unix-epoch for the Hinnant algorithm.
    const days_since_unix: i32 = @intCast(@as(i64, @intFromFloat(floored)) - 25569);
    const ymd = daysSinceUnixEpochToYMD(days_since_unix);

    // Time of day: fractional part × 86 400 seconds, rounded.
    // Clamp to 86399 so a rounding edge like `1.99999999` never
    // produces hour=24 / minute=60 / second=60; we lose one second
    // of resolution there.
    const total_s_f: f64 = @round((serial - floored) * 86400.0);
    const total_s: i64 = @intFromFloat(@min(total_s_f, 86399.0));
    return .{
        .year = ymd.year,
        .month = ymd.month,
        .day = ymd.day,
        .hour = @intCast(@divFloor(total_s, 3600)),
        .minute = @intCast(@mod(@divFloor(total_s, 60), 60)),
        .second = @intCast(@mod(total_s, 60)),
    };
}

pub const Cell = union(enum) {
    empty,
    /// Slice into the Book's internal buffers — not owned by Cell.
    string: []const u8,
    integer: i64,
    /// Non-integer number. `integer` is preferred when the value is a
    /// round integer because xlsx stores everything as text under `<v>`
    /// and many "integer" columns round-trip through float.
    number: f64,
    boolean: bool,
};

pub const Sheet = struct {
    name: []const u8,
    path: []const u8, // e.g. "xl/worksheets/sheet1.xml"
};

/// A1-style cell reference broken into components. Column is 0-based
/// (A=0, B=1, …, XFD=16383); row is 1-based to match Excel's native
/// convention (row 1 is the first row). Matches the axis layout used
/// by `Rows.next()`'s returned cell slices.
pub const CellRef = struct { col: u32, row: u32 };

/// A rectangular merged cell range read from the worksheet's
/// `<mergeCells>` block. `top_left` has the smaller `(col, row)` pair;
/// `bottom_right` has the larger. Both corners are inclusive.
pub const MergeRange = struct {
    top_left: CellRef,
    bottom_right: CellRef,
};

/// A hyperlink attached to a cell or cell range. Two flavours:
///   - **External**: `url` holds the resolved `Target` from the
///     sheet's `_rels/sheet{N}.xml.rels` file; `location` is empty.
///   - **Internal**: `location` holds the raw `location` attribute
///     (e.g. `Sheet2!A1`); `url` is empty.
/// Exactly one of the two is non-empty per entry. Range corners
/// follow `MergeRange` normalisation; `top_left == bottom_right`
/// for single-cell hyperlinks.
pub const Hyperlink = struct {
    top_left: CellRef,
    bottom_right: CellRef,
    /// External URL target. Slice into the sheet's rels XML buffer —
    /// valid for the Book's lifetime. Empty for internal hyperlinks.
    url: []const u8,
    /// Internal (same-workbook) target, e.g. `Sheet2!A1`. Slice into
    /// the sheet XML buffer. Empty for external hyperlinks.
    location: []const u8,
};

/// Data-validation kind, mirroring OOXML `ST_DataValidationType`.
/// `.unknown` covers any type we don't recognise (forward-compat with
/// generators that introduce new variants) — callers can still read
/// `formula1` / `formula2` verbatim.
pub const DataValidationKind = enum {
    list,
    whole,
    decimal,
    date,
    time,
    text_length,
    custom,
    unknown,
};

/// Numeric comparison operator, mirroring OOXML
/// `ST_DataValidationOperator`. Null when the source validation omits
/// `operator=` (list / custom validations never use one).
pub const DataValidationOperator = enum {
    between,
    not_between,
    equal,
    not_equal,
    less_than,
    less_than_or_equal,
    greater_than,
    greater_than_or_equal,
};

/// A data-validation entry parsed from a sheet's `<dataValidations>`
/// block. All kinds surface — list callers get `values` plus a
/// literal formula1; numeric / date / time / length / custom callers
/// get `kind`, `op`, `formula1`, and (for between / not_between)
/// `formula2`. Strings in `values`, `formula1`, and `formula2` are
/// entity-decoded and owned by the Book.
pub const DataValidation = struct {
    top_left: CellRef,
    bottom_right: CellRef,
    kind: DataValidationKind = .list,
    /// Null when `kind` is `.list`, `.custom`, `.unknown`, or when the
    /// source XML omits `operator=` (Excel defaults such omissions to
    /// `.between` but we preserve the absence so round-trips are exact).
    op: ?DataValidationOperator = null,
    /// First formula content, entity-decoded. For `.list` this is the
    /// literal `"a,b,c"` wrapped CSV or a range reference
    /// (`$A$1:$A$10`); the parsed CSV form is in `values`. Empty when
    /// the source had no `<formula1>` element.
    formula1: []const u8 = "",
    /// Second formula content, entity-decoded. Empty unless the
    /// validation uses `.between` or `.not_between`.
    formula2: []const u8 = "",
    /// Dropdown options. Empty when the source validation isn't a
    /// list or when `formula1` is a range reference (which we don't
    /// expand to the referenced cells on read).
    values: []const []const u8 = &.{},
};

/// Domain errors surfaced by this module. The public API uses inferred
/// error sets so callers get the full union of these plus whatever
/// std.zip / std.fs / std.compress.flate decide to return.
/// A single formatting run inside a shared-string entry. Excel emits
/// rich-text via `<si><r><rPr>...</rPr><t>...</t></r>...</si>` where
/// every `<r>` can carry its own font properties. Only entries with
/// at least one `<r>` wrapper produce runs — plain `<si><t>...</t></si>`
/// entries return null from `Book.richRuns`.
///
/// `text` borrows from `sst_arena` (decoded if entities, else a slice
/// directly into the raw SST xml). `bold` and `italic` are the two
/// properties this reader currently surfaces — other `<rPr>` children
/// (color, size, font) are skipped today and will be added in a
/// follow-up iter without breaking this shape.
pub const RichRun = struct {
    text: []const u8,
    bold: bool = false,
    italic: bool = false,
};

pub const DomainError = error{
    NotAnXlsx,
    BadZip,
    MissingWorkbook,
    MissingSheet,
    UnsupportedCompression,
    MalformedXml,
};

// ─── Book ────────────────────────────────────────────────────────────

pub const Book = struct {
    allocator: Allocator,
    /// Decompressed `xl/sharedStrings.xml` (nullable — small xlsx files
    /// with only inline strings omit this part).
    shared_strings_xml: ?[]u8 = null,
    /// Index into `shared_strings_xml` (or into `sst_arena` if we had
    /// to decode entities / concatenate rich-text runs). Order matches
    /// the SST in the file.
    shared_strings: [][]const u8 = &.{},
    /// Bump arena for SST-owned strings. Bulk-freed on `deinit`.
    /// One arena reset per parseSharedStrings call is all the cleanup
    /// we need — saves ~1 malloc per entry vs the previous per-string
    /// ArrayList path on workloads like worldbank_catalog (1,144 SST).
    sst_arena: std.heap.ArenaAllocator,
    /// (name, path) for each `<sheet>` in the workbook, in declared order.
    sheets: []Sheet = &.{},
    /// Decompressed bytes of each sheet's XML, keyed by path.
    sheet_data: std.StringHashMapUnmanaged([]u8) = .{},
    /// Merged cell ranges per sheet, parsed from `<mergeCells>` at
    /// open time. Keyed by the sheet path. Sheets without merges are
    /// absent from the map; callers should use
    /// `mergedRanges(sheet)` which normalises the missing case to an
    /// empty slice.
    merged_ranges: std.StringHashMapUnmanaged([]MergeRange) = .{},
    /// Decompressed per-sheet `_rels/sheet{N}.xml.rels` files, keyed
    /// by the *sheet* path (not the rels path) so callers never have
    /// to compute the rels filename themselves. Drives hyperlink
    /// URL resolution.
    sheet_rels_data: std.StringHashMapUnmanaged([]u8) = .{},
    /// Hyperlinks per sheet, resolved at open time by cross-referencing
    /// `<hyperlinks>` in the sheet XML against the matching rels file.
    /// Keyed by the sheet path. Sheets without hyperlinks are absent;
    /// use `hyperlinks(sheet)` for the missing-sheet normalisation.
    hyperlinks_by_sheet: std.StringHashMapUnmanaged([]Hyperlink) = .{},
    /// List-type data validations per sheet, parsed from the
    /// `<dataValidations>` block. Same keyed-by-sheet-path shape as
    /// `hyperlinks_by_sheet`; use `dataValidations(sheet)` for the
    /// empty-on-missing normalisation.
    data_validations_by_sheet: std.StringHashMapUnmanaged([]DataValidation) = .{},
    /// Rich-text runs per shared-string index, keyed by SST position.
    /// Only populated for entries that used `<r>` wrappers (plain
    /// `<si><t>...</t></si>` entries skip the map entirely). Run
    /// slices are owned by `sst_arena`; run text slices either
    /// borrow from the raw SST xml or live in `sst_arena`.
    rich_runs_by_sst_idx: std.AutoHashMapUnmanaged(usize, []RichRun) = .{},
    /// Owned backing storage for every string referenced by `sheets`,
    /// sheet_data keys, and entity-decoded shared strings.
    strings: std.ArrayListUnmanaged([]u8) = .{},

    /// Open and parse the workbook skeleton. Sheet XML is eagerly
    /// decompressed (xlsx files we target are small — ~300 KB — and
    /// streaming through std.zip is awkward).
    pub fn open(allocator: Allocator, path: []const u8) !Book {
        var book: Book = .{
            .allocator = allocator,
            .sst_arena = std.heap.ArenaAllocator.init(allocator),
        };
        errdefer book.deinit();

        var file = try std.fs.cwd().openFile(path, .{});
        defer file.close();

        var buf: [4096]u8 = undefined;
        var file_reader = file.reader(&buf);

        var iter = std.zip.Iterator.init(&file_reader) catch return error.BadZip;

        // We need three categories of files from the archive:
        //   xl/sharedStrings.xml     — optional
        //   xl/workbook.xml          — required
        //   xl/_rels/workbook.xml.rels — required (sheet id → target path)
        //   xl/worksheets/sheet*.xml — data
        var rels_xml: ?[]u8 = null;
        defer if (rels_xml) |r| allocator.free(r);

        var workbook_xml: ?[]u8 = null;
        defer if (workbook_xml) |w| allocator.free(w);

        // Track pending sheet paths we still have to extract (collected
        // from the CDFH walk; resolved after we've parsed workbook.xml).
        // We just extract every xl/worksheets/*.xml we see — cheaper
        // than a second pass.
        var filename_buf: [512]u8 = undefined;

        while (iter.next() catch return error.BadZip) |entry| {
            if (entry.filename_len == 0 or entry.filename_len > filename_buf.len) continue;

            // Read filename (lives in the CDFH after the fixed-size header).
            try file_reader.seekTo(entry.header_zip_offset + @sizeOf(std.zip.CentralDirectoryFileHeader));
            const filename = filename_buf[0..entry.filename_len];
            file_reader.interface.readSliceAll(filename) catch return error.BadZip;

            if (std.mem.eql(u8, filename, "xl/sharedStrings.xml")) {
                book.shared_strings_xml = try extractEntryToBuffer(allocator, entry, &file_reader);
            } else if (std.mem.eql(u8, filename, "xl/workbook.xml")) {
                workbook_xml = try extractEntryToBuffer(allocator, entry, &file_reader);
            } else if (std.mem.eql(u8, filename, "xl/_rels/workbook.xml.rels")) {
                rels_xml = try extractEntryToBuffer(allocator, entry, &file_reader);
            } else if (std.mem.startsWith(u8, filename, "xl/worksheets/_rels/") and
                std.mem.endsWith(u8, filename, ".xml.rels"))
            {
                // Per-sheet rels (hyperlinks, drawings, etc.). Key the
                // entry by the *sheet* path so the hyperlink resolver
                // can do a direct lookup:
                //   "xl/worksheets/_rels/sheet1.xml.rels"
                //     → "xl/worksheets/sheet1.xml"
                const rels_prefix = "xl/worksheets/_rels/".len;
                const rels_suffix = ".rels".len;
                const bare = filename[rels_prefix .. filename.len - rels_suffix];
                const sheet_key = try std.fmt.allocPrint(allocator, "xl/worksheets/{s}", .{bare});
                errdefer allocator.free(sheet_key);
                try book.strings.append(allocator, sheet_key);
                const data = try extractEntryToBuffer(allocator, entry, &file_reader);
                try book.sheet_rels_data.put(allocator, sheet_key, data);
            } else if (std.mem.startsWith(u8, filename, "xl/worksheets/") and
                std.mem.endsWith(u8, filename, ".xml"))
            {
                // Own the key outright so the HashMap entry stays valid
                // for the Book's lifetime.
                const key = try allocator.dupe(u8, filename);
                errdefer allocator.free(key);
                try book.strings.append(allocator, key);
                const data = try extractEntryToBuffer(allocator, entry, &file_reader);
                try book.sheet_data.put(allocator, key, data);
            }
        }

        const wb = workbook_xml orelse return error.MissingWorkbook;
        const rels = rels_xml orelse return error.MissingWorkbook;

        try parseWorkbookSheets(&book, wb, rels);
        if (book.shared_strings_xml) |sst| try parseSharedStrings(&book, sst);

        // Parse merged ranges + hyperlinks + data validations per
        // sheet. All three are cheap — most sheets have none of them,
        // and each parser bails on the first miss.
        var sheet_it = book.sheet_data.iterator();
        while (sheet_it.next()) |entry| {
            try parseMergedRangesForSheet(&book, entry.key_ptr.*, entry.value_ptr.*);
            try parseHyperlinksForSheet(&book, entry.key_ptr.*, entry.value_ptr.*);
            try parseDataValidationsForSheet(&book, entry.key_ptr.*, entry.value_ptr.*);
        }

        return book;
    }

    /// Merged cell ranges declared in this sheet's `<mergeCells>`
    /// block, or an empty slice if none. The returned slice is owned
    /// by the Book and valid until `deinit`.
    pub fn mergedRanges(self: *const Book, sheet: Sheet) []const MergeRange {
        return self.merged_ranges.get(sheet.path) orelse &.{};
    }

    /// External-URL hyperlinks declared on this sheet, with each
    /// `r:id` resolved through the sheet's `_rels` file. Returns an
    /// empty slice for sheets without a `<hyperlinks>` block or
    /// without a matching rels file. The returned slice + `url`
    /// strings inside it are owned by the Book and valid until
    /// `deinit`.
    pub fn hyperlinks(self: *const Book, sheet: Sheet) []const Hyperlink {
        return self.hyperlinks_by_sheet.get(sheet.path) orelse &.{};
    }

    /// List-type data validations (dropdowns) declared on this sheet.
    /// Returns an empty slice for sheets without a `<dataValidations>`
    /// block or with only non-list variants. The returned slice,
    /// every inner `values` slice, and each value string are owned
    /// by the Book and valid until `deinit`.
    pub fn dataValidations(self: *const Book, sheet: Sheet) []const DataValidation {
        return self.data_validations_by_sheet.get(sheet.path) orelse &.{};
    }

    /// Rich-text runs for a shared-string entry, or null when the
    /// entry is a plain single-run string (the common case —
    /// spreadsheets without inline formatting never enter the rich
    /// path and pay zero overhead). Returned slices and text lifetimes
    /// match the Book.
    pub fn richRuns(self: *const Book, sst_idx: usize) ?[]const RichRun {
        return self.rich_runs_by_sst_idx.get(sst_idx);
    }

    pub fn deinit(self: *Book) void {
        const a = self.allocator;
        if (self.shared_strings_xml) |s| a.free(s);
        a.free(self.shared_strings);
        self.sst_arena.deinit();

        var mit = self.merged_ranges.valueIterator();
        while (mit.next()) |v| a.free(v.*);
        self.merged_ranges.deinit(a);

        var srit = self.sheet_rels_data.valueIterator();
        while (srit.next()) |v| a.free(v.*);
        self.sheet_rels_data.deinit(a);

        var hit = self.hyperlinks_by_sheet.valueIterator();
        while (hit.next()) |v| a.free(v.*);
        self.hyperlinks_by_sheet.deinit(a);

        var dvit = self.data_validations_by_sheet.valueIterator();
        while (dvit.next()) |v| {
            for (v.*) |dv| a.free(dv.values);
            a.free(v.*);
        }
        self.data_validations_by_sheet.deinit(a);

        // Rich-run slices live in sst_arena (already freed above); only
        // the hashmap spine is on the general allocator.
        self.rich_runs_by_sst_idx.deinit(a);

        var it = self.sheet_data.valueIterator();
        while (it.next()) |v| a.free(v.*);
        self.sheet_data.deinit(a);

        a.free(self.sheets);
        for (self.strings.items) |s| a.free(s);
        self.strings.deinit(a);
        self.* = undefined;
    }

    /// Find the sheet by display name (case-sensitive). Returns null if
    /// the workbook doesn't have it.
    pub fn sheetByName(self: *const Book, name: []const u8) ?Sheet {
        for (self.sheets) |s| {
            if (std.mem.eql(u8, s.name, name)) return s;
        }
        return null;
    }

    /// Iterator over the sheet's rows. Yields dense `[]Cell` (one slot
    /// per column present in the widest cell reference in that row, with
    /// gaps filled by `.empty`). Caller retains ownership of the
    /// internal row buffer via `Rows.deinit()`.
    pub fn rows(self: *const Book, sheet: Sheet, allocator: Allocator) !Rows {
        const xml = self.sheet_data.get(sheet.path) orelse return error.MissingSheet;
        return .{
            .xml = xml,
            .pos = 0,
            .shared_strings = self.shared_strings,
            .allocator = allocator,
            .row_cells = .{},
            .arena = std.heap.ArenaAllocator.init(allocator),
        };
    }
};

// ─── Rows iterator ───────────────────────────────────────────────────

pub const Rows = struct {
    xml: []const u8,
    pos: usize,
    shared_strings: []const []const u8,
    allocator: Allocator,
    row_cells: std.ArrayListUnmanaged(Cell),
    /// Bump arena for per-row decoded strings. Reset (O(1)) at the
    /// start of each `next()` call — previous row's owned strings
    /// become invalid, which matches the documented contract. Compared
    /// to the older per-string malloc/free list this saves ~one free
    /// per entity-bearing or rich-text cell per row.
    arena: std.heap.ArenaAllocator,

    pub fn deinit(self: *Rows) void {
        self.arena.deinit();
        self.row_cells.deinit(self.allocator);
        self.* = undefined;
    }

    /// Returns the next row's cells, or null at end-of-sheet. Returned
    /// slice is valid until the next call to `next()` (or until
    /// `deinit()`). Cell string contents are either shared-string slices
    /// (owned by the Book), xml-backed slices (stable for the Book's
    /// lifetime), or row-owned slices that are invalidated on the next
    /// call (arena reset).
    pub fn next(self: *Rows) !?[]const Cell {
        self.row_cells.clearRetainingCapacity();
        _ = self.arena.reset(.retain_capacity);

        while (findTagOpen(self.xml, self.pos, "row")) |row_start| {
            self.pos = row_start.after_open;
            // Consume cells until </row>.
            try self.consumeRow();
            return self.row_cells.items;
        }
        return null;
    }

    fn consumeRow(self: *Rows) !void {
        while (true) {
            // Look for <c ... />, <c ...>...</c>, or </row>.
            const next_lt = std.mem.indexOfScalarPos(u8, self.xml, self.pos, '<') orelse return error.MalformedXml;
            self.pos = next_lt;
            if (std.mem.startsWith(u8, self.xml[self.pos..], "</row>")) {
                self.pos += "</row>".len;
                return;
            }
            if (std.mem.startsWith(u8, self.xml[self.pos..], "<c")) {
                try self.consumeCell();
            } else {
                // Skip unknown tag
                const end = std.mem.indexOfScalarPos(u8, self.xml, self.pos, '>') orelse return error.MalformedXml;
                self.pos = end + 1;
            }
        }
    }

    fn consumeCell(self: *Rows) !void {
        // Parse `<c r="A3" t="s" s="1">…</c>` or `<c r="A3"/>`
        const tag_end_rel = std.mem.indexOfAnyPos(u8, self.xml, self.pos, "/>") orelse return error.MalformedXml;
        _ = tag_end_rel;
        const gt = std.mem.indexOfScalarPos(u8, self.xml, self.pos, '>') orelse return error.MalformedXml;
        const is_self_closing = gt > 0 and self.xml[gt - 1] == '/';
        const attrs = self.xml[self.pos + 2 .. if (is_self_closing) gt - 1 else gt];

        const r_attr = getAttr(attrs, "r") orelse return error.MalformedXml;
        const col_idx = try columnIndexFromRef(r_attr);
        const cell_type = getAttr(attrs, "t") orelse "n"; // default numeric

        // Grow row_cells to cover col_idx; fill gaps with .empty.
        while (self.row_cells.items.len <= col_idx) {
            try self.row_cells.append(self.allocator, .empty);
        }

        if (is_self_closing) {
            // Empty cell; already .empty.
            self.pos = gt + 1;
            return;
        }
        self.pos = gt + 1;

        // Parse body until </c>.
        // For t="inlineStr" the body is <is><t>text</t></is>; otherwise
        // it's <v>text</v> (maybe preceded by <f>formula</f>, which we
        // ignore).
        const cell_close = "</c>";
        const body_end = std.mem.indexOfPos(u8, self.xml, self.pos, cell_close) orelse return error.MalformedXml;
        const body = self.xml[self.pos..body_end];
        self.pos = body_end + cell_close.len;

        const cell: Cell = if (std.mem.eql(u8, cell_type, "inlineStr"))
            .{ .string = try self.decodeInlineString(body) }
        else if (std.mem.eql(u8, cell_type, "s"))
            try self.resolveSharedString(body)
        else if (std.mem.eql(u8, cell_type, "str"))
            .{ .string = try self.decodeVValue(body) }
        else if (std.mem.eql(u8, cell_type, "b")) blk: {
            // Empty `<v></v>` is still valid XML; treat it as false rather
            // than crashing on the [0] index.
            const raw = extractVValue(body) orelse "";
            break :blk .{ .boolean = raw.len > 0 and raw[0] == '1' };
        } else if (std.mem.eql(u8, cell_type, "e"))
            .{ .string = try self.decodeVValue(body) }
        else
            try parseNumericCell(extractVValue(body) orelse "");

        self.row_cells.items[col_idx] = cell;
    }

    fn resolveSharedString(self: *Rows, body: []const u8) !Cell {
        const idx_text = extractVValue(body) orelse return error.MalformedXml;
        const idx = std.fmt.parseInt(u32, idx_text, 10) catch return error.MalformedXml;
        if (idx >= self.shared_strings.len) return error.MalformedXml;
        return .{ .string = self.shared_strings[idx] };
    }

    /// Decode a `<v>…</v>` cell body. If the raw text has no entities,
    /// return the xml-backed slice directly (no allocation). Otherwise
    /// allocate an owned slice in the row arena.
    fn decodeVValue(self: *Rows, body: []const u8) ![]const u8 {
        const raw = extractVValue(body) orelse return "";
        return try self.internOrBorrow(raw);
    }

    /// Decode inline-string body `<is>(<r>)?<t>text</t>(</r>)?</is>`.
    /// Single-`<t>` bodies without entities borrow from xml. Anything
    /// else (rich-text runs, entities) gets an owned allocation in the
    /// row arena.
    fn decodeInlineString(self: *Rows, body: []const u8) ![]const u8 {
        // Count <t> runs to decide borrow vs own.
        var t_count: usize = 0;
        var probe: usize = 0;
        var only_start: usize = 0;
        var only_end: usize = 0;
        var has_entities = false;
        while (std.mem.indexOfPos(u8, body, probe, "<t")) |t_start| {
            const gt = std.mem.indexOfScalarPos(u8, body, t_start, '>') orelse return error.MalformedXml;
            if (body[gt - 1] == '/') {
                probe = gt + 1;
                continue;
            }
            const t_close = std.mem.indexOfPos(u8, body, gt + 1, "</t>") orelse return error.MalformedXml;
            const span = body[gt + 1 .. t_close];
            if (std.mem.indexOfScalar(u8, span, '&') != null) has_entities = true;
            if (t_count == 0) {
                only_start = gt + 1;
                only_end = t_close;
            }
            t_count += 1;
            probe = t_close + "</t>".len;
        }

        if (t_count == 0) return "";
        if (t_count == 1 and !has_entities) {
            return body[only_start..only_end];
        }

        // Multi-run or entity-bearing — allocate into the row arena.
        const a = self.arena.allocator();
        var buf: std.ArrayListUnmanaged(u8) = .{};
        var i: usize = 0;
        while (std.mem.indexOfPos(u8, body, i, "<t")) |t_start| {
            const gt = std.mem.indexOfScalarPos(u8, body, t_start, '>') orelse return error.MalformedXml;
            if (body[gt - 1] == '/') {
                i = gt + 1;
                continue;
            }
            const t_close = std.mem.indexOfPos(u8, body, gt + 1, "</t>") orelse return error.MalformedXml;
            try appendDecoded(a, &buf, body[gt + 1 .. t_close]);
            i = t_close + "</t>".len;
        }
        return try buf.toOwnedSlice(a);
    }

    /// Return `raw` unchanged if it needs no decoding; otherwise allocate
    /// an owned slice in the row arena (invalidated at the next `next()` call).
    fn internOrBorrow(self: *Rows, raw: []const u8) ![]const u8 {
        if (std.mem.indexOfScalar(u8, raw, '&') == null) return raw;
        const a = self.arena.allocator();
        var buf: std.ArrayListUnmanaged(u8) = .{};
        try appendDecoded(a, &buf, raw);
        return try buf.toOwnedSlice(a);
    }
};

// ─── Helpers: XML + column math ──────────────────────────────────────

/// Returns the index of the next `<tag` occurrence and the offset just
/// after the `<tag>` opening. Null if none.
const TagOpen = struct { start: usize, after_open: usize };

fn findTagOpen(xml: []const u8, from: usize, tag: []const u8) ?TagOpen {
    var i = from;
    while (std.mem.indexOfPos(u8, xml, i, "<")) |lt| {
        const after_lt = lt + 1;
        if (after_lt + tag.len <= xml.len and std.mem.eql(u8, xml[after_lt .. after_lt + tag.len], tag)) {
            // Must be followed by `/`, `>`, space, or `/>` — i.e. a whole tag, not a prefix.
            const next_c = if (after_lt + tag.len < xml.len) xml[after_lt + tag.len] else return null;
            if (next_c == ' ' or next_c == '\t' or next_c == '>' or next_c == '/') {
                const gt = std.mem.indexOfScalarPos(u8, xml, lt, '>') orelse return null;
                return .{ .start = lt, .after_open = gt + 1 };
            }
        }
        i = after_lt;
    }
    return null;
}

/// Extract a string attribute value from an attribute region (everything
/// between `<tag` and `>`). Returns the quoted value verbatim (no entity
/// decoding — attribute values in the xlsx files we care about don't
/// carry entities beyond shared_strings).
fn getAttr(attrs: []const u8, name: []const u8) ?[]const u8 {
    var i: usize = 0;
    while (i < attrs.len) {
        // Skip whitespace
        while (i < attrs.len and std.ascii.isWhitespace(attrs[i])) i += 1;
        if (i >= attrs.len) break;
        // Scan attribute name
        const name_start = i;
        while (i < attrs.len and attrs[i] != '=' and !std.ascii.isWhitespace(attrs[i])) i += 1;
        const attr_name = attrs[name_start..i];
        // Skip `=` and optional whitespace, then quote
        while (i < attrs.len and (attrs[i] == '=' or std.ascii.isWhitespace(attrs[i]))) i += 1;
        if (i >= attrs.len or (attrs[i] != '"' and attrs[i] != '\'')) break;
        const quote = attrs[i];
        i += 1;
        const val_start = i;
        while (i < attrs.len and attrs[i] != quote) i += 1;
        const val = attrs[val_start..i];
        if (i < attrs.len) i += 1; // past closing quote
        if (std.mem.eql(u8, attr_name, name)) return val;
    }
    return null;
}

fn extractVValue(body: []const u8) ?[]const u8 {
    const v_open = std.mem.indexOf(u8, body, "<v>") orelse return null;
    const v_close = std.mem.indexOfPos(u8, body, v_open + 3, "</v>") orelse return null;
    return body[v_open + 3 .. v_close];
}

fn parseNumericCell(text: []const u8) !Cell {
    if (text.len == 0) return .empty;
    // Integer first (tight range for Alfred's use case)
    if (std.fmt.parseInt(i64, text, 10)) |n| {
        return .{ .integer = n };
    } else |_| {}
    if (std.fmt.parseFloat(f64, text)) |f| {
        return .{ .number = f };
    } else |_| {
        return .{ .string = text };
    }
}

/// "A1" → 0, "B3" → 1, "AA10" → 26, "AB10" → 27.
fn columnIndexFromRef(ref: []const u8) !usize {
    var idx: usize = 0;
    var i: usize = 0;
    while (i < ref.len and ref[i] >= 'A' and ref[i] <= 'Z') : (i += 1) {
        idx = idx * 26 + (ref[i] - 'A' + 1);
    }
    if (i == 0) return error.MalformedXml;
    return idx - 1;
}

// ─── Entity decoding ─────────────────────────────────────────────────

fn appendDecoded(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    text: []const u8,
) !void {
    var i: usize = 0;
    while (i < text.len) {
        const amp = std.mem.indexOfScalarPos(u8, text, i, '&');
        if (amp == null) {
            try out.appendSlice(allocator, text[i..]);
            return;
        }
        try out.appendSlice(allocator, text[i..amp.?]);
        i = amp.?;
        const semi = std.mem.indexOfScalarPos(u8, text, i, ';') orelse return error.MalformedXml;
        const entity = text[i + 1 .. semi];
        if (std.mem.eql(u8, entity, "amp")) {
            try out.append(allocator, '&');
        } else if (std.mem.eql(u8, entity, "lt")) {
            try out.append(allocator, '<');
        } else if (std.mem.eql(u8, entity, "gt")) {
            try out.append(allocator, '>');
        } else if (std.mem.eql(u8, entity, "quot")) {
            try out.append(allocator, '"');
        } else if (std.mem.eql(u8, entity, "apos")) {
            try out.append(allocator, '\'');
        } else if (entity.len > 1 and entity[0] == '#') {
            const cp: u21 = if (entity[1] == 'x')
                std.fmt.parseInt(u21, entity[2..], 16) catch return error.MalformedXml
            else
                std.fmt.parseInt(u21, entity[1..], 10) catch return error.MalformedXml;
            var utf8_buf: [4]u8 = undefined;
            const n = std.unicode.utf8Encode(cp, &utf8_buf) catch return error.MalformedXml;
            try out.appendSlice(allocator, utf8_buf[0..n]);
        } else {
            // Unknown entity — preserve verbatim rather than erroring out.
            try out.append(allocator, '&');
            try out.appendSlice(allocator, entity);
            try out.append(allocator, ';');
        }
        i = semi + 1;
    }
}

// ─── workbook.xml + rels parsing ─────────────────────────────────────

fn parseWorkbookSheets(book: *Book, wb_xml: []const u8, rels_xml: []const u8) !void {
    // Parse rels first into an id → target map.
    var rel_map: std.StringHashMapUnmanaged([]const u8) = .{};
    defer rel_map.deinit(book.allocator);

    var i: usize = 0;
    while (std.mem.indexOfPos(u8, rels_xml, i, "<Relationship")) |rel_start| {
        const gt = std.mem.indexOfScalarPos(u8, rels_xml, rel_start, '>') orelse break;
        const attrs = rels_xml[rel_start + "<Relationship".len .. gt];
        const id = getAttr(attrs, "Id") orelse {
            i = gt + 1;
            continue;
        };
        const target = getAttr(attrs, "Target") orelse {
            i = gt + 1;
            continue;
        };
        try rel_map.put(book.allocator, id, target);
        i = gt + 1;
    }

    // Walk <sheet name="..." r:id="..."/> in workbook.xml.
    var sheets: std.ArrayListUnmanaged(Sheet) = .{};
    errdefer sheets.deinit(book.allocator);

    i = 0;
    while (std.mem.indexOfPos(u8, wb_xml, i, "<sheet")) |sh_start| {
        // Must be a standalone `<sheet` tag — `<sheets>` also matches.
        const next_c = if (sh_start + "<sheet".len < wb_xml.len) wb_xml[sh_start + "<sheet".len] else break;
        if (next_c != ' ' and next_c != '\t' and next_c != '/') {
            i = sh_start + 1;
            continue;
        }
        const gt = std.mem.indexOfScalarPos(u8, wb_xml, sh_start, '>') orelse break;
        const attrs = wb_xml[sh_start + "<sheet".len .. gt];
        const name = getAttr(attrs, "name") orelse {
            i = gt + 1;
            continue;
        };
        const rid = getAttr(attrs, "r:id") orelse {
            i = gt + 1;
            continue;
        };
        const target = rel_map.get(rid) orelse {
            i = gt + 1;
            continue;
        };

        // Path is relative to xl/ — prepend if not absolute. Own it.
        var path_buf: std.ArrayListUnmanaged(u8) = .{};
        errdefer path_buf.deinit(book.allocator);
        if (std.mem.startsWith(u8, target, "/")) {
            try path_buf.appendSlice(book.allocator, target[1..]);
        } else {
            try path_buf.appendSlice(book.allocator, "xl/");
            try path_buf.appendSlice(book.allocator, target);
        }
        const path = try path_buf.toOwnedSlice(book.allocator);
        try book.strings.append(book.allocator, path);

        // Name needs entity decoding (hotels with & in their names).
        var name_buf: std.ArrayListUnmanaged(u8) = .{};
        errdefer name_buf.deinit(book.allocator);
        try appendDecoded(book.allocator, &name_buf, name);
        const name_decoded = try name_buf.toOwnedSlice(book.allocator);
        try book.strings.append(book.allocator, name_decoded);

        try sheets.append(book.allocator, .{ .name = name_decoded, .path = path });
        i = gt + 1;
    }

    book.sheets = try sheets.toOwnedSlice(book.allocator);
}

// ─── sharedStrings.xml parsing ───────────────────────────────────────

/// Parse one corner of an A1-style reference ("B12") into `{col, row}`.
/// Column is 0-based (A=0, B=1, …), row is 1-based (row1=1). Rejects
/// empty input, lowercase, missing digits, and row=0.
fn parseA1Ref(s: []const u8) !CellRef {
    if (s.len == 0) return error.MalformedXml;
    var i: usize = 0;
    var col: u32 = 0;
    while (i < s.len and s[i] >= 'A' and s[i] <= 'Z') : (i += 1) {
        col = col * 26 + (s[i] - 'A' + 1);
    }
    if (i == 0 or i == s.len) return error.MalformedXml;
    var row: u32 = 0;
    while (i < s.len and s[i] >= '0' and s[i] <= '9') : (i += 1) {
        row = row * 10 + (s[i] - '0');
    }
    if (i != s.len or row == 0) return error.MalformedXml;
    return .{ .col = col - 1, .row = row };
}

/// Parse an A1-style range ("A1:B2"). Top-left corner must precede or
/// equal bottom-right on both axes; we normalise order so callers can
/// rely on `top_left ≤ bottom_right` component-wise even if the source
/// XML listed the corners in the other order.
fn parseA1Range(ref: []const u8) !MergeRange {
    const colon = std.mem.indexOfScalar(u8, ref, ':') orelse {
        // Single-cell "range" — degenerate but legal per the reader's
        // contract. Promote to a 1×1 rectangle.
        const p = try parseA1Ref(ref);
        return .{ .top_left = p, .bottom_right = p };
    };
    const a = try parseA1Ref(ref[0..colon]);
    const b = try parseA1Ref(ref[colon + 1 ..]);
    return .{
        .top_left = .{ .col = @min(a.col, b.col), .row = @min(a.row, b.row) },
        .bottom_right = .{ .col = @max(a.col, b.col), .row = @max(a.row, b.row) },
    };
}

/// Walk a sheet XML's `<mergeCells>` section and collect ranges into
/// `book.merged_ranges`. No-op when the sheet has no merges.
fn parseMergedRangesForSheet(book: *Book, sheet_path: []const u8, xml: []const u8) !void {
    const mc_start = std.mem.indexOf(u8, xml, "<mergeCells") orelse return;
    const mc_end = std.mem.indexOfPos(u8, xml, mc_start, "</mergeCells>") orelse return;
    const block = xml[mc_start..mc_end];

    var ranges: std.ArrayListUnmanaged(MergeRange) = .{};
    errdefer ranges.deinit(book.allocator);

    // Honour the `count="N"` attribute when present — lets us
    // pre-size the backing array and skip geometric grows.
    const hint: usize = blk: {
        const c = std.mem.indexOfPos(u8, block, 0, "count=\"") orelse break :blk 0;
        const start = c + "count=\"".len;
        const end = std.mem.indexOfScalarPos(u8, block, start, '"') orelse break :blk 0;
        break :blk std.fmt.parseInt(usize, block[start..end], 10) catch 0;
    };
    if (hint > 0) try ranges.ensureTotalCapacity(book.allocator, hint);

    const needle = "<mergeCell ref=\"";
    var i: usize = 0;
    while (std.mem.indexOfPos(u8, block, i, needle)) |start| {
        const ref_start = start + needle.len;
        const ref_end = std.mem.indexOfScalarPos(u8, block, ref_start, '"') orelse break;
        // Skip malformed ranges rather than fail the whole open —
        // callers with partially-valid workbooks should still get the
        // rest of their data.
        if (parseA1Range(block[ref_start..ref_end])) |r| {
            try ranges.append(book.allocator, r);
        } else |_| {}
        i = ref_end + 1;
    }

    if (ranges.items.len == 0) {
        ranges.deinit(book.allocator);
        return;
    }
    const slice = try ranges.toOwnedSlice(book.allocator);
    errdefer book.allocator.free(slice);
    try book.merged_ranges.put(book.allocator, sheet_path, slice);
}

/// Resolve a rels id (e.g. "rId3") to its `Target` attribute in the
/// given rels XML. Returns a slice into `rels_xml` so the caller
/// doesn't need to copy. `null` when the id isn't present (e.g. the
/// sheet XML references an id that the rels file omits — malformed,
/// skip the entry).
fn findRelTarget(rels_xml: []const u8, rid: []const u8) ?[]const u8 {
    // Find each `<Relationship ... Id="rid" ... Target="..."/>`. We do a
    // linear scan keyed on `Id="…"`; the rels files are tiny (few KB
    // at most) so no smarter index is needed.
    var probe: usize = 0;
    while (std.mem.indexOfPos(u8, rels_xml, probe, "<Relationship")) |rel_start| {
        const rel_end = std.mem.indexOfScalarPos(u8, rels_xml, rel_start, '>') orelse return null;
        const attrs = rels_xml[rel_start..rel_end];
        // Check Id="rid".
        const id_key = "Id=\"";
        const id_pos = std.mem.indexOf(u8, attrs, id_key) orelse {
            probe = rel_end + 1;
            continue;
        };
        const id_start = id_pos + id_key.len;
        const id_close = std.mem.indexOfScalarPos(u8, attrs, id_start, '"') orelse {
            probe = rel_end + 1;
            continue;
        };
        if (!std.mem.eql(u8, attrs[id_start..id_close], rid)) {
            probe = rel_end + 1;
            continue;
        }
        // Matched — pull out Target="…".
        const tgt_key = "Target=\"";
        const tgt_pos = std.mem.indexOfPos(u8, attrs, id_close, tgt_key) orelse return null;
        const tgt_start = tgt_pos + tgt_key.len;
        const tgt_close = std.mem.indexOfScalarPos(u8, attrs, tgt_start, '"') orelse return null;
        return attrs[tgt_start..tgt_close];
    }
    return null;
}

/// Walk a sheet XML's `<hyperlinks>` section, cross-reference each
/// `r:id` against the sheet's rels file, and collect resolved entries
/// into `book.hyperlinks_by_sheet`. No-op when the sheet has no
/// hyperlinks or no rels file.
fn parseHyperlinksForSheet(book: *Book, sheet_path: []const u8, xml: []const u8) !void {
    const hl_start = std.mem.indexOf(u8, xml, "<hyperlinks") orelse return;
    const hl_end = std.mem.indexOfPos(u8, xml, hl_start, "</hyperlinks>") orelse return;
    const block = xml[hl_start..hl_end];

    // A sheet may carry both external (r:id → rels → Target) and
    // internal (location="…") hyperlinks; the rels file is only
    // required for the former. Fall through with `null` when the
    // sheet has no per-sheet rels — internal entries still parse.
    const rels_xml: ?[]const u8 = book.sheet_rels_data.get(sheet_path);

    var entries: std.ArrayListUnmanaged(Hyperlink) = .{};
    errdefer entries.deinit(book.allocator);

    var probe: usize = 0;
    while (std.mem.indexOfPos(u8, block, probe, "<hyperlink")) |hl| {
        const end = std.mem.indexOfScalarPos(u8, block, hl, '>') orelse break;
        const attrs = block[hl..end];
        probe = end + 1;

        const ref_key = "ref=\"";
        const ref_pos = std.mem.indexOf(u8, attrs, ref_key) orelse continue;
        const ref_start = ref_pos + ref_key.len;
        const ref_close = std.mem.indexOfScalarPos(u8, attrs, ref_start, '"') orelse continue;
        const ref = attrs[ref_start..ref_close];

        const range = parseA1Range(ref) catch continue;

        // Prefer external (r:id) when both are present — a valid OOXML
        // entry has one or the other, but defensive ordering keeps us
        // working on workbooks that generate both by mistake.
        var url: []const u8 = "";
        var location: []const u8 = "";
        if (std.mem.indexOf(u8, attrs, "r:id=\"")) |rid_pos| {
            const rid_start = rid_pos + "r:id=\"".len;
            const rid_close = std.mem.indexOfScalarPos(u8, attrs, rid_start, '"') orelse continue;
            const rid = attrs[rid_start..rid_close];
            if (rels_xml) |rx| {
                url = findRelTarget(rx, rid) orelse continue;
            } else {
                continue;
            }
        } else if (std.mem.indexOf(u8, attrs, "location=\"")) |loc_pos| {
            const loc_start = loc_pos + "location=\"".len;
            const loc_close = std.mem.indexOfScalarPos(u8, attrs, loc_start, '"') orelse continue;
            location = attrs[loc_start..loc_close];
        } else {
            // Neither r:id nor location — malformed, skip.
            continue;
        }

        try entries.append(book.allocator, .{
            .top_left = range.top_left,
            .bottom_right = range.bottom_right,
            .url = url,
            .location = location,
        });
    }

    if (entries.items.len == 0) {
        entries.deinit(book.allocator);
        return;
    }
    const slice = try entries.toOwnedSlice(book.allocator);
    errdefer book.allocator.free(slice);
    try book.hyperlinks_by_sheet.put(book.allocator, sheet_path, slice);
}

/// Split a list-type validation's formula1 content into dropdown
/// values. Excel wraps the joined CSV in double-quotes (either raw
/// `"` or XML-escaped `&quot;`). Values are entity-decoded into
/// `book.sst_arena` so they live for the Book's lifetime. Returns
/// a freshly allocated `[][]const u8` owned by `book.allocator` (to
/// be stored inside `DataValidation.values`), or `null` when the
/// formula1 isn't a literal-list form (it's a range reference or
/// malformed).
fn splitFormula1List(book: *Book, formula1: []const u8) !?[][]const u8 {
    // Accept either literal-quote or XML-escaped-quote wrapping.
    const trimmed = if (std.mem.startsWith(u8, formula1, "&quot;") and
        std.mem.endsWith(u8, formula1, "&quot;"))
        formula1[6 .. formula1.len - 6]
    else if (std.mem.startsWith(u8, formula1, "\"") and
        std.mem.endsWith(u8, formula1, "\""))
        formula1[1 .. formula1.len - 1]
    else
        return null;
    if (trimmed.len == 0) return null;

    var out: std.ArrayListUnmanaged([]const u8) = .{};
    errdefer out.deinit(book.allocator);

    const arena = book.sst_arena.allocator();
    var start: usize = 0;
    var i: usize = 0;
    while (i <= trimmed.len) : (i += 1) {
        const at_end = i == trimmed.len;
        if (at_end or trimmed[i] == ',') {
            const raw = trimmed[start..i];
            // Entity-decode into the SST arena so callers get clean strings.
            if (std.mem.indexOfScalar(u8, raw, '&') == null) {
                try out.append(book.allocator, raw);
            } else {
                var buf: std.ArrayListUnmanaged(u8) = try .initCapacity(arena, raw.len);
                try appendDecoded(arena, &buf, raw);
                try out.append(book.allocator, try buf.toOwnedSlice(arena));
            }
            start = i + 1;
        }
    }

    return try out.toOwnedSlice(book.allocator);
}

fn parseDvKind(s: []const u8) DataValidationKind {
    if (std.mem.eql(u8, s, "list")) return .list;
    if (std.mem.eql(u8, s, "whole")) return .whole;
    if (std.mem.eql(u8, s, "decimal")) return .decimal;
    if (std.mem.eql(u8, s, "date")) return .date;
    if (std.mem.eql(u8, s, "time")) return .time;
    if (std.mem.eql(u8, s, "textLength")) return .text_length;
    if (std.mem.eql(u8, s, "custom")) return .custom;
    return .unknown;
}

fn parseDvOperator(s: []const u8) ?DataValidationOperator {
    if (std.mem.eql(u8, s, "between")) return .between;
    if (std.mem.eql(u8, s, "notBetween")) return .not_between;
    if (std.mem.eql(u8, s, "equal")) return .equal;
    if (std.mem.eql(u8, s, "notEqual")) return .not_equal;
    if (std.mem.eql(u8, s, "lessThan")) return .less_than;
    if (std.mem.eql(u8, s, "lessThanOrEqual")) return .less_than_or_equal;
    if (std.mem.eql(u8, s, "greaterThan")) return .greater_than;
    if (std.mem.eql(u8, s, "greaterThanOrEqual")) return .greater_than_or_equal;
    return null;
}

/// Return the text between `<tag>` and `</tag>` inside `body`, or
/// null when the element is absent. No decoding — raw XML text is
/// returned so callers can decide whether to keep entity-escaped
/// form (e.g. list parser) or decode (e.g. formula1/formula2 surface).
fn extractElementContent(body: []const u8, tag: []const u8) ?[]const u8 {
    var open_buf: [32]u8 = undefined;
    var close_buf: [32]u8 = undefined;
    const open = std.fmt.bufPrint(&open_buf, "<{s}>", .{tag}) catch return null;
    const close = std.fmt.bufPrint(&close_buf, "</{s}>", .{tag}) catch return null;
    const o = std.mem.indexOf(u8, body, open) orelse return null;
    const start = o + open.len;
    const c = std.mem.indexOfPos(u8, body, start, close) orelse return null;
    return body[start..c];
}

/// Decode XML entities in `raw` and return a slice owned by the Book
/// (fresh allocation in `sst_arena` when decoding is needed, otherwise
/// a zero-copy slice into the sheet XML itself).
fn decodeFormulaInto(book: *Book, raw: []const u8) ![]const u8 {
    if (std.mem.indexOfScalar(u8, raw, '&') == null) return raw;
    const arena = book.sst_arena.allocator();
    var buf: std.ArrayListUnmanaged(u8) = try .initCapacity(arena, raw.len);
    try appendDecoded(arena, &buf, raw);
    return try buf.toOwnedSlice(arena);
}

/// Walk a sheet XML's `<dataValidations>` section and collect every
/// entry into `book.data_validations_by_sheet`, surfacing kind,
/// operator, formula1 and formula2 so non-list validations round-trip.
fn parseDataValidationsForSheet(book: *Book, sheet_path: []const u8, xml: []const u8) !void {
    const dv_start = std.mem.indexOf(u8, xml, "<dataValidations") orelse return;
    const dv_end = std.mem.indexOfPos(u8, xml, dv_start, "</dataValidations>") orelse return;
    const block = xml[dv_start..dv_end];

    var entries: std.ArrayListUnmanaged(DataValidation) = .{};
    errdefer entries.deinit(book.allocator);

    var i: usize = 0;
    while (std.mem.indexOfPos(u8, block, i, "<dataValidation")) |dv| {
        const hdr_end = std.mem.indexOfScalarPos(u8, block, dv, '>') orelse break;
        const attrs = block[dv..hdr_end];
        i = hdr_end + 1;

        // sqref="…"
        const sqref_key = "sqref=\"";
        const sq_pos = std.mem.indexOf(u8, attrs, sqref_key) orelse continue;
        const sq_start = sq_pos + sqref_key.len;
        const sq_close = std.mem.indexOfScalarPos(u8, attrs, sq_start, '"') orelse continue;
        const sqref = attrs[sq_start..sq_close];
        const r = parseA1Range(sqref) catch continue;

        // type="…" → DataValidationKind. Excel omits type for list by
        // default in some generators, so "list" is the implicit fallback.
        var kind: DataValidationKind = .list;
        const type_key = "type=\"";
        if (std.mem.indexOf(u8, attrs, type_key)) |tp| {
            const t_start = tp + type_key.len;
            const t_close = std.mem.indexOfScalarPos(u8, attrs, t_start, '"') orelse continue;
            kind = parseDvKind(attrs[t_start..t_close]);
        }

        // operator="…" → DataValidationOperator (null when absent).
        var op: ?DataValidationOperator = null;
        const op_key = "operator=\"";
        if (std.mem.indexOf(u8, attrs, op_key)) |op_pos| {
            const o_start = op_pos + op_key.len;
            const o_close = std.mem.indexOfScalarPos(u8, attrs, o_start, '"') orelse continue;
            op = parseDvOperator(attrs[o_start..o_close]);
        }

        var values: []const []const u8 = &.{};
        var formula1: []const u8 = "";
        var formula2: []const u8 = "";

        // If self-closing `<dataValidation … />` there's no body.
        const is_self_closing = hdr_end > 0 and block[hdr_end - 1] == '/';
        if (!is_self_closing) {
            const dv_close = std.mem.indexOfPos(u8, block, hdr_end, "</dataValidation>") orelse continue;
            const body = block[hdr_end + 1 .. dv_close];
            i = dv_close + "</dataValidation>".len;

            if (extractElementContent(body, "formula1")) |raw_f1| {
                formula1 = try decodeFormulaInto(book, raw_f1);
                if (kind == .list) {
                    if (try splitFormula1List(book, raw_f1)) |parsed| {
                        values = parsed;
                    }
                }
            }
            if (extractElementContent(body, "formula2")) |raw_f2| {
                formula2 = try decodeFormulaInto(book, raw_f2);
            }
        }

        try entries.append(book.allocator, .{
            .top_left = r.top_left,
            .bottom_right = r.bottom_right,
            .kind = kind,
            .op = op,
            .formula1 = formula1,
            .formula2 = formula2,
            .values = values,
        });
    }

    if (entries.items.len == 0) {
        entries.deinit(book.allocator);
        return;
    }
    const slice = try entries.toOwnedSlice(book.allocator);
    errdefer book.allocator.free(slice);
    try book.data_validations_by_sheet.put(book.allocator, sheet_path, slice);
}

/// Scan a slice of `<rPr>...</rPr>` content for the two font flags
/// this reader currently surfaces (bold / italic). `<b/>` and `<i/>`
/// are self-closing in every OOXML generator I've checked; explicit
/// `val="false"` is rare but honoured. Returns defaults when the
/// slice is empty or the flags are absent.
fn parseRprFlags(rpr: []const u8) RichRun {
    var bold = false;
    var italic = false;
    // Match `<b/>`, `<b ...>`, `<b val="1"/>`, etc. Skip `<b val="0"/>`.
    if (std.mem.indexOf(u8, rpr, "<b")) |bp| {
        if (bp + 2 < rpr.len) {
            const next = rpr[bp + 2];
            if (next == '/' or next == '>' or next == ' ') {
                bold = !hasFalseVal(rpr[bp..]);
            }
        }
    }
    if (std.mem.indexOf(u8, rpr, "<i")) |ip| {
        if (ip + 2 < rpr.len) {
            const next = rpr[ip + 2];
            if (next == '/' or next == '>' or next == ' ') {
                italic = !hasFalseVal(rpr[ip..]);
            }
        }
    }
    return .{ .text = "", .bold = bold, .italic = italic };
}

/// Returns true if the tag body (up to the next `>`) contains
/// `val="0"` or `val="false"`. OOXML treats missing val as true.
fn hasFalseVal(tag: []const u8) bool {
    const gt = std.mem.indexOfScalar(u8, tag, '>') orelse return false;
    const body = tag[0..gt];
    return std.mem.indexOf(u8, body, "val=\"0\"") != null or
        std.mem.indexOf(u8, body, "val=\"false\"") != null;
}

fn parseSharedStrings(book: *Book, sst_xml: []u8) !void {
    // Pre-size via the uniqueCount hint in the <sst> tag when present —
    // OOXML generators are reliable about this attribute, so we get a
    // right-sized backing store on the first append and skip all
    // geometric grows. Fall back to 64 for generators that omit it.
    const hint: usize = blk: {
        const open = std.mem.indexOf(u8, sst_xml, "<sst") orelse break :blk 64;
        const gt = std.mem.indexOfScalarPos(u8, sst_xml, open, '>') orelse break :blk 64;
        const attrs = sst_xml[open..gt];
        const u = std.mem.indexOf(u8, attrs, "uniqueCount=\"") orelse break :blk 64;
        const start = u + "uniqueCount=\"".len;
        const end = std.mem.indexOfScalarPos(u8, attrs, start, '"') orelse break :blk 64;
        break :blk std.fmt.parseInt(usize, attrs[start..end], 10) catch 64;
    };

    var strings: std.ArrayListUnmanaged([]const u8) = .{};
    errdefer strings.deinit(book.allocator);
    try strings.ensureTotalCapacity(book.allocator, hint);
    const arena_alloc = book.sst_arena.allocator();

    // Single-pass byte walker driven by `indexOfScalarPos('<')` — the
    // SIMD-accelerated scalar scan is meaningfully faster than
    // `indexOfPos` on 2-3 byte needles. Each `<` hit peeks 1-2 bytes
    // to identify the tag type, then either consumes it (updating
    // the local accumulator) or advances past it.
    //
    // CRITICAL — every code path must strictly advance either `i`
    // (outer) or `j` (inner) on each iteration, including on
    // malformed input. The earlier attempt at this rewrite lacked
    // an `i` advance on the "inner loop exhausted body without
    // finding </si>" path, which hung the fuzz. The explicit
    // `i = si_gt + 1; continue :outer` below is that fix.
    const xml = sst_xml;
    var i: usize = 0;

    outer: while (i < xml.len) {
        const i_prev = i;

        // Find next `<si` — scalar `<` scan + 2-byte peek.
        const lt = std.mem.indexOfScalarPos(u8, xml, i, '<') orelse break;
        if (lt + 3 > xml.len) break;
        if (xml[lt + 1] != 's' or xml[lt + 2] != 'i') {
            i = lt + 1;
            std.debug.assert(i > i_prev);
            continue;
        }
        // Opening `<si` tag end — `<si>` or `<si/>` or `<si attr="…">`.
        const si_gt = std.mem.indexOfScalarPos(u8, xml, lt + 3, '>') orelse break;
        if (si_gt > 0 and xml[si_gt - 1] == '/') {
            try strings.append(book.allocator, "");
            i = si_gt + 1;
            std.debug.assert(i > i_prev);
            continue;
        }

        // Walk body — commit to one of three finishing paths without
        // re-scanning:
        //   t_count == 0  → ""
        //   t_count == 1  → borrow (no entities) or decode-into-arena
        //   t_count ≥ 2   → multi-run; buf collects the decoded concat
        var t_count: usize = 0;
        var first_span: []const u8 = "";
        var first_has_ent = false;
        var buf: std.ArrayListUnmanaged(u8) = .empty;

        // Rich-text tracking — populated only when we see a `<r>` tag.
        // `pending_flags` holds formatting from the most recent `<rPr>`
        // within the current `<r>` block; applied to the next `<t>`
        // encountered. Reset at each new `<r>`.
        var runs: std.ArrayListUnmanaged(RichRun) = .empty;
        var saw_r = false;
        var pending_flags: RichRun = .{ .text = "" };

        const sst_idx = strings.items.len;

        var j: usize = si_gt + 1;
        body: while (j < xml.len) {
            const j_prev = j;
            const next_lt = std.mem.indexOfScalarPos(u8, xml, j, '<') orelse break :body;
            if (next_lt + 2 > xml.len) break :body;
            const c1 = xml[next_lt + 1];

            if (c1 == 't' and (next_lt + 2 == xml.len or
                xml[next_lt + 2] == '>' or xml[next_lt + 2] == ' ' or xml[next_lt + 2] == '/'))
            {
                // `<t>`, `<t attr="…">`, or `<t/>`.
                const t_gt = std.mem.indexOfScalarPos(u8, xml, next_lt + 2, '>') orelse break :body;
                if (t_gt > 0 and xml[t_gt - 1] == '/') {
                    j = t_gt + 1;
                    std.debug.assert(j > j_prev);
                    continue :body;
                }
                const t_close = std.mem.indexOfPos(u8, xml, t_gt + 1, "</t>") orelse break :body;
                const span = xml[t_gt + 1 .. t_close];
                const span_has_ent = std.mem.indexOfScalar(u8, span, '&') != null;

                if (t_count == 0) {
                    first_span = span;
                    first_has_ent = span_has_ent;
                } else {
                    // Promote to `buf` on the 2nd <t>. Seed with the
                    // decoded first span, then append the current span.
                    if (buf.capacity == 0) {
                        const cap: usize = first_span.len + span.len + 8;
                        buf = try .initCapacity(arena_alloc, cap);
                        if (first_has_ent) {
                            try appendDecoded(arena_alloc, &buf, first_span);
                        } else {
                            buf.appendSliceAssumeCapacity(first_span);
                        }
                    }
                    try buf.ensureUnusedCapacity(arena_alloc, span.len);
                    if (span_has_ent) {
                        try appendDecoded(arena_alloc, &buf, span);
                    } else {
                        buf.appendSliceAssumeCapacity(span);
                    }
                }
                t_count += 1;

                // Record a rich-text run only when we've entered at
                // least one `<r>` wrapper — otherwise this is a plain
                // `<si><t>...</t></si>` and richRuns() returns null.
                if (saw_r) {
                    const run_text: []const u8 = if (span_has_ent) blk: {
                        var rb: std.ArrayListUnmanaged(u8) = try .initCapacity(arena_alloc, span.len);
                        try appendDecoded(arena_alloc, &rb, span);
                        break :blk try rb.toOwnedSlice(arena_alloc);
                    } else span;
                    try runs.append(arena_alloc, .{
                        .text = run_text,
                        .bold = pending_flags.bold,
                        .italic = pending_flags.italic,
                    });
                }

                j = t_close + "</t>".len;
                std.debug.assert(j > j_prev);
            } else if (c1 == 'r' and next_lt + 2 < xml.len and
                (xml[next_lt + 2] == '>' or xml[next_lt + 2] == ' '))
            {
                // `<r>` — enter a rich-text run block. Every `<r>` gets
                // its own formatting; reset pending_flags so a run
                // without `<rPr>` defaults to unstyled.
                const r_gt = std.mem.indexOfScalarPos(u8, xml, next_lt + 2, '>') orelse break :body;
                saw_r = true;
                pending_flags = .{ .text = "" };
                j = r_gt + 1;
                std.debug.assert(j > j_prev);
            } else if (c1 == 'r' and next_lt + 3 < xml.len and
                xml[next_lt + 2] == 'P' and xml[next_lt + 3] == 'r')
            {
                // `<rPr>...</rPr>` — parse bold / italic, skip body.
                const rpr_close = std.mem.indexOfPos(u8, xml, next_lt, "</rPr>") orelse break :body;
                pending_flags = parseRprFlags(xml[next_lt .. rpr_close + "</rPr>".len]);
                j = rpr_close + "</rPr>".len;
                std.debug.assert(j > j_prev);
            } else if (c1 == '/' and next_lt + 5 <= xml.len and
                xml[next_lt + 2] == 's' and xml[next_lt + 3] == 'i' and xml[next_lt + 4] == '>')
            {
                // `</si>` — emit and advance past it.
                if (t_count == 0) {
                    try strings.append(book.allocator, "");
                } else if (t_count == 1) {
                    if (first_has_ent) {
                        var b: std.ArrayListUnmanaged(u8) = try .initCapacity(arena_alloc, first_span.len);
                        try appendDecoded(arena_alloc, &b, first_span);
                        try strings.append(book.allocator, try b.toOwnedSlice(arena_alloc));
                    } else {
                        // Fast borrow path — zero allocations.
                        try strings.append(book.allocator, first_span);
                    }
                } else {
                    try strings.append(book.allocator, try buf.toOwnedSlice(arena_alloc));
                }

                // Rich runs — only stash when at least one `<r>` was
                // seen AND at least one run carried text. Keeps the
                // common-case (no rich text) map empty.
                if (saw_r and runs.items.len > 0) {
                    const owned = try runs.toOwnedSlice(arena_alloc);
                    try book.rich_runs_by_sst_idx.put(book.allocator, sst_idx, owned);
                }

                i = next_lt + 5;
                std.debug.assert(i > i_prev);
                continue :outer;
            } else {
                // Skip any other tag (unmatched `<t` patterns, etc.).
                // indexOfScalarPos from `next_lt + 1` guarantees
                // monotonic progress.
                const skip_gt = std.mem.indexOfScalarPos(u8, xml, next_lt + 1, '>') orelse break :body;
                j = skip_gt + 1;
                std.debug.assert(j > j_prev);
            }
        }

        // Inner body loop fell through without finding `</si>` —
        // malformed SST entry. Advance `i` past the opening `<si` we
        // found so the outer loop makes monotonic progress (the iter16
        // bug was forgetting this and re-entering the same bad `<si>`
        // forever on fuzz-random input).
        i = si_gt + 1;
        std.debug.assert(i > i_prev);
    }

    book.shared_strings = try strings.toOwnedSlice(book.allocator);
}

// ─── zip → buffer ────────────────────────────────────────────────────

fn extractEntryToBuffer(
    allocator: Allocator,
    entry: std.zip.Iterator.Entry,
    stream: *std.fs.File.Reader,
) ![]u8 {
    switch (entry.compression_method) {
        .store, .deflate => {},
        else => return error.UnsupportedCompression,
    }

    // Read + verify LocalFileHeader.
    try stream.seekTo(entry.file_offset);
    const local = stream.interface.takeStruct(std.zip.LocalFileHeader, .little) catch return error.BadZip;
    if (!std.mem.eql(u8, &local.signature, &std.zip.local_file_header_sig)) return error.BadZip;

    try stream.seekTo(entry.file_offset + @sizeOf(std.zip.LocalFileHeader) + local.filename_len + local.extra_len);

    const out = try allocator.alloc(u8, entry.uncompressed_size);
    errdefer allocator.free(out);
    var writer = std.Io.Writer.fixed(out);

    switch (entry.compression_method) {
        .store => {
            stream.interface.streamExact64(&writer, entry.uncompressed_size) catch return error.BadZip;
        },
        .deflate => {
            var flate_buffer: [std.compress.flate.max_window_len]u8 = undefined;
            var decompress = std.compress.flate.Decompress.init(&stream.interface, .raw, &flate_buffer);
            decompress.reader.streamExact64(&writer, entry.uncompressed_size) catch return error.BadZip;
        },
        else => unreachable,
    }

    return out;
}

// ─── Tests ───────────────────────────────────────────────────────────

test "fromExcelSerial: known reference dates + rejection matrix" {
    // Reference points — all confirmed against Excel's DATEVALUE:
    //   61      = 1900-03-01 (first serial past the leap-year bug window)
    //   25569   = 1970-01-01 (Unix epoch)
    //   40000   = 2009-07-06
    //   43831   = 2020-01-01
    //   45658   = 2025-01-01
    //   2958465 = 9999-12-31 (max)
    const cases = [_]struct { s: f64, y: u16, m: u8, d: u8, h: u8, mi: u8, se: u8 }{
        .{ .s = 61.0, .y = 1900, .m = 3, .d = 1, .h = 0, .mi = 0, .se = 0 },
        .{ .s = 25569.0, .y = 1970, .m = 1, .d = 1, .h = 0, .mi = 0, .se = 0 },
        .{ .s = 40000.0, .y = 2009, .m = 7, .d = 6, .h = 0, .mi = 0, .se = 0 },
        .{ .s = 43831.0, .y = 2020, .m = 1, .d = 1, .h = 0, .mi = 0, .se = 0 },
        .{ .s = 45658.0, .y = 2025, .m = 1, .d = 1, .h = 0, .mi = 0, .se = 0 },
        .{ .s = 2958465.0, .y = 9999, .m = 12, .d = 31, .h = 0, .mi = 0, .se = 0 },
        // Times-of-day — noon, 3:30 PM, end-of-minute.
        .{ .s = 45658.5, .y = 2025, .m = 1, .d = 1, .h = 12, .mi = 0, .se = 0 },
        .{ .s = 45658.0 + (15.0 * 3600.0 + 30.0 * 60.0) / 86400.0, .y = 2025, .m = 1, .d = 1, .h = 15, .mi = 30, .se = 0 },
        .{ .s = 45658.0 + 59.0 / 86400.0, .y = 2025, .m = 1, .d = 1, .h = 0, .mi = 0, .se = 59 },
    };
    for (cases) |c| {
        const got = fromExcelSerial(c.s) orelse {
            std.debug.print("fromExcelSerial returned null for serial {d}\n", .{c.s});
            return error.UnexpectedNull;
        };
        try std.testing.expectEqual(@as(u16, c.y), got.year);
        try std.testing.expectEqual(@as(u8, c.m), got.month);
        try std.testing.expectEqual(@as(u8, c.d), got.day);
        try std.testing.expectEqual(@as(u8, c.h), got.hour);
        try std.testing.expectEqual(@as(u8, c.mi), got.minute);
        try std.testing.expectEqual(@as(u8, c.se), got.second);
    }

    // Rejection: NaN, infinity, pre-1900-leap-bug window, past 9999.
    try std.testing.expect(fromExcelSerial(std.math.nan(f64)) == null);
    try std.testing.expect(fromExcelSerial(std.math.inf(f64)) == null);
    try std.testing.expect(fromExcelSerial(-std.math.inf(f64)) == null);
    try std.testing.expect(fromExcelSerial(-1.0) == null);
    try std.testing.expect(fromExcelSerial(0.0) == null);
    try std.testing.expect(fromExcelSerial(60.0) == null); // fictitious 1900-02-29
    try std.testing.expect(fromExcelSerial(60.9999) == null);
    try std.testing.expect(fromExcelSerial(2958466.0) == null);
    try std.testing.expect(fromExcelSerial(1e20) == null);
}

test "fuzz fromExcelSerial: finite serials never panic + accepted range invariants" {
    const iters = fuzz_default_iters;
    var prng = std.Random.DefaultPrng.init(0xD4725EA1);
    const rng = prng.random();
    for (0..iters) |_| {
        // Mix of in-range serials, out-of-range, and edge cases.
        const serial: f64 = switch (rng.intRangeAtMost(u8, 0, 9)) {
            0 => std.math.nan(f64),
            1 => std.math.inf(f64),
            2 => -std.math.inf(f64),
            3 => -rng.float(f64) * 1e10,
            4 => rng.float(f64) * 60.0, // pre-leap-bug window
            5 => 2958466.0 + rng.float(f64) * 1e6, // past max
            // The rest: in-range values.
            else => 61.0 + rng.float(f64) * (2958465.0 - 61.0),
        };
        if (fromExcelSerial(serial)) |dt| {
            // Invariants must hold on any accepted input.
            try std.testing.expect(dt.year >= 1900 and dt.year <= 9999);
            try std.testing.expect(dt.month >= 1 and dt.month <= 12);
            try std.testing.expect(dt.day >= 1 and dt.day <= 31);
            try std.testing.expect(dt.hour <= 23);
            try std.testing.expect(dt.minute <= 59);
            try std.testing.expect(dt.second <= 59);
        }
    }
}

test "parseA1Ref: basic A1 parsing + rejection" {
    try std.testing.expectEqualDeep(CellRef{ .col = 0, .row = 1 }, try parseA1Ref("A1"));
    try std.testing.expectEqualDeep(CellRef{ .col = 1, .row = 2 }, try parseA1Ref("B2"));
    try std.testing.expectEqualDeep(CellRef{ .col = 25, .row = 99 }, try parseA1Ref("Z99"));
    try std.testing.expectEqualDeep(CellRef{ .col = 26, .row = 1 }, try parseA1Ref("AA1"));
    try std.testing.expectEqualDeep(CellRef{ .col = 16383, .row = 1048576 }, try parseA1Ref("XFD1048576"));

    try std.testing.expectError(error.MalformedXml, parseA1Ref(""));
    try std.testing.expectError(error.MalformedXml, parseA1Ref("A0")); // row 0
    try std.testing.expectError(error.MalformedXml, parseA1Ref("1A")); // row before col
    try std.testing.expectError(error.MalformedXml, parseA1Ref("A")); // no row
    try std.testing.expectError(error.MalformedXml, parseA1Ref("1")); // no col
    try std.testing.expectError(error.MalformedXml, parseA1Ref("a1")); // lowercase
}

test "parseA1Range: rectangle parsing + corner normalisation" {
    const r = try parseA1Range("A1:B2");
    try std.testing.expectEqualDeep(CellRef{ .col = 0, .row = 1 }, r.top_left);
    try std.testing.expectEqualDeep(CellRef{ .col = 1, .row = 2 }, r.bottom_right);

    // Single-cell form is accepted (1×1 rectangle).
    const s = try parseA1Range("C3");
    try std.testing.expectEqualDeep(CellRef{ .col = 2, .row = 3 }, s.top_left);
    try std.testing.expectEqualDeep(CellRef{ .col = 2, .row = 3 }, s.bottom_right);

    // Order-swapped corners get normalised — some historical
    // generators emit "B2:A1" where they meant "A1:B2".
    const swapped = try parseA1Range("B2:A1");
    try std.testing.expectEqualDeep(CellRef{ .col = 0, .row = 1 }, swapped.top_left);
    try std.testing.expectEqualDeep(CellRef{ .col = 1, .row = 2 }, swapped.bottom_right);
}

test "Book.mergedRanges: round-trip through writer + reader" {
    const tmp_path = "/tmp/zlsx_reader_merged_roundtrip.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("Sheet1");
        try sheet.addMergedCell("A1:C1");
        try sheet.addMergedCell("B5:D7");
        try sheet.addMergedCell("XFD1048575:XFD1048576");
        try sheet.writeRow(&.{.{ .string = "hdr" }});
        try w.save(tmp_path);
    }

    var book = try Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    const ranges = book.mergedRanges(book.sheets[0]);
    try std.testing.expectEqual(@as(usize, 3), ranges.len);
    try std.testing.expectEqualDeep(CellRef{ .col = 0, .row = 1 }, ranges[0].top_left);
    try std.testing.expectEqualDeep(CellRef{ .col = 2, .row = 1 }, ranges[0].bottom_right);
    try std.testing.expectEqualDeep(CellRef{ .col = 1, .row = 5 }, ranges[1].top_left);
    try std.testing.expectEqualDeep(CellRef{ .col = 3, .row = 7 }, ranges[1].bottom_right);
    try std.testing.expectEqualDeep(CellRef{ .col = 16383, .row = 1048575 }, ranges[2].top_left);
    try std.testing.expectEqualDeep(CellRef{ .col = 16383, .row = 1048576 }, ranges[2].bottom_right);
}

test "Book.hyperlinks: round-trip through writer + reader" {
    const tmp_path = "/tmp/zlsx_reader_hyperlinks_roundtrip.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("Links");
        try sheet.addHyperlink("A1", "https://example.com/path?q=1&x=2");
        try sheet.addHyperlink("B2:C3", "mailto:foo@example.com");
        try sheet.addHyperlink("D5", "https://docs.example.com/");
        try sheet.writeRow(&.{.{ .string = "click" }});
        try w.save(tmp_path);
    }

    var book = try Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    const links = book.hyperlinks(book.sheets[0]);
    try std.testing.expectEqual(@as(usize, 3), links.len);

    // rId1 → A1 single cell → top_left == bottom_right.
    try std.testing.expectEqualDeep(CellRef{ .col = 0, .row = 1 }, links[0].top_left);
    try std.testing.expectEqualDeep(CellRef{ .col = 0, .row = 1 }, links[0].bottom_right);
    // Writer xml-escapes `&` to `&amp;` on emit; reader must NOT
    // un-escape here — the contract is that `url` is the raw
    // `Target` attribute, and entity decoding on URLs is a caller
    // decision (they round-trip fine through every major consumer).
    try std.testing.expectEqualStrings("https://example.com/path?q=1&amp;x=2", links[0].url);

    try std.testing.expectEqualDeep(CellRef{ .col = 1, .row = 2 }, links[1].top_left);
    try std.testing.expectEqualDeep(CellRef{ .col = 2, .row = 3 }, links[1].bottom_right);
    try std.testing.expectEqualStrings("mailto:foo@example.com", links[1].url);

    try std.testing.expectEqualDeep(CellRef{ .col = 3, .row = 5 }, links[2].top_left);
    try std.testing.expectEqualStrings("https://docs.example.com/", links[2].url);
}

test "Book.dataValidations: round-trip through writer + reader" {
    const tmp_path = "/tmp/zlsx_reader_dv_roundtrip.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("Pick");
        try sheet.addDataValidationList("A2:A10", &.{ "Red", "Green", "Blue" });
        try sheet.addDataValidationList("C3", &.{"Single"});
        // XML-special chars — writer escapes `R&D` → `R&amp;D`;
        // reader must decode back on the way out.
        try sheet.addDataValidationList("B2", &.{ "R&D", "Q<A", "x>y" });
        try sheet.writeRow(&.{.{ .string = "hdr" }});
        try w.save(tmp_path);
    }

    var book = try Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    const dvs = book.dataValidations(book.sheets[0]);
    try std.testing.expectEqual(@as(usize, 3), dvs.len);

    // 0: A2:A10 → three values.
    try std.testing.expectEqualDeep(CellRef{ .col = 0, .row = 2 }, dvs[0].top_left);
    try std.testing.expectEqualDeep(CellRef{ .col = 0, .row = 10 }, dvs[0].bottom_right);
    try std.testing.expectEqual(@as(usize, 3), dvs[0].values.len);
    try std.testing.expectEqualStrings("Red", dvs[0].values[0]);
    try std.testing.expectEqualStrings("Green", dvs[0].values[1]);
    try std.testing.expectEqualStrings("Blue", dvs[0].values[2]);

    // 1: C3 single cell, one value.
    try std.testing.expectEqualDeep(CellRef{ .col = 2, .row = 3 }, dvs[1].top_left);
    try std.testing.expectEqualDeep(CellRef{ .col = 2, .row = 3 }, dvs[1].bottom_right);
    try std.testing.expectEqual(@as(usize, 1), dvs[1].values.len);
    try std.testing.expectEqualStrings("Single", dvs[1].values[0]);

    // 2: B2 with XML-special chars, must be decoded on the way out.
    try std.testing.expectEqualDeep(CellRef{ .col = 1, .row = 2 }, dvs[2].top_left);
    try std.testing.expectEqual(@as(usize, 3), dvs[2].values.len);
    try std.testing.expectEqualStrings("R&D", dvs[2].values[0]);
    try std.testing.expectEqualStrings("Q<A", dvs[2].values[1]);
    try std.testing.expectEqualStrings("x>y", dvs[2].values[2]);
}

test "Book.dataValidations: numeric + custom round-trip kind / op / formula1 / formula2" {
    const tmp_path = "/tmp/zlsx_reader_dv_numeric_roundtrip.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("Num");
        try sheet.addDataValidationNumeric("B2:B10", .whole, .between, "1", "100");
        try sheet.addDataValidationNumeric("C3", .decimal, .greater_than, "0", null);
        try sheet.addDataValidationNumeric("D4", .date, .less_than, "45658", null);
        try sheet.addDataValidationNumeric("E5", .text_length, .between, "3", "20");
        // Custom formula uses XML-special char `<`; reader must decode it.
        try sheet.addDataValidationCustom("F6", "AND(F6>0,F6<LEN(A1))");
        // List still round-trips alongside the new kinds.
        try sheet.addDataValidationList("G7", &.{ "Yes", "No" });
        try sheet.writeRow(&.{.{ .string = "hdr" }});
        try w.save(tmp_path);
    }

    var book = try Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    const dvs = book.dataValidations(book.sheets[0]);
    // List validations emit first (iter13 path), then the numeric /
    // custom range block (iter23 path). Writer ordering preserves this.
    try std.testing.expectEqual(@as(usize, 6), dvs.len);

    // 0: list — pre-existing iter13 path.
    try std.testing.expectEqual(DataValidationKind.list, dvs[0].kind);
    try std.testing.expectEqual(@as(?DataValidationOperator, null), dvs[0].op);
    try std.testing.expectEqual(@as(usize, 2), dvs[0].values.len);
    try std.testing.expectEqualStrings("Yes", dvs[0].values[0]);
    try std.testing.expectEqualStrings("No", dvs[0].values[1]);

    // 1: whole between 1..100
    try std.testing.expectEqual(DataValidationKind.whole, dvs[1].kind);
    try std.testing.expectEqual(@as(?DataValidationOperator, .between), dvs[1].op);
    try std.testing.expectEqualStrings("1", dvs[1].formula1);
    try std.testing.expectEqualStrings("100", dvs[1].formula2);
    try std.testing.expectEqual(@as(usize, 0), dvs[1].values.len);

    // 2: decimal greater_than 0
    try std.testing.expectEqual(DataValidationKind.decimal, dvs[2].kind);
    try std.testing.expectEqual(@as(?DataValidationOperator, .greater_than), dvs[2].op);
    try std.testing.expectEqualStrings("0", dvs[2].formula1);
    try std.testing.expectEqualStrings("", dvs[2].formula2);

    // 3: date less_than 45658
    try std.testing.expectEqual(DataValidationKind.date, dvs[3].kind);
    try std.testing.expectEqual(@as(?DataValidationOperator, .less_than), dvs[3].op);
    try std.testing.expectEqualStrings("45658", dvs[3].formula1);

    // 4: text_length between 3..20
    try std.testing.expectEqual(DataValidationKind.text_length, dvs[4].kind);
    try std.testing.expectEqual(@as(?DataValidationOperator, .between), dvs[4].op);
    try std.testing.expectEqualStrings("3", dvs[4].formula1);
    try std.testing.expectEqualStrings("20", dvs[4].formula2);

    // 5: custom — no operator, formula1 contains `<` (XML-encoded `&lt;`
    // on disk, decoded on read).
    try std.testing.expectEqual(DataValidationKind.custom, dvs[5].kind);
    try std.testing.expectEqual(@as(?DataValidationOperator, null), dvs[5].op);
    try std.testing.expectEqualStrings("AND(F6>0,F6<LEN(A1))", dvs[5].formula1);
    try std.testing.expectEqualStrings("", dvs[5].formula2);
}

test "Book.dataValidations: empty slice for sheets without validations" {
    const tmp_path = "/tmp/zlsx_reader_no_dv.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("Plain");
        try sheet.writeRow(&.{.{ .string = "a" }});
        try w.save(tmp_path);
    }

    var book = try Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 0), book.dataValidations(book.sheets[0]).len);
}

test "Book.richRuns: rich-text SST entries expose per-run bold/italic" {
    // Direct parseSharedStrings drive — avoids needing the writer to
    // emit rich text (it doesn't yet). Covers: plain `<t>` → null runs,
    // single `<r>` with `<b/>`, multiple `<r>` with mixed flags,
    // `val="0"` explicit-false, and entity-decoded run text.
    const sst_xml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" ++
        "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"5\">" ++
        "<si><t>plain</t></si>" ++
        "<si><r><rPr><b/></rPr><t>bold</t></r></si>" ++
        "<si><r><rPr><b/><i/></rPr><t>bold-italic</t></r></si>" ++
        "<si><r><rPr><b/></rPr><t>A</t></r><r><rPr><i/></rPr><t> B</t></r><r><t> C</t></r></si>" ++
        "<si><r><rPr><b val=\"0\"/><i/></rPr><t>R&amp;D</t></r></si>" ++
        "</sst>";

    var book: Book = .{
        .allocator = std.testing.allocator,
        .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer book.deinit();
    const owned = try std.testing.allocator.dupe(u8, sst_xml);
    book.shared_strings_xml = owned;
    try parseSharedStrings(&book, owned);

    try std.testing.expectEqual(@as(usize, 5), book.shared_strings.len);

    // Flat strings round-trip correctly for both plain and rich paths.
    try std.testing.expectEqualStrings("plain", book.shared_strings[0]);
    try std.testing.expectEqualStrings("bold", book.shared_strings[1]);
    try std.testing.expectEqualStrings("bold-italic", book.shared_strings[2]);
    try std.testing.expectEqualStrings("A B C", book.shared_strings[3]);
    try std.testing.expectEqualStrings("R&D", book.shared_strings[4]);

    // Plain SST entries return null from richRuns — zero map overhead.
    try std.testing.expectEqual(@as(?[]const RichRun, null), book.richRuns(0));

    // Single-run bold.
    const r1 = book.richRuns(1) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(usize, 1), r1.len);
    try std.testing.expectEqualStrings("bold", r1[0].text);
    try std.testing.expectEqual(true, r1[0].bold);
    try std.testing.expectEqual(false, r1[0].italic);

    // Single-run bold + italic.
    const r2 = book.richRuns(2) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(usize, 1), r2.len);
    try std.testing.expectEqual(true, r2[0].bold);
    try std.testing.expectEqual(true, r2[0].italic);

    // Multi-run: bold / italic / plain.
    const r3 = book.richRuns(3) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(usize, 3), r3.len);
    try std.testing.expectEqualStrings("A", r3[0].text);
    try std.testing.expectEqual(true, r3[0].bold);
    try std.testing.expectEqual(false, r3[0].italic);
    try std.testing.expectEqualStrings(" B", r3[1].text);
    try std.testing.expectEqual(false, r3[1].bold);
    try std.testing.expectEqual(true, r3[1].italic);
    try std.testing.expectEqualStrings(" C", r3[2].text);
    try std.testing.expectEqual(false, r3[2].bold);
    try std.testing.expectEqual(false, r3[2].italic);

    // `val="0"` on bold overrides the presence of `<b/>`; italic still true.
    // Entity-decoded text must come through clean.
    const r4 = book.richRuns(4) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(usize, 1), r4.len);
    try std.testing.expectEqualStrings("R&D", r4[0].text);
    try std.testing.expectEqual(false, r4[0].bold);
    try std.testing.expectEqual(true, r4[0].italic);
}

test "Book.hyperlinks: internal hyperlinks (location) round-trip + mixed external/internal" {
    const tmp_path = "/tmp/zlsx_reader_internal_hyperlinks.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        _ = try w.addSheet("Main");
        var s2 = try w.addSheet("Details");
        try s2.addInternalHyperlink("A1", "Main!A1");
        try s2.addInternalHyperlink("B2:C2", "'Main'!B2");
        try s2.addHyperlink("D1", "https://example.com/"); // mixed
        try s2.writeRow(&.{.{ .string = "x" }});
        // Rejection path.
        try std.testing.expectError(error.InvalidHyperlinkLocation, s2.addInternalHyperlink("A3", ""));
        try std.testing.expectError(error.InvalidHyperlinkRange, s2.addInternalHyperlink("", "Main!A1"));
        try w.save(tmp_path);
    }

    var book = try Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    const links = book.hyperlinks(book.sheets[1]); // "Details"
    try std.testing.expectEqual(@as(usize, 3), links.len);

    // External hyperlinks come first in our writer's emission order
    // (r:id references emit before `location`-bearing entries), so
    // the external one is links[0].
    try std.testing.expectEqualStrings("https://example.com/", links[0].url);
    try std.testing.expectEqualStrings("", links[0].location);

    // Then the two internal entries.
    try std.testing.expectEqualStrings("", links[1].url);
    try std.testing.expectEqualStrings("Main!A1", links[1].location);
    try std.testing.expectEqualDeep(CellRef{ .col = 0, .row = 1 }, links[1].top_left);

    try std.testing.expectEqualStrings("", links[2].url);
    // Writer xml-escapes `'` → `&apos;` on emit; reader preserves the
    // raw attribute bytes (matches the `url` contract for external
    // hyperlinks — decoding is the caller's choice).
    try std.testing.expectEqualStrings("&apos;Main&apos;!B2", links[2].location);
    try std.testing.expectEqualDeep(CellRef{ .col = 1, .row = 2 }, links[2].top_left);
    try std.testing.expectEqualDeep(CellRef{ .col = 2, .row = 2 }, links[2].bottom_right);
}

test "Book.hyperlinks: internal-only sheet (no _rels file needed)" {
    const tmp_path = "/tmp/zlsx_reader_internal_only.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        _ = try w.addSheet("Main");
        var s2 = try w.addSheet("TOC");
        try s2.addInternalHyperlink("A1", "Main!A1");
        try s2.writeRow(&.{.{ .string = "ToC" }});
        try w.save(tmp_path);
    }

    var book = try Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();
    const links = book.hyperlinks(book.sheets[1]);
    try std.testing.expectEqual(@as(usize, 1), links.len);
    try std.testing.expectEqualStrings("Main!A1", links[0].location);
    try std.testing.expectEqualStrings("", links[0].url);
}

test "Book.hyperlinks: empty slice for sheets without hyperlinks" {
    const tmp_path = "/tmp/zlsx_reader_no_hyperlinks.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("Plain");
        try sheet.writeRow(&.{.{ .string = "no-links" }});
        try w.save(tmp_path);
    }

    var book = try Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 0), book.hyperlinks(book.sheets[0]).len);
}

test "Book.mergedRanges: empty slice for sheets without merges" {
    const tmp_path = "/tmp/zlsx_reader_no_merged.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("Sheet1");
        try sheet.writeRow(&.{.{ .string = "a" }});
        try w.save(tmp_path);
    }

    var book = try Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 0), book.mergedRanges(book.sheets[0]).len);
}

test "columnIndexFromRef" {
    try std.testing.expectEqual(@as(usize, 0), try columnIndexFromRef("A1"));
    try std.testing.expectEqual(@as(usize, 1), try columnIndexFromRef("B2"));
    try std.testing.expectEqual(@as(usize, 25), try columnIndexFromRef("Z99"));
    try std.testing.expectEqual(@as(usize, 26), try columnIndexFromRef("AA1"));
    try std.testing.expectEqual(@as(usize, 27), try columnIndexFromRef("AB1"));
    try std.testing.expectEqual(@as(usize, 51), try columnIndexFromRef("AZ1"));
    try std.testing.expectEqual(@as(usize, 52), try columnIndexFromRef("BA1"));
}

test "getAttr" {
    try std.testing.expectEqualStrings("A1", getAttr(" r=\"A1\" t=\"s\"", "r").?);
    try std.testing.expectEqualStrings("s", getAttr(" r=\"A1\" t=\"s\"", "t").?);
    try std.testing.expect(getAttr(" r=\"A1\" t=\"s\"", "x") == null);
    try std.testing.expectEqualStrings("", getAttr(" r=\"\" t=\"s\"", "r").?);
}

test "appendDecoded entities" {
    const alloc = std.testing.allocator;
    var buf: std.ArrayListUnmanaged(u8) = .{};
    defer buf.deinit(alloc);
    try appendDecoded(alloc, &buf, "Smith &amp; Co &lt;HQ&gt; &#233;");
    try std.testing.expectEqualStrings("Smith & Co <HQ> é", buf.items);
}

test "parseNumericCell" {
    try std.testing.expectEqual(Cell{ .integer = 42 }, try parseNumericCell("42"));
    try std.testing.expectEqual(Cell{ .integer = -7 }, try parseNumericCell("-7"));
    const f = try parseNumericCell("3.14");
    try std.testing.expect(f == .number);
    try std.testing.expect(@abs(f.number - 3.14) < 1e-9);
    try std.testing.expectEqual(Cell.empty, try parseNumericCell(""));
}

// ─── Fuzz suite (PRNG-driven) ────────────────────────────────────────
//
// These tests generate random inputs with a seeded PRNG and run every
// public-facing and internal parser against them. The contract each
// fuzz target enforces: **no crashes, no OOM panics, no infinite
// loops, no unreachable triggered** on any byte input. An error
// return is acceptable — the public API treats malformed input as a
// recoverable error, not a process-abort.
//
// Iteration count is picked up from the XLSX_FUZZ_ITERS environment
// variable at test time, defaulting to a small number so regular
// `zig build test` stays fast. For deep fuzzing:
//
//     XLSX_FUZZ_ITERS=1_000_000 zig build test
//
// Why PRNG + not `std.testing.fuzz`: Zig 0.15's coverage-guided fuzz
// (`zig build test --fuzz`) crashes on macOS due to a Mach-O parsing
// bug in `std.Build.Fuzz.addEntryPoint`. Dumb fuzzing with a seeded
// PRNG is portable, reproducible (the seed is logged), and catches
// the same classes of bugs for state-machine parsers.

const fuzz_default_iters: usize = 1_000;
const fuzz_max_input_len: usize = 4_096;

fn fuzzIterations() usize {
    const env = std.process.getEnvVarOwned(std.heap.page_allocator, "XLSX_FUZZ_ITERS") catch return fuzz_default_iters;
    defer std.heap.page_allocator.free(env);
    // Strip underscores so humans can write "1_000_000".
    var digits_buf: [32]u8 = undefined;
    var di: usize = 0;
    for (env) |c| {
        if (c == '_') continue;
        if (di == digits_buf.len) break;
        digits_buf[di] = c;
        di += 1;
    }
    return std.fmt.parseInt(usize, digits_buf[0..di], 10) catch fuzz_default_iters;
}

fn fuzzSeed() u64 {
    if (std.process.getEnvVarOwned(std.heap.page_allocator, "XLSX_FUZZ_SEED")) |s| {
        defer std.heap.page_allocator.free(s);
        return std.fmt.parseInt(u64, s, 10) catch 0xA1F8ED;
    } else |_| {
        return @bitCast(std.time.milliTimestamp());
    }
}

fn randomInput(rng: std.Random, buf: []u8) []u8 {
    const len = rng.intRangeAtMost(usize, 0, buf.len);
    rng.bytes(buf[0..len]);
    return buf[0..len];
}

test "fuzz parseA1Ref + parseA1Range: adversarial bytes never panic" {
    const iters = fuzz_default_iters;
    var prng = std.Random.DefaultPrng.init(0x3F2A1E);
    const rng = prng.random();

    var buf: [20]u8 = undefined;
    for (0..iters) |_| {
        const len = rng.intRangeAtMost(usize, 0, buf.len);
        for (0..len) |i| buf[i] = rng.int(u8);
        const s = buf[0..len];

        if (parseA1Ref(s)) |ref| {
            // Any accepted input must map to an in-Excel-range cell.
            try std.testing.expect(ref.col <= 16383);
            try std.testing.expect(ref.row >= 1 and ref.row <= 1048576);
        } else |err| {
            try std.testing.expectEqual(error.MalformedXml, err);
        }

        if (parseA1Range(s)) |r| {
            try std.testing.expect(r.top_left.col <= r.bottom_right.col);
            try std.testing.expect(r.top_left.row <= r.bottom_right.row);
        } else |err| {
            try std.testing.expectEqual(error.MalformedXml, err);
        }
    }
}

test "fuzz columnIndexFromRef" {
    const iters = fuzzIterations();
    const seed = fuzzSeed();
    var prng = std.Random.DefaultPrng.init(seed);
    const rng = prng.random();
    var buf: [64]u8 = undefined;
    for (0..iters) |_| {
        const input = randomInput(rng, &buf);
        _ = columnIndexFromRef(input) catch {};
    }
}

test "fuzz parseNumericCell" {
    const iters = fuzzIterations();
    var prng = std.Random.DefaultPrng.init(fuzzSeed());
    const rng = prng.random();
    var buf: [128]u8 = undefined;
    for (0..iters) |_| {
        const input = randomInput(rng, &buf);
        _ = parseNumericCell(input) catch {};
    }
}

test "fuzz appendDecoded" {
    const iters = fuzzIterations();
    var prng = std.Random.DefaultPrng.init(fuzzSeed());
    const rng = prng.random();
    var buf: [512]u8 = undefined;
    for (0..iters) |_| {
        const input = randomInput(rng, &buf);
        var out: std.ArrayListUnmanaged(u8) = .{};
        defer out.deinit(std.testing.allocator);
        appendDecoded(std.testing.allocator, &out, input) catch {};
    }
}

test "fuzz getAttr" {
    const iters = fuzzIterations();
    var prng = std.Random.DefaultPrng.init(fuzzSeed());
    const rng = prng.random();
    var buf: [256]u8 = undefined;
    var name_buf: [16]u8 = undefined;
    for (0..iters) |_| {
        const input = randomInput(rng, &buf);
        const name_len = rng.intRangeAtMost(usize, 0, name_buf.len);
        rng.bytes(name_buf[0..name_len]);
        _ = getAttr(input, name_buf[0..name_len]);
    }
}

test "fuzz findTagOpen" {
    const iters = fuzzIterations();
    var prng = std.Random.DefaultPrng.init(fuzzSeed());
    const rng = prng.random();
    var buf: [512]u8 = undefined;
    const tags = [_][]const u8{ "c", "row", "si", "sheet", "Relationship", "t" };
    for (0..iters) |_| {
        const input = randomInput(rng, &buf);
        for (tags) |tag| {
            _ = findTagOpen(input, 0, tag);
        }
    }
}

test "fuzz extractVValue" {
    const iters = fuzzIterations();
    var prng = std.Random.DefaultPrng.init(fuzzSeed());
    const rng = prng.random();
    var buf: [256]u8 = undefined;
    for (0..iters) |_| {
        const input = randomInput(rng, &buf);
        _ = extractVValue(input);
    }
}

test "fuzz parseSharedStrings" {
    const iters = fuzzIterations();
    var prng = std.Random.DefaultPrng.init(fuzzSeed());
    const rng = prng.random();
    var buf: [fuzz_max_input_len]u8 = undefined;
    for (0..iters) |_| {
        const input = randomInput(rng, &buf);
        var book: Book = .{ .allocator = std.testing.allocator, .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator) };
        defer book.deinit();
        // parseSharedStrings may borrow spans from the xml buffer when
        // no entity decoding is needed; dupe it so the buffer outlives
        // the book regardless of borrowing choice.
        const owned = std.testing.allocator.dupe(u8, input) catch continue;
        book.shared_strings_xml = owned;
        parseSharedStrings(&book, owned) catch {};
    }
}

test "fuzz parseWorkbookSheets" {
    const iters = fuzzIterations();
    var prng = std.Random.DefaultPrng.init(fuzzSeed());
    const rng = prng.random();
    var buf: [fuzz_max_input_len]u8 = undefined;
    for (0..iters) |_| {
        const input = randomInput(rng, &buf);
        const mid = input.len / 2;
        const wb = input[0..mid];
        const rels = input[mid..];
        var book: Book = .{ .allocator = std.testing.allocator, .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator) };
        defer book.deinit();
        parseWorkbookSheets(&book, wb, rels) catch {};
    }
}

// ─── Mutation fuzzing ────────────────────────────────────────────────
//
// Random-byte fuzzing rarely produces inputs that advance the XML
// state machines past the first `<`. These tests start from real XML
// templates and apply random mutations (byte flips, deletions,
// insertions, duplications) so parser branches deep in the state
// machine actually get hit.

const sst_template =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
    "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"3\" uniqueCount=\"3\">" ++
    "<si><t>Hello</t></si>" ++
    "<si><r><rPr><b/></rPr><t>World &amp; more</t></r></si>" ++
    "<si><t xml:space=\"preserve\">  spaced  </t></si>" ++
    "</sst>";

const workbook_template =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
    "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " ++
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" ++
    "<sheets>" ++
    "<sheet name=\"Alpha\" sheetId=\"1\" r:id=\"rId1\"/>" ++
    "<sheet name=\"Beta &amp; Gamma\" sheetId=\"2\" r:id=\"rId2\"/>" ++
    "</sheets></workbook>";

const rels_template =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" ++
    "<Relationship Id=\"rId1\" Type=\"http://example.com/ws\" Target=\"worksheets/sheet1.xml\"/>" ++
    "<Relationship Id=\"rId2\" Type=\"http://example.com/ws\" Target=\"worksheets/sheet2.xml\"/>" ++
    "</Relationships>";

/// Mutate `src` into `dst` with a small number of random edits. The
/// caller provides the destination buffer so the allocator cost is
/// eliminated from the fuzz hot loop.
fn mutate(rng: std.Random, src: []const u8, dst: []u8) []u8 {
    // Copy src into dst, possibly truncated.
    const base_len = @min(src.len, dst.len);
    @memcpy(dst[0..base_len], src[0..base_len]);
    var len = base_len;

    // Apply 1-8 random edits.
    const edits = rng.intRangeAtMost(u8, 1, 8);
    for (0..edits) |_| {
        if (len == 0) break;
        const op = rng.intRangeAtMost(u8, 0, 3);
        switch (op) {
            0 => {
                // Byte flip
                const i = rng.intRangeLessThan(usize, 0, len);
                dst[i] = rng.int(u8);
            },
            1 => {
                // Byte delete
                const i = rng.intRangeLessThan(usize, 0, len);
                std.mem.copyForwards(u8, dst[i .. len - 1], dst[i + 1 .. len]);
                len -= 1;
            },
            2 => {
                // Byte insert at random position
                if (len + 1 > dst.len) continue;
                const i = rng.intRangeAtMost(usize, 0, len);
                std.mem.copyBackwards(u8, dst[i + 1 .. len + 1], dst[i..len]);
                dst[i] = rng.int(u8);
                len += 1;
            },
            3 => {
                // Duplicate a short run from within the current mutated
                // content. The bounds must be drawn from `len` — not
                // `src.len` — because mutation may have grown dst past
                // src's tail.
                const run = rng.intRangeAtMost(usize, 1, @min(16, len));
                if (len + run > dst.len) continue;
                const from = rng.intRangeAtMost(usize, 0, len - run);
                const to = rng.intRangeAtMost(usize, 0, len);
                // Save the run before shifting (the shift may clobber it).
                var saved: [16]u8 = undefined;
                @memcpy(saved[0..run], dst[from .. from + run]);
                std.mem.copyBackwards(u8, dst[to + run .. len + run], dst[to..len]);
                @memcpy(dst[to .. to + run], saved[0..run]);
                len += run;
            },
            else => unreachable,
        }
    }
    return dst[0..len];
}

test "fuzz parseSharedStrings mutations" {
    const iters = fuzzIterations();
    var prng = std.Random.DefaultPrng.init(fuzzSeed());
    const rng = prng.random();
    var dst: [fuzz_max_input_len]u8 = undefined;
    for (0..iters) |_| {
        const input = mutate(rng, sst_template, &dst);
        var book: Book = .{ .allocator = std.testing.allocator, .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator) };
        defer book.deinit();
        const owned = std.testing.allocator.dupe(u8, input) catch continue;
        book.shared_strings_xml = owned;
        parseSharedStrings(&book, owned) catch {};
    }
}

test "fuzz parseWorkbookSheets mutations" {
    const iters = fuzzIterations();
    var prng = std.Random.DefaultPrng.init(fuzzSeed());
    const rng = prng.random();
    var wb_dst: [fuzz_max_input_len]u8 = undefined;
    var rels_dst: [fuzz_max_input_len]u8 = undefined;
    for (0..iters) |_| {
        const wb = mutate(rng, workbook_template, &wb_dst);
        const rels = mutate(rng, rels_template, &rels_dst);
        var book: Book = .{ .allocator = std.testing.allocator, .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator) };
        defer book.deinit();
        parseWorkbookSheets(&book, wb, rels) catch {};
    }
}

/// Realistic worksheet XML exercising every cell type (`inlineStr`,
/// shared-string `s`, `str`, boolean `b`, error `e`, numeric default),
/// empty cells, self-closing cells, multi-row layout, and a stray
/// `<f>` formula child.
const sheet_template =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" ++
    "<sheetData>" ++
    "<row r=\"1\">" ++
    "<c r=\"A1\" t=\"inlineStr\"><is><t>alpha</t></is></c>" ++
    "<c r=\"B1\" t=\"inlineStr\"><is><t>beta &amp; gamma</t></is></c>" ++
    "<c r=\"D1\"/>" ++ // self-closing
    "<c r=\"E1\" t=\"s\"><v>0</v></c>" ++
    "</row>" ++
    "<row r=\"2\">" ++
    "<c r=\"A2\"><v>42</v></c>" ++
    "<c r=\"B2\"><v>-3.14</v></c>" ++
    "<c r=\"C2\" t=\"b\"><v>1</v></c>" ++
    "<c r=\"D2\" t=\"e\"><v>#N/A</v></c>" ++
    "<c r=\"E2\" t=\"str\"><f>A1</f><v>computed</v></c>" ++
    "</row>" ++
    "<row r=\"3\"><c r=\"Z3\" t=\"inlineStr\"><is><r><rPr><b/></rPr><t>rich</t></r><r><t> text</t></r></is></c></row>" ++
    "</sheetData></worksheet>";

/// Run the row iterator end-to-end on an xml buffer. Used by multiple
/// fuzz tests — factored so both random-byte and mutation paths share
/// the same state-machine exerciser.
fn consumeAllRows(alloc: std.mem.Allocator, shared_strings: []const []const u8, xml: []const u8) void {
    var rows: Rows = .{
        .xml = xml,
        .pos = 0,
        .shared_strings = shared_strings,
        .allocator = alloc,
        .row_cells = .{},
        .arena = std.heap.ArenaAllocator.init(alloc),
    };
    defer rows.deinit();
    var count: usize = 0;
    while (rows.next() catch null) |_| : (count += 1) {
        if (count > 4096) break; // defensive cap against pathological loops
    }
}

test "fuzz Rows.next mutations on real sheet XML" {
    const iters = fuzzIterations();
    var prng = std.Random.DefaultPrng.init(fuzzSeed());
    const rng = prng.random();
    var dst: [fuzz_max_input_len]u8 = undefined;
    // A 1-element SST so cells with t="s" can resolve index 0.
    const sst = [_][]const u8{"shared-entry"};
    for (0..iters) |_| {
        const input = mutate(rng, sheet_template, &dst);
        consumeAllRows(std.testing.allocator, &sst, input);
    }
}

test "fuzz Rows.next on random bytes" {
    const iters = fuzzIterations();
    var prng = std.Random.DefaultPrng.init(fuzzSeed());
    const rng = prng.random();
    var buf: [fuzz_max_input_len]u8 = undefined;
    const sst = [_][]const u8{"shared-entry"};
    for (0..iters) |_| {
        const input = randomInput(rng, &buf);
        consumeAllRows(std.testing.allocator, &sst, input);
    }
}

test "fuzz appendDecoded mutations" {
    // Entity-dense template to exercise every branch of the decoder.
    const entity_template = "prefix &amp; &lt; &gt; &quot; &apos; &#233; &#xE9; &unknown; trailing";
    const iters = fuzzIterations();
    var prng = std.Random.DefaultPrng.init(fuzzSeed());
    const rng = prng.random();
    var dst: [512]u8 = undefined;
    for (0..iters) |_| {
        const input = mutate(rng, entity_template, &dst);
        var out: std.ArrayListUnmanaged(u8) = .{};
        defer out.deinit(std.testing.allocator);
        appendDecoded(std.testing.allocator, &out, input) catch {};
    }
}

test "fuzz Book.open against arbitrary bytes" {
    // Almost every byte-string is rejected at the zip signature check,
    // but the error path itself must be crash-free. A small fraction
    // of inputs will accidentally pass the zip header and exercise the
    // XML parsers downstream.
    const iters = fuzzIterations() / 4; // file IO is expensive; scale down
    var prng = std.Random.DefaultPrng.init(fuzzSeed());
    const rng = prng.random();
    var buf: [fuzz_max_input_len]u8 = undefined;
    for (0..iters) |_| {
        const input = randomInput(rng, &buf);

        var tmp = std.testing.tmpDir(.{});
        defer tmp.cleanup();
        tmp.dir.writeFile(.{ .sub_path = "fuzz.xlsx", .data = input }) catch continue;
        const path = tmp.dir.realpathAlloc(std.testing.allocator, "fuzz.xlsx") catch continue;
        defer std.testing.allocator.free(path);

        var book = Book.open(std.testing.allocator, path) catch continue;
        defer book.deinit();

        for (book.sheets) |sheet| {
            var rows = book.rows(sheet, std.testing.allocator) catch continue;
            defer rows.deinit();
            var count: usize = 0;
            while (rows.next() catch null) |_| : (count += 1) {
                if (count > 64) break;
            }
        }
    }
}

// ─── Writer re-exports ───────────────────────────────────────────────
//
// Expose the writer from the public zlsx module so downstream consumers
// can do `@import("zlsx").Writer` and `@import("zlsx").SheetWriter`. The
// writer lives in src/writer.zig and imports xlsx.Cell back; Zig handles
// this mutual import cleanly because neither side introspects the other
// at comptime.

pub const Writer = @import("writer.zig").Writer;
pub const SheetWriter = @import("writer.zig").SheetWriter;

// ─── Deep reader fuzz: SST index pointing past the SST table ────────
//
// A malicious xlsx could reference a shared-string index that's beyond
// the SST table's length. The reader's `Rows.next` resolves the index,
// so it must bounds-check. This fuzz target synthesizes cells with
// deliberately-high SST indices and confirms the reader doesn't crash.

test "fuzz Rows.next on synthetic cells with out-of-range SST indices" {
    const iters = fuzzIterations();
    var prng = std.Random.DefaultPrng.init(fuzzSeed());
    const rng = prng.random();

    // Build a tiny SST of 3 entries (indices 0..2 valid).
    const sst_entries = [_][]const u8{ "alpha", "beta", "gamma" };

    // For each iteration, build synthetic sheet XML with a cell whose
    // `<v>` value is a random u32 (likely out-of-range). Rows.next
    // must either surface a MalformedXml error or return an empty
    // slice — never crash.
    var sheet_xml_buf: std.ArrayListUnmanaged(u8) = .{};
    defer sheet_xml_buf.deinit(std.testing.allocator);

    for (0..iters) |_| {
        sheet_xml_buf.clearRetainingCapacity();
        try sheet_xml_buf.appendSlice(std.testing.allocator, "<sheetData>");
        const n_cells = rng.intRangeAtMost(usize, 1, 8);
        for (0..n_cells) |i| {
            const col_letter: u8 = @intCast('A' + (i % 26));
            const idx = rng.int(u32);
            try sheet_xml_buf.print(
                std.testing.allocator,
                "<row r=\"1\"><c r=\"{c}1\" t=\"s\"><v>{d}</v></c></row>",
                .{ col_letter, idx },
            );
        }
        try sheet_xml_buf.appendSlice(std.testing.allocator, "</sheetData>");

        var rows: Rows = .{
            .xml = sheet_xml_buf.items,
            .pos = 0,
            .shared_strings = &sst_entries,
            .allocator = std.testing.allocator,
            .row_cells = .{},
            .arena = std.heap.ArenaAllocator.init(std.testing.allocator),
        };
        defer rows.deinit();

        // Consume the rows — may error, must not panic.
        while (rows.next() catch null) |_| {}
    }
}
