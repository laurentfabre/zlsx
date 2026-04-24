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

/// Gregorian `{year, month, day}` → days-since-1970-01-01 using
/// Hinnant's days_from_civil. Inverse of `daysSinceUnixEpochToYMD`.
/// Input must be a real calendar date (year 1..=9999, month 1..=12,
/// day 1..=31 with the month's actual max) — `toExcelSerial` gates
/// this upstream.
fn ymdToDaysSinceUnixEpoch(year: i32, month: u8, day: u8) i32 {
    const y: i32 = if (month <= 2) year - 1 else year;
    const era: i32 = if (y >= 0) @divFloor(y, 400) else @divFloor(y - 399, 400);
    const yoe: i32 = y - era * 400;
    const m: i32 = @intCast(month);
    const doy: i32 = @divFloor(153 * (if (m > 2) m - 3 else m + 9) + 2, 5) + @as(i32, @intCast(day)) - 1;
    const doe: i32 = yoe * 365 + @divFloor(yoe, 4) - @divFloor(yoe, 100) + doy;
    return era * 146097 + doe - 719468;
}

/// Inverse of `fromExcelSerial`: convert a calendar `DateTime` into
/// the Excel serial-date number that writes produce. Returns `null`
/// when the input is outside the round-trippable range the reader
/// decodes cleanly:
///   - year < 1900 or > 9999
///   - month / day / hour / minute / second outside their legal
///     Gregorian ranges (1..=12 / 1..=31-per-month / 0..=23 / 0..=59)
///   - dates on or before 1900-03-01 — fromExcelSerial rejects
///     serials < 61 because of the 1900 leap-year bug, so this
///     matches that exclusion and keeps write/read symmetric.
///
/// Pair with `Writer.addStyle(Style{.number_format="yyyy-mm-dd"})`
/// to emit a date cell that Excel displays correctly and the reader
/// decodes via `Rows.parseDate`.
pub fn toExcelSerial(dt: DateTime) ?f64 {
    if (dt.year < 1900 or dt.year > 9999) return null;
    if (dt.month < 1 or dt.month > 12) return null;
    if (dt.hour > 23 or dt.minute > 59 or dt.second > 59) return null;

    // Day-of-month bounds vary with month + leap year.
    const days_in_month = [_]u8{ 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
    var max_day = days_in_month[dt.month - 1];
    if (dt.month == 2) {
        const is_leap =
            (@mod(dt.year, 4) == 0 and @mod(dt.year, 100) != 0) or
            @mod(dt.year, 400) == 0;
        if (is_leap) max_day = 29;
    }
    if (dt.day < 1 or dt.day > max_day) return null;

    const days_since_unix = ymdToDaysSinceUnixEpoch(@intCast(dt.year), dt.month, dt.day);
    // Excel serial 1 = 1900-01-01, unix epoch is 1970-01-01 —
    // offset 25569 bridges them. Matches the constant used in
    // fromExcelSerial.
    const excel_days = days_since_unix + 25569;
    const time_frac: f64 =
        (@as(f64, @floatFromInt(dt.hour)) * 3600 +
            @as(f64, @floatFromInt(dt.minute)) * 60 +
            @as(f64, @floatFromInt(dt.second))) / 86400.0;
    const serial = @as(f64, @floatFromInt(excel_days)) + time_frac;

    // Match fromExcelSerial's lower-bound rejection — dates that
    // decode as null shouldn't encode back to a non-null serial.
    if (serial < 61.0) return null;
    return serial;
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
/// directly into the raw SST xml). `font_name` likewise borrows from
/// `sst_arena` when non-null. Theme colors (`<color theme="…"/>`) are
/// not resolved today — we only expose explicit `rgb="AARRGGBB"`
/// values.
pub const RichRun = struct {
    text: []const u8,
    bold: bool = false,
    italic: bool = false,
    /// ARGB color from `<color rgb="FFFFFFFF"/>`. Null when the run
    /// uses a theme color (`<color theme="…"/>` — not resolved here)
    /// or no color is declared at all.
    color_argb: ?u32 = null,
    /// Font size in points, from `<sz val="11"/>`. Null when absent.
    size: ?f32 = null,
    /// Font family name, from `<rFont val="Calibri"/>`. Empty string
    /// when absent (distinguishes "no rFont" from "rFont val=''").
    font_name: []const u8 = "",
};

/// One side of a cell border, parsed from `<left>/<right>/<top>/
/// <bottom>/<diagonal>` inside a `<border>` element. `style` is the
/// OOXML `style="…"` attribute (e.g. "thin", "medium", "thick",
/// "double", "dashed"); empty when the side has no border or the
/// element was self-closing. `color_argb` tracks the child
/// `<color rgb="…"/>`; theme / indexed colors aren't resolved.
pub const BorderSide = struct {
    style: []const u8 = "",
    color_argb: ?u32 = null,
};

/// Cell border, parsed from `xl/styles.xml` `<borders>`. Every side
/// is always present in the struct — absent sides have `style=""`.
/// The writer-side OOXML borders block covers the same five slots.
pub const Border = struct {
    left: BorderSide = .{},
    right: BorderSide = .{},
    top: BorderSide = .{},
    bottom: BorderSide = .{},
    diagonal: BorderSide = .{},
};

/// A cell comment / note parsed from `xl/comments*.xml`. `text` is
/// always the concatenated plain-text form. `runs` mirrors the SST
/// rich-text path — populated only when the source body had at least
/// one `<r>` wrapper with formatting (non-null → multi-run rich
/// comment; null → single-run plain text). All slices entity-decoded
/// and arena-owned.
pub const Comment = struct {
    top_left: CellRef,
    author: []const u8,
    text: []const u8,
    runs: ?[]const RichRun = null,
};

/// Cell fill, parsed from `xl/styles.xml` `<fills>`. `pattern` is the
/// OOXML `patternType` attribute (e.g. "none", "solid", "darkDown",
/// "gray125"); borrows from `styles_xml`. `fg_color_argb` / `bg_color_argb`
/// track `<fgColor rgb="…"/>` / `<bgColor rgb="…"/>` — theme and
/// indexed colors aren't resolved (null).
pub const Fill = struct {
    pattern: []const u8 = "none",
    fg_color_argb: ?u32 = null,
    bg_color_argb: ?u32 = null,
};

/// Font properties for a cell, parsed from `xl/styles.xml` `<fonts>`.
/// Shared with `RichRun` semantically but kept as a distinct type so
/// callers don't have to populate a meaningless `text` field. Theme
/// colors aren't resolved — only explicit `<color rgb="AARRGGBB"/>`
/// populates `color_argb`. `name` is empty when the font had no
/// `<name val="…"/>` child.
pub const Font = struct {
    bold: bool = false,
    italic: bool = false,
    color_argb: ?u32 = null,
    size: ?f32 = null,
    name: []const u8 = "",
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

/// Owner of the open zip file handle, its reader buffer, and cached
/// central-directory offsets for lazily-extractable parts. Heap-boxed
/// because `std.fs.File.Reader` embeds a pointer into its own buffer:
/// moving it by value (e.g. returning `Book` from `openLazy`) would
/// invalidate that pointer. One indirection sidesteps the whole class
/// of bug.
///
/// Slice A caches sheet / comments / vml offsets even though every
/// matching entry is still eagerly extracted in the same pass —
/// populating the maps now makes the slice-B swap to on-demand
/// extraction a one-line change per part.
const ZipArchive = struct {
    file: std.fs.File,
    reader_buf: [4096]u8 = undefined,
    reader: std.fs.File.Reader,
    sheet_offsets: std.StringHashMapUnmanaged(Entry) = .{},
    comments_offsets: std.StringHashMapUnmanaged(Entry) = .{},
    vml_offsets: std.StringHashMapUnmanaged(Entry) = .{},

    const Entry = struct {
        file_offset: u64,
        uncompressed_size: u64,
        compression_method: std.zip.CompressionMethod,
    };

    /// Open the file, walk the central directory once, populate offset
    /// maps. On success, `self` owns the file handle; `deinit` closes
    /// it. The file stays open for the caller's lifetime so subsequent
    /// `extractEntry` calls can re-seek.
    fn open(allocator: Allocator, archive_path: []const u8) !*ZipArchive {
        const self = try allocator.create(ZipArchive);
        errdefer allocator.destroy(self);

        self.* = .{
            .file = try std.fs.cwd().openFile(archive_path, .{}),
            .reader = undefined,
            .sheet_offsets = .{},
            .comments_offsets = .{},
            .vml_offsets = .{},
        };
        errdefer self.file.close();

        // reader embeds a pointer into self.reader_buf — must run after
        // the struct lives at its final heap address.
        self.reader = self.file.reader(&self.reader_buf);

        errdefer self.sheet_offsets.deinit(allocator);
        errdefer self.comments_offsets.deinit(allocator);
        errdefer self.vml_offsets.deinit(allocator);

        return self;
    }

    fn deinit(self: *ZipArchive, allocator: Allocator) void {
        self.sheet_offsets.deinit(allocator);
        self.comments_offsets.deinit(allocator);
        self.vml_offsets.deinit(allocator);
        self.file.close();
        allocator.destroy(self);
    }

    /// Re-extract the bytes of a previously-cached part by re-seeking
    /// into the archive. Looks in the sheet / comments / vml maps in
    /// that order. Returns `null` if the key isn't cached anywhere
    /// (caller bug: every key lives in some map from the openLazy walk,
    /// or it shouldn't have been handed to us).
    fn extractByKey(self: *ZipArchive, allocator: Allocator, key: []const u8) !?[]u8 {
        const cached = self.sheet_offsets.get(key) orelse
            self.comments_offsets.get(key) orelse
            self.vml_offsets.get(key) orelse return null;

        // Fabricate a minimal `std.zip.Iterator.Entry` so the existing
        // `extractEntryToBuffer` helper is re-usable. Only three fields
        // are read by the extractor (file_offset, uncompressed_size,
        // compression_method) — the rest stay zero-valued.
        const entry = std.zip.Iterator.Entry{
            .version_needed_to_extract = 0,
            .flags = @bitCast(@as(u16, 0)),
            .compression_method = cached.compression_method,
            .last_modification_time = 0,
            .last_modification_date = 0,
            .header_zip_offset = 0,
            .crc32 = 0,
            .filename_len = 0,
            .compressed_size = 0,
            .uncompressed_size = cached.uncompressed_size,
            .file_offset = cached.file_offset,
        };
        return try extractEntryToBuffer(allocator, entry, &self.reader);
    }
};

pub const Book = struct {
    allocator: Allocator,
    /// Open zip archive backing this book. Owns the file handle and
    /// reader. Set by `openLazy`; torn down LAST in `deinit` (after
    /// every borrowed part buffer is freed).
    archive: ?*ZipArchive = null,
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
    /// Raw comments XML parts keyed by their archive path
    /// (`xl/comments1.xml`, `xl/comments2.xml`, …). Retained so
    /// per-sheet rels → Target lookups can resolve to a buffer
    /// the parser borrows from.
    comments_data: std.StringHashMapUnmanaged([]u8) = .{},
    /// Comments per sheet, resolved at open time by following each
    /// sheet's rels → comments target. Keyed by sheet path; missing
    /// sheets use the `comments(sheet)` empty-slice normalisation.
    comments_by_sheet: std.StringHashMapUnmanaged([]Comment) = .{},
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
    /// Raw `xl/styles.xml` bytes (nullable — workbooks with no
    /// formatting omit this file). Kept around because the parsed
    /// tables below borrow from it for custom number-format strings.
    styles_xml: ?[]u8 = null,
    /// Raw `xl/theme/theme1.xml` bytes (nullable — minimal xlsx files
    /// skip the theme part). Kept around because the parsed theme
    /// color table below borrows from it.
    theme_xml: ?[]u8 = null,
    /// 12-entry theme color palette extracted from `xl/theme/theme1.xml`.
    /// Index order mirrors OOXML's `<color theme="N"/>` semantics:
    ///   0 = lt1, 1 = dk1, 2 = lt2, 3 = dk2,
    ///   4..9 = accent1..accent6,
    ///   10 = hlink, 11 = folHlink.
    /// Empty when no theme.xml or the parser failed — callers get
    /// `null` on `<color theme="N"/>` resolution in that case,
    /// matching pre-iter52 behavior.
    theme_colors: []u32 = &.{},
    /// `cellXfs` table — maps a cell's `s="N"` attribute to a numFmt
    /// id. Owned slice; index matches the xfId. Empty when the
    /// workbook has no styles.xml.
    cell_xf_numfmt_ids: []u32 = &.{},
    /// Custom numFmts, keyed by numFmtId. Built-in ids (0-163) aren't
    /// stored here — they're resolved via `builtinNumberFormat`.
    /// Values borrow from `styles_xml`.
    custom_num_fmts: std.AutoHashMapUnmanaged(u32, []const u8) = .{},
    /// `fonts` table from styles.xml — index matches `fontId`.
    /// Values borrow from `styles_xml` for the `name` slices.
    fonts: []Font = &.{},
    /// One fontId per `<xf>` under `<cellXfs>`, matching
    /// `cell_xf_numfmt_ids`. Use `Book.cellFont(style_idx)` for the
    /// resolved lookup.
    cell_xf_font_ids: []u32 = &.{},
    /// `fills` table from styles.xml — index matches `fillId`. Values
    /// borrow from `styles_xml` for `pattern` slices.
    fills: []Fill = &.{},
    /// One fillId per `<xf>` under `<cellXfs>`. Use
    /// `Book.cellFill(style_idx)` for the resolved lookup.
    cell_xf_fill_ids: []u32 = &.{},
    /// `borders` table from styles.xml — index matches `borderId`.
    /// BorderSide.style slices borrow from `styles_xml`.
    borders: []Border = &.{},
    /// One borderId per `<xf>` under `<cellXfs>`. Use
    /// `Book.cellBorder(style_idx)` for the resolved lookup.
    cell_xf_border_ids: []u32 = &.{},
    /// Owned backing storage for every string referenced by `sheets`,
    /// sheet_data keys, and entity-decoded shared strings.
    strings: std.ArrayListUnmanaged([]u8) = .{},

    /// Open and parse the workbook skeleton. Sheet XML is eagerly
    /// decompressed (xlsx files we target are small — ~300 KB — and
    /// streaming through std.zip is awkward).
    ///
    /// Facade over `openLazy` that releases the source file handle
    /// before returning, so callers retain today's behavior of being
    /// able to rename / overwrite / delete the source xlsx while the
    /// Book is in use (matters on Windows file locks; harmless
    /// elsewhere). Slice B will split `loadEagerParts` out of
    /// `openLazy` so only the `open` facade pays for it.
    pub fn open(allocator: Allocator, path: []const u8) !Book {
        var book = try Book.openLazy(allocator, path);
        errdefer book.deinit();
        try book.loadEagerParts();
        book.closeArchive();
        return book;
    }

    /// Close the backing archive early. Slice A's `open` facade calls
    /// this after eager loading so the source file doesn't stay locked
    /// for the Book's lifetime. Slice B's streaming path keeps the
    /// archive alive by not calling this.
    fn closeArchive(self: *Book) void {
        if (self.archive) |arch| {
            arch.deinit(self.allocator);
            self.archive = null;
        }
    }

    /// Open the archive, walk the central directory, populate
    /// **workbook-wide** state (sheets list, SST, theme, styles, per-
    /// sheet rels), but leave **per-sheet** XML and metadata unloaded
    /// until `rows(sheet)` / `preloadSheet(sheet)` is called. The
    /// archive file handle stays open for the Book's lifetime; use
    /// `Book.open` if the source file must be releasable immediately.
    ///
    /// Getter-contract:
    /// - `numberFormat`, `cellFont`, `cellFill`, `cellBorder`,
    ///   `isDateFormat`, `richRuns`, `sharedStrings` — **always**
    ///   work, populated by `openLazy`.
    /// - `mergedRanges(sheet)`, `hyperlinks(sheet)`,
    ///   `dataValidations(sheet)`, `comments(sheet)` — return the
    ///   empty slice for a sheet that has not yet been loaded. Call
    ///   `preloadSheet(sheet)` or any `rows(sheet)` iteration first
    ///   to populate. `Book.open` preloads everything on open, so
    ///   non-lazy callers never see this.
    pub fn openLazy(allocator: Allocator, path: []const u8) !Book {
        var book: Book = .{
            .allocator = allocator,
            .sst_arena = std.heap.ArenaAllocator.init(allocator),
        };
        errdefer book.deinit();

        book.archive = try ZipArchive.open(allocator, path);

        var iter = std.zip.Iterator.init(&book.archive.?.reader) catch return error.BadZip;

        // We need three categories of files from the archive:
        //   xl/sharedStrings.xml     — optional
        //   xl/workbook.xml          — required
        //   xl/_rels/workbook.xml.rels — required (sheet id → target path)
        //   xl/worksheets/sheet*.xml — data
        var rels_xml: ?[]u8 = null;
        defer if (rels_xml) |r| allocator.free(r);

        var workbook_xml: ?[]u8 = null;
        defer if (workbook_xml) |w| allocator.free(w);

        var filename_buf: [512]u8 = undefined;
        const archive = book.archive.?;
        const file_reader = &archive.reader;

        while (iter.next() catch return error.BadZip) |entry| {
            if (entry.filename_len == 0 or entry.filename_len > filename_buf.len) continue;

            // Read filename (lives in the CDFH after the fixed-size header).
            try file_reader.seekTo(entry.header_zip_offset + @sizeOf(std.zip.CentralDirectoryFileHeader));
            const filename = filename_buf[0..entry.filename_len];
            file_reader.interface.readSliceAll(filename) catch return error.BadZip;

            const cached: ZipArchive.Entry = .{
                .file_offset = entry.file_offset,
                .uncompressed_size = entry.uncompressed_size,
                .compression_method = entry.compression_method,
            };

            if (std.mem.eql(u8, filename, "xl/sharedStrings.xml")) {
                book.shared_strings_xml = try extractEntryToBuffer(allocator, entry, file_reader);
            } else if (std.mem.eql(u8, filename, "xl/styles.xml")) {
                book.styles_xml = try extractEntryToBuffer(allocator, entry, file_reader);
            } else if (std.mem.eql(u8, filename, "xl/theme/theme1.xml")) {
                book.theme_xml = try extractEntryToBuffer(allocator, entry, file_reader);
            } else if (std.mem.startsWith(u8, filename, "xl/comments") and
                std.mem.endsWith(u8, filename, ".xml"))
            {
                // Cache the offset only — comments are extracted on
                // demand by `ensureCommentsLoaded`. `Book.open`'s
                // facade walks every sheet through `loadEagerParts`,
                // which pulls these in via the rels resolver.
                const key = try allocator.dupe(u8, filename);
                errdefer allocator.free(key);
                try book.strings.append(allocator, key);
                try archive.comments_offsets.put(allocator, key, cached);
            } else if (std.mem.eql(u8, filename, "xl/workbook.xml")) {
                workbook_xml = try extractEntryToBuffer(allocator, entry, file_reader);
            } else if (std.mem.eql(u8, filename, "xl/_rels/workbook.xml.rels")) {
                rels_xml = try extractEntryToBuffer(allocator, entry, file_reader);
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
                const data = try extractEntryToBuffer(allocator, entry, file_reader);
                try book.sheet_rels_data.put(allocator, sheet_key, data);
            } else if (std.mem.startsWith(u8, filename, "xl/worksheets/") and
                std.mem.endsWith(u8, filename, ".xml"))
            {
                // Cache the offset only — sheet XML is extracted by
                // `ensureSheetLoaded` (called from `Book.rows` on the
                // lazy path, or iterated eagerly by `loadEagerParts`
                // on the `Book.open` facade). Own the key outright so
                // the HashMap entry stays valid for the Book's lifetime.
                const key = try allocator.dupe(u8, filename);
                errdefer allocator.free(key);
                try book.strings.append(allocator, key);
                try archive.sheet_offsets.put(allocator, key, cached);
            } else if (std.mem.startsWith(u8, filename, "xl/drawings/vmlDrawing") and
                std.mem.endsWith(u8, filename, ".vml"))
            {
                // VML offsets are cached for slice B's lazy comments
                // path. Today we don't load vml drawings eagerly, so
                // just stash the offset (key owned by book.strings).
                const key = try allocator.dupe(u8, filename);
                errdefer allocator.free(key);
                try book.strings.append(allocator, key);
                try archive.vml_offsets.put(allocator, key, cached);
            }
        }

        const wb = workbook_xml orelse return error.MissingWorkbook;
        const rels = rels_xml orelse return error.MissingWorkbook;

        try parseWorkbookSheets(&book, wb, rels);
        if (book.shared_strings_xml) |sst| try parseSharedStrings(&book, sst);

        // Workbook-wide style + theme tables MUST parse here (not in
        // `loadEagerParts`): every `numberFormat`, `cellFont`,
        // `cellFill`, `cellBorder`, `isDateFormat` lookup is a random-
        // access read against these tables, independent of which sheet
        // the caller is on. Parsing them during `openLazy` means
        // callers who want just workbook metadata never pay for sheet
        // extraction. Theme colors MUST parse before styles.xml so
        // style parsers can resolve `<color theme="N"/>` references.
        if (book.theme_xml) |tx| try parseTheme(&book, tx);
        if (book.styles_xml) |sx| try parseStyles(&book, sx);

        // NOTE: per-sheet metadata (merged ranges, hyperlinks,
        // validations, comments) is NOT populated here. `Book.open`
        // calls `loadEagerParts` explicitly after `openLazy` returns.
        // Streaming / on-demand callers must call `preloadSheet(s)`
        // or `rows(s)` to populate each sheet's side-indices.
        return book;
    }

    /// Populate per-sheet side-indices (merged ranges, hyperlinks,
    /// data validations, comments) for every sheet in the workbook.
    /// Split from `openLazy` so slice B's streaming path can skip it
    /// on per-sheet-iteration workloads.
    ///
    /// Intentionally private + non-idempotent: calling twice would
    /// leak the previously parsed per-sheet hashmaps. The only valid
    /// caller is `Book.open`.
    fn loadEagerParts(self: *Book) !void {
        // Iterate declared sheets rather than the (now empty) cache.
        // Each `ensureSheetLoaded` extracts the sheet XML on demand
        // and runs the merged / hyperlinks / validations / comments
        // parsers once — same net work the pre-slice-B path did, just
        // routed through the single entry point the lazy path uses.
        for (self.sheets) |sheet| {
            _ = try self.ensureSheetLoaded(sheet.path);
        }
    }

    /// Force per-sheet side-indices (merged ranges, hyperlinks,
    /// validations, comments) for a single sheet. Intended for
    /// `openLazy` callers that want the metadata getters
    /// (`mergedRanges`, `hyperlinks`, `dataValidations`, `comments`)
    /// to return populated slices for a specific sheet without having
    /// to iterate its rows.
    ///
    /// Idempotent: second call on the same sheet is a hashmap hit.
    /// After `Book.open`, every sheet has already been preloaded, so
    /// calling this is a no-op.
    pub fn preloadSheet(self: *Book, sheet: Sheet) !void {
        _ = try self.ensureSheetLoaded(sheet.path);
    }

    /// Return the XML bytes for `sheet_path`, extracting from the
    /// archive on first access and running the per-sheet side-index
    /// parsers (merged ranges, hyperlinks, data validations, comments).
    ///
    /// Idempotent: gated on `sheet_data.contains(path)` so a second
    /// call is a pure hashmap hit and does not re-parse the side
    /// indices. If extraction succeeds but a downstream parser fails,
    /// the sheet XML stays in the cache — partial side-index state
    /// is tolerated because each parser is no-op-on-missing and a
    /// retry via the `contains` gate short-circuits before re-entering
    /// the parsers.
    ///
    /// Returns `error.ArchiveClosed` when the book was opened via
    /// `Book.open` (which closes the archive after `loadEagerParts`)
    /// AND the sheet wasn't cached by that prior load — impossible
    /// in practice, but the error path exists for correctness.
    fn ensureSheetLoaded(self: *Book, sheet_path: []const u8) ![]const u8 {
        if (self.sheet_data.get(sheet_path)) |existing| return existing;

        const archive = self.archive orelse return error.ArchiveClosed;
        const data = (try archive.extractByKey(self.allocator, sheet_path)) orelse
            return error.MissingSheet;

        // `committed` flips once `sheet_data.put` succeeds — from that
        // point on, ownership of `data` + `owned_key` has moved into
        // the book and the allocation-cleanup errdefers must become
        // no-ops. Otherwise a later parser error would free memory the
        // book's `sheet_data` still references, causing a double-free
        // on `deinit` via the valueIterator loop.
        var committed = false;
        errdefer if (!committed) self.allocator.free(data);

        // The caller's `sheet_path` may be a transient slice (e.g. a
        // fresh allocation from parseWorkbookSheets). Dupe into
        // `self.strings` so the hashmap key lifetime matches the book
        // instead of the caller.
        const owned_key = try self.allocator.dupe(u8, sheet_path);
        errdefer if (!committed) self.allocator.free(owned_key);
        try self.strings.append(self.allocator, owned_key);
        errdefer if (!committed) {
            _ = self.strings.pop();
        };

        try self.sheet_data.put(self.allocator, owned_key, data);
        committed = true;

        // From here on, `sheet_data` owns `data` + `owned_key`. Partial
        // side-index state on a later parse error is tolerated — each
        // parser is no-op-on-missing and a retry via the `contains`
        // gate short-circuits before re-entering them.
        try parseMergedRangesForSheet(self, owned_key, data);
        try parseHyperlinksForSheet(self, owned_key, data);
        try parseDataValidationsForSheet(self, owned_key, data);
        try self.ensureCommentsLoadedForSheet(owned_key);
        try parseCommentsForSheet(self, owned_key);

        return data;
    }

    /// Resolve this sheet's rels to a comments part (if any) and
    /// ensure the comments XML is extracted into `comments_data`.
    /// Idempotent: repeated calls are a hashmap hit.
    fn ensureCommentsLoadedForSheet(self: *Book, sheet_path: []const u8) !void {
        const rels_xml = self.sheet_rels_data.get(sheet_path) orelse return;

        // Scan the rels file for a Target ending in /comments*.xml —
        // same rule as `parseCommentsForSheet`. We only need to walk
        // until we find the first one (a sheet carries at most one
        // comments part).
        const target_key = "Target=\"";
        var i: usize = 0;
        while (std.mem.indexOfPos(u8, rels_xml, i, target_key)) |tp| {
            const s = tp + target_key.len;
            const e = std.mem.indexOfScalarPos(u8, rels_xml, s, '"') orelse return;
            i = e + 1;
            const target = rels_xml[s..e];
            if (std.mem.indexOf(u8, target, "comments") == null) continue;
            const basename = if (std.mem.lastIndexOfScalar(u8, target, '/')) |slash|
                target[slash + 1 ..]
            else
                target;
            var path_buf: [64]u8 = undefined;
            const full = std.fmt.bufPrint(&path_buf, "xl/{s}", .{basename}) catch return;
            try self.ensureCommentsLoaded(full);
            return;
        }
    }

    /// Extract a comments part by archive path. Idempotent.
    fn ensureCommentsLoaded(self: *Book, comments_path: []const u8) !void {
        if (self.comments_data.contains(comments_path)) return;

        const archive = self.archive orelse return error.ArchiveClosed;
        const data = (try archive.extractByKey(self.allocator, comments_path)) orelse return;
        errdefer self.allocator.free(data);

        // Key must outlive the caller's buffer; the openLazy walk
        // already allocated a stable key in `self.strings`, so prefer
        // that slice. Fall back to a fresh dupe if we somehow hit a
        // path the walk didn't cache (shouldn't happen — extractByKey
        // would have returned null first).
        const stable_key: []const u8 = blk: {
            if (archive.comments_offsets.getKey(comments_path)) |k| break :blk k;
            const dup = try self.allocator.dupe(u8, comments_path);
            errdefer self.allocator.free(dup);
            try self.strings.append(self.allocator, dup);
            break :blk dup;
        };
        try self.comments_data.put(self.allocator, stable_key, data);
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

    /// Number-format code for a cell's style index (the value of the
    /// `s="…"` attribute, surfaced via `Rows.styleIndices()`). Returns
    /// null when the workbook has no styles.xml, the index is out of
    /// range, or the resolved numFmt id has no built-in mapping and
    /// no custom override (rare — typically means a malformed file).
    ///
    /// The returned slice borrows from `styles_xml` (custom formats)
    /// or a constant string table (built-ins); valid for the Book's
    /// lifetime either way.
    pub fn numberFormat(self: *const Book, style_idx: u32) ?[]const u8 {
        if (style_idx >= self.cell_xf_numfmt_ids.len) return null;
        const fmt_id = self.cell_xf_numfmt_ids[style_idx];
        if (self.custom_num_fmts.get(fmt_id)) |code| return code;
        return builtinNumberFormat(fmt_id);
    }

    /// True when the resolved number-format code is a date / time /
    /// date-time pattern. Callers can combine with
    /// `xlsx.fromExcelSerial(cell.number)` to auto-materialise a
    /// `DateTime` without having to interpret the pattern themselves.
    pub fn isDateFormat(self: *const Book, style_idx: u32) bool {
        const code = self.numberFormat(style_idx) orelse return false;
        return isDateFormatCode(code);
    }

    /// Font properties for a cell, resolved via the style index
    /// (`Rows.styleIndices()`). Returns null when the workbook has
    /// no styles.xml, the index is out of range, or the referenced
    /// fontId doesn't resolve (malformed file). `name` slice borrows
    /// from `styles_xml`; lifetime matches the Book.
    pub fn cellFont(self: *const Book, style_idx: u32) ?Font {
        if (style_idx >= self.cell_xf_font_ids.len) return null;
        const font_id = self.cell_xf_font_ids[style_idx];
        if (font_id >= self.fonts.len) return null;
        return self.fonts[font_id];
    }

    /// Fill properties for a cell, resolved via the style index.
    /// Returns null when the workbook has no styles.xml, the index is
    /// out of range, or the referenced fillId doesn't resolve. Unlike
    /// `cellFont` the common case returns a Fill with `pattern="none"`
    /// — absence of any `<patternFill>` colors is valid and means
    /// "no fill" rather than "malformed".
    pub fn cellFill(self: *const Book, style_idx: u32) ?Fill {
        if (style_idx >= self.cell_xf_fill_ids.len) return null;
        const fill_id = self.cell_xf_fill_ids[style_idx];
        if (fill_id >= self.fills.len) return null;
        return self.fills[fill_id];
    }

    /// Cell comments attached to `sheet` (from `xl/comments*.xml`
    /// discovered via the sheet's rels). Empty for sheets without
    /// comments. Returned slice, comment strings, and author strings
    /// are all owned by the Book.
    pub fn comments(self: *const Book, sheet: Sheet) []const Comment {
        return self.comments_by_sheet.get(sheet.path) orelse &.{};
    }

    /// Border properties for a cell, resolved via the style index.
    /// Returns null on out-of-range indices or workbooks without
    /// styles.xml. Sides without a border surface with `style=""`.
    pub fn cellBorder(self: *const Book, style_idx: u32) ?Border {
        if (style_idx >= self.cell_xf_border_ids.len) return null;
        const border_id = self.cell_xf_border_ids[style_idx];
        if (border_id >= self.borders.len) return null;
        return self.borders[border_id];
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

        var cdit = self.comments_data.valueIterator();
        while (cdit.next()) |v| a.free(v.*);
        self.comments_data.deinit(a);

        var cmit = self.comments_by_sheet.valueIterator();
        while (cmit.next()) |v| a.free(v.*);
        self.comments_by_sheet.deinit(a);

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

        // Styles table — numFmt values borrow from styles_xml; only
        // the xfIds slice + hashmap spine are owned here.
        a.free(self.cell_xf_numfmt_ids);
        a.free(self.cell_xf_font_ids);
        a.free(self.cell_xf_fill_ids);
        a.free(self.cell_xf_border_ids);
        a.free(self.fonts);
        a.free(self.fills);
        a.free(self.borders);
        self.custom_num_fmts.deinit(a);
        if (self.styles_xml) |s| a.free(s);
        if (self.theme_xml) |s| a.free(s);
        a.free(self.theme_colors);

        var it = self.sheet_data.valueIterator();
        while (it.next()) |v| a.free(v.*);
        self.sheet_data.deinit(a);

        a.free(self.sheets);
        for (self.strings.items) |s| a.free(s);
        self.strings.deinit(a);

        // archive teardown runs LAST: its offset maps key-borrow from
        // `self.strings` (freed above) and its file handle is irrelevant
        // to the part buffers, but closing it before those frees is
        // harmless — we still do it last to document the intended order
        // and to keep slice B's lazy-extraction path trivially correct
        // (borrowed buffers must all be freed before the file closes).
        if (self.archive) |arch| arch.deinit(a);
        self.archive = null;

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
    ///
    /// Takes `*Book` (not `*const Book`) as of iter54 slice B: the
    /// lazy path populates `sheet_data` and the per-sheet side indices
    /// on first call. Rows itself still only borrows from the book.
    pub fn rows(self: *Book, sheet: Sheet, allocator: Allocator) !Rows {
        const xml = try self.ensureSheetLoaded(sheet.path);
        return .{
            .xml = xml,
            .pos = 0,
            .shared_strings = self.shared_strings,
            .book = self,
            .allocator = allocator,
            .row_cells = .{},
            .row_styles = .{},
            .arena = std.heap.ArenaAllocator.init(allocator),
        };
    }
};

// ─── Rows iterator ───────────────────────────────────────────────────

pub const Rows = struct {
    xml: []const u8,
    pos: usize,
    shared_strings: []const []const u8,
    /// Weak reference into the Book so `parseDate` can check the
    /// cell's style index against the workbook's numFmt table. Rows
    /// never outlive their parent Book (constructed via
    /// `Book.rows(sheet, alloc)`), so the pointer is always live
    /// during iteration. Nullable only so internal fuzz helpers can
    /// drive the state machine without a Book; public callers always
    /// get a valid pointer from `Book.rows`.
    book: ?*const Book = null,
    allocator: Allocator,
    row_cells: std.ArrayListUnmanaged(Cell),
    /// Parallel to `row_cells`. Each slot holds the cell's `s="N"`
    /// attribute (the index into `Book.cell_xf_numfmt_ids`), or null
    /// when the source `<c>` had no `s` attribute. Filled with nulls
    /// for gap cells so positional indexing matches `row_cells`.
    row_styles: std.ArrayListUnmanaged(?u32),
    /// Bump arena for per-row decoded strings. Reset (O(1)) at the
    /// start of each `next()` call — previous row's owned strings
    /// become invalid, which matches the documented contract. Compared
    /// to the older per-string malloc/free list this saves ~one free
    /// per entity-bearing or rich-text cell per row.
    arena: std.heap.ArenaAllocator,

    pub fn deinit(self: *Rows) void {
        self.arena.deinit();
        self.row_cells.deinit(self.allocator);
        self.row_styles.deinit(self.allocator);
        self.* = undefined;
    }

    /// Style indices for the current row — one slot per `row_cells`
    /// slot, mirroring the same gap-filled layout. `null` means the
    /// `<c>` had no `s` attribute (i.e. the General format). Valid
    /// until the next `next()` call, same as `row_cells`.
    pub fn styleIndices(self: *const Rows) []const ?u32 {
        return self.row_styles.items;
    }

    /// Parse the current-row cell at `col_idx` as a date-styled
    /// number. Returns a `DateTime` when all three conditions hold:
    ///   - the cell is `.number` or `.integer`
    ///   - the cell's style index resolves to a date-like numFmt
    ///     (per `Book.isDateFormat`)
    ///   - the serial is in the valid Excel range (>= 61 per the
    ///     1900 leap-year bug exclusion, < 2958466)
    /// Otherwise returns null. Use this instead of manually chaining
    /// `styleIndices()` + `Book.isDateFormat` + `fromExcelSerial`.
    ///
    /// Valid until the next `next()` call (row lifetime).
    pub fn parseDate(self: *const Rows, col_idx: usize) ?DateTime {
        if (col_idx >= self.row_cells.items.len) return null;
        const cell = self.row_cells.items[col_idx];
        const num: f64 = switch (cell) {
            .number => |n| n,
            .integer => |n| @floatFromInt(n),
            else => return null,
        };
        if (col_idx >= self.row_styles.items.len) return null;
        const style_idx = self.row_styles.items[col_idx] orelse return null;
        const book = self.book orelse return null;
        if (!book.isDateFormat(style_idx)) return null;
        return fromExcelSerial(num);
    }

    /// Returns the next row's cells, or null at end-of-sheet. Returned
    /// slice is valid until the next call to `next()` (or until
    /// `deinit()`). Cell string contents are either shared-string slices
    /// (owned by the Book), xml-backed slices (stable for the Book's
    /// lifetime), or row-owned slices that are invalidated on the next
    /// call (arena reset).
    pub fn next(self: *Rows) !?[]const Cell {
        self.row_cells.clearRetainingCapacity();
        self.row_styles.clearRetainingCapacity();
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
        const style_attr = getAttr(attrs, "s");
        const style_idx: ?u32 = if (style_attr) |s|
            std.fmt.parseInt(u32, s, 10) catch null
        else
            null;

        // Grow row_cells + row_styles to cover col_idx; fill gaps
        // with `.empty` / `null`. Both arrays stay in lock-step so
        // callers can index them the same way.
        while (self.row_cells.items.len <= col_idx) {
            try self.row_cells.append(self.allocator, .empty);
            try self.row_styles.append(self.allocator, null);
        }
        self.row_styles.items[col_idx] = style_idx;

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
/// Resolve a sheet's `_rels` to find the `xl/comments*.xml` target
/// (if any), parse it into a `Comment[]`, and stash under
/// `book.comments_by_sheet`. No-op when the sheet has no rels or
/// no comments relationship. Authors/texts are copied into
/// `sst_arena` so they live for the Book's lifetime.
fn parseCommentsForSheet(book: *Book, sheet_path: []const u8) !void {
    const rels_xml = book.sheet_rels_data.get(sheet_path) orelse return;

    // Scan rels for a Target whose path ends with /comments*.xml.
    // OOXML uses a relationship Type of ".../relationships/comments"
    // but a Target suffix match is robust enough for our needs.
    const target_key = "Target=\"";
    var comments_xml: ?[]const u8 = null;
    var i: usize = 0;
    while (std.mem.indexOfPos(u8, rels_xml, i, target_key)) |tp| {
        const s = tp + target_key.len;
        const e = std.mem.indexOfScalarPos(u8, rels_xml, s, '"') orelse break;
        i = e + 1;
        const target = rels_xml[s..e];
        if (std.mem.indexOf(u8, target, "comments") == null) continue;
        // Targets in sheet rels are typically "../comments1.xml" —
        // normalise to "xl/comments1.xml" for the map lookup.
        const basename = if (std.mem.lastIndexOfScalar(u8, target, '/')) |slash|
            target[slash + 1 ..]
        else
            target;
        var path_buf: [64]u8 = undefined;
        const full = std.fmt.bufPrint(&path_buf, "xl/{s}", .{basename}) catch continue;
        if (book.comments_data.get(full)) |data| {
            comments_xml = data;
            break;
        }
    }
    const cxml = comments_xml orelse return;

    // Parse <authors>: one per <author>…</author>.
    var authors: std.ArrayListUnmanaged([]const u8) = .{};
    defer authors.deinit(book.allocator);
    const arena = book.sst_arena.allocator();
    if (std.mem.indexOf(u8, cxml, "<authors>")) |ap| {
        const close = std.mem.indexOfPos(u8, cxml, ap, "</authors>") orelse cxml.len;
        const block = cxml[ap..close];
        var j: usize = 0;
        while (std.mem.indexOfPos(u8, block, j, "<author>")) |op| {
            const body_start = op + "<author>".len;
            const close_tag = std.mem.indexOfPos(u8, block, body_start, "</author>") orelse break;
            const raw = block[body_start..close_tag];
            j = close_tag + "</author>".len;
            try authors.append(book.allocator, try arenaDupeDecoded(arena, raw));
        }
    }

    // Parse <commentList>: one <comment ref="A1" authorId="N">
    // <text>…</text></comment> per entry. Comment text may contain
    // rich runs — concatenate all <t>…</t> contents for the flat
    // `text` field.
    var entries: std.ArrayListUnmanaged(Comment) = .{};
    errdefer entries.deinit(book.allocator);
    const cl_pos = std.mem.indexOf(u8, cxml, "<commentList>") orelse return;
    const cl_end = std.mem.indexOfPos(u8, cxml, cl_pos, "</commentList>") orelse return;
    const cl_block = cxml[cl_pos..cl_end];
    var k: usize = 0;
    while (std.mem.indexOfPos(u8, cl_block, k, "<comment ")) |cp| {
        const hdr_end = std.mem.indexOfScalarPos(u8, cl_block, cp, '>') orelse break;
        const attrs = cl_block[cp..hdr_end];
        k = hdr_end + 1;

        const ref_key = "ref=\"";
        const ref_pos = std.mem.indexOf(u8, attrs, ref_key) orelse continue;
        const rs = ref_pos + ref_key.len;
        const re = std.mem.indexOfScalarPos(u8, attrs, rs, '"') orelse continue;
        const ref_str = attrs[rs..re];
        const cell_ref = parseA1Ref(ref_str) catch continue;

        const aid_key = "authorId=\"";
        var author_id: usize = 0;
        if (std.mem.indexOf(u8, attrs, aid_key)) |ap| {
            const as = ap + aid_key.len;
            const ae = std.mem.indexOfScalarPos(u8, attrs, as, '"') orelse continue;
            author_id = std.fmt.parseInt(usize, attrs[as..ae], 10) catch 0;
        }
        const author_str: []const u8 = if (author_id < authors.items.len)
            authors.items[author_id]
        else
            "";

        // Body: concat every <t>…</t> inside <text>…</text>.
        const c_close = std.mem.indexOfPos(u8, cl_block, hdr_end, "</comment>") orelse break;
        const body = cl_block[hdr_end + 1 .. c_close];
        k = c_close + "</comment>".len;

        // Walk the comment body collecting both a flat-text concat
        // AND per-run metadata when the body is rich-text (any `<r>`).
        // Mirrors the iter26 SST parser's structure: `<r>` toggles
        // `saw_r`; `<rPr>` populates `pending_flags`; each `<t>`
        // consumed inside `<r>` appends a RichRun.
        var text_buf: std.ArrayListUnmanaged(u8) = .{};
        defer text_buf.deinit(arena);
        var runs_list: std.ArrayListUnmanaged(RichRun) = .{};
        errdefer runs_list.deinit(arena);
        var saw_r = false;
        var pending_flags: RichRun = .{ .text = "" };

        var t: usize = 0;
        while (t < body.len) {
            const next_lt = std.mem.indexOfScalarPos(u8, body, t, '<') orelse break;
            if (next_lt + 2 > body.len) break;
            const c1 = body[next_lt + 1];

            if (c1 == 't' and next_lt + 2 < body.len and
                (body[next_lt + 2] == '>' or body[next_lt + 2] == ' ' or body[next_lt + 2] == '/'))
            {
                // `<t>`, `<t xml:space="…">`, or `<t/>`.
                const tgt = std.mem.indexOfScalarPos(u8, body, next_lt + 2, '>') orelse break;
                if (tgt > 0 and body[tgt - 1] == '/') {
                    t = tgt + 1;
                    continue;
                }
                const tclose = std.mem.indexOfPos(u8, body, tgt + 1, "</t>") orelse break;
                const span = body[tgt + 1 .. tclose];
                const span_has_ent = std.mem.indexOfScalar(u8, span, '&') != null;

                try appendDecoded(arena, &text_buf, span);

                if (saw_r) {
                    const run_text: []const u8 = if (span_has_ent) blk: {
                        var rb: std.ArrayListUnmanaged(u8) = try .initCapacity(arena, span.len);
                        try appendDecoded(arena, &rb, span);
                        break :blk try rb.toOwnedSlice(arena);
                    } else span;
                    try runs_list.append(arena, .{
                        .text = run_text,
                        .bold = pending_flags.bold,
                        .italic = pending_flags.italic,
                        .color_argb = pending_flags.color_argb,
                        .size = pending_flags.size,
                        .font_name = pending_flags.font_name,
                    });
                }
                t = tclose + "</t>".len;
            } else if (c1 == 'r' and next_lt + 2 < body.len and
                (body[next_lt + 2] == '>' or body[next_lt + 2] == ' '))
            {
                // `<r>` — new rich-text run; reset formatting.
                const r_gt = std.mem.indexOfScalarPos(u8, body, next_lt + 2, '>') orelse break;
                saw_r = true;
                pending_flags = .{ .text = "" };
                t = r_gt + 1;
            } else if (c1 == 'r' and next_lt + 3 < body.len and
                body[next_lt + 2] == 'P' and body[next_lt + 3] == 'r')
            {
                // `<rPr>...</rPr>` — parse formatting for the current run.
                const rpr_close = std.mem.indexOfPos(u8, body, next_lt, "</rPr>") orelse break;
                pending_flags = try parseRprFlags(book, arena, body[next_lt .. rpr_close + "</rPr>".len]);
                t = rpr_close + "</rPr>".len;
            } else {
                // Skip any other tag (`</r>`, `<text>`, etc.) — advance
                // past the `>` so the outer loop makes progress.
                const skip_gt = std.mem.indexOfScalarPos(u8, body, next_lt + 1, '>') orelse break;
                t = skip_gt + 1;
            }
        }

        const runs_slice: ?[]RichRun = if (saw_r and runs_list.items.len > 0)
            try runs_list.toOwnedSlice(arena)
        else blk: {
            runs_list.deinit(arena);
            break :blk null;
        };

        try entries.append(book.allocator, .{
            .top_left = cell_ref,
            .author = author_str,
            .text = try text_buf.toOwnedSlice(arena),
            .runs = runs_slice,
        });
    }

    if (entries.items.len == 0) {
        entries.deinit(book.allocator);
        return;
    }
    const slice = try entries.toOwnedSlice(book.allocator);
    errdefer book.allocator.free(slice);
    try book.comments_by_sheet.put(book.allocator, sheet_path, slice);
}

/// Arena-allocate a decoded copy of `raw`. If `raw` has no XML
/// entities, dupe as-is; otherwise decode on the way in.
fn arenaDupeDecoded(arena: Allocator, raw: []const u8) ![]const u8 {
    if (std.mem.indexOfScalar(u8, raw, '&') == null) return try arena.dupe(u8, raw);
    var buf: std.ArrayListUnmanaged(u8) = try .initCapacity(arena, raw.len);
    try appendDecoded(arena, &buf, raw);
    return try buf.toOwnedSlice(arena);
}

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

/// Scan a slice of `<rPr>...</rPr>` content for the font properties
/// this reader surfaces: bold, italic, rgb color, size, and font
/// name. `<b/>` / `<i/>` are self-closing in every OOXML generator
/// I've checked; explicit `val="false"` is rare but honoured.
/// `font_name` is duped into `arena` when present; callers get an
/// empty string otherwise. Theme colors are deliberately not
/// resolved — only explicit `rgb="AARRGGBB"` values set
/// `color_argb`.
fn parseRprFlags(book: *const Book, arena: Allocator, rpr: []const u8) !RichRun {
    var run: RichRun = .{ .text = "" };

    // Match `<b/>`, `<b ...>`, `<b val="1"/>`, etc. Skip `<b val="0"/>`.
    if (std.mem.indexOf(u8, rpr, "<b")) |bp| {
        if (bp + 2 < rpr.len) {
            const next = rpr[bp + 2];
            if (next == '/' or next == '>' or next == ' ') {
                run.bold = !hasFalseVal(rpr[bp..]);
            }
        }
    }
    if (std.mem.indexOf(u8, rpr, "<i")) |ip| {
        if (ip + 2 < rpr.len) {
            const next = rpr[ip + 2];
            if (next == '/' or next == '>' or next == ' ') {
                run.italic = !hasFalseVal(rpr[ip..]);
            }
        }
    }

    // `<color rgb="AARRGGBB"/>` parses the direct hex form.
    // `<color theme="N"/>` falls back to `book.theme_colors` — iter52
    // made theme-color resolution routine. Either attribute populates
    // `run.color_argb` when resolvable.
    if (std.mem.indexOf(u8, rpr, "<color")) |cp| {
        const gt = std.mem.indexOfScalarPos(u8, rpr, cp, '>') orelse 0;
        if (gt > cp) {
            const tag = rpr[cp..gt];
            run.color_argb = parseColorAttr(book, tag);
        }
    }

    // `<sz val="11"/>` or `<sz val="10.5"/>`.
    if (std.mem.indexOf(u8, rpr, "<sz")) |sp| {
        const gt = std.mem.indexOfScalarPos(u8, rpr, sp, '>') orelse 0;
        if (gt > sp) {
            const tag = rpr[sp..gt];
            const val_key = "val=\"";
            if (std.mem.indexOf(u8, tag, val_key)) |vp| {
                const vs = vp + val_key.len;
                if (std.mem.indexOfScalarPos(u8, tag, vs, '"')) |ve| {
                    run.size = std.fmt.parseFloat(f32, tag[vs..ve]) catch null;
                }
            }
        }
    }

    // `<rFont val="Calibri"/>` (in rich-text) or `<name val="…"/>`
    // (inside `<font>` elements in styles.xml) — same semantics, dupe
    // into arena. Try the rich-text form first; if absent, fall back
    // to the styles-form. The two never co-occur.
    const name_marker = if (std.mem.indexOf(u8, rpr, "<rFont")) |p| p else blk: {
        if (std.mem.indexOf(u8, rpr, "<name")) |np| {
            // Guard against a `<name>` that isn't a self-closing attribute tag
            // (shouldn't happen in valid OOXML, but be defensive).
            if (np + 5 < rpr.len and (rpr[np + 5] == ' ' or rpr[np + 5] == '/'))
                break :blk np;
        }
        break :blk null;
    };
    if (name_marker) |fp| {
        const gt = std.mem.indexOfScalarPos(u8, rpr, fp, '>') orelse 0;
        if (gt > fp) {
            const tag = rpr[fp..gt];
            const val_key = "val=\"";
            if (std.mem.indexOf(u8, tag, val_key)) |vp| {
                const vs = vp + val_key.len;
                if (std.mem.indexOfScalarPos(u8, tag, vs, '"')) |ve| {
                    run.font_name = try arena.dupe(u8, tag[vs..ve]);
                }
            }
        }
    }

    return run;
}

/// Returns true if the tag body (up to the next `>`) contains
/// `val="0"` or `val="false"`. OOXML treats missing val as true.
fn hasFalseVal(tag: []const u8) bool {
    const gt = std.mem.indexOfScalar(u8, tag, '>') orelse return false;
    const body = tag[0..gt];
    return std.mem.indexOf(u8, body, "val=\"0\"") != null or
        std.mem.indexOf(u8, body, "val=\"false\"") != null;
}

/// Walk `xl/styles.xml` and populate:
///   - `book.cell_xf_numfmt_ids` — one u32 numFmtId per `<xf>` under
///     `<cellXfs>` (callers look up by the `s="N"` index on cells).
///   - `book.custom_num_fmts` — map of numFmtId → format code for
///     every `<numFmt numFmtId="…" formatCode="…"/>` (built-in ids
///     don't appear here; they're resolved via `builtinNumberFormat`).
/// Values borrow from `styles_xml` so the file stays mapped for the
/// Book's lifetime. Malformed or missing sections degrade silently —
/// an xlsx with no cellXfs just returns null from `numberFormat`.
/// Walk `xl/theme/theme1.xml` and populate `book.theme_colors` with
/// the 12 palette entries OOXML's `<color theme="N"/>` indexes into.
///
/// OOXML layout: the `<a:clrScheme>` element holds the 12 palette
/// entries in schema order:
///   dk1, lt1, dk2, lt2, accent1..accent6, hlink, folHlink.
/// But `<color theme="N"/>` uses a different index order (lt1 and
/// dk1 are swapped):
///   0 = lt1, 1 = dk1, 2 = lt2, 3 = dk2, 4..9 = accent1..6,
///   10 = hlink, 11 = folHlink.
/// Each entry is either `<a:srgbClr val="HEXHEX"/>` or
/// `<a:sysClr val="…" lastClr="HEXHEX"/>` — we read the lastClr
/// fallback for sysClr since there's no way to resolve "windowText"
/// without the OS.
fn parseTheme(book: *Book, xml: []const u8) !void {
    const scheme_open = std.mem.indexOf(u8, xml, "<a:clrScheme") orelse return;
    const scheme_close = std.mem.indexOfPos(u8, xml, scheme_open, "</a:clrScheme>") orelse return;
    const scheme = xml[scheme_open..scheme_close];

    // Read each child element of clrScheme in schema order. Expected
    // tags: <a:dk1>, <a:lt1>, <a:dk2>, <a:lt2>, <a:accent1> .. 6,
    // <a:hlink>, <a:folHlink>. Each wraps either srgbClr or sysClr.
    const tags = [_][]const u8{
        "<a:dk1>",     "<a:lt1>",
        "<a:dk2>",     "<a:lt2>",
        "<a:accent1>", "<a:accent2>",
        "<a:accent3>", "<a:accent4>",
        "<a:accent5>", "<a:accent6>",
        "<a:hlink>",   "<a:folHlink>",
    };
    // `schema_colors[i]` follows the XML order above.
    var schema_colors: [12]u32 = undefined;
    var got: [12]bool = .{false} ** 12;
    for (tags, 0..) |tag, i| {
        const open = std.mem.indexOf(u8, scheme, tag) orelse continue;
        const start = open + tag.len;
        // Close tag mirrors the open: `<a:dk1>` → `</a:dk1>`.
        var close_buf: [24]u8 = undefined;
        const close_tag = std.fmt.bufPrint(&close_buf, "</{s}", .{tag[1..]}) catch continue;
        const close = std.mem.indexOfPos(u8, scheme, start, close_tag) orelse continue;
        const body = scheme[start..close];
        schema_colors[i] = extractThemeColorValue(body) orelse continue;
        got[i] = true;
    }

    // Remap schema order → `<color theme="N"/>` index order.
    // ECMA-376: both dk/lt pairs are flipped — theme index order
    // is lt1, dk1, lt2, dk2, accent1..6, hlink, folHlink, whereas
    // the XML schema orders them dk1, lt1, dk2, lt2, ...
    const swap_pairs = [_][2]usize{
        .{ 0, 1 }, .{ 1, 0 }, // lt1 ← schema[1], dk1 ← schema[0]
        .{ 2, 3 },   .{ 3, 2 }, // lt2 ← schema[3], dk2 ← schema[2]
        .{ 4, 4 },   .{ 5, 5 },
        .{ 6, 6 },   .{ 7, 7 },
        .{ 8, 8 },   .{ 9, 9 },
        .{ 10, 10 }, .{ 11, 11 },
    };
    var out: [12]u32 = undefined;
    var out_len: usize = 0;
    for (swap_pairs) |pair| {
        const theme_idx = pair[0];
        const schema_idx = pair[1];
        if (!got[schema_idx]) break;
        out[theme_idx] = schema_colors[schema_idx];
        out_len = @max(out_len, theme_idx + 1);
    }

    if (out_len > 0) {
        book.theme_colors = try book.allocator.dupe(u32, out[0..out_len]);
    }
}

/// Pull an ARGB from a theme color element body — one of:
///   `<a:srgbClr val="HEXHEX"/>`    → parse val
///   `<a:sysClr val="…" lastClr="HEXHEX"/>` → parse lastClr
/// Theme hex is RGB only (6 digits); upcast to ARGB with FF alpha.
fn extractThemeColorValue(body: []const u8) ?u32 {
    const Tag = struct { open: []const u8, attr: []const u8 };
    const candidates = [_]Tag{
        .{ .open = "<a:srgbClr", .attr = "val=\"" },
        .{ .open = "<a:sysClr", .attr = "lastClr=\"" },
    };
    for (candidates) |c| {
        if (std.mem.indexOf(u8, body, c.open)) |pos| {
            const gt = std.mem.indexOfScalarPos(u8, body, pos, '>') orelse continue;
            const tag = body[pos..gt];
            if (std.mem.indexOf(u8, tag, c.attr)) |ap| {
                const s = ap + c.attr.len;
                const e = std.mem.indexOfScalarPos(u8, tag, s, '"') orelse continue;
                const hex = tag[s..e];
                if (hex.len != 6) continue;
                const rgb = std.fmt.parseInt(u32, hex, 16) catch continue;
                return 0xFF00_0000 | rgb; // promote to ARGB with opaque alpha
            }
        }
    }
    return null;
}

/// Look up theme color index `N` in `book.theme_colors`, returning
/// null when out of range or when the workbook has no theme table.
fn resolveThemeColor(book: *const Book, theme_idx: u32) ?u32 {
    if (theme_idx >= book.theme_colors.len) return null;
    return book.theme_colors[theme_idx];
}

/// Parse an ARGB from an OOXML color tag attribute list. Supports:
///   rgb="AARRGGBB"   → direct hex
///   theme="N"        → lookup via book.theme_colors
///   indexed="N"      → legacy indexed palette (not resolved — returns null)
/// Tint attribute is ignored (would require HSL math we don't ship).
/// Returns null when none of the above parse successfully.
fn parseColorAttr(book: *const Book, tag: []const u8) ?u32 {
    const rgb_key = "rgb=\"";
    if (std.mem.indexOf(u8, tag, rgb_key)) |rp| {
        const rs = rp + rgb_key.len;
        if (std.mem.indexOfScalarPos(u8, tag, rs, '"')) |re| {
            if (std.fmt.parseInt(u32, tag[rs..re], 16)) |v| return v else |_| {}
        }
    }
    const theme_key = "theme=\"";
    if (std.mem.indexOf(u8, tag, theme_key)) |tp| {
        const ts = tp + theme_key.len;
        if (std.mem.indexOfScalarPos(u8, tag, ts, '"')) |te| {
            if (std.fmt.parseInt(u32, tag[ts..te], 10)) |idx|
                return resolveThemeColor(book, idx)
            else |_| {}
        }
    }
    return null;
}

fn parseStyles(book: *Book, xml: []const u8) !void {
    // fonts — optional. Shape: `<fonts count="N"><font>...</font>...</fonts>`.
    // Each <font> has the same child shape as <rPr> inside <si>, so we
    // reuse parseRprFlags + project its text='' variant into a Font.
    if (std.mem.indexOf(u8, xml, "<fonts")) |fp| {
        const fp_end = std.mem.indexOfPos(u8, xml, fp, "</fonts>") orelse xml.len;
        const block = xml[fp..fp_end];
        var fonts_list: std.ArrayListUnmanaged(Font) = .{};
        errdefer fonts_list.deinit(book.allocator);
        var i: usize = 0;
        while (std.mem.indexOfPos(u8, block, i, "<font")) |font_pos| {
            // Skip <fonts> self-tag (already consumed via index above).
            if (font_pos + 5 < block.len) {
                const after = block[font_pos + 5];
                if (after != '>' and after != ' ') {
                    i = font_pos + 5;
                    continue;
                }
            }
            // Self-closing `<font/>` → default empty.
            const gt = std.mem.indexOfScalarPos(u8, block, font_pos, '>') orelse break;
            if (gt > 0 and block[gt - 1] == '/') {
                try fonts_list.append(book.allocator, .{});
                i = gt + 1;
                continue;
            }
            const font_close = std.mem.indexOfPos(u8, block, gt, "</font>") orelse break;
            const body = block[font_pos .. font_close + "</font>".len];
            const rr = try parseRprFlags(book, book.sst_arena.allocator(), body);
            try fonts_list.append(book.allocator, .{
                .bold = rr.bold,
                .italic = rr.italic,
                .color_argb = rr.color_argb,
                .size = rr.size,
                .name = rr.font_name,
            });
            i = font_close + "</font>".len;
        }
        book.fonts = try fonts_list.toOwnedSlice(book.allocator);
    }

    // fills — optional. Shape:
    //   <fills count="N">
    //     <fill><patternFill patternType="none"/></fill>
    //     <fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/><bgColor indexed="64"/></patternFill></fill>
    //   </fills>
    if (std.mem.indexOf(u8, xml, "<fills")) |fp| {
        const fp_end = std.mem.indexOfPos(u8, xml, fp, "</fills>") orelse xml.len;
        const block = xml[fp..fp_end];
        var fills_list: std.ArrayListUnmanaged(Fill) = .{};
        errdefer fills_list.deinit(book.allocator);
        var i: usize = 0;
        while (std.mem.indexOfPos(u8, block, i, "<fill")) |fill_pos| {
            if (fill_pos + 5 < block.len) {
                const after = block[fill_pos + 5];
                if (after != '>' and after != ' ' and after != '/') {
                    // `<fills ...>` — outer wrapper, skip the tag.
                    i = fill_pos + 5;
                    continue;
                }
            }
            const gt = std.mem.indexOfScalarPos(u8, block, fill_pos, '>') orelse break;
            if (gt > 0 and block[gt - 1] == '/') {
                try fills_list.append(book.allocator, .{});
                i = gt + 1;
                continue;
            }
            const fill_close = std.mem.indexOfPos(u8, block, gt, "</fill>") orelse break;
            const body = block[gt + 1 .. fill_close];
            try fills_list.append(book.allocator, parseFillBody(book, body));
            i = fill_close + "</fill>".len;
        }
        book.fills = try fills_list.toOwnedSlice(book.allocator);
    }

    // borders — optional. Shape:
    //   <borders count="N">
    //     <border>
    //       <left style="thin"><color rgb="FF000000"/></left>
    //       <right/>
    //       <top style="thin"><color rgb="FF000000"/></top>
    //       <bottom/>
    //       <diagonal/>
    //     </border>
    //   </borders>
    if (std.mem.indexOf(u8, xml, "<borders")) |bp| {
        const bp_end = std.mem.indexOfPos(u8, xml, bp, "</borders>") orelse xml.len;
        const block = xml[bp..bp_end];
        var borders_list: std.ArrayListUnmanaged(Border) = .{};
        errdefer borders_list.deinit(book.allocator);
        var i: usize = 0;
        while (std.mem.indexOfPos(u8, block, i, "<border")) |border_pos| {
            if (border_pos + 7 < block.len) {
                const after = block[border_pos + 7];
                // Skip the outer `<borders …>` wrapper tag.
                if (after != '>' and after != ' ' and after != '/') {
                    i = border_pos + 7;
                    continue;
                }
            }
            const gt = std.mem.indexOfScalarPos(u8, block, border_pos, '>') orelse break;
            if (gt > 0 and block[gt - 1] == '/') {
                try borders_list.append(book.allocator, .{});
                i = gt + 1;
                continue;
            }
            const border_close = std.mem.indexOfPos(u8, block, gt, "</border>") orelse break;
            const body = block[gt + 1 .. border_close];
            try borders_list.append(book.allocator, parseBorderBody(book, body));
            i = border_close + "</border>".len;
        }
        book.borders = try borders_list.toOwnedSlice(book.allocator);
    }

    // numFmts — optional. Shape: `<numFmts count="…"><numFmt numFmtId="164" formatCode="…"/>…</numFmts>`.
    if (std.mem.indexOf(u8, xml, "<numFmts")) |nfs_pos| {
        const nfs_end = std.mem.indexOfPos(u8, xml, nfs_pos, "</numFmts>") orelse xml.len;
        const block = xml[nfs_pos..nfs_end];
        var i: usize = 0;
        while (std.mem.indexOfPos(u8, block, i, "<numFmt ")) |nf| {
            const gt = std.mem.indexOfScalarPos(u8, block, nf, '>') orelse break;
            const attrs = block[nf..gt];
            i = gt + 1;
            const id_key = "numFmtId=\"";
            const code_key = "formatCode=\"";
            const id_pos = std.mem.indexOf(u8, attrs, id_key) orelse continue;
            const id_start = id_pos + id_key.len;
            const id_end = std.mem.indexOfScalarPos(u8, attrs, id_start, '"') orelse continue;
            const id = std.fmt.parseInt(u32, attrs[id_start..id_end], 10) catch continue;
            const code_pos = std.mem.indexOf(u8, attrs, code_key) orelse continue;
            const code_start = code_pos + code_key.len;
            const code_end = std.mem.indexOfScalarPos(u8, attrs, code_start, '"') orelse continue;
            try book.custom_num_fmts.put(book.allocator, id, attrs[code_start..code_end]);
        }
    }

    // cellXfs — the slot callers index into via `s="N"`. Shape:
    // `<cellXfs count="…"><xf numFmtId="0" fontId="0" fillId="0" …/>…</cellXfs>`.
    const xfs_pos = std.mem.indexOf(u8, xml, "<cellXfs") orelse return;
    const xfs_end = std.mem.indexOfPos(u8, xml, xfs_pos, "</cellXfs>") orelse return;
    const xfs_block = xml[xfs_pos..xfs_end];

    var ids: std.ArrayListUnmanaged(u32) = .{};
    errdefer ids.deinit(book.allocator);
    var font_ids: std.ArrayListUnmanaged(u32) = .{};
    errdefer font_ids.deinit(book.allocator);
    var fill_ids: std.ArrayListUnmanaged(u32) = .{};
    errdefer fill_ids.deinit(book.allocator);
    var border_ids: std.ArrayListUnmanaged(u32) = .{};
    errdefer border_ids.deinit(book.allocator);
    var i: usize = 0;
    while (std.mem.indexOfPos(u8, xfs_block, i, "<xf")) |xp| {
        // Guard against longer tags that share the `<xf` prefix
        // (OOXML doesn't define any, but future-proof the scan the
        // same way `<font` / `<fill` / `<border` above do).
        if (xp + 3 >= xfs_block.len) break;
        const after = xfs_block[xp + 3];
        if (after != ' ' and after != '>' and after != '/') {
            i = xp + 3;
            continue;
        }
        const gt = std.mem.indexOfScalarPos(u8, xfs_block, xp, '>') orelse break;
        const attrs = xfs_block[xp..gt];
        i = gt + 1;
        try ids.append(book.allocator, parseXfAttrU32(attrs, "numFmtId=\"", 0));
        try font_ids.append(book.allocator, parseXfAttrU32(attrs, "fontId=\"", 0));
        try fill_ids.append(book.allocator, parseXfAttrU32(attrs, "fillId=\"", 0));
        try border_ids.append(book.allocator, parseXfAttrU32(attrs, "borderId=\"", 0));
    }
    book.cell_xf_numfmt_ids = try ids.toOwnedSlice(book.allocator);
    book.cell_xf_font_ids = try font_ids.toOwnedSlice(book.allocator);
    book.cell_xf_fill_ids = try fill_ids.toOwnedSlice(book.allocator);
    book.cell_xf_border_ids = try border_ids.toOwnedSlice(book.allocator);
}

/// Parse a single `<border>` body (between `<border>` and `</border>`)
/// into a `Border`. Handles self-closing side tags (no border) and
/// regular `<left style="thin"><color rgb="…"/></left>` shapes.
fn parseBorderBody(book: *const Book, body: []const u8) Border {
    return .{
        .left = parseBorderSide(book, body, "left"),
        .right = parseBorderSide(book, body, "right"),
        .top = parseBorderSide(book, body, "top"),
        .bottom = parseBorderSide(book, body, "bottom"),
        .diagonal = parseBorderSide(book, body, "diagonal"),
    };
}

/// Look up one side element (`<left>`, `<right>`, `<top>`, `<bottom>`,
/// `<diagonal>`) in the border body. Returns default `BorderSide`
/// when the side is self-closing or absent. Parses `style="…"`
/// attribute and `<color rgb="…"/>` child. Slices borrow from the
/// input body (which itself borrows from `styles_xml`).
fn parseBorderSide(book: *const Book, body: []const u8, tag: []const u8) BorderSide {
    // All callers pass the literal OOXML side names — longest is
    // "diagonal" (8 bytes). Future tag additions that overflow these
    // stack buffers would otherwise `bufPrint → error.NoSpaceLeft`
    // and silently return an empty side. Bail early at build-time
    // instead.
    std.debug.assert(tag.len + 3 <= 16); // "<tag" + nul slack
    std.debug.assert(tag.len + 4 <= 24); // "</tag>" + nul slack
    var open_buf: [16]u8 = undefined;
    const open = std.fmt.bufPrint(&open_buf, "<{s}", .{tag}) catch return .{};
    const pos = std.mem.indexOf(u8, body, open) orelse return .{};
    const after_tag = pos + open.len;
    if (after_tag >= body.len) return .{};
    // The next char must be space, `>`, or `/` (self-closing); anything
    // else means we matched a longer tag like `<topics>` that happens
    // to share a prefix — bail out.
    const next = body[after_tag];
    if (next != ' ' and next != '>' and next != '/') return .{};

    const gt = std.mem.indexOfScalarPos(u8, body, pos, '>') orelse return .{};
    const tag_body = body[pos..gt];
    var side: BorderSide = .{};
    const style_key = "style=\"";
    if (std.mem.indexOf(u8, tag_body, style_key)) |sp| {
        const s = sp + style_key.len;
        if (std.mem.indexOfScalarPos(u8, tag_body, s, '"')) |e| {
            side.style = tag_body[s..e];
        }
    }
    // Self-closing → no color child.
    if (gt > 0 and body[gt - 1] == '/') return side;

    // Otherwise look for a <color rgb="…"/> before the matching close tag.
    var close_buf: [24]u8 = undefined;
    const close_s = std.fmt.bufPrint(&close_buf, "</{s}>", .{tag}) catch return side;
    const close_pos = std.mem.indexOfPos(u8, body, gt, close_s) orelse return side;
    const inner = body[gt + 1 .. close_pos];
    if (std.mem.indexOf(u8, inner, "<color")) |cp| {
        const cgt = std.mem.indexOfScalarPos(u8, inner, cp, '>') orelse return side;
        side.color_argb = parseColorAttr(book, inner[cp..cgt]);
    }
    return side;
}

/// Parse a `<fill>` body slice (everything between `<fill>` and
/// `</fill>`). Handles the common shape `<patternFill patternType="…">
/// <fgColor rgb="…"/><bgColor rgb="…"/></patternFill>` plus the
/// no-op `<patternFill patternType="none"/>` variant. Unresolved
/// theme / indexed colors leave `*_color_argb` null.
fn parseFillBody(book: *const Book, body: []const u8) Fill {
    var out: Fill = .{};
    // patternType on <patternFill>.
    const pt_key = "patternType=\"";
    if (std.mem.indexOf(u8, body, pt_key)) |pp| {
        const s = pp + pt_key.len;
        if (std.mem.indexOfScalarPos(u8, body, s, '"')) |e| {
            out.pattern = body[s..e];
        }
    }
    // <fgColor rgb="AARRGGBB"/> or <fgColor theme="N"/>
    if (std.mem.indexOf(u8, body, "<fgColor")) |fp| {
        const gt = std.mem.indexOfScalarPos(u8, body, fp, '>') orelse 0;
        if (gt > fp) {
            out.fg_color_argb = parseColorAttr(book, body[fp..gt]);
        }
    }
    // <bgColor rgb="AARRGGBB"/> or <bgColor theme="N"/>
    if (std.mem.indexOf(u8, body, "<bgColor")) |bp| {
        const gt = std.mem.indexOfScalarPos(u8, body, bp, '>') orelse 0;
        if (gt > bp) {
            out.bg_color_argb = parseColorAttr(book, body[bp..gt]);
        }
    }
    return out;
}

/// Parse a `key="N"` u32 attribute from an `<xf>` attrs blob, defaulting
/// when the key is absent or malformed. OOXML's cellXfs entries are
/// generated programmatically and have predictable shape; we don't need
/// a full XML attr parser for a handful of integer slots.
fn parseXfAttrU32(attrs: []const u8, key: []const u8, default: u32) u32 {
    const kp = std.mem.indexOf(u8, attrs, key) orelse return default;
    const s = kp + key.len;
    const e = std.mem.indexOfScalarPos(u8, attrs, s, '"') orelse return default;
    return std.fmt.parseInt(u32, attrs[s..e], 10) catch default;
}

/// Resolve a built-in number-format id to its OOXML pattern. Returns
/// null for ids that are neither built-in nor in the custom table —
/// callers should treat that as "General" themselves. IDs 0-49 are
/// the classical built-in set; 14-22 and 45-47 are the date/time
/// relevant ones.
fn builtinNumberFormat(id: u32) ?[]const u8 {
    return switch (id) {
        0 => "General",
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        14 => "m/d/yyyy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yyyy h:mm",
        37 => "#,##0 ;(#,##0)",
        38 => "#,##0 ;[Red](#,##0)",
        39 => "#,##0.00;(#,##0.00)",
        40 => "#,##0.00;[Red](#,##0.00)",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mm:ss.0",
        48 => "##0.0E+0",
        49 => "@",
        else => null,
    };
}

/// Heuristic: does the format code describe a date / time / datetime?
/// Covers the 9 built-in date IDs and custom codes that contain an
/// unquoted `y`, `m`, `d`, `h`, or `s` token outside brackets. Skips
/// quoted literals (`"dd"` shouldn't trigger) and the `[Red]` / `[h]`
/// color / duration modifiers.
fn isDateFormatCode(code: []const u8) bool {
    var in_quote = false;
    var in_bracket = false;
    for (code) |c| {
        switch (c) {
            '"' => in_quote = !in_quote,
            '[' => in_bracket = true,
            ']' => in_bracket = false,
            'y', 'Y', 'd', 'D', 'h', 'H', 's', 'S' => {
                if (!in_quote and !in_bracket) return true;
            },
            'm', 'M' => {
                // `m` is ambiguous (month vs minute), but any presence
                // outside quotes/brackets means the code has a
                // date-or-time component — both are date-y enough.
                if (!in_quote and !in_bracket) return true;
            },
            else => {},
        }
    }
    return false;
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
                        .color_argb = pending_flags.color_argb,
                        .size = pending_flags.size,
                        .font_name = pending_flags.font_name,
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
                pending_flags = try parseRprFlags(book, arena_alloc, xml[next_lt .. rpr_close + "</rPr>".len]);
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

test "toExcelSerial: inverse of fromExcelSerial on round-trippable range" {
    // Round-trip every known reference date through both helpers.
    const cases = [_]struct {
        y: u16,
        m: u8,
        d: u8,
        want: f64,
    }{
        .{ .y = 1900, .m = 3, .d = 1, .want = 61.0 }, // earliest supported
        .{ .y = 1970, .m = 1, .d = 1, .want = 25569.0 }, // unix epoch
        .{ .y = 2000, .m = 1, .d = 1, .want = 36526.0 },
        .{ .y = 2023, .m = 1, .d = 1, .want = 44927.0 },
        .{ .y = 2024, .m = 2, .d = 29, .want = 45351.0 }, // leap day
        .{ .y = 9999, .m = 12, .d = 31, .want = 2958465.0 }, // upper bound
    };
    for (cases) |c| {
        const dt = DateTime{ .year = c.y, .month = c.m, .day = c.d, .hour = 0, .minute = 0, .second = 0 };
        const got = toExcelSerial(dt) orelse {
            std.debug.print("toExcelSerial returned null for {d}-{d}-{d}\n", .{ c.y, c.m, c.d });
            return error.TestUnexpectedResult;
        };
        try std.testing.expectEqual(c.want, got);
        // Round-trip: serial → DateTime → same fields.
        const back = fromExcelSerial(got) orelse return error.TestUnexpectedResult;
        try std.testing.expectEqualDeep(dt, back);
    }

    // Time-of-day: 2023-06-15 12:34:56 → serial + fractional part.
    const noon_ish = DateTime{ .year = 2023, .month = 6, .day = 15, .hour = 12, .minute = 34, .second = 56 };
    const s = toExcelSerial(noon_ish) orelse return error.TestUnexpectedResult;
    const back = fromExcelSerial(s) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualDeep(noon_ish, back);

    // Rejection paths.
    try std.testing.expectEqual(@as(?f64, null), toExcelSerial(.{ .year = 1899, .month = 12, .day = 31, .hour = 0, .minute = 0, .second = 0 }));
    try std.testing.expectEqual(@as(?f64, null), toExcelSerial(.{ .year = 1900, .month = 2, .day = 28, .hour = 0, .minute = 0, .second = 0 })); // pre-leap-bug exclusion
    try std.testing.expectEqual(@as(?f64, null), toExcelSerial(.{ .year = 2023, .month = 2, .day = 29, .hour = 0, .minute = 0, .second = 0 })); // non-leap Feb 29
    try std.testing.expectEqual(@as(?f64, null), toExcelSerial(.{ .year = 2023, .month = 13, .day = 1, .hour = 0, .minute = 0, .second = 0 }));
    try std.testing.expectEqual(@as(?f64, null), toExcelSerial(.{ .year = 2023, .month = 4, .day = 31, .hour = 0, .minute = 0, .second = 0 })); // April has 30 days
    try std.testing.expectEqual(@as(?f64, null), toExcelSerial(.{ .year = 2023, .month = 1, .day = 1, .hour = 24, .minute = 0, .second = 0 })); // hour out of range
}

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

test "writer.addComment: emits comments that round-trip through Book.comments" {
    const tmp_path = "/tmp/zlsx_writer_comments_roundtrip.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("S");
        try sheet.addComment("B2", "Alice", "review this");
        try sheet.addComment("C3", "Bob & Co", "R&D notes");
        try sheet.addComment("D4", "Alice", "follow-up"); // same author reused
        try sheet.writeRow(&.{.{ .string = "hdr" }});
        try w.save(tmp_path);

        // Rejection paths — single-cell refs only, non-empty.
        try std.testing.expectError(error.InvalidCommentRef, sheet.addComment("", "a", "b"));
        try std.testing.expectError(error.InvalidCommentRef, sheet.addComment("A1:B2", "a", "b"));
    }

    var book = try Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    const cs = book.comments(book.sheets[0]);
    try std.testing.expectEqual(@as(usize, 3), cs.len);

    try std.testing.expectEqualDeep(CellRef{ .col = 1, .row = 2 }, cs[0].top_left);
    try std.testing.expectEqualStrings("Alice", cs[0].author);
    try std.testing.expectEqualStrings("review this", cs[0].text);

    try std.testing.expectEqualDeep(CellRef{ .col = 2, .row = 3 }, cs[1].top_left);
    try std.testing.expectEqualStrings("Bob & Co", cs[1].author); // entity-decoded
    try std.testing.expectEqualStrings("R&D notes", cs[1].text);

    try std.testing.expectEqualDeep(CellRef{ .col = 3, .row = 4 }, cs[2].top_left);
    try std.testing.expectEqualStrings("Alice", cs[2].author); // author table dedup
    try std.testing.expectEqualStrings("follow-up", cs[2].text);
}

test "writer.addComment: XML-special chars in author + text round-trip via entity-decoding" {
    const tmp_path = "/tmp/zlsx_comments_xml_special.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("S");
        try sheet.addComment("A1", "R&D <Lead>", "review <this> & \"that\" quickly");
        try sheet.writeRow(&.{.{ .string = "x" }});
        try w.save(tmp_path);
    }
    var book = try Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    const cs = book.comments(book.sheets[0]);
    try std.testing.expectEqual(@as(usize, 1), cs.len);
    // Every XML-special survives escape-on-write + decode-on-read.
    try std.testing.expectEqualStrings("R&D <Lead>", cs[0].author);
    try std.testing.expectEqualStrings("review <this> & \"that\" quickly", cs[0].text);
}

test "writer.addComment: multi-sheet with comments keeps per-sheet rels independent" {
    const tmp_path = "/tmp/zlsx_comments_multi_sheet.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s1 = try w.addSheet("S1");
        try s1.addComment("A1", "Alice", "on S1");
        try s1.writeRow(&.{.{ .string = "x" }});
        // S2 has NO comments — its rels file must stay absent and the
        // sheet XML must not carry `<legacyDrawing>`.
        var s2 = try w.addSheet("S2");
        try s2.writeRow(&.{.{ .string = "y" }});
        // S3 has its own comments — both its rels and its commentsN.xml
        // must be emitted with the correct N suffix.
        var s3 = try w.addSheet("S3");
        try s3.addComment("B2", "Bob", "on S3");
        try s3.addComment("C3", "Carol", "also on S3");
        try s3.writeRow(&.{.{ .string = "z" }});
        try w.save(tmp_path);
    }
    var book = try Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    try std.testing.expectEqual(@as(usize, 1), book.comments(book.sheets[0]).len);
    try std.testing.expectEqual(@as(usize, 0), book.comments(book.sheets[1]).len);
    try std.testing.expectEqual(@as(usize, 2), book.comments(book.sheets[2]).len);

    // Cross-check: sheet 2's "on S1" comment doesn't leak into S3.
    const s3 = book.comments(book.sheets[2]);
    try std.testing.expectEqualStrings("Bob", s3[0].author);
    try std.testing.expectEqualStrings("on S3", s3[0].text);
    try std.testing.expectEqualStrings("Carol", s3[1].author);
}

test "writer.addComment: 50 comments in one sheet stress-test authors + rels" {
    const tmp_path = "/tmp/zlsx_comments_stress.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("S");
        var ref_buf: [8]u8 = undefined;
        for (0..50) |i| {
            // Refs march down column A: A1, A2, …, A50.
            const ref = try std.fmt.bufPrint(&ref_buf, "A{d}", .{i + 1});
            // Two alternating authors to exercise the dedup table.
            const author = if (i % 2 == 0) "Alice" else "Bob";
            var text_buf: [32]u8 = undefined;
            const text = try std.fmt.bufPrint(&text_buf, "note #{d}", .{i});
            try sheet.addComment(ref, author, text);
        }
        try sheet.writeRow(&.{.{ .string = "hdr" }});
        try w.save(tmp_path);
    }
    var book = try Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();
    const cs = book.comments(book.sheets[0]);
    try std.testing.expectEqual(@as(usize, 50), cs.len);
    // Spot-check the first, middle, and last — all must round-trip
    // through the same authors-table indirection without drift.
    try std.testing.expectEqualStrings("Alice", cs[0].author);
    try std.testing.expectEqualStrings("note #0", cs[0].text);
    try std.testing.expectEqualStrings("Bob", cs[25].author);
    try std.testing.expectEqualStrings("note #25", cs[25].text);
    try std.testing.expectEqualStrings("Bob", cs[49].author);
    try std.testing.expectEqualStrings("note #49", cs[49].text);
}

test "writer.writeRichRow: emits rich-text SST entries readable by Book.richRuns" {
    const tmp_path = "/tmp/zlsx_writer_rich_roundtrip.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("S");
        // Row mixes plain + rich + integer.
        try sheet.writeRichRow(&.{
            .{ .string = "label" },
            .{ .rich = &.{
                .{ .text = "hello ", .bold = true },
                .{ .text = "world", .italic = true, .color_argb = 0xFFFF0000 },
            } },
            .{ .integer = 42 },
        });
        try w.save(tmp_path);
    }

    var book = try Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    // SST ordering: plain "label" first, rich entry second.
    try std.testing.expectEqualStrings("label", book.shared_strings[0]);
    try std.testing.expectEqualStrings("hello world", book.shared_strings[1]);

    // Plain entry has no runs; rich entry has two.
    try std.testing.expectEqual(@as(?[]const RichRun, null), book.richRuns(0));
    const runs = book.richRuns(1) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(usize, 2), runs.len);
    try std.testing.expectEqualStrings("hello ", runs[0].text);
    try std.testing.expectEqual(true, runs[0].bold);
    try std.testing.expectEqual(false, runs[0].italic);
    try std.testing.expectEqualStrings("world", runs[1].text);
    try std.testing.expectEqual(true, runs[1].italic);
    try std.testing.expectEqual(@as(?u32, 0xFFFF0000), runs[1].color_argb);
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

test "Rows.parseDate: auto-convert date-styled cells through the reader" {
    // iter46 convenience: instead of chaining styleIndices +
    // isDateFormat + fromExcelSerial, callers just call parseDate.
    // Must return null for non-date-styled numbers, null for
    // string cells, and the correct DateTime for date-styled
    // numerics.
    const tmp_path = "/tmp/zlsx_rows_parse_date.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        const date_style = try w.addStyle(.{ .number_format = "yyyy-mm-dd" });
        const pct_style = try w.addStyle(.{ .number_format = "0.00%" });
        var sheet = try w.addSheet("S");
        // Row 1 — hdr. Row 2 — date col + pct col + plain integer + text.
        try sheet.writeRow(&.{.{ .string = "hdr" }});
        try sheet.writeRowStyled(
            &.{
                .{ .number = 44927 }, // 2023-01-01
                .{ .number = 0.25 }, // styled % — not a date
                .{ .integer = 42 }, // no style — not a date
                .{ .string = "txt" }, // string cell
            },
            &.{ date_style, pct_style, 0, 0 },
        );
        try w.save(tmp_path);
    }

    var book = try Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    _ = try rows.next(); // hdr
    _ = try rows.next(); // data row — populates row_cells + row_styles

    // col 0: date-styled 44927 → 2023-01-01
    const d0 = rows.parseDate(0) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(u16, 2023), d0.year);
    try std.testing.expectEqual(@as(u8, 1), d0.month);
    try std.testing.expectEqual(@as(u8, 1), d0.day);

    // col 1: percentage-styled — not a date.
    try std.testing.expectEqual(@as(?DateTime, null), rows.parseDate(1));
    // col 2: no style — not a date.
    try std.testing.expectEqual(@as(?DateTime, null), rows.parseDate(2));
    // col 3: string cell — not a date.
    try std.testing.expectEqual(@as(?DateTime, null), rows.parseDate(3));
    // col 99: out of range — null.
    try std.testing.expectEqual(@as(?DateTime, null), rows.parseDate(99));
}

test "Book.numberFormat + isDateFormat: built-in, custom, and per-cell style lookup" {
    // Write a workbook with a styled column, then read it back and
    // check that every moving part lines up: styles.xml is extracted,
    // cellXfs is parsed, per-cell `s="N"` is tracked via
    // Rows.styleIndices(), numberFormat() resolves built-ins + custom,
    // and isDateFormat() gets the heuristic right.
    const tmp_path = "/tmp/zlsx_reader_numfmt.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        const date_style = try w.addStyle(.{ .number_format = "yyyy-mm-dd" });
        const pct_style = try w.addStyle(.{ .number_format = "0.00%" });
        var sheet = try w.addSheet("S");
        try sheet.writeRow(&.{.{ .string = "hdr" }});
        try sheet.writeRowStyled(
            &.{ .{ .number = 44927 }, .{ .number = 0.25 }, .{ .integer = 42 } },
            &.{ date_style, pct_style, 0 },
        );
        try w.save(tmp_path);
    }

    var book = try Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    // Pull the second row and check styleIndices alignment.
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    _ = try rows.next(); // hdr row
    const cells = (try rows.next()).?;
    const styles = rows.styleIndices();
    try std.testing.expectEqual(@as(usize, 3), cells.len);
    try std.testing.expectEqual(@as(usize, 3), styles.len);

    // The date-style cell: number format "yyyy-mm-dd" is custom,
    // must surface via numberFormat and register as a date.
    const s0 = styles[0] orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualStrings("yyyy-mm-dd", book.numberFormat(s0).?);
    try std.testing.expectEqual(true, book.isDateFormat(s0));

    // Percentage cell: custom numFmt, not a date.
    const s1 = styles[1] orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualStrings("0.00%", book.numberFormat(s1).?);
    try std.testing.expectEqual(false, book.isDateFormat(s1));

    // Plain integer cell uses the writer's default xfId 0 (General).
    if (styles[2]) |s2| {
        try std.testing.expectEqual(false, book.isDateFormat(s2));
    }
}

test "Book.cellFont: round-trips bold / color / size / name from xl/styles.xml" {
    const tmp_path = "/tmp/zlsx_reader_cell_font.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        const bold_style = try w.addStyle(.{
            .font_bold = true,
            .font_color_argb = 0xFFFF0000,
            .font_size = 14,
            .font_name = "Courier New",
        });
        const plain_style = try w.addStyle(.{ .font_italic = true });
        var sheet = try w.addSheet("S");
        try sheet.writeRowStyled(
            &.{ .{ .string = "bold-red" }, .{ .string = "italic" }, .{ .string = "bare" } },
            &.{ bold_style, plain_style, 0 },
        );
        try w.save(tmp_path);
    }

    var book = try Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    _ = try rows.next();
    const styles = rows.styleIndices();

    const s0 = styles[0] orelse return error.TestUnexpectedResult;
    const f0 = book.cellFont(s0) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(true, f0.bold);
    try std.testing.expectEqual(false, f0.italic);
    try std.testing.expectEqual(@as(?u32, 0xFFFF0000), f0.color_argb);
    try std.testing.expectEqual(@as(?f32, 14.0), f0.size);
    try std.testing.expectEqualStrings("Courier New", f0.name);

    const s1 = styles[1] orelse return error.TestUnexpectedResult;
    const f1 = book.cellFont(s1) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(false, f1.bold);
    try std.testing.expectEqual(true, f1.italic);

    // Default style (idx 0) resolves to the writer's default font —
    // cellFont should still return non-null.
    const s2 = styles[2] orelse 0;
    try std.testing.expect(book.cellFont(s2) != null);
}

test "parseCommentsForSheet: rich-text comment bodies populate Comment.runs" {
    // iter53 — comments that use `<r><rPr>` wrappers surface the
    // runs alongside the flat text. Plain-text bodies still produce
    // `runs = null` (the zero-overhead path).
    var book: Book = .{
        .allocator = std.testing.allocator,
        .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer book.deinit();

    const sheet_path = "xl/worksheets/sheet1.xml";
    const rels_xml =
        \\<?xml version="1.0"?>
        \\<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        \\<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
        \\</Relationships>
    ;
    const comments_xml =
        \\<?xml version="1.0"?>
        \\<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        \\<authors><author>Alice</author></authors>
        \\<commentList>
        \\<comment ref="A1" authorId="0"><text><t>plain body</t></text></comment>
        \\<comment ref="B2" authorId="0"><text><r><rPr><b/><color rgb="FFFF0000"/></rPr><t>bold red </t></r><r><rPr><i/></rPr><t>italic tail</t></r></text></comment>
        \\</commentList></comments>
    ;

    const owned_rels = try std.testing.allocator.dupe(u8, rels_xml);
    try book.sheet_rels_data.put(std.testing.allocator, sheet_path, owned_rels);
    const owned_comments = try std.testing.allocator.dupe(u8, comments_xml);
    const comments_key = try std.testing.allocator.dupe(u8, "xl/comments1.xml");
    try book.strings.append(std.testing.allocator, comments_key);
    try book.comments_data.put(std.testing.allocator, comments_key, owned_comments);

    try parseCommentsForSheet(&book, sheet_path);

    const cs = book.comments_by_sheet.get(sheet_path) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(usize, 2), cs.len);

    // Plain comment: flat text, no runs.
    try std.testing.expectEqualStrings("plain body", cs[0].text);
    try std.testing.expectEqual(@as(?[]const RichRun, null), cs[0].runs);

    // Rich comment: flat text is concatenated runs, and `runs` is populated.
    try std.testing.expectEqualStrings("bold red italic tail", cs[1].text);
    const runs = cs[1].runs orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(usize, 2), runs.len);
    try std.testing.expectEqualStrings("bold red ", runs[0].text);
    try std.testing.expectEqual(true, runs[0].bold);
    try std.testing.expectEqual(false, runs[0].italic);
    try std.testing.expectEqual(@as(?u32, 0xFFFF0000), runs[0].color_argb);
    try std.testing.expectEqualStrings("italic tail", runs[1].text);
    try std.testing.expectEqual(false, runs[1].bold);
    try std.testing.expectEqual(true, runs[1].italic);
    try std.testing.expectEqual(@as(?u32, null), runs[1].color_argb);
}

test "parseCommentsForSheet: authors + refs + flattened rich-text bodies" {
    // Drive the parser directly with pre-populated sheet_rels_data +
    // comments_data maps so we don't need the writer (which doesn't
    // emit comments) or a real xlsx archive on disk. Mirrors the
    // Python end-to-end test but stays inside Zig for CI coverage
    // even when the Python stack doesn't run.
    var book: Book = .{
        .allocator = std.testing.allocator,
        .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer book.deinit();

    const sheet_path = "xl/worksheets/sheet1.xml";
    const rels_xml =
        \\<?xml version="1.0"?>
        \\<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        \\<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
        \\</Relationships>
    ;
    const comments_xml =
        \\<?xml version="1.0"?>
        \\<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        \\<authors><author>Alice</author><author>Bob &amp; Co</author></authors>
        \\<commentList>
        \\<comment ref="B2" authorId="0"><text><r><t>review this</t></r></text></comment>
        \\<comment ref="C3" authorId="1"><text><r><rPr><b/></rPr><t xml:space="preserve">R&amp;D </t></r><r><t>notes</t></r></text></comment>
        \\</commentList></comments>
    ;

    const owned_rels = try std.testing.allocator.dupe(u8, rels_xml);
    try book.sheet_rels_data.put(std.testing.allocator, sheet_path, owned_rels);
    const owned_comments = try std.testing.allocator.dupe(u8, comments_xml);
    const comments_key = try std.testing.allocator.dupe(u8, "xl/comments1.xml");
    try book.strings.append(std.testing.allocator, comments_key);
    try book.comments_data.put(std.testing.allocator, comments_key, owned_comments);

    try parseCommentsForSheet(&book, sheet_path);

    const cs = book.comments_by_sheet.get(sheet_path) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(usize, 2), cs.len);
    try std.testing.expectEqualDeep(CellRef{ .col = 1, .row = 2 }, cs[0].top_left);
    try std.testing.expectEqualStrings("Alice", cs[0].author);
    try std.testing.expectEqualStrings("review this", cs[0].text);
    try std.testing.expectEqualDeep(CellRef{ .col = 2, .row = 3 }, cs[1].top_left);
    try std.testing.expectEqualStrings("Bob & Co", cs[1].author);
    try std.testing.expectEqualStrings("R&D notes", cs[1].text);
}

test "parseTheme: 12-entry palette + <color theme=N/> resolution in fonts/fills/borders" {
    const theme_xml =
        \\<?xml version="1.0"?>
        \\<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        \\<a:themeElements><a:clrScheme name="Office">
        \\<a:dk1><a:sysClr val="windowText" lastClr="010101"/></a:dk1>
        \\<a:lt1><a:sysClr val="window" lastClr="020202"/></a:lt1>
        \\<a:dk2><a:srgbClr val="030303"/></a:dk2>
        \\<a:lt2><a:srgbClr val="040404"/></a:lt2>
        \\<a:accent1><a:srgbClr val="AA0000"/></a:accent1>
        \\<a:accent2><a:srgbClr val="00AA00"/></a:accent2>
        \\<a:accent3><a:srgbClr val="0000AA"/></a:accent3>
        \\<a:accent4><a:srgbClr val="AAAA00"/></a:accent4>
        \\<a:accent5><a:srgbClr val="AA00AA"/></a:accent5>
        \\<a:accent6><a:srgbClr val="00AAAA"/></a:accent6>
        \\<a:hlink><a:srgbClr val="0563C1"/></a:hlink>
        \\<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
        \\</a:clrScheme></a:themeElements></a:theme>
    ;
    var book: Book = .{
        .allocator = std.testing.allocator,
        .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer book.deinit();
    try parseTheme(&book, theme_xml);

    // 12 entries, remapped to theme-index order.
    try std.testing.expectEqual(@as(usize, 12), book.theme_colors.len);
    // theme 0 = lt1 (from schema lt1/sysClr lastClr=020202, ARGB 0xFF020202).
    try std.testing.expectEqual(@as(u32, 0xFF020202), book.theme_colors[0]);
    // theme 1 = dk1 (from schema dk1/sysClr lastClr=010101).
    try std.testing.expectEqual(@as(u32, 0xFF010101), book.theme_colors[1]);
    // theme 2 = lt2 (from schema lt2 = 040404).
    try std.testing.expectEqual(@as(u32, 0xFF040404), book.theme_colors[2]);
    // theme 3 = dk2 (from schema dk2 = 030303).
    try std.testing.expectEqual(@as(u32, 0xFF030303), book.theme_colors[3]);
    // theme 4 = accent1 = AA0000.
    try std.testing.expectEqual(@as(u32, 0xFFAA0000), book.theme_colors[4]);
    // theme 11 = folHlink = 954F72.
    try std.testing.expectEqual(@as(u32, 0xFF954F72), book.theme_colors[11]);

    // Now drive parseStyles with a font that references theme color 4.
    // Must resolve to 0xFFAA0000 (accent1) — pre-iter52 this would be null.
    const styles_xml =
        \\<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        \\<fonts count="2">
        \\<font><sz val="11"/><color theme="4"/><name val="Calibri"/></font>
        \\<font><b/><color rgb="FFFF0000"/></font>
        \\</fonts>
        \\<fills count="1">
        \\<fill><patternFill patternType="solid"><fgColor theme="6"/><bgColor theme="11"/></patternFill></fill>
        \\</fills>
        \\<borders count="1">
        \\<border><left style="thin"><color theme="1"/></left></border>
        \\</borders>
        \\<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
        \\</styleSheet>
    ;
    try parseStyles(&book, styles_xml);

    // Font 0 referenced theme=4 (accent1) → AA0000 with opaque alpha.
    try std.testing.expectEqual(@as(?u32, 0xFFAA0000), book.fonts[0].color_argb);
    // Font 1 kept explicit rgb.
    try std.testing.expectEqual(@as(?u32, 0xFFFF0000), book.fonts[1].color_argb);
    // Fill 0 fg=theme 6 → accent3 = 0000AA.
    try std.testing.expectEqual(@as(?u32, 0xFF0000AA), book.fills[0].fg_color_argb);
    // Fill 0 bg=theme 11 → folHlink = 954F72.
    try std.testing.expectEqual(@as(?u32, 0xFF954F72), book.fills[0].bg_color_argb);
    // Border left.color → theme 1 = dk1 = 010101.
    try std.testing.expectEqual(@as(?u32, 0xFF010101), book.borders[0].left.color_argb);
}

test "parseStyles: cellXfs handles <xf/> and <xf> variants (no attrs)" {
    // Regression guard for an iter35 audit finding: `<xf ` (with
    // required trailing space) silently dropped bare `<xf/>` or
    // `<xf>` entries, shifting every subsequent style index. This
    // fixture mixes all three shapes plus a trailing attributed
    // entry; if the parser reverts to requiring a space, the count
    // comes back as 1 instead of 4 and the xf at index 3 loses its
    // fontId.
    const styles_xml =
        \\<?xml version="1.0"?>
        \\<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        \\<cellXfs count="4">
        \\<xf/>
        \\<xf></xf>
        \\<xf numFmtId="0"/>
        \\<xf numFmtId="14" fontId="2" fillId="1" borderId="0"/>
        \\</cellXfs>
        \\</styleSheet>
    ;
    var book: Book = .{
        .allocator = std.testing.allocator,
        .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer book.deinit();
    try parseStyles(&book, styles_xml);

    try std.testing.expectEqual(@as(usize, 4), book.cell_xf_numfmt_ids.len);
    try std.testing.expectEqual(@as(u32, 0), book.cell_xf_numfmt_ids[0]);
    try std.testing.expectEqual(@as(u32, 0), book.cell_xf_numfmt_ids[1]);
    try std.testing.expectEqual(@as(u32, 0), book.cell_xf_numfmt_ids[2]);
    try std.testing.expectEqual(@as(u32, 14), book.cell_xf_numfmt_ids[3]);
    try std.testing.expectEqual(@as(u32, 2), book.cell_xf_font_ids[3]);
    try std.testing.expectEqual(@as(u32, 1), book.cell_xf_fill_ids[3]);
}

test "Book.cellBorder: round-trip sided styles + color through writer" {
    const tmp_path = "/tmp/zlsx_reader_cell_border.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        const boxed = try w.addStyle(.{
            .border_left = .{ .style = .thin, .color_argb = 0xFF000000 },
            .border_right = .{ .style = .thin, .color_argb = 0xFF000000 },
            .border_top = .{ .style = .medium, .color_argb = 0xFFFF0000 },
            .border_bottom = .{ .style = .medium, .color_argb = 0xFFFF0000 },
        });
        var sheet = try w.addSheet("S");
        try sheet.writeRowStyled(
            &.{ .{ .string = "boxed" }, .{ .string = "plain" } },
            &.{ boxed, 0 },
        );
        try w.save(tmp_path);
    }

    var book = try Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    _ = try rows.next();
    const styles = rows.styleIndices();

    const s0 = styles[0] orelse return error.TestUnexpectedResult;
    const b0 = book.cellBorder(s0) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualStrings("thin", b0.left.style);
    try std.testing.expectEqual(@as(?u32, 0xFF000000), b0.left.color_argb);
    try std.testing.expectEqualStrings("thin", b0.right.style);
    try std.testing.expectEqualStrings("medium", b0.top.style);
    try std.testing.expectEqual(@as(?u32, 0xFFFF0000), b0.top.color_argb);
    try std.testing.expectEqualStrings("medium", b0.bottom.style);
    try std.testing.expectEqualStrings("", b0.diagonal.style);

    // Default style resolves to a Border with all empty sides.
    const s1 = styles[1] orelse 0;
    const b1 = book.cellBorder(s1) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualStrings("", b1.left.style);
    try std.testing.expectEqualStrings("", b1.top.style);
}

test "Book.cellFill: round-trip solid fg/bg and pattern 'none' default" {
    const tmp_path = "/tmp/zlsx_reader_cell_fill.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        const red_fill = try w.addStyle(.{
            .fill_pattern = .solid,
            .fill_fg_argb = 0xFFFF0000,
        });
        var sheet = try w.addSheet("S");
        try sheet.writeRowStyled(
            &.{ .{ .string = "filled" }, .{ .string = "plain" } },
            &.{ red_fill, 0 },
        );
        try w.save(tmp_path);
    }

    var book = try Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    _ = try rows.next();
    const styles = rows.styleIndices();

    const s0 = styles[0] orelse return error.TestUnexpectedResult;
    const f0 = book.cellFill(s0) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualStrings("solid", f0.pattern);
    try std.testing.expectEqual(@as(?u32, 0xFFFF0000), f0.fg_color_argb);

    // Style 0 resolves to the writer's default fill (patternType="none").
    const s1 = styles[1] orelse 0;
    const f1 = book.cellFill(s1) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualStrings("none", f1.pattern);
    try std.testing.expectEqual(@as(?u32, null), f1.fg_color_argb);
}

test "Book.richRuns: color / size / font_name from <rPr>" {
    const sst_xml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" ++
        "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"3\">" ++
        // Full-fat rPr: bold + color + size + font
        "<si><r><rPr><b/><sz val=\"14\"/><color rgb=\"FFFF0000\"/><rFont val=\"Arial\"/></rPr><t>styled</t></r></si>" ++
        // Theme-only color — should NOT populate color_argb
        "<si><r><rPr><color theme=\"1\"/><sz val=\"11.5\"/></rPr><t>themed</t></r></si>" ++
        // No rPr children we care about
        "<si><r><t>bare</t></r></si>" ++
        "</sst>";

    var book: Book = .{
        .allocator = std.testing.allocator,
        .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer book.deinit();
    const owned = try std.testing.allocator.dupe(u8, sst_xml);
    book.shared_strings_xml = owned;
    try parseSharedStrings(&book, owned);

    // Styled: every property surfaced.
    const r0 = book.richRuns(0) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(usize, 1), r0.len);
    try std.testing.expectEqual(true, r0[0].bold);
    try std.testing.expectEqual(@as(?u32, 0xFFFF0000), r0[0].color_argb);
    try std.testing.expectEqual(@as(?f32, 14.0), r0[0].size);
    try std.testing.expectEqualStrings("Arial", r0[0].font_name);

    // Theme color is intentionally skipped; size still parses.
    const r1 = book.richRuns(1) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(?u32, null), r1[0].color_argb);
    try std.testing.expectEqual(@as(?f32, 11.5), r1[0].size);
    try std.testing.expectEqualStrings("", r1[0].font_name);

    // Bare run: all optionals null / empty.
    const r2 = book.richRuns(2) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(?u32, null), r2[0].color_argb);
    try std.testing.expectEqual(@as(?f32, null), r2[0].size);
    try std.testing.expectEqualStrings("", r2[0].font_name);
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

test "fuzz parseStyles" {
    const iters = fuzzIterations();
    var prng = std.Random.DefaultPrng.init(fuzzSeed());
    const rng = prng.random();
    var buf: [fuzz_max_input_len]u8 = undefined;
    for (0..iters) |_| {
        const input = randomInput(rng, &buf);
        var book: Book = .{ .allocator = std.testing.allocator, .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator) };
        defer book.deinit();
        parseStyles(&book, input) catch {};
    }
}

test "fuzz parseCommentsForSheet" {
    const iters = fuzzIterations();
    var prng = std.Random.DefaultPrng.init(fuzzSeed());
    const rng = prng.random();
    var buf: [fuzz_max_input_len]u8 = undefined;
    const sheet_path = "xl/worksheets/sheet1.xml";
    for (0..iters) |_| {
        const input = randomInput(rng, &buf);
        const mid = input.len / 2;
        const rels_input = input[0..mid];
        const comments_input = input[mid..];
        var book: Book = .{ .allocator = std.testing.allocator, .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator) };
        defer book.deinit();

        const owned_rels = std.testing.allocator.dupe(u8, rels_input) catch continue;
        book.sheet_rels_data.put(std.testing.allocator, sheet_path, owned_rels) catch {
            std.testing.allocator.free(owned_rels);
            continue;
        };
        const owned_comments = std.testing.allocator.dupe(u8, comments_input) catch continue;
        const comments_key = std.testing.allocator.dupe(u8, "xl/comments1.xml") catch {
            std.testing.allocator.free(owned_comments);
            continue;
        };
        book.strings.append(std.testing.allocator, comments_key) catch {
            std.testing.allocator.free(owned_comments);
            std.testing.allocator.free(comments_key);
            continue;
        };
        book.comments_data.put(std.testing.allocator, comments_key, owned_comments) catch {
            std.testing.allocator.free(owned_comments);
            continue;
        };

        parseCommentsForSheet(&book, sheet_path) catch {};
    }
}

test "fuzz writer.addComment: adversarial author + text never crash emission" {
    // Writer-side fuzz: random bytes fed into addComment (ref, author,
    // text). The ref path is the interesting one — a bad ref returns
    // InvalidCommentRef before the author/text ever reach the emit
    // buffers. A pass ref with adversarial author/text must still
    // emit valid XML (entity-escaped) so save() completes without
    // a panic.
    const iters = fuzzIterations();
    var prng = std.Random.DefaultPrng.init(fuzzSeed());
    const rng = prng.random();
    var buf: [fuzz_max_input_len]u8 = undefined;

    const writer = @import("writer.zig");
    const tmp_path = "/tmp/zlsx_fuzz_addcomment.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    for (0..iters) |_| {
        const input = randomInput(rng, &buf);
        // Split the random bytes into ref / author / text slices.
        const third = input.len / 3;
        const ref = input[0..third];
        const author = input[third .. 2 * third];
        const text = input[2 * third ..];

        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = w.addSheet("S") catch continue;
        // Most refs will be rejected at validate time; when they pass
        // (e.g. random bytes happen to spell "A1"), the emit path
        // must tolerate any author/text content.
        _ = sheet.addComment(ref, author, text) catch {};
        sheet.writeRow(&.{.{ .string = "x" }}) catch continue;
        w.save(tmp_path) catch {};
    }
}

test "fuzz writer.addConditionalFormat + addDxf: adversarial inputs never crash emission" {
    // Writer-side fuzz for iter40-41 surfaces. Random bytes feed
    // into addConditionalFormatCellIs / Expression (range / formula
    // slots) + addDxf (argb values). Most random ranges get
    // rejected at validate time; valid ones with adversarial
    // formulas must still emit valid CF XML so save() completes.
    const iters = fuzzIterations();
    var prng = std.Random.DefaultPrng.init(fuzzSeed());
    const rng = prng.random();
    var buf: [fuzz_max_input_len]u8 = undefined;

    const writer = @import("writer.zig");
    const tmp_path = "/tmp/zlsx_fuzz_addcf.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    const ops = [_]writer.CfOperator{
        .less_than,          .equal,                 .greater_than,
        .between,            .not_between,           .not_equal,
        .less_than_or_equal, .greater_than_or_equal,
    };

    for (0..iters) |_| {
        const input = randomInput(rng, &buf);
        const third = input.len / 3;
        const range = input[0..third];
        const formula1 = input[third .. 2 * third];
        const formula2 = input[2 * third ..];

        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        // addDxf with adversarial ARGB-like u32: always succeeds
        // (Dxf has no intake validation beyond u32 bounds).
        const dxf_id = w.addDxf(.{
            .font_bold = (input.len & 1) != 0,
            .font_italic = (input.len & 2) != 0,
            .font_color_argb = if (input.len > 4) std.mem.readInt(u32, input[0..4], .little) else null,
            .fill_fg_argb = null,
        }) catch continue;

        var sheet = w.addSheet("S") catch continue;
        const op = ops[@intCast(input.len % ops.len)];
        // cellIs — random operator; formula2 only when the op
        // demands two formulas (between / not_between per OOXML).
        const needs_two = op == .between or op == .not_between;
        const f2: ?[]const u8 = if (needs_two) formula2 else null;
        _ = sheet.addConditionalFormatCellIs(range, op, formula1, f2, dxf_id) catch {};
        _ = sheet.addConditionalFormatExpression(range, formula1, dxf_id) catch {};
        sheet.writeRow(&.{.{ .string = "x" }}) catch continue;
        w.save(tmp_path) catch {};
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
        .row_styles = .{},
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
            .row_styles = .{},
            .arena = std.heap.ArenaAllocator.init(std.testing.allocator),
        };
        defer rows.deinit();

        // Consume the rows — may error, must not panic.
        while (rows.next() catch null) |_| {}
    }
}

// ─── iter54 slice B: lazy sheet / comments extraction ───────────────

test "openLazy path loads sheets lazily, yields identical state to open()" {
    // Build a small multi-sheet workbook with merged ranges, hyperlinks,
    // data validations, and comments so every per-sheet side-index gets
    // a round-trip comparison.
    const tmp_path = "/tmp/zlsx_iter54_lazy_struct.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();

        var s1 = try w.addSheet("Alpha");
        try s1.addMergedCell("A1:B1");
        try s1.addHyperlink("A2", "https://example.com/a");
        try s1.writeRow(&.{ .{ .string = "hdr1" }, .{ .string = "hdr2" } });
        try s1.writeRow(&.{ .{ .number = 1.0 }, .{ .number = 2.0 } });
        try s1.addComment("A1", "me", "hello");

        var s2 = try w.addSheet("Beta");
        try s2.writeRow(&.{.{ .string = "x" }});
        try s2.writeRow(&.{.{ .number = 42.0 }});

        try w.save(tmp_path);
    }

    // Eager path.
    var eager = try Book.open(std.testing.allocator, tmp_path);
    defer eager.deinit();

    // Lazy path — right after open, sheet_data is empty.
    var lazy = try Book.openLazy(std.testing.allocator, tmp_path);
    defer lazy.deinit();

    try std.testing.expectEqual(@as(usize, 0), lazy.sheet_data.count());
    try std.testing.expectEqual(eager.sheets.len, eager.sheet_data.count());
    try std.testing.expectEqual(eager.sheets.len, lazy.sheets.len);

    // Iterate each sheet via lazy.rows — must match eager.rows row-for-row.
    for (eager.sheets, lazy.sheets) |eager_sheet, lazy_sheet| {
        try std.testing.expectEqualStrings(eager_sheet.path, lazy_sheet.path);
        var er = try eager.rows(eager_sheet, std.testing.allocator);
        defer er.deinit();
        var lr = try lazy.rows(lazy_sheet, std.testing.allocator);
        defer lr.deinit();
        while (true) {
            const er_row = try er.next();
            const lr_row = try lr.next();
            if (er_row == null and lr_row == null) break;
            try std.testing.expect(er_row != null and lr_row != null);
            try std.testing.expectEqual(er_row.?.len, lr_row.?.len);
            for (er_row.?, lr_row.?) |ec, lc| {
                try std.testing.expectEqualDeep(ec, lc);
            }
        }
    }

    // After iterating every sheet, lazy sheet_data matches eager.
    try std.testing.expectEqual(lazy.sheets.len, lazy.sheet_data.count());
    try std.testing.expectEqual(eager.merged_ranges.count(), lazy.merged_ranges.count());
    try std.testing.expectEqual(eager.hyperlinks_by_sheet.count(), lazy.hyperlinks_by_sheet.count());
    try std.testing.expectEqual(eager.data_validations_by_sheet.count(), lazy.data_validations_by_sheet.count());
    try std.testing.expectEqual(eager.comments_by_sheet.count(), lazy.comments_by_sheet.count());
}

test "openLazy -> ensureSheetLoaded is idempotent" {
    const tmp_path = "/tmp/zlsx_iter54_lazy_idempotent.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("One");
        try s.addMergedCell("A1:A2");
        try s.writeRow(&.{.{ .string = "x" }});
        try s.writeRow(&.{.{ .number = 1.0 }});
        try w.save(tmp_path);
    }

    var book = try Book.openLazy(std.testing.allocator, tmp_path);
    defer book.deinit();

    try std.testing.expectEqual(@as(usize, 0), book.sheet_data.count());

    {
        var r = try book.rows(book.sheets[0], std.testing.allocator);
        defer r.deinit();
        while (try r.next()) |_| {}
    }
    try std.testing.expectEqual(@as(usize, 1), book.sheet_data.count());
    const merged_after_first = book.merged_ranges.count();

    // Second call: cache hit, no double-parse.
    {
        var r = try book.rows(book.sheets[0], std.testing.allocator);
        defer r.deinit();
        while (try r.next()) |_| {}
    }
    try std.testing.expectEqual(@as(usize, 1), book.sheet_data.count());
    try std.testing.expectEqual(merged_after_first, book.merged_ranges.count());
}

test "openLazy parses workbook-wide styles + theme (no sheet load required)" {
    const tmp_path = "/tmp/zlsx_iter54_lazy_styles.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        // Any non-General number format forces a custom numFmt entry
        // that lives in styles.xml — exactly the path openLazy must
        // still parse so numberFormat lookups work before any sheet
        // is loaded.
        const s_idx = try w.addStyle(.{ .number_format = "0.00%" });
        var s = try w.addSheet("One");
        try s.writeRowStyled(&.{.{ .number = 0.5 }}, &.{s_idx});
        try w.save(tmp_path);
    }

    var book = try Book.openLazy(std.testing.allocator, tmp_path);
    defer book.deinit();

    // Sheet XML must still be unloaded.
    try std.testing.expectEqual(@as(usize, 0), book.sheet_data.count());

    // Styles / theme tables must be populated — numberFormat is the
    // workbook-wide lookup that regressed in slice B before this fix.
    try std.testing.expect(book.cell_xf_numfmt_ids.len > 0);
    const fmt = book.numberFormat(@intCast(book.cell_xf_numfmt_ids.len - 1));
    try std.testing.expect(fmt != null);
    try std.testing.expect(std.mem.indexOf(u8, fmt.?, "%") != null);
}

test "preloadSheet populates per-sheet metadata without rows() iteration" {
    const tmp_path = "/tmp/zlsx_iter54_preload_sheet.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("One");
        try s.addMergedCell("A1:C1");
        try s.writeRow(&.{.{ .string = "merged" }});
        try s.writeRow(&.{.{ .number = 1.0 }});
        try w.save(tmp_path);
    }

    var book = try Book.openLazy(std.testing.allocator, tmp_path);
    defer book.deinit();

    // Before preload, metadata getter returns empty on the lazy path.
    try std.testing.expectEqual(@as(usize, 0), book.mergedRanges(book.sheets[0]).len);

    try book.preloadSheet(book.sheets[0]);

    // Sheet XML is now cached; metadata getter returns populated.
    try std.testing.expectEqual(@as(usize, 1), book.sheet_data.count());
    try std.testing.expectEqual(@as(usize, 1), book.mergedRanges(book.sheets[0]).len);

    // Second preload is idempotent.
    try book.preloadSheet(book.sheets[0]);
    try std.testing.expectEqual(@as(usize, 1), book.sheet_data.count());
}
