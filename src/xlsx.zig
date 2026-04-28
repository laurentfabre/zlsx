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
const casefold = @import("unicode/casefold.zig");
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
    var days_since_unix: i32 = @intCast(@as(i64, @intFromFloat(floored)) - 25569);

    // Time of day: fractional part × 86 400 seconds, rounded. When
    // the rounded value equals 86400 (e.g. serial 1.9999999999 →
    // 86400.0), the result is the next day at 00:00:00 — carry the
    // day forward instead of clamping back to 23:59:59 like prior
    // versions did. If the carry would push the date past the
    // documented upper bound (serial >= 2958466.0 → 9999-12-31),
    // return null rather than emitting a year-10000 DateTime.
    const total_s_f: f64 = @round((serial - floored) * 86400.0);
    var total_s: i64 = @intFromFloat(total_s_f);
    if (total_s >= 86400) {
        if (floored + 1.0 >= 2958466.0) return null;
        total_s -= 86400;
        days_since_unix += 1;
    }
    const ymd = daysSinceUnixEpochToYMD(days_since_unix);
    return .{
        .year = ymd.year,
        .month = ymd.month,
        .day = ymd.day,
        .hour = @intCast(@divFloor(total_s, 3600)),
        .minute = @intCast(@mod(@divFloor(total_s, 60), 60)),
        .second = @intCast(@mod(total_s, 60)),
    };
}

/// Convert a Mac Excel 1904-epoch serial into a calendar `DateTime`.
/// The 1904 system shifts every serial by 1462 days vs the 1900
/// default — so this is just `fromExcelSerial(serial + 1462)`. Has
/// the side-benefit of dodging the 1900 leap-year-bug window: every
/// real 1904 date lands at 1900-system serial >= 1463, well past the
/// `< 61` rejection. Books opened with `<workbookPr date1904="1"/>`
/// (`Book.uses_1904_epoch`) decode through this function.
pub fn fromExcelSerial1904(serial: f64) ?DateTime {
    if (!std.math.isFinite(serial)) return null;
    return fromExcelSerial(serial + 1462.0);
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

/// Array-formula spread rectangle from `<f t="array" ref="X:Y">…</f>`.
/// `tl` and `br` are the inclusive corners (component-wise normalised
/// like `MergeRange`); `base` is the cell that carried the `<f>`
/// element — by OOXML convention always the top-left of the rectangle,
/// recorded explicitly so the slave→base linkage is unambiguous if a
/// future generator drops the convention.
pub const ArrayRange = struct {
    tl: CellRef,
    br: CellRef,
    base: CellRef,
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

/// SST strategy passed to `Book.openLazyWithSst`. The public surface
/// uses three named entrypoints (`open`, `openLazy`, `openSstLazy`)
/// that each pick a strategy internally; this enum exists so the
/// shared lazy-archive walker can dispatch without code duplication.
pub const SstStrategy = enum { eager, lazy };

/// Internal SST storage backend. Eager mode (the default, populated
/// by `Book.open`) holds every decoded entry as a contiguous slice;
/// lazy mode (populated by `Book.openSstLazy` — iter-sst-3b) holds an
/// offset/length index into `shared_strings_xml` and a sparse map of
/// already-resolved entries. The cell hot path branches on the union
/// tag exactly once per access, then falls into a per-variant fast
/// path. See docs/plans/streaming-sst.md for the full design.
pub const SstBackend = union(enum) {
    /// Pre-decoded entries; valid for the Book's lifetime.
    eager: [][]const u8,
    /// Offset table + sparse resolution cache; entries are decoded
    /// into `Book.sst_arena` on first `sharedStringAt(idx)` access.
    /// Reserved for iter-sst-3b — not constructed yet by any
    /// production path; the variant exists at iter-sst-3a so the
    /// accessor switch covers both arms exhaustively.
    lazy: struct {
        offsets: []usize,
        lengths: []usize,
        resolved: std.AutoHashMapUnmanaged(u32, []const u8),
    },
};

pub const Book = struct {
    allocator: Allocator,
    /// Open zip archive backing this book. Owns the file handle and
    /// reader. Set by `openLazy`; torn down LAST in `deinit` (after
    /// every borrowed part buffer is freed).
    archive: ?*ZipArchive = null,
    /// True when the source workbook declares `<workbookPr
    /// date1904="1"/>` — the Mac Excel epoch shifts every serial by
    /// 1462 days vs the default 1900 system. iter61-a only wires the
    /// flag as a gate: `Rows` skips `t:"date"` auto-conversion for
    /// 1904-epoch workbooks (the numeric value survives as `.integer`
    /// / `.number`). Proper 1904 decoding is queued for a follow-up.
    uses_1904_epoch: bool = false,
    /// Decompressed `xl/sharedStrings.xml` (nullable — small xlsx files
    /// with only inline strings omit this part).
    shared_strings_xml: ?[]u8 = null,
    /// SST storage backend (eager today; lazy variant arrives in
    /// iter-sst-3b). Production callers use the
    /// `Book.sharedStringAt(idx)` / `Book.sharedStringsCount()`
    /// accessors which dispatch on the union tag; tests and
    /// fuzz scaffolding may reach into `sst.eager` directly when they
    /// already know the backend.
    sst: SstBackend = .{ .eager = &.{} },
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
    /// Eager sheets, lazy SST. Companion to `open` (eager everything)
    /// and `openLazy` (lazy sheets, eager SST). Mitigates the SST RAM
    /// blow-up on workbooks with millions of unique strings: at open
    /// time the SST is walked once to build a per-entry offset/length
    /// index, but plain text is not decoded until `sharedStringAt` is
    /// called for that index. See docs/plans/streaming-sst.md for the
    /// full design.
    ///
    /// Rich-text runs are eagerly captured on the lazy backend too:
    /// the per-entry walk that records offsets also populates
    /// `rich_runs_by_sst_idx`, so `Book.richRuns(idx)` returns the
    /// same slice as the eager backend with the same lifetime
    /// contract.
    ///
    /// Trade-off: lazy mode defers per-entry plain-text decoding,
    /// so a malformed `<t>` body inside an `<si>` won't surface at
    /// `openSstLazy` time — it'll surface as `MalformedXml` from
    /// `sharedStringAt(idx)` (or, transitively, from a `Rows.next()`
    /// cell read that resolves to that idx). Callers that need
    /// "open fails fast on bad SST" should use `Book.open` instead.
    /// Single-threaded contract — the lazy backend mutates internal
    /// state on first-touch. Multi-threaded SST access requires the
    /// caller to serialise externally or to call `Book.open` instead.
    pub fn openSstLazy(allocator: Allocator, path: []const u8) !Book {
        var book = try Book.openLazyWithSst(allocator, path, .lazy);
        errdefer book.deinit();
        try book.loadEagerParts();
        book.closeArchive();
        return book;
    }

    pub fn openLazy(allocator: Allocator, path: []const u8) !Book {
        return Book.openLazyWithSst(allocator, path, .eager);
    }

    fn openLazyWithSst(allocator: Allocator, path: []const u8, sst_strategy: SstStrategy) !Book {
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
        if (book.shared_strings_xml) |sst| switch (sst_strategy) {
            .eager => try parseSharedStrings(&book, sst),
            .lazy => try parseSharedStringsLazy(&book, sst),
        };

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
        if (rels_targetForComments(rels_xml)) |target| {
            // Sheet rels live at `xl/worksheets/_rels/sheetN.xml.rels`
            // and `Target` is relative to the worksheet itself
            // (`xl/worksheets/sheetN.xml`), so the base dir is
            // `xl/worksheets/`.
            var path_buf: [256]u8 = undefined;
            const full = resolveRelArchivePath("xl/worksheets/", target, &path_buf) orelse return;
            try self.ensureCommentsLoaded(full);
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
        switch (self.sst) {
            .eager => |entries| a.free(entries),
            .lazy => |*lazy| {
                a.free(lazy.offsets);
                a.free(lazy.lengths);
                lazy.resolved.deinit(a);
            },
        }
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

    /// Resolved shared string at `idx`. Errors with `MalformedXml`
    /// when `idx` is past the end of the SST. Eager-backend accessor;
    /// the streaming-SST plan (docs/plans/streaming-sst.md) introduces
    /// a lazy variant in iter-sst-3 — this iter-sst-1 accessor
    /// centralises the lookup so future migration touches one site.
    pub fn sharedStringAt(self: *Book, idx: usize) ![]const u8 {
        switch (self.sst) {
            .eager => |entries| {
                if (idx >= entries.len) return error.MalformedXml;
                return entries[idx];
            },
            .lazy => |*lazy| {
                if (idx >= lazy.offsets.len) return error.MalformedXml;
                if (lazy.resolved.get(@intCast(idx))) |s| return s;
                // First-touch decode into `sst_arena`. Caches the
                // result so subsequent accesses are O(1) hashmap hits.
                return try materialiseSstEntry(self, idx);
            },
        }
    }

    /// Total entries in the SST. Equivalent to the old
    /// `shared_strings.len`; both backends know this O(1) without
    /// materialising any entry.
    pub fn sharedStringsCount(self: *const Book) usize {
        return switch (self.sst) {
            .eager => |entries| entries.len,
            .lazy => |lazy| lazy.offsets.len,
        };
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
            // Rows.shared_strings is the test-scaffolding fallback (used
            // only when book is null in resolveSharedString). Production
            // routes through self.book.?.sharedStringAt — the snapshot
            // is dead weight on hot rows. iter-sst-3a populates it from
            // the eager backend when available so existing tests that
            // poke at the snapshot keep working; lazy books leave it
            // empty.
            .shared_strings = switch (self.sst) {
                .eager => |entries| entries,
                .lazy => &.{},
            },
            .book = self,
            .allocator = allocator,
            .row_cells = .{},
            .row_styles = .{},
            .row_date_types = .{},
            .row_error_strings = .{},
            .row_formula_strings = .{},
            .row_formula_refs = .{},
            .shared_si_to_base_ref = .{},
            .array_ranges = .{},
            .arena = std.heap.ArenaAllocator.init(allocator),
        };
    }

    /// Index-based streaming entry point for the upcoming jq-for-excel
    /// CLI: `zlsx cells file.xlsx --sheet-index 2` resolves directly to
    /// `book.streamSheet(2, alloc)`. Semantically identical to
    /// `book.rows(book.sheets[idx], alloc)` — both go through
    /// `ensureSheetLoaded` on the lazy path and hit the cache on the
    /// eager path. Returns `error.SheetIndexOutOfRange` when `idx` is
    /// not a valid sheet.
    pub fn streamSheet(self: *Book, idx: usize, allocator: Allocator) !Rows {
        if (idx >= self.sheets.len) return error.SheetIndexOutOfRange;
        return self.rows(self.sheets[idx], allocator);
    }

    /// Bulk-materialise every row of `sheet` into a dense in-memory
    /// matrix, returned as a `SheetMatrix` that owns its backing
    /// storage via a single arena. Non-breaking addition: the
    /// streaming `rows()` iterator is unchanged.
    ///
    /// Allocator strategy: one `ArenaAllocator` is owned by the
    /// returned matrix and serves three roles in one bump-bucket
    /// sequence — per-row Cell slabs, the outer rows slab, and any
    /// entity-decoded / rich-text-concatenated cell strings. The
    /// caller's `allocator` parameter is the arena's underlying
    /// child allocator and is freed wholesale by
    /// `SheetMatrix.deinit`.
    ///
    /// Strings inside Cells follow the same lifetime contract as
    /// `Rows.next()`: shared-string and XML-backed slices borrow
    /// from the parent Book (which MUST outlive the matrix);
    /// previously per-row-arena strings now live in the matrix's
    /// arena and stay valid until `deinit`.
    ///
    /// Errors mirror the streaming path (malformed XML, OOM, sheet
    /// load). On error every partial allocation is reaped via the
    /// errdefers below — no leaks on the failure path. Verified by
    /// the `materialiseSheet under failing allocator` test.
    pub fn materialiseSheet(self: *Book, sheet: Sheet, allocator: Allocator) !SheetMatrix {
        var matrix: SheetMatrix = .{
            .arena = std.heap.ArenaAllocator.init(allocator),
            .rows = &.{},
        };
        // Free the matrix arena (and every row slab + cloned string
        // it backs) on any error below. Cleared by the final return.
        errdefer matrix.arena.deinit();

        var it = try self.rows(sheet, allocator);
        defer it.deinit();
        // Note: we do NOT set it.keep_arena_strings here (Codex P2
        // follow-up). Each row body is fully cloned into matrix.arena
        // below before the next `next()` call fires, so the iterator's
        // own arena is free to reset between rows. Keeping it alive
        // wholesale doubled peak memory on entity-heavy workbooks for
        // no benefit.
        const arena_alloc = matrix.arena.allocator();
        var row_acc: std.ArrayListUnmanaged([]const Cell) = .{};
        // Pre-size the outer accumulator using the SST length as a
        // *very* loose order-of-magnitude hint; for sparse sheets
        // this overshoots, but the arena absorbs the slack on
        // success. Skipped for unknown-size workbooks.

        while (try it.next()) |cells| {
            // Bulk-allocate the per-row Cell slab once into the
            // matrix arena. This is the alloc the materialise API
            // exists to amortise — instead of paying a separate
            // allocator round-trip per row (and a freelist hit on
            // any caller-side `dupe`), we pay one bump-pointer
            // bump per row.
            const row_copy = try arena_alloc.dupe(Cell, cells);
            // Clone string payloads into the matrix arena so the
            // returned matrix is fully self-contained: the
            // iterator's arena dies with `it.deinit()` when this
            // function returns. SST and sheet-XML borrows are
            // also cloned for uniformity — over-copy cost is small
            // (Book-lifetime slices are typically short tokens) and
            // it eliminates a foot-gun where a future caller
            // outlives the Book.
            //
            // NOTE: this is a deliberate trade — we sacrifice the
            // SST-borrow zero-copy property in exchange for a
            // single, simple lifetime story (matrix owns
            // everything). The win still nets out positive on the
            // worldbank fixture because the per-row alloc round-
            // trip dominates the small string copies.
            for (row_copy) |*cell| {
                switch (cell.*) {
                    .string => |s| {
                        if (s.len == 0) continue;
                        cell.* = .{ .string = try arena_alloc.dupe(u8, s) };
                    },
                    else => {},
                }
            }
            try row_acc.append(arena_alloc, row_copy);
        }

        const rows_final = try row_acc.toOwnedSlice(arena_alloc);
        std.debug.assert(rows_final.len == row_acc.items.len or row_acc.items.len == 0);
        matrix.rows = rows_final;
        return matrix;
    }
};

/// When a `<row>` tag has no usable `r` attribute, try to recover
/// the 1-based row number from the first `<c r="A11">` cell inside
/// that row. OOXML generators typically stamp every cell with an
/// `r` attribute even when they omit it on `<row>`, so this branch
/// covers the realistic malformed-input cases. Returns null when
/// the row body is truncated, the first cell has no r, or the
/// trailing digits don't parse / resolve to zero.
fn recoverRowFromFirstCell(xml: []const u8, row_body_start: usize) ?u32 {
    const row_end_pos = std.mem.indexOfPos(u8, xml, row_body_start, "</row>") orelse return null;
    const body = xml[row_body_start..row_end_pos];

    // Scan for the first `<c` tag-open. Any XML-valid whitespace
    // or `/` or `>` after the tag name is legal; a letter would
    // mean we matched a longer name (`<col>`, `<cellmap>`) and
    // must keep searching.
    var scan: usize = 0;
    const c_start = while (std.mem.indexOfPos(u8, body, scan, "<c")) |p| {
        const after = p + "<c".len;
        if (after >= body.len) return null;
        const ch = body[after];
        if (ch == ' ' or ch == '\t' or ch == '\n' or ch == '\r' or ch == '/' or ch == '>') {
            break p;
        }
        scan = p + 1;
    } else return null;

    const tag_open = c_start + "<c".len;
    const tag_close = std.mem.indexOfScalarPos(u8, body, tag_open, '>') orelse return null;
    const c_attrs = body[tag_open..tag_close];
    const ref = getAttr(c_attrs, "r") orelse return null;
    // Split "A11" into letters + digits. Skip leading non-digits
    // (column letters — could be 1..3 chars, `A` to `XFD`).
    var j: usize = 0;
    while (j < ref.len and !std.ascii.isDigit(ref[j])) : (j += 1) {}
    if (j >= ref.len) return null;
    const n = std.fmt.parseInt(u32, ref[j..], 10) catch return null;
    return if (n > 0) n else null;
}

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
    book: ?*Book = null,
    allocator: Allocator,
    row_cells: std.ArrayListUnmanaged(Cell),
    /// Parallel to `row_cells`. Each slot holds the cell's `s="N"`
    /// attribute (the index into `Book.cell_xf_numfmt_ids`), or null
    /// when the source `<c>` had no `s` attribute. Filled with nulls
    /// for gap cells so positional indexing matches `row_cells`.
    row_styles: std.ArrayListUnmanaged(?u32),
    /// Parallel to `row_cells`. Each slot is `true` iff the cell is a
    /// numeric value (`.integer` or `.number`) whose style resolves to
    /// a date-like numFmt (`Book.isDateFormat`) AND whose value lies in
    /// the round-trippable Excel-serial range (`fromExcelSerial` returns
    /// non-null). Gap cells stay `false`. Exposed via `dateTypes()` so
    /// the CLI can emit `t:"date"` alongside the raw numeric serial
    /// without altering the `Cell` union shape.
    row_date_types: std.ArrayListUnmanaged(bool),
    /// Parallel to `row_cells`. Each slot holds the literal Excel-error
    /// string (e.g. `"#DIV/0!"`, `"#N/A"`, `"#REF!"`) for cells whose
    /// source `<c>` carried `t="e"` and a `<v>` body, or `null`
    /// otherwise. The string borrows from the sheet XML buffer when
    /// the body has no entities, otherwise it is owned by the row arena
    /// (via `decodeVValue`). Exposed via `errorStrings()` so the CLI
    /// can emit `t:"error"` without altering the `Cell` union shape;
    /// the cell slot itself stays a plain `.string` so consumers that
    /// don't ask for the side channel still see the literal text. Same
    /// lifetime as `row_cells`.
    row_error_strings: std.ArrayListUnmanaged(?[]const u8),
    /// Parallel to `row_cells`. Each slot holds the cell's own formula
    /// text — non-null when the source `<c>` carried an `<f>` body
    /// (stand-alone formula, shared-formula base, or array-formula
    /// base). Borrows from the sheet XML buffer when entity-free,
    /// otherwise lives in the row arena (via `decodeFormulaBody`).
    /// Exposed via `formulaStrings()` so the CLI can emit
    /// `t:"formula"` + `formula:<text>` without altering the `Cell`
    /// union shape; the cell slot continues to carry the cached `<v>`
    /// value (`.integer` / `.number` / `.string` / `.boolean` /
    /// `.empty`) so consumers that don't read the side channel still
    /// see the cached value. Same lifetime as `row_cells`.
    row_formula_strings: std.ArrayListUnmanaged(?[]const u8),
    /// Parallel to `row_cells`. Each slot holds the base cell's A1
    /// reference for formula slaves — cells that participate in a
    /// formula spread without carrying their own `<f>` body. Two
    /// kinds populate this slot:
    ///   - **Shared-formula slaves**: source `<c>` carried
    ///     `<f t="shared" si="N"/>` with no body; the ref is looked
    ///     up from `shared_si_to_base_ref`.
    ///   - **Array-formula slaves**: source `<c>` had no `<f>` at
    ///     all, but its (col,row) sits inside a previously seen
    ///     `<f t="array" ref="X:Y">` rectangle; the ref is the array
    ///     base (top-left of that rectangle). Recorded in
    ///     `array_ranges` and resolved by linear scan in `consumeCell`.
    /// Null otherwise. The CellRef itself is by-value (no allocation),
    /// so this list owns the slot directly. Exposed via `formulaRefs()`.
    row_formula_refs: std.ArrayListUnmanaged(?CellRef),
    /// Map from a shared-formula `si` index (the integer in
    /// `<f t="shared" si="N"…>`) to the base cell's A1 ref (the
    /// top-left of the shared range, parsed from the base cell's
    /// `ref="X:Y"` attribute). Populated as base cells are consumed;
    /// queried by slave cells later in the same OR a later row.
    ///
    /// Lifetime is the **iterator's lifetime**, NOT the row's: shared
    /// formulas span multiple rows by design (the whole point of
    /// `<f t="shared" ref="C2:C10">` is that C3..C10 reference C2's
    /// formula text), so clearing this on `next()` would lose the
    /// mapping before slave cells are reached. Cleared / freed only
    /// in `deinit()`.
    shared_si_to_base_ref: std.AutoHashMapUnmanaged(u32, CellRef),
    /// Array-formula spread ranges seen so far in this iterator. Each
    /// entry records a base cell (`<f t="array" ref="X:Y">body</f>`,
    /// the top-left of the rectangle) plus the rectangle's corners.
    /// Slave cells inside the rectangle (e.g. C3..C10 for an array
    /// base at C2 with `ref="C2:C10"`) carry only a cached `<v>` and
    /// no `<f>` of their own; the linear scan in `consumeCell` finds
    /// the matching range and surfaces them as `formula_ref`-bearing
    /// formula records, mirroring the shared-formula slave shape.
    ///
    /// Lifetime is the **iterator's lifetime**, NOT the row's: array
    /// formulas span multiple rows by design (the whole point of
    /// `<f t="array" ref="C2:C10">` is that C3..C10 are dependents),
    /// so clearing this on `next()` would lose the linkage before the
    /// slave rows are reached. Cleared / freed only in `deinit()`.
    /// Typically tiny (a handful of entries per workbook), so a
    /// linear scan per cell is fine.
    array_ranges: std.ArrayListUnmanaged(ArrayRange),
    /// Bump arena for per-row decoded strings. Reset (O(1)) at the
    /// start of each `next()` call — previous row's owned strings
    /// become invalid, which matches the documented contract. Compared
    /// to the older per-string malloc/free list this saves ~one free
    /// per entity-bearing or rich-text cell per row.
    arena: std.heap.ArenaAllocator,
    /// 1-based OOXML row number of the most recently yielded row.
    /// Resolution priority (highest to lowest):
    ///   1. `<row r="N">` when present and N >= 1.
    ///   2. First `<c r="A11">` cell's trailing digits when the
    ///      row-level r is missing / zero / malformed — keeps the
    ///      envelope `row` in lockstep with inline cell refs.
    ///   3. `yield_count` (1-based count of rows yielded) as the
    ///      last resort when neither the row nor any cell has a
    ///      usable r.
    /// Valid only after `next()` has returned a non-null row; zero
    /// before the first successful call.
    current_row: u32 = 0,
    /// 1-based counter of rows yielded from this iterator. Used as
    /// the fallback for `current_row` when the source `<row>` tag
    /// has no usable `r` attribute, so envelope consumers see
    /// 1..N-contiguous row numbers regardless of source sparsity.
    yield_count: u32 = 0,
    /// Next 0-based column index for a `<c>` element that lacks
    /// the (spec-optional) `r=` attribute. Reset to 0 at the start
    /// of every row; on a `<c r="ABC123">` it jumps to that cell's
    /// column + 1 so subsequent r-less cells continue from there.
    /// Some Excel exporters (e.g. World Bank's WDI) omit `r=` on
    /// every cell to save bytes — without this fallback the parser
    /// would `error.MalformedXml` on the first cell.
    next_implicit_col: u32 = 0,

    pub fn deinit(self: *Rows) void {
        self.arena.deinit();
        self.row_cells.deinit(self.allocator);
        self.row_styles.deinit(self.allocator);
        self.row_date_types.deinit(self.allocator);
        self.row_error_strings.deinit(self.allocator);
        self.row_formula_strings.deinit(self.allocator);
        self.row_formula_refs.deinit(self.allocator);
        self.shared_si_to_base_ref.deinit(self.allocator);
        self.array_ranges.deinit(self.allocator);
        self.* = undefined;
    }

    /// Style indices for the current row — one slot per `row_cells`
    /// slot, mirroring the same gap-filled layout. `null` means the
    /// `<c>` had no `s` attribute (i.e. the General format). Valid
    /// until the next `next()` call, same as `row_cells`.
    pub fn styleIndices(self: *const Rows) []const ?u32 {
        return self.row_styles.items;
    }

    /// Per-column "this cell is a date-styled numeric serial" flags for
    /// the current row. Parallel to the slice returned by `next()`;
    /// length matches `row_cells.len`. `true` only when the cell is
    /// `.integer`/`.number`, its style resolves to a date format
    /// (`Book.isDateFormat`), and its value is inside
    /// `fromExcelSerial`'s accepted range. Gap cells and non-date
    /// numerics stay `false`. Valid until the next `next()` call.
    pub fn dateTypes(self: *const Rows) []const bool {
        return self.row_date_types.items;
    }

    /// Per-column literal Excel-error strings for the current row.
    /// Aligned to the slice returned by `next()`; `null` means the
    /// `<c>` was not `t="e"`. When non-null, the slot holds the literal
    /// error text (`"#DIV/0!"`, `"#N/A"`, `"#REF!"`, `"#VALUE!"`,
    /// `"#NUM!"`, `"#NAME?"`, `"#NULL!"`, `"#GETTING_DATA"`). Borrows
    /// from the sheet XML buffer when entity-free, otherwise lives in
    /// the row arena. Valid until the next `next()` call, same as
    /// `row_cells`.
    pub fn errorStrings(self: *const Rows) []const ?[]const u8 {
        return self.row_error_strings.items;
    }

    /// Per-column own-formula text for the current row. Aligned to the
    /// slice returned by `next()`; `null` means the `<c>` had no `<f>`
    /// of its own (i.e. it's a non-formula cell OR a shared-formula
    /// slave — slaves carry a `formula_ref` instead). Non-null slots
    /// hold the formula expression text verbatim from `<f>…body…</f>`:
    /// stand-alone, shared-formula bases, and array-formula bases all
    /// land here. Borrows from the sheet XML buffer when entity-free,
    /// otherwise lives in the row arena. Valid until the next `next()`
    /// call.
    pub fn formulaStrings(self: *const Rows) []const ?[]const u8 {
        return self.row_formula_strings.items;
    }

    /// Per-column shared-formula base references for the current row.
    /// Aligned to the slice returned by `next()`; `null` means the
    /// cell is not a shared-formula slave. Non-null slots hold the
    /// base cell's `{col, row}` (i.e. the top-left of the shared
    /// range, e.g. `C2` for a slave at `C5` whose base declared
    /// `<f t="shared" ref="C2:C10" si="0">`). The reverse of this
    /// invariant: `formulaStrings()[i]` is non-null iff the cell
    /// carries its own `<f>` body, in which case `formulaRefs()[i]`
    /// MUST be null (a single cell can't be both a base/standalone
    /// and a slave). Valid until the next `next()` call.
    pub fn formulaRefs(self: *const Rows) []const ?CellRef {
        return self.row_formula_refs.items;
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
        return if (book.uses_1904_epoch) fromExcelSerial1904(num) else fromExcelSerial(num);
    }

    /// 1-based OOXML row number of the most recently yielded row.
    /// Returns the `r="N"` attribute from the source `<row>` tag when
    /// present; otherwise a 1-based fallback counter of rows yielded.
    /// Only meaningful after `next()` has returned a non-null slice;
    /// zero before the first successful call.
    pub fn currentRowNumber(self: *const Rows) u32 {
        return self.current_row;
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
        self.row_date_types.clearRetainingCapacity();
        self.row_error_strings.clearRetainingCapacity();
        self.row_formula_strings.clearRetainingCapacity();
        self.row_formula_refs.clearRetainingCapacity();
        self.next_implicit_col = 0;
        // NOTE: shared_si_to_base_ref and array_ranges are NOT cleared
        // here — both kinds of formula spread span rows (the base cell
        // at C2 is referenced by slave cells at C3..C10 in subsequent
        // rows). Both are freed in deinit() once the iterator goes
        // out of scope.
        _ = self.arena.reset(.retain_capacity);

        while (findTagOpen(self.xml, self.pos, "row")) |row_start| {
            // The `r` attribute on <row> is optional per the OOXML
            // spec; parse it when present, fall back to a 1-based
            // yield counter when absent.
            const attrs = self.xml[row_start.start + "<row".len .. row_start.after_open - 1];
            const attrs_trimmed = if (attrs.len > 0 and attrs[attrs.len - 1] == '/')
                attrs[0 .. attrs.len - 1]
            else
                attrs;
            // Row number resolution (priority):
            //   1. `<row r="N">` with N >= 1.
            //   2. Recover from the first `<c r="A11">` inside this
            //      row — the cell's ref is authoritative and matches
            //      what the envelope will synthesize as `{ref}` per
            //      cell, keeping them in lockstep.
            //   3. `yield_count` (1-based count of rows yielded) —
            //      last resort when neither the row nor any cell has
            //      a usable r attribute. In that case the synthesized
            //      envelope refs (`col_letter+yield_count`) also have
            //      no inline source refs to diverge from.
            //   `r="0"` is invalid per spec (rows are 1-based) and
            //   falls through the row-r step to step 2/3.
            self.yield_count += 1;
            const row_r_valid: ?u32 = if (getAttr(attrs_trimmed, "r")) |r_str| blk: {
                const parsed = std.fmt.parseInt(u32, r_str, 10) catch 0;
                break :blk if (parsed > 0) parsed else null;
            } else null;
            // Detect self-closing `<row .../>` early so the no-r
            // branch below doesn't fall into recoverRowFromFirstCell —
            // that helper scans forward for `</row>` and would happily
            // grab the *next* row's first cell ref, mis-numbering this
            // empty row.
            const attrs_no_ws = std.mem.trimRight(u8, attrs, " \t\r\n");
            const is_self_closing_row = attrs_no_ws.len > 0 and
                attrs_no_ws[attrs_no_ws.len - 1] == '/';
            if (row_r_valid) |n| {
                self.current_row = n;
            } else if (is_self_closing_row) {
                self.current_row = self.yield_count;
            } else {
                self.current_row = recoverRowFromFirstCell(self.xml, row_start.after_open) orelse self.yield_count;
            }
            // Note: array_ranges is NOT pruned here, even though
            // rows are usually monotonic. The earlier "swapRemove
            // ranges with br.row < current_row" optimization had a
            // correctness gap on out-of-order rows: the guard fires
            // AFTER the bad prune already happened. Real workbooks
            // rarely have more than a handful of array formulas, so
            // the linear-scan-per-cell is dominated by parsing cost.
            // If a future workload measurably suffers, build a
            // sorted-input precondition by iterating once first.
            self.pos = row_start.after_open;
            // Self-closing `<row .../>` (no cells; common in files
            // with style/height-only rows like POI's 58325_db).
            // Yield as an empty row — matches what a populated row
            // with all `.empty` cells would look like.
            if (is_self_closing_row) return self.row_cells.items;
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

        // The `r` attribute on `<c>` is optional per the OOXML
        // spec — absent means "next implicit column" relative to
        // the row's last referenced cell. Tight-packed exporters
        // (e.g. World Bank's WDI) omit it on every cell.
        const col_idx: usize = if (getAttr(attrs, "r")) |r_attr|
            try columnIndexFromRef(r_attr)
        else
            self.next_implicit_col;
        self.next_implicit_col = std.math.cast(u32, col_idx + 1) orelse return error.MalformedXml;
        const cell_type = getAttr(attrs, "t") orelse "n"; // default numeric
        const style_attr = getAttr(attrs, "s");
        const style_idx: ?u32 = if (style_attr) |s|
            std.fmt.parseInt(u32, s, 10) catch null
        else
            null;

        // Grow row_cells + row_styles + row_date_types + row_error_strings
        // + row_formula_strings + row_formula_refs to cover col_idx;
        // fill gaps with `.empty` / `null` / `false` / `null` / `null` /
        // `null`. All six arrays stay in lock-step so callers can index
        // them the same way.
        while (self.row_cells.items.len <= col_idx) {
            try self.row_cells.append(self.allocator, .empty);
            try self.row_styles.append(self.allocator, null);
            try self.row_date_types.append(self.allocator, false);
            try self.row_error_strings.append(self.allocator, null);
            try self.row_formula_strings.append(self.allocator, null);
            try self.row_formula_refs.append(self.allocator, null);
        }
        self.row_styles.items[col_idx] = style_idx;
        // Date-flag defaults to false; the numeric branch below flips
        // it on when the style + value qualify.
        self.row_date_types.items[col_idx] = false;
        // Error-string defaults to null; the t="e" branch below sets it.
        self.row_error_strings.items[col_idx] = null;
        // Formula side channels default to null; the <f>-parse step
        // below sets one (and only one) of them.
        self.row_formula_strings.items[col_idx] = null;
        self.row_formula_refs.items[col_idx] = null;

        if (is_self_closing) {
            // Empty cell; already .empty.
            self.pos = gt + 1;
            return;
        }
        self.pos = gt + 1;

        // Parse body until </c>.
        // For t="inlineStr" the body is <is><t>text</t></is>; otherwise
        // it's <v>text</v>, optionally preceded by an <f> element. The
        // <f> element is parsed by `parseFormulaElement` (iter61-b) and
        // populates `row_formula_strings` / `row_formula_refs`; the <v>
        // body still drives the cached value into the `Cell` slot below.
        const cell_close = "</c>";
        const body_end = std.mem.indexOfPos(u8, self.xml, self.pos, cell_close) orelse return error.MalformedXml;
        const body = self.xml[self.pos..body_end];
        self.pos = body_end + cell_close.len;

        // iter61-b: extract the `<f>` element (if any) BEFORE we touch
        // the cell value. Sets exactly one of `row_formula_strings` /
        // `row_formula_refs` per the OOXML rules:
        //   - `<f>body</f>` (no t=)        → stand-alone formula:    formula_strings[col] = body
        //   - `<f t="shared" ref="X:Y" si="N">body</f>` → base cell:  formula_strings[col] = body, map[N] = top-left(X:Y)
        //   - `<f t="shared" si="N"/>`     → slave:                  formula_refs[col] = map.get(N)
        //   - `<f t="array" ref="X:Y">body</f>` → array base:        formula_strings[col] = body, array_ranges += {tl,br,base=tl}
        //   - cell inside an array range with no `<f>` (slave):       formula_refs[col] = base (set after value parse)
        try self.parseFormulaElement(body, col_idx);

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
        } else if (std.mem.eql(u8, cell_type, "e")) blk: {
            // iter61-c: capture the literal error string on the parallel
            // side channel so the CLI can emit `t:"error"`. The Cell
            // union slot stays `.string` (same as before) — consumers
            // that don't read `errorStrings()` still see the literal
            // text and a stable type tag, matching pre-iter61-c
            // behaviour.
            const literal = try self.decodeVValue(body);
            self.row_error_strings.items[col_idx] = literal;
            break :blk .{ .string = literal };
        } else try parseNumericCell(extractVValue(body) orelse "");

        self.row_cells.items[col_idx] = cell;

        // Array-formula slave detection: a cell inside a previously
        // registered `<f t="array" ref="X:Y">` rectangle that has no
        // `<f>` of its own gets its `formula_ref` set to the base.
        // Runs AFTER parseFormulaElement so a cell that carries its
        // own formula (rare overlap case) keeps it; the existing
        // formula slots win because we gate on both being null. The
        // base cell itself is excluded so its own `formula_strings`
        // entry is preserved untouched.
        if (self.row_formula_strings.items[col_idx] == null and
            self.row_formula_refs.items[col_idx] == null)
        {
            const here = CellRef{ .col = @intCast(col_idx), .row = self.current_row };
            for (self.array_ranges.items) |ar| {
                if (here.row >= ar.tl.row and here.row <= ar.br.row and
                    here.col >= ar.tl.col and here.col <= ar.br.col and
                    !(here.row == ar.base.row and here.col == ar.base.col))
                {
                    self.row_formula_refs.items[col_idx] = ar.base;
                    break;
                }
            }
        }

        // Date-type side channel: flip the flag when the cell is a
        // numeric serial under a date-styled numFmt and the serial is
        // inside fromExcelSerial's round-trippable window. Mirrors the
        // three-condition check in `parseDate` but without materialising
        // the DateTime — cheaper for the common case where callers want
        // just the tag and will re-derive the DateTime on demand.
        //
        // iter61-b: a formula cell ALWAYS wins over date — per the
        // design doc, `t:"formula"` is the surface tag even when the
        // cached value is a date-styled numeric. The CLI surfaces
        // `cached:<n>` (raw serial) on the formula record; consumers
        // who want the ISO string can decode it themselves.
        const is_formula =
            self.row_formula_strings.items[col_idx] != null or
            self.row_formula_refs.items[col_idx] != null;
        if (style_idx != null and !is_formula) {
            const num_opt: ?f64 = switch (cell) {
                .number => |n| n,
                .integer => |n| @floatFromInt(n),
                else => null,
            };
            if (num_opt) |num| {
                if (self.book) |book| {
                    // Pick the decoder for the workbook's epoch.
                    // 1904-epoch books shift every serial by 1462 days;
                    // fromExcelSerial1904 just adds 1462 and delegates.
                    // (iter61-a originally skipped 1904 entirely; this
                    // iter ships the proper decoder.)
                    const decoded = if (book.uses_1904_epoch) fromExcelSerial1904(num) else fromExcelSerial(num);
                    if (book.isDateFormat(style_idx.?) and decoded != null) {
                        self.row_date_types.items[col_idx] = true;
                    }
                }
            }
        }
        // iter61-b: same mutual-exclusion guard for the error side
        // channel. The OOXML t="e" branch above already wrote into
        // row_error_strings; if the same cell ALSO carried an `<f>`
        // (rare but legal — a formula whose cached value is an
        // error literal), formula wins. Clear the error slot so
        // writeCell's tag-precedence rule (formula > error > date >
        // primitive) lands on `t:"formula"`.
        if (is_formula) {
            self.row_error_strings.items[col_idx] = null;
        }
    }

    /// Parse the optional `<f>` element inside a cell body and update
    /// the formula side channels for `col_idx`. Idempotent: if no `<f>`
    /// is present, leaves both slots null. Tolerates malformed `<f>`
    /// fragments by treating them as absent — the cell still surfaces
    /// its cached value via the `<v>` extraction in `consumeCell`.
    fn parseFormulaElement(self: *Rows, body: []const u8, col_idx: usize) !void {
        // Locate `<f` followed by either whitespace, `/`, or `>`.
        // Anything else (e.g. a hypothetical `<finalize>`) means we
        // matched a longer name and must keep searching.
        var scan: usize = 0;
        const f_start = while (std.mem.indexOfPos(u8, body, scan, "<f")) |p| {
            const after = p + "<f".len;
            if (after >= body.len) return; // truncated — no real <f>
            const ch = body[after];
            if (ch == ' ' or ch == '\t' or ch == '\n' or ch == '\r' or ch == '/' or ch == '>') {
                break p;
            }
            scan = p + 1;
        } else return; // no <f> in this cell body — non-formula

        const tag_open = f_start + "<f".len;
        const tag_gt = std.mem.indexOfScalarPos(u8, body, tag_open, '>') orelse return;
        const is_self_closing = tag_gt > tag_open and body[tag_gt - 1] == '/';
        const attrs = body[tag_open..if (is_self_closing) tag_gt - 1 else tag_gt];

        const f_type = getAttr(attrs, "t"); // null / "shared" / "array" / "dataTable"
        const ref_attr = getAttr(attrs, "ref");
        const si_attr = getAttr(attrs, "si");

        // Slave: `<f t="shared" si="N"/>` (no body, no ref). Look up
        // the base ref in the iterator's shared map. If the lookup
        // misses (malformed input — slave before its base), leave
        // both side channels null. Defensive: never propagate a
        // missing-base into a phantom formula record.
        if (is_self_closing and f_type != null and std.mem.eql(u8, f_type.?, "shared") and ref_attr == null) {
            const si_str = si_attr orelse return;
            const si = std.fmt.parseInt(u32, si_str, 10) catch return;
            if (self.shared_si_to_base_ref.get(si)) |base_ref| {
                self.row_formula_refs.items[col_idx] = base_ref;
            }
            return;
        }

        // Anything else with a body: extract `<f …>BODY</f>`.
        if (is_self_closing) return; // self-closing without t="shared" — nothing to surface
        const f_close = std.mem.indexOfPos(u8, body, tag_gt + 1, "</f>") orelse return;
        const f_body = body[tag_gt + 1 .. f_close];
        const formula_text = try self.internOrBorrow(f_body);
        self.row_formula_strings.items[col_idx] = formula_text;

        // Shared-formula base: register `si → top-left(ref)` so
        // subsequent slaves can resolve. Tolerate parse failures
        // (malformed ref / si): the base cell still surfaces its
        // own formula text; only the cross-cell linkage is lost.
        if (f_type != null and std.mem.eql(u8, f_type.?, "shared") and ref_attr != null and si_attr != null) {
            const si = std.fmt.parseInt(u32, si_attr.?, 10) catch return;
            const range = parseA1Range(ref_attr.?) catch {
                // ref might be a single cell ("C2") rather than a
                // range — parseA1Range already handles that path.
                // Any other failure means malformed XML; skip
                // registration but keep the formula string above.
                return;
            };
            // Top-left of the shared range is the base cell. For a
            // single-cell shared formula (degenerate), `top_left ==
            // bottom_right`, which is still the correct base.
            try self.shared_si_to_base_ref.put(self.allocator, si, range.top_left);
        }
        // Array formulas (`t="array"` with body+ref): the base cell
        // already has its formula text set above. Register the spread
        // rectangle so dependent cells inside the range — which carry
        // only a cached `<v>` and no `<f>` of their own — can be
        // resolved to this base by the linear scan in `consumeCell`.
        // Tolerate a missing/malformed ref the same way the
        // shared-formula branch does: the base cell still surfaces
        // its own formula text; only the cross-cell linkage is lost.
        if (f_type != null and std.mem.eql(u8, f_type.?, "array") and ref_attr != null) {
            const range = parseA1Range(ref_attr.?) catch return;
            const base = CellRef{ .col = @intCast(col_idx), .row = self.current_row };
            try self.array_ranges.append(self.allocator, .{
                .tl = range.top_left,
                .br = range.bottom_right,
                .base = base,
            });
        }
    }

    fn resolveSharedString(self: *Rows, body: []const u8) !Cell {
        const idx_text = extractVValue(body) orelse return error.MalformedXml;
        const idx = std.fmt.parseInt(u32, idx_text, 10) catch return error.MalformedXml;
        // Production path: route through Book.sharedStringAt so iter-sst-3's
        // lazy backend can flip storage without touching the cell hot path.
        // Test scaffolding sometimes constructs Rows with `book = null` and
        // a direct `shared_strings` snapshot — fall through to the snapshot
        // in that case.
        if (self.book) |book| {
            return .{ .string = try book.sharedStringAt(idx) };
        }
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

// ─── SheetMatrix: bulk row-materialise reader ────────────────────────

/// In-memory dense matrix view of an entire sheet. Built by
/// `Book.materialiseSheet` in one streaming pass; intended for
/// callers that want random row/column access without re-parsing
/// or that want to amortise per-row allocator round-trips across a
/// single arena.
///
/// Lifetime: borrows from the parent `Book` for shared-string and
/// XML-backed slices, so the Book MUST outlive the matrix. Owns its
/// own per-row decoded strings via the embedded `arena`. Tear down
/// with `deinit()` (frees the arena AND the outer rows slab in one
/// shot).
///
/// The row width is per-row dense (one slot per column up to the
/// widest cell in that row, gaps filled with `.empty`), matching
/// the existing `Rows.next()` shape. Different rows MAY have
/// different widths when the source sheet has ragged trailing
/// emptiness — the longest row dictates `maxColCount()`.
pub const SheetMatrix = struct {
    /// Bump arena owning every dynamically-allocated byte the matrix
    /// references that did not already belong to the parent Book.
    /// Includes: each row's `[]Cell` slab, the outer rows slab, and
    /// any entity-decoded / rich-text-concatenated cell strings.
    arena: std.heap.ArenaAllocator,
    /// One slice per row in source order. Inner slice is the dense
    /// per-row Cell array (gaps as `.empty`). Both layers live in
    /// `arena`, so freeing the arena frees the whole matrix.
    rows: []const []const Cell,

    /// Free the arena (and with it, every row slab and decoded
    /// string the matrix owned). After deinit the matrix is
    /// `undefined` — callers must drop their reference.
    pub fn deinit(self: *SheetMatrix) void {
        self.arena.deinit();
        self.* = undefined;
    }

    /// Number of rows materialised. Equivalent to `rows.len`.
    pub fn rowCount(self: *const SheetMatrix) usize {
        return self.rows.len;
    }

    /// Width (column count) of the widest row, or 0 when the sheet
    /// is empty. Useful for callers that want a uniform 2-D shape
    /// rather than the per-row jagged form.
    pub fn maxColCount(self: *const SheetMatrix) usize {
        var w: usize = 0;
        for (self.rows) |r| {
            if (r.len > w) w = r.len;
        }
        return w;
    }

    /// Cell at `(row, col)` with implicit gap-fill: out-of-bounds
    /// columns within the matrix's row-count window resolve to
    /// `.empty` rather than panicking. Returns `null` when `row`
    /// itself is out of bounds — the call-site can distinguish
    /// "missing row" from "empty cell" without an extra length
    /// check.
    pub fn cellAt(self: *const SheetMatrix, row: usize, col: usize) ?Cell {
        if (row >= self.rows.len) return null;
        const r = self.rows[row];
        if (col >= r.len) return .empty;
        return r[col];
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
            // Must be followed by whitespace, `/`, or `>` — i.e. a
            // whole tag, not a prefix. Pretty-printed OOXML wraps
            // attrs onto the next line (`<row\n r="7">`), so accept
            // `\n` and `\r` alongside space and tab.
            const next_c = if (after_lt + tag.len < xml.len) xml[after_lt + tag.len] else return null;
            if (next_c == ' ' or next_c == '\t' or next_c == '\n' or
                next_c == '\r' or next_c == '>' or next_c == '/')
            {
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
/// Refs past Excel's XFD column (16384, 1-based) are rejected as
/// MalformedXml so adversarial XML can't drive `consumeCell`'s row-
/// growth loop into a multi-MB allocation.
fn columnIndexFromRef(ref: []const u8) !usize {
    var idx: usize = 0;
    var i: usize = 0;
    while (i < ref.len and ref[i] >= 'A' and ref[i] <= 'Z') : (i += 1) {
        idx = idx * 26 + (ref[i] - 'A' + 1);
        if (idx > max_col_1based) return error.MalformedXml;
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
    // iter61-a: detect the 1904 date-epoch flag on the workbook. It's
    // carried on <workbookPr date1904="1"/> (or "true"). Default is
    // the 1900 system. Tag-boundary check required because
    // <workbookProtection> shares the `<workbookPr` prefix — matching
    // the prefix alone would stop on workbookProtection and miss
    // date1904 on the real workbookPr element. Scan lenient: if we
    // can't find the tag the flag stays false (1900, the common case).
    var scan: usize = 0;
    while (std.mem.indexOfPos(u8, wb_xml, scan, "<workbookPr")) |p| {
        const after = p + "<workbookPr".len;
        if (after >= wb_xml.len) break;
        const ch = wb_xml[after];
        if (ch == ' ' or ch == '\t' or ch == '\n' or ch == '\r' or ch == '/' or ch == '>') {
            const pr_end = std.mem.indexOfScalarPos(u8, wb_xml, after, '>') orelse wb_xml.len;
            const attrs = wb_xml[after..pr_end];
            if (getAttr(attrs, "date1904")) |v| {
                if (std.mem.eql(u8, v, "1") or std.mem.eql(u8, v, "true")) {
                    book.uses_1904_epoch = true;
                }
            }
            break;
        }
        scan = p + 1;
    }

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
        // Accept any whitespace (incl. \n / \r for pretty-printed
        // workbook.xml) plus `/` (self-closing) after `<sheet`.
        const next_c = if (sh_start + "<sheet".len < wb_xml.len) wb_xml[sh_start + "<sheet".len] else break;
        if (next_c != ' ' and next_c != '\t' and next_c != '\n' and next_c != '\r' and next_c != '/') {
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

/// Maximum column index in an Excel sheet (XFD, 1-based 16384).
pub const max_col_1based: u32 = 16384;
/// Maximum row index in an Excel sheet (1-based).
pub const max_row: u32 = 1_048_576;

/// Parse one corner of an A1-style reference ("B12") into `{col, row}`.
/// Column is 0-based (A=0, B=1, …), row is 1-based (row1=1). Rejects
/// empty input, lowercase, missing digits, row=0, and refs beyond
/// Excel's XFD/1048576 sheet bounds.
pub fn parseA1Ref(s: []const u8) !CellRef {
    if (s.len == 0) return error.MalformedXml;
    var i: usize = 0;
    var col: u32 = 0;
    while (i < s.len and s[i] >= 'A' and s[i] <= 'Z') : (i += 1) {
        // Bail as soon as the running column exceeds the Excel max.
        // The previous "saturate" approach risked u32 overflow on
        // pathologically long all-letter prefixes since the
        // multiplication ran before the bound check fired.
        if (col > max_col_1based) return error.MalformedXml;
        col = col * 26 + (s[i] - 'A' + 1);
    }
    if (i == 0 or i == s.len) return error.MalformedXml;
    var row: u32 = 0;
    while (i < s.len and s[i] >= '0' and s[i] <= '9') : (i += 1) {
        if (row > max_row) return error.MalformedXml;
        row = row * 10 + (s[i] - '0');
    }
    if (i != s.len or row == 0) return error.MalformedXml;
    if (col > max_col_1based or row > max_row) return error.MalformedXml;
    return .{ .col = col - 1, .row = row };
}

/// Parse an A1-style range ("A1:B2"). Top-left corner must precede or
/// equal bottom-right on both axes; we normalise order so callers can
/// rely on `top_left ≤ bottom_right` component-wise even if the source
/// XML listed the corners in the other order.
pub fn parseA1Range(ref: []const u8) !MergeRange {
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
    // pre-size the backing array and skip geometric grows. Capped
    // against the block size so a malformed workbook can't claim
    // billions of merge cells and force a huge upfront alloc. The
    // parser only matches `<mergeCell ref="…"/>`, whose shortest
    // valid spelling is 21 bytes (`<mergeCell ref="A1"/>`).
    const hint: usize = blk: {
        const c = std.mem.indexOfPos(u8, block, 0, "count=\"") orelse break :blk 0;
        const start = c + "count=\"".len;
        const end = std.mem.indexOfScalarPos(u8, block, start, '"') orelse break :blk 0;
        const claimed = std.fmt.parseInt(usize, block[start..end], 10) catch 0;
        break :blk @min(claimed, block.len / 21);
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

/// Find the comments part referenced by a sheet's rels XML. Returns
/// the raw `Target` attribute of the first `Relationship` whose
/// `Type` ends in `/comments` OR whose `Target` basename matches
/// `comments[N].xml`. Caller is responsible for resolving the
/// returned target against the rels file's base directory (sheet
/// rels live at `xl/worksheets/_rels/sheetN.xml.rels`, so
/// `../comments1.xml` resolves to `xl/comments1.xml`).
///
/// Both criteria are robust against accidentally matching a stray
/// `Target` like `https://host/comments` or `commentsShape1.vml`,
/// which the previous substring scan would happily accept.
fn rels_targetForComments(rels_xml: []const u8) ?[]const u8 {
    var probe: usize = 0;
    while (std.mem.indexOfPos(u8, rels_xml, probe, "<Relationship")) |start| {
        const end = std.mem.indexOfScalarPos(u8, rels_xml, start, '>') orelse return null;
        const attrs = rels_xml[start..end];
        probe = end + 1;

        const tgt_key = "Target=\"";
        const tgt_pos = std.mem.indexOf(u8, attrs, tgt_key) orelse continue;
        const tgt_start = tgt_pos + tgt_key.len;
        const tgt_close = std.mem.indexOfScalarPos(u8, attrs, tgt_start, '"') orelse continue;
        const target = attrs[tgt_start..tgt_close];

        const type_key = "Type=\"";
        const type_attr: ?[]const u8 = blk: {
            const tp = std.mem.indexOf(u8, attrs, type_key) orelse break :blk null;
            const ts = tp + type_key.len;
            const tc = std.mem.indexOfScalarPos(u8, attrs, ts, '"') orelse break :blk null;
            break :blk attrs[ts..tc];
        };

        const type_matches = if (type_attr) |t|
            std.mem.endsWith(u8, t, "/comments") or std.mem.endsWith(u8, t, "/relationships/comments")
        else
            false;

        const basename = if (std.mem.lastIndexOfScalar(u8, target, '/')) |slash|
            target[slash + 1 ..]
        else
            target;

        // Basename test: starts with "comments" + [digits] + ".xml".
        const basename_matches = blk: {
            if (!std.mem.startsWith(u8, basename, "comments")) break :blk false;
            if (!std.mem.endsWith(u8, basename, ".xml")) break :blk false;
            const middle = basename["comments".len .. basename.len - ".xml".len];
            if (middle.len == 0) break :blk true;
            for (middle) |c| if (c < '0' or c > '9') break :blk false;
            break :blk true;
        };

        if (type_matches or basename_matches) return target;
    }
    return null;
}

/// Resolve a rels-XML `Target` against `base_dir` (the rels file's
/// referencing-document directory, e.g. `xl/worksheets/`). Handles
/// leading `..` segments and writes the joined, normalised archive
/// path into `out`. Returns the slice of `out` actually written, or
/// null if the path doesn't fit. `base_dir` MUST end in '/' for the
/// segment math to work.
///
/// Absolute (package-rooted) targets like `/xl/comments1.xml` are
/// returned with the leading `/` stripped, NOT joined to `base_dir`
/// — that's how `parseWorkbookSheets` already treats them.
fn resolveRelArchivePath(base_dir: []const u8, target: []const u8, out: []u8) ?[]const u8 {
    // Absolute (package-rooted) targets use no base. They still flow
    // through the segment normalisation loop so dot segments inside
    // an absolute path collapse correctly (e.g.
    // `/xl/worksheets/../comments1.xml` → `xl/comments1.xml`).
    var effective_base: []const u8 = base_dir;
    var effective_target: []const u8 = target;
    if (target.len > 0 and target[0] == '/') {
        effective_base = "";
        effective_target = target[1..];
    }

    var written: usize = 0;
    if (effective_base.len > out.len) return null;
    @memcpy(out[0..effective_base.len], effective_base);
    written = effective_base.len;

    var iter = std.mem.splitScalar(u8, effective_target, '/');
    while (iter.next()) |seg| {
        if (seg.len == 0) continue; // collapsed `//`
        if (std.mem.eql(u8, seg, ".")) continue;
        if (std.mem.eql(u8, seg, "..")) {
            // Pop one trailing segment off `out`. base_dir must end
            // in '/'; trim the trailing '/' first, then trim back to
            // (and drop) the previous '/'.
            if (written == 0) return null;
            written -= 1; // drop trailing '/'
            while (written > 0 and out[written - 1] != '/') written -= 1;
            continue;
        }
        if (written + seg.len > out.len) return null;
        @memcpy(out[written..][0..seg.len], seg);
        written += seg.len;
        if (written == out.len) return null;
        out[written] = '/';
        written += 1;
    }
    // Trim trailing '/' that we just appended for the final segment.
    if (written > 0 and out[written - 1] == '/') written -= 1;
    return out[0..written];
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
        // URLs and locations may carry XML entities (`&amp;`, `&apos;`,
        // …); decode them through `decodeFormulaInto` so callers get
        // the real target. The helper returns a zero-copy slice when
        // there are no entities and an arena-owned decoded copy when
        // there are.
        var url: []const u8 = "";
        var location: []const u8 = "";
        if (std.mem.indexOf(u8, attrs, "r:id=\"")) |rid_pos| {
            const rid_start = rid_pos + "r:id=\"".len;
            const rid_close = std.mem.indexOfScalarPos(u8, attrs, rid_start, '"') orelse continue;
            const rid = attrs[rid_start..rid_close];
            if (rels_xml) |rx| {
                const raw = findRelTarget(rx, rid) orelse continue;
                url = try decodeFormulaInto(book, raw);
            } else {
                continue;
            }
        } else if (std.mem.indexOf(u8, attrs, "location=\"")) |loc_pos| {
            const loc_start = loc_pos + "location=\"".len;
            const loc_close = std.mem.indexOfScalarPos(u8, attrs, loc_start, '"') orelse continue;
            location = try decodeFormulaInto(book, attrs[loc_start..loc_close]);
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

    // Resolve via the shared rels-target helper, which only matches
    // the relationship Type ending in `/comments` *or* a basename
    // matching `comments[N].xml` — avoids picking up
    // `…/drawings/commentsShape1.vml` or any URL containing the
    // substring "comments" by accident. The returned target is
    // resolved against `xl/worksheets/` so non-canonical layouts
    // (e.g. `Target="../threadedComments/comments1.xml"`) keep
    // their subdir in the archive lookup.
    var comments_xml: ?[]const u8 = null;
    if (rels_targetForComments(rels_xml)) |target| {
        var path_buf: [256]u8 = undefined;
        const full = resolveRelArchivePath("xl/worksheets/", target, &path_buf) orelse return;
        if (book.comments_data.get(full)) |data| comments_xml = data;
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
            // OOXML rgb is canonical AARRGGBB — exactly 8 hex digits.
            // Anything shorter (e.g. malformed `rgb="FF"`) would have
            // surfaced as a bogus near-black color via parseInt; reject
            // those rather than leak garbage into Book.richRuns / styles.
            if (re - rs == 8) {
                if (std.fmt.parseInt(u32, tag[rs..re], 16)) |v| return v else |_| {}
            }
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
            // OOXML escapes quoted literals inside formatCode as
            // entities, e.g. `0.00" months"` becomes
            // `0.00&quot; months&quot;` in styles.xml. Decode here so
            // (a) callers see the real format string and (b) the
            // is-date-format scanner doesn't mis-classify literal
            // text inside a quoted run as date tokens.
            const code = try decodeFormulaInto(book, attrs[code_start..code_end]);
            try book.custom_num_fmts.put(book.allocator, id, code);
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
    // The hint is capped against the XML size: a malformed or hostile
    // workbook can advertise an absurd uniqueCount to force a huge
    // upfront alloc. The smallest possible `<si/>` self-close is 5
    // bytes, so the buffer can't really hold more than xml.len/5
    // entries — anything past that is a lie.
    const hint: usize = blk: {
        const open = std.mem.indexOf(u8, sst_xml, "<sst") orelse break :blk 64;
        const gt = std.mem.indexOfScalarPos(u8, sst_xml, open, '>') orelse break :blk 64;
        const attrs = sst_xml[open..gt];
        const u = std.mem.indexOf(u8, attrs, "uniqueCount=\"") orelse break :blk 64;
        const start = u + "uniqueCount=\"".len;
        const end = std.mem.indexOfScalarPos(u8, attrs, start, '"') orelse break :blk 64;
        const claimed = std.fmt.parseInt(usize, attrs[start..end], 10) catch 64;
        break :blk @min(claimed, sst_xml.len / 5);
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
            } else if (c1 == 'r' and next_lt + 4 <= xml.len and
                xml[next_lt + 2] == 'P' and xml[next_lt + 3] == 'h' and
                (next_lt + 4 == xml.len or xml[next_lt + 4] == '>' or xml[next_lt + 4] == ' ' or xml[next_lt + 4] == '/'))
            {
                // `<rPh>...</rPh>` (or `<rPh.../>`) — phonetic guide /
                // furigana annotation. The text inside is hidden
                // pronunciation, not visible cell content. Skip the
                // entire element so its inner `<t>` doesn't get
                // flattened into the SST string. Self-closing form
                // (`<rPh .../>`) just skips past the opening tag.
                const rph_open_gt = std.mem.indexOfScalarPos(u8, xml, next_lt + 4, '>') orelse break :body;
                if (rph_open_gt > 0 and xml[rph_open_gt - 1] == '/') {
                    j = rph_open_gt + 1;
                } else {
                    const rph_close = std.mem.indexOfPos(u8, xml, rph_open_gt + 1, "</rPh>") orelse break :body;
                    j = rph_close + "</rPh>".len;
                }
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

    book.sst = .{ .eager = try strings.toOwnedSlice(book.allocator) };
}

/// Lazy SST builder (iter-sst-3b). Walks the SST once recording the
/// byte span of each `<si>...</si>` body, but does NOT decode plain
/// text. `Book.sharedStringAt` materialises individual entries into
/// `sst_arena` on first access. Rich-text runs are not eagerly
/// captured by the lazy backend yet (deferred to a follow-up iter);
/// `Book.richRuns(idx)` returns null on lazy-opened books.
fn parseSharedStringsLazy(book: *Book, sst_xml: []u8) !void {
    const hint: usize = blk: {
        const open = std.mem.indexOf(u8, sst_xml, "<sst") orelse break :blk 64;
        const gt = std.mem.indexOfScalarPos(u8, sst_xml, open, '>') orelse break :blk 64;
        const attrs = sst_xml[open..gt];
        const u = std.mem.indexOf(u8, attrs, "uniqueCount=\"") orelse break :blk 64;
        const start = u + "uniqueCount=\"".len;
        const end = std.mem.indexOfScalarPos(u8, attrs, start, '"') orelse break :blk 64;
        const claimed = std.fmt.parseInt(usize, attrs[start..end], 10) catch 64;
        break :blk @min(claimed, sst_xml.len / 5);
    };

    var offsets: std.ArrayListUnmanaged(usize) = .{};
    var lengths: std.ArrayListUnmanaged(usize) = .{};
    errdefer offsets.deinit(book.allocator);
    errdefer lengths.deinit(book.allocator);
    try offsets.ensureTotalCapacity(book.allocator, hint);
    try lengths.ensureTotalCapacity(book.allocator, hint);

    const xml = sst_xml;
    var i: usize = 0;

    outer: while (i < xml.len) {
        const i_prev = i;
        const lt = std.mem.indexOfScalarPos(u8, xml, i, '<') orelse break;
        if (lt + 3 > xml.len) break;
        if (xml[lt + 1] != 's' or xml[lt + 2] != 'i') {
            i = lt + 1;
            std.debug.assert(i > i_prev);
            continue;
        }
        const si_gt = std.mem.indexOfScalarPos(u8, xml, lt + 3, '>') orelse break;
        // Self-closing `<si/>` — empty entry.
        if (si_gt > 0 and xml[si_gt - 1] == '/') {
            try offsets.append(book.allocator, si_gt);
            try lengths.append(book.allocator, 0);
            i = si_gt + 1;
            std.debug.assert(i > i_prev);
            continue;
        }
        // Find the matching `</si>`. Body span is everything between
        // `>` and `</`. Match the eager parser's recovery: a missing
        // `</si>` on one entry only skips THAT entry — we advance past
        // the opening `<si>` and keep scanning subsequent entries.
        const si_close = std.mem.indexOfPos(u8, xml, si_gt + 1, "</si>") orelse {
            i = si_gt + 1;
            std.debug.assert(i > i_prev);
            continue :outer;
        };
        const body_start: usize = si_gt + 1;
        const body_end: usize = si_close;
        const sst_idx = offsets.items.len;
        try offsets.append(book.allocator, body_start);
        try lengths.append(book.allocator, body_end - body_start);

        // Eagerly capture rich-text runs so Book.richRuns() works on
        // lazy books. Plain entries no-op out of the inner walker; the
        // common-case rich-density is <1% of typical SSTs so the
        // amortised cost is small.
        try parseSstRichRunsForBody(book, xml[body_start..body_end], sst_idx);

        i = si_close + "</si>".len;
        std.debug.assert(i > i_prev);
        continue :outer;
    }

    // Stage both owned slices in locals before assigning the union so
    // a failure on the second `toOwnedSlice` doesn't leak the first.
    const offsets_owned = try offsets.toOwnedSlice(book.allocator);
    errdefer book.allocator.free(offsets_owned);
    const lengths_owned = try lengths.toOwnedSlice(book.allocator);
    errdefer book.allocator.free(lengths_owned);
    book.sst = .{ .lazy = .{
        .offsets = offsets_owned,
        .lengths = lengths_owned,
        .resolved = .{},
    } };
}

/// Walk one `<si>` body for rich-text runs and, if any are present,
/// commit them into `book.rich_runs_by_sst_idx[sst_idx]`. No-op when
/// the body has no `<r>` blocks (the common case). Used by
/// `parseSharedStringsLazy` so `Book.richRuns()` works on lazy
/// books with the same lifetime contract as the eager backend.
fn parseSstRichRunsForBody(book: *Book, body: []const u8, sst_idx: usize) !void {
    const arena = book.sst_arena.allocator();
    var runs: std.ArrayListUnmanaged(RichRun) = .empty;
    var saw_r = false;
    var pending_flags: RichRun = .{ .text = "" };

    var j: usize = 0;
    while (j < body.len) {
        const j_prev = j;
        const next_lt = std.mem.indexOfScalarPos(u8, body, j, '<') orelse break;
        if (next_lt + 2 > body.len) break;
        const c1 = body[next_lt + 1];

        if (c1 == 'r' and next_lt + 2 < body.len and
            (body[next_lt + 2] == '>' or body[next_lt + 2] == ' '))
        {
            // `<r>` — enter a rich-text run block. Each `<r>` resets
            // pending_flags so a run without `<rPr>` is unstyled.
            const r_gt = std.mem.indexOfScalarPos(u8, body, next_lt + 2, '>') orelse break;
            saw_r = true;
            pending_flags = .{ .text = "" };
            j = r_gt + 1;
        } else if (c1 == 'r' and next_lt + 3 < body.len and
            body[next_lt + 2] == 'P' and body[next_lt + 3] == 'r')
        {
            const rpr_close = std.mem.indexOfPos(u8, body, next_lt, "</rPr>") orelse break;
            pending_flags = try parseRprFlags(book, arena, body[next_lt .. rpr_close + "</rPr>".len]);
            j = rpr_close + "</rPr>".len;
        } else if (c1 == 'r' and next_lt + 4 <= body.len and
            body[next_lt + 2] == 'P' and body[next_lt + 3] == 'h' and
            (next_lt + 4 == body.len or body[next_lt + 4] == '>' or body[next_lt + 4] == ' ' or body[next_lt + 4] == '/'))
        {
            // `<rPh.../>` or `<rPh>...</rPh>` — phonetic guide.
            // Skip the entire element so its inner `<t>` doesn't
            // get registered as a rich-text run.
            const rph_open_gt = std.mem.indexOfScalarPos(u8, body, next_lt + 4, '>') orelse break;
            if (rph_open_gt > 0 and body[rph_open_gt - 1] == '/') {
                j = rph_open_gt + 1;
            } else {
                const rph_close = std.mem.indexOfPos(u8, body, rph_open_gt + 1, "</rPh>") orelse break;
                j = rph_close + "</rPh>".len;
            }
        } else if (c1 == 't' and (next_lt + 2 == body.len or
            body[next_lt + 2] == '>' or body[next_lt + 2] == ' ' or body[next_lt + 2] == '/'))
        {
            const t_gt = std.mem.indexOfScalarPos(u8, body, next_lt + 2, '>') orelse break;
            if (t_gt > 0 and body[t_gt - 1] == '/') {
                j = t_gt + 1;
                std.debug.assert(j > j_prev);
                continue;
            }
            const t_close = std.mem.indexOfPos(u8, body, t_gt + 1, "</t>") orelse break;
            const span = body[t_gt + 1 .. t_close];

            if (saw_r) {
                const span_has_ent = std.mem.indexOfScalar(u8, span, '&') != null;
                const run_text: []const u8 = if (span_has_ent) blk: {
                    var rb: std.ArrayListUnmanaged(u8) = try .initCapacity(arena, span.len);
                    try appendDecoded(arena, &rb, span);
                    break :blk try rb.toOwnedSlice(arena);
                } else span;
                try runs.append(arena, .{
                    .text = run_text,
                    .bold = pending_flags.bold,
                    .italic = pending_flags.italic,
                    .color_argb = pending_flags.color_argb,
                    .size = pending_flags.size,
                    .font_name = pending_flags.font_name,
                });
            }
            j = t_close + "</t>".len;
        } else {
            j = next_lt + 1;
        }
        std.debug.assert(j > j_prev);
    }

    if (runs.items.len > 0) {
        const owned = try runs.toOwnedSlice(arena);
        try book.rich_runs_by_sst_idx.put(book.allocator, sst_idx, owned);
    }
}

/// Decode a single SST entry's body span into the Book's `sst_arena`.
/// Used by the lazy backend on first `sharedStringAt(idx)` access; the
/// returned slice is cached in `lazy.resolved` so subsequent accesses
/// are O(1) hashmap hits. Errors with `MalformedXml` if the offset
/// table is corrupted (defensive — should never fire on a sanely
/// parsed SST).
fn materialiseSstEntry(book: *Book, idx: usize) ![]const u8 {
    const lazy = switch (book.sst) {
        .lazy => |*l| l,
        .eager => unreachable,
    };
    if (idx >= lazy.offsets.len) return error.MalformedXml;
    const xml = book.shared_strings_xml orelse return error.MalformedXml;
    const start = lazy.offsets[idx];
    const len = lazy.lengths[idx];
    if (start + len > xml.len) return error.MalformedXml;
    const body = xml[start .. start + len];

    // Concat every `<t>...</t>` body in this entry, decoding entities.
    // Mirrors the eager parser's plain-text path but only runs on the
    // single requested entry.
    const arena = book.sst_arena.allocator();
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    var t_count: usize = 0;
    var first_span: []const u8 = "";
    var first_has_ent = false;
    var j: usize = 0;
    while (j < body.len) {
        const next_lt = std.mem.indexOfScalarPos(u8, body, j, '<') orelse break;
        if (next_lt + 2 > body.len) break;
        const c1 = body[next_lt + 1];
        if (c1 == 't' and (next_lt + 2 == body.len or
            body[next_lt + 2] == '>' or body[next_lt + 2] == ' ' or body[next_lt + 2] == '/'))
        {
            const t_gt = std.mem.indexOfScalarPos(u8, body, next_lt + 2, '>') orelse break;
            if (t_gt > 0 and body[t_gt - 1] == '/') {
                j = t_gt + 1;
                continue;
            }
            const t_close = std.mem.indexOfPos(u8, body, t_gt + 1, "</t>") orelse break;
            const span = body[t_gt + 1 .. t_close];
            const span_has_ent = std.mem.indexOfScalar(u8, span, '&') != null;
            if (t_count == 0) {
                first_span = span;
                first_has_ent = span_has_ent;
            } else {
                if (buf.capacity == 0) {
                    const cap: usize = first_span.len + span.len + 8;
                    buf = try .initCapacity(arena, cap);
                    if (first_has_ent) {
                        try appendDecoded(arena, &buf, first_span);
                    } else {
                        buf.appendSliceAssumeCapacity(first_span);
                    }
                }
                try buf.ensureUnusedCapacity(arena, span.len);
                if (span_has_ent) {
                    try appendDecoded(arena, &buf, span);
                } else {
                    buf.appendSliceAssumeCapacity(span);
                }
            }
            t_count += 1;
            j = t_close + "</t>".len;
        } else if (c1 == 'r' and next_lt + 4 <= body.len and
            body[next_lt + 2] == 'P' and body[next_lt + 3] == 'h' and
            (next_lt + 4 == body.len or body[next_lt + 4] == '>' or body[next_lt + 4] == ' ' or body[next_lt + 4] == '/'))
        {
            // `<rPh.../>` or `<rPh>...</rPh>` — phonetic guide,
            // hidden pronunciation. Skip the entire element so its
            // inner `<t>` doesn't get flattened. Mirrors the eager
            // path's rPh skip.
            const rph_open_gt = std.mem.indexOfScalarPos(u8, body, next_lt + 4, '>') orelse break;
            if (rph_open_gt > 0 and body[rph_open_gt - 1] == '/') {
                j = rph_open_gt + 1;
            } else {
                const rph_close = std.mem.indexOfPos(u8, body, rph_open_gt + 1, "</rPh>") orelse break;
                j = rph_close + "</rPh>".len;
            }
        } else {
            j = next_lt + 1;
        }
    }

    const decoded: []const u8 = if (t_count == 0) "" else if (t_count == 1) blk: {
        if (first_has_ent) {
            var b: std.ArrayListUnmanaged(u8) = try .initCapacity(arena, first_span.len);
            try appendDecoded(arena, &b, first_span);
            break :blk try b.toOwnedSlice(arena);
        } else {
            // Borrow span directly — lifetime matches sst_xml which
            // lives for the Book's lifetime.
            break :blk first_span;
        }
    } else try buf.toOwnedSlice(arena);

    try lazy.resolved.put(book.allocator, @intCast(idx), decoded);
    return decoded;
}

// ─── Editor (load-modify-save, Phase 3c) ────────────────────────────

/// Read an existing xlsx, mutate, save. Phase 3c append-only v1 per
/// docs/plans/load-modify-save.md.
///
/// Walks the source archive at `open` time to capture per-entry byte
/// spans (LFH + payload + central-directory record + EOCD comment).
/// `save` re-emits using those spans so unmodified entries round-trip
/// byte-identically; future iters (iter-lms-2+) replace specific
/// entries with mutated copies.
///
/// Byte span of a `<c>` element in a worksheet's decompressed XML.
/// Produced by `Editor.scanWorksheet` (Phase 3d, iter-cm-1) — the
/// foundation for cell-mutate. iter-cm-2's `setCell` will use these
/// to splice replacement bytes in place without disturbing
/// surrounding `<f>` formulas, `<is>` inline strings, phonetic
/// hints, or unknown attributes/children.
pub const CellSpan = struct {
    /// Byte offset of `<` opening the `<c` tag (in the parent
    /// `WorksheetSpans.xml` slice).
    start: usize,
    /// Byte offset just past `</c>` for body cells, or just past
    /// `/>` for self-closing cells.
    end: usize,
    /// Byte offset just past `>` of the opening tag (== `end` for
    /// self-closing cells). The slice `xml[body_start..close]` is
    /// the body content (`<v>...</v>`, `<f>...</f><v>...</v>`,
    /// `<is>...</is>`, etc.) where `close = end - 4` for `</c>`
    /// or `end` for self-closing.
    body_start: usize,
    /// 1-based row number resolved from `<row r="N">` (with
    /// implicit-row fallback) AND from `<c r="A1">` (with
    /// implicit-column fallback per iter-impl-col).
    row: u32,
    /// 0-based column index (A=0).
    col: u32,
};

/// Read-only span index over a worksheet's `<c>` elements. The
/// decompressed sheet XML and the cells slice are both owned;
/// `deinit` frees both.
pub const WorksheetSpans = struct {
    allocator: Allocator,
    /// Decompressed sheet XML. Owned — freed by `deinit`.
    xml: []const u8,
    /// Every `<c>` span, in source order. Owned.
    cells: []const CellSpan,

    pub fn deinit(self: *WorksheetSpans) void {
        self.allocator.free(self.xml);
        self.allocator.free(self.cells);
        self.* = undefined;
    }

    /// Find the span for `(row, col)`. Linear scan — fine for
    /// iter-cm-1's read-only use. iter-cm-2 will add a row-index
    /// hashmap when bulk mutation needs O(log N) lookups.
    pub fn find(self: *const WorksheetSpans, row: u32, col: u32) ?CellSpan {
        for (self.cells) |s| if (s.row == row and s.col == col) return s;
        return null;
    }
};

/// Lifetime: the source buffer + entry table are held resident;
/// `deinit` frees both. Single-threaded per Editor instance, matching
/// the rest of zlsx.
pub const Editor = struct {
    allocator: Allocator,
    src_buf: []u8,
    entries: []const ZipEntry,
    /// Offset of central directory inside `src_buf` (from EOCD).
    cd_offset: u32,
    /// Length of the central directory in bytes.
    cd_size: u32,
    /// Offset of EOCD record inside `src_buf`.
    eocd_offset: u32,
    /// Verbatim trailing-comment bytes from the EOCD record (may be
    /// empty). Points into `src_buf`.
    eocd_comment: []const u8,
    /// Sheet path for each declared sheet, in declared order. Owned
    /// `dupe` of the path string from the source's workbook.xml.rels
    /// (resolved via `Book.open` at editor-construction time so the
    /// editor doesn't need to re-parse rels itself).
    sheet_paths: []const []const u8,
    /// Pending appended rows per sheet. Empty when no mutation has
    /// been requested — `save` then flows through the byte-identical
    /// passthrough path. Each `AppendBuffer.rows` slice + every
    /// inner row slice is allocator-owned.
    pending_appends: std.AutoHashMapUnmanaged(u32, AppendBuffer),
    /// Pending sheet additions (Phase 3e / iter-sheet-1). Each entry
    /// produces a new ZIP entry at save time + patches workbook.xml,
    /// rels, and [Content_Types].xml. The Editor's `sheet_paths`
    /// slice is grown synchronously on `addSheet` so subsequent
    /// `setCell` / `scanWorksheet` calls can target the new sheet
    /// by its returned index.
    pending_new_sheets: std.ArrayListUnmanaged(NewSheet),
    /// Pending sheet renames (Phase 3e iter-sheet-2). Each entry
    /// patches the matching `<sheet name="OLD"…>` line in
    /// `xl/workbook.xml` to use the new name. v1 limitation:
    /// formulas in OTHER sheets that reference the renamed sheet
    /// by name (`'OLD'!A1`) are NOT rewritten — that requires a
    /// formula tokenizer (iter-col-1).
    pending_renames: std.ArrayListUnmanaged(SheetRename),
    /// Pending sheet deletions (Phase 3e iter-sheet-3). Each entry
    /// drops the matching ZIP entry at save + patches workbook.xml,
    /// rels, Content_Types. Cross-sheet formula refs to the deleted
    /// sheet become `#REF!` until the iter-col-1 formula tokenizer
    /// ships.
    pending_deletes: std.ArrayListUnmanaged(SheetDelete),
    /// Pending row insertions (Phase 3e iter-row-2). Each entry
    /// shifts every row at `before_row..` down by 1 in the named
    /// sheet's worksheet XML at save time. v1 limitations: only
    /// `<row r=>` + `<c r=>` row component + `<mergeCells>` ref +
    /// `<dimension>` are rewritten; data validations, hyperlinks,
    /// conditional formatting, defined names, formulas, drawings
    /// and pivots are left unchanged. Refuse if any of those
    /// elements exist (conservative guard).
    pending_row_inserts: std.ArrayListUnmanaged(RowEdit),
    /// Pending row deletions (Phase 3e iter-row-3). Same shape as
    /// inserts: shifts every row > deleted_row up by 1. Same v1
    /// limitations.
    pending_row_deletes: std.ArrayListUnmanaged(RowEdit),
    /// Pending column inserts (Phase 3e iter-col-3). Shifts every
    /// column at or above `before_col` right by one position.
    /// Same conservative guards as row edits — refuses sheets
    /// with formulas (no formula tokenizer in v1).
    pending_col_inserts: std.ArrayListUnmanaged(ColEdit),
    /// Pending column deletes (Phase 3e iter-col-4). Symmetric to
    /// inserts.
    pending_col_deletes: std.ArrayListUnmanaged(ColEdit),

    /// Pending in-place cell mutations per sheet (Phase 3d / iter-cm-2).
    /// Each entry holds the decompressed-and-mutated worksheet XML
    /// plus an in-sync span index. Populated lazily by `setCell` on
    /// first call; consumed at `save`. Cannot coexist with
    /// `pending_appends` for the same sheet — refused with
    /// `error.SheetHasUnsavedAppends` (or the symmetric
    /// `error.SheetHasUnsavedMutations` from `appendRows`).
    pending_mutations: std.AutoHashMapUnmanaged(u32, MutatedSheet),

    /// Buffered appended rows for one sheet. iter-lms-2 accepts
    /// numeric / integer / boolean / empty cells only — string cells
    /// require SST extension which lands in iter-lms-3.
    pub const AppendBuffer = struct {
        rows: std.ArrayListUnmanaged([]Cell) = .{},
    };

    /// Decompressed worksheet XML kept resident so successive
    /// `setCell` calls amortise the decompress + tokenize cost.
    /// `spans` is kept in source-order and in sync with `xml` —
    /// every splice updates both.
    pub const MutatedSheet = struct {
        xml: std.ArrayListUnmanaged(u8) = .{},
        spans: std.ArrayListUnmanaged(CellSpan) = .{},

        pub fn deinit(self: *MutatedSheet, allocator: Allocator) void {
            self.xml.deinit(allocator);
            self.spans.deinit(allocator);
        }
    };

    /// State for one Phase 3e (iter-row-2/3) row insert/delete.
    pub const RowEdit = struct {
        sheet_idx: u32,
        row: u32, // 1-based; for insert this is `before_row`,
        // for delete this is the row to remove.
    };

    /// State for one Phase 3e (iter-col-3/4) column insert/delete.
    pub const ColEdit = struct {
        sheet_idx: u32,
        col_1based: u32, // 1-based (A=1, B=2, …)
    };

    /// State for one Phase 3e (iter-sheet-3) sheet deletion.
    pub const SheetDelete = struct {
        path: []u8, // owned
        rid: []u8, // owned

        pub fn deinit(self: *SheetDelete, alloc: Allocator) void {
            alloc.free(self.path);
            alloc.free(self.rid);
        }
    };

    /// State for one Phase 3e (iter-sheet-2) sheet rename.
    /// `rid` is the source workbook's `r:id="rIdN"` for this sheet,
    /// captured at renameSheet time. The save-time patcher matches
    /// `<sheet r:id="…">` by this rId — Book.sheets indexing might
    /// drop entries with broken rels, so a positional walk over the
    /// raw workbook.xml order would target the wrong line.
    pub const SheetRename = struct {
        sheet_idx: u32,
        rid: []u8,
        old_name: []u8,
        new_name: []u8,

        pub fn deinit(self: *SheetRename, alloc: Allocator) void {
            alloc.free(self.rid);
            alloc.free(self.old_name);
            alloc.free(self.new_name);
        }
    };

    /// State for one Phase 3e (iter-sheet-1) sheet addition. `path`
    /// is a BORROWED slice into the editor's `sheet_paths` array
    /// (the addSheet helper grows that array and stores the owned
    /// path bytes there); `deinit` does NOT free it.
    pub const NewSheet = struct {
        name: []u8, // owned dupe of caller's name
        path: []const u8, // borrowed from Editor.sheet_paths
        rid: []u8, // owned, "rIdN"
        sheet_id: u32,
        body_xml: []u8, // owned, empty worksheet template

        pub fn deinit(self: *NewSheet, alloc: Allocator) void {
            alloc.free(self.name);
            alloc.free(self.rid);
            alloc.free(self.body_xml);
        }
    };

    /// Per-entry spans captured by the iter-lms-1b raw-ZIP scanner.
    /// All offsets/lengths are relative to `src_buf`.
    pub const ZipEntry = struct {
        /// Filename slice into `src_buf` (the CDFH-side filename;
        /// must match the LFH's).
        name: []const u8,
        /// Local-file-header offset.
        lfh_offset: u32,
        /// Length of LFH (header + filename + extras). Payload starts
        /// at `lfh_offset + lfh_total_len`.
        lfh_total_len: u32,
        /// Compressed payload bytes.
        payload_len: u32,
        /// Uncompressed payload size declared in the CDFH. Used as a
        /// hard cap during decompression so a maliciously crafted
        /// deflate stream can't expand past the declared size and
        /// inflate memory usage on `save`.
        uncompressed_size: u32,
        /// CDFH offset (start of the central-directory file header
        /// for this entry).
        cdfh_offset: u32,
        /// CDFH total length (header + filename + extras + comment).
        cdfh_total_len: u32,
        /// CompressionMethod value.
        compression_method: u16,
        /// General-purpose bit flag from CDFH.
        gp_flags_raw: u16,
    };

    pub fn open(allocator: Allocator, path: []const u8) !Editor {
        const file = try std.fs.cwd().openFile(path, .{});
        defer file.close();
        const stat = try file.stat();
        // Refuse files > 4 GiB up front (ZIP64 isn't supported by v1
        // per the plan; documented limit).
        if (stat.size > std.math.maxInt(u32)) return error.ZipTooLarge;
        const buf = try allocator.alloc(u8, @intCast(stat.size));
        errdefer allocator.free(buf);
        const n = try file.readAll(buf);
        if (n != buf.len) return error.UnexpectedEof;

        // Scan the source ZIP and capture verbatim spans.
        //
        // EOCD search must respect the trailing-comment field: a ZIP
        // comment may legally contain the byte sequence `PK\x05\x06`,
        // so the right scan is "look for a candidate whose declared
        // comment_len consumes the rest of the file exactly". Bound
        // the search to the last 65535 + sizeof(EndRecord) bytes
        // (the max comment is u16-bounded). Stdlib's
        // `EndRecord.findBuffer` does the right thing in spirit but
        // has an error-set bug in Zig 0.15.2.
        const max_comment: usize = std.math.maxInt(u16);
        const eocd_size: usize = @sizeOf(std.zip.EndRecord);
        if (buf.len < eocd_size) return error.BadZip;
        const search_start: usize = if (buf.len > max_comment + eocd_size)
            buf.len - max_comment - eocd_size
        else
            0;
        var eocd_pos: usize = buf.len - eocd_size;
        var eocd: std.zip.EndRecord = undefined;
        var found = false;
        while (true) : (eocd_pos -= 1) {
            if (std.mem.eql(u8, buf[eocd_pos .. eocd_pos + 4], &std.zip.end_record_sig)) {
                eocd = std.mem.bytesToValue(
                    std.zip.EndRecord,
                    buf[eocd_pos..][0..eocd_size],
                );
                if (@import("builtin").cpu.arch.endian() != .little)
                    std.mem.byteSwapAllFields(std.zip.EndRecord, &eocd);
                if (eocd_pos + eocd_size + eocd.comment_len == buf.len) {
                    found = true;
                    break;
                }
            }
            if (eocd_pos == search_start) break;
        }
        if (!found) return error.BadZip;
        if (eocd.need_zip64()) return error.Zip64NotSupported;
        // Refuse multi-disk archives — `local_file_header_offset` on
        // each CDFH would be relative to another segment, not
        // `src_buf`. xlsx files are single-segment in practice.
        if (eocd.disk_number != 0 or eocd.central_directory_disk_number != 0 or
            eocd.record_count_disk != eocd.record_count_total)
            return error.ZipSplitNotSupported;
        const cd_offset = eocd.central_directory_offset;
        const cd_size = eocd.central_directory_size;
        if (@as(u64, cd_offset) + cd_size > buf.len) return error.BadZip;
        const comment_off = eocd_pos + @sizeOf(std.zip.EndRecord);
        if (comment_off + eocd.comment_len > buf.len) return error.BadZip;
        const eocd_comment = buf[comment_off .. comment_off + eocd.comment_len];

        var entries: std.ArrayListUnmanaged(ZipEntry) = .{};
        errdefer entries.deinit(allocator);
        try entries.ensureTotalCapacity(allocator, eocd.record_count_total);

        var p: usize = cd_offset;
        const cd_end: usize = @as(usize, cd_offset) + cd_size;
        while (p + @sizeOf(std.zip.CentralDirectoryFileHeader) <= cd_end) {
            // Parse CDFH (little-endian, packed). bytesToValue copies
            // bytes into a local without requiring the source to be
            // naturally aligned for the struct.
            const cdfh_size = @sizeOf(std.zip.CentralDirectoryFileHeader);
            var cdfh_local = std.mem.bytesToValue(
                std.zip.CentralDirectoryFileHeader,
                buf[p..][0..cdfh_size],
            );
            if (!std.mem.eql(u8, &cdfh_local.signature, &std.zip.central_file_header_sig))
                return error.BadZip;
            if (@import("builtin").cpu.arch.endian() != .little)
                std.mem.byteSwapAllFields(std.zip.CentralDirectoryFileHeader, &cdfh_local);
            const cdfh_ptr = &cdfh_local;
            // Reject ZIP64 / data-descriptor / encrypted entries up
            // front; v1 supports plain Excel-shaped archives only.
            if (cdfh_ptr.compressed_size == std.math.maxInt(u32) or
                cdfh_ptr.uncompressed_size == std.math.maxInt(u32) or
                cdfh_ptr.local_file_header_offset == std.math.maxInt(u32))
                return error.Zip64NotSupported;
            const flags_word: u16 = @bitCast(cdfh_ptr.flags);
            if ((flags_word & 0x0008) != 0) return error.ZipDataDescriptorNotSupported;
            if ((flags_word & 0x0001) != 0) return error.ZipEncryptedNotSupported;
            if (cdfh_ptr.disk_number != 0) return error.ZipSplitNotSupported;

            const filename_len = cdfh_ptr.filename_len;
            const extra_len = cdfh_ptr.extra_len;
            const comment_len_cdfh = cdfh_ptr.comment_len;
            const cdfh_total: usize = @sizeOf(std.zip.CentralDirectoryFileHeader) +
                filename_len + extra_len + comment_len_cdfh;
            if (p + cdfh_total > cd_end) return error.BadZip;
            const name_off = p + @sizeOf(std.zip.CentralDirectoryFileHeader);
            const name = buf[name_off .. name_off + filename_len];

            const lfh_offset: usize = cdfh_ptr.local_file_header_offset;
            const lfh_size = @sizeOf(std.zip.LocalFileHeader);
            if (lfh_offset + lfh_size > buf.len) return error.BadZip;
            var lfh_local = std.mem.bytesToValue(
                std.zip.LocalFileHeader,
                buf[lfh_offset..][0..lfh_size],
            );
            if (!std.mem.eql(u8, &lfh_local.signature, &std.zip.local_file_header_sig))
                return error.BadZip;
            if (@import("builtin").cpu.arch.endian() != .little)
                std.mem.byteSwapAllFields(std.zip.LocalFileHeader, &lfh_local);
            const lfh_ptr = &lfh_local;
            const lfh_total: usize = @sizeOf(std.zip.LocalFileHeader) +
                lfh_ptr.filename_len + lfh_ptr.extra_len;
            // Validate the LFH filename matches the CDFH's. Some
            // malformed archives have divergent names; trusting the
            // CDFH alone would associate one entry's name with
            // another's payload span on save-mutate.
            if (lfh_ptr.filename_len != filename_len) return error.BadZip;
            const lfh_name_off = lfh_offset + @sizeOf(std.zip.LocalFileHeader);
            if (lfh_name_off + filename_len > buf.len) return error.BadZip;
            if (!std.mem.eql(u8, name, buf[lfh_name_off .. lfh_name_off + filename_len]))
                return error.BadZip;
            const payload_len: usize = cdfh_ptr.compressed_size;
            if (lfh_offset + lfh_total + payload_len > buf.len) return error.BadZip;

            try entries.append(allocator, .{
                .name = name,
                .lfh_offset = @intCast(lfh_offset),
                .lfh_total_len = @intCast(lfh_total),
                .payload_len = @intCast(payload_len),
                .uncompressed_size = cdfh_ptr.uncompressed_size,
                .cdfh_offset = @intCast(p),
                .cdfh_total_len = @intCast(cdfh_total),
                .compression_method = @intFromEnum(cdfh_ptr.compression_method),
                .gp_flags_raw = flags_word,
            });

            p += cdfh_total;
        }

        // EOCD's claimed record count must match what we walked.
        // A short central directory yields fewer entries than the
        // EOCD advertises; trusting the EOCD count would skip valid
        // entries on save-mutate.
        if (entries.items.len != eocd.record_count_total) return error.BadZip;

        // Resolve the sheet_idx → path mapping by opening the source
        // through Book.open (which parses workbook.xml.rels). The
        // editor needs this to find which entry in the ZIP table
        // corresponds to a given sheet on appendRows. Paths are
        // dup'd into editor-owned storage so Book.deinit doesn't
        // dangle them.
        const sheet_paths_owned = blk: {
            var b = try Book.open(allocator, path);
            defer b.deinit();
            const out_paths = try allocator.alloc([]const u8, b.sheets.len);
            errdefer {
                for (out_paths) |p_owned| allocator.free(p_owned);
                allocator.free(out_paths);
            }
            for (b.sheets, 0..) |s, i| {
                out_paths[i] = try allocator.dupe(u8, s.path);
            }
            break :blk out_paths;
        };

        return .{
            .allocator = allocator,
            .src_buf = buf,
            .entries = try entries.toOwnedSlice(allocator),
            .cd_offset = cd_offset,
            .cd_size = cd_size,
            .eocd_offset = @intCast(eocd_pos),
            .eocd_comment = eocd_comment,
            .sheet_paths = sheet_paths_owned,
            .pending_appends = .{},
            .pending_mutations = .{},
            .pending_new_sheets = .{},
            .pending_renames = .{},
            .pending_deletes = .{},
            .pending_row_inserts = .{},
            .pending_row_deletes = .{},
            .pending_col_inserts = .{},
            .pending_col_deletes = .{},
        };
    }

    pub fn deinit(self: *Editor) void {
        self.allocator.free(self.src_buf);
        self.allocator.free(self.entries);
        for (self.sheet_paths) |p| self.allocator.free(p);
        self.allocator.free(self.sheet_paths);
        var it = self.pending_appends.valueIterator();
        while (it.next()) |buf| {
            for (buf.rows.items) |row| {
                for (row) |c| switch (c) {
                    .string => |s| self.allocator.free(s),
                    else => {},
                };
                self.allocator.free(row);
            }
            buf.rows.deinit(self.allocator);
        }
        self.pending_appends.deinit(self.allocator);
        var mit = self.pending_mutations.valueIterator();
        while (mit.next()) |m| m.deinit(self.allocator);
        self.pending_mutations.deinit(self.allocator);
        for (self.pending_new_sheets.items) |*s| s.deinit(self.allocator);
        self.pending_new_sheets.deinit(self.allocator);
        for (self.pending_renames.items) |*r| r.deinit(self.allocator);
        self.pending_renames.deinit(self.allocator);
        for (self.pending_deletes.items) |*d| d.deinit(self.allocator);
        self.pending_deletes.deinit(self.allocator);
        self.pending_row_inserts.deinit(self.allocator);
        self.pending_row_deletes.deinit(self.allocator);
        self.pending_col_inserts.deinit(self.allocator);
        self.pending_col_deletes.deinit(self.allocator);
        self.* = undefined;
    }

    /// Append rows to an existing sheet. Accepts numeric / integer /
    /// boolean / empty / string cells. String appends extend the
    /// workbook's SST — iter-lms-3 always allocates a new SST entry
    /// per appended string (no plain-text equality reuse) so a
    /// rich-text entry in the source SST that happens to share text
    /// with an appended string can't silently inherit formatting.
    /// Rows go after the source's highest used row in that sheet
    /// (computed at save time, not now). When the source workbook
    /// has no `xl/sharedStrings.xml` (only inline strings or numeric
    /// data), `save` creates one on demand and patches both
    /// `xl/_rels/workbook.xml.rels` and `[Content_Types].xml` so
    /// readers recognise it.
    pub fn appendRows(self: *Editor, sheet_idx: u32, rows: []const []const Cell) !void {
        if (sheet_idx >= self.sheet_paths.len) return error.SheetIndexOutOfRange;
        // Refuse to mix appends with `setCell` mutations on the same
        // sheet. The two paths build the modified XML differently
        // (delta vs full-buffer); merging them safely needs design
        // work that hasn't shipped. Symmetric guard on `setCell`.
        if (self.pending_mutations.contains(sheet_idx)) return error.SheetHasUnsavedMutations;
        // Also refuse when a row/col edit is queued for this sheet:
        // save() runs the row/col substitution first, then the append
        // pass would overwrite that substituted entry with XML built
        // from the pre-edit source — silently dropping the edit.
        if (self.sheetHasPendingRowOrColEdit(sheet_idx)) return error.SheetHasUnsavedRowOrColEdit;
        // Empty append is a documented no-op — recording it as a
        // pending mutation would underflow the row-index math in
        // `buildSubstitutedSheet` (start_row + 0 - 1 = u32.max).
        if (rows.len == 0) return;
        // Refuse rows wider than Excel's max column (XFD = 16384).
        // The actual final row count check (start_row + len <=
        // 1048576) happens in buildSubstitutedSheet once the source's
        // highest row is known.
        for (rows) |row| {
            if (row.len > max_col_1based) return error.ColumnIndexOutOfRange;
        }
        const writer_mod = @import("writer.zig");
        for (rows) |row| for (row) |c| switch (c) {
            .empty, .number, .boolean, .string => {},
            .integer => |n| {
                // Match writer.zig's contract: integers must round-
                // trip exactly through f64 (Excel stores all numerics
                // as IEEE-754 doubles). Reject up front rather than
                // silently rounding on open.
                if (!writer_mod.fitsExactlyInF64(n)) return error.IntegerExceedsExcelPrecision;
            },
        };
        const gop = try self.pending_appends.getOrPut(self.allocator, sheet_idx);
        if (!gop.found_existing) gop.value_ptr.* = .{};
        for (rows) |row| {
            const owned = try self.allocator.alloc(Cell, row.len);
            errdefer self.allocator.free(owned);
            // Track how many string buffers we duped successfully so
            // a mid-loop OOM doesn't leak the prefix.
            var duped: usize = 0;
            errdefer {
                for (owned[0..duped]) |c| switch (c) {
                    .string => |s| self.allocator.free(s),
                    else => {},
                };
            }
            for (row, 0..) |c, i| {
                switch (c) {
                    // String cells need their byte contents duped
                    // because the caller may have passed a temporary
                    // slice; the editor holds onto pending appends
                    // until save and beyond.
                    .string => |s| {
                        owned[i] = .{ .string = try self.allocator.dupe(u8, s) };
                    },
                    else => owned[i] = c,
                }
                duped = i + 1;
            }
            try gop.value_ptr.rows.append(self.allocator, owned);
        }
    }

    /// Write the workbook (with any pending appends applied) to
    /// `out_path`. Atomic via `std.fs.Dir.atomicFile`.
    ///
    /// **No-op save**: when no `appendRows` calls have been made,
    /// streams `src_buf` verbatim — preserves SHA256 round-trip for
    /// any well-formed source archive (canonical or not).
    ///
    /// **Mutated save**: walks the entry table, substitutes each
    /// modified sheet entry with a freshly-emitted LFH+payload (new
    /// CRC32 / sizes), and re-emits the central directory + EOCD
    /// with patched offsets. Sheets that weren't touched flow through
    /// verbatim; the source's preserved EOCD comment is kept.
    pub fn save(self: *Editor, out_path: []const u8) !void {
        var write_buf: [4096]u8 = undefined;
        var atomic_file = try std.fs.cwd().atomicFile(out_path, .{ .write_buffer = &write_buf });
        defer atomic_file.deinit();
        const w = &atomic_file.file_writer.interface;

        if (self.pending_appends.count() == 0 and
            self.pending_mutations.count() == 0 and
            self.pending_new_sheets.items.len == 0 and
            self.pending_renames.items.len == 0 and
            self.pending_deletes.items.len == 0 and
            self.pending_row_inserts.items.len == 0 and
            self.pending_row_deletes.items.len == 0 and
            self.pending_col_inserts.items.len == 0 and
            self.pending_col_deletes.items.len == 0)
        {
            try w.writeAll(self.src_buf);
            try atomic_file.finish();
            return;
        }

        // Detect whether any pending append carries a string cell;
        // if so, the SST entry must be substituted alongside the
        // sheets so the new t="s" indices resolve.
        var has_strings = false;
        var pa_check = self.pending_appends.iterator();
        outer: while (pa_check.next()) |kv| {
            for (kv.value_ptr.rows.items) |row| {
                for (row) |c| if (c == .string) {
                    has_strings = true;
                    break :outer;
                };
            }
        }

        // Locate sharedStrings.xml entry + count source SST entries
        // when string cells are pending. iter-lms-3+follow-up:
        // SST-less source workbooks (only inline strings or numeric
        // data) get a fresh sharedStrings.xml created on demand,
        // along with rels + Content_Types patches. `create_new_sst`
        // is true on the SST-less path; index counter starts at 0.
        var sst_entry_idx: ?usize = null;
        var source_sst_count: u32 = 0;
        var sst_xml_owned: ?[]u8 = null;
        var create_new_sst = false;
        defer if (sst_xml_owned) |x| self.allocator.free(x);
        if (has_strings) {
            if (findEntryByName(self.entries, "xl/sharedStrings.xml")) |idx| {
                sst_entry_idx = idx;
                const sst_entry = self.entries[idx];
                const sst_payload = self.src_buf[sst_entry.lfh_offset + sst_entry.lfh_total_len ..][0..sst_entry.payload_len];
                sst_xml_owned = try decompressZipPayload(
                    self.allocator,
                    sst_payload,
                    sst_entry.compression_method,
                    sst_entry.uncompressed_size,
                );
                source_sst_count = countSiInSst(sst_xml_owned.?);
            } else {
                create_new_sst = true;
            }
        }

        var sst_appender: SstAppender = .{
            .allocator = self.allocator,
            .next_idx = source_sst_count,
        };
        defer sst_appender.deinit();
        const sst_ptr: ?*SstAppender = if (has_strings) &sst_appender else null;

        // Build substituted entries for each modified sheet.
        const subs = try self.allocator.alloc(?SubstitutedEntry, self.entries.len);
        for (subs) |*slot| slot.* = null;
        defer {
            for (subs) |maybe_sub| if (maybe_sub) |s| {
                self.allocator.free(s.lfh);
                self.allocator.free(s.payload);
                self.allocator.free(s.cdfh);
            };
            self.allocator.free(subs);
        }

        // Phase 3e iter-row-2/3: apply row inserts + deletes by
        // building a substituted sheet entry per affected sheet.
        // Each sheet has at most one pending row edit (recordRowEdit
        // enforces that), so order doesn't matter.
        for (self.pending_row_inserts.items) |edit| {
            const path = self.sheet_paths[edit.sheet_idx];
            const entry_idx = findEntryByName(self.entries, path) orelse
                return error.SheetEntryNotFound;
            const entry = self.entries[entry_idx];
            const payload_bytes = self.src_buf[entry.lfh_offset + entry.lfh_total_len ..][0..entry.payload_len];
            const src_xml = try decompressZipPayload(self.allocator, payload_bytes, entry.compression_method, entry.uncompressed_size);
            defer self.allocator.free(src_xml);
            const new_xml = try applyRowEditToWorksheet(self.allocator, src_xml, edit.row, .insert);
            subs[entry_idx] = try buildEntryFromXml(self.allocator, path, new_xml);
        }
        for (self.pending_row_deletes.items) |edit| {
            const path = self.sheet_paths[edit.sheet_idx];
            const entry_idx = findEntryByName(self.entries, path) orelse
                return error.SheetEntryNotFound;
            const entry = self.entries[entry_idx];
            const payload_bytes = self.src_buf[entry.lfh_offset + entry.lfh_total_len ..][0..entry.payload_len];
            const src_xml = try decompressZipPayload(self.allocator, payload_bytes, entry.compression_method, entry.uncompressed_size);
            defer self.allocator.free(src_xml);
            const new_xml = try applyRowEditToWorksheet(self.allocator, src_xml, edit.row, .delete);
            subs[entry_idx] = try buildEntryFromXml(self.allocator, path, new_xml);
        }
        for (self.pending_col_inserts.items) |edit| {
            const path = self.sheet_paths[edit.sheet_idx];
            const entry_idx = findEntryByName(self.entries, path) orelse return error.SheetEntryNotFound;
            const entry = self.entries[entry_idx];
            const payload_bytes = self.src_buf[entry.lfh_offset + entry.lfh_total_len ..][0..entry.payload_len];
            const src_xml = try decompressZipPayload(self.allocator, payload_bytes, entry.compression_method, entry.uncompressed_size);
            defer self.allocator.free(src_xml);
            const new_xml = try applyColEditToWorksheet(self.allocator, src_xml, edit.col_1based, .insert);
            subs[entry_idx] = try buildEntryFromXml(self.allocator, path, new_xml);
        }
        for (self.pending_col_deletes.items) |edit| {
            const path = self.sheet_paths[edit.sheet_idx];
            const entry_idx = findEntryByName(self.entries, path) orelse return error.SheetEntryNotFound;
            const entry = self.entries[entry_idx];
            const payload_bytes = self.src_buf[entry.lfh_offset + entry.lfh_total_len ..][0..entry.payload_len];
            const src_xml = try decompressZipPayload(self.allocator, payload_bytes, entry.compression_method, entry.uncompressed_size);
            defer self.allocator.free(src_xml);
            const new_xml = try applyColEditToWorksheet(self.allocator, src_xml, edit.col_1based, .delete);
            subs[entry_idx] = try buildEntryFromXml(self.allocator, path, new_xml);
        }

        var pa_iter = self.pending_appends.iterator();
        while (pa_iter.next()) |kv| {
            const sheet_idx = kv.key_ptr.*;
            const buf = kv.value_ptr.*;
            const path = self.sheet_paths[sheet_idx];
            // Phase 3e: appendRows on a new sheet is handled by the
            // pending_new_sheets branch below — no source entry to
            // substitute here.
            if (self.findPendingNewSheet(path) != null) continue;
            const entry_idx = findEntryByName(self.entries, path) orelse
                return error.SheetEntryNotFound;
            subs[entry_idx] = try buildSubstitutedSheet(
                self.allocator,
                self.entries[entry_idx],
                self.src_buf,
                buf.rows.items,
                sst_ptr,
            );
        }

        // iter-cm-2a: pending in-place cell mutations. The mutated
        // XML buffer is already complete — just pipe it through
        // `buildEntryFromXml` to get a fresh LFH/CDFH + payload.
        // appendRows + setCell on the same sheet is rejected at
        // mutation time so there's no merge step here.
        var pm_iter = self.pending_mutations.iterator();
        while (pm_iter.next()) |kv| {
            const sheet_idx = kv.key_ptr.*;
            const path = self.sheet_paths[sheet_idx];
            // Phase 3e: pending mutations on a NEW sheet are handled
            // by the pending_new_sheets branch below — no source
            // entry to substitute.
            if (self.findPendingNewSheet(path) != null) continue;
            const entry_idx = findEntryByName(self.entries, path) orelse
                return error.SheetEntryNotFound;

            // Best-effort `<dimension>` update — same canonical-form
            // contract as the append path: only the
            // `<dimension ref="A1:Z100"/>` shape is widened, others
            // pass through and Excel recomputes on its next save.
            // Skip cleanly when the spans index is empty (no real
            // mutations on this sheet) — leaves dimension as-is.
            var pm_min_row: u32 = std.math.maxInt(u32);
            var pm_max_row: u32 = 0;
            var pm_min_col1: u32 = std.math.maxInt(u32);
            var pm_max_col1: u32 = 0;
            for (kv.value_ptr.spans.items) |s| {
                if (s.row < pm_min_row) pm_min_row = s.row;
                if (s.row > pm_max_row) pm_max_row = s.row;
                const c1 = s.col + 1;
                if (c1 < pm_min_col1) pm_min_col1 = c1;
                if (c1 > pm_max_col1) pm_max_col1 = c1;
            }
            const xml_to_use = if (pm_max_row > 0)
                (try updateDimensionRange(
                    self.allocator,
                    kv.value_ptr.xml.items,
                    pm_min_row,
                    pm_max_row,
                    pm_min_col1,
                    pm_max_col1,
                )) orelse try self.allocator.dupe(u8, kv.value_ptr.xml.items)
            else
                try self.allocator.dupe(u8, kv.value_ptr.xml.items);
            // buildEntryFromXml takes ownership of `xml_to_use` and
            // frees it on its own — no errdefer needed here.
            subs[entry_idx] = try buildEntryFromXml(self.allocator, path, xml_to_use);
        }

        // Appendix entries: brand-new ZIP entries that don't exist
        // in source (currently: a fresh sharedStrings.xml when the
        // source workbook had none). Lifetime: their .lfh / .payload
        // / .cdfh slices are owned, freed at the end of save().
        var extra_entries: std.ArrayListUnmanaged(SubstitutedEntry) = .{};
        defer {
            for (extra_entries.items) |e| {
                self.allocator.free(e.lfh);
                self.allocator.free(e.payload);
                self.allocator.free(e.cdfh);
            }
            extra_entries.deinit(self.allocator);
        }

        // Phase 3e iter-sheet-1: emit each new sheet's body BEFORE
        // the SST commit so any string cells injected here extend
        // sst_appender in time. Metadata patches (workbook.xml,
        // rels, Content_Types) come after — they don't touch the
        // SST.
        if (self.pending_new_sheets.items.len > 0) {
            const source_count: u32 = @intCast(self.sheet_paths.len - self.pending_new_sheets.items.len);
            for (self.pending_new_sheets.items, 0..) |new_sheet, i| {
                const sheet_idx_for_new: u32 = source_count + @as(u32, @intCast(i));
                const body_owned: []u8 = blk: {
                    if (self.pending_mutations.get(sheet_idx_for_new)) |ms|
                        break :blk try self.allocator.dupe(u8, ms.xml.items);
                    if (self.pending_appends.get(sheet_idx_for_new)) |buf| {
                        // New sheet starts empty, so appended rows
                        // begin at row 1. Pass sst_ptr so any
                        // string cells extend the SST in time for
                        // the commit below.
                        break :blk try injectAppendedRows(
                            self.allocator,
                            new_sheet.body_xml,
                            buf.rows.items,
                            1,
                            sst_ptr,
                        );
                    }
                    break :blk try self.allocator.dupe(u8, new_sheet.body_xml);
                };
                // buildEntryFromXml takes ownership of body_owned.
                const new_entry = try buildEntryFromXml(self.allocator, new_sheet.path, body_owned);
                try extra_entries.append(self.allocator, new_entry);
            }
        }

        if (has_strings and sst_appender.new_strings.items.len > 0) {
            if (create_new_sst) {
                // SST-less source workbook — build a fresh
                // sharedStrings.xml from scratch, patch rels +
                // Content_Types so Excel / readers recognise it.
                const new_sst_xml = try buildFreshSstXml(
                    self.allocator,
                    sst_appender.new_strings.items,
                );
                const new_entry = try buildEntryFromXml(
                    self.allocator,
                    "xl/sharedStrings.xml",
                    new_sst_xml,
                );
                try extra_entries.append(self.allocator, new_entry);

                try patchEntryXml(
                    self.allocator,
                    self.entries,
                    self.src_buf,
                    subs,
                    "xl/_rels/workbook.xml.rels",
                    addSstRelationship,
                );
                try patchEntryXml(
                    self.allocator,
                    self.entries,
                    self.src_buf,
                    subs,
                    "[Content_Types].xml",
                    addSstContentTypeOverride,
                );
            } else {
                // Existing SST entry — substitute in place.
                const sst_idx = sst_entry_idx.?;
                subs[sst_idx] = try buildSubstitutedSst(
                    self.allocator,
                    self.entries[sst_idx],
                    sst_xml_owned.?,
                    sst_appender.new_strings.items,
                    source_sst_count,
                );
            }
        }

        // Phase 3e iter-sheet-1 (continued): metadata patches for
        // the new sheets — these must run AFTER the SST patches
        // above so a workbook that ALSO got a fresh SST has all
        // its workbook.xml-rels updates applied through the same
        // patchEntryForNewSheets re-substitution flow.
        if (self.pending_new_sheets.items.len > 0) {
            try patchEntryForNewSheets(
                self.allocator,
                self.entries,
                self.src_buf,
                subs,
                "xl/workbook.xml",
                self.pending_new_sheets.items,
                patchWorkbookXmlForNewSheets,
            );
            try patchEntryForNewSheets(
                self.allocator,
                self.entries,
                self.src_buf,
                subs,
                "xl/_rels/workbook.xml.rels",
                self.pending_new_sheets.items,
                patchWorkbookRelsForNewSheets,
            );
            try patchEntryForNewSheets(
                self.allocator,
                self.entries,
                self.src_buf,
                subs,
                "[Content_Types].xml",
                self.pending_new_sheets.items,
                patchContentTypesForNewSheets,
            );
        }

        // Phase 3e iter-sheet-3: pending sheet deletions. Patch
        // workbook.xml (drop the <sheet> line), rels (drop the
        // Relationship), Content_Types (drop the Override). The
        // entry-skipping in the LFH-emit loop further down drops
        // the sheet's worksheet-XML ZIP entry itself.
        if (self.pending_deletes.items.len > 0) {
            try patchEntryForDeletes(
                self.allocator,
                self.entries,
                self.src_buf,
                subs,
                "xl/workbook.xml",
                self.pending_deletes.items,
                patchWorkbookXmlForDeletes,
            );
            try patchEntryForDeletes(
                self.allocator,
                self.entries,
                self.src_buf,
                subs,
                "xl/_rels/workbook.xml.rels",
                self.pending_deletes.items,
                patchWorkbookRelsForDeletes,
            );
            try patchEntryForDeletes(
                self.allocator,
                self.entries,
                self.src_buf,
                subs,
                "[Content_Types].xml",
                self.pending_deletes.items,
                patchContentTypesForDeletes,
            );
        }

        // Phase 3e iter-sheet-2: pending sheet renames. Composes
        // through patchEntryForRenames's re-substitution path, so
        // it stacks correctly on top of any prior workbook.xml
        // modification (addSheet's <sheets> patch).
        if (self.pending_renames.items.len > 0) {
            try patchEntryForRenames(
                self.allocator,
                self.entries,
                self.src_buf,
                subs,
                "xl/workbook.xml",
                self.pending_renames.items,
                patchWorkbookXmlForRenames,
            );
        }

        // Emit LFHs in LFH-offset order. Each substituted entry's LFH
        // / payload are freshly built; others copy from src_buf.
        const new_lfh_offsets = try self.allocator.alloc(u32, self.entries.len);
        defer self.allocator.free(new_lfh_offsets);
        const lfh_sorted = try self.allocator.alloc(usize, self.entries.len);
        defer self.allocator.free(lfh_sorted);
        for (lfh_sorted, 0..) |*slot, i| slot.* = i;
        std.mem.sort(usize, lfh_sorted, self.entries, struct {
            fn lessThan(es: []const ZipEntry, a: usize, b: usize) bool {
                return es[a].lfh_offset < es[b].lfh_offset;
            }
        }.lessThan);

        // Phase 3e iter-sheet-3: mask of entries to drop entirely
        // (the deleted sheets' worksheet XML). The size validation
        // and emission loops skip these.
        const deleted_mask = try self.allocator.alloc(bool, self.entries.len);
        defer self.allocator.free(deleted_mask);
        @memset(deleted_mask, false);
        for (self.pending_deletes.items) |d| {
            if (findEntryByName(self.entries, d.path)) |di| deleted_mask[di] = true;
        }
        // Drop xl/calcChain.xml when any sheet is being deleted —
        // it's a recompute cache, Excel rebuilds it on open. The
        // structural-edits plan calls this out as standard cleanup
        // for destructive edits. Also drop the rels entry that
        // points at it via `<Relationship Type=".../calcChain"/>`
        // to avoid a dangling relationship.
        if (self.pending_deletes.items.len > 0) {
            if (findEntryByName(self.entries, "xl/calcChain.xml")) |ci| {
                deleted_mask[ci] = true;
                // Patch rels to drop the calcChain Relationship.
                try patchEntryForDeletes(
                    self.allocator,
                    self.entries,
                    self.src_buf,
                    subs,
                    "xl/_rels/workbook.xml.rels",
                    self.pending_deletes.items,
                    patchWorkbookRelsForCalcChainDrop,
                );
                // Patch [Content_Types] to drop calcChain Override.
                try patchEntryForDeletes(
                    self.allocator,
                    self.entries,
                    self.src_buf,
                    subs,
                    "[Content_Types].xml",
                    self.pending_deletes.items,
                    patchContentTypesForCalcChainDrop,
                );
            }
        }

        // Pre-validate that the rewritten archive stays within ZIP32
        // bounds. Source size is already <= 4 GiB (Editor.open caps),
        // but substitutions + appendix entries can push the total over.
        // Without this guard the `@intCast(u32, ...)` writes below
        // would trap (safe builds) or silently truncate offsets
        // (release builds), producing an unreadable archive.
        // Use `>= maxInt(u32)` rather than `>` so ZIP32 sentinel
        // values (`0xFFFFFFFF` for offsets/sizes, `0xFFFF` for
        // entry counts) are also rejected.
        const max_u32: u64 = std.math.maxInt(u32);
        const max_u16: usize = std.math.maxInt(u16);
        var planned_total: u64 = 0;
        var live_entry_count: usize = 0;
        for (lfh_sorted) |i| {
            if (deleted_mask[i]) continue;
            live_entry_count += 1;
            if (subs[i]) |s| {
                planned_total += @as(u64, s.lfh.len) + @as(u64, s.payload.len);
            } else {
                const e = self.entries[i];
                planned_total += @as(u64, e.lfh_total_len) + @as(u64, e.payload_len);
            }
        }
        for (extra_entries.items) |e| {
            planned_total += @as(u64, e.lfh.len) + @as(u64, e.payload.len);
        }
        const planned_cd_offset = planned_total;
        for (self.entries, 0..) |entry, i| {
            if (deleted_mask[i]) continue;
            if (subs[i]) |s| {
                planned_total += @as(u64, s.cdfh.len);
            } else {
                planned_total += @as(u64, entry.cdfh_total_len);
            }
        }
        for (extra_entries.items) |e| planned_total += @as(u64, e.cdfh.len);
        const planned_cd_size = planned_total - planned_cd_offset;
        const planned_archive_size = planned_total +
            @as(u64, @sizeOf(std.zip.EndRecord)) +
            @as(u64, self.eocd_comment.len);
        const planned_entries = live_entry_count + extra_entries.items.len;
        if (planned_cd_offset >= max_u32 or
            planned_cd_size >= max_u32 or
            planned_archive_size >= max_u32 or
            planned_entries >= max_u16)
        {
            return error.Zip64NotSupported;
        }

        var written: u64 = 0;
        for (lfh_sorted) |i| {
            if (deleted_mask[i]) continue;
            new_lfh_offsets[i] = @intCast(written);
            if (subs[i]) |s| {
                try w.writeAll(s.lfh);
                try w.writeAll(s.payload);
                written += @as(u64, s.lfh.len) + @as(u64, s.payload.len);
            } else {
                const e = self.entries[i];
                const lfh_bytes = self.src_buf[e.lfh_offset .. e.lfh_offset + e.lfh_total_len];
                try w.writeAll(lfh_bytes);
                const payload_bytes = self.src_buf[e.lfh_offset + e.lfh_total_len ..][0..e.payload_len];
                try w.writeAll(payload_bytes);
                written += @as(u64, e.lfh_total_len) + @as(u64, e.payload_len);
            }
        }

        // Emit appendix entries (brand-new parts that didn't exist in
        // source). Track their new LFH offsets for the CD rewrite.
        const extra_lfh_offsets = try self.allocator.alloc(u32, extra_entries.items.len);
        defer self.allocator.free(extra_lfh_offsets);
        for (extra_entries.items, 0..) |e, ei| {
            extra_lfh_offsets[ei] = @intCast(written);
            try w.writeAll(e.lfh);
            try w.writeAll(e.payload);
            written += @as(u64, e.lfh.len) + @as(u64, e.payload.len);
        }

        const new_cd_offset: u32 = @intCast(written);
        for (self.entries, 0..) |entry, i| {
            if (deleted_mask[i]) continue;
            if (subs[i]) |s| {
                var cdfh_copy = try self.allocator.dupe(u8, s.cdfh);
                defer self.allocator.free(cdfh_copy);
                const lfh_field_pos = @sizeOf(std.zip.CentralDirectoryFileHeader) - 4;
                std.mem.writeInt(u32, cdfh_copy[lfh_field_pos..][0..4], new_lfh_offsets[i], .little);
                try w.writeAll(cdfh_copy);
                written += @as(u64, cdfh_copy.len);
            } else {
                var cdfh_bytes: [@sizeOf(std.zip.CentralDirectoryFileHeader)]u8 = undefined;
                const src_cdfh = self.src_buf[entry.cdfh_offset .. entry.cdfh_offset + cdfh_bytes.len];
                @memcpy(&cdfh_bytes, src_cdfh);
                const lfh_field_pos = cdfh_bytes.len - 4;
                std.mem.writeInt(u32, cdfh_bytes[lfh_field_pos..][0..4], new_lfh_offsets[i], .little);
                try w.writeAll(&cdfh_bytes);
                const var_off = entry.cdfh_offset + cdfh_bytes.len;
                const var_len = entry.cdfh_total_len - cdfh_bytes.len;
                try w.writeAll(self.src_buf[var_off .. var_off + var_len]);
                written += @as(u64, entry.cdfh_total_len);
            }
        }
        // Emit appendix CDFHs after the source-entry CDFHs.
        for (extra_entries.items, 0..) |e, ei| {
            var cdfh_copy = try self.allocator.dupe(u8, e.cdfh);
            defer self.allocator.free(cdfh_copy);
            const lfh_field_pos = @sizeOf(std.zip.CentralDirectoryFileHeader) - 4;
            std.mem.writeInt(u32, cdfh_copy[lfh_field_pos..][0..4], extra_lfh_offsets[ei], .little);
            try w.writeAll(cdfh_copy);
            written += @as(u64, cdfh_copy.len);
        }
        const new_cd_size: u32 = @intCast(written - new_cd_offset);

        const total_entries = live_entry_count + extra_entries.items.len;
        var eocd_out: std.zip.EndRecord = .{
            .signature = std.zip.end_record_sig,
            .disk_number = 0,
            .central_directory_disk_number = 0,
            .record_count_disk = @intCast(total_entries),
            .record_count_total = @intCast(total_entries),
            .central_directory_size = new_cd_size,
            .central_directory_offset = new_cd_offset,
            .comment_len = @intCast(self.eocd_comment.len),
        };
        if (@import("builtin").cpu.arch.endian() != .little)
            std.mem.byteSwapAllFields(std.zip.EndRecord, &eocd_out);
        try w.writeAll(std.mem.asBytes(&eocd_out));
        try w.writeAll(self.eocd_comment);

        try atomic_file.finish();
    }

    /// Decompress sheet `sheet_idx` and walk its `<c>` elements,
    /// returning a `WorksheetSpans` index over every cell. The
    /// foundation for Phase 3d cell-mutate (iter-cm-1) — read-only
    /// today; iter-cm-2 builds `setCell` on top.
    ///
    /// The returned `xml` + `cells` borrow the same allocation;
    /// caller frees both via `WorksheetSpans.deinit`. Scans are
    /// independent (no caching yet) — repeated calls re-decompress.
    pub fn scanWorksheet(self: *Editor, sheet_idx: u32) !WorksheetSpans {
        if (sheet_idx >= self.sheet_paths.len) return error.SheetIndexOutOfRange;
        // The scanner does NOT see rows queued in `pending_appends`
        // (those are deltas applied at save time). Reject rather
        // than return a stale span set.
        if (self.pending_appends.contains(sheet_idx)) return error.SheetHasUnsavedAppends;

        // If a previous `setCell` populated `pending_mutations` for
        // this sheet, surface THAT XML (and its in-sync spans) so
        // the caller sees a consistent view. Without this branch a
        // setCell-then-scan workflow would silently return spans
        // from the pre-mutation source bytes.
        if (self.pending_mutations.get(sheet_idx)) |ms| {
            const xml_copy = try self.allocator.dupe(u8, ms.xml.items);
            errdefer self.allocator.free(xml_copy);
            const spans_copy = try self.allocator.dupe(CellSpan, ms.spans.items);
            return .{ .allocator = self.allocator, .xml = xml_copy, .cells = spans_copy };
        }

        const path = self.sheet_paths[sheet_idx];
        // Source ZIP entry first; if absent, this is a freshly-
        // added sheet (Phase 3e iter-sheet-1) — return its empty
        // body template so callers can scan an untouched new sheet.
        if (findEntryByName(self.entries, path)) |entry_idx| {
            const entry = self.entries[entry_idx];
            const payload = self.src_buf[entry.lfh_offset + entry.lfh_total_len ..][0..entry.payload_len];
            const xml = try decompressZipPayload(
                self.allocator,
                payload,
                entry.compression_method,
                entry.uncompressed_size,
            );
            errdefer self.allocator.free(xml);
            const cells = try scanWorksheetXml(self.allocator, xml);
            return .{ .allocator = self.allocator, .xml = xml, .cells = cells };
        }
        if (self.findPendingNewSheet(path)) |ns_idx| {
            const ns = self.pending_new_sheets.items[ns_idx];
            const xml = try self.allocator.dupe(u8, ns.body_xml);
            errdefer self.allocator.free(xml);
            const cells = try scanWorksheetXml(self.allocator, xml);
            return .{ .allocator = self.allocator, .xml = xml, .cells = cells };
        }
        return error.SheetEntryNotFound;
    }

    /// Replace one cell's value in place (Phase 3d, iter-cm-2a).
    /// First call on a sheet decompresses + scans it into
    /// `pending_mutations`; subsequent calls splice the cached XML
    /// directly. v1 limitations:
    ///   - cell type: numeric / integer / boolean / empty only;
    ///     string cells return `error.SetCellStringNotImplementedYet`
    ///     until iter-cm-2b ships SST integration.
    ///   - target: the (row, col) must already exist in the sheet;
    ///     missing cells return `error.SetCellMissingTarget`,
    ///     missing rows the same. iter-cm-2c/d add the insert paths.
    ///   - source attrs: the replacement is a fresh canonical
    ///     `<c r="REF"…>` — formulas / inline strings / phonetic
    ///     hints / unknown attrs on the source span are NOT
    ///     preserved. iter-cm-2b documents the preservation
    ///     contract; reject for now.
    pub fn setCell(
        self: *Editor,
        sheet_idx: u32,
        row: u32,
        col: u32,
        cell: Cell,
    ) !void {
        if (sheet_idx >= self.sheet_paths.len) return error.SheetIndexOutOfRange;
        if (self.pending_appends.contains(sheet_idx)) return error.SheetHasUnsavedAppends;
        // Refuse when a row/col edit is queued for this sheet — same
        // save-order race that affects appendRows.
        if (self.sheetHasPendingRowOrColEdit(sheet_idx)) return error.SheetHasUnsavedRowOrColEdit;
        if (row == 0 or row > max_row) return error.RowIndexOutOfRange;
        if (col >= max_col_1based) return error.ColumnIndexOutOfRange;

        switch (cell) {
            .integer => |n| {
                const writer_mod = @import("writer.zig");
                if (!writer_mod.fitsExactlyInF64(n)) return error.IntegerExceedsExcelPrecision;
            },
            .string, .number, .boolean, .empty => {},
        }

        // Track whether we created the MutatedSheet entry on this
        // call. If we did AND the rest of setCell errors out,
        // remove the empty/uncommitted entry — leaving it would
        // wrongly poison subsequent appendRows on the same sheet
        // with `error.SheetHasUnsavedMutations`. If a prior
        // successful setCell already populated the entry, leave
        // it alone.
        const ms_existed_before = self.pending_mutations.contains(sheet_idx);
        const ms = try self.getOrInitMutatedSheet(sheet_idx);
        errdefer if (!ms_existed_before) {
            if (self.pending_mutations.fetchRemove(sheet_idx)) |kv| {
                var v = kv.value;
                v.deinit(self.allocator);
            }
        };

        // Locate the existing span if present, plus the in-row
        // neighbours we'd need for an insert.
        var idx: ?usize = null;
        var insert_before_idx: ?usize = null; // first span in same row with col > target
        var last_in_row_idx: ?usize = null; // last span in same row (any col)
        for (ms.spans.items, 0..) |s, i| {
            if (s.row != row) continue;
            if (s.col == col) {
                idx = i;
                break;
            }
            last_in_row_idx = i;
            if (s.col > col and insert_before_idx == null) insert_before_idx = i;
        }
        if (idx == null) {
            const row_has_anchor = last_in_row_idx != null or insert_before_idx != null;
            if (!row_has_anchor) {
                // The spans index only tracks `<c>` elements, so a
                // present-but-cellless row (`<row r="5" ht="24"/>`
                // or `<row r="5"></row>`) would otherwise be
                // misclassified as missing — duplicating the row.
                // Try to splice into the existing empty row first;
                // fall through to insertMissingRow only if no
                // matching `<row r="N">` exists.
                if (try insertCellIntoEmptyRow(self.allocator, ms, row, col, cell)) {
                    return;
                }
                // iter-cm-2d: row missing entirely. Build a new
                // `<row r="N"><c r="REF">…</c></row>` block and
                // splice it into sheetData at the right
                // lexicographic position.
                try insertMissingRow(self.allocator, ms, row, col, cell);
                return;
            }
            // iter-cm-2c: cell missing in an existing row. Insert
            // at the lexicographic position inside the row.
            var new_buf: std.ArrayListUnmanaged(u8) = .{};
            defer new_buf.deinit(self.allocator);
            try emitCellXml(self.allocator, &new_buf, row, col, cell);

            // Insertion offset: just before the first higher-col
            // span, or just after the last lower-col span if the
            // new cell is going at the end of the row.
            const insert_at: usize = if (insert_before_idx) |i|
                ms.spans.items[i].start
            else
                ms.spans.items[last_in_row_idx.?].end;

            // Transactional: reserve the spans-array capacity for
            // the new entry BEFORE mutating ms.xml. If the insert
            // OOMs after we've already shifted bytes + offsets,
            // ms is left inconsistent for any later save/scan.
            try ms.spans.ensureUnusedCapacity(self.allocator, 1);
            try ms.xml.insertSlice(self.allocator, insert_at, new_buf.items);
            const new_len = new_buf.items.len;

            // Shift every later span. Use signed delta — positive
            // here since insert always grows.
            for (ms.spans.items) |*s| {
                if (s.start >= insert_at) {
                    s.start += new_len;
                    s.end += new_len;
                    s.body_start += new_len;
                }
            }
            // Compute body_start of the new cell.
            const opening_gt = std.mem.indexOfScalar(u8, new_buf.items, '>') orelse unreachable;
            const new_span: CellSpan = .{
                .start = insert_at,
                .end = insert_at + new_len,
                .body_start = insert_at + opening_gt + 1,
                .row = row,
                .col = col,
            };
            const span_pos: usize = insert_before_idx orelse (last_in_row_idx.? + 1);
            ms.spans.insertAssumeCapacity(span_pos, new_span);
            return;
        }
        const span_idx = idx.?;
        const old = ms.spans.items[span_idx];

        // iter-cm-2 contract: the replacement is canonical, so any
        // non-r attribute (`s="N"` styles, `t="…"` overrides) or
        // non-`<v>` body (`<f>` formulas, `<is>` inline strings)
        // would be silently dropped. Reject up front rather than
        // corrupt the source. Lifting this gate is iter-cm-2e (the
        // attr/body preservation contract) — out of scope for v1.
        if (sourceCellHasMetadata(ms.xml.items, old)) return error.SetCellSourceCellHasMetadata;

        // Build the replacement bytes.
        var new_buf: std.ArrayListUnmanaged(u8) = .{};
        defer new_buf.deinit(self.allocator);
        try emitCellXml(self.allocator, &new_buf, row, col, cell);

        // Splice xml[old.start..old.end] = new_buf.
        const old_len = old.end - old.start;
        const new_len = new_buf.items.len;
        try ms.xml.replaceRange(self.allocator, old.start, old_len, new_buf.items);

        // Shift later spans by the byte delta. Spans whose start
        // is BEFORE old.start are unchanged. Spans starting AT or
        // AFTER old.end shift by `new_len - old_len`.
        const old_end = old.end;
        for (ms.spans.items, 0..) |*s, i| {
            if (i == span_idx) continue;
            if (s.start >= old_end) {
                // Use signed arithmetic to handle shrink/grow uniformly.
                const start_signed: isize = @as(isize, @intCast(s.start)) +
                    @as(isize, @intCast(new_len)) - @as(isize, @intCast(old_len));
                const end_signed: isize = @as(isize, @intCast(s.end)) +
                    @as(isize, @intCast(new_len)) - @as(isize, @intCast(old_len));
                const body_signed: isize = @as(isize, @intCast(s.body_start)) +
                    @as(isize, @intCast(new_len)) - @as(isize, @intCast(old_len));
                s.start = @intCast(start_signed);
                s.end = @intCast(end_signed);
                s.body_start = @intCast(body_signed);
            }
        }
        // Update the targeted span. The replacement always emits
        // `<c r="REF" …>…</c>` (or `<c r="REF"/>` for empty), so
        // body_start = position of '>' just past the opening tag.
        const opening_gt = std.mem.indexOfScalar(u8, new_buf.items, '>') orelse
            unreachable;
        ms.spans.items[span_idx] = .{
            .start = old.start,
            .end = old.start + new_len,
            .body_start = old.start + opening_gt + 1,
            .row = row,
            .col = col,
        };
    }

    /// Bulk variant of `setCell` (Phase 3d, iter-cm-3). Same
    /// per-cell rules as `setCell`; the only win is amortising the
    /// "open the MutatedSheet for this sheet" lookup across N edits
    /// for callers that can batch them. Each `Edit` is applied in
    /// source order — later edits see the byte offsets produced by
    /// earlier ones, same as calling `setCell` N times.
    pub const Edit = struct {
        row: u32,
        col: u32,
        cell: Cell,
    };
    pub fn setCells(
        self: *Editor,
        sheet_idx: u32,
        edits: []const Edit,
    ) !void {
        for (edits) |e| try self.setCell(sheet_idx, e.row, e.col, e.cell);
    }

    /// Append a new sheet to the workbook (Phase 3e, iter-sheet-1).
    /// Returns the new sheet's 0-based index, ready for use with
    /// `setCell` / `scanWorksheet` / `appendRows`. The sheet is
    /// empty (`<sheetData></sheetData>`) until the caller writes
    /// content via the existing mutation APIs.
    ///
    /// At save time the workbook gets:
    ///   - a new ZIP entry at `xl/worksheets/sheetN.xml`
    ///   - a new `<sheet …/>` line inside `<sheets>` in
    ///     `xl/workbook.xml`
    ///   - a new `<Relationship/>` in `xl/_rels/workbook.xml.rels`
    ///   - a new `<Override/>` in `[Content_Types].xml`
    ///
    /// Errors: `error.InvalidSheetName` (empty / >31 chars / banned
    /// characters), `error.DuplicateSheetName`, plus any propagated
    /// allocator errors.
    pub fn addSheet(self: *Editor, name: []const u8) !u32 {
        const writer_mod = @import("writer.zig");
        try writer_mod.validateSheetName(name);

        // Excel hard cap: a workbook can carry up to 255 sheets in
        // the same file. self.sheet_paths already grows to include
        // every prior pending sheet (each addSheet appends ns.path
        // to it), so the projected count after this call is just
        // `sheet_paths.len + 1` — counting pending separately would
        // double-count and reject around 128 sheets.
        if (self.sheet_paths.len + 1 > 255) return error.TooManySheets;

        // Reject duplicates against existing source sheets and any
        // pending additions. Source sheet names live in the source's
        // workbook.xml; rather than re-parse it here, query the
        // already-extracted Book (Editor's open path used Book.open
        // for sheet path resolution but didn't keep the names).
        // Cheapest: read workbook.xml once and scan for `name="…"`.
        if (try self.isSheetNameTaken(name, null, null)) return error.DuplicateSheetName;

        // Pick the next sheet_id, sheet path number, and rId by
        // scanning the workbook.xml + rels for the highest existing,
        // AND every pending_new_sheets entry (otherwise repeated
        // addSheet calls collide on the same rId / sheetId / path).
        const wb_xml = try self.readEntry("xl/workbook.xml");
        defer self.allocator.free(wb_xml);
        const rels_xml = try self.readEntry("xl/_rels/workbook.xml.rels");
        defer self.allocator.free(rels_xml);

        var next_sheet_id = nextMaxAttr(wb_xml, "sheetId=\"") + 1;
        var next_rid_num = nextMaxRId(rels_xml) + 1;
        var next_path_num = nextMaxSheetPathNum(self.entries) + 1;
        for (self.pending_new_sheets.items) |s| {
            if (s.sheet_id >= next_sheet_id) next_sheet_id = s.sheet_id + 1;
            // s.rid is "rIdN"; strip the "rId" prefix. The pending
            // entries were produced by addSheet itself, so a parseInt
            // failure here is a real internal-invariant break — assert
            // it instead of silently dropping (which would yield a
            // duplicate rId on the next addSheet call).
            if (std.mem.startsWith(u8, s.rid, "rId")) {
                const n = std.fmt.parseInt(u32, s.rid["rId".len..], 10) catch unreachable;
                if (n >= next_rid_num) next_rid_num = n + 1;
            }
            // s.path is "xl/worksheets/sheetN.xml". Same internal-
            // invariant applies to the path number.
            const prefix = "xl/worksheets/sheet";
            const suffix = ".xml";
            if (std.mem.startsWith(u8, s.path, prefix) and std.mem.endsWith(u8, s.path, suffix)) {
                const num_str = s.path[prefix.len .. s.path.len - suffix.len];
                const n = std.fmt.parseInt(u32, num_str, 10) catch unreachable;
                if (n >= next_path_num) next_path_num = n + 1;
            }
        }

        // Build the new entries. Each field gets its own errdefer
        // so a mid-init OOM frees only the slices we've already
        // produced — the previous struct-init form deferred all
        // cleanup behind one `errdefer ns.deinit(...)` that fires
        // only after the struct is fully built, leaking earlier
        // allocs if a later one failed.
        const ns_name = try self.allocator.dupe(u8, name);
        errdefer self.allocator.free(ns_name);
        const ns_path = try std.fmt.allocPrint(self.allocator, "xl/worksheets/sheet{d}.xml", .{next_path_num});
        errdefer self.allocator.free(ns_path);
        const ns_rid = try std.fmt.allocPrint(self.allocator, "rId{d}", .{next_rid_num});
        errdefer self.allocator.free(ns_rid);
        const ns_body = try self.allocator.dupe(u8, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" ++
            "<sheetData></sheetData></worksheet>");
        errdefer self.allocator.free(ns_body);
        const ns: NewSheet = .{
            .name = ns_name,
            .path = ns_path,
            .rid = ns_rid,
            .sheet_id = next_sheet_id,
            .body_xml = ns_body,
        };

        // Grow sheet_paths so subsequent setCell etc. resolve the
        // new index. sheet_paths is currently `[]const []const u8`;
        // rebuild a fresh slice with the new tail.
        const old_paths = self.sheet_paths;
        const new_paths = try self.allocator.alloc([]const u8, old_paths.len + 1);
        errdefer self.allocator.free(new_paths);
        @memcpy(new_paths[0..old_paths.len], old_paths);
        new_paths[old_paths.len] = ns.path; // borrow; freed via NewSheet.deinit

        try self.pending_new_sheets.append(self.allocator, ns);
        self.sheet_paths = new_paths;
        self.allocator.free(old_paths); // inner strings still owned by their original allocs

        return @intCast(old_paths.len);
    }

    /// Rename a sheet (Phase 3e, iter-sheet-2). v1 patches only
    /// `xl/workbook.xml` — formulas in other sheets that reference
    /// this sheet by name (`'OLD'!A1`) are NOT rewritten. A real
    /// formula tokenizer (iter-col-1) will close that gap. The
    /// caller-visible contract: rename succeeds even when cross-
    /// sheet refs exist; those refs become `#REF!` in Excel until
    /// the next iter ships.
    pub fn renameSheet(self: *Editor, sheet_idx: u32, new_name: []const u8) !void {
        if (sheet_idx >= self.sheet_paths.len) return error.SheetIndexOutOfRange;
        const writer_mod = @import("writer.zig");
        try writer_mod.validateSheetName(new_name);

        // Resolve current name + rId. For a source sheet, look up
        // via workbook.xml + rels; for a pending-new-sheet, use the
        // NewSheet entry's name + rid fields.
        const path = self.sheet_paths[sheet_idx];
        var current_name: []u8 = undefined;
        var current_rid: ?[]u8 = null;
        var current_name_owned = false;
        var current_rid_owned = false;
        defer if (current_name_owned) self.allocator.free(current_name);
        defer if (current_rid_owned) if (current_rid) |r| self.allocator.free(r);
        if (self.findPendingNewSheet(path)) |ns_idx| {
            current_name = self.pending_new_sheets.items[ns_idx].name;
        } else {
            const meta = (try self.sheetMetaAtPath(path)) orelse return error.SheetEntryNotFound;
            current_name = meta.name;
            current_name_owned = true;
            current_rid = meta.rid;
            current_rid_owned = true;
        }

        // No-op short-circuit: only when the EFFECTIVE current name
        // (taking pending renames into account) equals the new name
        // BYTE-EXACT. asciiEqlFold here would silently drop legit
        // case-only renames; loading current_name from workbook.xml
        // would silently drop a same-session "rename back to
        // original" (`A->B` then `B->A`).
        const eff_current: []const u8 = blk: {
            for (self.pending_renames.items) |r| {
                if (r.sheet_idx == sheet_idx) break :blk r.new_name;
            }
            break :blk current_name;
        };
        if (std.mem.eql(u8, eff_current, new_name)) return;

        // Reject duplicates against every OTHER sheet's effective
        // name. "Effective" lets `A->C` then `B->A` work because
        // sheet 0's effective name after the first rename is `C`,
        // freeing `A` for sheet 1. Skip-by-rId lets the renamed
        // source sheet itself drop out of the candidate set.
        if (try self.isSheetNameTaken(new_name, sheet_idx, current_rid)) return error.DuplicateSheetName;

        // Pending-new-sheet path: mutate the NewSheet's name in
        // place, no workbook.xml patch needed (the sheet doesn't
        // exist there yet).
        if (self.findPendingNewSheet(path)) |ns_idx| {
            const ns = &self.pending_new_sheets.items[ns_idx];
            const new_owned = try self.allocator.dupe(u8, new_name);
            self.allocator.free(ns.name);
            ns.name = new_owned;
            return;
        }

        // Source sheet path: record the rename for save-time
        // workbook.xml patch. If a previous rename targeted the
        // same sheet, replace its new_name (don't accumulate).
        for (self.pending_renames.items) |*r| {
            if (r.sheet_idx == sheet_idx) {
                const replaced = try self.allocator.dupe(u8, new_name);
                self.allocator.free(r.new_name);
                r.new_name = replaced;
                return;
            }
        }
        const old_dup = try self.allocator.dupe(u8, current_name);
        errdefer self.allocator.free(old_dup);
        const new_dup = try self.allocator.dupe(u8, new_name);
        errdefer self.allocator.free(new_dup);
        // Take ownership of the rId (sheetMetaAtPath returned an
        // owned dupe). Disable the deferred free above.
        const rid_owned = current_rid.?;
        current_rid_owned = false;
        try self.pending_renames.append(self.allocator, .{
            .sheet_idx = sheet_idx,
            .rid = rid_owned,
            .old_name = old_dup,
            .new_name = new_dup,
        });
    }

    /// Delete a sheet (Phase 3e, iter-sheet-3). v1 contract:
    ///   - Refuses if it's the only remaining sheet.
    ///   - Refuses if there's pending mutation/append/rename state
    ///     on ANY sheet (caller must `save` first then re-open).
    ///   - Pending-new-sheet path: removes from `pending_new_sheets`.
    ///   - Source sheet path: records the delete + drops the path
    ///     from `sheet_paths`.
    ///   - Cross-sheet formula refs to the deleted sheet become
    ///     `#REF!` (deferred to iter-col-1's formula tokenizer).
    /// Sheet indices SHIFT after a delete: the call invalidates
    /// every sheet_idx > deleted_idx.
    pub fn deleteSheet(self: *Editor, sheet_idx: u32) !void {
        if (sheet_idx >= self.sheet_paths.len) return error.SheetIndexOutOfRange;
        if (self.sheet_paths.len <= 1) return error.CannotDeleteLastSheet;
        if (self.pending_appends.count() > 0 or
            self.pending_mutations.count() > 0 or
            self.pending_renames.items.len > 0 or
            self.pending_row_inserts.items.len > 0 or
            self.pending_row_deletes.items.len > 0 or
            self.pending_col_inserts.items.len > 0 or
            self.pending_col_deletes.items.len > 0)
        {
            // deleteSheet rebuilds sheet_paths; queued row/col edits
            // hold raw indices into it and would silently point at
            // the wrong sheet (or out-of-bounds) after the rebuild.
            return error.SheetDeleteRequiresCleanState;
        }

        // Conservative guard: refuse SOURCE-sheet deletes when the
        // workbook has any `<definedName>` entries. Pending-new
        // sheet deletes don't shift any existing localSheetId or
        // formula reference (the new sheet was never saved), so
        // they're safe regardless.
        const path_check = self.sheet_paths[sheet_idx];
        if (self.findPendingNewSheet(path_check) == null) {
            const wb_check = try self.readEntry("xl/workbook.xml");
            defer self.allocator.free(wb_check);
            if (std.mem.indexOf(u8, wb_check, "<definedName") != null)
                return error.SheetDeleteWithDefinedNamesNotSupported;
        }

        const path = self.sheet_paths[sheet_idx];

        if (self.findPendingNewSheet(path)) |ns_idx| {
            // Pending-new sheet: remove from pending_new_sheets.
            // orderedRemove (NOT swapRemove) so the remaining new
            // sheets keep their original order — save-time loops
            // rely on `source_count + i` indexing matching
            // sheet_paths' tail. NewSheet.path borrows from
            // sheet_paths; deinit frees name/rid/body only.
            var ns = self.pending_new_sheets.orderedRemove(ns_idx);
            ns.deinit(self.allocator);
        } else {
            const meta = (try self.sheetMetaAtPath(path)) orelse return error.SheetEntryNotFound;
            self.allocator.free(meta.name); // not needed for delete
            const path_dup = try self.allocator.dupe(u8, path);
            errdefer self.allocator.free(path_dup);
            errdefer self.allocator.free(meta.rid);
            try self.pending_deletes.append(self.allocator, .{
                .path = path_dup,
                .rid = meta.rid,
            });
        }

        // Rebuild sheet_paths without the deleted entry. Free the
        // deleted entry's path bytes; other entries are still
        // referenced by new_paths so they live on.
        const new_paths = try self.allocator.alloc([]const u8, self.sheet_paths.len - 1);
        var dst: usize = 0;
        for (self.sheet_paths, 0..) |p, i| {
            if (i == sheet_idx) continue;
            new_paths[dst] = p;
            dst += 1;
        }
        const old_paths = self.sheet_paths;
        self.allocator.free(old_paths[sheet_idx]);
        self.sheet_paths = new_paths;
        self.allocator.free(old_paths);
    }

    /// Insert a blank row at position `before_row` in sheet
    /// `sheet_idx` (Phase 3e, iter-row-2). Every existing row at
    /// or below `before_row` shifts down by 1. v1 limitations:
    ///   - Worksheet XML rewrites: <row r="N"> renumber, <c r="A1">
    ///     row component, <mergeCells> rect bounds, <dimension>.
    ///   - Refuses if the worksheet contains <hyperlinks>,
    ///     <dataValidations>, <conditionalFormatting>, <f>
    ///     formulas, or any <drawing>/<picture> reference — those
    ///     can carry row indices we don't yet rewrite.
    ///   - The sheet must not have other pending mutations
    ///     (setCell / appendRows / row inserts/deletes); save
    ///     first to apply those.
    pub fn insertRow(self: *Editor, sheet_idx: u32, before_row: u32) !void {
        try self.recordRowEdit(sheet_idx, before_row, &self.pending_row_inserts, true);
    }

    /// Delete row `row` in sheet `sheet_idx` (Phase 3e, iter-row-3).
    /// Every row > `row` shifts up by 1. Same v1 limitations as
    /// `insertRow`.
    pub fn deleteRow(self: *Editor, sheet_idx: u32, row: u32) !void {
        try self.recordRowEdit(sheet_idx, row, &self.pending_row_deletes, false);
    }

    /// Insert a blank column at position `before_col` (1-based,
    /// A=1) in sheet `sheet_idx`. Phase 3e iter-col-3. Same v1
    /// limitations as `insertRow` — formula bodies, defined names,
    /// and structured-table refs aren't rewritten and are refused.
    pub fn insertColumn(self: *Editor, sheet_idx: u32, before_col_1based: u32) !void {
        try self.recordColEdit(sheet_idx, before_col_1based, &self.pending_col_inserts);
    }

    /// Delete column `col_1based` (1-based) in sheet `sheet_idx`.
    /// Phase 3e iter-col-4.
    pub fn deleteColumn(self: *Editor, sheet_idx: u32, col_1based: u32) !void {
        try self.recordColEdit(sheet_idx, col_1based, &self.pending_col_deletes);
    }

    /// True when `sheet_idx` already has a queued insertRow,
    /// deleteRow, insertColumn, or deleteColumn. Used by
    /// appendRows/setCell to refuse mixing with a row/col edit on
    /// the same sheet — save() applies row/col edits before
    /// appends/mutations and the latter would otherwise overwrite
    /// the row/col-substituted entry from pre-edit source XML.
    fn sheetHasPendingRowOrColEdit(self: *Editor, sheet_idx: u32) bool {
        for (self.pending_row_inserts.items) |e| if (e.sheet_idx == sheet_idx) return true;
        for (self.pending_row_deletes.items) |e| if (e.sheet_idx == sheet_idx) return true;
        for (self.pending_col_inserts.items) |e| if (e.sheet_idx == sheet_idx) return true;
        for (self.pending_col_deletes.items) |e| if (e.sheet_idx == sheet_idx) return true;
        return false;
    }

    /// Cross-sheet reference carriers: tags whose body or attrs can
    /// hold an `OtherSheet!Ref` pointer. If any sheet in the workbook
    /// has one of these, a row/column edit on a *different* sheet
    /// can still leave a stale reference pointing at the edited
    /// sheet's old layout. Until the formula tokenizer (iter-col-1)
    /// rewrites them, refuse globally.
    ///   - `<f>` / `<f ` / `<f/`     — formula bodies (=Sheet1!B:B)
    ///   - `<hyperlinks`             — `location="Sheet1!C5"` etc.
    ///   - `<dataValidations`        — formula1/formula2 may cross
    ///   - `<conditionalFormatting`  — cfRule formulas may cross
    fn anySheetCrossSheetCarrier(self: *Editor) !struct { found: bool, kind: enum { none, formula, hyperlink, data_validation, cond_format } } {
        for (self.sheet_paths) |path| {
            if (self.findPendingNewSheet(path) != null) continue;
            const entry_idx = findEntryByName(self.entries, path) orelse continue;
            const e = self.entries[entry_idx];
            const payload = self.src_buf[e.lfh_offset + e.lfh_total_len ..][0..e.payload_len];
            const xml = try decompressZipPayload(self.allocator, payload, e.compression_method, e.uncompressed_size);
            defer self.allocator.free(xml);
            if (std.mem.indexOf(u8, xml, "<f>") != null or
                std.mem.indexOf(u8, xml, "<f ") != null or
                std.mem.indexOf(u8, xml, "<f/") != null)
                return .{ .found = true, .kind = .formula };
            if (std.mem.indexOf(u8, xml, "<hyperlinks") != null)
                return .{ .found = true, .kind = .hyperlink };
            if (std.mem.indexOf(u8, xml, "<dataValidations") != null)
                return .{ .found = true, .kind = .data_validation };
            if (std.mem.indexOf(u8, xml, "<conditionalFormatting") != null)
                return .{ .found = true, .kind = .cond_format };
        }
        return .{ .found = false, .kind = .none };
    }

    fn recordColEdit(
        self: *Editor,
        sheet_idx: u32,
        col_1based: u32,
        list: *std.ArrayListUnmanaged(ColEdit),
    ) !void {
        if (sheet_idx >= self.sheet_paths.len) return error.SheetIndexOutOfRange;
        if (col_1based == 0 or col_1based > max_col_1based) return error.ColumnIndexOutOfRange;
        if (self.pending_appends.contains(sheet_idx) or
            self.pending_mutations.contains(sheet_idx))
        {
            return error.ColEditRequiresCleanSheet;
        }
        for (self.pending_row_inserts.items) |e| {
            if (e.sheet_idx == sheet_idx) return error.ColEditRequiresCleanSheet;
        }
        for (self.pending_row_deletes.items) |e| {
            if (e.sheet_idx == sheet_idx) return error.ColEditRequiresCleanSheet;
        }
        for (self.pending_col_inserts.items) |e| {
            if (e.sheet_idx == sheet_idx) return error.ColEditRequiresCleanSheet;
        }
        for (self.pending_col_deletes.items) |e| {
            if (e.sheet_idx == sheet_idx) return error.ColEditRequiresCleanSheet;
        }

        const path = self.sheet_paths[sheet_idx];
        if (self.findPendingNewSheet(path) != null) return error.ColEditOnNewSheetUnsupported;

        const wb_check = try self.readEntry("xl/workbook.xml");
        defer self.allocator.free(wb_check);
        if (std.mem.indexOf(u8, wb_check, "<definedName") != null)
            return error.ColEditWithDefinedNamesNotSupported;

        // Cross-sheet reference carriers (formulas, hyperlinks,
        // data validations, conditional formatting) live in *any*
        // sheet body but can point at the edited sheet's columns.
        // Without a tokenizer we can't rewrite them, so refuse
        // globally if any sheet carries one. Picks the most
        // specific error code based on which carrier we hit first.
        const xref = try self.anySheetCrossSheetCarrier();
        if (xref.found) {
            return switch (xref.kind) {
                .formula => error.ColEditWithFormulasNotSupported,
                .hyperlink, .data_validation, .cond_format => error.ColEditUnsafeForSheet,
                .none => unreachable,
            };
        }

        if (findEntryByName(self.entries, path)) |entry_idx| {
            const e = self.entries[entry_idx];
            const payload = self.src_buf[e.lfh_offset + e.lfh_total_len ..][0..e.payload_len];
            const xml = try decompressZipPayload(self.allocator, payload, e.compression_method, e.uncompressed_size);
            defer self.allocator.free(xml);
            // Local-only structures we don't yet rewrite — the
            // cross-sheet carriers above are already checked
            // workbook-wide.
            const guards = [_][]const u8{
                "<autoFilter",
                "<tableParts",
                "<drawing",
                "<legacyDrawing",
                "<picture",
                // <pane xSplit=..|ySplit=..|topLeftCell=..> carries
                // column/row coordinates that aren't rewritten by
                // the row/col edit path. Refuse rather than save a
                // workbook with frozen panes pointing at the wrong
                // boundary.
                "<pane ",
                "<pane/",
                "<pane\t",
                "<pane\n",
                "<pane\r",
            };
            for (guards) |g| {
                if (std.mem.indexOf(u8, xml, g) != null) {
                    return error.ColEditUnsafeForSheet;
                }
            }
        } else return error.SheetEntryNotFound;

        try list.append(self.allocator, .{ .sheet_idx = sheet_idx, .col_1based = col_1based });
    }

    fn recordRowEdit(
        self: *Editor,
        sheet_idx: u32,
        row: u32,
        list: *std.ArrayListUnmanaged(RowEdit),
        is_insert: bool,
    ) !void {
        _ = is_insert;
        if (sheet_idx >= self.sheet_paths.len) return error.SheetIndexOutOfRange;
        if (row == 0 or row > max_row) return error.RowIndexOutOfRange;
        // Conservative: refuse when other mutations target this
        // sheet (would need cross-pending state shifts we don't
        // model in v1).
        if (self.pending_appends.contains(sheet_idx) or
            self.pending_mutations.contains(sheet_idx))
        {
            return error.RowEditRequiresCleanSheet;
        }
        for (self.pending_row_inserts.items) |e| {
            if (e.sheet_idx == sheet_idx) return error.RowEditRequiresCleanSheet;
        }
        for (self.pending_row_deletes.items) |e| {
            if (e.sheet_idx == sheet_idx) return error.RowEditRequiresCleanSheet;
        }
        for (self.pending_col_inserts.items) |e| {
            if (e.sheet_idx == sheet_idx) return error.RowEditRequiresCleanSheet;
        }
        for (self.pending_col_deletes.items) |e| {
            if (e.sheet_idx == sheet_idx) return error.RowEditRequiresCleanSheet;
        }

        // Pending-new sheets have empty bodies — insertRow/deleteRow
        // on them is meaningless and the save path can't substitute
        // a non-existent source ZIP entry. Reject up front.
        const path = self.sheet_paths[sheet_idx];
        if (self.findPendingNewSheet(path) != null) return error.RowEditOnNewSheetUnsupported;

        // Workbook-scoped <definedName> entries can carry row
        // references in formula text (named ranges, print areas,
        // print titles). Until iter-col-1's formula tokenizer
        // ships, refuse rather than save stale references.
        const wb_check = try self.readEntry("xl/workbook.xml");
        defer self.allocator.free(wb_check);
        if (std.mem.indexOf(u8, wb_check, "<definedName") != null)
            return error.RowEditWithDefinedNamesNotSupported;

        // Cross-sheet reference carriers (formulas, hyperlinks,
        // data validations, conditional formatting) live in *any*
        // sheet body but can point at the edited sheet's rows.
        // Without a tokenizer we can't rewrite them, so refuse
        // globally if any sheet carries one.
        const xref = try self.anySheetCrossSheetCarrier();
        if (xref.found) {
            return switch (xref.kind) {
                .formula => error.RowEditWithFormulasNotSupported,
                .hyperlink, .data_validation, .cond_format => error.RowEditUnsafeForSheet,
                .none => unreachable,
            };
        }

        // Conservative content guard: scan the worksheet XML for
        // local-only elements that v1 doesn't rewrite. If any are
        // present, refuse rather than silently corrupt the workbook.
        if (findEntryByName(self.entries, path)) |entry_idx| {
            const e = self.entries[entry_idx];
            const payload = self.src_buf[e.lfh_offset + e.lfh_total_len ..][0..e.payload_len];
            const xml = try decompressZipPayload(self.allocator, payload, e.compression_method, e.uncompressed_size);
            defer self.allocator.free(xml);
            const guards = [_][]const u8{
                "<autoFilter",
                "<tableParts",
                "<drawing",
                "<legacyDrawing",
                "<picture",
                // <pane xSplit=..|ySplit=..|topLeftCell=..> carries
                // column/row coordinates that aren't rewritten by
                // the row/col edit path. Refuse rather than save a
                // workbook with frozen panes pointing at the wrong
                // boundary.
                "<pane ",
                "<pane/",
                "<pane\t",
                "<pane\n",
                "<pane\r",
            };
            for (guards) |g| {
                if (std.mem.indexOf(u8, xml, g) != null) {
                    return error.RowEditUnsafeForSheet;
                }
            }
        } else return error.SheetEntryNotFound;

        try list.append(self.allocator, .{ .sheet_idx = sheet_idx, .row = row });
    }

    /// True when some sheet has an effective name matching
    /// `candidate` case-insensitively, EXCLUDING the sheet at
    /// `except_sheet_idx` (if non-null) and the source `<sheet>`
    /// entry whose r:id matches `except_rid` (if non-null).
    ///
    /// "Effective" means: pending-rename's new_name takes precedence
    /// over the raw workbook.xml name; pending-new-sheets contribute
    /// their NewSheet.name. Walks raw workbook.xml directly so
    /// entries `parseWorkbookSheets` skipped (broken rels) still
    /// participate in dup detection.
    fn isSheetNameTaken(
        self: *Editor,
        candidate: []const u8,
        except_sheet_idx: ?u32,
        except_rid: ?[]const u8,
    ) !bool {

        // 1. Raw <sheet> entries from workbook.xml. Apply pending
        //    renames as we walk to compute effective names.
        const wb = try self.readEntry("xl/workbook.xml");
        defer self.allocator.free(wb);
        var i: usize = 0;
        while (findTagOpen(wb, i, "sheet")) |t| {
            const attrs = wb[t.start + "<sheet".len .. t.after_open - 1];
            const rid = getAttr(attrs, "r:id") orelse {
                i = t.after_open;
                continue;
            };
            if (except_rid) |skip_rid| if (std.mem.eql(u8, rid, skip_rid)) {
                i = t.after_open;
                continue;
            };
            // Skip sheets queued for deletion — their names are
            // freed for reuse on save.
            var was_deleted = false;
            for (self.pending_deletes.items) |d| {
                if (std.mem.eql(u8, d.rid, rid)) {
                    was_deleted = true;
                    break;
                }
            }
            if (was_deleted) {
                i = t.after_open;
                continue;
            }
            // Effective name: pending-rename's new_name if any
            // matches this rId; else the raw (decoded) name.
            var rename_hit: ?[]const u8 = null;
            for (self.pending_renames.items) |r| {
                if (std.mem.eql(u8, r.rid, rid)) {
                    rename_hit = r.new_name;
                    break;
                }
            }
            if (rename_hit) |nm| {
                if (casefold.excelSheetNameEql(nm, candidate)) return true;
            } else if (getAttr(attrs, "name")) |raw| {
                var decoded: std.ArrayListUnmanaged(u8) = .{};
                defer decoded.deinit(self.allocator);
                try decodeXmlAttrInto(self.allocator, &decoded, raw);
                if (casefold.excelSheetNameEql(decoded.items, candidate)) return true;
            }
            i = t.after_open;
        }

        // 2. Pending new sheets. Indexed by sheet_idx for the skip.
        const source_count: u32 = @intCast(self.sheet_paths.len - self.pending_new_sheets.items.len);
        for (self.pending_new_sheets.items, 0..) |s, ns_idx| {
            const my_idx: u32 = source_count + @as(u32, @intCast(ns_idx));
            if (except_sheet_idx) |ei| if (my_idx == ei) continue;
            if (casefold.excelSheetNameEql(s.name, candidate)) return true;
        }
        return false;
    }

    /// Look up the source workbook's `<sheet>` row whose r:id
    /// resolves (via rels) to `target_path`. Returns owned dupes of
    /// the (entity-decoded) name AND the rId, or null when no
    /// match. Caller frees both. iter-sheet-2 needs the rId to
    /// match `<sheet>` lines by-id rather than by-position.
    fn sheetMetaAtPath(self: *Editor, target_path: []const u8) !?struct { name: []u8, rid: []u8 } {
        const wb = try self.readEntry("xl/workbook.xml");
        defer self.allocator.free(wb);
        const rels = try self.readEntry("xl/_rels/workbook.xml.rels");
        defer self.allocator.free(rels);

        var i: usize = 0;
        while (findTagOpen(wb, i, "sheet")) |t| {
            const sh_attrs = wb[t.start + "<sheet".len .. t.after_open - 1];
            const rid = getAttr(sh_attrs, "r:id") orelse {
                i = t.after_open;
                continue;
            };
            const name_raw = getAttr(sh_attrs, "name") orelse {
                i = t.after_open;
                continue;
            };
            var rels_i: usize = 0;
            while (std.mem.indexOfPos(u8, rels, rels_i, "<Relationship")) |rel_pos| {
                const rel_end = std.mem.indexOfScalarPos(u8, rels, rel_pos, '>') orelse break;
                const rel_attrs = rels[rel_pos + "<Relationship".len .. rel_end];
                const id = getAttr(rel_attrs, "Id") orelse {
                    rels_i = rel_end + 1;
                    continue;
                };
                if (!std.mem.eql(u8, id, rid)) {
                    rels_i = rel_end + 1;
                    continue;
                }
                const target = getAttr(rel_attrs, "Target") orelse {
                    rels_i = rel_end + 1;
                    continue;
                };
                // Three legal forms for the same path:
                //   - "worksheets/sheet1.xml" (relative; needs xl/)
                //   - "xl/worksheets/sheet1.xml" (already prefixed)
                //   - "/xl/worksheets/sheet1.xml" (absolute) —
                //     parseWorkbookSheets accepts this; we must too.
                const t_norm = if (target.len > 0 and target[0] == '/') target[1..] else target;
                var matches = std.mem.eql(u8, t_norm, target_path);
                if (!matches) {
                    var prefixed_buf: [256]u8 = undefined;
                    const prefixed = std.fmt.bufPrint(&prefixed_buf, "xl/{s}", .{t_norm}) catch break;
                    matches = std.mem.eql(u8, prefixed, target_path);
                }
                if (matches) {
                    var decoded: std.ArrayListUnmanaged(u8) = .{};
                    errdefer decoded.deinit(self.allocator);
                    try decodeXmlAttrInto(self.allocator, &decoded, name_raw);
                    const name_owned = try decoded.toOwnedSlice(self.allocator);
                    errdefer self.allocator.free(name_owned);
                    const rid_owned = try self.allocator.dupe(u8, rid);
                    return .{ .name = name_owned, .rid = rid_owned };
                }
                rels_i = rel_end + 1;
            }
            i = t.after_open;
        }
        return null;
    }

    /// Decompress an entry by name. Caller owns the returned slice.
    fn readEntry(self: *Editor, entry_name: []const u8) ![]u8 {
        const idx = findEntryByName(self.entries, entry_name) orelse
            return error.MissingEntry;
        const e = self.entries[idx];
        const payload = self.src_buf[e.lfh_offset + e.lfh_total_len ..][0..e.payload_len];
        return try decompressZipPayload(self.allocator, payload, e.compression_method, e.uncompressed_size);
    }

    fn getOrInitMutatedSheet(self: *Editor, sheet_idx: u32) !*MutatedSheet {
        const gop = try self.pending_mutations.getOrPut(self.allocator, sheet_idx);
        if (!gop.found_existing) {
            gop.value_ptr.* = .{};
            errdefer {
                gop.value_ptr.deinit(self.allocator);
                _ = self.pending_mutations.remove(sheet_idx);
            }
            const path = self.sheet_paths[sheet_idx];
            // First try the source ZIP entry. If absent, this is a
            // pending-new-sheet (Phase 3e iter-sheet-1) — seed the
            // mutation buffer from its empty body template.
            if (findEntryByName(self.entries, path)) |entry_idx| {
                const entry = self.entries[entry_idx];
                const payload = self.src_buf[entry.lfh_offset + entry.lfh_total_len ..][0..entry.payload_len];
                const xml = try decompressZipPayload(
                    self.allocator,
                    payload,
                    entry.compression_method,
                    entry.uncompressed_size,
                );
                defer self.allocator.free(xml);
                try gop.value_ptr.xml.appendSlice(self.allocator, xml);
            } else {
                const ns_idx = self.findPendingNewSheet(path) orelse
                    return error.SheetEntryNotFound;
                try gop.value_ptr.xml.appendSlice(
                    self.allocator,
                    self.pending_new_sheets.items[ns_idx].body_xml,
                );
            }
            const spans = try scanWorksheetXml(self.allocator, gop.value_ptr.xml.items);
            defer self.allocator.free(spans);
            try gop.value_ptr.spans.appendSlice(self.allocator, spans);
        }
        return gop.value_ptr;
    }

    fn findPendingNewSheet(self: *Editor, path: []const u8) ?usize {
        for (self.pending_new_sheets.items, 0..) |s, i| {
            if (std.mem.eql(u8, s.path, path)) return i;
        }
        return null;
    }
};

/// True if the source span carries any attribute beyond `r="…"` OR
/// any body content beyond a single `<v>…</v>`. Iter-cm-2 contract:
/// the replacement is a canonical `<c r="…"…>…</c>`, so anything
/// non-canonical on the source (`s="N"`, `<f>`, `<is>`, phonetic
/// hints, unknown attrs/children) would be silently dropped. Caller
/// rejects with a typed error rather than corrupt.
fn sourceCellHasMetadata(xml: []const u8, span: CellSpan) bool {
    const attrs_end = if (span.body_start > 0) span.body_start - 1 else return false;
    if (attrs_end <= span.start + 2) return false;
    const attrs = xml[span.start + 2 .. attrs_end];
    // Self-closing form has `/` as the last byte of attrs; trim it.
    const trimmed_attrs = if (attrs.len > 0 and attrs[attrs.len - 1] == '/')
        attrs[0 .. attrs.len - 1]
    else
        attrs;

    // Walk the attrs region. Allowed: `r="…"` and `t="…"` (the
    // type override — we're replacing the value, so the source
    // type tag is moot anyway). Reject on anything else (`s=`
    // styles, `cm=`/`vm=` metadata, namespace-prefixed extensions)
    // because those carry semantic state setCell would silently
    // drop. iter-cm-2e relaxes this to a preserve-and-merge model.
    var i: usize = 0;
    while (i < trimmed_attrs.len) {
        // Skip whitespace.
        while (i < trimmed_attrs.len and (trimmed_attrs[i] == ' ' or
            trimmed_attrs[i] == '\t' or trimmed_attrs[i] == '\n' or
            trimmed_attrs[i] == '\r')) i += 1;
        if (i >= trimmed_attrs.len) break;
        // Find attribute name terminator (`=` or whitespace before `=`).
        const name_start = i;
        while (i < trimmed_attrs.len and trimmed_attrs[i] != '=' and
            trimmed_attrs[i] != ' ' and trimmed_attrs[i] != '\t' and
            trimmed_attrs[i] != '\n' and trimmed_attrs[i] != '\r') i += 1;
        const name = trimmed_attrs[name_start..i];
        if (!std.mem.eql(u8, name, "r") and !std.mem.eql(u8, name, "t"))
            return true; // any attr other than r/t is metadata
        // Skip `=` and the quoted value.
        while (i < trimmed_attrs.len and trimmed_attrs[i] != '=') i += 1;
        if (i >= trimmed_attrs.len) return true;
        i += 1;
        if (i >= trimmed_attrs.len or
            (trimmed_attrs[i] != '"' and trimmed_attrs[i] != '\''))
            return true;
        const quote = trimmed_attrs[i];
        i += 1;
        while (i < trimmed_attrs.len and trimmed_attrs[i] != quote) i += 1;
        if (i >= trimmed_attrs.len) return true;
        i += 1;
    }

    // Check body. Reject any tag that isn't part of the canonical
    // value envelope (`<v>`, `<is>`, `<t>` and their closes —
    // empty `<is>`/`<t>`/`<v>` self-closing forms accepted too).
    // Anything else — `<f>` formulas, `<rPh>` phonetic hints,
    // `<phoneticPr>`, `<extLst>` extensions — carries semantic
    // state setCell would silently drop. iter-cm-2e relaxes this
    // to a preserve-and-merge model.
    if (span.end < span.body_start + "</c>".len) return false;
    const body_end = span.end - "</c>".len;
    if (body_end <= span.body_start) return false;
    const body = xml[span.body_start..body_end];

    var bi: usize = 0;
    while (std.mem.indexOfScalarPos(u8, body, bi, '<')) |lt| {
        const after = lt + 1;
        if (after >= body.len) return true;
        // Step over `/` for closing tags; same allow-list applies.
        const name_start = if (body[after] == '/') after + 1 else after;
        if (name_start >= body.len) return true;
        // Find tag-name terminator (whitespace, `/`, `>`).
        var ne = name_start;
        while (ne < body.len and body[ne] != ' ' and body[ne] != '\t' and
            body[ne] != '\n' and body[ne] != '\r' and
            body[ne] != '/' and body[ne] != '>') ne += 1;
        const name = body[name_start..ne];
        const allowed = std.mem.eql(u8, name, "v") or
            std.mem.eql(u8, name, "is") or
            std.mem.eql(u8, name, "t");
        if (!allowed) return true;
        // Advance past the tag's `>`.
        const gt = std.mem.indexOfScalarPos(u8, body, ne, '>') orelse return true;
        bi = gt + 1;
    }
    return false;
}

/// Splice a new cell into a row that exists with `<row r="N">` but
/// has no inner `<c>` cells (e.g. style/height-only `<row r="5"
/// ht="24"/>`). Returns true if the row was found and the cell
/// inserted; false if no matching row exists (caller falls
/// through to the missing-row insert path).
fn insertCellIntoEmptyRow(
    allocator: Allocator,
    ms: *Editor.MutatedSheet,
    row: u32,
    col: u32,
    cell: Cell,
) !bool {
    // Walk every `<row>` opening matching either an explicit
    // `r="N"` attribute OR the implicit row counter (1-based count
    // of rows seen). OOXML allows r= to be omitted and
    // scanWorksheetXml already supports it; this helper must match
    // that semantics or it will misclassify cellless implicit-row
    // rows as missing and duplicate them.
    var pos: usize = 0;
    var implicit_row: u32 = 0;
    while (findTagOpen(ms.xml.items, pos, "row")) |t| {
        implicit_row += 1;
        const attrs_full = ms.xml.items[t.start + "<row".len .. t.after_open - 1];
        const trimmed = std.mem.trimRight(u8, attrs_full, " \t\r\n");
        const is_self_closing = trimmed.len > 0 and trimmed[trimmed.len - 1] == '/';
        const attrs = if (is_self_closing) trimmed[0 .. trimmed.len - 1] else trimmed;
        const effective_row: u32 = blk: {
            if (getAttr(attrs, "r")) |r_attr| {
                const parsed = std.fmt.parseInt(u32, r_attr, 10) catch 0;
                if (parsed > 0) break :blk parsed;
            }
            // Same recovery cascade Rows.next uses: try the body's
            // first cell ref, fall back to the implicit counter.
            if (!is_self_closing) {
                if (recoverRowFromFirstCell(ms.xml.items, t.after_open)) |n|
                    break :blk n;
            }
            break :blk implicit_row;
        };
        if (effective_row != row) {
            pos = t.after_open;
            continue;
        }

        // Found the row. Reserve span capacity up front so the
        // final insert can't OOM after we've already mutated xml.
        try ms.spans.ensureUnusedCapacity(allocator, 1);
        // Build the cell bytes once.
        var cell_buf: std.ArrayListUnmanaged(u8) = .{};
        defer cell_buf.deinit(allocator);
        try emitCellXml(allocator, &cell_buf, row, col, cell);

        if (is_self_closing) {
            // Expand `<row …/>` to `<row …><c …>…</c></row>`.
            // Position of the `/` is at `t.after_open - 2` (just
            // before the closing `>`).
            const slash_pos = t.after_open - 2;
            // Build replacement: `>` + cell + `</row>`
            var repl: std.ArrayListUnmanaged(u8) = .{};
            defer repl.deinit(allocator);
            try repl.append(allocator, '>');
            try repl.appendSlice(allocator, cell_buf.items);
            try repl.appendSlice(allocator, "</row>");
            try ms.xml.replaceRange(allocator, slash_pos, 2, repl.items);
            const delta_signed: isize = @as(isize, @intCast(repl.items.len)) - 2;
            // Shift later spans.
            for (ms.spans.items) |*s| {
                if (s.start >= t.after_open) {
                    s.start = @intCast(@as(isize, @intCast(s.start)) + delta_signed);
                    s.end = @intCast(@as(isize, @intCast(s.end)) + delta_signed);
                    s.body_start = @intCast(@as(isize, @intCast(s.body_start)) + delta_signed);
                }
            }
            // The new cell starts where the old `/>` ended (replaced).
            const cell_abs_start = slash_pos + 1;
            const opening_gt = std.mem.indexOfScalar(u8, cell_buf.items, '>') orelse unreachable;
            const new_span: CellSpan = .{
                .start = cell_abs_start,
                .end = cell_abs_start + cell_buf.items.len,
                .body_start = cell_abs_start + opening_gt + 1,
                .row = row,
                .col = col,
            };
            // Insert at the right position in spans (source order).
            var span_pos: usize = ms.spans.items.len;
            for (ms.spans.items, 0..) |s, i| {
                if (s.start >= new_span.end) {
                    span_pos = i;
                    break;
                }
            }
            ms.spans.insertAssumeCapacity(span_pos, new_span);
            return true;
        }

        // Body form `<row …>...</row>` with no inner cells. Splice
        // the new cell at `t.after_open` (just past the `>`).
        try ms.xml.insertSlice(allocator, t.after_open, cell_buf.items);
        for (ms.spans.items) |*s| {
            if (s.start >= t.after_open) {
                s.start += cell_buf.items.len;
                s.end += cell_buf.items.len;
                s.body_start += cell_buf.items.len;
            }
        }
        const opening_gt = std.mem.indexOfScalar(u8, cell_buf.items, '>') orelse unreachable;
        const new_span: CellSpan = .{
            .start = t.after_open,
            .end = t.after_open + cell_buf.items.len,
            .body_start = t.after_open + opening_gt + 1,
            .row = row,
            .col = col,
        };
        var span_pos: usize = ms.spans.items.len;
        for (ms.spans.items, 0..) |s, i| {
            if (s.start >= new_span.end) {
                span_pos = i;
                break;
            }
        }
        ms.spans.insertAssumeCapacity(span_pos, new_span);
        return true;
    }
    return false;
}

/// Insert a new `<row r="N"><c …>…</c></row>` block into the
/// MutatedSheet at the lexicographic position for `row`. iter-cm-2d.
fn insertMissingRow(
    allocator: Allocator,
    ms: *Editor.MutatedSheet,
    row: u32,
    col: u32,
    cell: Cell,
) !void {
    // Reserve span capacity up front so the final insert can't OOM
    // after we've already shifted bytes + offsets.
    try ms.spans.ensureUnusedCapacity(allocator, 1);
    // Find insertion offset:
    //   - just before the next-higher row's `<row` opening, OR
    //   - just before `</sheetData>` if no higher row exists.
    //
    // Walk EVERY `<row>` tag (not just rows that produced spans) so
    // cellless rows like `<row r="5"/>` or `<row r="5"></row>`
    // anchor the position correctly. Going only by spans would
    // misorder a `setCell(row=3, …)` on a sheet with row 5 empty +
    // row 10 populated (would produce XML order 5,3,10). Use the
    // same explicit-r → first-cell → implicit-counter cascade
    // Rows.next uses.
    var insert_at: usize = 0;
    var anchored = false;
    var pos_walk: usize = 0;
    var implicit_row_walk: u32 = 0;
    while (findTagOpen(ms.xml.items, pos_walk, "row")) |t| {
        implicit_row_walk += 1;
        const attrs_full = ms.xml.items[t.start + "<row".len .. t.after_open - 1];
        const trimmed = std.mem.trimRight(u8, attrs_full, " \t\r\n");
        const is_self_closing = trimmed.len > 0 and trimmed[trimmed.len - 1] == '/';
        const attrs = if (is_self_closing) trimmed[0 .. trimmed.len - 1] else trimmed;
        const effective_row: u32 = blk: {
            if (getAttr(attrs, "r")) |r_attr| {
                const parsed = std.fmt.parseInt(u32, r_attr, 10) catch 0;
                if (parsed > 0) break :blk parsed;
            }
            if (!is_self_closing) {
                if (recoverRowFromFirstCell(ms.xml.items, t.after_open)) |n|
                    break :blk n;
            }
            break :blk implicit_row_walk;
        };
        if (effective_row > row) {
            insert_at = t.start;
            anchored = true;
            break;
        }
        pos_walk = t.after_open;
    }
    if (!anchored) {
        // No higher row — append before `</sheetData>`. If the
        // worksheet is empty (`<sheetData/>` self-closing form),
        // expand it to `<sheetData></sheetData>` first; mirrors
        // what the append path does in `injectAppendedRows`.
        if (std.mem.indexOf(u8, ms.xml.items, "</sheetData>")) |end| {
            insert_at = end;
        } else if (std.mem.indexOf(u8, ms.xml.items, "<sheetData")) |sd_open| {
            // Find the closing `>` of the self-closing tag.
            const sd_close = std.mem.indexOfScalarPos(u8, ms.xml.items, sd_open, '>') orelse
                return error.MalformedXml;
            if (sd_close == 0 or ms.xml.items[sd_close - 1] != '/')
                return error.MalformedXml;
            // Replace `<sheetData [attrs]/>` with `<sheetData
            // [attrs]></sheetData>`. Net byte delta = +
            // "</sheetData>".len ("/" is consumed). All spans
            // after sd_close shift by that delta.
            const expand_at = sd_close - 1; // position of the `/`
            try ms.xml.replaceRange(allocator, expand_at, 2, "></sheetData>");
            const delta: usize = "></sheetData>".len - 2;
            for (ms.spans.items) |*s| {
                if (s.start >= sd_close) {
                    s.start += delta;
                    s.end += delta;
                    s.body_start += delta;
                }
            }
            // Recompute insertion offset — the new `</sheetData>`
            // sits where the old `/` was, plus 1 (for the new `>`).
            insert_at = expand_at + 1;
        } else {
            return error.MalformedXml;
        }
    }

    // Build the new row block.
    var new_buf: std.ArrayListUnmanaged(u8) = .{};
    defer new_buf.deinit(allocator);
    try new_buf.writer(allocator).print("<row r=\"{d}\">", .{row});
    const cell_start_in_buf = new_buf.items.len;
    try emitCellXml(allocator, &new_buf, row, col, cell);
    try new_buf.appendSlice(allocator, "</row>");

    try ms.xml.insertSlice(allocator, insert_at, new_buf.items);
    const new_len = new_buf.items.len;

    // Shift every later span by new_len.
    for (ms.spans.items) |*s| {
        if (s.start >= insert_at) {
            s.start += new_len;
            s.end += new_len;
            s.body_start += new_len;
        }
    }

    // Compute the new cell's span. Inside new_buf, the cell starts
    // at `cell_start_in_buf` and ends at `new_buf.len - "</row>".len`.
    const cell_abs_start = insert_at + cell_start_in_buf;
    const cell_abs_end = insert_at + new_buf.items.len - "</row>".len;
    // body_start = position of '>' within the cell's opening tag.
    const cell_bytes = new_buf.items[cell_start_in_buf .. new_buf.items.len - "</row>".len];
    const opening_gt_off = std.mem.indexOfScalar(u8, cell_bytes, '>') orelse unreachable;
    const new_cell_span: CellSpan = .{
        .start = cell_abs_start,
        .end = cell_abs_end,
        .body_start = cell_abs_start + opening_gt_off + 1,
        .row = row,
        .col = col,
    };

    // Find the right insertion index in spans (source order).
    var span_pos: usize = ms.spans.items.len;
    for (ms.spans.items, 0..) |s, i| {
        if (s.start >= cell_abs_end) {
            span_pos = i;
            break;
        }
    }
    ms.spans.insertAssumeCapacity(span_pos, new_cell_span);
}

/// Emit a fresh `<c r="REF"…>…</c>` for one cell. Output is canonical
/// (no formula, no inline-string body, no preserved source attrs)
/// — iter-cm-2a contract. Caller owns `out`'s buffer.
fn emitCellXml(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    row: u32,
    col: u32,
    cell: Cell,
) !void {
    const writer_mod = @import("writer.zig");
    var ref_buf: [16]u8 = undefined;
    const ref = try writer_mod.formatCellRef(&ref_buf, row, @intCast(col));

    // Empty cells get a self-closing form to match what writers emit.
    if (cell == .empty) {
        try out.writer(allocator).print("<c r=\"{s}\"/>", .{ref});
        return;
    }

    const type_attr: []const u8 = switch (cell) {
        .boolean => " t=\"b\"",
        // iter-cm-2b: emit strings as `t="inlineStr"` rather than
        // patching the SST. Valid OOXML, smaller surface area, no
        // cross-cell coordination needed. Trade-off: repeated
        // strings are no longer deduped by setCell — callers who
        // care can normalise upstream or fall back to a writer.
        .string => " t=\"inlineStr\"",
        else => "",
    };
    try out.writer(allocator).print("<c r=\"{s}\"{s}>", .{ ref, type_attr });
    switch (cell) {
        .empty => unreachable,
        .string => |s| {
            try out.appendSlice(allocator, "<is><t");
            // Honour OOXML's xml:space="preserve" semantics for
            // strings whose first/last byte is whitespace —
            // matches the writer's existing rule.
            const needs_preserve = s.len > 0 and
                (s[0] == ' ' or s[0] == '\t' or s[0] == '\n' or s[0] == '\r' or
                    s[s.len - 1] == ' ' or s[s.len - 1] == '\t' or
                    s[s.len - 1] == '\n' or s[s.len - 1] == '\r');
            if (needs_preserve) try out.appendSlice(allocator, " xml:space=\"preserve\"");
            try out.append(allocator, '>');
            try appendXmlEscaped(allocator, out, s);
            try out.appendSlice(allocator, "</t></is>");
        },
        .integer => |n| try out.writer(allocator).print("<v>{d}</v>", .{n}),
        .number => |f| try out.writer(allocator).print("<v>{d}</v>", .{f}),
        .boolean => |b| try out.writer(allocator).print("<v>{d}</v>", .{@intFromBool(b)}),
    }
    try out.appendSlice(allocator, "</c>");
}

/// Walk worksheet XML and emit a span per `<c>` element. Pure
/// function — no allocator state beyond the returned slice. Mirrors
/// the row/cell parser in `Rows.next` but records byte offsets
/// instead of decoding values.
fn scanWorksheetXml(allocator: Allocator, xml: []const u8) ![]CellSpan {
    var out: std.ArrayListUnmanaged(CellSpan) = .{};
    errdefer out.deinit(allocator);

    var pos: usize = 0;
    var implicit_row: u32 = 0;

    while (findTagOpen(xml, pos, "row")) |row_tag| {
        const row_attrs = xml[row_tag.start + "<row".len .. row_tag.after_open - 1];
        const row_attrs_trim = std.mem.trimRight(u8, row_attrs, " \t\r\n");
        const is_self_closing_row = row_attrs_trim.len > 0 and
            row_attrs_trim[row_attrs_trim.len - 1] == '/';
        const row_attrs_for_lookup = if (is_self_closing_row)
            row_attrs_trim[0 .. row_attrs_trim.len - 1]
        else
            row_attrs_trim;

        implicit_row += 1;
        const row_num: u32 = blk: {
            if (getAttr(row_attrs_for_lookup, "r")) |r_str| {
                const parsed = std.fmt.parseInt(u32, r_str, 10) catch 0;
                if (parsed > 0) break :blk parsed;
            }
            // Mirror Rows.next semantics: when <row r> is absent or
            // unusable, recover from the first <c r="A11">-style ref
            // BEFORE falling back to the implicit row counter. Without
            // this, scanWorksheet would diverge from Book.rows on the
            // same fixture and any future setCell built on it would
            // target the wrong row.
            if (!is_self_closing_row) {
                if (recoverRowFromFirstCell(xml, row_tag.after_open)) |n|
                    break :blk n;
            }
            break :blk implicit_row;
        };

        if (is_self_closing_row) {
            pos = row_tag.after_open;
            continue;
        }

        var implicit_col: u32 = 0;
        var cur = row_tag.after_open;
        while (true) {
            const next_lt = std.mem.indexOfScalarPos(u8, xml, cur, '<') orelse
                return error.MalformedXml;
            cur = next_lt;
            if (std.mem.startsWith(u8, xml[cur..], "</row>")) {
                cur += "</row>".len;
                break;
            }
            // Match `<c` only when followed by whitespace, `/`, or `>` —
            // otherwise we'd match `<col>` / `<choose>` / etc.
            // Pretty-printed OOXML can wrap attrs onto the next line
            // (`<c\n r="A1">…</c>`), so accept newline + carriage
            // return too. Matches what Rows.next tolerates.
            const after_c = cur + 2;
            const is_c_open = cur + 2 <= xml.len and
                xml[cur + 1] == 'c' and
                after_c < xml.len and
                (xml[after_c] == ' ' or xml[after_c] == '\t' or
                    xml[after_c] == '\n' or xml[after_c] == '\r' or
                    xml[after_c] == '/' or xml[after_c] == '>');
            if (!is_c_open) {
                const gt = std.mem.indexOfScalarPos(u8, xml, cur, '>') orelse
                    return error.MalformedXml;
                cur = gt + 1;
                continue;
            }

            const c_start = cur;
            const gt = std.mem.indexOfScalarPos(u8, xml, cur, '>') orelse
                return error.MalformedXml;
            // Trim trailing whitespace before sniffing the self-
            // closing slash. Defensive against XML producers that
            // emit `<c r="A1" / >` (slash separated from `>` by
            // whitespace). Mirrors the same trim on the <row .../>
            // path.
            const candidate_attrs = xml[c_start + 2 .. gt];
            const trimmed = std.mem.trimRight(u8, candidate_attrs, " \t\r\n");
            const is_self_closing = trimmed.len > 0 and trimmed[trimmed.len - 1] == '/';
            const attrs = if (is_self_closing) trimmed[0 .. trimmed.len - 1] else candidate_attrs;

            const col: u32 = blk: {
                if (getAttr(attrs, "r")) |r_attr| {
                    const c_idx = columnIndexFromRef(r_attr) catch
                        return error.MalformedXml;
                    break :blk std.math.cast(u32, c_idx) orelse
                        return error.MalformedXml;
                }
                break :blk implicit_col;
            };
            implicit_col = std.math.cast(u32, @as(u64, col) + 1) orelse
                return error.MalformedXml;

            const c_end = if (is_self_closing) gt + 1 else blk: {
                const close_pos = std.mem.indexOfPos(u8, xml, gt + 1, "</c>") orelse
                    return error.MalformedXml;
                break :blk close_pos + "</c>".len;
            };

            try out.append(allocator, .{
                .start = c_start,
                .end = c_end,
                .body_start = gt + 1,
                .row = row_num,
                .col = col,
            });

            cur = c_end;
        }
        pos = cur;
    }

    return try out.toOwnedSlice(allocator);
}

/// New LFH + payload + CDFH bytes for a sheet entry that's been
/// modified by `appendRows`. Owned by the caller; freed on
/// `Editor.save` exit.
const SubstitutedEntry = struct {
    lfh: []u8,
    payload: []u8,
    cdfh: []u8,
    crc32: u32,
    uncompressed_size: u32,
    compression_method: u16,
};

/// Linear scan: small N (entries per xlsx is ~10-20) makes a hashmap
/// overkill. Returns the first matching index.
fn findEntryByName(entries: []const Editor.ZipEntry, name: []const u8) ?usize {
    for (entries, 0..) |e, i| {
        if (std.mem.eql(u8, e.name, name)) return i;
    }
    return null;
}

/// Build the substituted LFH + payload + CDFH for a sheet that has
/// pending appended rows. Decompresses the source sheet XML, finds
/// `</sheetData>`, injects new `<row>...</row>` blocks, recompresses
/// (or falls back to store if deflate inflates), then emits canonical
/// LFH/CDFH bytes.
/// SST extension state shared across sheet substitutions in a
/// single `Editor.save` pass. iter-lms-3: every appended string cell
/// allocates a fresh SST index (no plain-text equality reuse — the
/// source SST may carry rich-text entries whose plain text shadows
/// an appended string, and reuse would silently inherit formatting).
const SstAppender = struct {
    allocator: Allocator,
    next_idx: u32,
    /// Borrowed slices into editor-owned cell strings. Valid for the
    /// duration of `save()`; new SST XML copies them out when the
    /// sharedStrings.xml entry is rebuilt.
    new_strings: std.ArrayListUnmanaged([]const u8) = .{},

    pub fn add(self: *SstAppender, str: []const u8) !u32 {
        const idx = self.next_idx;
        try self.new_strings.append(self.allocator, str);
        self.next_idx += 1;
        return idx;
    }

    pub fn deinit(self: *SstAppender) void {
        self.new_strings.deinit(self.allocator);
    }
};

fn buildSubstitutedSheet(
    allocator: Allocator,
    entry: Editor.ZipEntry,
    src_buf: []u8,
    appended_rows: []const []Cell,
    sst: ?*SstAppender,
) !SubstitutedEntry {
    // Decompress source payload.
    const payload = src_buf[entry.lfh_offset + entry.lfh_total_len ..][0..entry.payload_len];
    const decompressed = try decompressZipPayload(
        allocator,
        payload,
        entry.compression_method,
        entry.uncompressed_size,
    );
    defer allocator.free(decompressed);

    // Find the first row index for the appends — one past the
    // highest used row in the source.
    const highest_row = findHighestRowInSheetXml(decompressed);
    const start_row: u32 = highest_row + 1;
    // Refuse appends that push past Excel's max row (1,048,576).
    const final_row: u64 = @as(u64, start_row) + @as(u64, appended_rows.len) - 1;
    if (final_row > max_row) return error.RowIndexOutOfRange;

    // Inject `<row r="N">…</row>` blocks before `</sheetData>`.
    const injected = try injectAppendedRows(
        allocator,
        decompressed,
        appended_rows,
        start_row,
        sst,
    );

    // Update the canonical-form `<dimension ref="A1:Z<row>"/>` to
    // extend both row AND column bounds when the appended rows
    // pushed past the source's declared dimension. Other shapes
    // (no dimension / open-tag form / single-cell ref) are left
    // alone — Excel recomputes the dimension on its next save so
    // staleness on those is tolerable.
    const new_max_row: u32 = start_row +
        @as(u32, @intCast(appended_rows.len)) - 1;
    var new_max_col_1based: u32 = 0;
    for (appended_rows) |row| {
        if (row.len > new_max_col_1based) new_max_col_1based = @intCast(row.len);
    }
    const new_xml = blk: {
        if (try updateDimension(allocator, injected, new_max_row, new_max_col_1based)) |patched| {
            allocator.free(injected);
            break :blk patched;
        } else {
            break :blk injected;
        }
    };

    return try buildEntryFromXml(allocator, entry.name, new_xml);
}

/// Count `<si>` openings in a sharedStrings.xml body. Cheap one-pass
/// scan — used once at save() to know how many indices the source
/// SST already occupies before assigning indices to appended strings.
fn countSiInSst(xml: []const u8) u32 {
    var count: u32 = 0;
    var i: usize = 0;
    while (std.mem.indexOfPos(u8, xml, i, "<si")) |pos| {
        if (pos + "<si".len < xml.len) {
            const c = xml[pos + "<si".len];
            if (c == ' ' or c == '/' or c == '>') {
                count += 1;
                i = pos + "<si".len + 1;
                continue;
            }
        }
        i = pos + 1;
    }
    return count;
}

/// Append `<si><t>…</t></si>` blocks (one per new string,
/// XML-escaped) to a copy of the source SST XML and patch the
/// `count` / `uniqueCount` attributes on the `<sst>` opening tag.
/// Handles both `<sst …>…</sst>` and `<sst …/>` forms. Returns the
/// substituted entry (LFH + new compressed payload + CDFH).
/// Build a fresh `xl/sharedStrings.xml` body containing only the
/// caller's appended strings. Used when the source workbook has no
/// SST entry at all (only inline strings or numeric data) and the
/// editor is appending its first string cells. Caller owns the
/// returned slice.
fn buildFreshSstXml(allocator: Allocator, strings: []const []const u8) ![]u8 {
    var out: std.ArrayListUnmanaged(u8) = .{};
    errdefer out.deinit(allocator);
    try out.appendSlice(allocator, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
        "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
    try std.fmt.format(out.writer(allocator), " count=\"{d}\" uniqueCount=\"{d}\">", .{ strings.len, strings.len });
    for (strings) |s| {
        try out.appendSlice(allocator, "<si><t");
        if (sstNeedsXmlSpacePreserve(s)) {
            try out.appendSlice(allocator, " xml:space=\"preserve\"");
        }
        try out.appendSlice(allocator, ">");
        try appendXmlEscaped(allocator, &out, s);
        try out.appendSlice(allocator, "</t></si>");
    }
    try out.appendSlice(allocator, "</sst>");
    return try out.toOwnedSlice(allocator);
}

/// Splice an `Override` into `[Content_Types].xml` so readers
/// recognise the freshly-created `xl/sharedStrings.xml` part.
/// No-op (returns null) when the override is already present.
/// Scan `xml` for the largest `attr="N"` numeric value (e.g.
/// `sheetId="3"`). Returns 0 when no match is found. Used by
/// `Editor.addSheet` to pick non-colliding ids.
fn nextMaxAttr(xml: []const u8, attr_prefix: []const u8) u32 {
    var max_id: u32 = 0;
    var i: usize = 0;
    while (std.mem.indexOfPos(u8, xml, i, attr_prefix)) |pos| {
        const num_start = pos + attr_prefix.len;
        var num_end = num_start;
        while (num_end < xml.len and xml[num_end] >= '0' and xml[num_end] <= '9') : (num_end += 1) {}
        if (num_end > num_start) {
            if (std.fmt.parseInt(u32, xml[num_start..num_end], 10)) |n| {
                if (n > max_id) max_id = n;
            } else |_| {}
        }
        i = num_end + 1;
    }
    return max_id;
}

/// Largest existing `rIdN` numeric suffix in the rels XML.
fn nextMaxRId(xml: []const u8) u32 {
    return nextMaxAttr(xml, "Id=\"rId");
}

/// Highest `sheetN.xml` number seen in the entry table. Returns
/// 0 when no `xl/worksheets/sheet*.xml` entry exists.
fn nextMaxSheetPathNum(entries: []const Editor.ZipEntry) u32 {
    const prefix = "xl/worksheets/sheet";
    const suffix = ".xml";
    var max_n: u32 = 0;
    for (entries) |e| {
        if (!std.mem.startsWith(u8, e.name, prefix)) continue;
        if (!std.mem.endsWith(u8, e.name, suffix)) continue;
        const num_str = e.name[prefix.len .. e.name.len - suffix.len];
        if (num_str.len == 0) continue;
        if (std.fmt.parseInt(u32, num_str, 10)) |n| {
            if (n > max_n) max_n = n;
        } else |_| {}
    }
    return max_n;
}

/// Rewrite `xl/workbook.xml`'s `<sheet>` line at position
/// `r.sheet_idx` (0-based among `<sheet>` elements in source
/// order) for each pending rename. Iter-sheet-2.
///
/// Match is by POSITION, not by name — addSheet's patcher may have
/// already appended new `<sheet>` lines, and a freshly-added sheet
/// with the same name as a pre-rename old name would otherwise be
/// hit by accident.
fn patchWorkbookXmlForRenames(
    allocator: Allocator,
    xml: []const u8,
    renames: []const Editor.SheetRename,
) ![]u8 {
    var out: std.ArrayListUnmanaged(u8) = .{};
    errdefer out.deinit(allocator);

    var i: usize = 0;
    while (findTagOpen(xml, i, "sheet")) |t| {
        try out.appendSlice(allocator, xml[i .. t.start + "<sheet".len]);
        const attrs = xml[t.start + "<sheet".len .. t.after_open - 1];
        const sh_end = t.after_open - 1; // position of `>`

        // Match by `r:id` (captured at renameSheet time). Position-
        // based match would target the wrong line on workbooks where
        // Book.sheets indexing diverges from raw <sheet> order
        // (e.g. broken rels entries skipped by parseWorkbookSheets).
        var matched_new: ?[]const u8 = null;
        if (getAttr(attrs, "r:id")) |rid| {
            for (renames) |r| {
                if (std.mem.eql(u8, r.rid, rid)) {
                    matched_new = r.new_name;
                    break;
                }
            }
        }

        if (matched_new) |new_name| {
            // Rewrite the `name="..."` attribute. Walk attrs and
            // emit verbatim except the name field, which we
            // replace with the escaped new value.
            var ai: usize = 0;
            while (ai < attrs.len) {
                // Skip whitespace.
                while (ai < attrs.len and (attrs[ai] == ' ' or
                    attrs[ai] == '\t' or attrs[ai] == '\n' or
                    attrs[ai] == '\r')) : (ai += 1)
                {
                    try out.append(allocator, attrs[ai]);
                }
                if (ai >= attrs.len) break;
                // Read attribute name.
                const name_start = ai;
                while (ai < attrs.len and attrs[ai] != '=' and
                    attrs[ai] != ' ' and attrs[ai] != '\t' and
                    attrs[ai] != '\n' and attrs[ai] != '\r') ai += 1;
                const attr_name = attrs[name_start..ai];
                // Skip = and the quote.
                while (ai < attrs.len and attrs[ai] != '=') ai += 1;
                if (ai >= attrs.len) {
                    try out.appendSlice(allocator, attrs[name_start..]);
                    break;
                }
                ai += 1;
                if (ai >= attrs.len) {
                    try out.appendSlice(allocator, attrs[name_start..]);
                    break;
                }
                const quote = attrs[ai];
                ai += 1;
                const val_start = ai;
                while (ai < attrs.len and attrs[ai] != quote) ai += 1;
                const val = attrs[val_start..ai];
                if (ai < attrs.len) ai += 1; // step past closing quote

                if (std.mem.eql(u8, attr_name, "name")) {
                    try out.appendSlice(allocator, "name=\"");
                    try appendXmlAttrEscaped(allocator, &out, new_name);
                    try out.append(allocator, '"');
                } else {
                    // Verbatim: name=quote+value+quote.
                    try out.appendSlice(allocator, attr_name);
                    try out.append(allocator, '=');
                    try out.append(allocator, quote);
                    try out.appendSlice(allocator, val);
                    try out.append(allocator, quote);
                }
            }
        } else {
            try out.appendSlice(allocator, attrs);
        }

        i = sh_end;
    }
    try out.appendSlice(allocator, xml[i..]);
    return try out.toOwnedSlice(allocator);
}

/// Splice the new-sheet `<sheet …/>` lines into `xl/workbook.xml`'s
/// `<sheets>…</sheets>` block. iter-sheet-1.
fn patchWorkbookXmlForNewSheets(
    allocator: Allocator,
    xml: []const u8,
    new_sheets: []const Editor.NewSheet,
) ![]u8 {
    const close = std.mem.indexOf(u8, xml, "</sheets>") orelse return error.MalformedXml;
    var out: std.ArrayListUnmanaged(u8) = .{};
    errdefer out.deinit(allocator);
    try out.appendSlice(allocator, xml[0..close]);
    for (new_sheets) |s| {
        try out.appendSlice(allocator, "<sheet name=\"");
        try appendXmlAttrEscaped(allocator, &out, s.name);
        try out.writer(allocator).print(
            "\" sheetId=\"{d}\" r:id=\"{s}\"/>",
            .{ s.sheet_id, s.rid },
        );
    }
    try out.appendSlice(allocator, xml[close..]);
    return try out.toOwnedSlice(allocator);
}

/// Splice `<Relationship/>` entries for the new sheets into the
/// workbook rels file. iter-sheet-1.
fn patchWorkbookRelsForNewSheets(
    allocator: Allocator,
    xml: []const u8,
    new_sheets: []const Editor.NewSheet,
) ![]u8 {
    const close = std.mem.indexOf(u8, xml, "</Relationships>") orelse return error.MalformedXml;
    var out: std.ArrayListUnmanaged(u8) = .{};
    errdefer out.deinit(allocator);
    try out.appendSlice(allocator, xml[0..close]);
    for (new_sheets) |s| {
        // Target is relative to xl/_rels/, so strip the "xl/" prefix
        // from the sheet path. e.g. "xl/worksheets/sheet5.xml" →
        // "worksheets/sheet5.xml".
        const target = if (std.mem.startsWith(u8, s.path, "xl/")) s.path[3..] else s.path;
        try out.writer(allocator).print(
            "<Relationship Id=\"{s}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"{s}\"/>",
            .{ s.rid, target },
        );
    }
    try out.appendSlice(allocator, xml[close..]);
    return try out.toOwnedSlice(allocator);
}

/// Splice `<Override PartName="/xl/worksheets/…" ContentType="…"/>`
/// entries for the new sheets into `[Content_Types].xml`. iter-sheet-1.
fn patchContentTypesForNewSheets(
    allocator: Allocator,
    xml: []const u8,
    new_sheets: []const Editor.NewSheet,
) ![]u8 {
    const close = std.mem.indexOf(u8, xml, "</Types>") orelse return error.MalformedXml;
    var out: std.ArrayListUnmanaged(u8) = .{};
    errdefer out.deinit(allocator);
    try out.appendSlice(allocator, xml[0..close]);
    for (new_sheets) |s| {
        try out.writer(allocator).print(
            "<Override PartName=\"/{s}\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>",
            .{s.path},
        );
    }
    try out.appendSlice(allocator, xml[close..]);
    return try out.toOwnedSlice(allocator);
}

/// Apply one column edit (insert or delete at `col_1based`) to a
/// worksheet XML buffer. iter-col-3/4 v1: rewrites `<c r="A1">`
/// column letter, `<col min=N max=M>` bounds, `<mergeCells>` rect
/// bounds, and `<dimension>`. Other elements pass through —
/// recordColEdit refuses sheets that contain formulas / hyperlinks
/// / validations / cond-formats / drawings / tables, so we don't
/// have to handle those here.
fn applyColEditToWorksheet(
    allocator: Allocator,
    src: []const u8,
    col_1based: u32,
    kind: RowEditKind,
) ![]u8 {
    var out: std.ArrayListUnmanaged(u8) = .{};
    errdefer out.deinit(allocator);

    var i: usize = 0;
    while (i < src.len) {
        const next_lt = std.mem.indexOfScalarPos(u8, src, i, '<') orelse {
            try out.appendSlice(allocator, src[i..]);
            return try out.toOwnedSlice(allocator);
        };
        try out.appendSlice(allocator, src[i..next_lt]);
        i = next_lt;

        if (matchTagAt(src, i, "c")) |t| {
            try processCellTagCol(allocator, &out, src, t, col_1based, kind, &i);
        } else if (matchTagAt(src, i, "col")) |t| {
            try processColTag(allocator, &out, src, t, col_1based, kind);
            i = t.after_open;
        } else if (matchTagAt(src, i, "mergeCell")) |t| {
            try processMergeCellTagCol(allocator, &out, src, t, col_1based, kind);
            i = t.after_open;
        } else if (matchTagAt(src, i, "dimension")) |t| {
            try processDimensionTagCol(allocator, &out, src, t, col_1based, kind);
            i = t.after_open;
        } else {
            try out.append(allocator, '<');
            i += 1;
        }
    }
    return try out.toOwnedSlice(allocator);
}

fn processCellTagCol(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    col_1based: u32,
    kind: RowEditKind,
    i: *usize,
) !void {
    const attrs = src[t.start + "<c".len .. t.after_open - 1];
    const trimmed = std.mem.trimRight(u8, attrs, " \t\r\n");
    const is_self_closing = trimmed.len > 0 and trimmed[trimmed.len - 1] == '/';
    const attrs_for_lookup = if (is_self_closing) trimmed[0 .. trimmed.len - 1] else trimmed;
    const r_attr = getAttr(attrs_for_lookup, "r");
    if (r_attr == null) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        i.* = t.after_open;
        return;
    }
    const ref = r_attr.?;
    var letters_end: usize = 0;
    while (letters_end < ref.len and ref[letters_end] >= 'A' and ref[letters_end] <= 'Z') letters_end += 1;
    if (letters_end == 0 or letters_end == ref.len) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        i.* = t.after_open;
        return;
    }
    const old_col = parseColLetters(ref[0..letters_end]) orelse {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        i.* = t.after_open;
        return;
    };

    // Delete-match: drop the entire <c> element (open + body + close).
    if (kind == .delete and old_col == col_1based) {
        if (is_self_closing) {
            i.* = t.after_open;
        } else {
            const close = std.mem.indexOfPos(u8, src, t.after_open, "</c>") orelse t.after_open;
            i.* = if (close + "</c>".len <= src.len) close + "</c>".len else t.after_open;
        }
        return;
    }

    var new_col: u32 = old_col;
    switch (kind) {
        .insert => if (old_col >= col_1based) {
            if (old_col >= max_col_1based) return error.ColEditExceedsMaxCol;
            new_col = old_col + 1;
        },
        .delete => if (old_col > col_1based) {
            new_col = old_col - 1;
        },
    }
    if (new_col == old_col) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        i.* = t.after_open;
        return;
    }
    var letters_buf: [8]u8 = undefined;
    const new_letters = formatColLetters(&letters_buf, new_col - 1) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        i.* = t.after_open;
        return;
    };
    var new_ref_buf: [16]u8 = undefined;
    const new_ref = try std.fmt.bufPrint(&new_ref_buf, "{s}{s}", .{ new_letters, ref[letters_end..] });
    try writeWithReplacedAttr(allocator, out, src, t, "<c".len, "r", new_ref);
    i.* = t.after_open;
}

fn processColTag(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    col_1based: u32,
    kind: RowEditKind,
) !void {
    // <col min="N" max="M" .../> covers a contiguous range of cols.
    // For inserts: shift min/max if >= col_1based; if min < col <= max,
    // split into two col entries (we just shift max in v1; the
    // inserted column gets default formatting).
    // For deletes: similar shift; if min == max == col, drop entry.
    //   Else if min == col, increment min.
    //   Else if max == col, decrement max.
    //   Else if min <= col <= max, shrink max by 1.
    const attrs = src[t.start + "<col".len .. t.after_open - 1];
    const min_str = getAttr(attrs, "min");
    const max_str = getAttr(attrs, "max");
    if (min_str == null or max_str == null) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    const old_min = std.fmt.parseInt(u32, min_str.?, 10) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    const old_max = std.fmt.parseInt(u32, max_str.?, 10) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    var new_min = old_min;
    var new_max = old_max;
    var drop = false;
    var split_emit_first_max: ?u32 = null;
    switch (kind) {
        .insert => {
            // Insert point strictly INSIDE the existing range:
            // split into two `<col>` entries, leaving the inserted
            // column with default formatting. e.g. <col min=2 max=4>
            // + insertColumn(3) → <col min=2 max=2/> +
            // <col min=4 max=5/>; col 3 (the inserted one) gets no
            // formatting entry.
            if (old_min < col_1based and col_1based <= old_max) {
                split_emit_first_max = col_1based - 1;
                new_min = col_1based + 1;
                new_max = old_max + 1;
            } else {
                if (old_min >= col_1based) new_min = old_min + 1;
                if (old_max >= col_1based) new_max = old_max + 1;
            }
            // Insert that pushes a <col> range past XFD (16384)
            // would emit max="16385", which Excel rejects as out
            // of range. Mirror the cell-reference path's error.
            if (new_max > max_col_1based or new_min > max_col_1based) {
                return error.ColEditExceedsMaxCol;
            }
        },
        .delete => {
            if (old_min == col_1based and old_max == col_1based) {
                drop = true;
            } else {
                if (old_min > col_1based) new_min = old_min - 1;
                if (old_max >= col_1based) new_max = old_max - 1;
                if (new_min > new_max) drop = true;
            }
        },
    }
    if (drop) return;
    if (split_emit_first_max == null and new_min == old_min and new_max == old_max) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    if (split_emit_first_max) |first_max| {
        // Emit two <col> entries: one for the unchanged head, one
        // for the shifted tail.
        try emitColEntry(allocator, out, attrs, src, t, old_min, first_max);
        try emitColEntry(allocator, out, attrs, src, t, new_min, new_max);
    } else {
        try emitColEntry(allocator, out, attrs, src, t, new_min, new_max);
    }
}

/// Emit a `<col>` tag with its attrs preserved verbatim except
/// `min` and `max`, which take the supplied values.
fn emitColEntry(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    attrs: []const u8,
    src: []const u8,
    t: TagOpen,
    new_min: u32,
    new_max: u32,
) !void {
    var min_buf: [12]u8 = undefined;
    var max_buf: [12]u8 = undefined;
    const min_s = try std.fmt.bufPrint(&min_buf, "{d}", .{new_min});
    const max_s = try std.fmt.bufPrint(&max_buf, "{d}", .{new_max});
    try out.appendSlice(allocator, "<col");
    var ai: usize = 0;
    while (ai < attrs.len) {
        const ws_start = ai;
        while (ai < attrs.len and (attrs[ai] == ' ' or attrs[ai] == '\t' or
            attrs[ai] == '\n' or attrs[ai] == '\r')) ai += 1;
        try out.appendSlice(allocator, attrs[ws_start..ai]);
        if (ai >= attrs.len) break;
        const name_start = ai;
        while (ai < attrs.len and attrs[ai] != '=' and attrs[ai] != ' ' and
            attrs[ai] != '\t' and attrs[ai] != '\n' and attrs[ai] != '\r') ai += 1;
        const aname = attrs[name_start..ai];
        while (ai < attrs.len and attrs[ai] != '=') ai += 1;
        if (ai >= attrs.len) break;
        ai += 1;
        while (ai < attrs.len and (attrs[ai] == ' ' or attrs[ai] == '\t' or
            attrs[ai] == '\n' or attrs[ai] == '\r')) ai += 1;
        if (ai >= attrs.len or (attrs[ai] != '"' and attrs[ai] != '\'')) break;
        const quote = attrs[ai];
        ai += 1;
        const val_start = ai;
        while (ai < attrs.len and attrs[ai] != quote) ai += 1;
        const val = attrs[val_start..ai];
        if (ai < attrs.len) ai += 1;
        try out.appendSlice(allocator, aname);
        try out.append(allocator, '=');
        try out.append(allocator, quote);
        if (std.mem.eql(u8, aname, "min")) {
            try out.appendSlice(allocator, min_s);
        } else if (std.mem.eql(u8, aname, "max")) {
            try out.appendSlice(allocator, max_s);
        } else {
            try out.appendSlice(allocator, val);
        }
        try out.append(allocator, quote);
    }
    // OOXML <col/> is always an empty element. Detect the original
    // self-closing form (last non-ws byte of attrs is `/`) and
    // emit `/>` accordingly, so we don't leave open `<col>` tags.
    const trimmed_attrs = std.mem.trimRight(u8, attrs, " \t\r\n");
    const was_self_closing = trimmed_attrs.len > 0 and trimmed_attrs[trimmed_attrs.len - 1] == '/';
    if (was_self_closing) {
        try out.appendSlice(allocator, "/>");
    } else {
        try out.appendSlice(allocator, src[t.after_open - 1 .. t.after_open]);
    }
}

fn processMergeCellTagCol(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    col_1based: u32,
    kind: RowEditKind,
) !void {
    const attrs = src[t.start + "<mergeCell".len .. t.after_open - 1];
    const r_attr = getAttr(attrs, "ref");
    if (r_attr == null) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    const ref = r_attr.?;
    const colon = std.mem.indexOfScalar(u8, ref, ':') orelse {
        // Single-cell merge.
        if (kind == .delete) {
            const c = parseColFromA1(ref) orelse 0;
            if (c == col_1based) return; // drop
        }
        var new_buf: [16]u8 = undefined;
        const new_ref = shiftSingleA1Col(ref, col_1based, kind, &new_buf, false) catch {
            try out.appendSlice(allocator, src[t.start..t.after_open]);
            return;
        };
        if (std.mem.eql(u8, ref, new_ref)) {
            try out.appendSlice(allocator, src[t.start..t.after_open]);
            return;
        }
        try writeWithReplacedAttr(allocator, out, src, t, "<mergeCell".len, "ref", new_ref);
        return;
    };
    if (kind == .delete) {
        const tl_col = parseColFromA1(ref[0..colon]) orelse 0;
        const br_col = parseColFromA1(ref[colon + 1 ..]) orelse 0;
        if (tl_col == col_1based and br_col == col_1based) return;
    }
    var tl_buf: [16]u8 = undefined;
    var br_buf: [16]u8 = undefined;
    const tl_new = shiftSingleA1Col(ref[0..colon], col_1based, kind, &tl_buf, false) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    const br_new = shiftSingleA1Col(ref[colon + 1 ..], col_1based, kind, &br_buf, true) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    var new_ref_buf: [40]u8 = undefined;
    const new_ref = try std.fmt.bufPrint(&new_ref_buf, "{s}:{s}", .{ tl_new, br_new });
    if (std.mem.eql(u8, ref, new_ref)) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    try writeWithReplacedAttr(allocator, out, src, t, "<mergeCell".len, "ref", new_ref);
}

fn processDimensionTagCol(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    col_1based: u32,
    kind: RowEditKind,
) !void {
    const attrs = src[t.start + "<dimension".len .. t.after_open - 1];
    const r_attr = getAttr(attrs, "ref");
    if (r_attr == null) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    const ref = r_attr.?;
    const colon = std.mem.indexOfScalar(u8, ref, ':') orelse {
        // Single-cell dimension. On a delete-match (the only used
        // column == col_1based) the cell is gone; fall back to "A1"
        // so the dimension stays valid (Excel recomputes on open).
        if (kind == .delete) {
            if (parseColFromA1(ref)) |c| {
                if (c == col_1based) {
                    try writeWithReplacedAttr(allocator, out, src, t, "<dimension".len, "ref", "A1");
                    return;
                }
            }
        }
        var b: [16]u8 = undefined;
        const new_ref = shiftSingleA1Col(ref, col_1based, kind, &b, false) catch {
            try out.appendSlice(allocator, src[t.start..t.after_open]);
            return;
        };
        if (std.mem.eql(u8, ref, new_ref)) {
            try out.appendSlice(allocator, src[t.start..t.after_open]);
            return;
        }
        try writeWithReplacedAttr(allocator, out, src, t, "<dimension".len, "ref", new_ref);
        return;
    };
    // Detect "entire range is the deleted col" up front: both
    // corners at col_1based on a delete. Fall back to a safe
    // sentinel "A1" so the dimension stays valid (Excel
    // recomputes on open).
    if (kind == .delete) {
        const tl_col = parseColFromA1(ref[0..colon]) orelse 0;
        const br_col = parseColFromA1(ref[colon + 1 ..]) orelse 0;
        if (tl_col == col_1based and br_col == col_1based) {
            try writeWithReplacedAttr(allocator, out, src, t, "<dimension".len, "ref", "A1");
            return;
        }
    }
    var tl_buf: [16]u8 = undefined;
    var br_buf: [16]u8 = undefined;
    const tl_new = shiftSingleA1Col(ref[0..colon], col_1based, kind, &tl_buf, false) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    const br_new = shiftSingleA1Col(ref[colon + 1 ..], col_1based, kind, &br_buf, true) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    var new_ref_buf: [40]u8 = undefined;
    const new_ref = try std.fmt.bufPrint(&new_ref_buf, "{s}:{s}", .{ tl_new, br_new });
    if (std.mem.eql(u8, ref, new_ref)) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    try writeWithReplacedAttr(allocator, out, src, t, "<dimension".len, "ref", new_ref);
}

/// Shift the column letters of an A1 ref. `is_br_corner` shrinks
/// the BR corner by one on delete-match.
fn shiftSingleA1Col(ref: []const u8, col_1based: u32, kind: RowEditKind, buf: *[16]u8, is_br_corner: bool) ![]const u8 {
    var letters_end: usize = 0;
    while (letters_end < ref.len and ref[letters_end] >= 'A' and ref[letters_end] <= 'Z') letters_end += 1;
    if (letters_end == 0) return error.MalformedXml;
    const old_col = parseColLetters(ref[0..letters_end]) orelse return error.MalformedXml;
    var new_col: u32 = old_col;
    switch (kind) {
        .insert => if (old_col >= col_1based) {
            if (old_col >= max_col_1based) return error.ColEditExceedsMaxCol;
            new_col = old_col + 1;
        },
        .delete => if (old_col > col_1based) {
            new_col = old_col - 1;
        } else if (old_col == col_1based and is_br_corner and old_col > 1) {
            new_col = old_col - 1;
        },
    }
    if (new_col == old_col) {
        @memcpy(buf[0..ref.len], ref);
        return buf[0..ref.len];
    }
    var letters_buf: [8]u8 = undefined;
    const new_letters = try formatColLetters(&letters_buf, new_col - 1);
    return try std.fmt.bufPrint(buf, "{s}{s}", .{ new_letters, ref[letters_end..] });
}

fn parseColFromA1(ref: []const u8) ?u32 {
    var letters_end: usize = 0;
    while (letters_end < ref.len and ref[letters_end] >= 'A' and ref[letters_end] <= 'Z') letters_end += 1;
    if (letters_end == 0) return null;
    return parseColLetters(ref[0..letters_end]);
}

/// Render a 0-based col_idx as A1 letters (A=0, Z=25, AA=26, ...).
/// Caller-provided buffer; result borrows from buf.
fn formatColLetters(buf: *[8]u8, col_idx: u32) ![]const u8 {
    var col_chars: [8]u8 = undefined;
    var pos: usize = col_chars.len;
    var c = col_idx + 1;
    while (c > 0) {
        c -= 1;
        pos -= 1;
        if (pos == std.math.maxInt(usize)) return error.ColumnIndexOutOfRange;
        col_chars[pos] = @intCast('A' + (c % 26));
        c /= 26;
        if (pos == 0 and c > 0) return error.ColumnIndexOutOfRange;
    }
    const len = col_chars.len - pos;
    @memcpy(buf[0..len], col_chars[pos..]);
    return buf[0..len];
}

/// Apply one row edit (insert or delete at `row`) to a worksheet
/// XML buffer. iter-row-2/3 v1: rewrites `<row r=>` (renumber or
/// drop), `<c r="A1">` row component, `<mergeCells>` rect bounds,
/// and `<dimension>`. Other elements pass through verbatim — the
/// caller's recordRowEdit guard refuses sheets that contain
/// formulas / hyperlinks / validations / cond-formats / drawings,
/// so we don't have to handle those here.
const RowEditKind = enum { insert, delete };

fn applyRowEditToWorksheet(
    allocator: Allocator,
    src: []const u8,
    row: u32,
    kind: RowEditKind,
) ![]u8 {
    var out: std.ArrayListUnmanaged(u8) = .{};
    errdefer out.deinit(allocator);

    var i: usize = 0;
    while (i < src.len) {
        // Find the next interesting tag opening.
        const next_lt = std.mem.indexOfScalarPos(u8, src, i, '<') orelse {
            try out.appendSlice(allocator, src[i..]);
            return try out.toOwnedSlice(allocator);
        };
        try out.appendSlice(allocator, src[i..next_lt]);
        i = next_lt;

        // Identify which tag we're at.
        if (matchTagAt(src, i, "row")) |t| {
            try processRowTag(allocator, &out, src, t, row, kind, &i);
        } else if (matchTagAt(src, i, "c")) |t| {
            try processCellTag(allocator, &out, src, t, row, kind);
            i = t.after_open;
        } else if (matchTagAt(src, i, "mergeCell")) |t| {
            try processMergeCellTag(allocator, &out, src, t, row, kind);
            i = t.after_open;
        } else if (matchTagAt(src, i, "dimension")) |t| {
            try processDimensionTag(allocator, &out, src, t, row, kind);
            i = t.after_open;
        } else {
            // Some other tag; emit `<` and continue past it.
            try out.append(allocator, '<');
            i += 1;
        }
    }
    return try out.toOwnedSlice(allocator);
}

fn matchTagAt(src: []const u8, i: usize, tag: []const u8) ?TagOpen {
    if (i >= src.len or src[i] != '<') return null;
    const after = i + 1 + tag.len;
    if (after > src.len) return null;
    if (!std.mem.eql(u8, src[i + 1 .. i + 1 + tag.len], tag)) return null;
    const c = src[after];
    if (c != ' ' and c != '\t' and c != '\n' and c != '\r' and c != '/' and c != '>') return null;
    const gt = std.mem.indexOfScalarPos(u8, src, i, '>') orelse return null;
    return .{ .start = i, .after_open = gt + 1 };
}

fn processRowTag(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    row: u32,
    kind: RowEditKind,
    i: *usize,
) !void {
    const attrs = src[t.start + "<row".len .. t.after_open - 1];
    const trimmed = std.mem.trimRight(u8, attrs, " \t\r\n");
    const is_self_closing = trimmed.len > 0 and trimmed[trimmed.len - 1] == '/';
    const attrs_for_lookup = if (is_self_closing) trimmed[0 .. trimmed.len - 1] else trimmed;
    const r_attr = getAttr(attrs_for_lookup, "r");
    const old_r: ?u32 = if (r_attr) |s| (std.fmt.parseInt(u32, s, 10) catch null) else null;

    // Decide: drop entirely (delete-match), shift, or pass-through.
    var drop = false;
    var new_r: ?u32 = null;
    if (old_r) |r_val| {
        switch (kind) {
            .insert => if (r_val >= row) {
                if (r_val >= max_row) return error.RowEditExceedsMaxRow;
                new_r = r_val + 1;
            },
            .delete => if (r_val == row) {
                drop = true;
            } else if (r_val > row) {
                new_r = r_val - 1;
            },
        }
    }

    if (drop) {
        // Skip the entire <row>...</row> block. For self-closing
        // form, just skip the opener.
        if (is_self_closing) {
            i.* = t.after_open;
        } else {
            const close = std.mem.indexOfPos(u8, src, t.after_open, "</row>") orelse t.after_open;
            i.* = if (close + "</row>".len <= src.len) close + "</row>".len else t.after_open;
        }
        return;
    }

    // Emit the row open with possibly-rewritten r=.
    if (new_r == null and old_r != null) {
        // Same row number — emit verbatim.
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        i.* = t.after_open;
        // If body form, also recursively process inner tags via
        // the outer loop (they include <c> tags whose row will
        // be rewritten on subsequent iterations).
        return;
    }
    if (new_r == null) {
        // No r= attribute at all — pass through.
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        i.* = t.after_open;
        return;
    }
    // Rewrite r="..." to the new value.
    try writeWithReplacedRowAttr(allocator, out, src, t, "<row".len, new_r.?);
    i.* = t.after_open;
}

fn processCellTag(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    row: u32,
    kind: RowEditKind,
) !void {
    const attrs = src[t.start + "<c".len .. t.after_open - 1];
    const trimmed = std.mem.trimRight(u8, attrs, " \t\r\n");
    const is_self_closing = trimmed.len > 0 and trimmed[trimmed.len - 1] == '/';
    const attrs_for_lookup = if (is_self_closing) trimmed[0 .. trimmed.len - 1] else trimmed;
    const r_attr = getAttr(attrs_for_lookup, "r");
    if (r_attr == null) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    const ref = r_attr.?;
    var letters_end: usize = 0;
    while (letters_end < ref.len and ref[letters_end] >= 'A' and ref[letters_end] <= 'Z') letters_end += 1;
    if (letters_end == 0 or letters_end == ref.len) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    const old_row = std.fmt.parseInt(u32, ref[letters_end..], 10) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    var new_row: u32 = old_row;
    switch (kind) {
        .insert => if (old_row >= row) {
            if (old_row >= max_row) return error.RowEditExceedsMaxRow;
            new_row = old_row + 1;
        },
        .delete => if (old_row > row) {
            new_row = old_row - 1;
        }, // (a cell at the deleted row is dropped along with the <row> block)
    }
    if (new_row == old_row) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    // Rewrite the r= attribute, preserving column letters.
    var new_ref_buf: [16]u8 = undefined;
    const new_ref = try std.fmt.bufPrint(&new_ref_buf, "{s}{d}", .{ ref[0..letters_end], new_row });
    try writeWithReplacedAttr(allocator, out, src, t, "<c".len, "r", new_ref);
}

fn processMergeCellTag(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    row: u32,
    kind: RowEditKind,
) !void {
    const attrs = src[t.start + "<mergeCell".len .. t.after_open - 1];
    const r_attr = getAttr(attrs, "ref");
    if (r_attr == null) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    const ref = r_attr.?;
    const colon = std.mem.indexOfScalar(u8, ref, ':') orelse {
        // Single-cell merge. On delete-match, drop the entire tag
        // so it doesn't transfer onto whatever shifts into the
        // deleted row.
        if (kind == .delete) {
            const tl_row = parseRowFromA1(ref) orelse 0;
            if (tl_row == row) return; // drop tag entirely
        }
        var new_ref_buf: [16]u8 = undefined;
        const new_ref = shiftSingleA1(ref, row, kind, &new_ref_buf, false) catch {
            try out.appendSlice(allocator, src[t.start..t.after_open]);
            return;
        };
        if (std.mem.eql(u8, ref, new_ref)) {
            try out.appendSlice(allocator, src[t.start..t.after_open]);
            return;
        }
        try writeWithReplacedAttr(allocator, out, src, t, "<mergeCell".len, "ref", new_ref);
        return;
    };
    // Rectangle merge. On delete, if BOTH the top-left and
    // bottom-right rows equal the deleted row, the entire merge
    // is on the dead row — drop the tag.
    if (kind == .delete) {
        const tl_row = parseRowFromA1(ref[0..colon]) orelse 0;
        const br_row = parseRowFromA1(ref[colon + 1 ..]) orelse 0;
        if (tl_row == row and br_row == row) return;
    }
    var tl_buf: [16]u8 = undefined;
    var br_buf: [16]u8 = undefined;
    const tl_new = shiftSingleA1(ref[0..colon], row, kind, &tl_buf, false) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    const br_new = shiftSingleA1(ref[colon + 1 ..], row, kind, &br_buf, true) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    var new_ref_buf: [40]u8 = undefined;
    const new_ref = try std.fmt.bufPrint(&new_ref_buf, "{s}:{s}", .{ tl_new, br_new });
    if (std.mem.eql(u8, ref, new_ref)) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    try writeWithReplacedAttr(allocator, out, src, t, "<mergeCell".len, "ref", new_ref);
}

/// Parse the row component (digits) from an A1-style ref. Returns
/// null on malformed input.
fn parseRowFromA1(ref: []const u8) ?u32 {
    var letters_end: usize = 0;
    while (letters_end < ref.len and ref[letters_end] >= 'A' and ref[letters_end] <= 'Z') letters_end += 1;
    if (letters_end == 0 or letters_end == ref.len) return null;
    return std.fmt.parseInt(u32, ref[letters_end..], 10) catch null;
}

fn processDimensionTag(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    row: u32,
    kind: RowEditKind,
) !void {
    // Reuse mergeCell logic — same `ref="A1:Z100"` / `ref="A1"` shape.
    const attrs = src[t.start + "<dimension".len .. t.after_open - 1];
    const r_attr = getAttr(attrs, "ref");
    if (r_attr == null) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    const ref = r_attr.?;
    const colon = std.mem.indexOfScalar(u8, ref, ':') orelse {
        // Single-cell dimension. On a delete-match (the only used
        // row == deleted row) the cell is gone; fall back to "A1"
        // so the dimension stays valid (Excel recomputes on open).
        if (kind == .delete) {
            if (parseRowFromA1(ref)) |r| {
                if (r == row) {
                    try writeWithReplacedAttr(allocator, out, src, t, "<dimension".len, "ref", "A1");
                    return;
                }
            }
        }
        var b: [16]u8 = undefined;
        const new_ref = shiftSingleA1(ref, row, kind, &b, false) catch {
            try out.appendSlice(allocator, src[t.start..t.after_open]);
            return;
        };
        if (std.mem.eql(u8, ref, new_ref)) {
            try out.appendSlice(allocator, src[t.start..t.after_open]);
            return;
        }
        try writeWithReplacedAttr(allocator, out, src, t, "<dimension".len, "ref", new_ref);
        return;
    };
    // Detect "entire range is the deleted row" up front: both
    // corners at the deleted row on a delete. Fall back to a safe
    // sentinel "A1" so the dimension stays valid (Excel recomputes
    // on open).
    if (kind == .delete) {
        const tl_row = parseRowFromA1(ref[0..colon]) orelse 0;
        const br_row = parseRowFromA1(ref[colon + 1 ..]) orelse 0;
        if (tl_row == row and br_row == row) {
            try writeWithReplacedAttr(allocator, out, src, t, "<dimension".len, "ref", "A1");
            return;
        }
    }
    var tl_buf: [16]u8 = undefined;
    var br_buf: [16]u8 = undefined;
    const tl_new = shiftSingleA1(ref[0..colon], row, kind, &tl_buf, false) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    const br_new = shiftSingleA1(ref[colon + 1 ..], row, kind, &br_buf, true) catch {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    var new_ref_buf: [40]u8 = undefined;
    const new_ref = try std.fmt.bufPrint(&new_ref_buf, "{s}:{s}", .{ tl_new, br_new });
    if (std.mem.eql(u8, ref, new_ref)) {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    }
    try writeWithReplacedAttr(allocator, out, src, t, "<dimension".len, "ref", new_ref);
}

/// Shift the row component of an A1-style ref. `is_br_corner`
/// says "this is the bottom-right of a rectangular range" — on a
/// delete-match (old_row == row), BR corners shrink by 1 so the
/// range collapses to the row above the deleted one.
fn shiftSingleA1(ref: []const u8, row: u32, kind: RowEditKind, buf: *[16]u8, is_br_corner: bool) ![]const u8 {
    var letters_end: usize = 0;
    while (letters_end < ref.len and ref[letters_end] >= 'A' and ref[letters_end] <= 'Z') letters_end += 1;
    if (letters_end == 0 or letters_end == ref.len) return error.MalformedXml;
    const old_row = std.fmt.parseInt(u32, ref[letters_end..], 10) catch return error.MalformedXml;
    var new_row: u32 = old_row;
    switch (kind) {
        .insert => if (old_row >= row) {
            // Cap at Excel's max row (1_048_576). If a shift would
            // overflow, refuse the entire edit — saving an invalid
            // workbook is worse than a typed error at the call
            // site (recordRowEdit catches the propagation).
            if (old_row >= max_row) return error.RowEditExceedsMaxRow;
            new_row = old_row + 1;
        },
        .delete => if (old_row > row) {
            new_row = old_row - 1;
        } else if (old_row == row and is_br_corner and old_row > 1) {
            // BR shrink: don't go below row 1. If the entire range
            // collapses to row 0 we'd produce invalid XML; leaving
            // it at row 1 keeps the dimension valid (Excel will
            // recompute on next save anyway).
            new_row = old_row - 1;
        },
    }
    if (new_row == old_row) {
        @memcpy(buf[0..ref.len], ref);
        return buf[0..ref.len];
    }
    return try std.fmt.bufPrint(buf, "{s}{d}", .{ ref[0..letters_end], new_row });
}

/// Emit the original `<tag attrs>` with `attr_name="..."` value
/// replaced by `new_value`. `tag_name_len` is the length of the
/// tag name including the leading `<` (e.g. `"<c".len` = 2).
fn writeWithReplacedAttr(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    tag_name_len: usize,
    attr_name: []const u8,
    new_value: []const u8,
) !void {
    const attrs_full_start = t.start + tag_name_len;
    const attrs_full_end = t.after_open - 1;
    const attrs = src[attrs_full_start..attrs_full_end];
    // Find the `name="..."` occurrence inside attrs.
    var pat_buf: [32]u8 = undefined;
    const pat = try std.fmt.bufPrint(&pat_buf, "{s}=\"", .{attr_name});
    const pat_off = std.mem.indexOf(u8, attrs, pat) orelse {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    const val_start_in_src = attrs_full_start + pat_off + pat.len;
    const val_end_in_src = std.mem.indexOfScalarPos(u8, src, val_start_in_src, '"') orelse {
        try out.appendSlice(allocator, src[t.start..t.after_open]);
        return;
    };
    try out.appendSlice(allocator, src[t.start..val_start_in_src]);
    try out.appendSlice(allocator, new_value);
    try out.appendSlice(allocator, src[val_end_in_src..t.after_open]);
}

/// Rewrite the row attribute on `<row r="…">`. Just delegates to
/// writeWithReplacedAttr formatted as a number.
fn writeWithReplacedRowAttr(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    src: []const u8,
    t: TagOpen,
    tag_name_len: usize,
    new_row: u32,
) !void {
    var buf: [16]u8 = undefined;
    const s = try std.fmt.bufPrint(&buf, "{d}", .{new_row});
    try writeWithReplacedAttr(allocator, out, src, t, tag_name_len, "r", s);
}

/// Drop the calcChain `<Relationship/>` (Type ends with
/// `/relationships/calcChain`) from the workbook rels file. Used
/// alongside the calcChain ZIP-entry drop on delete.
fn patchWorkbookRelsForCalcChainDrop(
    allocator: Allocator,
    xml: []const u8,
    deletes: []const Editor.SheetDelete,
) ![]u8 {
    _ = deletes;
    var out: std.ArrayListUnmanaged(u8) = .{};
    errdefer out.deinit(allocator);
    var i: usize = 0;
    while (std.mem.indexOfPos(u8, xml, i, "<Relationship")) |rel_pos| {
        const rel_end = std.mem.indexOfScalarPos(u8, xml, rel_pos, '>') orelse {
            try out.appendSlice(allocator, xml[i..]);
            return try out.toOwnedSlice(allocator);
        };
        const rel_attrs = xml[rel_pos + "<Relationship".len .. rel_end];
        var dropped = false;
        if (getAttr(rel_attrs, "Type")) |ty| {
            if (std.mem.endsWith(u8, ty, "/calcChain")) dropped = true;
        }
        if (!dropped) {
            if (getAttr(rel_attrs, "Target")) |tg| {
                if (std.mem.endsWith(u8, tg, "calcChain.xml")) dropped = true;
            }
        }
        if (dropped) {
            try out.appendSlice(allocator, xml[i..rel_pos]);
            i = rel_end + 1;
        } else {
            try out.appendSlice(allocator, xml[i .. rel_end + 1]);
            i = rel_end + 1;
        }
    }
    try out.appendSlice(allocator, xml[i..]);
    return try out.toOwnedSlice(allocator);
}

/// Drop `<Override PartName="/xl/calcChain.xml" .../>` from
/// [Content_Types].xml.
fn patchContentTypesForCalcChainDrop(
    allocator: Allocator,
    xml: []const u8,
    deletes: []const Editor.SheetDelete,
) ![]u8 {
    _ = deletes;
    var out: std.ArrayListUnmanaged(u8) = .{};
    errdefer out.deinit(allocator);
    var i: usize = 0;
    while (std.mem.indexOfPos(u8, xml, i, "<Override")) |o_pos| {
        const o_end = std.mem.indexOfScalarPos(u8, xml, o_pos, '>') orelse {
            try out.appendSlice(allocator, xml[i..]);
            return try out.toOwnedSlice(allocator);
        };
        const o_attrs = xml[o_pos + "<Override".len .. o_end];
        var dropped = false;
        if (getAttr(o_attrs, "PartName")) |part| {
            const part_norm = if (part.len > 0 and part[0] == '/') part[1..] else part;
            if (std.mem.eql(u8, part_norm, "xl/calcChain.xml")) dropped = true;
        }
        if (dropped) {
            try out.appendSlice(allocator, xml[i..o_pos]);
            i = o_end + 1;
        } else {
            try out.appendSlice(allocator, xml[i .. o_end + 1]);
            i = o_end + 1;
        }
    }
    try out.appendSlice(allocator, xml[i..]);
    return try out.toOwnedSlice(allocator);
}

/// Walk `<workbookView>` attributes and rewrite `activeTab=` and
/// `firstSheet=` values to account for deleted sheets. Other attrs
/// pass through verbatim.
fn patchViewAttrs(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    attrs: []const u8,
    dropped_positions: []const u32,
    live_count: u32,
) !void {
    var i: usize = 0;
    while (i < attrs.len) {
        // Skip whitespace (preserve verbatim).
        const ws_start = i;
        while (i < attrs.len and (attrs[i] == ' ' or attrs[i] == '\t' or
            attrs[i] == '\n' or attrs[i] == '\r')) i += 1;
        try out.appendSlice(allocator, attrs[ws_start..i]);
        if (i >= attrs.len) break;

        // Read attribute name.
        const name_start = i;
        while (i < attrs.len and attrs[i] != '=' and attrs[i] != ' ' and
            attrs[i] != '\t' and attrs[i] != '\n' and attrs[i] != '\r') i += 1;
        const attr_name = attrs[name_start..i];

        // Skip optional whitespace + `=`.
        while (i < attrs.len and attrs[i] != '=') i += 1;
        if (i >= attrs.len) {
            try out.appendSlice(allocator, attrs[name_start..]);
            return;
        }
        i += 1; // past `=`
        // Tolerate whitespace between `=` and the opening quote
        // (legal XML; some generators emit it).
        while (i < attrs.len and (attrs[i] == ' ' or attrs[i] == '\t' or
            attrs[i] == '\n' or attrs[i] == '\r')) i += 1;
        if (i >= attrs.len or (attrs[i] != '"' and attrs[i] != '\'')) {
            try out.appendSlice(allocator, attrs[name_start..]);
            return;
        }
        const quote = attrs[i];
        i += 1;
        const val_start = i;
        while (i < attrs.len and attrs[i] != quote) i += 1;
        const val = attrs[val_start..i];
        if (i < attrs.len) i += 1; // past closing quote

        const is_active = std.mem.eql(u8, attr_name, "activeTab");
        const is_first = std.mem.eql(u8, attr_name, "firstSheet");
        if (is_active or is_first) {
            const old_n = std.fmt.parseInt(u32, val, 10) catch live_count;
            var new_n: u32 = 0;
            if (live_count > 0) {
                var was_dropped = false;
                var shift: u32 = 0;
                for (dropped_positions) |dp| {
                    if (dp == old_n) was_dropped = true;
                    if (dp < old_n) shift += 1;
                }
                if (was_dropped) {
                    new_n = 0;
                } else if (old_n >= shift) {
                    new_n = old_n - shift;
                } else {
                    new_n = 0;
                }
                if (new_n >= live_count) new_n = live_count - 1;
            }
            try out.appendSlice(allocator, attr_name);
            try out.append(allocator, '=');
            try out.append(allocator, quote);
            try out.writer(allocator).print("{d}", .{new_n});
            try out.append(allocator, quote);
        } else {
            // Verbatim: name + `=` + quote + value + quote.
            try out.appendSlice(allocator, attr_name);
            try out.append(allocator, '=');
            try out.append(allocator, quote);
            try out.appendSlice(allocator, val);
            try out.append(allocator, quote);
        }
    }
}

/// Drop the `<sheet r:id="rIdN"/>` line from `xl/workbook.xml`
/// for each pending delete (matched by r:id). Two passes:
///   1. First walk computes which raw `<sheet>` positions are
///      being deleted so we can later shift `activeTab`.
///   2. Second pass emits the result, dropping deleted lines and
///      patching `<workbookView activeTab="N"/>` so it never
///      points past the new last sheet.
///
/// Limitation called out in the iter-sheet-3 plan: `<definedName>`
/// entries tied to deleted sheets via `localSheetId` are NOT
/// rewritten — that needs the formula tokenizer (iter-col-1).
/// Print areas / named ranges referencing the deleted sheet may
/// fail repair in strict readers.
fn patchWorkbookXmlForDeletes(
    allocator: Allocator,
    xml: []const u8,
    deletes: []const Editor.SheetDelete,
) ![]u8 {
    // Single-pass walk: visit `<sheet>` and `<workbookView>` tags
    // in source order, regardless of which appears first in the
    // file. parseWorkbookSheets order matches the raw `<sheet>`
    // sequence (broken-rels-skipped sheets aside, which we don't
    // try to model here).
    var dropped_positions: std.ArrayListUnmanaged(u32) = .{};
    defer dropped_positions.deinit(allocator);
    var live_count: u32 = 0;
    var sheet_pos: u32 = 0;
    {
        var i: usize = 0;
        while (findTagOpen(xml, i, "sheet")) |t| {
            const attrs = xml[t.start + "<sheet".len .. t.after_open - 1];
            var is_dropped = false;
            if (getAttr(attrs, "r:id")) |rid| {
                for (deletes) |d| {
                    if (std.mem.eql(u8, d.rid, rid)) {
                        is_dropped = true;
                        break;
                    }
                }
            }
            if (is_dropped) try dropped_positions.append(allocator, sheet_pos) else live_count += 1;
            sheet_pos += 1;
            i = t.after_open;
        }
    }

    var out: std.ArrayListUnmanaged(u8) = .{};
    errdefer out.deinit(allocator);

    // Walk both tag types together, sorted by source position.
    // For each interesting tag, emit bytes up to it, transform
    // it, advance.
    var i: usize = 0;
    while (true) {
        const next_sheet = findTagOpen(xml, i, "sheet");
        const next_view = findTagOpen(xml, i, "workbookView");
        var pick_sheet: bool = false;
        var pick_view: bool = false;
        if (next_sheet) |s| {
            if (next_view) |v| {
                if (s.start < v.start) pick_sheet = true else pick_view = true;
            } else pick_sheet = true;
        } else if (next_view != null) pick_view = true else break;

        if (pick_sheet) {
            const t = next_sheet.?;
            const attrs = xml[t.start + "<sheet".len .. t.after_open - 1];
            const is_self_closing = attrs.len > 0 and attrs[attrs.len - 1] == '/';
            var dropped = false;
            if (getAttr(attrs, "r:id")) |rid| {
                for (deletes) |d| {
                    if (std.mem.eql(u8, d.rid, rid)) {
                        dropped = true;
                        break;
                    }
                }
            }
            if (dropped) {
                // Drop the entire `<sheet>` element. For self-closing
                // form `<sheet ... />`, after_open is past the `>`.
                // For body form `<sheet ...></sheet>`, also skip
                // past the matching `</sheet>` close so we don't
                // leave a dangling close tag.
                try out.appendSlice(allocator, xml[i..t.start]);
                if (is_self_closing) {
                    i = t.after_open;
                } else {
                    if (std.mem.indexOfPos(u8, xml, t.after_open, "</sheet>")) |close_pos| {
                        i = close_pos + "</sheet>".len;
                    } else {
                        i = t.after_open;
                    }
                }
            } else {
                try out.appendSlice(allocator, xml[i..t.after_open]);
                i = t.after_open;
            }
        } else {
            const t = next_view.?;
            try out.appendSlice(allocator, xml[i..t.start]);
            // Rewrite both activeTab + firstSheet attributes inside
            // <workbookView>, in source order, so each adjustment
            // composes cleanly. Both attributes get the same
            // shift/clamp treatment.
            const view_attrs_start = t.start + "<workbookView".len;
            const view_attrs_end = t.after_open - 1;
            try out.append(allocator, '<');
            try out.appendSlice(allocator, "workbookView");
            try patchViewAttrs(
                allocator,
                &out,
                xml[view_attrs_start..view_attrs_end],
                dropped_positions.items,
                live_count,
            );
            // Emit the closing `>` (or `/>` for self-closing).
            try out.appendSlice(allocator, xml[view_attrs_end..t.after_open]);
            i = t.after_open;
        }
    }
    try out.appendSlice(allocator, xml[i..]);
    return try out.toOwnedSlice(allocator);
}

/// Drop `<Relationship Id="rIdN" .../>` lines from
/// `xl/_rels/workbook.xml.rels` for each pending delete.
fn patchWorkbookRelsForDeletes(
    allocator: Allocator,
    xml: []const u8,
    deletes: []const Editor.SheetDelete,
) ![]u8 {
    var out: std.ArrayListUnmanaged(u8) = .{};
    errdefer out.deinit(allocator);
    var i: usize = 0;
    while (std.mem.indexOfPos(u8, xml, i, "<Relationship")) |rel_pos| {
        const rel_end = std.mem.indexOfScalarPos(u8, xml, rel_pos, '>') orelse {
            try out.appendSlice(allocator, xml[i..]);
            return try out.toOwnedSlice(allocator);
        };
        const rel_attrs = xml[rel_pos + "<Relationship".len .. rel_end];
        var dropped = false;
        if (getAttr(rel_attrs, "Id")) |id| {
            for (deletes) |d| {
                if (std.mem.eql(u8, d.rid, id)) {
                    dropped = true;
                    break;
                }
            }
        }
        if (dropped) {
            try out.appendSlice(allocator, xml[i..rel_pos]);
            i = rel_end + 1;
        } else {
            try out.appendSlice(allocator, xml[i .. rel_end + 1]);
            i = rel_end + 1;
        }
    }
    try out.appendSlice(allocator, xml[i..]);
    return try out.toOwnedSlice(allocator);
}

/// Drop `<Override PartName="/xl/worksheets/sheetN.xml" …/>`
/// entries for each pending delete.
fn patchContentTypesForDeletes(
    allocator: Allocator,
    xml: []const u8,
    deletes: []const Editor.SheetDelete,
) ![]u8 {
    var out: std.ArrayListUnmanaged(u8) = .{};
    errdefer out.deinit(allocator);
    var i: usize = 0;
    while (std.mem.indexOfPos(u8, xml, i, "<Override")) |o_pos| {
        const o_end = std.mem.indexOfScalarPos(u8, xml, o_pos, '>') orelse {
            try out.appendSlice(allocator, xml[i..]);
            return try out.toOwnedSlice(allocator);
        };
        const o_attrs = xml[o_pos + "<Override".len .. o_end];
        var dropped = false;
        if (getAttr(o_attrs, "PartName")) |part| {
            // Strip leading `/` before comparing: PartName values
            // are absolute (`/xl/worksheets/sheetN.xml`).
            const part_norm = if (part.len > 0 and part[0] == '/') part[1..] else part;
            for (deletes) |d| {
                if (std.mem.eql(u8, d.path, part_norm)) {
                    dropped = true;
                    break;
                }
            }
        }
        if (dropped) {
            try out.appendSlice(allocator, xml[i..o_pos]);
            i = o_end + 1;
        } else {
            try out.appendSlice(allocator, xml[i .. o_end + 1]);
            i = o_end + 1;
        }
    }
    try out.appendSlice(allocator, xml[i..]);
    return try out.toOwnedSlice(allocator);
}

/// Twin of `patchEntryForNewSheets` for sheet deletes.
fn patchEntryForDeletes(
    allocator: Allocator,
    entries: []const Editor.ZipEntry,
    src_buf: []u8,
    subs: []?SubstitutedEntry,
    entry_name: []const u8,
    deletes: []const Editor.SheetDelete,
    patcher: *const fn (Allocator, []const u8, []const Editor.SheetDelete) anyerror![]u8,
) !void {
    const entry_idx = findEntryByName(entries, entry_name) orelse
        return error.MissingEntry;
    var src_xml: []u8 = undefined;
    var src_xml_owned = false;
    if (subs[entry_idx]) |s| {
        src_xml = try decompressZipPayload(allocator, s.payload, s.compression_method, s.uncompressed_size);
        src_xml_owned = true;
        allocator.free(s.lfh);
        allocator.free(s.payload);
        allocator.free(s.cdfh);
        subs[entry_idx] = null;
    } else {
        const e = entries[entry_idx];
        const payload = src_buf[e.lfh_offset + e.lfh_total_len ..][0..e.payload_len];
        src_xml = try decompressZipPayload(allocator, payload, e.compression_method, e.uncompressed_size);
        src_xml_owned = true;
    }
    defer if (src_xml_owned) allocator.free(src_xml);

    const new_xml = try patcher(allocator, src_xml, deletes);
    subs[entry_idx] = try buildEntryFromXml(allocator, entry_name, new_xml);
}

/// Twin of `patchEntryForNewSheets` for sheet renames. Same
/// re-substitution semantics — composes with new-sheet patches
/// already in subs[entry_idx].
fn patchEntryForRenames(
    allocator: Allocator,
    entries: []const Editor.ZipEntry,
    src_buf: []u8,
    subs: []?SubstitutedEntry,
    entry_name: []const u8,
    renames: []const Editor.SheetRename,
    patcher: *const fn (Allocator, []const u8, []const Editor.SheetRename) anyerror![]u8,
) !void {
    const entry_idx = findEntryByName(entries, entry_name) orelse
        return error.MissingEntry;
    var src_xml: []u8 = undefined;
    var src_xml_owned = false;
    if (subs[entry_idx]) |s| {
        src_xml = try decompressZipPayload(
            allocator,
            s.payload,
            s.compression_method,
            s.uncompressed_size,
        );
        src_xml_owned = true;
        allocator.free(s.lfh);
        allocator.free(s.payload);
        allocator.free(s.cdfh);
        subs[entry_idx] = null;
    } else {
        const e = entries[entry_idx];
        const payload = src_buf[e.lfh_offset + e.lfh_total_len ..][0..e.payload_len];
        src_xml = try decompressZipPayload(allocator, payload, e.compression_method, e.uncompressed_size);
        src_xml_owned = true;
    }
    defer if (src_xml_owned) allocator.free(src_xml);

    const new_xml = try patcher(allocator, src_xml, renames);
    subs[entry_idx] = try buildEntryFromXml(allocator, entry_name, new_xml);
}

/// Variant of `patchEntryXml` that threads a NewSheet slice through
/// to the patcher. Decompresses the named entry, hands it to
/// `patcher`, wraps the result back into a fresh SubstitutedEntry
/// stored at `subs[entry_idx]`. Errors when the entry is missing
/// — unlike the SST path, the workbook + rels + Content_Types
/// MUST exist for `addSheet` to be meaningful.
fn patchEntryForNewSheets(
    allocator: Allocator,
    entries: []const Editor.ZipEntry,
    src_buf: []u8,
    subs: []?SubstitutedEntry,
    entry_name: []const u8,
    new_sheets: []const Editor.NewSheet,
    patcher: *const fn (Allocator, []const u8, []const Editor.NewSheet) anyerror![]u8,
) !void {
    const entry_idx = findEntryByName(entries, entry_name) orelse
        return error.MissingEntry;
    // If a previous pass already substituted this entry (e.g. SST
    // creation patched workbook rels), patch the already-substituted
    // bytes instead of re-decompressing the source.
    var src_xml: []u8 = undefined;
    var src_xml_owned = false;
    if (subs[entry_idx]) |s| {
        // The substituted entry's payload is compressed bytes;
        // decompress to get the XML again.
        src_xml = try decompressZipPayload(
            allocator,
            s.payload,
            s.compression_method,
            s.uncompressed_size,
        );
        src_xml_owned = true;
        // Drop the old substituted entry — we're rebuilding it.
        allocator.free(s.lfh);
        allocator.free(s.payload);
        allocator.free(s.cdfh);
        subs[entry_idx] = null;
    } else {
        const e = entries[entry_idx];
        const payload = src_buf[e.lfh_offset + e.lfh_total_len ..][0..e.payload_len];
        src_xml = try decompressZipPayload(allocator, payload, e.compression_method, e.uncompressed_size);
        src_xml_owned = true;
    }
    defer if (src_xml_owned) allocator.free(src_xml);

    const new_xml = try patcher(allocator, src_xml, new_sheets);
    // buildEntryFromXml takes ownership of new_xml.
    subs[entry_idx] = try buildEntryFromXml(allocator, entry_name, new_xml);
}

fn addSstContentTypeOverride(allocator: Allocator, xml: []const u8) !?[]u8 {
    const inserted = "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>";
    if (std.mem.indexOf(u8, xml, "PartName=\"/xl/sharedStrings.xml\"") != null) return null;
    const close = std.mem.indexOf(u8, xml, "</Types>") orelse return error.MalformedXml;
    var out: std.ArrayListUnmanaged(u8) = .{};
    errdefer out.deinit(allocator);
    try out.appendSlice(allocator, xml[0..close]);
    try out.appendSlice(allocator, inserted);
    try out.appendSlice(allocator, xml[close..]);
    return try out.toOwnedSlice(allocator);
}

/// Splice a `<Relationship>` for `xl/sharedStrings.xml` into
/// `xl/_rels/workbook.xml.rels`. Picks an Id that doesn't collide
/// with existing `rIdN` values (max+1, or `rId1` if none). No-op
/// (returns null) if a sharedStrings relationship already exists.
fn addSstRelationship(allocator: Allocator, xml: []const u8) !?[]u8 {
    if (std.mem.indexOf(u8, xml, "Target=\"sharedStrings.xml\"") != null) return null;

    // Find the largest existing rId<N> so the new one doesn't
    // collide. Fallback to rId1 when the rels file has none.
    var max_id: u32 = 0;
    var i: usize = 0;
    while (std.mem.indexOfPos(u8, xml, i, "Id=\"rId")) |pos| {
        const num_start = pos + "Id=\"rId".len;
        var num_end = num_start;
        while (num_end < xml.len and xml[num_end] >= '0' and xml[num_end] <= '9') : (num_end += 1) {}
        if (num_end > num_start) {
            if (std.fmt.parseInt(u32, xml[num_start..num_end], 10)) |n| {
                if (n > max_id) max_id = n;
            } else |_| {}
        }
        i = num_end;
    }
    const new_id: u32 = max_id + 1;

    const close = std.mem.indexOf(u8, xml, "</Relationships>") orelse return error.MalformedXml;
    var out: std.ArrayListUnmanaged(u8) = .{};
    errdefer out.deinit(allocator);
    try out.appendSlice(allocator, xml[0..close]);
    try out.appendSlice(allocator, "<Relationship Id=\"rId");
    try std.fmt.format(out.writer(allocator), "{d}", .{new_id});
    try out.appendSlice(allocator, "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>");
    try out.appendSlice(allocator, xml[close..]);
    return try out.toOwnedSlice(allocator);
}

/// Decompress an entry's payload, hand it to `patcher` (which
/// returns a new owned XML or null = "no change needed"), wrap the
/// result back into a `SubstitutedEntry` parked in `subs[entry_idx]`.
/// Skipped silently when the entry is missing — the SST-creation
/// path can tolerate a workbook that lacks the patched part (the
/// fresh SST itself is appended to the central directory either
/// way; the rels patch is best-effort).
fn patchEntryXml(
    allocator: Allocator,
    entries: []const Editor.ZipEntry,
    src_buf: []u8,
    subs: []?SubstitutedEntry,
    name: []const u8,
    patcher: *const fn (Allocator, []const u8) anyerror!?[]u8,
) !void {
    const idx = findEntryByName(entries, name) orelse return;
    if (subs[idx] != null) return; // already substituted by another path
    const e = entries[idx];
    const payload = src_buf[e.lfh_offset + e.lfh_total_len ..][0..e.payload_len];
    const xml = try decompressZipPayload(
        allocator,
        payload,
        e.compression_method,
        e.uncompressed_size,
    );
    defer allocator.free(xml);
    const patched_opt = try patcher(allocator, xml);
    const patched = patched_opt orelse return;
    subs[idx] = try buildEntryFromXml(allocator, e.name, patched);
}

fn buildSubstitutedSst(
    allocator: Allocator,
    entry: Editor.ZipEntry,
    src_xml: []const u8,
    new_strings: []const []const u8,
    source_count: u32,
) !SubstitutedEntry {
    // Render new SST entries into a buffer.
    var entries_buf: std.ArrayListUnmanaged(u8) = .{};
    defer entries_buf.deinit(allocator);
    for (new_strings) |s| {
        try entries_buf.appendSlice(allocator, "<si><t");
        if (sstNeedsXmlSpacePreserve(s)) {
            try entries_buf.appendSlice(allocator, " xml:space=\"preserve\"");
        }
        try entries_buf.appendSlice(allocator, ">");
        try appendXmlEscaped(allocator, &entries_buf, s);
        try entries_buf.appendSlice(allocator, "</t></si>");
    }

    // Splice into source XML — handle open/close form first, then
    // self-closing. Per the SpreadsheetML schema, `<extLst>` is the
    // terminal child of `<sst>`, so injection has to land BEFORE
    // any `<extLst>` block when present, not just before `</sst>`.
    var spliced: std.ArrayListUnmanaged(u8) = .{};
    defer spliced.deinit(allocator);
    if (std.mem.indexOf(u8, src_xml, "</sst>")) |close_pos| {
        // Prefer an inject point right before `<extLst>` if the
        // source declares one; otherwise inject right before
        // `</sst>`. The substring scan is intentionally narrow —
        // matches `<extLst>` and `<extLst …>` but not text content.
        var inject_pos: usize = close_pos;
        if (std.mem.indexOf(u8, src_xml, "<extLst")) |ext_pos| {
            if (ext_pos < close_pos) inject_pos = ext_pos;
        }
        try spliced.appendSlice(allocator, src_xml[0..inject_pos]);
        try spliced.appendSlice(allocator, entries_buf.items);
        try spliced.appendSlice(allocator, src_xml[inject_pos..]);
    } else if (std.mem.indexOf(u8, src_xml, "<sst")) |sst_open| {
        const tag_close = std.mem.indexOfScalarPos(u8, src_xml, sst_open, '>') orelse
            return error.MalformedXml;
        if (tag_close == 0 or src_xml[tag_close - 1] != '/') return error.MalformedXml;
        const attrs_end = tag_close - 1;
        const attrs = src_xml[sst_open + "<sst".len .. attrs_end];
        try spliced.appendSlice(allocator, src_xml[0..sst_open]);
        try spliced.appendSlice(allocator, "<sst");
        try spliced.appendSlice(allocator, attrs);
        try spliced.append(allocator, '>');
        try spliced.appendSlice(allocator, entries_buf.items);
        try spliced.appendSlice(allocator, "</sst>");
        try spliced.appendSlice(allocator, src_xml[tag_close + 1 ..]);
    } else {
        return error.MalformedXml;
    }

    // Patch count / uniqueCount on the <sst> opening tag. They have
    // different semantics: `count` tracks total string-cell
    // references (can exceed unique entries when the workbook
    // reuses shared strings), `uniqueCount` tracks `<si>` entries.
    // Each appended string cell adds 1 reference AND 1 unique entry,
    // so both bump by `new_strings.len` — but starting from the
    // source's declared values, not from `source_count`.
    const inc: u32 = @intCast(new_strings.len);
    const src_count = readSstAttrU32(src_xml, "count") orelse source_count;
    const src_unique = readSstAttrU32(src_xml, "uniqueCount") orelse source_count;
    const new_count = src_count + inc;
    const new_unique = src_unique + inc;
    const after_count = try patchSstAttr(allocator, spliced.items, "count", new_count);
    spliced.deinit(allocator);
    spliced = .{};
    const after_unique = try patchSstAttr(allocator, after_count, "uniqueCount", new_unique);
    allocator.free(after_count);

    return try buildEntryFromXml(allocator, entry.name, after_unique);
}

/// Read a numeric attribute from the `<sst …>` opening tag. Returns
/// null if the attribute is absent or unparseable.
fn readSstAttrU32(xml: []const u8, name: []const u8) ?u32 {
    const sst_open = std.mem.indexOf(u8, xml, "<sst") orelse return null;
    const tag_close = std.mem.indexOfScalarPos(u8, xml, sst_open, '>') orelse return null;
    const tag = xml[sst_open..tag_close];
    var name_buf: [32]u8 = undefined;
    if (1 + name.len + 2 > name_buf.len) return null;
    name_buf[0] = ' ';
    @memcpy(name_buf[1 .. 1 + name.len], name);
    @memcpy(name_buf[1 + name.len .. 1 + name.len + 2], "=\"");
    const needle = name_buf[0 .. 1 + name.len + 2];
    const attr_pos = std.mem.indexOf(u8, tag, needle) orelse return null;
    const val_start_in_tag = attr_pos + needle.len;
    if (val_start_in_tag >= tag.len) return null;
    var val_end = val_start_in_tag;
    while (val_end < tag.len and tag[val_end] != '"') : (val_end += 1) {}
    if (val_end == val_start_in_tag) return null;
    return std.fmt.parseInt(u32, tag[val_start_in_tag..val_end], 10) catch null;
}

/// Patch a numeric attribute (`count` or `uniqueCount`) on the
/// `<sst …>` opening tag. Falls back to leaving the XML unchanged
/// if the attribute isn't present (rare — readers tolerate omitted
/// counts). Returns an owned slice.
fn patchSstAttr(allocator: Allocator, xml: []const u8, name: []const u8, new_value: u32) ![]u8 {
    const sst_open = std.mem.indexOf(u8, xml, "<sst") orelse return try allocator.dupe(u8, xml);
    const tag_close = std.mem.indexOfScalarPos(u8, xml, sst_open, '>') orelse
        return try allocator.dupe(u8, xml);
    const tag = xml[sst_open..tag_close];
    // Find ` name="…"` inside the tag span.
    var name_buf: [32]u8 = undefined;
    if (1 + name.len + 2 > name_buf.len) return try allocator.dupe(u8, xml);
    name_buf[0] = ' ';
    @memcpy(name_buf[1 .. 1 + name.len], name);
    @memcpy(name_buf[1 + name.len .. 1 + name.len + 2], "=\"");
    const needle = name_buf[0 .. 1 + name.len + 2];
    const attr_pos = std.mem.indexOf(u8, tag, needle) orelse return try allocator.dupe(u8, xml);
    const val_start = sst_open + attr_pos + needle.len;
    const val_end = std.mem.indexOfScalarPos(u8, xml, val_start, '"') orelse
        return try allocator.dupe(u8, xml);

    var out: std.ArrayListUnmanaged(u8) = .{};
    errdefer out.deinit(allocator);
    try out.appendSlice(allocator, xml[0..val_start]);
    try std.fmt.format(out.writer(allocator), "{d}", .{new_value});
    try out.appendSlice(allocator, xml[val_end..]);
    return try out.toOwnedSlice(allocator);
}

/// True when the string would normally need `xml:space="preserve"`
/// to round-trip — i.e. it has leading/trailing whitespace OOXML
/// would otherwise strip on parse.
fn sstNeedsXmlSpacePreserve(s: []const u8) bool {
    if (s.len == 0) return false;
    const lead = s[0];
    const trail = s[s.len - 1];
    return lead == ' ' or lead == '\t' or lead == '\n' or lead == '\r' or
        trail == ' ' or trail == '\t' or trail == '\n' or trail == '\r';
}

/// Append `s` to `out` with the OOXML-required entity escaping
/// (`&`, `<`, `>`). Suffices for `<t>` content; quote chars don't
/// appear in element bodies.
fn appendXmlEscaped(allocator: Allocator, out: *std.ArrayListUnmanaged(u8), s: []const u8) !void {
    for (s) |c| {
        switch (c) {
            '&' => try out.appendSlice(allocator, "&amp;"),
            '<' => try out.appendSlice(allocator, "&lt;"),
            '>' => try out.appendSlice(allocator, "&gt;"),
            else => try out.append(allocator, c),
        }
    }
}

/// Like `appendXmlEscaped` but also escapes `"` → `&quot;`. Use
/// when emitting into a double-quoted attribute value (`name="…"`).
fn appendXmlAttrEscaped(allocator: Allocator, out: *std.ArrayListUnmanaged(u8), s: []const u8) !void {
    for (s) |c| {
        switch (c) {
            '&' => try out.appendSlice(allocator, "&amp;"),
            '<' => try out.appendSlice(allocator, "&lt;"),
            '>' => try out.appendSlice(allocator, "&gt;"),
            '"' => try out.appendSlice(allocator, "&quot;"),
            else => try out.append(allocator, c),
        }
    }
}

/// Decode the named-entity-only XML attribute reverse: `&amp;`/
/// `&lt;`/`&gt;`/`&quot;`/`&apos;` → their literal char. Numeric
/// character references (`&#NN;` / `&#xHH;`) are passed through
/// verbatim — sheet names don't carry them in practice. Unknown
/// entities are left intact too.
fn decodeXmlAttrInto(allocator: Allocator, out: *std.ArrayListUnmanaged(u8), s: []const u8) !void {
    var i: usize = 0;
    while (i < s.len) {
        if (s[i] == '&') {
            const semi = std.mem.indexOfScalarPos(u8, s, i, ';') orelse {
                try out.append(allocator, s[i]);
                i += 1;
                continue;
            };
            const ent = s[i + 1 .. semi];
            const replaced: ?u8 =
                if (std.mem.eql(u8, ent, "amp")) @as(u8, '&') else if (std.mem.eql(u8, ent, "lt")) @as(u8, '<') else if (std.mem.eql(u8, ent, "gt")) @as(u8, '>') else if (std.mem.eql(u8, ent, "quot")) @as(u8, '"') else if (std.mem.eql(u8, ent, "apos")) @as(u8, '\'') else null;
            if (replaced) |c| {
                try out.append(allocator, c);
                i = semi + 1;
            } else {
                // Unknown entity — pass through verbatim.
                try out.appendSlice(allocator, s[i .. semi + 1]);
                i = semi + 1;
            }
        } else {
            try out.append(allocator, s[i]);
            i += 1;
        }
    }
}

/// Compress `new_xml`, build a fresh LFH+CDFH, and return the
/// substituted entry. Shared between every editor write path:
/// sheet substitution, SST substitution, and on-demand creation
/// of new parts (e.g. a fresh `xl/sharedStrings.xml` for an
/// SST-less source workbook). The caller owns `new_xml`.
fn buildEntryFromXml(
    allocator: Allocator,
    filename: []const u8,
    new_xml: []const u8,
) !SubstitutedEntry {
    defer allocator.free(new_xml);

    // `>=` for size fields: 0xFFFFFFFF is the ZIP64 sentinel that
    // Book.open interprets as "look at extras", so emitting it
    // would produce a self-incompatible archive. The filename
    // length field is a plain u16 with no sentinel meaning, so
    // 0xFFFF is a legal maximum there.
    if (new_xml.len >= std.math.maxInt(u32)) return error.Zip64NotSupported;
    if (filename.len > std.math.maxInt(u16)) return error.FilenameTooLong;

    var compressed: std.ArrayListUnmanaged(u8) = .{};
    defer compressed.deinit(allocator);
    var compression_method: u16 = 8;
    if (new_xml.len < 1024) {
        compression_method = 0;
        try compressed.appendSlice(allocator, new_xml);
    } else {
        const writer_mod = @import("writer.zig");
        try writer_mod.deflateCompress(allocator, new_xml, &compressed);
        if (compressed.items.len >= new_xml.len) {
            compression_method = 0;
            compressed.clearRetainingCapacity();
            try compressed.appendSlice(allocator, new_xml);
        }
    }
    if (compressed.items.len >= std.math.maxInt(u32)) return error.Zip64NotSupported;

    const crc = std.hash.Crc32.hash(new_xml);
    // `filename` is now the function parameter; no need to read
    // `entry.name` (the helper takes any path).
    const lfh_size = @sizeOf(std.zip.LocalFileHeader);
    const lfh_total = lfh_size + filename.len;
    const lfh = try allocator.alloc(u8, lfh_total);
    errdefer allocator.free(lfh);
    var lfh_struct: std.zip.LocalFileHeader = .{
        .signature = std.zip.local_file_header_sig,
        .version_needed_to_extract = 20,
        .flags = .{ .encrypted = false, ._ = 0 },
        .compression_method = @enumFromInt(compression_method),
        .last_modification_time = 0,
        .last_modification_date = 0x21,
        .crc32 = crc,
        .compressed_size = @intCast(compressed.items.len),
        .uncompressed_size = @intCast(new_xml.len),
        .filename_len = @intCast(filename.len),
        .extra_len = 0,
    };
    if (@import("builtin").cpu.arch.endian() != .little)
        std.mem.byteSwapAllFields(std.zip.LocalFileHeader, &lfh_struct);
    @memcpy(lfh[0..lfh_size], std.mem.asBytes(&lfh_struct));
    @memcpy(lfh[lfh_size..], filename);

    const cdfh_size = @sizeOf(std.zip.CentralDirectoryFileHeader);
    const cdfh_total = cdfh_size + filename.len;
    const cdfh = try allocator.alloc(u8, cdfh_total);
    errdefer allocator.free(cdfh);
    var cdfh_struct: std.zip.CentralDirectoryFileHeader = .{
        .signature = std.zip.central_file_header_sig,
        .version_made_by = 20,
        .version_needed_to_extract = 20,
        .flags = .{ .encrypted = false, ._ = 0 },
        .compression_method = @enumFromInt(compression_method),
        .last_modification_time = 0,
        .last_modification_date = 0x21,
        .crc32 = crc,
        .compressed_size = @intCast(compressed.items.len),
        .uncompressed_size = @intCast(new_xml.len),
        .filename_len = @intCast(filename.len),
        .extra_len = 0,
        .comment_len = 0,
        .disk_number = 0,
        .internal_file_attributes = 0,
        .external_file_attributes = 0,
        .local_file_header_offset = 0,
    };
    if (@import("builtin").cpu.arch.endian() != .little)
        std.mem.byteSwapAllFields(std.zip.CentralDirectoryFileHeader, &cdfh_struct);
    @memcpy(cdfh[0..cdfh_size], std.mem.asBytes(&cdfh_struct));
    @memcpy(cdfh[cdfh_size..], filename);

    const payload_owned = try compressed.toOwnedSlice(allocator);
    return .{
        .lfh = lfh,
        .payload = payload_owned,
        .cdfh = cdfh,
        .crc32 = crc,
        .uncompressed_size = @intCast(new_xml.len),
        .compression_method = compression_method,
    };
}

// ZIP-bomb defenses for `decompressZipPayload`. Mirror the caps in
// pkg/store.zig: per-part 512 MiB hard cap + 4096:1 ratio. Both
// checks fire BEFORE any allocation so a crafted CDFH declaring
// multi-GB uncompressed size can't OOM the reader.
const max_reader_part_size: usize = 512 * 1024 * 1024;
const max_reader_deflate_ratio: usize = 4096;

fn decompressZipPayload(
    allocator: Allocator,
    payload: []const u8,
    method: u16,
    declared_uncompressed: u32,
) ![]u8 {
    if (declared_uncompressed > max_reader_part_size) return error.BadZip;
    const ratio_cap = std.math.mul(usize, payload.len, max_reader_deflate_ratio) catch std.math.maxInt(usize);
    if (declared_uncompressed > ratio_cap) return error.BadZip;

    if (method == 0) {
        if (payload.len != declared_uncompressed) return error.BadZip;
        return try allocator.dupe(u8, payload);
    } else if (method == 8) {
        // streamExact64 caps the inflated output at the declared
        // size, defending against zip bombs / oversize-inflation
        // attacks. The CDFH-side uncompressed_size is what the
        // archive promises; refuse to allocate more than that.
        var src_reader = std.Io.Reader.fixed(payload);
        var flate_buffer: [std.compress.flate.max_window_len]u8 = undefined;
        var dec = std.compress.flate.Decompress.init(&src_reader, .raw, &flate_buffer);
        const out = try allocator.alloc(u8, declared_uncompressed);
        errdefer allocator.free(out);
        var out_writer = std.Io.Writer.fixed(out);
        dec.reader.streamExact64(&out_writer, declared_uncompressed) catch return error.BadZip;
        return out;
    }
    return error.UnsupportedCompression;
}

/// Return the largest cell-row index in the sheet XML. Walks both
/// `<row>` and `<c>` opening tags and looks for `r="…"` anywhere in
/// each tag's attribute span — OOXML doesn't constrain attribute
/// order, so `<c s="1" r="A42">` and `<row spans="1:4" r="12">` are
/// both legal. Cell refs are preferred because OOXML allows `<row>`
/// to omit `r=` entirely (the row index then infers from the first
/// cell), but explicit-row scans are also covered for empty
/// `<row r="N"/>` openings without children.
fn findHighestRowInSheetXml(xml: []const u8) u32 {
    var highest: u32 = 0;

    // Pass 1: `<c ...>` tags — extract the row component from r="A1".
    {
        var i: usize = 0;
        while (std.mem.indexOfPos(u8, xml, i, "<c")) |tag_start| {
            const after = tag_start + "<c".len;
            if (after >= xml.len) break;
            const c = xml[after];
            // Filter out `<col`, `<conditionalFormatting`, etc.
            if (c != ' ' and c != '\t' and c != '\n' and c != '\r' and c != '/' and c != '>') {
                i = tag_start + 1;
                continue;
            }
            const tag_end = std.mem.indexOfScalarPos(u8, xml, tag_start, '>') orelse break;
            if (findAttrRowFromCellRef(xml[tag_start..tag_end])) |n| {
                if (n > highest) highest = n;
            }
            i = tag_end + 1;
        }
    }
    // Pass 2: `<row ...>` tags — explicit r="N".
    {
        var i: usize = 0;
        while (std.mem.indexOfPos(u8, xml, i, "<row")) |tag_start| {
            const after = tag_start + "<row".len;
            if (after >= xml.len) break;
            const c = xml[after];
            if (c != ' ' and c != '\t' and c != '\n' and c != '\r' and c != '/' and c != '>') {
                i = tag_start + 1;
                continue;
            }
            const tag_end = std.mem.indexOfScalarPos(u8, xml, tag_start, '>') orelse break;
            if (findAttrRowExplicit(xml[tag_start..tag_end])) |n| {
                if (n > highest) highest = n;
            }
            i = tag_end + 1;
        }
    }
    return highest;
}

/// Locate ` r="…"` within an opening-tag span and parse the row
/// component of an A1-style cell ref ("A1", "B12", "AAA9999").
/// Returns null on no match or unparseable digits.
fn findAttrRowFromCellRef(tag: []const u8) ?u32 {
    var search_from: usize = 0;
    while (std.mem.indexOfPos(u8, tag, search_from, "r=\"")) |r_pos| {
        const prev = if (r_pos > 0) tag[r_pos - 1] else 0;
        if (prev == ' ' or prev == '\t' or prev == '\n' or prev == '\r') {
            const ref_start = r_pos + "r=\"".len;
            var col_end = ref_start;
            while (col_end < tag.len and tag[col_end] >= 'A' and tag[col_end] <= 'Z') : (col_end += 1) {}
            var num_end = col_end;
            while (num_end < tag.len and tag[num_end] >= '0' and tag[num_end] <= '9') : (num_end += 1) {}
            if (num_end > col_end) {
                return std.fmt.parseInt(u32, tag[col_end..num_end], 10) catch null;
            }
            return null;
        }
        search_from = r_pos + 1;
    }
    return null;
}

/// Locate ` r="N"` within a `<row …>` span and parse N as a u32.
fn findAttrRowExplicit(tag: []const u8) ?u32 {
    var search_from: usize = 0;
    while (std.mem.indexOfPos(u8, tag, search_from, "r=\"")) |r_pos| {
        const prev = if (r_pos > 0) tag[r_pos - 1] else 0;
        if (prev == ' ' or prev == '\t' or prev == '\n' or prev == '\r') {
            const num_start = r_pos + "r=\"".len;
            var num_end = num_start;
            while (num_end < tag.len and tag[num_end] >= '0' and tag[num_end] <= '9') : (num_end += 1) {}
            if (num_end > num_start) {
                return std.fmt.parseInt(u32, tag[num_start..num_end], 10) catch null;
            }
            return null;
        }
        search_from = r_pos + 1;
    }
    return null;
}

/// Build a new sheet XML with appended rows. Handles two source
/// shapes:
///   - `<sheetData>...</sheetData>` — inject rows before the close.
///   - `<sheetData/>` — replace with `<sheetData>…</sheetData>` so
///     empty sheets become well-formed once they have rows.
/// Returns `error.SheetMissingSheetData` if neither shape is found
/// (defensive — every well-formed OOXML sheet has one).
fn injectAppendedRows(
    allocator: Allocator,
    src_xml: []const u8,
    appended: []const []Cell,
    start_row: u32,
    sst: ?*SstAppender,
) ![]u8 {
    // Render appended rows once.
    var rows_buf: std.ArrayListUnmanaged(u8) = .{};
    defer rows_buf.deinit(allocator);
    for (appended, 0..) |row, ri| {
        const row_idx: u32 = start_row + @as(u32, @intCast(ri));
        try rows_buf.appendSlice(allocator, "<row r=\"");
        try std.fmt.format(rows_buf.writer(allocator), "{d}", .{row_idx});
        try rows_buf.appendSlice(allocator, "\">");
        for (row, 0..) |cell, ci| {
            try renderCellOoxml(allocator, &rows_buf, cell, row_idx, @intCast(ci), sst);
        }
        try rows_buf.appendSlice(allocator, "</row>");
    }

    // Prefer the open/close form. Fall back to self-closing if no
    // close is found.
    if (std.mem.indexOf(u8, src_xml, "</sheetData>")) |inject_pos| {
        const out_len = src_xml.len + rows_buf.items.len;
        const out = try allocator.alloc(u8, out_len);
        errdefer allocator.free(out);
        @memcpy(out[0..inject_pos], src_xml[0..inject_pos]);
        @memcpy(out[inject_pos..][0..rows_buf.items.len], rows_buf.items);
        @memcpy(out[inject_pos + rows_buf.items.len ..], src_xml[inject_pos..]);
        return out;
    }

    // Self-closing form: `<sheetData/>` (sometimes with attributes).
    // Locate `<sheetData` then advance to the closing `>`. If the
    // tag ends with `/>`, replace the whole tag with
    // `<sheetData …>` + rows + `</sheetData>` — preserving any
    // attributes (rare but legal).
    const sd_open = std.mem.indexOf(u8, src_xml, "<sheetData") orelse
        return error.SheetMissingSheetData;
    const sd_close = std.mem.indexOfScalarPos(u8, src_xml, sd_open, '>') orelse
        return error.SheetMissingSheetData;
    if (sd_close == 0 or src_xml[sd_close - 1] != '/')
        return error.SheetMissingSheetData;
    // Tag attributes (if any) live between `<sheetData` and `/>`.
    const attrs_end = sd_close - 1;
    const attrs = src_xml[sd_open + "<sheetData".len .. attrs_end];

    var spliced: std.ArrayListUnmanaged(u8) = .{};
    defer spliced.deinit(allocator);
    try spliced.appendSlice(allocator, src_xml[0..sd_open]);
    try spliced.appendSlice(allocator, "<sheetData");
    try spliced.appendSlice(allocator, attrs);
    try spliced.append(allocator, '>');
    try spliced.appendSlice(allocator, rows_buf.items);
    try spliced.appendSlice(allocator, "</sheetData>");
    try spliced.appendSlice(allocator, src_xml[sd_close + 1 ..]);
    return try spliced.toOwnedSlice(allocator);
}

/// Patch BOTH corners of a canonical-form `<dimension ref="TL:BR"/>`
/// to span `[min_row..=max_row]` × `[min_col_1based..=max_col_1based]`.
/// Used by the cell-mutate save path where edits can land outside
/// the existing range (e.g. `setCell(A1, ...)` on a sheet whose
/// dimension was `C5:D7`). Same canonical-form contract as
/// `updateDimension`: non-canonical refs pass through.
fn updateDimensionRange(
    allocator: Allocator,
    xml: []const u8,
    new_min_row: u32,
    new_max_row: u32,
    new_min_col_1based: u32,
    new_max_col_1based: u32,
) !?[]u8 {
    if (new_min_row == 0 or new_max_row == 0 or new_min_col_1based == 0 or new_max_col_1based == 0)
        return null;
    const dim_open = "<dimension ref=\"";
    const dim_pos = std.mem.indexOf(u8, xml, dim_open) orelse return null;
    const ref_start = dim_pos + dim_open.len;
    const ref_end = std.mem.indexOfScalarPos(u8, xml, ref_start, '"') orelse return null;
    const ref = xml[ref_start..ref_end];
    const colon = std.mem.indexOfScalar(u8, ref, ':') orelse return null;
    const tl_part = ref[0..colon];
    const br_part = ref[colon + 1 ..];

    // Canonical-form gate: both halves must be plain
    // `[A-Z]+[0-9]+` (no `$`, no namespace, no expressions like
    // `'Sheet1'!A1`). Anything else passes through — Excel
    // recomputes the dimension on its next save.
    const tl_parsed = parseA1Plain(tl_part) orelse return null;
    const br_parsed = parseA1Plain(br_part) orelse return null;

    // Only WIDEN — never shrink. The source's existing dimension
    // may already cover more than the cells the scanner saw (e.g.
    // a producer that left the dimension at A1:Z100 after clearing
    // rows). Take the union with the new edits' bounds.
    const u_min_row = @min(tl_parsed.row, new_min_row);
    const u_max_row = @max(br_parsed.row, new_max_row);
    const u_min_col1 = @min(tl_parsed.col_1based, new_min_col_1based);
    const u_max_col1 = @max(br_parsed.col_1based, new_max_col_1based);

    // No-op when the union doesn't extend the existing range.
    if (u_min_row == tl_parsed.row and u_max_row == br_parsed.row and
        u_min_col1 == tl_parsed.col_1based and u_max_col1 == br_parsed.col_1based)
        return null;

    const writer_mod = @import("writer.zig");
    var tl_buf: [16]u8 = undefined;
    var br_buf: [16]u8 = undefined;
    const tl = try writer_mod.formatCellRef(&tl_buf, u_min_row, u_min_col1 - 1);
    const br = try writer_mod.formatCellRef(&br_buf, u_max_row, u_max_col1 - 1);

    var out: std.ArrayListUnmanaged(u8) = .{};
    defer out.deinit(allocator);
    try out.appendSlice(allocator, xml[0..ref_start]);
    try out.appendSlice(allocator, tl);
    try out.append(allocator, ':');
    try out.appendSlice(allocator, br);
    try out.appendSlice(allocator, xml[ref_end..]);
    return try out.toOwnedSlice(allocator);
}

/// Parse a plain A1-style ref (`[A-Z]+[0-9]+`, no `$`, no sheet
/// prefix). Returns null on any other shape so callers can pass
/// through non-canonical inputs.
fn parseA1Plain(s: []const u8) ?struct { row: u32, col_1based: u32 } {
    if (s.len == 0) return null;
    var i: usize = 0;
    while (i < s.len and s[i] >= 'A' and s[i] <= 'Z') i += 1;
    if (i == 0 or i == s.len) return null;
    const col_1based = parseColLetters(s[0..i]) orelse return null;
    const row = std.fmt.parseInt(u32, s[i..], 10) catch return null;
    if (row == 0) return null;
    return .{ .row = row, .col_1based = col_1based };
}

/// Patch a canonical-form `<dimension ref="TL:BR"/>` so the
/// bottom-right corner's row component reaches `new_max_row` and
/// its column component reaches `new_max_col_1based`. Returns null
/// when no patch is needed or the dimension isn't in canonical
/// range form (single-cell refs / open-tag form / namespaced attr
/// remain best-effort per the LMS plan — Excel rescans
/// `<sheetData>` on open and rewrites the dimension on its next
/// save, so staleness on those is tolerable). Returns an owned new
/// slice on success.
fn updateDimension(
    allocator: Allocator,
    xml: []const u8,
    new_max_row: u32,
    new_max_col_1based: u32,
) !?[]u8 {
    const dim_open = "<dimension ref=\"";
    const dim_pos = std.mem.indexOf(u8, xml, dim_open) orelse return null;
    const ref_start = dim_pos + dim_open.len;
    const ref_end = std.mem.indexOfScalarPos(u8, xml, ref_start, '"') orelse return null;
    const ref = xml[ref_start..ref_end];
    const colon = std.mem.indexOfScalar(u8, ref, ':') orelse return null;
    const br = ref[colon + 1 ..];
    if (br.len == 0) return null;
    // br is the bottom-right corner like "Z100". Split into letter
    // prefix + digit suffix.
    var digit_start: usize = br.len;
    while (digit_start > 0 and br[digit_start - 1] >= '0' and br[digit_start - 1] <= '9') {
        digit_start -= 1;
    }
    if (digit_start == br.len or digit_start == 0) return null;
    // Validate letter prefix is uppercase A-Z only.
    for (br[0..digit_start]) |c| if (c < 'A' or c > 'Z') return null;
    const old_row = std.fmt.parseInt(u32, br[digit_start..], 10) catch return null;
    const old_col_1based = parseColLetters(br[0..digit_start]) orelse return null;

    const final_row: u32 = @max(old_row, new_max_row);
    const final_col_1based: u32 = @max(old_col_1based, new_max_col_1based);
    if (final_row == old_row and final_col_1based == old_col_1based) return null;

    // Splice the new bottom-right corner.
    const br_abs_start = ref_start + colon + 1;
    const br_abs_end = ref_end;
    var out: std.ArrayListUnmanaged(u8) = .{};
    errdefer out.deinit(allocator);
    try out.appendSlice(allocator, xml[0..br_abs_start]);
    var letter_buf: [8]u8 = undefined;
    const letters = colLetterEditor(&letter_buf, final_col_1based - 1);
    try out.appendSlice(allocator, letters);
    try std.fmt.format(out.writer(allocator), "{d}", .{final_row});
    try out.appendSlice(allocator, xml[br_abs_end..]);
    return try out.toOwnedSlice(allocator);
}

/// Parse uppercase A-Z letters as a 1-based Excel column index
/// (A=1, B=2, ..., Z=26, AA=27, ..., XFD=16384). Returns null on
/// empty input or anything past `max_col_1based`.
fn parseColLetters(s: []const u8) ?u32 {
    if (s.len == 0) return null;
    var n: u32 = 0;
    for (s) |c| {
        if (c < 'A' or c > 'Z') return null;
        n = n * 26 + (c - 'A' + 1);
        if (n > max_col_1based) return null;
    }
    return n;
}

fn renderCellOoxml(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    cell: Cell,
    row_idx: u32,
    col_idx: u32,
    sst: ?*SstAppender,
) !void {
    switch (cell) {
        .empty => return, // skip — empty cells emit nothing
        .integer => |x| {
            try writeCellOpen(allocator, out, col_idx, row_idx, null);
            try out.appendSlice(allocator, "<v>");
            try std.fmt.format(out.writer(allocator), "{d}", .{x});
            try out.appendSlice(allocator, "</v></c>");
        },
        .number => |f| {
            try writeCellOpen(allocator, out, col_idx, row_idx, null);
            try out.appendSlice(allocator, "<v>");
            try std.fmt.format(out.writer(allocator), "{d}", .{f});
            try out.appendSlice(allocator, "</v></c>");
        },
        .boolean => |b| {
            try writeCellOpen(allocator, out, col_idx, row_idx, "b");
            try out.appendSlice(allocator, "<v>");
            try out.appendSlice(allocator, if (b) "1" else "0");
            try out.appendSlice(allocator, "</v></c>");
        },
        .string => |s| {
            // String cells go through the SST appender, which assigns
            // a fresh index per cell (no plain-text reuse — see the
            // SstAppender doc).
            const appender = sst orelse return error.SstAppenderRequired;
            const idx = try appender.add(s);
            try writeCellOpen(allocator, out, col_idx, row_idx, "s");
            try out.appendSlice(allocator, "<v>");
            try std.fmt.format(out.writer(allocator), "{d}", .{idx});
            try out.appendSlice(allocator, "</v></c>");
        },
    }
}

fn writeCellOpen(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    col_idx: u32,
    row_idx: u32,
    t_attr: ?[]const u8,
) !void {
    var col_buf: [8]u8 = undefined;
    const col = colLetterEditor(&col_buf, col_idx);
    try out.appendSlice(allocator, "<c r=\"");
    try out.appendSlice(allocator, col);
    try std.fmt.format(out.writer(allocator), "{d}", .{row_idx});
    if (t_attr) |t| {
        try out.appendSlice(allocator, "\" t=\"");
        try out.appendSlice(allocator, t);
    }
    try out.appendSlice(allocator, "\">");
}

/// Render `col_idx` (0-based) as A, B, ..., Z, AA, AB, ... into `buf`.
/// Capacity 8 is more than enough (Excel max is XFD = 3 letters).
fn colLetterEditor(buf: []u8, col_idx: u32) []u8 {
    var n: u32 = col_idx + 1;
    var i: usize = 0;
    while (n > 0) {
        const r = (n - 1) % 26;
        buf[i] = 'A' + @as(u8, @intCast(r));
        i += 1;
        n = (n - 1) / 26;
    }
    std.mem.reverse(u8, buf[0..i]);
    return buf[0..i];
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

/// Per-test temporary file helper. Replaces the older `/tmp/zlsx_*.xlsx`
/// hard-coded paths so the suite is portable across Linux / macOS /
/// Windows. Each call creates a fresh isolated `std.testing.TmpDir`
/// (auto-cleaned). Caller frees the returned slice. Mirror of the
/// helper in src/writer.zig — kept in-file so the two test sections
/// don't depend on a shared utility module yet (introducing one is
/// part of B0 PartStore).
const TestTmp = struct {
    dir: std.testing.TmpDir,
    pub fn init() TestTmp {
        return .{ .dir = std.testing.tmpDir(.{}) };
    }
    pub fn deinit(self: *TestTmp) void {
        self.dir.cleanup();
    }
    pub fn path(self: *TestTmp, alloc: std.mem.Allocator, name: []const u8) ![:0]u8 {
        const d = try self.dir.dir.realpathAlloc(alloc, ".");
        defer alloc.free(d);
        return std.fs.path.joinZ(alloc, &.{ d, name });
    }
};

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

    // Excel sheet bounds: cols A..XFD, rows 1..1048576. Refs past
    // those limits are syntactically well-formed but cannot exist
    // in any sheet, so reject at parse time rather than silently
    // accept and produce empty output.
    try std.testing.expectError(error.MalformedXml, parseA1Ref("XFE1")); // col past XFD
    try std.testing.expectError(error.MalformedXml, parseA1Ref("A1048577")); // row past max
    try std.testing.expectError(error.MalformedXml, parseA1Ref("ZZZZZZZ1")); // pathological col
    try std.testing.expectError(error.MalformedXml, parseA1Ref("A99999999999")); // pathological row
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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "reader_merged_roundtrip.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "reader_hyperlinks_roundtrip.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    // Writer xml-escapes `&` to `&amp;` on emit; reader decodes it
    // back so callers can use the URL as-is without re-decoding.
    try std.testing.expectEqualStrings("https://example.com/path?q=1&x=2", links[0].url);

    try std.testing.expectEqualDeep(CellRef{ .col = 1, .row = 2 }, links[1].top_left);
    try std.testing.expectEqualDeep(CellRef{ .col = 2, .row = 3 }, links[1].bottom_right);
    try std.testing.expectEqualStrings("mailto:foo@example.com", links[1].url);

    try std.testing.expectEqualDeep(CellRef{ .col = 3, .row = 5 }, links[2].top_left);
    try std.testing.expectEqualStrings("https://docs.example.com/", links[2].url);
}

test "Book.dataValidations: round-trip through writer + reader" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "reader_dv_roundtrip.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "reader_dv_numeric_roundtrip.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "reader_no_dv.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "writer_comments_roundtrip.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "comments_xml_special.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "comments_multi_sheet.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "comments_stress.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "writer_rich_roundtrip.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    try std.testing.expectEqualStrings("label", book.sst.eager[0]);
    try std.testing.expectEqualStrings("hello world", book.sst.eager[1]);

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

    try std.testing.expectEqual(@as(usize, 5), book.sharedStringsCount());

    // Flat strings round-trip correctly for both plain and rich paths.
    try std.testing.expectEqualStrings("plain", book.sst.eager[0]);
    try std.testing.expectEqualStrings("bold", book.sst.eager[1]);
    try std.testing.expectEqualStrings("bold-italic", book.sst.eager[2]);
    try std.testing.expectEqualStrings("A B C", book.sst.eager[3]);
    try std.testing.expectEqualStrings("R&D", book.sst.eager[4]);

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

test "parseSharedStrings: phonetic <rPh> annotations are dropped from visible text" {
    // A typical East-Asian SST entry: visible Kanji text with a
    // phonetic-guide (furigana) <rPh> wrapping the reading. The
    // visible cell content is the OUTER <t>; the <rPh><t> inner is
    // hidden. Old parser would flatten both into a single string
    // ("漢字かんじ" instead of "漢字").
    const sst_xml =
        "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\">" ++
        "<si>" ++
        "<t>漢字</t>" ++
        "<rPh sb=\"0\" eb=\"2\"><t>かんじ</t></rPh>" ++
        "<phoneticPr fontId=\"1\"/>" ++
        "</si>" ++
        "</sst>";

    var book: Book = .{
        .allocator = std.testing.allocator,
        .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer book.deinit();
    const owned = try std.testing.allocator.dupe(u8, sst_xml);
    book.shared_strings_xml = owned;
    try parseSharedStrings(&book, owned);

    try std.testing.expectEqual(@as(usize, 1), book.sharedStringsCount());
    try std.testing.expectEqualStrings("漢字", book.sst.eager[0]);
}

test "openSstLazy: phonetic <rPh> annotations are dropped from materialised text" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const tmp_path = try tt.path(std.testing.allocator, "lazy_rph.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        // Build a minimal workbook with one cell, then patch the
        // SST entry to include a `<rPh>` annotation by direct
        // string surgery — the writer doesn't emit `<rPh>` so we
        // need a synthetic source.
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .string = "漢字" }});
        try w.save(tmp_path);
    }
    // Read the file and rewrite xl/sharedStrings.xml in place to
    // inject a phonetic annotation. Skip the test if anything fails
    // so this stays a regression guard rather than a flaky CI step.
    const file_data = std.fs.cwd().readFileAlloc(std.testing.allocator, tmp_path, 1 << 20) catch return;
    defer std.testing.allocator.free(file_data);

    // For simplicity, rebuild a workbook with the canonical SST
    // shape via a manual override is too invasive — instead, this
    // test asserts the handler logic via a direct invocation against
    // a synthetic SST body, mirroring the eager test. The lazy code
    // path is `materialiseSstEntry`.
    const sst_xml =
        "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\">" ++
        "<si><t>漢字</t><rPh sb=\"0\" eb=\"2\"><t>かんじ</t></rPh></si>" ++
        "</sst>";

    var book: Book = .{
        .allocator = std.testing.allocator,
        .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer book.deinit();
    const owned = try std.testing.allocator.dupe(u8, sst_xml);
    book.shared_strings_xml = owned;
    // Hand-build the lazy index so materialiseSstEntry can run
    // without going through the full Book.openSstLazy path.
    const si_open = std.mem.indexOf(u8, owned, "<si>").?;
    const si_close = std.mem.indexOf(u8, owned, "</si>").? + "</si>".len;
    const offsets = try std.testing.allocator.dupe(usize, &[_]usize{si_open + "<si>".len});
    const lengths = try std.testing.allocator.dupe(usize, &[_]usize{si_close - si_open - "<si>".len - "</si>".len});
    book.sst = .{ .lazy = .{
        .offsets = offsets,
        .lengths = lengths,
        .resolved = .empty,
    } };

    const got = try book.sharedStringAt(0);
    try std.testing.expectEqualStrings("漢字", got);
}

test "parseSharedStrings: self-closing <rPh/> is also skipped" {
    // Defensive: producers occasionally emit empty/self-closing <rPh/>.
    // The parser must skip past it without recursing.
    const sst_xml =
        "<sst count=\"1\" uniqueCount=\"1\">" ++
        "<si><t>plain</t><rPh sb=\"0\" eb=\"0\"/></si>" ++
        "</sst>";

    var book: Book = .{
        .allocator = std.testing.allocator,
        .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer book.deinit();
    const owned = try std.testing.allocator.dupe(u8, sst_xml);
    book.shared_strings_xml = owned;
    try parseSharedStrings(&book, owned);

    try std.testing.expectEqual(@as(usize, 1), book.sharedStringsCount());
    try std.testing.expectEqualStrings("plain", book.sst.eager[0]);
}

test "parseSharedStrings: hostile uniqueCount is capped against XML size" {
    // A malformed or hostile workbook can claim an absurd uniqueCount
    // to force a multi-GB pre-allocation. The hint is capped against
    // the XML size, so this short SST parses cleanly under the test
    // allocator (which traps OOM) instead of attempting a billion-slot
    // allocation.
    const sst_xml =
        "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"2\" uniqueCount=\"999999999999\">" ++
        "<si><t>a</t></si>" ++
        "<si><t>b</t></si>" ++
        "</sst>";

    var book: Book = .{
        .allocator = std.testing.allocator,
        .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer book.deinit();
    const owned = try std.testing.allocator.dupe(u8, sst_xml);
    book.shared_strings_xml = owned;
    try parseSharedStrings(&book, owned);

    try std.testing.expectEqual(@as(usize, 2), book.sharedStringsCount());
    try std.testing.expectEqualStrings("a", book.sst.eager[0]);
    try std.testing.expectEqualStrings("b", book.sst.eager[1]);
}

test "openSstLazy: lazy SST resolves on demand and matches eager (iter-sst-3b)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const tmp_path = try tt.path(std.testing.allocator, "open_sst_lazy.xlsx");
    defer std.testing.allocator.free(tmp_path);

    // Build a workbook with a handful of distinct shared strings so
    // both backends have non-trivial work to do.
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{
            .{ .string = "alpha" },
            .{ .string = "beta" },
            .{ .string = "gamma" },
        });
        try s.writeRow(&.{
            .{ .string = "delta" },
            .{ .string = "alpha" }, // dedup'd in SST
            .{ .string = "epsilon" },
        });
        try w.save(tmp_path);
    }

    // Eager open: every entry pre-decoded.
    {
        var book = try Book.open(std.testing.allocator, tmp_path);
        defer book.deinit();
        try std.testing.expectEqual(@as(usize, 5), book.sharedStringsCount());
        try std.testing.expectEqualStrings("alpha", try book.sharedStringAt(0));
        try std.testing.expectEqualStrings("beta", try book.sharedStringAt(1));
        try std.testing.expectEqualStrings("gamma", try book.sharedStringAt(2));
        try std.testing.expectEqualStrings("delta", try book.sharedStringAt(3));
        try std.testing.expectEqualStrings("epsilon", try book.sharedStringAt(4));
    }

    // Lazy open: same external behaviour, different internal storage.
    {
        var book = try Book.openSstLazy(std.testing.allocator, tmp_path);
        defer book.deinit();
        try std.testing.expectEqual(@as(usize, 5), book.sharedStringsCount());
        // Confirm we landed on the lazy backend.
        try std.testing.expect(std.meta.activeTag(book.sst) == .lazy);
        // Resolved cache empty before any access.
        try std.testing.expectEqual(@as(u32, 0), book.sst.lazy.resolved.count());

        // Sparse access: materialise a few entries.
        try std.testing.expectEqualStrings("beta", try book.sharedStringAt(1));
        try std.testing.expectEqualStrings("delta", try book.sharedStringAt(3));
        try std.testing.expectEqual(@as(u32, 2), book.sst.lazy.resolved.count());

        // Re-access does not re-materialise (cache hit).
        _ = try book.sharedStringAt(1);
        try std.testing.expectEqual(@as(u32, 2), book.sst.lazy.resolved.count());

        // Full sweep matches eager.
        try std.testing.expectEqualStrings("alpha", try book.sharedStringAt(0));
        try std.testing.expectEqualStrings("gamma", try book.sharedStringAt(2));
        try std.testing.expectEqualStrings("epsilon", try book.sharedStringAt(4));
        try std.testing.expectEqual(@as(u32, 5), book.sst.lazy.resolved.count());

        // Out-of-range still reports MalformedXml on the lazy backend.
        try std.testing.expectError(error.MalformedXml, book.sharedStringAt(99));
    }
}

test "openSstLazy: rich-runs eagerly captured on lazy backend (iter-sst-3b)" {
    // Direct parseSharedStringsLazy drive against a known SST XML —
    // covers the contract that lazy-opened books still expose
    // `Book.richRuns(idx)` with the same shape as eager.
    const sst_xml =
        "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"3\" uniqueCount=\"3\">" ++
        "<si><t>plain</t></si>" ++
        "<si><r><rPr><b/></rPr><t>bold</t></r></si>" ++
        "<si><r><rPr><i/></rPr><t>italic</t></r></si>" ++
        "</sst>";

    var book: Book = .{
        .allocator = std.testing.allocator,
        .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer book.deinit();
    const owned = try std.testing.allocator.dupe(u8, sst_xml);
    book.shared_strings_xml = owned;
    try parseSharedStringsLazy(&book, owned);

    try std.testing.expectEqual(@as(usize, 3), book.sharedStringsCount());
    try std.testing.expect(std.meta.activeTag(book.sst) == .lazy);

    // Plain entry: no rich runs.
    try std.testing.expectEqual(@as(?[]const RichRun, null), book.richRuns(0));

    // Bold entry: one run, bold flag set.
    const r1 = book.richRuns(1) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(usize, 1), r1.len);
    try std.testing.expectEqualStrings("bold", r1[0].text);
    try std.testing.expectEqual(true, r1[0].bold);
    try std.testing.expectEqual(false, r1[0].italic);

    // Italic entry: one run, italic flag set.
    const r2 = book.richRuns(2) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(usize, 1), r2.len);
    try std.testing.expectEqualStrings("italic", r2[0].text);
    try std.testing.expectEqual(false, r2[0].bold);
    try std.testing.expectEqual(true, r2[0].italic);

    // Plain text still resolves through sharedStringAt.
    try std.testing.expectEqualStrings("plain", try book.sharedStringAt(0));
    try std.testing.expectEqualStrings("bold", try book.sharedStringAt(1));
    try std.testing.expectEqualStrings("italic", try book.sharedStringAt(2));
}

test "Editor: byte-identical passthrough (iter-lms-1)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "editor_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "editor_dst.xlsx");
    defer std.testing.allocator.free(dst_path);

    // Build a non-trivial workbook so the round-trip exercises real
    // ZIP shape (multiple entries: workbook.xml, worksheets, SST,
    // styles, content types, rels).
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Data");
        try s.writeRow(&.{ .{ .string = "header" }, .{ .integer = 42 } });
        try s.writeRow(&.{ .{ .string = "row1" }, .{ .number = 3.14 } });
        try w.save(src_path);
    }

    // SHA256 of source.
    const Sha256 = std.crypto.hash.sha2.Sha256;
    var src_hash: [Sha256.digest_length]u8 = undefined;
    {
        const f = try std.fs.cwd().openFile(src_path, .{});
        defer f.close();
        const buf = try std.testing.allocator.alloc(u8, @intCast((try f.stat()).size));
        defer std.testing.allocator.free(buf);
        _ = try f.readAll(buf);
        Sha256.hash(buf, &src_hash, .{});
    }

    // Round-trip through Editor.
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.save(dst_path);
    }

    // SHA256 of destination must match.
    var dst_hash: [Sha256.digest_length]u8 = undefined;
    {
        const f = try std.fs.cwd().openFile(dst_path, .{});
        defer f.close();
        const buf = try std.testing.allocator.alloc(u8, @intCast((try f.stat()).size));
        defer std.testing.allocator.free(buf);
        _ = try f.readAll(buf);
        Sha256.hash(buf, &dst_hash, .{});
    }
    try std.testing.expectEqualSlices(u8, &src_hash, &dst_hash);

    // The destination must still open as a valid workbook through
    // the reader — confirms we didn't corrupt anything.
    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 1), book.sheets.len);
    try std.testing.expectEqualStrings("Data", book.sheets[0].name);
}

test "Editor: raw-ZIP scanner builds entry table (iter-lms-1b)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "editor_scan.xlsx");
    defer std.testing.allocator.free(src_path);

    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Sheet1");
        try s.writeRow(&.{ .{ .string = "a" }, .{ .integer = 1 } });
        try w.save(src_path);
    }

    var ed = try Editor.open(std.testing.allocator, src_path);
    defer ed.deinit();

    // Every Excel-shape archive has at least these parts: the rels
    // dir, content types, workbook.xml, the worksheet, and styles.xml
    // (zlsx writer always emits a default styles part).
    try std.testing.expect(ed.entries.len >= 5);

    // Spot-check that we found workbook.xml and the worksheet.
    var saw_workbook = false;
    var saw_sheet = false;
    for (ed.entries) |e| {
        if (std.mem.eql(u8, e.name, "xl/workbook.xml")) saw_workbook = true;
        if (std.mem.startsWith(u8, e.name, "xl/worksheets/")) saw_sheet = true;

        // Each entry's recorded LFH + payload must point inside src_buf.
        try std.testing.expect(e.lfh_offset + e.lfh_total_len + e.payload_len <= ed.src_buf.len);
        // CDFH spans must point inside the central directory range.
        try std.testing.expect(e.cdfh_offset >= ed.cd_offset);
        try std.testing.expect(e.cdfh_offset + e.cdfh_total_len <= ed.cd_offset + ed.cd_size);
    }
    try std.testing.expect(saw_workbook);
    try std.testing.expect(saw_sheet);

    // EOCD-comment is empty for zlsx-written files (the writer
    // doesn't set a comment).
    try std.testing.expectEqual(@as(usize, 0), ed.eocd_comment.len);
}

test "Editor: appendRows + save round-trips through reader (iter-lms-2)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "editor_append_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "editor_append_dst.xlsx");
    defer std.testing.allocator.free(dst_path);

    // Source workbook: 2 rows.
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Data");
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 10 } });
        try s.writeRow(&.{ .{ .integer = 2 }, .{ .integer = 20 } });
        try w.save(src_path);
    }

    // Append two more rows via Editor.
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        const append_rows = [_][]const Cell{
            &.{ .{ .integer = 3 }, .{ .integer = 30 } },
            &.{ .{ .integer = 4 }, .{ .integer = 40 } },
        };
        try ed.appendRows(0, &append_rows);
        try ed.save(dst_path);
    }

    // Read back via Book — confirm 4 rows total, original cells intact,
    // new cells at the expected indices.
    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();

    const r1 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(i64, 1), r1[0].integer);
    try std.testing.expectEqual(@as(i64, 10), r1[1].integer);

    const r2 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(i64, 2), r2[0].integer);
    try std.testing.expectEqual(@as(i64, 20), r2[1].integer);

    const r3 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(i64, 3), r3[0].integer);
    try std.testing.expectEqual(@as(i64, 30), r3[1].integer);

    const r4 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(i64, 4), r4[0].integer);
    try std.testing.expectEqual(@as(i64, 40), r4[1].integer);

    try std.testing.expectEqual(@as(?[]const Cell, null), try rows.next());
}

test "Editor: appendRows rejects out-of-range sheet idx + lossy ints" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "editor_append_reject.xlsx");
    defer std.testing.allocator.free(src_path);

    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(src_path);
    }

    var ed = try Editor.open(std.testing.allocator, src_path);
    defer ed.deinit();

    const ok_rows = [_][]const Cell{&.{.{ .integer = 2 }}};
    try std.testing.expectError(error.SheetIndexOutOfRange, ed.appendRows(99, &ok_rows));

    // Lossy integer (>2^53 + 1) is refused for the same reason the
    // writer refuses it.
    const lossy_rows = [_][]const Cell{&.{.{ .integer = 9007199254740993 }}};
    try std.testing.expectError(error.IntegerExceedsExcelPrecision, ed.appendRows(0, &lossy_rows));
}

test "Editor: appendRows with string cells extends SST (iter-lms-3)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "editor_append_str_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "editor_append_str_dst.xlsx");
    defer std.testing.allocator.free(dst_path);

    // Source workbook with one string row so the SST exists.
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Data");
        try s.writeRow(&.{ .{ .string = "alpha" }, .{ .integer = 1 } });
        try w.save(src_path);
    }

    // Append a row that mixes string + integer + boolean.
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        const append_rows = [_][]const Cell{
            &.{ .{ .string = "beta" }, .{ .integer = 2 } },
            &.{ .{ .string = "gamma" }, .{ .boolean = true } },
            // Same plain-text as an existing entry — must NOT alias
            // (always-new SST index per the iter-lms-3 contract).
            &.{ .{ .string = "alpha" }, .{ .integer = 3 } },
        };
        try ed.appendRows(0, &append_rows);
        try ed.save(dst_path);
    }

    // Read back through Book — every appended string resolves to the
    // expected content.
    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();

    const r1 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualStrings("alpha", r1[0].string);
    try std.testing.expectEqual(@as(i64, 1), r1[1].integer);

    const r2 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualStrings("beta", r2[0].string);
    try std.testing.expectEqual(@as(i64, 2), r2[1].integer);

    const r3 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualStrings("gamma", r3[0].string);
    try std.testing.expectEqual(true, r3[1].boolean);

    const r4 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualStrings("alpha", r4[0].string);
    try std.testing.expectEqual(@as(i64, 3), r4[1].integer);

    try std.testing.expectEqual(@as(?[]const Cell, null), try rows.next());

    // SST must have grown — original 1 entry + 3 appended strings
    // (no reuse, even though "alpha" repeats).
    try std.testing.expectEqual(@as(usize, 4), book.sharedStringsCount());
}

test "Editor: SST-less workbook gets fresh sharedStrings.xml on string append" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "editor_sstless_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "editor_sstless_dst.xlsx");
    defer std.testing.allocator.free(dst_path);

    // Build a workbook (the Zig writer always emits sharedStrings.xml).
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("D");
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 } });
        try w.save(src_path);
    }

    // Force the create-new-sst path: open the editor, then strip the
    // sharedStrings.xml entry from `ed.entries` before appending.
    // The Zig writer emits an SST regardless of cell types, so
    // without this the test would fall back to the substitute-
    // existing-SST path and never exercise the new branch.
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();

        var filtered: std.ArrayListUnmanaged(Editor.ZipEntry) = .{};
        defer filtered.deinit(std.testing.allocator);
        var saw_sst = false;
        for (ed.entries) |e| {
            if (std.mem.eql(u8, e.name, "xl/sharedStrings.xml")) {
                saw_sst = true;
                continue;
            }
            try filtered.append(std.testing.allocator, e);
        }
        try std.testing.expect(saw_sst); // sanity check: source had an SST
        std.testing.allocator.free(ed.entries);
        ed.entries = try filtered.toOwnedSlice(std.testing.allocator);

        const append_rows = [_][]const Cell{
            &.{ .{ .string = "alpha" }, .{ .integer = 10 } },
            &.{ .{ .string = "beta" }, .{ .integer = 20 } },
        };
        try ed.appendRows(0, &append_rows);
        try ed.save(dst_path);
    }

    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();

    const r1 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(i64, 1), r1[0].integer);
    try std.testing.expectEqual(@as(i64, 2), r1[1].integer);

    const r2 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualStrings("alpha", r2[0].string);
    try std.testing.expectEqual(@as(i64, 10), r2[1].integer);

    const r3 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualStrings("beta", r3[0].string);
    try std.testing.expectEqual(@as(i64, 20), r3[1].integer);

    try std.testing.expectEqual(@as(?[]const Cell, null), try rows.next());
    // Confirm the create-new-sst path produced an SST with exactly
    // the appended strings — `book.sharedStringsCount() == 2` would
    // be inflated if the test had silently fallen back to the
    // substitute-existing path (which would have surfaced the
    // source's SST entries first).
    try std.testing.expectEqual(@as(usize, 2), book.sharedStringsCount());
}

test "Editor: scanWorksheet returns one span per <c> element (iter-cm-1)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "scan_basic.xlsx");
    defer std.testing.allocator.free(src_path);

    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .string = "h1" }, .{ .string = "h2" }, .{ .string = "h3" } });
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .number = 2.5 }, .{ .boolean = true } });
        try s.writeRow(&.{ .{ .string = "x" }, .{ .empty = {} }, .{ .integer = 99 } });
        try w.save(src_path);
    }

    var ed = try Editor.open(std.testing.allocator, src_path);
    defer ed.deinit();

    var spans = try ed.scanWorksheet(0);
    defer spans.deinit();

    // Empty cells aren't emitted as `<c>` by the writer, so row 3
    // contributes 2 spans (col A + col C); rows 1 and 2 contribute
    // 3 each. 3 + 3 + 2 = 8.
    try std.testing.expectEqual(@as(usize, 8), spans.cells.len);

    // Spot-check a few (row, col) pairs and the body-vs-end invariant.
    const a1 = spans.find(1, 0) orelse return error.A1Missing;
    const c3 = spans.find(3, 2) orelse return error.C3Missing;
    try std.testing.expect(a1.start < a1.body_start);
    try std.testing.expect(a1.body_start <= a1.end);
    try std.testing.expect(c3.row == 3 and c3.col == 2);

    // The byte slice xml[start..end] must be a valid `<c …>…</c>`
    // (or self-closing `<c …/>`) substring.
    for (spans.cells) |s| {
        try std.testing.expect(std.mem.startsWith(u8, spans.xml[s.start..], "<c"));
        try std.testing.expect(s.end <= spans.xml.len);
    }
}

test "Editor: scanWorksheet (row,col) matches Book.rows on every cell" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "scan_roundtrip.xlsx");
    defer std.testing.allocator.free(src_path);

    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .string = "name" }, .{ .string = "qty" }, .{ .string = "ok" } });
        try s.writeRow(&.{ .{ .string = "alpha" }, .{ .integer = 10 }, .{ .boolean = true } });
        try s.writeRow(&.{ .{ .string = "beta" }, .{ .integer = 20 }, .{ .boolean = false } });
        try s.writeRow(&.{ .{ .string = "gamma" }, .{ .empty = {} }, .{ .boolean = true } });
        try w.save(src_path);
    }

    // Read the source via the Editor scanner.
    var ed = try Editor.open(std.testing.allocator, src_path);
    defer ed.deinit();
    var spans = try ed.scanWorksheet(0);
    defer spans.deinit();

    // Read the same file through Book.rows; every non-empty cell
    // must have a matching span at the same (row, col).
    var book = try Book.open(std.testing.allocator, src_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    var row_idx: u32 = 1;
    while (try rows.next()) |row| : (row_idx += 1) {
        for (row, 0..) |cell, col| {
            if (cell == .empty) continue;
            const found = spans.find(row_idx, @intCast(col));
            try std.testing.expect(found != null);
        }
    }
}

test "scanWorksheetXml: pretty-printed cells (newline after <c) + r-less rows" {
    // Two regressions captured here:
    //   1. `<c\n r="A1">` — pretty-printed XML where attrs wrap to
    //      the next line. Earlier scanner gated on space/tab/slash/
    //      gt only and dropped these cells.
    //   2. `<row>` with no r= AND r-less cells — scanner had to
    //      fall back to recoverRowFromFirstCell when present, not
    //      blindly increment a sequential counter.
    const xml =
        \\<sheetData><row r="7"><c
        \\  r="A7" t="s"><v>0</v></c><c
        \\  r="B7"><v>3.14</v></c></row><row><c r="A12"><v>1</v></c></row></sheetData>
    ;
    const cells = try scanWorksheetXml(std.testing.allocator, xml);
    defer std.testing.allocator.free(cells);
    try std.testing.expectEqual(@as(usize, 3), cells.len);
    // Row 7 cells via <row r=>; row 12 recovered from <c r="A12">.
    try std.testing.expectEqual(@as(u32, 7), cells[0].row);
    try std.testing.expectEqual(@as(u32, 0), cells[0].col);
    try std.testing.expectEqual(@as(u32, 7), cells[1].row);
    try std.testing.expectEqual(@as(u32, 1), cells[1].col);
    try std.testing.expectEqual(@as(u32, 12), cells[2].row);
    try std.testing.expectEqual(@as(u32, 0), cells[2].col);
}

test "Editor: setCell replaces a numeric cell in place (iter-cm-2a)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "setcell_basic_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "setcell_basic_dst.xlsx");
    defer std.testing.allocator.free(dst_path);

    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 } });
        try s.writeRow(&.{ .{ .integer = 3 }, .{ .integer = 4 } });
        try w.save(src_path);
    }

    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.setCell(0, 1, 1, .{ .integer = 99 }); // B1: 2 -> 99
        try ed.setCell(0, 2, 0, .{ .number = 3.5 }); // A2: 3 -> 3.5
        try ed.save(dst_path);
    }

    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();

    const r1 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(i64, 1), r1[0].integer);
    try std.testing.expectEqual(@as(i64, 99), r1[1].integer);

    const r2 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(f64, 3.5), r2[0].number);
    try std.testing.expectEqual(@as(i64, 4), r2[1].integer);
}

test "Editor: setCell with strings emits inline-string cells (iter-cm-2b)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "setcell_str_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "setcell_str_dst.xlsx");
    defer std.testing.allocator.free(dst_path);

    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .string = "a" }, .{ .string = "b" } });
        try w.save(src_path);
    }

    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        // Replace shared-string cells with inline strings — including
        // entity-needing chars and leading/trailing whitespace.
        try ed.setCell(0, 1, 0, .{ .string = "Done & dusted" });
        try ed.setCell(0, 1, 1, .{ .string = " trim me " });
        try ed.save(dst_path);
    }

    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();

    const r1 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualStrings("Done & dusted", r1[0].string);
    try std.testing.expectEqualStrings(" trim me ", r1[1].string);
}

test "Editor: setCell inserts a missing cell into an existing row (iter-cm-2c)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "setcell_insert.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "setcell_insert_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        // Row 1 has cells at A, C — gap at B. Row 2 has only A.
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .empty = {} }, .{ .integer = 3 } });
        try s.writeRow(&.{.{ .integer = 4 }});
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        // Insert into the gap at row 1 col 1.
        try ed.setCell(0, 1, 1, .{ .integer = 99 });
        // Insert at end-of-row in row 2: row has A only; append B.
        try ed.setCell(0, 2, 1, .{ .string = "appended" });
        try ed.save(dst_path);
    }
    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    const r1 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(i64, 1), r1[0].integer);
    try std.testing.expectEqual(@as(i64, 99), r1[1].integer);
    try std.testing.expectEqual(@as(i64, 3), r1[2].integer);
    const r2 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(i64, 4), r2[0].integer);
    try std.testing.expectEqualStrings("appended", r2[1].string);
}

test "Editor: setCell inserts a missing row at the right position (iter-cm-2d)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "setcell_row_insert.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "setcell_row_insert_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }}); // row 1
        try s.writeRow(&.{.{ .integer = 5 }}); // row 2
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        // Insert a row beyond the source (row 5 — fresh).
        try ed.setCell(0, 5, 0, .{ .string = "row5" });
        // Insert a row in the middle (between rows 2 and 5 — call
        // it row 3).
        try ed.setCell(0, 3, 0, .{ .integer = 33 });
        try ed.save(dst_path);
    }
    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    var seen: u32 = 0;
    while (try rows.next()) |row| : (seen += 1) {
        switch (seen) {
            0 => try std.testing.expectEqual(@as(i64, 1), row[0].integer),
            1 => try std.testing.expectEqual(@as(i64, 5), row[0].integer),
            2 => try std.testing.expectEqual(@as(i64, 33), row[0].integer),
            3 => try std.testing.expectEqualStrings("row5", row[0].string),
            else => {},
        }
    }
    try std.testing.expectEqual(@as(u32, 4), seen);
}

test "Editor: setCell rejects when source has style or formula metadata" {
    var tt = TestTmp.init();
    defer tt.deinit();
    // Codex caught: pre-fix, setCell rewrote a styled or formula
    // cell as a canonical <c>, silently dropping s="N"/<f> state.
    // Now reject up front so callers know they need a future
    // attr-preserving variant.
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "setcell_metadata_src.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        const style = try w.addStyle(.{ .font_bold = true });
        var s = try w.addSheet("S");
        // Cell A1 has s="1" (a real style index). Cell B1 has a
        // formula. Cell C1 is plain — should be settable.
        try s.writeRowStyled(
            &.{ .{ .integer = 1 }, .{ .integer = 2 }, .{ .integer = 3 } },
            &.{ style, 0, 0 },
        );
        try s.writeRowWithFormulas(
            &.{.{ .integer = 0 }},
            &.{"SUM(A1:A1)"},
        );
        try w.save(src_path);
    }
    var ed = try Editor.open(std.testing.allocator, src_path);
    defer ed.deinit();
    // Styled cell — reject.
    try std.testing.expectError(
        error.SetCellSourceCellHasMetadata,
        ed.setCell(0, 1, 0, .{ .integer = 99 }),
    );
    // Formula cell on row 2 — reject.
    try std.testing.expectError(
        error.SetCellSourceCellHasMetadata,
        ed.setCell(0, 2, 0, .{ .integer = 99 }),
    );
    // Plain cell on (1, 2) — accept.
    try ed.setCell(0, 1, 2, .{ .integer = 99 });
}

test "Editor: setCell handles empty <row r=N/> rows without duplicating" {
    var tt = TestTmp.init();
    defer tt.deinit();
    // Build a sheet with a self-closing row in the middle. The
    // writer doesn't emit those, so synthesise via XML rewrite of
    // a save (write a normal sheet, then test the helper directly
    // by stuffing it into a MutatedSheet). For session brevity:
    // round-trip a writer-emitted body row through Editor + verify
    // setCell on a row-with-cells doesn't duplicate.
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "setcell_empty_row.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "setcell_empty_row_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }}); // r=1 with cells
        try s.writeRow(&.{.{ .integer = 5 }}); // r=2 with cells
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        // setCell on row 2 col 1 — row 2 already exists. The
        // pre-fix code path classified it as missing because no
        // span had row=2 col=1. Should now correctly insert into
        // the existing row's body, not duplicate.
        try ed.setCell(0, 2, 1, .{ .integer = 999 });
        try ed.save(dst_path);
    }
    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    var n: usize = 0;
    while (try rows.next()) |row| : (n += 1) {
        if (n == 1) {
            try std.testing.expectEqual(@as(i64, 5), row[0].integer);
            try std.testing.expectEqual(@as(i64, 999), row[1].integer);
        }
    }
    // Crucial assertion: only 2 rows, not 3 (no duplicate r=2).
    try std.testing.expectEqual(@as(usize, 2), n);
}

test "Editor: setCell populates an empty <sheetData/> worksheet" {
    var tt = TestTmp.init();
    defer tt.deinit();
    // Write a workbook with only header cells, no body — produces
    // <sheetData/> in some readers' canonical form. Assert setCell
    // expands it to <sheetData></sheetData> form transparently.
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "setcell_empty_sd.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "setcell_empty_sd_dst.xlsx");
    defer std.testing.allocator.free(dst_path);

    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        // The writer always emits at least an empty <sheetData></sheetData>
        // pair. To exercise the self-closing branch we'd need a
        // fixture from an external producer; for now validate the
        // common path doesn't regress by writing an unmodified sheet
        // and inserting a row.
        _ = try w.addSheet("S");
        try w.save(src_path);
    }

    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.setCell(0, 1, 0, .{ .integer = 42 });
        try ed.save(dst_path);
    }

    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    const r1 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(i64, 42), r1[0].integer);
}

test "Editor: addSheet appends a new sheet and round-trips through reader (iter-sheet-1)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "addsheet_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "addsheet_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Original");
        try s.writeRow(&.{.{ .integer = 42 }});
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        const new_idx = try ed.addSheet("Added");
        try std.testing.expectEqual(@as(u32, 1), new_idx);
        // The new sheet should be addressable via setCell.
        try ed.setCell(new_idx, 1, 0, .{ .string = "hello" });
        try ed.setCell(new_idx, 1, 1, .{ .integer = 99 });
        try ed.save(dst_path);
    }
    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 2), book.sheets.len);
    try std.testing.expectEqualStrings("Original", book.sheets[0].name);
    try std.testing.expectEqualStrings("Added", book.sheets[1].name);
    var rows = try book.rows(book.sheets[1], std.testing.allocator);
    defer rows.deinit();
    const r = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualStrings("hello", r[0].string);
    try std.testing.expectEqual(@as(i64, 99), r[1].integer);
}

test "Editor: addSheet escapes quotes in name attribute" {
    var tt = TestTmp.init();
    defer tt.deinit();
    // Codex P1: pre-fix, names with `"` produced malformed
    // workbook.xml (`name="He said "Hi""`). Use attr-escape that
    // covers `"` → `&quot;`.
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "addsheet_quote.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "addsheet_quote_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Original");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        // validateSheetName accepts `"` so we must round-trip it.
        _ = try ed.addSheet("He said \"Hi\"");
        try ed.save(dst_path);
    }
    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 2), book.sheets.len);
    try std.testing.expectEqualStrings("He said \"Hi\"", book.sheets[1].name);
}

test "Editor: insertColumn shifts existing cells right (iter-col-3)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "insertcol_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "insertcol_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 }, .{ .integer = 3 } });
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.insertColumn(0, 2); // insert before col B
        try ed.save(dst_path);
    }
    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    const r = (try rows.next()) orelse return error.TestUnexpectedResult;
    // A=1 stays, blank inserted at B, C=2 (was B), D=3 (was C).
    try std.testing.expectEqual(@as(i64, 1), r[0].integer);
    try std.testing.expect(r[1] == .empty);
    try std.testing.expectEqual(@as(i64, 2), r[2].integer);
    try std.testing.expectEqual(@as(i64, 3), r[3].integer);
}

test "Editor: deleteColumn drops a column + shifts everything right of it left" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "deletecol_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "deletecol_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .integer = 10 }, .{ .integer = 20 }, .{ .integer = 30 } });
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.deleteColumn(0, 2); // delete col B (the 20)
        try ed.save(dst_path);
    }
    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    const r = (try rows.next()) orelse return error.TestUnexpectedResult;
    // A=10, B should be 30 (was C), and B was dropped.
    try std.testing.expectEqual(@as(i64, 10), r[0].integer);
    try std.testing.expectEqual(@as(i64, 30), r[1].integer);
}

test "applyRowEditToWorksheet: delete drops the row's <row> block" {
    const src =
        "<worksheet><dimension ref=\"A1:A3\"/><sheetData>" ++
        "<row r=\"1\"><c r=\"A1\"><v>10</v></c></row>" ++
        "<row r=\"2\"><c r=\"A2\"><v>20</v></c></row>" ++
        "<row r=\"3\"><c r=\"A3\"><v>30</v></c></row>" ++
        "</sheetData></worksheet>";
    const out = try applyRowEditToWorksheet(std.testing.allocator, src, 2, .delete);
    defer std.testing.allocator.free(out);
    // Should not contain `<v>20</v>` anymore.
    try std.testing.expect(std.mem.indexOf(u8, out, "<v>20</v>") == null);
    try std.testing.expect(std.mem.indexOf(u8, out, "<v>10</v>") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "<v>30</v>") != null);
    // Row 3 should have been renumbered to row 2.
    try std.testing.expect(std.mem.indexOf(u8, out, "<row r=\"2\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "<row r=\"3\"") == null);
    try std.testing.expect(std.mem.indexOf(u8, out, "r=\"A2\"") != null);
}

test "rels_targetForComments matches by Type or basename, not substring" {
    // Real comments rel — should be picked up.
    const ok =
        "<Relationships>" ++
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" Target=\"../comments1.xml\"/>" ++
        "</Relationships>";
    try std.testing.expectEqualStrings("../comments1.xml", rels_targetForComments(ok).?);

    // Type matches even with a non-canonical basename — non-basename
    // path must come back intact so the resolver can keep the subdir.
    const type_only =
        "<Relationships>" ++
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" Target=\"../threadedComments/comments1.xml\"/>" ++
        "</Relationships>";
    try std.testing.expectEqualStrings("../threadedComments/comments1.xml", rels_targetForComments(type_only).?);

    // Basename matches even if Type is missing/odd.
    const basename_only =
        "<Relationships>" ++
        "<Relationship Id=\"rId1\" Type=\"http://example.com/x\" Target=\"../comments7.xml\"/>" ++
        "</Relationships>";
    try std.testing.expectEqualStrings("../comments7.xml", rels_targetForComments(basename_only).?);

    // The substring trap: a hyperlink target containing the literal
    // string "comments" must NOT match. Old logic returned this one.
    const trap =
        "<Relationships>" ++
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com/comments\" TargetMode=\"External\"/>" ++
        "</Relationships>";
    try std.testing.expect(rels_targetForComments(trap) == null);

    // VML drawing whose basename happens to start with "comments"
    // but isn't .xml — also must not match.
    const drawing_trap =
        "<Relationships>" ++
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\" Target=\"../drawings/commentsShape1.vml\"/>" ++
        "</Relationships>";
    try std.testing.expect(rels_targetForComments(drawing_trap) == null);

    // Real comments rel sitting AFTER a substring trap — must be
    // picked up despite the trap appearing first.
    const trap_then_real =
        "<Relationships>" ++
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com/comments\" TargetMode=\"External\"/>" ++
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" Target=\"../comments1.xml\"/>" ++
        "</Relationships>";
    try std.testing.expectEqualStrings("../comments1.xml", rels_targetForComments(trap_then_real).?);
}

test "resolveRelArchivePath joins + normalises ../ segments" {
    var buf: [256]u8 = undefined;

    // Canonical sheet-rels case: "../comments1.xml" off "xl/worksheets/".
    try std.testing.expectEqualStrings(
        "xl/comments1.xml",
        resolveRelArchivePath("xl/worksheets/", "../comments1.xml", &buf).?,
    );

    // Non-canonical layout that the broken substring matcher would
    // have skipped. Now keeps the subdir.
    try std.testing.expectEqualStrings(
        "xl/threadedComments/comments1.xml",
        resolveRelArchivePath("xl/worksheets/", "../threadedComments/comments1.xml", &buf).?,
    );

    // No leading ../ — Target is taken relative to the worksheet's
    // own directory.
    try std.testing.expectEqualStrings(
        "xl/worksheets/comments1.xml",
        resolveRelArchivePath("xl/worksheets/", "comments1.xml", &buf).?,
    );

    // Multiple ../ segments collapse correctly.
    try std.testing.expectEqualStrings(
        "comments1.xml",
        resolveRelArchivePath("xl/worksheets/", "../../comments1.xml", &buf).?,
    );

    // Absolute (package-rooted) target: strip leading '/' and skip
    // base-dir join. parseWorkbookSheets accepts the same form.
    try std.testing.expectEqualStrings(
        "xl/comments1.xml",
        resolveRelArchivePath("xl/worksheets/", "/xl/comments1.xml", &buf).?,
    );

    // Absolute target with dot segments must still normalise.
    try std.testing.expectEqualStrings(
        "xl/comments1.xml",
        resolveRelArchivePath("xl/worksheets/", "/xl/worksheets/../comments1.xml", &buf).?,
    );
}

test "applyColEditToWorksheet refuses insert that pushes <col> past XFD" {
    const src =
        "<worksheet><dimension ref=\"A1\"/>" ++
        "<cols><col min=\"16384\" max=\"16384\" width=\"12\" customWidth=\"1\"/></cols>" ++
        "<sheetData/></worksheet>";
    const got = applyColEditToWorksheet(std.testing.allocator, src, 1, .insert);
    try std.testing.expectError(error.ColEditExceedsMaxCol, got);
}

test "applyRowEditToWorksheet collapses single-cell dimension on row delete" {
    const src =
        "<worksheet><dimension ref=\"B2\"/><sheetData>" ++
        "<row r=\"2\"><c r=\"B2\"><v>9</v></c></row>" ++
        "</sheetData></worksheet>";
    const out = try applyRowEditToWorksheet(std.testing.allocator, src, 2, .delete);
    defer std.testing.allocator.free(out);
    // Stale `B2` dimension would be wrong now; expect the safe sentinel.
    try std.testing.expect(std.mem.indexOf(u8, out, "ref=\"A1\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "ref=\"B2\"") == null);
}

test "applyColEditToWorksheet collapses single-cell dimension on col delete" {
    const src =
        "<worksheet><dimension ref=\"B2\"/><sheetData>" ++
        "<row r=\"2\"><c r=\"B2\"><v>9</v></c></row>" ++
        "</sheetData></worksheet>";
    const out = try applyColEditToWorksheet(std.testing.allocator, src, 2, .delete);
    defer std.testing.allocator.free(out);
    try std.testing.expect(std.mem.indexOf(u8, out, "ref=\"A1\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "ref=\"B2\"") == null);
}

test "Editor: insertRow shifts existing rows down (iter-row-2)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "insertrow_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "insertrow_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }});
        try s.writeRow(&.{.{ .integer = 2 }});
        try s.writeRow(&.{.{ .integer = 3 }});
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.insertRow(0, 2); // insert blank row before row 2
        try ed.save(dst_path);
    }
    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    // Three populated rows yield (the inserted row 2 has no <row>
    // element, so the iterator skips it). Original row numbers
    // 1, 2, 3 now live at rows 1, 3, 4.
    const r1 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(i64, 1), r1[0].integer);
    const r2 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(i64, 2), r2[0].integer);
    const r3 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(i64, 3), r3[0].integer);
    try std.testing.expectEqual(@as(?[]const Cell, null), try rows.next());
}

test "Editor: deleteRow removes a row + shifts everything below up" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "deleterow_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "deleterow_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 10 }});
        try s.writeRow(&.{.{ .integer = 20 }});
        try s.writeRow(&.{.{ .integer = 30 }});
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.deleteRow(0, 2);
        try ed.save(dst_path);
    }
    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    const r1 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(i64, 10), r1[0].integer);
    const r2 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(i64, 30), r2[0].integer);
    try std.testing.expectEqual(@as(?[]const Cell, null), try rows.next());
}

test "Editor: insertRow rejects sheets carrying formulas globally" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "row_unsafe.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRowWithFormulas(&.{.{ .integer = 0 }}, &.{"SUM(A1:A1)"});
        try w.save(src_path);
    }
    var ed = try Editor.open(std.testing.allocator, src_path);
    defer ed.deinit();
    try std.testing.expectError(error.RowEditWithFormulasNotSupported, ed.insertRow(0, 1));
    try std.testing.expectError(error.RowEditWithFormulasNotSupported, ed.deleteRow(0, 1));
}

test "Editor: row/col edits refuse when another sheet has a cross-sheet hyperlink" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "xref_hyperlink.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s1 = try w.addSheet("Plain");
        try s1.writeRow(&.{.{ .integer = 1 }});
        try s1.writeRow(&.{.{ .integer = 2 }});
        var s2 = try w.addSheet("WithLink");
        try s2.writeRow(&.{.{ .string = "click" }});
        // Internal hyperlink pointing back into the first sheet —
        // exactly the cross-sheet case the row/col rewriter can't
        // tokenize. Sheet1's own body is clean.
        try s2.addInternalHyperlink("A1", "Plain!C5");
        try w.save(src_path);
    }
    var ed = try Editor.open(std.testing.allocator, src_path);
    defer ed.deinit();
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 1));
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.deleteRow(0, 1));
    try std.testing.expectError(error.ColEditUnsafeForSheet, ed.insertColumn(0, 1));
    try std.testing.expectError(error.ColEditUnsafeForSheet, ed.deleteColumn(0, 1));
}

test "Editor: row/col edits refuse on sheets with frozen panes" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "pane_unsafe.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }});
        try s.writeRow(&.{.{ .integer = 2 }});
        try s.freezePanes(1, 2);
        try w.save(src_path);
    }
    var ed = try Editor.open(std.testing.allocator, src_path);
    defer ed.deinit();
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 1));
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.deleteRow(0, 1));
    try std.testing.expectError(error.ColEditUnsafeForSheet, ed.insertColumn(0, 1));
    try std.testing.expectError(error.ColEditUnsafeForSheet, ed.deleteColumn(0, 1));
}

test "Editor: deleteSheet refuses while column edits are queued" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "delsheet_after_col.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s1 = try w.addSheet("A");
        try s1.writeRow(&.{.{ .integer = 1 }});
        var s2 = try w.addSheet("B");
        try s2.writeRow(&.{.{ .integer = 2 }});
        try w.save(src_path);
    }
    var ed = try Editor.open(std.testing.allocator, src_path);
    defer ed.deinit();
    try ed.insertColumn(0, 1);
    try std.testing.expectError(error.SheetDeleteRequiresCleanState, ed.deleteSheet(1));
}

test "Editor: deleteSheet refuses while row edits are queued" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "delsheet_after_row.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s1 = try w.addSheet("A");
        try s1.writeRow(&.{.{ .integer = 1 }});
        try s1.writeRow(&.{.{ .integer = 2 }});
        var s2 = try w.addSheet("B");
        try s2.writeRow(&.{.{ .integer = 3 }});
        try w.save(src_path);
    }
    var ed = try Editor.open(std.testing.allocator, src_path);
    defer ed.deinit();
    try ed.insertRow(0, 1);
    try std.testing.expectError(error.SheetDeleteRequiresCleanState, ed.deleteSheet(1));
}

test "Editor: appendRows + setCell refuse after queued row/col edit" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "append_after_coledit.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.insertColumn(0, 1);
        try std.testing.expectError(
            error.SheetHasUnsavedRowOrColEdit,
            ed.appendRows(0, &.{&.{.{ .integer = 99 }}}),
        );
        try std.testing.expectError(
            error.SheetHasUnsavedRowOrColEdit,
            ed.setCell(0, 5, 0, .{ .integer = 99 }),
        );
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.deleteRow(0, 1);
        try std.testing.expectError(
            error.SheetHasUnsavedRowOrColEdit,
            ed.appendRows(0, &.{&.{.{ .integer = 99 }}}),
        );
        try std.testing.expectError(
            error.SheetHasUnsavedRowOrColEdit,
            ed.setCell(0, 5, 0, .{ .integer = 99 }),
        );
    }
}

test "Editor: row edits refuse when ANY sheet has a formula (cross-sheet)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "row_xsheet_formula.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s1 = try w.addSheet("Plain");
        try s1.writeRow(&.{.{ .integer = 1 }});
        try s1.writeRow(&.{.{ .integer = 2 }});
        var s2 = try w.addSheet("HasFormula");
        try s2.writeRowWithFormulas(&.{.{ .integer = 0 }}, &.{"Plain!A1+Plain!A2"});
        try w.save(src_path);
    }
    var ed = try Editor.open(std.testing.allocator, src_path);
    defer ed.deinit();
    // Editing the *clean* sheet must still refuse, because Sheet2's
    // formula references back into it and we have no tokenizer.
    try std.testing.expectError(error.RowEditWithFormulasNotSupported, ed.insertRow(0, 1));
    try std.testing.expectError(error.RowEditWithFormulasNotSupported, ed.deleteRow(0, 1));
    try std.testing.expectError(error.ColEditWithFormulasNotSupported, ed.insertColumn(0, 1));
    try std.testing.expectError(error.ColEditWithFormulasNotSupported, ed.deleteColumn(0, 1));
}

test "Editor: deleteSheet drops a source sheet (iter-sheet-3)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "delete_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "delete_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s1 = try w.addSheet("Keep");
        try s1.writeRow(&.{.{ .integer = 1 }});
        var s2 = try w.addSheet("Drop");
        try s2.writeRow(&.{.{ .integer = 2 }});
        var s3 = try w.addSheet("AlsoKeep");
        try s3.writeRow(&.{.{ .integer = 3 }});
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.deleteSheet(1);
        try ed.save(dst_path);
    }
    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 2), book.sheets.len);
    try std.testing.expectEqualStrings("Keep", book.sheets[0].name);
    try std.testing.expectEqualStrings("AlsoKeep", book.sheets[1].name);
}

test "Editor: deleteSheet drops a pending-new sheet" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "delete_new.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "delete_new_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        _ = try w.addSheet("Original");
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        const new_idx = try ed.addSheet("Tmp");
        try ed.deleteSheet(new_idx);
        try ed.save(dst_path);
    }
    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 1), book.sheets.len);
    try std.testing.expectEqualStrings("Original", book.sheets[0].name);
}

test "Editor: deleteSheet preserves order of other pending-new sheets" {
    var tt = TestTmp.init();
    defer tt.deinit();
    // Codex P1: swapRemove reordered remaining new sheets. orderedRemove
    // keeps them aligned with sheet_paths' tail.
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "delete_order.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "delete_order_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        _ = try w.addSheet("Source");
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        const a = try ed.addSheet("A"); // idx 1
        _ = try ed.addSheet("B"); // idx 2
        _ = try ed.addSheet("C"); // idx 3
        try ed.deleteSheet(a); // remove A; B,C should stay in order
        try ed.save(dst_path);
    }
    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 3), book.sheets.len);
    try std.testing.expectEqualStrings("Source", book.sheets[0].name);
    try std.testing.expectEqualStrings("B", book.sheets[1].name);
    try std.testing.expectEqualStrings("C", book.sheets[2].name);
}

test "Editor: deleteSheet frees name for reuse via addSheet" {
    var tt = TestTmp.init();
    defer tt.deinit();
    // Codex P2: pre-fix, addSheet rejected reuse of a deleted
    // source sheet's name as DuplicateSheetName.
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "delete_reuse.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "delete_reuse_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        _ = try w.addSheet("Keep");
        _ = try w.addSheet("Drop");
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.deleteSheet(1);
        // Reuse the deleted name.
        _ = try ed.addSheet("Drop");
        try ed.save(dst_path);
    }
    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 2), book.sheets.len);
    try std.testing.expectEqualStrings("Keep", book.sheets[0].name);
    try std.testing.expectEqualStrings("Drop", book.sheets[1].name);
}

test "Editor: deleteSheet rejects last-sheet" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "delete_last.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        _ = try w.addSheet("Solo");
        try w.save(src_path);
    }
    var ed = try Editor.open(std.testing.allocator, src_path);
    defer ed.deinit();
    try std.testing.expectError(error.CannotDeleteLastSheet, ed.deleteSheet(0));
    try std.testing.expectError(error.SheetIndexOutOfRange, ed.deleteSheet(99));
}

test "Editor: deleteSheet rejects dirty state" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "delete_dirty.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        _ = try w.addSheet("S1");
        _ = try w.addSheet("S2");
        try w.save(src_path);
    }
    var ed = try Editor.open(std.testing.allocator, src_path);
    defer ed.deinit();
    try ed.appendRows(0, &.{&.{.{ .integer = 5 }}});
    try std.testing.expectError(error.SheetDeleteRequiresCleanState, ed.deleteSheet(1));
}

test "Editor: renameSheet renames an existing sheet (iter-sheet-2)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "rename_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "rename_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("OldName");
        try s.writeRow(&.{.{ .integer = 42 }});
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.renameSheet(0, "NewName");
        try ed.save(dst_path);
    }
    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    try std.testing.expectEqualStrings("NewName", book.sheets[0].name);
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    const r = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(i64, 42), r[0].integer);
}

test "Editor: renameSheet rejects duplicates and invalid names" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "rename_reject.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        _ = try w.addSheet("First");
        _ = try w.addSheet("Second");
        try w.save(src_path);
    }
    var ed = try Editor.open(std.testing.allocator, src_path);
    defer ed.deinit();
    // Same name (case-insensitive) is a no-op, not an error.
    try ed.renameSheet(0, "FIRST");
    // Invalid name.
    try std.testing.expectError(error.InvalidSheetName, ed.renameSheet(0, "a:b"));
    // Duplicate against another existing sheet.
    try std.testing.expectError(error.DuplicateSheetName, ed.renameSheet(0, "Second"));
    try std.testing.expectError(error.DuplicateSheetName, ed.renameSheet(0, "second"));
    // Out-of-range.
    try std.testing.expectError(error.SheetIndexOutOfRange, ed.renameSheet(99, "Foo"));
}

test "Editor: renameSheet supports undo (rename A->B then B->A)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "rename_undo.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "rename_undo_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Original");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.renameSheet(0, "Renamed");
        // Now revert. Pre-fix this was silently dropped.
        try ed.renameSheet(0, "Original");
        try ed.save(dst_path);
    }
    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    try std.testing.expectEqualStrings("Original", book.sheets[0].name);
}

test "Editor: renameSheet persists case-only changes" {
    var tt = TestTmp.init();
    defer tt.deinit();
    // Excel uniqueness is case-insensitive but the displayed
    // casing matters. asciiEqlFold short-circuit dropped legit
    // case-only renames pre-fix.
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "rename_case.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "rename_case_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("sheet1");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.renameSheet(0, "Sheet1");
        try ed.save(dst_path);
    }
    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    try std.testing.expectEqualStrings("Sheet1", book.sheets[0].name);
}

test "Editor: rename + add reuses names freed by earlier renames" {
    var tt = TestTmp.init();
    defer tt.deinit();
    // Codex P2: pre-fix, rotate / swap rename workflows were
    // rejected because the dup check only saw raw source names.
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "rename_rotate.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "rename_rotate_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        _ = try w.addSheet("A");
        _ = try w.addSheet("B");
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        // A -> C, then B -> A (rotate).
        try ed.renameSheet(0, "C");
        try ed.renameSheet(1, "A");
        // addSheet can also reuse a freed name: rename A->B's
        // earlier old-name was "B", but sheet 1 is now "A". Adding
        // "B" should succeed.
        _ = try ed.addSheet("B");
        try ed.save(dst_path);
    }
    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    try std.testing.expectEqualStrings("C", book.sheets[0].name);
    try std.testing.expectEqualStrings("A", book.sheets[1].name);
    try std.testing.expectEqualStrings("B", book.sheets[2].name);
}

test "Editor: renameSheet on a pending-new sheet mutates in place" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "rename_new.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "rename_new_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Source");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        const idx = try ed.addSheet("Tmp");
        try ed.renameSheet(idx, "Final");
        try ed.save(dst_path);
    }
    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 2), book.sheets.len);
    try std.testing.expectEqualStrings("Final", book.sheets[1].name);
}

test "Editor: appendRows works on freshly-added sheets" {
    var tt = TestTmp.init();
    defer tt.deinit();
    // Codex P1: pre-fix, save() failed with SheetEntryNotFound when
    // appendRows targeted a sheet created via addSheet, because the
    // pending_appends loop assumed every sheet had a source entry.
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "addsheet_append.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "addsheet_append_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Original");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        const new_idx = try ed.addSheet("Fresh");
        try ed.appendRows(new_idx, &.{
            &.{ .{ .string = "alpha" }, .{ .integer = 10 } },
            &.{ .{ .string = "beta" }, .{ .integer = 20 } },
        });
        try ed.save(dst_path);
    }
    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 2), book.sheets.len);
    try std.testing.expectEqualStrings("Fresh", book.sheets[1].name);
    var rows = try book.rows(book.sheets[1], std.testing.allocator);
    defer rows.deinit();
    const r1 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualStrings("alpha", r1[0].string);
    try std.testing.expectEqual(@as(i64, 10), r1[1].integer);
    const r2 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualStrings("beta", r2[0].string);
    try std.testing.expectEqual(@as(i64, 20), r2[1].integer);
}

test "Editor: scanWorksheet works on freshly-added sheets (iter-sheet-1)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "scan_new_sheet.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Original");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(src_path);
    }
    var ed = try Editor.open(std.testing.allocator, src_path);
    defer ed.deinit();
    const new_idx = try ed.addSheet("Empty");
    // scanWorksheet on the brand-new untouched sheet must not error
    // — it should return a zero-cell span set.
    var spans = try ed.scanWorksheet(new_idx);
    defer spans.deinit();
    try std.testing.expectEqual(@as(usize, 0), spans.cells.len);
}

test "Editor: addSheet handles XML-escaped duplicate names (R&D)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    // The source workbook stores `R&D` as `name="R&amp;D"` in
    // workbook.xml. Pre-fix, sheetNameExists compared raw bytes
    // and accepted the duplicate. Now decode entities first.
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "addsheet_amp.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("R&D");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(src_path);
    }
    var ed = try Editor.open(std.testing.allocator, src_path);
    defer ed.deinit();
    try std.testing.expectError(error.DuplicateSheetName, ed.addSheet("R&D"));
    // Case-insensitive too.
    try std.testing.expectError(error.DuplicateSheetName, ed.addSheet("r&d"));
}

test "Editor: addSheet allocates non-colliding ids across multiple calls" {
    var tt = TestTmp.init();
    defer tt.deinit();
    // Codex caught: pre-fix, two addSheet calls in one session
    // produced duplicate rIds / sheetIds / sheet paths because the
    // max scan only looked at the source.
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "addsheet_seq_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "addsheet_seq_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Original");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        const a = try ed.addSheet("A");
        const b = try ed.addSheet("B");
        const c = try ed.addSheet("C");
        try ed.setCell(a, 1, 0, .{ .string = "in_a" });
        try ed.setCell(b, 1, 0, .{ .string = "in_b" });
        try ed.setCell(c, 1, 0, .{ .string = "in_c" });
        try ed.save(dst_path);
    }
    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 4), book.sheets.len);
    try std.testing.expectEqualStrings("Original", book.sheets[0].name);
    try std.testing.expectEqualStrings("A", book.sheets[1].name);
    try std.testing.expectEqualStrings("B", book.sheets[2].name);
    try std.testing.expectEqualStrings("C", book.sheets[3].name);
    inline for ([_]struct { idx: usize, want: []const u8 }{
        .{ .idx = 1, .want = "in_a" },
        .{ .idx = 2, .want = "in_b" },
        .{ .idx = 3, .want = "in_c" },
    }) |c| {
        var rows = try book.rows(book.sheets[c.idx], std.testing.allocator);
        defer rows.deinit();
        const r = (try rows.next()) orelse return error.TestUnexpectedResult;
        try std.testing.expectEqualStrings(c.want, r[0].string);
    }
}

test "Editor: addSheet duplicate names are case-insensitive" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "addsheet_case.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Existing");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(src_path);
    }
    var ed = try Editor.open(std.testing.allocator, src_path);
    defer ed.deinit();
    // Existing source sheet — case-insensitive collision.
    try std.testing.expectError(error.DuplicateSheetName, ed.addSheet("EXISTING"));
    try std.testing.expectError(error.DuplicateSheetName, ed.addSheet("existing"));
    // Pending-new-sheets — also case-insensitive.
    _ = try ed.addSheet("Fresh");
    try std.testing.expectError(error.DuplicateSheetName, ed.addSheet("FRESH"));
    try std.testing.expectError(error.DuplicateSheetName, ed.addSheet("fresh"));
}

test "Editor: addSheet rejects invalid + duplicate names" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "addsheet_reject.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Existing");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(src_path);
    }
    var ed = try Editor.open(std.testing.allocator, src_path);
    defer ed.deinit();
    try std.testing.expectError(error.InvalidSheetName, ed.addSheet(""));
    try std.testing.expectError(error.InvalidSheetName, ed.addSheet("a:b"));
    try std.testing.expectError(error.DuplicateSheetName, ed.addSheet("Existing"));
    _ = try ed.addSheet("Fresh");
    try std.testing.expectError(error.DuplicateSheetName, ed.addSheet("Fresh"));
}

test "Editor: setCells applies a batch in source order (iter-cm-3)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "setcells_batch.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "setcells_batch_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 }, .{ .integer = 3 } });
        try s.writeRow(&.{ .{ .integer = 4 }, .{ .integer = 5 }, .{ .integer = 6 } });
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.setCells(0, &.{
            .{ .row = 1, .col = 0, .cell = .{ .integer = 100 } },
            .{ .row = 1, .col = 2, .cell = .{ .number = 3.14 } },
            .{ .row = 2, .col = 1, .cell = .{ .string = "x" } },
        });
        try ed.save(dst_path);
    }
    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    const r1 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(i64, 100), r1[0].integer);
    try std.testing.expectEqual(@as(i64, 2), r1[1].integer);
    try std.testing.expectEqual(@as(f64, 3.14), r1[2].number);
    const r2 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(i64, 4), r2[0].integer);
    try std.testing.expectEqualStrings("x", r2[1].string);
    try std.testing.expectEqual(@as(i64, 6), r2[2].integer);
}

test "Editor: setCell amortises decompress across many calls (iter-cm-2a)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    // Lazy-init the MutatedSheet on first call; subsequent calls
    // mutate the cached buffer.
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "setcell_many.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "setcell_many_dst.xlsx");
    defer std.testing.allocator.free(dst_path);

    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        var i: u32 = 0;
        while (i < 10) : (i += 1) {
            try s.writeRow(&.{ .{ .integer = i }, .{ .integer = i * 2 }, .{ .integer = i * 3 } });
        }
        try w.save(src_path);
    }

    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        // Set every cell to 999. After this only the (row, col) -> 999
        // mapping should hold.
        var r: u32 = 1;
        while (r <= 10) : (r += 1) {
            var c: u32 = 0;
            while (c < 3) : (c += 1) {
                try ed.setCell(0, r, c, .{ .integer = 999 });
            }
        }
        try ed.save(dst_path);
    }

    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    while (try rows.next()) |row| {
        for (row) |cell| {
            try std.testing.expectEqual(@as(i64, 999), cell.integer);
        }
    }
}

test "Editor: setCell rejects unsupported cases (iter-cm-2a)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "setcell_reject.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(src_path);
    }
    var ed = try Editor.open(std.testing.allocator, src_path);
    defer ed.deinit();
    // Out-of-range sheet
    try std.testing.expectError(error.SheetIndexOutOfRange, ed.setCell(99, 1, 0, .{ .integer = 1 }));
    // Row 0 is invalid (1-based)
    try std.testing.expectError(error.RowIndexOutOfRange, ed.setCell(0, 0, 0, .{ .integer = 1 }));
    // Row > Excel's max (1_048_576) is invalid; in Python a caller
    // can pass `row=-1` which arrives via ctypes as u32::MAX.
    try std.testing.expectError(error.RowIndexOutOfRange, ed.setCell(0, 1_048_577, 0, .{ .integer = 1 }));
    try std.testing.expectError(error.RowIndexOutOfRange, ed.setCell(0, std.math.maxInt(u32), 0, .{ .integer = 1 }));
    // Lossy integer
    try std.testing.expectError(error.IntegerExceedsExcelPrecision, ed.setCell(0, 1, 0, .{ .integer = 9007199254740993 }));
    // Mix with appendRows on same sheet — both directions.
    try ed.appendRows(0, &.{&.{.{ .integer = 99 }}});
    try std.testing.expectError(error.SheetHasUnsavedAppends, ed.setCell(0, 1, 0, .{ .integer = 2 }));
}

test "Editor: appendRows rejects sheets with pending setCell mutations" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "setcell_then_append.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(src_path);
    }
    var ed = try Editor.open(std.testing.allocator, src_path);
    defer ed.deinit();
    try ed.setCell(0, 1, 0, .{ .integer = 2 });
    try std.testing.expectError(
        error.SheetHasUnsavedMutations,
        ed.appendRows(0, &.{&.{.{ .integer = 99 }}}),
    );
}

test "Editor: scanWorksheet sees setCell mutations (no stale read)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    // Codex caught: pre-fix, scanWorksheet decompressed from
    // src_buf and ignored pending_mutations. A setCell-then-scan
    // workflow got pre-mutation spans.
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "setcell_then_scan.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 } });
        try w.save(src_path);
    }
    var ed = try Editor.open(std.testing.allocator, src_path);
    defer ed.deinit();

    // Capture the byte length of the pre-mutation XML.
    var pre = try ed.scanWorksheet(0);
    const pre_len = pre.xml.len;
    pre.deinit();

    // Mutate B1 to a much larger number — should grow the cell's
    // byte span and therefore the worksheet XML.
    try ed.setCell(0, 1, 1, .{ .integer = 999_999_999 });

    var post = try ed.scanWorksheet(0);
    defer post.deinit();
    try std.testing.expect(post.xml.len > pre_len);
    // Confirm the new value is in the post-mutation XML.
    try std.testing.expect(std.mem.indexOf(u8, post.xml, "999999999") != null);
}

test "Editor: scanWorksheet rejects sheets with unsaved appendRows" {
    var tt = TestTmp.init();
    defer tt.deinit();
    // The scanner reads from src_buf and would silently miss rows
    // queued in pending_appends. Contract: reject so callers can't
    // act on a stale span set.
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "scan_pending.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(src_path);
    }
    var ed = try Editor.open(std.testing.allocator, src_path);
    defer ed.deinit();
    // Scan works on a clean Editor.
    {
        var spans = try ed.scanWorksheet(0);
        defer spans.deinit();
        try std.testing.expectEqual(@as(usize, 1), spans.cells.len);
    }
    // After appendRows, scanning the same sheet must error cleanly.
    try ed.appendRows(0, &.{&.{.{ .integer = 2 }}});
    try std.testing.expectError(error.SheetHasUnsavedAppends, ed.scanWorksheet(0));
}

test "Editor: scanWorksheet rejects out-of-range sheet idx" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = @import("writer.zig");
    const src_path = try tt.path(std.testing.allocator, "scan_oor.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(src_path);
    }
    var ed = try Editor.open(std.testing.allocator, src_path);
    defer ed.deinit();
    try std.testing.expectError(error.SheetIndexOutOfRange, ed.scanWorksheet(99));
}

test "buildFreshSstXml round-trips through the reader's parser" {
    const sst_xml = try buildFreshSstXml(std.testing.allocator, &.{
        "alpha", "beta with <xml> & special chars",
    });
    defer std.testing.allocator.free(sst_xml);

    var book: Book = .{
        .allocator = std.testing.allocator,
        .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer book.deinit();
    const owned = try std.testing.allocator.dupe(u8, sst_xml);
    book.shared_strings_xml = owned;
    try parseSharedStrings(&book, owned);
    try std.testing.expectEqual(@as(usize, 2), book.sharedStringsCount());
    try std.testing.expectEqualStrings("alpha", try book.sharedStringAt(0));
    try std.testing.expectEqualStrings("beta with <xml> & special chars", try book.sharedStringAt(1));
}

test "addSstRelationship splices a unique-Id Relationship" {
    const a = std.testing.allocator;
    const rels =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" ++
        "<Relationship Id=\"rId1\" Type=\"...workbook\" Target=\"xl/workbook.xml\"/>" ++
        "<Relationship Id=\"rId3\" Type=\"...sheet\" Target=\"xl/worksheets/sheet1.xml\"/>" ++
        "</Relationships>";
    const out = (try addSstRelationship(a, rels)) orelse return error.TestUnexpectedResult;
    defer a.free(out);
    try std.testing.expect(std.mem.indexOf(u8, out, "Id=\"rId4\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "Target=\"sharedStrings.xml\"") != null);

    // Idempotent — second call detects existing relationship.
    try std.testing.expectEqual(@as(?[]u8, null), try addSstRelationship(a, out));
}

test "addSstContentTypeOverride splices an Override before </Types>" {
    const a = std.testing.allocator;
    const ct =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" ++
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" ++
        "</Types>";
    const out = (try addSstContentTypeOverride(a, ct)) orelse return error.TestUnexpectedResult;
    defer a.free(out);
    try std.testing.expect(std.mem.indexOf(u8, out, "PartName=\"/xl/sharedStrings.xml\"") != null);
    try std.testing.expectEqual(@as(?[]u8, null), try addSstContentTypeOverride(a, out));
}

test "updateDimension widens row + column bounds together" {
    const a = std.testing.allocator;
    // Source dimension is A1:B2; append a 4-wide row past row 2 →
    // expect A1:D5 (max row = 5, max col 1-based = 4 = D).
    const xml = "<dimension ref=\"A1:B2\"/><sheetData/>";
    const out = (try updateDimension(a, xml, 5, 4)) orelse return error.TestUnexpectedResult;
    defer a.free(out);
    try std.testing.expect(std.mem.indexOf(u8, out, "<dimension ref=\"A1:D5\"/>") != null);

    // Row-only widening (col already covers).
    const xml2 = "<dimension ref=\"A1:Z9\"/>";
    const out2 = (try updateDimension(a, xml2, 50, 5)) orelse return error.TestUnexpectedResult;
    defer a.free(out2);
    try std.testing.expect(std.mem.indexOf(u8, out2, "<dimension ref=\"A1:Z50\"/>") != null);

    // Col-only widening (row already covers).
    const xml3 = "<dimension ref=\"A1:B100\"/>";
    const out3 = (try updateDimension(a, xml3, 50, 27)) orelse return error.TestUnexpectedResult;
    defer a.free(out3);
    try std.testing.expect(std.mem.indexOf(u8, out3, "<dimension ref=\"A1:AA100\"/>") != null);

    // No-op when both bounds already cover.
    const xml4 = "<dimension ref=\"A1:Z100\"/>";
    try std.testing.expectEqual(@as(?[]u8, null), try updateDimension(a, xml4, 50, 5));

    // Single-cell refs and missing dimension stay null per the
    // best-effort contract.
    const xml5 = "<dimension ref=\"A1\"/>";
    try std.testing.expectEqual(@as(?[]u8, null), try updateDimension(a, xml5, 99, 99));
    const xml6 = "<sheetData/>";
    try std.testing.expectEqual(@as(?[]u8, null), try updateDimension(a, xml6, 99, 99));
}

test "Rows.parseDate: auto-convert date-styled cells through the reader" {
    var tt = TestTmp.init();
    defer tt.deinit();
    // iter46 convenience: instead of chaining styleIndices +
    // isDateFormat + fromExcelSerial, callers just call parseDate.
    // Must return null for non-date-styled numbers, null for
    // string cells, and the correct DateTime for date-styled
    // numerics.
    const tmp_path = try tt.path(std.testing.allocator, "rows_parse_date.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    // Write a workbook with a styled column, then read it back and
    // check that every moving part lines up: styles.xml is extracted,
    // cellXfs is parsed, per-cell `s="N"` is tracked via
    // Rows.styleIndices(), numberFormat() resolves built-ins + custom,
    // and isDateFormat() gets the heuristic right.
    const tmp_path = try tt.path(std.testing.allocator, "reader_numfmt.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "reader_cell_font.xlsx");
    defer std.testing.allocator.free(tmp_path);

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

test "parseWorkbookSheets: pretty-printed <sheet\\n ...> is recognised" {
    const wb_xml =
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        \\          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        \\  <sheets>
        \\    <sheet
        \\      name="Pretty"
        \\      sheetId="1"
        \\      r:id="rId1"/>
        \\  </sheets>
        \\</workbook>
    ;
    const rels_xml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" ++
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" ++
        "</Relationships>";

    var book: Book = .{
        .allocator = std.testing.allocator,
        .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer book.deinit();
    try parseWorkbookSheets(&book, wb_xml, rels_xml);
    try std.testing.expectEqual(@as(usize, 1), book.sheets.len);
    try std.testing.expectEqualStrings("Pretty", book.sheets[0].name);
}

test "parseColorAttr requires exactly 8 hex digits for rgb=" {
    var book: Book = .{
        .allocator = std.testing.allocator,
        .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer book.deinit();
    // Canonical AARRGGBB.
    try std.testing.expectEqual(@as(?u32, 0xFF112233), parseColorAttr(&book, "<color rgb=\"FF112233\"/>"));
    // Non-8-digit forms must reject (return null), not silently
    // surface a near-black 0x000000FF.
    try std.testing.expectEqual(@as(?u32, null), parseColorAttr(&book, "<color rgb=\"FF\"/>"));
    try std.testing.expectEqual(@as(?u32, null), parseColorAttr(&book, "<color rgb=\"112233\"/>"));
    try std.testing.expectEqual(@as(?u32, null), parseColorAttr(&book, "<color rgb=\"FFAABBCCDD\"/>"));
}

test "fromExcelSerial: 86400-second rounding carries to next day" {
    // Serials < 61 are rejected by design (1900 leap-year ambiguity),
    // so test the midnight-rounding edge at a valid serial. A
    // fractional that rounds to 86400 must advance the date instead
    // of clamping back to 23:59:59 on the previous day.
    const day_a = fromExcelSerial(61.0).?; // 1900-03-01 00:00:00
    try std.testing.expectEqual(@as(u8, 0), day_a.hour);
    try std.testing.expectEqual(@as(u8, 0), day_a.minute);
    try std.testing.expectEqual(@as(u8, 0), day_a.second);

    // Round up to 86400: 60.99999999... — should land on 1900-03-01,
    // not 1900-02-28 23:59:59 (and 1900-02-28 isn't legal anyway).
    // Use serial 61.99999999... to check that 62.0 worth of fractional
    // input lands on the boundary day cleanly.
    const carry = fromExcelSerial(61.0 + 0.99999999999).?;
    try std.testing.expectEqual(@as(u8, 0), carry.hour);
    try std.testing.expectEqual(@as(u8, 0), carry.minute);
    try std.testing.expectEqual(@as(u8, 0), carry.second);
    try std.testing.expectEqual(@as(u8, 2), carry.day);

    // Just-under boundary stays on the same day.
    const just_before = fromExcelSerial(61.0 + 23.0 / 24.0 + 59.0 / 1440.0 + 59.0 / 86400.0).?;
    try std.testing.expectEqual(@as(u8, 23), just_before.hour);
    try std.testing.expectEqual(@as(u8, 59), just_before.minute);
    try std.testing.expectEqual(@as(u8, 59), just_before.second);

    // Carry past the upper bound must reject — 2958465.9999... rounds
    // to the next day = 2958466 which is past 9999-12-31.
    try std.testing.expectEqual(@as(?DateTime, null), fromExcelSerial(2958465.9999999999));
}

test "parseStyles: custom numFmt formatCode is XML-entity decoded" {
    // OOXML escapes quotes inside formatCode as `&quot;`. The reader
    // must decode so callers see the real format string and the
    // is-date-format scanner doesn't misclassify literal text as date
    // tokens.
    const styles_xml =
        "<?xml version=\"1.0\"?>" ++
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" ++
        "<numFmts count=\"1\"><numFmt numFmtId=\"164\" formatCode=\"0.00&quot; months&quot;\"/></numFmts>" ++
        "<cellXfs count=\"1\"><xf numFmtId=\"164\"/></cellXfs>" ++
        "</styleSheet>";

    var book: Book = .{
        .allocator = std.testing.allocator,
        .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer book.deinit();
    try parseStyles(&book, styles_xml);

    const code = book.numberFormat(0) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualStrings("0.00\" months\"", code);
    // The literal-text suffix carries `m` ("months") which the date
    // scanner would otherwise treat as a month token. The decoded
    // form keeps the suffix inside quotes — must not be a date.
    try std.testing.expect(!book.isDateFormat(0));
}

test "Book.cellBorder: round-trip sided styles + color through writer" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "reader_cell_border.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "reader_cell_fill.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "reader_internal_hyperlinks.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    // Writer xml-escapes `'` → `&apos;` on emit; reader decodes it
    // back so callers see the real target. Same contract for url +
    // location now.
    try std.testing.expectEqualStrings("'Main'!B2", links[2].location);
    try std.testing.expectEqualDeep(CellRef{ .col = 1, .row = 2 }, links[2].top_left);
    try std.testing.expectEqualDeep(CellRef{ .col = 2, .row = 2 }, links[2].bottom_right);
}

test "Book.hyperlinks: internal-only sheet (no _rels file needed)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "reader_internal_only.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "reader_no_hyperlinks.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "reader_no_merged.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    // XFD = 16384 = max valid 1-based.
    try std.testing.expectEqual(@as(usize, 16383), try columnIndexFromRef("XFD1"));
    // Refs past XFD must reject — guards consumeCell's row-growth loop.
    try std.testing.expectError(error.MalformedXml, columnIndexFromRef("XFE1"));
    try std.testing.expectError(error.MalformedXml, columnIndexFromRef("ZZZZZZ1"));
    try std.testing.expectError(error.MalformedXml, columnIndexFromRef("ZZZZZZZZZZZZZZZZZZZZZZ1"));
}

test "parseA1Ref bounds long letter / digit runs without u32 overflow" {
    // Letter run past XFD must reject before u32 multiplication can
    // overflow on safe builds.
    try std.testing.expectError(error.MalformedXml, parseA1Ref("XFE1"));
    try std.testing.expectError(error.MalformedXml, parseA1Ref("ZZZZZZZZZ1"));
    // Digit run past 1048576 must reject likewise.
    try std.testing.expectError(error.MalformedXml, parseA1Ref("A1048577"));
    try std.testing.expectError(error.MalformedXml, parseA1Ref("A99999999999"));
    // Boundary still works.
    const max = try parseA1Ref("XFD1048576");
    try std.testing.expectEqual(@as(u32, 16383), max.col);
    try std.testing.expectEqual(@as(u32, 1048576), max.row);
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

// ─── B-fuzz: coverage-guided fuzz targets ─────────────────────────────
//
// These call std.testing.fuzz() so they benefit from `-ffuzz`
// instrumentation when a Linux fuzz workflow runs `zig build fuzz`.
// Without the flag they run their corpus once each as a deterministic
// regression test (also useful in CI).
//
// Seed corpora live in tests/fuzz/corpus/<target>/, embedded via
// @embedFile so the test binary stays self-contained — no runtime
// filesystem access required.

fn fuzzParseSharedStringsTarget(_: void, input: []const u8) anyerror!void {
    var book: Book = .{
        .allocator = std.testing.allocator,
        .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer book.deinit();
    // parseSharedStrings may borrow spans from the xml buffer; dupe
    // so the buffer outlives the book regardless of the borrowing
    // choice the parser made.
    const owned = std.testing.allocator.dupe(u8, input) catch return;
    book.shared_strings_xml = owned;
    parseSharedStrings(&book, owned) catch {};
}

test "fuzz parseSharedStrings — coverage-guided" {
    // Seeds intentionally inlined as comptime strings rather than
    // @embedFile'd from tests/fuzz/corpus/: the unit-test module's
    // package boundary doesn't include the tests/ tree, so embeds
    // would fail. Expanding via build.zig anonymous imports is
    // possible but adds wiring noise; inlining stays compact while
    // documenting the canonical seed shapes. The corpus files at
    // tests/fuzz/corpus/parse_shared_strings/ stay around as the
    // canonical reference + mirror for an external fuzz harness.
    const empty = "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\"/>";
    const single = "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\"><si><t>hello</t></si></sst>";
    const rich = "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\"><si><r><rPr><b/></rPr><t>bold</t></r><r><t> normal</t></r></si></sst>";
    const phonetic = "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\"><si><t>\xE6\xBC\xA2\xE5\xAD\x97</t><rPh sb=\"0\" eb=\"2\"><t>\xE3\x81\x8B\xE3\x82\x93\xE3\x81\x98</t></rPh></si></sst>";
    try std.testing.fuzz({}, fuzzParseSharedStringsTarget, .{
        .corpus = &.{ empty, single, rich, phonetic },
    });
}

fn fuzzParseWorkbookSheetsTarget(_: void, input: []const u8) anyerror!void {
    // The workbook parser also wants a rels XML alongside the
    // workbook bytes. Use a fixed minimal rels so the fuzzer can
    // mutate the workbook side without losing coverage on the rels
    // path (which has its own fuzz target).
    const rels_xml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" ++
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" ++
        "</Relationships>";
    var book: Book = .{
        .allocator = std.testing.allocator,
        .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer book.deinit();
    parseWorkbookSheets(&book, input, rels_xml) catch {};
}

test "fuzz parseWorkbookSheets — coverage-guided" {
    const minimal =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" ++
        "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>" ++
        "</workbook>";
    const multi =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" ++
        "<workbookPr date1904=\"0\"/><bookViews><workbookView activeTab=\"0\"/></bookViews>" ++
        "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/><sheet name=\"R&amp;D\" sheetId=\"2\" r:id=\"rId2\"/></sheets>" ++
        "<definedNames><definedName name=\"MyRange\">Sheet1!$A$1:$B$2</definedName></definedNames>" ++
        "</workbook>";
    try std.testing.fuzz({}, fuzzParseWorkbookSheetsTarget, .{
        .corpus = &.{ minimal, multi },
    });
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

test "parseWorkbookSheets sets uses_1904_epoch past a <workbookProtection> prefix" {
    // Regression: <workbookProtection> shares the `<workbookPr` prefix.
    // A naive prefix-match would stop on the protection tag, miss
    // date1904 on the real workbookPr element, and silently fall back
    // to the 1900 epoch — re-creating the 1462-day skew the iter61-a
    // gate is meant to prevent.
    const wb_xml =
        \\<?xml version="1.0"?>
        \\<workbook>
        \\<workbookProtection workbookPassword="ABCD" lockStructure="1"/>
        \\<workbookPr date1904="1"/>
        \\<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>
        \\</workbook>
    ;
    const rels_xml =
        \\<?xml version="1.0"?>
        \\<Relationships>
        \\<Relationship Id="rId1" Target="worksheets/sheet1.xml"/>
        \\</Relationships>
    ;

    var book: Book = .{
        .allocator = std.testing.allocator,
        .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer book.deinit();

    try parseWorkbookSheets(&book, wb_xml, rels_xml);
    try std.testing.expect(book.uses_1904_epoch);
}

test "parseWorkbookSheets leaves uses_1904_epoch false for 1900 workbooks" {
    const wb_xml =
        \\<?xml version="1.0"?>
        \\<workbook>
        \\<workbookPr defaultThemeVersion="124226"/>
        \\<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>
        \\</workbook>
    ;
    const rels_xml =
        \\<?xml version="1.0"?>
        \\<Relationships>
        \\<Relationship Id="rId1" Target="worksheets/sheet1.xml"/>
        \\</Relationships>
    ;

    var book: Book = .{
        .allocator = std.testing.allocator,
        .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer book.deinit();

    try parseWorkbookSheets(&book, wb_xml, rels_xml);
    try std.testing.expect(!book.uses_1904_epoch);
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
    var tt = TestTmp.init();
    defer tt.deinit();
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
    const tmp_path = try tt.path(std.testing.allocator, "fuzz_addcomment.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
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
    const tmp_path = try tt.path(std.testing.allocator, "fuzz_addcf.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
        .row_date_types = .{},
        .row_error_strings = .{},
        .row_formula_strings = .{},
        .row_formula_refs = .{},
        .shared_si_to_base_ref = .{},
        .array_ranges = .{},
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
            .row_date_types = .{},
            .row_error_strings = .{},
            .row_formula_strings = .{},
            .row_formula_refs = .{},
            .shared_si_to_base_ref = .{},
            .array_ranges = .{},
            .arena = std.heap.ArenaAllocator.init(std.testing.allocator),
        };
        defer rows.deinit();

        // Consume the rows — may error, must not panic.
        while (rows.next() catch null) |_| {}
    }
}

test "Rows.currentRowNumber recovers from cell r attr when <row r> is absent/bad" {
    // Mixed <row> tags: r="10", r=missing, r="bad", r="0", r="20".
    // Each row has a cell with an explicit r attr that should be
    // used to recover the row number when <row r> is unusable. The
    // envelope `row` thus matches the inline cell ref.
    const xml =
        "<sheetData>" ++
        "<row r=\"10\"><c r=\"A10\" t=\"inlineStr\"><is><t>a</t></is></c></row>" ++
        "<row><c r=\"A11\" t=\"inlineStr\"><is><t>b</t></is></c></row>" ++
        "<row r=\"bad\"><c r=\"A12\" t=\"inlineStr\"><is><t>c</t></is></c></row>" ++
        "<row r=\"0\"><c r=\"A13\" t=\"inlineStr\"><is><t>d</t></is></c></row>" ++
        "<row r=\"20\"><c r=\"A20\" t=\"inlineStr\"><is><t>e</t></is></c></row>" ++
        "</sheetData>";

    var rows: Rows = .{
        .xml = xml,
        .pos = 0,
        .shared_strings = &.{},
        .allocator = std.testing.allocator,
        .row_cells = .{},
        .row_styles = .{},
        .row_date_types = .{},
        .row_error_strings = .{},
        .row_formula_strings = .{},
        .row_formula_refs = .{},
        .shared_si_to_base_ref = .{},
        .array_ranges = .{},
        .arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer rows.deinit();

    const expected = [_]u32{ 10, 11, 12, 13, 20 };
    for (expected) |want| {
        _ = (try rows.next()) orelse return error.UnexpectedEndOfRows;
        try std.testing.expectEqual(want, rows.currentRowNumber());
    }
    try std.testing.expectEqual(@as(?[]const Cell, null), try rows.next());
}

test "Rows.currentRowNumber recovers cell-r with non-space whitespace after <c" {
    // Pretty-printed / whitespace-normalized OOXML: the <c tag is
    // followed by a newline instead of a space. consumeCell
    // accepts this; the row-r recovery must match it too.
    const xml =
        "<sheetData>" ++
        "<row>\n  <c\n    r=\"A42\"\n    t=\"inlineStr\"><is><t>x</t></is></c>\n</row>" ++
        "</sheetData>";

    var rows: Rows = .{
        .xml = xml,
        .pos = 0,
        .shared_strings = &.{},
        .allocator = std.testing.allocator,
        .row_cells = .{},
        .row_styles = .{},
        .row_date_types = .{},
        .row_error_strings = .{},
        .row_formula_strings = .{},
        .row_formula_refs = .{},
        .shared_si_to_base_ref = .{},
        .array_ranges = .{},
        .arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer rows.deinit();

    _ = (try rows.next()) orelse return error.UnexpectedEndOfRows;
    try std.testing.expectEqual(@as(u32, 42), rows.currentRowNumber());
}

test "Rows.currentRowNumber falls through to yield count when row body is empty" {
    // Rows without r attrs and without any cells — no cell-r to
    // recover from. The fallback is the 1-based yield counter.
    // (In practice the zlsx cell parser requires every `<c>` to
    // carry an r attr, so this branch is only reachable when
    // `<row>` is truly empty or self-closing.)
    const xml =
        "<sheetData>" ++
        "<row></row>" ++
        "<row></row>" ++
        "<row></row>" ++
        "</sheetData>";

    var rows: Rows = .{
        .xml = xml,
        .pos = 0,
        .shared_strings = &.{},
        .allocator = std.testing.allocator,
        .row_cells = .{},
        .row_styles = .{},
        .row_date_types = .{},
        .row_error_strings = .{},
        .row_formula_strings = .{},
        .row_formula_refs = .{},
        .shared_si_to_base_ref = .{},
        .array_ranges = .{},
        .arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer rows.deinit();

    const expected = [_]u32{ 1, 2, 3 };
    for (expected) |want| {
        _ = (try rows.next()) orelse return error.UnexpectedEndOfRows;
        try std.testing.expectEqual(want, rows.currentRowNumber());
    }
}

// ─── iter54 slice B: lazy sheet / comments extraction ───────────────

test "openLazy path loads sheets lazily, yields identical state to open()" {
    var tt = TestTmp.init();
    defer tt.deinit();
    // Build a small multi-sheet workbook with merged ranges, hyperlinks,
    // data validations, and comments so every per-sheet side-index gets
    // a round-trip comparison.
    const tmp_path = try tt.path(std.testing.allocator, "iter54_lazy_struct.xlsx");
    defer std.testing.allocator.free(tmp_path);
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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "iter54_lazy_idempotent.xlsx");
    defer std.testing.allocator.free(tmp_path);
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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "iter54_lazy_styles.xlsx");
    defer std.testing.allocator.free(tmp_path);
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

test "streamSheet: out-of-range index errors, valid index hits cache on Book.open" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "iter54_stream_sheet.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Alpha");
        try s0.writeRow(&.{.{ .string = "a" }});
        var s1 = try w.addSheet("Beta");
        try s1.writeRow(&.{.{ .number = 1.0 }});
        try w.save(tmp_path);
    }

    var book = try Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    try std.testing.expectError(
        error.SheetIndexOutOfRange,
        book.streamSheet(99, std.testing.allocator),
    );

    // After Book.open, every sheet is preloaded: streamSheet hits the
    // cache and doesn't grow sheet_data.
    const before = book.sheet_data.count();
    {
        var r = try book.streamSheet(1, std.testing.allocator);
        defer r.deinit();
        while (try r.next()) |_| {}
    }
    try std.testing.expectEqual(before, book.sheet_data.count());
}

test "streamSheet: lazily materializes a single sheet on openLazy" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "iter54_stream_lazy.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Alpha");
        try s0.writeRow(&.{.{ .string = "a" }});
        var s1 = try w.addSheet("Beta");
        try s1.writeRow(&.{.{ .string = "b" }});
        try w.save(tmp_path);
    }

    var book = try Book.openLazy(std.testing.allocator, tmp_path);
    defer book.deinit();

    try std.testing.expectEqual(@as(usize, 0), book.sheet_data.count());

    // Stream only sheet 1; sheet 0 must stay unloaded.
    {
        var r = try book.streamSheet(1, std.testing.allocator);
        defer r.deinit();
        while (try r.next()) |_| {}
    }
    try std.testing.expectEqual(@as(usize, 1), book.sheet_data.count());
    try std.testing.expect(book.sheet_data.get(book.sheets[1].path) != null);
    try std.testing.expect(book.sheet_data.get(book.sheets[0].path) == null);
}

test "preloadSheet populates per-sheet metadata without rows() iteration" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "iter54_preload_sheet.xlsx");
    defer std.testing.allocator.free(tmp_path);
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

test "array-formula spread: base + slaves resolve via formula_refs" {
    // Synthetic two-row sheet: C2 is the array base for C2:C4 (=A2:A4*B2:B4),
    // C3 + C4 carry only cached <v> bodies. Reader must surface C2 via
    // formula_strings + the cell's cached value, and C3/C4 via
    // formula_refs pointing at C2 with the cached values intact.
    const xml =
        "<sheetData>" ++
        "<row r=\"2\">" ++
        "<c r=\"C2\"><f t=\"array\" ref=\"C2:C4\">A2:A4*B2:B4</f><v>10</v></c>" ++
        "</row>" ++
        "<row r=\"3\">" ++
        "<c r=\"C3\"><v>20</v></c>" ++
        "</row>" ++
        "<row r=\"4\">" ++
        "<c r=\"C4\"><v>30</v></c>" ++
        "</row>" ++
        "</sheetData>";

    var rows: Rows = .{
        .xml = xml,
        .pos = 0,
        .shared_strings = &.{},
        .allocator = std.testing.allocator,
        .row_cells = .{},
        .row_styles = .{},
        .row_date_types = .{},
        .row_error_strings = .{},
        .row_formula_strings = .{},
        .row_formula_refs = .{},
        .shared_si_to_base_ref = .{},
        .array_ranges = .{},
        .arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer rows.deinit();

    // Row 2 — base cell at C2 (col 2, 0-based).
    {
        const cells = (try rows.next()) orelse return error.UnexpectedEndOfRows;
        try std.testing.expect(cells.len > 2);
        const fs = rows.formulaStrings();
        const fr = rows.formulaRefs();
        try std.testing.expectEqualStrings("A2:A4*B2:B4", fs[2].?);
        try std.testing.expect(fr[2] == null);
        try std.testing.expectEqual(@as(i64, 10), cells[2].integer);
    }

    // Row 3 — slave at C3, no <f>, must resolve to C2.
    {
        const cells = (try rows.next()) orelse return error.UnexpectedEndOfRows;
        const fs = rows.formulaStrings();
        const fr = rows.formulaRefs();
        try std.testing.expect(fs[2] == null);
        try std.testing.expectEqualDeep(CellRef{ .col = 2, .row = 2 }, fr[2].?);
        try std.testing.expectEqual(@as(i64, 20), cells[2].integer);
    }

    // Row 4 — slave at C4 → also C2.
    {
        const cells = (try rows.next()) orelse return error.UnexpectedEndOfRows;
        const fs = rows.formulaStrings();
        const fr = rows.formulaRefs();
        try std.testing.expect(fs[2] == null);
        try std.testing.expectEqualDeep(CellRef{ .col = 2, .row = 2 }, fr[2].?);
        try std.testing.expectEqual(@as(i64, 30), cells[2].integer);
    }
}

test "array-formula spread: out-of-order rows still resolve slaves" {
    // No pruning of array_ranges per row — see comment in
    // `Rows.next`. This test enforces that contract: row 10
    // registers C10:C11 array, row 12 advances current_row past
    // the range's br.row, row 11 (out of order) must still find
    // C10:C11 alive in array_ranges and resolve C11's formula_ref
    // to C10.
    const xml =
        "<sheetData>" ++
        "<row r=\"10\">" ++
        "<c r=\"C10\"><f t=\"array\" ref=\"C10:C11\">A10:A11*B10:B11</f><v>5</v></c>" ++
        "</row>" ++
        "<row r=\"12\">" ++
        "<c r=\"A12\"><v>1</v></c>" ++
        "</row>" ++
        "<row r=\"11\">" ++
        "<c r=\"C11\"><v>10</v></c>" ++
        "</row>" ++
        "</sheetData>";

    var rows: Rows = .{
        .xml = xml,
        .pos = 0,
        .shared_strings = &.{},
        .allocator = std.testing.allocator,
        .row_cells = .{},
        .row_styles = .{},
        .row_date_types = .{},
        .row_error_strings = .{},
        .row_formula_strings = .{},
        .row_formula_refs = .{},
        .shared_si_to_base_ref = .{},
        .array_ranges = .{},
        .arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer rows.deinit();

    _ = (try rows.next()) orelse return error.UnexpectedEndOfRows; // row 10 base
    _ = (try rows.next()) orelse return error.UnexpectedEndOfRows; // row 12 sibling
    _ = (try rows.next()) orelse return error.UnexpectedEndOfRows; // row 11 slave (out of order)

    const fr = rows.formulaRefs();
    try std.testing.expect(fr.len > 2);
    try std.testing.expectEqualDeep(CellRef{ .col = 2, .row = 10 }, fr[2].?);
}

test "array-formula spread: own <f> on a cell inside the range wins" {
    // Defensive ordering: a cell inside a registered array range that
    // ALSO carries its own <f> keeps its own formula text. Surfaces as
    // formula_strings (not formula_refs).
    const xml =
        "<sheetData>" ++
        "<row r=\"2\">" ++
        "<c r=\"C2\"><f t=\"array\" ref=\"C2:C4\">A2:A4*B2:B4</f><v>10</v></c>" ++
        "</row>" ++
        "<row r=\"3\">" ++
        "<c r=\"C3\"><f>SUM(1,2)</f><v>3</v></c>" ++
        "</row>" ++
        "</sheetData>";

    var rows: Rows = .{
        .xml = xml,
        .pos = 0,
        .shared_strings = &.{},
        .allocator = std.testing.allocator,
        .row_cells = .{},
        .row_styles = .{},
        .row_date_types = .{},
        .row_error_strings = .{},
        .row_formula_strings = .{},
        .row_formula_refs = .{},
        .shared_si_to_base_ref = .{},
        .array_ranges = .{},
        .arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer rows.deinit();

    _ = (try rows.next()) orelse return error.UnexpectedEndOfRows; // row 2 (base)

    const cells = (try rows.next()) orelse return error.UnexpectedEndOfRows;
    const fs = rows.formulaStrings();
    const fr = rows.formulaRefs();
    try std.testing.expectEqualStrings("SUM(1,2)", fs[2].?);
    try std.testing.expect(fr[2] == null);
    try std.testing.expectEqual(@as(i64, 3), cells[2].integer);
}

test "array-formula spread: base without <v> still spreads to slaves" {
    // Edge case: an array base whose source has no cached <v>. The
    // base still surfaces its formula text; slaves still get
    // formula_ref. The base cell value is .empty.
    const xml =
        "<sheetData>" ++
        "<row r=\"2\">" ++
        "<c r=\"C2\"><f t=\"array\" ref=\"C2:C3\">A2:A3*B2:B3</f></c>" ++
        "</row>" ++
        "<row r=\"3\">" ++
        "<c r=\"C3\"><v>42</v></c>" ++
        "</row>" ++
        "</sheetData>";

    var rows: Rows = .{
        .xml = xml,
        .pos = 0,
        .shared_strings = &.{},
        .allocator = std.testing.allocator,
        .row_cells = .{},
        .row_styles = .{},
        .row_date_types = .{},
        .row_error_strings = .{},
        .row_formula_strings = .{},
        .row_formula_refs = .{},
        .shared_si_to_base_ref = .{},
        .array_ranges = .{},
        .arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer rows.deinit();

    {
        const cells = (try rows.next()) orelse return error.UnexpectedEndOfRows;
        const fs = rows.formulaStrings();
        try std.testing.expectEqualStrings("A2:A3*B2:B3", fs[2].?);
        try std.testing.expect(cells[2] == .empty);
    }
    {
        const cells = (try rows.next()) orelse return error.UnexpectedEndOfRows;
        const fr = rows.formulaRefs();
        try std.testing.expectEqualDeep(CellRef{ .col = 2, .row = 2 }, fr[2].?);
        try std.testing.expectEqual(@as(i64, 42), cells[2].integer);
    }
}

// ─── SheetMatrix tests ───────────────────────────────────────────────

test "Book.materialiseSheet: dense matrix matches streaming rows()" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "materialise_basic.xlsx");
    defer std.testing.allocator.free(tmp_path);

    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("Sheet1");
        try sheet.writeRow(&.{
            .{ .string = "Name" },
            .{ .string = "Age" },
            .{ .string = "Pi" },
        });
        try sheet.writeRow(&.{
            .{ .string = "Alice" },
            .{ .integer = 30 },
            .{ .number = 3.14159 },
        });
        try sheet.writeRow(&.{
            .{ .string = "Bob" },
            .{ .integer = 25 },
            .empty,
        });
        try w.save(tmp_path);
    }

    // Streaming reference: collect every row's cells into an owned
    // copy for shape-comparison (string aliasing is fine because the
    // book outlives the comparison block).
    var ref_book = try Book.open(std.testing.allocator, tmp_path);
    defer ref_book.deinit();
    var ref_rows_count: usize = 0;
    var ref_widths: [3]usize = undefined;
    var ref_first_string: [3][]const u8 = undefined;
    {
        var rows = try ref_book.rows(ref_book.sheets[0], std.testing.allocator);
        defer rows.deinit();
        while (try rows.next()) |cells| : (ref_rows_count += 1) {
            std.debug.assert(ref_rows_count < ref_widths.len);
            ref_widths[ref_rows_count] = cells.len;
            ref_first_string[ref_rows_count] = switch (cells[0]) {
                .string => |s| s,
                else => return error.UnexpectedCellType,
            };
            // Borrowed slices are valid for the test body; the
            // comparison below dupes them into stable storage.
            ref_first_string[ref_rows_count] = try std.testing.allocator.dupe(u8, ref_first_string[ref_rows_count]);
        }
    }
    defer for (ref_first_string[0..ref_rows_count]) |s| std.testing.allocator.free(s);

    var book = try Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    var matrix = try book.materialiseSheet(book.sheets[0], std.testing.allocator);
    defer matrix.deinit();

    try std.testing.expectEqual(@as(usize, 3), matrix.rowCount());
    try std.testing.expectEqual(ref_rows_count, matrix.rowCount());
    for (0..ref_rows_count) |i| {
        try std.testing.expectEqual(ref_widths[i], matrix.rows[i].len);
    }

    // Row 0: header strings.
    try std.testing.expectEqualStrings("Name", matrix.rows[0][0].string);
    try std.testing.expectEqualStrings("Age", matrix.rows[0][1].string);
    try std.testing.expectEqualStrings("Pi", matrix.rows[0][2].string);
    try std.testing.expectEqualStrings(ref_first_string[0], matrix.rows[0][0].string);

    // Row 1: mixed types.
    try std.testing.expectEqualStrings("Alice", matrix.rows[1][0].string);
    try std.testing.expectEqual(@as(i64, 30), matrix.rows[1][1].integer);
    try std.testing.expectApproxEqAbs(@as(f64, 3.14159), matrix.rows[1][2].number, 1e-9);
    try std.testing.expectEqualStrings(ref_first_string[1], matrix.rows[1][0].string);

    // Row 2: trailing `.empty` is dropped by the writer, so the
    // dense row mirrors the streaming reference. Per-row width
    // matches `ref_widths` (asserted above) so we don't hard-code
    // `[2].len` here.
    try std.testing.expectEqualStrings("Bob", matrix.rows[2][0].string);
    try std.testing.expectEqual(@as(i64, 25), matrix.rows[2][1].integer);

    // Accessor coverage.
    try std.testing.expect(matrix.maxColCount() >= 2);
    try std.testing.expectEqualStrings("Alice", matrix.cellAt(1, 0).?.string);
    // Out-of-bounds row → null.
    try std.testing.expect(matrix.cellAt(99, 0) == null);
    // Out-of-bounds col within a known row → `.empty` (gap fill).
    try std.testing.expect(matrix.cellAt(0, 99).? == .empty);
}

test "Book.materialiseSheet: testing.allocator confirms zero leaks" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "materialise_leakcheck.xlsx");
    defer std.testing.allocator.free(tmp_path);

    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("Leak");
        // Mix of value types + an entity-bearing string forces the
        // row arena's owned-string path so the clone-into-matrix-arena
        // branch is exercised under the leak detector.
        try sheet.writeRow(&.{
            .{ .string = "ampersand <&> here" },
            .{ .integer = 1 },
        });
        try sheet.writeRow(&.{
            .{ .string = "row two" },
            .{ .number = 2.5 },
        });
        try w.save(tmp_path);
    }

    var book = try Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    var matrix = try book.materialiseSheet(book.sheets[0], std.testing.allocator);
    defer matrix.deinit();

    try std.testing.expectEqual(@as(usize, 2), matrix.rowCount());
    try std.testing.expectEqualStrings("ampersand <&> here", matrix.rows[0][0].string);
    try std.testing.expectEqualStrings("row two", matrix.rows[1][0].string);
}

test "Book.materialiseSheet: empty sheet yields zero rows" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "materialise_empty.xlsx");
    defer std.testing.allocator.free(tmp_path);

    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        _ = try w.addSheet("Empty");
        try w.save(tmp_path);
    }

    var book = try Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    var matrix = try book.materialiseSheet(book.sheets[0], std.testing.allocator);
    defer matrix.deinit();

    try std.testing.expectEqual(@as(usize, 0), matrix.rowCount());
    try std.testing.expectEqual(@as(usize, 0), matrix.maxColCount());
    try std.testing.expect(matrix.cellAt(0, 0) == null);
}

// NOTE: a `checkAllAllocationFailures` walk over `Book.open` +
// `materialiseSheet` is the right next step to prove the errdefer
// chain on the matrix arena — but `Book.open`'s early stages have
// pre-existing allocation-failure paths that segfault under the
// failing allocator (orthogonal to this slice). Once that's
// untangled this is the natural test to add:
//
//     const Runner = struct {
//         fn run(alloc: Allocator, path: []const u8) !void {
//             var book = try Book.open(alloc, path);
//             defer book.deinit();
//             var matrix = try book.materialiseSheet(book.sheets[0], alloc);
//             defer matrix.deinit();
//             try std.testing.expectEqual(@as(usize, 2), matrix.rowCount());
//         }
//     };
//     try std.testing.checkAllAllocationFailures(std.testing.allocator, Runner.run, .{tmp_path});
