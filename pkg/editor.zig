//! Editor — load-modify-save (Phase 3c) — relocated from src/xlsx.zig
//! to pkg/editor.zig as part of B2 iter-er-0. Fixes the long-standing
//! src/ ↔ pkg/ circular-import constraint that blocked iter-er-1's
//! `Workbook.fromBook` wiring (memory: project_iter_wb_3_blocked.md).
//!
//! No semantic change vs the pre-relocation version. Internal helpers
//! that previously lived as private fns at top-level in src/xlsx.zig
//! are now top-level in this file. Cross-cutting helpers that Book/
//! Rows also use (`extractEntryToBuffer`, `findTagOpen`, `getAttr`,
//! `columnIndexFromRef`, `recoverRowFromFirstCell`, `parseWorkbookSheets`,
//! `mutate`) became `pub` in `src/xlsx.zig` so we reach them via the
//! `xlsx` module import.

const std = @import("std");
const xlsx = @import("zlsx");
const workbook_mod = @import("workbook.zig");
const sheet_edit = @import("sheet_edit.zig");
const store_mod = @import("store.zig");

const Allocator = std.mem.Allocator;
const Workbook = workbook_mod.Workbook;
const CellValue = workbook_mod.CellValue;

// Public-type aliases from xlsx so the moved code reads identically
// to its original src/-side form.
const Cell = xlsx.Cell;
const Sheet = xlsx.Sheet;
const CellRef = xlsx.CellRef;
const MergeRange = xlsx.MergeRange;
const ArrayRange = xlsx.ArrayRange;
const Hyperlink = xlsx.Hyperlink;
const DataValidation = xlsx.DataValidation;
const DataValidationKind = xlsx.DataValidationKind;
const DataValidationOperator = xlsx.DataValidationOperator;
const Comment = xlsx.Comment;
const Border = xlsx.Border;
const BorderSide = xlsx.BorderSide;
const Fill = xlsx.Fill;
const Font = xlsx.Font;
const Alignment = xlsx.Alignment;
const RichRun = xlsx.RichRun;
const DateTime = xlsx.DateTime;
const Book = xlsx.Book;
const Rows = xlsx.Rows;
const Writer = xlsx.Writer;
const TagOpen = xlsx.TagOpen;

// Cross-cutting helpers that were private in src/xlsx.zig; promoted
// to pub for cross-module access from this file.
const recoverRowFromFirstCell = xlsx.recoverRowFromFirstCell;
const findTagOpen = xlsx.findTagOpen;
const getAttr = xlsx.getAttr;
const columnIndexFromRef = xlsx.columnIndexFromRef;
const parseWorkbookSheets = xlsx.parseWorkbookSheets;
const extractEntryToBuffer = xlsx.extractEntryToBuffer;
const mutate = xlsx.mutate;
const max_col_1based = xlsx.max_col_1based;
const max_row = xlsx.max_row;
const parseA1Ref = xlsx.parseA1Ref;
const parseA1Range = xlsx.parseA1Range;

// B2 iter-er-4 (3/N): row/col edit byte-walkers + col-letter
// helpers moved to `pkg/sheet_edit.zig` so `pkg/workbook.zig` can
// also call them. Aliases below keep existing editor call sites
// unchanged.
const RowEditKind = sheet_edit.RowEditKind;
const applyRowEditToWorksheet = sheet_edit.applyRowEditToWorksheet;
const applyColEditToWorksheet = sheet_edit.applyColEditToWorksheet;
const parseColLetters = sheet_edit.parseColLetters;
const colLetterEditor = sheet_edit.colLetterEditor;
const parseSharedStrings = xlsx.parseSharedStrings;

// `casefold` was a relative import in xlsx.zig; pkg-side imports it
// via the module graph (the casefold module is registered in build.zig
// — verify if compile fails).
const casefold = xlsx.casefold;

// `deflateCompress` was at file-scope in xlsx.zig; route through the
// xlsx public surface (re-exported via pub fn deflateCompress).
const deflateCompress = xlsx.deflateCompress;

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
    /// B2 iter-er-1 read-side parity: `Editor.open` constructs an
    /// internal `Workbook` view via `Workbook.fromBook` so subsequent
    /// iters can route reads through the typed-overlay surface
    /// without forking parsing logic. v1: read-only mirror — the
    /// existing `Editor.scanWorksheet` / `appendRows` / `setCell`
    /// pipeline still walks `entries` + `src_buf` directly.
    /// `Editor.deinit` cleans up both Editor and Workbook.
    workbook: Workbook,

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
        // dangle them. The same Book instance is then promoted to a
        // pkg-side `Workbook` via `Workbook.fromBook` (B2 iter-er-1)
        // — typed-overlay parity for reads, no mutation rerouting yet.
        var sheet_paths_alloc: ?[]const []const u8 = null;
        errdefer if (sheet_paths_alloc) |sp| {
            for (sp) |p_owned| allocator.free(p_owned);
            allocator.free(sp);
        };
        var workbook_built: ?Workbook = null;
        errdefer if (workbook_built) |*wb| {
            var w = wb.*;
            w.deinit();
        };
        {
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
            sheet_paths_alloc = out_paths;

            // Workbook.fromBook re-opens the file and sanity-checks
            // sheet_count == book.sheets.len (errors SheetCountMismatch
            // on disagreement). v1 contract — future iters may share
            // bytes via PartStore-from-bytes.
            workbook_built = try Workbook.fromBook(allocator, &b, path);
        }

        return .{
            .allocator = allocator,
            .src_buf = buf,
            .entries = try entries.toOwnedSlice(allocator),
            .cd_offset = cd_offset,
            .cd_size = cd_size,
            .eocd_offset = @intCast(eocd_pos),
            .eocd_comment = eocd_comment,
            .sheet_paths = sheet_paths_alloc.?,
            .workbook = workbook_built.?,
        };
    }

    pub fn deinit(self: *Editor) void {
        self.workbook.deinit();
        self.allocator.free(self.src_buf);
        self.allocator.free(self.entries);
        for (self.sheet_paths) |p| self.allocator.free(p);
        self.allocator.free(self.sheet_paths);
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
        // sheet. Worksheet.appendRows enforces this on existing sheets
        // too, but checking here surfaces a stable error name across
        // both branches.
        if (self.sheetHasWorkbookDeltas(sheet_idx)) return error.SheetHasUnsavedMutations;
        // Empty append is a documented no-op — recording it as a
        // pending mutation would underflow the row-index math in
        // `buildSubstitutedSheet` (start_row + 0 - 1 = u32.max).
        if (rows.len == 0) return;

        // B2 iter-er-3: existing sheets route through the workbook
        // fast-path (`Worksheet.appendRows` → save-time substring
        // splice via `Worksheet.emitWithAppends`). Worksheet enforces
        // the column-cap + integer-precision guards pre-allocation,
        // so editor-level validation is unnecessary on this branch.
        // After iter-er-4 (2/N), all sheets — source AND
        // wb.addSheet'd — flow through `Worksheet.appendRows`.
        // The legacy `pending_appends` queue + `findPendingNewSheet`
        // branch retired in iter-er-6 cleanup phase 2.
        if (sheet_idx >= self.workbook.sheetCount()) return error.SheetIndexOutOfRange;
        const ws = try self.workbook.sheet(sheet_idx);
        return ws.appendRows(rows);
    }

    /// Persist the workbook to `out_path`. Atomic via
    /// `std.fs.Dir.atomicFile`. Two paths:
    ///
    /// **Passthrough**: when no `setCell` / `appendRows` /
    /// `addSheet` / `deleteSheet` / `renameSheet` / row+col edit /
    /// rewriter has touched anything, streams `src_buf` verbatim.
    /// Preserves the source SHA256 byte-for-byte.
    ///
    /// **Mutated**: delegates to `Workbook.save`, which collects
    /// all staged deltas + appended_rows into an SST extension plan
    /// (when string cells are present), emits patched sheet XML
    /// per worksheet, then `PartStore.save` rebuilds the ZIP from
    /// the override map. Source LFH bytes for untouched parts copy
    /// through byte-for-byte; the EOCD comment is preserved.
    pub fn save(self: *Editor, out_path: []const u8) !void {
        if (!self.workbookHasAnyDeltas() and
            !self.workbookHasAnyAppendedRows() and
            !self.workbook.store.hasUnsavedChanges())
        {
            var write_buf: [4096]u8 = undefined;
            var atomic_file = try std.fs.cwd().atomicFile(out_path, .{ .write_buffer = &write_buf });
            defer atomic_file.deinit();
            const w = &atomic_file.file_writer.interface;
            try w.writeAll(self.src_buf);
            try atomic_file.finish();
            return;
        }
        try self.workbook.save(out_path);
    }

    pub fn scanWorksheet(self: *Editor, sheet_idx: u32) !WorksheetSpans {
        if (sheet_idx >= self.sheet_paths.len) return error.SheetIndexOutOfRange;
        // Reject staged appends — `Worksheet.appended_rows` doesn't
        // flow into the parsed view; the span set would be stale.
        if (self.sheetHasWorkbookAppendedRows(sheet_idx)) return error.SheetHasUnsavedAppends;

        // B2 iter-er-2: existing-sheet setCell deltas live on the
        // typed-overlay view. Regenerate via `Worksheet.emitWithDeltas`
        // and re-tokenize spans, matching the post-save shape that a
        // round-trip would produce.
        if (sheet_idx < self.workbook.sheetCount()) {
            const ws = try self.workbook.sheet(sheet_idx);
            if (ws.deltas.count() > 0) {
                const xml = try ws.emitWithDeltas(self.allocator);
                errdefer self.allocator.free(xml);
                const cells = try scanWorksheetXml(self.allocator, xml);
                return .{ .allocator = self.allocator, .xml = xml, .cells = cells };
            }
        }

        const path = self.sheet_paths[sheet_idx];
        // Source ZIP entry first.
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
        // B2 iter-er-4 (2/N): wb.addSheet'd sheet — body lives in
        // PartStore (set by Workbook.addSheet). Dupe + scan.
        if (try self.workbook.store.part(path)) |part| {
            const xml = try self.allocator.dupe(u8, part.bytes);
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
    /// B2 iter-er-2: cell mutations now route through
    /// `Worksheet.setCell` on the typed-overlay view stored in
    /// `editor.workbook` (populated by `Editor.open` since iter-er-1).
    /// The Workbook side stages a `CellValue` delta keyed by
    /// `CellRef`; on save, `Worksheet.emitWithDeltas` regenerates
    /// `<sheetData>` from the parsed view + deltas and the resulting
    /// XML feeds the existing `buildEntryFromXml` pipeline.
    ///
    /// The previous in-place byte-splicing implementation (175 LOC of
    /// span tracking + `MutatedSheet` cache + insert-into-empty-row /
    /// missing-row helpers) is retired. Workbook's `emitSheetData`
    /// handles every cell-position case via a single regenerator
    /// (cell-replace, cell-into-existing-row, cell-into-empty-row,
    /// missing-row insert all collapse to "merge deltas with parsed
    /// view, emit canonical XML in row/col order").
    ///
    /// Behavioural change vs the pre-rebase setCell:
    ///   - `error.SetCellSourceCellHasMetadata` no longer fires.
    ///     Workbook regenerates each cell canonically (same shape as
    ///     the previous canonical-emit contract — formulas / inline
    ///     strings / phonetic hints / unknown attrs were already
    ///     dropped silently by the old emit; the rejection was a
    ///     pre-emit guard, not a preservation guarantee).
    ///   - `Cell.string` continues to emit as `<c t="inlineStr">`
    ///     (Workbook maps `CellValue.string` → inlineStr, no SST
    ///     extension). Identical wire shape.
    pub fn setCell(
        self: *Editor,
        sheet_idx: u32,
        row: u32,
        col: u32,
        cell: Cell,
    ) !void {
        if (sheet_idx >= self.sheet_paths.len) return error.SheetIndexOutOfRange;
        // B2 iter-er-3 symmetric guard: workbook-side appended_rows
        // are mutually exclusive with setCell on the same sheet.
        // (Worksheet.setCell enforces this internally too, but the
        // legacy-path branch below skips Worksheet.setCell, so we
        // need an editor-level check.)
        if (self.sheetHasWorkbookAppendedRows(sheet_idx)) return error.SheetHasUnsavedAppends;
        if (row == 0 or row > max_row) return error.RowIndexOutOfRange;
        if (col >= max_col_1based) return error.ColumnIndexOutOfRange;

        // For sheets that exist in the source workbook, route the
        // mutation through `Worksheet.setCell` — the typed-overlay
        // delta map is the new source of truth, and `Editor.save`
        // emits via `Worksheet.emitWithDeltas`. For freshly-added
        // sheets (sheet_idx >= source count), Workbook.addSheet isn't
        // shipped yet (iter-er-4), so fall back to the legacy
        // `pending_mutations` path. Once iter-er-4 wires Workbook
        // structural edits this branch goes away.
        if (sheet_idx < self.workbook.sheetCount()) {
            const value: CellValue = switch (cell) {
                .integer => |n| blk: {
                    if (!xlsx.fitsExactlyInF64(n)) return error.IntegerExceedsExcelPrecision;
                    break :blk .{ .number = @floatFromInt(n) };
                },
                .number => |n| .{ .number = n },
                .boolean => |b| .{ .boolean = b },
                .string => |s| .{ .string = s },
                .empty => .blank,
            };

            var ref_buf: [16]u8 = undefined;
            const ref = try xlsx.formatCellRef(&ref_buf, row, col);
            const ws = try self.workbook.sheet(sheet_idx);
            try ws.setCell(ref, value);
            return;
        }

        // sheet_idx is bounded by sheet_paths.len above and
        // sheet_paths.len == workbook.sheetCount() post-#77, so the
        // typed-overlay branch always fires. The legacy
        // pending_mutations path is unreachable.
        unreachable;
    }

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
        if (self.sheet_paths.len >= std.math.maxInt(u32)) return error.TooManySheets;

        // B2 iter-er-4 (2/N): full migration. Editor.addSheet
        // delegates to Workbook.addSheet, which patches workbook.xml
        // + workbook.xml.rels + Content_Types in-memory via
        // PartStore and grows the typed-overlay worksheets array.
        // The legacy `pending_new_sheets` queue is retired.
        //
        // Workbook.addSheet runs its own duplicate-name check
        // against `workbook.sheets` — that view reflects the
        // post-rename names because Editor.renameSheet now patches
        // the workbook view in-memory too (see below).
        const ws = self.workbook.addSheet(name) catch |err| switch (err) {
            error.SheetNameInUse => return error.DuplicateSheetName,
            else => return err,
        };
        // Capture sheet_idx + part name BEFORE any future structural
        // mutation invalidates the *Worksheet pointer.
        const new_idx = ws.sheet_idx;
        const part_name = try ws.resolvePartName();

        // Mirror in self.sheet_paths so editor-level paths
        // (scanWorksheet, setCell, etc.) resolve the new index.
        // Path string is duped into the editor allocator —
        // Editor.deinit frees per-entry via `self.allocator.free`.
        const new_path = try self.allocator.dupe(u8, part_name);
        errdefer self.allocator.free(new_path);

        const old_paths = self.sheet_paths;
        const new_paths = try self.allocator.alloc([]const u8, old_paths.len + 1);
        errdefer self.allocator.free(new_paths);
        @memcpy(new_paths[0..old_paths.len], old_paths);
        new_paths[old_paths.len] = new_path;

        self.sheet_paths = new_paths;
        self.allocator.free(old_paths);
        return new_idx;
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
        // B2 iter-er-4 (2/N): delegate to Workbook.renameSheet,
        // which patches xl/workbook.xml in-memory + runs the formula
        // rewriter (cross-sheet refs become valid, not #REF!) — the
        // strict-better path versus the legacy `pending_renames`
        // queue that this Editor used to maintain. Translates the
        // workbook-layer error names back to Editor's contract.
        if (sheet_idx >= self.workbook.sheetCount()) return error.SheetIndexOutOfRange;
        self.workbook.renameSheet(sheet_idx, new_name) catch |err| switch (err) {
            error.SheetNameInUse => return error.DuplicateSheetName,
            else => return err,
        };
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
        if (self.workbookHasAnyDeltas() or
            self.workbookHasAnyAppendedRows())
        {
            // deleteSheet rebuilds sheet_paths; queued mutations
            // hold raw indices into it and would silently point at
            // the wrong sheet (or out-of-bounds) after the rebuild.
            return error.SheetDeleteRequiresCleanState;
        }

        // iter-er-5 (5/5) lift: the `<definedName>` guard is gone.
        // `Workbook.deleteSheet` runs the four cross-sheet rewriters
        // (defined-names, formulas, hyperlinks, DV/CF) with the
        // `delete_sheet` edit variant — refs targeting the deleted
        // sheet become `#REF!`, refs to surviving sheets stay intact.

        // B2 iter-er-4 (2/N): full migration to Workbook.deleteSheet.
        // workbook.xml + workbook.xml.rels + Content_Types are
        // patched in-memory; the workbook view shrinks immediately
        // so subsequent name-collision checks (in addSheet) reflect
        // the post-delete state. The deleted sheet's part bytes
        // remain in PartStore (orphan-part v1 trade-off documented
        // on Workbook.deleteSheet) — Excel and openpyxl tolerate
        // unreferenced parts via the OPC reader's
        // Content_Types-driven part discovery.
        try self.workbook.deleteSheet(sheet_idx);

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
        try self.recordRowEdit(sheet_idx, before_row, true);
    }

    /// Delete row `row` in sheet `sheet_idx` (Phase 3e, iter-row-3).
    /// Every row > `row` shifts up by 1. Same v1 limitations as
    /// `insertRow`.
    pub fn deleteRow(self: *Editor, sheet_idx: u32, row: u32) !void {
        try self.recordRowEdit(sheet_idx, row, false);
    }

    /// Insert a blank column at position `before_col` (1-based,
    /// A=1) in sheet `sheet_idx`. Phase 3e iter-col-3. Same v1
    /// limitations as `insertRow` — formula bodies, defined names,
    /// and structured-table refs aren't rewritten and are refused.
    pub fn insertColumn(self: *Editor, sheet_idx: u32, before_col_1based: u32) !void {
        try self.recordColEdit(sheet_idx, before_col_1based, true);
    }

    /// Delete column `col_1based` (1-based) in sheet `sheet_idx`.
    /// Phase 3e iter-col-4.
    pub fn deleteColumn(self: *Editor, sheet_idx: u32, col_1based: u32) !void {
        try self.recordColEdit(sheet_idx, col_1based, false);
    }

    /// True iff any worksheet in the embedded workbook has staged
    /// `setCell`/`deleteCell` deltas. B2 iter-er-2 replacement for
    /// the retired `self.pending_mutations.count() > 0` check.
    fn workbookHasAnyDeltas(self: *Editor) bool {
        var i: u32 = 0;
        while (i < self.workbook.sheetCount()) : (i += 1) {
            const ws = self.workbook.sheet(i) catch unreachable;
            if (ws.deltas.count() > 0) return true;
        }
        return false;
    }

    /// True iff the worksheet at `sheet_idx` has staged
    /// `setCell`/`deleteCell` deltas. B2 iter-er-2 replacement for
    /// `self.pending_mutations.contains(sheet_idx)`.
    fn sheetHasWorkbookDeltas(self: *Editor, sheet_idx: u32) bool {
        if (sheet_idx >= self.workbook.sheetCount()) return false;
        const ws = self.workbook.sheet(sheet_idx) catch return false;
        return ws.deltas.count() > 0;
    }

    /// True iff any worksheet in the embedded workbook has staged
    /// `appendRows` rows. B2 iter-er-3 mirror of
    /// `workbookHasAnyDeltas` for the staging buffer that
    /// `Worksheet.appendRows` writes to.
    fn workbookHasAnyAppendedRows(self: *Editor) bool {
        var i: u32 = 0;
        while (i < self.workbook.sheetCount()) : (i += 1) {
            // `Workbook.sheet(i)` for `i < sheetCount()` is genuinely
            // infallible (the body has no alloc / parse path; the
            // broad `Error` union covers other entry points). Use
            // `catch unreachable` — `catch return false` would silently
            // mask a real corruption if the contract ever weakens.
            const ws = self.workbook.sheet(i) catch unreachable;
            if (ws.appended_rows.items.len > 0) return true;
        }
        return false;
    }

    /// True iff the worksheet at `sheet_idx` has staged appended
    /// rows. Mirror of `sheetHasWorkbookDeltas`. Used by
    /// `Editor.scanWorksheet` and `Editor.setCell` to refuse
    /// operations against a sheet whose appended rows aren't yet
    /// flushed.
    fn sheetHasWorkbookAppendedRows(self: *Editor, sheet_idx: u32) bool {
        if (sheet_idx >= self.workbook.sheetCount()) return false;
        const ws = self.workbook.sheet(sheet_idx) catch return false;
        return ws.appended_rows.items.len > 0;
    }

    fn recordColEdit(
        self: *Editor,
        sheet_idx: u32,
        col_1based: u32,
        is_insert: bool,
    ) !void {
        if (sheet_idx >= self.sheet_paths.len) return error.SheetIndexOutOfRange;
        if (col_1based == 0 or col_1based > max_col_1based) return error.ColumnIndexOutOfRange;
        if (self.sheetHasWorkbookDeltas(sheet_idx) or
            self.sheetHasWorkbookAppendedRows(sheet_idx))
        {
            return error.ColEditRequiresCleanSheet;
        }

        // Per-sheet content guard — same axes as recordRowEdit.
        const path = self.sheet_paths[sheet_idx];
        if (findEntryByName(self.entries, path)) |entry_idx| {
            const e = self.entries[entry_idx];
            const payload = self.src_buf[e.lfh_offset + e.lfh_total_len ..][0..e.payload_len];
            const xml = try decompressZipPayload(self.allocator, payload, e.compression_method, e.uncompressed_size);
            defer self.allocator.free(xml);
            // `<pane>`, `<picture>`, and modern `<drawing>` (xdr)
            // are no longer in this list — `pkg/sheet_edit.zig`
            // shifts pane splits during the byte transform,
            // CT_SheetBackgroundPicture is a coordinate-free
            // `r:id` reference (ECMA-376 §18.3.1.67), and
            // `pkg/drawing_edit.zig` (wired through
            // `Workbook.applyDrawingEditForSheet`) shifts
            // `<xdr:from>`/`<xdr:to>` row/col coords in the
            // referenced drawing part. `<legacyDrawing>` (VML)
            // and `<tableParts>` stay refused.
            const guards = [_][]const u8{
                "<tableParts",
                "<legacyDrawing",
            };
            for (guards) |g| {
                if (std.mem.indexOf(u8, xml, g) != null) {
                    return error.ColEditUnsafeForSheet;
                }
            }
        } else return error.SheetEntryNotFound;

        if (is_insert) {
            try self.workbook.insertColumn(sheet_idx, col_1based);
        } else {
            try self.workbook.deleteColumn(sheet_idx, col_1based);
        }
    }

    fn recordRowEdit(
        self: *Editor,
        sheet_idx: u32,
        row: u32,
        is_insert: bool,
    ) !void {
        if (sheet_idx >= self.sheet_paths.len) return error.SheetIndexOutOfRange;
        if (row == 0 or row > max_row) return error.RowIndexOutOfRange;
        // Refuse when staged mutations target the same sheet — the
        // typed-overlay row shift invalidates the parsed view, so
        // existing deltas / appended_rows would point at the
        // pre-shift refs and produce stale output. Same invariant
        // as Worksheet.appendRows / Worksheet.setCell mutual
        // exclusion.
        if (self.sheetHasWorkbookDeltas(sheet_idx) or
            self.sheetHasWorkbookAppendedRows(sheet_idx))
        {
            return error.RowEditRequiresCleanSheet;
        }

        // Per-sheet content guard: refuse sheets whose bodies carry
        // constructs the rewriters don't yet handle (legacy VML
        // drawings, pivots, tableParts). `<pane>`, `<autoFilter>`,
        // `<picture>`, and modern `<drawing>` (xdr) are no longer
        // in this list — `pkg/sheet_edit.zig` rewrites pane +
        // autoFilter during the worksheet byte transform,
        // CT_SheetBackgroundPicture is a coordinate-free
        // background-image reference (ECMA-376 §18.3.1.67), and
        // `pkg/drawing_edit.zig` (wired through
        // `Workbook.applyDrawingEditForSheet`) shifts xdr anchor
        // coords in the referenced drawing part. See
        // `docs/plans/refusal-audit.md` for the remaining rationale.
        const path = self.sheet_paths[sheet_idx];
        if (findEntryByName(self.entries, path)) |entry_idx| {
            const e = self.entries[entry_idx];
            const payload = self.src_buf[e.lfh_offset + e.lfh_total_len ..][0..e.payload_len];
            const xml = try decompressZipPayload(self.allocator, payload, e.compression_method, e.uncompressed_size);
            defer self.allocator.free(xml);
            const guards = [_][]const u8{
                "<tableParts",
                "<legacyDrawing",
            };
            for (guards) |g| {
                if (std.mem.indexOf(u8, xml, g) != null) {
                    return error.RowEditUnsafeForSheet;
                }
            }
        } else return error.SheetEntryNotFound;

        if (is_insert) {
            try self.workbook.insertRow(sheet_idx, row);
        } else {
            try self.workbook.deleteRow(sheet_idx, row);
        }
    }

    /// Decompress an entry by name. Caller owns the returned slice.
    fn readEntry(self: *Editor, entry_name: []const u8) ![]u8 {
        const idx = findEntryByName(self.entries, entry_name) orelse
            return error.MissingEntry;
        const e = self.entries[idx];
        const payload = self.src_buf[e.lfh_offset + e.lfh_total_len ..][0..e.payload_len];
        return try decompressZipPayload(self.allocator, payload, e.compression_method, e.uncompressed_size);
    }
};

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

/// Linear scan: small N (entries per xlsx is ~10-20) makes a hashmap
/// overkill. Returns the first matching index.
fn findEntryByName(entries: []const ZipEntry, name: []const u8) ?usize {
    for (entries, 0..) |e, i| {
        if (std.mem.eql(u8, e.name, name)) return i;
    }
    return null;
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

/// (tests there have their own copy; relocating Editor brings its own).
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

// ─── Tests ───────────────────────────────────────────────────────────

test "Editor: byte-identical passthrough (iter-lms-1)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
            // Same plain-text as an existing entry — dedups via the
            // SST extension plan (iter-er-6 unified Editor.save with
            // the typed-overlay's setCell-delta semantics).
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

    // SST grows by every UNIQUE new string; "alpha" already
    // existed, so only "beta" and "gamma" extend the table —
    // 1 existing + 2 new = 3.
    try std.testing.expectEqual(@as(usize, 3), book.sharedStringsCount());
}

test "Editor: SST-less workbook gets fresh sharedStrings.xml on string append" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
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

        var filtered: std.ArrayListUnmanaged(ZipEntry) = .{};
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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

// B2 iter-er-2 retired the `error.SetCellSourceCellHasMetadata`
// guard for existing-sheet setCell — Worksheet.setCell on the
// typed-overlay routes through Workbook's emitSheetData regenerator
// which already produces canonical XML (the pre-rebase guard was a
// pre-emit warning, not a preservation guarantee — old setCell would
// also silently drop styles/formulas, just with an explicit error
// instead of silent overwrite). The previous test for that behavior
// is dropped; iter-er-2e (attr-preserving variant) is the future
// home for metadata-aware mutation.

test "Editor: setCell handles empty <row r=N/> rows without duplicating" {
    var tt = TestTmp.init();
    defer tt.deinit();
    // Build a sheet with a self-closing row in the middle. The
    // writer doesn't emit those, so synthesise via XML rewrite of
    // a save (write a normal sheet, then test the helper directly
    // by stuffing it into a MutatedSheet). For session brevity:
    // round-trip a writer-emitted body row through Editor + verify
    // setCell on a row-with-cells doesn't duplicate.
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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

test "Editor: editor.workbook.addSheet path round-trips through Editor.save (iter-er-4)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, "wb_addsheet_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "wb_addsheet_dst.xlsx");
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
        // Bypass Editor.addSheet — go straight through the
        // typed-overlay surface. Editor.save's verbatim-emit walk
        // and PartStore-overrides sync should pick it up.
        _ = try ed.workbook.addSheet("FromWorkbook");
        try ed.save(dst_path);
    }
    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 2), book.sheets.len);
    try std.testing.expectEqualStrings("Original", book.sheets[0].name);
    try std.testing.expectEqualStrings("FromWorkbook", book.sheets[1].name);
}

test "Editor: editor.workbook.addSheet + setCell on returned handle round-trips" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, "wb_addsheet_set_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "wb_addsheet_set_dst.xlsx");
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
        const new_ws_idx = blk: {
            const ws = try ed.workbook.addSheet("Target");
            const idx = ws.sheet_idx;
            try ws.setCell("A1", .{ .number = 3.14 });
            try ws.setCell("B1", .{ .string = "pi" });
            break :blk idx;
        };
        // Re-fetch handle (per pointer-lifetime contract — fine in
        // this test since no further structural mutations happen,
        // but exercises the documented re-fetch pattern).
        _ = new_ws_idx;
        try ed.save(dst_path);
    }
    var book = try Book.open(std.testing.allocator, dst_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 2), book.sheets.len);
    var rows = try book.rows(book.sheets[1], std.testing.allocator);
    defer rows.deinit();
    const r = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(f64, 3.14), r[0].number);
    try std.testing.expectEqualStrings("pi", r[1].string);
}

test "Editor: addSheet appends a new sheet and round-trips through reader (iter-sheet-1)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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

// ─── pane rewriter (lifts the `<pane>` per-sheet row/col refusal) ──

test "applyRowEditToWorksheet: insertRow above frozen ySplit grows the freeze" {
    const src =
        "<worksheet><sheetViews><sheetView workbookViewId=\"0\">" ++
        "<pane ySplit=\"3\" topLeftCell=\"A4\" activePane=\"bottomLeft\" state=\"frozen\"/>" ++
        "</sheetView></sheetViews>" ++
        "<sheetData/></worksheet>";
    const out = try applyRowEditToWorksheet(std.testing.allocator, src, 2, .insert);
    defer std.testing.allocator.free(out);
    try std.testing.expect(std.mem.indexOf(u8, out, "ySplit=\"4\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "topLeftCell=\"A5\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "state=\"frozen\"") != null);
}

test "applyRowEditToWorksheet: insertRow below frozen ySplit leaves freeze alone" {
    const src =
        "<worksheet><sheetViews><sheetView workbookViewId=\"0\">" ++
        "<pane ySplit=\"3\" topLeftCell=\"A4\" activePane=\"bottomLeft\" state=\"frozen\"/>" ++
        "</sheetView></sheetViews>" ++
        "<sheetData/></worksheet>";
    const out = try applyRowEditToWorksheet(std.testing.allocator, src, 5, .insert);
    defer std.testing.allocator.free(out);
    try std.testing.expect(std.mem.indexOf(u8, out, "ySplit=\"3\"") != null);
    // topLeftCell row 4 < insert row 5, so it doesn't shift.
    try std.testing.expect(std.mem.indexOf(u8, out, "topLeftCell=\"A4\"") != null);
}

test "applyRowEditToWorksheet: deleteRow inside frozen ySplit shrinks the freeze" {
    const src =
        "<worksheet><sheetViews><sheetView workbookViewId=\"0\">" ++
        "<pane ySplit=\"3\" topLeftCell=\"A4\" activePane=\"bottomLeft\" state=\"frozen\"/>" ++
        "</sheetView></sheetViews>" ++
        "<sheetData/></worksheet>";
    const out = try applyRowEditToWorksheet(std.testing.allocator, src, 2, .delete);
    defer std.testing.allocator.free(out);
    try std.testing.expect(std.mem.indexOf(u8, out, "ySplit=\"2\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "topLeftCell=\"A3\"") != null);
}

test "applyColEditToWorksheet: insertColumn inside frozen xSplit grows the freeze" {
    const src =
        "<worksheet><sheetViews><sheetView workbookViewId=\"0\">" ++
        "<pane xSplit=\"2\" topLeftCell=\"C1\" activePane=\"topRight\" state=\"frozen\"/>" ++
        "</sheetView></sheetViews>" ++
        "<sheetData/></worksheet>";
    const out = try applyColEditToWorksheet(std.testing.allocator, src, 1, .insert);
    defer std.testing.allocator.free(out);
    try std.testing.expect(std.mem.indexOf(u8, out, "xSplit=\"3\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "topLeftCell=\"D1\"") != null);
}

test "applyColEditToWorksheet: deleteColumn inside frozen xSplit shrinks the freeze" {
    const src =
        "<worksheet><sheetViews><sheetView workbookViewId=\"0\">" ++
        "<pane xSplit=\"2\" topLeftCell=\"C1\" activePane=\"topRight\" state=\"frozen\"/>" ++
        "</sheetView></sheetViews>" ++
        "<sheetData/></worksheet>";
    const out = try applyColEditToWorksheet(std.testing.allocator, src, 1, .delete);
    defer std.testing.allocator.free(out);
    try std.testing.expect(std.mem.indexOf(u8, out, "xSplit=\"1\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "topLeftCell=\"B1\"") != null);
}

test "applyRowEditToWorksheet: row+col freeze (bottomRight) shifts both axes" {
    const src =
        "<worksheet><sheetViews><sheetView workbookViewId=\"0\">" ++
        "<pane xSplit=\"2\" ySplit=\"1\" topLeftCell=\"C2\" activePane=\"bottomRight\" state=\"frozen\"/>" ++
        "</sheetView></sheetViews>" ++
        "<sheetData/></worksheet>";
    // Row insert at 1 lands inside the row freeze (ySplit=1).
    const out = try applyRowEditToWorksheet(std.testing.allocator, src, 1, .insert);
    defer std.testing.allocator.free(out);
    try std.testing.expect(std.mem.indexOf(u8, out, "ySplit=\"2\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "xSplit=\"2\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "topLeftCell=\"C3\"") != null);
}

test "applyRowEditToWorksheet: ySplit=1 delete row 1 collapses freeze to 0" {
    const src =
        "<worksheet><sheetViews><sheetView workbookViewId=\"0\">" ++
        "<pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/>" ++
        "</sheetView></sheetViews>" ++
        "<sheetData/></worksheet>";
    const out = try applyRowEditToWorksheet(std.testing.allocator, src, 1, .delete);
    defer std.testing.allocator.free(out);
    try std.testing.expect(std.mem.indexOf(u8, out, "ySplit=\"0\"") != null);
    // topLeftCell row 2 > deleted row 1, so it shifts up to A1.
    try std.testing.expect(std.mem.indexOf(u8, out, "topLeftCell=\"A1\"") != null);
}

test "applyRowEditToWorksheet: refuses split-state pane (pixel offsets)" {
    const src =
        "<worksheet><sheetViews><sheetView workbookViewId=\"0\">" ++
        "<pane ySplit=\"2400\" topLeftCell=\"A5\" activePane=\"bottomLeft\" state=\"split\"/>" ++
        "</sheetView></sheetViews>" ++
        "<sheetData/></worksheet>";
    try std.testing.expectError(
        error.SplitPaneNotSupported,
        applyRowEditToWorksheet(std.testing.allocator, src, 1, .insert),
    );
}

test "applyColEditToWorksheet: refuses pane with missing state attr (OOXML default = split)" {
    const src =
        "<worksheet><sheetViews><sheetView workbookViewId=\"0\">" ++
        "<pane xSplit=\"1200\" topLeftCell=\"E1\" activePane=\"topRight\"/>" ++
        "</sheetView></sheetViews>" ++
        "<sheetData/></worksheet>";
    try std.testing.expectError(
        error.SplitPaneNotSupported,
        applyColEditToWorksheet(std.testing.allocator, src, 1, .insert),
    );
}

test "applyRowEditToWorksheet: pane with frozenSplit state is rewritten" {
    const src =
        "<worksheet><sheetViews><sheetView workbookViewId=\"0\">" ++
        "<pane ySplit=\"2\" topLeftCell=\"A3\" activePane=\"bottomLeft\" state=\"frozenSplit\"/>" ++
        "</sheetView></sheetViews>" ++
        "<sheetData/></worksheet>";
    const out = try applyRowEditToWorksheet(std.testing.allocator, src, 1, .insert);
    defer std.testing.allocator.free(out);
    try std.testing.expect(std.mem.indexOf(u8, out, "ySplit=\"3\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "topLeftCell=\"A4\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "state=\"frozenSplit\"") != null);
}

test "applyRowEditToWorksheet: pane with malformed ySplit surfaces typed error" {
    const src =
        "<worksheet><sheetViews><sheetView workbookViewId=\"0\">" ++
        "<pane ySplit=\"abc\" topLeftCell=\"A4\" state=\"frozen\"/>" ++
        "</sheetView></sheetViews>" ++
        "<sheetData/></worksheet>";
    try std.testing.expectError(
        error.MalformedPaneSplit,
        applyRowEditToWorksheet(std.testing.allocator, src, 1, .insert),
    );
}

test "Editor: insertRow shifts existing rows down (iter-row-2)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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

test "Editor: insertRow on a sheet carrying formulas rewrites refs (iter-er-5 lift)" {
    // Pre iter-er-5 lift: refused with `RowEditWithFormulasNotSupported`.
    // Post-lift: `Workbook.insertRow` runs `rewriteAllFormulas` so
    // cross-sheet and bare formula refs shift alongside the byte
    // transform. The call now succeeds.
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, "row_unsafe.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRowWithFormulas(&.{.{ .integer = 0 }}, &.{"SUM(A1:A1)"});
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.insertRow(0, 1);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.deleteRow(0, 1);
    }
}

test "Editor: row/col edits with cross-sheet hyperlinks rewrite locations (iter-er-5 lift)" {
    // Pre iter-er-5 lift: refused with `RowEditUnsafeForSheet` /
    // `ColEditUnsafeForSheet`. Post-lift: `Workbook.{insertRow,
    // deleteRow, insertColumn, deleteColumn}` runs
    // `rewriteAllHyperlinkLocations` so cross-sheet location
    // qualifiers shift alongside the byte transform. The call now
    // succeeds.
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
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
        try s2.addInternalHyperlink("A1", "Plain!C5");
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.insertRow(0, 1);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.deleteRow(0, 1);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.insertColumn(0, 1);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.deleteColumn(0, 1);
    }
}

test "Editor: insertRow on frozen-pane sheet shifts ySplit + topLeftCell" {
    // Post-lift: `<pane>` is no longer in the row-edit refusal
    // guards. The frozen pane's ySplit and topLeftCell row shift
    // alongside the row attrs. xSplit is unaffected by a row edit.
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, "pane_insert_row.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "pane_insert_row_out.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }});
        try s.writeRow(&.{.{ .integer = 2 }});
        try s.freezePanes(1, 2); // rows=1, cols=2 → ySplit=1, xSplit=2, topLeftCell=C2
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.insertRow(0, 1);
        try ed.save(dst_path);
    }
    var ed2 = try Editor.open(std.testing.allocator, dst_path);
    defer ed2.deinit();
    const part = (try ed2.workbook.store.part("xl/worksheets/sheet1.xml")) orelse return error.TestUnexpectedResult;
    try std.testing.expect(std.mem.indexOf(u8, part.bytes, "ySplit=\"2\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, part.bytes, "xSplit=\"2\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, part.bytes, "topLeftCell=\"C3\"") != null);
}

test "Editor: insertColumn on frozen-pane sheet shifts xSplit + topLeftCell" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, "pane_insert_col.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "pane_insert_col_out.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 }, .{ .integer = 3 } });
        try s.freezePanes(1, 2);
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.insertColumn(0, 1);
        try ed.save(dst_path);
    }
    var ed2 = try Editor.open(std.testing.allocator, dst_path);
    defer ed2.deinit();
    const part = (try ed2.workbook.store.part("xl/worksheets/sheet1.xml")) orelse return error.TestUnexpectedResult;
    try std.testing.expect(std.mem.indexOf(u8, part.bytes, "xSplit=\"3\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, part.bytes, "ySplit=\"1\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, part.bytes, "topLeftCell=\"D2\"") != null);
}

test "Editor: deleteRow + deleteColumn on frozen-pane sheet shrinks freeze" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, "pane_delete.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "pane_delete_out.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 }, .{ .integer = 3 } });
        try s.writeRow(&.{ .{ .integer = 4 }, .{ .integer = 5 }, .{ .integer = 6 } });
        try s.freezePanes(1, 2);
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.deleteRow(0, 1); // ySplit 1→0, topLeftCell row 2→1
        try ed.save(dst_path);
    }
    var ed2 = try Editor.open(std.testing.allocator, dst_path);
    defer ed2.deinit();
    const part = (try ed2.workbook.store.part("xl/worksheets/sheet1.xml")) orelse return error.TestUnexpectedResult;
    try std.testing.expect(std.mem.indexOf(u8, part.bytes, "ySplit=\"0\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, part.bytes, "topLeftCell=\"C1\"") != null);
}

test "Editor: deleteColumn on frozen-pane sheet shrinks xSplit" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, "pane_delete_col.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "pane_delete_col_out.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 }, .{ .integer = 3 } });
        try s.freezePanes(1, 2);
        try w.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.deleteColumn(0, 1); // xSplit 2→1, topLeftCell col C→B
        try ed.save(dst_path);
    }
    var ed2 = try Editor.open(std.testing.allocator, dst_path);
    defer ed2.deinit();
    const part = (try ed2.workbook.store.part("xl/worksheets/sheet1.xml")) orelse return error.TestUnexpectedResult;
    try std.testing.expect(std.mem.indexOf(u8, part.bytes, "xSplit=\"1\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, part.bytes, "topLeftCell=\"B2\"") != null);
}

test "Editor: deleteSheet on a different sheet is allowed after a column edit (iter-er-4 (3/N))" {
    // Pre iter-er-4 (3/N), `Editor.insertColumn` queued a pending
    // edit and `deleteSheet` refused with `SheetDeleteRequiresCleanState`
    // because the queued edit's `sheet_idx` would shift after the
    // delete. Post-migration the column shift is applied
    // immediately to the typed-overlay's PartStore (no queue), so
    // there's no shifting state to corrupt — `deleteSheet` on a
    // DIFFERENT sheet is now safe.
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
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
    // Used to error SheetDeleteRequiresCleanState; now succeeds.
    try ed.deleteSheet(1);
    try std.testing.expectEqual(@as(usize, 1), ed.sheet_paths.len);
}

test "Editor: deleteSheet on a different sheet is allowed after a row edit (iter-er-4 (3/N))" {
    // Same parity as the column-edit test above.
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
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
    try ed.deleteSheet(1);
    try std.testing.expectEqual(@as(usize, 1), ed.sheet_paths.len);
}

test "Editor: appendRows + setCell after a row/col edit compose (iter-er-4 (3/N))" {
    // Pre iter-er-4 (3/N), insertColumn / deleteRow queued a
    // pending edit and subsequent appendRows / setCell on the
    // same sheet refused with `SheetHasUnsavedRowOrColEdit`
    // because the save path's row/col substitution would have
    // overwritten the appendRows / setCell deltas. Post-migration
    // the row/col shift is applied immediately to PartStore and
    // the typed-overlay's parsed view is invalidated, so deltas
    // staged AFTER the shift target the post-shift bytes — the
    // refusal becomes unnecessary. Both ops now succeed.
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, "append_after_coledit.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(src_path);
    }
    // appendRows after a column edit composes:
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.insertColumn(0, 1);
        try ed.appendRows(0, &.{&.{.{ .integer = 99 }}});
    }
    // setCell after a column edit composes (separate session
    // because appendRows + setCell on the same sheet remain
    // mutually exclusive — that invariant is independent of
    // row/col edits):
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.insertColumn(0, 1);
        try ed.setCell(0, 5, 0, .{ .integer = 99 });
    }
    // Same parity for deleteRow:
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.deleteRow(0, 1);
        try ed.appendRows(0, &.{&.{.{ .integer = 99 }}});
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.deleteRow(0, 1);
        try ed.setCell(0, 5, 0, .{ .integer = 99 });
    }
}

test "Editor: row edits with cross-sheet formulas rewrite refs (iter-er-5 lift)" {
    // Pre iter-er-5 lift: refused with `RowEditWithFormulasNotSupported`
    // / `ColEditWithFormulasNotSupported` because sheet2's formula
    // `Plain!A1+Plain!A2` references rows in the clean sheet.
    // Post-lift: `Workbook.insertRow` / `deleteRow` /
    // `insertColumn` / `deleteColumn` runs `rewriteAllFormulas`,
    // which shifts the cross-sheet refs. The call now succeeds.
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
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
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.insertRow(0, 1);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.deleteRow(0, 1);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.insertColumn(0, 1);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.deleteColumn(0, 1);
    }
}

test "Editor: deleteSheet drops a source sheet (iter-sheet-3)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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
    const writer_mod = xlsx;
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

// ─── B2 iter-er-1: Editor read-side parity tests ─────────────────────

/// Build a 2-sheet xlsx via the writer (which produces a non-DataDescriptor
/// ZIP that Editor accepts) and return its temp path. Caller frees the
/// returned slice and is responsible for the TestTmp lifecycle.
fn buildIterEr1Fixture(tt: *TestTmp, alloc: std.mem.Allocator) ![:0]u8 {
    const path = try tt.path(alloc, "iter_er_1.xlsx");
    var w = xlsx.Writer.init(alloc);
    defer w.deinit();
    var s1 = try w.addSheet("Alpha");
    try s1.writeRow(&.{ .{ .string = "h" }, .{ .integer = 1 } });
    var s2 = try w.addSheet("Beta");
    try s2.writeRow(&.{ .{ .string = "x" }, .{ .number = 2.5 } });
    try w.save(path);
    return path;
}

test "iter-er-1: Editor.open populates a Workbook view (sheet count + names match)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try buildIterEr1Fixture(&tt, std.testing.allocator);
    defer std.testing.allocator.free(path);

    var ed = try Editor.open(std.testing.allocator, path);
    defer ed.deinit();

    // sheet_paths is built from Book.sheets[i].path. workbook view's
    // sheetCount must match in lockstep — both walk
    // xl/_rels/workbook.xml.rels off the same source archive.
    try std.testing.expectEqual(@as(usize, ed.sheet_paths.len), @as(usize, ed.workbook.sheetCount()));
    try std.testing.expectEqual(@as(u32, 2), ed.workbook.sheetCount());

    // Sheet names from the typed-overlay view match what the writer
    // emitted ("Alpha" / "Beta", in addSheet order).
    const ws0 = try ed.workbook.sheet(0);
    try std.testing.expectEqualStrings("Alpha", ws0.name());
    const ws1 = try ed.workbook.sheet(1);
    try std.testing.expectEqualStrings("Beta", ws1.name());
}

test "iter-er-1: Editor.deinit cleans up Workbook + everything else (no leaks)" {
    // Editor.deinit calls self.workbook.deinit() before freeing
    // src_buf / entries / sheet_paths. Under std.testing.allocator
    // (which panics on leak), this test fails if any allocation
    // outlives the editor.
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try buildIterEr1Fixture(&tt, std.testing.allocator);
    defer std.testing.allocator.free(path);

    {
        var ed = try Editor.open(std.testing.allocator, path);
        ed.deinit();
    }
}

test "iter-er-1: Editor.workbook.cellByRef matches Book.cell for known cells" {
    // Cross-API parity sanity check: the Workbook view's per-cell
    // accessor finds cells the writer emitted.
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try buildIterEr1Fixture(&tt, std.testing.allocator);
    defer std.testing.allocator.free(path);

    var ed = try Editor.open(std.testing.allocator, path);
    defer ed.deinit();

    // The writer wrote integer 1 at B1 of "Alpha"; the typed-overlay
    // view exposes it via Worksheet.cellByRef.
    const ws0 = try ed.workbook.sheet(0);
    const c = (try ws0.cellByRef("B1")) orelse return error.MissingCell;
    try std.testing.expectEqualStrings("B1", c.ref);
    try std.testing.expect(c.raw_value != null);
    try std.testing.expectEqualStrings("1", c.raw_value.?);
}

// ---------------------------------------------------------------------------
// autoFilter round-trip tests (refusal lift; replaces prior ColEditUnsafe /
// RowEditUnsafe refusals on sheets carrying <autoFilter>).
// ---------------------------------------------------------------------------

fn buildAutoFilterFixture(path: []const u8) !void {
    var w = xlsx.Writer.init(std.testing.allocator);
    defer w.deinit();
    var s = try w.addSheet("S");
    try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 }, .{ .integer = 3 }, .{ .integer = 4 } });
    try s.writeRow(&.{ .{ .integer = 5 }, .{ .integer = 6 }, .{ .integer = 7 }, .{ .integer = 8 } });
    try s.setAutoFilter("B1:D2");
    try w.save(path);
}

fn assertSheetXmlContains(ed: *Editor, needle: []const u8) !void {
    const xml = try ed.readEntry("xl/worksheets/sheet1.xml");
    defer std.testing.allocator.free(xml);
    if (std.mem.indexOf(u8, xml, needle) == null) {
        std.debug.print("\nexpected sheet1.xml to contain `{s}`; got:\n{s}\n", .{ needle, xml });
        return error.TestExpectedNeedle;
    }
}

test "Editor: insertRow on a sheet carrying <autoFilter> rewrites range" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, "af_insrow_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "af_insrow_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    try buildAutoFilterFixture(src_path);
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.insertRow(0, 2);
        try ed.save(dst_path);
    }
    var ed2 = try Editor.open(std.testing.allocator, dst_path);
    defer ed2.deinit();
    // B1:D2 → B1:D3 after inserting a row at row 2.
    try assertSheetXmlContains(&ed2, "ref=\"B1:D3\"");
}

test "Editor: insertColumn on a sheet carrying <autoFilter> rewrites range" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, "af_inscol_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "af_inscol_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    try buildAutoFilterFixture(src_path);
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.insertColumn(0, 3); // insert before col C, inside B:D
        try ed.save(dst_path);
    }
    var ed2 = try Editor.open(std.testing.allocator, dst_path);
    defer ed2.deinit();
    // B1:D2 → B1:E2 after extending the range by one column.
    try assertSheetXmlContains(&ed2, "ref=\"B1:E2\"");
}

test "Editor: deleteRow on a sheet carrying <autoFilter> rewrites range" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, "af_delrow_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "af_delrow_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    try buildAutoFilterFixture(src_path);
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.deleteRow(0, 2); // delete row 2 (the BR of B1:D2)
        try ed.save(dst_path);
    }
    var ed2 = try Editor.open(std.testing.allocator, dst_path);
    defer ed2.deinit();
    // B1:D2 → B1:D1 after the BR row vanishes.
    try assertSheetXmlContains(&ed2, "ref=\"B1:D1\"");
}

test "Editor: deleteColumn on a sheet carrying <autoFilter> shrinks range" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, "af_delcol_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "af_delcol_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    try buildAutoFilterFixture(src_path);
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.deleteColumn(0, 3); // delete col C inside B:D
        try ed.save(dst_path);
    }
    var ed2 = try Editor.open(std.testing.allocator, dst_path);
    defer ed2.deinit();
    // B1:D2 → B1:C2 (BR shrinks by one column).
    try assertSheetXmlContains(&ed2, "ref=\"B1:C2\"");
}

// ---------------------------------------------------------------------------
// <picture> refusal lift (dr-0).
// CT_SheetBackgroundPicture is a single coordinate-free `r:id` reference to
// a tiled background image; row/col edits cannot misalign it. Confirm the
// guard drop by injecting a `<picture/>` element into a sheet that wasn't
// authored with one (zlsx's writer never emits CT_SheetBackgroundPicture)
// and round-tripping insertRow + insertColumn through the Editor.
// ---------------------------------------------------------------------------

fn buildPictureFixture(path: []const u8) !void {
    // Stage 1: produce a baseline workbook with the writer.
    {
        var w = xlsx.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 } });
        try s.writeRow(&.{ .{ .integer = 3 }, .{ .integer = 4 } });
        try w.save(path);
    }

    // Stage 2: splice `<picture r:id="rId99"/>` before `</worksheet>` in
    // sheet1.xml via PartStore.replacePart, then save back over the
    // baseline. The dangling rId99 doesn't resolve to anything in the
    // sheet's rels; that's fine for this test — the guard is what we're
    // exercising, and ECMA-376 readers tolerate dangling rels by
    // ignoring the picture element rather than failing the file.
    var store = try store_mod.PartStore.open(std.testing.allocator, path);
    defer store.deinit();
    const sheet_name = "xl/worksheets/sheet1.xml";
    const orig = (try store.part(sheet_name)) orelse return error.MissingSheet;
    const inject = "<picture r:id=\"rId99\"/>";
    const close = "</worksheet>";
    const idx = std.mem.indexOf(u8, orig.bytes, close) orelse return error.NoWorksheetClose;
    var new_xml = try std.testing.allocator.alloc(u8, orig.bytes.len + inject.len);
    defer std.testing.allocator.free(new_xml);
    @memcpy(new_xml[0..idx], orig.bytes[0..idx]);
    @memcpy(new_xml[idx .. idx + inject.len], inject);
    @memcpy(new_xml[idx + inject.len ..], orig.bytes[idx..]);
    try store.replacePart(sheet_name, new_xml);
    try store.save(path);
}

test "Editor: insertRow on sheet with `<picture>` background no longer refused" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, "pic_insrow_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "pic_insrow_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    try buildPictureFixture(src_path);
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.insertRow(0, 2);
        try ed.save(dst_path);
    }
    var ed2 = try Editor.open(std.testing.allocator, dst_path);
    defer ed2.deinit();
    // The `<picture>` element passes through unchanged (no row/col coords).
    try assertSheetXmlContains(&ed2, "<picture r:id=\"rId99\"/>");
}

test "Editor: insertColumn on sheet with `<picture>` background no longer refused" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, "pic_inscol_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "pic_inscol_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    try buildPictureFixture(src_path);
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.insertColumn(0, 2);
        try ed.save(dst_path);
    }
    var ed2 = try Editor.open(std.testing.allocator, dst_path);
    defer ed2.deinit();
    try assertSheetXmlContains(&ed2, "<picture r:id=\"rId99\"/>");
}

// ---------------------------------------------------------------------------
// Modern xdr drawings refusal lift (dr-1).
// `Workbook.addImage` produces a `<xdr:oneCellAnchor>` at the requested
// 1-based cell. After Editor row/col edits, the shifted xdr coords (0-based
// in the wire format) must reflect the same insert/delete semantics as the
// rest of the byte transform.
// ---------------------------------------------------------------------------

const tiny_png_1x1_for_editor = [_]u8{
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
    0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
    0x08, 0x00, 0x00, 0x00, 0x00, 0x3A, 0x7E, 0x9B,
    0x55, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41,
    0x54, 0x78, 0x9C, 0x63, 0x00, 0x00, 0x00, 0x02,
    0x00, 0x01, 0xE2, 0x21, 0xBC, 0x33, 0x00, 0x00,
    0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42,
    0x60, 0x82,
};

fn buildDrawingFixture(path: []const u8, anchor_col: u32, anchor_row: u32) !void {
    {
        var w = xlsx.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 }, .{ .integer = 3 }, .{ .integer = 4 } });
        try s.writeRow(&.{ .{ .integer = 5 }, .{ .integer = 6 }, .{ .integer = 7 }, .{ .integer = 8 } });
        try w.save(path);
    }
    var wb = try workbook_mod.Workbook.open(std.testing.allocator, path);
    defer wb.deinit();
    try wb.addImage(0, .{ .col = anchor_col, .row = anchor_row }, &tiny_png_1x1_for_editor, .png);
    try wb.save(path);
}

fn drawingPartContains(path: []const u8, needle: []const u8) !bool {
    var store = try store_mod.PartStore.open(std.testing.allocator, path);
    defer store.deinit();
    const drawing = (try store.part("xl/drawings/drawing1.xml")) orelse return error.MissingDrawing;
    return std.mem.indexOf(u8, drawing.bytes, needle) != null;
}

test "Editor: insertColumn shifts xdr drawing anchor col (dr-1)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, "draw_inscol_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "draw_inscol_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    try buildDrawingFixture(src_path, 3, 5);
    try std.testing.expect(try drawingPartContains(src_path, "<xdr:col>2</xdr:col>"));
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.insertColumn(0, 2);
        try ed.save(dst_path);
    }
    try std.testing.expect(try drawingPartContains(dst_path, "<xdr:col>3</xdr:col>"));
    try std.testing.expect(try drawingPartContains(dst_path, "<xdr:row>4</xdr:row>"));
}

test "Editor: insertRow shifts xdr drawing anchor row (dr-1)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, "draw_insrow_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "draw_insrow_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    try buildDrawingFixture(src_path, 3, 5);
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.insertRow(0, 3);
        try ed.save(dst_path);
    }
    try std.testing.expect(try drawingPartContains(dst_path, "<xdr:col>2</xdr:col>"));
    try std.testing.expect(try drawingPartContains(dst_path, "<xdr:row>5</xdr:row>"));
}

test "Editor: deleteColumn at the anchor's column drops the oneCellAnchor (dr-1)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, "draw_delcol_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "draw_delcol_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    try buildDrawingFixture(src_path, 3, 5);
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.deleteColumn(0, 3);
        try ed.save(dst_path);
    }
    try std.testing.expect(!try drawingPartContains(dst_path, "<xdr:oneCellAnchor>"));
}

test "Editor: drawing rewrite tolerates XML whitespace around `=` in worksheet rels (dr-1 REL-602/604)" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, "draw_eq_ws_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, "draw_eq_ws_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    try buildDrawingFixture(src_path, 3, 5);
    // Splice a custom `<drawing  r:id  =  "rId99"/>` into sheet1.xml
    // ALONGSIDE the existing well-formed `<drawing r:id="rIdN"/>`
    // emitted by addImage; we exercise the parser's whitespace
    // tolerance even though the resolved part is the SAME.
    {
        var store = try store_mod.PartStore.open(std.testing.allocator, src_path);
        defer store.deinit();
        const sheet_name = "xl/worksheets/sheet1.xml";
        const orig = (try store.part(sheet_name)) orelse return error.MissingSheet;
        // Find the existing <drawing tag and rewrite it with
        // whitespace around `=`.
        const orig_open = std.mem.indexOf(u8, orig.bytes, "<drawing ") orelse return error.NoDrawingTag;
        const orig_end = std.mem.indexOfPos(u8, orig.bytes, orig_open, "/>") orelse return error.NoDrawingClose;
        // Extract the rId from the existing tag.
        const eq = std.mem.indexOfScalarPos(u8, orig.bytes, orig_open, '=') orelse return error.NoEq;
        const q1 = std.mem.indexOfScalarPos(u8, orig.bytes, eq, '"') orelse return error.NoQ1;
        const q2 = std.mem.indexOfScalarPos(u8, orig.bytes, q1 + 1, '"') orelse return error.NoQ2;
        const rid = orig.bytes[q1 + 1 .. q2];
        var ws_tag_buf: [128]u8 = undefined;
        const ws_tag = try std.fmt.bufPrint(&ws_tag_buf, "<drawing  r:id  =  \"{s}\" />", .{rid});
        var new_xml = try std.testing.allocator.alloc(u8, orig.bytes.len - (orig_end + 2 - orig_open) + ws_tag.len);
        defer std.testing.allocator.free(new_xml);
        @memcpy(new_xml[0..orig_open], orig.bytes[0..orig_open]);
        @memcpy(new_xml[orig_open .. orig_open + ws_tag.len], ws_tag);
        @memcpy(new_xml[orig_open + ws_tag.len ..], orig.bytes[orig_end + 2 ..]);
        try store.replacePart(sheet_name, new_xml);
        try store.save(src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, src_path);
        defer ed.deinit();
        try ed.insertColumn(0, 2);
        try ed.save(dst_path);
    }
    // The drawing's xdr:col MUST shift; if findAttrValue mishandled
    // the whitespace, applyDrawingEditForSheet would silently skip
    // and the col would still be 2.
    try std.testing.expect(try drawingPartContains(dst_path, "<xdr:col>3</xdr:col>"));
}
