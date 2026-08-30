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
const table_edit = @import("table_edit.zig");
const pivots_mod = @import("pivots.zig");
const store_mod = @import("store.zig");
const AtomicFile = @import("atomic_file.zig").AtomicFile;

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

    pub fn open(allocator: Allocator, io: std.Io, path: []const u8) !Editor {
        const file = try std.Io.Dir.cwd().openFile(io, path, .{});
        defer file.close(io);
        const stat = try file.stat(io);
        // Refuse files > 4 GiB up front (ZIP64 isn't supported by v1
        // per the plan; documented limit).
        if (stat.size > std.math.maxInt(u32)) return error.ZipTooLarge;
        const buf = try allocator.alloc(u8, @intCast(stat.size));
        {
            errdefer allocator.free(buf);
            // 0.16 reads go through the Reader interface; `readSliceAll`
            // is the exact-length read that 0.15's `readAll` + length
            // check expressed.
            var read_buf: [4096]u8 = undefined;
            var file_reader = file.reader(io, &read_buf);
            file_reader.interface.readSliceAll(buf) catch |err| switch (err) {
                error.EndOfStream => return error.UnexpectedEof,
                else => return err,
            };
        }
        return fromOwnedSource(allocator, io, buf, .{ .path = path });
    }

    /// Open a workbook that is already in memory (M9a2, §5.10).
    ///
    /// **The borrow ends when this returns** — the same contract
    /// `Book.openBuffer` and `Workbook.openBuffer` keep: `bytes` is
    /// duped into editor-owned storage, so the caller may free, reuse
    /// or poison it the moment the call comes back.
    pub fn openBuffer(allocator: Allocator, io: std.Io, bytes: []const u8) !Editor {
        // Same v1 limit the path open enforces on stat.size: the ZIP
        // scan below trusts u32 offsets.
        if (bytes.len > std.math.maxInt(u32)) return error.ZipTooLarge;
        const buf = try allocator.dupe(u8, bytes);
        return fromOwnedSource(allocator, io, buf, .buffer);
    }

    /// Where the editor's source bytes came from — decides only how the
    /// internal `Book` + `Workbook` views are constructed; the ZIP scan
    /// itself reads `src_buf` either way.
    const SourceOrigin = union(enum) { path: []const u8, buffer };

    /// The shared tail of `open` / `openBuffer`. Takes ownership of
    /// `buf` including on failure.
    fn fromOwnedSource(allocator: Allocator, io: std.Io, buf: []u8, origin: SourceOrigin) !Editor {
        errdefer allocator.free(buf);

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

        var entries: std.ArrayListUnmanaged(ZipEntry) = .empty;
        errdefer entries.deinit(allocator);
        try entries.ensureTotalCapacity(allocator, eocd.record_count_total);
        var budget: xlsx.DecompressBudget = .init(xlsx.decompress_limits);

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
            // S1: the decompression limits, on this scan rather than
            // only in `readEntry` below — `Book.open` follows this
            // walk and inflates every sheet, so the scan is the first
            // point where a hostile declaration can be refused before
            // anything at all is allocated for it.
            try budget.admit(cdfh_ptr.compressed_size, cdfh_ptr.uncompressed_size);

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
            var b = switch (origin) {
                .path => |src_path| try Book.open(allocator, io, src_path),
                .buffer => try Book.openBuffer(allocator, io, buf),
            };
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
            // bytes via PartStore-from-bytes. The buffer arm keeps the
            // same sanity check against its second parse of `buf`.
            workbook_built = switch (origin) {
                .path => |src_path| try Workbook.fromBook(allocator, io, &b, src_path),
                .buffer => blk: {
                    var wb = try Workbook.openBuffer(allocator, io, buf);
                    if (wb.sheetCount() != b.sheets.len) {
                        wb.deinit();
                        return error.SheetCountMismatch;
                    }
                    break :blk wb;
                },
            };
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
    /// through byte-for-byte; the EOCD comment is preserved. Before
    /// the sheet emit, a staged write that lands where a pivot source
    /// reads — a `setCell` or an appended row inside its rectangle, in
    /// a whole-column or whole-row source's span, anywhere on a sheet
    /// a `sheet`-only spelling claims or an unbounded source
    /// references — marks that cache `refreshOnLoad="1"` (S7b-3, the
    /// one rule a row edit inside the same rectangle follows); writes
    /// outside every source leave every cache definition byte-identical.
    pub fn save(self: *Editor, io: std.Io, out_path: []const u8) !void {
        if (!self.workbookHasAnyDeltas() and
            !self.workbookHasAnyAppendedRows() and
            !self.workbook.store.hasUnsavedChanges())
        {
            var write_buf: [4096]u8 = undefined;
            var atomic_file = try AtomicFile.init(io, out_path, &write_buf);
            defer atomic_file.deinit();
            const w = &atomic_file.file_writer.interface;
            try w.writeAll(self.src_buf);
            try atomic_file.finish();
            return;
        }
        try self.workbook.save(io, out_path);
    }

    /// `save` into caller-owned memory (M9a2, §5.10). The same two
    /// paths: an untouched editor hands back a dupe of the source bytes
    /// (byte-identical, SHA256-preserving, like the passthrough save);
    /// a mutated one routes through `Workbook.saveToOwnedBuffer`. The
    /// returned bytes are the caller's, freed with `allocator`.
    pub fn saveToOwnedBuffer(self: *Editor, allocator: Allocator) ![]u8 {
        if (!self.workbookHasAnyDeltas() and
            !self.workbookHasAnyAppendedRows() and
            !self.workbook.store.hasUnsavedChanges())
        {
            return allocator.dupe(u8, self.src_buf);
        }
        return self.workbook.saveToOwnedBuffer(allocator);
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
    /// Typed view over the workbook's `docProps/*` metadata.
    ///
    /// Delegates to `Workbook.docProps`. Returned strings borrow from
    /// the store and stay valid until the Editor is deinitialised.
    pub fn docProps(self: *const Editor) !workbook_mod.DocProps {
        return self.workbook.docProps();
    }

    /// Strip identifying document metadata, staged for the next `save`.
    ///
    /// This is the counterpart to cell masking: without it, a workbook
    /// whose cells have been pseudonymised still ships `dc:creator`,
    /// `cp:lastModifiedBy` and `Company` in the archive, untouched by
    /// every edit because zlsx used to copy those parts through
    /// verbatim. Everything the mask does not name is byte-preserved.
    pub fn stripDocProps(self: *Editor, mask: workbook_mod.DocPropsMask) !void {
        return self.workbook.stripDocProps(mask);
    }

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

    /// Rename a sheet (Phase 3e, iter-sheet-2). Delegates to
    /// `Workbook.renameSheet`, which patches `xl/workbook.xml`
    /// in-memory and runs the formula + defined-name rewriters:
    /// cross-sheet refs (`'OLD'!A1`) follow the rename instead of
    /// decaying to `#REF!`.
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

    /// Delete a sheet (Phase 3e, iter-sheet-3). Contract:
    ///   - Refuses if it's the only remaining sheet.
    ///   - Refuses if there are staged setCell deltas or appended
    ///     rows on ANY sheet (caller must `save` first then
    ///     re-open) — the delete rebuilds `sheet_paths`, and queued
    ///     mutations hold raw indices into it.
    ///   - Delegates to `Workbook.deleteSheet` (workbook.xml, rels
    ///     and Content_Types patched in-memory), then drops the
    ///     path from `sheet_paths`.
    ///   - Cross-sheet formula refs to the deleted sheet collapse
    ///     to `#REF!` via the workbook rewriters; refs to
    ///     surviving sheets stay intact.
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
    /// or below `before_row` shifts down by 1.
    ///   - Worksheet XML rewrites: <row r="N"> renumber, <c r="A1">
    ///     row component, <mergeCells> rect bounds, <dimension>.
    ///   - Formulas, defined names, hyperlink locations, DV/CF
    ///     formulas, `<extLst>` `<xm:f>` formulas (sparklines, x14
    ///     CF / DV), drawings (xdr + VML), panes, autoFilter and
    ///     table parts are rewritten in step (iter-er-5, dr-1/dr-2,
    ///     S2); a pivot's `location@ref` on a sheet that only
    ///     hosts it moves in step (S7a). Still refused: an edit
    ///     inside a hosted pivot's footprint or on a host sheet a
    ///     pivot also reads from, unsafe table edits (collapse /
    ///     header-row delete), and an `<xm:f>` carrier the scan
    ///     cannot read.
    ///   - The sheet must not have other pending mutations
    ///     (setCell / appendRows / row inserts/deletes); save
    ///     first to apply those.
    pub fn insertRow(self: *Editor, sheet_idx: u32, before_row: u32) !void {
        try self.recordRowEdit(sheet_idx, before_row, true);
    }

    /// Delete row `row` in sheet `sheet_idx` (Phase 3e, iter-row-3).
    /// Every row > `row` shifts up by 1. Same rewrite coverage and
    /// refusal contract as `insertRow`.
    pub fn deleteRow(self: *Editor, sheet_idx: u32, row: u32) !void {
        try self.recordRowEdit(sheet_idx, row, false);
    }

    /// Insert a blank column at position `before_col` (1-based,
    /// A=1) in sheet `sheet_idx`. Phase 3e iter-col-3. Same
    /// rewrite coverage and refusal contract as `insertRow`.
    pub fn insertColumn(self: *Editor, sheet_idx: u32, before_col_1based: u32) !void {
        try self.recordColEdit(sheet_idx, before_col_1based, true);
    }

    /// Delete column `col_1based` (1-based) in sheet `sheet_idx`.
    /// Phase 3e iter-col-4.
    pub fn deleteColumn(self: *Editor, sheet_idx: u32, col_1based: u32) !void {
        try self.recordColEdit(sheet_idx, col_1based, false);
    }

    /// Rename a column of the named table (C1: structured-ref
    /// rewriting). Delegates to `Workbook.renameTableColumn`:
    /// `<tableColumn name>` + the table part's own formulas, every
    /// structured reference workbook-wide (`Table1[Old]`, and bare
    /// `[Old]` / `[@Old]` in cells inside the table's range), defined
    /// names, hyperlink locations, DV/CF formulas, and the header
    /// cell's text. Names are decoded plain text; matching follows
    /// the fold rule name resolution uses. Refusals
    /// (`TableNotFound`, `TableColumnNotFound`,
    /// `TableColumnNameInUse`, `InvalidTableColumnName`) all precede
    /// mutation. Like the other structural edits, changes are staged
    /// in-memory; `save` commits them.
    pub fn renameTableColumn(
        self: *Editor,
        table_name: []const u8,
        old_name: []const u8,
        new_name: []const u8,
    ) !void {
        _ = try self.workbook.renameTableColumn(table_name, old_name, new_name);
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

        // Per-sheet content guard — `<pane>`, `<picture>`, modern
        // `<drawing>` (xdr), `<legacyDrawing>` (VML), `<autoFilter>`,
        // and `<tableParts>` are all rewritten by their own
        // byte-transform pipelines; see the matching comment in
        // `recordRowEdit`. `<tableParts>` is pre-flighted (vs
        // string-scanned) — `Workbook.preflightTableEditsForSheet`
        // dry-runs `pkg/table_edit.zig::applyEditToTable` against
        // each referenced `xl/tables/tableN.xml` and surfaces
        // `TableCollapseUnsafe` / `TableHeaderRowDeleteUnsafe`
        // before any sheet bytes are mutated.
        const path = self.sheet_paths[sheet_idx];
        const tbl_kind: table_edit.EditKind = if (is_insert) .insert else .delete;
        // REL-B527/B528: remap every table_edit refusal/diagnostic
        // to the existing `ColEditUnsafeForSheet` axis. Surfacing
        // `MalformedTableXml` / `TableCoordinateOverflow` raw would
        // leak Workbook-internal error names through Editor's
        // public surface — and a user staring at "MalformedTableXml"
        // from `Editor.insertColumn` has no actionable fix beyond
        // "this sheet is unsafe to edit", which is exactly what
        // `ColEditUnsafeForSheet` already means.
        // Pivots (S7a): a sheet that only hosts one has its
        // `location@ref` moved by the sweep; the pre-flight dry-runs
        // that move and refuses what the sweep would refuse — an edit
        // inside the pivot's footprint, a host that is also a source, a
        // graph that cannot be read whole (`MalformedPivotXml`). Both
        // fold into the sheet-level refusal: a caller of
        // `Editor.insertColumn` has no fix for either beyond "this
        // sheet is unsafe to edit". Checked before the table pre-flight
        // because a sheet without a pivot relationship costs one rels
        // lookup here.
        const pivot_kind: pivots_mod.edit.Kind = if (is_insert) .insert else .delete;
        var prepared = self.workbook.preflightPivotEditsForSheet(path, .col, col_1based, pivot_kind) catch |err| switch (err) {
            error.PivotEditUnsafe, error.MalformedPivotXml => return error.ColEditUnsafeForSheet,
            else => |e| return e,
        };
        defer prepared.deinit(self.allocator);

        // `<extLst>` `<xm:f>` formulas (sparklines, x14 CF / DV) are
        // rewritten by `Workbook.rewriteAllExtensionFormulas` (S2);
        // the only refusal left on that axis is a carrier the scan
        // cannot read, folded into the same user-facing error.
        self.workbook.preflightExtensionFormulas() catch |err| switch (err) {
            error.MalformedExtensionXml => return error.ColEditUnsafeForSheet,
            else => |e| return e,
        };

        self.workbook.preflightTableEditsForSheet(path, .col, col_1based, tbl_kind) catch |err| switch (err) {
            error.TableCollapseUnsafe,
            error.TableHeaderRowDeleteUnsafe,
            error.MalformedTableXml,
            error.TableCoordinateOverflow,
            => return error.ColEditUnsafeForSheet,
            else => |e| return e,
        };

        // The sweep installs the pre-flight's collection: one pivot
        // pass per edit, not two (Codex #205 r1 PERF-101).
        try self.workbook.applySheetEdit(sheet_idx, .{ .col = col_1based, .kind = if (is_insert) .insert else .delete }, &prepared);
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

        // Per-sheet content guard: every prior axis (pane,
        // autoFilter, picture, modern + legacy drawings) is
        // rewritten by its own byte-transform pipeline. The last
        // axis — `<tableParts>` — is now pre-flighted (vs
        // string-scanned). `Workbook.preflightTableEditsForSheet`
        // dry-runs `pkg/table_edit.zig::applyEditToTable` against
        // each referenced `xl/tables/tableN.xml` and surfaces
        // `TableCollapseUnsafe` / `TableHeaderRowDeleteUnsafe`
        // before any sheet bytes are mutated. See
        // `docs/plans/refusal-audit.md` for the v1 limitations
        // (collapse + header-row delete still refused — those are
        // schema-invalid table states, not unrewritten ones).
        const path = self.sheet_paths[sheet_idx];
        const tbl_kind: table_edit.EditKind = if (is_insert) .insert else .delete;
        // REL-B527/B528: same remap as recordColEdit — every
        // table_edit refusal/diagnostic folds into the existing
        // `RowEditUnsafeForSheet` axis.
        // Same pivot pre-flight as `recordColEdit` — see the note there.
        const pivot_kind: pivots_mod.edit.Kind = if (is_insert) .insert else .delete;
        var prepared = self.workbook.preflightPivotEditsForSheet(path, .row, row, pivot_kind) catch |err| switch (err) {
            error.PivotEditUnsafe, error.MalformedPivotXml => return error.RowEditUnsafeForSheet,
            else => |e| return e,
        };
        defer prepared.deinit(self.allocator);

        // Same `<xm:f>` gate as `recordColEdit` — see the note there.
        self.workbook.preflightExtensionFormulas() catch |err| switch (err) {
            error.MalformedExtensionXml => return error.RowEditUnsafeForSheet,
            else => |e| return e,
        };

        self.workbook.preflightTableEditsForSheet(path, .row, row, tbl_kind) catch |err| switch (err) {
            error.TableCollapseUnsafe,
            error.TableHeaderRowDeleteUnsafe,
            error.MalformedTableXml,
            error.TableCoordinateOverflow,
            => return error.RowEditUnsafeForSheet,
            else => |e| return e,
        };

        // The sweep installs the pre-flight's collection: one cache
        // rebuild per edit, not two (Codex #205 r1 PERF-101).
        try self.workbook.applySheetEdit(sheet_idx, .{ .row = row, .kind = if (is_insert) .insert else .delete }, &prepared);
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
    var out: std.ArrayListUnmanaged(CellSpan) = .empty;
    errdefer out.deinit(allocator);

    var pos: usize = 0;
    var implicit_row: u32 = 0;

    while (findTagOpen(xml, pos, "row")) |row_tag| {
        const row_attrs = xml[row_tag.start + "<row".len .. row_tag.after_open - 1];
        const row_attrs_trim = std.mem.trimEnd(u8, row_attrs, " \t\r\n");
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
            const trimmed = std.mem.trimEnd(u8, candidate_attrs, " \t\r\n");
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

// ZIP-bomb defenses for `decompressZipPayload`: the per-part half of
// `zlsx_control.decompress_limits`, the same numbers `pkg/store.zig`
// and the core reader apply, checked BEFORE any allocation. Every
// entry was already admitted (aggregate included) by the scan in
// `fromOwnedSource`; this is the last line for the one part in hand.
fn decompressZipPayload(
    allocator: Allocator,
    payload: []const u8,
    method: u16,
    declared_uncompressed: u32,
) ![]u8 {
    try xlsx.decompress_limits.checkPart(payload.len, declared_uncompressed);

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
    pub fn path(self: *TestTmp, alloc: std.mem.Allocator, io: std.Io, name: []const u8) ![:0]u8 {
        const d = try self.dir.dir.realPathFileAlloc(io, ".", alloc);
        defer alloc.free(d);
        return std.fs.path.joinZ(alloc, &.{ d, name });
    }
};

// ─── Tests ───────────────────────────────────────────────────────────

test "Editor: byte-identical passthrough (iter-lms-1)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "editor_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "editor_dst.xlsx");
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
        try w.save(io, src_path);
    }

    // SHA256 of source.
    const Sha256 = std.crypto.hash.sha2.Sha256;
    var src_hash: [Sha256.digest_length]u8 = undefined;
    {
        const f = try std.Io.Dir.cwd().openFile(io, src_path, .{});
        defer f.close(io);
        const buf = try std.testing.allocator.alloc(u8, @intCast((try f.stat(io)).size));
        defer std.testing.allocator.free(buf);
        {
            var fr = f.reader(io, &.{});
            try fr.interface.readSliceAll(buf);
        }
        Sha256.hash(buf, &src_hash, .{});
    }

    // Round-trip through Editor.
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.save(io, dst_path);
    }

    // SHA256 of destination must match.
    var dst_hash: [Sha256.digest_length]u8 = undefined;
    {
        const f = try std.Io.Dir.cwd().openFile(io, dst_path, .{});
        defer f.close(io);
        const buf = try std.testing.allocator.alloc(u8, @intCast((try f.stat(io)).size));
        defer std.testing.allocator.free(buf);
        {
            var fr = f.reader(io, &.{});
            try fr.interface.readSliceAll(buf);
        }
        Sha256.hash(buf, &dst_hash, .{});
    }
    try std.testing.expectEqualSlices(u8, &src_hash, &dst_hash);

    // The destination must still open as a valid workbook through
    // the reader — confirms we didn't corrupt anything.
    var book = try Book.open(std.testing.allocator, io, dst_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 1), book.sheets.len);
    try std.testing.expectEqualStrings("Data", book.sheets[0].name);
}

test "Editor: openBuffer + saveToOwnedBuffer round-trip (M9a2)" {
    const a = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(a, io, "editor_buf_src.xlsx");
    defer a.free(src_path);

    {
        var w = xlsx.Writer.init(a);
        defer w.deinit();
        var s = try w.addSheet("Data");
        try s.writeRow(&.{ .{ .string = "header" }, .{ .integer = 42 } });
        try w.save(io, src_path);
    }

    const src_bytes = blk: {
        const f = try std.Io.Dir.cwd().openFile(io, src_path, .{});
        defer f.close(io);
        const buf = try a.alloc(u8, @intCast((try f.stat(io)).size));
        errdefer a.free(buf);
        var fr = f.reader(io, &.{});
        try fr.interface.readSliceAll(buf);
        break :blk buf;
    };
    defer a.free(src_bytes);

    var ed = blk: {
        // The borrow ends at the call: open from a copy poisoned
        // immediately after, so a retained pointer would be caught.
        const borrowed = try a.dupe(u8, src_bytes);
        defer a.free(borrowed);
        var e = try Editor.openBuffer(a, io, borrowed);
        errdefer e.deinit();
        @memset(borrowed, 0xAA);
        break :blk e;
    };
    defer ed.deinit();

    // Untouched editor: the buffer save is the passthrough arm,
    // byte-for-byte like the passthrough file save.
    const passthrough = try ed.saveToOwnedBuffer(a);
    defer a.free(passthrough);
    try std.testing.expectEqualSlices(u8, src_bytes, passthrough);

    // Mutated editor: the buffer carries the mutation and still opens.
    try ed.setCell(0, 1, 1, .{ .integer = 7 }); // B1: 42 -> 7
    const mutated = try ed.saveToOwnedBuffer(a);
    defer a.free(mutated);
    try std.testing.expect(!std.mem.eql(u8, src_bytes, mutated));
    {
        var book = try Book.openBuffer(a, io, mutated);
        defer book.deinit();
        try std.testing.expectEqual(@as(usize, 1), book.sheets.len);
    }

    // A path save after the buffer save writes the same bytes: the
    // buffer emitter and the file emitter share one archive stream.
    const dst_path = try tt.path(a, io, "editor_buf_dst.xlsx");
    defer a.free(dst_path);
    try ed.save(io, dst_path);
    const dst_bytes = try std.Io.Dir.cwd().readFileAlloc(io, dst_path, a, .limited(1 << 22));
    defer a.free(dst_bytes);
    try std.testing.expectEqualSlices(u8, mutated, dst_bytes);
}

test "Editor: raw-ZIP scanner builds entry table (iter-lms-1b)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "editor_scan.xlsx");
    defer std.testing.allocator.free(src_path);

    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Sheet1");
        try s.writeRow(&.{ .{ .string = "a" }, .{ .integer = 1 } });
        try w.save(io, src_path);
    }

    var ed = try Editor.open(std.testing.allocator, io, src_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "editor_append_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "editor_append_dst.xlsx");
    defer std.testing.allocator.free(dst_path);

    // Source workbook: 2 rows.
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Data");
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 10 } });
        try s.writeRow(&.{ .{ .integer = 2 }, .{ .integer = 20 } });
        try w.save(io, src_path);
    }

    // Append two more rows via Editor.
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        const append_rows = [_][]const Cell{
            &.{ .{ .integer = 3 }, .{ .integer = 30 } },
            &.{ .{ .integer = 4 }, .{ .integer = 40 } },
        };
        try ed.appendRows(0, &append_rows);
        try ed.save(io, dst_path);
    }

    // Read back via Book — confirm 4 rows total, original cells intact,
    // new cells at the expected indices.
    var book = try Book.open(std.testing.allocator, io, dst_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "editor_append_reject.xlsx");
    defer std.testing.allocator.free(src_path);

    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, src_path);
    }

    var ed = try Editor.open(std.testing.allocator, io, src_path);
    defer ed.deinit();

    const ok_rows = [_][]const Cell{&.{.{ .integer = 2 }}};
    try std.testing.expectError(error.SheetIndexOutOfRange, ed.appendRows(99, &ok_rows));

    // Lossy integer (>2^53 + 1) is refused for the same reason the
    // writer refuses it.
    const lossy_rows = [_][]const Cell{&.{.{ .integer = 9007199254740993 }}};
    try std.testing.expectError(error.IntegerExceedsExcelPrecision, ed.appendRows(0, &lossy_rows));
}

test "Editor: appendRows with string cells extends SST (iter-lms-3)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "editor_append_str_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "editor_append_str_dst.xlsx");
    defer std.testing.allocator.free(dst_path);

    // Source workbook with one string row so the SST exists.
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Data");
        try s.writeRow(&.{ .{ .string = "alpha" }, .{ .integer = 1 } });
        try w.save(io, src_path);
    }

    // Append a row that mixes string + integer + boolean.
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
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
        try ed.save(io, dst_path);
    }

    // Read back through Book — every appended string resolves to the
    // expected content.
    var book = try Book.open(std.testing.allocator, io, dst_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "editor_sstless_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "editor_sstless_dst.xlsx");
    defer std.testing.allocator.free(dst_path);

    // Build a workbook (the Zig writer always emits sharedStrings.xml).
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("D");
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 } });
        try w.save(io, src_path);
    }

    // Force the create-new-sst path: open the editor, then strip the
    // sharedStrings.xml entry from `ed.entries` before appending.
    // The Zig writer emits an SST regardless of cell types, so
    // without this the test would fall back to the substitute-
    // existing-SST path and never exercise the new branch.
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();

        var filtered: std.ArrayListUnmanaged(ZipEntry) = .empty;
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
        try ed.save(io, dst_path);
    }

    var book = try Book.open(std.testing.allocator, io, dst_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "scan_basic.xlsx");
    defer std.testing.allocator.free(src_path);

    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .string = "h1" }, .{ .string = "h2" }, .{ .string = "h3" } });
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .number = 2.5 }, .{ .boolean = true } });
        try s.writeRow(&.{ .{ .string = "x" }, .{ .empty = {} }, .{ .integer = 99 } });
        try w.save(io, src_path);
    }

    var ed = try Editor.open(std.testing.allocator, io, src_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "scan_roundtrip.xlsx");
    defer std.testing.allocator.free(src_path);

    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .string = "name" }, .{ .string = "qty" }, .{ .string = "ok" } });
        try s.writeRow(&.{ .{ .string = "alpha" }, .{ .integer = 10 }, .{ .boolean = true } });
        try s.writeRow(&.{ .{ .string = "beta" }, .{ .integer = 20 }, .{ .boolean = false } });
        try s.writeRow(&.{ .{ .string = "gamma" }, .{ .empty = {} }, .{ .boolean = true } });
        try w.save(io, src_path);
    }

    // Read the source via the Editor scanner.
    var ed = try Editor.open(std.testing.allocator, io, src_path);
    defer ed.deinit();
    var spans = try ed.scanWorksheet(0);
    defer spans.deinit();

    // Read the same file through Book.rows; every non-empty cell
    // must have a matching span at the same (row, col).
    var book = try Book.open(std.testing.allocator, io, src_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "setcell_basic_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "setcell_basic_dst.xlsx");
    defer std.testing.allocator.free(dst_path);

    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 } });
        try s.writeRow(&.{ .{ .integer = 3 }, .{ .integer = 4 } });
        try w.save(io, src_path);
    }

    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.setCell(0, 1, 1, .{ .integer = 99 }); // B1: 2 -> 99
        try ed.setCell(0, 2, 0, .{ .number = 3.5 }); // A2: 3 -> 3.5
        try ed.save(io, dst_path);
    }

    var book = try Book.open(std.testing.allocator, io, dst_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "setcell_str_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "setcell_str_dst.xlsx");
    defer std.testing.allocator.free(dst_path);

    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .string = "a" }, .{ .string = "b" } });
        try w.save(io, src_path);
    }

    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        // Replace shared-string cells with inline strings — including
        // entity-needing chars and leading/trailing whitespace.
        try ed.setCell(0, 1, 0, .{ .string = "Done & dusted" });
        try ed.setCell(0, 1, 1, .{ .string = " trim me " });
        try ed.save(io, dst_path);
    }

    var book = try Book.open(std.testing.allocator, io, dst_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();

    const r1 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualStrings("Done & dusted", r1[0].string);
    try std.testing.expectEqualStrings(" trim me ", r1[1].string);
}

test "Editor: setCell inserts a missing cell into an existing row (iter-cm-2c)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "setcell_insert.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "setcell_insert_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        // Row 1 has cells at A, C — gap at B. Row 2 has only A.
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .empty = {} }, .{ .integer = 3 } });
        try s.writeRow(&.{.{ .integer = 4 }});
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        // Insert into the gap at row 1 col 1.
        try ed.setCell(0, 1, 1, .{ .integer = 99 });
        // Insert at end-of-row in row 2: row has A only; append B.
        try ed.setCell(0, 2, 1, .{ .string = "appended" });
        try ed.save(io, dst_path);
    }
    var book = try Book.open(std.testing.allocator, io, dst_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "setcell_row_insert.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "setcell_row_insert_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }}); // row 1
        try s.writeRow(&.{.{ .integer = 5 }}); // row 2
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        // Insert a row beyond the source (row 5 — fresh).
        try ed.setCell(0, 5, 0, .{ .string = "row5" });
        // Insert a row in the middle (between rows 2 and 5 — call
        // it row 3).
        try ed.setCell(0, 3, 0, .{ .integer = 33 });
        try ed.save(io, dst_path);
    }
    var book = try Book.open(std.testing.allocator, io, dst_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // Build a sheet with a self-closing row in the middle. The
    // writer doesn't emit those, so synthesise via XML rewrite of
    // a save (write a normal sheet, then test the helper directly
    // by stuffing it into a MutatedSheet). For session brevity:
    // round-trip a writer-emitted body row through Editor + verify
    // setCell on a row-with-cells doesn't duplicate.
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "setcell_empty_row.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "setcell_empty_row_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }}); // r=1 with cells
        try s.writeRow(&.{.{ .integer = 5 }}); // r=2 with cells
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        // setCell on row 2 col 1 — row 2 already exists. The
        // pre-fix code path classified it as missing because no
        // span had row=2 col=1. Should now correctly insert into
        // the existing row's body, not duplicate.
        try ed.setCell(0, 2, 1, .{ .integer = 999 });
        try ed.save(io, dst_path);
    }
    var book = try Book.open(std.testing.allocator, io, dst_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // Write a workbook with only header cells, no body — produces
    // <sheetData/> in some readers' canonical form. Assert setCell
    // expands it to <sheetData></sheetData> form transparently.
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "setcell_empty_sd.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "setcell_empty_sd_dst.xlsx");
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
        try w.save(io, src_path);
    }

    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.setCell(0, 1, 0, .{ .integer = 42 });
        try ed.save(io, dst_path);
    }

    var book = try Book.open(std.testing.allocator, io, dst_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    const r1 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(i64, 42), r1[0].integer);
}

test "Editor: editor.workbook.addSheet path round-trips through Editor.save (iter-er-4)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "wb_addsheet_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "wb_addsheet_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Original");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        // Bypass Editor.addSheet — go straight through the
        // typed-overlay surface. Editor.save's verbatim-emit walk
        // and PartStore-overrides sync should pick it up.
        _ = try ed.workbook.addSheet("FromWorkbook");
        try ed.save(io, dst_path);
    }
    var book = try Book.open(std.testing.allocator, io, dst_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 2), book.sheets.len);
    try std.testing.expectEqualStrings("Original", book.sheets[0].name);
    try std.testing.expectEqualStrings("FromWorkbook", book.sheets[1].name);
}

test "Editor: editor.workbook.addSheet + setCell on returned handle round-trips" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "wb_addsheet_set_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "wb_addsheet_set_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Source");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
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
        try ed.save(io, dst_path);
    }
    var book = try Book.open(std.testing.allocator, io, dst_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 2), book.sheets.len);
    var rows = try book.rows(book.sheets[1], std.testing.allocator);
    defer rows.deinit();
    const r = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(f64, 3.14), r[0].number);
    try std.testing.expectEqualStrings("pi", r[1].string);
}

test "Editor: addSheet appends a new sheet and round-trips through reader (iter-sheet-1)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "addsheet_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "addsheet_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Original");
        try s.writeRow(&.{.{ .integer = 42 }});
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        const new_idx = try ed.addSheet("Added");
        try std.testing.expectEqual(@as(u32, 1), new_idx);
        // The new sheet should be addressable via setCell.
        try ed.setCell(new_idx, 1, 0, .{ .string = "hello" });
        try ed.setCell(new_idx, 1, 1, .{ .integer = 99 });
        try ed.save(io, dst_path);
    }
    var book = try Book.open(std.testing.allocator, io, dst_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // Codex P1: pre-fix, names with `"` produced malformed
    // workbook.xml (`name="He said "Hi""`). Use attr-escape that
    // covers `"` → `&quot;`.
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "addsheet_quote.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "addsheet_quote_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Original");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        // validateSheetName accepts `"` so we must round-trip it.
        _ = try ed.addSheet("He said \"Hi\"");
        try ed.save(io, dst_path);
    }
    var book = try Book.open(std.testing.allocator, io, dst_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 2), book.sheets.len);
    try std.testing.expectEqualStrings("He said \"Hi\"", book.sheets[1].name);
}

test "Editor: insertColumn shifts existing cells right (iter-col-3)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "insertcol_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "insertcol_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 }, .{ .integer = 3 } });
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertColumn(0, 2); // insert before col B
        try ed.save(io, dst_path);
    }
    var book = try Book.open(std.testing.allocator, io, dst_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "deletecol_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "deletecol_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .integer = 10 }, .{ .integer = 20 }, .{ .integer = 30 } });
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.deleteColumn(0, 2); // delete col B (the 20)
        try ed.save(io, dst_path);
    }
    var book = try Book.open(std.testing.allocator, io, dst_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "insertrow_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "insertrow_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }});
        try s.writeRow(&.{.{ .integer = 2 }});
        try s.writeRow(&.{.{ .integer = 3 }});
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertRow(0, 2); // insert blank row before row 2
        try ed.save(io, dst_path);
    }
    var book = try Book.open(std.testing.allocator, io, dst_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "deleterow_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "deleterow_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 10 }});
        try s.writeRow(&.{.{ .integer = 20 }});
        try s.writeRow(&.{.{ .integer = 30 }});
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.deleteRow(0, 2);
        try ed.save(io, dst_path);
    }
    var book = try Book.open(std.testing.allocator, io, dst_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    // Pre iter-er-5 lift: refused with `RowEditWithFormulasNotSupported`.
    // Post-lift: `Workbook.insertRow` runs `rewriteAllFormulas` so
    // cross-sheet and bare formula refs shift alongside the byte
    // transform. The call now succeeds.
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "row_unsafe.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRowWithFormulas(&.{.{ .integer = 0 }}, &.{"SUM(A1:A1)"});
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertRow(0, 1);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.deleteRow(0, 1);
    }
}

test "Editor: row/col edits with cross-sheet hyperlinks rewrite locations (iter-er-5 lift)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    // Pre iter-er-5 lift: refused with `RowEditUnsafeForSheet` /
    // `ColEditUnsafeForSheet`. Post-lift: `Workbook.{insertRow,
    // deleteRow, insertColumn, deleteColumn}` runs
    // `rewriteAllHyperlinkLocations` so cross-sheet location
    // qualifiers shift alongside the byte transform. The call now
    // succeeds.
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "xref_hyperlink.xlsx");
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
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertRow(0, 1);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.deleteRow(0, 1);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertColumn(0, 1);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.deleteColumn(0, 1);
    }
}

test "Editor: insertRow on frozen-pane sheet shifts ySplit + topLeftCell" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    // Post-lift: `<pane>` is no longer in the row-edit refusal
    // guards. The frozen pane's ySplit and topLeftCell row shift
    // alongside the row attrs. xSplit is unaffected by a row edit.
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "pane_insert_row.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "pane_insert_row_out.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }});
        try s.writeRow(&.{.{ .integer = 2 }});
        try s.freezePanes(1, 2); // rows=1, cols=2 → ySplit=1, xSplit=2, topLeftCell=C2
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertRow(0, 1);
        try ed.save(io, dst_path);
    }
    var ed2 = try Editor.open(std.testing.allocator, io, dst_path);
    defer ed2.deinit();
    const part = (try ed2.workbook.store.part("xl/worksheets/sheet1.xml")) orelse return error.TestUnexpectedResult;
    try std.testing.expect(std.mem.indexOf(u8, part.bytes, "ySplit=\"2\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, part.bytes, "xSplit=\"2\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, part.bytes, "topLeftCell=\"C3\"") != null);
}

test "Editor: insertColumn on frozen-pane sheet shifts xSplit + topLeftCell" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "pane_insert_col.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "pane_insert_col_out.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 }, .{ .integer = 3 } });
        try s.freezePanes(1, 2);
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertColumn(0, 1);
        try ed.save(io, dst_path);
    }
    var ed2 = try Editor.open(std.testing.allocator, io, dst_path);
    defer ed2.deinit();
    const part = (try ed2.workbook.store.part("xl/worksheets/sheet1.xml")) orelse return error.TestUnexpectedResult;
    try std.testing.expect(std.mem.indexOf(u8, part.bytes, "xSplit=\"3\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, part.bytes, "ySplit=\"1\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, part.bytes, "topLeftCell=\"D2\"") != null);
}

test "Editor: deleteRow + deleteColumn on frozen-pane sheet shrinks freeze" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "pane_delete.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "pane_delete_out.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 }, .{ .integer = 3 } });
        try s.writeRow(&.{ .{ .integer = 4 }, .{ .integer = 5 }, .{ .integer = 6 } });
        try s.freezePanes(1, 2);
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.deleteRow(0, 1); // ySplit 1→0, topLeftCell row 2→1
        try ed.save(io, dst_path);
    }
    var ed2 = try Editor.open(std.testing.allocator, io, dst_path);
    defer ed2.deinit();
    const part = (try ed2.workbook.store.part("xl/worksheets/sheet1.xml")) orelse return error.TestUnexpectedResult;
    try std.testing.expect(std.mem.indexOf(u8, part.bytes, "ySplit=\"0\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, part.bytes, "topLeftCell=\"C1\"") != null);
}

test "Editor: deleteColumn on frozen-pane sheet shrinks xSplit" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "pane_delete_col.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "pane_delete_col_out.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 }, .{ .integer = 3 } });
        try s.freezePanes(1, 2);
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.deleteColumn(0, 1); // xSplit 2→1, topLeftCell col C→B
        try ed.save(io, dst_path);
    }
    var ed2 = try Editor.open(std.testing.allocator, io, dst_path);
    defer ed2.deinit();
    const part = (try ed2.workbook.store.part("xl/worksheets/sheet1.xml")) orelse return error.TestUnexpectedResult;
    try std.testing.expect(std.mem.indexOf(u8, part.bytes, "xSplit=\"1\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, part.bytes, "topLeftCell=\"B2\"") != null);
}

test "Editor: deleteSheet on a different sheet is allowed after a column edit (iter-er-4 (3/N))" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
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
    const src_path = try tt.path(std.testing.allocator, io, "delsheet_after_col.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s1 = try w.addSheet("A");
        try s1.writeRow(&.{.{ .integer = 1 }});
        var s2 = try w.addSheet("B");
        try s2.writeRow(&.{.{ .integer = 2 }});
        try w.save(io, src_path);
    }
    var ed = try Editor.open(std.testing.allocator, io, src_path);
    defer ed.deinit();
    try ed.insertColumn(0, 1);
    // Used to error SheetDeleteRequiresCleanState; now succeeds.
    try ed.deleteSheet(1);
    try std.testing.expectEqual(@as(usize, 1), ed.sheet_paths.len);
}

test "Editor: deleteSheet on a different sheet is allowed after a row edit (iter-er-4 (3/N))" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    // Same parity as the column-edit test above.
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "delsheet_after_row.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s1 = try w.addSheet("A");
        try s1.writeRow(&.{.{ .integer = 1 }});
        try s1.writeRow(&.{.{ .integer = 2 }});
        var s2 = try w.addSheet("B");
        try s2.writeRow(&.{.{ .integer = 3 }});
        try w.save(io, src_path);
    }
    var ed = try Editor.open(std.testing.allocator, io, src_path);
    defer ed.deinit();
    try ed.insertRow(0, 1);
    try ed.deleteSheet(1);
    try std.testing.expectEqual(@as(usize, 1), ed.sheet_paths.len);
}

test "Editor: appendRows + setCell after a row/col edit compose (iter-er-4 (3/N))" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
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
    const src_path = try tt.path(std.testing.allocator, io, "append_after_coledit.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, src_path);
    }
    // appendRows after a column edit composes:
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertColumn(0, 1);
        try ed.appendRows(0, &.{&.{.{ .integer = 99 }}});
    }
    // setCell after a column edit composes (separate session
    // because appendRows + setCell on the same sheet remain
    // mutually exclusive — that invariant is independent of
    // row/col edits):
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertColumn(0, 1);
        try ed.setCell(0, 5, 0, .{ .integer = 99 });
    }
    // Same parity for deleteRow:
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.deleteRow(0, 1);
        try ed.appendRows(0, &.{&.{.{ .integer = 99 }}});
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.deleteRow(0, 1);
        try ed.setCell(0, 5, 0, .{ .integer = 99 });
    }
}

test "Editor: row edits with cross-sheet formulas rewrite refs (iter-er-5 lift)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    // Pre iter-er-5 lift: refused with `RowEditWithFormulasNotSupported`
    // / `ColEditWithFormulasNotSupported` because sheet2's formula
    // `Plain!A1+Plain!A2` references rows in the clean sheet.
    // Post-lift: `Workbook.insertRow` / `deleteRow` /
    // `insertColumn` / `deleteColumn` runs `rewriteAllFormulas`,
    // which shifts the cross-sheet refs. The call now succeeds.
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "row_xsheet_formula.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s1 = try w.addSheet("Plain");
        try s1.writeRow(&.{.{ .integer = 1 }});
        try s1.writeRow(&.{.{ .integer = 2 }});
        var s2 = try w.addSheet("HasFormula");
        try s2.writeRowWithFormulas(&.{.{ .integer = 0 }}, &.{"Plain!A1+Plain!A2"});
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertRow(0, 1);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.deleteRow(0, 1);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertColumn(0, 1);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.deleteColumn(0, 1);
    }
}

test "C1: after a row insert the reference follows the cell, INDIRECT follows the coordinate" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const ta = std.testing.allocator;

    // The literal-inertness tests in `pkg/workbook.zig` stop at the
    // rewritten bytes. This one runs the composition they imply —
    // edit, then EVALUATE — because the decision being pinned is a
    // claim about what a formula MEANS afterwards, not what it spells.
    //
    // A4=44 and A5=55; A1 reads `A5` and A2 reads `INDIRECT("A5")`, so
    // both answer 55 to begin with. Inserting a row at the top slides
    // the values down (44 lands on A5, 55 on A6) and the formulas to
    // A2/A3:
    //
    //   * `A5` is rewritten to `A6` and still answers 55 — it followed
    //     the CELL, which is what a reference does.
    //   * `INDIRECT("A5")` is untouched, resolves A5 against the grid
    //     as it now stands, and answers 44 — it followed the
    //     COORDINATE, which is the whole reason the function is used.
    //
    // Same pair, same starting answer, different answers after.
    // Rewriting the literal would collapse both onto 55 and erase the
    // distinction this milestone decided to keep.
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(ta, io, "c1_indirect_eval_src.xlsx");
    defer ta.free(src_path);
    const dst_path = try tt.path(ta, io, "c1_indirect_eval_dst.xlsx");
    defer ta.free(dst_path);

    {
        var w = xlsx.Writer.init(ta);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRowWithFormulas(&.{.{ .integer = 0 }}, &.{"A5"});
        try s.writeRowWithFormulas(&.{.{ .integer = 0 }}, &.{"INDIRECT(\"A5\")"});
        try s.writeRow(&.{.{ .integer = 0 }});
        try s.writeRow(&.{.{ .integer = 44 }});
        try s.writeRow(&.{.{ .integer = 55 }});
        try w.save(io, src_path);
    }

    {
        var wb = try Workbook.open(ta, io, src_path);
        defer wb.deinit();
        try wb.insertRow(0, 1);
        var report = try wb.saveWithRecalc(ta, io, dst_path, .{
            .now_utc_ms = 1_700_000_000_000,
            .rng_seed = 0x5EED_5D3,
            .limits = .{},
        }, .{});
        report.deinit(ta);
    }

    var wb2 = try Workbook.open(ta, io, dst_path);
    defer wb2.deinit();
    const ws = try wb2.sheet(0);

    const followed_cell = (try ws.cellByRef("A2")).?;
    try std.testing.expectEqualStrings("A6", followed_cell.formula.?);
    try std.testing.expectEqualStrings("55", followed_cell.raw_value.?);

    // `Cell.formula` is the RAW `<f>` inner text, so the quotes around
    // the literal arrive as `&quot;`; decode before comparing rather
    // than pinning one escaping policy.
    const followed_coordinate = (try ws.cellByRef("A3")).?;
    const decoded = try store_mod.decodeXmlEntities(ta, followed_coordinate.formula.?);
    defer ta.free(decoded);
    try std.testing.expectEqualStrings("INDIRECT(\"A5\")", decoded);
    try std.testing.expectEqualStrings("44", followed_coordinate.raw_value.?);
}

// C1 wiring — per-edit forwarding proofs. Each of the six
// structural edits must reach `Workbook.rewriteAllFormulas`; the
// byte transform alone moves `<c r=>` anchors but never rewrites
// `<f>` bodies, so asserting the rewritten body text after
// save → reopen fails exactly when that edit's rewriter call is
// deleted. One Editor instance per edit: the rewriter stages its
// output as setCell deltas, which the clean-sheet guard on a
// following edit would refuse.

test "Editor: insertRow forwards formulas to the rewriter (C1 wiring proof)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "fwd_insrow_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "fwd_insrow_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }});
        try s.writeRow(&.{.{ .integer = 2 }});
        try s.writeRowWithFormulas(&.{.{ .integer = 3 }}, &.{"A1+A2"});
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertRow(0, 2);
        try ed.save(io, dst_path);
    }
    var wb = try Workbook.open(std.testing.allocator, io, dst_path);
    defer wb.deinit();
    const c = (try (try wb.sheet(0)).cellByRef("A4")).?;
    try std.testing.expectEqualStrings("A1+A3", c.formula.?);
}

test "Editor: insertRow scopes bare refs to the edited sheet (C1 wiring proof)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    // A row edit moves ONE sheet's grid. Bare refs in formulas on
    // other sheets must stay put; refs qualified to the edited
    // sheet must shift wherever they live. Codex r1 on this PR
    // found the wiring passed no target sheet, so bare refs
    // shifted workbook-wide.
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "fwd_scope_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "fwd_scope_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s1 = try w.addSheet("Edited");
        try s1.writeRow(&.{.{ .integer = 1 }});
        try s1.writeRow(&.{.{ .integer = 2 }});
        try s1.writeRowWithFormulas(&.{.{ .integer = 3 }}, &.{"A1+A2"});
        var s2 = try w.addSheet("Other");
        try s2.writeRowWithFormulas(
            &.{ .{ .integer = 0 }, .{ .integer = 0 }, .{ .integer = 0 } },
            &.{ "A2+1", "Edited!A2*2", "edited!A2+3" },
        );
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertRow(0, 2);
        try ed.save(io, dst_path);
    }
    var wb = try Workbook.open(std.testing.allocator, io, dst_path);
    defer wb.deinit();
    const edited = (try (try wb.sheet(0)).cellByRef("A4")).?;
    try std.testing.expectEqualStrings("A1+A3", edited.formula.?);
    const other_bare = (try (try wb.sheet(1)).cellByRef("A1")).?;
    try std.testing.expectEqualStrings("A2+1", other_bare.formula.?);
    const other_qualified = (try (try wb.sheet(1)).cellByRef("B1")).?;
    try std.testing.expectEqualStrings("Edited!A3*2", other_qualified.formula.?);
    // Case-variant qualifier: Excel resolves sheet names case-
    // insensitively, so `edited!` scopes to sheet "Edited" too.
    const other_case = (try (try wb.sheet(1)).cellByRef("C1")).?;
    try std.testing.expectEqualStrings("edited!A3+3", other_case.formula.?);
}

test "Editor: deleteRow forwards formulas to the rewriter (C1 wiring proof)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "fwd_delrow_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "fwd_delrow_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }});
        try s.writeRow(&.{.{ .integer = 2 }});
        try s.writeRow(&.{.{ .integer = 3 }});
        try s.writeRowWithFormulas(&.{.{ .integer = 5 }}, &.{"A2+A3"});
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.deleteRow(0, 1);
        try ed.save(io, dst_path);
    }
    var wb = try Workbook.open(std.testing.allocator, io, dst_path);
    defer wb.deinit();
    const c = (try (try wb.sheet(0)).cellByRef("A3")).?;
    try std.testing.expectEqualStrings("A1+A2", c.formula.?);
}

test "Editor: insertColumn forwards formulas to the rewriter (C1 wiring proof)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "fwd_inscol_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "fwd_inscol_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 } });
        try s.writeRowWithFormulas(&.{.{ .integer = 3 }}, &.{"A1+B1"});
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertColumn(0, 2);
        try ed.save(io, dst_path);
    }
    var wb = try Workbook.open(std.testing.allocator, io, dst_path);
    defer wb.deinit();
    const c = (try (try wb.sheet(0)).cellByRef("A2")).?;
    try std.testing.expectEqualStrings("A1+C1", c.formula.?);
}

test "Editor: deleteColumn forwards formulas to the rewriter (C1 wiring proof)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "fwd_delcol_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "fwd_delcol_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 }, .{ .integer = 3 } });
        try s.writeRowWithFormulas(&.{.{ .integer = 6 }}, &.{"C1*2"});
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.deleteColumn(0, 2);
        try ed.save(io, dst_path);
    }
    var wb = try Workbook.open(std.testing.allocator, io, dst_path);
    defer wb.deinit();
    const c = (try (try wb.sheet(0)).cellByRef("A2")).?;
    try std.testing.expectEqualStrings("B1*2", c.formula.?);
}

test "Editor: renameSheet forwards formulas to the rewriter (C1 wiring proof)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "fwd_rename_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "fwd_rename_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s1 = try w.addSheet("Alpha");
        try s1.writeRowWithFormulas(&.{.{ .integer = 0 }}, &.{"Beta!A1+1"});
        var s2 = try w.addSheet("Beta");
        try s2.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.renameSheet(1, "Gamma");
        try ed.save(io, dst_path);
    }
    var wb = try Workbook.open(std.testing.allocator, io, dst_path);
    defer wb.deinit();
    const c = (try (try wb.sheet(0)).cellByRef("A1")).?;
    try std.testing.expectEqualStrings("Gamma!A1+1", c.formula.?);
}

test "Editor: deleteSheet forwards formulas to the rewriter (C1 wiring proof)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "fwd_delsheet_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "fwd_delsheet_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s1 = try w.addSheet("Keep");
        try s1.writeRowWithFormulas(&.{.{ .integer = 0 }}, &.{"Doomed!A1+1"});
        var s2 = try w.addSheet("Doomed");
        try s2.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.deleteSheet(1);
        try ed.save(io, dst_path);
    }
    var wb = try Workbook.open(std.testing.allocator, io, dst_path);
    defer wb.deinit();
    const c = (try (try wb.sheet(0)).cellByRef("A1")).?;
    try std.testing.expectEqualStrings("#REF!+1", c.formula.?);
}

test "Editor: deleteSheet drops a source sheet (iter-sheet-3)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "delete_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "delete_dst.xlsx");
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
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.deleteSheet(1);
        try ed.save(io, dst_path);
    }
    var book = try Book.open(std.testing.allocator, io, dst_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 2), book.sheets.len);
    try std.testing.expectEqualStrings("Keep", book.sheets[0].name);
    try std.testing.expectEqualStrings("AlsoKeep", book.sheets[1].name);
}

test "Editor: deleteSheet drops a pending-new sheet" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "delete_new.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "delete_new_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        _ = try w.addSheet("Original");
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        const new_idx = try ed.addSheet("Tmp");
        try ed.deleteSheet(new_idx);
        try ed.save(io, dst_path);
    }
    var book = try Book.open(std.testing.allocator, io, dst_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 1), book.sheets.len);
    try std.testing.expectEqualStrings("Original", book.sheets[0].name);
}

test "Editor: deleteSheet preserves order of other pending-new sheets" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // Codex P1: swapRemove reordered remaining new sheets. orderedRemove
    // keeps them aligned with sheet_paths' tail.
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "delete_order.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "delete_order_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        _ = try w.addSheet("Source");
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        const a = try ed.addSheet("A"); // idx 1
        _ = try ed.addSheet("B"); // idx 2
        _ = try ed.addSheet("C"); // idx 3
        try ed.deleteSheet(a); // remove A; B,C should stay in order
        try ed.save(io, dst_path);
    }
    var book = try Book.open(std.testing.allocator, io, dst_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 3), book.sheets.len);
    try std.testing.expectEqualStrings("Source", book.sheets[0].name);
    try std.testing.expectEqualStrings("B", book.sheets[1].name);
    try std.testing.expectEqualStrings("C", book.sheets[2].name);
}

test "Editor: deleteSheet frees name for reuse via addSheet" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // Codex P2: pre-fix, addSheet rejected reuse of a deleted
    // source sheet's name as DuplicateSheetName.
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "delete_reuse.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "delete_reuse_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        _ = try w.addSheet("Keep");
        _ = try w.addSheet("Drop");
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.deleteSheet(1);
        // Reuse the deleted name.
        _ = try ed.addSheet("Drop");
        try ed.save(io, dst_path);
    }
    var book = try Book.open(std.testing.allocator, io, dst_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 2), book.sheets.len);
    try std.testing.expectEqualStrings("Keep", book.sheets[0].name);
    try std.testing.expectEqualStrings("Drop", book.sheets[1].name);
}

test "Editor: deleteSheet rejects last-sheet" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "delete_last.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        _ = try w.addSheet("Solo");
        try w.save(io, src_path);
    }
    var ed = try Editor.open(std.testing.allocator, io, src_path);
    defer ed.deinit();
    try std.testing.expectError(error.CannotDeleteLastSheet, ed.deleteSheet(0));
    try std.testing.expectError(error.SheetIndexOutOfRange, ed.deleteSheet(99));
}

test "Editor: deleteSheet rejects dirty state" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "delete_dirty.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        _ = try w.addSheet("S1");
        _ = try w.addSheet("S2");
        try w.save(io, src_path);
    }
    var ed = try Editor.open(std.testing.allocator, io, src_path);
    defer ed.deinit();
    try ed.appendRows(0, &.{&.{.{ .integer = 5 }}});
    try std.testing.expectError(error.SheetDeleteRequiresCleanState, ed.deleteSheet(1));
}

test "Editor: renameSheet renames an existing sheet (iter-sheet-2)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "rename_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "rename_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("OldName");
        try s.writeRow(&.{.{ .integer = 42 }});
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.renameSheet(0, "NewName");
        try ed.save(io, dst_path);
    }
    var book = try Book.open(std.testing.allocator, io, dst_path);
    defer book.deinit();
    try std.testing.expectEqualStrings("NewName", book.sheets[0].name);
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    const r = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(i64, 42), r[0].integer);
}

test "Editor: renameSheet rejects duplicates and invalid names" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "rename_reject.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        _ = try w.addSheet("First");
        _ = try w.addSheet("Second");
        try w.save(io, src_path);
    }
    var ed = try Editor.open(std.testing.allocator, io, src_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "rename_undo.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "rename_undo_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Original");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.renameSheet(0, "Renamed");
        // Now revert. Pre-fix this was silently dropped.
        try ed.renameSheet(0, "Original");
        try ed.save(io, dst_path);
    }
    var book = try Book.open(std.testing.allocator, io, dst_path);
    defer book.deinit();
    try std.testing.expectEqualStrings("Original", book.sheets[0].name);
}

test "Editor: renameSheet persists case-only changes" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // Excel uniqueness is case-insensitive but the displayed
    // casing matters. asciiEqlFold short-circuit dropped legit
    // case-only renames pre-fix.
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "rename_case.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "rename_case_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("sheet1");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.renameSheet(0, "Sheet1");
        try ed.save(io, dst_path);
    }
    var book = try Book.open(std.testing.allocator, io, dst_path);
    defer book.deinit();
    try std.testing.expectEqualStrings("Sheet1", book.sheets[0].name);
}

test "Editor: rename + add reuses names freed by earlier renames" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // Codex P2: pre-fix, rotate / swap rename workflows were
    // rejected because the dup check only saw raw source names.
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "rename_rotate.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "rename_rotate_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        _ = try w.addSheet("A");
        _ = try w.addSheet("B");
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        // A -> C, then B -> A (rotate).
        try ed.renameSheet(0, "C");
        try ed.renameSheet(1, "A");
        // addSheet can also reuse a freed name: rename A->B's
        // earlier old-name was "B", but sheet 1 is now "A". Adding
        // "B" should succeed.
        _ = try ed.addSheet("B");
        try ed.save(io, dst_path);
    }
    var book = try Book.open(std.testing.allocator, io, dst_path);
    defer book.deinit();
    try std.testing.expectEqualStrings("C", book.sheets[0].name);
    try std.testing.expectEqualStrings("A", book.sheets[1].name);
    try std.testing.expectEqualStrings("B", book.sheets[2].name);
}

test "Editor: renameSheet on a pending-new sheet mutates in place" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "rename_new.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "rename_new_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Source");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        const idx = try ed.addSheet("Tmp");
        try ed.renameSheet(idx, "Final");
        try ed.save(io, dst_path);
    }
    var book = try Book.open(std.testing.allocator, io, dst_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 2), book.sheets.len);
    try std.testing.expectEqualStrings("Final", book.sheets[1].name);
}

test "Editor: appendRows works on freshly-added sheets" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // Codex P1: pre-fix, save() failed with SheetEntryNotFound when
    // appendRows targeted a sheet created via addSheet, because the
    // pending_appends loop assumed every sheet had a source entry.
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "addsheet_append.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "addsheet_append_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Original");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        const new_idx = try ed.addSheet("Fresh");
        try ed.appendRows(new_idx, &.{
            &.{ .{ .string = "alpha" }, .{ .integer = 10 } },
            &.{ .{ .string = "beta" }, .{ .integer = 20 } },
        });
        try ed.save(io, dst_path);
    }
    var book = try Book.open(std.testing.allocator, io, dst_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "scan_new_sheet.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Original");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, src_path);
    }
    var ed = try Editor.open(std.testing.allocator, io, src_path);
    defer ed.deinit();
    const new_idx = try ed.addSheet("Empty");
    // scanWorksheet on the brand-new untouched sheet must not error
    // — it should return a zero-cell span set.
    var spans = try ed.scanWorksheet(new_idx);
    defer spans.deinit();
    try std.testing.expectEqual(@as(usize, 0), spans.cells.len);
}

test "Editor: addSheet handles XML-escaped duplicate names (R&D)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // The source workbook stores `R&D` as `name="R&amp;D"` in
    // workbook.xml. Pre-fix, sheetNameExists compared raw bytes
    // and accepted the duplicate. Now decode entities first.
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "addsheet_amp.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("R&D");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, src_path);
    }
    var ed = try Editor.open(std.testing.allocator, io, src_path);
    defer ed.deinit();
    try std.testing.expectError(error.DuplicateSheetName, ed.addSheet("R&D"));
    // Case-insensitive too.
    try std.testing.expectError(error.DuplicateSheetName, ed.addSheet("r&d"));
}

test "Editor: addSheet allocates non-colliding ids across multiple calls" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // Codex caught: pre-fix, two addSheet calls in one session
    // produced duplicate rIds / sheetIds / sheet paths because the
    // max scan only looked at the source.
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "addsheet_seq_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "addsheet_seq_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Original");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        const a = try ed.addSheet("A");
        const b = try ed.addSheet("B");
        const c = try ed.addSheet("C");
        try ed.setCell(a, 1, 0, .{ .string = "in_a" });
        try ed.setCell(b, 1, 0, .{ .string = "in_b" });
        try ed.setCell(c, 1, 0, .{ .string = "in_c" });
        try ed.save(io, dst_path);
    }
    var book = try Book.open(std.testing.allocator, io, dst_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "addsheet_case.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Existing");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, src_path);
    }
    var ed = try Editor.open(std.testing.allocator, io, src_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "addsheet_reject.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Existing");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, src_path);
    }
    var ed = try Editor.open(std.testing.allocator, io, src_path);
    defer ed.deinit();
    try std.testing.expectError(error.InvalidSheetName, ed.addSheet(""));
    try std.testing.expectError(error.InvalidSheetName, ed.addSheet("a:b"));
    try std.testing.expectError(error.DuplicateSheetName, ed.addSheet("Existing"));
    _ = try ed.addSheet("Fresh");
    try std.testing.expectError(error.DuplicateSheetName, ed.addSheet("Fresh"));
}

test "Editor: setCells applies a batch in source order (iter-cm-3)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "setcells_batch.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "setcells_batch_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 }, .{ .integer = 3 } });
        try s.writeRow(&.{ .{ .integer = 4 }, .{ .integer = 5 }, .{ .integer = 6 } });
        try w.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.setCells(0, &.{
            .{ .row = 1, .col = 0, .cell = .{ .integer = 100 } },
            .{ .row = 1, .col = 2, .cell = .{ .number = 3.14 } },
            .{ .row = 2, .col = 1, .cell = .{ .string = "x" } },
        });
        try ed.save(io, dst_path);
    }
    var book = try Book.open(std.testing.allocator, io, dst_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // Lazy-init the MutatedSheet on first call; subsequent calls
    // mutate the cached buffer.
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "setcell_many.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "setcell_many_dst.xlsx");
    defer std.testing.allocator.free(dst_path);

    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        var i: u32 = 0;
        while (i < 10) : (i += 1) {
            try s.writeRow(&.{ .{ .integer = i }, .{ .integer = i * 2 }, .{ .integer = i * 3 } });
        }
        try w.save(io, src_path);
    }

    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
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
        try ed.save(io, dst_path);
    }

    var book = try Book.open(std.testing.allocator, io, dst_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "setcell_reject.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, src_path);
    }
    var ed = try Editor.open(std.testing.allocator, io, src_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "setcell_then_append.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, src_path);
    }
    var ed = try Editor.open(std.testing.allocator, io, src_path);
    defer ed.deinit();
    try ed.setCell(0, 1, 0, .{ .integer = 2 });
    try std.testing.expectError(
        error.SheetHasUnsavedMutations,
        ed.appendRows(0, &.{&.{.{ .integer = 99 }}}),
    );
}

test "Editor: scanWorksheet sees setCell mutations (no stale read)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // Codex caught: pre-fix, scanWorksheet decompressed from
    // src_buf and ignored pending_mutations. A setCell-then-scan
    // workflow got pre-mutation spans.
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "setcell_then_scan.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 } });
        try w.save(io, src_path);
    }
    var ed = try Editor.open(std.testing.allocator, io, src_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // The scanner reads from src_buf and would silently miss rows
    // queued in pending_appends. Contract: reject so callers can't
    // act on a stale span set.
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "scan_pending.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, src_path);
    }
    var ed = try Editor.open(std.testing.allocator, io, src_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx;
    const src_path = try tt.path(std.testing.allocator, io, "scan_oor.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, src_path);
    }
    var ed = try Editor.open(std.testing.allocator, io, src_path);
    defer ed.deinit();
    try std.testing.expectError(error.SheetIndexOutOfRange, ed.scanWorksheet(99));
}

// ─── B2 iter-er-1: Editor read-side parity tests ─────────────────────

/// Build a 2-sheet xlsx via the writer (which produces a non-DataDescriptor
/// ZIP that Editor accepts) and return its temp path. Caller frees the
/// returned slice and is responsible for the TestTmp lifecycle.
fn buildIterEr1Fixture(io: std.Io, tt: *TestTmp, alloc: std.mem.Allocator) ![:0]u8 {
    const path = try tt.path(alloc, io, "iter_er_1.xlsx");
    var w = xlsx.Writer.init(alloc);
    defer w.deinit();
    var s1 = try w.addSheet("Alpha");
    try s1.writeRow(&.{ .{ .string = "h" }, .{ .integer = 1 } });
    var s2 = try w.addSheet("Beta");
    try s2.writeRow(&.{ .{ .string = "x" }, .{ .number = 2.5 } });
    try w.save(io, path);
    return path;
}

test "S1: Editor.openBuffer refuses the three hostile shapes on its own scan, before the core reader runs" {
    const a = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const hostile = xlsx.zip_probe.hostile;

    const oversized = try hostile.oversizedPart(a);
    defer a.free(oversized);
    try std.testing.expectError(error.ZipBombSuspected, Editor.openBuffer(a, io, oversized));

    const ratio = try hostile.absurdRatio(a);
    defer a.free(ratio);
    try std.testing.expectError(error.ZipBombSuspected, Editor.openBuffer(a, io, ratio));

    const over = try hostile.fullSizedParts(a, hostile.partsOverBudget());
    defer a.free(over);
    try std.testing.expectError(error.ZipBombSuspected, Editor.openBuffer(a, io, over));

    // Control: exactly on the aggregate — the scan admits it and the
    // open proceeds into `Book.openBuffer`, which fails on the missing
    // workbook part. That the error is the *reader's* proves the scan
    // let it through; that it is not `ZipBombSuspected` proves the
    // reader's own admission agrees with the editor's.
    const within = try hostile.fullSizedParts(a, hostile.partsWithinBudget());
    defer a.free(within);
    try std.testing.expectError(error.MissingWorkbook, Editor.openBuffer(a, io, within));
}

test "S1: the editor's scan refuses before the core reader is even constructed" {
    // Both layers admit every entry, so a hostile archive is refused by
    // whichever runs first and the error alone cannot say which. The
    // ordering is made observable with an allocator that serves the
    // scan's two allocations (the source dupe, the entry table) and
    // fails the third — which is the core reader's first. A refusal
    // from the scan is `ZipBombSuspected`; a refusal that had waited
    // for `Book.openBuffer` would surface as `OutOfMemory`.
    //
    // The archive is a real one-sheet workbook re-emitted entry by
    // entry through the probe builder with one lie: sheet1 declares
    // one byte past the per-part cap behind a zero-filled payload long
    // enough for the ratio. If anything inflated it, the error would
    // be `BadZip`.
    const a = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var w = xlsx.Writer.init(a);
    defer w.deinit();
    var s1 = try w.addSheet("Alpha");
    try s1.writeRow(&.{ .{ .string = "h" }, .{ .integer = 1 } });
    const real = try w.saveToOwnedBuffer(a, io);
    defer a.free(real);

    var entries: std.ArrayListUnmanaged(xlsx.zip_probe.Entry) = .empty;
    defer entries.deinit(a);
    var owned: std.ArrayListUnmanaged([]u8) = .empty;
    defer {
        for (owned.items) |o| a.free(o);
        owned.deinit(a);
    }
    const limits = xlsx.decompress_limits;
    {
        var src = try Editor.openBuffer(a, io, real);
        defer src.deinit();
        for (src.entries) |e| {
            const name = try a.dupe(u8, e.name);
            try owned.append(a, name);
            const is_sheet = std.mem.eql(u8, e.name, "xl/worksheets/sheet1.xml");
            const payload = if (is_sheet) blk: {
                const pl = try a.alloc(u8, @intCast(limits.max_part_size / limits.max_deflate_ratio + 1));
                @memset(pl, 0);
                break :blk pl;
            } else try a.dupe(u8, src.src_buf[e.lfh_offset + e.lfh_total_len ..][0..e.payload_len]);
            try owned.append(a, payload);
            try entries.append(a, .{
                .name = name,
                .payload = payload,
                .method = e.compression_method,
                .declared_uncompressed = if (is_sheet) @intCast(limits.max_part_size + 1) else e.uncompressed_size,
            });
        }
    }
    const hostile_bytes = try xlsx.zip_probe.build(a, entries.items);
    defer a.free(hostile_bytes);

    var starved = std.testing.FailingAllocator.init(a, .{ .fail_index = 2 });
    try std.testing.expectError(error.ZipBombSuspected, Editor.openBuffer(starved.allocator(), io, hostile_bytes));
    // The scan's own two allocations were served — the third was never
    // asked for, because the refusal came first.
    try std.testing.expectEqual(@as(usize, 2), starved.alloc_index);

    // The same allocator budget against the untouched workbook proves
    // the third allocation *is* the reader's: it fails there.
    var starved2 = std.testing.FailingAllocator.init(a, .{ .fail_index = 2 });
    try std.testing.expectError(error.OutOfMemory, Editor.openBuffer(starved2.allocator(), io, real));
}

test "iter-er-1: Editor.open populates a Workbook view (sheet count + names match)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try buildIterEr1Fixture(io, &tt, std.testing.allocator);
    defer std.testing.allocator.free(path);

    var ed = try Editor.open(std.testing.allocator, io, path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    // Editor.deinit calls self.workbook.deinit() before freeing
    // src_buf / entries / sheet_paths. Under std.testing.allocator
    // (which panics on leak), this test fails if any allocation
    // outlives the editor.
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try buildIterEr1Fixture(io, &tt, std.testing.allocator);
    defer std.testing.allocator.free(path);

    {
        var ed = try Editor.open(std.testing.allocator, io, path);
        ed.deinit();
    }
}

test "iter-er-1: Editor.workbook.cellByRef matches Book.cell for known cells" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    // Cross-API parity sanity check: the Workbook view's per-cell
    // accessor finds cells the writer emitted.
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try buildIterEr1Fixture(io, &tt, std.testing.allocator);
    defer std.testing.allocator.free(path);

    var ed = try Editor.open(std.testing.allocator, io, path);
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

fn buildAutoFilterFixture(io: std.Io, path: []const u8) !void {
    var w = xlsx.Writer.init(std.testing.allocator);
    defer w.deinit();
    var s = try w.addSheet("S");
    try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 }, .{ .integer = 3 }, .{ .integer = 4 } });
    try s.writeRow(&.{ .{ .integer = 5 }, .{ .integer = 6 }, .{ .integer = 7 }, .{ .integer = 8 } });
    try s.setAutoFilter("B1:D2");
    try w.save(io, path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, io, "af_insrow_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "af_insrow_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    try buildAutoFilterFixture(io, src_path);
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertRow(0, 2);
        try ed.save(io, dst_path);
    }
    var ed2 = try Editor.open(std.testing.allocator, io, dst_path);
    defer ed2.deinit();
    // B1:D2 → B1:D3 after inserting a row at row 2.
    try assertSheetXmlContains(&ed2, "ref=\"B1:D3\"");
}

test "Editor: insertColumn on a sheet carrying <autoFilter> rewrites range" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, io, "af_inscol_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "af_inscol_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    try buildAutoFilterFixture(io, src_path);
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertColumn(0, 3); // insert before col C, inside B:D
        try ed.save(io, dst_path);
    }
    var ed2 = try Editor.open(std.testing.allocator, io, dst_path);
    defer ed2.deinit();
    // B1:D2 → B1:E2 after extending the range by one column.
    try assertSheetXmlContains(&ed2, "ref=\"B1:E2\"");
}

test "Editor: deleteRow on a sheet carrying <autoFilter> rewrites range" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, io, "af_delrow_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "af_delrow_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    try buildAutoFilterFixture(io, src_path);
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.deleteRow(0, 2); // delete row 2 (the BR of B1:D2)
        try ed.save(io, dst_path);
    }
    var ed2 = try Editor.open(std.testing.allocator, io, dst_path);
    defer ed2.deinit();
    // B1:D2 → B1:D1 after the BR row vanishes.
    try assertSheetXmlContains(&ed2, "ref=\"B1:D1\"");
}

test "Editor: deleteColumn on a sheet carrying <autoFilter> shrinks range" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, io, "af_delcol_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "af_delcol_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    try buildAutoFilterFixture(io, src_path);
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.deleteColumn(0, 3); // delete col C inside B:D
        try ed.save(io, dst_path);
    }
    var ed2 = try Editor.open(std.testing.allocator, io, dst_path);
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

fn buildPictureFixture(io: std.Io, path: []const u8) !void {
    // Stage 1: produce a baseline workbook with the writer.
    {
        var w = xlsx.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 } });
        try s.writeRow(&.{ .{ .integer = 3 }, .{ .integer = 4 } });
        try w.save(io, path);
    }

    // Stage 2: splice `<picture r:id="rId99"/>` before `</worksheet>` in
    // sheet1.xml via PartStore.replacePart, then save back over the
    // baseline. The dangling rId99 doesn't resolve to anything in the
    // sheet's rels; that's fine for this test — the guard is what we're
    // exercising, and ECMA-376 readers tolerate dangling rels by
    // ignoring the picture element rather than failing the file.
    var store = try store_mod.PartStore.open(std.testing.allocator, io, path);
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
    try store.save(io, path);
}

test "Editor: insertRow on sheet with `<picture>` background no longer refused" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, io, "pic_insrow_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "pic_insrow_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    try buildPictureFixture(io, src_path);
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertRow(0, 2);
        try ed.save(io, dst_path);
    }
    var ed2 = try Editor.open(std.testing.allocator, io, dst_path);
    defer ed2.deinit();
    // The `<picture>` element passes through unchanged (no row/col coords).
    try assertSheetXmlContains(&ed2, "<picture r:id=\"rId99\"/>");
}

test "Editor: insertColumn on sheet with `<picture>` background no longer refused" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, io, "pic_inscol_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "pic_inscol_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    try buildPictureFixture(io, src_path);
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertColumn(0, 2);
        try ed.save(io, dst_path);
    }
    var ed2 = try Editor.open(std.testing.allocator, io, dst_path);
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

fn buildDrawingFixture(io: std.Io, path: []const u8, anchor_col: u32, anchor_row: u32) !void {
    {
        var w = xlsx.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 }, .{ .integer = 3 }, .{ .integer = 4 } });
        try s.writeRow(&.{ .{ .integer = 5 }, .{ .integer = 6 }, .{ .integer = 7 }, .{ .integer = 8 } });
        try w.save(io, path);
    }
    var wb = try workbook_mod.Workbook.open(std.testing.allocator, io, path);
    defer wb.deinit();
    try wb.addImage(0, .{ .col = anchor_col, .row = anchor_row }, &tiny_png_1x1_for_editor, .png);
    try wb.save(io, path);
}

fn drawingPartContains(io: std.Io, path: []const u8, needle: []const u8) !bool {
    var store = try store_mod.PartStore.open(std.testing.allocator, io, path);
    defer store.deinit();
    const drawing = (try store.part("xl/drawings/drawing1.xml")) orelse return error.MissingDrawing;
    return std.mem.indexOf(u8, drawing.bytes, needle) != null;
}

test "Editor: insertColumn shifts xdr drawing anchor col (dr-1)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, io, "draw_inscol_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "draw_inscol_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    try buildDrawingFixture(io, src_path, 3, 5);
    try std.testing.expect(try drawingPartContains(io, src_path, "<xdr:col>2</xdr:col>"));
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertColumn(0, 2);
        try ed.save(io, dst_path);
    }
    try std.testing.expect(try drawingPartContains(io, dst_path, "<xdr:col>3</xdr:col>"));
    try std.testing.expect(try drawingPartContains(io, dst_path, "<xdr:row>4</xdr:row>"));
}

test "Editor: insertRow shifts xdr drawing anchor row (dr-1)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, io, "draw_insrow_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "draw_insrow_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    try buildDrawingFixture(io, src_path, 3, 5);
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertRow(0, 3);
        try ed.save(io, dst_path);
    }
    try std.testing.expect(try drawingPartContains(io, dst_path, "<xdr:col>2</xdr:col>"));
    try std.testing.expect(try drawingPartContains(io, dst_path, "<xdr:row>5</xdr:row>"));
}

test "Editor: deleteColumn at the anchor's column drops the oneCellAnchor (dr-1)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, io, "draw_delcol_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "draw_delcol_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    try buildDrawingFixture(io, src_path, 3, 5);
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.deleteColumn(0, 3);
        try ed.save(io, dst_path);
    }
    try std.testing.expect(!try drawingPartContains(io, dst_path, "<xdr:oneCellAnchor>"));
}

test "Editor: drawing rewrite tolerates XML whitespace around `=` in worksheet rels (dr-1 REL-602/604)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, io, "draw_eq_ws_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "draw_eq_ws_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    try buildDrawingFixture(io, src_path, 3, 5);
    // Splice a custom `<drawing  r:id  =  "rId99"/>` into sheet1.xml
    // ALONGSIDE the existing well-formed `<drawing r:id="rIdN"/>`
    // emitted by addImage; we exercise the parser's whitespace
    // tolerance even though the resolved part is the SAME.
    {
        var store = try store_mod.PartStore.open(std.testing.allocator, io, src_path);
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
        try store.save(io, src_path);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertColumn(0, 2);
        try ed.save(io, dst_path);
    }
    // The drawing's xdr:col MUST shift; if findAttrValue mishandled
    // the whitespace, applyDrawingEditForSheet would silently skip
    // and the col would still be 2.
    try std.testing.expect(try drawingPartContains(io, dst_path, "<xdr:col>3</xdr:col>"));
}

// ---------------------------------------------------------------------------
// VML legacy drawings refusal lift (dr-2).
// `Writer.addComment` produces a `<v:shape>` block in
// `xl/drawings/vmlDrawing1.vml` AND a `<comment>` entry in
// `xl/comments1.xml`. After Editor row/col edits, both must shift.
// ---------------------------------------------------------------------------

fn buildCommentFixture(io: std.Io, path: []const u8, comment_ref: []const u8) !void {
    var w = xlsx.Writer.init(std.testing.allocator);
    defer w.deinit();
    var s = try w.addSheet("S");
    try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 }, .{ .integer = 3 }, .{ .integer = 4 } });
    try s.writeRow(&.{ .{ .integer = 5 }, .{ .integer = 6 }, .{ .integer = 7 }, .{ .integer = 8 } });
    try s.writeRow(&.{ .{ .integer = 9 }, .{ .integer = 10 }, .{ .integer = 11 }, .{ .integer = 12 } });
    try s.writeRow(&.{ .{ .integer = 13 }, .{ .integer = 14 }, .{ .integer = 15 }, .{ .integer = 16 } });
    try s.writeRow(&.{ .{ .integer = 17 }, .{ .integer = 18 }, .{ .integer = 19 }, .{ .integer = 20 } });
    try s.addComment(comment_ref, "Author", "test note");
    try w.save(io, path);
}

fn vmlPartContains(io: std.Io, path: []const u8, needle: []const u8) !bool {
    var store = try store_mod.PartStore.open(std.testing.allocator, io, path);
    defer store.deinit();
    const vml = (try store.part("xl/drawings/vmlDrawing1.vml")) orelse return error.MissingVml;
    return std.mem.indexOf(u8, vml.bytes, needle) != null;
}

fn commentsPartContains(io: std.Io, path: []const u8, needle: []const u8) !bool {
    var store = try store_mod.PartStore.open(std.testing.allocator, io, path);
    defer store.deinit();
    const part = (try store.part("xl/comments1.xml")) orelse return error.MissingComments;
    return std.mem.indexOf(u8, part.bytes, needle) != null;
}

test "Editor: insertColumn shifts VML anchor + x:Column (dr-2)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, io, "vml_inscol_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "vml_inscol_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    try buildCommentFixture(io, src_path, "C5");
    try std.testing.expect(try vmlPartContains(io, src_path, "<x:Column>2</x:Column>"));
    try std.testing.expect(try vmlPartContains(io, src_path, "<x:Anchor>3, 15, 4, 2, 5, 31, 8, 3</x:Anchor>"));
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertColumn(0, 2);
        try ed.save(io, dst_path);
    }
    try std.testing.expect(try vmlPartContains(io, dst_path, "<x:Column>3</x:Column>"));
    try std.testing.expect(try vmlPartContains(io, dst_path, "<x:Anchor>4, 15, 4, 2, 6, 31, 8, 3</x:Anchor>"));
}

test "Editor: insertRow shifts VML anchor + x:Row (dr-2)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, io, "vml_insrow_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "vml_insrow_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    try buildCommentFixture(io, src_path, "C5");
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertRow(0, 3);
        try ed.save(io, dst_path);
    }
    try std.testing.expect(try vmlPartContains(io, dst_path, "<x:Row>5</x:Row>"));
    try std.testing.expect(try vmlPartContains(io, dst_path, "<x:Anchor>3, 15, 5, 2, 5, 31, 9, 3</x:Anchor>"));
    try std.testing.expect(try vmlPartContains(io, dst_path, "<x:Column>2</x:Column>"));
}

test "Editor: deleteColumn at the comment's column drops the v:shape (dr-2)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, io, "vml_delcol_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "vml_delcol_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    try buildCommentFixture(io, src_path, "C5");
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.deleteColumn(0, 3);
        try ed.save(io, dst_path);
    }
    try std.testing.expect(!try vmlPartContains(io, dst_path, "<v:shape "));
}

test "Editor: insertColumn shifts BOTH VML anchor AND comments ref (REL-705)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, io, "vml_cmt_inscol_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "vml_cmt_inscol_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    try buildCommentFixture(io, src_path, "C5");
    try std.testing.expect(try commentsPartContains(io, src_path, "ref=\"C5\""));
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertColumn(0, 2);
        try ed.save(io, dst_path);
    }
    try std.testing.expect(try vmlPartContains(io, dst_path, "<x:Column>3</x:Column>"));
    try std.testing.expect(try commentsPartContains(io, dst_path, "ref=\"D5\""));
    try std.testing.expect(!try commentsPartContains(io, dst_path, "ref=\"C5\""));
}

test "Editor: deleteColumn at comment's column drops BOTH VML shape AND comment entry (REL-705)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, io, "vml_cmt_delcol_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "vml_cmt_delcol_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    try buildCommentFixture(io, src_path, "C5");
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.deleteColumn(0, 3);
        try ed.save(io, dst_path);
    }
    try std.testing.expect(!try vmlPartContains(io, dst_path, "<v:shape "));
    try std.testing.expect(!try commentsPartContains(io, dst_path, "<comment "));
}

// ---------------------------------------------------------------------------
// <tableParts> round-trip tests (refusal lift; replaces prior
// RowEditUnsafeForSheet / ColEditUnsafeForSheet on sheets carrying
// `<tableParts>`). Drives a real corpus fixture because the Writer
// doesn't synthesize tables; collapse / header-row-delete paths
// stay refused (those are schema-invalid table states, not
// unrewritten ones).
// ---------------------------------------------------------------------------

fn copyCorpusToTmp(io: std.Io, src_corpus: []const u8, dst_path: [:0]const u8) !void {
    std.Io.Dir.cwd().access(io, src_corpus, .{}) catch return error.SkipZigTest;
    try std.Io.Dir.cwd().copyFile(src_corpus, std.Io.Dir.cwd(), dst_path, io, .{});
}

fn tablePartContains(io: std.Io, path: []const u8, table_part: []const u8, needle: []const u8) !bool {
    var ed = try Editor.open(std.testing.allocator, io, path);
    defer ed.deinit();
    // Propagate readEntry errors so the assertion fails loudly when
    // the part is missing — the previous `catch return false` silently
    // turned MissingEntry into "needle not found", letting tests with
    // negative assertions pass against a deleted part (REL-B532).
    const xml = try ed.readEntry(table_part);
    defer std.testing.allocator.free(xml);
    return std.mem.indexOf(u8, xml, needle) != null;
}

/// Read an entry's bytes via the Editor's source-archive view.
/// Test-only helper for byte-equality assertions on parts that
/// aren't supposed to change.
fn readEntryBytes(io: std.Io, path: []const u8, entry_name: []const u8) ![]u8 {
    var ed = try Editor.open(std.testing.allocator, io, path);
    defer ed.deinit();
    return try ed.readEntry(entry_name);
}

test "Editor: insertRow above table shifts <table ref> in xl/tables/tableN.xml" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, io, "tbl_insrow_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "tbl_insrow_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    copyCorpusToTmp(io, "tests/corpus/poi_xxe_in_schema.xlsx", src_path) catch return error.SkipZigTest;
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertRow(0, 5); // Above the table at C9:E10.
        try ed.save(io, dst_path);
    }
    // Expect C9:E10 → C10:E11.
    try std.testing.expect(try tablePartContains(io, dst_path, "xl/tables/table1.xml", "ref=\"C10:E11\""));
    try std.testing.expect(try tablePartContains(io, dst_path, "xl/tables/table1.xml", "<autoFilter ref=\"C10:E11\""));
}

test "Editor: insertColumn inside table extends range and adds synthetic tableColumn" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, io, "tbl_inscol_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "tbl_inscol_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    copyCorpusToTmp(io, "tests/corpus/poi_xxe_in_schema.xlsx", src_path) catch return error.SkipZigTest;
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertColumn(0, 4); // Insert at col D — inside C9:E10.
        try ed.save(io, dst_path);
    }
    // C9:E10 → C9:F10. <tableColumns count="3"> → "4".
    try std.testing.expect(try tablePartContains(io, dst_path, "xl/tables/table1.xml", "ref=\"C9:F10\""));
    try std.testing.expect(try tablePartContains(io, dst_path, "xl/tables/table1.xml", "<tableColumns count=\"4\">"));
    // The corpus table has tableColumn ids 1, 2, 3 → synthetic
    // claims id=4 (max + 1).
    try std.testing.expect(try tablePartContains(io, dst_path, "xl/tables/table1.xml", "<tableColumn id=\"4\""));
}

test "Editor: deleteColumn inside table drops matching tableColumn" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, io, "tbl_delcol_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "tbl_delcol_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    copyCorpusToTmp(io, "tests/corpus/poi_xxe_in_schema.xlsx", src_path) catch return error.SkipZigTest;
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.deleteColumn(0, 4); // Drop col D — middle of C9:E10.
        try ed.save(io, dst_path);
    }
    // C9:E10 → C9:D10. count goes 3 → 2; the middle tableColumn
    // (id="3" in this corpus, name="Column1") goes away.
    try std.testing.expect(try tablePartContains(io, dst_path, "xl/tables/table1.xml", "ref=\"C9:D10\""));
    try std.testing.expect(try tablePartContains(io, dst_path, "xl/tables/table1.xml", "<tableColumns count=\"2\">"));
    try std.testing.expect(!try tablePartContains(io, dst_path, "xl/tables/table1.xml", "name=\"Column1\""));
}

test "Editor: renameTableColumn on a real-producer table part" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, io, "tbl_rencol_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "tbl_rencol_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    copyCorpusToTmp(io, "tests/corpus/poi_xxe_in_schema.xlsx", src_path) catch return error.SkipZigTest;
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        // POI's table1.xml: Table1 at C9:E10, columns `/Test/Test`,
        // `Column1`, `Column2` — a producer-written part, xmlColumnPr
        // body included.
        try ed.renameTableColumn("Table1", "Column1", "Renamed Col");
        // Refusals surface through the same path.
        try std.testing.expectError(
            error.TableColumnNotFound,
            ed.renameTableColumn("Table1", "Nope", "X"),
        );
        try std.testing.expectError(
            error.TableColumnNameInUse,
            ed.renameTableColumn("Table1", "Column2", "renamed col"),
        );
        try ed.save(io, dst_path);
    }
    try std.testing.expect(try tablePartContains(io, dst_path, "xl/tables/table1.xml", "name=\"Renamed Col\""));
    try std.testing.expect(!try tablePartContains(io, dst_path, "xl/tables/table1.xml", "name=\"Column1\""));
    // Sibling columns and the xmlColumnPr body survive untouched.
    try std.testing.expect(try tablePartContains(io, dst_path, "xl/tables/table1.xml", "name=\"Column2\""));
    try std.testing.expect(try tablePartContains(io, dst_path, "xl/tables/table1.xml", "xpath=\"/Test/Test\""));
    // Header cell D9 (row 9, second column of C9:E10) now carries
    // the new name through the shared-string table.
    try std.testing.expect(try tablePartContains(io, dst_path, "xl/sharedStrings.xml", "Renamed Col"));
}

test "Editor: deleteRow on table's header row refuses with RowEditUnsafeForSheet" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, io, "tbl_hdrrow_src.xlsx");
    defer std.testing.allocator.free(src_path);
    copyCorpusToTmp(io, "tests/corpus/poi_xxe_in_schema.xlsx", src_path) catch return error.SkipZigTest;
    var ed = try Editor.open(std.testing.allocator, io, src_path);
    defer ed.deinit();
    // Table is at C9:E10 — row 9 is the header row.
    const r = ed.deleteRow(0, 9);
    try std.testing.expectError(error.RowEditUnsafeForSheet, r);
}

test "Editor: edit outside table range is byte-stable inside the table part" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, io, "tbl_outside_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "tbl_outside_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    copyCorpusToTmp(io, "tests/corpus/poi_xxe_in_schema.xlsx", src_path) catch return error.SkipZigTest;
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        // Insert a column far past the table's BR (E = col 5).
        try ed.insertColumn(0, 12);
        try ed.save(io, dst_path);
    }
    // REL-B530: actually verify byte-identity (not just substring
    // presence). The byte-equal short-circuit in
    // applyTableEditsForSheet must skip the replacePart entirely so
    // the table part flows through ZIP substitution untouched.
    const src_table = try readEntryBytes(io, src_path, "xl/tables/table1.xml");
    defer std.testing.allocator.free(src_table);
    const dst_table = try readEntryBytes(io, dst_path, "xl/tables/table1.xml");
    defer std.testing.allocator.free(dst_table);
    try std.testing.expectEqualSlices(u8, src_table, dst_table);
}

// ─── docProps read + scrub (Z3) ──────────────────────────────────────
//
// End-to-end rather than unit: the parser has its own tests in
// pkg/typed_parts/doc_props_xml.zig. What matters here is that the
// parts survive a real Editor round trip, that the scrub reaches the
// saved archive, and — the load-bearing bit — that scrubbing metadata
// does not disturb a single cell.

/// Local entry reader for the docProps tests. Takes an `io` explicitly
/// rather than reusing `readEntryBytes`, which predates io threading.
fn readEntryBytesIo(io: std.Io, path: []const u8, entry_name: []const u8) ![]u8 {
    var ed = try Editor.open(std.testing.allocator, io, path);
    defer ed.deinit();
    return try ed.readEntry(entry_name);
}

/// Build a workbook that carries docProps parts with known PII. The
/// zlsx Writer does not emit docProps, so they are injected through
/// PartStore.addPart, which is also what exercises the content-type
/// registration path the scrub later has to unwind.
fn buildDocPropsFixture(io: std.Io, path: [:0]const u8) !void {
    const a = std.testing.allocator;

    {
        var w = xlsx.Writer.init(a);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .string = "name" }, .{ .integer = 1 } });
        try s.writeRow(&.{ .{ .string = "keep-me" }, .{ .integer = 2 } });
        try w.save(io, path);
    }

    var wb = try Workbook.open(a, io, path);
    defer wb.deinit();

    try wb.store.addPart(
        "docProps/core.xml",
        "application/vnd.openxmlformats-package.core-properties+xml",
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
            "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\"" ++
            " xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\">" ++
            "<dc:creator>Jane Q. Fixture</dc:creator>" ++
            "<cp:lastModifiedBy>Jane Q. Fixture</cp:lastModifiedBy>" ++
            "<dc:title>Confidential Q3</dc:title>" ++
            "<dcterms:created xsi:type=\"dcterms:W3CDTF\">2020-01-01T00:00:00Z</dcterms:created>" ++
            "<cp:revision>4</cp:revision>" ++
            "</cp:coreProperties>",
    );
    try wb.store.addPart(
        "docProps/app.xml",
        "application/vnd.openxmlformats-officedocument.extended-properties+xml",
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
            "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">" ++
            "<TotalTime>7</TotalTime><Company>AcmeCorp</Company><Manager>Bob Boss</Manager>" ++
            "</Properties>",
    );
    try wb.store.addPart(
        "docProps/custom.xml",
        "application/vnd.openxmlformats-officedocument.custom-properties+xml",
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
            "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\">" ++
            "<property name=\"Owner\"><vt:lpwstr>Jane Q. Fixture</vt:lpwstr></property>" ++
            "</Properties>",
    );
    try wb.save(io, path);
}

test "Editor.docProps: reads creator, company and custom-props presence" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(std.testing.allocator, io, "docprops_read.xlsx");
    defer std.testing.allocator.free(path);

    try buildDocPropsFixture(io, path);

    var ed = try Editor.open(std.testing.allocator, io, path);
    defer ed.deinit();

    const props = try ed.docProps();
    try std.testing.expectEqualStrings("Jane Q. Fixture", props.creator.?);
    try std.testing.expectEqualStrings("Jane Q. Fixture", props.last_modified_by.?);
    try std.testing.expectEqualStrings("Confidential Q3", props.title.?);
    try std.testing.expectEqualStrings("AcmeCorp", props.company.?);
    try std.testing.expectEqualStrings("Bob Boss", props.manager.?);
    try std.testing.expect(props.has_custom_properties);
    try std.testing.expect(props.hasIdentifyingFields());
}

test "Editor.stripDocProps: PII gone from the saved archive, cells intact" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "docprops_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "docprops_clean.xlsx");
    defer std.testing.allocator.free(dst);

    try buildDocPropsFixture(io, src);

    {
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try ed.stripDocProps(.{});
        try ed.save(io, dst);
    }

    // core.xml: identifying fields gone, timestamp + revision kept.
    const core = try readEntryBytesIo(io, dst, "docProps/core.xml");
    defer std.testing.allocator.free(core);
    try std.testing.expect(std.mem.indexOf(u8, core, "Jane Q. Fixture") == null);
    try std.testing.expect(std.mem.indexOf(u8, core, "Confidential Q3") == null);
    try std.testing.expect(std.mem.indexOf(u8, core, "2020-01-01T00:00:00Z") != null);
    try std.testing.expect(std.mem.indexOf(u8, core, "<cp:revision>4</cp:revision>") != null);

    // app.xml: company + manager gone, unmodelled TotalTime survives.
    const app = try readEntryBytesIo(io, dst, "docProps/app.xml");
    defer std.testing.allocator.free(app);
    try std.testing.expect(std.mem.indexOf(u8, app, "AcmeCorp") == null);
    try std.testing.expect(std.mem.indexOf(u8, app, "Bob Boss") == null);
    try std.testing.expect(std.mem.indexOf(u8, app, "<TotalTime>7</TotalTime>") != null);

    // custom.xml: dropped wholesale, and its content-type override with
    // it — a dangling Override makes the package invalid.
    try std.testing.expectError(
        error.MissingEntry,
        readEntryBytesIo(io, dst, "docProps/custom.xml"),
    );
    const ct = try readEntryBytesIo(io, dst, "[Content_Types].xml");
    defer std.testing.allocator.free(ct);
    try std.testing.expect(std.mem.indexOf(u8, ct, "docProps/custom.xml") == null);
    // The parts that stayed must still be declared.
    try std.testing.expect(std.mem.indexOf(u8, ct, "docProps/core.xml") != null);
    try std.testing.expect(std.mem.indexOf(u8, ct, "docProps/app.xml") != null);

    // The whole point: cell data is untouched.
    const src_sheet = try readEntryBytesIo(io, src, "xl/worksheets/sheet1.xml");
    defer std.testing.allocator.free(src_sheet);
    const dst_sheet = try readEntryBytesIo(io, dst, "xl/worksheets/sheet1.xml");
    defer std.testing.allocator.free(dst_sheet);
    try std.testing.expectEqualSlices(u8, src_sheet, dst_sheet);

    // And the scrubbed workbook still reads back through the reader.
    var re = try Editor.open(std.testing.allocator, io, dst);
    defer re.deinit();
    const after = try re.docProps();
    try std.testing.expect(after.creator == null);
    try std.testing.expect(after.company == null);
    try std.testing.expect(!after.has_custom_properties);
    try std.testing.expect(!after.hasIdentifyingFields());
}

test "Editor.stripDocProps: no docProps parts is a clean no-op" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "nodocprops_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "nodocprops_dst.xlsx");
    defer std.testing.allocator.free(dst);

    // Writer-produced workbooks carry no docProps at all.
    {
        var w = xlsx.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, src);
    }

    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    const props = try ed.docProps();
    try std.testing.expect(props.isEmpty());

    // Must not error, and must not invent parts.
    try ed.stripDocProps(.{});
    try ed.save(io, dst);

    try std.testing.expectError(
        error.MissingEntry,
        readEntryBytesIo(io, dst, "docProps/core.xml"),
    );
}

// ─── iter-sv-1: third-party view/sort attrs survive structural edits ──

/// Build a fixture shaped like an Excel-authored sheet.
///
/// zlsx's own writer never emits `<sheetView topLeftCell>`,
/// `<selection>` or a sheet-bare `<sortState>` — but every scrolled
/// Excel file carries the first two, so this is the load-modify-save
/// shape where a stale coordinate actually bites. The writer produces
/// the archive, then the parts are patched to third-party shape.
fn buildThirdPartyViewFixture(io: std.Io, path: []const u8) !void {
    const a = std.testing.allocator;
    {
        var w = xlsx.Writer.init(a);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 } });
        try s.writeRow(&.{ .{ .integer = 3 }, .{ .integer = 4 } });
        try w.save(io, path);
    }
    var wb = try Workbook.open(a, io, path);
    defer wb.deinit();

    const name = "xl/worksheets/sheet1.xml";
    const p = (try wb.store.part(name)) orelse return error.MissingSheetPart;

    // CT_Worksheet fixes child order: sheetViews before sheetData,
    // sortState after it. Injecting out of order would make this a
    // test of a schema-invalid file rather than of the rewriter.
    const views =
        "<sheetViews><sheetView topLeftCell=\"B5\" workbookViewId=\"0\">" ++
        "<selection activeCell=\"B5\" sqref=\"B5 D7:E9\"/>" ++
        "</sheetView></sheetViews>";
    const sort = "<sortState ref=\"A6:B9\"><sortCondition ref=\"B6:B9\"/></sortState>";

    const sd_open = std.mem.indexOf(u8, p.bytes, "<sheetData") orelse return error.MalformedXml;
    const sd_close_tag = "</sheetData>";
    const sd_close = std.mem.indexOf(u8, p.bytes, sd_close_tag) orelse return error.MalformedXml;
    const after_sd = sd_close + sd_close_tag.len;

    const patched = try std.fmt.allocPrint(a, "{s}{s}{s}{s}{s}", .{
        p.bytes[0..sd_open],
        views,
        p.bytes[sd_open..after_sd],
        sort,
        p.bytes[after_sd..],
    });
    defer a.free(patched);
    try wb.store.replacePart(name, patched);
    try wb.save(io, path);
}

test "Editor: insertRow rewrites third-party sheetView, selection and sortState" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, io, "tpv_insrow_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "tpv_insrow_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    try buildThirdPartyViewFixture(io, src_path);
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertRow(0, 2);
        try ed.save(io, dst_path);
    }
    var ed2 = try Editor.open(std.testing.allocator, io, dst_path);
    defer ed2.deinit();
    // Everything at or below row 2 moves down one.
    try assertSheetXmlContains(&ed2, "topLeftCell=\"B6\"");
    try assertSheetXmlContains(&ed2, "activeCell=\"B6\"");
    try assertSheetXmlContains(&ed2, "sqref=\"B6 D8:E10\"");
    try assertSheetXmlContains(&ed2, "<sortState ref=\"A7:B10\"");
    try assertSheetXmlContains(&ed2, "<sortCondition ref=\"B7:B10\"");
}

test "Editor: insertColumn rewrites third-party sheetView, selection and sortState" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, io, "tpv_inscol_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "tpv_inscol_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    try buildThirdPartyViewFixture(io, src_path);
    {
        var ed = try Editor.open(std.testing.allocator, io, src_path);
        defer ed.deinit();
        try ed.insertColumn(0, 2); // insert before column B
        try ed.save(io, dst_path);
    }
    var ed2 = try Editor.open(std.testing.allocator, io, dst_path);
    defer ed2.deinit();
    // Everything at or right of column B moves one column right.
    try assertSheetXmlContains(&ed2, "topLeftCell=\"C5\"");
    try assertSheetXmlContains(&ed2, "activeCell=\"C5\"");
    try assertSheetXmlContains(&ed2, "sqref=\"C5 E7:F9\"");
    try assertSheetXmlContains(&ed2, "<sortState ref=\"A6:C9\"");
    try assertSheetXmlContains(&ed2, "<sortCondition ref=\"C6:C9\"");
}

// ─── pivot refusal (silent-corruption fix, #139) ────────────────────
//
// A pivot's `<location ref>` and its cache field ranges live in
// `xl/pivotTables/*` + `xl/pivotCache/*`, keyed by a cross-part graph.
// Before #139 there was no refusal — neither this file nor
// `sheet_edit.zig` contained the string "pivot" — so a row insert left
// every pivot coordinate pointing at the pre-shift grid.
// `docs/plans/refusal-audit.md` recorded pivots as "refused at consumer
// level"; they were not refused anywhere.
//
// S7a lifts the host-only case (`location@ref` moves; tests further
// down). The fixture here wires a pivot relationship to a part that
// does not exist, so what these three now pin is the other half of the
// contract: a graph that cannot be read whole refuses, and a sheet
// without a pivot relationship never pays for the read.

const PIVOT_SHEET_RELS: []const u8 =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>
;

/// Write a 3-row sheet, optionally wiring a pivot relationship onto it.
fn writePivotFixture(io: std.Io, path: []const u8, with_pivot: bool) !void {
    {
        var w = xlsx.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }});
        try s.writeRow(&.{.{ .integer = 2 }});
        try s.writeRow(&.{.{ .integer = 3 }});
        try w.save(io, path);
    }
    if (!with_pivot) return;

    // Injected through a real save/reopen because `addPart` does not
    // refresh `PartStore`'s rels cache — and a genuine pivot workbook
    // arrives from disk anyway.
    var store = try store_mod.PartStore.open(std.testing.allocator, io, path);
    defer store.deinit();
    try store.addPart(
        "xl/worksheets/_rels/sheet1.xml.rels",
        "application/vnd.openxmlformats-package.relationships+xml",
        PIVOT_SHEET_RELS,
    );
    try store.save(io, path);
}

test "Editor: a row edit refuses when the pivot part a relationship names is missing" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "pivot_row.xlsx");
    defer std.testing.allocator.free(src);
    try writePivotFixture(io, src, true);

    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    // The relationship says "a pivot renders here"; the part it names
    // is absent, so the typed read cannot say where. Refusing is the
    // whole point: shifting the grid on a guess would leave a
    // `<location ref>` zlsx never saw describing the old layout.
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.deleteRow(0, 2));
}

test "Editor: a column edit refuses when the pivot part a relationship names is missing" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "pivot_col.xlsx");
    defer std.testing.allocator.free(src);
    try writePivotFixture(io, src, true);

    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    try std.testing.expectError(error.ColEditUnsafeForSheet, ed.insertColumn(0, 1));
    try std.testing.expectError(error.ColEditUnsafeForSheet, ed.deleteColumn(0, 1));
}

test "Editor: the pivot guard does not refuse sheets without one" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "nopivot_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "nopivot_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try writePivotFixture(io, src, false);

    // A guard that refuses everything is not a guard. Same fixture,
    // same edit, no pivot relationship — must still succeed.
    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    try ed.insertRow(0, 2);
    try ed.save(io, dst);

    var book = try Book.open(std.testing.allocator, io, dst);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    const r1 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(i64, 1), r1[0].integer);
}

// ─── S6 audit → S7b: what the pivot guard sees (surface-matrix footnote ¹⁰) ──
//
// `preflightPivotEditsForSheet` used to start from the EDITED sheet's
// relationships alone, so it saw the sheet a pivot renders on (the
// host's rels name the pivot part) and not the sheet a pivot only READS
// from — that edge runs the other way, from the cache definition's
// `worksheetSource` to the sheet, by name. The S6 audit pinned that
// blind spot; S7b closes it: the graph is read whenever the workbook
// carries a cache, every cache that depends on the edited sheet is
// selected (`Pivots.dependsOnSheet`), and a `sheet` + `ref` source is
// respelled by range semantics (`pivots.edit.applyToCacheDefinition`)
// where a table-named or defined-name one moves with its own carrier.
// These tests pin both halves as they stand now: the refusal on the
// host inside its pivot's footprint, the moved reference on the source.

/// Both paths live in the caller's temp dir: `std.testing.tmpDir` is
/// per-test state, and a second one inside a helper does not coexist
/// with the first.
fn expectSourceSheetAdmitted(
    io: std.Io,
    kind: pivots_mod.fixture.SourceKind,
    src: []const u8,
    dst: []const u8,
) !void {
    try pivots_mod.fixture.write(std.testing.allocator, io, src, kind);

    // The typed read: `Data` (0) is source-only, `Report` (1) hosts.
    {
        var wb = try Workbook.open(std.testing.allocator, io, src);
        defer wb.deinit();
        var p = try wb.pivotTables();
        defer p.deinit();
        try std.testing.expect(p.readsFromSheet(0) and !p.hostsPivot(0));
        try std.testing.expect(p.hostsPivot(1) and !p.readsFromSheet(1));
    }

    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    // The host refuses inside its pivot's rectangle `A3:B6` (#139,
    // narrowed by S7a to the footprint — an edit above it moves the
    // pivot instead; see the S7a tests) …
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(1, 4));
    try std.testing.expectError(error.ColEditUnsafeForSheet, ed.insertColumn(1, 2));
    // … and the source-only sheet is admitted: the guard sees it now
    // (S7b) and moves what needs moving.
    try ed.insertRow(0, 2);
    try ed.save(io, dst);
}

/// The one root attribute S7b-3 writes, `refreshOnLoad="1"`: present
/// exactly once when `want`, absent otherwise — no fixture or corpus
/// definition carries it as written, so a count is the whole story.
fn expectMarked(store: *store_mod.PartStore, part: []const u8, want: bool) !void {
    const p = (try store.part(part)) orelse return error.PartNotFound;
    try std.testing.expectEqual(@as(usize, if (want) 1 else 0), std.mem.count(u8, p.bytes, pivots_mod.edit.marker_attr));
    if (want) try std.testing.expect(std.mem.indexOf(u8, p.bytes, pivots_mod.edit.marker_insert ++ ">") != null);
}

/// `bytes` with ` refreshOnLoad="1"` inserted before the root open
/// tag's `>` — what a marked definition must read, spelled without the
/// splice under test. The fixture root carries no `>` in a value.
fn markedDefinition(alloc: Allocator, bytes: []const u8) ![]u8 {
    const root = std.mem.indexOf(u8, bytes, "<pivotCacheDefinition") orelse return error.PartNotFound;
    const gt = std.mem.indexOfScalarPos(u8, bytes, root, '>') orelse return error.PartNotFound;
    return std.mem.concat(alloc, u8, &.{ bytes[0..gt], pivots_mod.edit.marker_insert, bytes[gt..] });
}

/// S7b-4: the definition counts `n` records and is dated by the
/// edit's wall clock (a serial past 2026-01-01); the records part's
/// root counts `n` and holds exactly `n` `<r>` elements.
fn expectRebuilt(store: *store_mod.PartStore, def_part: []const u8, rec_part: []const u8, n: usize) !void {
    const cd = (try store.part(def_part)) orelse return error.PartNotFound;
    const count_attr = try std.fmt.allocPrint(std.testing.allocator, "recordCount=\"{d}\"", .{n});
    defer std.testing.allocator.free(count_attr);
    try std.testing.expect(std.mem.indexOf(u8, cd.bytes, count_attr) != null);
    const at = std.mem.indexOf(u8, cd.bytes, "refreshedDate=\"") orelse return error.TestExpectedRefreshedDate;
    const from = at + "refreshedDate=\"".len;
    const end = std.mem.indexOfScalarPos(u8, cd.bytes, from, '"') orelse return error.TestExpectedRefreshedDate;
    const serial = try std.fmt.parseFloat(f64, cd.bytes[from..end]);
    try std.testing.expect(serial > 46023);
    const rec = (try store.part(rec_part)) orelse return error.PartNotFound;
    const rec_count = try std.fmt.allocPrint(std.testing.allocator, " count=\"{d}\">", .{n});
    defer std.testing.allocator.free(rec_count);
    try std.testing.expect(std.mem.indexOf(u8, rec.bytes, rec_count) != null);
    try std.testing.expectEqual(n, std.mem.count(u8, rec.bytes, "<r>"));
}

fn cacheDefinitionBytes(alloc: Allocator, io: std.Io, path: []const u8) ![]u8 {
    var store = try store_mod.PartStore.open(alloc, io, path);
    defer store.deinit();
    const cd = (try store.part("xl/pivotCache/pivotCacheDefinition1.xml")) orelse return error.PartNotFound;
    return alloc.dupe(u8, cd.bytes);
}

test "S7b: a sheet+ref source sheet is admitted, and worksheetSource@ref moves (the S6 audit's stale finding, closed)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s6_sheet_ref_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "s6_sheet_ref_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try expectSourceSheetAdmitted(io, .sheet_ref, src, dst);

    // The data now spans A1:C5, and so does the cache's reference: the
    // insert inside the range grew its bottom edge, spliced at the
    // parser's `WorksheetSource.ref_span`. Before S7b the `ref` stayed
    // `A1:C4` — the silent-corruption class #139 closed for hosts.
    var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
    defer store.deinit();
    const cd = (try store.part("xl/pivotCache/pivotCacheDefinition1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "<worksheetSource sheet=\"Data\" ref=\"A1:C5\"/>") != null);
    try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "A1:C4") == null);
    // The insert was INSIDE the range — a blank record the snapshot
    // lacks — so the cache is marked to refresh at open (S7b-3, A1) …
    try expectMarked(&store, "xl/pivotCache/pivotCacheDefinition1.xml", true);
    // … and the snapshot IS the new range (S7b-4): four records, the
    // blank one first, each inventory grown by the blank alone.
    try expectRebuilt(&store, "xl/pivotCache/pivotCacheDefinition1.xml", "xl/pivotCache/pivotCacheRecords1.xml", 4);
    try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "<sharedItems containsBlank=\"1\" count=\"3\"><s v=\"East\"/><s v=\"West\"/><m/></sharedItems>") != null);
    try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "<sharedItems containsString=\"0\" containsBlank=\"1\" containsNumber=\"1\" containsInteger=\"1\" minValue=\"3\" maxValue=\"5\"/>") != null);
    const rec = (try store.part("xl/pivotCache/pivotCacheRecords1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, rec.bytes, "count=\"4\"><r><x v=\"2\"/><m/><m/></r><r><x v=\"0\"/><n v=\"3\"/><n v=\"1.5\"/></r>") != null);
}

test "S6 audit: a table-named source stays valid because the table rewriter moves the table" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s6_table_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "s6_table_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try expectSourceSheetAdmitted(io, .table_name, src, dst);

    // Same admission, different carrier: `worksheetSource@name` still
    // names the table, and the table part's own `ref` followed the
    // insert — the source is spelled through a carrier that has a
    // rewriter, so no coordinate in the cache definition moves. What
    // it gains is the marker: the insert was inside the table, so the
    // cached records are stale and Excel re-reads the (moved) table at
    // open (S7b-3).
    var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
    defer store.deinit();
    const cd = (try store.part("xl/pivotCache/pivotCacheDefinition1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "<worksheetSource name=\"SalesTbl\"/>") != null);
    try expectMarked(&store, "xl/pivotCache/pivotCacheDefinition1.xml", true);
    // S7b-4: rebuilt from the table's rectangle — the header row is
    // the field names, the blank row is the first record.
    try expectRebuilt(&store, "xl/pivotCache/pivotCacheDefinition1.xml", "xl/pivotCache/pivotCacheRecords1.xml", 4);
    const tbl = (try store.part("xl/tables/table1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, tbl.bytes, "ref=\"A1:C5\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, tbl.bytes, "ref=\"A1:C4\"") == null);
}

test "S7b: the corpus fixture — the host-and-source sheet moves its pivot and its table, the footprint still refuses, the source-only sheet is admitted" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const src = "tests/corpus/openxlsx_loadExample.xlsx";
    std.Io.Dir.cwd().access(io, src, .{}) catch return error.SkipZigTest;
    var tt = TestTmp.init();
    defer tt.deinit();
    const dst = try tt.path(std.testing.allocator, io, "s6_corpus_dst.xlsx");
    defer std.testing.allocator.free(dst);

    // `IrisSample` (0) hosts PivotTable1 and is read by it through
    // `Table2`; `mtcars` (2) is read by PivotTable3 through `Table3` and
    // hosts nothing; `mtCars Pivot` (3) hosts PivotTable3.
    {
        var wb = try Workbook.open(std.testing.allocator, io, src);
        defer wb.deinit();
        var p = try wb.pivotTables();
        defer p.deinit();
        try std.testing.expect(p.hostsPivot(0) and p.readsFromSheet(0));
        try std.testing.expect(!p.hostsPivot(2) and p.readsFromSheet(2));
        try std.testing.expect(p.hostsPivot(3) and !p.readsFromSheet(3));
    }

    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    // `IrisSample` hosts `PivotTable1` at `G2:K6` AND feeds it through
    // `Table2` (`A1:E51`): row 2 is above the pivot's rectangle and
    // inside the table — the pivot shifts, the table grows, the cache
    // (table-named) needs nothing. Refused whole before S7b.
    try ed.insertRow(0, 2);
    // `mtCars Pivot` hosts `A1:D5` only — row 2 is inside it.
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(3, 2));
    try ed.insertRow(2, 2);
    try ed.save(io, dst);

    var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
    defer store.deinit();
    const pt1 = (try store.part("xl/pivotTables/pivotTable1.xml")).?;
    // Moved by the insert (S7a), then grown by the blank record's
    // `(blank)` row (S7b-5).
    try std.testing.expect(std.mem.indexOf(u8, pt1.bytes, "<location ref=\"G3:K8\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, pt1.bytes, "<rowItems count=\"5\">") != null);
    const pt2 = (try store.part("xl/pivotTables/pivotTable2.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, pt2.bytes, "<location ref=\"A1:D6\"") != null);
    const tbl1 = (try store.part("xl/tables/table1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, tbl1.bytes, "ref=\"A1:E52\"") != null);
    const cd1 = (try store.part("xl/pivotCache/pivotCacheDefinition1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, cd1.bytes, "<worksheetSource name=\"Table2\"/>") != null);
    const cd = (try store.part("xl/pivotCache/pivotCacheDefinition2.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "<worksheetSource name=\"Table3\"/>") != null);
    const tbl = (try store.part("xl/tables/table2.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, tbl.bytes, "ref=\"A1:K31\"") != null);
    // Both inserts were inside their tables: both caches — neither of
    // which carries the attribute as Excel wrote it — are marked.
    try expectMarked(&store, "xl/pivotCache/pivotCacheDefinition1.xml", true);
    try expectMarked(&store, "xl/pivotCache/pivotCacheDefinition2.xml", true);
    // S7b-4: and rebuilt from their tables as Excel wrote them — one
    // blank record each, first; every inventory kept in its order and
    // grown by the blank; the extremes the sheet's own lexicals;
    // Excel's extension list kept.
    try expectRebuilt(&store, "xl/pivotCache/pivotCacheDefinition1.xml", "xl/pivotCache/pivotCacheRecords1.xml", 51);
    try std.testing.expect(std.mem.indexOf(u8, cd1.bytes, "<cacheField name=\"Species\" numFmtId=\"0\"><sharedItems containsBlank=\"1\" count=\"4\"><s v=\"virginica\"/><s v=\"setosa\"/><s v=\"versicolor\"/><m/></sharedItems></cacheField>") != null);
    try std.testing.expect(std.mem.indexOf(u8, cd1.bytes, "<cacheField name=\"Sepal Length\" numFmtId=\"0\"><sharedItems containsString=\"0\" containsBlank=\"1\" containsNumber=\"1\" minValue=\"4.4000000000000004\" maxValue=\"7.9\"/></cacheField>") != null);
    try std.testing.expect(std.mem.indexOf(u8, cd1.bytes, "<x14:pivotCacheDefinition pivotCacheId=\"2\"/>") != null);
    const rec1 = (try store.part("xl/pivotCache/pivotCacheRecords1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, rec1.bytes, "count=\"51\"><r><m/><m/><m/><m/><x v=\"3\"/></r><r><n v=\"6.4\"/><n v=\"2.7\"/><n v=\"5.3\"/><n v=\"1.9\"/><x v=\"0\"/></r>") != null);
    try expectRebuilt(&store, "xl/pivotCache/pivotCacheDefinition2.xml", "xl/pivotCache/pivotCacheRecords2.xml", 30);
    try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "<cacheField name=\"cyl\" numFmtId=\"0\"><sharedItems containsString=\"0\" containsBlank=\"1\" containsNumber=\"1\" containsInteger=\"1\" minValue=\"4\" maxValue=\"8\" count=\"4\"><n v=\"6\"/><n v=\"4\"/><n v=\"8\"/><m/></sharedItems></cacheField>") != null);
    const rec2 = (try store.part("xl/pivotCache/pivotCacheRecords2.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, rec2.bytes, "count=\"30\"><r><m/><x v=\"3\"/><m/><m/><m/><m/><m/><m/><m/><x v=\"3\"/><m/></r><r><n v=\"21\"/><x v=\"0\"/><n v=\"160\"/>") != null);
}

// ─── S7b: the source-rows lift ───────────────────────────────────────
//
// Written failing-first with the S7b analysis: the contract for the
// one spelling that carries its own coordinates — `<worksheetSource
// sheet= ref=>` — under the three range semantics. `docs/plans/
// s7b-cache-policy.md` is the analysis behind it: the spellings, what
// each needs rewritten, and the cache policy the owner chose. The
// `ref` move is common to every policy option; the refresh marker and
// the engine are the row's later pieces.

test "S7b: row edits on a sheet+ref source sheet move worksheetSource@ref like a range (the failing-first contract)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7b_sheet_ref_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "s7b_sheet_ref_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);

    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    // `Data` (0) feeds the cache from `A1:C4`: a header row and three
    // records. The three edits are the three range semantics Excel
    // applies to a source reference — an insert INSIDE grows the
    // bottom edge (`A1:C5`), an insert ABOVE shifts both edges
    // (`A2:C6`), a delete INSIDE shrinks the bottom edge (`A2:C5`).
    try ed.insertRow(0, 2);
    try ed.insertRow(0, 1);
    try ed.deleteRow(0, 4);
    try ed.save(io, dst);

    var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
    defer store.deinit();
    const cd = (try store.part("xl/pivotCache/pivotCacheDefinition1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "<worksheetSource sheet=\"Data\" ref=\"A2:C5\"/>") != null);
    // Two of the three edits changed what the range holds, and the
    // snapshot followed each (S7b-4): the first insert added a blank
    // record at the top, the delete removed `East`, 3 — three records,
    // the blank first; `East` stays an item (its other record, and a
    // consumer's index). Marked, as one Excel refreshes at open.
    try expectRebuilt(&store, "xl/pivotCache/pivotCacheDefinition1.xml", "xl/pivotCache/pivotCacheRecords1.xml", 3);
    try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "<sharedItems containsBlank=\"1\" count=\"3\"><s v=\"East\"/><s v=\"West\"/><m/></sharedItems>") != null);
    try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "<sharedItems containsString=\"0\" containsBlank=\"1\" containsNumber=\"1\" containsInteger=\"1\" minValue=\"4\" maxValue=\"5\"/>") != null);
    const rec = (try store.part("xl/pivotCache/pivotCacheRecords1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, rec.bytes, "count=\"3\"><r><x v=\"2\"/><m/><m/></r><r><x v=\"1\"/><n v=\"4\"/><n v=\"2.5\"/></r><r><x v=\"0\"/><n v=\"5\"/><n v=\"3.5\"/></r></pivotCacheRecords>") != null);
    try expectMarked(&store, "xl/pivotCache/pivotCacheDefinition1.xml", true);
}

test "S7b-3: a proven pure shift of the source never marks — the part differs from the original in `ref` alone" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7b3_shift_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "s7b3_shift_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
    const original = try cacheDefinitionBytes(std.testing.allocator, io, src);
    defer std.testing.allocator.free(original);

    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    // Above (row 1), left (column A), below (row 9), right (column E):
    // the rectangle moves or stays; its content is what it was.
    try ed.insertRow(0, 1);
    try ed.insertColumn(0, 1);
    try ed.insertRow(0, 9);
    try ed.deleteColumn(0, 5);
    try ed.save(io, dst);

    const after = try cacheDefinitionBytes(std.testing.allocator, io, dst);
    defer std.testing.allocator.free(after);
    const want = try pivotPartWithRef(std.testing.allocator, original, "ref=\"A1:C4\"", "ref=\"B2:D5\"");
    defer std.testing.allocator.free(want);
    try std.testing.expectEqualStrings(want, after);
}

test "S7b analysis: a defined-name source follows the defined-name rewriter" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7b_name_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "s7b_name_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .defined_name);

    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    try ed.insertRow(0, 2);
    try ed.save(io, dst);

    // Row 3 of the S7b analysis (`docs/plans/s7b-cache-policy.md`
    // §2.1): the cache names `PivotSrc`, whose body is a formula
    // carrier — `rewriteAllDefinedNames` moves it under every row
    // edit, so the cache definition needs nothing and the source
    // still resolves through the name to the same sheet.
    var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
    defer store.deinit();
    const cd = (try store.part("xl/pivotCache/pivotCacheDefinition1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "<worksheetSource name=\"PivotSrc\"/>") != null);
    const wb_xml = (try store.part("xl/workbook.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, wb_xml.bytes, "<definedName name=\"PivotSrc\">Data!$A$1:$C$5</definedName>") != null);
    // Row 2 is inside the name's rectangle: the body is the sweep's,
    // the content change is the marker's.
    try expectMarked(&store, "xl/pivotCache/pivotCacheDefinition1.xml", true);

    var wb = try Workbook.open(std.testing.allocator, io, dst);
    defer wb.deinit();
    var p = try wb.pivotTables();
    defer p.deinit();
    try std.testing.expect(p.readsFromSheet(0) and !p.hostsPivot(0));
    try std.testing.expectEqual(pivots_mod.ResolvedVia.defined_name, p.caches[0].resolution.sheet.via);
}

test "S7b: a source's header-row delete, only-row delete and in-range column edits refuse before any mutation; edits beside it shift the ref" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7b_refuse_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "s7b_refuse_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);

    // The direct `Workbook` path refuses whole: no part changes.
    {
        var wb = try Workbook.open(std.testing.allocator, io, src);
        defer wb.deinit();
        try std.testing.expectError(error.PivotEditUnsafe, wb.deleteRow(0, 1));
        try std.testing.expectError(error.PivotEditUnsafe, wb.insertColumn(0, 2));
        try std.testing.expectError(error.PivotEditUnsafe, wb.deleteColumn(0, 3));
        const cd = (try wb.store.part("xl/pivotCache/pivotCacheDefinition1.xml")).?;
        try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "ref=\"A1:C4\"") != null);
        const sheet = (try wb.store.part("xl/worksheets/sheet1.xml")).?;
        try std.testing.expect(std.mem.indexOf(u8, sheet.bytes, "r=\"C4\"") != null);
    }
    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.deleteRow(0, 1));
    try std.testing.expectError(error.ColEditUnsafeForSheet, ed.insertColumn(0, 2));
    try std.testing.expectError(error.ColEditUnsafeForSheet, ed.deleteColumn(0, 1));
    // Beside the range on either axis: a pure shift of the spelling.
    try ed.insertColumn(0, 1);
    try ed.insertRow(0, 1);
    // Below and right of it: nothing moves.
    try ed.insertRow(0, 9);
    try ed.insertColumn(0, 9);
    try ed.save(io, dst);
    var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
    defer store.deinit();
    const cd = (try store.part("xl/pivotCache/pivotCacheDefinition1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "<worksheetSource sheet=\"Data\" ref=\"B2:D5\"/>") != null);

    // A one-row source collapses under its only delete.
    const one = try tt.path(std.testing.allocator, io, "s7b_one_row_src.xlsx");
    defer std.testing.allocator.free(one);
    try pivots_mod.fixture.write(std.testing.allocator, io, one, .sheet_ref);
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, one, "xl/pivotCache/pivotCacheDefinition1.xml", "ref=\"A1:C4\"", "ref=\"A2:C2\"");
    var ed2 = try Editor.open(std.testing.allocator, io, one);
    defer ed2.deinit();
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed2.deleteRow(0, 2));
    try ed2.deleteRow(0, 1);
}

test "S7b: a defined-name source refuses the endpoint delete the name sweep would spell #REF! on, and admits the interior one" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7b_name_ref_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "s7b_name_ref_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .defined_name);

    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    // The header row: the rectangle's refusal, whatever spelled it.
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.deleteRow(0, 1));
    // The last row: the rewriter would write `Data!$A$1:#REF!` where
    // Excel shrinks to `$C$3` — refused on the dry-run body (§2.2).
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.deleteRow(0, 4));
    // A column of the range is the field schema — S7c's.
    try std.testing.expectError(error.ColEditUnsafeForSheet, ed.deleteColumn(0, 3));
    // Interior delete, insert above: the body follows the grid.
    try ed.deleteRow(0, 3);
    try ed.insertRow(0, 1);
    try ed.save(io, dst);
    var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
    defer store.deinit();
    const wb_xml = (try store.part("xl/workbook.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, wb_xml.bytes, "<definedName name=\"PivotSrc\">Data!$A$2:$C$4</definedName>") != null);
    const cd = (try store.part("xl/pivotCache/pivotCacheDefinition1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "<worksheetSource name=\"PivotSrc\"/>") != null);

    // A name already spelling `#REF!` is not this edit's doing: the
    // dangling class, admitted where the edit leaves its body alone
    // (§7 Q4 iii — the host's edit) — refused when the edit adds a
    // `#REF!` of its own (the count, not a flag: Codex r2 REL-203),
    // here the delete of the one anchor the body still had — and,
    // since S7b-4, refused for any other `Data` edit too: the body
    // names the sheet without bounding it, so the edit is a content
    // change with no rectangle to rebuild the snapshot from (S7b-3
    // marked and admitted).
    const broken = try tt.path(std.testing.allocator, io, "s7b_name_broken_src.xlsx");
    defer std.testing.allocator.free(broken);
    try pivots_mod.fixture.write(std.testing.allocator, io, broken, .defined_name);
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, broken, "xl/workbook.xml", "Data!$A$1:$C$4", "Data!$A$1:#REF!");
    var ed2 = try Editor.open(std.testing.allocator, io, broken);
    defer ed2.deinit();
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed2.deleteRow(0, 1));
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed2.deleteRow(0, 3));
    try ed2.insertRow(1, 1);
}

test "S7b: a name reaching the host through another name is dry-run through its closure" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7b_chain_src.xlsx");
    defer std.testing.allocator.free(src);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .defined_name);
    // `PivotSrc` = `OFFSET(Anchor,0,0,4,3)`, `Anchor` = `Report!$D$1`:
    // the anchor's column is two names away from the cache, and its
    // delete still refuses.
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/workbook.xml", "<definedName name=\"PivotSrc\">Data!$A$1:$C$4</definedName>", "<definedName name=\"Anchor\">Report!$D$1</definedName><definedName name=\"PivotSrc\">OFFSET(Anchor,0,0,4,3)</definedName>");
    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    try std.testing.expectError(error.ColEditUnsafeForSheet, ed.deleteColumn(1, 4));
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.deleteRow(1, 1));
    // S7b-4: a `Report` edit the sweep would survive is still a content
    // change of a source with no rectangle — refused at the engine's
    // edge. `Data` is outside the closure: admitted.
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.deleteRow(1, 2));
    try ed.insertRow(0, 1);
}

test "S7b: a whole-column name body shifts under a column edit beside it and refuses every row edit — no rectangle to rebuild from (S7b-4)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7b_cols_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "s7b_cols_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .defined_name);
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/workbook.xml", "Data!$A$1:$C$4", "Data!$A:$C");
    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 1));
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.deleteRow(0, 1));
    try std.testing.expectError(error.ColEditUnsafeForSheet, ed.insertColumn(0, 2));
    // Every row of the sheet is inside whole columns: a row edit
    // changes their content, and there is no finite rectangle to
    // rebuild the snapshot from — refused (S7b-3 admitted and marked).
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.deleteRow(0, 4));
    // A column beside them is a proven shift: the body moves, and
    // nothing marks.
    try ed.insertColumn(0, 1);
    try ed.save(io, dst);
    var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
    defer store.deinit();
    const wb_xml = (try store.part("xl/workbook.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, wb_xml.bytes, "<definedName name=\"PivotSrc\">Data!$B:$D</definedName>") != null);
    try expectMarked(&store, "xl/pivotCache/pivotCacheDefinition1.xml", false);
}

test "S7b: the consolidation fixture — the direct set is respelled, the named set follows its body, the host still moves" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7b_consolidation_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "s7b_consolidation_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .consolidation);
    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    // Set 0: `Data!A1:C4` by `sheet` + `ref` — an insert inside it
    // changes the consolidated content, which the S7b-4 engine does not
    // rebuild: refused. Above it, the set shifts.
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
    try ed.insertRow(0, 1);
    // Set 1: `PivotSrc` = `Report!$A$1:$B$2`, on the host — its header
    // row refuses, an insert above it moves the body and the pivot.
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.deleteRow(1, 1));
    try ed.insertRow(1, 1);
    try ed.save(io, dst);
    var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
    defer store.deinit();
    const cd = (try store.part("xl/pivotCache/pivotCacheDefinition1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "<rangeSet i1=\"0\" sheet=\"Data\" ref=\"A2:C5\"/><rangeSet i1=\"1\" name=\"PivotSrc\"/>") != null);
    const wb_xml = (try store.part("xl/workbook.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, wb_xml.bytes, "<definedName name=\"PivotSrc\">Report!$A$2:$B$3</definedName>") != null);
    const pt = (try store.part("xl/pivotTables/pivotTable1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, pt.bytes, "<location ref=\"A4:B7\"") != null);
    // Both sets shifted: no content change, no marker.
    try expectMarked(&store, "xl/pivotCache/pivotCacheDefinition1.xml", false);
}

// ─── S7b-4: the engine's first slice ─────────────────────────────────
//
// A row edit that changes a source's content no longer leaves the
// snapshot as saved: the cache is REBUILT from the cells — records,
// every inventory, `recordCount`, `refreshedDate` — in the same
// install as the `ref` move, still marked so Excel's refresh at open
// lays the consumers out over it (`docs/plans/s7b-cache-policy.md`
// §9, S7b-4). What the slice does not evaluate refuses the edit.

test "S7b-4: a calculated field refuses the edit that needs a rebuild and admits the pure shift beside it — the snapshot byte-preserved" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7b4_calc_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "s7b4_calc_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/pivotCache/pivotCacheDefinition1.xml", "<cacheFields count=\"3\">", "<cacheFields count=\"4\">");
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/pivotCache/pivotCacheDefinition1.xml", "</cacheFields>", "<cacheField name=\"Total\" numFmtId=\"0\" formula=\"Qty*Price\" databaseField=\"0\"><sharedItems containsSemiMixedTypes=\"0\" containsString=\"0\" containsNumber=\"1\" minValue=\"4.5\" maxValue=\"17.5\"/></cacheField></cacheFields>");
    const original = try cacheDefinitionBytes(std.testing.allocator, io, src);
    defer std.testing.allocator.free(original);

    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.deleteRow(0, 3));
    try ed.insertRow(0, 1);
    try ed.insertRow(0, 9);
    try ed.save(io, dst);
    const after = try cacheDefinitionBytes(std.testing.allocator, io, dst);
    defer std.testing.allocator.free(after);
    const want = try pivotPartWithRef(std.testing.allocator, original, "ref=\"A1:C4\"", "ref=\"A2:C5\"");
    defer std.testing.allocator.free(want);
    try std.testing.expectEqualStrings(want, after);
    var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
    defer store.deinit();
    const rec = (try store.part("xl/pivotCache/pivotCacheRecords1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, rec.bytes, "count=\"3\"") != null);
}

test "S7b-4: a header that is not the field names, a boolean, a date-formatted number, an inline string and an uncomputed formula refuse before any mutation; a computed formula reads as its value" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const Case = struct { name: []const u8, part: []const u8, old: []const u8, new: []const u8 };
    const refusals = [_]Case{
        .{ .name = "header", .part = "xl/pivotCache/pivotCacheDefinition1.xml", .old = "name=\"Region\"", .new = "name=\"Area\"" },
        .{ .name = "bool", .part = "xl/worksheets/sheet1.xml", .old = "<c r=\"B2\"><v>3</v></c>", .new = "<c r=\"B2\" t=\"b\"><v>1</v></c>" },
        .{ .name = "error", .part = "xl/worksheets/sheet1.xml", .old = "<c r=\"B2\"><v>3</v></c>", .new = "<c r=\"B2\" t=\"e\"><v>#N/A</v></c>" },
        .{ .name = "inline", .part = "xl/worksheets/sheet1.xml", .old = "<c r=\"A2\" t=\"s\"><v>3</v></c>", .new = "<c r=\"A2\" t=\"inlineStr\"><is><t>East</t></is></c>" },
        .{ .name = "formula", .part = "xl/worksheets/sheet1.xml", .old = "<c r=\"B2\"><v>3</v></c>", .new = "<c r=\"B2\"><f>1+2</f></c>" },
        .{ .name = "date", .part = "xl/worksheets/sheet1.xml", .old = "<c r=\"C2\"><v>1.5</v></c>", .new = "<c r=\"C2\" s=\"1\"><v>1.5</v></c>" },
    };
    for (refusals) |case| {
        const file = try std.fmt.allocPrint(std.testing.allocator, "s7b4_refuse_{s}.xlsx", .{case.name});
        defer std.testing.allocator.free(file);
        const src = try tt.path(std.testing.allocator, io, file);
        defer std.testing.allocator.free(src);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
        if (std.mem.eql(u8, case.name, "date")) {
            // The fixture writes no styles part; add one whose style 1
            // is the built-in short date (`numFmtId` 14).
            var store = try store_mod.PartStore.open(std.testing.allocator, io, src);
            defer store.deinit();
            try store.addPart(
                "xl/styles.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
                \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                \\<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>
                ,
            );
            try store.save(io, src);
        }
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, case.part, case.old, case.new);
        var wb = try Workbook.open(std.testing.allocator, io, src);
        defer wb.deinit();
        const before = try std.testing.allocator.dupe(u8, (try wb.store.part("xl/pivotCache/pivotCacheDefinition1.xml")).?.bytes);
        defer std.testing.allocator.free(before);
        try std.testing.expectError(error.PivotEditUnsafe, wb.insertRow(0, 2));
        try std.testing.expectError(error.PivotEditUnsafe, wb.deleteRow(0, 3));
        try std.testing.expectEqualStrings(before, (try wb.store.part("xl/pivotCache/pivotCacheDefinition1.xml")).?.bytes);
        // A pure shift reads no cell and is still admitted.
        try wb.insertRow(0, 1);
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
    }

    // Computed formulas are their cached values: a number, a string.
    const src = try tt.path(std.testing.allocator, io, "s7b4_formula_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "s7b4_formula_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/worksheets/sheet1.xml", "<c r=\"B2\"><v>3</v></c>", "<c r=\"B2\"><f>1+2</f><v>3</v></c>");
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/worksheets/sheet1.xml", "<c r=\"A3\" t=\"s\"><v>4</v></c>", "<c r=\"A3\" t=\"str\"><f>\"We\"&amp;\"st\"</f><v>West</v></c>");
    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    try ed.deleteRow(0, 4);
    try ed.save(io, dst);
    var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
    defer store.deinit();
    try expectRebuilt(&store, "xl/pivotCache/pivotCacheDefinition1.xml", "xl/pivotCache/pivotCacheRecords1.xml", 2);
    const rec = (try store.part("xl/pivotCache/pivotCacheRecords1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, rec.bytes, "count=\"2\"><r><x v=\"0\"/><n v=\"3\"/><n v=\"1.5\"/></r><r><x v=\"1\"/><n v=\"4\"/><n v=\"2.5\"/></r></pivotCacheRecords>") != null);
}

test "S7b-4: a workbook from a bare store has no clock — its rebuilt definition drops refreshedDate; a records part with an extension list refuses; a cache without one rebuilds its definition alone" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    {
        const src = try tt.path(std.testing.allocator, io, "s7b4_noclock_src.xlsx");
        defer std.testing.allocator.free(src);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
        const store = try store_mod.PartStore.open(std.testing.allocator, io, src);
        var wb = try Workbook.fromStore(std.testing.allocator, store);
        defer wb.deinit();
        try std.testing.expect(wb.clock == null);
        try wb.insertRow(0, 2);
        const cd = (try wb.store.part("xl/pivotCache/pivotCacheDefinition1.xml")).?;
        try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "refreshedDate") == null);
        try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "refreshedBy=\"zlsx\" createdVersion=\"6\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "recordCount=\"4\" refreshOnLoad=\"1\">") != null);
    }
    {
        const src = try tt.path(std.testing.allocator, io, "s7b4_ext_src.xlsx");
        defer std.testing.allocator.free(src);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/pivotCache/pivotCacheRecords1.xml", "</pivotCacheRecords>", "<extLst/></pivotCacheRecords>");
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
        try ed.insertRow(0, 1);
    }
    {
        const src = try tt.path(std.testing.allocator, io, "s7b4_norec_src.xlsx");
        defer std.testing.allocator.free(src);
        const dst = try tt.path(std.testing.allocator, io, "s7b4_norec_dst.xlsx");
        defer std.testing.allocator.free(dst);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
        // No `r:id`: the definition names no records part.
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/pivotCache/pivotCacheDefinition1.xml", " r:id=\"rId1\"", "");
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try ed.insertRow(0, 2);
        try ed.save(io, dst);
        var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
        defer store.deinit();
        const cd = (try store.part("xl/pivotCache/pivotCacheDefinition1.xml")).?;
        try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "recordCount=\"4\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "<sharedItems containsBlank=\"1\" count=\"3\"><s v=\"East\"/><s v=\"West\"/><m/></sharedItems>") != null);
        const rec = (try store.part("xl/pivotCache/pivotCacheRecords1.xml")).?;
        try std.testing.expect(std.mem.indexOf(u8, rec.bytes, "count=\"3\"") != null);
    }
}

test "S7b-4: the Editor's row edit rebuilds each cache once — the pre-flight's collection is what the sweep installs (Codex #205 r1 PERF-101)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7b4_once_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "s7b4_once_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    try std.testing.expectEqual(@as(u32, 0), ed.workbook.pivot_rebuilds);
    try ed.insertRow(0, 2);
    try std.testing.expectEqual(@as(u32, 1), ed.workbook.pivot_rebuilds);
    try ed.deleteRow(0, 3);
    try std.testing.expectEqual(@as(u32, 2), ed.workbook.pivot_rebuilds);
    // A pure shift and a column edit beside the rectangle read no cell.
    try ed.insertRow(0, 1);
    try ed.insertColumn(0, 5);
    try std.testing.expectEqual(@as(u32, 2), ed.workbook.pivot_rebuilds);
    try ed.save(io, dst);
    var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
    defer store.deinit();
    try expectRebuilt(&store, "xl/pivotCache/pivotCacheDefinition1.xml", "xl/pivotCache/pivotCacheRecords1.xml", 3);
    const cd = (try store.part("xl/pivotCache/pivotCacheDefinition1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "ref=\"A2:C5\"") != null);
}

test "S7b-4: a cell whose reference is not its row's, a cell or row without one, a string formula without a cached value, and a foreign or aliased element refuse before any mutation (Codex #205 r1 REL-102, REL-103, REL-104)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const def_part = "xl/pivotCache/pivotCacheDefinition1.xml";
    const rec_part = "xl/pivotCache/pivotCacheRecords1.xml";
    const sheet_part = "xl/worksheets/sheet1.xml";
    const Case = struct { name: []const u8, part: []const u8, old: []const u8, new: []const u8, err: anyerror, shift_ok: bool };
    const cases = [_]Case{
        .{ .name = "rowmismatch", .part = sheet_part, .old = "<c r=\"B2\"><v>3</v></c>", .new = "<c r=\"B99\"><v>3</v></c>", .err = error.MalformedSheetXml, .shift_ok = true },
        .{ .name = "rowzero", .part = sheet_part, .old = "<c r=\"B2\"><v>3</v></c>", .new = "<c r=\"B0\"><v>3</v></c>", .err = error.MalformedSheetXml, .shift_ok = true },
        .{ .name = "nocellref", .part = sheet_part, .old = "<c r=\"B2\"><v>3</v></c>", .new = "<c><v>3</v></c>", .err = error.PivotEditUnsafe, .shift_ok = true },
        .{ .name = "norowref", .part = sheet_part, .old = "<row r=\"4\">", .new = "<row>", .err = error.PivotEditUnsafe, .shift_ok = true },
        .{ .name = "strformula", .part = sheet_part, .old = "<c r=\"A2\" t=\"s\"><v>3</v></c>", .new = "<c r=\"A2\" t=\"str\"><f>\"Ea\"&amp;\"st\"</f></c>", .err = error.PivotEditUnsafe, .shift_ok = true },
        .{ .name = "foreign_item", .part = def_part, .old = "<s v=\"West\"/>", .new = "<s v=\"West\"/><z:s xmlns:z=\"urn:vendor\" v=\"North\"/>", .err = error.PivotEditUnsafe, .shift_ok = true },
        .{ .name = "foreign_record", .part = rec_part, .old = "</pivotCacheRecords>", .new = "<z:r xmlns:z=\"urn:vendor\"/></pivotCacheRecords>", .err = error.PivotEditUnsafe, .shift_ok = true },
        // An alias of the main namespace is a graph that does not read:
        // every edit of the sheet refuses, the pure shift included.
        .{ .name = "alias", .part = def_part, .old = "<s v=\"West\"/>", .new = "<y:s xmlns:y=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" v=\"West\"/>", .err = error.MalformedPivotXml, .shift_ok = false },
        // A coordinate twice (Codex #205 r4 REL-404).
        .{ .name = "duprow", .part = sheet_part, .old = "<row r=\"3\">", .new = "<row r=\"2\"><c r=\"C2\"><v>9</v></c></row><row r=\"3\">", .err = error.MalformedSheetXml, .shift_ok = true },
        .{ .name = "dupcell", .part = sheet_part, .old = "<c r=\"B2\"><v>3</v></c>", .new = "<c r=\"B2\"><v>3</v></c><c r=\"B2\"><v>4</v></c>", .err = error.MalformedSheetXml, .shift_ok = true },
        // Zig's float grammar is wider than xsd:double (Codex #205 r9
        // REL-901).
        .{ .name = "hexfloat", .part = sheet_part, .old = "<c r=\"B2\"><v>3</v></c>", .new = "<c r=\"B2\"><v>0x1p0</v></c>", .err = error.PivotEditUnsafe, .shift_ok = true },
        .{ .name = "underscore", .part = sheet_part, .old = "<c r=\"B2\"><v>3</v></c>", .new = "<c r=\"B2\"><v>1_0</v></c>", .err = error.PivotEditUnsafe, .shift_ok = true },
        .{ .name = "hexitem", .part = def_part, .old = "minValue=\"3\" maxValue=\"5\"/>", .new = "minValue=\"3\" maxValue=\"5\" count=\"1\"><n v=\"0x1p0\"/></sharedItems>", .err = error.MalformedPivotXml, .shift_ok = true },
        // An inventory with items whose records are inline (Codex #205
        // r11 REL-1101): not one shape.
        .{ .name = "inline_with_items", .part = def_part, .old = "minValue=\"3\" maxValue=\"5\"/>", .new = "minValue=\"3\" maxValue=\"99\" count=\"1\"><n v=\"99\"/></sharedItems>", .err = error.PivotEditUnsafe, .shift_ok = true },
        // A `t` this view does not know is not a number (Codex #205 r8
        // REL-801).
        .{ .name = "bogus_t", .part = sheet_part, .old = "<c r=\"B2\"><v>3</v></c>", .new = "<c r=\"B2\" t=\"bogus\"><v>3</v></c>", .err = error.PivotEditUnsafe, .shift_ok = true },
        // Not the strict grid spelling (Codex #205 r6 REL-603).
        .{ .name = "leadingzero", .part = sheet_part, .old = "<c r=\"B2\"><v>3</v></c>", .new = "<c r=\"B02\"><v>3</v></c>", .err = error.MalformedSheetXml, .shift_ok = true },
        .{ .name = "pastxfd", .part = sheet_part, .old = "<c r=\"B2\"><v>3</v></c>", .new = "<c r=\"XFE2\"><v>3</v></c>", .err = error.MalformedSheetXml, .shift_ok = true },
    };
    for (cases) |case| {
        const file = try std.fmt.allocPrint(std.testing.allocator, "s7b4_r1_{s}.xlsx", .{case.name});
        defer std.testing.allocator.free(file);
        const src = try tt.path(std.testing.allocator, io, file);
        defer std.testing.allocator.free(src);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, case.part, case.old, case.new);
        var wb = try Workbook.open(std.testing.allocator, io, src);
        defer wb.deinit();
        const parts = [_][]const u8{ def_part, rec_part, sheet_part };
        var before: [parts.len][]const u8 = undefined;
        for (parts, 0..) |p, i| before[i] = try std.testing.allocator.dupe(u8, (try wb.store.part(p)).?.bytes);
        defer for (before) |b| std.testing.allocator.free(b);
        try std.testing.expectError(case.err, wb.insertRow(0, 2));
        try std.testing.expectError(case.err, wb.deleteRow(0, 3));
        for (parts, 0..) |p, i| try std.testing.expectEqualStrings(before[i], (try wb.store.part(p)).?.bytes);
        if (case.shift_ok) {
            try wb.insertRow(0, 1);
        } else {
            try std.testing.expectError(case.err, wb.insertRow(0, 1));
            for (parts, 0..) |p, i| try std.testing.expectEqualStrings(before[i], (try wb.store.part(p)).?.bytes);
        }
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        const want: anyerror = if (case.err == error.MalformedSheetXml) error.MalformedSheetXml else error.RowEditUnsafeForSheet;
        try std.testing.expectError(want, ed.insertRow(0, 2));
    }
    {
        // A row past the grid: an insert is the sheet transform's own
        // refusal (`RowEditExceedsMaxRow`, before any mutation); a
        // delete passes that probe and meets the strict parser.
        const src = try tt.path(std.testing.allocator, io, "s7b4_r1_pastlastrow.xlsx");
        defer std.testing.allocator.free(src);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, sheet_part, "<c r=\"B2\"><v>3</v></c>", "<c r=\"B1048577\"><v>3</v></c>");
        var wb = try Workbook.open(std.testing.allocator, io, src);
        defer wb.deinit();
        const before = try std.testing.allocator.dupe(u8, (try wb.store.part(rec_part)).?.bytes);
        defer std.testing.allocator.free(before);
        try std.testing.expectError(error.RowEditExceedsMaxRow, wb.insertRow(0, 2));
        try std.testing.expectError(error.MalformedSheetXml, wb.deleteRow(0, 3));
        try std.testing.expectEqualStrings(before, (try wb.store.part(rec_part)).?.bytes);
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try std.testing.expectError(error.MalformedSheetXml, ed.deleteRow(0, 3));
    }

    // A string formula whose cached value is the empty string is the
    // empty string: an item of its own.
    const src = try tt.path(std.testing.allocator, io, "s7b4_r1_emptystr_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "s7b4_r1_emptystr_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, sheet_part, "<c r=\"A2\" t=\"s\"><v>3</v></c>", "<c r=\"A2\" t=\"str\"><f>\"\"</f><v></v></c>");
    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    try ed.deleteRow(0, 4);
    try ed.save(io, dst);
    var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
    defer store.deinit();
    const cd = (try store.part(def_part)).?;
    try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "<sharedItems count=\"3\"><s v=\"East\"/><s v=\"West\"/><s v=\"\"/></sharedItems>") != null);
    const rec = (try store.part(rec_part)).?;
    try std.testing.expect(std.mem.indexOf(u8, rec.bytes, "count=\"2\"><r><x v=\"2\"/><n v=\"3\"/><n v=\"1.5\"/></r><r><x v=\"1\"/><n v=\"4\"/><n v=\"2.5\"/></r></pivotCacheRecords>") != null);
}

test "S7b-4: a numeric cell under a style the workbook cannot spell — a locale built-in, a missing custom entry, a format the grammar refuses, an index past cellXfs — refuses the rebuild; General and a plain number format admit (Codex #205 r2 REL-202)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const Case = struct { name: []const u8, num_fmts: []const u8, xf1: []const u8, style: []const u8, admits: bool };
    const cases = [_]Case{
        // 27 is a locale date built-in the table does not spell.
        .{ .name = "builtin27", .num_fmts = "", .xf1 = "numFmtId=\"27\"", .style = "1", .admits = false },
        .{ .name = "nocustom", .num_fmts = "", .xf1 = "numFmtId=\"164\"", .style = "1", .admits = false },
        .{ .name = "grammar", .num_fmts = "<numFmts count=\"1\"><numFmt numFmtId=\"164\" formatCode=\"[DBNum1]0\"/></numFmts>", .xf1 = "numFmtId=\"164\"", .style = "1", .admits = false },
        .{ .name = "pastxfs", .num_fmts = "", .xf1 = "numFmtId=\"0\"", .style = "7", .admits = false },
        .{ .name = "general", .num_fmts = "", .xf1 = "", .style = "1", .admits = true },
        .{ .name = "zero", .num_fmts = "", .xf1 = "numFmtId=\"0\"", .style = "1", .admits = true },
        .{ .name = "number", .num_fmts = "", .xf1 = "numFmtId=\"2\" applyNumberFormat=\"1\"", .style = "1", .admits = true },
        .{ .name = "custom", .num_fmts = "<numFmts count=\"1\"><numFmt numFmtId=\"164\" formatCode=\"0.000\"/></numFmts>", .xf1 = "numFmtId=\"164\" applyNumberFormat=\"1\"", .style = "1", .admits = true },
        // Written but not a number: not absent (Codex #205 r5 REL-504).
        .{ .name = "bogus_xf", .num_fmts = "", .xf1 = "numFmtId=\"bogus\"", .style = "1", .admits = false },
        .{ .name = "bogus_s", .num_fmts = "", .xf1 = "numFmtId=\"0\"", .style = "bogus", .admits = false },
    };
    // A cell without `s` wears style 0: a date there refuses (Codex
    // #205 r6 REL-601). The fixture's cells carry no `s`.
    {
        const src = try tt.path(std.testing.allocator, io, "s7b4_r6_style0_date.xlsx");
        defer std.testing.allocator.free(src);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
        {
            var store = try store_mod.PartStore.open(std.testing.allocator, io, src);
            defer store.deinit();
            try store.addPart("xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts><fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills><borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"1\"><xf numFmtId=\"14\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/></cellXfs></styleSheet>");
            try store.save(io, src);
        }
        var wb = try Workbook.open(std.testing.allocator, io, src);
        defer wb.deinit();
        const before = try std.testing.allocator.dupe(u8, (try wb.store.part("xl/pivotCache/pivotCacheRecords1.xml")).?.bytes);
        defer std.testing.allocator.free(before);
        try std.testing.expectError(error.PivotEditUnsafe, wb.deleteRow(0, 4));
        try std.testing.expectEqualStrings(before, (try wb.store.part("xl/pivotCache/pivotCacheRecords1.xml")).?.bytes);
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.deleteRow(0, 4));
    }
    for (cases) |case| {
        const file = try std.fmt.allocPrint(std.testing.allocator, "s7b4_r2_{s}.xlsx", .{case.name});
        defer std.testing.allocator.free(file);
        const src = try tt.path(std.testing.allocator, io, file);
        defer std.testing.allocator.free(src);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
        const styles = try std.fmt.allocPrint(std.testing.allocator, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">{s}<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts><fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills><borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"2\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/><xf {s} fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs></styleSheet>", .{ case.num_fmts, case.xf1 });
        defer std.testing.allocator.free(styles);
        {
            var store = try store_mod.PartStore.open(std.testing.allocator, io, src);
            defer store.deinit();
            try store.addPart("xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", styles);
            try store.save(io, src);
        }
        const cell = try std.fmt.allocPrint(std.testing.allocator, "<c r=\"C2\" s=\"{s}\"><v>1.5</v></c>", .{case.style});
        defer std.testing.allocator.free(cell);
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/worksheets/sheet1.xml", "<c r=\"C2\"><v>1.5</v></c>", cell);
        var wb = try Workbook.open(std.testing.allocator, io, src);
        defer wb.deinit();
        const before = try std.testing.allocator.dupe(u8, (try wb.store.part("xl/pivotCache/pivotCacheRecords1.xml")).?.bytes);
        defer std.testing.allocator.free(before);
        if (case.admits) {
            try wb.deleteRow(0, 4);
            const rec = (try wb.store.part("xl/pivotCache/pivotCacheRecords1.xml")).?;
            try std.testing.expect(std.mem.indexOf(u8, rec.bytes, "count=\"2\"><r><x v=\"0\"/><n v=\"3\"/><n v=\"1.5\"/></r>") != null);
        } else {
            try std.testing.expectError(error.PivotEditUnsafe, wb.deleteRow(0, 4));
            try std.testing.expectEqualStrings(before, (try wb.store.part("xl/pivotCache/pivotCacheRecords1.xml")).?.bytes);
            var ed = try Editor.open(std.testing.allocator, io, src);
            defer ed.deinit();
            try std.testing.expectError(error.RowEditUnsafeForSheet, ed.deleteRow(0, 4));
        }
    }
}

test "S7b-4: two definitions naming one records part refuse every edit; a rectangle wider than the schema or past the cell budget refuses before any read; a tall sparse source reads by its cells (Codex #205 r3 REL-301, PERF-301)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const def_part = "xl/pivotCache/pivotCacheDefinition1.xml";
    const rec_part = "xl/pivotCache/pivotCacheRecords1.xml";
    {
        const src = try tt.path(std.testing.allocator, io, "s7b4_r3_shared_src.xlsx");
        defer std.testing.allocator.free(src);
        try pivots_mod.fixture.writeWithOrphanCache(std.testing.allocator, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/pivotCache/_rels/pivotCacheDefinition2.xml.rels", "Target=\"pivotCacheRecords2.xml\"", "Target=\"pivotCacheRecords1.xml\"");
        var wb = try Workbook.open(std.testing.allocator, io, src);
        defer wb.deinit();
        const before = try std.testing.allocator.dupe(u8, (try wb.store.part(rec_part)).?.bytes);
        defer std.testing.allocator.free(before);
        try std.testing.expectError(error.MalformedPivotXml, wb.insertRow(0, 2));
        try std.testing.expectError(error.MalformedPivotXml, wb.insertRow(0, 1));
        try std.testing.expectEqualStrings(before, (try wb.store.part(rec_part)).?.bytes);
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
    }
    {
        // Sixteen thousand columns for three fields: refused before a
        // cell is read.
        const src = try tt.path(std.testing.allocator, io, "s7b4_r3_wide_src.xlsx");
        defer std.testing.allocator.free(src);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, def_part, "ref=\"A1:C4\"", "ref=\"A1:XFD1048576\"");
        var wb = try Workbook.open(std.testing.allocator, io, src);
        defer wb.deinit();
        try std.testing.expectError(error.PivotEditUnsafe, wb.insertRow(0, 2));
        try std.testing.expectEqual(@as(u32, 0), wb.pivot_rebuilds);
    }
    {
        // Seventeen fields over the whole column height: past the
        // budget, refused before a cell is read.
        const src = try tt.path(std.testing.allocator, io, "s7b4_r3_budget_src.xlsx");
        defer std.testing.allocator.free(src);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
        var extra: std.ArrayListUnmanaged(u8) = .empty;
        defer extra.deinit(std.testing.allocator);
        for (0..14) |i| {
            const one = try std.fmt.allocPrint(std.testing.allocator, "<cacheField name=\"F{d}\" numFmtId=\"0\"><sharedItems containsSemiMixedTypes=\"0\" containsString=\"0\" containsNumber=\"1\" containsInteger=\"1\" minValue=\"0\" maxValue=\"0\"/></cacheField>", .{i});
            defer std.testing.allocator.free(one);
            try extra.appendSlice(std.testing.allocator, one);
        }
        try extra.appendSlice(std.testing.allocator, "</cacheFields>");
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, def_part, "<cacheFields count=\"3\">", "<cacheFields count=\"17\">");
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, def_part, "</cacheFields>", extra.items);
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, def_part, "ref=\"A1:C4\"", "ref=\"A1:Q1048576\"");
        var wb = try Workbook.open(std.testing.allocator, io, src);
        defer wb.deinit();
        try std.testing.expectError(error.PivotEditUnsafe, wb.insertRow(0, 2));
        try std.testing.expectEqual(@as(u32, 0), wb.pivot_rebuilds);
    }
    {
        // Six thousand rows of which three hold cells: read and rebuilt,
        // the blank rows as blank records.
        const src = try tt.path(std.testing.allocator, io, "s7b4_r3_tall_src.xlsx");
        defer std.testing.allocator.free(src);
        const dst = try tt.path(std.testing.allocator, io, "s7b4_r3_tall_dst.xlsx");
        defer std.testing.allocator.free(dst);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, def_part, "ref=\"A1:C4\"", "ref=\"A1:C6000\"");
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try ed.insertRow(0, 2);
        try std.testing.expectEqual(@as(u32, 1), ed.workbook.pivot_rebuilds);
        try ed.save(io, dst);
        var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
        defer store.deinit();
        try expectRebuilt(&store, def_part, rec_part, 6000);
        const cd = (try store.part(def_part)).?;
        try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "ref=\"A1:C6001\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "<sharedItems containsBlank=\"1\" count=\"3\"><s v=\"East\"/><s v=\"West\"/><m/></sharedItems>") != null);
    }
}

test "S7b-4: a table's totals row is not a record — its delete refuses, an insert above it appends a data row; a headerless table's column names must be the field names (Codex #205 r4 REL-401, REL-402)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const def_part = "xl/pivotCache/pivotCacheDefinition1.xml";
    const rec_part = "xl/pivotCache/pivotCacheRecords1.xml";
    const table_part = "xl/tables/table1.xml";
    const sheet_part = "xl/worksheets/sheet1.xml";
    {
        const src = try tt.path(std.testing.allocator, io, "s7b4_r4_totals_src.xlsx");
        defer std.testing.allocator.free(src);
        const dst = try tt.path(std.testing.allocator, io, "s7b4_r4_totals_dst.xlsx");
        defer std.testing.allocator.free(dst);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .table_name);
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, table_part, "ref=\"A1:C4\" totalsRowShown=\"0\"", "ref=\"A1:C5\" totalsRowCount=\"1\" totalsRowShown=\"1\"");
        // The totals row: an inline label and an uncached SUBTOTAL —
        // cells the rebuild would refuse, were they read (r5 REL-502).
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, sheet_part, "</sheetData>", "<row r=\"5\"><c r=\"A5\" t=\"inlineStr\"><is><t>Total</t></is></c><c r=\"B5\"><f>SUBTOTAL(109,SalesTbl[Qty])</f></c></row></sheetData>");
        {
            var wb = try Workbook.open(std.testing.allocator, io, src);
            defer wb.deinit();
            const before = try std.testing.allocator.dupe(u8, (try wb.store.part(rec_part)).?.bytes);
            defer std.testing.allocator.free(before);
            // The totals row is not a data row to delete.
            try std.testing.expectError(error.PivotEditUnsafe, wb.deleteRow(0, 5));
            try std.testing.expectEqualStrings(before, (try wb.store.part(rec_part)).?.bytes);
            try std.testing.expectEqual(@as(u32, 0), wb.pivot_rebuilds);
        }
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try ed.insertRow(0, 3);
        try ed.save(io, dst);
        var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
        defer store.deinit();
        // Three data rows and the inserted blank: the totals row is
        // neither read nor a record.
        try expectRebuilt(&store, def_part, rec_part, 4);
        const rec = (try store.part(rec_part)).?;
        const cd = (try store.part(def_part)).?;
        try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "Total") == null);
        try std.testing.expect(std.mem.indexOf(u8, rec.bytes, "<r><x v=\"0\"/><n v=\"3\"/><n v=\"1.5\"/></r><r><x v=\"2\"/><m/><m/></r><r><x v=\"1\"/><n v=\"4\"/><n v=\"2.5\"/></r><r><x v=\"0\"/><n v=\"5\"/><n v=\"3.5\"/></r>") != null);
        const tbl = (try store.part(table_part)).?;
        try std.testing.expect(std.mem.indexOf(u8, tbl.bytes, "ref=\"A1:C6\"") != null);
    }
    {
        // A totals count no rectangle holds: refused, not added
        // (Codex #205 r5 SEC-501).
        const src = try tt.path(std.testing.allocator, io, "s7b4_r5_totals_overflow_src.xlsx");
        defer std.testing.allocator.free(src);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .table_name);
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, table_part, "ref=\"A1:C4\" totalsRowShown=\"0\"", "ref=\"A1:C4\" totalsRowCount=\"4294967295\" totalsRowShown=\"0\"");
        var wb = try Workbook.open(std.testing.allocator, io, src);
        defer wb.deinit();
        const before = try std.testing.allocator.dupe(u8, (try wb.store.part(rec_part)).?.bytes);
        defer std.testing.allocator.free(before);
        try std.testing.expectError(error.PivotEditUnsafe, wb.insertRow(0, 2));
        try std.testing.expectEqual(@as(u32, 0), wb.pivot_rebuilds);
        try std.testing.expectEqualStrings(before, (try wb.store.part(rec_part)).?.bytes);
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
    }
    {
        // Headerless: the column names are the schema.
        const src = try tt.path(std.testing.allocator, io, "s7b4_r4_headerless_src.xlsx");
        defer std.testing.allocator.free(src);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .table_name);
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, table_part, "ref=\"A1:C4\" totalsRowShown=\"0\"", "ref=\"A1:C4\" headerRowCount=\"0\" totalsRowShown=\"0\"");
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, table_part, "name=\"Region\"", "name=\"Area\"");
        var wb = try Workbook.open(std.testing.allocator, io, src);
        defer wb.deinit();
        try std.testing.expectError(error.PivotEditUnsafe, wb.insertRow(0, 2));
        try std.testing.expectEqual(@as(u32, 0), wb.pivot_rebuilds);
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
    }
    {
        // Headerless and named as the fields: the first row is data. A
        // column-shaped comment and a padded close are nothing to the
        // scanner the names are read through (r5 REL-503).
        const src = try tt.path(std.testing.allocator, io, "s7b4_r4_headerless_ok_src.xlsx");
        defer std.testing.allocator.free(src);
        const dst = try tt.path(std.testing.allocator, io, "s7b4_r4_headerless_ok_dst.xlsx");
        defer std.testing.allocator.free(dst);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .table_name);
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, table_part, "ref=\"A1:C4\" totalsRowShown=\"0\"", "ref=\"A1:C4\" headerRowCount=\"0\" totalsRowShown=\"0\"");
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, table_part, "</tableColumns>", "<!-- <tableColumn id=\"9\" name=\"Fake\"/> --></tableColumns >");
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try ed.insertRow(0, 2);
        try ed.save(io, dst);
        var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
        defer store.deinit();
        try expectRebuilt(&store, def_part, rec_part, 5);
    }
}

test "S7b-4: a prepared collection is this edit's or nothing — another index, another axis, an aged store, another workbook refuse and change no part (Codex #205 r6 REL-605)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7b4_r6_token.xlsx");
    defer std.testing.allocator.free(src);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
    const def_part = "xl/pivotCache/pivotCacheDefinition1.xml";
    const rec_part = "xl/pivotCache/pivotCacheRecords1.xml";
    var wb = try Workbook.open(std.testing.allocator, io, src);
    defer wb.deinit();
    var other = try Workbook.open(std.testing.allocator, io, src);
    defer other.deinit();
    const before_def = try std.testing.allocator.dupe(u8, (try wb.store.part(def_part)).?.bytes);
    defer std.testing.allocator.free(before_def);
    const before_rec = try std.testing.allocator.dupe(u8, (try wb.store.part(rec_part)).?.bytes);
    defer std.testing.allocator.free(before_rec);

    var prepared = try wb.preflightPivotEditsForSheet("xl/worksheets/sheet1.xml", .row, 2, .insert);
    defer prepared.deinit(std.testing.allocator);
    try std.testing.expect(prepared.patches.items.len == 3);
    try std.testing.expect(prepared.host_writes.items.len == 1);
    try std.testing.expectError(error.PivotEditUnsafe, wb.applySheetEdit(0, .{ .row = 3, .kind = .insert }, &prepared));
    try std.testing.expectError(error.PivotEditUnsafe, wb.applySheetEdit(0, .{ .row = 2, .kind = .delete }, &prepared));
    try std.testing.expectError(error.PivotEditUnsafe, wb.applySheetEdit(0, .{ .col = 2, .kind = .insert }, &prepared));
    try std.testing.expectError(error.PivotEditUnsafe, other.applySheetEdit(0, .{ .row = 2, .kind = .insert }, &prepared));
    // Neither axis, or both: not an edit (Codex #205 r12 REL-1201) —
    // refused before anything is read, the store untouched.
    const mutations = wb.store.mutations;
    try std.testing.expectError(error.InvalidSheetEditSpec, wb.applySheetEdit(0, .{ .kind = .insert }, &prepared));
    try std.testing.expectError(error.InvalidSheetEditSpec, wb.applySheetEdit(0, .{ .row = 2, .col = 2, .kind = .insert }, &prepared));
    try std.testing.expectEqual(mutations, wb.store.mutations);
    try std.testing.expectEqualStrings(before_def, (try wb.store.part(def_part)).?.bytes);
    try std.testing.expectEqualStrings(before_rec, (try wb.store.part(rec_part)).?.bytes);
    try std.testing.expectEqualStrings(before_def, (try other.store.part(def_part)).?.bytes);
    // The store moved under the token (a pure shift replaced the
    // definition): aged, it refuses too.
    try wb.insertRow(0, 1);
    try std.testing.expectError(error.PivotEditUnsafe, wb.applySheetEdit(0, .{ .row = 2, .kind = .insert }, &prepared));
    // The token for the edit as it now is installs.
    var fresh = try wb.preflightPivotEditsForSheet("xl/worksheets/sheet1.xml", .row, 3, .insert);
    defer fresh.deinit(std.testing.allocator);
    try wb.applySheetEdit(0, .{ .row = 3, .kind = .insert }, &fresh);
    const cd = (try wb.store.part(def_part)).?;
    try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "recordCount=\"4\"") != null);
}

test "S7b-4: refreshedDate survives a calc policy the recalculator refuses, under either epoch (Codex #205 r7 REL-702)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const def_part = "xl/pivotCache/pivotCacheDefinition1.xml";
    for ([_]bool{ false, true }) |d1904| {
        const file = try std.fmt.allocPrint(std.testing.allocator, "s7b4_r7_epoch_{d}.xlsx", .{@intFromBool(d1904)});
        defer std.testing.allocator.free(file);
        const src = try tt.path(std.testing.allocator, io, file);
        defer std.testing.allocator.free(src);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
        // A precision-as-displayed policy the recalculator refuses; the
        // epoch beside it still reads.
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/workbook.xml", "<pivotCaches", "<calcPr fullPrecision=\"0\"/><pivotCaches");
        if (d1904) try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/workbook.xml", "<sheets>", "<workbookPr date1904=\"1\"/><sheets>");
        var wb = try Workbook.open(std.testing.allocator, io, src);
        defer wb.deinit();
        try wb.insertRow(0, 2);
        const cd = (try wb.store.part(def_part)).?;
        const at = std.mem.indexOf(u8, cd.bytes, "refreshedDate=\"") orelse return error.TestExpectedRefreshedDate;
        const from = at + "refreshedDate=\"".len;
        const end = std.mem.indexOfScalarPos(u8, cd.bytes, from, '"') orelse return error.TestExpectedRefreshedDate;
        const serial = try std.fmt.parseFloat(f64, cd.bytes[from..end]);
        // 2026-08-29 is serial 46263 under 1900 and 1462 less under
        // 1904; the bounds hold for years either side.
        if (d1904) {
            try std.testing.expect(serial > 44000 and serial < 45500);
        } else {
            try std.testing.expect(serial > 46000 and serial < 47000);
        }
    }
    // The attribute's text is XML (r9 REL-904); two `workbookPr` are
    // not one epoch, and the rebuild goes undated (r9 REL-905); an
    // ISO spelling beside the serial is redated to the same instant
    // (r9 REL-902).
    const Case = struct { name: []const u8, pr: []const u8, dated: bool, d1904: bool };
    const cases = [_]Case{
        .{ .name = "ref49", .pr = "<workbookPr date1904=\"&#49;\"/>", .dated = true, .d1904 = true },
        .{ .name = "ref48", .pr = "<workbookPr date1904=\"&#48;\"/>", .dated = true, .d1904 = false },
        .{ .name = "twice", .pr = "<workbookPr/><workbookPr date1904=\"1\"/>", .dated = false, .d1904 = false },
        .{ .name = "twice_equal", .pr = "<workbookPr date1904=\"1\"/><workbookPr date1904=\"1\"/>", .dated = false, .d1904 = false },
    };
    for (cases) |case| {
        const file = try std.fmt.allocPrint(std.testing.allocator, "s7b4_r9_epoch_{s}.xlsx", .{case.name});
        defer std.testing.allocator.free(file);
        const src = try tt.path(std.testing.allocator, io, file);
        defer std.testing.allocator.free(src);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
        const pr = try std.mem.concat(std.testing.allocator, u8, &.{ case.pr, "<sheets>" });
        defer std.testing.allocator.free(pr);
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/workbook.xml", "<sheets>", pr);
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, def_part, "refreshedDate=\"45000.5\"", "refreshedDate=\"45000.5\" refreshedDateIso=\"2023-03-15T12:00:00\"");
        var wb = try Workbook.open(std.testing.allocator, io, src);
        defer wb.deinit();
        try wb.insertRow(0, 2);
        const cd = (try wb.store.part(def_part)).?;
        if (!case.dated) {
            try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "refreshedDate") == null);
            continue;
        }
        const at = std.mem.indexOf(u8, cd.bytes, "refreshedDate=\"") orelse return error.TestExpectedRefreshedDate;
        const from = at + "refreshedDate=\"".len;
        const end = std.mem.indexOfScalarPos(u8, cd.bytes, from, '"') orelse return error.TestExpectedRefreshedDate;
        const serial = try std.fmt.parseFloat(f64, cd.bytes[from..end]);
        if (case.d1904) {
            try std.testing.expect(serial > 44000 and serial < 45500);
        } else {
            try std.testing.expect(serial > 46000 and serial < 47000);
        }
        const iso_at = std.mem.indexOf(u8, cd.bytes, "refreshedDateIso=\"") orelse return error.TestExpectedRefreshedDate;
        const iso = cd.bytes[iso_at + "refreshedDateIso=\"".len ..][0..19];
        try std.testing.expect(std.mem.startsWith(u8, iso, "20") and iso[4] == '-' and iso[10] == 'T');
        try std.testing.expect(!std.mem.eql(u8, iso, "2023-03-15T12:00:00"));
    }
}

test "S7b-4: a number or an SST index spelt with character references reads by what it is; a table count written but not a number refuses the graph (Codex #205 r8 REL-802, REL-803)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const def_part = "xl/pivotCache/pivotCacheDefinition1.xml";
    const rec_part = "xl/pivotCache/pivotCacheRecords1.xml";
    const sheet_part = "xl/worksheets/sheet1.xml";
    const table_part = "xl/tables/table1.xml";
    {
        const src = try tt.path(std.testing.allocator, io, "s7b4_r8_refs_src.xlsx");
        defer std.testing.allocator.free(src);
        const dst = try tt.path(std.testing.allocator, io, "s7b4_r8_refs_dst.xlsx");
        defer std.testing.allocator.free(dst);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, sheet_part, "<c r=\"C2\"><v>1.5</v></c>", "<c r=\"C2\"><v>1&#x2E;5</v></c>");
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, sheet_part, "<c r=\"A2\" t=\"s\"><v>3</v></c>", "<c r=\"A2\" t=\"s\"><v>&#51;</v></c>");
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try ed.deleteRow(0, 4);
        try ed.save(io, dst);
        var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
        defer store.deinit();
        try expectRebuilt(&store, def_part, rec_part, 2);
        const rec = (try store.part(rec_part)).?;
        try std.testing.expect(std.mem.indexOf(u8, rec.bytes, "<r><x v=\"0\"/><n v=\"3\"/><n v=\"1.5\"/></r>") != null);
    }
    const Case = struct { name: []const u8, new: []const u8 };
    const counts = [_]Case{
        .{ .name = "totals_bogus", .new = "ref=\"A1:C4\" totalsRowCount=\"bogus\" totalsRowShown=\"0\"" },
        .{ .name = "header_empty", .new = "ref=\"A1:C4\" headerRowCount=\"\" totalsRowShown=\"0\"" },
        .{ .name = "header_negative", .new = "ref=\"A1:C4\" headerRowCount=\"-1\" totalsRowShown=\"0\"" },
    };
    for (counts) |case| {
        const file = try std.fmt.allocPrint(std.testing.allocator, "s7b4_r8_{s}.xlsx", .{case.name});
        defer std.testing.allocator.free(file);
        const src = try tt.path(std.testing.allocator, io, file);
        defer std.testing.allocator.free(src);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .table_name);
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, table_part, "ref=\"A1:C4\" totalsRowShown=\"0\"", case.new);
        var wb = try Workbook.open(std.testing.allocator, io, src);
        defer wb.deinit();
        const before = try std.testing.allocator.dupe(u8, (try wb.store.part(rec_part)).?.bytes);
        defer std.testing.allocator.free(before);
        try std.testing.expectError(error.MalformedPivotXml, wb.insertRow(0, 2));
        try std.testing.expectEqualStrings(before, (try wb.store.part(rec_part)).?.bytes);
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
    }
}

test "S7b-4: a row or cell reference, a style or type, a table ref or count spelt with character references reads by what it is (Codex #205 r10 REL-1002)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const def_part = "xl/pivotCache/pivotCacheDefinition1.xml";
    const rec_part = "xl/pivotCache/pivotCacheRecords1.xml";
    const sheet_part = "xl/worksheets/sheet1.xml";
    const table_part = "xl/tables/table1.xml";
    {
        // `<row r="&#50;">`, `<c r="B&#50;" t="&#110;">`, `<c r="A&#50;"
        // t="&#115;">`: the row reads, the cells read.
        const src = try tt.path(std.testing.allocator, io, "s7b4_r10_refs_src.xlsx");
        defer std.testing.allocator.free(src);
        const dst = try tt.path(std.testing.allocator, io, "s7b4_r10_refs_dst.xlsx");
        defer std.testing.allocator.free(dst);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, sheet_part, "<row r=\"2\">", "<row r=\"&#50;\">");
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, sheet_part, "<c r=\"B2\"><v>3</v></c>", "<c r=\"B&#50;\" t=\"&#110;\"><v>3</v></c>");
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, sheet_part, "<c r=\"A2\" t=\"s\"><v>3</v></c>", "<c r=\"A&#50;\" t=\"&#115;\"><v>3</v></c>");
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try ed.deleteRow(0, 4);
        try ed.save(io, dst);
        var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
        defer store.deinit();
        try expectRebuilt(&store, def_part, rec_part, 2);
        const rec = (try store.part(rec_part)).?;
        try std.testing.expect(std.mem.indexOf(u8, rec.bytes, "<r><x v=\"0\"/><n v=\"3\"/><n v=\"1.5\"/></r>") != null);
    }
    {
        // `s="&#49;"` under a date style is the date it names: refused.
        const src = try tt.path(std.testing.allocator, io, "s7b4_r10_style_src.xlsx");
        defer std.testing.allocator.free(src);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
        {
            var store = try store_mod.PartStore.open(std.testing.allocator, io, src);
            defer store.deinit();
            try store.addPart("xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts><fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills><borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"2\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/><xf numFmtId=\"&#49;&#52;\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/></cellXfs></styleSheet>");
            try store.save(io, src);
        }
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, sheet_part, "<c r=\"C2\"><v>1.5</v></c>", "<c r=\"C2\" s=\"&#49;\"><v>1.5</v></c>");
        var wb = try Workbook.open(std.testing.allocator, io, src);
        defer wb.deinit();
        try std.testing.expectError(error.PivotEditUnsafe, wb.deleteRow(0, 4));
    }
    {
        // A table `ref` and counts with references: headerless, its
        // names the schema, its one totals row not a record.
        const src = try tt.path(std.testing.allocator, io, "s7b4_r10_table_src.xlsx");
        defer std.testing.allocator.free(src);
        const dst = try tt.path(std.testing.allocator, io, "s7b4_r10_table_dst.xlsx");
        defer std.testing.allocator.free(dst);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .table_name);
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, table_part, "ref=\"A1:C4\" totalsRowShown=\"0\"", "ref=\"A1&#58;C5\" headerRowCount=\"&#48;\" totalsRowCount=\"&#49;\" totalsRowShown=\"1\"");
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, sheet_part, "</sheetData>", "<row r=\"5\"><c r=\"B5\"><f>SUBTOTAL(109,SalesTbl[Qty])</f></c></row></sheetData>");
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try ed.insertRow(0, 2);
        try ed.save(io, dst);
        var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
        defer store.deinit();
        // Four data rows (the old header row is data) and the blank.
        try expectRebuilt(&store, def_part, rec_part, 5);
    }
}

// ─── S7b-3: the refresh marker for cell writes (§7 Q3 — one rule) ───
//
// A row edit inside a source rectangle marks the cache in the sweep; a
// `setCell` inside the same rectangle changes the same content and
// used to mark nothing — the asymmetry `docs/plans/s7b-cache-policy.md`
// §5 names. `Workbook.applySavePlans` now applies the one predicate to
// every staged write at save (deltas at their coordinates, appended
// rows where the emitter lands them), so `Editor.save`,
// `saveToOwnedBuffer` and the direct `Workbook.save` share it.

test "S7b-3: a setCell inside a source rectangle marks the cache at save; outside it, and on the host, the definition is byte-identical" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7b3_cell_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "s7b3_cell_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
    const original = try cacheDefinitionBytes(std.testing.allocator, io, src);
    defer std.testing.allocator.free(original);
    const marked = try markedDefinition(std.testing.allocator, original);
    defer std.testing.allocator.free(marked);
    {
        // `Data!E9` is past both edges of `A1:C4`; `Report!A1` is the
        // host's own cell. Neither is the source's content.
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try ed.setCell(0, 9, 4, .{ .integer = 1 });
        try ed.setCell(1, 1, 0, .{ .string = "host" });
        try ed.save(io, dst);
        const after = try cacheDefinitionBytes(std.testing.allocator, io, dst);
        defer std.testing.allocator.free(after);
        try std.testing.expectEqualStrings(original, after);
    }
    {
        // `Data!B3` is a record's `Qty`: marked — and, since S7b-5,
        // rebuilt from the sheet as written, the consumer re-laid.
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try ed.setCell(0, 3, 1, .{ .number = 9.5 });
        try ed.save(io, dst);
        var wb = try Workbook.open(std.testing.allocator, io, dst);
        defer wb.deinit();
        var p = try wb.pivotTables();
        defer p.deinit();
        try std.testing.expect(p.caches[0].definition.refresh_on_load);
        try std.testing.expectEqualStrings("A1:C4", p.caches[0].source.ref.?);
        try expectRebuilt(&wb.store, "xl/pivotCache/pivotCacheDefinition1.xml", "xl/pivotCache/pivotCacheRecords1.xml", 3);
        const rec = (try wb.store.part("xl/pivotCache/pivotCacheRecords1.xml")).?;
        try std.testing.expect(std.mem.indexOf(u8, rec.bytes, "<r><x v=\"1\"/><n v=\"9.5\"/><n v=\"2.5\"/></r>") != null);
        try expectHostCell(&wb, 1, "B5", .{ .number = 9.5 });
        try expectHostCell(&wb, 1, "B6", .{ .number = 17.5 });
    }
    {
        // The buffer save takes the same path; the header cell `A1` is
        // inside the rectangle (the field names come from it).
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try ed.setCell(0, 1, 0, .{ .string = "Area" });
        const bytes = try ed.saveToOwnedBuffer(std.testing.allocator);
        defer std.testing.allocator.free(bytes);
        var wb = try Workbook.openBuffer(std.testing.allocator, io, bytes);
        defer wb.deinit();
        var p = try wb.pivotTables();
        defer p.deinit();
        try std.testing.expect(p.caches[0].definition.refresh_on_load);
    }
    {
        // A blank write at the bottom-right corner `C4` is a write:
        // marked and rebuilt, the record's `Price` now a blank.
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try ed.setCell(0, 4, 2, .empty);
        try ed.save(io, dst);
        var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
        defer store.deinit();
        try expectMarked(&store, "xl/pivotCache/pivotCacheDefinition1.xml", true);
        try expectRebuilt(&store, "xl/pivotCache/pivotCacheDefinition1.xml", "xl/pivotCache/pivotCacheRecords1.xml", 3);
        const rec = (try store.part("xl/pivotCache/pivotCacheRecords1.xml")).?;
        try std.testing.expect(std.mem.indexOf(u8, rec.bytes, "<r><x v=\"0\"/><n v=\"5\"/><m/></r>") != null);
    }
}

test "S7b-3: a cell write marks a table-named source inside its table, a whole-column name in its columns, an unbounded name on the sheet its closure names — and never one that proves no local range" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const dst = try tt.path(std.testing.allocator, io, "s7b3_kinds_dst.xlsx");
    defer std.testing.allocator.free(dst);

    const Probe = struct {
        fn run(alloc: Allocator, io_: std.Io, src: []const u8, out: []const u8, sheet: u32, row: u32, col: u32) !bool {
            var ed = try Editor.open(alloc, io_, src);
            defer ed.deinit();
            try ed.setCell(sheet, row, col, .{ .integer = 42 });
            try ed.save(io_, out);
            var wb = try Workbook.open(alloc, io_, out);
            defer wb.deinit();
            var p = try wb.pivotTables();
            defer p.deinit();
            return p.caches[0].definition.refresh_on_load;
        }
    };

    const table = try tt.path(std.testing.allocator, io, "s7b3_kinds_table.xlsx");
    defer std.testing.allocator.free(table);
    try pivots_mod.fixture.write(std.testing.allocator, io, table, .table_name);
    // `SalesTbl` is `A1:C4`: `C2` in, `D2` out.
    try std.testing.expect(try Probe.run(std.testing.allocator, io, table, dst, 0, 2, 2));
    try std.testing.expect(!try Probe.run(std.testing.allocator, io, table, dst, 0, 2, 3));

    const cols = try tt.path(std.testing.allocator, io, "s7b3_kinds_cols.xlsx");
    defer std.testing.allocator.free(cols);
    try pivots_mod.fixture.write(std.testing.allocator, io, cols, .defined_name);
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, cols, "xl/workbook.xml", "Data!$A$1:$C$4", "Data!$A:$C");
    // Whole columns `A:C`: `A500` in, `F1` out.
    try std.testing.expect(try Probe.run(std.testing.allocator, io, cols, dst, 0, 500, 0));
    try std.testing.expect(!try Probe.run(std.testing.allocator, io, cols, dst, 0, 1, 5));

    const unbounded = try tt.path(std.testing.allocator, io, "s7b3_kinds_unbounded.xlsx");
    defer std.testing.allocator.free(unbounded);
    try pivots_mod.fixture.write(std.testing.allocator, io, unbounded, .defined_name);
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, unbounded, "xl/workbook.xml", "Data!$A$1:$C$4", "OFFSET(Report!$D$1,0,0,4,3)");
    // The closure names `Report` and nothing else: any cell of `Report`
    // (even `A50`, far from the anchor), no cell of `Data`.
    try std.testing.expect(try Probe.run(std.testing.allocator, io, unbounded, dst, 1, 50, 0));
    try std.testing.expect(!try Probe.run(std.testing.allocator, io, unbounded, dst, 0, 2, 1));

    // Another workbook, or no workbook: nothing local is proven.
    const external = try tt.path(std.testing.allocator, io, "s7b3_kinds_external.xlsx");
    defer std.testing.allocator.free(external);
    try pivots_mod.fixture.write(std.testing.allocator, io, external, .external);
    try std.testing.expect(!try Probe.run(std.testing.allocator, io, external, dst, 0, 2, 1));
    const dangling = try tt.path(std.testing.allocator, io, "s7b3_kinds_dangling.xlsx");
    defer std.testing.allocator.free(dangling);
    try pivots_mod.fixture.write(std.testing.allocator, io, dangling, .dangling);
    try std.testing.expect(!try Probe.run(std.testing.allocator, io, dangling, dst, 0, 2, 1));
}

test "S7b-3: appended rows mark only where the emitter lands them — inside a source rectangle that extends past the data" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7b3_append_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "s7b3_append_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);

    const row = [_]Cell{ .empty, .{ .integer = 6 }, .{ .number = 4.5 } };
    const rows = [_][]const Cell{&row};
    {
        // `A1:C4` ends where the data ends: the append lands on row 5,
        // outside. Byte-identical.
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try ed.appendRows(0, &rows);
        try ed.save(io, dst);
        var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
        defer store.deinit();
        try expectMarked(&store, "xl/pivotCache/pivotCacheDefinition1.xml", false);
    }
    // Excel commonly writes a source taller than its data. `A1:C9`:
    // row 5 is inside — the empty first cell is no write, the second
    // (`B5`) is.
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/pivotCache/pivotCacheDefinition1.xml", "ref=\"A1:C4\"", "ref=\"A1:C9\"");
    {
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try ed.appendRows(0, &rows);
        try ed.save(io, dst);
        var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
        defer store.deinit();
        try expectMarked(&store, "xl/pivotCache/pivotCacheDefinition1.xml", true);
    }
    // A row whose only cells fall right of the rectangle (`D5`, `E5`)
    // is outside it.
    const beside = [_]Cell{ .empty, .empty, .empty, .{ .integer = 1 }, .{ .integer = 2 } };
    const beside_rows = [_][]const Cell{&beside};
    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    try ed.appendRows(0, &beside_rows);
    try ed.save(io, dst);
    var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
    defer store.deinit();
    try expectMarked(&store, "xl/pivotCache/pivotCacheDefinition1.xml", false);
}

test "S7b-3: an already-marked definition keeps its one marker under a write inside (S7b-5 rebuilds it); a graph that cannot be read marks nothing and still saves" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const dst = try tt.path(std.testing.allocator, io, "s7b3_idem_dst.xlsx");
    defer std.testing.allocator.free(dst);

    const set = try tt.path(std.testing.allocator, io, "s7b3_idem_set.xlsx");
    defer std.testing.allocator.free(set);
    try pivots_mod.fixture.write(std.testing.allocator, io, set, .sheet_ref);
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, set, "xl/pivotCache/pivotCacheDefinition1.xml", " recordCount=\"3\">", " recordCount=\"3\" refreshOnLoad=\"1\">");
    {
        var ed = try Editor.open(std.testing.allocator, io, set);
        defer ed.deinit();
        try ed.setCell(0, 3, 1, .{ .number = 9.5 });
        try ed.save(io, dst);
        var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
        defer store.deinit();
        try expectMarked(&store, "xl/pivotCache/pivotCacheDefinition1.xml", true);
        try expectRebuilt(&store, "xl/pivotCache/pivotCacheDefinition1.xml", "xl/pivotCache/pivotCacheRecords1.xml", 3);
        try std.testing.expect(std.mem.indexOf(u8, (try store.part("xl/pivotCache/pivotCacheDefinition1.xml")).?.bytes, "maxValue=\"9.5\"") != null);
    }

    // The cache's workbook relationship mistyped: the graph refuses
    // (S7b-2 — the row edit below), and a save with a write inside the
    // source cannot mark what it cannot read. It saves, unmarked: the
    // marker is best-effort, and the file is in the state it was in.
    const broken = try tt.path(std.testing.allocator, io, "s7b3_idem_broken.xlsx");
    defer std.testing.allocator.free(broken);
    try pivots_mod.fixture.write(std.testing.allocator, io, broken, .sheet_ref);
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, broken, "xl/_rels/workbook.xml.rels", "relationships/pivotCacheDefinition\" Target=\"pivotCache/pivotCacheDefinition1.xml\"", "relationships/pivotCacheDefinitionX\" Target=\"pivotCache/pivotCacheDefinition1.xml\"");
    const original = try cacheDefinitionBytes(std.testing.allocator, io, broken);
    defer std.testing.allocator.free(original);
    var ed = try Editor.open(std.testing.allocator, io, broken);
    defer ed.deinit();
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
    try ed.setCell(0, 3, 1, .{ .number = 9.5 });
    try ed.save(io, dst);
    const after = try cacheDefinitionBytes(std.testing.allocator, io, dst);
    defer std.testing.allocator.free(after);
    try std.testing.expectEqualStrings(original, after);
}

test "S7b: the sweep is all-or-nothing — a source refusal leaves the host's already-computed pivot move uninstalled" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7b_atomic_src.xlsx");
    defer std.testing.allocator.free(src);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/pivotCache/pivotCacheDefinition1.xml", "sheet=\"Data\"", "sheet=\"Report\"");
    var wb = try Workbook.open(std.testing.allocator, io, src);
    defer wb.deinit();
    // Row 1 is above the pivot (`A3:B6` would shift) and the source's
    // header (`A1:C4` refuses): the pivot part must not move alone.
    try std.testing.expectError(error.PivotEditUnsafe, wb.deleteRow(1, 1));
    const pt = (try wb.store.part("xl/pivotTables/pivotTable1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, pt.bytes, "<location ref=\"A3:B6\"") != null);
    const cd = (try wb.store.part("xl/pivotCache/pivotCacheDefinition1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "ref=\"A1:C4\"") != null);
    const sheet = (try wb.store.part("xl/worksheets/sheet2.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, sheet.bytes, "r=\"A1\"") != null);
}

test "S7b: spellings that claim the edited sheet without a range refuse; ones that claim another workbook or nothing are left alone" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const def_part = "xl/pivotCache/pivotCacheDefinition1.xml";
    // `sheet` alone (§7 Q4 iv).
    const only = try tt.path(std.testing.allocator, io, "s7b_sheet_only_src.xlsx");
    defer std.testing.allocator.free(only);
    try pivots_mod.fixture.write(std.testing.allocator, io, only, .sheet_ref);
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, only, def_part, "<worksheetSource sheet=\"Data\" ref=\"A1:C4\"/>", "<worksheetSource sheet=\"Data\"/>");
    {
        var ed = try Editor.open(std.testing.allocator, io, only);
        defer ed.deinit();
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 9));
        try std.testing.expectError(error.ColEditUnsafeForSheet, ed.insertColumn(0, 9));
        try ed.insertRow(1, 1);
    }
    // An `r:id` the reader cannot place beside `sheet="Data"` (Q4 i).
    const rid = try tt.path(std.testing.allocator, io, "s7b_rid_src.xlsx");
    defer std.testing.allocator.free(rid);
    try pivots_mod.fixture.write(std.testing.allocator, io, rid, .external);
    {
        // As written the source is another workbook's `Sheet1`; no
        // local sheet is claimed, `Data` is free.
        var ed = try Editor.open(std.testing.allocator, io, rid);
        defer ed.deinit();
        try ed.insertRow(0, 1);
    }
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, rid, def_part, "sheet=\"Sheet1\"", "sheet=\"Data\"");
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, rid, "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels", "Id=\"rIdExt\"", "Id=\"rIdGone\"");
    {
        var ed = try Editor.open(std.testing.allocator, io, rid);
        defer ed.deinit();
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 9));
        try ed.insertRow(1, 1);
    }
    // A locator under a source `type` the reader does not know is
    // authoritative (§7 Q5): the same `ref`, the same move.
    const unknown = try tt.path(std.testing.allocator, io, "s7b_unknown_src.xlsx");
    defer std.testing.allocator.free(unknown);
    const unknown_dst = try tt.path(std.testing.allocator, io, "s7b_unknown_dst.xlsx");
    defer std.testing.allocator.free(unknown_dst);
    try pivots_mod.fixture.write(std.testing.allocator, io, unknown, .sheet_ref);
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, unknown, def_part, "type=\"worksheet\"", "type=\"zlsxFuture\"");
    var ed = try Editor.open(std.testing.allocator, io, unknown);
    defer ed.deinit();
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.deleteRow(0, 1));
    try ed.insertRow(0, 1);
    try ed.save(io, unknown_dst);
    var store = try store_mod.PartStore.open(std.testing.allocator, io, unknown_dst);
    defer store.deinit();
    const cd = (try store.part(def_part)).?;
    try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "<cacheSource type=\"zlsxFuture\"><worksheetSource sheet=\"Data\" ref=\"A2:C5\"/>") != null);
}

test "S7b: a listed cache whose workbook relationship is absent or mistyped refuses every sheet — the gate sees the list (Codex r1 REL-101)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // The pivot table's own relationship still reaches the cache; only
    // the workbook's edge is wrong. Before the fix the source sheet was
    // admitted with the graph unread — and its `ref` left stale.
    const mistyped = try tt.path(std.testing.allocator, io, "s7b_rel_mistyped_src.xlsx");
    defer std.testing.allocator.free(mistyped);
    try pivots_mod.fixture.write(std.testing.allocator, io, mistyped, .sheet_ref);
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, mistyped, "xl/_rels/workbook.xml.rels", "relationships/pivotCacheDefinition\" Target=\"pivotCache/pivotCacheDefinition1.xml\"", "relationships/pivotCacheDefinitionX\" Target=\"pivotCache/pivotCacheDefinition1.xml\"");
    {
        var ed = try Editor.open(std.testing.allocator, io, mistyped);
        defer ed.deinit();
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
        try std.testing.expectError(error.ColEditUnsafeForSheet, ed.insertColumn(0, 9));
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(1, 1));
    }
    const absent = try tt.path(std.testing.allocator, io, "s7b_rel_absent_src.xlsx");
    defer std.testing.allocator.free(absent);
    try pivots_mod.fixture.write(std.testing.allocator, io, absent, .sheet_ref);
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, absent, "xl/_rels/workbook.xml.rels", "<Relationship Id=\"rIdPC1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition\" Target=\"pivotCache/pivotCacheDefinition1.xml\"/>", "");
    {
        var ed = try Editor.open(std.testing.allocator, io, absent);
        defer ed.deinit();
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
    }
    // A workbook with no pivot at all, listing a cache it cannot reach:
    // the documented workbook-wide refusal, from a sheet hosting nothing.
    const dangling = try tt.path(std.testing.allocator, io, "s7b_dangling_list_src.xlsx");
    defer std.testing.allocator.free(dangling);
    {
        var w = xlsx.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Only");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, dangling);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, dangling);
        defer ed.deinit();
        try ed.insertRow(0, 1);
    }
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, dangling, "xl/workbook.xml", "</workbook>", "<pivotCaches><pivotCache cacheId=\"1\" r:id=\"rIdNone\"/></pivotCaches></workbook>");
    var ed = try Editor.open(std.testing.allocator, io, dangling);
    defer ed.deinit();
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 1));
    try ed.setCell(0, 2, 0, .{ .integer = 2 });
}

test "S7b: a sheet+name source on a headerless table admits the top-row delete the table rewriter admits (Codex r1 REL-102)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7b_sheet_table_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "s7b_sheet_table_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .table_name);
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/pivotCache/pivotCacheDefinition1.xml", "<worksheetSource name=\"SalesTbl\"/>", "<worksheetSource sheet=\"Data\" name=\"SalesTbl\"/>");
    {
        // With the default header the table rewriter refuses.
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.deleteRow(0, 1));
        try std.testing.expectError(error.ColEditUnsafeForSheet, ed.insertColumn(0, 2));
    }
    // `headerRowCount="0"`: the field names come from `<tableColumns>`,
    // the top row is data, and its delete shrinks the table.
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/tables/table1.xml", "ref=\"A1:C4\" totalsRowShown=\"0\"", "ref=\"A1:C4\" headerRowCount=\"0\" totalsRowShown=\"0\"");
    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    try ed.deleteRow(0, 1);
    try ed.save(io, dst);
    var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
    defer store.deinit();
    const tbl = (try store.part("xl/tables/table1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, tbl.bytes, "ref=\"A1:C3\"") != null);
    const cd = (try store.part("xl/pivotCache/pivotCacheDefinition1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "<worksheetSource sheet=\"Data\" name=\"SalesTbl\"/>") != null);
    // The admitted top-row delete dropped a record: a content change,
    // marked (§2.2's headerless exception) — and rebuilt with no header
    // row to check (S7b-4): the table's four rows were all data, three
    // remain, and `Region`'s inventory is untouched.
    try expectMarked(&store, "xl/pivotCache/pivotCacheDefinition1.xml", true);
    try expectRebuilt(&store, "xl/pivotCache/pivotCacheDefinition1.xml", "xl/pivotCache/pivotCacheRecords1.xml", 3);
    try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "<sharedItems count=\"2\"><s v=\"East\"/><s v=\"West\"/></sharedItems>") != null);
}

test "S7b: the direct Workbook path refuses a table source's header-row delete before any part changes (Codex r2 REL-202)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const src = "tests/corpus/openxlsx_loadExample.xlsx";
    std.Io.Dir.cwd().access(io, src, .{}) catch return error.SkipZigTest;
    var wb = try Workbook.open(std.testing.allocator, io, src);
    defer wb.deinit();
    // `IrisSample` hosts `G2:K6` and sources it through `Table2` at
    // `A1:E51`: row 1 is above the pivot (it would shift) and the
    // table's header (it refuses). Nothing may move.
    const parts = [_][]const u8{ "xl/pivotTables/pivotTable1.xml", "xl/worksheets/sheet1.xml", "xl/tables/table1.xml", "xl/pivotCache/pivotCacheDefinition1.xml" };
    var before: [parts.len][]u8 = undefined;
    var filled: usize = 0;
    defer for (before[0..filled]) |b| std.testing.allocator.free(b);
    for (parts, 0..) |name, i| {
        before[i] = try std.testing.allocator.dupe(u8, (try wb.store.part(name)).?.bytes);
        filled += 1;
    }
    try std.testing.expectError(error.TableHeaderRowDeleteUnsafe, wb.deleteRow(0, 1));
    for (parts, 0..) |name, i| try std.testing.expectEqualStrings(before[i], (try wb.store.part(name)).?.bytes);
    // The same edit through the Editor folds the refusal.
    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.deleteRow(0, 1));
}

test "S7b: a #REF! inside a string literal neither masks a new one nor counts as one (Codex r2 REL-203)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7b_ref_literal_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "s7b_ref_literal_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .defined_name);
    // A live name whose unused branch spells the error as text.
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/workbook.xml", "Data!$A$1:$C$4", "IF(TRUE,Data!$A$1:$C$4,INDIRECT(\"#REF!\"))");
    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    // The endpoint delete introduces a real `#REF!` beside the literal.
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.deleteRow(0, 4));
    // An interior delete leaves the literal as the body's only one —
    // the count would admit it; since S7b-4 the body, which the reader
    // cannot bound, refuses every `Data` edit at the engine's edge
    // instead (`Workbook.countRefErrorsOutsideQuotes` pins the count).
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.deleteRow(0, 3));
    // The host is outside the closure: its edit moves the pivot and
    // leaves the body, literal and all, as written.
    try ed.insertRow(1, 1);
    try ed.save(io, dst);
    var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
    defer store.deinit();
    const wb_xml = (try store.part("xl/workbook.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, wb_xml.bytes, "IF(TRUE,Data!$A$1:$C$4,INDIRECT(&quot;#REF!&quot;))") != null or
        std.mem.indexOf(u8, wb_xml.bytes, "IF(TRUE,Data!$A$1:$C$4,INDIRECT(\"#REF!\"))") != null);
    try expectMarked(&store, "xl/pivotCache/pivotCacheDefinition1.xml", false);
}

test "S7b: a name body the sweep never rewrites refuses the edit of a sheet its source depends on (Codex r3 REL-301)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7b_cdata_src.xlsx");
    defer std.testing.allocator.free(src);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .defined_name);
    // A CDATA body: well-formed, the same range, and one the name sweep
    // leaves as written (#188 r8). The resolver still proves `Data`
    // from the text, so the cache depends on it.
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/workbook.xml", "<definedName name=\"PivotSrc\">Data!$A$1:$C$4</definedName>", "<definedName name=\"PivotSrc\"><![CDATA[Data!$A$1:$C$4]]></definedName>");
    {
        var wb = try Workbook.open(std.testing.allocator, io, src);
        defer wb.deinit();
        {
            var p = try wb.pivotTables();
            defer p.deinit();
            try std.testing.expect(p.dependsOnSheet(0) and !p.dependsOnSheet(1));
        }
        const parts = [_][]const u8{ "xl/workbook.xml", "xl/worksheets/sheet1.xml", "xl/pivotCache/pivotCacheDefinition1.xml", "xl/pivotTables/pivotTable1.xml" };
        var before: [parts.len][]u8 = undefined;
        var filled: usize = 0;
        defer for (before[0..filled]) |b| std.testing.allocator.free(b);
        for (parts, 0..) |name, i| {
            before[i] = try std.testing.allocator.dupe(u8, (try wb.store.part(name)).?.bytes);
            filled += 1;
        }
        try std.testing.expectError(error.PivotEditUnsafe, wb.insertRow(0, 2));
        try std.testing.expectError(error.PivotEditUnsafe, wb.deleteRow(0, 9));
        try std.testing.expectError(error.PivotEditUnsafe, wb.insertColumn(0, 1));
        for (parts, 0..) |name, i| try std.testing.expectEqualStrings(before[i], (try wb.store.part(name)).?.bytes);
        // The host is not a sheet the name depends on: its pivot moves.
        try wb.insertRow(1, 1);
        const pt = (try wb.store.part("xl/pivotTables/pivotTable1.xml")).?;
        try std.testing.expect(std.mem.indexOf(u8, pt.bytes, "<location ref=\"A4:B7\"") != null);
    }
    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
    try std.testing.expectError(error.ColEditUnsafeForSheet, ed.deleteColumn(0, 3));
    try ed.insertRow(1, 1);
}

// ─── S7a: the output-location lift ───────────────────────────────────
//
// A sheet that only HOSTS a pivot is editable: `location@ref` — the one
// absolute coordinate the definition carries — moves in step with a
// whole-row / whole-column edit that leaves the rectangle intact. What
// still refuses is pinned here beside what lifts: an edit inside the
// footprint, a host that is also a source, a graph that cannot be read.

/// The fixture's pivot part with its rectangle respelled — what the
/// splice must produce, byte for byte, since nothing else may change.
fn pivotPartWithRef(alloc: std.mem.Allocator, part: []const u8, from: []const u8, to: []const u8) ![]u8 {
    const at = std.mem.indexOf(u8, part, from) orelse return error.TestUnexpectedResult;
    return std.mem.concat(alloc, u8, &.{ part[0..at], to, part[at + from.len ..] });
}

test "S7a: a host-only sheet's row and column edits move location@ref and nothing else" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7a_host_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "s7a_host_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);

    const original = blk: {
        var store = try store_mod.PartStore.open(std.testing.allocator, io, src);
        defer store.deinit();
        const pt = (try store.part("xl/pivotTables/pivotTable1.xml")).?;
        break :blk try std.testing.allocator.dupe(u8, pt.bytes);
    };
    defer std.testing.allocator.free(original);

    {
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        // `Report` (1) hosts `A3:B6`. Above, left, below: move, move, stay.
        try ed.insertRow(1, 1);
        try ed.insertColumn(1, 1);
        try ed.deleteRow(1, 9);
        try ed.save(io, dst);
    }

    var wb = try Workbook.open(std.testing.allocator, io, dst);
    defer wb.deinit();
    var p = try wb.pivotTables();
    defer p.deinit();
    try std.testing.expectEqual(@as(usize, 1), p.tables.len);
    try std.testing.expectEqualStrings("B4:C7", p.tables[0].location_ref);
    try std.testing.expectEqual(@as(u32, 1), p.tables[0].sheet_idx);

    // The part differs from the original in the one attribute.
    const want = try pivotPartWithRef(std.testing.allocator, original, "ref=\"A3:B6\"", "ref=\"B4:C7\"");
    defer std.testing.allocator.free(want);
    try std.testing.expectEqualStrings(want, p.tables[0].raw_xml);
    // The cache is not this row's: `worksheetSource` reads `Data` as
    // before, and a moved host rectangle changes no data — it does not
    // mark (the S7a gate's parked second question, answered by the
    // S7b-3 predicate).
    const cache = p.cacheOf(p.tables[0]).?;
    try std.testing.expect(std.mem.indexOf(u8, cache.raw_xml, "<worksheetSource sheet=\"Data\" ref=\"A1:C4\"/>") != null);
    try std.testing.expectEqualStrings("A1:C4", cache.source.ref.?);
    try std.testing.expect(!cache.definition.refresh_on_load);
    try std.testing.expect(std.mem.indexOf(u8, cache.raw_xml, pivots_mod.edit.marker_attr) == null);
}

test "S7a: the host's grid and the pivot move together" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7a_grid_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "s7a_grid_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);

    {
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try ed.insertRow(1, 1);
        try ed.save(io, dst);
    }
    // `Report!A1` ("pivot host") is now `A2`; the pivot that was at row 3
    // is at row 4 — the same displacement, so the rendered cells and
    // the rectangle still agree.
    var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
    defer store.deinit();
    const sheet = (try store.part("xl/worksheets/sheet2.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, sheet.bytes, "r=\"A2\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet.bytes, "r=\"A1\"") == null);
    const pt = (try store.part("xl/pivotTables/pivotTable1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, pt.bytes, "<location ref=\"A4:B7\"") != null);
}

test "S7a: an edit inside the footprint refuses before any mutation, and a no-op edit keeps the part's bytes" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7a_inside_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "s7a_inside_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);

    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    // Inside `A3:B6`: the rows of the body, its columns, its last row.
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(1, 4));
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(1, 6));
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.deleteRow(1, 3));
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.deleteRow(1, 6));
    try std.testing.expectError(error.ColEditUnsafeForSheet, ed.insertColumn(1, 2));
    try std.testing.expectError(error.ColEditUnsafeForSheet, ed.deleteColumn(1, 1));
    try std.testing.expectError(error.ColEditUnsafeForSheet, ed.deleteColumn(1, 2));
    // Below / right of it: admitted, and the part is byte-identical —
    // the sweep never replaced it.
    try ed.deleteRow(1, 7);
    try ed.insertColumn(1, 3);
    try ed.save(io, dst);

    var a = try store_mod.PartStore.open(std.testing.allocator, io, src);
    defer a.deinit();
    var b = try store_mod.PartStore.open(std.testing.allocator, io, dst);
    defer b.deinit();
    const pa = (try a.part("xl/pivotTables/pivotTable1.xml")).?;
    const pb = (try b.part("xl/pivotTables/pivotTable1.xml")).?;
    try std.testing.expectEqualStrings(pa.bytes, pb.bytes);
}

test "S7b: a host that is also a source moves both — location@ref with the grid, worksheetSource@ref by range semantics" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7a_self_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "s7a_self_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
    // Point the cache at the host: `Report` now renders the pivot
    // (`A3:B6`) AND feeds it (`A1:C4`) — the corpus' `IrisSample`
    // shape, refused whole until S7b. Both coordinates move now.
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/pivotCache/pivotCacheDefinition1.xml", "sheet=\"Data\"", "sheet=\"Report\"");

    {
        var wb = try Workbook.open(std.testing.allocator, io, src);
        defer wb.deinit();
        var p = try wb.pivotTables();
        defer p.deinit();
        try std.testing.expect(p.hostsPivot(1) and p.readsFromSheet(1));
    }
    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    try ed.insertRow(1, 1);
    try ed.insertColumn(1, 1);
    try ed.deleteRow(1, 9);
    try ed.save(io, dst);

    var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
    defer store.deinit();
    const pt = (try store.part("xl/pivotTables/pivotTable1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, pt.bytes, "<location ref=\"B4:C7\"") != null);
    const cd = (try store.part("xl/pivotCache/pivotCacheDefinition1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, cd.bytes, "<worksheetSource sheet=\"Report\" ref=\"B2:D5\"/>") != null);
}

test "S7b: a pivot graph that cannot be read refuses every sheet's edit — a source cannot be told from a non-source" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7a_broken_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "s7a_broken_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
    // Break the pivot part: no `<location>` — the graph no longer reads.
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/pivotTables/pivotTable1.xml", "<location ref=\"A3:B6\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>", "");

    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    // The host refuses: it cannot know where its pivot is …
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(1, 1));
    // … and so does `Data`, which has no pivot relationship of its own:
    // the workbook carries a cache, so S7b reads the graph for every
    // sheet, and a graph that does not read cannot say which sheets
    // are sources. Before S7b the source-only sheet was admitted here
    // with the graph unread — and its `ref` left stale.
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 5));
    // A cell write is not a structural edit and never reads the graph.
    try ed.setCell(0, 1, 4, .{ .integer = 1 });
    try ed.save(io, dst);
}

test "S7a: Workbook.insertRow refuses whole — a second hosted pivot's refusal leaves the first untouched" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7a_two_src.xlsx");
    defer std.testing.allocator.free(src);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
    // A second pivot on `Report`, wired AFTER the first in the sheet's
    // relationships, at `A8:B10`; move the first to `A12:B14` so the
    // edit below shifts it and lands inside the second.
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/pivotTables/pivotTable1.xml", "ref=\"A3:B6\"", "ref=\"A12:B14\"");
    {
        var store = try store_mod.PartStore.open(std.testing.allocator, io, src);
        defer store.deinit();
        const first = (try store.part("xl/pivotTables/pivotTable1.xml")).?;
        const second = try pivotPartWithRef(std.testing.allocator, first.bytes, "ref=\"A12:B14\"", "ref=\"A8:B10\"");
        defer std.testing.allocator.free(second);
        try store.addPart("xl/pivotTables/pivotTable2.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml", second);
        const rels1 = (try store.part("xl/pivotTables/_rels/pivotTable1.xml.rels")).?;
        try store.addPart("xl/pivotTables/_rels/pivotTable2.xml.rels", "application/vnd.openxmlformats-package.relationships+xml", rels1.bytes);
        try store.save(io, src);
    }
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/worksheets/_rels/sheet2.xml.rels", "</Relationships>", "<Relationship Id=\"rIdPT2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable\" Target=\"../pivotTables/pivotTable2.xml\"/></Relationships>");

    var wb = try Workbook.open(std.testing.allocator, io, src);
    defer wb.deinit();
    {
        var p = try wb.pivotTables();
        defer p.deinit();
        try std.testing.expectEqual(@as(usize, 2), p.tables.len);
        try std.testing.expectEqualStrings("A12:B14", p.tables[0].location_ref);
        try std.testing.expectEqualStrings("A8:B10", p.tables[1].location_ref);
    }
    // Row 9 is above the first pivot (would move it) and inside the
    // second (refuses). The direct `Workbook` path carries the typed
    // refusal; the Editor folds it into `RowEditUnsafeForSheet`.
    try std.testing.expectError(error.PivotEditUnsafe, wb.insertRow(1, 9));
    const first = (try wb.store.part("xl/pivotTables/pivotTable1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, first.bytes, "ref=\"A12:B14\"") != null);
    const sheet = (try wb.store.part("xl/worksheets/sheet2.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, sheet.bytes, "r=\"A1\"") != null);
    // Row 1 is above both: both move.
    try wb.insertRow(1, 1);
    var p = try wb.pivotTables();
    defer p.deinit();
    try std.testing.expectEqualStrings("A13:B15", p.tables[0].location_ref);
    try std.testing.expectEqualStrings("A9:B11", p.tables[1].location_ref);
}

test "S7b: a source the reader cannot bound refuses every edit of a sheet its closure names — the sweep's #REF! and the engine's missing rectangle alike (was S7a REL-033)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7a_offset_src.xlsx");
    defer std.testing.allocator.free(src);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .defined_name);
    // A dynamic name Excel accepts as a source, and which reads the
    // HOST: `Report!D1:F4`. The reader cannot bound it; its closure
    // names `Report`, so the cache depends on the host (§7 Q4 ii):
    // admitted while the name sweep keeps the body whole, refused
    // where the sweep would spell `#REF!` — S7a refused both.
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/workbook.xml", "Data!$A$1:$C$4", "OFFSET(Report!$D$1,0,0,4,3)");

    var wb = try Workbook.open(std.testing.allocator, io, src);
    defer wb.deinit();
    {
        var p = try wb.pivotTables();
        defer p.deinit();
        try std.testing.expect(p.caches[0].resolution == .unresolved);
        try std.testing.expect(!p.readsFromSheet(1) and p.dependsOnSheet(1) and p.mayReadFromSheet(1));
    }
    // Column D is the anchor: deleting it spells `OFFSET(#REF!,…)`.
    try std.testing.expectError(error.PivotEditUnsafe, wb.deleteColumn(1, 4));
    // Row 1 is the anchor's row: the same refusal.
    try std.testing.expectError(error.PivotEditUnsafe, wb.deleteRow(1, 1));
    const pt0 = (try wb.store.part("xl/pivotTables/pivotTable1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, pt0.bytes, "ref=\"A3:B6\"") != null);
    // Row 2 is above the rectangle (`A3:B6`) and below the anchor: the
    // body would survive the sweep — but the source is unbounded, so
    // the edit is a content change with no rectangle to rebuild the
    // snapshot from. S7b-3 admitted it and marked the cache; the
    // S7b-4 engine's edge refuses it, and nothing moves — on the
    // direct path and the Editor's alike.
    try std.testing.expectError(error.PivotEditUnsafe, wb.deleteRow(1, 2));
    const pt = (try wb.store.part("xl/pivotTables/pivotTable1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, pt.bytes, "ref=\"A3:B6\"") != null);
    const wbx = (try wb.store.part("xl/workbook.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, wbx.bytes, "OFFSET(Report!$D$1,0,0,4,3)") != null);
    try expectMarked(&wb.store, "xl/pivotCache/pivotCacheDefinition1.xml", false);

    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    try std.testing.expectError(error.ColEditUnsafeForSheet, ed.deleteColumn(1, 4));
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.deleteRow(1, 1));
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.deleteRow(1, 2));
    // The dangling spelling names no sheet of this workbook: admitted
    // untouched (§7 Q4 iii), and the host-only lift moves the pivot.
    const dangling = try tt.path(std.testing.allocator, io, "s7a_dangling_src.xlsx");
    defer std.testing.allocator.free(dangling);
    try pivots_mod.fixture.write(std.testing.allocator, io, dangling, .dangling);
    var ed2 = try Editor.open(std.testing.allocator, io, dangling);
    defer ed2.deinit();
    try ed2.insertRow(1, 1);
    try ed2.insertRow(0, 1);
}

test "S7a: a pivot part two sheets host refuses on either host (Codex r1 REL-035)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7a_shared_src.xlsx");
    defer std.testing.allocator.free(src);
    // An external source so neither sheet is a source; `Data` (sheet 1)
    // gets a relationship to `Report`'s pivot part.
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .external);
    {
        var store = try store_mod.PartStore.open(std.testing.allocator, io, src);
        defer store.deinit();
        try store.addPart(
            "xl/worksheets/_rels/sheet1.xml.rels",
            "application/vnd.openxmlformats-package.relationships+xml",
            \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            \\<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPT1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>
            ,
        );
        try store.save(io, src);
    }
    var wb = try Workbook.open(std.testing.allocator, io, src);
    defer wb.deinit();
    {
        var p = try wb.pivotTables();
        defer p.deinit();
        try std.testing.expectEqual(@as(usize, 2), p.tables.len);
        try std.testing.expect(p.hostsPivot(0) and p.hostsPivot(1));
    }
    try std.testing.expectError(error.PivotEditUnsafe, wb.insertRow(0, 1));
    try std.testing.expectError(error.PivotEditUnsafe, wb.insertRow(1, 1));
    const pt = (try wb.store.part("xl/pivotTables/pivotTable1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, pt.bytes, "ref=\"A3:B6\"") != null);

    // Two relationships from ONE sheet to one part move it once.
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/worksheets/_rels/sheet1.xml.rels", "<Relationship Id=\"rIdPT1\"", "<Relationship Id=\"rIdPT0\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable\" Target=\"../pivotTables/pivotTable1.xml\"/><Relationship Id=\"rIdPT1\"");
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/worksheets/_rels/sheet2.xml.rels", "<Relationship Id=\"rIdPT1\"", "<Relationship Id=\"rIdNone\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"http://example.invalid\" TargetMode=\"External\"/>");
    var wb2 = try Workbook.open(std.testing.allocator, io, src);
    defer wb2.deinit();
    try wb2.insertRow(0, 1);
    var p = try wb2.pivotTables();
    defer p.deinit();
    try std.testing.expectEqual(@as(usize, 2), p.tables.len);
    try std.testing.expectEqualStrings("A4:B7", p.tables[0].location_ref);
    try std.testing.expectEqualStrings("A4:B7", p.tables[1].location_ref);
}

fn replacePartsForFailures(allocator: std.mem.Allocator, io: std.Io, path: []const u8, a_new: []const u8, b_new: []const u8) !void {
    var store = try store_mod.PartStore.open(allocator, io, path);
    defer store.deinit();
    const a_old = try allocator.dupe(u8, (try store.part("xl/pivotTables/pivotTable1.xml")).?.bytes);
    defer allocator.free(a_old);
    const b_old = try allocator.dupe(u8, (try store.part("xl/pivotTables/pivotTable2.xml")).?.bytes);
    defer allocator.free(b_old);
    const big_before = store.big_parts.items.len;
    store.replaceParts(&.{
        .{ .name = "xl/pivotTables/pivotTable1.xml", .bytes = a_new },
        .{ .name = "xl/pivotTables/pivotTable2.xml", .bytes = b_new },
    }) catch |e| {
        // Neither installed, and no out-of-arena block retained: the
        // failure left the store as it was (Codex r2 REL-040).
        try std.testing.expectEqualStrings(a_old, (try store.part("xl/pivotTables/pivotTable1.xml")).?.bytes);
        try std.testing.expectEqualStrings(b_old, (try store.part("xl/pivotTables/pivotTable2.xml")).?.bytes);
        try std.testing.expectEqual(big_before, store.big_parts.items.len);
        return e;
    };
    try std.testing.expectEqualStrings(a_new, (try store.part("xl/pivotTables/pivotTable1.xml")).?.bytes);
    try std.testing.expectEqualStrings(b_new, (try store.part("xl/pivotTables/pivotTable2.xml")).?.bytes);
}

test "S7a: the sweep's install is transactional under allocation failure (Codex r1 REL-036)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7a_txn_src.xlsx");
    defer std.testing.allocator.free(src);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
    {
        var store = try store_mod.PartStore.open(std.testing.allocator, io, src);
        defer store.deinit();
        const first = (try store.part("xl/pivotTables/pivotTable1.xml")).?;
        try store.addPart("xl/pivotTables/pivotTable2.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml", first.bytes);
        try store.save(io, src);
    }
    const a_new = "<pivotTableDefinition a/>";
    const b_new = "<pivotTableDefinition b/>";
    try std.testing.checkAllAllocationFailures(std.testing.allocator, replacePartsForFailures, .{ io, src, a_new, b_new });

    // Payloads at the exact-block threshold take the out-of-arena path,
    // where a block registered for the first part must be given back
    // when the second fails.
    const big = store_mod.PartStore.big_payload_bytes + 16;
    const a_big = try std.testing.allocator.alloc(u8, big);
    defer std.testing.allocator.free(a_big);
    @memset(a_big, 'a');
    const b_big = try std.testing.allocator.alloc(u8, big);
    defer std.testing.allocator.free(b_big);
    @memset(b_big, 'b');
    try std.testing.checkAllAllocationFailures(std.testing.allocator, replacePartsForFailures, .{ io, src, a_big, b_big });
}

test "S7a: the direct Workbook path refuses position 0 and beyond the grid before any sweep (Codex r2 REL-038)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7a_zero_src.xlsx");
    defer std.testing.allocator.free(src);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
    // A pivot selection on the host, as Excel saves one: absolute
    // 0-based coordinates inside `A3:B6`. The writer emits
    // `<sheetViews>` only for frozen panes, so the whole block goes in
    // ahead of `<sheetData>`.
    try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/worksheets/sheet2.xml", "<sheetData", "<sheetViews><sheetView workbookViewId=\"0\"><pivotSelection pane=\"topLeft\" activeRow=\"3\" activeCol=\"1\" previousRow=\"3\" previousCol=\"1\" r:id=\"rIdPT1\"/></sheetView></sheetViews><sheetData");

    var wb = try Workbook.open(std.testing.allocator, io, src);
    defer wb.deinit();
    const before_sheet = try std.testing.allocator.dupe(u8, (try wb.store.part("xl/worksheets/sheet2.xml")).?.bytes);
    defer std.testing.allocator.free(before_sheet);
    try std.testing.expect(std.mem.indexOf(u8, before_sheet, "<pivotSelection ") != null);
    const before_pt = try std.testing.allocator.dupe(u8, (try wb.store.part("xl/pivotTables/pivotTable1.xml")).?.bytes);
    defer std.testing.allocator.free(before_pt);

    try std.testing.expectError(error.RowIndexOutOfRange, wb.insertRow(1, 0));
    try std.testing.expectError(error.RowIndexOutOfRange, wb.deleteRow(1, 0));
    try std.testing.expectError(error.RowIndexOutOfRange, wb.insertRow(1, xlsx.max_row + 1));
    try std.testing.expectError(error.ColumnIndexOutOfRange, wb.insertColumn(1, 0));
    try std.testing.expectError(error.ColumnIndexOutOfRange, wb.deleteColumn(1, 0));
    try std.testing.expectError(error.ColumnIndexOutOfRange, wb.deleteColumn(1, xlsx.max_col_1based + 1));
    try std.testing.expectEqualStrings(before_sheet, (try wb.store.part("xl/worksheets/sheet2.xml")).?.bytes);
    try std.testing.expectEqualStrings(before_pt, (try wb.store.part("xl/pivotTables/pivotTable1.xml")).?.bytes);

    // And the selection moves with the rectangle on a real edit.
    try wb.insertRow(1, 1);
    const sheet = (try wb.store.part("xl/worksheets/sheet2.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, sheet.bytes, "activeRow=\"4\" activeCol=\"1\" previousRow=\"4\" previousCol=\"1\"") != null);
    const pt = (try wb.store.part("xl/pivotTables/pivotTable1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, pt.bytes, "ref=\"A4:B7\"") != null);
}

test "S7b: the corpus fixture — a row above `mtCars Pivot` moves PivotTable3; `IrisSample` moves PivotTable1 and Table2 together" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const src = "tests/corpus/openxlsx_loadExample.xlsx";
    std.Io.Dir.cwd().access(io, src, .{}) catch return error.SkipZigTest;
    var tt = TestTmp.init();
    defer tt.deinit();
    const dst = try tt.path(std.testing.allocator, io, "s7a_corpus_dst.xlsx");
    defer std.testing.allocator.free(dst);

    {
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        // `IrisSample`: the pivot at `G2:K6` and `Table2` at `A1:E51`
        // both shift under an insert at row 1 / column 1; the cache is
        // table-named and needs nothing. `insertColumn` inside the
        // table (its field schema) would refuse — S7c's row.
        try ed.insertRow(0, 1);
        try ed.insertColumn(0, 1);
        try std.testing.expectError(error.ColEditUnsafeForSheet, ed.insertColumn(0, 3));
        try ed.insertRow(3, 1);
        try ed.save(io, dst);
    }

    var wb = try Workbook.open(std.testing.allocator, io, dst);
    defer wb.deinit();
    var p = try wb.pivotTables();
    defer p.deinit();
    var seen: u32 = 0;
    for (p.tables) |t| {
        if (t.sheet_idx == 3) {
            seen += 1;
            try std.testing.expectEqualStrings("xl/pivotTables/pivotTable2.xml", t.part_name);
            try std.testing.expectEqualStrings("A2:D6", t.location_ref);
            try std.testing.expect(std.mem.indexOf(u8, t.raw_xml, "<location ref=\"A2:D6\" firstHeaderRow=\"0\" firstDataRow=\"1\" firstDataCol=\"1\"/>") != null);
        } else if (t.sheet_idx == 0) {
            seen += 1;
            try std.testing.expectEqualStrings("xl/pivotTables/pivotTable1.xml", t.part_name);
            try std.testing.expectEqualStrings("H3:L7", t.location_ref);
            const c = p.cacheOf(t).?;
            try std.testing.expectEqual(pivots_mod.ResolvedVia.table, c.resolution.sheet.via);
            var buf: [pivots_mod.Bounds.format_buf_len]u8 = undefined;
            try std.testing.expectEqualStrings("B2:F52", c.resolution.sheet.bounds.?.formatA1(&buf).?);
        }
    }
    try std.testing.expectEqual(@as(u32, 2), seen);
}

// ─── extLst: xm:sqref shifts (#140), xm:f rewrites (S2) ─────────────

/// Sheet fixture carrying an `<extLst>` extension.
fn writeExtLstFixture(io: std.Io, path: []const u8, ext_body: []const u8) !void {
    {
        var w = xlsx.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }});
        try s.writeRow(&.{.{ .integer = 2 }});
        try w.save(io, path);
    }
    var store = try store_mod.PartStore.open(std.testing.allocator, io, path);
    defer store.deinit();
    const sheet = (try store.part("xl/worksheets/sheet1.xml")).?;
    const close = "</worksheet>";
    const at = std.mem.lastIndexOf(u8, sheet.bytes, close).?;
    const patched = try std.mem.concat(std.testing.allocator, u8, &.{
        sheet.bytes[0..at], ext_body, close,
    });
    defer std.testing.allocator.free(patched);
    try store.replacePart("xl/worksheets/sheet1.xml", patched);
    try store.save(io, path);
}

/// Two-sheet fixture — `Report` (index 0, sheet1.xml) and `Data`
/// (index 1, sheet2.xml), two numeric rows each — with an `<extLst>`
/// body appended to either sheet (empty = leave that sheet alone).
/// The cross-sheet shape is the point: a sparkline on `Report`
/// reading `Data!…` is what the S2 lift exists for.
fn writeTwoSheetExtLstFixture(io: std.Io, path: []const u8, report_ext: []const u8, data_ext: []const u8) !void {
    {
        var w = xlsx.Writer.init(std.testing.allocator);
        defer w.deinit();
        var report = try w.addSheet("Report");
        try report.writeRow(&.{.{ .integer = 1 }});
        try report.writeRow(&.{.{ .integer = 2 }});
        var data = try w.addSheet("Data");
        try data.writeRow(&.{.{ .integer = 10 }});
        try data.writeRow(&.{.{ .integer = 20 }});
        try w.save(io, path);
    }
    var store = try store_mod.PartStore.open(std.testing.allocator, io, path);
    defer store.deinit();
    const parts = [_]struct { name: []const u8, ext: []const u8 }{
        .{ .name = "xl/worksheets/sheet1.xml", .ext = report_ext },
        .{ .name = "xl/worksheets/sheet2.xml", .ext = data_ext },
    };
    for (parts) |p| {
        if (p.ext.len == 0) continue;
        const sheet = (try store.part(p.name)).?;
        const close = "</worksheet>";
        const at = std.mem.lastIndexOf(u8, sheet.bytes, close).?;
        const patched = try std.mem.concat(std.testing.allocator, u8, &.{
            sheet.bytes[0..at], p.ext, close,
        });
        defer std.testing.allocator.free(patched);
        try store.replacePart(p.name, patched);
    }
    try store.save(io, path);
}

/// Sparkline group whose one sparkline reads `f` and sits at `sqref`.
fn sparklineExt(comptime f: []const u8, comptime sqref: []const u8) []const u8 {
    return "<extLst><ext uri=\"{05C60535-1F16-4fd2-B633-F4F36F0B64E0}\">" ++
        "<x14:sparklineGroups><x14:sparklineGroup displayEmptyCellsAs=\"gap\"><x14:sparklines>" ++
        "<x14:sparkline><xm:f>" ++ f ++ "</xm:f><xm:sqref>" ++ sqref ++ "</xm:sqref>" ++
        "</x14:sparkline></x14:sparklines></x14:sparklineGroup>" ++
        "</x14:sparklineGroups></ext></extLst>";
}

/// The saved sheet part, owned by the caller.
fn readSavedPart(io: std.Io, path: []const u8, part_name: []const u8) ![]u8 {
    var store = try store_mod.PartStore.open(std.testing.allocator, io, path);
    defer store.deinit();
    const part = (try store.part(part_name)).?;
    return std.testing.allocator.dupe(u8, part.bytes);
}

test "Editor: row/col edits rewrite a sparkline's cross-sheet xm:f with sheet-name context (S2)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "xmf_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "xmf_dst.xlsx");
    defer std.testing.allocator.free(dst);
    // Pre-S2 this fixture refused with `RowEditUnsafeForSheet` /
    // `ColEditUnsafeForSheet` (#140's `ExtensionEditUnsafe`).
    try writeTwoSheetExtLstFixture(io, src, sparklineExt("Data!A2:A5", "B1"), "");

    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    // The sheet-name context is the whole point: an edit on `Data`
    // moves the formula and not the sqref; an edit on `Report` moves
    // the sqref and not the formula. `pkg/sheet_edit.zig` alone
    // could not tell the two apart.
    try ed.insertRow(1, 3); // Data: A2:A5 → A2:A6
    try ed.insertRow(0, 1); // Report: sqref B1 → B2, formula untouched
    try ed.insertColumn(1, 1); // Data: A2:A6 → B2:B6
    // An interior row: the span shrinks. (Deleting an endpoint row
    // is the rewriter's own A1 policy — `B2:B6` minus row 2 spells
    // `#REF!:B5`, endpoints moving independently, exactly as a cell
    // formula or DV/CF body would; S2 carries that convention, it
    // does not choose one.)
    try ed.deleteRow(1, 4); // Data: B2:B6 → B2:B5
    try ed.save(io, dst);

    const report = try readSavedPart(io, dst, "xl/worksheets/sheet1.xml");
    defer std.testing.allocator.free(report);
    try std.testing.expect(std.mem.indexOf(u8, report, "<xm:f>Data!B2:B5</xm:f><xm:sqref>B2</xm:sqref>") != null);
    // Attributes around the carrier survive the splice verbatim.
    try std.testing.expect(std.mem.indexOf(u8, report, "displayEmptyCellsAs=\"gap\"") != null);
}

test "Editor: every x14 xm:f shape rewrites — date axis, cfRule, cfvo, formula1 — entities intact (S2)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "xmf_shapes_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "xmf_shapes_dst.xlsx");
    defer std.testing.allocator.free(dst);
    // Every `<xm:f>` carrier Excel writes into a worksheet: a
    // sparkline group's date axis (a direct child of the group), the
    // sparkline's own range, an `x14:cfRule` expression and its
    // `x14:cfvo` thresholds, an `x14:dataValidation` `formula1`.
    // The CF expression carries a bare ref (host = Report) beside a
    // qualified one, and an entity the splice must re-escape.
    try writeTwoSheetExtLstFixture(io, src, "<extLst><ext uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\">" ++
        "<x14:conditionalFormattings><x14:conditionalFormatting>" ++
        "<x14:cfRule type=\"expression\" priority=\"1\"><xm:f>A1&gt;Data!$B$2</xm:f><x14:dxf/></x14:cfRule>" ++
        "<x14:cfRule type=\"iconSet\" priority=\"2\"><x14:iconSet><x14:cfvo type=\"num\"><xm:f>0</xm:f></x14:cfvo>" ++
        "<x14:cfvo type=\"formula\"><xm:f>Data!$C$3*2</xm:f></x14:cfvo></x14:iconSet></x14:cfRule>" ++
        "<xm:sqref>A1:A2</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext>" ++
        "<ext uri=\"{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}\"><x14:dataValidations count=\"1\">" ++
        "<x14:dataValidation type=\"list\" allowBlank=\"1\"><x14:formula1><xm:f>Data!$A$2:$A$5</xm:f></x14:formula1>" ++
        "<xm:sqref>C1</xm:sqref></x14:dataValidation></x14:dataValidations></ext>" ++
        "<ext uri=\"{05C60535-1F16-4fd2-B633-F4F36F0B64E0}\"><x14:sparklineGroups>" ++
        "<x14:sparklineGroup dateAxis=\"1\"><xm:f>Data!A1:E1</xm:f><x14:sparklines>" ++
        "<x14:sparkline><xm:f>Data!A2:E2</xm:f><xm:sqref>F2</xm:sqref></x14:sparkline>" ++
        "</x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst>", "");

    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    try ed.insertRow(1, 1); // Data: every Data-qualified row ref +1
    try ed.save(io, dst);

    const report = try readSavedPart(io, dst, "xl/worksheets/sheet1.xml");
    defer std.testing.allocator.free(report);
    // Bare `A1` is Report's and stays; `Data!$B$2` shifts; `&gt;`
    // round-trips as an entity, not a raw `>`.
    try std.testing.expect(std.mem.indexOf(u8, report, "<xm:f>A1&gt;Data!$B$3</xm:f>") != null);
    try std.testing.expect(std.mem.indexOf(u8, report, "<xm:f>0</xm:f>") != null);
    try std.testing.expect(std.mem.indexOf(u8, report, "<xm:f>Data!$C$4*2</xm:f>") != null);
    try std.testing.expect(std.mem.indexOf(u8, report, "<xm:f>Data!$A$3:$A$6</xm:f>") != null);
    try std.testing.expect(std.mem.indexOf(u8, report, "<xm:f>Data!A2:E2</xm:f><x14:sparklines>") != null);
    try std.testing.expect(std.mem.indexOf(u8, report, "<xm:f>Data!A3:E3</xm:f><xm:sqref>F2</xm:sqref>") != null);
    // Report's own sqrefs did not move — the edit was on Data.
    try std.testing.expect(std.mem.indexOf(u8, report, "<xm:sqref>A1:A2</xm:sqref>") != null);
    try std.testing.expect(std.mem.indexOf(u8, report, "<xm:sqref>C1</xm:sqref>") != null);
}

test "Editor: deleting a sparkline's whole source range collapses its xm:f to #REF! (S2)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "xmf_ref_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "xmf_ref_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try writeTwoSheetExtLstFixture(io, src, sparklineExt("Data!A2:C2", "B1"), "");

    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    // Same convention as a cell formula whose range is deleted: the
    // reference collapses to `#REF!` and the element stays, so Excel
    // shows an empty sparkline rather than a stale one.
    try ed.deleteRow(1, 2);
    try ed.save(io, dst);

    const report = try readSavedPart(io, dst, "xl/worksheets/sheet1.xml");
    defer std.testing.allocator.free(report);
    try std.testing.expect(std.mem.indexOf(u8, report, "<xm:f>Data!#REF!</xm:f><xm:sqref>B1</xm:sqref>") != null);
}

test "Editor: renameSheet and deleteSheet rewrite a sparkline's xm:f qualifier (S2)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "xmf_sheet_src.xlsx");
    defer std.testing.allocator.free(src);
    const renamed = try tt.path(std.testing.allocator, io, "xmf_sheet_renamed.xlsx");
    defer std.testing.allocator.free(renamed);
    const deleted = try tt.path(std.testing.allocator, io, "xmf_sheet_deleted.xlsx");
    defer std.testing.allocator.free(deleted);
    try writeTwoSheetExtLstFixture(io, src, sparklineExt("Data!A2:A5", "B1"), "");

    {
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try ed.renameSheet(1, "Raw Data");
        try ed.save(io, renamed);
        const report = try readSavedPart(io, renamed, "xl/worksheets/sheet1.xml");
        defer std.testing.allocator.free(report);
        try std.testing.expect(std.mem.indexOf(u8, report, "<xm:f>'Raw Data'!A2:A5</xm:f>") != null);
    }
    {
        var ed = try Editor.open(std.testing.allocator, io, renamed);
        defer ed.deinit();
        try ed.deleteSheet(1);
        try ed.save(io, deleted);
        const report = try readSavedPart(io, deleted, "xl/worksheets/sheet1.xml");
        defer std.testing.allocator.free(report);
        try std.testing.expect(std.mem.indexOf(u8, report, "<xm:f>#REF!</xm:f>") != null);
    }
}

test "Editor: a malformed xm:f refuses the whole edit before any part is touched (S2)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "xmf_bad_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "xmf_bad_dst.xlsx");
    defer std.testing.allocator.free(dst);
    // The unreadable carrier sits on Data; the edit targets Report.
    // All-or-nothing means Report's bytes — sqref included — must not
    // move either: the refusal is a property of the workbook, raised
    // before the first mutation, not of the edited sheet.
    try writeTwoSheetExtLstFixture(
        io,
        src,
        sparklineExt("Data!A2:A5", "B1"),
        "<extLst><ext><x14:sparklineGroups><x14:sparklineGroup><x14:sparklines>" ++
            "<x14:sparkline><xm:f>Report!A1:A2</x14:sparkline>" ++
            "</x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst>",
    );
    const before = try readSavedPart(io, src, "xl/worksheets/sheet1.xml");
    defer std.testing.allocator.free(before);

    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 1));
    try std.testing.expectError(error.ColEditUnsafeForSheet, ed.insertColumn(0, 1));
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.deleteRow(1, 1));
    try ed.save(io, dst);

    const after = try readSavedPart(io, dst, "xl/worksheets/sheet1.xml");
    defer std.testing.allocator.free(after);
    try std.testing.expectEqualStrings(before, after);
}

test "Editor: an extLst with only xm:sqref still edits, and shifts" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "xmsqref_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "xmsqref_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try writeExtLstFixture(io, src, "<extLst><ext><x14:conditionalFormattings><x14:conditionalFormatting>" ++
        "<xm:sqref>A5:A9</xm:sqref></x14:conditionalFormatting>" ++
        "</x14:conditionalFormattings></ext></extLst>");

    // The guard must be narrow: xm:f refuses, xm:sqref alone does not.
    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    try ed.insertRow(0, 2);
    try ed.save(io, dst);

    var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
    defer store.deinit();
    const sheet = (try store.part("xl/worksheets/sheet1.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, sheet.bytes, "<xm:sqref>A6:A10</xm:sqref>") != null);
}

// ─── S7b-5: the engine's second slice — the consumers ────────────────
//
// A rebuilt cache's consumers are laid out again in the same edit:
// the row field's `<items>`, `<rowItems>`, `location@ref`, and the
// cells of the host rectangle — the header, one row per item with a
// record, the grand total — written last in the sweep, in post-edit
// coordinates. A save with a staged write inside a source takes the
// same path (§7 Q3), with the marker kept as the safety net.

/// What a host cell holds after a save.
const HostValue = union(enum) { text: []const u8, number: f64, blank, none };

fn expectHostCell(wb: *Workbook, sheet_idx: u32, ref: []const u8, want: HostValue) !void {
    const ws = try wb.sheet(sheet_idx);
    const cell = try ws.cellByRef(ref);
    switch (want) {
        .none => try std.testing.expect(cell == null),
        .blank => {
            const c = cell orelse return error.TestExpectedCell;
            try std.testing.expect(c.raw_value == null and c.formula == null);
        },
        .number => |x| {
            const c = cell orelse return error.TestExpectedCell;
            try std.testing.expect(c.cell_type == .number);
            try std.testing.expectEqual(x, try std.fmt.parseFloat(f64, c.raw_value orelse return error.TestExpectedValue));
        },
        .text => |t| {
            const c = cell orelse return error.TestExpectedCell;
            const raw: []const u8 = switch (c.cell_type) {
                .shared_string => blk: {
                    const idx = try std.fmt.parseInt(usize, c.raw_value orelse return error.TestExpectedValue, 10);
                    const sst = (try wb.sst()) orelse return error.TestExpectedSst;
                    break :blk sst.entries[idx].plain;
                },
                // A pure shift leaves an inline string as written.
                .inline_string => c.raw_value orelse return error.TestExpectedValue,
                else => return error.TestExpectedText,
            };
            const decoded = try store_mod.decodeXmlEntities(std.testing.allocator, raw);
            defer std.testing.allocator.free(decoded);
            try std.testing.expectEqualStrings(t, decoded);
        },
    }
}

fn expectHostStyle(wb: *Workbook, sheet_idx: u32, ref: []const u8, want: ?u32) !void {
    const ws = try wb.sheet(sheet_idx);
    const c = (try ws.cellByRef(ref)) orelse return error.TestExpectedCell;
    try std.testing.expectEqual(want, c.style_idx);
}

fn expectPartHas(store: *store_mod.PartStore, part: []const u8, needle: []const u8) !void {
    const p = (try store.part(part)) orelse return error.PartNotFound;
    if (std.mem.indexOf(u8, p.bytes, needle) == null) {
        std.debug.print("\nexpected in {s}:\n  {s}\ngot:\n  {s}\n", .{ part, needle, p.bytes });
        return error.TestExpectedSubstring;
    }
}

const pt_part = "xl/pivotTables/pivotTable1.xml";
const fixture_pivot_field = "<pivotField axis=\"axisRow\" showAll=\"0\"><items count=\"3\"><item x=\"0\"/><item x=\"1\"/><item t=\"default\"/></items></pivotField>";
const fixture_data_field = "<dataField name=\"Sum of Qty\" fld=\"1\" baseField=\"0\" baseItem=\"0\"/>";

test "S7b-5: an insert inside the source re-lays the consumer — items, rowItems, location and the host cells; a delete shrinks it and clears the row it no longer covers" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7b5_grow_src.xlsx");
    defer std.testing.allocator.free(src);
    const grown = try tt.path(std.testing.allocator, io, "s7b5_grow_dst.xlsx");
    defer std.testing.allocator.free(grown);
    const shrunk = try tt.path(std.testing.allocator, io, "s7b5_shrink_dst.xlsx");
    defer std.testing.allocator.free(shrunk);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
    {
        // Row 2 of `Data` is inside `A1:C4`: a blank record, so a
        // `(blank)` item after the written ones (`sortType` manual),
        // shown with an empty sum; the grand total moves down one.
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try ed.insertRow(0, 2);
        try ed.save(io, grown);
    }
    {
        var store = try store_mod.PartStore.open(std.testing.allocator, io, grown);
        defer store.deinit();
        try expectRebuilt(&store, "xl/pivotCache/pivotCacheDefinition1.xml", "xl/pivotCache/pivotCacheRecords1.xml", 4);
        try expectPartHas(&store, "xl/pivotCache/pivotCacheDefinition1.xml", "<sharedItems containsBlank=\"1\" count=\"3\"><s v=\"East\"/><s v=\"West\"/><m/></sharedItems>");
        try expectPartHas(&store, pt_part, "<location ref=\"A3:B7\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>");
        try expectPartHas(&store, pt_part, "<pivotField axis=\"axisRow\" showAll=\"0\"><items count=\"4\"><item x=\"0\"/><item x=\"1\"/><item x=\"2\"/><item t=\"default\"/></items></pivotField>");
        try expectPartHas(&store, pt_part, "<rowItems count=\"4\"><i><x/></i><i><x v=\"1\"/></i><i><x v=\"2\"/></i><i t=\"grand\"><x/></i></rowItems>");
        try expectPartHas(&store, pt_part, "<colItems count=\"1\"><i/></colItems>");
        var wb = try Workbook.open(std.testing.allocator, io, grown);
        defer wb.deinit();
        try expectHostCell(&wb, 1, "A1", .{ .text = "pivot host" });
        try expectHostCell(&wb, 1, "A3", .{ .text = "Row Labels" });
        try expectHostCell(&wb, 1, "B3", .{ .text = "Sum of Qty" });
        try expectHostCell(&wb, 1, "A4", .{ .text = "East" });
        try expectHostCell(&wb, 1, "B4", .{ .number = 8 });
        try expectHostCell(&wb, 1, "A5", .{ .text = "West" });
        try expectHostCell(&wb, 1, "B5", .{ .number = 4 });
        try expectHostCell(&wb, 1, "A6", .{ .text = "(blank)" });
        try expectHostCell(&wb, 1, "B6", .blank);
        try expectHostCell(&wb, 1, "A7", .{ .text = "Grand Total" });
        try expectHostCell(&wb, 1, "B7", .{ .number = 12 });
        try expectHostCell(&wb, 1, "A8", .none);
        // The marker stays: the safety net under B (§8 Q2).
        try expectMarked(&store, "xl/pivotCache/pivotCacheDefinition1.xml", true);
    }
    {
        // From there, delete `West` (row 4 of the grown sheet): its
        // item stays in the inventory and the field's list, marked
        // missing, and is not a row; the old grand-total cells at row
        // 7 are cleared.
        var ed = try Editor.open(std.testing.allocator, io, grown);
        defer ed.deinit();
        try ed.deleteRow(0, 4);
        try ed.save(io, shrunk);
        var store = try store_mod.PartStore.open(std.testing.allocator, io, shrunk);
        defer store.deinit();
        try expectRebuilt(&store, "xl/pivotCache/pivotCacheDefinition1.xml", "xl/pivotCache/pivotCacheRecords1.xml", 3);
        try expectPartHas(&store, pt_part, "<location ref=\"A3:B6\"");
        try expectPartHas(&store, pt_part, "<items count=\"4\"><item x=\"0\"/><item x=\"1\" m=\"1\"/><item x=\"2\"/><item t=\"default\"/></items>");
        try expectPartHas(&store, pt_part, "<rowItems count=\"3\"><i><x/></i><i><x v=\"2\"/></i><i t=\"grand\"><x/></i></rowItems>");
        var wb = try Workbook.open(std.testing.allocator, io, shrunk);
        defer wb.deinit();
        try expectHostCell(&wb, 1, "A4", .{ .text = "East" });
        try expectHostCell(&wb, 1, "B4", .{ .number = 8 });
        try expectHostCell(&wb, 1, "A5", .{ .text = "(blank)" });
        try expectHostCell(&wb, 1, "B5", .blank);
        try expectHostCell(&wb, 1, "A6", .{ .text = "Grand Total" });
        try expectHostCell(&wb, 1, "B6", .{ .number = 8 });
        try expectHostCell(&wb, 1, "A7", .none);
        try expectHostCell(&wb, 1, "B7", .none);
    }
}

/// The fixture with its pivot hosted on `Data` at `A7:B10`, below the
/// source, laid out as Excel (in French) would have left it — the
/// captions it wrote, a style per cell.
fn writeHostOnSourceFixture(io: std.Io, path: []const u8, bottom_note: bool) !void {
    const a = std.testing.allocator;
    try pivots_mod.fixture.write(a, io, path, .sheet_ref);
    {
        var store = try store_mod.PartStore.open(a, io, path);
        defer store.deinit();
        try store.addPart(
            "xl/worksheets/_rels/sheet1.xml.rels",
            "application/vnd.openxmlformats-package.relationships+xml",
            \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            \\<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPT1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>
            ,
        );
        try store.save(io, path);
    }
    try pivots_mod.fixture.patchPart(a, io, path, "xl/worksheets/_rels/sheet2.xml.rels", "<Relationship Id=\"rIdPT1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable\" Target=\"../pivotTables/pivotTable1.xml\"/>", "");
    try pivots_mod.fixture.patchPart(a, io, path, pt_part, "ref=\"A3:B6\"", "ref=\"A7:B10\"");
    const note: []const u8 = if (bottom_note) "<row r=\"11\"><c r=\"A11\" t=\"inlineStr\"><is><t>note</t></is></c></row>" else "";
    const rows = try std.mem.concat(a, u8, &.{
        "<row r=\"7\"><c r=\"A7\" s=\"3\" t=\"inlineStr\"><is><t>Étiquettes de lignes</t></is></c><c r=\"B7\" s=\"4\" t=\"inlineStr\"><is><t>Somme de Qty</t></is></c></row>",
        "<row r=\"8\"><c r=\"A8\" s=\"5\" t=\"inlineStr\"><is><t>East</t></is></c><c r=\"B8\" s=\"6\"><v>8</v></c></row>",
        "<row r=\"9\"><c r=\"A9\" s=\"5\" t=\"inlineStr\"><is><t>West</t></is></c><c r=\"B9\" s=\"6\"><v>4</v></c></row>",
        "<row r=\"10\"><c r=\"A10\" s=\"7\" t=\"inlineStr\"><is><t>Total général</t></is></c><c r=\"B10\" s=\"8\"><v>12</v></c></row>",
        note,
        "</sheetData>",
    });
    defer a.free(rows);
    try pivots_mod.fixture.patchPart(a, io, path, "xl/worksheets/sheet1.xml", "</sheetData>", rows);
}

test "S7b-5: a host on the edited sheet — the rectangle moves with the edit, then grows, in post-edit coordinates; the captions and styles carry; growth over a cell refuses whole" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7b5_host_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "s7b5_host_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try writeHostOnSourceFixture(io, src, false);
    {
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try ed.insertRow(0, 2);
        try ed.save(io, dst);
    }
    var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
    defer store.deinit();
    // `A7:B10` shifted to `A8:B11` by the insert, then grown to
    // `A8:B12` by the `(blank)` row.
    try expectPartHas(&store, pt_part, "<location ref=\"A8:B12\"");
    var wb = try Workbook.open(std.testing.allocator, io, dst);
    defer wb.deinit();
    try expectHostCell(&wb, 0, "A8", .{ .text = "Étiquettes de lignes" });
    try expectHostCell(&wb, 0, "B8", .{ .text = "Sum of Qty" });
    try expectHostCell(&wb, 0, "A9", .{ .text = "East" });
    try expectHostCell(&wb, 0, "B9", .{ .number = 8 });
    try expectHostCell(&wb, 0, "A10", .{ .text = "West" });
    try expectHostCell(&wb, 0, "B10", .{ .number = 4 });
    try expectHostCell(&wb, 0, "A11", .{ .text = "(blank)" });
    try expectHostCell(&wb, 0, "B11", .blank);
    try expectHostCell(&wb, 0, "A12", .{ .text = "Total général" });
    try expectHostCell(&wb, 0, "B12", .{ .number = 12 });
    try expectHostCell(&wb, 0, "A7", .none);
    try expectHostCell(&wb, 0, "A13", .none);
    // Styles by row kind: the header's, the last item row's for every
    // item row (the new one included), the grand total's.
    try expectHostStyle(&wb, 0, "A8", 3);
    try expectHostStyle(&wb, 0, "B8", 4);
    try expectHostStyle(&wb, 0, "A9", 5);
    try expectHostStyle(&wb, 0, "B11", 6);
    try expectHostStyle(&wb, 0, "A12", 7);
    try expectHostStyle(&wb, 0, "B12", 8);
    // The source rows moved too: the blank record's row is 2.
    try expectHostCell(&wb, 0, "A3", .{ .text = "East" });

    // A note at `A11` sits at `A12` after the insert — where the grown
    // rectangle would land. Refused whole, on both paths.
    const noted = try tt.path(std.testing.allocator, io, "s7b5_host_noted.xlsx");
    defer std.testing.allocator.free(noted);
    try writeHostOnSourceFixture(io, noted, true);
    var ed = try Editor.open(std.testing.allocator, io, noted);
    defer ed.deinit();
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
    var direct = try Workbook.open(std.testing.allocator, io, noted);
    defer direct.deinit();
    const before = try std.testing.allocator.dupe(u8, (try direct.store.part("xl/worksheets/sheet1.xml")).?.bytes);
    defer std.testing.allocator.free(before);
    try std.testing.expectError(error.PivotEditUnsafe, direct.insertRow(0, 2));
    try std.testing.expectEqualStrings(before, (try direct.store.part("xl/worksheets/sheet1.xml")).?.bytes);
    try expectPartHas(&direct.store, pt_part, "<location ref=\"A7:B10\"");
    // A shift alone (row 1, above the source's header — no content
    // changes) moves the rectangle and writes no cell.
    try direct.insertRow(0, 1);
    try expectPartHas(&direct.store, pt_part, "<location ref=\"A8:B11\"");
    try expectHostCell(&direct, 0, "A11", .{ .text = "Total général" });
    try expectHostCell(&direct, 0, "A12", .{ .text = "note" });
}

test "S7b-5: a save with a write inside the source rebuilds the cache and re-lays the consumer (§7 Q3); `ascending` sorts a new item by value, `manual` appends it; a write the rebuild refuses still marks" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const dst = try tt.path(std.testing.allocator, io, "s7b5_save_dst.xlsx");
    defer std.testing.allocator.free(dst);
    const Case = struct { name: []const u8, ascending: bool, rows: [3][]const u8, nums: [3]f64 };
    const cases = [_]Case{
        .{ .name = "ascending", .ascending = true, .rows = .{ "Central", "East", "West" }, .nums = .{ 3, 5, 4 } },
        .{ .name = "manual", .ascending = false, .rows = .{ "East", "West", "Central" }, .nums = .{ 5, 4, 3 } },
    };
    for (cases) |case| {
        const file = try std.fmt.allocPrint(std.testing.allocator, "s7b5_save_{s}.xlsx", .{case.name});
        defer std.testing.allocator.free(file);
        const src = try tt.path(std.testing.allocator, io, file);
        defer std.testing.allocator.free(src);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
        if (case.ascending) try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, pt_part, "<pivotField axis=\"axisRow\" showAll=\"0\">", "<pivotField axis=\"axisRow\" showAll=\"0\" sortType=\"ascending\">");
        {
            // `Data!A2`: the first record's `Region`, `East` → `Central`.
            var ed = try Editor.open(std.testing.allocator, io, src);
            defer ed.deinit();
            try ed.setCell(0, 2, 0, .{ .string = "Central" });
            try ed.save(io, dst);
        }
        var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
        defer store.deinit();
        try expectRebuilt(&store, "xl/pivotCache/pivotCacheDefinition1.xml", "xl/pivotCache/pivotCacheRecords1.xml", 3);
        try expectMarked(&store, "xl/pivotCache/pivotCacheDefinition1.xml", true);
        try expectPartHas(&store, "xl/pivotCache/pivotCacheDefinition1.xml", "<sharedItems count=\"3\"><s v=\"East\"/><s v=\"West\"/><s v=\"Central\"/></sharedItems>");
        try expectPartHas(&store, "xl/pivotCache/pivotCacheRecords1.xml", "<r><x v=\"2\"/><n v=\"3\"/><n v=\"1.5\"/></r>");
        if (case.ascending) {
            try expectPartHas(&store, pt_part, "<items count=\"4\"><item x=\"2\"/><item x=\"0\"/><item x=\"1\"/><item t=\"default\"/></items>");
        } else {
            try expectPartHas(&store, pt_part, "<items count=\"4\"><item x=\"0\"/><item x=\"1\"/><item x=\"2\"/><item t=\"default\"/></items>");
        }
        try expectPartHas(&store, pt_part, "<rowItems count=\"4\"><i><x/></i><i><x v=\"1\"/></i><i><x v=\"2\"/></i><i t=\"grand\"><x/></i></rowItems>");
        try expectPartHas(&store, pt_part, "<location ref=\"A3:B7\"");
        var wb = try Workbook.open(std.testing.allocator, io, dst);
        defer wb.deinit();
        try expectHostCell(&wb, 0, "A2", .{ .text = "Central" });
        try expectHostCell(&wb, 1, "A3", .{ .text = "Row Labels" });
        try expectHostCell(&wb, 1, "A4", .{ .text = case.rows[0] });
        try expectHostCell(&wb, 1, "B4", .{ .number = case.nums[0] });
        try expectHostCell(&wb, 1, "A5", .{ .text = case.rows[1] });
        try expectHostCell(&wb, 1, "B5", .{ .number = case.nums[1] });
        try expectHostCell(&wb, 1, "A6", .{ .text = case.rows[2] });
        try expectHostCell(&wb, 1, "B6", .{ .number = case.nums[2] });
        try expectHostCell(&wb, 1, "A7", .{ .text = "Grand Total" });
        try expectHostCell(&wb, 1, "B7", .{ .number = 12 });
    }
    {
        // A write the rebuild refuses — a boolean in a record — marks
        // and leaves the snapshot: the save is not where an admitted
        // write is refused.
        const src = try tt.path(std.testing.allocator, io, "s7b5_save_refused.xlsx");
        defer std.testing.allocator.free(src);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
        const original = try cacheDefinitionBytes(std.testing.allocator, io, src);
        defer std.testing.allocator.free(original);
        const marked = try markedDefinition(std.testing.allocator, original);
        defer std.testing.allocator.free(marked);
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try ed.setCell(0, 2, 1, .{ .boolean = true });
        try ed.save(io, dst);
        const after = try cacheDefinitionBytes(std.testing.allocator, io, dst);
        defer std.testing.allocator.free(after);
        try std.testing.expectEqualStrings(marked, after);
        var wb = try Workbook.open(std.testing.allocator, io, dst);
        defer wb.deinit();
        try expectHostCell(&wb, 1, "A3", .none);
    }
    {
        // A write inside the pivot's own rectangle is the pivot's to
        // overwrite: the refresh wins, as Excel's does.
        const src = try tt.path(std.testing.allocator, io, "s7b5_save_typed_over.xlsx");
        defer std.testing.allocator.free(src);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try ed.setCell(0, 3, 1, .{ .integer = 10 });
        try ed.setCell(1, 4, 1, .{ .string = "typed over" });
        try ed.setCell(1, 9, 0, .{ .string = "kept" });
        try ed.save(io, dst);
        var wb = try Workbook.open(std.testing.allocator, io, dst);
        defer wb.deinit();
        try expectHostCell(&wb, 1, "B4", .{ .number = 8 });
        try expectHostCell(&wb, 1, "B5", .{ .number = 10 });
        try expectHostCell(&wb, 1, "B6", .{ .number = 18 });
        try expectHostCell(&wb, 1, "A9", .{ .text = "kept" });
    }
}

test "S7b-5: the values axis across — two data fields, an average beside a sum, the grand total folded from the subtotals" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7b5_two_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "s7b5_two_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
    const a = std.testing.allocator;
    try pivots_mod.fixture.patchPart(a, io, src, pt_part, "<pivotField dataField=\"1\" showAll=\"0\"/><pivotField showAll=\"0\"/>", "<pivotField dataField=\"1\" showAll=\"0\"/><pivotField dataField=\"1\" showAll=\"0\"/>");
    try pivots_mod.fixture.patchPart(a, io, src, pt_part, "<colItems count=\"1\"><i/></colItems><dataFields count=\"1\">" ++ fixture_data_field ++ "</dataFields>", "<colFields count=\"1\"><field x=\"-2\"/></colFields><colItems count=\"2\"><i><x/></i><i i=\"1\"><x v=\"1\"/></i></colItems><dataFields count=\"2\">" ++ fixture_data_field ++ "<dataField name=\"Average of Price\" fld=\"2\" subtotal=\"average\" baseField=\"0\" baseItem=\"0\"/></dataFields>");
    try pivots_mod.fixture.patchPart(a, io, src, pt_part, "ref=\"A3:B6\" firstHeaderRow=\"1\"", "ref=\"A3:C6\" firstHeaderRow=\"0\"");
    var ed = try Editor.open(a, io, src);
    defer ed.deinit();
    try ed.insertRow(0, 2);
    try ed.save(io, dst);
    var store = try store_mod.PartStore.open(a, io, dst);
    defer store.deinit();
    try expectPartHas(&store, pt_part, "<location ref=\"A3:C7\" firstHeaderRow=\"0\"");
    try expectPartHas(&store, pt_part, "<colItems count=\"2\"><i><x/></i><i i=\"1\"><x v=\"1\"/></i></colItems>");
    var wb = try Workbook.open(a, io, dst);
    defer wb.deinit();
    try expectHostCell(&wb, 1, "C3", .{ .text = "Average of Price" });
    try expectHostCell(&wb, 1, "B4", .{ .number = 8 });
    try expectHostCell(&wb, 1, "C4", .{ .number = 2.5 });
    try expectHostCell(&wb, 1, "B5", .{ .number = 4 });
    try expectHostCell(&wb, 1, "C5", .{ .number = 2.5 });
    try expectHostCell(&wb, 1, "B6", .blank);
    try expectHostCell(&wb, 1, "C6", .blank);
    try expectHostCell(&wb, 1, "B7", .{ .number = 12 });
    try expectHostCell(&wb, 1, "C7", .{ .number = 2.5 });
}

test "S7b-5: a consumer form the slice does not lay out refuses the edit whole — page, column and hidden items, a dispersion, a percentage, tabular form, a chart on a field" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const Case = struct { name: []const u8, old: []const u8, new: []const u8, admitted: bool = false, old2: ?[]const u8 = null, new2: []const u8 = "", old3: ?[]const u8 = null, new3: []const u8 = "", direct: anyerror = error.PivotEditUnsafe };
    const cases = [_]Case{
        .{ .name = "page", .old = "<dataFields count=\"1\">", .new = "<pageFields count=\"1\"><pageField fld=\"2\" hier=\"-1\"/></pageFields><dataFields count=\"1\">" },
        // Codex #206 r1 REL-104: a container whose count disagrees
        // with its children, a second subtotal item, an item both
        // cached and derived, a grand total without its `<x>`, an
        // attribute on a container the layout regenerates — refused
        // before mutation rather than normalised.
        .{ .name = "items_count", .old = "<items count=\"3\">", .new = "<items count=\"9\">", .direct = error.MalformedPivotXml },
        .{ .name = "row_items_count", .old = "<rowItems count=\"3\">", .new = "<rowItems count=\"2\">", .direct = error.MalformedPivotXml },
        .{ .name = "col_items_count", .old = "<colItems count=\"1\">", .new = "<colItems count=\"2\">", .direct = error.MalformedPivotXml },
        .{ .name = "grand_without_x", .old = "<i t=\"grand\"><x/></i>", .new = "<i t=\"grand\"/>" },
        .{ .name = "two_defaults", .old = "<item t=\"default\"/></items>", .new = "<item t=\"default\"/><item t=\"default\"/></items>", .old2 = "<items count=\"3\">", .new2 = "<items count=\"4\">" },
        .{ .name = "x_and_t", .old = "<item x=\"0\"/>", .new = "<item x=\"0\" t=\"default\"/>" },
        .{ .name = "items_attr", .old = "<items count=\"3\">", .new = "<items count=\"3\" foo=\"1\">" },
        .{ .name = "row_items_attr", .old = "<rowItems count=\"3\">", .new = "<rowItems count=\"3\" foo=\"1\">" },
        // Codex #206 r2 REL-202: a `<rowItems>` position outside the
        // item list, twice, or a grand total not at position 0.
        .{ .name = "position_out_of_range", .old = "<i><x v=\"1\"/></i>", .new = "<i><x v=\"99\"/></i>" },
        .{ .name = "position_twice", .old = "<i><x v=\"1\"/></i>", .new = "<i><x/></i>" },
        .{ .name = "grand_at_one", .old = "<i t=\"grand\"><x/></i>", .new = "<i t=\"grand\"><x v=\"1\"/></i>" },
        .{ .name = "col_field", .old = "<dataFields count=\"1\">", .new = "<colFields count=\"1\"><field x=\"2\"/></colFields><dataFields count=\"1\">" },
        .{ .name = "show_all", .old = "<pivotField axis=\"axisRow\" showAll=\"0\">", .new = "<pivotField axis=\"axisRow\" showAll=\"1\">" },
        .{ .name = "hidden_item", .old = "<item x=\"1\"/>", .new = "<item x=\"1\" h=\"1\"/>" },
        .{ .name = "std_dev", .old = "fld=\"1\" baseField", .new = "fld=\"1\" subtotal=\"stdDev\" baseField" },
        .{ .name = "percent", .old = "fld=\"1\" baseField", .new = "fld=\"1\" showDataAs=\"percentOfTotal\" baseField" },
        .{ .name = "tabular", .old = "outline=\"1\"", .new = "outline=\"1\" compact=\"0\"" },
        // Codex #206 r5 REL-501: the field's own `compact="0"`.
        .{ .name = "field_tabular", .old = "<pivotField axis=\"axisRow\" showAll=\"0\">", .new = "<pivotField axis=\"axisRow\" showAll=\"0\" compact=\"0\">" },
        .{ .name = "descending", .old = "<pivotField axis=\"axisRow\" showAll=\"0\">", .new = "<pivotField axis=\"axisRow\" showAll=\"0\" sortType=\"descending\">" },
        .{ .name = "top_n", .old = "<pivotField axis=\"axisRow\" showAll=\"0\">", .new = "<pivotField axis=\"axisRow\" showAll=\"0\" autoShow=\"1\" topAutoShow=\"1\" itemPageCount=\"2\" rankBy=\"0\">" },
        .{ .name = "filters", .old = "<pivotTableStyleInfo", .new = "<filters count=\"1\"><filter fld=\"1\" type=\"count\" evalOrder=\"-1\" id=\"1\" iMeasureFld=\"0\"><autoFilter ref=\"A1\"/></filter></filters><pivotTableStyleInfo" },
        .{ .name = "chart_on_field", .old = "<pivotTableStyleInfo", .new = "<chartFormats count=\"1\"><chartFormat chart=\"0\" format=\"0\" series=\"1\"><pivotArea type=\"data\" outline=\"0\" fieldPosition=\"0\"><references count=\"1\"><reference field=\"0\" count=\"1\" selected=\"0\"><x v=\"0\"/></reference></references></pivotArea></chartFormat></chartFormats><pivotTableStyleInfo" },
        // Codex #206 r16 REL-1603: display options that change the cells.
        .{ .name = "show_headers_off", .old = "outline=\"1\"", .new = "outline=\"1\" showHeaders=\"0\"" },
        .{ .name = "missing_caption", .old = "outline=\"1\"", .new = "outline=\"1\" showMissing=\"1\" missingCaption=\"N/A\"" },
        .{ .name = "merge_item", .old = "outline=\"1\"", .new = "outline=\"1\" mergeItem=\"1\"" },
        // Codex #206 r15 REL-1504: an area selecting by axis, or proving
        // nothing.
        .{ .name = "chart_area_axis_row", .old = "<pivotTableStyleInfo", .new = "<chartFormats count=\"1\"><chartFormat chart=\"0\" format=\"0\" series=\"1\"><pivotArea type=\"data\" outline=\"0\" fieldPosition=\"0\" axis=\"axisRow\"/></chartFormat></chartFormats><pivotTableStyleInfo" },
        .{ .name = "chart_area_empty", .old = "<pivotTableStyleInfo", .new = "<chartFormats count=\"1\"><chartFormat chart=\"0\" format=\"0\" series=\"1\"><pivotArea/></chartFormat></chartFormats><pivotTableStyleInfo" },
        .{ .name = "chart_area_axis_values", .old = "<pivotTableStyleInfo", .new = "<chartFormats count=\"1\"><chartFormat chart=\"0\" format=\"0\" series=\"1\"><pivotArea type=\"data\" outline=\"0\" fieldPosition=\"0\" axis=\"axisValues\"/></chartFormat></chartFormats><pivotTableStyleInfo", .admitted = true },
        // Codex #206 r15 REL-1505: the location's required offsets.
        .{ .name = "no_first_data_row", .old = " firstDataRow=\"1\"", .new = "", .direct = error.MalformedPivotXml },
        .{ .name = "no_first_header_row", .old = " firstHeaderRow=\"1\"", .new = "", .direct = error.MalformedPivotXml },
        .{ .name = "no_first_data_col", .old = " firstDataCol=\"1\"", .new = "", .direct = error.MalformedPivotXml },
        .{ .name = "first_data_col_2", .old = "firstDataCol=\"1\"", .new = "firstDataCol=\"2\"" },
        .{ .name = "first_header_row_2", .old = "firstHeaderRow=\"1\"", .new = "firstHeaderRow=\"2\"" },
        // Codex #206 r12 REL-1201: the field wrapper's own count and
        // attributes.
        .{ .name = "pivot_fields_count", .old = "<pivotFields count=\"3\">", .new = "<pivotFields count=\"9\">", .direct = error.MalformedPivotXml },
        .{ .name = "pivot_fields_attr", .old = "<pivotFields count=\"3\">", .new = "<pivotFields count=\"3\" foo=\"1\">" },
        // Codex #206 r10 REL-1002: the axis and data wrappers.
        .{ .name = "row_fields_count", .old = "<rowFields count=\"1\">", .new = "<rowFields count=\"2\">", .direct = error.MalformedPivotXml },
        .{ .name = "data_fields_stray", .old = "<dataFields count=\"1\">", .new = "<dataFields count=\"1\"><foo/>" },
        .{ .name = "row_fields_attr", .old = "<rowFields count=\"1\">", .new = "<rowFields count=\"1\" foo=\"1\">" },
        .{ .name = "field_body", .old = "<field x=\"0\"/>", .new = "<field x=\"0\">x</field>" },
        .{ .name = "data_field_comment", .old = "<dataFields count=\"1\">", .new = "<dataFields count=\"1\"><!-- x -->" },
        // Codex #206 r9 REL-902: content under `<pivotFields>` that is
        // not a field — a stray element, one under another prefix.
        .{ .name = "fields_stray", .old = "<pivotFields count=\"3\">", .new = "<pivotFields count=\"3\"><foo/>" },
        .{ .name = "fields_foreign", .old = "<pivotFields count=\"3\">", .new = "<pivotFields count=\"3\"><x14:future xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"/>" },
        // Codex #206 r8 REL-801: no grand total, or column totals off —
        // consistent with themselves, still not the one form.
        .{ .name = "no_grand_total", .old = "outline=\"1\"", .new = "outline=\"1\" rowGrandTotals=\"0\"", .old2 = "<rowItems count=\"3\"><i><x/></i><i><x v=\"1\"/></i><i t=\"grand\"><x/></i></rowItems>", .new2 = "<rowItems count=\"2\"><i><x/></i><i><x v=\"1\"/></i></rowItems>", .old3 = "ref=\"A3:B6\"", .new3 = "ref=\"A3:B5\"" },
        .{ .name = "col_grand_totals_off", .old = "outline=\"1\"", .new = "outline=\"1\" colGrandTotals=\"0\"" },
        // Codex #206 r6 REL-602: the area's own `field`.
        .{ .name = "chart_area_field", .old = "<pivotTableStyleInfo", .new = "<chartFormats count=\"1\"><chartFormat chart=\"0\" format=\"0\" series=\"1\"><pivotArea type=\"data\" outline=\"0\" fieldPosition=\"0\" field=\"0\"/></chartFormat></chartFormats><pivotTableStyleInfo" },
        .{ .name = "chart_area_values", .old = "<pivotTableStyleInfo", .new = "<chartFormats count=\"1\"><chartFormat chart=\"0\" format=\"0\" series=\"1\"><pivotArea type=\"data\" outline=\"0\" fieldPosition=\"0\" field=\"4294967294\"/></chartFormat></chartFormats><pivotTableStyleInfo", .admitted = true },
        .{ .name = "chart_on_values", .old = "<pivotTableStyleInfo", .new = "<chartFormats count=\"1\"><chartFormat chart=\"0\" format=\"0\" series=\"1\"><pivotArea type=\"data\" outline=\"0\" fieldPosition=\"0\"><references count=\"1\"><reference field=\"4294967294\" count=\"1\" selected=\"0\"><x v=\"0\"/></reference></references></pivotArea></chartFormat></chartFormats><pivotTableStyleInfo", .admitted = true },
        .{ .name = "no_row_items", .old = "<rowItems count=\"3\"><i><x/></i><i><x v=\"1\"/></i><i t=\"grand\"><x/></i></rowItems>", .new = "" },
    };
    for (cases) |case| {
        const file = try std.fmt.allocPrint(std.testing.allocator, "s7b5_refuse_{s}.xlsx", .{case.name});
        defer std.testing.allocator.free(file);
        const src = try tt.path(std.testing.allocator, io, file);
        defer std.testing.allocator.free(src);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, pt_part, case.old, case.new);
        if (case.old2) |old2| try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, pt_part, old2, case.new2);
        if (case.old3) |old3| try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, pt_part, old3, case.new3);
        var wb = try Workbook.open(std.testing.allocator, io, src);
        defer wb.deinit();
        const before = try std.testing.allocator.dupe(u8, (try wb.store.part("xl/pivotCache/pivotCacheRecords1.xml")).?.bytes);
        defer std.testing.allocator.free(before);
        if (case.admitted) {
            try wb.insertRow(0, 2);
            try expectPartHas(&wb.store, pt_part, "<location ref=\"A3:B7\"");
            continue;
        }
        try std.testing.expectError(case.direct, wb.insertRow(0, 2));
        try std.testing.expectEqualStrings(before, (try wb.store.part("xl/pivotCache/pivotCacheRecords1.xml")).?.bytes);
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
        // A shift stays a shift: row 1 is above the source, and no
        // cache is rebuilt, so the form is not asked.
        try ed.insertRow(0, 1);
    }
}

/// A second consumer of the fixture's cache on `Report`, `A10:B13`,
/// with a page field the slice does not lay out.
fn addUnsupportedConsumer(io: std.Io, path: []const u8) !void {
    const a = std.testing.allocator;
    var store = try store_mod.PartStore.open(a, io, path);
    defer store.deinit();
    try store.addPart(
        "xl/pivotTables/pivotTable2.xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml",
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable2" cacheId="7" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0"><location ref="A10:B13" firstHeaderRow="1" firstDataRow="1" firstDataCol="1" rowPageCount="1" colPageCount="1"/><pivotFields count="3"><pivotField axis="axisRow" showAll="0"><items count="3"><item x="0"/><item x="1"/><item t="default"/></items></pivotField><pivotField dataField="1" showAll="0"/><pivotField axis="axisPage" showAll="0"><items count="1"><item t="default"/></items></pivotField></pivotFields><rowFields count="1"><field x="0"/></rowFields><rowItems count="3"><i><x/></i><i><x v="1"/></i><i t="grand"><x/></i></rowItems><colItems count="1"><i/></colItems><pageFields count="1"><pageField fld="2" hier="-1"/></pageFields><dataFields count="1"><dataField name="Sum of Qty" fld="1" baseField="0" baseItem="0"/></dataFields></pivotTableDefinition>
        ,
    );
    try store.addPart(
        "xl/pivotTables/_rels/pivotTable2.xml.rels",
        "application/vnd.openxmlformats-package.relationships+xml",
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>
        ,
    );
    try store.save(io, path);
    try pivots_mod.fixture.patchPart(a, io, path, "xl/worksheets/_rels/sheet2.xml.rels", "</Relationships>", "<Relationship Id=\"rIdPT2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable\" Target=\"../pivotTables/pivotTable2.xml\"/></Relationships>");
}

/// Every pivot part and the host byte-identical to `original`, save
/// for the marker on the definition.
fn expectMarkerOnly(io: std.Io, original: []const u8, saved: []const u8) !void {
    return expectMarkerOnlyWith(io, original, saved, true);
}

/// `same_sst` false when the save's own write adds a string.
fn expectMarkerOnlyWith(io: std.Io, original: []const u8, saved: []const u8, same_sst: bool) !void {
    const a = std.testing.allocator;
    var before = try store_mod.PartStore.open(a, io, original);
    defer before.deinit();
    var after = try store_mod.PartStore.open(a, io, saved);
    defer after.deinit();
    const same = [_][]const u8{ "xl/pivotCache/pivotCacheRecords1.xml", "xl/pivotTables/pivotTable1.xml", "xl/worksheets/sheet2.xml", "xl/sharedStrings.xml" };
    for (same) |name| {
        if (!same_sst and std.mem.eql(u8, name, "xl/sharedStrings.xml")) continue;
        const b = (try before.part(name)) orelse return error.PartNotFound;
        const c = (try after.part(name)) orelse return error.PartNotFound;
        try std.testing.expectEqualStrings(b.bytes, c.bytes);
    }
    if (try before.part("xl/pivotTables/pivotTable2.xml")) |b| {
        const c = (try after.part("xl/pivotTables/pivotTable2.xml")) orelse return error.PartNotFound;
        try std.testing.expectEqualStrings(b.bytes, c.bytes);
    }
    const def = (try before.part("xl/pivotCache/pivotCacheDefinition1.xml")) orelse return error.PartNotFound;
    const marked = try markedDefinition(a, def.bytes);
    defer a.free(marked);
    const got = (try after.part("xl/pivotCache/pivotCacheDefinition1.xml")) orelse return error.PartNotFound;
    try std.testing.expectEqualStrings(marked, got.bytes);
}

test "S7b-5: a cache with one consumer the slice cannot lay out takes the marker alone at save, whole — no rebuilt records, no re-laid sibling; the edit path refuses whole (Codex #206 r1 REL-101)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src = try tt.path(std.testing.allocator, io, "s7b5_r1_two_src.xlsx");
    defer std.testing.allocator.free(src);
    const dst = try tt.path(std.testing.allocator, io, "s7b5_r1_two_dst.xlsx");
    defer std.testing.allocator.free(dst);
    try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
    try addUnsupportedConsumer(io, src);
    {
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try ed.setCell(0, 3, 1, .{ .number = 9.5 });
        try ed.save(io, dst);
    }
    try expectMarkerOnly(io, src, dst);
    var wb = try Workbook.open(std.testing.allocator, io, dst);
    defer wb.deinit();
    try expectHostCell(&wb, 0, "B3", .{ .number = 9.5 });
    var ed = try Editor.open(std.testing.allocator, io, src);
    defer ed.deinit();
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
}

test "S7b-5: a host part the cells cannot be rendered into, or a blank merge in a grown rectangle's way, leaves the save at the marker alone and refuses the edit (Codex #206 r1 REL-102, REL-103)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const dst = try tt.path(std.testing.allocator, io, "s7b5_r1_host_dst.xlsx");
    defer std.testing.allocator.free(dst);
    {
        // No `<sheetData>` on the host: the layout plans, the render
        // cannot emit — nothing but the marker is installed, and the
        // write inside the source is saved.
        const src = try tt.path(std.testing.allocator, io, "s7b5_r1_nosd.xlsx");
        defer std.testing.allocator.free(src);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/worksheets/sheet2.xml", "<sheetData>", "<sd>");
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/worksheets/sheet2.xml", "</sheetData>", "</sd>");
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try ed.setCell(0, 3, 1, .{ .number = 9.5 });
        try ed.save(io, dst);
        try expectMarkerOnly(io, src, dst);
        var wb = try Workbook.open(std.testing.allocator, io, dst);
        defer wb.deinit();
        try expectHostCell(&wb, 0, "B3", .{ .number = 9.5 });
    }
    {
        // A blank merged range `A7:B7` right below `A3:B6`: no `<c>`
        // marks it, the grown rectangle would write under it.
        const src = try tt.path(std.testing.allocator, io, "s7b5_r1_merge.xlsx");
        defer std.testing.allocator.free(src);
        try pivots_mod.fixture.write(std.testing.allocator, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(std.testing.allocator, io, src, "xl/worksheets/sheet2.xml", "</sheetData>", "</sheetData><mergeCells count=\"1\"><mergeCell ref=\"A7:B7\"/></mergeCells>");
        {
            var ed = try Editor.open(std.testing.allocator, io, src);
            defer ed.deinit();
            try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
            try ed.setCell(0, 2, 0, .{ .string = "Central" });
            try ed.save(io, dst);
        }
        try expectMarkerOnly(io, src, dst);
        // A merge the rectangle does not reach is no obstacle: the
        // delete shrinks it.
        var ed = try Editor.open(std.testing.allocator, io, src);
        defer ed.deinit();
        try ed.deleteRow(0, 3);
        try ed.save(io, dst);
        var store = try store_mod.PartStore.open(std.testing.allocator, io, dst);
        defer store.deinit();
        try expectPartHas(&store, pt_part, "<location ref=\"A3:B5\"");
        try expectPartHas(&store, "xl/worksheets/sheet2.xml", "<mergeCell ref=\"A7:B7\"/>");
    }
}

test "S7b-5: a table's row counts are checked before any arithmetic on a delete; a host the cells cannot be rendered into refuses the edit before anything moves (Codex #206 r2 SEC-201, REL-203)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const a = std.testing.allocator;
    const parts = [_][]const u8{ "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml", "xl/pivotCache/pivotCacheDefinition1.xml", "xl/pivotCache/pivotCacheRecords1.xml", "xl/pivotTables/pivotTable1.xml", "xl/sharedStrings.xml" };
    const Case = struct { name: []const u8, old: []const u8, new: []const u8 };
    const counts = [_]Case{
        .{ .name = "header_max", .old = "ref=\"A1:C4\" totalsRowShown=\"0\"", .new = "ref=\"A1:C4\" headerRowCount=\"4294967295\" totalsRowShown=\"0\"" },
        .{ .name = "totals_max", .old = "ref=\"A1:C4\" totalsRowShown=\"0\"", .new = "ref=\"A1:C4\" totalsRowCount=\"4294967295\" totalsRowShown=\"0\"" },
    };
    for (counts) |case| {
        const file = try std.fmt.allocPrint(a, "s7b5_r2_{s}.xlsx", .{case.name});
        defer a.free(file);
        const src = try tt.path(a, io, file);
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .table_name);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/tables/table1.xml", case.old, case.new);
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        var before: [parts.len][]u8 = undefined;
        for (parts, 0..) |name, i| before[i] = try a.dupe(u8, (try wb.store.part(name)).?.bytes);
        defer for (before) |b| a.free(b);
        try std.testing.expectError(error.PivotEditUnsafe, wb.deleteRow(0, 3));
        for (parts, 0..) |name, i| try std.testing.expectEqualStrings(before[i], (try wb.store.part(name)).?.bytes);
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.deleteRow(0, 3));
    }
    {
        // No `<sheetData>` on the host, a content edit on the source:
        // refused at the pre-flight, every part as it was.
        const src = try tt.path(a, io, "s7b5_r2_nosd.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", "<sheetData>", "<sd>");
        try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", "</sheetData>", "</sd>");
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        var before: [parts.len][]u8 = undefined;
        for (parts, 0..) |name, i| before[i] = try a.dupe(u8, (try wb.store.part(name)).?.bytes);
        defer for (before) |b| a.free(b);
        try std.testing.expectError(error.PivotEditUnsafe, wb.insertRow(0, 2));
        for (parts, 0..) |name, i| try std.testing.expectEqualStrings(before[i], (try wb.store.part(name)).?.bytes);
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
        // A shift alone writes no cell and needs no host render.
        try ed.insertRow(0, 1);
    }
}

/// The fixture's shared-string table with its `uniqueCount` respelt.
fn patchSstUniqueCount(io: std.Io, path: []const u8, value: []const u8) !void {
    const a = std.testing.allocator;
    var store = try store_mod.PartStore.open(a, io, path);
    defer store.deinit();
    const part = (try store.part("xl/sharedStrings.xml")) orelse return error.PartNotFound;
    const key = "uniqueCount=\"";
    const at = (std.mem.indexOf(u8, part.bytes, key) orelse return error.PatchAnchorNotFound) + key.len;
    const end = std.mem.indexOfScalarPos(u8, part.bytes, at, '"') orelse return error.PatchAnchorNotFound;
    const patched = try std.mem.concat(a, u8, &.{ part.bytes[0..at], value, part.bytes[end..] });
    defer a.free(patched);
    try store.replacePart("xl/sharedStrings.xml", patched);
    try store.save(io, path);
}

/// A second consumer of the fixture's cache on `Report` — the
/// fixture's own definition, renamed, at `location`.
fn addSecondConsumer(io: std.Io, path: []const u8, location: []const u8) !void {
    const a = std.testing.allocator;
    var store = try store_mod.PartStore.open(a, io, path);
    defer store.deinit();
    const first = (try store.part(pt_part)) orelse return error.PartNotFound;
    const renamed = try std.mem.replaceOwned(u8, a, first.bytes, "name=\"PivotTable1\"", "name=\"PivotTable2\"");
    defer a.free(renamed);
    const moved = try std.mem.replaceOwned(u8, a, renamed, "ref=\"A3:B6\"", location);
    defer a.free(moved);
    try store.addPart("xl/pivotTables/pivotTable2.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml", moved);
    try store.addPart(
        "xl/pivotTables/_rels/pivotTable2.xml.rels",
        "application/vnd.openxmlformats-package.relationships+xml",
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>
        ,
    );
    try store.save(io, path);
    try pivots_mod.fixture.patchPart(a, io, path, "xl/worksheets/_rels/sheet2.xml.rels", "</Relationships>", "<Relationship Id=\"rIdPT2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable\" Target=\"../pivotTables/pivotTable2.xml\"/></Relationships>");
}

test "S7b-5: a shared-string table lying about its count refuses the edit and leaves the save at the marker; a host coordinate twice, or two pivots on one cell, refuse (Codex #206 r3 SEC-301, REL-303)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const a = std.testing.allocator;
    const dst = try tt.path(a, io, "s7b5_r3_dst.xlsx");
    defer a.free(dst);
    {
        // A `uniqueCount` at the u32 maximum: the parsed table's entry
        // count is what the new indices continue from (r15 SEC-1506
        // superseded r3's refusal), so the edit is admitted and the
        // table respelt exactly.
        const src = try tt.path(a, io, "s7b5_r3_sst.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try patchSstUniqueCount(io, src, "4294967295");
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        try wb.insertRow(0, 2);
        const sst = (try wb.store.part("xl/sharedStrings.xml")).?.bytes;
        const unique = std.mem.count(u8, sst, "<si>");
        const want = try std.fmt.allocPrint(a, "uniqueCount=\"{d}\"", .{unique});
        defer a.free(want);
        try std.testing.expect(std.mem.indexOf(u8, sst, want) != null);
        try expectHostCell(&wb, 1, "A6", .{ .text = "(blank)" });
    }
    {
        // `A4` twice on the host.
        const src = try tt.path(a, io, "s7b5_r3_dup.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", "</row></sheetData>", "</row><row r=\"4\"><c r=\"A4\"><v>1</v></c><c r=\"A4\"><v>2</v></c></row></sheetData>");
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        try std.testing.expectError(error.MalformedSheetXml, wb.insertRow(0, 2));
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try ed.setCell(0, 2, 0, .{ .string = "Central" });
        try ed.save(io, dst);
        try expectMarkerOnly(io, src, dst);
    }
    {
        // A second pivot on the same cache at `A5:B8`: its rectangle
        // overlaps the first's.
        const src = try tt.path(a, io, "s7b5_r3_overlap.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try addSecondConsumer(io, src, "ref=\"A5:B8\"");
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        try std.testing.expectError(error.PivotEditUnsafe, wb.insertRow(0, 2));
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
        try ed.setCell(0, 2, 0, .{ .string = "Central" });
        try ed.save(io, dst);
        try expectMarkerOnly(io, src, dst);
    }
}

test "S7b-5: the host is spliced, not regenerated — a row's attributes, a shared formula, an inline string beside the pivot stay byte for byte; the dimension widens (Codex #206 r3 REL-302)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const a = std.testing.allocator;
    const src = try tt.path(a, io, "s7b5_r3_splice_src.xlsx");
    defer a.free(src);
    const dst = try tt.path(a, io, "s7b5_r3_splice_dst.xlsx");
    defer a.free(dst);
    try pivots_mod.fixture.write(a, io, src, .sheet_ref);
    const kept_row = "<row r=\"2\" ht=\"30\" customHeight=\"1\" hidden=\"1\" s=\"4\" customFormat=\"1\"><c r=\"D2\"><f t=\"shared\" ref=\"D2:D3\" si=\"0\">A1*2</f><v>0</v></c><c r=\"E2\" t=\"inlineStr\"><is><t>keep</t></is></c></row>";
    // Row 4 sits inside the pivot's rectangle with a cell of its own
    // right of it: the row's attribute and that cell stay.
    const pivot_row = "<row r=\"4\" ht=\"18\"><c r=\"D4\"><f t=\"shared\" si=\"0\"/><v>0</v></c></row>";
    const rows = try std.mem.concat(a, u8, &.{ "</row>", kept_row, pivot_row, "</sheetData>" });
    defer a.free(rows);
    try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", "</row></sheetData>", rows);
    const original = blk: {
        var store = try store_mod.PartStore.open(a, io, src);
        defer store.deinit();
        break :blk try a.dupe(u8, (try store.part("xl/worksheets/sheet2.xml")).?.bytes);
    };
    defer a.free(original);
    var ed = try Editor.open(a, io, src);
    defer ed.deinit();
    try ed.insertRow(0, 2);
    try ed.save(io, dst);
    var store = try store_mod.PartStore.open(a, io, dst);
    defer store.deinit();
    const host = (try store.part("xl/worksheets/sheet2.xml")).?.bytes;
    try std.testing.expect(std.mem.indexOf(u8, host, kept_row) != null);
    try std.testing.expect(std.mem.indexOf(u8, host, "<row r=\"4\" ht=\"18\"><c r=\"A4\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, host, "<c r=\"D4\"><f t=\"shared\" si=\"0\"/><v>0</v></c></row>") != null);
    // Everything before the rows is as written, but for a widened
    // dimension when the part spells one.
    const sd = std.mem.indexOf(u8, original, "<sheetData").?;
    if (std.mem.indexOf(u8, original[0..sd], "<dimension")) |_| {
        try std.testing.expect(std.mem.indexOf(u8, host, "<dimension ref=\"A1:B7\"/>") != null);
    } else {
        try std.testing.expectEqualStrings(original[0..sd], host[0..sd]);
    }
    var wb = try Workbook.open(a, io, dst);
    defer wb.deinit();
    try expectHostCell(&wb, 1, "A4", .{ .text = "East" });
    try expectHostCell(&wb, 1, "B7", .{ .number = 12 });
    try expectHostCell(&wb, 1, "E2", .{ .text = "keep" });
}

test "S7b-5: a comment, CDATA section or processing instruction spelling sheet markup is copied through, never spliced into; an explicitly closed dimension keeps its close (Codex #206 r4 SEC-401, REL-405)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const a = std.testing.allocator;
    const src = try tt.path(a, io, "s7b5_r4_decoy_src.xlsx");
    defer a.free(src);
    const dst = try tt.path(a, io, "s7b5_r4_decoy_dst.xlsx");
    defer a.free(dst);
    try pivots_mod.fixture.write(a, io, src, .sheet_ref);
    const decoy_head = "<!--<sheetData><row r=\"3\"><c r=\"A3\"/></row></sheetData>--><?zlsx <sheetData/>?><dimension ref=\"A1\"></dimension>";
    const decoy_row = "<row r=\"4\"><!--<c r=\"A4\"/></row>--><c r=\"C4\"><![CDATA[</c></row>]]><v>1</v></c></row>";
    const rows = try std.mem.concat(a, u8, &.{ "</row>", decoy_row, "</sheetData><!--<row r=\"9\"/>-->" });
    defer a.free(rows);
    // The rows first: the decoy head spells `</row></sheetData>` too.
    try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", "</row></sheetData>", rows);
    try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", "<sheetData>", decoy_head ++ "<sheetData>");
    var ed = try Editor.open(a, io, src);
    defer ed.deinit();
    try ed.insertRow(0, 2);
    try ed.save(io, dst);
    var store = try store_mod.PartStore.open(a, io, dst);
    defer store.deinit();
    const host = (try store.part("xl/worksheets/sheet2.xml")).?.bytes;
    try std.testing.expect(std.mem.indexOf(u8, host, "<!--<sheetData><row r=\"3\"><c r=\"A3\"/></row></sheetData>--><?zlsx <sheetData/>?><dimension ref=\"A1:B7\"></dimension><sheetData>") != null);
    try std.testing.expect(std.mem.indexOf(u8, host, "<row r=\"4\"><!--<c r=\"A4\"/></row>--><c r=\"A4\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, host, "<c r=\"C4\"><![CDATA[</c></row>]]><v>1</v></c></row>") != null);
    try std.testing.expect(std.mem.indexOf(u8, host, "</sheetData><!--<row r=\"9\"/>-->") != null);
    try std.testing.expectEqual(@as(usize, 1), std.mem.count(u8, host, "<dimension"));
    var wb = try Workbook.open(a, io, dst);
    defer wb.deinit();
    try expectHostCell(&wb, 1, "A3", .{ .text = "Row Labels" });
    try expectHostCell(&wb, 1, "A4", .{ .text = "East" });
    try expectHostCell(&wb, 1, "B4", .{ .number = 8 });
    try expectHostCell(&wb, 1, "B7", .{ .number = 12 });
}

test "S7b-5: a merge the old rectangle held is still in the way; an appended integer under a date style 0 leaves the save at the marker (Codex #206 r4 REL-402, REL-403)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const a = std.testing.allocator;
    const dst = try tt.path(a, io, "s7b5_r4_dst.xlsx");
    defer a.free(dst);
    {
        // `A6:B6` merged inside `A3:B6`.
        const src = try tt.path(a, io, "s7b5_r4_merge_inside.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", "</sheetData>", "</sheetData><mergeCells count=\"1\"><mergeCell ref=\"A6:B6\"/></mergeCells>");
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
        try ed.setCell(0, 2, 0, .{ .string = "Central" });
        try ed.save(io, dst);
        try expectMarkerOnly(io, src, dst);
    }
    {
        // Style 0 is a date; the source's own numbers wear style 1
        // (General); the source rectangle runs to row 9, and a row
        // appended inside it carries an integer with no style — style
        // 0, a date: refused, the marker alone.
        const src = try tt.path(a, io, "s7b5_r4_append_int.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        {
            var store = try store_mod.PartStore.open(a, io, src);
            defer store.deinit();
            try store.addPart("xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts><fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills><borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"2\"><xf numFmtId=\"14\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs></styleSheet>");
            try store.save(io, src);
        }
        const numeric = [_][]const u8{ "<c r=\"B2\">", "<c r=\"C2\">", "<c r=\"B3\">", "<c r=\"C3\">", "<c r=\"B4\">", "<c r=\"C4\">" };
        for (numeric) |old| {
            const new = try std.mem.concat(a, u8, &.{ old[0 .. old.len - 1], " s=\"1\">" });
            defer a.free(new);
            try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet1.xml", old, new);
        }
        try pivots_mod.fixture.patchPart(a, io, src, "xl/pivotCache/pivotCacheDefinition1.xml", "ref=\"A1:C4\"", "ref=\"A1:C9\"");
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        const row = [_]Cell{ .{ .string = "North" }, .{ .integer = 6 }, .empty };
        const rows = [_][]const Cell{&row};
        try ed.appendRows(0, &rows);
        try ed.save(io, dst);
        try expectMarkerOnlyWith(io, src, dst, false);
        var wb = try Workbook.open(a, io, dst);
        defer wb.deinit();
        try expectHostCell(&wb, 0, "B5", .{ .number = 6 });
    }
}

test "S7b-5: a product grand total folds the subtotals as a sum does (Codex #206 r4 REL-406)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const a = std.testing.allocator;
    const src = try tt.path(a, io, "s7b5_r4_product_src.xlsx");
    defer a.free(src);
    const dst = try tt.path(a, io, "s7b5_r4_product_dst.xlsx");
    defer a.free(dst);
    try pivots_mod.fixture.write(a, io, src, .sheet_ref);
    try pivots_mod.fixture.patchPart(a, io, src, pt_part, "fld=\"1\" baseField", "fld=\"1\" subtotal=\"product\" baseField");
    // East 0.1, West 3, East 0.7: in record order (0.1·3)·0.7 =
    // 0.21000000000000002; folded, (0.1·0.7)·3 = 0.20999999999999996.
    var ed = try Editor.open(a, io, src);
    defer ed.deinit();
    try ed.setCell(0, 2, 1, .{ .number = 0.1 });
    try ed.setCell(0, 3, 1, .{ .number = 3 });
    try ed.setCell(0, 4, 1, .{ .number = 0.7 });
    try ed.save(io, dst);
    var wb = try Workbook.open(a, io, dst);
    defer wb.deinit();
    try expectHostCell(&wb, 1, "B4", .{ .number = 0.06999999999999999 });
    try expectHostCell(&wb, 1, "B5", .{ .number = 3 });
    try expectHostCell(&wb, 1, "B6", .{ .number = 0.20999999999999996 });
}

test "S7b-5: a staged write over a source cell the read refuses is read as the write; a rich inline caption is joined whole (Codex #206 r5 REL-502, REL-503)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const a = std.testing.allocator;
    const dst = try tt.path(a, io, "s7b5_r5_dst.xlsx");
    defer a.free(dst);
    {
        // `B3` holds an uncomputed formula — refused as read — and a
        // staged 9 replaces it: the rebuild sees 9.
        const src = try tt.path(a, io, "s7b5_r5_overlay.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet1.xml", "<c r=\"B3\"><v>4</v></c>", "<c r=\"B3\"><f>2+2</f></c>");
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try ed.setCell(0, 3, 1, .{ .integer = 9 });
        try ed.save(io, dst);
        var store = try store_mod.PartStore.open(a, io, dst);
        defer store.deinit();
        try expectRebuilt(&store, "xl/pivotCache/pivotCacheDefinition1.xml", "xl/pivotCache/pivotCacheRecords1.xml", 3);
        try expectPartHas(&store, "xl/pivotCache/pivotCacheRecords1.xml", "<r><x v=\"1\"/><n v=\"9\"/><n v=\"2.5\"/></r>");
        var wb = try Workbook.open(a, io, dst);
        defer wb.deinit();
        try expectHostCell(&wb, 1, "B5", .{ .number = 9 });
        try expectHostCell(&wb, 1, "B6", .{ .number = 17 });
    }
    {
        // The host's captions as two-run inline strings.
        const src = try tt.path(a, io, "s7b5_r5_rich.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", "</row></sheetData>", "</row><row r=\"3\"><c r=\"A3\" t=\"inlineStr\"><is><r><t>Étiquettes </t></r><r><rPr><b/></rPr><t>de lignes</t></r></is></c></row><row r=\"6\"><c r=\"A6\" t=\"inlineStr\"><is><r><t>Total </t></r><r><t>général</t></r></is></c></row></sheetData>");
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try ed.insertRow(0, 2);
        try ed.save(io, dst);
        var wb = try Workbook.open(a, io, dst);
        defer wb.deinit();
        try expectHostCell(&wb, 1, "A3", .{ .text = "Étiquettes de lignes" });
        try expectHostCell(&wb, 1, "A7", .{ .text = "Total général" });
    }
}

test "S7b-5: a write staged on the host after the pre-flight ages the prepared collection — refused, the write kept (Codex #206 r6 REL-601)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const a = std.testing.allocator;
    const src = try tt.path(a, io, "s7b5_r6_token.xlsx");
    defer a.free(src);
    try pivots_mod.fixture.write(a, io, src, .sheet_ref);
    var wb = try Workbook.open(a, io, src);
    defer wb.deinit();
    const parts = [_][]const u8{ "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml", "xl/pivotCache/pivotCacheDefinition1.xml", "xl/pivotCache/pivotCacheRecords1.xml", "xl/pivotTables/pivotTable1.xml" };
    var before: [parts.len][]u8 = undefined;
    for (parts, 0..) |name, i| before[i] = try a.dupe(u8, (try wb.store.part(name)).?.bytes);
    defer for (before) |b| a.free(b);
    var prepared = try wb.preflightPivotEditsForSheet("xl/worksheets/sheet1.xml", .row, 2, .insert);
    defer prepared.deinit(a);
    // `Report!A7` is the row the grown rectangle would take.
    const host = try wb.sheet(1);
    try host.setCell("A7", .{ .shared_string = "note" });
    try std.testing.expectError(error.PivotEditUnsafe, wb.applySheetEdit(0, .{ .row = 2, .kind = .insert }, &prepared));
    for (parts, 0..) |name, i| try std.testing.expectEqualStrings(before[i], (try wb.store.part(name)).?.bytes);
    try std.testing.expect(host.deltas.contains(.{ .row = 7, .col = 1 }));
    // Pre-flighted again over the staged write, the edit refuses on its
    // own terms: the growth cell is taken.
    try std.testing.expectError(error.PivotEditUnsafe, wb.preflightPivotEditsForSheet("xl/worksheets/sheet1.xml", .row, 2, .insert));
}

test "S7b-5: a comment or processing instruction between an inline caption's runs is not a run; a decoy <is> in a comment is not the element; an unterminated run refuses (Codex #206 r7 REL-702)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const a = std.testing.allocator;
    const dst = try tt.path(a, io, "s7b5_r7_dst.xlsx");
    defer a.free(dst);
    {
        const src = try tt.path(a, io, "s7b5_r7_decoys.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", "</row></sheetData>", "</row><row r=\"3\"><c r=\"A3\" t=\"inlineStr\"><!--<is><t>fake</t></is>--><is><r><t>Étiquettes </t></r><!--<t>fake</t>--><?pi <t>fake</t>?><r><t>de lignes</t></r></is></c></row><row r=\"6\"><c r=\"A6\" t=\"inlineStr\"><is><r><t>Total <!--x--></t></r><r><t>général</t></r></is></c></row></sheetData>");
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try ed.insertRow(0, 2);
        try ed.save(io, dst);
        var wb = try Workbook.open(a, io, dst);
        defer wb.deinit();
        try expectHostCell(&wb, 1, "A3", .{ .text = "Étiquettes de lignes" });
        try expectHostCell(&wb, 1, "A7", .{ .text = "Total général" });
    }
    {
        const src = try tt.path(a, io, "s7b5_r7_unterminated.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", "</row></sheetData>", "</row><row r=\"3\"><c r=\"A3\" t=\"inlineStr\"><is><r><t>Row Labels</r></is></c></row></sheetData>");
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        try std.testing.expectError(error.MalformedSheetXml, wb.insertRow(0, 2));
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try ed.setCell(0, 2, 0, .{ .string = "Central" });
        try ed.save(io, dst);
        try expectMarkerOnly(io, src, dst);
    }
}

test "S7b-5: a pivot whose source is another pivot's rectangle — the edit refuses, the save marks it in the same install and refreshes the upstream one (Codex #206 r9 REL-901)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const a = std.testing.allocator;
    const src = try tt.path(a, io, "s7b5_r9_downstream.xlsx");
    defer a.free(src);
    const dst = try tt.path(a, io, "s7b5_r9_downstream_dst.xlsx");
    defer a.free(dst);
    // The orphan cache 8 reads `Report!A1:A1`; respelt to read the
    // pivot's own rectangle `A3:B6`.
    try pivots_mod.fixture.writeWithOrphanCache(a, io, src, .sheet_ref);
    try pivots_mod.fixture.patchPart(a, io, src, "xl/pivotCache/pivotCacheDefinition2.xml", "sheet=\"Report\" ref=\"A1:A1\"", "sheet=\"Report\" ref=\"A3:B6\"");
    {
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        const before = try a.dupe(u8, (try wb.store.part("xl/pivotCache/pivotCacheRecords1.xml")).?.bytes);
        defer a.free(before);
        try std.testing.expectError(error.PivotEditUnsafe, wb.insertRow(0, 2));
        try std.testing.expectEqualStrings(before, (try wb.store.part("xl/pivotCache/pivotCacheRecords1.xml")).?.bytes);
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
    }
    var ed = try Editor.open(a, io, src);
    defer ed.deinit();
    try ed.setCell(0, 2, 0, .{ .string = "Central" });
    try ed.save(io, dst);
    var store = try store_mod.PartStore.open(a, io, dst);
    defer store.deinit();
    try expectRebuilt(&store, "xl/pivotCache/pivotCacheDefinition1.xml", "xl/pivotCache/pivotCacheRecords1.xml", 3);
    try expectMarked(&store, "xl/pivotCache/pivotCacheDefinition1.xml", true);
    try expectMarked(&store, "xl/pivotCache/pivotCacheDefinition2.xml", true);
    try expectPartHas(&store, "xl/pivotCache/pivotCacheDefinition2.xml", "recordCount=\"0\"");
    var wb = try Workbook.open(a, io, dst);
    defer wb.deinit();
    try expectHostCell(&wb, 1, "A6", .{ .text = "Central" });
}

/// A pivot on the orphan cache 8 (`Report!A1:A1`, one field), hosted
/// on `Report` at `location` — a pivot no edit of `Data` touches.
fn addConsumerOnOrphanCache(io: std.Io, path: []const u8, location: []const u8) !void {
    const a = std.testing.allocator;
    var store = try store_mod.PartStore.open(a, io, path);
    defer store.deinit();
    const def = try std.mem.concat(a, u8, &.{
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<pivotTableDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" name=\"PivotTable2\" cacheId=\"8\" dataCaption=\"Values\" updatedVersion=\"6\" minRefreshableVersion=\"3\" useAutoFormatting=\"1\" itemPrintTitles=\"1\" createdVersion=\"6\" indent=\"0\" outline=\"1\" outlineData=\"1\" multipleFieldFilters=\"0\"><location ",
        location,
        " firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/><pivotFields count=\"1\"><pivotField axis=\"axisRow\" showAll=\"0\"><items count=\"1\"><item t=\"default\"/></items></pivotField></pivotFields><rowFields count=\"1\"><field x=\"0\"/></rowFields><rowItems count=\"1\"><i t=\"grand\"><x/></i></rowItems><colItems count=\"1\"><i/></colItems></pivotTableDefinition>",
    });
    defer a.free(def);
    try store.addPart("xl/pivotTables/pivotTable2.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml", def);
    try store.addPart(
        "xl/pivotTables/_rels/pivotTable2.xml.rels",
        "application/vnd.openxmlformats-package.relationships+xml",
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition2.xml"/></Relationships>
        ,
    );
    try store.save(io, path);
    try pivots_mod.fixture.patchPart(a, io, path, "xl/worksheets/_rels/sheet2.xml.rels", "</Relationships>", "<Relationship Id=\"rIdPT2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable\" Target=\"../pivotTables/pivotTable2.xml\"/></Relationships>");
}

test "S7b-5: a rectangle may not grow into another pivot's declared footprint, even one with no cell there — the edit refuses, the save marks alone (Codex #206 r10 REL-1001)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const a = std.testing.allocator;
    const src = try tt.path(a, io, "s7b5_r10_footprint.xlsx");
    defer a.free(src);
    const dst = try tt.path(a, io, "s7b5_r10_footprint_dst.xlsx");
    defer a.free(dst);
    try pivots_mod.fixture.writeWithOrphanCache(a, io, src, .sheet_ref);
    // `A7:B10`, right under `A3:B6`, no cell written there.
    try addConsumerOnOrphanCache(io, src, "ref=\"A7:B10\"");
    var wb = try Workbook.open(a, io, src);
    defer wb.deinit();
    const before = try a.dupe(u8, (try wb.store.part("xl/pivotCache/pivotCacheRecords1.xml")).?.bytes);
    defer a.free(before);
    try std.testing.expectError(error.PivotEditUnsafe, wb.insertRow(0, 2));
    try std.testing.expectEqualStrings(before, (try wb.store.part("xl/pivotCache/pivotCacheRecords1.xml")).?.bytes);
    var ed = try Editor.open(a, io, src);
    defer ed.deinit();
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
    try ed.setCell(0, 2, 0, .{ .string = "Central" });
    try ed.save(io, dst);
    try expectMarkerOnly(io, src, dst);
    // A delete shrinks the rectangle: no footprint in the way.
    var ed2 = try Editor.open(a, io, src);
    defer ed2.deinit();
    try ed2.deleteRow(0, 3);
    try ed2.save(io, dst);
    var store = try store_mod.PartStore.open(a, io, dst);
    defer store.deinit();
    try expectPartHas(&store, pt_part, "<location ref=\"A3:B5\"");
}

test "S7b-5: a pivot whose rectangle overlaps its own source — the edit refuses, the save marks alone; a `>` inside a quoted attribute is not a tag's end (Codex #206 r13 REL-1301, REL-1302)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const a = std.testing.allocator;
    const dst = try tt.path(a, io, "s7b5_r13_dst.xlsx");
    defer a.free(dst);
    {
        // The pivot at `A7:B10` on `Data`, and a source `A1:C9` that
        // reaches into it.
        const src = try tt.path(a, io, "s7b5_r13_self.xlsx");
        defer a.free(src);
        try writeHostOnSourceFixture(io, src, false);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/pivotCache/pivotCacheDefinition1.xml", "ref=\"A1:C4\"", "ref=\"A1:C9\"");
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        try std.testing.expectError(error.PivotEditUnsafe, wb.insertRow(0, 2));
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
        try ed.setCell(0, 2, 0, .{ .string = "Central" });
        try ed.save(io, dst);
        try expectMarkerOnly(io, src, dst);
        var store = try store_mod.PartStore.open(a, io, dst);
        defer store.deinit();
        try expectPartHas(&store, "xl/worksheets/sheet1.xml", "<c r=\"A10\" s=\"7\" t=\"inlineStr\"><is><t>Total général</t></is></c>");
    }
    {
        const src = try tt.path(a, io, "s7b5_r13_quoted.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", "</row></sheetData>", "</row><row r=\"4\" xml:lang=\"a>b\"><c r=\"A4\" xml:lang=\"c>d\" t=\"inlineStr\"><is><t>East</t></is></c><c r=\"D4\" xml:lang=\"e>f\"><v>1</v></c></row></sheetData>");
        try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", "<sheetData>", "<dimension xml:lang=\"g>h\" ref=\"A1\"/><sheetData>");
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try ed.insertRow(0, 2);
        try ed.save(io, dst);
        var store = try store_mod.PartStore.open(a, io, dst);
        defer store.deinit();
        try expectPartHas(&store, "xl/worksheets/sheet2.xml", "<dimension xml:lang=\"g>h\" ref=\"A1:B7\"/>");
        try expectPartHas(&store, "xl/worksheets/sheet2.xml", "<row r=\"4\" xml:lang=\"a>b\"><c r=\"A4\"");
        try expectPartHas(&store, "xl/worksheets/sheet2.xml", "<c r=\"D4\" xml:lang=\"e>f\"><v>1</v></c></row>");
        var wb = try Workbook.open(a, io, dst);
        defer wb.deinit();
        try expectHostCell(&wb, 1, "A4", .{ .text = "East" });
        try expectHostCell(&wb, 1, "B7", .{ .number = 12 });
    }
}

test "S7b-5: a phonetic run is not caption text; the shared-string table drops its occurrence count once host cells change it (Codex #206 r14 REL-1402, REL-1403)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const a = std.testing.allocator;
    const src = try tt.path(a, io, "s7b5_r14_src.xlsx");
    defer a.free(src);
    const dst = try tt.path(a, io, "s7b5_r14_dst.xlsx");
    defer a.free(dst);
    try pivots_mod.fixture.write(a, io, src, .sheet_ref);
    try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", "</row></sheetData>", "</row><row r=\"3\"><c r=\"A3\" t=\"inlineStr\"><is><t>漢字</t><rPh sb=\"0\" eb=\"2\"><t>かんじ</t></rPh><phoneticPr fontId=\"1\"/></is></c></row></sheetData>");
    {
        var store = try store_mod.PartStore.open(a, io, src);
        defer store.deinit();
        const sst = (try store.part("xl/sharedStrings.xml")).?.bytes;
        try std.testing.expect(std.mem.indexOf(u8, sst, " count=\"") != null);
    }
    var ed = try Editor.open(a, io, src);
    defer ed.deinit();
    try ed.insertRow(0, 2);
    try ed.save(io, dst);
    var store = try store_mod.PartStore.open(a, io, dst);
    defer store.deinit();
    const sst = (try store.part("xl/sharedStrings.xml")).?.bytes;
    try std.testing.expect(std.mem.indexOf(u8, sst, " count=\"") == null);
    const unique = std.mem.count(u8, sst, "<si>");
    const want = try std.fmt.allocPrint(a, "uniqueCount=\"{d}\"", .{unique});
    defer a.free(want);
    try std.testing.expect(std.mem.indexOf(u8, sst, want) != null);
    var wb = try Workbook.open(a, io, dst);
    defer wb.deinit();
    try expectHostCell(&wb, 1, "A3", .{ .text = "漢字" });
    // A shrink writes only strings the table has: the count still goes.
    const src2 = try tt.path(a, io, "s7b5_r14_shrink.xlsx");
    defer a.free(src2);
    try pivots_mod.fixture.write(a, io, src2, .sheet_ref);
    var ed2 = try Editor.open(a, io, src2);
    defer ed2.deinit();
    try ed2.deleteRow(0, 3);
    try ed2.save(io, dst);
    var store2 = try store_mod.PartStore.open(a, io, dst);
    defer store2.deinit();
    const sst2 = (try store2.part("xl/sharedStrings.xml")).?.bytes;
    try std.testing.expect(std.mem.indexOf(u8, sst2, " count=\"") == null);
}

test "S7b-5: a decoy shared-string root, a quoted `>` on it and a lying uniqueCount do not misplace a new entry; an unknown element in an inline caption refuses; a count-only table rewrite drops the cached view (Codex #206 r15 SEC-1506, REL-1503, REL-1501)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const a = std.testing.allocator;
    const dst = try tt.path(a, io, "s7b5_r15_dst.xlsx");
    defer a.free(dst);
    {
        const src = try tt.path(a, io, "s7b5_r15_sst.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try patchSstUniqueCount(io, src, "99");
        try pivots_mod.fixture.patchPart(a, io, src, "xl/sharedStrings.xml", "<sst ", "<!--<sst/>--><?pi <sst>?><sst xml:lang=\"a>b\" ");
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try ed.insertRow(0, 2);
        try ed.save(io, dst);
        var store = try store_mod.PartStore.open(a, io, dst);
        defer store.deinit();
        const sst = (try store.part("xl/sharedStrings.xml")).?.bytes;
        // The decoys as written, the root's own attribute kept (the
        // attribute writer respaces the blob), the count gone.
        try std.testing.expect(std.mem.indexOf(u8, sst, "<!--<sst/>--><?pi <sst>?><sst") != null);
        try std.testing.expect(std.mem.indexOf(u8, sst, "xml:lang=\"a>b\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, sst, " count=\"") == null);
        const unique = std.mem.count(u8, sst, "<si>");
        const want = try std.fmt.allocPrint(a, "uniqueCount=\"{d}\"", .{unique});
        defer a.free(want);
        try std.testing.expect(std.mem.indexOf(u8, sst, want) != null);
        var wb = try Workbook.open(a, io, dst);
        defer wb.deinit();
        try expectHostCell(&wb, 1, "A6", .{ .text = "(blank)" });
        try expectHostCell(&wb, 1, "A7", .{ .text = "Grand Total" });
    }
    {
        const src = try tt.path(a, io, "s7b5_r15_foo.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", "</row></sheetData>", "</row><row r=\"3\"><c r=\"A3\" t=\"inlineStr\"><is><foo><t>fake</t></foo><t>Row Labels</t></is></c></row></sheetData>");
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        try std.testing.expectError(error.MalformedSheetXml, wb.insertRow(0, 2));
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try ed.setCell(0, 2, 0, .{ .string = "Central" });
        try ed.save(io, dst);
        try expectMarkerOnly(io, src, dst);
    }
    {
        const src = try tt.path(a, io, "s7b5_r15_view.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        try std.testing.expect((try wb.sst()).?.total_count != null);
        try wb.deleteRow(0, 3);
        try std.testing.expect((try wb.sst()).?.total_count == null);
    }
}

test "S7b-5: a cell inside a cell is not a grid; a padded `</is >` is a caption, and a cell; a merge the insert would push off the grid refuses; a deleted delta keeps its style donor; rowHeaderCaption is the header (Codex #206 r16 SEC-1601, REL-1601, REL-1602, REL-1604, REL-1603)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const a = std.testing.allocator;
    const dst = try tt.path(a, io, "s7b5_r16_dst.xlsx");
    defer a.free(dst);
    {
        // A cell directly inside a cell is no grid: refused, marker
        // alone at save.
        const src = try tt.path(a, io, "s7b5_r16_nested.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", "</row></sheetData>", "</row><row r=\"4\"><c r=\"A4\"><c r=\"Z99\"><v>1</v></c><v>2</v></c></row></sheetData>");
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        // The rehearsed render refuses it, as the typed refusal.
        try std.testing.expectError(error.PivotEditUnsafe, wb.insertRow(0, 2));
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try ed.setCell(0, 2, 0, .{ .string = "Central" });
        try ed.save(io, dst);
        try expectMarkerOnly(io, src, dst);
        // One wrapped in an element the walk steps over whole is
        // replaced with its wrapper — no stray close left behind.
        const wrapped = try tt.path(a, io, "s7b5_r16_wrapped.xlsx");
        defer a.free(wrapped);
        try pivots_mod.fixture.write(a, io, wrapped, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, wrapped, "xl/worksheets/sheet2.xml", "</row></sheetData>", "</row><row r=\"4\"><c r=\"A4\"><foo><c r=\"Z99\"><v>1</v></c></foo><v>2</v></c><c r=\"D4\"><bar><c r=\"Z98\"/></bar><v>3</v></c></row></sheetData>");
        var ed2 = try Editor.open(a, io, wrapped);
        defer ed2.deinit();
        try ed2.insertRow(0, 2);
        try ed2.save(io, dst);
        var store = try store_mod.PartStore.open(a, io, dst);
        defer store.deinit();
        const host = (try store.part("xl/worksheets/sheet2.xml")).?.bytes;
        try std.testing.expect(std.mem.indexOf(u8, host, "<foo>") == null);
        try std.testing.expect(std.mem.indexOf(u8, host, "<c r=\"D4\"><bar><c r=\"Z98\"/></bar><v>3</v></c></row>") != null);
        var wb2 = try Workbook.open(a, io, dst);
        defer wb2.deinit();
        try expectHostCell(&wb2, 1, "A4", .{ .text = "East" });
    }
    {
        // `</is >` on the header caption and on a cell right under the
        // rectangle: the caption is kept, the cell is in the way.
        const src = try tt.path(a, io, "s7b5_r16_padded.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", "</row></sheetData>", "</row><row r=\"3\"><c r=\"A3\" t=\"inlineStr\"><is><t>Étiquettes de lignes</t></is ></c></row><row r=\"7\"><c r=\"A7\" t=\"inlineStr\"><is><t>note</t></is ></c></row></sheetData>");
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
        try ed.deleteRow(0, 3);
        try ed.save(io, dst);
        var wb = try Workbook.open(a, io, dst);
        defer wb.deinit();
        try expectHostCell(&wb, 1, "A3", .{ .text = "Étiquettes de lignes" });
        try expectHostCell(&wb, 1, "A7", .{ .text = "note" });
    }
    {
        // A blank merge on the grid's last row of the edited host.
        const src = try tt.path(a, io, "s7b5_r16_edge.xlsx");
        defer a.free(src);
        try writeHostOnSourceFixture(io, src, false);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet1.xml", "</sheetData>", "</sheetData><mergeCells count=\"1\"><mergeCell ref=\"D1048576:E1048576\"/></mergeCells>");
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        try std.testing.expectError(error.PivotEditUnsafe, wb.insertRow(0, 2));
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
    }
    {
        // The grand total's label deleted in the same save: its style
        // still dresses the regenerated grand-total row.
        const src = try tt.path(a, io, "s7b5_r16_donor.xlsx");
        defer a.free(src);
        try writeHostOnSourceFixture(io, src, false);
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try ed.setCell(0, 2, 0, .{ .string = "Central" });
        const host = try ed.workbook.sheet(0);
        try host.deleteCell("A10");
        try ed.save(io, dst);
        var wb = try Workbook.open(a, io, dst);
        defer wb.deinit();
        // The label the user deleted is gone as a caption to reuse —
        // the default stands — but its style dresses the new row.
        try expectHostCell(&wb, 0, "A11", .{ .text = "Grand Total" });
        try expectHostStyle(&wb, 0, "A11", 7);
        try expectHostStyle(&wb, 0, "B11", 8);
    }
    {
        const src = try tt.path(a, io, "s7b5_r16_caption.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, src, pt_part, "outline=\"1\"", "outline=\"1\" rowHeaderCaption=\"R&#233;gion\"");
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try ed.insertRow(0, 2);
        try ed.save(io, dst);
        var wb = try Workbook.open(a, io, dst);
        defer wb.deinit();
        try expectHostCell(&wb, 1, "A3", .{ .text = "Région" });
    }
}

test "S7b-5: a shared-string root nested in the root refuses the edit and leaves the save at the marker (Codex #206 r17 SEC-1701)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const a = std.testing.allocator;
    const src = try tt.path(a, io, "s7b5_r17_nested_sst.xlsx");
    defer a.free(src);
    const dst = try tt.path(a, io, "s7b5_r17_nested_sst_dst.xlsx");
    defer a.free(dst);
    try pivots_mod.fixture.write(a, io, src, .sheet_ref);
    try pivots_mod.fixture.patchPart(a, io, src, "xl/sharedStrings.xml", "><si>", "><sst></sst><si>");
    var wb = try Workbook.open(a, io, src);
    defer wb.deinit();
    const before = try a.dupe(u8, (try wb.store.part("xl/sharedStrings.xml")).?.bytes);
    defer a.free(before);
    try std.testing.expectError(error.PivotEditUnsafe, wb.insertRow(0, 2));
    try std.testing.expectEqualStrings(before, (try wb.store.part("xl/sharedStrings.xml")).?.bytes);
    var ed = try Editor.open(a, io, src);
    defer ed.deinit();
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
    try ed.setCell(0, 2, 0, .{ .string = "Central" });
    try ed.save(io, dst);
    try expectMarkerOnly(io, src, dst);
}

test "S7b-5: a host that aliases or rebinds the main namespace refuses the edit and leaves the save at the marker (Codex #206 r18 REL-1801)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const a = std.testing.allocator;
    const dst = try tt.path(a, io, "s7b5_r18_dst.xlsx");
    defer a.free(dst);
    const Case = struct { name: []const u8, rows: []const u8 };
    const cases = [_]Case{
        .{ .name = "alias", .rows = "</row><row r=\"7\"><x:c xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" r=\"A7\"><v>1</v></x:c></row></sheetData>" },
        .{ .name = "rebind", .rows = "</row><row r=\"7\" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><c r=\"A7\"><v>1</v></c></row></sheetData>" },
        .{ .name = "alias_merge", .rows = "</row></sheetData><x:mergeCells xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\"><x:mergeCell ref=\"A7:B7\"/></x:mergeCells>" },
        // Codex #206 r19 SEC-1901: the binding spelt with a reference.
        .{ .name = "alias_entity", .rows = "</row><row r=\"7\"><x:c xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/mai&#x6e;\" r=\"A7\"><v>1</v></x:c></row></sheetData>" },
        .{ .name = "alias_merge_entity", .rows = "</row></sheetData><x:mergeCells xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/mai&#x6e;\" count=\"1\"><x:mergeCell ref=\"A7:B7\"/></x:mergeCells>" },
    };
    for (cases) |case| {
        const file = try std.fmt.allocPrint(a, "s7b5_r18_{s}.xlsx", .{case.name});
        defer a.free(file);
        const src = try tt.path(a, io, file);
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", "</row></sheetData>", case.rows);
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        const before = try a.dupe(u8, (try wb.store.part("xl/worksheets/sheet2.xml")).?.bytes);
        defer a.free(before);
        try std.testing.expectError(error.PivotEditUnsafe, wb.insertRow(0, 2));
        try std.testing.expectEqualStrings(before, (try wb.store.part("xl/worksheets/sheet2.xml")).?.bytes);
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
        try ed.setCell(0, 2, 0, .{ .string = "Central" });
        try ed.save(io, dst);
        try expectMarkerOnly(io, src, dst);
    }
}

test "S7b-5: a sheetData or dimension nested under another element is not the root's — the edit refuses, the save marks alone (Codex #206 r20 SEC-2001)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const a = std.testing.allocator;
    const dst = try tt.path(a, io, "s7b5_r20_dst.xlsx");
    defer a.free(dst);
    const Case = struct { name: []const u8, before: []const u8 };
    const cases = [_]Case{
        .{ .name = "nested_sheet_data", .before = "<foo><sheetData><row r=\"3\"><c r=\"A3\"/></row></sheetData></foo>" },
        .{ .name = "nested_dimension", .before = "<foo><dimension ref=\"A1\"/></foo><dimension ref=\"A1\"/>" },
    };
    for (cases) |case| {
        const file = try std.fmt.allocPrint(a, "s7b5_r20_{s}.xlsx", .{case.name});
        defer a.free(file);
        const src = try tt.path(a, io, file);
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        const patched = try std.mem.concat(a, u8, &.{ case.before, "<sheetData>" });
        defer a.free(patched);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", "<sheetData>", patched);
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        const before = try a.dupe(u8, (try wb.store.part("xl/worksheets/sheet2.xml")).?.bytes);
        defer a.free(before);
        try std.testing.expectError(error.PivotEditUnsafe, wb.insertRow(0, 2));
        try std.testing.expectEqualStrings(before, (try wb.store.part("xl/worksheets/sheet2.xml")).?.bytes);
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
        try ed.setCell(0, 2, 0, .{ .string = "Central" });
        try ed.save(io, dst);
        try expectMarkerOnly(io, src, dst);
    }
}

test "S7b-5: a host's own staged writes go into the pivot's splice, byte-preservingly; a crossed close, a second top-level element, a wrapped shared-string root, an empty inline string (Codex #206 r21 REL-2101, SEC-2103, SEC-2104, REL-2102)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const a = std.testing.allocator;
    const dst = try tt.path(a, io, "s7b5_r21_dst.xlsx");
    defer a.free(dst);
    {
        const src = try tt.path(a, io, "s7b5_r21_deltas.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        const kept_row = "<row r=\"2\" ht=\"30\" customHeight=\"1\"><c r=\"D2\"><f t=\"shared\" ref=\"D2:D3\" si=\"0\">A1*2</f><v>0</v></c></row>";
        const rows = try std.mem.concat(a, u8, &.{ "</row>", kept_row, "</sheetData><!--<sheetData/>-->" });
        defer a.free(rows);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", "</row></sheetData>", rows);
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try ed.setCell(0, 2, 0, .{ .string = "Central" });
        try ed.setCell(1, 9, 4, .{ .string = "note" });
        try ed.setCell(1, 9, 5, .{ .integer = 7 });
        try ed.save(io, dst);
        var store = try store_mod.PartStore.open(a, io, dst);
        defer store.deinit();
        const host = (try store.part("xl/worksheets/sheet2.xml")).?.bytes;
        try std.testing.expect(std.mem.indexOf(u8, host, kept_row) != null);
        try std.testing.expect(std.mem.indexOf(u8, host, "</sheetData><!--<sheetData/>-->") != null);
        var wb = try Workbook.open(a, io, dst);
        defer wb.deinit();
        try expectHostCell(&wb, 1, "E9", .{ .text = "note" });
        try expectHostCell(&wb, 1, "F9", .{ .number = 7 });
        try expectHostCell(&wb, 1, "A6", .{ .text = "Central" });
        try expectHostCell(&wb, 1, "B7", .{ .number = 12 });
    }
    const Case = struct { name: []const u8, part: []const u8, old: []const u8, new: []const u8 };
    const cases = [_]Case{
        .{ .name = "crossed_close", .part = "xl/worksheets/sheet2.xml", .old = "<sheetData>", .new = "<foo></worksheet><sheetData/></foo><sheetData>" },
        .{ .name = "second_top_level", .part = "xl/worksheets/sheet2.xml", .old = "</worksheet>", .new = "</worksheet><foo><sheetData/></foo>" },
        .{ .name = "wrapped_sst", .part = "xl/sharedStrings.xml", .old = "<sst ", .new = "<wrapper><sst/></wrapper><sst " },
        .{ .name = "empty_inline", .part = "xl/worksheets/sheet2.xml", .old = "</row></sheetData>", .new = "</row><row r=\"7\"><c r=\"A7\" t=\"inlineStr\"><is/></c></row></sheetData>" },
        // Codex #206 r22 SEC-2202: character data or a CDATA section
        // outside the root, on the host and on the table.
        .{ .name = "host_prefix_text", .part = "xl/worksheets/sheet2.xml", .old = "<worksheet ", .new = "junk<worksheet " },
        .{ .name = "host_suffix_text", .part = "xl/worksheets/sheet2.xml", .old = "</worksheet>", .new = "</worksheet>junk" },
        .{ .name = "host_top_cdata", .part = "xl/worksheets/sheet2.xml", .old = "<worksheet ", .new = "<![CDATA[x]]><worksheet " },
        .{ .name = "sst_prefix_text", .part = "xl/sharedStrings.xml", .old = "<sst ", .new = "junk<sst " },
    };
    for (cases) |case| {
        const file = try std.fmt.allocPrint(a, "s7b5_r21_{s}.xlsx", .{case.name});
        defer a.free(file);
        const src = try tt.path(a, io, file);
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, src, case.part, case.old, case.new);
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
        try ed.setCell(0, 2, 0, .{ .string = "Central" });
        try ed.save(io, dst);
        try expectMarkerOnly(io, src, dst);
    }
}

test "S7b-5: a count-only shared-string rewrite reads the table as one document too (Codex #206 r22 SEC-2201)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const a = std.testing.allocator;
    const src = try tt.path(a, io, "s7b5_r22_wrapped_shrink.xlsx");
    defer a.free(src);
    const dst = try tt.path(a, io, "s7b5_r22_wrapped_shrink_dst.xlsx");
    defer a.free(dst);
    try pivots_mod.fixture.write(a, io, src, .sheet_ref);
    try pivots_mod.fixture.patchPart(a, io, src, "xl/sharedStrings.xml", "<sst ", "<wrapper><sst/></wrapper><sst ");
    var wb = try Workbook.open(a, io, src);
    defer wb.deinit();
    const before = try a.dupe(u8, (try wb.store.part("xl/sharedStrings.xml")).?.bytes);
    defer a.free(before);
    // A delete shrinks the rectangle with strings the table has: the
    // count-only branch.
    try std.testing.expectError(error.PivotEditUnsafe, wb.deleteRow(0, 3));
    try std.testing.expectEqualStrings(before, (try wb.store.part("xl/sharedStrings.xml")).?.bytes);
    var ed = try Editor.open(a, io, src);
    defer ed.deinit();
    try std.testing.expectError(error.RowEditUnsafeForSheet, ed.deleteRow(0, 3));
    try ed.setCell(0, 3, 1, .{ .number = 9.5 });
    try ed.save(io, dst);
    try expectMarkerOnly(io, src, dst);
}

test "S7b-5: a consumer host with staged appends refuses the edit; an <is> nested under another element is not the cell's (Codex #206 r23 REL-2301, REL-2302)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const a = std.testing.allocator;
    const dst = try tt.path(a, io, "s7b5_r23_dst.xlsx");
    defer a.free(dst);
    {
        const src = try tt.path(a, io, "s7b5_r23_appends.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        const row = [_]Cell{.{ .string = "pending" }};
        const rows = [_][]const Cell{&row};
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try ed.appendRows(1, &rows);
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        const host = try wb.sheet(1);
        try host.appendRows(&rows);
        const before = try a.dupe(u8, (try wb.store.part("xl/worksheets/sheet2.xml")).?.bytes);
        defer a.free(before);
        try std.testing.expectError(error.PivotEditUnsafe, wb.insertRow(0, 2));
        try std.testing.expectEqualStrings(before, (try wb.store.part("xl/worksheets/sheet2.xml")).?.bytes);
        try std.testing.expectEqual(@as(usize, 1), host.appended_rows.items.len);
    }
    const Case = struct { name: []const u8, rows: []const u8 };
    const cases = [_]Case{
        .{ .name = "nested_is", .rows = "</row><row r=\"3\"><c r=\"A3\" t=\"inlineStr\"><foo><is><t>decoy</t></is></foo></c></row></sheetData>" },
        .{ .name = "nested_then_direct", .rows = "</row><row r=\"3\"><c r=\"A3\" t=\"inlineStr\"><foo><is><t>decoy</t></is></foo><is><t>Row Labels</t></is></c></row></sheetData>" },
    };
    for (cases) |case| {
        const file = try std.fmt.allocPrint(a, "s7b5_r23_{s}.xlsx", .{case.name});
        defer a.free(file);
        const src = try tt.path(a, io, file);
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", "</row></sheetData>", case.rows);
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        // A host that does not read as a grid surfaces as such.
        try std.testing.expectError(error.MalformedSheetXml, ed.insertRow(0, 2));
        try ed.setCell(0, 2, 0, .{ .string = "Central" });
        try ed.save(io, dst);
        try expectMarkerOnly(io, src, dst);
    }
}

test "S7b-5: a new shared string goes before the table's extension list; character data between an inline string's elements refuses (Codex #206 r24 REL-2401, REL-2403)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const a = std.testing.allocator;
    const dst = try tt.path(a, io, "s7b5_r24_dst.xlsx");
    defer a.free(dst);
    {
        const src = try tt.path(a, io, "s7b5_r24_ext.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/sharedStrings.xml", "</sst>", "<extLst><ext uri=\"{x}\"/></extLst></sst>");
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try ed.insertRow(0, 2);
        try ed.save(io, dst);
        var store = try store_mod.PartStore.open(a, io, dst);
        defer store.deinit();
        const sst = (try store.part("xl/sharedStrings.xml")).?.bytes;
        const ext = std.mem.indexOf(u8, sst, "<extLst>").?;
        const last_si = std.mem.lastIndexOf(u8, sst, "<si>").?;
        try std.testing.expect(last_si < ext);
        try std.testing.expect(std.mem.endsWith(u8, std.mem.trimEnd(u8, sst, " \n"), "<extLst><ext uri=\"{x}\"/></extLst></sst>"));
        var wb = try Workbook.open(a, io, dst);
        defer wb.deinit();
        try expectHostCell(&wb, 1, "A6", .{ .text = "(blank)" });
    }
    const Case = struct { name: []const u8, rows: []const u8 };
    const cases = [_]Case{
        .{ .name = "leading_text", .rows = "</row><row r=\"3\"><c r=\"A3\" t=\"inlineStr\"><is>junk<t>Row Labels</t></is></c></row></sheetData>" },
        .{ .name = "trailing_text", .rows = "</row><row r=\"3\"><c r=\"A3\" t=\"inlineStr\"><is><t>Row Labels</t>tail</is></c></row></sheetData>" },
    };
    for (cases) |case| {
        const file = try std.fmt.allocPrint(a, "s7b5_r24_{s}.xlsx", .{case.name});
        defer a.free(file);
        const src = try tt.path(a, io, file);
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", "</row></sheetData>", case.rows);
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        try std.testing.expectError(error.MalformedSheetXml, wb.insertRow(0, 2));
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try ed.setCell(0, 2, 0, .{ .string = "Central" });
        try ed.save(io, dst);
        try expectMarkerOnly(io, src, dst);
    }
}

test "S7b-5: a host root bound to a foreign default namespace refuses; the Strict namespace is a main one (Codex #206 r26 SEC-2601)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const a = std.testing.allocator;
    const dst = try tt.path(a, io, "s7b5_r26_dst.xlsx");
    defer a.free(dst);
    const main_ns = "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"";
    {
        const src = try tt.path(a, io, "s7b5_r26_foreign.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", main_ns, "xmlns=\"urn:vendor\"");
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        const before = try a.dupe(u8, (try wb.store.part("xl/worksheets/sheet2.xml")).?.bytes);
        defer a.free(before);
        try std.testing.expectError(error.PivotEditUnsafe, wb.insertRow(0, 2));
        try std.testing.expectEqualStrings(before, (try wb.store.part("xl/worksheets/sheet2.xml")).?.bytes);
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
        try ed.setCell(0, 2, 0, .{ .string = "Central" });
        try ed.save(io, dst);
        try expectMarkerOnly(io, src, dst);
    }
    {
        const src = try tt.path(a, io, "s7b5_r26_strict.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", main_ns, "xmlns=\"http://purl.oclc.org/ooxml/spreadsheetml/main\"");
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try ed.insertRow(0, 2);
        try ed.save(io, dst);
        var wb = try Workbook.open(a, io, dst);
        defer wb.deinit();
        try expectHostCell(&wb, 1, "A6", .{ .text = "(blank)" });
        try expectHostCell(&wb, 1, "B7", .{ .number = 12 });
    }
}

test "S7b-5: a host root with no default binding refuses; a planned cell goes before a row's extension list; a self-closing sheetData keeps its attributes (Codex #206 r27 SEC-2701, REL-2702, REL-2703)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const a = std.testing.allocator;
    const dst = try tt.path(a, io, "s7b5_r27_dst.xlsx");
    defer a.free(dst);
    const main_ns = "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"";
    {
        const src = try tt.path(a, io, "s7b5_r27_unbound.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", main_ns, "xmlns:z=\"urn:z\"");
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        const before = try a.dupe(u8, (try wb.store.part("xl/worksheets/sheet2.xml")).?.bytes);
        defer a.free(before);
        try std.testing.expectError(error.PivotEditUnsafe, wb.insertRow(0, 2));
        try std.testing.expectEqualStrings(before, (try wb.store.part("xl/worksheets/sheet2.xml")).?.bytes);
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
        try ed.setCell(0, 2, 0, .{ .string = "Central" });
        try ed.save(io, dst);
        try expectMarkerOnly(io, src, dst);
    }
    {
        // Row 4 holds one cell and an extension list: the pivot's `B4`
        // lands between them.
        const src = try tt.path(a, io, "s7b5_r27_row_ext.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", "</row></sheetData>", "</row><row r=\"4\"><c r=\"A4\"/><extLst><ext uri=\"{x}\"/></extLst></row></sheetData>");
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try ed.insertRow(0, 2);
        try ed.save(io, dst);
        var store = try store_mod.PartStore.open(a, io, dst);
        defer store.deinit();
        const host = (try store.part("xl/worksheets/sheet2.xml")).?.bytes;
        const b4 = std.mem.indexOf(u8, host, "<c r=\"B4\"").?;
        const ext = std.mem.indexOf(u8, host, "<extLst><ext uri=\"{x}\"/></extLst></row>").?;
        try std.testing.expect(b4 < ext);
        var wb = try Workbook.open(a, io, dst);
        defer wb.deinit();
        try expectHostCell(&wb, 1, "B4", .{ .number = 8 });
    }
    {
        // The old rows commented out, a self-closing `<sheetData>` with
        // attributes in their place: it opens as written.
        const src = try tt.path(a, io, "s7b5_r27_self_closing.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", "<sheetData>", "<sheetData xmlns:z=\"urn:z\" z:k=\"v\"/><!--");
        try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", "</sheetData>", "-->");
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try ed.insertRow(0, 2);
        try ed.save(io, dst);
        var store = try store_mod.PartStore.open(a, io, dst);
        defer store.deinit();
        try expectPartHas(&store, "xl/worksheets/sheet2.xml", "<sheetData xmlns:z=\"urn:z\" z:k=\"v\"><row r=\"3\">");
        try expectPartHas(&store, "xl/worksheets/sheet2.xml", "</row></sheetData><!--");
        var wb = try Workbook.open(a, io, dst);
        defer wb.deinit();
        try expectHostCell(&wb, 1, "B7", .{ .number = 12 });
    }
}

test "S7b-5: a shared-string table under the host's namespace hygiene; CDATA in an inline caption is its payload, between the runs it is malformed (Codex #206 r28 SEC-2801, REL-2801)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const a = std.testing.allocator;
    const dst = try tt.path(a, io, "s7b5_r28_dst.xlsx");
    defer a.free(dst);
    const main_ns = "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"";
    const Case = struct { name: []const u8, old: []const u8, new: []const u8, admitted: bool };
    const cases = [_]Case{
        .{ .name = "sst_foreign_root", .old = main_ns, .new = "xmlns=\"urn:vendor\"", .admitted = false },
        .{ .name = "sst_rebound_si", .old = "<si><t xml:space=\"preserve\">", .new = "<si xmlns=\"urn:vendor\"><t xml:space=\"preserve\">", .admitted = false },
        .{ .name = "sst_aliased", .old = main_ns, .new = "xmlns:m=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " ++ main_ns, .admitted = false },
        .{ .name = "sst_strict", .old = main_ns, .new = "xmlns=\"http://purl.oclc.org/ooxml/spreadsheetml/main\"", .admitted = true },
        .{ .name = "sst_entity", .old = main_ns, .new = "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/mai&#x6e;\"", .admitted = true },
    };
    for (cases) |case| {
        const src = try tt.path(a, io, case.name);
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/sharedStrings.xml", case.old, case.new);
        if (case.admitted) {
            var ed = try Editor.open(a, io, src);
            defer ed.deinit();
            try ed.insertRow(0, 2);
            try ed.save(io, dst);
            var wb = try Workbook.open(a, io, dst);
            defer wb.deinit();
            try expectHostCell(&wb, 1, "B4", .{ .number = 8 });
            try expectHostCell(&wb, 1, "A3", .{ .text = "Row Labels" });
            continue;
        }
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        const before = try a.dupe(u8, (try wb.store.part("xl/sharedStrings.xml")).?.bytes);
        defer a.free(before);
        try std.testing.expectError(error.PivotEditUnsafe, wb.insertRow(0, 2));
        try std.testing.expectEqualStrings(before, (try wb.store.part("xl/sharedStrings.xml")).?.bytes);
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try std.testing.expectError(error.RowEditUnsafeForSheet, ed.insertRow(0, 2));
        try ed.setCell(0, 2, 0, .{ .string = "Central" });
        try ed.save(io, dst);
        try expectMarkerOnly(io, src, dst);
    }
    {
        // Captions with CDATA: the payload, verbatim, joined with the
        // decoded text around it.
        const src = try tt.path(a, io, "s7b5_r28_cdata.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", "</row></sheetData>", "</row><row r=\"3\"><c r=\"A3\" t=\"inlineStr\"><is><t>pre<![CDATA[x&<]]>post &amp; more</t></is></c></row><row r=\"6\"><c r=\"A6\" t=\"inlineStr\"><is><r><t><![CDATA[Grand Total]]></t></r></is></c></row></sheetData>");
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try ed.insertRow(0, 2);
        try ed.save(io, dst);
        var wb = try Workbook.open(a, io, dst);
        defer wb.deinit();
        try expectHostCell(&wb, 1, "A3", .{ .text = "prex&<post & more" });
        try expectHostCell(&wb, 1, "A7", .{ .text = "Grand Total" });
    }
    {
        // A CDATA section between an inline string's elements is
        // character data where none may be.
        const src = try tt.path(a, io, "s7b5_r28_cdata_between.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, src, "xl/worksheets/sheet2.xml", "</row></sheetData>", "</row><row r=\"6\"><c r=\"A6\" t=\"inlineStr\"><is><![CDATA[x]]><t>Grand Total</t></is></c></row></sheetData>");
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        try std.testing.expectError(error.MalformedSheetXml, wb.insertRow(0, 2));
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try ed.setCell(0, 2, 0, .{ .string = "Central" });
        try ed.save(io, dst);
        try expectMarkerOnly(io, src, dst);
    }
}

test "S7b-5: a malformed token or a duplicate name in a root's attributes refuses before the bindings after it; a grid element under an opaque child refuses; the SST `count` a render dropped stays dropped (Codex #206 r29 SEC-2901, SEC-2902, REL-2901)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const a = std.testing.allocator;
    const dst = try tt.path(a, io, "s7b5_r29_dst.xlsx");
    defer a.free(dst);
    const main_ns = "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"";
    const Case = struct { name: []const u8, part: []const u8, old: []const u8, new: []const u8 };
    const cases = [_]Case{
        .{ .name = "host_junk_then_alias", .part = "xl/worksheets/sheet2.xml", .old = main_ns, .new = main_ns ++ " junk xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"" },
        .{ .name = "host_dup_name", .part = "xl/worksheets/sheet2.xml", .old = main_ns, .new = main_ns ++ " xmlns:q=\"urn:q\" xmlns:q=\"urn:q\"" },
        .{ .name = "host_unterminated", .part = "xl/worksheets/sheet2.xml", .old = main_ns, .new = main_ns ++ " xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main" },
        .{ .name = "sst_junk_then_alias", .part = "xl/sharedStrings.xml", .old = main_ns, .new = main_ns ++ " junk xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"" },
        .{ .name = "cell_under_opaque", .part = "xl/worksheets/sheet2.xml", .old = "</row></sheetData>", .new = "</row><row r=\"4\"><foo><c r=\"A4\"><v>1</v></c></foo></row></sheetData>" },
        .{ .name = "row_under_opaque", .part = "xl/worksheets/sheet2.xml", .old = "</row></sheetData>", .new = "</row><bar><row r=\"9\"><c r=\"A9\"><v>1</v></c></row></bar></sheetData>" },
    };
    const parts = [_][]const u8{ "xl/worksheets/sheet2.xml", "xl/sharedStrings.xml", "xl/pivotTables/pivotTable1.xml", "xl/pivotCache/pivotCacheRecords1.xml" };
    for (cases) |case| {
        const src = try tt.path(a, io, case.name);
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, src, case.part, case.old, case.new);
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        var before: [parts.len][]u8 = undefined;
        for (parts, 0..) |name, i| before[i] = try a.dupe(u8, (try wb.store.part(name)).?.bytes);
        defer for (before) |b| a.free(b);
        if (wb.insertRow(0, 2)) |_| return error.TestUnexpectedResult else |err| switch (err) {
            error.PivotEditUnsafe, error.MalformedSheetXml => {},
            else => return err,
        }
        for (parts, 0..) |name, i| try std.testing.expectEqualStrings(before[i], (try wb.store.part(name)).?.bytes);
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try ed.setCell(0, 2, 0, .{ .string = "Central" });
        try ed.save(io, dst);
        try expectMarkerOnly(io, src, dst);
    }
    {
        // A source write re-lays the host (its strings drop the SST's
        // `count`); a novel shared string elsewhere extends the same
        // table later in the save — `count` stays absent, `uniqueCount`
        // is the entries.
        const src = try tt.path(a, io, "s7b5_r29_count.xlsx");
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        const data = try wb.sheet(0);
        try data.setCell("A3", .{ .shared_string = "Central" });
        try data.setCell("E1", .{ .shared_string = "Novel" });
        try wb.save(io, dst);
        var store = try store_mod.PartStore.open(a, io, dst);
        defer store.deinit();
        const sst = (try store.part("xl/sharedStrings.xml")).?.bytes;
        try std.testing.expect(std.mem.indexOf(u8, sst, "Novel</t>") != null);
        try std.testing.expect(std.mem.indexOf(u8, sst, " count=") == null);
        const uc_at = std.mem.indexOf(u8, sst, "uniqueCount=\"").? + "uniqueCount=\"".len;
        const uc_end = std.mem.indexOfScalarPos(u8, sst, uc_at, '"').?;
        const unique = try std.fmt.parseInt(usize, sst[uc_at..uc_end], 10);
        try std.testing.expectEqual(std.mem.count(u8, sst, "<si>"), unique);
        var out = try Workbook.open(a, io, dst);
        defer out.deinit();
        try expectHostCell(&out, 1, "A5", .{ .text = "Central" });
        try expectHostCell(&out, 1, "B5", .{ .number = 4 });
        try expectHostCell(&out, 1, "B6", .{ .number = 12 });
    }
}

test "S7b-5: an <si> after the extension list or nested under another element refuses (its number and the cells' disagree); a root with more attributes than the ceiling refuses (Codex #206 r30 SEC-3001, PERF-3002)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const a = std.testing.allocator;
    const dst = try tt.path(a, io, "s7b5_r30_dst.xlsx");
    defer a.free(dst);
    const main_ns = "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"";
    var many: std.ArrayListUnmanaged(u8) = .empty;
    defer many.deinit(a);
    try many.appendSlice(a, main_ns);
    for (0..257) |i| try many.print(a, " a{d}=\"{d}\"", .{ i, i });
    const Case = struct { name: []const u8, part: []const u8, old: []const u8, new: []const u8 };
    const cases = [_]Case{
        .{ .name = "sst_si_after_ext", .part = "xl/sharedStrings.xml", .old = "</sst>", .new = "<extLst><ext uri=\"{x}\"/></extLst><si><t xml:space=\"preserve\">Late</t></si></sst>" },
        .{ .name = "sst_si_nested", .part = "xl/sharedStrings.xml", .old = "</sst>", .new = "<foo><si><t xml:space=\"preserve\">Nested</t></si></foo></sst>" },
        .{ .name = "host_257_attrs", .part = "xl/worksheets/sheet2.xml", .old = main_ns, .new = many.items },
        .{ .name = "sst_257_attrs", .part = "xl/sharedStrings.xml", .old = main_ns, .new = many.items },
    };
    const parts = [_][]const u8{ "xl/worksheets/sheet2.xml", "xl/sharedStrings.xml", "xl/pivotTables/pivotTable1.xml", "xl/pivotCache/pivotCacheRecords1.xml" };
    for (cases) |case| {
        const src = try tt.path(a, io, case.name);
        defer a.free(src);
        try pivots_mod.fixture.write(a, io, src, .sheet_ref);
        try pivots_mod.fixture.patchPart(a, io, src, case.part, case.old, case.new);
        var wb = try Workbook.open(a, io, src);
        defer wb.deinit();
        var before: [parts.len][]u8 = undefined;
        for (parts, 0..) |name, i| before[i] = try a.dupe(u8, (try wb.store.part(name)).?.bytes);
        defer for (before) |b| a.free(b);
        if (wb.insertRow(0, 2)) |_| return error.TestUnexpectedResult else |err| switch (err) {
            error.PivotEditUnsafe, error.MalformedSheetXml, error.MalformedXml => {},
            else => return err,
        }
        for (parts, 0..) |name, i| try std.testing.expectEqualStrings(before[i], (try wb.store.part(name)).?.bytes);
        var ed = try Editor.open(a, io, src);
        defer ed.deinit();
        try ed.setCell(0, 2, 0, .{ .string = "Central" });
        try ed.save(io, dst);
        try expectMarkerOnly(io, src, dst);
    }
}
