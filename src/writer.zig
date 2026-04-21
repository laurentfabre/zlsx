//! xlsx writer — fresh-file emission (Phase 3a MVP).
//!
//! Scope
//! -----
//! * Single-call `create → addSheet → writeRow → save` flow.
//! * Multiple sheets.
//! * Cell types: empty, string (shared + deduped), integer, number, boolean.
//! * Output: OOXML with zip-store (no deflate on write). Excel, LibreOffice,
//!   and zlsx's own reader all accept stored zip xlsx.
//!
//! Out of scope (later phases)
//! ---------------------------
//! * Styles, fonts, fills, borders — Phase 3b (openpyxl-parity).
//! * Load + edit + save round-trip — Phase 3c.
//! * Formulas, merged regions, rich text, inline strings on write.

const std = @import("std");
const xlsx = @import("xlsx.zig");

const Allocator = std.mem.Allocator;

/// Returns true iff `n` can be represented exactly as an IEEE-754 double
/// (which is how spreadsheets store numeric cells). Integers with more
/// than 53 significant bits (after stripping trailing zeros) round on
/// open; those are rejected up front by `writeRow`.
///
/// Notes:
/// * `2^53` fits (one significant bit after stripping trailing zeros).
/// * `2^53 + 1` does not (54 significant bits).
/// * `2^54`, `3 * 2^52`, `2^62`, etc. all fit — magnitude is irrelevant,
///   only the count of bits after stripping trailing zeros matters.
fn fitsExactlyInF64(n: i64) bool {
    if (n == 0) return true;
    // Take absolute value as u64 so std.math.minInt(i64) = -2^63 is
    // representable (it flips to 2^63 which fits in u64 unchanged).
    const abs_n: u64 = if (n < 0) @as(u64, @intCast(-(n + 1))) + 1 else @intCast(n);
    const trailing = @ctz(abs_n);
    const shifted = abs_n >> @intCast(trailing);
    const bit_len = 64 - @clz(shifted);
    return bit_len <= 53;
}

// ─── OOXML skeleton strings ──────────────────────────────────────────

const CONTENT_TYPES_HEAD: []const u8 =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    \\<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    \\<Default Extension="xml" ContentType="application/xml"/>
    \\<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    \\<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
;
const CONTENT_TYPES_TAIL: []const u8 = "</Types>";

const ROOT_RELS: []const u8 =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    \\<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
    \\</Relationships>
;

const WORKBOOK_HEAD: []const u8 =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>
;
const WORKBOOK_TAIL: []const u8 = "</sheets></workbook>";

const WORKBOOK_RELS_HEAD: []const u8 =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
;
const WORKBOOK_RELS_TAIL: []const u8 = "</Relationships>";

const WORKSHEET_HEAD: []const u8 =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
;
const WORKSHEET_TAIL: []const u8 = "</sheetData></worksheet>";

const SST_HEAD_FMT: []const u8 =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{d}" uniqueCount="{d}">
;
const SST_TAIL: []const u8 = "</sst>";

// Static skeleton for xl/styles.xml. Fixed sections (fills at index 0=none
// and 1=gray125, empty border, default cellStyleXfs entry, "Normal"
// cellStyle) — fonts and cellXfs get appended dynamically.
const STYLES_HEAD: []const u8 =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
;
const STYLES_FONTS_DEFAULT: []const u8 =
    \\<font><sz val="11"/><name val="Calibri"/></font>
;
const STYLES_FILLS: []const u8 =
    \\<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
;
const STYLES_BORDERS: []const u8 =
    \\<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
;
const STYLES_CELL_STYLE_XFS: []const u8 =
    \\<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
;
const STYLES_DEFAULT_CELL_XF: []const u8 =
    \\<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
;
const STYLES_CELL_STYLES: []const u8 =
    \\<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
;
const STYLES_TAIL: []const u8 = "</styleSheet>";

// ─── Writer public API ───────────────────────────────────────────────

/// Cell style registered via `Writer.addStyle`. Start with the minimum
/// surface that openpyxl users reach for and grow from here — each new
/// field defaults to "unset" so registering `.{ .font_bold = true }`
/// produces the same styles.xml as registering `.{}` with bold.
///
/// Phase 3b stages:
///   - stage 1 (this release): font bold/italic
///   - stage 2: font name/size/color, horizontal alignment, wrap_text
///   - stage 3: fills (patternType, fg/bg rgb)
///   - stage 4: borders (left/right/top/bottom + style + color)
///   - stage 5: number formats, column widths, freeze panes, auto_filter
pub const Style = struct {
    font_bold: bool = false,
    font_italic: bool = false,
};

pub const Writer = struct {
    allocator: Allocator,
    // Accumulated sheet writers (owned).
    sheets: std.ArrayListUnmanaged(*SheetWriter) = .{},
    // Shared-string table: unique strings + lookup from content → index.
    sst_strings: std.ArrayListUnmanaged([]u8) = .{},
    sst_index: std.StringHashMapUnmanaged(u32) = .{},
    // Total number of string-typed cells written across all sheets
    // (informational — OOXML's <sst count="..."> field).
    sst_count: u64 = 0,
    // Registered styles (unique). Index 0 in the emitted <cellXfs> is the
    // default no-style entry; user styles start at 1 so the value
    // returned from `addStyle()` can be used directly as the cell's
    // `s="N"` attribute.
    styles: std.ArrayListUnmanaged(Style) = .{},

    pub fn init(allocator: Allocator) Writer {
        return .{ .allocator = allocator };
    }

    pub fn deinit(self: *Writer) void {
        for (self.sheets.items) |s| {
            s.deinit();
            self.allocator.destroy(s);
        }
        self.sheets.deinit(self.allocator);
        for (self.sst_strings.items) |s| self.allocator.free(s);
        self.sst_strings.deinit(self.allocator);
        self.sst_index.deinit(self.allocator);
        self.styles.deinit(self.allocator);
        self.* = undefined;
    }

    /// Register a cell style and return its `s="…"` index. Dedupes —
    /// registering the same `Style` twice returns the same index. The
    /// returned value is 1-based (cellXfs[0] is the default no-style
    /// record, reserved).
    pub fn addStyle(self: *Writer, style: Style) !u32 {
        for (self.styles.items, 0..) |existing, i| {
            if (std.meta.eql(existing, style)) return @intCast(i + 1);
        }
        try self.styles.append(self.allocator, style);
        return @intCast(self.styles.items.len);
    }

    /// Add a sheet and return a handle to append rows. Sheet is owned by
    /// the Writer — do not free the returned pointer.
    pub fn addSheet(self: *Writer, name: []const u8) !*SheetWriter {
        const sw = try self.allocator.create(SheetWriter);
        errdefer self.allocator.destroy(sw);
        sw.* = try SheetWriter.init(self, name);
        try self.sheets.append(self.allocator, sw);
        return sw;
    }

    /// Return the 0-based SST index for `s`. Dedups; copies the string
    /// into the writer's pool on first sight so callers don't need to
    /// keep it alive.
    fn sstIntern(self: *Writer, s: []const u8) !u32 {
        if (self.sst_index.get(s)) |idx| return idx;
        const owned = try self.allocator.dupe(u8, s);
        errdefer self.allocator.free(owned);
        const idx: u32 = @intCast(self.sst_strings.items.len);
        try self.sst_strings.append(self.allocator, owned);
        try self.sst_index.put(self.allocator, owned, idx);
        return idx;
    }

    /// Serialise everything and write to `path`. Overwrites.
    pub fn save(self: *Writer, path: []const u8) !void {
        if (self.sheets.items.len == 0) return error.NoSheets;

        var zip_buf: std.ArrayListUnmanaged(u8) = .{};
        defer zip_buf.deinit(self.allocator);

        var zw = ZipWriter.init(self.allocator, &zip_buf);
        defer zw.deinit();

        const alloc = self.allocator;

        const have_styles = self.styles.items.len > 0;

        // 1. [Content_Types].xml
        {
            var ct: std.ArrayListUnmanaged(u8) = .{};
            defer ct.deinit(alloc);
            try ct.appendSlice(alloc, CONTENT_TYPES_HEAD);
            for (self.sheets.items, 0..) |_, i| {
                try ct.print(
                    alloc,
                    "<Override PartName=\"/xl/worksheets/sheet{d}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>",
                    .{i + 1},
                );
            }
            if (have_styles) {
                try ct.appendSlice(
                    alloc,
                    "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>",
                );
            }
            try ct.appendSlice(alloc, CONTENT_TYPES_TAIL);
            try zw.addEntry("[Content_Types].xml", ct.items);
        }

        // 2. _rels/.rels (static)
        try zw.addEntry("_rels/.rels", ROOT_RELS);

        // 3. xl/workbook.xml
        {
            var wb: std.ArrayListUnmanaged(u8) = .{};
            defer wb.deinit(alloc);
            try wb.appendSlice(alloc, WORKBOOK_HEAD);
            for (self.sheets.items, 0..) |sw, i| {
                // Sheet names can contain XML-special chars (e.g. "R&D",
                // "x<y"); escape them before inlining into the attribute.
                try wb.appendSlice(alloc, "<sheet name=\"");
                try appendXmlEscaped(alloc, &wb, sw.name);
                try wb.print(alloc, "\" sheetId=\"{d}\" r:id=\"rId{d}\"/>", .{ i + 1, i + 1 });
            }
            try wb.appendSlice(alloc, WORKBOOK_TAIL);
            try zw.addEntry("xl/workbook.xml", wb.items);
        }

        // 4. xl/_rels/workbook.xml.rels
        {
            var rels: std.ArrayListUnmanaged(u8) = .{};
            defer rels.deinit(alloc);
            try rels.appendSlice(alloc, WORKBOOK_RELS_HEAD);
            for (self.sheets.items, 0..) |_, i| {
                try rels.print(
                    alloc,
                    "<Relationship Id=\"rId{d}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{d}.xml\"/>",
                    .{ i + 1, i + 1 },
                );
            }
            // Shared strings relationship id follows after sheets.
            try rels.print(
                alloc,
                "<Relationship Id=\"rId{d}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>",
                .{self.sheets.items.len + 1},
            );
            if (have_styles) {
                try rels.print(
                    alloc,
                    "<Relationship Id=\"rId{d}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>",
                    .{self.sheets.items.len + 2},
                );
            }
            try rels.appendSlice(alloc, WORKBOOK_RELS_TAIL);
            try zw.addEntry("xl/_rels/workbook.xml.rels", rels.items);
        }

        // 5. xl/worksheets/sheetN.xml
        for (self.sheets.items, 0..) |sw, i| {
            var full: std.ArrayListUnmanaged(u8) = .{};
            defer full.deinit(alloc);
            try full.appendSlice(alloc, WORKSHEET_HEAD);
            try full.appendSlice(alloc, sw.body.items);
            try full.appendSlice(alloc, WORKSHEET_TAIL);

            var name_buf: [64]u8 = undefined;
            const entry_name = try std.fmt.bufPrint(&name_buf, "xl/worksheets/sheet{d}.xml", .{i + 1});
            try zw.addEntry(entry_name, full.items);
        }

        // 6. xl/sharedStrings.xml
        {
            var sst: std.ArrayListUnmanaged(u8) = .{};
            defer sst.deinit(alloc);
            try sst.print(alloc, SST_HEAD_FMT, .{ self.sst_count, self.sst_strings.items.len });
            for (self.sst_strings.items) |s| {
                try sst.appendSlice(alloc, "<si><t xml:space=\"preserve\">");
                try appendXmlEscaped(alloc, &sst, s);
                try sst.appendSlice(alloc, "</t></si>");
            }
            try sst.appendSlice(alloc, SST_TAIL);
            try zw.addEntry("xl/sharedStrings.xml", sst.items);
        }

        // 7. xl/styles.xml — only when the caller registered any styles.
        // Keeps the "no styles" path byte-identical to v0.2.0-0.2.3 output.
        if (have_styles) try emitStylesXml(alloc, &zw, self.styles.items);

        try zw.finalize();

        var file = try std.fs.cwd().createFile(path, .{ .truncate = true });
        defer file.close();
        try file.writeAll(zip_buf.items);
    }
};

// ─── SheetWriter ─────────────────────────────────────────────────────

pub const SheetWriter = struct {
    parent: *Writer,
    // Owned copy of the sheet name.
    name: []u8,
    // Accumulated `<row>` elements; emitted inside <sheetData> on save.
    body: std.ArrayListUnmanaged(u8) = .{},
    // 1-based row index (xlsx convention).
    next_row: u32 = 1,

    fn init(parent: *Writer, name: []const u8) !SheetWriter {
        return .{
            .parent = parent,
            .name = try parent.allocator.dupe(u8, name),
        };
    }

    fn deinit(self: *SheetWriter) void {
        self.parent.allocator.free(self.name);
        self.body.deinit(self.parent.allocator);
        self.* = undefined;
    }

    /// Write a row of cells. Empty cells are omitted from the output
    /// (OOXML treats missing cells as empty). Strings are interned into
    /// the parent's SST.
    pub fn writeRow(self: *SheetWriter, cells: []const xlsx.Cell) !void {
        return self.writeRowImpl(cells, null);
    }

    /// Write a row with per-cell style indices. `styles.len` must equal
    /// `cells.len`; use `0` (the default no-style slot) for cells that
    /// should inherit the default formatting. Style indices come from
    /// `Writer.addStyle` / `zlsx_writer_add_style`.
    pub fn writeRowStyled(
        self: *SheetWriter,
        cells: []const xlsx.Cell,
        styles: []const u32,
    ) !void {
        if (styles.len != cells.len) return error.StyleCountMismatch;
        return self.writeRowImpl(cells, styles);
    }

    fn writeRowImpl(
        self: *SheetWriter,
        cells: []const xlsx.Cell,
        styles: ?[]const u32,
    ) !void {
        // Pre-validate integers BEFORE mutating `self.body`. This keeps
        // writeRow atomic on IntegerExceedsExcelPrecision so the caller
        // can catch the error and retry / skip that row without ending
        // up with a half-emitted <row> in the sheet body.
        for (cells) |cell| switch (cell) {
            .integer => |n| if (!fitsExactlyInF64(n)) return error.IntegerExceedsExcelPrecision,
            else => {},
        };

        const alloc = self.parent.allocator;
        try self.body.print(alloc, "<row r=\"{d}\">", .{self.next_row});

        for (cells, 0..) |cell, col_idx| {
            const style_id: u32 = if (styles) |s| s[col_idx] else 0;
            // `<c>` elements for empty cells are only emitted when a
            // non-default style is applied — otherwise OOXML's
            // "missing cell = empty" rule keeps the sheet smaller.
            if (cell == .empty and style_id == 0) continue;

            var ref_buf: [16]u8 = undefined;
            const ref = try formatCellRef(&ref_buf, self.next_row, @intCast(col_idx));

            switch (cell) {
                .empty => {
                    // Styled-but-empty cell: emit just `<c r="…" s="N"/>`.
                    try self.body.print(alloc, "<c r=\"{s}\" s=\"{d}\"/>", .{ ref, style_id });
                },
                .string => |s| {
                    const idx = try self.parent.sstIntern(s);
                    self.parent.sst_count += 1;
                    if (style_id == 0) {
                        try self.body.print(alloc, "<c r=\"{s}\" t=\"s\"><v>{d}</v></c>", .{ ref, idx });
                    } else {
                        try self.body.print(alloc, "<c r=\"{s}\" s=\"{d}\" t=\"s\"><v>{d}</v></c>", .{ ref, style_id, idx });
                    }
                },
                .integer => |n| {
                    if (style_id == 0) {
                        try self.body.print(alloc, "<c r=\"{s}\"><v>{d}</v></c>", .{ ref, n });
                    } else {
                        try self.body.print(alloc, "<c r=\"{s}\" s=\"{d}\"><v>{d}</v></c>", .{ ref, style_id, n });
                    }
                },
                .number => |f| {
                    // {d} renders the shortest round-trip decimal; Excel
                    // accepts decimal or scientific notation in <v>.
                    if (style_id == 0) {
                        try self.body.print(alloc, "<c r=\"{s}\"><v>{d}</v></c>", .{ ref, f });
                    } else {
                        try self.body.print(alloc, "<c r=\"{s}\" s=\"{d}\"><v>{d}</v></c>", .{ ref, style_id, f });
                    }
                },
                .boolean => |b| {
                    if (style_id == 0) {
                        try self.body.print(alloc, "<c r=\"{s}\" t=\"b\"><v>{d}</v></c>", .{ ref, @intFromBool(b) });
                    } else {
                        try self.body.print(alloc, "<c r=\"{s}\" s=\"{d}\" t=\"b\"><v>{d}</v></c>", .{ ref, style_id, @intFromBool(b) });
                    }
                },
            }
        }

        try self.body.appendSlice(alloc, "</row>");
        self.next_row += 1;
    }
};

// ─── Helpers ─────────────────────────────────────────────────────────

fn formatCellRef(buf: *[16]u8, row: u32, col_idx: u32) ![]u8 {
    // Column letter (1-based in xlsx: A=1, Z=26, AA=27 …).
    var col_chars: [8]u8 = undefined;
    var pos: usize = col_chars.len;
    var c = col_idx + 1;
    while (c > 0) {
        c -= 1;
        pos -= 1;
        col_chars[pos] = 'A' + @as(u8, @intCast(c % 26));
        c /= 26;
    }
    const letters = col_chars[pos..];
    return std.fmt.bufPrint(buf, "{s}{d}", .{ letters, row });
}

/// Emit xl/styles.xml based on the registered style list. Fonts are
/// deduped into a `<fonts>` list (default font at index 0; user styles
/// each reference a new font entry, which is wasteful compared to
/// deduping fonts separately from styles but is fine for the MVP scope
/// of Phase 3b stage 1). `<cellXfs>` gets the default entry at index 0
/// plus one entry per user style.
fn emitStylesXml(
    alloc: Allocator,
    zw: *ZipWriter,
    styles: []const Style,
) !void {
    var buf: std.ArrayListUnmanaged(u8) = .{};
    defer buf.deinit(alloc);

    try buf.appendSlice(alloc, STYLES_HEAD);

    // <fonts>: default at index 0 + one per user style.
    try buf.print(alloc, "<fonts count=\"{d}\">", .{styles.len + 1});
    try buf.appendSlice(alloc, STYLES_FONTS_DEFAULT);
    for (styles) |s| {
        try buf.appendSlice(alloc, "<font>");
        if (s.font_bold) try buf.appendSlice(alloc, "<b/>");
        if (s.font_italic) try buf.appendSlice(alloc, "<i/>");
        try buf.appendSlice(alloc, "<sz val=\"11\"/><name val=\"Calibri\"/></font>");
    }
    try buf.appendSlice(alloc, "</fonts>");

    try buf.appendSlice(alloc, STYLES_FILLS);
    try buf.appendSlice(alloc, STYLES_BORDERS);
    try buf.appendSlice(alloc, STYLES_CELL_STYLE_XFS);

    // <cellXfs>: default at index 0 + one per user style. Style i (1-based
    // to callers) references font[i] so each user style gets its own font.
    try buf.print(alloc, "<cellXfs count=\"{d}\">", .{styles.len + 1});
    try buf.appendSlice(alloc, STYLES_DEFAULT_CELL_XF);
    for (styles, 0..) |_, i| {
        try buf.print(
            alloc,
            "<xf numFmtId=\"0\" fontId=\"{d}\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>",
            .{i + 1},
        );
    }
    try buf.appendSlice(alloc, "</cellXfs>");

    try buf.appendSlice(alloc, STYLES_CELL_STYLES);
    try buf.appendSlice(alloc, STYLES_TAIL);

    try zw.addEntry("xl/styles.xml", buf.items);
}

fn appendXmlEscaped(alloc: Allocator, out: *std.ArrayListUnmanaged(u8), s: []const u8) !void {
    for (s) |ch| switch (ch) {
        '<' => try out.appendSlice(alloc, "&lt;"),
        '>' => try out.appendSlice(alloc, "&gt;"),
        '&' => try out.appendSlice(alloc, "&amp;"),
        '"' => try out.appendSlice(alloc, "&quot;"),
        '\'' => try out.appendSlice(alloc, "&apos;"),
        else => try out.append(alloc, ch),
    };
}

// ─── ZIP writer (stored, no deflate) ─────────────────────────────────

/// Minimal zip archive builder. Appends file entries to a byte buffer;
/// `finalize()` emits the central directory + end-of-central-directory
/// trailer. All entries use compression method 0 (stored). This keeps
/// the write path simple; Excel and libreoffice both accept stored xlsx.
const ZipWriter = struct {
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    // Per-entry info accumulated for the central directory.
    entries: std.ArrayListUnmanaged(EntryMeta) = .{},

    const EntryMeta = struct {
        name: []u8, // owned copy
        crc32: u32,
        size: u32,
        local_offset: u32,
    };

    fn init(alloc: Allocator, out: *std.ArrayListUnmanaged(u8)) ZipWriter {
        return .{ .allocator = alloc, .out = out };
    }

    fn deinit(self: *ZipWriter) void {
        for (self.entries.items) |e| self.allocator.free(e.name);
        self.entries.deinit(self.allocator);
    }

    fn addEntry(self: *ZipWriter, name: []const u8, data: []const u8) !void {
        const alloc = self.allocator;
        if (data.len > std.math.maxInt(u32)) return error.EntryTooLarge;
        if (name.len > std.math.maxInt(u16)) return error.NameTooLong;

        const crc = std.hash.Crc32.hash(data);
        const offset: u32 = @intCast(self.out.items.len);

        // Local file header.
        const hdr: std.zip.LocalFileHeader = .{
            .signature = std.zip.local_file_header_sig,
            .version_needed_to_extract = 20,
            .flags = .{ .encrypted = false, ._ = 0 },
            .compression_method = .store,
            .last_modification_time = 0,
            .last_modification_date = 0x21, // 1980-01-01, minimum valid
            .crc32 = crc,
            .compressed_size = @intCast(data.len),
            .uncompressed_size = @intCast(data.len),
            .filename_len = @intCast(name.len),
            .extra_len = 0,
        };
        try appendStruct(alloc, self.out, std.zip.LocalFileHeader, hdr);
        try self.out.appendSlice(alloc, name);
        try self.out.appendSlice(alloc, data);

        const owned_name = try alloc.dupe(u8, name);
        errdefer alloc.free(owned_name);
        try self.entries.append(alloc, .{
            .name = owned_name,
            .crc32 = crc,
            .size = @intCast(data.len),
            .local_offset = offset,
        });
    }

    fn finalize(self: *ZipWriter) !void {
        const alloc = self.allocator;
        const cd_start: u32 = @intCast(self.out.items.len);

        for (self.entries.items) |e| {
            const cd: std.zip.CentralDirectoryFileHeader = .{
                .signature = std.zip.central_file_header_sig,
                .version_made_by = 20,
                .version_needed_to_extract = 20,
                .flags = .{ .encrypted = false, ._ = 0 },
                .compression_method = .store,
                .last_modification_time = 0,
                .last_modification_date = 0x21,
                .crc32 = e.crc32,
                .compressed_size = e.size,
                .uncompressed_size = e.size,
                .filename_len = @intCast(e.name.len),
                .extra_len = 0,
                .comment_len = 0,
                .disk_number = 0,
                .internal_file_attributes = 0,
                .external_file_attributes = 0,
                .local_file_header_offset = e.local_offset,
            };
            try appendStruct(alloc, self.out, std.zip.CentralDirectoryFileHeader, cd);
            try self.out.appendSlice(alloc, e.name);
        }

        const cd_end: u32 = @intCast(self.out.items.len);
        const cd_size = cd_end - cd_start;

        const end: std.zip.EndRecord = .{
            .signature = std.zip.end_record_sig,
            .disk_number = 0,
            .central_directory_disk_number = 0,
            .record_count_disk = @intCast(self.entries.items.len),
            .record_count_total = @intCast(self.entries.items.len),
            .central_directory_size = cd_size,
            .central_directory_offset = cd_start,
            .comment_len = 0,
        };
        try appendStruct(alloc, self.out, std.zip.EndRecord, end);
    }
};

fn appendStruct(alloc: Allocator, out: *std.ArrayListUnmanaged(u8), comptime T: type, value: T) !void {
    const bytes = std.mem.asBytes(&value);
    try out.appendSlice(alloc, bytes);
}

// ─── Tests ───────────────────────────────────────────────────────────

test "formatCellRef: A1, B2, Z1, AA1, AAA1" {
    var buf: [16]u8 = undefined;
    try std.testing.expectEqualStrings("A1", try formatCellRef(&buf, 1, 0));
    try std.testing.expectEqualStrings("B2", try formatCellRef(&buf, 2, 1));
    try std.testing.expectEqualStrings("Z1", try formatCellRef(&buf, 1, 25));
    try std.testing.expectEqualStrings("AA1", try formatCellRef(&buf, 1, 26));
    try std.testing.expectEqualStrings("AAA1", try formatCellRef(&buf, 1, 702));
}

test "appendXmlEscaped covers all 5 entities" {
    var buf: std.ArrayListUnmanaged(u8) = .{};
    defer buf.deinit(std.testing.allocator);
    try appendXmlEscaped(std.testing.allocator, &buf, "a<b>c&d\"e'f");
    try std.testing.expectEqualStrings("a&lt;b&gt;c&amp;d&quot;e&apos;f", buf.items);
}

test "Writer: empty workbook fails with NoSheets" {
    var w = Writer.init(std.testing.allocator);
    defer w.deinit();
    try std.testing.expectError(error.NoSheets, w.save("/tmp/zlsx_empty.xlsx"));
}

test "Writer: single-sheet round-trip via zlsx reader" {
    const tmp_path = "/tmp/zlsx_writer_test.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();

        var sheet = try w.addSheet("Summary");
        try sheet.writeRow(&.{
            .{ .string = "Name" },
            .{ .string = "Age" },
            .{ .string = "Active" },
            .{ .string = "Pi" },
        });
        try sheet.writeRow(&.{
            .{ .string = "Alice" },
            .{ .integer = 30 },
            .{ .boolean = true },
            .{ .number = 3.14159 },
        });
        try sheet.writeRow(&.{
            .{ .string = "Bob" },
            .{ .integer = 25 },
            .{ .boolean = false },
            .empty,
        });

        try w.save(tmp_path);
    }

    // Read it back.
    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    try std.testing.expectEqual(@as(usize, 1), book.sheets.len);
    try std.testing.expectEqualStrings("Summary", book.sheets[0].name);

    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();

    const r1 = (try rows.next()).?;
    try std.testing.expectEqual(@as(usize, 4), r1.len);
    try std.testing.expectEqualStrings("Name", r1[0].string);
    try std.testing.expectEqualStrings("Age", r1[1].string);
    try std.testing.expectEqualStrings("Active", r1[2].string);
    try std.testing.expectEqualStrings("Pi", r1[3].string);

    const r2 = (try rows.next()).?;
    try std.testing.expectEqualStrings("Alice", r2[0].string);
    try std.testing.expectEqual(@as(i64, 30), r2[1].integer);
    try std.testing.expectEqual(true, r2[2].boolean);
    try std.testing.expectApproxEqAbs(@as(f64, 3.14159), r2[3].number, 1e-9);

    const r3 = (try rows.next()).?;
    try std.testing.expectEqualStrings("Bob", r3[0].string);
    try std.testing.expectEqual(@as(i64, 25), r3[1].integer);
    try std.testing.expectEqual(false, r3[2].boolean);
    // r3[3] may be .empty or may be absent depending on reader's row-width
    // policy; don't assert length.

    try std.testing.expectEqual(@as(?[]const xlsx.Cell, null), try rows.next());
}

test "Writer: multi-sheet round-trip + SST dedup" {
    const tmp_path = "/tmp/zlsx_writer_multisheet.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();

        var s1 = try w.addSheet("Alpha");
        try s1.writeRow(&.{.{ .string = "hello" }});
        try s1.writeRow(&.{.{ .string = "world" }});

        var s2 = try w.addSheet("Beta");
        // "hello" dedupes against s1's SST entry.
        try s2.writeRow(&.{.{ .string = "hello" }});
        try s2.writeRow(&.{.{ .string = "zig" }});

        try w.save(tmp_path);

        // 3 unique strings after dedup: hello, world, zig.
        try std.testing.expectEqual(@as(usize, 3), w.sst_strings.items.len);
        // 4 string-cell writes total.
        try std.testing.expectEqual(@as(u64, 4), w.sst_count);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 2), book.sheets.len);
    try std.testing.expectEqualStrings("Alpha", book.sheets[0].name);
    try std.testing.expectEqualStrings("Beta", book.sheets[1].name);
    try std.testing.expectEqual(@as(usize, 3), book.shared_strings.len);
}

test "Writer: xml entities in strings are escaped" {
    const tmp_path = "/tmp/zlsx_writer_entities.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("S");
        try sheet.writeRow(&.{.{ .string = "a<b & c>d \"e\" 'f'" }});
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    const r = (try rows.next()).?;
    try std.testing.expectEqualStrings("a<b & c>d \"e\" 'f'", r[0].string);
}

test "Writer: styles — bold + italic round-trip" {
    const tmp_path = "/tmp/zlsx_writer_styles_bold.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    var registered_bold: u32 = 0;
    var registered_italic: u32 = 0;

    {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();

        registered_bold = try w.addStyle(.{ .font_bold = true });
        registered_italic = try w.addStyle(.{ .font_italic = true });

        // Dedup: registering the same style again returns the same index.
        const again = try w.addStyle(.{ .font_bold = true });
        try std.testing.expectEqual(registered_bold, again);

        // Style indices are 1-based (0 is the default no-style slot).
        try std.testing.expect(registered_bold >= 1);
        try std.testing.expect(registered_italic != registered_bold);

        var s = try w.addSheet("S");
        try s.writeRowStyled(
            &.{ .{ .string = "bold" }, .{ .string = "italic" }, .{ .string = "plain" } },
            &.{ registered_bold, registered_italic, 0 },
        );
        // Unstyled path still works alongside styled rows.
        try s.writeRow(&.{.{ .string = "unstyled row" }});

        // styles.len != cells.len → error.StyleCountMismatch
        try std.testing.expectError(error.StyleCountMismatch, s.writeRowStyled(
            &.{.{ .string = "x" }},
            &.{},
        ));

        try w.save(tmp_path);
    }

    // The reader ignores styles but the file must still parse cleanly
    // and contain the cell values we wrote. Also grep the raw archive
    // for xl/styles.xml + applyFont markers so we know styles.xml was
    // actually emitted.
    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    const r1 = (try rows.next()).?;
    try std.testing.expectEqualStrings("bold", r1[0].string);
    try std.testing.expectEqualStrings("italic", r1[1].string);
    try std.testing.expectEqualStrings("plain", r1[2].string);
    const r2 = (try rows.next()).?;
    try std.testing.expectEqualStrings("unstyled row", r2[0].string);

    // Read xl/styles.xml raw out of the archive and check for the bold +
    // italic markers + applyFont attribute — proves the styles.xml
    // emission path actually ran.
    const styles_xml = blk: {
        var file = try std.fs.cwd().openFile(tmp_path, .{});
        defer file.close();
        var fbuf: [4096]u8 = undefined;
        var fr = file.reader(&fbuf);
        var iter = try std.zip.Iterator.init(&fr);
        var filename_buf: [64]u8 = undefined;
        while (try iter.next()) |entry| {
            if (entry.filename_len > filename_buf.len) continue;
            try fr.seekTo(entry.header_zip_offset + @sizeOf(std.zip.CentralDirectoryFileHeader));
            const filename = filename_buf[0..entry.filename_len];
            try fr.interface.readSliceAll(filename);
            if (std.mem.eql(u8, filename, "xl/styles.xml")) {
                break :blk try extractEntryForTest(std.testing.allocator, entry, &fr);
            }
        }
        return error.StylesXmlNotFound;
    };
    defer std.testing.allocator.free(styles_xml);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "<b/>") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "<i/>") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "applyFont=\"1\"") != null);
}

/// Test helper: mirror the reader's extractEntryToBuffer but keep it
/// local so this test file doesn't reach into xlsx.zig internals.
fn extractEntryForTest(
    allocator: Allocator,
    entry: std.zip.Iterator.Entry,
    stream: anytype,
) ![]u8 {
    try stream.seekTo(entry.file_offset);
    const local = try stream.interface.takeStruct(std.zip.LocalFileHeader, .little);
    try stream.seekTo(entry.file_offset + @sizeOf(std.zip.LocalFileHeader) + local.filename_len + local.extra_len);
    const out = try allocator.alloc(u8, entry.uncompressed_size);
    errdefer allocator.free(out);
    var w = std.Io.Writer.fixed(out);
    switch (entry.compression_method) {
        .store => try stream.interface.streamExact64(&w, entry.uncompressed_size),
        .deflate => {
            var flate_buffer: [std.compress.flate.max_window_len]u8 = undefined;
            var decompress = std.compress.flate.Decompress.init(&stream.interface, .raw, &flate_buffer);
            try decompress.reader.streamExact64(&w, entry.uncompressed_size);
        },
        else => unreachable,
    }
    return out;
}

test "Writer: sheet names with XML-special chars are escaped" {
    const tmp_path = "/tmp/zlsx_writer_sheet_escape.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();
        // Sheet names with ampersand, angles, and quote — all common
        // in real workbooks ("R&D", "x<y", 'He said "hi"').
        _ = try w.addSheet("R&D");
        _ = try w.addSheet("x<y");
        const s3 = try w.addSheet("quote\"it");
        try s3.writeRow(&.{.{ .string = "marker" }});
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 3), book.sheets.len);
    try std.testing.expectEqualStrings("R&D", book.sheets[0].name);
    try std.testing.expectEqualStrings("x<y", book.sheets[1].name);
    try std.testing.expectEqualStrings("quote\"it", book.sheets[2].name);
}

test "Writer: reject only integers that round on IEEE-754 conversion" {
    var w = Writer.init(std.testing.allocator);
    defer w.deinit();
    var sheet = try w.addSheet("S");

    // Exactly representable — must succeed.
    try sheet.writeRow(&.{.{ .integer = 1 << 53 }}); // 2^53
    try sheet.writeRow(&.{.{ .integer = 1 << 54 }}); // 2^54 — magnitude is fine
    try sheet.writeRow(&.{.{ .integer = 1 << 62 }}); // 2^62 — still fits
    try sheet.writeRow(&.{.{ .integer = 3 * (@as(i64, 1) << 52) }}); // 2 significant bits
    try sheet.writeRow(&.{.{ .integer = -(1 << 54) }}); // negative power of two
    try sheet.writeRow(&.{.{ .integer = std.math.minInt(i64) }}); // -2^63

    // NOT exactly representable — 54+ significant bits.
    try std.testing.expectError(
        error.IntegerExceedsExcelPrecision,
        sheet.writeRow(&.{.{ .integer = (1 << 53) + 1 }}),
    );
    try std.testing.expectError(
        error.IntegerExceedsExcelPrecision,
        sheet.writeRow(&.{.{ .integer = -((1 << 53) + 1) }}),
    );
    try std.testing.expectError(
        error.IntegerExceedsExcelPrecision,
        sheet.writeRow(&.{.{ .integer = std.math.maxInt(i64) }}), // 2^63 - 1
    );
}

test "Writer: writeRow is atomic on IntegerExceedsExcelPrecision" {
    const tmp_path = "/tmp/zlsx_writer_atomic.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    var w = Writer.init(std.testing.allocator);
    defer w.deinit();
    var sheet = try w.addSheet("S");

    // First row succeeds.
    try sheet.writeRow(&.{.{ .string = "ok" }});

    // Second row fails validation — the bad integer is after a good cell,
    // so a non-atomic writer would have already appended `<row>` + the
    // first `<c>` before hitting the error.
    try std.testing.expectError(
        error.IntegerExceedsExcelPrecision,
        sheet.writeRow(&.{
            .{ .string = "first" },
            .{ .integer = (1 << 53) + 1 }, // bad
            .{ .string = "third" },
        }),
    );

    // Third row succeeds and becomes row 2 (next_row wasn't advanced).
    try sheet.writeRow(&.{.{ .string = "after" }});

    try w.save(tmp_path);

    // Reading back proves the file is well-formed: no partial row leaked.
    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();

    const r1 = (try rows.next()).?;
    try std.testing.expectEqualStrings("ok", r1[0].string);
    const r2 = (try rows.next()).?;
    try std.testing.expectEqualStrings("after", r2[0].string);
    try std.testing.expectEqual(@as(?[]const xlsx.Cell, null), try rows.next());
}

test "fitsExactlyInF64 matches round-trip reference" {
    // Sanity check: fitsExactlyInF64(n) iff (f64 round-trip == n).
    const test_values = [_]i64{
        0,             1,                       -1,
        1 << 52,       (1 << 52) - 1,           (1 << 52) + 1,
        1 << 53,       (1 << 53) - 1,           (1 << 53) + 1,
        1 << 54,       3 * (@as(i64, 1) << 52), 1 << 62,
        (1 << 62) + 1, std.math.maxInt(i64),    std.math.minInt(i64),
    };
    for (test_values) |n| {
        const f: f64 = @floatFromInt(n);
        // Round-trip reference — only valid when f is in i64 range.
        const lossless_via_roundtrip = blk: {
            if (f >= 9.223372036854776e18 or f < -9.223372036854776e18) break :blk false;
            const back: i64 = @intFromFloat(f);
            break :blk back == n;
        };
        try std.testing.expectEqual(lossless_via_roundtrip, fitsExactlyInF64(n));
    }
}

test "Writer: exposed via @import(\"xlsx.zig\") namespace re-export" {
    // This ensures the re-export at the bottom of xlsx.zig actually
    // compiles — downstream consumers rely on @import("zlsx").Writer.
    const W = xlsx.Writer;
    const SW = xlsx.SheetWriter;
    comptime {
        _ = W;
        _ = SW;
    }
}
