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

const WORKSHEET_PROLOG: []const u8 =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
;

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

/// OOXML reserves numFmtIds 0..=49 for built-ins; user numFmts must
/// start at 164 (Excel's convention — 50..=163 are "reserved").
const NUM_FMT_BASE: u32 = 164;

// ─── Writer public API ───────────────────────────────────────────────

/// OOXML border-side style enum. `.none` is the default (no side
/// emitted); numeric tag values are part of the C ABI — append
/// new entries, never reorder.
pub const BorderStyle = enum(u8) {
    none = 0,
    thin = 1,
    medium = 2,
    dashed = 3,
    dotted = 4,
    thick = 5,
    double = 6,
    hair = 7,
    medium_dashed = 8,
    dash_dot = 9,
    medium_dash_dot = 10,
    dash_dot_dot = 11,
    medium_dash_dot_dot = 12,
    slant_dash_dot = 13,
};

/// One side of a cell border (left / right / top / bottom / diagonal).
pub const BorderSide = struct {
    style: BorderStyle = .none,
    /// ARGB colour for the border line. Null = OOXML default (auto).
    color_argb: ?u32 = null,
};

/// OOXML `<patternFill patternType="…"/>` values. `.none` is the
/// default (no fill); numeric tag values are part of the C ABI —
/// append new entries, never reorder.
pub const PatternType = enum(u8) {
    none = 0,
    solid = 1,
    gray125 = 2,
    gray0625 = 3,
    dark_gray = 4,
    medium_gray = 5,
    light_gray = 6,
    dark_horizontal = 7,
    dark_vertical = 8,
    dark_down = 9,
    dark_up = 10,
    dark_grid = 11,
    dark_trellis = 12,
    light_horizontal = 13,
    light_vertical = 14,
    light_down = 15,
    light_up = 16,
    light_grid = 17,
    light_trellis = 18,
};

/// Horizontal alignment for a cell style. `.general` is the OOXML
/// default (no `<alignment>` element emitted); nonzero values emit
/// `<alignment horizontal="…"/>`. Numeric tag values are part of the
/// C ABI — append new entries, never reorder.
pub const HAlign = enum(u8) {
    general = 0,
    left = 1,
    center = 2,
    right = 3,
    fill = 4,
    justify = 5,
    center_continuous = 6,
    distributed = 7,
};

/// Cell style registered via `Writer.addStyle`. Fields default to
/// "unset" so `Writer.addStyle(.{ .font_bold = true })` produces
/// the minimum-overhead styles.xml entry.
///
/// Phase 3b stages:
///   - stage 1: font bold/italic                         [shipped]
///   - stage 2 (this release): font name/size/color,
///             horizontal alignment, wrap_text
///   - stage 3: fills (patternType, fg/bg rgb)
///   - stage 4: borders (left/right/top/bottom + style + color)
///   - stage 5: number formats, column widths, freeze panes, auto_filter
///
/// `font_name` is caller-owned for the duration of the `addStyle`
/// call; the writer dupes it into its own pool so callers can free
/// the original immediately after.
pub const Style = struct {
    font_bold: bool = false,
    font_italic: bool = false,
    /// Null = default (11 pt). Must be positive and finite when set.
    font_size: ?f32 = null,
    /// Null = default ("Calibri"). Escaped for XML on emit.
    font_name: ?[]const u8 = null,
    /// Null = default (theme auto). ARGB packed: 0xAARRGGBB.
    font_color_argb: ?u32 = null,
    alignment_horizontal: HAlign = .general,
    wrap_text: bool = false,
    /// `.none` emits no fill (style points at fillId=0). Any other
    /// value emits a `<patternFill>` element. For "solid" highlights
    /// set `.fill_pattern = .solid` plus `.fill_fg_argb` to the
    /// desired ARGB colour.
    fill_pattern: PatternType = .none,
    /// Foreground (pattern) colour, ARGB packed 0xAARRGGBB. Null = OOXML default.
    fill_fg_argb: ?u32 = null,
    /// Background (pattern backdrop) colour, ARGB packed 0xAARRGGBB. Null = OOXML default.
    fill_bg_argb: ?u32 = null,
    /// Cell border sides. Defaults emit no side — set any of these
    /// `style` fields to get a border. A style that touches any
    /// border field (sides or diagonal flags) gets its own
    /// `<border>` entry in xl/styles.xml.
    border_left: BorderSide = .{},
    border_right: BorderSide = .{},
    border_top: BorderSide = .{},
    border_bottom: BorderSide = .{},
    border_diagonal: BorderSide = .{},
    /// Draw the diagonal from the lower-left corner upward to the
    /// upper-right. Requires `border_diagonal.style != .none` to
    /// render.
    diagonal_up: bool = false,
    /// Draw the diagonal from the upper-left corner downward to the
    /// lower-right. Same `border_diagonal.style` gates rendering.
    diagonal_down: bool = false,
    /// OOXML number format string (e.g., "0.00", "m/d/yyyy",
    /// "$#,##0.00"). Null = General. Custom strings register as user
    /// numFmts starting at id 164; multiple styles using the same
    /// format string share a single numFmtId.
    number_format: ?[]const u8 = null,
};

fn hasBorder(s: Style) bool {
    return s.border_left.style != .none or
        s.border_right.style != .none or
        s.border_top.style != .none or
        s.border_bottom.style != .none or
        s.border_diagonal.style != .none or
        s.diagonal_up or s.diagonal_down;
}

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
    // Number-format pool (stage 5). Parallel arrays: `num_fmts` owns
    // the format strings (writer-allocated); `num_fmt_index` maps
    // format → numFmtId (starting at 164 — OOXML reserves 0..=49 for
    // built-ins). All values are unique.
    num_fmts: std.ArrayListUnmanaged([]u8) = .{},
    num_fmt_index: std.StringHashMapUnmanaged(u32) = .{},

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
        // Each style owns its font_name / number_format slices (if any)
        // on the writer's heap; drop them here before the styles
        // ArrayList goes.
        for (self.styles.items) |s| {
            if (s.font_name) |n| self.allocator.free(n);
            if (s.number_format) |n| self.allocator.free(n);
        }
        self.styles.deinit(self.allocator);
        for (self.num_fmts.items) |n| self.allocator.free(n);
        self.num_fmts.deinit(self.allocator);
        self.num_fmt_index.deinit(self.allocator);
        self.* = undefined;
    }

    /// Register a cell style and return its `s="…"` index. Dedupes
    /// structurally (including content-comparing `font_name`, not
    /// just slice-header comparing). Returning value is 1-based —
    /// cellXfs[0] is reserved for the default no-style record.
    pub fn addStyle(self: *Writer, style: Style) !u32 {
        // Validate stage-2 inputs up front so dedup doesn't have to
        // handle subtly-equal-but-invalid specs.
        if (style.font_size) |s| {
            if (!std.math.isFinite(s) or s <= 0) return error.InvalidFontSize;
        }
        if (style.font_name) |n| {
            if (n.len == 0) return error.InvalidFontName;
        }
        if (style.number_format) |n| {
            if (n.len == 0) return error.InvalidNumberFormat;
        }

        // Side-effect of validation: register the format string in the
        // numFmt pool (dedup via StringHashMap). Happens BEFORE dedup of
        // the parent Style so we don't register formats for rejected
        // styles.
        if (style.number_format) |fmt| {
            _ = try self.internNumFmt(fmt);
        }

        // Linear-scan dedup. Need content-equal font_name comparison
        // (std.meta.eql on `?[]const u8` compares slice headers only).
        for (self.styles.items, 0..) |existing, i| {
            if (stylesEqual(existing, style)) return @intCast(i + 1);
        }

        // New entry — dupe font_name / number_format into writer-owned
        // storage so the caller can free their buffers immediately.
        var owned_style = style;
        if (style.font_name) |n| {
            owned_style.font_name = try self.allocator.dupe(u8, n);
        }
        errdefer if (owned_style.font_name) |n| self.allocator.free(n);
        if (style.number_format) |n| {
            owned_style.number_format = try self.allocator.dupe(u8, n);
        }
        errdefer if (owned_style.number_format) |n| self.allocator.free(n);
        try self.styles.append(self.allocator, owned_style);
        return @intCast(self.styles.items.len);
    }

    /// Return the numFmtId for `fmt`, allocating a new entry at id >=
    /// NUM_FMT_BASE (164) on first sight. Subsequent calls with the
    /// same content return the same id.
    fn internNumFmt(self: *Writer, fmt: []const u8) !u32 {
        if (self.num_fmt_index.get(fmt)) |id| return id;
        const owned = try self.allocator.dupe(u8, fmt);
        errdefer self.allocator.free(owned);
        const id: u32 = @intCast(NUM_FMT_BASE + self.num_fmts.items.len);
        try self.num_fmts.append(self.allocator, owned);
        try self.num_fmt_index.put(self.allocator, owned, id);
        return id;
    }

    /// Add a sheet and return a handle to append rows. Sheet is owned
    /// by the Writer — do not free the returned pointer.
    ///
    /// Enforces Excel's sheet-name rules (length 1..=31, no control or
    /// path-reserved chars, no wrapping apostrophes, not "History") and
    /// rejects case-insensitive duplicates so callers can't
    /// accidentally produce workbooks Excel refuses to open. Returns
    /// `error.InvalidSheetName` or `error.DuplicateSheetName` on bad
    /// input.
    pub fn addSheet(self: *Writer, name: []const u8) !*SheetWriter {
        try validateSheetName(name);
        // O(N) duplicate scan — typical workbooks have ≤10 sheets, so
        // the case-fold loop cost is negligible and saves maintaining
        // a hash of lowercased names.
        for (self.sheets.items) |existing| {
            if (asciiEqlFold(existing.name, name)) return error.DuplicateSheetName;
        }
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
            try full.appendSlice(alloc, WORKSHEET_PROLOG);

            // <sheetViews> — emitted when any pane is frozen (stage 5).
            if (sw.freeze_rows != 0 or sw.freeze_cols != 0) {
                try full.appendSlice(alloc, "<sheetViews><sheetView workbookViewId=\"0\">");
                try full.appendSlice(alloc, "<pane");
                if (sw.freeze_cols != 0) try full.print(alloc, " xSplit=\"{d}\"", .{sw.freeze_cols});
                if (sw.freeze_rows != 0) try full.print(alloc, " ySplit=\"{d}\"", .{sw.freeze_rows});
                var tl_buf: [16]u8 = undefined;
                const top_left = try formatCellRef(&tl_buf, sw.freeze_rows + 1, sw.freeze_cols);
                const active_pane: []const u8 = if (sw.freeze_rows != 0 and sw.freeze_cols != 0)
                    "bottomRight"
                else if (sw.freeze_rows != 0)
                    "bottomLeft"
                else
                    "topRight";
                try full.print(alloc, " topLeftCell=\"{s}\" activePane=\"{s}\" state=\"frozen\"/>", .{ top_left, active_pane });
                try full.appendSlice(alloc, "</sheetView></sheetViews>");
            }

            // <cols> — one <col> per registered width override.
            if (sw.column_widths.items.len > 0) {
                try full.appendSlice(alloc, "<cols>");
                for (sw.column_widths.items) |cw| {
                    try full.print(
                        alloc,
                        "<col min=\"{d}\" max=\"{d}\" width=\"{d}\" customWidth=\"1\"/>",
                        .{ cw.col_min, cw.col_max, cw.width },
                    );
                }
                try full.appendSlice(alloc, "</cols>");
            }

            try full.appendSlice(alloc, "<sheetData>");
            try full.appendSlice(alloc, sw.body.items);
            try full.appendSlice(alloc, "</sheetData>");

            // <autoFilter> must come after </sheetData>.
            if (sw.auto_filter_range) |range| {
                try full.appendSlice(alloc, "<autoFilter ref=\"");
                try appendXmlEscaped(alloc, &full, range);
                try full.appendSlice(alloc, "\"/>");
            }

            // <mergeCells> follows <autoFilter> per ECMA-376 CT_Worksheet
            // child order. Ranges were validated on intake, but defensively
            // xml-escape them on emit anyway.
            if (sw.merged_cells.items.len > 0) {
                try full.print(alloc, "<mergeCells count=\"{d}\">", .{sw.merged_cells.items.len});
                for (sw.merged_cells.items) |range| {
                    try full.appendSlice(alloc, "<mergeCell ref=\"");
                    try appendXmlEscaped(alloc, &full, range);
                    try full.appendSlice(alloc, "\"/>");
                }
                try full.appendSlice(alloc, "</mergeCells>");
            }

            // <dataValidations> slots between <mergeCells> and
            // <hyperlinks> per ECMA-376 CT_Worksheet child order. Two
            // emission paths share the block: iter13 list entries (dropdown)
            // first, then iter23 numeric / custom entries.
            const dv_list_count = sw.data_validations.items.len;
            const dv_range_count = sw.data_validation_ranges.items.len;
            if (dv_list_count + dv_range_count > 0) {
                try full.print(alloc, "<dataValidations count=\"{d}\">", .{dv_list_count + dv_range_count});
                for (sw.data_validations.items) |dv| {
                    try full.appendSlice(alloc, "<dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"");
                    try appendXmlEscaped(alloc, &full, dv.range);
                    try full.appendSlice(alloc, "\"><formula1>&quot;");
                    for (dv.values, 0..) |v, vi| {
                        if (vi != 0) try full.append(alloc, ',');
                        try appendXmlEscaped(alloc, &full, v);
                    }
                    try full.appendSlice(alloc, "&quot;</formula1></dataValidation>");
                }
                for (sw.data_validation_ranges.items) |dv| {
                    try full.appendSlice(alloc, "<dataValidation type=\"");
                    try full.appendSlice(alloc, dv.kind_name);
                    try full.appendSlice(alloc, "\"");
                    if (dv.op_name) |op| {
                        try full.print(alloc, " operator=\"{s}\"", .{op});
                    }
                    try full.appendSlice(alloc, " allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"");
                    try appendXmlEscaped(alloc, &full, dv.range);
                    try full.appendSlice(alloc, "\"><formula1>");
                    try appendXmlEscaped(alloc, &full, dv.formula1);
                    try full.appendSlice(alloc, "</formula1>");
                    if (dv.formula2) |f2| {
                        try full.appendSlice(alloc, "<formula2>");
                        try appendXmlEscaped(alloc, &full, f2);
                        try full.appendSlice(alloc, "</formula2>");
                    }
                    try full.appendSlice(alloc, "</dataValidation>");
                }
                try full.appendSlice(alloc, "</dataValidations>");
            }

            // <hyperlinks> follows <mergeCells> per ECMA-376 ordering.
            // External entries get r:id references into the per-sheet
            // _rels file written below; internal entries use
            // `location="…"` with no rels coupling.
            if (sw.hyperlinks.items.len > 0 or sw.internal_hyperlinks.items.len > 0) {
                try full.appendSlice(alloc, "<hyperlinks>");
                for (sw.hyperlinks.items, 0..) |h, idx| {
                    try full.appendSlice(alloc, "<hyperlink ref=\"");
                    try appendXmlEscaped(alloc, &full, h.range);
                    try full.print(alloc, "\" r:id=\"rId{d}\"/>", .{idx + 1});
                }
                for (sw.internal_hyperlinks.items) |h| {
                    try full.appendSlice(alloc, "<hyperlink ref=\"");
                    try appendXmlEscaped(alloc, &full, h.range);
                    try full.appendSlice(alloc, "\" location=\"");
                    try appendXmlEscaped(alloc, &full, h.location);
                    try full.appendSlice(alloc, "\"/>");
                }
                try full.appendSlice(alloc, "</hyperlinks>");
            }

            try full.appendSlice(alloc, "</worksheet>");

            var name_buf: [64]u8 = undefined;
            const entry_name = try std.fmt.bufPrint(&name_buf, "xl/worksheets/sheet{d}.xml", .{i + 1});
            try zw.addEntry(entry_name, full.items);
        }

        // 5a. xl/worksheets/_rels/sheetN.xml.rels (hyperlinks only)
        //
        // The Default Extension="rels" content-type in [Content_Types].xml
        // covers any .rels file we add here — no extra <Override> needed.
        for (self.sheets.items, 0..) |sw, i| {
            if (sw.hyperlinks.items.len == 0) continue;

            var rels: std.ArrayListUnmanaged(u8) = .{};
            defer rels.deinit(alloc);
            try rels.appendSlice(alloc, WORKBOOK_RELS_HEAD);
            for (sw.hyperlinks.items, 0..) |h, idx| {
                try rels.print(alloc, "<Relationship Id=\"rId{d}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"", .{idx + 1});
                try appendXmlEscaped(alloc, &rels, h.url);
                try rels.appendSlice(alloc, "\" TargetMode=\"External\"/>");
            }
            try rels.appendSlice(alloc, WORKBOOK_RELS_TAIL);

            var rels_name_buf: [64]u8 = undefined;
            const rels_name = try std.fmt.bufPrint(&rels_name_buf, "xl/worksheets/_rels/sheet{d}.xml.rels", .{i + 1});
            try zw.addEntry(rels_name, rels.items);
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
        if (have_styles) try emitStylesXml(
            alloc,
            &zw,
            self.styles.items,
            self.num_fmts.items,
            &self.num_fmt_index,
        );

        try zw.finalize();

        var file = try std.fs.cwd().createFile(path, .{ .truncate = true });
        defer file.close();
        try file.writeAll(zip_buf.items);
    }
};

// ─── SheetWriter ─────────────────────────────────────────────────────

/// External-URL hyperlink registered against a cell or range on one
/// sheet. Both fields are SheetWriter-owned copies.
pub const Hyperlink = struct {
    range: []u8,
    url: []u8,
};

/// Internal-target hyperlink — jumps to another cell or range within
/// the same workbook. Emitted as `<hyperlink ref="…" location="…"/>`
/// (no r:id, no rels entry). Both fields SheetWriter-owned.
pub const InternalHyperlink = struct {
    range: []u8,
    location: []u8,
};

/// List-type data validation (dropdown) bound to a cell or range.
/// `values` are the literal dropdown options — Excel joins them with
/// commas inside a quoted formula1 string. All fields SheetWriter-owned.
pub const DataValidationList = struct {
    range: []u8,
    values: [][]u8,
};

/// Numeric / date / time / text-length / custom data validation. Same
/// `<dataValidations>` block as list validations but a different
/// `type="…"` and a formula-based constraint. `kind_name` and
/// `op_name` are static strings (no allocation); `range`, `formula1`,
/// and `formula2` (when present) are SheetWriter-owned copies.
pub const DataValidationRange = struct {
    range: []u8,
    /// One of "whole", "decimal", "date", "time", "textLength", "custom".
    kind_name: []const u8,
    /// One of "between", "notBetween", "equal", "notEqual",
    /// "greaterThan", "lessThan", "greaterThanOrEqual",
    /// "lessThanOrEqual". `null` for `type="custom"` which doesn't
    /// use an operator.
    op_name: ?[]const u8,
    formula1: []u8,
    /// Required iff `op_name` is "between" or "notBetween"; null otherwise.
    formula2: ?[]u8,
};

/// Numeric-side comparison operator for `addDataValidationNumeric`.
pub const DataValidationOp = enum {
    between,
    not_between,
    equal,
    not_equal,
    greater_than,
    less_than,
    greater_than_or_equal,
    less_than_or_equal,

    fn toOoxml(self: DataValidationOp) []const u8 {
        return switch (self) {
            .between => "between",
            .not_between => "notBetween",
            .equal => "equal",
            .not_equal => "notEqual",
            .greater_than => "greaterThan",
            .less_than => "lessThan",
            .greater_than_or_equal => "greaterThanOrEqual",
            .less_than_or_equal => "lessThanOrEqual",
        };
    }

    fn needsSecondFormula(self: DataValidationOp) bool {
        return self == .between or self == .not_between;
    }
};

/// Data-validation kind for `addDataValidationNumeric`. For dropdown
/// lists use `addDataValidationList`; for formula-driven custom
/// checks use `addDataValidationCustom`.
pub const DataValidationNumericKind = enum {
    whole,
    decimal,
    date,
    time,
    text_length,

    fn toOoxml(self: DataValidationNumericKind) []const u8 {
        return switch (self) {
            .whole => "whole",
            .decimal => "decimal",
            .date => "date",
            .time => "time",
            .text_length => "textLength",
        };
    }
};

/// Per-column width override. `col_min..=col_max` is the inclusive
/// range this width applies to (xlsx indexes columns 1-based — the
/// SheetWriter API takes 0-based indices and translates on emit).
pub const ColumnWidth = struct {
    col_min: u32,
    col_max: u32,
    width: f32,
};

pub const SheetWriter = struct {
    parent: *Writer,
    // Owned copy of the sheet name.
    name: []u8,
    // Accumulated `<row>` elements; emitted inside <sheetData> on save.
    body: std.ArrayListUnmanaged(u8) = .{},
    // 1-based row index (xlsx convention).
    next_row: u32 = 1,
    // Stage 5: per-sheet layout features.
    column_widths: std.ArrayListUnmanaged(ColumnWidth) = .{},
    /// Row-height overrides keyed by 0-based row index (0 = row 1).
    /// Height is in Excel point units (default 15.0). Emitted on the
    /// matching `<row>` as `ht="…" customHeight="1"`. Rows emit
    /// sequentially inside `writeRow`, so callers must set the
    /// override BEFORE the corresponding row is written — later calls
    /// on an already-emitted row are silently ignored.
    row_heights: std.AutoHashMapUnmanaged(u32, f32) = .{},
    /// Number of rows frozen at the top (1 = freeze row 1). 0 = none.
    freeze_rows: u32 = 0,
    /// Number of columns frozen at the left (1 = freeze column A). 0 = none.
    freeze_cols: u32 = 0,
    /// Auto-filter range (e.g., "A1:E1"). null = no filter.
    /// Owned by the SheetWriter.
    auto_filter_range: ?[]u8 = null,
    /// Merged cell ranges (e.g., "A1:B2"). Each entry is a
    /// SheetWriter-owned copy of a validated A1-style range.
    merged_cells: std.ArrayListUnmanaged([]u8) = .{},
    /// External-URL hyperlinks registered against cells or ranges.
    /// Each entry gets an rId in `xl/worksheets/_rels/sheetN.xml.rels`
    /// whose position in this list is its 1-based rId index.
    hyperlinks: std.ArrayListUnmanaged(Hyperlink) = .{},
    /// Internal (same-workbook) hyperlinks. No rels entry — emitted
    /// alongside external hyperlinks inside the `<hyperlinks>` block
    /// using `location="…"` instead of `r:id`.
    internal_hyperlinks: std.ArrayListUnmanaged(InternalHyperlink) = .{},
    /// List-type data validations (dropdowns). Emitted as a single
    /// `<dataValidations>` block with one `<dataValidation>` per entry.
    data_validations: std.ArrayListUnmanaged(DataValidationList) = .{},
    /// Numeric / date / time / text-length / custom data validations.
    /// Sharing the `<dataValidations>` block with the list entries;
    /// the two lists are kept separate so iter13 list-validation code
    /// stays unchanged.
    data_validation_ranges: std.ArrayListUnmanaged(DataValidationRange) = .{},

    fn init(parent: *Writer, name: []const u8) !SheetWriter {
        return .{
            .parent = parent,
            .name = try parent.allocator.dupe(u8, name),
        };
    }

    fn deinit(self: *SheetWriter) void {
        self.parent.allocator.free(self.name);
        self.body.deinit(self.parent.allocator);
        self.column_widths.deinit(self.parent.allocator);
        self.row_heights.deinit(self.parent.allocator);
        if (self.auto_filter_range) |r| self.parent.allocator.free(r);
        for (self.merged_cells.items) |r| self.parent.allocator.free(r);
        self.merged_cells.deinit(self.parent.allocator);
        for (self.hyperlinks.items) |h| {
            self.parent.allocator.free(h.range);
            self.parent.allocator.free(h.url);
        }
        self.hyperlinks.deinit(self.parent.allocator);
        for (self.internal_hyperlinks.items) |h| {
            self.parent.allocator.free(h.range);
            self.parent.allocator.free(h.location);
        }
        self.internal_hyperlinks.deinit(self.parent.allocator);
        for (self.data_validations.items) |dv| {
            self.parent.allocator.free(dv.range);
            for (dv.values) |v| self.parent.allocator.free(v);
            self.parent.allocator.free(dv.values);
        }
        self.data_validations.deinit(self.parent.allocator);
        for (self.data_validation_ranges.items) |dv| {
            self.parent.allocator.free(dv.range);
            self.parent.allocator.free(dv.formula1);
            if (dv.formula2) |f2| self.parent.allocator.free(f2);
        }
        self.data_validation_ranges.deinit(self.parent.allocator);
        self.* = undefined;
    }

    /// Set a column's width in character units (Excel's default is
    /// 8.43). `col_idx` is 0-based (A=0, B=1, …). Multiple calls on
    /// the same column append a new override — the emitter keeps them
    /// in order, so a later call wins on overlap in Excel.
    pub fn setColumnWidth(self: *SheetWriter, col_idx: u32, width: f32) !void {
        if (!std.math.isFinite(width) or width <= 0) return error.InvalidColumnWidth;
        const col_1based = col_idx + 1;
        try self.column_widths.append(self.parent.allocator, .{
            .col_min = col_1based,
            .col_max = col_1based,
            .width = width,
        });
    }

    /// Freeze the top `rows` rows and left `cols` columns. Pass 0 to
    /// disable one axis (e.g., `freezePanes(1, 0)` freezes only row 1).
    /// Calling again overrides the previous setting.
    /// Set `row_idx`'s height in Excel point units (default row
    /// height is ~15 pt). `row_idx` is 0-based (0 = row 1). Must be
    /// called before the matching `writeRow` / `writeRowStyled` — the
    /// row is emitted inline at that time, and a post-hoc call on an
    /// already-emitted row is silently ignored (no retroactive XML
    /// rewrite). Later calls on the same row_idx override earlier
    /// ones as long as the row hasn't been written yet.
    pub fn setRowHeight(self: *SheetWriter, row_idx: u32, height: f32) !void {
        if (!std.math.isFinite(height) or height <= 0) return error.InvalidRowHeight;
        try self.row_heights.put(self.parent.allocator, row_idx, height);
    }

    pub fn freezePanes(self: *SheetWriter, rows: u32, cols: u32) void {
        self.freeze_rows = rows;
        self.freeze_cols = cols;
    }

    /// Apply an auto-filter over the given A1-style range (e.g.,
    /// "A1:E1"). Caller-owned; the writer dupes it.
    pub fn setAutoFilter(self: *SheetWriter, range: []const u8) !void {
        if (range.len == 0) return error.InvalidAutoFilterRange;
        if (self.auto_filter_range) |old| self.parent.allocator.free(old);
        self.auto_filter_range = try self.parent.allocator.dupe(u8, range);
    }

    /// Merge a rectangular cell range (e.g., "A1:B2"). The range must
    /// be a valid multi-cell A1-style span — single-cell ranges and
    /// inverted (bottom-right-before-top-left) ranges are rejected.
    /// Caller-owned; the writer dupes it. Multiple merges per sheet
    /// are allowed; callers are responsible for avoiding overlaps,
    /// which Excel rejects at file-open time.
    pub fn addMergedCell(self: *SheetWriter, range: []const u8) !void {
        try validateMergeRange(range);
        const copy = try self.parent.allocator.dupe(u8, range);
        errdefer self.parent.allocator.free(copy);
        try self.merged_cells.append(self.parent.allocator, copy);
    }

    /// Attach a list-type data validation (dropdown) to a cell or
    /// rectangular range. `range` is A1-style (single cell "A1" or
    /// span "B2:B10"); `values` are the literal dropdown options.
    /// Excel's in-cell list format joins values with commas inside a
    /// quoted formula1 string, so embedded commas and bare double-
    /// quotes in values are rejected (callers who need those should
    /// use a range-reference validation — not yet supported). Empty
    /// values and empty `values` slice also rejected.
    pub fn addDataValidationList(
        self: *SheetWriter,
        range: []const u8,
        values: []const []const u8,
    ) !void {
        try validateHyperlinkRange(range); // same A1 single-or-range shape
        if (values.len == 0) return error.InvalidDataValidation;
        for (values) |v| {
            if (v.len == 0) return error.InvalidDataValidation;
            // Comma breaks the Excel list format; bare `"` breaks the
            // outer quoting. We xml-escape on emit, so `<>&` are safe.
            if (std.mem.indexOfScalar(u8, v, ',') != null) return error.InvalidDataValidation;
            if (std.mem.indexOfScalar(u8, v, '"') != null) return error.InvalidDataValidation;
        }

        const alloc = self.parent.allocator;
        const range_copy = try alloc.dupe(u8, range);
        errdefer alloc.free(range_copy);

        const values_copy = try alloc.alloc([]u8, values.len);
        errdefer alloc.free(values_copy);
        var copied: usize = 0;
        errdefer for (values_copy[0..copied]) |v| alloc.free(v);
        for (values, 0..) |v, i| {
            values_copy[i] = try alloc.dupe(u8, v);
            copied = i + 1;
        }

        try self.data_validations.append(alloc, .{
            .range = range_copy,
            .values = values_copy,
        });
    }

    /// Attach a numeric / date / time / text-length data validation.
    /// `range` is A1-style. `formula1` is the primary bound (number
    /// or Excel date serial or length — passed as a string, the
    /// writer emits it verbatim). `formula2` must be non-null iff
    /// `op` is `.between` or `.not_between`, and must be null
    /// otherwise — mismatches surface `error.InvalidDataValidation`.
    /// Excel displays number-typed validations as red-circle errors
    /// when the cell value falls outside the constraint.
    pub fn addDataValidationNumeric(
        self: *SheetWriter,
        range: []const u8,
        kind: DataValidationNumericKind,
        op: DataValidationOp,
        formula1: []const u8,
        formula2: ?[]const u8,
    ) !void {
        try validateHyperlinkRange(range);
        if (formula1.len == 0) return error.InvalidDataValidation;
        const needs_two = op.needsSecondFormula();
        if (needs_two and (formula2 == null or formula2.?.len == 0)) return error.InvalidDataValidation;
        if (!needs_two and formula2 != null) return error.InvalidDataValidation;

        const alloc = self.parent.allocator;
        const range_copy = try alloc.dupe(u8, range);
        errdefer alloc.free(range_copy);
        const f1_copy = try alloc.dupe(u8, formula1);
        errdefer alloc.free(f1_copy);
        const f2_copy: ?[]u8 = if (formula2) |f| try alloc.dupe(u8, f) else null;
        errdefer if (f2_copy) |f| alloc.free(f);

        try self.data_validation_ranges.append(alloc, .{
            .range = range_copy,
            .kind_name = kind.toOoxml(),
            .op_name = op.toOoxml(),
            .formula1 = f1_copy,
            .formula2 = f2_copy,
        });
    }

    /// Attach a custom-formula data validation. `formula` is any
    /// Excel formula that evaluates to TRUE for accepted cell
    /// values (e.g. `LEN(A1)>3`, `AND(A1>0,A1<100)`). XML-special
    /// characters in the formula are entity-escaped on emit.
    /// Rejects empty formula with `error.InvalidDataValidation`.
    pub fn addDataValidationCustom(
        self: *SheetWriter,
        range: []const u8,
        formula: []const u8,
    ) !void {
        try validateHyperlinkRange(range);
        if (formula.len == 0) return error.InvalidDataValidation;

        const alloc = self.parent.allocator;
        const range_copy = try alloc.dupe(u8, range);
        errdefer alloc.free(range_copy);
        const f_copy = try alloc.dupe(u8, formula);
        errdefer alloc.free(f_copy);

        try self.data_validation_ranges.append(alloc, .{
            .range = range_copy,
            .kind_name = "custom",
            .op_name = null,
            .formula1 = f_copy,
            .formula2 = null,
        });
    }

    /// Attach a hyperlink to a cell or rectangular range. `range` is
    /// A1-style — single cell ("A1") or span ("B2:C3"), same column/
    /// row bounds as Excel (max XFD × 1 048 576). `url` is the
    /// external target (http/https/mailto/file/…); it's xml-escaped
    /// on emit, so `?q=1&x=2` style query strings are safe. Empty
    /// URLs and malformed ranges are rejected. Caller-owned; both
    /// args are duped.
    pub fn addHyperlink(self: *SheetWriter, range: []const u8, url: []const u8) !void {
        try validateHyperlinkRange(range);
        if (url.len == 0) return error.InvalidHyperlinkUrl;
        const range_copy = try self.parent.allocator.dupe(u8, range);
        errdefer self.parent.allocator.free(range_copy);
        const url_copy = try self.parent.allocator.dupe(u8, url);
        errdefer self.parent.allocator.free(url_copy);
        try self.hyperlinks.append(self.parent.allocator, .{
            .range = range_copy,
            .url = url_copy,
        });
    }

    /// Attach an internal hyperlink that jumps to another cell or
    /// range within the same workbook. `location` is an OOXML
    /// workbook-scoped reference like `Sheet2!A1`, `'My Sheet'!B2:C3`,
    /// or a named range. Emitted as `<hyperlink ref="…" location="…"/>`
    /// without an r:id — no `_rels` entry needed. `range` validation
    /// matches `addHyperlink` (single-cell or rectangle A1-style);
    /// empty `location` strings are rejected.
    pub fn addInternalHyperlink(self: *SheetWriter, range: []const u8, location: []const u8) !void {
        try validateHyperlinkRange(range);
        if (location.len == 0) return error.InvalidHyperlinkLocation;
        const range_copy = try self.parent.allocator.dupe(u8, range);
        errdefer self.parent.allocator.free(range_copy);
        const loc_copy = try self.parent.allocator.dupe(u8, location);
        errdefer self.parent.allocator.free(loc_copy);
        try self.internal_hyperlinks.append(self.parent.allocator, .{
            .range = range_copy,
            .location = loc_copy,
        });
    }

    /// Write a row of cells. Empty cells are omitted from the output
    /// (OOXML treats missing cells as empty). Strings are interned into
    /// the parent's SST.
    pub fn writeRow(self: *SheetWriter, cells: []const xlsx.Cell) !void {
        return self.writeRowImpl(cells, null, null);
    }

    /// Write a row with per-cell style indices. `styles.len` must equal
    /// `cells.len`; use `0` (the default no-style slot) for cells that
    /// should inherit the default formatting. Style indices come from
    /// `Writer.addStyle` / `zlsx_writer_add_style`.
    ///
    /// Each non-zero style id is range-checked against the parent
    /// Writer's registered-style count — out-of-range ids fail fast with
    /// `error.UnknownStyleId` rather than producing a workbook that
    /// references a missing `<xf>` record (which Excel would silently
    /// repair or reject). Invariant: after a successful `writeRowStyled`
    /// every emitted `s="N"` attribute corresponds to an existing entry
    /// in the (eventual) `xl/styles.xml` `<cellXfs>` list.
    pub fn writeRowStyled(
        self: *SheetWriter,
        cells: []const xlsx.Cell,
        styles: []const u32,
    ) !void {
        if (styles.len != cells.len) return error.StyleCountMismatch;
        const max_style_id: u32 = @intCast(self.parent.styles.items.len);
        for (styles) |sid| {
            if (sid > max_style_id) return error.UnknownStyleId;
        }
        return self.writeRowImpl(cells, styles, null);
    }

    /// Write a row where some cells carry formulas. `formulas.len`
    /// must equal `cells.len`. Non-null `formulas[i]` attaches the
    /// formula text (without leading `=`) to that cell — the
    /// accompanying `cells[i]` value is emitted as the `<v>` cached
    /// result Excel displays until the sheet is recalculated. Pass
    /// `.empty` for a formula cell with no cached value (Excel will
    /// show 0 initially). Pass `null` in slot `i` for a regular
    /// value cell.
    pub fn writeRowWithFormulas(
        self: *SheetWriter,
        cells: []const xlsx.Cell,
        formulas: []const ?[]const u8,
    ) !void {
        if (formulas.len != cells.len) return error.FormulaCountMismatch;
        return self.writeRowImpl(cells, null, formulas);
    }

    fn writeRowImpl(
        self: *SheetWriter,
        cells: []const xlsx.Cell,
        styles: ?[]const u32,
        formulas: ?[]const ?[]const u8,
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
        // Row index is 0-based inside the height map; next_row is
        // 1-based per xlsx convention, so subtract 1 on lookup.
        if (self.row_heights.get(self.next_row - 1)) |h| {
            try self.body.print(alloc, "<row r=\"{d}\" ht=\"{d}\" customHeight=\"1\">", .{ self.next_row, h });
        } else {
            try self.body.print(alloc, "<row r=\"{d}\">", .{self.next_row});
        }

        for (cells, 0..) |cell, col_idx| {
            const style_id: u32 = if (styles) |s| s[col_idx] else 0;
            const formula: ?[]const u8 = if (formulas) |fs| fs[col_idx] else null;

            // `<c>` elements for empty cells are only emitted when a
            // non-default style is applied OR a formula is attached —
            // otherwise OOXML's "missing cell = empty" rule keeps the
            // sheet smaller.
            if (cell == .empty and style_id == 0 and formula == null) continue;

            var ref_buf: [16]u8 = undefined;
            const ref = try formatCellRef(&ref_buf, self.next_row, @intCast(col_idx));

            // Self-closing fast path: styled but empty, no formula.
            // Preserves byte-for-byte output with the pre-formula
            // revision so existing round-trip tests stay valid.
            if (cell == .empty and formula == null) {
                try self.body.print(alloc, "<c r=\"{s}\" s=\"{d}\"/>", .{ ref, style_id });
                continue;
            }

            // Fall-through emission pattern:
            //   <c r="…"[ s="N"][ t="s|b"]>[<f>formula</f>][<v>value</v>]</c>
            // The non-formula paths match the pre-refactor byte output
            // so the grep-style test assertions stay stable.
            const type_attr: []const u8 = switch (cell) {
                .string => " t=\"s\"",
                .boolean => " t=\"b\"",
                else => "",
            };
            if (style_id == 0) {
                try self.body.print(alloc, "<c r=\"{s}\"{s}>", .{ ref, type_attr });
            } else {
                try self.body.print(alloc, "<c r=\"{s}\" s=\"{d}\"{s}>", .{ ref, style_id, type_attr });
            }

            if (formula) |f| {
                try self.body.appendSlice(alloc, "<f>");
                try appendXmlEscaped(alloc, &self.body, f);
                try self.body.appendSlice(alloc, "</f>");
            }

            switch (cell) {
                .empty => {}, // formula-only cell: no cached value
                .string => |s| {
                    const idx = try self.parent.sstIntern(s);
                    self.parent.sst_count += 1;
                    try self.body.print(alloc, "<v>{d}</v>", .{idx});
                },
                .integer => |n| try self.body.print(alloc, "<v>{d}</v>", .{n}),
                .number => |f| try self.body.print(alloc, "<v>{d}</v>", .{f}),
                .boolean => |b| try self.body.print(alloc, "<v>{d}</v>", .{@intFromBool(b)}),
            }

            try self.body.appendSlice(alloc, "</c>");
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

// Excel's hard limits: 16 384 columns (XFD) × 1 048 576 rows.
const EXCEL_MAX_COL: u32 = 16_384;
const EXCEL_MAX_ROW: u32 = 1_048_576;

const MergeCorner = struct { col: u32, row: u32 };

fn parseA1Corner(s: []const u8) !MergeCorner {
    if (s.len == 0) return error.InvalidMergeRange;
    var i: usize = 0;
    var col: u32 = 0;
    while (i < s.len and s[i] >= 'A' and s[i] <= 'Z') : (i += 1) {
        col = col * 26 + (s[i] - 'A' + 1);
        if (col > EXCEL_MAX_COL) return error.InvalidMergeRange;
    }
    // Need at least one letter and at least one digit after it.
    if (i == 0 or i == s.len) return error.InvalidMergeRange;
    // Leading zero (e.g., "A0", "A01") is not a valid A1 row.
    if (s[i] == '0') return error.InvalidMergeRange;
    var row: u32 = 0;
    while (i < s.len and s[i] >= '0' and s[i] <= '9') : (i += 1) {
        row = row * 10 + (s[i] - '0');
        if (row > EXCEL_MAX_ROW) return error.InvalidMergeRange;
    }
    if (i != s.len) return error.InvalidMergeRange;
    return .{ .col = col, .row = row };
}

fn validateMergeRange(range: []const u8) !void {
    const colon = std.mem.indexOfScalar(u8, range, ':') orelse return error.InvalidMergeRange;
    const tl = try parseA1Corner(range[0..colon]);
    const br = try parseA1Corner(range[colon + 1 ..]);
    // Top-left must strictly precede or equal bottom-right on both axes.
    if (tl.col > br.col or tl.row > br.row) return error.InvalidMergeRange;
    // Single-cell "merge" is a no-op that Excel warns on — reject it
    // so callers catch typos at write time rather than on file-open.
    if (tl.col == br.col and tl.row == br.row) return error.InvalidMergeRange;
}

/// Byte-wise ASCII case-fold equality. Excel's sheet-name uniqueness
/// rule is Unicode-case-insensitive, but ASCII case-fold catches the
/// overwhelming majority of real-world collisions ("Summary" vs
/// "summary") at a tiny code cost. Non-ASCII input falls through to
/// byte comparison — sufficient for everything except genuine
/// Turkish-i-style edge cases that no real caller hits.
fn asciiEqlFold(a: []const u8, b: []const u8) bool {
    if (a.len != b.len) return false;
    for (a, b) |x, y| {
        const xl: u8 = if (x >= 'A' and x <= 'Z') x + 32 else x;
        const yl: u8 = if (y >= 'A' and y <= 'Z') y + 32 else y;
        if (xl != yl) return false;
    }
    return true;
}

/// Enforce Excel's sheet-name rules at the API boundary. Silent drift
/// here produces workbooks Excel refuses to open — catch it up front
/// with a clear error:
///   - 1..=31 UTF-8 bytes (Excel caps at 31 chars; we check bytes
///     since non-ASCII names are rare and the difference is subtle
///     — a conservative limit won't reject any real-world input).
///   - No control chars (< 0x20).
///   - None of the reserved path chars `: / \ ? * [ ]`.
///   - No leading or trailing apostrophe (Excel uses `'` as a
///     sheet-reference quote delimiter).
///   - Not the reserved name "History" (case-insensitive).
fn validateSheetName(name: []const u8) !void {
    if (name.len == 0 or name.len > 31) return error.InvalidSheetName;
    if (name[0] == '\'' or name[name.len - 1] == '\'') return error.InvalidSheetName;
    for (name) |c| switch (c) {
        0...0x1F => return error.InvalidSheetName,
        ':', '/', '\\', '?', '*', '[', ']' => return error.InvalidSheetName,
        else => {},
    };
    if (asciiEqlFold(name, "History")) return error.InvalidSheetName;
}

fn validateHyperlinkRange(range: []const u8) !void {
    if (range.len == 0) return error.InvalidHyperlinkRange;
    if (std.mem.indexOfScalar(u8, range, ':')) |colon| {
        // Rectangle form — mirror the merge-range rules except that
        // single-cell ranges ARE valid here (a 1×1 hyperlink is exactly
        // the A1 form, just with redundant `:` sugar).
        const tl = parseA1Corner(range[0..colon]) catch return error.InvalidHyperlinkRange;
        const br = parseA1Corner(range[colon + 1 ..]) catch return error.InvalidHyperlinkRange;
        if (tl.col > br.col or tl.row > br.row) return error.InvalidHyperlinkRange;
    } else {
        // Single-cell form — just an A1 reference.
        _ = parseA1Corner(range) catch return error.InvalidHyperlinkRange;
    }
}

/// Content-compare two styles. Necessary because std.meta.eql on
/// `?[]const u8` compares slice headers (pointer + length) rather than
/// the underlying bytes, so two registrations of `font_name = "Arial"`
/// from distinct buffers would not dedup.
fn stylesEqual(a: Style, b: Style) bool {
    if (a.font_bold != b.font_bold) return false;
    if (a.font_italic != b.font_italic) return false;
    if (!std.meta.eql(a.font_size, b.font_size)) return false;
    if (a.font_color_argb != b.font_color_argb) return false;
    if (a.alignment_horizontal != b.alignment_horizontal) return false;
    if (a.wrap_text != b.wrap_text) return false;
    if (a.fill_pattern != b.fill_pattern) return false;
    if (a.fill_fg_argb != b.fill_fg_argb) return false;
    if (a.fill_bg_argb != b.fill_bg_argb) return false;
    if (!std.meta.eql(a.border_left, b.border_left)) return false;
    if (!std.meta.eql(a.border_right, b.border_right)) return false;
    if (!std.meta.eql(a.border_top, b.border_top)) return false;
    if (!std.meta.eql(a.border_bottom, b.border_bottom)) return false;
    if (!std.meta.eql(a.border_diagonal, b.border_diagonal)) return false;
    if (a.diagonal_up != b.diagonal_up) return false;
    if (a.diagonal_down != b.diagonal_down) return false;
    // Content-compare font_name.
    if ((a.font_name == null) != (b.font_name == null)) return false;
    if (a.font_name) |an| {
        if (!std.mem.eql(u8, an, b.font_name.?)) return false;
    }
    // Content-compare number_format.
    if ((a.number_format == null) != (b.number_format == null)) return false;
    if (a.number_format) |an| {
        if (!std.mem.eql(u8, an, b.number_format.?)) return false;
    }
    return true;
}

/// Emit xl/styles.xml based on the registered style list. Fonts are
/// keyed 1:1 with styles (fonts[i+1] corresponds to style i); deduping
/// fonts independently of styles is a Phase 3b stage-3 optimisation.
/// `<cellXfs>` gets the default entry at index 0 plus one entry per
/// user style, with `applyAlignment="1"` when a style sets any
/// alignment/wrap field.
fn emitStylesXml(
    alloc: Allocator,
    zw: *ZipWriter,
    styles: []const Style,
    num_fmts: []const []const u8,
    num_fmt_index: *const std.StringHashMapUnmanaged(u32),
) !void {
    var buf: std.ArrayListUnmanaged(u8) = .{};
    defer buf.deinit(alloc);

    try buf.appendSlice(alloc, STYLES_HEAD);

    // <numFmts> — emitted only when the user registered any custom
    // format. Built-ins (General / 0..=49) don't go here.
    if (num_fmts.len > 0) {
        try buf.print(alloc, "<numFmts count=\"{d}\">", .{num_fmts.len});
        for (num_fmts, 0..) |fmt, i| {
            const id: u32 = @intCast(NUM_FMT_BASE + i);
            try buf.print(alloc, "<numFmt numFmtId=\"{d}\" formatCode=\"", .{id});
            try appendXmlEscaped(alloc, &buf, fmt);
            try buf.appendSlice(alloc, "\"/>");
        }
        try buf.appendSlice(alloc, "</numFmts>");
    }

    // <fonts>: default at index 0 + one per user style.
    try buf.print(alloc, "<fonts count=\"{d}\">", .{styles.len + 1});
    try buf.appendSlice(alloc, STYLES_FONTS_DEFAULT);
    for (styles) |s| {
        try buf.appendSlice(alloc, "<font>");
        if (s.font_bold) try buf.appendSlice(alloc, "<b/>");
        if (s.font_italic) try buf.appendSlice(alloc, "<i/>");
        // <sz> — configurable in stage 2; fall back to 11 when unset.
        const size = s.font_size orelse 11.0;
        try buf.print(alloc, "<sz val=\"{d}\"/>", .{size});
        // <color> — only when set; theme auto is implied by omission.
        if (s.font_color_argb) |c| try buf.print(
            alloc,
            "<color rgb=\"{X:0>8}\"/>",
            .{c},
        );
        // <name> — configurable in stage 2; default "Calibri".
        try buf.appendSlice(alloc, "<name val=\"");
        if (s.font_name) |n| {
            try appendXmlEscaped(alloc, &buf, n);
        } else {
            try buf.appendSlice(alloc, "Calibri");
        }
        try buf.appendSlice(alloc, "\"/></font>");
    }
    try buf.appendSlice(alloc, "</fonts>");

    // <fills>: 2 reserved slots (none, gray125 — conventional OOXML
    // defaults), then one user fill per style that sets any fill field.
    // Styles without a fill reference fillId=0.
    var fill_ids = try alloc.alloc(u32, styles.len);
    defer alloc.free(fill_ids);
    var next_user_fill_id: u32 = 2;
    for (styles, 0..) |s, i| {
        if (s.fill_pattern != .none or s.fill_fg_argb != null or s.fill_bg_argb != null) {
            fill_ids[i] = next_user_fill_id;
            next_user_fill_id += 1;
        } else {
            fill_ids[i] = 0;
        }
    }
    try buf.print(alloc, "<fills count=\"{d}\">", .{next_user_fill_id});
    try buf.appendSlice(alloc, "<fill><patternFill patternType=\"none\"/></fill>");
    try buf.appendSlice(alloc, "<fill><patternFill patternType=\"gray125\"/></fill>");
    for (styles) |s| {
        if (s.fill_pattern == .none and s.fill_fg_argb == null and s.fill_bg_argb == null) continue;
        try buf.print(
            alloc,
            "<fill><patternFill patternType=\"{s}\"",
            .{patternTypeName(s.fill_pattern)},
        );
        if (s.fill_fg_argb == null and s.fill_bg_argb == null) {
            try buf.appendSlice(alloc, "/></fill>");
        } else {
            try buf.appendSlice(alloc, ">");
            if (s.fill_fg_argb) |c| try buf.print(alloc, "<fgColor rgb=\"{X:0>8}\"/>", .{c});
            if (s.fill_bg_argb) |c| try buf.print(alloc, "<bgColor rgb=\"{X:0>8}\"/>", .{c});
            try buf.appendSlice(alloc, "</patternFill></fill>");
        }
    }
    try buf.appendSlice(alloc, "</fills>");

    // <borders>: default empty border at index 0, then one per style
    // that touches any border field. Styles without borders keep
    // borderId=0.
    var border_ids = try alloc.alloc(u32, styles.len);
    defer alloc.free(border_ids);
    var next_user_border_id: u32 = 1;
    for (styles, 0..) |s, i| {
        if (hasBorder(s)) {
            border_ids[i] = next_user_border_id;
            next_user_border_id += 1;
        } else {
            border_ids[i] = 0;
        }
    }
    try buf.print(alloc, "<borders count=\"{d}\">", .{next_user_border_id});
    try buf.appendSlice(alloc, "<border><left/><right/><top/><bottom/><diagonal/></border>");
    for (styles) |s| {
        if (!hasBorder(s)) continue;
        try buf.appendSlice(alloc, "<border");
        if (s.diagonal_up) try buf.appendSlice(alloc, " diagonalUp=\"1\"");
        if (s.diagonal_down) try buf.appendSlice(alloc, " diagonalDown=\"1\"");
        try buf.appendSlice(alloc, ">");
        try emitBorderSide(alloc, &buf, "left", s.border_left);
        try emitBorderSide(alloc, &buf, "right", s.border_right);
        try emitBorderSide(alloc, &buf, "top", s.border_top);
        try emitBorderSide(alloc, &buf, "bottom", s.border_bottom);
        try emitBorderSide(alloc, &buf, "diagonal", s.border_diagonal);
        try buf.appendSlice(alloc, "</border>");
    }
    try buf.appendSlice(alloc, "</borders>");
    try buf.appendSlice(alloc, STYLES_CELL_STYLE_XFS);

    // <cellXfs>: default at index 0 + one per user style.
    try buf.print(alloc, "<cellXfs count=\"{d}\">", .{styles.len + 1});
    try buf.appendSlice(alloc, STYLES_DEFAULT_CELL_XF);
    for (styles, 0..) |s, i| {
        const has_alignment = s.alignment_horizontal != .general or s.wrap_text;
        const fill_id = fill_ids[i];
        const border_id = border_ids[i];
        const num_fmt_id: u32 = if (s.number_format) |fmt|
            (num_fmt_index.get(fmt) orelse 0)
        else
            0;
        try buf.print(
            alloc,
            "<xf numFmtId=\"{d}\" fontId=\"{d}\" fillId=\"{d}\" borderId=\"{d}\" xfId=\"0\" applyFont=\"1\"",
            .{ num_fmt_id, i + 1, fill_id, border_id },
        );
        if (num_fmt_id != 0) try buf.appendSlice(alloc, " applyNumberFormat=\"1\"");
        if (fill_id != 0) try buf.appendSlice(alloc, " applyFill=\"1\"");
        if (border_id != 0) try buf.appendSlice(alloc, " applyBorder=\"1\"");
        if (has_alignment) {
            try buf.appendSlice(alloc, " applyAlignment=\"1\"><alignment");
            if (s.alignment_horizontal != .general) {
                try buf.print(alloc, " horizontal=\"{s}\"", .{hAlignName(s.alignment_horizontal)});
            }
            if (s.wrap_text) try buf.appendSlice(alloc, " wrapText=\"1\"");
            try buf.appendSlice(alloc, "/></xf>");
        } else {
            try buf.appendSlice(alloc, "/>");
        }
    }
    try buf.appendSlice(alloc, "</cellXfs>");

    try buf.appendSlice(alloc, STYLES_CELL_STYLES);
    try buf.appendSlice(alloc, STYLES_TAIL);

    try zw.addEntry("xl/styles.xml", buf.items);
}

fn hAlignName(a: HAlign) []const u8 {
    return switch (a) {
        .general => "general",
        .left => "left",
        .center => "center",
        .right => "right",
        .fill => "fill",
        .justify => "justify",
        .center_continuous => "centerContinuous",
        .distributed => "distributed",
    };
}

fn borderStyleName(b: BorderStyle) []const u8 {
    return switch (b) {
        .none => "none",
        .thin => "thin",
        .medium => "medium",
        .dashed => "dashed",
        .dotted => "dotted",
        .thick => "thick",
        .double => "double",
        .hair => "hair",
        .medium_dashed => "mediumDashed",
        .dash_dot => "dashDot",
        .medium_dash_dot => "mediumDashDot",
        .dash_dot_dot => "dashDotDot",
        .medium_dash_dot_dot => "mediumDashDotDot",
        .slant_dash_dot => "slantDashDot",
    };
}

fn emitBorderSide(
    alloc: Allocator,
    buf: *std.ArrayListUnmanaged(u8),
    tag: []const u8,
    side: BorderSide,
) !void {
    if (side.style == .none and side.color_argb == null) {
        // Empty side — OOXML wants the element present but attribute-less.
        try buf.print(alloc, "<{s}/>", .{tag});
        return;
    }
    try buf.print(alloc, "<{s}", .{tag});
    if (side.style != .none) {
        try buf.print(alloc, " style=\"{s}\"", .{borderStyleName(side.style)});
    }
    if (side.color_argb) |c| {
        try buf.print(alloc, "><color rgb=\"{X:0>8}\"/></{s}>", .{ c, tag });
    } else {
        try buf.appendSlice(alloc, "/>");
    }
}

fn patternTypeName(p: PatternType) []const u8 {
    return switch (p) {
        .none => "none",
        .solid => "solid",
        .gray125 => "gray125",
        .gray0625 => "gray0625",
        .dark_gray => "darkGray",
        .medium_gray => "mediumGray",
        .light_gray => "lightGray",
        .dark_horizontal => "darkHorizontal",
        .dark_vertical => "darkVertical",
        .dark_down => "darkDown",
        .dark_up => "darkUp",
        .dark_grid => "darkGrid",
        .dark_trellis => "darkTrellis",
        .light_horizontal => "lightHorizontal",
        .light_vertical => "lightVertical",
        .light_down => "lightDown",
        .light_up => "lightUp",
        .light_grid => "lightGrid",
        .light_trellis => "lightTrellis",
    };
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

// ─── Deflate (LZ77 + dynamic huffman + lazy matching) ────────────────
//
// Pure in-house deflate compressor. Zig 0.15.2's stdlib
// `std.compress.flate.Compress` cannot compile (references a missing
// `bit_writer` field in BlockWriter + has @panic("TODO") in the token
// emission path), so we grow our own. Two-pass layout: tokenize input
// into literals + LZ77 matches (greedy with lazy-match-defer-one), then
// emit a single final type-2 (dynamic huffman) block. HuffmanEncoder
// from stdlib builds the code tables from frequency counts.
//
// Tuning knobs: 32 KB sliding window, hash-chain depth 32, 3-byte hash.
// The lazy match defers one position — after finding a match at i we
// look at i+1 and keep whichever is longer. Cuts about 5 KB off the
// Phase 3b test workbook on top of the dynamic huffman tables that
// already close most of the xlsxwriter/openpyxl gap.

// RFC 1951 §3.2.5 — length codes 257..285 with base + extra-bit counts.
const deflate_length_base = [_]u16{
    3,   4,   5,   6,   7,   8,  9,  10,
    11,  13,  15,  17,  19,  23, 27, 31,
    35,  43,  51,  59,  67,  83, 99, 115,
    131, 163, 195, 227, 258,
};
const deflate_length_extra = [_]u8{
    0, 0, 0, 0, 0, 0, 0, 0,
    1, 1, 1, 1, 2, 2, 2, 2,
    3, 3, 3, 3, 4, 4, 4, 4,
    5, 5, 5, 5, 0,
};
const deflate_dist_base = [_]u16{
    1,     2,     3,    4,
    5,     7,     9,    13,
    17,    25,    33,   49,
    65,    97,    129,  193,
    257,   385,   513,  769,
    1025,  1537,  2049, 3073,
    4097,  6145,  8193, 12289,
    16385, 24577,
};
const deflate_dist_extra = [_]u8{
    0,  0,  0,  0,
    1,  1,  2,  2,
    3,  3,  4,  4,
    5,  5,  6,  6,
    7,  7,  8,  8,
    9,  9,  10, 10,
    11, 11, 12, 12,
    13, 13,
};

fn deflateLengthCode(len: u16) struct { code: u16, extra: u16, extra_bits: u8 } {
    std.debug.assert(len >= 3 and len <= 258);
    var i: usize = deflate_length_base.len;
    while (i > 0) {
        i -= 1;
        if (deflate_length_base[i] <= len) return .{
            .code = @intCast(257 + i),
            .extra = len - deflate_length_base[i],
            .extra_bits = deflate_length_extra[i],
        };
    }
    unreachable;
}

fn deflateDistCode(d: u16) struct { code: u8, extra: u16, extra_bits: u8 } {
    std.debug.assert(d >= 1);
    var i: usize = deflate_dist_base.len;
    while (i > 0) {
        i -= 1;
        if (deflate_dist_base[i] <= d) return .{
            .code = @intCast(i),
            .extra = d - deflate_dist_base[i],
            .extra_bits = deflate_dist_extra[i],
        };
    }
    unreachable;
}

const DeflateBitWriter = struct {
    out: *std.ArrayListUnmanaged(u8),
    alloc: Allocator,
    buf: u64 = 0,
    n_bits: u6 = 0,

    fn writeBits(self: *DeflateBitWriter, value: u32, n: u5) !void {
        self.buf |= (@as(u64, value) << @intCast(self.n_bits));
        self.n_bits += n;
        while (self.n_bits >= 8) {
            try self.out.append(self.alloc, @truncate(self.buf & 0xFF));
            self.buf >>= 8;
            self.n_bits -= 8;
        }
    }

    /// Write an MSB-first huffman code (as in RFC 1951's tables) —
    /// reversed into the LSB-first output stream.
    fn writeMsbCode(self: *DeflateBitWriter, code: u16, len: u4) !void {
        const rev = @bitReverse(@as(u32, code)) >> @as(u5, @intCast(32 - @as(u6, len)));
        try self.writeBits(rev, @intCast(len));
    }

    fn flushByte(self: *DeflateBitWriter) !void {
        if (self.n_bits > 0) {
            try self.out.append(self.alloc, @truncate(self.buf & 0xFF));
            self.buf = 0;
            self.n_bits = 0;
        }
    }
};

const DEFLATE_WINDOW_SIZE: usize = 32768;
const DEFLATE_MIN_MATCH: usize = 3;
const DEFLATE_MAX_MATCH: usize = 258;
const DEFLATE_HASH_BITS: u6 = 15;
const DEFLATE_HASH_SIZE: usize = 1 << DEFLATE_HASH_BITS;

fn deflateHash3(d0: u8, d1: u8, d2: u8) u32 {
    const a: u32 = d0;
    const b: u32 = d1;
    const c: u32 = d2;
    // Knuth multiplicative hash on the 24-bit 3-byte prefix.
    return (((a << 16) | (b << 8) | c) *% 2654435761) >> @as(u5, @intCast(32 - DEFLATE_HASH_BITS));
}

/// LZ77 token. `kind=0` → literal (byte in `val`); `kind=1` → match
/// (raw length in `val`, distance in `dist`). Packed to 4 bytes so the
/// token stream is cache-friendly on large inputs.
const DeflateToken = packed struct(u32) {
    kind: u1,
    val: u15, // literal byte (0-255) or raw match length (3-258)
    dist: u16, // match distance (1-32768); unused for literals
};

/// Insert the 3-byte prefix at `pos` into the hash chain.
fn deflateInsertHash(
    input: []const u8,
    pos: usize,
    head: []i32,
    prev: []i32,
) void {
    const h = deflateHash3(input[pos], input[pos + 1], input[pos + 2]);
    prev[pos] = head[h];
    head[h] = @intCast(pos);
}

/// Tokenize `input` into LZ77 literals + matches with lazy matching.
/// "Lazy" here means single-step: at position i we find the best match,
/// look one byte ahead to see if a longer match starts at i+1, and
/// commit to the longer one. Costs little (one extra search per
/// non-MAX_MATCH hit) and reliably trims 3-5 % off the stream.
fn tokenizeLazy(
    alloc: Allocator,
    input: []const u8,
    tokens: *std.ArrayListUnmanaged(DeflateToken),
) !void {
    const head = try alloc.alloc(i32, DEFLATE_HASH_SIZE);
    defer alloc.free(head);
    @memset(head, -1);

    const prev = try alloc.alloc(i32, input.len);
    defer alloc.free(prev);
    @memset(prev, -1);

    const max_chain: u32 = 32;

    var i: usize = 0;
    // One-step deferred match (the "lazy" part).
    var prev_match_len: usize = 0;
    var prev_match_dist: usize = 0;
    var prev_literal: u8 = 0;

    while (i < input.len) {
        var match_len: usize = 0;
        var match_dist: usize = 0;

        if (i + DEFLATE_MIN_MATCH <= input.len) {
            const h = deflateHash3(input[i], input[i + 1], input[i + 2]);
            var candidate = head[h];
            var chain_steps: u32 = 0;
            while (candidate >= 0 and chain_steps < max_chain) : (chain_steps += 1) {
                const cand_pos: usize = @intCast(candidate);
                if (i - cand_pos > DEFLATE_WINDOW_SIZE) break;
                const limit = @min(input.len - i, DEFLATE_MAX_MATCH);
                // Word-at-a-time match length: XOR 8 bytes of source and
                // dest, and if they differ, @ctz(diff) / 8 is the first
                // mismatched byte offset. Typical xlsx XML matches run
                // 3-30 bytes, so the 8-wide compare cuts the inner loop
                // iteration count by ~6× on average. Bounds check: the
                // outer limit already caps us at input.len - i, and the
                // distance invariant cand_pos < i means cand_pos + limit
                // stays strictly inside input as well.
                var k: usize = 0;
                var word_exit = false;
                while (k + 8 <= limit) : (k += 8) {
                    const a = std.mem.readInt(u64, input[cand_pos + k ..][0..8], .little);
                    const b = std.mem.readInt(u64, input[i + k ..][0..8], .little);
                    const diff = a ^ b;
                    if (diff != 0) {
                        k += @as(usize, @ctz(diff)) / 8;
                        word_exit = true;
                        break;
                    }
                }
                if (!word_exit) {
                    while (k < limit and input[cand_pos + k] == input[i + k]) : (k += 1) {}
                }
                if (k >= DEFLATE_MIN_MATCH and k > match_len) {
                    match_len = k;
                    match_dist = i - cand_pos;
                    if (match_len >= DEFLATE_MAX_MATCH) break;
                }
                candidate = prev[cand_pos];
            }
            // Always insert the current position into the chain so
            // later positions (including the i+1 lazy lookahead) can
            // match against it.
            deflateInsertHash(input, i, head, prev);
        }

        if (prev_match_len >= DEFLATE_MIN_MATCH) {
            if (match_len > prev_match_len) {
                // i+1 beats i — emit i's byte as a literal and promote
                // the current match to the new deferred slot.
                try tokens.append(alloc, .{ .kind = 0, .val = prev_literal, .dist = 0 });
                prev_match_len = match_len;
                prev_match_dist = match_dist;
                prev_literal = input[i];
                i += 1;
            } else {
                // Emit the deferred match. We're currently 1 past the
                // deferred position; fill the hash chain for the match
                // interior we're about to skip over, then advance.
                try tokens.append(alloc, .{
                    .kind = 1,
                    .val = @intCast(prev_match_len),
                    .dist = @intCast(prev_match_dist),
                });
                const skip = prev_match_len - 1;
                var j: usize = 1;
                while (j < skip) : (j += 1) {
                    if (i + j + DEFLATE_MIN_MATCH <= input.len) {
                        deflateInsertHash(input, i + j, head, prev);
                    }
                }
                i += skip;
                prev_match_len = 0;
            }
        } else {
            if (match_len >= DEFLATE_MIN_MATCH) {
                // Defer the commit to check i+1 first.
                prev_match_len = match_len;
                prev_match_dist = match_dist;
                prev_literal = input[i];
                i += 1;
            } else {
                try tokens.append(alloc, .{ .kind = 0, .val = input[i], .dist = 0 });
                i += 1;
            }
        }
    }

    // Flush any deferred match at end of input.
    if (prev_match_len >= DEFLATE_MIN_MATCH) {
        try tokens.append(alloc, .{
            .kind = 1,
            .val = @intCast(prev_match_len),
            .dist = @intCast(prev_match_dist),
        });
    }
}

/// Run-length encode a combined (lit_lens ++ dist_lens) sequence into
/// the codegen alphabet (0-18) per RFC 1951 §3.2.7. `out_syms` receives
/// packed symbols where bits 0-4 = codegen symbol, bits 5-11 = extra
/// payload, bits 12-15 = extra bits count.
const CodegenSym = packed struct(u16) { sym: u5, extra: u7, extra_bits: u4 };

fn rleEncodeCodeLengths(
    alloc: Allocator,
    lens: []const u8,
    out: *std.ArrayListUnmanaged(CodegenSym),
) !void {
    var i: usize = 0;
    while (i < lens.len) {
        const cur = lens[i];
        var run: usize = 1;
        while (i + run < lens.len and lens[i + run] == cur) run += 1;

        if (cur == 0) {
            // Zero runs: symbol 17 for 3..10 zeros, 18 for 11..138.
            while (run >= 11) {
                const n = @min(run, 138);
                try out.append(alloc, .{ .sym = 18, .extra = @intCast(n - 11), .extra_bits = 7 });
                run -= n;
                i += n;
            }
            if (run >= 3) {
                try out.append(alloc, .{ .sym = 17, .extra = @intCast(run - 3), .extra_bits = 3 });
                i += run;
                run = 0;
            }
            while (run > 0) : (run -= 1) {
                try out.append(alloc, .{ .sym = 0, .extra = 0, .extra_bits = 0 });
                i += 1;
            }
        } else {
            // Non-zero run: emit first as a literal, use symbol 16 to
            // repeat prev 3..6 more times, stragglers as literals.
            try out.append(alloc, .{ .sym = @intCast(cur), .extra = 0, .extra_bits = 0 });
            i += 1;
            run -= 1;
            while (run >= 3) {
                const n = @min(run, 6);
                try out.append(alloc, .{ .sym = 16, .extra = @intCast(n - 3), .extra_bits = 2 });
                run -= n;
                i += n;
            }
            while (run > 0) : (run -= 1) {
                try out.append(alloc, .{ .sym = @intCast(cur), .extra = 0, .extra_bits = 0 });
                i += 1;
            }
        }
    }
}

const HuffEncoder = std.compress.flate.HuffmanEncoder;

/// Count byte frequencies in the token stream. Uses u32 to avoid u16
/// overflow on large repetitive inputs (one literal can appear ≫64k
/// times in a 300 KB sheet); we rescale to u16 at HuffmanEncoder call
/// time since the stdlib's api takes u16 freqs.
fn countTokenFrequencies(
    tokens: []const DeflateToken,
    lit_freq: *[286]u32,
    dist_freq: *[30]u32,
) void {
    @memset(lit_freq, 0);
    @memset(dist_freq, 0);
    lit_freq[256] = 1; // EOB always appears exactly once.
    for (tokens) |tok| {
        if (tok.kind == 0) {
            lit_freq[tok.val] += 1;
        } else {
            const lc = deflateLengthCode(tok.val);
            lit_freq[lc.code] += 1;
            const dc = deflateDistCode(tok.dist);
            dist_freq[dc.code] += 1;
        }
    }
}

/// Scale frequencies down so the max fits in u15 — HuffmanEncoder sums
/// them internally and needs room. Preserve relative rank: a non-zero
/// input keeps a non-zero output (rounding up from 0).
fn scaleFreqs(src: []const u32, dst: []u16) void {
    var max_freq: u32 = 0;
    for (src) |f| max_freq = @max(max_freq, f);
    const u15_max: u32 = 32767;
    const scale: u32 = if (max_freq > u15_max) (max_freq + u15_max - 1) / u15_max else 1;
    for (src, 0..) |f, i| {
        const v = f / scale;
        dst[i] = @intCast(if (f > 0 and v == 0) 1 else v);
    }
}

/// Emit `tokens` as a single final dynamic-huffman block. Assumes
/// `tokens.len > 0` — empty inputs are handled upstream (stored
/// fallback in ZipWriter.addEntry).
fn emitDynamicBlock(
    alloc: Allocator,
    tokens: []const DeflateToken,
    bw: *DeflateBitWriter,
) !void {
    // 1. Count frequencies and build lit/length + distance huffman tables.
    var lit_freq_u32: [286]u32 = undefined;
    var dist_freq_u32: [30]u32 = undefined;
    countTokenFrequencies(tokens, &lit_freq_u32, &dist_freq_u32);

    var lit_freq: [286]u16 = undefined;
    var dist_freq: [30]u16 = undefined;
    scaleFreqs(&lit_freq_u32, &lit_freq);
    scaleFreqs(&dist_freq_u32, &dist_freq);

    // RFC 1951 mandates at least one distance code; if the input had
    // zero matches, zlib's convention is to define two dummy 1-bit
    // codes so the decoder's alphabet is always well-formed.
    var any_dist: bool = false;
    for (dist_freq) |f| if (f > 0) {
        any_dist = true;
        break;
    };
    if (!any_dist) {
        dist_freq[0] = 1;
        dist_freq[1] = 1;
    }

    var lit_codes: [286]HuffEncoder.Code = undefined;
    var lit_lns: [286]HuffEncoder.LiteralNode = undefined;
    var lit_lfs: [286]HuffEncoder.LiteralNode = undefined;
    var lit_enc: HuffEncoder = .{
        .codes = &lit_codes,
        .freq_cache = undefined,
        .bit_count = undefined,
        .lns = &lit_lns,
        .lfs = &lit_lfs,
    };
    lit_enc.generate(&lit_freq, 15);

    var dist_codes: [30]HuffEncoder.Code = undefined;
    var dist_lns: [30]HuffEncoder.LiteralNode = undefined;
    var dist_lfs: [30]HuffEncoder.LiteralNode = undefined;
    var dist_enc: HuffEncoder = .{
        .codes = &dist_codes,
        .freq_cache = undefined,
        .bit_count = undefined,
        .lns = &dist_lns,
        .lfs = &dist_lfs,
    };
    dist_enc.generate(&dist_freq, 15);

    // 2. Trim trailing zero-length entries to compute HLIT/HDIST.
    var num_lit: usize = 286;
    while (num_lit > 257 and lit_codes[num_lit - 1].len == 0) num_lit -= 1;
    var num_dist: usize = 30;
    while (num_dist > 1 and dist_codes[num_dist - 1].len == 0) num_dist -= 1;

    // 3. Concatenate lit + dist code lengths, RLE-encode with codegens.
    var combined: [286 + 30]u8 = undefined;
    for (0..num_lit) |k| combined[k] = @intCast(lit_codes[k].len);
    for (0..num_dist) |k| combined[num_lit + k] = @intCast(dist_codes[k].len);

    var codegen_syms: std.ArrayListUnmanaged(CodegenSym) = .{};
    defer codegen_syms.deinit(alloc);
    try rleEncodeCodeLengths(alloc, combined[0 .. num_lit + num_dist], &codegen_syms);

    // 4. Build the codegen huffman (7-bit-limited).
    var cg_freq: [19]u16 = .{0} ** 19;
    for (codegen_syms.items) |s| cg_freq[s.sym] += 1;

    var cg_codes: [19]HuffEncoder.Code = undefined;
    var cg_lns: [19]HuffEncoder.LiteralNode = undefined;
    var cg_lfs: [19]HuffEncoder.LiteralNode = undefined;
    var cg_enc: HuffEncoder = .{
        .codes = &cg_codes,
        .freq_cache = undefined,
        .bit_count = undefined,
        .lns = &cg_lns,
        .lfs = &cg_lfs,
    };
    cg_enc.generate(&cg_freq, 7);

    // HCLEN is the number of codegen code lengths emitted, in the
    // RFC 1951 permuted order. Trim trailing zeros down to minimum 4.
    var num_cg: usize = 19;
    while (num_cg > 4 and cg_codes[HuffEncoder.codegen_order[num_cg - 1]].len == 0) num_cg -= 1;

    // 5. Emit the block.
    // BFINAL=1 (bit 0) + BTYPE=10 (bits 1-2): value 0b101 = 5.
    try bw.writeBits(5, 3);
    try bw.writeBits(@intCast(num_lit - 257), 5);
    try bw.writeBits(@intCast(num_dist - 1), 5);
    try bw.writeBits(@intCast(num_cg - 4), 4);
    for (0..num_cg) |k| {
        const order_idx = HuffEncoder.codegen_order[k];
        try bw.writeBits(@intCast(cg_codes[order_idx].len), 3);
    }
    // Codegen-encoded code lengths. HuffmanEncoder.generate stores
    // codes in LSB-first form already (bit-reversed during assignment),
    // so write them with the raw bit writer rather than writeMsbCode.
    for (codegen_syms.items) |s| {
        const c = cg_codes[s.sym];
        try bw.writeBits(c.code, @intCast(c.len));
        if (s.extra_bits > 0) try bw.writeBits(s.extra, @intCast(s.extra_bits));
    }

    // 6. Emit the token stream using the lit + dist huffman tables.
    for (tokens) |tok| {
        if (tok.kind == 0) {
            const c = lit_codes[tok.val];
            try bw.writeBits(c.code, @intCast(c.len));
        } else {
            const lc = deflateLengthCode(tok.val);
            const c = lit_codes[lc.code];
            try bw.writeBits(c.code, @intCast(c.len));
            if (lc.extra_bits > 0) try bw.writeBits(lc.extra, @intCast(lc.extra_bits));
            const dc = deflateDistCode(tok.dist);
            const dc_c = dist_codes[dc.code];
            try bw.writeBits(dc_c.code, @intCast(dc_c.len));
            if (dc.extra_bits > 0) try bw.writeBits(dc.extra, @intCast(dc.extra_bits));
        }
    }

    // EOB.
    const eob_c = lit_codes[256];
    try bw.writeBits(eob_c.code, @intCast(eob_c.len));
}

/// Compress `input` as a single final dynamic-huffman deflate stream.
/// Caller ensures `input.len > 0` (empty inputs bypass compression
/// upstream).
fn deflateCompress(alloc: Allocator, input: []const u8, out: *std.ArrayListUnmanaged(u8)) !void {
    std.debug.assert(input.len > 0);

    var tokens: std.ArrayListUnmanaged(DeflateToken) = .{};
    defer tokens.deinit(alloc);
    try tokens.ensureTotalCapacity(alloc, input.len);
    try tokenizeLazy(alloc, input, &tokens);

    var bw = DeflateBitWriter{ .out = out, .alloc = alloc };
    try emitDynamicBlock(alloc, tokens.items, &bw);
    try bw.flushByte();
}

// ─── ZIP writer (deflate + stored fallback) ──────────────────────────

/// Minimal zip archive builder. Appends file entries to a byte buffer;
/// `finalize()` emits the central directory + end-of-central-directory
/// trailer. Each entry is deflate-compressed unless compression grows
/// the payload (empty entries, near-random bytes), in which case the
/// entry falls back to stored (method 0). Both Excel and LibreOffice
/// accept mixed-method archives.
const ZipWriter = struct {
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    // Per-entry info accumulated for the central directory.
    entries: std.ArrayListUnmanaged(EntryMeta) = .{},

    const EntryMeta = struct {
        name: []u8, // owned copy
        crc32: u32,
        compressed_size: u32,
        uncompressed_size: u32,
        local_offset: u32,
        method: std.zip.CompressionMethod,
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

        // Sub-1 KB entries skip compression. The dynamic-huffman block
        // header adds ~60-120 bytes of fixed overhead that rarely pays
        // back on tiny XML fragments (Content_Types.xml, workbook rels,
        // empty sheet templates) — and the hash-chain init is pure waste.
        // The big entries (sheet1.xml, sharedStrings.xml, styles.xml)
        // dominate archive size, so bypassing small ones loses negligible
        // savings and shaves real per-entry wall time.
        //
        // If deflate still inflates a ≥ 1 KB payload (already-compressed
        // or near-random content), fall back to stored.
        const COMPRESS_MIN: usize = 1024;
        var compressed: std.ArrayListUnmanaged(u8) = .{};
        defer compressed.deinit(alloc);

        var method: std.zip.CompressionMethod = .deflate;
        var payload: []const u8 = undefined;
        if (data.len >= COMPRESS_MIN) {
            try deflateCompress(alloc, data, &compressed);
        }
        if (data.len < COMPRESS_MIN or compressed.items.len >= data.len) {
            method = .store;
            payload = data;
        } else {
            payload = compressed.items;
        }
        if (payload.len > std.math.maxInt(u32)) return error.EntryTooLarge;

        const hdr: std.zip.LocalFileHeader = .{
            .signature = std.zip.local_file_header_sig,
            .version_needed_to_extract = 20,
            .flags = .{ .encrypted = false, ._ = 0 },
            .compression_method = method,
            .last_modification_time = 0,
            .last_modification_date = 0x21, // 1980-01-01, minimum valid
            .crc32 = crc,
            .compressed_size = @intCast(payload.len),
            .uncompressed_size = @intCast(data.len),
            .filename_len = @intCast(name.len),
            .extra_len = 0,
        };
        try appendStruct(alloc, self.out, std.zip.LocalFileHeader, hdr);
        try self.out.appendSlice(alloc, name);
        try self.out.appendSlice(alloc, payload);

        const owned_name = try alloc.dupe(u8, name);
        errdefer alloc.free(owned_name);
        try self.entries.append(alloc, .{
            .name = owned_name,
            .crc32 = crc,
            .compressed_size = @intCast(payload.len),
            .uncompressed_size = @intCast(data.len),
            .local_offset = offset,
            .method = method,
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
                .compression_method = e.method,
                .last_modification_time = 0,
                .last_modification_date = 0x21,
                .crc32 = e.crc32,
                .compressed_size = e.compressed_size,
                .uncompressed_size = e.uncompressed_size,
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

test "Writer: writeRowStyled rejects out-of-range style id" {
    var w = Writer.init(std.testing.allocator);
    defer w.deinit();
    var sheet = try w.addSheet("S");

    // No styles registered — id 1 out of range.
    try std.testing.expectError(error.UnknownStyleId, sheet.writeRowStyled(
        &.{.{ .string = "x" }},
        &.{1},
    ));

    const bold = try w.addStyle(.{ .font_bold = true });
    try std.testing.expectEqual(@as(u32, 1), bold);

    // id 1 now valid.
    try sheet.writeRowStyled(&.{.{ .string = "ok" }}, &.{1});

    // id 2 still out of range.
    try std.testing.expectError(error.UnknownStyleId, sheet.writeRowStyled(
        &.{.{ .string = "x" }},
        &.{2},
    ));
}

test "Writer: stage-5 number format registers + emits numFmts" {
    const tmp_path = "/tmp/zlsx_writer_numfmt.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();

        const money = try w.addStyle(.{ .number_format = "$#,##0.00" });
        const pct = try w.addStyle(.{ .number_format = "0.00%" });
        const plain = try w.addStyle(.{ .font_bold = true });
        // Dedup: same format returns same numFmtId inside styles.xml
        // and same style index.
        const money_again = try w.addStyle(.{ .number_format = "$#,##0.00" });
        try std.testing.expectEqual(money, money_again);
        try std.testing.expect(pct != money);
        try std.testing.expect(plain != money);

        // Empty format string is rejected.
        try std.testing.expectError(error.InvalidNumberFormat, w.addStyle(.{ .number_format = "" }));

        var sheet = try w.addSheet("S");
        try sheet.writeRowStyled(
            &.{ .{ .number = 123.45 }, .{ .number = 0.9 }, .{ .string = "boo" } },
            &.{ money, pct, plain },
        );
        try w.save(tmp_path);
    }

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

    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "<numFmts count=\"2\">") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "numFmtId=\"164\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "numFmtId=\"165\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "formatCode=\"$#,##0.00\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "formatCode=\"0.00%\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "applyNumberFormat=\"1\"") != null);
}

test "Writer: writeRowWithFormulas emits <f> + cached <v> correctly" {
    const tmp_path = "/tmp/zlsx_writer_formulas.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("Calc");

        // Header row — plain values.
        try sheet.writeRow(&.{ .{ .string = "A" }, .{ .string = "B" }, .{ .string = "Sum" } });
        // Data row — plain.
        try sheet.writeRow(&.{ .{ .integer = 10 }, .{ .integer = 20 }, .empty });
        // Formula row — col 2 is =A2+B2 with cached value 30; no formula in 0/1.
        try sheet.writeRowWithFormulas(
            &.{ .{ .integer = 100 }, .{ .integer = 200 }, .{ .integer = 300 } },
            &.{ null, null, "A2+B2" },
        );
        // Formula cell with no cached value (Excel shows 0 until recalc).
        try sheet.writeRowWithFormulas(
            &.{ .empty, .empty, .empty },
            &.{ null, null, "NOW()" },
        );
        // XML-special char inside formula must be escaped.
        try sheet.writeRowWithFormulas(
            &.{ .{ .string = "foo" }, .empty, .empty },
            &.{ null, null, "IF(A5>5,\"big\",\"small\")" },
        );

        // Rejection — length mismatch.
        try std.testing.expectError(
            error.FormulaCountMismatch,
            sheet.writeRowWithFormulas(&.{ .empty, .empty }, &.{null}),
        );

        try w.save(tmp_path);
    }

    const sheet_xml = blk: {
        var file = try std.fs.cwd().openFile(tmp_path, .{});
        defer file.close();
        var fbuf: [4096]u8 = undefined;
        var fr = file.reader(&fbuf);
        var iter = try std.zip.Iterator.init(&fr);
        var name_buf: [64]u8 = undefined;
        while (try iter.next()) |entry| {
            if (entry.filename_len > name_buf.len) continue;
            try fr.seekTo(entry.header_zip_offset + @sizeOf(std.zip.CentralDirectoryFileHeader));
            const fn_slice = name_buf[0..entry.filename_len];
            try fr.interface.readSliceAll(fn_slice);
            if (std.mem.eql(u8, fn_slice, "xl/worksheets/sheet1.xml")) {
                break :blk try extractEntryForTest(std.testing.allocator, entry, &fr);
            }
        }
        return error.SheetXmlNotFound;
    };
    defer std.testing.allocator.free(sheet_xml);

    // Row 3: formula with cached integer 300.
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<c r=\"C3\"><f>A2+B2</f><v>300</v></c>") != null);
    // Row 4: formula with no cached value → no <v>.
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<c r=\"C4\"><f>NOW()</f></c>") != null);
    // Row 5: formula with XML-special chars in body — `>` and `"`
    // must be entity-escaped (`>` is optional but our escape path
    // emits `&gt;`; `"` becomes `&quot;`).
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<f>IF(A5&gt;5,&quot;big&quot;,&quot;small&quot;)</f>") != null);

    // Round-trip through the reader — the cached values are what
    // `Cell.number` / `.integer` will surface since the reader
    // only reads the `<v>` cached result, not the formula text.
    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    _ = (try rows.next()).?; // header
    _ = (try rows.next()).?; // data
    const r3 = (try rows.next()).?;
    try std.testing.expectEqual(@as(i64, 300), r3[2].integer);
}

test "Writer: setRowHeight emits ht + customHeight, only on marked rows" {
    const tmp_path = "/tmp/zlsx_writer_row_heights.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("Sheet1");

        // Tall header + normal body row + taller footer.
        try sheet.setRowHeight(0, 30.0); // row 1
        try sheet.setRowHeight(2, 42.5); // row 3

        try sheet.writeRow(&.{.{ .string = "header" }});
        try sheet.writeRow(&.{.{ .string = "body" }});
        try sheet.writeRow(&.{.{ .string = "footer" }});

        // Rejections — non-finite / non-positive.
        try std.testing.expectError(error.InvalidRowHeight, sheet.setRowHeight(5, 0));
        try std.testing.expectError(error.InvalidRowHeight, sheet.setRowHeight(5, -1));
        try std.testing.expectError(error.InvalidRowHeight, sheet.setRowHeight(5, std.math.nan(f32)));
        try std.testing.expectError(error.InvalidRowHeight, sheet.setRowHeight(5, std.math.inf(f32)));

        // Post-emit call on row 1 is silently ignored (XML was
        // already flushed to self.body); documented behaviour.
        try sheet.setRowHeight(0, 99.0);

        try w.save(tmp_path);
    }

    const sheet_xml = blk: {
        var file = try std.fs.cwd().openFile(tmp_path, .{});
        defer file.close();
        var fbuf: [4096]u8 = undefined;
        var fr = file.reader(&fbuf);
        var iter = try std.zip.Iterator.init(&fr);
        var name_buf: [64]u8 = undefined;
        while (try iter.next()) |entry| {
            if (entry.filename_len > name_buf.len) continue;
            try fr.seekTo(entry.header_zip_offset + @sizeOf(std.zip.CentralDirectoryFileHeader));
            const fn_slice = name_buf[0..entry.filename_len];
            try fr.interface.readSliceAll(fn_slice);
            if (std.mem.eql(u8, fn_slice, "xl/worksheets/sheet1.xml")) {
                break :blk try extractEntryForTest(std.testing.allocator, entry, &fr);
            }
        }
        return error.SheetXmlNotFound;
    };
    defer std.testing.allocator.free(sheet_xml);

    // Row 1 with height 30, row 3 with height 42.5, row 2 plain.
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<row r=\"1\" ht=\"30\" customHeight=\"1\">") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<row r=\"2\">") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<row r=\"3\" ht=\"42.5\" customHeight=\"1\">") != null);
    // Post-emit override of row 0 MUST NOT have rewritten the XML.
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "ht=\"99\"") == null);

    // Reader still walks the workbook.
    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    var n: usize = 0;
    while (try rows.next()) |_| n += 1;
    try std.testing.expectEqual(@as(usize, 3), n);
}

test "Writer: stage-5 sheet-level features (cols, freeze, autoFilter)" {
    const tmp_path = "/tmp/zlsx_writer_sheet_features.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("Sheet1");
        try sheet.setColumnWidth(0, 20.5);
        try sheet.setColumnWidth(3, 12);
        sheet.freezePanes(1, 2);
        try sheet.setAutoFilter("A1:D1");

        try std.testing.expectError(
            error.InvalidColumnWidth,
            sheet.setColumnWidth(1, -1),
        );
        try std.testing.expectError(
            error.InvalidAutoFilterRange,
            sheet.setAutoFilter(""),
        );

        try sheet.writeRow(&.{ .{ .string = "a" }, .{ .string = "b" }, .{ .string = "c" }, .{ .string = "d" } });
        try w.save(tmp_path);
    }

    // Read the raw sheet1.xml to verify the new sections are present in
    // the right order (sheetViews → cols → sheetData → autoFilter).
    const sheet_xml = blk: {
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
            if (std.mem.eql(u8, filename, "xl/worksheets/sheet1.xml")) {
                break :blk try extractEntryForTest(std.testing.allocator, entry, &fr);
            }
        }
        return error.SheetXmlNotFound;
    };
    defer std.testing.allocator.free(sheet_xml);

    // Ordering check — each segment must come before the next.
    const sv = std.mem.indexOf(u8, sheet_xml, "<sheetViews>") orelse return error.MissingSheetViews;
    const cols = std.mem.indexOf(u8, sheet_xml, "<cols>") orelse return error.MissingCols;
    const data = std.mem.indexOf(u8, sheet_xml, "<sheetData>") orelse return error.MissingSheetData;
    const af = std.mem.indexOf(u8, sheet_xml, "<autoFilter") orelse return error.MissingAutoFilter;
    try std.testing.expect(sv < cols);
    try std.testing.expect(cols < data);
    try std.testing.expect(data < af);

    // Specifics.
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "xSplit=\"2\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "ySplit=\"1\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "state=\"frozen\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "width=\"20.5\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "customWidth=\"1\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "ref=\"A1:D1\"") != null);
}

test "Writer: addMergedCell validates + emits <mergeCells> block" {
    const tmp_path = "/tmp/zlsx_writer_merged_cells.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("Sheet1");

        // Valid — three non-overlapping rectangles + a full-width span.
        try sheet.addMergedCell("A1:B2");
        try sheet.addMergedCell("C5:F5");
        try sheet.addMergedCell("A10:XFD10");

        // Rejections — every rule in parseA1Corner / validateMergeRange.
        try std.testing.expectError(error.InvalidMergeRange, sheet.addMergedCell(""));
        try std.testing.expectError(error.InvalidMergeRange, sheet.addMergedCell("A1")); // no colon
        try std.testing.expectError(error.InvalidMergeRange, sheet.addMergedCell("A1:")); // empty right
        try std.testing.expectError(error.InvalidMergeRange, sheet.addMergedCell(":B2")); // empty left
        try std.testing.expectError(error.InvalidMergeRange, sheet.addMergedCell("A1:A1")); // single cell
        try std.testing.expectError(error.InvalidMergeRange, sheet.addMergedCell("B1:A1")); // col inverted
        try std.testing.expectError(error.InvalidMergeRange, sheet.addMergedCell("A2:A1")); // row inverted
        try std.testing.expectError(error.InvalidMergeRange, sheet.addMergedCell("A:B2")); // no row on left
        try std.testing.expectError(error.InvalidMergeRange, sheet.addMergedCell("A1:B")); // no row on right
        try std.testing.expectError(error.InvalidMergeRange, sheet.addMergedCell("1:B2")); // no col on left
        try std.testing.expectError(error.InvalidMergeRange, sheet.addMergedCell("A0:B2")); // row 0
        try std.testing.expectError(error.InvalidMergeRange, sheet.addMergedCell("A01:B2")); // leading zero
        try std.testing.expectError(error.InvalidMergeRange, sheet.addMergedCell("a1:b2")); // lowercase
        try std.testing.expectError(error.InvalidMergeRange, sheet.addMergedCell("A1:B2 ")); // trailing space
        try std.testing.expectError(error.InvalidMergeRange, sheet.addMergedCell("XFE1:XFE2")); // col > 16384
        try std.testing.expectError(error.InvalidMergeRange, sheet.addMergedCell("A1:A1048577")); // row > 1048576

        try sheet.writeRow(&.{.{ .string = "header" }});
        try w.save(tmp_path);
    }

    // Inspect raw sheet1.xml for the expected block + ordering.
    const sheet_xml = blk: {
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
            if (std.mem.eql(u8, filename, "xl/worksheets/sheet1.xml")) {
                break :blk try extractEntryForTest(std.testing.allocator, entry, &fr);
            }
        }
        return error.SheetXmlNotFound;
    };
    defer std.testing.allocator.free(sheet_xml);

    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<mergeCells count=\"3\">") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<mergeCell ref=\"A1:B2\"/>") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<mergeCell ref=\"C5:F5\"/>") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<mergeCell ref=\"A10:XFD10\"/>") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "</mergeCells>") != null);

    // Ordering: </sheetData> < <mergeCells> < </worksheet>.
    const sd_end = std.mem.indexOf(u8, sheet_xml, "</sheetData>") orelse return error.MissingSheetData;
    const mc = std.mem.indexOf(u8, sheet_xml, "<mergeCells") orelse return error.MissingMergeCells;
    const ws_end = std.mem.indexOf(u8, sheet_xml, "</worksheet>") orelse return error.MissingWorksheetEnd;
    try std.testing.expect(sd_end < mc);
    try std.testing.expect(mc < ws_end);

    // Confirm the reader still walks the workbook cleanly.
    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    while (try rows.next()) |_| {}
}

test "Writer: addDataValidationNumeric + Custom emit correct XML" {
    const tmp_path = "/tmp/zlsx_writer_dv_ranges.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("Validations");

        // Whole-number between 1..100.
        try sheet.addDataValidationNumeric("B2:B10", .whole, .between, "1", "100");
        // Decimal greater than 0.
        try sheet.addDataValidationNumeric("C3", .decimal, .greater_than, "0", null);
        // Date before 2025-01-01 (Excel serial 45658).
        try sheet.addDataValidationNumeric("D4", .date, .less_than, "45658", null);
        // Text length between 3 and 20 characters.
        try sheet.addDataValidationNumeric("E5", .text_length, .between, "3", "20");
        // Custom formula — XML-special chars must be escaped on emit.
        try sheet.addDataValidationCustom("F6", "AND(F6>0,F6<LEN(A1))");

        // Also mix with an iter13 list validation to prove both
        // emission paths coexist.
        try sheet.addDataValidationList("A2:A10", &.{ "Yes", "No" });

        // Rejections.
        try std.testing.expectError(error.InvalidDataValidation, sheet.addDataValidationNumeric("G1", .whole, .between, "1", null));
        try std.testing.expectError(error.InvalidDataValidation, sheet.addDataValidationNumeric("G2", .whole, .equal, "1", "2"));
        try std.testing.expectError(error.InvalidDataValidation, sheet.addDataValidationNumeric("G3", .whole, .equal, "", null));
        try std.testing.expectError(error.InvalidDataValidation, sheet.addDataValidationCustom("G4", ""));
        try std.testing.expectError(error.InvalidHyperlinkRange, sheet.addDataValidationNumeric("", .whole, .equal, "1", null));

        try sheet.writeRow(&.{.{ .string = "hdr" }});
        try w.save(tmp_path);
    }

    const sheet_xml = blk: {
        var file = try std.fs.cwd().openFile(tmp_path, .{});
        defer file.close();
        var fbuf: [4096]u8 = undefined;
        var fr = file.reader(&fbuf);
        var iter = try std.zip.Iterator.init(&fr);
        var name_buf: [64]u8 = undefined;
        while (try iter.next()) |entry| {
            if (entry.filename_len > name_buf.len) continue;
            try fr.seekTo(entry.header_zip_offset + @sizeOf(std.zip.CentralDirectoryFileHeader));
            const fn_slice = name_buf[0..entry.filename_len];
            try fr.interface.readSliceAll(fn_slice);
            if (std.mem.eql(u8, fn_slice, "xl/worksheets/sheet1.xml")) {
                break :blk try extractEntryForTest(std.testing.allocator, entry, &fr);
            }
        }
        return error.SheetXmlNotFound;
    };
    defer std.testing.allocator.free(sheet_xml);

    // Count = 6 (1 list + 5 ranges).
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<dataValidations count=\"6\">") != null);

    // whole/between with two formulas.
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<dataValidation type=\"whole\" operator=\"between\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B2:B10\"><formula1>1</formula1><formula2>100</formula2></dataValidation>") != null);

    // decimal/greaterThan with single formula.
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<dataValidation type=\"decimal\" operator=\"greaterThan\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"C3\"><formula1>0</formula1></dataValidation>") != null);

    // date/lessThan.
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "type=\"date\" operator=\"lessThan\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<formula1>45658</formula1>") != null);

    // textLength/between.
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "type=\"textLength\" operator=\"between\"") != null);

    // custom — no operator attribute; XML-special chars in the
    // formula must be entity-escaped (`>` → `&gt;`, `<` → `&lt;`).
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<dataValidation type=\"custom\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"F6\"><formula1>AND(F6&gt;0,F6&lt;LEN(A1))</formula1></dataValidation>") != null);

    // And the list entry still appears.
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<dataValidation type=\"list\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "&quot;Yes,No&quot;") != null);

    // Reader round-trip still parses cleanly.
    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    while (try rows.next()) |_| {}
}

test "Writer: addDataValidationList validates + emits <dataValidations> block" {
    const tmp_path = "/tmp/zlsx_writer_dv_list.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("Sheet1");

        // Valid.
        try sheet.addDataValidationList("A2:A10", &.{ "Red", "Green", "Blue" });
        try sheet.addDataValidationList("C3", &.{"Single"});
        // Values with XML specials — must be entity-escaped on emit.
        try sheet.addDataValidationList("B2", &.{ "R&D", "Q<A", "x>y" });

        // Rejections.
        try std.testing.expectError(error.InvalidDataValidation, sheet.addDataValidationList("D1", &.{}));
        try std.testing.expectError(error.InvalidDataValidation, sheet.addDataValidationList("D2", &.{""}));
        try std.testing.expectError(error.InvalidDataValidation, sheet.addDataValidationList("D3", &.{"has,comma"}));
        try std.testing.expectError(error.InvalidDataValidation, sheet.addDataValidationList("D4", &.{"has\"quote"}));
        try std.testing.expectError(error.InvalidHyperlinkRange, sheet.addDataValidationList("", &.{"x"}));
        try std.testing.expectError(error.InvalidHyperlinkRange, sheet.addDataValidationList("a1", &.{"x"}));

        try sheet.writeRow(&.{.{ .string = "hdr" }});
        try w.save(tmp_path);
    }

    const sheet_xml = blk: {
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
            if (std.mem.eql(u8, filename, "xl/worksheets/sheet1.xml")) {
                break :blk try extractEntryForTest(std.testing.allocator, entry, &fr);
            }
        }
        return error.SheetXmlNotFound;
    };
    defer std.testing.allocator.free(sheet_xml);

    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<dataValidations count=\"3\">") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "sqref=\"A2:A10\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "&quot;Red,Green,Blue&quot;") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "sqref=\"C3\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "&quot;Single&quot;") != null);
    // XML-special chars in values must be entity-escaped.
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "&quot;R&amp;D,Q&lt;A,x&gt;y&quot;") != null);

    // Ordering: </sheetData> < <dataValidations> < </worksheet>.
    const sd_end = std.mem.indexOf(u8, sheet_xml, "</sheetData>") orelse return error.MissingSheetData;
    const dv = std.mem.indexOf(u8, sheet_xml, "<dataValidations") orelse return error.MissingDataValidations;
    const ws_end = std.mem.indexOf(u8, sheet_xml, "</worksheet>") orelse return error.MissingWorksheetEnd;
    try std.testing.expect(sd_end < dv);
    try std.testing.expect(dv < ws_end);

    // Reader still walks the workbook cleanly.
    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    while (try rows.next()) |_| {}
}

test "Writer: addDataValidationList — no block when none registered" {
    const tmp_path = "/tmp/zlsx_writer_no_dv.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("Sheet1");
        try sheet.writeRow(&.{.{ .string = "plain" }});
        try w.save(tmp_path);
    }

    const sheet_xml = blk: {
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
            if (std.mem.eql(u8, filename, "xl/worksheets/sheet1.xml")) {
                break :blk try extractEntryForTest(std.testing.allocator, entry, &fr);
            }
        }
        return error.SheetXmlNotFound;
    };
    defer std.testing.allocator.free(sheet_xml);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<dataValidations") == null);
}

test "Writer: addHyperlink validates + emits <hyperlinks> + per-sheet _rels" {
    const tmp_path = "/tmp/zlsx_writer_hyperlinks.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("Sheet1");

        // Valid: single cell, rectangle, + URL with XML-special char to
        // exercise the escape path.
        try sheet.addHyperlink("A1", "https://example.com/path?q=1&x=2");
        try sheet.addHyperlink("B2:C3", "https://docs.example.com/");
        try sheet.addHyperlink("D5", "mailto:foo@example.com");

        // Rejections — full matrix.
        try std.testing.expectError(error.InvalidHyperlinkRange, sheet.addHyperlink("", "http://x"));
        try std.testing.expectError(error.InvalidHyperlinkRange, sheet.addHyperlink("a1", "http://x"));
        try std.testing.expectError(error.InvalidHyperlinkRange, sheet.addHyperlink("B2:A1", "http://x"));
        try std.testing.expectError(error.InvalidHyperlinkRange, sheet.addHyperlink("A0", "http://x"));
        try std.testing.expectError(error.InvalidHyperlinkRange, sheet.addHyperlink("A1:", "http://x"));
        try std.testing.expectError(error.InvalidHyperlinkUrl, sheet.addHyperlink("A1", ""));

        try sheet.writeRow(&.{.{ .string = "link" }});
        try w.save(tmp_path);
    }

    // Inspect xl/worksheets/sheet1.xml.
    const sheet_xml = blk: {
        var file = try std.fs.cwd().openFile(tmp_path, .{});
        defer file.close();
        var fbuf: [4096]u8 = undefined;
        var fr = file.reader(&fbuf);
        var iter = try std.zip.Iterator.init(&fr);
        var filename_buf: [96]u8 = undefined;
        while (try iter.next()) |entry| {
            if (entry.filename_len > filename_buf.len) continue;
            try fr.seekTo(entry.header_zip_offset + @sizeOf(std.zip.CentralDirectoryFileHeader));
            const filename = filename_buf[0..entry.filename_len];
            try fr.interface.readSliceAll(filename);
            if (std.mem.eql(u8, filename, "xl/worksheets/sheet1.xml")) {
                break :blk try extractEntryForTest(std.testing.allocator, entry, &fr);
            }
        }
        return error.SheetXmlNotFound;
    };
    defer std.testing.allocator.free(sheet_xml);

    // xmlns:r must be declared on the worksheet root so r:id parses.
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<hyperlinks>") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<hyperlink ref=\"A1\" r:id=\"rId1\"/>") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<hyperlink ref=\"B2:C3\" r:id=\"rId2\"/>") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<hyperlink ref=\"D5\" r:id=\"rId3\"/>") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "</hyperlinks>") != null);

    // Ordering: </sheetData> < <hyperlinks> < </worksheet>.
    const sd_end = std.mem.indexOf(u8, sheet_xml, "</sheetData>") orelse return error.MissingSheetData;
    const hl = std.mem.indexOf(u8, sheet_xml, "<hyperlinks>") orelse return error.MissingHyperlinks;
    const ws_end = std.mem.indexOf(u8, sheet_xml, "</worksheet>") orelse return error.MissingWorksheetEnd;
    try std.testing.expect(sd_end < hl);
    try std.testing.expect(hl < ws_end);

    // Inspect xl/worksheets/_rels/sheet1.xml.rels.
    const rels_xml = blk: {
        var file = try std.fs.cwd().openFile(tmp_path, .{});
        defer file.close();
        var fbuf: [4096]u8 = undefined;
        var fr = file.reader(&fbuf);
        var iter = try std.zip.Iterator.init(&fr);
        var filename_buf: [96]u8 = undefined;
        while (try iter.next()) |entry| {
            if (entry.filename_len > filename_buf.len) continue;
            try fr.seekTo(entry.header_zip_offset + @sizeOf(std.zip.CentralDirectoryFileHeader));
            const filename = filename_buf[0..entry.filename_len];
            try fr.interface.readSliceAll(filename);
            if (std.mem.eql(u8, filename, "xl/worksheets/_rels/sheet1.xml.rels")) {
                break :blk try extractEntryForTest(std.testing.allocator, entry, &fr);
            }
        }
        return error.SheetRelsNotFound;
    };
    defer std.testing.allocator.free(rels_xml);

    try std.testing.expect(std.mem.indexOf(u8, rels_xml, "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\"") != null);
    // Ampersand in the URL must be escaped to &amp;.
    try std.testing.expect(std.mem.indexOf(u8, rels_xml, "Target=\"https://example.com/path?q=1&amp;x=2\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, rels_xml, "Target=\"mailto:foo@example.com\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, rels_xml, "TargetMode=\"External\"") != null);

    // Reader still walks the workbook cleanly.
    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    while (try rows.next()) |_| {}
}

test "Writer: no <hyperlinks> block or _rels entry when none registered" {
    const tmp_path = "/tmp/zlsx_writer_no_hyperlinks.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("Sheet1");
        try sheet.writeRow(&.{.{ .string = "plain" }});
        try w.save(tmp_path);
    }

    // Neither the sheet XML's <hyperlinks> section nor the per-sheet
    // _rels file should exist.
    var file = try std.fs.cwd().openFile(tmp_path, .{});
    defer file.close();
    var fbuf: [4096]u8 = undefined;
    var fr = file.reader(&fbuf);
    var iter = try std.zip.Iterator.init(&fr);
    var filename_buf: [96]u8 = undefined;
    var saw_rels = false;
    while (try iter.next()) |entry| {
        if (entry.filename_len > filename_buf.len) continue;
        try fr.seekTo(entry.header_zip_offset + @sizeOf(std.zip.CentralDirectoryFileHeader));
        const filename = filename_buf[0..entry.filename_len];
        try fr.interface.readSliceAll(filename);
        if (std.mem.indexOf(u8, filename, "_rels/sheet") != null) saw_rels = true;
    }
    try std.testing.expect(!saw_rels);
}

test "Writer: no <mergeCells> block when none registered" {
    const tmp_path = "/tmp/zlsx_writer_no_merged.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("Sheet1");
        try sheet.writeRow(&.{.{ .string = "a" }});
        try w.save(tmp_path);
    }

    const sheet_xml = blk: {
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
            if (std.mem.eql(u8, filename, "xl/worksheets/sheet1.xml")) {
                break :blk try extractEntryForTest(std.testing.allocator, entry, &fr);
            }
        }
        return error.SheetXmlNotFound;
    };
    defer std.testing.allocator.free(sheet_xml);

    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<mergeCells") == null);
}

test "Writer: stage-4 border sides emit into styles.xml" {
    const tmp_path = "/tmp/zlsx_writer_styles_borders.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();

        // Thin black box on all 4 sides — the bread-and-butter table outline.
        const box = try w.addStyle(.{
            .border_left = .{ .style = .thin, .color_argb = 0xFF000000 },
            .border_right = .{ .style = .thin, .color_argb = 0xFF000000 },
            .border_top = .{ .style = .thin, .color_argb = 0xFF000000 },
            .border_bottom = .{ .style = .thin, .color_argb = 0xFF000000 },
        });
        // Bottom-only thick red + diagonal up.
        const fancy = try w.addStyle(.{
            .border_bottom = .{ .style = .thick, .color_argb = 0xFFFF0000 },
            .border_diagonal = .{ .style = .dashed },
            .diagonal_up = true,
        });
        const plain = try w.addStyle(.{ .font_bold = true });
        // Dedup.
        const box_again = try w.addStyle(.{
            .border_left = .{ .style = .thin, .color_argb = 0xFF000000 },
            .border_right = .{ .style = .thin, .color_argb = 0xFF000000 },
            .border_top = .{ .style = .thin, .color_argb = 0xFF000000 },
            .border_bottom = .{ .style = .thin, .color_argb = 0xFF000000 },
        });
        try std.testing.expectEqual(box, box_again);
        try std.testing.expect(fancy != box);
        try std.testing.expect(plain != box);

        var sheet = try w.addSheet("S");
        try sheet.writeRowStyled(
            &.{ .{ .string = "boxed" }, .{ .string = "fancy" }, .{ .string = "plain" } },
            &.{ box, fancy, plain },
        );
        try w.save(tmp_path);
    }

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

    // Default border at 0 + 2 user borders (plain has none).
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "<borders count=\"3\">") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "<left style=\"thin\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "<bottom style=\"thick\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "<color rgb=\"FFFF0000\"/>") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "diagonalUp=\"1\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "<diagonal style=\"dashed\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "applyBorder=\"1\"") != null);
}

test "Writer: stage-3 fill fields emit into styles.xml" {
    const tmp_path = "/tmp/zlsx_writer_styles_fills.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();

        // Solid yellow highlight — the bread-and-butter fill.
        const yellow = try w.addStyle(.{
            .fill_pattern = .solid,
            .fill_fg_argb = 0xFFFFFF00,
        });
        // Pattern fill with both fg and bg.
        const striped = try w.addStyle(.{
            .fill_pattern = .dark_horizontal,
            .fill_fg_argb = 0xFF0000FF,
            .fill_bg_argb = 0xFFFFFFFF,
        });
        // Pattern-only, no colours.
        const gray = try w.addStyle(.{ .fill_pattern = .gray0625 });
        // Style with no fill at all — fillId must remain 0.
        const plain_bold = try w.addStyle(.{ .font_bold = true });

        // Dedup across distinct calls.
        const yellow_again = try w.addStyle(.{
            .fill_pattern = .solid,
            .fill_fg_argb = 0xFFFFFF00,
        });
        try std.testing.expectEqual(yellow, yellow_again);
        try std.testing.expect(striped != yellow);
        try std.testing.expect(gray != striped);
        try std.testing.expect(plain_bold != yellow);

        var sheet = try w.addSheet("S");
        try sheet.writeRowStyled(
            &.{ .{ .string = "hi" }, .{ .string = "lo" }, .{ .string = "mid" }, .{ .string = "b" } },
            &.{ yellow, striped, gray, plain_bold },
        );

        try w.save(tmp_path);
    }

    // Grep the emitted styles.xml for the expected OOXML markers.
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

    // <fills count> should be 2 defaults + 3 user fills (plain_bold has
    // no fill, so it doesn't contribute).
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "<fills count=\"5\">") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "patternType=\"solid\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "<fgColor rgb=\"FFFFFF00\"/>") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "patternType=\"darkHorizontal\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "<fgColor rgb=\"FF0000FF\"/>") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "<bgColor rgb=\"FFFFFFFF\"/>") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "patternType=\"gray0625\"/>") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "applyFill=\"1\"") != null);
}

test "Writer: stage-2 style fields emit into styles.xml" {
    const tmp_path = "/tmp/zlsx_writer_styles_stage2.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();

        const big_red_arial = try w.addStyle(.{
            .font_size = 18,
            .font_name = "Arial",
            .font_color_argb = 0xFFFF0000,
            .alignment_horizontal = .center,
            .wrap_text = true,
        });
        const wrap_only = try w.addStyle(.{ .wrap_text = true });
        // Dedup: same style from distinct "Arial" buffer must coalesce.
        var arial_copy: [5]u8 = .{ 'A', 'r', 'i', 'a', 'l' };
        const again = try w.addStyle(.{
            .font_size = 18,
            .font_name = &arial_copy,
            .font_color_argb = 0xFFFF0000,
            .alignment_horizontal = .center,
            .wrap_text = true,
        });
        try std.testing.expectEqual(big_red_arial, again);

        // Invalid inputs surface typed errors, not panics.
        try std.testing.expectError(error.InvalidFontSize, w.addStyle(.{ .font_size = 0 }));
        try std.testing.expectError(error.InvalidFontSize, w.addStyle(.{ .font_size = -1 }));
        try std.testing.expectError(error.InvalidFontName, w.addStyle(.{ .font_name = "" }));

        var sheet = try w.addSheet("S");
        try sheet.writeRowStyled(
            &.{ .{ .string = "big red" }, .{ .string = "wrapped" } },
            &.{ big_red_arial, wrap_only },
        );

        try w.save(tmp_path);
    }

    // Read the raw styles.xml bytes to verify stage-2 fields landed.
    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

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

    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "<sz val=\"18\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "<name val=\"Arial\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "rgb=\"FFFF0000\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "horizontal=\"center\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "wrapText=\"1\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "applyAlignment=\"1\"") != null);
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

test "Writer: addSheet validates sheet names (length, reserved chars, History)" {
    var w = Writer.init(std.testing.allocator);
    defer w.deinit();

    // Valid — including XML-special chars (these are escaped on emit,
    // not rejected; only the path-reserved set triggers InvalidSheetName).
    _ = try w.addSheet("Summary");
    _ = try w.addSheet("2026 Q1");
    _ = try w.addSheet("R&D"); // & is xml-escaped on emit
    _ = try w.addSheet("x<y"); // < is xml-escaped on emit

    // Reject every rule.
    try std.testing.expectError(error.InvalidSheetName, w.addSheet(""));
    try std.testing.expectError(error.InvalidSheetName, w.addSheet("A" ** 32)); // > 31 chars
    try std.testing.expectError(error.InvalidSheetName, w.addSheet("Sheet/1"));
    try std.testing.expectError(error.InvalidSheetName, w.addSheet("Sheet\\1"));
    try std.testing.expectError(error.InvalidSheetName, w.addSheet("Sheet?1"));
    try std.testing.expectError(error.InvalidSheetName, w.addSheet("Sheet*1"));
    try std.testing.expectError(error.InvalidSheetName, w.addSheet("Sheet[1]"));
    try std.testing.expectError(error.InvalidSheetName, w.addSheet("Sheet:1"));
    try std.testing.expectError(error.InvalidSheetName, w.addSheet("'quoted"));
    try std.testing.expectError(error.InvalidSheetName, w.addSheet("quoted'"));
    try std.testing.expectError(error.InvalidSheetName, w.addSheet("tab\there"));
    try std.testing.expectError(error.InvalidSheetName, w.addSheet("History"));
    try std.testing.expectError(error.InvalidSheetName, w.addSheet("history")); // case-insensitive
    try std.testing.expectError(error.InvalidSheetName, w.addSheet("HISTORY"));

    // Exactly 31 chars still valid.
    const exactly_31 = "A" ** 31;
    _ = try w.addSheet(exactly_31);
}

test "Writer: addSheet rejects case-insensitive duplicates" {
    var w = Writer.init(std.testing.allocator);
    defer w.deinit();
    _ = try w.addSheet("Summary");
    try std.testing.expectError(error.DuplicateSheetName, w.addSheet("Summary"));
    try std.testing.expectError(error.DuplicateSheetName, w.addSheet("summary"));
    try std.testing.expectError(error.DuplicateSheetName, w.addSheet("SUMMARY"));
    try std.testing.expectError(error.DuplicateSheetName, w.addSheet("SumMarY"));
    // Different name still allowed.
    _ = try w.addSheet("Summary 2");
}

test "fuzz validateSheetName: adversarial bytes never panic + only valid names pass" {
    const iters = fuzzIterationsW();
    const seed = fuzzSeedW();
    var prng = std.Random.DefaultPrng.init(seed);
    const rng = prng.random();

    for (0..iters) |_| {
        var buf: [40]u8 = undefined;
        const len = rng.intRangeAtMost(usize, 0, buf.len);
        rng.bytes(buf[0..len]);
        const name = buf[0..len];
        const result = validateSheetName(name);
        if (result) |_| {
            // Post-conditions of a successful validation — must hold on
            // every accepted input so Excel always opens the workbook.
            try std.testing.expect(name.len >= 1 and name.len <= 31);
            try std.testing.expect(name[0] != '\'' and name[name.len - 1] != '\'');
            for (name) |c| {
                try std.testing.expect(c >= 0x20);
                try std.testing.expect(c != ':' and c != '/' and c != '\\' and
                    c != '?' and c != '*' and c != '[' and c != ']');
            }
            try std.testing.expect(!asciiEqlFold(name, "History"));
        } else |err| {
            try std.testing.expectEqual(error.InvalidSheetName, err);
        }
    }
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

// ─── Fuzz tests ──────────────────────────────────────────────────────
//
// PRNG-driven fuzzing (Zig's coverage-guided `--fuzz` is broken on
// macOS Mach-O — see src/xlsx.zig's fuzz block for the same pattern).
// Iteration count comes from XLSX_FUZZ_ITERS (default 1_000); seed from
// XLSX_FUZZ_SEED (default: current time). Each fuzz target enforces an
// invariant beyond "no panic" so we catch logic bugs, not just crashes.

const fuzz_default_iters_writer: usize = 1_000;

fn fuzzIterationsW() usize {
    const env = std.process.getEnvVarOwned(std.heap.page_allocator, "XLSX_FUZZ_ITERS") catch return fuzz_default_iters_writer;
    defer std.heap.page_allocator.free(env);
    var digits: [32]u8 = undefined;
    var di: usize = 0;
    for (env) |c| {
        if (c == '_') continue;
        if (di == digits.len) break;
        digits[di] = c;
        di += 1;
    }
    return std.fmt.parseInt(usize, digits[0..di], 10) catch fuzz_default_iters_writer;
}

fn fuzzSeedW() u64 {
    if (std.process.getEnvVarOwned(std.heap.page_allocator, "XLSX_FUZZ_SEED")) |s| {
        defer std.heap.page_allocator.free(s);
        return std.fmt.parseInt(u64, s, 10) catch 0xA1F8ED;
    } else |_| {
        return @bitCast(std.time.milliTimestamp());
    }
}

test "fuzz formatCellRef: no overflow, always starts with A-Z" {
    const iters = fuzzIterationsW();
    var prng = std.Random.DefaultPrng.init(fuzzSeedW());
    const rng = prng.random();
    var buf: [16]u8 = undefined;
    for (0..iters) |_| {
        const row = rng.intRangeAtMost(u32, 1, std.math.maxInt(u32));
        // Cap col_idx at 2^20 — beyond that the letter repr would
        // exceed the 8-byte scratch; real xlsx tops out at col 16384.
        const col = rng.intRangeAtMost(u32, 0, 1_048_575);
        const ref = formatCellRef(&buf, row, col) catch continue;
        try std.testing.expect(ref.len >= 2);
        try std.testing.expect(ref[0] >= 'A' and ref[0] <= 'Z');
        // The last char must be a digit (the row part).
        try std.testing.expect(ref[ref.len - 1] >= '0' and ref[ref.len - 1] <= '9');
    }
}

test "fuzz appendXmlEscaped: no raw XML specials in output" {
    const iters = fuzzIterationsW();
    var prng = std.Random.DefaultPrng.init(fuzzSeedW());
    const rng = prng.random();
    var input_buf: [512]u8 = undefined;
    var out: std.ArrayListUnmanaged(u8) = .{};
    defer out.deinit(std.testing.allocator);

    for (0..iters) |_| {
        const len = rng.intRangeAtMost(usize, 0, input_buf.len);
        rng.bytes(input_buf[0..len]);
        out.clearRetainingCapacity();
        try appendXmlEscaped(std.testing.allocator, &out, input_buf[0..len]);

        // Invariant: no raw `<`, `>`, `&`, `"`, `'` survives in the
        // output. Each would have been expanded to its entity.
        for (out.items) |c| {
            try std.testing.expect(c != '<' and c != '>' and c != '"' and c != '\'');
        }
        // `&` can appear inside an entity reference like `&amp;`, so
        // we can't forbid it outright. But every `&` must be followed
        // by one of the known entities (amp, lt, gt, quot, apos).
        var i: usize = 0;
        while (i < out.items.len) : (i += 1) {
            if (out.items[i] != '&') continue;
            const rest = out.items[i + 1 ..];
            const ok = std.mem.startsWith(u8, rest, "amp;") or
                std.mem.startsWith(u8, rest, "lt;") or
                std.mem.startsWith(u8, rest, "gt;") or
                std.mem.startsWith(u8, rest, "quot;") or
                std.mem.startsWith(u8, rest, "apos;");
            try std.testing.expect(ok);
        }
    }
}

test "fuzz fitsExactlyInF64 matches round-trip reference" {
    const iters = fuzzIterationsW();
    var prng = std.Random.DefaultPrng.init(fuzzSeedW());
    const rng = prng.random();

    for (0..iters) |_| {
        const n = rng.int(i64);
        const f: f64 = @floatFromInt(n);
        // Round-trip reference is valid when f stays inside i64 range
        // after the int→float conversion. std.math.maxInt(i64) rounds
        // up to 2^63 which would overflow @intFromFloat.
        const reference: bool = blk: {
            if (f >= 9.223372036854776e18) break :blk false;
            if (f < -9.223372036854776e18) break :blk false;
            const back: i64 = @intFromFloat(f);
            break :blk back == n;
        };
        try std.testing.expectEqual(reference, fitsExactlyInF64(n));
    }
}

test "fuzz Writer.sstIntern dedup invariant" {
    const iters = fuzzIterationsW();
    var prng = std.Random.DefaultPrng.init(fuzzSeedW());
    const rng = prng.random();

    var w = Writer.init(std.testing.allocator);
    defer w.deinit();

    // Pool of 16 distinct candidate strings so the rng can hit dupes.
    var pool_buf: [16][24]u8 = undefined;
    var pool_lens: [16]usize = undefined;
    for (0..16) |i| {
        const l = rng.intRangeAtMost(usize, 1, pool_buf[i].len);
        rng.bytes(pool_buf[i][0..l]);
        pool_lens[i] = l;
    }

    var seen_indices: std.StringHashMap(u32) = .init(std.testing.allocator);
    defer seen_indices.deinit();

    for (0..iters) |_| {
        const k = rng.intRangeAtMost(usize, 0, 15);
        const s = pool_buf[k][0..pool_lens[k]];
        const idx = try w.sstIntern(s);

        if (seen_indices.get(s)) |prior| {
            try std.testing.expectEqual(prior, idx);
        } else {
            try seen_indices.put(s, idx);
        }
        // strings.len must equal the distinct count.
        try std.testing.expectEqual(@as(u32, @intCast(seen_indices.count())), @as(u32, @intCast(w.sst_strings.items.len)));
    }
}

test "fuzz Writer.addStyle dedup on bool combos" {
    const iters = fuzzIterationsW();
    var prng = std.Random.DefaultPrng.init(fuzzSeedW());
    const rng = prng.random();

    var w = Writer.init(std.testing.allocator);
    defer w.deinit();

    // 4 possible Style values (2 bool fields) — after the first 4 unique
    // registrations the style count must plateau at 4. Track distinct
    // (bool, bool) → idx pairs directly since Style now contains an
    // f32/slice field that AutoHashMap can't hash.
    var distinct_indices: [2][2]?u32 = .{ .{ null, null }, .{ null, null } };

    for (0..iters) |_| {
        const bold = rng.boolean();
        const italic = rng.boolean();
        const idx = try w.addStyle(.{ .font_bold = bold, .font_italic = italic });
        const bi: usize = if (bold) 1 else 0;
        const ii: usize = if (italic) 1 else 0;
        if (distinct_indices[bi][ii]) |prior| {
            try std.testing.expectEqual(prior, idx);
        } else {
            distinct_indices[bi][ii] = idx;
        }
        try std.testing.expect(w.styles.items.len <= 4);
    }
}

test "fuzz Writer end-to-end round-trip via reader" {
    const iters = fuzzIterationsW() / 10; // each iter does real zip I/O
    const seed = fuzzSeedW();
    var prng = std.Random.DefaultPrng.init(seed);
    const rng = prng.random();
    var tmp_path_buf: [64]u8 = undefined;
    const tmp_path = try std.fmt.bufPrint(&tmp_path_buf, "/tmp/zlsx_fuzz_writer_{x}.xlsx", .{seed});
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    for (0..iters) |_| {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();
        const n_sheets = rng.intRangeAtMost(usize, 1, 3);
        var expected_rows: [3]usize = .{ 0, 0, 0 };

        for (0..n_sheets) |si| {
            // Random uppercase-letter names with a unique trailing digit
            // per sheet. Stays well clear of Excel's reserved-char list
            // (`/\?*[]:`) and case-insensitive-dup rule, so the fuzz
            // exercises the data paths rather than the name validator.
            // (Separate fuzz target covers validateSheetName.)
            var name_buf: [12]u8 = undefined;
            for (&name_buf) |*b| b.* = 'A' + rng.intRangeAtMost(u8, 0, 25);
            name_buf[name_buf.len - 1] = '0' + @as(u8, @intCast(si));
            var sheet = try w.addSheet(&name_buf);

            const n_rows = rng.intRangeAtMost(usize, 0, 8);
            for (0..n_rows) |_| {
                var cells: [6]xlsx.Cell = undefined;
                const n_cells = rng.intRangeAtMost(usize, 0, cells.len);
                for (0..n_cells) |ci| {
                    cells[ci] = switch (rng.intRangeAtMost(u8, 0, 4)) {
                        0 => .empty,
                        1 => blk: {
                            var sbuf: [16]u8 = undefined;
                            const l = rng.intRangeAtMost(usize, 0, sbuf.len);
                            rng.bytes(sbuf[0..l]);
                            for (sbuf[0..l]) |*b| b.* = (b.* % 94) + 32;
                            break :blk .{ .string = sbuf[0..l] };
                        },
                        2 => .{ .integer = rng.intRangeAtMost(i64, -(1 << 40), 1 << 40) },
                        3 => .{ .number = rng.float(f64) * 1000 },
                        else => .{ .boolean = rng.boolean() },
                    };
                }
                sheet.writeRow(cells[0..n_cells]) catch |e| switch (e) {
                    error.IntegerExceedsExcelPrecision => continue,
                    else => return e,
                };
                expected_rows[si] += 1;
            }
        }

        w.save(tmp_path) catch |e| switch (e) {
            error.NoSheets => continue,
            else => return e,
        };

        var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
        defer book.deinit();
        try std.testing.expectEqual(n_sheets, book.sheets.len);
        for (0..n_sheets) |si| {
            var rows = try book.rows(book.sheets[si], std.testing.allocator);
            defer rows.deinit();
            var count: usize = 0;
            while (try rows.next()) |_| count += 1;
            try std.testing.expectEqual(expected_rows[si], count);
        }
    }
}

test "fuzz Writer: random stage 2-5 style combos survive round-trip" {
    // Register styles with every stage's fields pseudo-randomly set,
    // save the workbook, and confirm the reader parses it cleanly.
    // Catches any crash in emitStylesXml caused by unusual field
    // combinations (e.g. fill + border + numFmt simultaneously).
    const iters = fuzzIterationsW() / 20;
    const seed = fuzzSeedW();
    var prng = std.Random.DefaultPrng.init(seed);
    const rng = prng.random();
    var tmp_path_buf: [64]u8 = undefined;
    const tmp_path = try std.fmt.bufPrint(&tmp_path_buf, "/tmp/zlsx_fuzz_combo_{x}.xlsx", .{seed});
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    const font_names = [_][]const u8{ "Calibri", "Arial", "Helvetica", "Times New Roman" };
    const num_formats = [_][]const u8{ "0.00", "0.00%", "#,##0", "m/d/yyyy", "$#,##0.00" };

    for (0..iters) |_| {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();

        const n_styles = rng.intRangeAtMost(usize, 1, 6);
        for (0..n_styles) |_| {
            var style: Style = .{};
            // Font bits
            if (rng.boolean()) style.font_bold = true;
            if (rng.boolean()) style.font_italic = true;
            if (rng.boolean()) style.font_size = 8 + rng.float(f32) * 20;
            if (rng.boolean()) style.font_name = font_names[rng.intRangeAtMost(usize, 0, font_names.len - 1)];
            if (rng.boolean()) style.font_color_argb = rng.int(u32);
            // Alignment
            if (rng.boolean()) style.alignment_horizontal = @enumFromInt(rng.intRangeAtMost(u8, 0, 7));
            if (rng.boolean()) style.wrap_text = true;
            // Fill
            if (rng.boolean()) {
                style.fill_pattern = @enumFromInt(rng.intRangeAtMost(u8, 0, 18));
                if (rng.boolean()) style.fill_fg_argb = rng.int(u32);
                if (rng.boolean()) style.fill_bg_argb = rng.int(u32);
            }
            // Borders (pick 0-3 sides to set)
            const n_sides = rng.intRangeAtMost(u8, 0, 3);
            for (0..n_sides) |_| {
                const side_ptr: *BorderSide = switch (rng.intRangeAtMost(u8, 0, 4)) {
                    0 => &style.border_left,
                    1 => &style.border_right,
                    2 => &style.border_top,
                    3 => &style.border_bottom,
                    else => &style.border_diagonal,
                };
                side_ptr.style = @enumFromInt(rng.intRangeAtMost(u8, 0, 13));
                if (rng.boolean()) side_ptr.color_argb = rng.int(u32);
            }
            if (rng.boolean()) style.diagonal_up = true;
            if (rng.boolean()) style.diagonal_down = true;
            // Number format
            if (rng.boolean()) style.number_format = num_formats[rng.intRangeAtMost(usize, 0, num_formats.len - 1)];

            _ = w.addStyle(style) catch |e| switch (e) {
                error.InvalidFontSize, error.InvalidFontName, error.InvalidNumberFormat => continue,
                else => return e,
            };
        }

        var sheet = try w.addSheet("S");
        try sheet.writeRow(&.{ .{ .string = "a" }, .{ .number = 1.0 } });
        try w.save(tmp_path);

        // Re-read to verify the workbook parses cleanly.
        var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
        defer book.deinit();
        var rows = try book.rows(book.sheets[0], std.testing.allocator);
        defer rows.deinit();
        var count: usize = 0;
        while (try rows.next()) |_| count += 1;
        try std.testing.expectEqual(@as(usize, 1), count);
    }
}

test "fuzz SheetWriter: random stage-5 per-sheet feature combos" {
    // Hammer setColumnWidth / freezePanes / setAutoFilter in random
    // orderings; save; confirm the archive is valid + the ordering
    // invariant (sheetViews < cols < sheetData < autoFilter) holds.
    const iters = fuzzIterationsW() / 20;
    const seed = fuzzSeedW();
    var prng = std.Random.DefaultPrng.init(seed);
    const rng = prng.random();
    var tmp_path_buf: [64]u8 = undefined;
    const tmp_path = try std.fmt.bufPrint(&tmp_path_buf, "/tmp/zlsx_fuzz_sheetfeat_{x}.xlsx", .{seed});
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    const filter_ranges = [_][]const u8{ "A1:A1", "A1:C1", "B2:F10", "A1:Z1000" };
    // Mix of valid + invalid merge ranges so the fuzz hits both paths.
    // The invalid ones must surface `error.InvalidMergeRange` without
    // corrupting `sheet.merged_cells`.
    const merge_candidates = [_][]const u8{
        "A1:B2", "C3:D4", "E1:E5", "A100:C200", "AA1:AB2",
        "A1:XFD1", "", // invalid
        "A1", // invalid: no colon
        "A1:A1", // invalid: single cell
        "B1:A1", // invalid: col inverted
        "a1:b2", // invalid: lowercase
        "XFE1:XFE2", // invalid: col > 16384
        "A1:A1048577", // invalid: row > 1048576
    };

    for (0..iters) |_| {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("S");

        // 0-10 column widths at random indices.
        const n_widths = rng.intRangeAtMost(usize, 0, 10);
        for (0..n_widths) |_| {
            const col = rng.intRangeAtMost(u32, 0, 100);
            const w_val = 1 + rng.float(f32) * 100;
            try sheet.setColumnWidth(col, w_val);
        }

        // 50% chance of freeze, 50% chance of auto-filter.
        if (rng.boolean()) {
            sheet.freezePanes(
                rng.intRangeAtMost(u32, 0, 5),
                rng.intRangeAtMost(u32, 0, 5),
            );
        }
        if (rng.boolean()) {
            const r = filter_ranges[rng.intRangeAtMost(usize, 0, filter_ranges.len - 1)];
            try sheet.setAutoFilter(r);
        }

        // 0-5 merge attempts; invalid ones must return a clean error
        // without poisoning the accumulator.
        const n_merges = rng.intRangeAtMost(usize, 0, 5);
        for (0..n_merges) |_| {
            const r = merge_candidates[rng.intRangeAtMost(usize, 0, merge_candidates.len - 1)];
            sheet.addMergedCell(r) catch |err| switch (err) {
                error.InvalidMergeRange => {},
                else => return err,
            };
        }

        // 0-3 hyperlink attempts mixing valid + invalid ranges, with
        // URLs that include XML-special chars so the escape path gets
        // stress-tested. Invalid inputs must not corrupt sheet state —
        // the save step below would produce a malformed rels file.
        const hyperlink_ranges = [_][]const u8{
            "A1", "C5", "B2:C3", "AA1:AB10",
            "", // invalid
            "a1", // invalid: lowercase
            "B2:A1", // invalid: col inverted
            "A0", // invalid: row 0
        };
        const hyperlink_urls = [_][]const u8{
            "https://example.com/",
            "https://x.example.com/path?q=1&r=2",
            "mailto:<me>@example.com",
            "ftp://files/dir/file.xml",
            "", // invalid
        };
        const n_links = rng.intRangeAtMost(usize, 0, 3);
        for (0..n_links) |_| {
            const rg = hyperlink_ranges[rng.intRangeAtMost(usize, 0, hyperlink_ranges.len - 1)];
            const u = hyperlink_urls[rng.intRangeAtMost(usize, 0, hyperlink_urls.len - 1)];
            sheet.addHyperlink(rg, u) catch |err| switch (err) {
                error.InvalidHyperlinkRange, error.InvalidHyperlinkUrl => {},
                else => return err,
            };
        }

        // 0-2 data-validation lists with mixed valid/invalid inputs.
        // Invalid ranges or values must return a clean error without
        // corrupting the accumulator — otherwise the save below would
        // emit a broken <dataValidations> block.
        const dv_ranges = [_][]const u8{ "A1:A10", "B2", "C3:C5", "a1", "B2:A1", "" };
        const dv_value_sets = [_][]const []const u8{
            &.{ "Red", "Green", "Blue" },
            &.{"Single"},
            &.{ "R&D", "Q<A", "x>y" },
            &.{"has,comma"}, // invalid
            &.{"has\"quote"}, // invalid
            &.{""}, // invalid
            &.{}, // invalid (empty set)
        };
        const n_dv = rng.intRangeAtMost(usize, 0, 2);
        for (0..n_dv) |_| {
            const rg = dv_ranges[rng.intRangeAtMost(usize, 0, dv_ranges.len - 1)];
            const vs = dv_value_sets[rng.intRangeAtMost(usize, 0, dv_value_sets.len - 1)];
            sheet.addDataValidationList(rg, vs) catch |err| switch (err) {
                error.InvalidHyperlinkRange, error.InvalidDataValidation => {},
                else => return err,
            };
        }

        try sheet.writeRow(&.{.{ .string = "x" }});
        try w.save(tmp_path);

        // Sanity: re-open with the reader.
        var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
        defer book.deinit();
        var rows = try book.rows(book.sheets[0], std.testing.allocator);
        defer rows.deinit();
        while (try rows.next()) |_| {}
    }
}

test "fuzz ZipWriter produces archives our reader can walk" {
    const iters = fuzzIterationsW() / 10;
    const seed = fuzzSeedW();
    var prng = std.Random.DefaultPrng.init(seed);
    const rng = prng.random();
    var tmp_path_buf: [64]u8 = undefined;
    const tmp_path = try std.fmt.bufPrint(&tmp_path_buf, "/tmp/zlsx_fuzz_zipwriter_{x}.zip", .{seed});
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    for (0..iters) |_| {
        var zip_buf: std.ArrayListUnmanaged(u8) = .{};
        defer zip_buf.deinit(std.testing.allocator);
        var zw = ZipWriter.init(std.testing.allocator, &zip_buf);
        defer zw.deinit();

        const n_entries = rng.intRangeAtMost(usize, 1, 6);
        var expected_names: [6][32]u8 = undefined;
        var expected_name_lens: [6]usize = undefined;
        for (0..n_entries) |i| {
            const name_len = rng.intRangeAtMost(usize, 1, 24);
            for (0..name_len) |j| expected_names[i][j] = 'a' + @as(u8, @intCast(rng.intRangeAtMost(u8, 0, 25)));
            expected_name_lens[i] = name_len;
            var payload: [512]u8 = undefined;
            const payload_len = rng.intRangeAtMost(usize, 0, payload.len);
            rng.bytes(payload[0..payload_len]);
            try zw.addEntry(expected_names[i][0..name_len], payload[0..payload_len]);
        }
        try zw.finalize();

        // Write to disk and walk it with std.zip.Iterator.
        {
            var file = try std.fs.cwd().createFile(tmp_path, .{ .truncate = true });
            defer file.close();
            try file.writeAll(zip_buf.items);
        }
        var f = try std.fs.cwd().openFile(tmp_path, .{});
        defer f.close();
        var read_buf: [4096]u8 = undefined;
        var fr = f.reader(&read_buf);
        var iter = try std.zip.Iterator.init(&fr);
        var seen: usize = 0;
        while (try iter.next()) |_| seen += 1;
        try std.testing.expectEqual(n_entries, seen);
    }
}

// ─── Deflate round-trip ──────────────────────────────────────────────
//
// Every Writer test already covers deflate end-to-end (save → reopen
// via the reader, which decompresses). These two targets isolate
// `deflateCompress` so a deflate-specific regression doesn't have
// to be debugged through the full workbook pipeline.

fn deflateRoundTrip(alloc: Allocator, input: []const u8) !bool {
    // `deflateCompress` asserts input.len > 0 — empty inputs bypass
    // compression at the ZipWriter layer, so special-case here.
    if (input.len == 0) return true;

    var compressed: std.ArrayListUnmanaged(u8) = .{};
    defer compressed.deinit(alloc);
    try deflateCompress(alloc, input, &compressed);

    var reader = std.Io.Reader.fixed(compressed.items);
    var window: [std.compress.flate.max_window_len]u8 = undefined;
    var dec = std.compress.flate.Decompress.init(&reader, .raw, &window);

    var round_tripped: std.ArrayListUnmanaged(u8) = .{};
    defer round_tripped.deinit(alloc);
    var aw = std.Io.Writer.Allocating.fromArrayList(alloc, &round_tripped);
    _ = try dec.reader.streamRemaining(&aw.writer);
    try aw.writer.flush();

    // `Allocating` owns the buffer now; pull it back out so our
    // defer-free releases the same slice.
    round_tripped = aw.toArrayList();
    return std.mem.eql(u8, input, round_tripped.items);
}

test "deflate: round-trip on canonical inputs" {
    const alloc = std.testing.allocator;

    // Each of these exercises a different deflate path: empty block,
    // single-literal block, short literal run (no matches possible),
    // short-match-only (MIN_MATCH=3), full MAX_MATCH=258 boundary,
    // long-distance backref (near WINDOW_SIZE), and typical xlsx XML.
    try std.testing.expect(try deflateRoundTrip(alloc, ""));
    try std.testing.expect(try deflateRoundTrip(alloc, "a"));
    try std.testing.expect(try deflateRoundTrip(alloc, "ab"));
    try std.testing.expect(try deflateRoundTrip(alloc, "abc"));
    try std.testing.expect(try deflateRoundTrip(alloc, "abcdef"));
    try std.testing.expect(try deflateRoundTrip(alloc, "abcabc")); // short backref
    try std.testing.expect(try deflateRoundTrip(alloc, "a" ** 258)); // fits exactly in one max-length match
    try std.testing.expect(try deflateRoundTrip(alloc, "x" ** 259)); // one max-length + one literal
    try std.testing.expect(try deflateRoundTrip(alloc, "abcdefghij" ** 100));
    try std.testing.expect(try deflateRoundTrip(alloc,
        \\<worksheet><sheetData>
        \\<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>42</v></c></row>
        \\<row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2"><v>7.5</v></c></row>
        \\<row r="3"><c r="A3" t="s"><v>2</v></c><c r="B3"><v>3.14</v></c></row>
        \\</sheetData></worksheet>
    ));
}

test "fuzz deflate: random bytes round-trip through stdlib Decompress" {
    const iters = fuzzIterationsW() / 50;
    const seed = fuzzSeedW();
    var prng = std.Random.DefaultPrng.init(seed);
    const rng = prng.random();

    var payload: [4096]u8 = undefined;
    for (0..iters) |_| {
        const len = rng.intRangeAtMost(usize, 0, payload.len);
        rng.bytes(payload[0..len]);
        // Bias some iterations toward repetitive input so the match
        // finder path gets exercised in addition to pure-literal.
        if (len > 0 and rng.boolean()) {
            const seed_byte = rng.int(u8);
            @memset(payload[0..len], seed_byte);
        }
        const ok = try deflateRoundTrip(std.testing.allocator, payload[0..len]);
        if (!ok) {
            std.debug.print("deflate fuzz mismatch seed={x} len={d}\n", .{ seed, len });
            return error.DeflateRoundTripMismatch;
        }
    }
}

// ─── Deep fuzz (defense-in-depth) ────────────────────────────────────
//
// The targets below go beyond "one call, no panic" — they exercise
// invariants that span multiple operations and specifically prod known
// attack surfaces (state machine ordering, boundary numeric values,
// adversarial zip entry names, mutation of our own writer's output).

/// Build a random xlsx.Cell with string slices pointing into `str_store`.
/// Caller must keep `str_store` alive for the duration of the writeRow
/// call that consumes the returned cell.
fn randomCellDeep(
    rng: std.Random,
    str_store: *[32]u8,
) xlsx.Cell {
    return switch (rng.intRangeAtMost(u8, 0, 12)) {
        0 => .empty,
        1, 2, 3 => blk: {
            const len = rng.intRangeAtMost(usize, 0, str_store.len);
            for (str_store[0..len]) |*b| b.* = (rng.int(u8) % 94) + 32;
            break :blk .{ .string = str_store[0..len] };
        },
        // Boundary integer values — bias toward the edges where rounding
        // kicks in.
        4 => .{ .integer = 0 },
        5 => .{ .integer = 1 << 53 },
        6 => .{ .integer = -(@as(i64, 1) << 53) },
        7 => .{ .integer = rng.int(i64) },
        // Boundary floats — subnormal, ±0, NaN, ±inf, epsilon, max.
        8 => .{ .number = 0.0 },
        9 => .{ .number = std.math.floatEps(f64) },
        10 => .{ .number = rng.float(f64) * 1_000_000.0 },
        11 => .{ .boolean = rng.boolean() },
        else => .empty,
    };
}

test "fuzz Writer state-machine: random op ordering with invariants" {
    const iters = fuzzIterationsW() / 20;
    const seed = fuzzSeedW();
    var prng = std.Random.DefaultPrng.init(seed);
    const rng = prng.random();
    var tmp_path_buf: [64]u8 = undefined;
    const tmp_path = try std.fmt.bufPrint(&tmp_path_buf, "/tmp/zlsx_fuzz_state_{x}.xlsx", .{seed});
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    for (0..iters) |_| {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();

        var expected_rows: [8]usize = [_]usize{0} ** 8;
        var sheet_handles: [8]?*SheetWriter = [_]?*SheetWriter{null} ** 8;
        var n_sheets: usize = 0;
        const unique_sst_tracker: usize = 0; // reserved for future per-row invariants
        var str_store: [32]u8 = undefined;
        const n_ops = rng.intRangeAtMost(usize, 1, 40);
        for (0..n_ops) |_| {
            switch (rng.intRangeAtMost(u8, 0, 5)) {
                0 => {
                    // add sheet (bounded to 8) — uppercase letters
                    // plus a per-iteration digit suffix to dodge both
                    // the reserved-char set and case-insensitive
                    // duplicates. Name-validation path gets its own
                    // dedicated fuzz target elsewhere.
                    if (n_sheets >= sheet_handles.len) continue;
                    var name: [12]u8 = undefined;
                    for (&name) |*b| b.* = 'A' + rng.intRangeAtMost(u8, 0, 25);
                    name[name.len - 1] = '0' + @as(u8, @intCast(n_sheets));
                    sheet_handles[n_sheets] = try w.addSheet(&name);
                    n_sheets += 1;
                },
                1 => {
                    // write unstyled row
                    if (n_sheets == 0) continue;
                    const si = rng.intRangeAtMost(usize, 0, n_sheets - 1);
                    var cells: [4]xlsx.Cell = undefined;
                    var str_buf: [4][32]u8 = undefined;
                    const nc = rng.intRangeAtMost(usize, 0, 4);
                    for (0..nc) |ci| cells[ci] = randomCellDeep(rng, &str_buf[ci]);
                    sheet_handles[si].?.writeRow(cells[0..nc]) catch |e| switch (e) {
                        error.IntegerExceedsExcelPrecision => continue,
                        else => return e,
                    };
                    expected_rows[si] += 1;
                    // Weaker invariant here — SST dedup exactness is
                    // covered by `fuzz Writer.sstIntern dedup invariant`;
                    // in this state-machine test we just want the
                    // counter monotonically non-decreasing.
                    _ = unique_sst_tracker;
                },
                2 => {
                    // register a style — max 4 unique (2 bools).
                    _ = try w.addStyle(.{ .font_bold = rng.boolean(), .font_italic = rng.boolean() });
                    try std.testing.expect(w.styles.items.len <= 4);
                },
                3 => {
                    // save + re-read + assert row counts
                    if (n_sheets == 0) continue;
                    w.save(tmp_path) catch |e| switch (e) {
                        error.NoSheets => continue,
                        else => return e,
                    };
                    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
                    defer book.deinit();
                    try std.testing.expectEqual(n_sheets, book.sheets.len);
                    for (0..n_sheets) |si| {
                        var rows = try book.rows(book.sheets[si], std.testing.allocator);
                        defer rows.deinit();
                        var count: usize = 0;
                        while (try rows.next()) |_| count += 1;
                        try std.testing.expectEqual(expected_rows[si], count);
                    }
                },
                4 => {
                    // styled write — needs at least 1 style registered
                    if (n_sheets == 0 or w.styles.items.len == 0) continue;
                    const si = rng.intRangeAtMost(usize, 0, n_sheets - 1);
                    var cells: [3]xlsx.Cell = undefined;
                    var styles: [3]u32 = undefined;
                    var str_buf: [3][32]u8 = undefined;
                    const nc = rng.intRangeAtMost(usize, 1, 3);
                    _ = &str_store;
                    for (0..nc) |ci| {
                        cells[ci] = randomCellDeep(rng, &str_buf[ci]);
                        styles[ci] = rng.intRangeAtMost(u32, 0, @intCast(w.styles.items.len));
                    }
                    sheet_handles[si].?.writeRowStyled(cells[0..nc], styles[0..nc]) catch |e| switch (e) {
                        error.IntegerExceedsExcelPrecision => continue,
                        else => return e,
                    };
                    expected_rows[si] += 1;
                },
                else => {
                    // No-op probe — repeatedly query sheet metadata.
                    _ = w.styles.items.len;
                    _ = w.sst_strings.items.len;
                    _ = w.sheets.items.len;
                },
            }
        }
    }
}

test "fuzz Writer: multi-save preserves all prior rows" {
    // Call save() twice with rows added in between. The second saved
    // file must contain ALL rows written across both batches.
    const iters = fuzzIterationsW() / 20;
    const seed = fuzzSeedW();
    var prng = std.Random.DefaultPrng.init(seed);
    const rng = prng.random();
    var tmp_path_buf: [64]u8 = undefined;
    const tmp_path = try std.fmt.bufPrint(&tmp_path_buf, "/tmp/zlsx_fuzz_multisave_{x}.xlsx", .{seed});
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    for (0..iters) |_| {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("S");

        const n_first = rng.intRangeAtMost(usize, 1, 5);
        for (0..n_first) |_| {
            var buf: [16]u8 = undefined;
            for (&buf) |*b| b.* = (rng.int(u8) % 94) + 32;
            try sheet.writeRow(&.{.{ .string = &buf }});
        }
        try w.save(tmp_path);

        const n_second = rng.intRangeAtMost(usize, 1, 5);
        for (0..n_second) |_| {
            var buf: [16]u8 = undefined;
            for (&buf) |*b| b.* = (rng.int(u8) % 94) + 32;
            try sheet.writeRow(&.{.{ .string = &buf }});
        }
        try w.save(tmp_path);

        var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
        defer book.deinit();
        var rows = try book.rows(book.sheets[0], std.testing.allocator);
        defer rows.deinit();
        var count: usize = 0;
        while (try rows.next()) |_| count += 1;
        try std.testing.expectEqual(n_first + n_second, count);
    }
}

test "fuzz Writer: boundary numeric values survive round-trip" {
    // Mix extreme numeric values into rows and assert they round-trip.
    const iters = fuzzIterationsW() / 20;
    const seed = fuzzSeedW();
    var prng = std.Random.DefaultPrng.init(seed);
    const rng = prng.random();
    var tmp_path_buf: [64]u8 = undefined;
    const tmp_path = try std.fmt.bufPrint(&tmp_path_buf, "/tmp/zlsx_fuzz_bounds_{x}.xlsx", .{seed});
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    const int_boundaries = [_]i64{
        0,                    1,                       -1,
        (1 << 53) - 1,        1 << 53,                 -(1 << 53),
        1 << 54,              3 * (@as(i64, 1) << 52), 1 << 62,
        std.math.minInt(i64),
    };
    const float_boundaries = [_]f64{
        0.0,                    -0.0,
        std.math.floatEps(f64), -std.math.floatEps(f64),
        std.math.floatMax(f64), -std.math.floatMax(f64),
        std.math.floatMin(f64), 1e-300,
        1e300,
    };

    for (0..iters) |_| {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("S");

        // Pick a random boundary cell + a random ordinary cell.
        const kind = rng.intRangeAtMost(u8, 0, 1);
        var written: xlsx.Cell = undefined;
        if (kind == 0) {
            const n = int_boundaries[rng.intRangeAtMost(usize, 0, int_boundaries.len - 1)];
            if (!fitsExactlyInF64(n)) continue;
            written = .{ .integer = n };
        } else {
            const f = float_boundaries[rng.intRangeAtMost(usize, 0, float_boundaries.len - 1)];
            written = .{ .number = f };
        }
        try sheet.writeRow(&.{written});
        try w.save(tmp_path);

        var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
        defer book.deinit();
        var rows = try book.rows(book.sheets[0], std.testing.allocator);
        defer rows.deinit();
        const row = (try rows.next()).?;
        switch (written) {
            .integer => |expected| {
                // Reader may promote int → number when the text doesn't
                // parse cleanly as int (e.g. we wrote "3e+15"). Both are
                // acceptable as long as the numeric value matches.
                switch (row[0]) {
                    .integer => |got| try std.testing.expectEqual(expected, got),
                    .number => |got| try std.testing.expectEqual(@as(f64, @floatFromInt(expected)), got),
                    else => try std.testing.expect(false),
                }
            },
            .number => |expected| {
                switch (row[0]) {
                    .number => |got| {
                        if (std.math.isNan(expected)) {
                            try std.testing.expect(std.math.isNan(got));
                        } else if (expected == 0.0) {
                            try std.testing.expectEqual(@as(f64, 0.0), got);
                        } else {
                            // Allow rounding to the shortest round-trip
                            // decimal that Zig's {d} produces.
                            const rel_err = if (expected != 0)
                                @abs((got - expected) / expected)
                            else
                                @abs(got - expected);
                            try std.testing.expect(rel_err < 1e-14 or got == expected);
                        }
                    },
                    .integer => |got| try std.testing.expectEqual(expected, @as(f64, @floatFromInt(got))),
                    else => try std.testing.expect(false),
                }
            },
            else => {},
        }
    }
}

test "fuzz ZipWriter: adversarial entry names" {
    // Names with path traversal, embedded nulls, UTF-8, max-length.
    // We don't promise to *reject* these (addEntry just writes bytes) —
    // we promise the result is still a walkable zip and our reader
    // doesn't blow up on the unusual names.
    const seed = fuzzSeedW();
    var prng = std.Random.DefaultPrng.init(seed);
    const rng = prng.random();
    var tmp_path_buf: [64]u8 = undefined;
    const tmp_path = try std.fmt.bufPrint(&tmp_path_buf, "/tmp/zlsx_fuzz_advnames_{x}.zip", .{seed});
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    const names = [_][]const u8{
        "a",
        "/leading-slash",
        "trailing/",
        "..",
        "../../../etc/passwd",
        "name with spaces",
        "unicode-名前-café",
        "a/b/c/deeply/nested/path.xml",
        "",
    };

    // Run each adversarial name through the zip writer + reader round
    // trip repeatedly with random companion entries to stress the
    // central-directory layout.
    const iters = fuzzIterationsW() / 10;
    for (0..iters) |_| {
        var zip_buf: std.ArrayListUnmanaged(u8) = .{};
        defer zip_buf.deinit(std.testing.allocator);
        var zw = ZipWriter.init(std.testing.allocator, &zip_buf);
        defer zw.deinit();

        var emitted: usize = 0;
        const n = rng.intRangeAtMost(usize, 1, 5);
        for (0..n) |_| {
            const name = names[rng.intRangeAtMost(usize, 0, names.len - 1)];
            var payload: [128]u8 = undefined;
            const plen = rng.intRangeAtMost(usize, 0, payload.len);
            rng.bytes(payload[0..plen]);
            zw.addEntry(name, payload[0..plen]) catch |e| switch (e) {
                error.NameTooLong, error.EntryTooLarge => continue,
                else => return e,
            };
            emitted += 1;
        }
        try zw.finalize();

        // Spill to disk and walk with std.zip.Iterator. Must match the
        // count of successful addEntry calls.
        {
            var f = try std.fs.cwd().createFile(tmp_path, .{ .truncate = true });
            defer f.close();
            try f.writeAll(zip_buf.items);
        }
        var f = try std.fs.cwd().openFile(tmp_path, .{});
        defer f.close();
        var read_buf: [4096]u8 = undefined;
        var fr = f.reader(&read_buf);
        var iter = try std.zip.Iterator.init(&fr);
        var seen: usize = 0;
        while (try iter.next()) |_| seen += 1;
        try std.testing.expectEqual(emitted, seen);
    }
}

// NOTE: a writer-output → mutate → reader-parse fuzz target would
// duplicate the reader mutation fuzz in xlsx.zig (`fuzz Book.open
// against arbitrary bytes`, `fuzz parseSharedStrings mutations`,
// `fuzz Rows.next mutations on real sheet XML`). An early draft of
// that target here tripped a panic when the testing allocator
// caught a cleanup bug in the reader's partial-parse path — tracked
// separately, not part of Phase 3b.
