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

/// Domain errors surfaced by this module. The public API uses inferred
/// error sets so callers get the full union of these plus whatever
/// std.zip / std.fs / std.compress.flate decide to return.
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
    /// Index into `shared_strings_xml` (or backing owned strings if we
    /// had to decode entities). Order matches the SST in the file.
    shared_strings: [][]const u8 = &.{},
    /// Owned decoded shared strings when entity decoding was needed.
    /// Empty if every SST entry was verbatim.
    shared_owned: std.ArrayListUnmanaged([]u8) = .{},
    /// (name, path) for each `<sheet>` in the workbook, in declared order.
    sheets: []Sheet = &.{},
    /// Decompressed bytes of each sheet's XML, keyed by path.
    sheet_data: std.StringHashMapUnmanaged([]u8) = .{},
    /// Owned backing storage for every string referenced by `sheets`,
    /// sheet_data keys, and entity-decoded shared strings.
    strings: std.ArrayListUnmanaged([]u8) = .{},

    /// Open and parse the workbook skeleton. Sheet XML is eagerly
    /// decompressed (xlsx files we target are small — ~300 KB — and
    /// streaming through std.zip is awkward).
    pub fn open(allocator: Allocator, path: []const u8) !Book {
        var book: Book = .{ .allocator = allocator };
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

        return book;
    }

    pub fn deinit(self: *Book) void {
        const a = self.allocator;
        if (self.shared_strings_xml) |s| a.free(s);
        a.free(self.shared_strings);
        for (self.shared_owned.items) |s| a.free(s);
        self.shared_owned.deinit(a);

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
            .owned = .{},
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
    /// Per-row owned strings: any string that required entity decoding
    /// or scanning across multiple `<t>` runs gets its own allocation
    /// here, so `[]const u8` slices in `Cell.string` never dangle when
    /// later cells in the same row allocate more text. Freed at the
    /// start of each `next()` call and at `deinit`.
    owned: std.ArrayListUnmanaged([]u8),

    pub fn deinit(self: *Rows) void {
        for (self.owned.items) |s| self.allocator.free(s);
        self.owned.deinit(self.allocator);
        self.row_cells.deinit(self.allocator);
        self.* = undefined;
    }

    fn clearOwned(self: *Rows) void {
        for (self.owned.items) |s| self.allocator.free(s);
        self.owned.clearRetainingCapacity();
    }

    /// Returns the next row's cells, or null at end-of-sheet. Returned
    /// slice is valid until the next call to `next()` (or until
    /// `deinit()`). Cell string contents are either shared-string slices
    /// (owned by the Book), xml-backed slices (stable for the Book's
    /// lifetime), or row-owned slices that are freed on the next call.
    pub fn next(self: *Rows) !?[]const Cell {
        self.row_cells.clearRetainingCapacity();
        self.clearOwned();

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
    /// allocate an owned slice tracked in `self.owned`.
    fn decodeVValue(self: *Rows, body: []const u8) ![]const u8 {
        const raw = extractVValue(body) orelse return "";
        return try self.internOrBorrow(raw);
    }

    /// Decode inline-string body `<is>(<r>)?<t>text</t>(</r>)?</is>`.
    /// Single-`<t>` bodies without entities borrow from xml. Anything
    /// else (rich-text runs, entities) gets an owned allocation in
    /// `self.owned`.
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

        // Multi-run or entity-bearing — allocate.
        var buf: std.ArrayListUnmanaged(u8) = .{};
        errdefer buf.deinit(self.allocator);
        var i: usize = 0;
        while (std.mem.indexOfPos(u8, body, i, "<t")) |t_start| {
            const gt = std.mem.indexOfScalarPos(u8, body, t_start, '>') orelse return error.MalformedXml;
            if (body[gt - 1] == '/') {
                i = gt + 1;
                continue;
            }
            const t_close = std.mem.indexOfPos(u8, body, gt + 1, "</t>") orelse return error.MalformedXml;
            try appendDecoded(self.allocator, &buf, body[gt + 1 .. t_close]);
            i = t_close + "</t>".len;
        }
        const owned = try buf.toOwnedSlice(self.allocator);
        try self.owned.append(self.allocator, owned);
        return owned;
    }

    /// Return `raw` unchanged if it needs no decoding; otherwise allocate
    /// an owned slice and track it in `self.owned`.
    fn internOrBorrow(self: *Rows, raw: []const u8) ![]const u8 {
        if (std.mem.indexOfScalar(u8, raw, '&') == null) return raw;
        var buf: std.ArrayListUnmanaged(u8) = .{};
        errdefer buf.deinit(self.allocator);
        try appendDecoded(self.allocator, &buf, raw);
        const owned = try buf.toOwnedSlice(self.allocator);
        try self.owned.append(self.allocator, owned);
        return owned;
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

fn parseSharedStrings(book: *Book, sst_xml: []u8) !void {
    var strings: std.ArrayListUnmanaged([]const u8) = .{};
    errdefer strings.deinit(book.allocator);

    var i: usize = 0;
    while (std.mem.indexOfPos(u8, sst_xml, i, "<si")) |si_start| {
        const si_gt = std.mem.indexOfScalarPos(u8, sst_xml, si_start, '>') orelse break;
        const is_self_closing = si_gt > 0 and sst_xml[si_gt - 1] == '/';
        if (is_self_closing) {
            try strings.append(book.allocator, "");
            i = si_gt + 1;
            continue;
        }
        const si_close = std.mem.indexOfPos(u8, sst_xml, si_gt, "</si>") orelse break;
        const body = sst_xml[si_gt + 1 .. si_close];

        // Concatenate every <t>…</t> in the body.
        var concat: std.ArrayListUnmanaged(u8) = .{};
        errdefer concat.deinit(book.allocator);
        var needs_decoding = false;

        var j: usize = 0;
        while (std.mem.indexOfPos(u8, body, j, "<t")) |t_start| {
            const t_gt = std.mem.indexOfScalarPos(u8, body, t_start, '>') orelse break;
            if (body[t_gt - 1] == '/') {
                j = t_gt + 1;
                continue;
            }
            const t_close = std.mem.indexOfPos(u8, body, t_gt + 1, "</t>") orelse break;
            const raw = body[t_gt + 1 .. t_close];
            if (std.mem.indexOfScalar(u8, raw, '&') != null) needs_decoding = true;
            try concat.appendSlice(book.allocator, raw);
            j = t_close + "</t>".len;
        }

        if (needs_decoding) {
            var decoded: std.ArrayListUnmanaged(u8) = .{};
            errdefer decoded.deinit(book.allocator);
            try appendDecoded(book.allocator, &decoded, concat.items);
            concat.deinit(book.allocator);
            const owned = try decoded.toOwnedSlice(book.allocator);
            try book.shared_owned.append(book.allocator, owned);
            try strings.append(book.allocator, owned);
        } else {
            // Point into sst_xml directly to save an allocation.
            // We need the slice to outlive concat — rescan the body to
            // locate the original span. For single-<t> entries this is
            // trivial; rich-text runs we handle via the owned path.
            const first_t = std.mem.indexOf(u8, body, "<t") orelse {
                concat.deinit(book.allocator);
                try strings.append(book.allocator, "");
                i = si_close + "</si>".len;
                continue;
            };
            const t_gt = std.mem.indexOfScalarPos(u8, body, first_t, '>') orelse break;
            const remaining = body[t_gt + 1 ..];
            const t_close_rel = std.mem.indexOf(u8, remaining, "</t>") orelse break;
            const span = remaining[0..t_close_rel];
            // If concat matches the single span we're pointing at, reuse.
            if (span.len == concat.items.len and std.mem.eql(u8, span, concat.items)) {
                concat.deinit(book.allocator);
                try strings.append(book.allocator, span);
            } else {
                // Rich-text run across multiple <t> — own the concat.
                const owned = try concat.toOwnedSlice(book.allocator);
                try book.shared_owned.append(book.allocator, owned);
                try strings.append(book.allocator, owned);
            }
        }

        i = si_close + "</si>".len;
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
        var book: Book = .{ .allocator = std.testing.allocator };
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
        var book: Book = .{ .allocator = std.testing.allocator };
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
        var book: Book = .{ .allocator = std.testing.allocator };
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
        var book: Book = .{ .allocator = std.testing.allocator };
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
        .owned = .{},
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
            .owned = .{},
        };
        defer rows.deinit();

        // Consume the rows — may error, must not panic.
        while (rows.next() catch null) |_| {}
    }
}
