//! `Workbook` — typed-overlay root for an OOXML spreadsheet
//! package (B1 iter-wb-2).
//!
//! Layered on top of `pkg/store.zig` (PartStore) and the
//! per-part typed parsers in `pkg/typed_parts/`. This iter is
//! read-only — `Workbook.open` + sheet lookup + per-sheet
//! cells / merges / hyperlinks / validations / conditional
//! formats / freeze. Mutation lands in iter-wb-4.
//!
//! Composition shape:
//!
//!     Workbook
//!     ├── store: PartStore           (owns arena for raw part bytes)
//!     ├── workbook: WorkbookXml      (typed view of xl/workbook.xml)
//!     ├── worksheets: []?Worksheet   (lazy slot per sheet, parsed on first access)
//!     ├── sst:    ?SstXml            (lazy)
//!     ├── styles: ?StylesXml         (lazy)
//!     └── arena_ws: ArenaAllocator   (per-Worksheet allocations)
//!
//! Each typed view (`WorkbookXml`, `SheetXml`, `SstXml`,
//! `StylesXml`) carries its own internal arena per the
//! `pkg/typed_parts/*.zig` contract. `Workbook.deinit` walks
//! all of them and reclaims.
//!
//! See `docs/plans/workbook-overlay.md` for the full plan.

const std = @import("std");
const Allocator = std.mem.Allocator;
const assert = std.debug.assert;

const store_mod = @import("store.zig");
const typed_parts = @import("typed_parts/root.zig");

const PartStore = store_mod.PartStore;

const workbook_xml_mod = typed_parts.workbook_xml;
const sheet_xml_mod = typed_parts.sheet_xml;
const sst_xml_mod = typed_parts.sst_xml;
const styles_xml_mod = typed_parts.styles_xml;

pub const Error = error{
    MissingWorkbookPart,
    MissingSheetPart,
    MissingRelationship,
    SheetIndexOutOfRange,
    SheetNotFound,
    SstIndexOutOfRange,
    SstEntryIsRich,
    InvalidCellRef,
    NoSheetData,
    UnsupportedCellValue,
    WriteFailed,
} || workbook_xml_mod.Error || sheet_xml_mod.ParseError || sst_xml_mod.Error || styles_xml_mod.Error || store_mod.Error ||
    std.fs.File.WriteError || std.fs.File.OpenError || std.fs.Dir.RenameError || std.fs.Dir.StatFileError;

/// Mutation primitive (B1 iter-wb-4). Strings emit as `inlineStr`
/// — cell-local text, no SST extension required. Formulas emit as
/// `<f>…</f>` with no cached `<v>`, so Excel recalculates on open.
///
/// `string` and `formula` slices borrow for `setCell`'s call only.
/// The delta map duplicates bytes into the Workbook allocator before
/// returning, so the caller can free / reuse the buffer as soon as
/// `setCell` returns.
pub const CellValue = union(enum) {
    blank: void,
    number: f64,
    boolean: bool,
    string: []const u8,
    /// Formula text (e.g. "SUM(A1:A10)" — no leading `=`). Emitted
    /// as `<f>…</f>` without a cached value; Excel recalculates the
    /// result on open. Caching computed results is a future iter
    /// (depends on D1 evaluator).
    formula: []const u8,
};

/// 1-based (row, col) — matches OOXML A1 conventions.
pub const CellRef = struct {
    row: u32,
    col: u32,

    pub fn eql(a: CellRef, b: CellRef) bool {
        return a.row == b.row and a.col == b.col;
    }
};

/// Parse an A1-style ref ("A1", "AA10") into a numeric CellRef.
/// Letters are case-insensitive. Returns `error.InvalidCellRef` for
/// any malformed input (no letters, no digits, leading-zero row,
/// out-of-range col [> Excel's 16384 limit] / row [> 1048576]).
pub fn parseA1Ref(ref: []const u8) Error!CellRef {
    if (ref.len < 2) return error.InvalidCellRef;
    var i: usize = 0;
    var col: u32 = 0;
    while (i < ref.len) : (i += 1) {
        const c = ref[i];
        const upper: u8 = if (c >= 'a' and c <= 'z') c - 32 else c;
        if (upper < 'A' or upper > 'Z') break;
        // col := col*26 + (upper - 'A' + 1); trapping arithmetic
        // catches overflow on absurd inputs ("AAAAAAAAAA").
        const inc: u32 = @as(u32, upper - 'A') + 1;
        col = std.math.mul(u32, col, 26) catch return error.InvalidCellRef;
        col = std.math.add(u32, col, inc) catch return error.InvalidCellRef;
    }
    if (i == 0) return error.InvalidCellRef; // no letters
    if (col > 16384) return error.InvalidCellRef; // Excel max column
    if (i == ref.len) return error.InvalidCellRef; // no digits
    if (ref[i] == '0') return error.InvalidCellRef; // leading zero forbidden
    var row: u32 = 0;
    while (i < ref.len) : (i += 1) {
        const c = ref[i];
        if (c < '0' or c > '9') return error.InvalidCellRef;
        const dig: u32 = c - '0';
        row = std.math.mul(u32, row, 10) catch return error.InvalidCellRef;
        row = std.math.add(u32, row, dig) catch return error.InvalidCellRef;
    }
    if (row == 0 or row > 1048576) return error.InvalidCellRef;
    return .{ .row = row, .col = col };
}

/// Format a CellRef as A1 ("A1", "AA10") into `buf`. Returns the
/// written slice. `buf.len >= 16` is sufficient for any in-range ref.
pub fn formatA1Ref(buf: []u8, ref: CellRef) []u8 {
    assert(ref.row >= 1 and ref.row <= 1048576);
    assert(ref.col >= 1 and ref.col <= 16384);
    assert(buf.len >= 16);

    // Letters: convert col (1-based) to base-26 with A=1..Z=26.
    var letters: [4]u8 = undefined;
    var n: usize = 0;
    var c: u32 = ref.col;
    while (c > 0) : (n += 1) {
        const r: u32 = (c - 1) % 26;
        letters[n] = @intCast(@as(u32, 'A') + r);
        c = (c - 1) / 26;
    }
    // Reverse letters into buf[0..n].
    var i: usize = 0;
    while (i < n) : (i += 1) buf[i] = letters[n - 1 - i];

    // Row digits — itoa.
    const row_str = std.fmt.bufPrint(buf[n..], "{d}", .{ref.row}) catch unreachable;
    return buf[0 .. n + row_str.len];
}

pub const Workbook = struct {
    allocator: Allocator,
    store: PartStore,

    /// Parsed `xl/workbook.xml`. Borrows from the PartStore arena
    /// for leaf strings; owns its own arena for spine slices.
    workbook: workbook_xml_mod.WorkbookXml,

    /// Lazy per-sheet typed view. Length == `workbook.sheets.len`.
    /// Each slot is `null` until `sheet(idx)` materialises it.
    worksheets: []Worksheet,

    /// Lazy workbook-scope views. Parsed on first access via
    /// `Workbook.sst()` / `Workbook.styles()`.
    sst_view: ?sst_xml_mod.SstXml = null,
    styles_view: ?styles_xml_mod.StylesXml = null,

    /// Open an .xlsx file as a typed `Workbook`.
    ///
    /// Errors if `xl/workbook.xml` is absent or malformed; otherwise
    /// every sheet is left lazy. `deinit` is required on success and
    /// on error after `open` returns successfully.
    pub fn open(allocator: Allocator, path: []const u8) Error!Workbook {
        assert(path.len > 0);

        var store = try PartStore.open(allocator, path);
        errdefer store.deinit();

        return try fromStore(allocator, store);
    }

    /// Construct a `Workbook` from an already-opened `PartStore`.
    /// Takes ownership of the store; `deinit` will tear it down.
    pub fn fromStore(allocator: Allocator, store: PartStore) Error!Workbook {
        var s = store;
        errdefer s.deinit();

        const wb_part = s.part("xl/workbook.xml") orelse
            return Error.MissingWorkbookPart;

        var workbook_view = try workbook_xml_mod.parse(allocator, wb_part.bytes);
        errdefer workbook_view.deinit(allocator);

        const ws_count = workbook_view.sheets.len;
        const slots = try allocator.alloc(Worksheet, ws_count);
        errdefer allocator.free(slots);

        for (slots, 0..) |*slot, i| slot.* = .{
            .workbook = undefined, // patched below; can't take address pre-return
            .sheet_idx = @intCast(i),
            .parsed = null,
            .resolved_part_name = null,
        };

        return .{
            .allocator = allocator,
            .store = s,
            .workbook = workbook_view,
            .worksheets = slots,
        };
    }

    pub fn deinit(self: *Workbook) void {
        for (self.worksheets) |*ws| ws.deinit(self.allocator);
        self.allocator.free(self.worksheets);

        if (self.sst_view) |*v| {
            var view = v.*;
            view.deinit(self.allocator);
        }
        if (self.styles_view) |*v| {
            var view = v.*;
            view.deinit(self.allocator);
        }
        self.workbook.deinit(self.allocator);
        self.store.deinit();
    }

    pub fn sheetCount(self: *const Workbook) u32 {
        assert(self.worksheets.len == self.workbook.sheets.len);
        return @intCast(self.worksheets.len);
    }

    /// Borrow a `Worksheet` handle by zero-based index. Materialises
    /// the typed view on first access; subsequent calls hit the cache.
    pub fn sheet(self: *Workbook, idx: u32) Error!*Worksheet {
        if (idx >= self.worksheets.len) return Error.SheetIndexOutOfRange;
        const ws = &self.worksheets[idx];
        // Patch the back-pointer on first observation. Done lazily so
        // the slot table can be allocated before `Workbook` exists.
        ws.workbook = self;
        return ws;
    }

    /// Borrow a `Worksheet` handle by sheet name (case-sensitive,
    /// no Unicode normalisation — match `WorkbookXml.Sheet.name`
    /// exactly). Returns `null` if no sheet with that name.
    pub fn sheetByName(self: *Workbook, name: []const u8) Error!?*Worksheet {
        assert(name.len > 0);
        for (self.workbook.sheets, 0..) |s, i| {
            if (std.mem.eql(u8, s.name, name)) return try self.sheet(@intCast(i));
        }
        return null;
    }

    /// Defined names from `xl/workbook.xml`. Borrowed from the
    /// `WorkbookXml` view; valid for the `Workbook`'s lifetime.
    pub fn definedNames(self: *const Workbook) []const workbook_xml_mod.DefinedName {
        return self.workbook.defined_names;
    }

    /// Defined names with no `localSheetId` attribute — workbook-scope
    /// names visible from every sheet. Allocator-owned (caller frees).
    pub fn definedNamesGlobal(self: *const Workbook, allocator: Allocator) Error![]workbook_xml_mod.DefinedName {
        var out: std.ArrayList(workbook_xml_mod.DefinedName) = .empty;
        errdefer out.deinit(allocator);
        for (self.workbook.defined_names) |dn| {
            if (dn.local_sheet_id == null) try out.append(allocator, dn);
        }
        return try out.toOwnedSlice(allocator);
    }

    /// Defined names scoped to a specific sheet (via `localSheetId`).
    /// Caller frees the returned slice.
    pub fn definedNamesForSheet(self: *const Workbook, allocator: Allocator, sheet_idx: u32) Error![]workbook_xml_mod.DefinedName {
        if (sheet_idx >= self.workbook.sheets.len) return Error.SheetIndexOutOfRange;
        var out: std.ArrayList(workbook_xml_mod.DefinedName) = .empty;
        errdefer out.deinit(allocator);
        for (self.workbook.defined_names) |dn| {
            if (dn.local_sheet_id) |sid| {
                if (sid == sheet_idx) try out.append(allocator, dn);
            }
        }
        return try out.toOwnedSlice(allocator);
    }

    /// Calc properties from `xl/workbook.xml`.
    pub fn calcProperties(self: *const Workbook) workbook_xml_mod.CalcProperties {
        return self.workbook.calc;
    }

    /// Convenience: SST entry `idx` as plain text. Errors on rich-run
    /// entries (caller must use `sst()` and walk `RichRun[]` directly).
    /// Returns the raw, undecoded slice — call `sst_xml.decodeText` to
    /// resolve `&amp;` etc.
    pub fn sstText(self: *Workbook, idx: u32) Error!?[]const u8 {
        const view = (try self.sst()) orelse return null;
        if (idx >= view.entries.len) return Error.SstIndexOutOfRange;
        switch (view.entries[idx]) {
            .plain => |s| return s,
            .rich => return Error.SstEntryIsRich,
        }
    }

    /// Lazily-parsed `xl/sharedStrings.xml`. Returns `null` if the
    /// workbook has no SST. Subsequent calls return the cached view.
    pub fn sst(self: *Workbook) Error!?*const sst_xml_mod.SstXml {
        if (self.sst_view != null) return &self.sst_view.?;
        const part = self.store.part("xl/sharedStrings.xml") orelse return null;
        self.sst_view = try sst_xml_mod.parse(self.allocator, part.bytes);
        return &self.sst_view.?;
    }

    /// Lazily-parsed `xl/styles.xml`. Returns `null` if absent.
    pub fn styles(self: *Workbook) Error!?*const styles_xml_mod.StylesXml {
        if (self.styles_view != null) return &self.styles_view.?;
        const part = self.store.part("xl/styles.xml") orelse return null;
        self.styles_view = try styles_xml_mod.parse(self.allocator, part.bytes);
        return &self.styles_view.?;
    }

    /// Persist all pending mutations to `path`. For each Worksheet
    /// with a non-empty delta map: regenerate the sheet's `<sheetData>`
    /// block from the typed view + deltas, splice into the source
    /// XML byte-preserving everything outside `<sheetData>`, push
    /// through PartStore.replacePart, then write the whole archive
    /// via PartStore.save.
    ///
    /// On success: every Worksheet's delta map is empty and any
    /// previously-cached `SheetXml` view is invalidated (next access
    /// re-parses from the new bytes).
    ///
    /// iter-wb-4 m1 limits: numeric / boolean / blank values only.
    /// Strings + formulas + new-row-insertion-with-original-cell-
    /// preservation come in m2.
    pub fn save(self: *Workbook, path: []const u8) Error!void {
        for (self.worksheets) |*ws| {
            if (ws.deltas.count() == 0) continue;
            _ = try ws.ensureParsed();
            const part_name = ws.resolved_part_name.?;
            const view = &ws.parsed.?;
            const source = blk: {
                const p = self.store.part(part_name) orelse return error.MissingSheetPart;
                break :blk p.bytes;
            };

            const new_xml = try emitSheetWithDeltas(self.allocator, source, view, &ws.deltas);
            defer self.allocator.free(new_xml);
            try self.store.replacePart(part_name, new_xml);

            freeDeltaStrings(self.allocator, &ws.deltas);
            ws.deltas.clearAndFree(self.allocator);
            // Invalidate the parsed view — its leaves borrowed from
            // the prior source bytes, which the caller may still see
            // as live (PartStore arena retains them) but the part's
            // logical content has changed.
            var stale = ws.parsed.?;
            stale.deinit(self.allocator);
            ws.parsed = null;
        }
        try self.store.save(path);
    }
};

// ─── Emit helpers (iter-wb-4 m1) ─────────────────────────────────────

/// Splice a regenerated `<sheetData>...</sheetData>` block into the
/// source sheet XML. Everything outside `<sheetData>` is copied
/// byte-for-byte. Returns a fresh allocator-owned slice.
fn emitSheetWithDeltas(
    allocator: Allocator,
    source: []const u8,
    view: *const sheet_xml_mod.SheetXml,
    deltas: *const std.AutoHashMapUnmanaged(CellRef, CellValue),
) Error![]u8 {
    assert(source.len > 0);

    const sd_idx = std.mem.indexOf(u8, source, "<sheetData") orelse
        return error.NoSheetData;
    const open_gt = std.mem.indexOfScalarPos(u8, source, sd_idx, '>') orelse
        return error.NoSheetData;
    const is_self_closing = open_gt > 0 and source[open_gt - 1] == '/';

    var prefix_end: usize = undefined;
    var suffix_start: usize = undefined;
    if (is_self_closing) {
        prefix_end = sd_idx; // we re-emit `<sheetData>` ourselves
        suffix_start = open_gt + 1;
    } else {
        prefix_end = open_gt + 1;
        suffix_start = std.mem.indexOfPos(u8, source, prefix_end, "</sheetData>") orelse
            return error.NoSheetData;
    }

    var out: std.ArrayList(u8) = .empty;
    errdefer out.deinit(allocator);
    try out.ensureTotalCapacity(allocator, source.len + 1024);

    try out.appendSlice(allocator, source[0..prefix_end]);
    if (is_self_closing) try out.appendSlice(allocator, "<sheetData>");

    try emitSheetData(allocator, &out, view, deltas);

    if (is_self_closing) try out.appendSlice(allocator, "</sheetData>");
    try out.appendSlice(allocator, source[suffix_start..]);

    return try out.toOwnedSlice(allocator);
}

const MergedCell = struct {
    ref: CellRef,
    style_idx: ?u32,
    payload: union(enum) {
        original: struct {
            cell_type: sheet_xml_mod.CellType,
            raw_value: ?[]const u8,
            formula: ?[]const u8,
        },
        delta: CellValue,
    },
};

fn mergedLessThan(_: void, a: MergedCell, b: MergedCell) bool {
    if (a.ref.row != b.ref.row) return a.ref.row < b.ref.row;
    return a.ref.col < b.ref.col;
}

fn emitSheetData(
    allocator: Allocator,
    out: *std.ArrayList(u8),
    view: *const sheet_xml_mod.SheetXml,
    deltas: *const std.AutoHashMapUnmanaged(CellRef, CellValue),
) Error!void {
    // 1. Collect existing cells (override with delta if matching).
    var merged: std.ArrayList(MergedCell) = .empty;
    defer merged.deinit(allocator);

    var seen: std.AutoHashMapUnmanaged(CellRef, void) = .{};
    defer seen.deinit(allocator);

    for (view.rows) |row| {
        for (row.cells) |c| {
            const cr = parseA1Ref(c.ref) catch continue;
            const overlay = deltas.get(cr);
            const mc: MergedCell = if (overlay) |dv| .{
                .ref = cr,
                .style_idx = c.style_idx,
                .payload = .{ .delta = dv },
            } else .{
                .ref = cr,
                .style_idx = c.style_idx,
                .payload = .{ .original = .{
                    .cell_type = c.cell_type,
                    .raw_value = c.raw_value,
                    .formula = c.formula,
                } },
            };
            try merged.append(allocator, mc);
            try seen.put(allocator, cr, {});
        }
    }

    // 2. Append delta-only cells (not matched to any existing cell).
    var dit = deltas.iterator();
    while (dit.next()) |entry| {
        if (seen.contains(entry.key_ptr.*)) continue;
        try merged.append(allocator, .{
            .ref = entry.key_ptr.*,
            .style_idx = null,
            .payload = .{ .delta = entry.value_ptr.* },
        });
    }

    // 3. Sort by (row, col) and group emit.
    std.sort.pdq(MergedCell, merged.items, {}, mergedLessThan);

    var i: usize = 0;
    var num_buf: [32]u8 = undefined;
    while (i < merged.items.len) {
        const row_idx = merged.items[i].ref.row;
        var j = i;
        while (j < merged.items.len and merged.items[j].ref.row == row_idx) : (j += 1) {}

        // <row r="N">
        try out.appendSlice(allocator, "<row r=\"");
        try out.appendSlice(allocator, try std.fmt.bufPrint(&num_buf, "{d}", .{row_idx}));
        try out.appendSlice(allocator, "\">");

        for (merged.items[i..j]) |mc| try emitCell(allocator, out, mc);

        try out.appendSlice(allocator, "</row>");
        i = j;
    }
}

fn emitCell(allocator: Allocator, out: *std.ArrayList(u8), mc: MergedCell) Error!void {
    var ref_buf: [16]u8 = undefined;
    const ref_str = formatA1Ref(&ref_buf, mc.ref);

    try out.appendSlice(allocator, "<c r=\"");
    try out.appendSlice(allocator, ref_str);
    try out.appendSlice(allocator, "\"");

    if (mc.style_idx) |s| {
        var s_buf: [16]u8 = undefined;
        try out.appendSlice(allocator, " s=\"");
        try out.appendSlice(allocator, try std.fmt.bufPrint(&s_buf, "{d}", .{s}));
        try out.appendSlice(allocator, "\"");
    }

    switch (mc.payload) {
        .original => |orig| {
            if (cellTypeAttr(orig.cell_type)) |t_attr| {
                try out.appendSlice(allocator, " t=\"");
                try out.appendSlice(allocator, t_attr);
                try out.appendSlice(allocator, "\"");
            }
            if (orig.raw_value == null and orig.formula == null) {
                try out.appendSlice(allocator, "/>");
                return;
            }
            try out.appendSlice(allocator, ">");
            if (orig.formula) |f| {
                try out.appendSlice(allocator, "<f>");
                try out.appendSlice(allocator, f);
                try out.appendSlice(allocator, "</f>");
            }
            if (orig.raw_value) |v| {
                if (orig.cell_type == .inline_string) {
                    try out.appendSlice(allocator, "<is><t>");
                    try out.appendSlice(allocator, v);
                    try out.appendSlice(allocator, "</t></is>");
                } else {
                    try out.appendSlice(allocator, "<v>");
                    try out.appendSlice(allocator, v);
                    try out.appendSlice(allocator, "</v>");
                }
            }
            try out.appendSlice(allocator, "</c>");
        },
        .delta => |dv| switch (dv) {
            .blank => {
                try out.appendSlice(allocator, "/>");
            },
            .number => |n| {
                try out.appendSlice(allocator, "><v>");
                var nbuf: [64]u8 = undefined;
                try out.appendSlice(allocator, try std.fmt.bufPrint(&nbuf, "{d}", .{n}));
                try out.appendSlice(allocator, "</v></c>");
            },
            .boolean => |b| {
                try out.appendSlice(allocator, " t=\"b\"><v>");
                try out.appendSlice(allocator, if (b) "1" else "0");
                try out.appendSlice(allocator, "</v></c>");
            },
            .string => |s| {
                try out.appendSlice(allocator, " t=\"inlineStr\"><is><t");
                // Preserve leading/trailing whitespace per OOXML.
                if (s.len > 0 and (s[0] == ' ' or s[s.len - 1] == ' ')) {
                    try out.appendSlice(allocator, " xml:space=\"preserve\"");
                }
                try out.appendSlice(allocator, ">");
                try appendXmlEscapedText(allocator, out, s);
                try out.appendSlice(allocator, "</t></is></c>");
            },
            .formula => |f| {
                // No cached value — Excel recalcs on open. Future iter
                // can stash a computed result inside `<v>` once a
                // formula evaluator (Tier D1) lands.
                try out.appendSlice(allocator, "><f>");
                try appendXmlEscapedText(allocator, out, f);
                try out.appendSlice(allocator, "</f></c>");
            },
        },
    }
}

fn cellTypeAttr(t: sheet_xml_mod.CellType) ?[]const u8 {
    return switch (t) {
        .number => null, // OOXML default; omit attribute
        .shared_string => "s",
        .boolean => "b",
        .formula_string => "str",
        .inline_string => "inlineStr",
        .error_value => "e",
        .date => "d",
    };
}

pub const Worksheet = struct {
    /// Back-pointer set lazily by `Workbook.sheet(idx)` (the slot table
    /// is allocated before the `Workbook` exists, so we patch in on
    /// first observation rather than at construction).
    workbook: *Workbook,
    sheet_idx: u32,

    /// Lazy typed view of the sheet part. `null` until first access.
    parsed: ?sheet_xml_mod.SheetXml,
    /// Cached resolved part name (e.g. "xl/worksheets/sheet1.xml").
    resolved_part_name: ?[]const u8,

    /// Pending mutations (B1 iter-wb-4 m1). Keyed by `CellRef`; the
    /// last `setCell` for a given ref wins. Empty after `Workbook.save`.
    deltas: std.AutoHashMapUnmanaged(CellRef, CellValue) = .{},

    pub fn deinit(self: *Worksheet, allocator: Allocator) void {
        if (self.parsed) |*p| {
            var view = p.*;
            view.deinit(allocator);
        }
        if (self.resolved_part_name) |part_name| allocator.free(part_name);
        freeDeltaStrings(allocator, &self.deltas);
        self.deltas.deinit(allocator);
    }

    /// Sheet name from the workbook's sheets list. Borrowed.
    pub fn name(self: *const Worksheet) []const u8 {
        return self.workbook.workbook.sheets[self.sheet_idx].name;
    }

    /// Workbook-assigned sheet ID (NOT the same as sheet_idx).
    pub fn sheetId(self: *const Worksheet) u32 {
        return self.workbook.workbook.sheets[self.sheet_idx].sheet_id;
    }

    pub fn state(self: *const Worksheet) workbook_xml_mod.SheetState {
        return self.workbook.workbook.sheets[self.sheet_idx].state;
    }

    /// Resolve the part name (e.g. "xl/worksheets/sheet1.xml") and
    /// parse the sheet XML if not already cached. Returns a const
    /// pointer to the cached view.
    pub fn ensureParsed(self: *Worksheet) Error!*const sheet_xml_mod.SheetXml {
        if (self.parsed != null) return &self.parsed.?;

        const wb = self.workbook;
        const r_id = wb.workbook.sheets[self.sheet_idx].r_id;
        if (r_id.len == 0) return Error.MissingRelationship;

        const wb_rels = wb.store.rels("xl/workbook.xml");
        var resolved: ?[]const u8 = null;
        for (wb_rels) |rel| {
            if (std.mem.eql(u8, rel.id, r_id)) {
                resolved = try wb.store.resolve("xl/workbook.xml", rel.target);
                break;
            }
        }
        const part_name = resolved orelse return Error.MissingRelationship;
        // Dupe so `resolved_part_name` lifetime is bound to Worksheet,
        // not to PartStore's arena (PartStore.resolve allocates into
        // its arena; safe to drop the dup if we trust the arena —
        // but explicit ownership is clearer).
        const owned = try wb.allocator.dupe(u8, part_name);
        errdefer wb.allocator.free(owned);
        self.resolved_part_name = owned;

        const part = wb.store.part(part_name) orelse return Error.MissingSheetPart;
        self.parsed = try sheet_xml_mod.parse(wb.allocator, part.bytes);
        return &self.parsed.?;
    }

    pub fn dimension(self: *Worksheet) Error!?sheet_xml_mod.Dimension {
        const view = try self.ensureParsed();
        return view.dimension;
    }

    pub fn rows(self: *Worksheet) Error![]const sheet_xml_mod.Row {
        const view = try self.ensureParsed();
        return view.rows;
    }

    pub fn merges(self: *Worksheet) Error![]const sheet_xml_mod.MergeRange {
        const view = try self.ensureParsed();
        return view.merges;
    }

    pub fn hyperlinks(self: *Worksheet) Error![]const sheet_xml_mod.Hyperlink {
        const view = try self.ensureParsed();
        return view.hyperlinks;
    }

    pub fn validations(self: *Worksheet) Error![]const sheet_xml_mod.DataValidation {
        const view = try self.ensureParsed();
        return view.validations;
    }

    pub fn conditionalFormats(self: *Worksheet) Error![]const sheet_xml_mod.ConditionalFormat {
        const view = try self.ensureParsed();
        return view.conditional_formats;
    }

    pub fn freezePane(self: *Worksheet) Error!?sheet_xml_mod.FreezePane {
        const view = try self.ensureParsed();
        return view.freeze;
    }

    /// Find a cell by its A1 reference (e.g. "A1", "B7"). Linear scan
    /// over the parsed rows/cells — sufficient for v1 read-only use.
    /// Match is case-insensitive on the column letters; row part is
    /// strict decimal. Returns `null` when no cell matches.
    ///
    /// Matched cells are returned by-value (small struct of borrowed
    /// slices); the underlying SheetXml owns the storage, so the
    /// returned Cell is valid for the Workbook's lifetime.
    pub fn cellByRef(self: *Worksheet, ref: []const u8) Error!?sheet_xml_mod.Cell {
        assert(ref.len > 0);
        const view = try self.ensureParsed();
        for (view.rows) |row| {
            for (row.cells) |c| {
                if (eqlAsciiIgnoreCase(c.ref, ref)) return c;
            }
        }
        return null;
    }

    /// Stage a mutation for cell at A1 ref `ref`. Persisted by
    /// `Workbook.save`. The last `setCell` call for a given ref wins.
    /// Numeric / boolean / blank values pass through by-value. String
    /// and formula values are duped into the Workbook allocator
    /// (caller can free the input slice as soon as `setCell` returns).
    ///
    /// String + formula inputs are validated against XML 1.0 —
    /// control bytes other than \t, \n, \r are rejected with
    /// `error.MalformedXml` to prevent emitting unparseable XML.
    pub fn setCell(self: *Worksheet, ref: []const u8, value: CellValue) Error!void {
        assert(ref.len > 0);
        const cr = try parseA1Ref(ref);
        const a = self.workbook.allocator;

        // Free any previous heap allocation for this ref so a
        // string/formula overwrite doesn't leak.
        if (self.deltas.get(cr)) |prev| {
            switch (prev) {
                .string => |s| a.free(s),
                .formula => |f| a.free(f),
                else => {},
            }
        }

        const stored: CellValue = switch (value) {
            .string => |s| blk: {
                if (!isXmlSafeText(s)) return error.MalformedXml;
                break :blk .{ .string = try a.dupe(u8, s) };
            },
            .formula => |f| blk: {
                if (!isXmlSafeText(f)) return error.MalformedXml;
                break :blk .{ .formula = try a.dupe(u8, f) };
            },
            else => value,
        };
        errdefer switch (stored) {
            .string => |s| a.free(s),
            .formula => |f| a.free(f),
            else => {},
        };

        try self.deltas.put(a, cr, stored);
    }
};

/// XML 1.0 §2.2: Char ::= #x9 | #xA | #xD | [#x20-#xD7FF] | …
/// Reject ASCII control bytes outside the allowed three. Bytes ≥ 0x80
/// pass through without interpretation — the input must already be
/// well-formed UTF-8.
fn isXmlSafeText(s: []const u8) bool {
    for (s) |b| {
        if (b < 0x20 and b != 0x09 and b != 0x0A and b != 0x0D) return false;
    }
    return true;
}

/// Free any string / formula allocations stashed in `deltas`. Called
/// both before `clearAndFree` (post-save) and `deinit` (Worksheet
/// teardown).
fn freeDeltaStrings(allocator: Allocator, deltas: *std.AutoHashMapUnmanaged(CellRef, CellValue)) void {
    var it = deltas.valueIterator();
    while (it.next()) |v| {
        switch (v.*) {
            .string => |s| allocator.free(s),
            .formula => |f| allocator.free(f),
            else => {},
        }
    }
}

/// Escape XML text content (`<`, `>`, `&`) into `out`. Quote chars
/// pass through — this helper is for ELEMENT-content escaping, not
/// attribute-value escaping. Caller pre-validates that bytes are
/// XML-1.0-safe via `isXmlSafeText`.
fn appendXmlEscapedText(allocator: Allocator, out: *std.ArrayList(u8), text: []const u8) Error!void {
    for (text) |b| {
        switch (b) {
            '<' => try out.appendSlice(allocator, "&lt;"),
            '>' => try out.appendSlice(allocator, "&gt;"),
            '&' => try out.appendSlice(allocator, "&amp;"),
            else => try out.append(allocator, b),
        }
    }
}

/// ASCII-case-insensitive equality. OOXML cell refs are ASCII letters
/// + decimal digits, so a Unicode-aware fold is unnecessary here.
fn eqlAsciiIgnoreCase(a: []const u8, b: []const u8) bool {
    if (a.len != b.len) return false;
    for (a, b) |ca, cb| {
        if (toAsciiLower(ca) != toAsciiLower(cb)) return false;
    }
    return true;
}

fn toAsciiLower(c: u8) u8 {
    return if (c >= 'A' and c <= 'Z') c + 32 else c;
}

// ─── Tests ────────────────────────────────────────────────────────────

test "Workbook.open: minimal corpus fixture exposes sheets" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;

    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();

    try std.testing.expectEqual(@as(u32, 2), wb.sheetCount());

    const s0 = try wb.sheet(0);
    try std.testing.expectEqualStrings("Sheet1", s0.name());

    const s1 = try wb.sheet(1);
    try std.testing.expectEqualStrings("Sheet2", s1.name());

    const out_of_range = wb.sheet(2);
    try std.testing.expectError(Error.SheetIndexOutOfRange, out_of_range);
}

test "Workbook.sheetByName: case-sensitive lookup, null on miss" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;

    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();

    const found = try wb.sheetByName("Sheet1");
    try std.testing.expect(found != null);
    try std.testing.expectEqual(@as(u32, 0), found.?.sheet_idx);

    const wrong_case = try wb.sheetByName("sheet1");
    try std.testing.expect(wrong_case == null);

    const missing = try wb.sheetByName("NoSuch");
    try std.testing.expect(missing == null);
}

test "Worksheet.ensureParsed: lazy cells/rows materialise on access" {
    const path = "tests/corpus/openpyxl_guess_types.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;

    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();

    const s0 = try wb.sheet(0);
    const ws_rows = try s0.rows();
    try std.testing.expect(ws_rows.len > 0);

    // Re-fetch hits the cache — same slice address.
    const ws_rows_cached = try s0.rows();
    try std.testing.expect(ws_rows.ptr == ws_rows_cached.ptr);
}

test "Workbook.sst: optional, lazily parsed, cached" {
    const path = "tests/corpus/worldbank_catalog.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;

    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();

    const sst1 = try wb.sst();
    try std.testing.expect(sst1 != null);
    try std.testing.expect(sst1.?.entries.len > 0);

    const sst2 = try wb.sst();
    try std.testing.expect(sst1 == sst2); // same pointer (cache hit)
}

test "Workbook.styles: lazily parsed, returns non-null on a real fixture" {
    const path = "tests/corpus/worldbank_catalog.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;

    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();

    const st = try wb.styles();
    try std.testing.expect(st != null);
}

test "Workbook.definedNames: surfaces empty list on fixture without names" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;

    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();

    // frictionless emits `<definedNames/>` (self-closing). Should
    // surface as an empty slice without erroring.
    const names = wb.definedNames();
    try std.testing.expectEqual(@as(usize, 0), names.len);
}

test "Workbook.definedNamesGlobal / definedNamesForSheet: split by scope" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;

    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();

    const global = try wb.definedNamesGlobal(std.testing.allocator);
    defer std.testing.allocator.free(global);
    try std.testing.expectEqual(@as(usize, 0), global.len);

    const for_s0 = try wb.definedNamesForSheet(std.testing.allocator, 0);
    defer std.testing.allocator.free(for_s0);
    try std.testing.expectEqual(@as(usize, 0), for_s0.len);

    const out_of_range = wb.definedNamesForSheet(std.testing.allocator, 9);
    try std.testing.expectError(Error.SheetIndexOutOfRange, out_of_range);
}

test "Workbook.sstText: plain entry returns the raw slice; rich errors" {
    const path = "tests/corpus/worldbank_catalog.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;

    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();

    // The corpus fixture's first SST entry should be plain text.
    const first = try wb.sstText(0);
    try std.testing.expect(first != null);
    try std.testing.expect(first.?.len > 0);

    const sst_view = (try wb.sst()).?;
    const oor = wb.sstText(@intCast(sst_view.entries.len));
    try std.testing.expectError(Error.SstIndexOutOfRange, oor);
}

test "Worksheet.cellByRef: A1-ref lookup matches case-insensitively" {
    const path = "tests/corpus/openpyxl_guess_types.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;

    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();

    const s0 = try wb.sheet(0);

    // First row of openpyxl_guess_types has a cell at A1.
    const cell_a1 = try s0.cellByRef("A1");
    try std.testing.expect(cell_a1 != null);

    // Lowercase ref hits the same cell.
    const cell_lower = try s0.cellByRef("a1");
    try std.testing.expect(cell_lower != null);
    try std.testing.expectEqualStrings(cell_a1.?.ref, cell_lower.?.ref);

    // Out-of-range ref returns null.
    const cell_zz = try s0.cellByRef("ZZ9999");
    try std.testing.expect(cell_zz == null);
}

test "parseA1Ref: well-formed refs map to (row, col)" {
    try std.testing.expectEqual(CellRef{ .row = 1, .col = 1 }, try parseA1Ref("A1"));
    try std.testing.expectEqual(CellRef{ .row = 10, .col = 2 }, try parseA1Ref("B10"));
    try std.testing.expectEqual(CellRef{ .row = 1, .col = 27 }, try parseA1Ref("AA1"));
    try std.testing.expectEqual(CellRef{ .row = 1048576, .col = 16384 }, try parseA1Ref("XFD1048576"));
    // Lowercase OK
    try std.testing.expectEqual(CellRef{ .row = 7, .col = 4 }, try parseA1Ref("d7"));
}

test "parseA1Ref: malformed input rejected" {
    try std.testing.expectError(Error.InvalidCellRef, parseA1Ref(""));
    try std.testing.expectError(Error.InvalidCellRef, parseA1Ref("A"));
    try std.testing.expectError(Error.InvalidCellRef, parseA1Ref("1"));
    try std.testing.expectError(Error.InvalidCellRef, parseA1Ref("A0"));
    try std.testing.expectError(Error.InvalidCellRef, parseA1Ref("A09"));
    try std.testing.expectError(Error.InvalidCellRef, parseA1Ref("XFE1")); // col > 16384
    try std.testing.expectError(Error.InvalidCellRef, parseA1Ref("A1048577")); // row > 1048576
    try std.testing.expectError(Error.InvalidCellRef, parseA1Ref("A1B"));
}

test "formatA1Ref: round-trips" {
    var buf: [16]u8 = undefined;
    try std.testing.expectEqualStrings("A1", formatA1Ref(&buf, .{ .row = 1, .col = 1 }));
    try std.testing.expectEqualStrings("Z99", formatA1Ref(&buf, .{ .row = 99, .col = 26 }));
    try std.testing.expectEqualStrings("AA1", formatA1Ref(&buf, .{ .row = 1, .col = 27 }));
    try std.testing.expectEqualStrings("XFD1048576", formatA1Ref(&buf, .{ .row = 1048576, .col = 16384 }));
}

test "Workbook.setCell + save: round-trip a number through PartStore" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    // Stage a temp output path under .zig-cache (always writable in
    // CI). Random suffix so parallel test binaries don't collide.
    var tmp_buf: [256]u8 = undefined;
    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));
    const tmp_path = try std.fmt.bufPrint(&tmp_buf, ".zig-cache/test-wb-setcell-{d}.xlsx", .{prng.random().int(u32)});

    {
        var wb = try Workbook.open(std.testing.allocator, src_path);
        defer wb.deinit();
        const s0 = try wb.sheet(0);
        try s0.setCell("A1", .{ .number = 42 });
        try s0.setCell("B2", .{ .number = -3.14 });
        try wb.save(tmp_path);
    }
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    // Re-open and verify the cells round-tripped.
    var wb2 = try Workbook.open(std.testing.allocator, tmp_path);
    defer wb2.deinit();
    const s0 = try wb2.sheet(0);
    const a1 = try s0.cellByRef("A1");
    try std.testing.expect(a1 != null);
    try std.testing.expect(a1.?.cell_type == .number);
    try std.testing.expect(a1.?.raw_value != null);
    try std.testing.expectEqualStrings("42", a1.?.raw_value.?);

    const b2 = try s0.cellByRef("B2");
    try std.testing.expect(b2 != null);
    try std.testing.expect(b2.?.cell_type == .number);
    try std.testing.expect(b2.?.raw_value != null);
    // Zig's "{d}" on -3.14 emits "-3.14"; checking exact bytes.
    try std.testing.expectEqualStrings("-3.14", b2.?.raw_value.?);
}

test "Workbook.setCell + save: boolean and blank land typed correctly" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    var tmp_buf: [256]u8 = undefined;
    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));
    const tmp_path = try std.fmt.bufPrint(&tmp_buf, ".zig-cache/test-wb-bool-{d}.xlsx", .{prng.random().int(u32)});

    {
        var wb = try Workbook.open(std.testing.allocator, src_path);
        defer wb.deinit();
        const s0 = try wb.sheet(0);
        try s0.setCell("A1", .{ .boolean = true });
        try s0.setCell("B1", .{ .boolean = false });
        try s0.setCell("C1", .blank);
        try wb.save(tmp_path);
    }
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    var wb2 = try Workbook.open(std.testing.allocator, tmp_path);
    defer wb2.deinit();
    const s0 = try wb2.sheet(0);

    const a1 = try s0.cellByRef("A1");
    try std.testing.expect(a1 != null);
    try std.testing.expect(a1.?.cell_type == .boolean);
    try std.testing.expectEqualStrings("1", a1.?.raw_value.?);

    const b1 = try s0.cellByRef("B1");
    try std.testing.expect(b1 != null);
    try std.testing.expect(b1.?.cell_type == .boolean);
    try std.testing.expectEqualStrings("0", b1.?.raw_value.?);

    const c1 = try s0.cellByRef("C1");
    // A blank cell (`<c r="C1"/>`) is still scanned by sheet_xml's
    // parser; raw_value is null and cell_type stays at the default.
    try std.testing.expect(c1 != null);
    try std.testing.expect(c1.?.raw_value == null);
}

test "Workbook.setCell: invalid ref errors before saving" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    var wb = try Workbook.open(std.testing.allocator, src_path);
    defer wb.deinit();
    const s0 = try wb.sheet(0);

    try std.testing.expectError(Error.InvalidCellRef, s0.setCell("A0", .{ .number = 1 }));
    try std.testing.expectError(Error.InvalidCellRef, s0.setCell("XFE1", .{ .number = 1 }));
}

test "Workbook.setCell: string round-trips via inlineStr" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    var tmp_buf: [256]u8 = undefined;
    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));
    const tmp_path = try std.fmt.bufPrint(&tmp_buf, ".zig-cache/test-wb-string-{d}.xlsx", .{prng.random().int(u32)});

    {
        var wb = try Workbook.open(std.testing.allocator, src_path);
        defer wb.deinit();
        const s0 = try wb.sheet(0);
        // Simple string
        try s0.setCell("A1", .{ .string = "Hello, world!" });
        // String with XML-special chars — must escape
        try s0.setCell("B1", .{ .string = "<a> & \"foo\"" });
        // Whitespace-bracketed string — must emit xml:space="preserve"
        try s0.setCell("C1", .{ .string = "  spaced  " });
        try wb.save(tmp_path);
    }
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    var wb2 = try Workbook.open(std.testing.allocator, tmp_path);
    defer wb2.deinit();
    const s0 = try wb2.sheet(0);

    const a1 = try s0.cellByRef("A1");
    try std.testing.expect(a1 != null);
    try std.testing.expect(a1.?.cell_type == .inline_string);
    try std.testing.expectEqualStrings("Hello, world!", a1.?.raw_value.?);

    const b1 = try s0.cellByRef("B1");
    try std.testing.expect(b1 != null);
    try std.testing.expect(b1.?.cell_type == .inline_string);
    // raw_value carries the escaped form — `<` → `&lt;` etc. Accept
    // either; sheet_xml.parse doesn't decode entities, so what we
    // emitted is what we read back.
    try std.testing.expectEqualStrings("&lt;a&gt; &amp; \"foo\"", b1.?.raw_value.?);

    const c1 = try s0.cellByRef("C1");
    try std.testing.expect(c1 != null);
    try std.testing.expect(c1.?.cell_type == .inline_string);
    try std.testing.expectEqualStrings("  spaced  ", c1.?.raw_value.?);
}

test "Workbook.setCell: control bytes in string rejected before save" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    var wb = try Workbook.open(std.testing.allocator, src_path);
    defer wb.deinit();
    const s0 = try wb.sheet(0);

    // \x00, \x01, \x1F all forbidden by XML 1.0
    try std.testing.expectError(error.MalformedXml, s0.setCell("A1", .{ .string = "bad\x00null" }));
    try std.testing.expectError(error.MalformedXml, s0.setCell("A2", .{ .string = "ctrl\x01here" }));
    try std.testing.expectError(error.MalformedXml, s0.setCell("A3", .{ .string = "esc\x1Fhere" }));

    // \t \n \r are explicitly allowed
    try s0.setCell("B1", .{ .string = "tab\there" });
    try s0.setCell("B2", .{ .string = "lf\nhere" });
    try s0.setCell("B3", .{ .string = "cr\rhere" });

    // Bytes ≥ 0x80 (UTF-8 continuation) pass through unchecked.
    try s0.setCell("C1", .{ .string = "café" });
}

test "Workbook.setCell: string overwrite frees prior allocation" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    var wb = try Workbook.open(std.testing.allocator, src_path);
    defer wb.deinit();
    const s0 = try wb.sheet(0);

    // String → string, twice, then drop without saving. If the
    // overwrite path leaks, std.testing.allocator catches it at
    // wb.deinit.
    try s0.setCell("A1", .{ .string = "first" });
    try s0.setCell("A1", .{ .string = "second" });
    try s0.setCell("A1", .{ .string = "third" });
}

test "Workbook.setCell: formula round-trips with no cached value" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    var tmp_buf: [256]u8 = undefined;
    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));
    const tmp_path = try std.fmt.bufPrint(&tmp_buf, ".zig-cache/test-wb-formula-{d}.xlsx", .{prng.random().int(u32)});

    {
        var wb = try Workbook.open(std.testing.allocator, src_path);
        defer wb.deinit();
        const s0 = try wb.sheet(0);
        try s0.setCell("A1", .{ .formula = "SUM(B1:B10)" });
        // Formula with XML-special chars (e.g. comparison operators)
        try s0.setCell("A2", .{ .formula = "IF(B1<C1, \"low\", \"high\")" });
        try wb.save(tmp_path);
    }
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    var wb2 = try Workbook.open(std.testing.allocator, tmp_path);
    defer wb2.deinit();
    const s0 = try wb2.sheet(0);

    const a1 = try s0.cellByRef("A1");
    try std.testing.expect(a1 != null);
    try std.testing.expect(a1.?.formula != null);
    try std.testing.expectEqualStrings("SUM(B1:B10)", a1.?.formula.?);
    // No cached value — Excel recalcs on open.
    try std.testing.expect(a1.?.raw_value == null);

    const a2 = try s0.cellByRef("A2");
    try std.testing.expect(a2 != null);
    try std.testing.expect(a2.?.formula != null);
    // The `<` was XML-escaped on emit; raw form is what we read back.
    try std.testing.expectEqualStrings("IF(B1&lt;C1, \"low\", \"high\")", a2.?.formula.?);
}

test "Workbook.setCell: formula control bytes rejected; overwrite leak-free" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    var wb = try Workbook.open(std.testing.allocator, src_path);
    defer wb.deinit();
    const s0 = try wb.sheet(0);

    try std.testing.expectError(error.MalformedXml, s0.setCell("A1", .{ .formula = "BAD\x00FN()" }));

    // Formula → formula → string → number overwrites: the heap-owned
    // variants must release on transition. std.testing.allocator
    // catches any leak at wb.deinit.
    try s0.setCell("B1", .{ .formula = "1+1" });
    try s0.setCell("B1", .{ .formula = "2+2" });
    try s0.setCell("B1", .{ .string = "now a string" });
    try s0.setCell("B1", .{ .number = 42 });
    try s0.setCell("B1", .{ .formula = "back to formula" });
}
