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
} || workbook_xml_mod.Error || sheet_xml_mod.ParseError || sst_xml_mod.Error || styles_xml_mod.Error || store_mod.Error;

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
};

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

    pub fn deinit(self: *Worksheet, allocator: Allocator) void {
        if (self.parsed) |*p| {
            var view = p.*;
            view.deinit(allocator);
        }
        if (self.resolved_part_name) |part_name| allocator.free(part_name);
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
};

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
