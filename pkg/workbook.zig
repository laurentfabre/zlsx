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
const zlsx = @import("zlsx");

const PartStore = store_mod.PartStore;

const workbook_xml_mod = typed_parts.workbook_xml;
const sheet_xml_mod = typed_parts.sheet_xml;
const sst_xml_mod = typed_parts.sst_xml;
const styles_xml_mod = typed_parts.styles_xml;

pub const Error = error{
    MissingWorkbookPart,
    MissingSheetPart,
    MissingRelationship,
    /// Workbook lacks `xl/_rels/workbook.xml.rels` — required to
    /// register a freshly-created `xl/sharedStrings.xml` part. Surfaces
    /// only when an SST extension is requested against a workbook that
    /// has no rels file at all (extremely malformed input).
    MissingWorkbookRels,
    SheetIndexOutOfRange,
    SheetNotFound,
    SstIndexOutOfRange,
    SstEntryIsRich,
    /// Existing `xl/_rels/workbook.xml.rels` is missing the closing
    /// `</Relationships>` tag — refused rather than producing an
    /// unparseable relationships file when extending the SST.
    MalformedWorkbookRels,
    InvalidCellRef,
    NoSheetData,
    UnsupportedCellValue,
    /// `Workbook.renameSheet` rejected `new_name`: empty, exceeds the
    /// length cap, or contains a forbidden character (`: \ / ? * [ ]`),
    /// or is the case-insensitive reserved name "history".
    InvalidSheetName,
    /// `Workbook.renameSheet` rejected `new_name`: an existing sheet
    /// (other than `sheet_idx` itself) already uses that name (case-
    /// insensitive ASCII compare; see method docstring for the
    /// Unicode-fold caveat).
    SheetNameInUse,
    /// Internal invariant: the existing sheet name in `WorkbookXml`
    /// exceeds 128 bytes. OOXML-conformant inputs cannot trip this —
    /// surfaces only on hand-crafted / corrupted workbook.xml.
    InternalSheetNameTooLong,
    /// `Workbook.renameSheet` could not locate the target `<sheet>`
    /// element in the source `xl/workbook.xml` bytes. Surfaces only
    /// if the file mutated under us between parse and patch.
    SheetElementNotFound,
    /// `Workbook.fromBook(book, path)` opened `path` but the resulting
    /// sheet count disagreed with `book.sheets.len`. Typically a path-
    /// drift bug in the caller (wrong path passed, file renamed,
    /// etc.). v1 of `fromBook` is a re-open + sanity-check shim;
    /// future iters may share bytes via PartStore-from-bytes.
    SheetCountMismatch,
    WriteFailed,
} || workbook_xml_mod.Error || sheet_xml_mod.ParseError || sst_xml_mod.Error || styles_xml_mod.Error || store_mod.Error ||
    zlsx.formula_rewriter.Error ||
    std.fs.File.WriteError || std.fs.File.OpenError || std.fs.Dir.RenameError || std.fs.Dir.StatFileError;

/// Mutation primitive (B1 iter-wb-4). Strings emit as `inlineStr`
/// — cell-local text, no SST extension required. `shared_string`
/// values flow through `xl/sharedStrings.xml` (m4): the workbook's
/// SST is extended (or created) on save and the cell emits as
/// `<c t="s"><v>idx</v></c>`. Formulas emit as `<f>…</f>` with no
/// cached `<v>`, so Excel recalculates on open.
///
/// `string`, `shared_string`, and `formula` slices borrow for
/// `setCell`'s call only. The delta map duplicates bytes into the
/// Workbook allocator before returning, so the caller can free /
/// reuse the buffer as soon as `setCell` returns.
pub const CellValue = union(enum) {
    blank: void,
    number: f64,
    boolean: bool,
    string: []const u8,
    /// Plain text routed through the workbook's shared-string table.
    /// On save, the SST is extended (or created) with the unique new
    /// strings and the cell emits as `t="s"` + numeric `<v>` index.
    /// De-dup is by exact-byte equality against existing SST plain
    /// entries (post-decode) and against other `.shared_string`
    /// deltas in the same save. Rich-text entries are NOT considered
    /// for de-dup (rare in writes).
    shared_string: []const u8,
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

    /// Lazy-open variant. Same shape as `open` for v1 — sheets are
    /// already lazy-materialised on first `Worksheet.ensureParsed()`,
    /// so there's no behavioural difference yet. The split exists so
    /// callers (and the iter-wb-6 RSS gate) can pin to the future-
    /// correct symbol; later iters may add an SST-lazy / drawings-lazy
    /// strategy here without changing call sites.
    pub fn openLazy(allocator: Allocator, path: []const u8) Error!Workbook {
        assert(path.len > 0);
        return Workbook.open(allocator, path);
    }

    /// Construct a `Workbook` from an already-opened `PartStore`.
    /// Takes ownership of the store; `deinit` will tear it down.
    pub fn fromStore(allocator: Allocator, store: PartStore) Error!Workbook {
        var s = store;
        errdefer s.deinit();

        const wb_part = try s.part("xl/workbook.xml") orelse
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

    /// Promote an already-opened `zlsx.Book` to a `Workbook`. Caller
    /// passes the path that was originally used to open `book`.
    ///
    /// **v1 contract — re-reads the file.** Today this is a thin
    /// wrapper around `Workbook.open(alloc, path)` plus a sanity check
    /// that the resulting sheet count matches `book.sheets.len`. The
    /// "without re-reading the file" promise from the workbook-overlay
    /// plan needs a `PartStore`-from-bytes constructor (or PartStore /
    /// Book sharing the underlying mmap) — out of scope for this iter.
    /// Use `book` for the migration-time consistency check; it is
    /// borrowed and the caller retains ownership.
    ///
    /// Errors `SheetCountMismatch` if `book` and the freshly-opened
    /// Workbook disagree on sheet count — typically a sign of a path
    /// drift bug in the caller (passed the wrong path, file was
    /// renamed, etc.).
    pub fn fromBook(allocator: Allocator, book: *const zlsx.Book, path: []const u8) Error!Workbook {
        assert(path.len > 0);
        var wb = try Workbook.open(allocator, path);
        errdefer wb.deinit();

        if (wb.sheetCount() != book.sheets.len) return error.SheetCountMismatch;
        return wb;
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
        const part = try self.store.part("xl/sharedStrings.xml") orelse return null;
        self.sst_view = try sst_xml_mod.parse(self.allocator, part.bytes);
        return &self.sst_view.?;
    }

    /// Lazily-parsed `xl/styles.xml`. Returns `null` if absent.
    pub fn styles(self: *Workbook) Error!?*const styles_xml_mod.StylesXml {
        if (self.styles_view != null) return &self.styles_view.?;
        const part = try self.store.part("xl/styles.xml") orelse return null;
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
    /// m2: strings + formulas. m4: shared-string mode (`<c t="s">`).
    pub fn save(self: *Workbook, path: []const u8) Error!void {
        // Phase 1: SST extension. Walk every worksheet's deltas for
        // `.shared_string` values and build a single text → index
        // map covering new strings across all sheets. If any are
        // present, regenerate `xl/sharedStrings.xml` BEFORE per-sheet
        // emit (per-sheet emit needs the assigned indices).
        var sst_plan = try buildSstExtensionPlan(self);
        defer sst_plan.deinit(self.allocator);

        if (sst_plan.has_new_strings) {
            try applySstExtensionPlan(self, &sst_plan);
        }

        for (self.worksheets) |*ws| {
            if (ws.deltas.count() == 0) continue;
            _ = try ws.ensureParsed();
            const part_name = ws.resolved_part_name.?;
            const view = &ws.parsed.?;
            const source = blk: {
                const p = try self.store.part(part_name) orelse return error.MissingSheetPart;
                break :blk p.bytes;
            };

            const new_xml = try emitSheetWithDeltas(
                self.allocator,
                source,
                view,
                &ws.deltas,
                &sst_plan,
            );
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
        // Invalidate cached SST view — its leaves borrowed from the
        // pre-extension SST bytes which `replacePart` swapped out.
        if (sst_plan.has_new_strings) {
            if (self.sst_view) |*v| {
                var view = v.*;
                view.deinit(self.allocator);
                self.sst_view = null;
            }
        }
        try self.store.save(path);
    }

    /// Read-only predicate: does the workbook carry any pending
    /// mutation that has not yet been flushed via `save`?
    ///
    /// Returns `true` if EITHER:
    ///   1. any `Worksheet.deltas` map is non-empty (uncommitted
    ///      `setCell` calls), OR
    ///   2. the underlying `PartStore` carries any part override
    ///      (uncommitted `replacePart` / `addPart` — produced by
    ///      `renameSheet`, `rewriteAllFormulas`, etc.).
    ///
    /// Important caveat — `PartStore.save` does NOT clear overrides
    /// after writing, so this predicate returns `true` even
    /// immediately after a successful `Workbook.save` if any
    /// PartStore-level override was applied during the workbook's
    /// lifetime. Callers needing a freshly-clean dirty bit must
    /// re-open the workbook from disk after `save`. The post-save
    /// behaviour for delta-only mutations IS clean (deltas are
    /// cleared in-place by `save`).
    pub fn hasUnsavedChanges(self: *const Workbook) bool {
        for (self.worksheets) |*ws| {
            if (ws.deltas.count() > 0) return true;
        }
        return self.store.hasUnsavedChanges();
    }

    /// Apply a structural-edit rewrite to every formula in every
    /// sheet. Walks each worksheet, materializes its SheetXml, runs
    /// `zlsx.formula_rewriter.rewriteFormula` on each cell that has
    /// `formula != null`, then stages the rewritten text via
    /// `Worksheet.setCell(ref, .{ .formula = new })`. Returns the
    /// number of cells rewritten (cells whose rewrite produced
    /// byte-identical output are NOT counted and don't grow the
    /// delta map).
    ///
    /// **This rewrites formulas only.** Row/col edits applied here
    /// shift formula references but do NOT structurally move cells —
    /// that's a follow-up iter (`Workbook.insertRow` etc.). For
    /// `rename_sheet` the workflow is coherent: pair this call with
    /// a manual `xl/workbook.xml` `<sheet name=>` rewrite. (A
    /// `Workbook.renameSheet` convenience is a future iter.)
    pub fn rewriteAllFormulas(
        self: *Workbook,
        edit: zlsx.formula_rewriter.RewriteEdit,
    ) Error!u32 {
        var count: u32 = 0;
        const a = self.allocator;
        var sheet_idx: u32 = 0;
        while (sheet_idx < self.sheetCount()) : (sheet_idx += 1) {
            const ws = try self.sheet(sheet_idx);
            const view = try ws.ensureParsed();
            const ws_name = ws.name();

            // Collect (ref, new_text) pairs first so we don't mutate
            // the Worksheet's delta map while iterating its parsed
            // view's row/cell slices.
            const Pending = struct { ref: []const u8, text: []u8 };
            var pending: std.ArrayList(Pending) = .empty;
            defer {
                for (pending.items) |p| a.free(p.text);
                pending.deinit(a);
            }

            for (view.rows) |row| {
                for (row.cells) |c| {
                    const f = c.formula orelse continue;
                    if (f.len == 0) continue;
                    const ctx = zlsx.formula_rewriter.RewriteContext{
                        .on_sheet = ws_name,
                        .target_sheet = null,
                        .edit = edit,
                    };
                    const rewritten = try zlsx.formula_rewriter.rewriteFormula(a, f, ctx);
                    if (std.mem.eql(u8, rewritten, f)) {
                        a.free(rewritten);
                        continue;
                    }
                    errdefer a.free(rewritten);
                    try pending.append(a, .{ .ref = c.ref, .text = rewritten });
                }
            }

            // Stage the deltas. `setCell` dupes the formula text
            // into its own allocation, so freeing `pending.items[i].text`
            // in the defer above is correct.
            for (pending.items) |p| {
                try ws.setCell(p.ref, .{ .formula = p.text });
                count += 1;
            }
        }
        return count;
    }

    /// Rename sheet at `sheet_idx` to `new_name`. Composes three steps
    /// atomically (in error semantics — partial work is left only on
    /// post-rewrite failures, see below):
    ///
    /// 1. Validate `new_name` per Excel rules (length, forbidden chars,
    ///    "history" reserved, no duplicate of any other sheet name).
    /// 2. Rewrite every formula in every sheet via
    ///    `rewriteAllFormulas(.{ .rename_sheet = ... })`. Cross-sheet
    ///    references targeting `old_name` get retargeted to `new_name`.
    /// 3. Patch `xl/workbook.xml` so the `<sheet name="OLD" .../>`
    ///    element for `sheet_idx` carries the new (XML-escaped) name.
    /// 4. Re-parse the in-memory `WorkbookXml` view from the freshly-
    ///    patched bytes so subsequent `wb.sheet(i).name()` returns the
    ///    new value without a `deinit + open` round-trip.
    ///
    /// **Lifecycle.** Step 2 stages formula deltas; they're persisted
    /// only by `Workbook.save`. The rewritten formulas live in each
    /// Worksheet's `deltas` map, NOT in its cached `parsed` view, so
    /// no `parsed = null` invalidation is required here. Caller still
    /// must call `save` to commit to disk.
    ///
    /// **Length cap.** v1 enforces a UTF-8 byte length of 1..127. Excel
    /// proper limits sheet names to 31 *characters* (Unicode codepoints,
    /// not bytes). For ASCII inputs the two coincide; for Unicode the
    /// byte cap is conservative — a follow-up iter can wire in
    /// `src/unicode/casefold.zig` for full character-count semantics.
    ///
    /// **Case folding.** Duplicate-name detection uses ASCII case-fold
    /// (a..z ↔ A..Z) only; "Sheet1" and "ŠHEET1" with non-ASCII letters
    /// fold differently than Excel does. Same follow-up iter applies.
    ///
    /// **Defined names.** Sheet-qualified `<definedName>` formulas
    /// (`Sheet2!$A$1` etc.) are NOT rewritten by this iter — only
    /// per-cell formulas via `rewriteAllFormulas`. Hyperlink targets
    /// pointing at the renamed sheet are likewise unaltered. A future
    /// iter (`m3-defnames-hyperlinks`) covers both.
    pub fn renameSheet(self: *Workbook, sheet_idx: u32, new_name: []const u8) Error!void {
        if (sheet_idx >= self.sheetCount()) return error.SheetIndexOutOfRange;
        try validateSheetName(new_name);
        try self.assertSheetNameAvailable(sheet_idx, new_name);

        // Capture old name into a stack copy: step 4 re-parses the
        // workbook view, freeing the arena that backs `sheets[i].name`.
        // We need the old bytes alive across step 2 (rewriter) and step
        // 3 (XML patch — the patch reads from the source bytes still
        // holding the old name).
        const old_name = self.workbook.sheets[sheet_idx].name;
        if (old_name.len == 0) return error.InternalSheetNameTooLong; // OOXML invariant
        if (old_name.len > 128) return error.InternalSheetNameTooLong;
        var old_buf: [128]u8 = undefined;
        @memcpy(old_buf[0..old_name.len], old_name);
        const old_name_owned = old_buf[0..old_name.len];

        // No-op rename: identical bytes. Skip rewriter (would error
        // .InvalidEdit on `old == new` is fine, but the cleaner contract
        // is "asking to rename to the current name is a successful
        // no-op").
        if (std.mem.eql(u8, old_name_owned, new_name)) return;

        _ = try self.rewriteAllFormulas(.{
            .rename_sheet = .{ .old = old_name_owned, .new = new_name },
        });

        try patchWorkbookXmlSheetName(self, sheet_idx, old_name_owned, new_name);
        try refreshWorkbookXmlView(self);

        // Postcondition: the in-memory view now reports the new name.
        assert(sheet_idx < self.workbook.sheets.len);
        assert(std.mem.eql(u8, self.workbook.sheets[sheet_idx].name, new_name));
    }

    /// Case-insensitive ASCII duplicate check. Skips the slot at
    /// `sheet_idx` itself so renaming a sheet to its own current name
    /// (modulo case) is permitted at this layer; a true no-op (exact
    /// byte match) short-circuits earlier in `renameSheet`.
    fn assertSheetNameAvailable(self: *const Workbook, sheet_idx: u32, new_name: []const u8) Error!void {
        assert(sheet_idx < self.workbook.sheets.len);
        assert(new_name.len > 0);
        for (self.workbook.sheets, 0..) |s, i| {
            if (i == sheet_idx) continue;
            if (asciiCaseInsensitiveEql(s.name, new_name)) return error.SheetNameInUse;
        }
    }
};

/// Validate a candidate sheet name per Excel's rules. v1 contract
/// (see `Workbook.renameSheet` docstring for the rationale):
///   - 1..127 bytes (UTF-8 byte length, NOT Unicode codepoints)
///   - none of `: \ / ? * [ ]`
///   - case-insensitive ASCII compare not equal to "history"
fn validateSheetName(name: []const u8) Error!void {
    if (name.len == 0) return error.InvalidSheetName;
    if (name.len > 127) return error.InvalidSheetName;
    for (name) |c| switch (c) {
        ':', '\\', '/', '?', '*', '[', ']' => return error.InvalidSheetName,
        else => {},
    };
    // Sheet-name reserved word. Excel rejects this case-insensitively.
    if (asciiCaseInsensitiveEql(name, "history")) return error.InvalidSheetName;
}

/// Lowercase-ASCII byte-equality. Non-ASCII bytes compare verbatim.
/// Documented limitation in `Workbook.renameSheet`.
fn asciiCaseInsensitiveEql(a: []const u8, b: []const u8) bool {
    if (a.len != b.len) return false;
    for (a, b) |x, y| {
        const xl: u8 = if (x >= 'A' and x <= 'Z') x + 32 else x;
        const yl: u8 = if (y >= 'A' and y <= 'Z') y + 32 else y;
        if (xl != yl) return false;
    }
    return true;
}

/// Walk the source `xl/workbook.xml` bytes, find the Nth `<sheet>`
/// element (1-based N == `sheet_idx + 1` since OOXML emits sheets in
/// document order), and rewrite its `name="..."` attribute to the
/// XML-escaped `new_name`. Re-emits the part via `store.replacePart`.
///
/// We match the specific element by index AND verify that its current
/// `name=` attribute equals `expected_old` — this is the pair
/// assertion: independent of whether `xl/workbook.xml` was emitted by
/// an external tool with surprising attribute ordering, we refuse to
/// rewrite an element whose old name doesn't match what we believe it
/// should be (`SheetElementNotFound`).
fn patchWorkbookXmlSheetName(
    self: *Workbook,
    sheet_idx: u32,
    expected_old: []const u8,
    new_name: []const u8,
) Error!void {
    assert(expected_old.len > 0);
    assert(new_name.len > 0);
    assert(sheet_idx < self.workbook.sheets.len);

    const part = try self.store.part("xl/workbook.xml") orelse return error.MissingWorkbookPart;
    const src = part.bytes;
    assert(src.len > 0);

    // Find the Nth `<sheet ` (note the trailing space — distinguishes
    // from `<sheets>`, `<sheetData>`, etc.) Also accept `<sheet/>`-
    // style self-close as a defensive fallback. We require an
    // attribute-bearing form for a name to be present, so primarily
    // match `<sheet ` and `<sheet\t` / `<sheet\n`.
    var search_from: usize = 0;
    var seen: u32 = 0;
    var elem_attrs_start: usize = 0;
    var elem_attrs_end: usize = 0;
    while (true) {
        const open = std.mem.indexOfPos(u8, src, search_from, "<sheet") orelse
            return error.SheetElementNotFound;
        const after = open + "<sheet".len;
        if (after >= src.len) return error.SheetElementNotFound;
        const boundary = src[after];
        // Distinguish `<sheet[ /\t\r\n]` from `<sheets`, `<sheetData`,
        // `<sheetView`, `<sheetFormatPr`, `<sheetPr`, `<sheetCalcPr`,
        // `<sheetProtection` and friends.
        const is_sheet_elem = switch (boundary) {
            ' ', '\t', '\r', '\n', '/' => true,
            else => false,
        };
        if (!is_sheet_elem) {
            search_from = after;
            continue;
        }
        // Find the closing `>` that terminates this open tag. Sheet
        // elements are leaves (`<sheet ... />` or `<sheet ...></sheet>`
        // with empty body); we just need the first `>` past `after`.
        const gt = std.mem.indexOfScalarPos(u8, src, after, '>') orelse
            return error.SheetElementNotFound;
        if (seen == sheet_idx) {
            elem_attrs_start = after;
            elem_attrs_end = if (gt > 0 and src[gt - 1] == '/') gt - 1 else gt;
            break;
        }
        seen += 1;
        search_from = gt + 1;
    }
    assert(elem_attrs_end >= elem_attrs_start);

    // Find `name="..."` (or `name='...'`) inside this element's
    // attribute span. Must be a real attribute, not a substring of
    // another attribute's value: we anchor on a preceding whitespace
    // OR the start of the attribute span.
    const attrs = src[elem_attrs_start..elem_attrs_end];
    const NameAttr = struct { value_start: usize, value_end: usize };
    const found: NameAttr = blk: {
        var i: usize = 0;
        while (i < attrs.len) {
            // Skip leading whitespace.
            while (i < attrs.len and (attrs[i] == ' ' or attrs[i] == '\t' or
                attrs[i] == '\r' or attrs[i] == '\n')) : (i += 1)
            {}
            if (i >= attrs.len) break;
            const key_start = i;
            while (i < attrs.len and attrs[i] != '=' and attrs[i] != ' ' and
                attrs[i] != '\t' and attrs[i] != '\r' and attrs[i] != '\n') : (i += 1)
            {}
            const key_end = i;
            // Skip = and any padding.
            while (i < attrs.len and (attrs[i] == ' ' or attrs[i] == '\t' or
                attrs[i] == '\r' or attrs[i] == '\n')) : (i += 1)
            {}
            if (i >= attrs.len or attrs[i] != '=') return error.SheetElementNotFound;
            i += 1;
            while (i < attrs.len and (attrs[i] == ' ' or attrs[i] == '\t' or
                attrs[i] == '\r' or attrs[i] == '\n')) : (i += 1)
            {}
            if (i >= attrs.len) return error.SheetElementNotFound;
            const quote = attrs[i];
            if (quote != '"' and quote != '\'') return error.SheetElementNotFound;
            i += 1;
            const val_start = i;
            while (i < attrs.len and attrs[i] != quote) : (i += 1) {}
            if (i >= attrs.len) return error.SheetElementNotFound;
            const val_end = i;
            i += 1; // past closing quote

            const key = attrs[key_start..key_end];
            if (std.mem.eql(u8, key, "name")) {
                break :blk .{
                    .value_start = elem_attrs_start + val_start,
                    .value_end = elem_attrs_start + val_end,
                };
            }
        }
        return error.SheetElementNotFound;
    };

    // Pair assertion: the element we found really IS the one we
    // intend to rewrite. The current name (still XML-escaped on the
    // wire — but for unescaped ASCII names like "Sheet1" the byte
    // comparison is correct) must match `expected_old`.
    if (!std.mem.eql(u8, src[found.value_start..found.value_end], expected_old)) {
        // Tolerate XML-escaped equivalents: if a name contains `&` or
        // `<` we'd see entities here; for ASCII-clean names this is
        // straight equality. If the wire form differs, it's not the
        // element we expected to rewrite.
        return error.SheetElementNotFound;
    }

    // Build the patched part: prefix + escaped new name + suffix.
    var out: std.ArrayList(u8) = .empty;
    defer out.deinit(self.allocator);
    try out.ensureTotalCapacity(self.allocator, src.len + new_name.len + 16);
    try out.appendSlice(self.allocator, src[0..found.value_start]);
    try appendXmlEscaped(self.allocator, &out, new_name);
    try out.appendSlice(self.allocator, src[found.value_end..]);

    try self.store.replacePart("xl/workbook.xml", out.items);
}

/// Append `s` to `out`, XML-escaping the five canonical entities
/// (`<`, `>`, `&`, `"`, `'`). Other bytes (including UTF-8
/// continuation bytes for non-ASCII characters) pass through verbatim.
fn appendXmlEscaped(allocator: Allocator, out: *std.ArrayList(u8), s: []const u8) !void {
    for (s) |c| switch (c) {
        '<' => try out.appendSlice(allocator, "&lt;"),
        '>' => try out.appendSlice(allocator, "&gt;"),
        '&' => try out.appendSlice(allocator, "&amp;"),
        '"' => try out.appendSlice(allocator, "&quot;"),
        '\'' => try out.appendSlice(allocator, "&apos;"),
        else => try out.append(allocator, c),
    };
}

/// Re-parse `xl/workbook.xml` from the (now-patched) PartStore bytes
/// and swap the typed view in place. The old view's arena is freed —
/// any external borrows of `wb.workbook.sheets[i].name` from before
/// `renameSheet` are invalidated. The contract says callers don't
/// hold those slices across mutation; this is the enforcement point.
fn refreshWorkbookXmlView(self: *Workbook) Error!void {
    const part = try self.store.part("xl/workbook.xml") orelse return error.MissingWorkbookPart;
    var fresh = try workbook_xml_mod.parse(self.allocator, part.bytes);
    errdefer fresh.deinit(self.allocator);

    // Length invariant: re-parse must agree on sheet count, otherwise
    // the slot table (worksheets[]) and the workbook view would drift.
    if (fresh.sheets.len != self.workbook.sheets.len) {
        return error.SheetCountMismatch;
    }

    self.workbook.deinit(self.allocator);
    self.workbook = fresh;
}

// ─── Emit helpers (iter-wb-4 m1) ─────────────────────────────────────

/// Splice a regenerated `<sheetData>...</sheetData>` block into the
/// source sheet XML. Everything outside `<sheetData>` is copied
/// byte-for-byte. Returns a fresh allocator-owned slice.
fn emitSheetWithDeltas(
    allocator: Allocator,
    source: []const u8,
    view: *const sheet_xml_mod.SheetXml,
    deltas: *const std.AutoHashMapUnmanaged(CellRef, CellValue),
    sst_plan: *const SstExtensionPlan,
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

    try emitSheetData(allocator, &out, view, deltas, sst_plan);

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
    sst_plan: *const SstExtensionPlan,
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

        for (merged.items[i..j]) |mc| try emitCell(allocator, out, mc, sst_plan);

        try out.appendSlice(allocator, "</row>");
        i = j;
    }
}

fn emitCell(
    allocator: Allocator,
    out: *std.ArrayList(u8),
    mc: MergedCell,
    sst_plan: *const SstExtensionPlan,
) Error!void {
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
            .shared_string => |s| {
                // Resolve the index assigned by the SST extension
                // pass. `getOrUnreachable` is safe here: every
                // `.shared_string` delta was registered into the
                // plan in `buildSstExtensionPlan` (precondition of
                // the save path).
                const idx = sst_plan.indexOf(s) orelse unreachable;
                try out.appendSlice(allocator, " t=\"s\"><v>");
                var ibuf: [16]u8 = undefined;
                try out.appendSlice(allocator, try std.fmt.bufPrint(&ibuf, "{d}", .{idx}));
                try out.appendSlice(allocator, "</v></c>");
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

// ─── SST extension (iter-wb-4 m4) ────────────────────────────────────

/// Plan for extending the workbook's shared-string table with new
/// strings collected from `.shared_string` deltas across every
/// worksheet's pending mutations.
///
/// Built upfront in `buildSstExtensionPlan` BEFORE per-sheet emit so
/// each `<c t="s">` knows its target index. `applySstExtensionPlan`
/// commits the plan to the `PartStore` (replacePart on existing SST,
/// addPart + workbook.xml.rels splice when SST is absent).
///
/// De-dup policy:
///   - linear scan against existing entries (decoded), then linear
///     scan against already-staged new strings.
///   - linear was chosen over a hashmap because (a) the typical
///     write workload stages a small handful of new strings while
///     the SST may carry thousands of existing entries; building a
///     hashmap of decoded existing entries up front is more work
///     than scanning per-new-string. (b) keeping the implementation
///     stdlib-only and trivially auditable matters more than constant-
///     factor speed at the SST sizes encountered in practice.
const ExistingMatch = struct {
    text: []const u8,
    index: u32,
};

const SstExtensionPlan = struct {
    /// True when at least one `.shared_string` delta required a fresh
    /// SST entry (i.e. the user-supplied text didn't already match an
    /// existing entry).
    has_new_strings: bool = false,
    /// Allocator owns: every entry of `new_strings` (duped on insert),
    /// the slice itself.
    new_strings: std.ArrayListUnmanaged([]const u8) = .empty,
    /// Side table: deltas whose text matched an existing SST entry.
    /// Allows `indexOf` to resolve those without rescanning the SST.
    /// Allocator owns each `text` slice.
    existing_matches: std.ArrayListUnmanaged(ExistingMatch) = .empty,
    /// Index of the FIRST new string within the regenerated SST. For
    /// an existing-SST workbook this is the existing entry count;
    /// for a freshly-created SST it's 0.
    base_index: u32 = 0,
    /// Tracks whether the SST part already existed at plan-build time.
    /// Drives the `replacePart` vs `addPart + workbook.xml.rels splice`
    /// branch in `applySstExtensionPlan`.
    sst_part_exists: bool = false,

    fn deinit(self: *SstExtensionPlan, allocator: Allocator) void {
        for (self.new_strings.items) |s| allocator.free(s);
        self.new_strings.deinit(allocator);
        for (self.existing_matches.items) |em| allocator.free(em.text);
        self.existing_matches.deinit(allocator);
        self.* = undefined;
    }

    /// Resolve the SST index for a (raw, unescaped) string staged
    /// via `setCell(.{ .shared_string = ... })`. Returns null if
    /// `s` was never staged into this plan (caller invariant: every
    /// `.shared_string` delta is registered before per-sheet emit).
    fn indexOf(self: *const SstExtensionPlan, s: []const u8) ?u32 {
        for (self.existing_matches.items) |em| {
            if (std.mem.eql(u8, em.text, s)) return em.index;
        }
        for (self.new_strings.items, 0..) |existing, i| {
            if (std.mem.eql(u8, existing, s)) {
                return self.base_index + @as(u32, @intCast(i));
            }
        }
        return null;
    }
};

/// Walk every worksheet's `.shared_string` deltas, de-dup against
/// the existing SST (when present) and against each other, and stage
/// the resulting unique-new-strings list into a plan. The plan owns
/// duplicates of every staged string; callers free via `plan.deinit`.
fn buildSstExtensionPlan(wb: *Workbook) Error!SstExtensionPlan {
    assert(@intFromPtr(wb) != 0);
    assert(@intFromPtr(wb.allocator.vtable) != 0);

    var plan: SstExtensionPlan = .{};
    errdefer plan.deinit(wb.allocator);

    // Quick scan: any `.shared_string` at all? Skips any work — and
    // crucially, skips parsing the SST — when the workbook has no
    // shared-string deltas pending.
    var any: bool = false;
    for (wb.worksheets) |*ws| {
        var it = ws.deltas.valueIterator();
        while (it.next()) |v| switch (v.*) {
            .shared_string => {
                any = true;
                break;
            },
            else => {},
        };
        if (any) break;
    }
    if (!any) return plan;

    // Resolve the existing SST's plain-entry count + decoded-text
    // slice for de-dup. Rich entries occupy indices but aren't
    // candidates for de-dup; a new string equal to a rich entry's
    // concatenated runs would still allocate a fresh `<si><t>...`.
    const existing_view = try wb.sst();
    if (existing_view) |view| {
        plan.sst_part_exists = true;
        plan.base_index = @intCast(view.entries.len);
    } else {
        plan.sst_part_exists = false;
        plan.base_index = 0;
    }

    // Pre-decode every existing plain entry once into an arena so
    // each new string compares against decoded text. Rich entries
    // get a sentinel empty slice (matched-equal would be wrong, but
    // an empty new string never equals a rich entry by construction
    // — `.plain` and `.rich` are disjoint).
    var decode_arena = std.heap.ArenaAllocator.init(wb.allocator);
    defer decode_arena.deinit();
    const da = decode_arena.allocator();

    var decoded_existing: [][]const u8 = &.{};
    if (existing_view) |view| {
        decoded_existing = try da.alloc([]const u8, view.entries.len);
        for (view.entries, 0..) |e, i| {
            decoded_existing[i] = switch (e) {
                .plain => |s| try sst_xml_mod.decodeText(da, s),
                .rich => "", // never matches a non-empty new string
            };
        }
    }

    // Walk deltas in worksheet order, then in iteration order. Order
    // is observable to test assertions, so document: "first occurrence
    // across (worksheet 0..N, iteration order) wins the lower index".
    for (wb.worksheets) |*ws| {
        var it = ws.deltas.valueIterator();
        while (it.next()) |v| {
            const s = switch (v.*) {
                .shared_string => |t| t,
                else => continue,
            };

            // De-dup against existing SST entries (decoded).
            var matched_existing: bool = false;
            for (decoded_existing) |de| {
                if (std.mem.eql(u8, de, s)) {
                    matched_existing = true;
                    break;
                }
            }
            if (matched_existing) continue;

            // De-dup against already-staged new strings.
            var dup_in_plan: bool = false;
            for (plan.new_strings.items) |existing| {
                if (std.mem.eql(u8, existing, s)) {
                    dup_in_plan = true;
                    break;
                }
            }
            if (dup_in_plan) continue;

            const owned = try wb.allocator.dupe(u8, s);
            errdefer wb.allocator.free(owned);
            try plan.new_strings.append(wb.allocator, owned);
        }
    }

    // The plan registered at least one new string only when at least
    // one delta failed to match an existing entry. If the user wrote
    // shared-strings that all already existed, has_new_strings stays
    // false and we skip SST regeneration entirely — but we still need
    // indexOf to resolve those existing entries. Patch base_index +
    // pre-load the matched existing entries so `indexOf` works.
    plan.has_new_strings = plan.new_strings.items.len > 0;

    // Whether or not we're regenerating, every `.shared_string` delta
    // must be reachable via plan.indexOf. For deltas that matched an
    // existing entry, register their (raw user) text → existing index
    // mapping by appending the user text under its existing index. We
    // do this by interleaving: re-walk deltas, for each shared_string
    // either it's already in plan.new_strings (just appended) or it
    // matched existing — we need to record the existing index.
    //
    // Simpler implementation: keep a parallel `existing_index_map`.
    // Done below via a second pass that uses the same de-dup logic
    // and populates `existing_match_index_for_each_new_string` which
    // is unused here; instead we extend `indexOf` to scan an
    // existing-match table. Defer that to a follow-up: in the common
    // case where a freshly-staged shared_string matches an existing
    // SST entry, we still want a valid emit path.

    // Fast path A: no new strings. Every delta matched an existing
    // SST entry — we skip regeneration. To keep emit correct, emit
    // a separate index lookup via a side-table keyed by raw user
    // text → existing index.
    if (!plan.has_new_strings and existing_view != null) {
        // Build the existing-match side table now.
        const view = existing_view.?;
        for (wb.worksheets) |*ws| {
            var it = ws.deltas.valueIterator();
            while (it.next()) |v| {
                const s = switch (v.*) {
                    .shared_string => |t| t,
                    else => continue,
                };
                // Already handled? linear scan keeps things simple.
                var already: bool = false;
                for (plan.existing_matches.items) |em| {
                    if (std.mem.eql(u8, em.text, s)) {
                        already = true;
                        break;
                    }
                }
                if (already) continue;
                // Find the matching existing index.
                var found_idx: u32 = std.math.maxInt(u32);
                for (decoded_existing, 0..) |de, i| {
                    if (std.mem.eql(u8, de, s)) {
                        found_idx = @intCast(i);
                        break;
                    }
                }
                assert(found_idx != std.math.maxInt(u32));
                const owned = try wb.allocator.dupe(u8, s);
                errdefer wb.allocator.free(owned);
                try plan.existing_matches.append(wb.allocator, .{ .text = owned, .index = found_idx });
            }
        }
        _ = view;
    }

    // Fast path B: there ARE new strings. We also need an existing-
    // match side table for any deltas whose text equals an existing
    // entry — those should resolve to the existing index, not a fresh
    // one. The de-dup loop above already skipped them from
    // plan.new_strings; populate the side table.
    if (plan.has_new_strings and existing_view != null) {
        for (wb.worksheets) |*ws| {
            var it = ws.deltas.valueIterator();
            while (it.next()) |v| {
                const s = switch (v.*) {
                    .shared_string => |t| t,
                    else => continue,
                };
                // If it's in plan.new_strings, it didn't match
                // existing — skip.
                var in_new: bool = false;
                for (plan.new_strings.items) |n| {
                    if (std.mem.eql(u8, n, s)) {
                        in_new = true;
                        break;
                    }
                }
                if (in_new) continue;

                // Skip if already registered in side table.
                var already: bool = false;
                for (plan.existing_matches.items) |em| {
                    if (std.mem.eql(u8, em.text, s)) {
                        already = true;
                        break;
                    }
                }
                if (already) continue;

                var found_idx: u32 = std.math.maxInt(u32);
                for (decoded_existing, 0..) |de, i| {
                    if (std.mem.eql(u8, de, s)) {
                        found_idx = @intCast(i);
                        break;
                    }
                }
                assert(found_idx != std.math.maxInt(u32));
                const owned = try wb.allocator.dupe(u8, s);
                errdefer wb.allocator.free(owned);
                try plan.existing_matches.append(wb.allocator, .{ .text = owned, .index = found_idx });
            }
        }
    }

    return plan;
}

/// Persist the SST extension plan to the PartStore. When the source
/// workbook had an existing `xl/sharedStrings.xml`, regenerate the
/// part's bytes (existing entries unchanged, new entries appended)
/// and `replacePart`. When absent, emit a fresh SST + register it via
/// `PartStore.addPart` + splice a `<Relationship>` into
/// `xl/_rels/workbook.xml.rels`.
fn applySstExtensionPlan(wb: *Workbook, plan: *const SstExtensionPlan) Error!void {
    assert(plan.has_new_strings);
    assert(plan.new_strings.items.len > 0);

    if (plan.sst_part_exists) {
        // Re-emit the SST part with the existing entries preserved
        // verbatim and the new strings appended.
        const existing_part = try wb.store.part("xl/sharedStrings.xml") orelse
            return Error.MissingWorkbookPart; // sst_part_exists invariant violated
        const new_xml = try emitSstXmlForExtension(
            wb.allocator,
            existing_part.bytes,
            plan.new_strings.items,
        );
        defer wb.allocator.free(new_xml);
        try wb.store.replacePart("xl/sharedStrings.xml", new_xml);
        return;
    }

    // Source had no SST. Emit a fresh one containing only the new
    // strings, register it as a new part with the correct content
    // type, then patch the workbook rels file.
    const fresh_xml = try emitFreshSstXml(wb.allocator, plan.new_strings.items);
    defer wb.allocator.free(fresh_xml);

    try wb.store.addPart(
        "xl/sharedStrings.xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
        fresh_xml,
    );

    // Splice a `<Relationship>` into `xl/_rels/workbook.xml.rels`.
    const rels_part = try wb.store.part("xl/_rels/workbook.xml.rels") orelse
        return Error.MissingWorkbookRels;
    const new_rels = try injectSstRelationship(wb.allocator, rels_part.bytes);
    defer wb.allocator.free(new_rels);
    try wb.store.replacePart("xl/_rels/workbook.xml.rels", new_rels);
}

/// Produce a regenerated `xl/sharedStrings.xml` with the original
/// entries preserved verbatim and one `<si><t>…</t></si>` per new
/// string appended. The `count` / `uniqueCount` attributes on `<sst>`
/// are rewritten to reflect the new totals; non-attribute markup
/// (xmlns, comments, PIs) is preserved as-is.
fn emitSstXmlForExtension(
    allocator: Allocator,
    src_xml: []const u8,
    new_strings: []const []const u8,
) Error![]u8 {
    assert(src_xml.len > 0);
    assert(new_strings.len > 0);

    // Locate `<sst …>` opening tag.
    const sst_open = std.mem.indexOf(u8, src_xml, "<sst") orelse
        return error.MalformedXml;
    const sst_open_gt = std.mem.indexOfScalarPos(u8, src_xml, sst_open, '>') orelse
        return error.MalformedXml;
    const is_self_closing = sst_open_gt > 0 and src_xml[sst_open_gt - 1] == '/';

    // Existing si count: parse uniqueCount attribute when present;
    // otherwise count `<si` opens in the body.
    const existing_si_count: u32 = blk: {
        const attrs = src_xml[sst_open .. sst_open_gt + 1];
        if (extractAttrValue(attrs, "uniqueCount")) |raw| {
            if (std.fmt.parseInt(u32, raw, 10)) |n| break :blk n else |_| {}
        }
        break :blk countSiOpens(src_xml);
    };
    const new_si_count: u32 = @intCast(new_strings.len);
    const total_si: u32 = existing_si_count + new_si_count;

    var out: std.ArrayList(u8) = .empty;
    errdefer out.deinit(allocator);
    try out.ensureTotalCapacity(allocator, src_xml.len + 64 * new_strings.len);

    // Copy bytes up to and INCLUDING `<sst`, then rewrite the
    // attribute blob with patched count/uniqueCount, then continue
    // from `>`.
    try out.appendSlice(allocator, src_xml[0 .. sst_open + "<sst".len]);

    // Walk the original attribute blob, replacing count/uniqueCount.
    const attr_start = sst_open + "<sst".len;
    const attr_end = sst_open_gt; // index of `>` (or `/>` slash)
    try writePatchedSstAttrs(
        allocator,
        &out,
        src_xml[attr_start..attr_end],
        total_si,
    );

    // If self-closing, transform into open form so we can append entries.
    if (is_self_closing) {
        try out.appendSlice(allocator, ">");
    } else {
        try out.appendSlice(allocator, ">");
    }

    if (is_self_closing) {
        // Source had `<sst …/>` with no body. Emit only the new entries
        // followed by a fresh `</sst>`.
        try appendNewSiEntries(allocator, &out, new_strings);
        try out.appendSlice(allocator, "</sst>");
        // Anything past the original `/>` is post-element trailing
        // bytes (rare, but preserve).
        if (sst_open_gt + 1 < src_xml.len) {
            try out.appendSlice(allocator, src_xml[sst_open_gt + 1 ..]);
        }
        return try out.toOwnedSlice(allocator);
    }

    // Normal form: copy body verbatim up to `</sst>`, then append
    // new entries, then `</sst>` + trailing.
    const body_start = sst_open_gt + 1;
    const close = std.mem.indexOfPos(u8, src_xml, body_start, "</sst>") orelse
        return error.MalformedXml;
    try out.appendSlice(allocator, src_xml[body_start..close]);
    try appendNewSiEntries(allocator, &out, new_strings);
    try out.appendSlice(allocator, src_xml[close..]);
    return try out.toOwnedSlice(allocator);
}

/// Build a complete `xl/sharedStrings.xml` from scratch. Used when
/// the source workbook had no SST part.
fn emitFreshSstXml(allocator: Allocator, new_strings: []const []const u8) Error![]u8 {
    assert(new_strings.len > 0);

    var out: std.ArrayList(u8) = .empty;
    errdefer out.deinit(allocator);
    try out.ensureTotalCapacity(allocator, 256 + 64 * new_strings.len);

    try out.appendSlice(allocator, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    try out.appendSlice(allocator, "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
    var nbuf: [32]u8 = undefined;
    try out.appendSlice(allocator, " count=\"");
    try out.appendSlice(allocator, try std.fmt.bufPrint(&nbuf, "{d}", .{new_strings.len}));
    try out.appendSlice(allocator, "\" uniqueCount=\"");
    try out.appendSlice(allocator, try std.fmt.bufPrint(&nbuf, "{d}", .{new_strings.len}));
    try out.appendSlice(allocator, "\">");
    try appendNewSiEntries(allocator, &out, new_strings);
    try out.appendSlice(allocator, "</sst>");
    return try out.toOwnedSlice(allocator);
}

/// Append one `<si><t>…</t></si>` per new string to `out`, with
/// `xml:space="preserve"` when the text has leading/trailing
/// whitespace that OOXML would otherwise strip.
fn appendNewSiEntries(
    allocator: Allocator,
    out: *std.ArrayList(u8),
    new_strings: []const []const u8,
) Error!void {
    for (new_strings) |s| {
        try out.appendSlice(allocator, "<si><t");
        if (sstNeedsXmlSpacePreserveLocal(s)) {
            try out.appendSlice(allocator, " xml:space=\"preserve\"");
        }
        try out.appendSlice(allocator, ">");
        try appendXmlEscapedText(allocator, out, s);
        try out.appendSlice(allocator, "</t></si>");
    }
}

/// Mirrors `src/xlsx.zig::sstNeedsXmlSpacePreserve`. Local copy keeps
/// `pkg/workbook.zig` independent of `src/`.
fn sstNeedsXmlSpacePreserveLocal(s: []const u8) bool {
    if (s.len == 0) return false;
    const lead = s[0];
    const trail = s[s.len - 1];
    return lead == ' ' or lead == '\t' or lead == '\n' or lead == '\r' or
        trail == ' ' or trail == '\t' or trail == '\n' or trail == '\r';
}

/// Walk the attribute blob between `<sst` and `>`, emitting it to
/// `out` with `count` and `uniqueCount` rewritten to `new_count`.
/// Attributes other than these two pass through byte-for-byte. If
/// neither attribute is present, both are appended.
fn writePatchedSstAttrs(
    allocator: Allocator,
    out: *std.ArrayList(u8),
    attrs: []const u8,
    new_count: u32,
) Error!void {
    var saw_count: bool = false;
    var saw_unique: bool = false;
    var i: usize = 0;
    while (i < attrs.len) {
        // Skip leading whitespace, but emit it.
        while (i < attrs.len and (attrs[i] == ' ' or attrs[i] == '\t' or
            attrs[i] == '\n' or attrs[i] == '\r'))
        {
            try out.append(allocator, attrs[i]);
            i += 1;
        }
        if (i >= attrs.len) break;
        // Slash (self-closing marker) or any other non-name char: emit + continue.
        if (attrs[i] == '/') {
            try out.append(allocator, attrs[i]);
            i += 1;
            continue;
        }
        // Identify attribute name = run of non-`=`, non-whitespace chars.
        const name_start = i;
        while (i < attrs.len and attrs[i] != '=' and attrs[i] != ' ' and
            attrs[i] != '\t' and attrs[i] != '\n' and attrs[i] != '\r' and
            attrs[i] != '/') : (i += 1)
        {}
        const name = attrs[name_start..i];
        // Skip whitespace before `=`.
        while (i < attrs.len and (attrs[i] == ' ' or attrs[i] == '\t' or
            attrs[i] == '\n' or attrs[i] == '\r')) : (i += 1)
        {}
        if (i >= attrs.len or attrs[i] != '=') {
            // Standalone token (e.g. trailing whitespace before `/>`).
            try out.appendSlice(allocator, name);
            continue;
        }
        i += 1; // past `=`
        // Skip whitespace, find quote.
        while (i < attrs.len and (attrs[i] == ' ' or attrs[i] == '\t' or
            attrs[i] == '\n' or attrs[i] == '\r')) : (i += 1)
        {}
        if (i >= attrs.len or (attrs[i] != '"' and attrs[i] != '\'')) {
            // Malformed attribute — emit verbatim, fall back to scanning to next whitespace.
            try out.appendSlice(allocator, name);
            try out.append(allocator, '=');
            continue;
        }
        const quote = attrs[i];
        const value_start = i + 1;
        const value_end = std.mem.indexOfScalarPos(u8, attrs, value_start, quote) orelse
            return error.MalformedXml;
        const raw_value = attrs[value_start..value_end];

        // Emit the attribute (rewriting count / uniqueCount).
        try out.append(allocator, ' ');
        if (std.mem.eql(u8, name, "count")) {
            saw_count = true;
            try writeCountAttr(allocator, out, "count", new_count);
        } else if (std.mem.eql(u8, name, "uniqueCount")) {
            saw_unique = true;
            try writeCountAttr(allocator, out, "uniqueCount", new_count);
        } else {
            try out.appendSlice(allocator, name);
            try out.append(allocator, '=');
            try out.append(allocator, quote);
            try out.appendSlice(allocator, raw_value);
            try out.append(allocator, quote);
        }
        i = value_end + 1;
    }
    if (!saw_count) try writeCountAttr(allocator, out, " count", new_count);
    if (!saw_unique) try writeCountAttr(allocator, out, " uniqueCount", new_count);
}

fn writeCountAttr(
    allocator: Allocator,
    out: *std.ArrayList(u8),
    name: []const u8,
    n: u32,
) Error!void {
    try out.appendSlice(allocator, name);
    try out.appendSlice(allocator, "=\"");
    var nbuf: [16]u8 = undefined;
    try out.appendSlice(allocator, try std.fmt.bufPrint(&nbuf, "{d}", .{n}));
    try out.append(allocator, '"');
}

/// Extract `name="value"` from an attribute blob (raw value, no
/// entity decoding). Returns null if `name` is absent. Boundary check
/// prevents `count` from matching `uniqueCount`.
fn extractAttrValue(blob: []const u8, name: []const u8) ?[]const u8 {
    assert(name.len > 0);
    var search_from: usize = 0;
    while (true) {
        const pos = std.mem.indexOfPos(u8, blob, search_from, name) orelse return null;
        const left_ok = pos == 0 or blob[pos - 1] == ' ' or blob[pos - 1] == '\t' or
            blob[pos - 1] == '\n' or blob[pos - 1] == '\r' or blob[pos - 1] == '<';
        const after = pos + name.len;
        if (after >= blob.len) return null;
        if (left_ok and blob[after] == '=') {
            const q_pos = after + 1;
            if (q_pos >= blob.len) return null;
            const quote = blob[q_pos];
            if (quote != '"' and quote != '\'') return null;
            const start = q_pos + 1;
            const end = std.mem.indexOfScalarPos(u8, blob, start, quote) orelse return null;
            return blob[start..end];
        }
        search_from = pos + 1;
    }
}

/// Count `<si` opens in `xml`. Used to recover the existing entry
/// count when `<sst>` has no `uniqueCount` attribute. Boundary check
/// keeps `<si` from matching `<silly` (the next char must be a tag
/// boundary).
fn countSiOpens(xml: []const u8) u32 {
    var n: u32 = 0;
    var i: usize = 0;
    while (std.mem.indexOfPos(u8, xml, i, "<si")) |pos| {
        const after = pos + 3;
        if (after >= xml.len) break;
        const c = xml[after];
        if (c == '>' or c == '/' or c == ' ' or c == '\t' or c == '\n' or c == '\r') {
            n += 1;
            i = after;
        } else {
            i = pos + 1;
        }
    }
    return n;
}

/// Splice a `<Relationship>` for `xl/sharedStrings.xml` into
/// `xl/_rels/workbook.xml.rels`. Picks an Id that doesn't collide
/// with existing `rIdN` values. No-op (returns the original bytes
/// duped) if a sharedStrings relationship already exists.
fn injectSstRelationship(allocator: Allocator, xml: []const u8) Error![]u8 {
    if (std.mem.indexOf(u8, xml, "/relationships/sharedStrings") != null) {
        return try allocator.dupe(u8, xml);
    }

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
        i = num_end + 1;
    }
    const new_id: u32 = max_id + 1;

    const close = std.mem.indexOf(u8, xml, "</Relationships>") orelse
        return error.MalformedWorkbookRels;

    var out: std.ArrayList(u8) = .empty;
    errdefer out.deinit(allocator);
    try out.ensureTotalCapacity(allocator, xml.len + 256);
    try out.appendSlice(allocator, xml[0..close]);
    try out.appendSlice(allocator, "<Relationship Id=\"rId");
    var nbuf: [16]u8 = undefined;
    try out.appendSlice(allocator, try std.fmt.bufPrint(&nbuf, "{d}", .{new_id}));
    try out.appendSlice(allocator, "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>");
    try out.appendSlice(allocator, xml[close..]);
    return try out.toOwnedSlice(allocator);
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
        // Free the prior part-name dupe if this Worksheet was
        // previously parsed-then-invalidated (e.g. by `Workbook.save`,
        // `renameSheet`, or test helpers that splice part bytes via
        // `PartStore.replacePart` and reset `parsed = null`). Without
        // this, every invalidate→re-parse cycle leaks the prior dupe.
        if (self.resolved_part_name) |prev| wb.allocator.free(prev);
        self.resolved_part_name = owned;

        const part = try wb.store.part(part_name) orelse return Error.MissingSheetPart;
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
        // string/formula/shared_string overwrite doesn't leak.
        if (self.deltas.get(cr)) |prev| {
            switch (prev) {
                .string => |s| a.free(s),
                .shared_string => |s| a.free(s),
                .formula => |f| a.free(f),
                else => {},
            }
        }

        const stored: CellValue = switch (value) {
            .string => |s| blk: {
                if (!isXmlSafeText(s)) return error.MalformedXml;
                break :blk .{ .string = try a.dupe(u8, s) };
            },
            .shared_string => |s| blk: {
                if (!isXmlSafeText(s)) return error.MalformedXml;
                break :blk .{ .shared_string = try a.dupe(u8, s) };
            },
            .formula => |f| blk: {
                if (!isXmlSafeText(f)) return error.MalformedXml;
                break :blk .{ .formula = try a.dupe(u8, f) };
            },
            else => value,
        };
        errdefer switch (stored) {
            .string => |s| a.free(s),
            .shared_string => |s| a.free(s),
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
            .shared_string => |s| allocator.free(s),
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

// ─── Test helpers ─────────────────────────────────────────────────────

/// Write a minimal SST-less .xlsx to `path` for testing the SST-
/// creation branch. Every part is STORED (compression method = 0)
/// so the file can be assembled without pulling in a deflate
/// dependency. Contents:
///   - `[Content_Types].xml` with no sharedStrings override
///   - `_rels/.rels` pointing to `xl/workbook.xml`
///   - `xl/workbook.xml` declaring a single sheet (rId1)
///   - `xl/_rels/workbook.xml.rels` with the sheet rel only
///   - `xl/worksheets/sheet1.xml` (empty `<sheetData/>`)
fn writeMinimalSstLessXlsx(allocator: Allocator, path: []const u8) !void {
    const Entry = struct { name: []const u8, body: []const u8 };

    const content_types =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" ++
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" ++
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" ++
        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" ++
        "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" ++
        "</Types>";
    const root_rels =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" ++
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" ++
        "</Relationships>";
    const workbook_xml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" ++
        "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>" ++
        "</workbook>";
    const workbook_rels =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" ++
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" ++
        "</Relationships>";
    const sheet_xml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" ++
        "<sheetData></sheetData>" ++
        "</worksheet>";

    const entries = [_]Entry{
        .{ .name = "[Content_Types].xml", .body = content_types },
        .{ .name = "_rels/.rels", .body = root_rels },
        .{ .name = "xl/workbook.xml", .body = workbook_xml },
        .{ .name = "xl/_rels/workbook.xml.rels", .body = workbook_rels },
        .{ .name = "xl/worksheets/sheet1.xml", .body = sheet_xml },
    };

    var buf: std.ArrayList(u8) = .empty;
    defer buf.deinit(allocator);

    const Lfh = struct { offset: u32, name: []const u8, body: []const u8, crc: u32 };
    var lfhs: std.ArrayList(Lfh) = .empty;
    defer lfhs.deinit(allocator);

    // Phase 1: write LFH + payload for each entry.
    for (entries) |e| {
        const off: u32 = @intCast(buf.items.len);
        const crc = std.hash.Crc32.hash(e.body);
        // LFH = 30 bytes + name + payload.
        var hdr: [30]u8 = undefined;
        std.mem.writeInt(u32, hdr[0..4], 0x04034b50, .little);
        std.mem.writeInt(u16, hdr[4..6], 20, .little); // version
        std.mem.writeInt(u16, hdr[6..8], 0, .little); // flags
        std.mem.writeInt(u16, hdr[8..10], 0, .little); // method = STORED
        std.mem.writeInt(u16, hdr[10..12], 0, .little); // mtime
        std.mem.writeInt(u16, hdr[12..14], 0, .little); // mdate
        std.mem.writeInt(u32, hdr[14..18], crc, .little);
        std.mem.writeInt(u32, hdr[18..22], @intCast(e.body.len), .little);
        std.mem.writeInt(u32, hdr[22..26], @intCast(e.body.len), .little);
        std.mem.writeInt(u16, hdr[26..28], @intCast(e.name.len), .little);
        std.mem.writeInt(u16, hdr[28..30], 0, .little); // extra len
        try buf.appendSlice(allocator, &hdr);
        try buf.appendSlice(allocator, e.name);
        try buf.appendSlice(allocator, e.body);
        try lfhs.append(allocator, .{ .offset = off, .name = e.name, .body = e.body, .crc = crc });
    }

    // Phase 2: central directory.
    const cd_off: u32 = @intCast(buf.items.len);
    for (lfhs.items) |l| {
        var cdfh: [46]u8 = undefined;
        std.mem.writeInt(u32, cdfh[0..4], 0x02014b50, .little);
        std.mem.writeInt(u16, cdfh[4..6], 20, .little);
        std.mem.writeInt(u16, cdfh[6..8], 20, .little);
        std.mem.writeInt(u16, cdfh[8..10], 0, .little);
        std.mem.writeInt(u16, cdfh[10..12], 0, .little);
        std.mem.writeInt(u16, cdfh[12..14], 0, .little);
        std.mem.writeInt(u16, cdfh[14..16], 0, .little);
        std.mem.writeInt(u32, cdfh[16..20], l.crc, .little);
        std.mem.writeInt(u32, cdfh[20..24], @intCast(l.body.len), .little);
        std.mem.writeInt(u32, cdfh[24..28], @intCast(l.body.len), .little);
        std.mem.writeInt(u16, cdfh[28..30], @intCast(l.name.len), .little);
        std.mem.writeInt(u16, cdfh[30..32], 0, .little);
        std.mem.writeInt(u16, cdfh[32..34], 0, .little);
        std.mem.writeInt(u16, cdfh[34..36], 0, .little);
        std.mem.writeInt(u16, cdfh[36..38], 0, .little);
        std.mem.writeInt(u32, cdfh[38..42], 0, .little);
        std.mem.writeInt(u32, cdfh[42..46], l.offset, .little);
        try buf.appendSlice(allocator, &cdfh);
        try buf.appendSlice(allocator, l.name);
    }
    const cd_size: u32 = @intCast(@as(u32, @intCast(buf.items.len)) - cd_off);

    // Phase 3: EOCD.
    var eocd: [22]u8 = undefined;
    std.mem.writeInt(u32, eocd[0..4], 0x06054b50, .little);
    std.mem.writeInt(u16, eocd[4..6], 0, .little);
    std.mem.writeInt(u16, eocd[6..8], 0, .little);
    std.mem.writeInt(u16, eocd[8..10], @intCast(lfhs.items.len), .little);
    std.mem.writeInt(u16, eocd[10..12], @intCast(lfhs.items.len), .little);
    std.mem.writeInt(u32, eocd[12..16], cd_size, .little);
    std.mem.writeInt(u32, eocd[16..20], cd_off, .little);
    std.mem.writeInt(u16, eocd[20..22], 0, .little);
    try buf.appendSlice(allocator, &eocd);

    // Write to disk.
    var f = try std.fs.cwd().createFile(path, .{});
    defer f.close();
    try f.writeAll(buf.items);
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

test "Workbook.setCell: shared_string round-trips on a fixture WITH existing SST" {
    const src_path = "tests/corpus/worldbank_catalog.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    var tmp_buf: [256]u8 = undefined;
    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));
    const tmp_path = try std.fmt.bufPrint(&tmp_buf, ".zig-cache/test-wb-sst-extend-{d}.xlsx", .{prng.random().int(u32)});

    const new_text = "zlsx-iter-wb-4-m4-sentinel-string";

    var pre_count: u32 = 0;
    {
        var wb = try Workbook.open(std.testing.allocator, src_path);
        defer wb.deinit();
        const sst_view = (try wb.sst()).?;
        pre_count = @intCast(sst_view.entries.len);
        try std.testing.expect(pre_count > 0);

        const s0 = try wb.sheet(0);
        try s0.setCell("Z999", .{ .shared_string = new_text });
        try wb.save(tmp_path);
    }
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    var wb2 = try Workbook.open(std.testing.allocator, tmp_path);
    defer wb2.deinit();

    // SST grew by exactly one and the new entry resolves to our text.
    const sst_view2 = (try wb2.sst()).?;
    try std.testing.expectEqual(@as(usize, pre_count + 1), sst_view2.entries.len);
    const tail_text = try wb2.sstText(pre_count);
    try std.testing.expect(tail_text != null);
    try std.testing.expectEqualStrings(new_text, tail_text.?);

    const s0 = try wb2.sheet(0);
    const c = try s0.cellByRef("Z999");
    try std.testing.expect(c != null);
    try std.testing.expect(c.?.cell_type == .shared_string);
    try std.testing.expect(c.?.raw_value != null);
    var ibuf: [16]u8 = undefined;
    const expected_idx_str = try std.fmt.bufPrint(&ibuf, "{d}", .{pre_count});
    try std.testing.expectEqualStrings(expected_idx_str, c.?.raw_value.?);
}

test "Workbook.setCell: shared_string creates SST on a workbook without one" {
    // None of the corpus fixtures lack an SST, so this test
    // synthesises a minimal SST-less xlsx in-memory (STORED entries
    // only — no deflate dependency) and writes it under .zig-cache.
    const alloc = std.testing.allocator;

    var tmp_buf: [256]u8 = undefined;
    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));
    const src_path = try std.fmt.bufPrint(&tmp_buf, ".zig-cache/test-wb-sstless-src-{d}.xlsx", .{prng.random().int(u32)});
    var tmp_buf2: [256]u8 = undefined;
    const tmp_path = try std.fmt.bufPrint(&tmp_buf2, ".zig-cache/test-wb-sstless-out-{d}.xlsx", .{prng.random().int(u32)});

    try writeMinimalSstLessXlsx(alloc, src_path);
    defer std.fs.cwd().deleteFile(src_path) catch {};

    const new_text = "fresh-sst-greeting";
    // Sanity: the synthetic source has no SST.
    {
        var wb_check = try Workbook.open(alloc, src_path);
        defer wb_check.deinit();
        const v = try wb_check.sst();
        try std.testing.expect(v == null);
    }

    {
        var wb = try Workbook.open(alloc, src_path);
        defer wb.deinit();
        const s0 = try wb.sheet(0);
        try s0.setCell("A1", .{ .shared_string = new_text });
        try wb.save(tmp_path);
    }
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    var wb2 = try Workbook.open(alloc, tmp_path);
    defer wb2.deinit();

    // SST part now exists with exactly one entry at index 0.
    const v2 = try wb2.sst();
    try std.testing.expect(v2 != null);
    try std.testing.expectEqual(@as(usize, 1), v2.?.entries.len);
    const t = try wb2.sstText(0);
    try std.testing.expect(t != null);
    try std.testing.expectEqualStrings(new_text, t.?);

    const s0 = try wb2.sheet(0);
    const c = try s0.cellByRef("A1");
    try std.testing.expect(c != null);
    try std.testing.expect(c.?.cell_type == .shared_string);
    try std.testing.expectEqualStrings("0", c.?.raw_value.?);
}

test "Workbook.setCell: shared_string de-dups identical text across cells" {
    const src_path = "tests/corpus/worldbank_catalog.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    var tmp_buf: [256]u8 = undefined;
    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));
    const tmp_path = try std.fmt.bufPrint(&tmp_buf, ".zig-cache/test-wb-sst-dedup-{d}.xlsx", .{prng.random().int(u32)});

    const new_text = "zlsx-dedup-target-string";

    var pre_count: u32 = 0;
    {
        var wb = try Workbook.open(std.testing.allocator, src_path);
        defer wb.deinit();
        pre_count = @intCast((try wb.sst()).?.entries.len);

        const s0 = try wb.sheet(0);
        try s0.setCell("Z998", .{ .shared_string = new_text });
        try s0.setCell("Z999", .{ .shared_string = new_text });
        try wb.save(tmp_path);
    }
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    var wb2 = try Workbook.open(std.testing.allocator, tmp_path);
    defer wb2.deinit();

    // SST grew by exactly ONE despite two cells writing the same text.
    const sst2 = (try wb2.sst()).?;
    try std.testing.expectEqual(@as(usize, pre_count + 1), sst2.entries.len);

    const s0 = try wb2.sheet(0);
    const c1 = try s0.cellByRef("Z998");
    const c2 = try s0.cellByRef("Z999");
    try std.testing.expect(c1 != null);
    try std.testing.expect(c2 != null);
    try std.testing.expect(c1.?.cell_type == .shared_string);
    try std.testing.expect(c2.?.cell_type == .shared_string);
    // Both cells reference the SAME index.
    try std.testing.expectEqualStrings(c1.?.raw_value.?, c2.?.raw_value.?);
    var ibuf: [16]u8 = undefined;
    const expected = try std.fmt.bufPrint(&ibuf, "{d}", .{pre_count});
    try std.testing.expectEqualStrings(expected, c1.?.raw_value.?);
}

test "Workbook.setCell: mixed inlineStr + shared_string in one save" {
    const src_path = "tests/corpus/worldbank_catalog.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    var tmp_buf: [256]u8 = undefined;
    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));
    const tmp_path = try std.fmt.bufPrint(&tmp_buf, ".zig-cache/test-wb-sst-mixed-{d}.xlsx", .{prng.random().int(u32)});

    const inline_text = "stays-inline-mixed-mode";
    const shared_text = "goes-to-sst-mixed-mode";

    var pre_count: u32 = 0;
    {
        var wb = try Workbook.open(std.testing.allocator, src_path);
        defer wb.deinit();
        pre_count = @intCast((try wb.sst()).?.entries.len);

        const s0 = try wb.sheet(0);
        try s0.setCell("Z997", .{ .string = inline_text });
        try s0.setCell("Z998", .{ .shared_string = shared_text });
        try wb.save(tmp_path);
    }
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    var wb2 = try Workbook.open(std.testing.allocator, tmp_path);
    defer wb2.deinit();
    const sst2 = (try wb2.sst()).?;
    try std.testing.expectEqual(@as(usize, pre_count + 1), sst2.entries.len);

    const s0 = try wb2.sheet(0);
    // Inline cell: t="inlineStr", raw_value carries the text directly.
    const c_inline = try s0.cellByRef("Z997");
    try std.testing.expect(c_inline != null);
    try std.testing.expect(c_inline.?.cell_type == .inline_string);
    try std.testing.expectEqualStrings(inline_text, c_inline.?.raw_value.?);

    // Shared-string cell: t="s", raw_value is the SST index.
    const c_shared = try s0.cellByRef("Z998");
    try std.testing.expect(c_shared != null);
    try std.testing.expect(c_shared.?.cell_type == .shared_string);
    var ibuf: [16]u8 = undefined;
    const expected_idx = try std.fmt.bufPrint(&ibuf, "{d}", .{pre_count});
    try std.testing.expectEqualStrings(expected_idx, c_shared.?.raw_value.?);
    // And the SST entry at that index resolves to our text.
    const tail = try wb2.sstText(pre_count);
    try std.testing.expectEqualStrings(shared_text, tail.?);
}

test "Workbook.fromBook: round-trip parity with Book.open on same path" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;

    var book = try zlsx.Book.open(std.testing.allocator, path);
    defer book.deinit();

    var wb = try Workbook.fromBook(std.testing.allocator, &book, path);
    defer wb.deinit();

    try std.testing.expectEqual(@as(usize, book.sheets.len), wb.sheetCount());
    var i: u32 = 0;
    while (i < wb.sheetCount()) : (i += 1) {
        const ws = try wb.sheet(i);
        try std.testing.expectEqualStrings(book.sheets[i].name, ws.name());
    }
}

test "Workbook.fromBook: independent lifetime — wb deinits before book" {
    const path = "tests/corpus/openpyxl_guess_types.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;

    var book = try zlsx.Book.open(std.testing.allocator, path);
    defer book.deinit();

    var wb = try Workbook.fromBook(std.testing.allocator, &book, path);
    // Tear down Workbook FIRST while book is still alive — must be
    // independent. std.testing.allocator catches any leak from a
    // mistaken shared-arena assumption.
    wb.deinit();
}

test "Workbook.fromBook: mismatched path errors SheetCountMismatch or opens cleanly" {
    const path_a = "tests/corpus/frictionless_2sheets.xlsx"; // 2 sheets
    const path_b = "tests/corpus/openpyxl_guess_types.xlsx"; // 1 sheet
    std.fs.cwd().access(path_a, .{}) catch return error.SkipZigTest;
    std.fs.cwd().access(path_b, .{}) catch return error.SkipZigTest;

    var book = try zlsx.Book.open(std.testing.allocator, path_a);
    defer book.deinit();

    // Pass `book` opened from path_a, but path_b — sheet counts differ
    // (2 vs 1) so fromBook surfaces the drift cleanly rather than
    // returning an inconsistent Workbook.
    const result = Workbook.fromBook(std.testing.allocator, &book, path_b);
    try std.testing.expectError(Error.SheetCountMismatch, result);
}

test "Workbook.rewriteAllFormulas: insert_rows shifts every formula's row refs in place" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    var tmp_buf: [256]u8 = undefined;
    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));
    const tmp_path = try std.fmt.bufPrint(&tmp_buf, ".zig-cache/test-rewrite-all-{d}.xlsx", .{prng.random().int(u32)});

    // Stage some formulas first, save, then re-open and rewrite.
    {
        var wb = try Workbook.open(std.testing.allocator, src_path);
        defer wb.deinit();
        const s0 = try wb.sheet(0);
        try s0.setCell("A1", .{ .formula = "SUM(B5:B10)" });
        try s0.setCell("B2", .{ .formula = "B7+1" });
        try s0.setCell("C3", .{ .formula = "B2*B5" }); // already-rewritten ref + a target ref
        try wb.save(tmp_path);
    }
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    var wb = try Workbook.open(std.testing.allocator, tmp_path);
    defer wb.deinit();

    // Insert 1 row at row 4 — every ref to row >= 4 shifts +1.
    const count = try wb.rewriteAllFormulas(.{
        .insert_rows = .{ .at = 4, .count = 1 },
    });
    // A1's SUM(B5:B10) → SUM(B6:B11) — 1 rewrite
    // B2's B7+1 → B8+1 — 1 rewrite
    // C3's B2*B5 → B2*B6 (B2 unchanged, B5 → B6) — 1 rewrite
    try std.testing.expectEqual(@as(u32, 3), count);

    var tmp2_buf: [256]u8 = undefined;
    const tmp2_path = try std.fmt.bufPrint(&tmp2_buf, ".zig-cache/test-rewrite-all-out-{d}.xlsx", .{prng.random().int(u32)});
    try wb.save(tmp2_path);
    defer std.fs.cwd().deleteFile(tmp2_path) catch {};

    var wb2 = try Workbook.open(std.testing.allocator, tmp2_path);
    defer wb2.deinit();
    const s0 = try wb2.sheet(0);
    const a1 = (try s0.cellByRef("A1")).?;
    try std.testing.expectEqualStrings("SUM(B6:B11)", a1.formula.?);
    const b2 = (try s0.cellByRef("B2")).?;
    try std.testing.expectEqualStrings("B8+1", b2.formula.?);
    const c3 = (try s0.cellByRef("C3")).?;
    try std.testing.expectEqualStrings("B2*B6", c3.formula.?);
}

test "Workbook.rewriteAllFormulas: no-op count == 0 on a workbook without formulas" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    var wb = try Workbook.open(std.testing.allocator, src_path);
    defer wb.deinit();

    // Pristine fixture has no <f> cells — nothing to rewrite.
    const count = try wb.rewriteAllFormulas(.{
        .insert_rows = .{ .at = 1, .count = 1 },
    });
    try std.testing.expectEqual(@as(u32, 0), count);
}

test "Workbook.rewriteAllFormulas: rename_sheet rewrites quoted sheet qualifiers" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    var tmp_buf: [256]u8 = undefined;
    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));
    const tmp_path = try std.fmt.bufPrint(&tmp_buf, ".zig-cache/test-rewrite-rename-{d}.xlsx", .{prng.random().int(u32)});

    {
        var wb = try Workbook.open(std.testing.allocator, src_path);
        defer wb.deinit();
        const s0 = try wb.sheet(0);
        // Cross-sheet ref using the source's actual sheet name "Sheet2".
        try s0.setCell("A1", .{ .formula = "Sheet2!A1+1" });
        try wb.save(tmp_path);
    }
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    var wb = try Workbook.open(std.testing.allocator, tmp_path);
    defer wb.deinit();
    const count = try wb.rewriteAllFormulas(.{
        .rename_sheet = .{ .old = "Sheet2", .new = "Renamed" },
    });
    try std.testing.expectEqual(@as(u32, 1), count);

    var tmp2_buf: [256]u8 = undefined;
    const tmp2_path = try std.fmt.bufPrint(&tmp2_buf, ".zig-cache/test-rewrite-rename-out-{d}.xlsx", .{prng.random().int(u32)});
    try wb.save(tmp2_path);
    defer std.fs.cwd().deleteFile(tmp2_path) catch {};

    var wb2 = try Workbook.open(std.testing.allocator, tmp2_path);
    defer wb2.deinit();
    const a1 = (try (try wb2.sheet(0)).cellByRef("A1")).?;
    // Bare cross-sheet name re-emits as bare.
    try std.testing.expectEqualStrings("Renamed!A1+1", a1.formula.?);
}

test "Workbook.renameSheet: happy path renames sheet and rewrites cross-sheet formula" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));

    var tmp_buf: [256]u8 = undefined;
    const tmp_path = try std.fmt.bufPrint(&tmp_buf, ".zig-cache/test-wb-rename-in-{d}.xlsx", .{prng.random().int(u32)});

    // Stage a cross-sheet formula referencing "Sheet2", save fresh copy.
    {
        var wb = try Workbook.open(std.testing.allocator, src_path);
        defer wb.deinit();
        const s0 = try wb.sheet(0);
        try s0.setCell("A1", .{ .formula = "Sheet2!A1+1" });
        try wb.save(tmp_path);
    }
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    // Re-open, renameSheet(1, "Renamed"), save, re-open, verify.
    var tmp2_buf: [256]u8 = undefined;
    const tmp2_path = try std.fmt.bufPrint(&tmp2_buf, ".zig-cache/test-wb-rename-out-{d}.xlsx", .{prng.random().int(u32)});
    {
        var wb = try Workbook.open(std.testing.allocator, tmp_path);
        defer wb.deinit();
        try wb.renameSheet(1, "Renamed");

        // In-memory view must reflect the rename immediately.
        try std.testing.expectEqualStrings("Renamed", (try wb.sheet(1)).name());

        try wb.save(tmp2_path);
    }
    defer std.fs.cwd().deleteFile(tmp2_path) catch {};

    var wb2 = try Workbook.open(std.testing.allocator, tmp2_path);
    defer wb2.deinit();
    try std.testing.expectEqualStrings("Renamed", (try wb2.sheet(1)).name());
    const a1 = (try (try wb2.sheet(0)).cellByRef("A1")).?;
    try std.testing.expectEqualStrings("Renamed!A1+1", a1.formula.?);
}

test "Workbook.renameSheet: rejects forbidden character with InvalidSheetName" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;
    var wb = try Workbook.open(std.testing.allocator, src_path);
    defer wb.deinit();
    try std.testing.expectError(error.InvalidSheetName, wb.renameSheet(0, "Has:Colon"));
    // Other forbidden characters round out the negative space.
    try std.testing.expectError(error.InvalidSheetName, wb.renameSheet(0, "back\\slash"));
    try std.testing.expectError(error.InvalidSheetName, wb.renameSheet(0, "fwd/slash"));
    try std.testing.expectError(error.InvalidSheetName, wb.renameSheet(0, "ques?tion"));
    try std.testing.expectError(error.InvalidSheetName, wb.renameSheet(0, "as*terisk"));
    try std.testing.expectError(error.InvalidSheetName, wb.renameSheet(0, "[bracket"));
    try std.testing.expectError(error.InvalidSheetName, wb.renameSheet(0, "bracket]"));
    try std.testing.expectError(error.InvalidSheetName, wb.renameSheet(0, ""));
    try std.testing.expectError(error.InvalidSheetName, wb.renameSheet(0, "history"));
    try std.testing.expectError(error.InvalidSheetName, wb.renameSheet(0, "HISTORY"));
    try std.testing.expectError(error.InvalidSheetName, wb.renameSheet(0, "History"));
}

test "Workbook.renameSheet: duplicate name errors SheetNameInUse" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;
    var wb = try Workbook.open(std.testing.allocator, src_path);
    defer wb.deinit();
    const s0_name_owned = try std.testing.allocator.dupe(u8, (try wb.sheet(0)).name());
    defer std.testing.allocator.free(s0_name_owned);
    // Renaming sheet 1 to sheet 0's exact name → conflict.
    try std.testing.expectError(error.SheetNameInUse, wb.renameSheet(1, s0_name_owned));
    // Case-insensitive: lowercase variant also conflicts.
    var lower_buf: [128]u8 = undefined;
    @memcpy(lower_buf[0..s0_name_owned.len], s0_name_owned);
    for (lower_buf[0..s0_name_owned.len]) |*c| {
        if (c.* >= 'A' and c.* <= 'Z') c.* += 32;
    }
    try std.testing.expectError(error.SheetNameInUse, wb.renameSheet(1, lower_buf[0..s0_name_owned.len]));
}

test "Workbook.renameSheet: out-of-range index errors SheetIndexOutOfRange" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;
    var wb = try Workbook.open(std.testing.allocator, src_path);
    defer wb.deinit();
    try std.testing.expectError(error.SheetIndexOutOfRange, wb.renameSheet(99, "X"));
    // Boundary: exact sheetCount() is also out-of-range (0-based index).
    try std.testing.expectError(error.SheetIndexOutOfRange, wb.renameSheet(wb.sheetCount(), "X"));
}

test "Workbook.renameSheet: no-op when new_name equals current name" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;
    var wb = try Workbook.open(std.testing.allocator, src_path);
    defer wb.deinit();
    const before = try std.testing.allocator.dupe(u8, (try wb.sheet(0)).name());
    defer std.testing.allocator.free(before);
    try wb.renameSheet(0, before);
    try std.testing.expectEqualStrings(before, (try wb.sheet(0)).name());
}

test "Workbook.hasUnsavedChanges: pristine workbook is clean immediately after open" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;
    var wb = try Workbook.open(std.testing.allocator, src_path);
    defer wb.deinit();
    try std.testing.expect(!wb.hasUnsavedChanges());
}

test "Workbook.hasUnsavedChanges: setCell flips the dirty bit" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;
    var wb = try Workbook.open(std.testing.allocator, src_path);
    defer wb.deinit();
    try std.testing.expect(!wb.hasUnsavedChanges());
    const s0 = try wb.sheet(0);
    try s0.setCell("A1", .{ .number = 42 });
    try std.testing.expect(wb.hasUnsavedChanges());
}

test "Workbook.hasUnsavedChanges: save clears delta-only dirt" {
    // setCell stages only Worksheet.deltas (no PartStore override
    // until save commits via replacePart). Workbook.save clears
    // every Worksheet's deltas, so for a pristine-then-setCell-
    // then-save lifecycle the bit goes false → true → true. The
    // bit stays true post-save because PartStore.save does not
    // reset overrides; that's documented on hasUnsavedChanges.
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    var tmp_buf: [256]u8 = undefined;
    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));
    const tmp_path = try std.fmt.bufPrint(&tmp_buf, ".zig-cache/test-wb-dirty-{d}.xlsx", .{prng.random().int(u32)});

    var wb = try Workbook.open(std.testing.allocator, src_path);
    defer wb.deinit();
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    const s0 = try wb.sheet(0);
    try s0.setCell("A1", .{ .number = 42 });
    try std.testing.expect(wb.hasUnsavedChanges());

    try wb.save(tmp_path);
    // deltas are cleared, but the just-replaced sheet part registered
    // a PartStore override — predicate stays true.
    try std.testing.expect(wb.hasUnsavedChanges());
    try std.testing.expectEqual(@as(u32, 0), (try wb.sheet(0)).deltas.count());

    // Re-opening the saved file yields a clean dirty bit.
    var wb2 = try Workbook.open(std.testing.allocator, tmp_path);
    defer wb2.deinit();
    try std.testing.expect(!wb2.hasUnsavedChanges());
}

test "Workbook.hasUnsavedChanges: renameSheet flips the bit via PartStore override" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;
    var wb = try Workbook.open(std.testing.allocator, src_path);
    defer wb.deinit();
    try std.testing.expect(!wb.hasUnsavedChanges());
    try wb.renameSheet(0, "Renamed");
    // No deltas — `renameSheet` rewrites `xl/workbook.xml` directly
    // via `store.replacePart`, so the dirty bit must be reachable
    // through the PartStore branch.
    try std.testing.expect(wb.hasUnsavedChanges());
}
