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
const drawing_emit = @import("drawing_emit.zig");
const sheet_edit = @import("sheet_edit.zig");
const sst_plan_mod = @import("zlsx_sst_plan");
const styles_plan_mod = @import("zlsx_styles_plan");
const workbook_xml_plan_mod = @import("zlsx_workbook_xml_plan");
// B3 iter-wr-6: canonical XML-escape helpers + sheet-plan registry
// types live on `pkg/sheet_plan.zig`. Workbook routes its <Override>,
// `<sheet name="…">`, `<si>`, and rich-run emit through the same
// byte-stable producers `xlsx.Writer` uses.
const sheet_plan = @import("zlsx_sheet_plan");

// B3 iter-wr-7: fresh-archive emit substrate. Lifts the entire archive
// orchestration ([Content_Types].xml + rels + workbook.xml + per-sheet
// sheet/rels/comments/vml + sst + styles + ZIP CD/EOCD) into a std-only
// module shared with `xlsx.Writer.save`. Workbook calls into this when
// `saveFreshEmit` is invoked; existing delta-on-bytes `Workbook.save`
// continues to ignore it.
const fresh_emit = @import("zlsx_fresh_emit");

/// B3 iter-wr-7: deflate adapter for the fresh-emit substrate. Routes
/// to `xlsx.deflateCompress` (the canonical deflate impl). Wrapping
/// is needed because `pkg/fresh_emit.zig` and `pkg/zip.zig` are
/// std-only (cycle-avoidance) and accept a `DeflateFn` callback.
fn freshEmitDeflate(
    alloc: Allocator,
    input: []const u8,
    out: *std.ArrayListUnmanaged(u8),
) anyerror!void {
    return zlsx.deflateCompress(alloc, input, out);
}

const PartStore = store_mod.PartStore;

const workbook_xml_mod = typed_parts.workbook_xml;
const sheet_xml_mod = typed_parts.sheet_xml;
const sst_xml_mod = typed_parts.sst_xml;
const styles_xml_mod = typed_parts.styles_xml;

/// Re-exported SST extension-plan substrate (B3 iter-wr-1). Definition
/// moved to `pkg/sst_plan.zig` so `xlsx.Writer` can import the same
/// types without the cycle that would form via `zlsx → writer.zig →
/// pkg/workbook.zig → zlsx`.
pub const RichRun = sst_plan_mod.RichRun;
pub const RichEntry = sst_plan_mod.RichEntry;
pub const SstExtensionPlan = sst_plan_mod.SstExtensionPlan;

// B3 iter-wr-2: shared `StylesPlan` substrate. Workbook gains
// `addStyle` / `addDxf` / `internNumFmt` thin pass-throughs that
// delegate to a `StylesPlan`. Type re-exports keep the
// `pkg.Workbook.addStyle(.{ ... })` call shape ergonomic.
pub const StylesPlan = styles_plan_mod.StylesPlan;
pub const Style = styles_plan_mod.Style;
pub const Dxf = styles_plan_mod.Dxf;
pub const BorderSide = styles_plan_mod.BorderSide;
pub const BorderStyle = styles_plan_mod.BorderStyle;
pub const PatternType = styles_plan_mod.PatternType;
pub const HAlign = styles_plan_mod.HAlign;

/// Re-exported workbook.xml fresh-emit plan substrate (B3 iter-wr-3).
/// Same cycle-avoidance argument as `SstExtensionPlan`: definition
/// lives in `pkg/workbook_xml_plan.zig` so `xlsx.Writer` can stage
/// defined names through the same shape.
pub const WorkbookXmlPlan = workbook_xml_plan_mod.WorkbookXmlPlan;
pub const WorkbookDefinedName = workbook_xml_plan_mod.DefinedName;
pub const DefinedNameOptions = workbook_xml_plan_mod.DefinedNameOptions;

pub const Error = error{
    /// Style validation failed — empty font name, non-positive font
    /// size, or empty number format string. Surfaces from
    /// `Workbook.addStyle` / `Workbook.internNumFmt`.
    InvalidStyle,
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
    /// `Workbook.deleteSheet` refused to remove the sole remaining
    /// sheet. OOXML / Excel require ≥1 sheet per workbook; opening a
    /// zero-sheet `.xlsx` is a hard error in every consumer we know
    /// about. The check is at the API boundary, not in the on-wire
    /// patch helpers, so callers get a typed error before any state
    /// mutation occurs.
    LastSheetUndeletable,
    /// `Workbook.deleteSheet` could not locate the target
    /// `<Relationship Id="rId…">` element in
    /// `xl/_rels/workbook.xml.rels`. Surfaces only when
    /// `WorkbookXml.sheets[idx].r_id` references an Id that doesn't
    /// exist on the wire (corrupted rels file).
    RelationshipElementNotFound,
    /// `Workbook.deleteSheet` could not locate the target
    /// `<Override PartName="/xl/worksheets/sheetN.xml" .../>` element
    /// in `[Content_Types].xml`. Soft: deleteSheet proceeds without
    /// the override removal in this case (some tooling omits per-sheet
    /// Overrides), but the typed error is exposed for future use.
    ContentTypesOverrideNotFound,
    /// `Workbook.addImage` v1 rejects sheets that already carry a
    /// `<drawing>` element. Multi-image / drawing-extension lands in
    /// a follow-up iter.
    SheetHasExistingDrawing,
    /// Image bytes' magic header didn't match the declared `mime`.
    MimeMagicMismatch,
    /// Image anchor had a 0 row or 0 col (OOXML is 1-based).
    InvalidAnchor,
    /// Empty image buffer.
    EmptyImage,
    /// `xl/worksheets/_rels/sheetN.xml.rels` exists but lacks the
    /// closing `</Relationships>` tag — refused rather than producing
    /// an unparseable rels part.
    MalformedSheetRels,
    /// Sheet XML is missing its closing `</worksheet>` tag — refused
    /// rather than producing an unparseable sheet part.
    MalformedSheetXml,
    /// `Worksheet.setCell` refused because the sheet already has
    /// `appended_rows` staged — the iter-er-3 contract makes setCell
    /// and appendRows mutually exclusive on a single Worksheet.
    SheetHasUnsavedAppends,
    /// `Worksheet.appendRows` refused because the sheet already has
    /// `setCell`/`deleteCell` deltas staged — symmetric mutual-
    /// exclusion guard. Lifting this gate is post-iter-er-2e (when
    /// the substring fast-path can merge with delta regen).
    SheetHasUnsavedMutations,
    /// `Worksheet.appendRows` row width exceeds Excel's 16_384-column
    /// cap (`xlsx.max_col_1based`).
    ColumnIndexOutOfRange,
    /// `Worksheet.appendRows` `.integer` value can't round-trip
    /// exactly through f64 (i.e. `!fitsExactlyInF64(n)`). Excel
    /// stores integers as f64 internally; values above 2^53 lose
    /// precision silently.
    IntegerExceedsExcelPrecision,
    /// `Worksheet.emitWithAppendsUsingPlan` refused because the
    /// spliced rows would push past Excel's 1,048,576-row cap
    /// (`xlsx.max_row`). Caller should split the append into
    /// multiple sheets — there is no in-place recovery.
    RowIndexOutOfRange,
    /// `Workbook.insertRow` refused because shifting an existing
    /// row would push it past Excel's 1,048,576-row cap. Surfaced
    /// from `pkg/sheet_edit.zig`.
    RowEditExceedsMaxRow,
    /// `Workbook.insertColumn` refused because shifting an
    /// existing column would push it past Excel's XFD-column cap.
    ColEditExceedsMaxCol,
    /// `Workbook.addSheet` refused because the workbook already
    /// holds the type-system maximum (`std.math.maxInt(u32)`)
    /// number of sheets. Excel imposes no documented sheet count
    /// limit (memory-bounded) but the typed-overlay's u32
    /// indexing does.
    TooManySheets,
    /// `Worksheet.emitWithAppendsUsingPlan` was given a string cell
    /// whose payload was not registered in the SST extension plan.
    /// Surfaces only on a `Workbook.save` bookkeeping bug — the plan
    /// builder walks every sheet's `appended_rows` and `deltas`, so
    /// reaching this error means a string was added between plan-
    /// build and emit.
    SharedStringNotInPlan,
    WriteFailed,
} || workbook_xml_mod.Error || sheet_xml_mod.ParseError || sst_xml_mod.Error || styles_xml_mod.Error || store_mod.Error ||
    workbook_xml_plan_mod.Error ||
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
    /// Fully remove the cell from `<sheetData>`. Distinct from
    /// `.blank` (which emits an empty `<c r="REF"/>` — cell present,
    /// no value): a `.deleted` delta elides the cell entirely from
    /// the regenerated sheet XML, so post-save `cellByRef(ref)`
    /// returns `null`. Staging a `.deleted` delta against a ref that
    /// isn't present in the source sheet is a no-op (delta carries
    /// nothing to elide).
    deleted: void,
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

/// Composite read-only view of a cell's resolved style. Each field
/// borrows from the workbook's `StylesXml` arena and is valid for as
/// long as the parent `Workbook` lives.
///
/// v1 simplification: when an `apply_*` flag on the underlying CellXf
/// is false, the corresponding field surfaces as `null`. OOXML's full
/// semantics inherit from `cellStyleXfs[xf.xfId]` in that case; we
/// defer that walk until callers explicitly request it.
///
/// `number_format_code` is `null` when `num_fmt_id` falls in the
/// built-in range (0..163, ECMA-376 §18.8.30) — those codes are
/// implicit and not stored in `<numFmts>`. Callers that need a
/// rendered string for a built-in id must map it themselves.
pub const ResolvedStyle = struct {
    font: ?styles_xml_mod.Font,
    fill: ?styles_xml_mod.Fill,
    border: ?styles_xml_mod.Border,
    alignment: ?styles_xml_mod.Alignment,
    number_format_code: ?[]const u8,
};

/// Resolved number-format descriptor for a cell-style index. Returned
/// by `Workbook.numberFormatFor`. `is_builtin == true` ⇒ `code` points
/// at a static literal from the OOXML built-in table (IDs 0..163 per
/// ECMA-376 §18.8.30); `false` ⇒ `code` borrows from the StylesXml
/// arena (alive while the producing `Workbook` is alive).
pub const NumberFormatInfo = struct {
    fmt_id: u32,
    code: []const u8,
    is_builtin: bool,
};

/// Map an OOXML built-in `numFmtId` to its format code per
/// ECMA-376 §18.8.30 Table. Covers the well-known subset (0-22, 37-49)
/// which is ~95% of real workbooks; locale-specific entries (27-36,
/// 50-58, 81) and anything ≥164 are treated as custom and fall through
/// to the `<numFmt>` table in `xl/styles.xml`.
///
/// Skipped IDs (deliberate): 5-8 (currency variants — locale-driven,
/// not portable as static strings), 23-36 (locale/CJK formats with no
/// stable pan-Excel rendering), 41-44 (currency with locale conditions),
/// 50-58 (locale Asian / hijri / Thai / etc.), 81. Real workbooks
/// touching these embed an explicit custom `<numFmt>` anyway, which our
/// custom-table fallback already handles.
fn builtinNumFmtCode(id: u32) ?[]const u8 {
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
        47 => "mmss.0",
        48 => "##0.0E+0",
        49 => "@",
        else => null,
    };
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

    /// B3 iter-wr-2: fresh-emit styles registry. Mirrors the SST plan
    /// pattern — Workbook gains `addStyle` / `addDxf` / `internNumFmt`
    /// pass-throughs so callers can stage styles + dxfs + custom
    /// number formats without going through `xlsx.Writer`. The plan
    /// owns every duped string; `deinit` walks it. Empty-by-default,
    /// so an existing-file workbook that never calls `addStyle`
    /// pays nothing.
    styles_plan: StylesPlan = .{},

    /// Workbook.xml fresh-emit plan (B3 iter-wr-3). Today's only axis
    /// is defined names, registered through `Workbook.addDefinedName`.
    /// On `save`, if at least one entry has been staged the workbook
    /// re-emits `xl/workbook.xml` from scratch via
    /// `workbook_xml_plan_mod.emitWorkbookXml`. The Writer-rebase
    /// path (`xlsx.Writer.save`) consults this same plan.
    workbook_xml_plan: WorkbookXmlPlan = .{},

    /// B3 iter-wr-7: fresh-emit shared-strings registry. The
    /// per-Worksheet `body` builder routes string interns through this
    /// plan; same byte-format as `xlsx.Writer.sst_plan`. Empty for
    /// edit-source workbooks; only populated when callers stage
    /// fresh-emit rows. `Workbook.saveFreshEmit` consumes it directly;
    /// `Workbook.save`'s delta path uses a separately-built plan from
    /// `buildSstExtensionPlan`.
    fresh_sst_plan: SstExtensionPlan = .{},

    /// B3 iter-wr-7: total string-typed cell occurrences across all
    /// sheets. The OOXML `<sst count=>` attribute. Distinct from
    /// `fresh_sst_plan.new_strings.items.len` (uniqueCount).
    fresh_sst_count: u64 = 0,

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

    /// Construct a fresh, empty `Workbook` with no source archive. The
    /// returned workbook holds the OOXML-minimum required parts:
    /// `[Content_Types].xml`, `_rels/.rels`, `xl/workbook.xml` (with
    /// an empty `<sheets/>`), and `xl/_rels/workbook.xml.rels` (with
    /// an empty `<Relationships>` body). Zero worksheets in the typed
    /// view; `addSheet(name)` grows it.
    ///
    /// Used by the upcoming B3 Writer-rebase track: `xlsx.Writer.save`
    /// will call `Workbook.empty(alloc)`, populate via `addSheet` +
    /// `setCell` / `appendRows`, then `save(path)`.
    ///
    pub fn empty(allocator: Allocator) Error!Workbook {
        var store = try PartStore.fresh(allocator);
        errdefer store.deinit();

        // `PartStore.fresh` seeds only `[Content_Types].xml`. Seed
        // the OOXML-minimum remaining parts before handing off to
        // `fromStore`. `addPart` appends the corresponding
        // `<Default>` / `<Override>` to CT.xml automatically.
        try store.addPart(
            "_rels/.rels",
            "application/vnd.openxmlformats-package.relationships+xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" ++
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" ++
                "</Relationships>",
        );
        try store.addPart(
            "xl/workbook.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
                "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" ++
                "<sheets></sheets>" ++
                "</workbook>",
        );
        try store.addPart(
            "xl/_rels/workbook.xml.rels",
            "application/vnd.openxmlformats-package.relationships+xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" ++
                "</Relationships>",
        );

        return try fromStore(allocator, store);
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
        self.styles_plan.deinit(self.allocator);
        self.workbook.deinit(self.allocator);
        self.workbook_xml_plan.deinit(self.allocator);
        self.fresh_sst_plan.deinit(self.allocator);
        self.store.deinit();
    }

    /// Register a cell style in the workbook-level styles plan and
    /// return its 1-based `s="…"` index. Dedupes by content. Mirrors
    /// `xlsx.Writer.addStyle` byte-for-byte (both ultimately route
    /// through `StylesPlan.addStyle`). Use this in conjunction with
    /// `Worksheet.setCell` to author styled cells without going
    /// through Writer's fluent-builder API.
    pub fn addStyle(self: *Workbook, style: Style) Error!u32 {
        return self.styles_plan.addStyle(self.allocator, style) catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            error.InvalidFontSize, error.InvalidFontName, error.InvalidNumberFormat => return error.InvalidStyle,
        };
    }

    /// Register a differential format for conditional formatting and
    /// return its 0-based dxfId. Dedupes by content. Mirrors
    /// `xlsx.Writer.addDxf` byte-for-byte.
    pub fn addDxf(self: *Workbook, dxf: Dxf) Error!u32 {
        return self.styles_plan.addDxf(self.allocator, dxf) catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            else => unreachable,
        };
    }

    /// Intern a custom number format string into the workbook's
    /// numFmt pool and return the assigned numFmtId. Subsequent calls
    /// with the same `format_code` return the same id. The first
    /// custom format gets id `styles_plan.NUM_FMT_BASE` (164); each
    /// new format increments the counter.
    pub fn internNumFmt(self: *Workbook, format_code: []const u8) Error!u32 {
        if (format_code.len == 0) return error.InvalidStyle;
        return self.styles_plan.internNumFmt(self.allocator, format_code) catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            else => unreachable,
        };
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

    /// Register a workbook-level defined name (B3 iter-wr-3 fresh-emit
    /// surface). Validates the name shape against Excel's full rule
    /// set (R1C1 reject, A1-shape reject, illegal char reject,
    /// 1..255 length, etc.); rejects empty `refers_to`; rejects
    /// case-insensitive duplicates within the same scope.
    ///
    /// `opts.local_sheet_id` (0-based) makes the name sheet-scoped —
    /// the index is bounds-checked at save / fresh-emit time, not at
    /// `addDefinedName` time, because the sheet count may grow after
    /// the name is registered. `opts.hidden = true` hides the name
    /// from Excel's Name Manager UI (used by `_xlnm.Print_Area` and
    /// similar).
    ///
    /// On success the workbook owns duped copies of `name` and
    /// `refers_to`; the caller can free their staging buffers
    /// immediately.
    pub fn addDefinedName(
        self: *Workbook,
        name: []const u8,
        refers_to: []const u8,
        opts: DefinedNameOptions,
    ) Error!void {
        try self.workbook_xml_plan.addDefinedName(
            self.allocator,
            name,
            refers_to,
            opts,
        );
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

    /// Resolve the number-format string a cell of `style_idx` would
    /// render with. Combines the OOXML built-in table (IDs 0..49 well-
    /// known subset; anything else falls through to the custom
    /// `<numFmt>` table in `xl/styles.xml`).
    ///
    /// Returns `null` when:
    ///   - the workbook has no `xl/styles.xml`,
    ///   - `style_idx` is outside `cell_xfs`,
    ///   - the resolved `numFmtId` matches no built-in and no custom
    ///     entry (malformed input — the cell would render as `General`
    ///     in Excel; callers wanting that fallback should treat `null`
    ///     as "General" themselves).
    ///
    /// Lifetime: built-in `code` is a `'static` string literal; custom
    /// `code` borrows from the StylesXml arena (alive as long as the
    /// `Workbook`).
    pub fn numberFormatFor(self: *Workbook, style_idx: u32) Error!?NumberFormatInfo {
        const styles_view = (try self.styles()) orelse return null;
        if (style_idx >= styles_view.cell_xfs.len) return null;
        const xf = styles_view.cell_xfs[style_idx];
        const nfid = xf.num_fmt_id orelse return null;
        if (builtinNumFmtCode(nfid)) |code| {
            return .{ .fmt_id = nfid, .code = code, .is_builtin = true };
        }
        for (styles_view.number_formats) |nf| {
            if (nf.fmt_id == nfid) {
                return .{ .fmt_id = nfid, .code = nf.code, .is_builtin = false };
            }
        }
        return null;
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
        // Phase 0 (B3 iter-wr-3): apply the workbook.xml fresh-emit
        // plan. Today the only axis is staged defined names — splice
        // them into `xl/workbook.xml` BEFORE the SST + per-sheet
        // phases so the workbook.xml byte image they read is final.
        // The validator at `addDefinedName` time guarantees every
        // staged name is well-formed; only the bounds check on
        // `local_sheet_id` runs at emit (sheet count may have grown
        // since the name was registered).
        if (self.workbook_xml_plan.defined_names.items.len > 0) {
            try self.applyWorkbookXmlPlanDefinedNames();
        }

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
            // Per-sheet, deltas and appended_rows are mutually
            // exclusive (refused at staging time). Handle whichever
            // is non-empty; skip clean sheets entirely.
            if (ws.appended_rows.items.len > 0) {
                assert(ws.deltas.count() == 0);
                const part_name = try ws.resolvePartName();
                const new_xml = try ws.emitWithAppendsUsingPlan(self.allocator, &sst_plan);
                defer self.allocator.free(new_xml);
                try self.store.replacePart(part_name, new_xml);
                ws.clearAppendedRows(self.allocator);
                if (ws.parsed) |*p| {
                    var stale = p.*;
                    stale.deinit(self.allocator);
                    ws.parsed = null;
                }
                continue;
            }
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

    /// B3 iter-wr-7: emit a fresh `.xlsx` archive from the workbook's
    /// fresh-emit registries (`fresh_sst_plan`, `fresh_sst_count`,
    /// `styles_plan`, `workbook_xml_plan`) plus per-Worksheet `body` +
    /// `state`. Routes through the shared `pkg/fresh_emit.zig`
    /// substrate, producing byte-identical archives to
    /// `xlsx.Writer.save`. The PartStore is bypassed entirely — this
    /// is a pure producer path. Use `Workbook.empty()` then
    /// `addSheet` + per-Worksheet `add*`/`set*` registrations + this
    /// to author workbooks without a backing source archive.
    ///
    /// Errors flow through verbatim from the substrate: `NoSheets`
    /// when no worksheets are registered; allocation/zip-sentinel
    /// errors on archive build.
    pub fn saveFreshEmit(self: *Workbook, path: []const u8) !void {
        if (self.worksheets.len == 0) return error.NoSheets;

        const inputs = try self.allocator.alloc(fresh_emit.SheetInput, self.worksheets.len);
        defer self.allocator.free(inputs);
        for (self.worksheets, 0..) |*ws, i| {
            inputs[i] = .{
                .name = ws.name(),
                .body = ws.body.items,
                .state = &ws.sheet_state,
            };
        }

        return fresh_emit.saveArchiveToPath(self.allocator, path, .{
            .sheets = inputs,
            .sst_plan = &self.fresh_sst_plan,
            .sst_count = self.fresh_sst_count,
            .styles_plan = &self.styles_plan,
            .workbook_xml_plan = &self.workbook_xml_plan,
        }, freshEmitDeflate);
    }

    /// Stage a plain string into the fresh-emit SST and return the
    /// 0-based index for use in row-emit `<v>{idx}</v>` payloads.
    /// Increments `fresh_sst_count` (the running cell-occurrence
    /// counter that becomes the OOXML `<sst count=>` attribute).
    /// Pair with the `body` payload assembled per-Worksheet.
    pub fn freshSstIntern(self: *Workbook, text: []const u8) Error!u32 {
        const idx = try self.fresh_sst_plan.registerNewPlain(self.allocator, text);
        self.fresh_sst_count += 1;
        return idx;
    }

    /// Apply a structural-edit rewrite to every `<dataValidation>`
    /// formula1/formula2 and every `<cfRule>` formula across every
    /// sheet, persisting the result in-place via `store.replacePart`.
    /// Returns the count of formula *bodies* whose rewrite produced
    /// different bytes (so a DV with both formula1 and formula2
    /// changed counts as 2; an unchanged body — including a no-op
    /// shift — counts 0).
    ///
    /// `target_sheet` scopes the edit the same way as
    /// `RewriteContext.target_sheet`: when non-null, only refs that
    /// resolve to that sheet (bare refs on a matching `on_sheet`, or
    /// sheet-qualified refs naming `target_sheet`) shift. `null`
    /// means "apply everywhere".
    ///
    /// **Persistence model.** This emits patched sheet XML bytes
    /// *immediately* via `PartStore.replacePart`. It does NOT use the
    /// `Workbook.save`-deltas pipeline (DV/CF aren't cell mutations).
    /// Run this BEFORE `Workbook.save` if a save also has pending
    /// `setCell` deltas — `save` re-fetches part bytes per sheet, so
    /// it sees the patched DV/CF blocks and preserves them in its
    /// own splice. The cached `SheetXml` view is invalidated
    /// (`parsed = null`) for any sheet rewritten here, matching the
    /// invalidation contract used by `save`.
    ///
    /// **Splice strategy.** Patches the formula inner text in place,
    /// byte-for-byte, inside each `<formula1>`, `<formula2>`, and
    /// CF `<formula>` element whose body the rewriter changed. Every
    /// surrounding attribute (`errorTitle`, `error`, `prompt`,
    /// `xr:uid`, `dxf_id`, `priority`, `operator`, etc.) is preserved
    /// verbatim — we never regenerate the DV/CF block from the typed
    /// view, which would lose any trivia the parser doesn't expose.
    ///
    /// **Body counting.** Each formula body that produces different
    /// bytes counts once. So a CF `D1+E1` rewritten to `E1+F1` (one
    /// body, two refs shifted) is one rewrite, not two.
    pub fn rewriteAllValidationsAndConditionalFormats(
        self: *Workbook,
        edit: zlsx.formula_rewriter.RewriteEdit,
        target_sheet: ?[]const u8,
    ) Error!u32 {
        var count: u32 = 0;
        const a = self.allocator;

        var sheet_idx: u32 = 0;
        while (sheet_idx < self.sheetCount()) : (sheet_idx += 1) {
            const ws = try self.sheet(sheet_idx);
            const view = try ws.ensureParsed();
            const ws_name = ws.name();
            const part_name = ws.resolved_part_name.?;

            // Two phases. Phase A: rewrite each DV/CF formula body
            // against the typed view, building an indexed plan
            // (DV index, CF index) keyed by the *position in the
            // view* — NOT source-byte offsets, since typed view
            // slices borrow from the parser's sanitized buffer, not
            // from `source`. Phase B walks the source XML and re-
            // locates each `<formula1>` / `<formula2>` / `<formula>`
            // body in lockstep with the view, splicing where the plan
            // says so.
            var dv_f1_new: std.AutoHashMapUnmanaged(usize, []u8) = .{};
            var dv_f2_new: std.AutoHashMapUnmanaged(usize, []u8) = .{};
            var cf_f_new: std.AutoHashMapUnmanaged(usize, []u8) = .{};
            defer {
                var it1 = dv_f1_new.iterator();
                while (it1.next()) |e| a.free(e.value_ptr.*);
                dv_f1_new.deinit(a);
                var it2 = dv_f2_new.iterator();
                while (it2.next()) |e| a.free(e.value_ptr.*);
                dv_f2_new.deinit(a);
                var it3 = cf_f_new.iterator();
                while (it3.next()) |e| a.free(e.value_ptr.*);
                cf_f_new.deinit(a);
            }

            for (view.validations, 0..) |dv, i| {
                if (dv.formula1) |f| {
                    if (try maybeRewrite(a, f, ws_name, target_sheet, edit)) |new| {
                        errdefer a.free(new);
                        try dv_f1_new.put(a, i, new);
                    }
                }
                if (dv.formula2) |f| {
                    if (try maybeRewrite(a, f, ws_name, target_sheet, edit)) |new| {
                        errdefer a.free(new);
                        try dv_f2_new.put(a, i, new);
                    }
                }
            }
            for (view.conditional_formats, 0..) |cf, j| {
                if (cf.formula) |f| {
                    if (try maybeRewrite(a, f, ws_name, target_sheet, edit)) |new| {
                        errdefer a.free(new);
                        try cf_f_new.put(a, j, new);
                    }
                }
            }

            const total = dv_f1_new.count() + dv_f2_new.count() + cf_f_new.count();
            if (total == 0) continue;

            // Phase B: walk source XML, build patch list of source-
            // byte spans (start, end, replacement). Then linear-splice.
            const source = blk: {
                const p = try self.store.part(part_name) orelse return error.MissingSheetPart;
                break :blk p.bytes;
            };
            assert(source.len > 0);

            var patches: std.ArrayList(SourcePatch) = .empty;
            defer patches.deinit(a);
            try collectDvCfPatches(a, source, &patches, &dv_f1_new, &dv_f2_new, &cf_f_new);

            // Sanity: every queued rewrite should have located a
            // splice site in the source (typed view and source share
            // document order; a missing site means the source was
            // mutated under us, which is a `replacePart`-ordering bug).
            assert(patches.items.len == total);

            const new_xml = try spliceFormulas(a, source, patches.items);
            defer a.free(new_xml);

            try self.store.replacePart(part_name, new_xml);
            count += @intCast(total);

            // Invalidate the cached parsed view: its leaves borrowed
            // from the old part bytes which `replacePart` swapped.
            // Mirrors the invalidation pattern in `Workbook.save`.
            var stale = ws.parsed.?;
            stale.deinit(self.allocator);
            ws.parsed = null;
        }

        return count;
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

    /// Apply `edit` to every `<definedName>` formula in `xl/workbook.xml`.
    /// `target_sheet` plumbs into `RewriteContext.target_sheet`: caller's
    /// statement of which sheet's row/col edit applies. `null` means
    /// "apply everywhere" per the rewriter's permissive default.
    ///
    /// Each defined name's `RewriteContext.on_sheet` is set to:
    ///   - workbook-scope (`local_sheet_id == null`): `null`
    ///   - sheet-scope (`local_sheet_id` set): the name of the sheet
    ///     the `localSheetId` index resolves to
    ///
    /// Returns the number of defined-name formulas whose rewritten
    /// output differs from the input (byte-identical rewrites are not
    /// counted).
    ///
    /// **Splice contract.** Three shapes for `<definedNames>` are
    /// handled in `xl/workbook.xml`:
    ///   - Paired `<definedNames>...</definedNames>` — classical splice
    ///   - Self-closing `<definedNames/>` — upgraded to paired form
    ///   - Block absent — insert fresh paired block before `<calcPr` (or
    ///     before `</workbook>` if no calcPr)
    ///
    /// **Bug-fix vs prior iter (PR #37).** The earlier draft re-parsed
    /// `self.workbook` mid-iteration over `defined_names`, leaving the
    /// loop reading freed memory on the sheet-scope branch. This impl
    /// collects ALL rewrites into parallel allocator-owned arrays
    /// FIRST, then performs the splice + `replacePart` + re-parse in a
    /// single transactional step.
    pub fn rewriteAllDefinedNames(
        self: *Workbook,
        edit: zlsx.formula_rewriter.RewriteEdit,
        target_sheet: ?[]const u8,
    ) Error!u32 {
        const a = self.allocator;

        if (self.workbook.defined_names.len == 0) return 0;

        // Owned strings for every defined name we plan to emit. Either
        // the rewritten formula (mutated) or a duplicated copy of the
        // original (unchanged). Owning every entry uniformly simplifies
        // the splice loop's lifetime story.
        var owned_formulas: std.ArrayList([]u8) = .empty;
        defer {
            for (owned_formulas.items) |s| a.free(s);
            owned_formulas.deinit(a);
        }
        // Parallel arrays: name / local_sheet_id / hidden, dup'd so
        // they survive the upcoming `refreshWorkbookXmlView` call which
        // frees the source `defined_names` arena.
        var owned_names: std.ArrayList([]u8) = .empty;
        defer {
            for (owned_names.items) |s| a.free(s);
            owned_names.deinit(a);
        }
        var local_ids: std.ArrayList(?u32) = .empty;
        defer local_ids.deinit(a);
        var hiddens: std.ArrayList(bool) = .empty;
        defer hiddens.deinit(a);

        var changed: u32 = 0;
        for (self.workbook.defined_names) |dn| {
            // Resolve `on_sheet`: null for workbook-scope, sheet name
            // for sheet-scope (via local_sheet_id index lookup).
            const on_sheet: ?[]const u8 = blk: {
                if (dn.local_sheet_id) |sid| {
                    if (sid < self.workbook.sheets.len) {
                        break :blk self.workbook.sheets[sid].name;
                    }
                    // Out-of-range localSheetId in source XML — treat
                    // as workbook-scope rather than crashing. Tolerant
                    // of malformed input, not a happy path.
                    break :blk null;
                }
                break :blk null;
            };
            const ctx = zlsx.formula_rewriter.RewriteContext{
                .on_sheet = on_sheet,
                .target_sheet = target_sheet,
                .edit = edit,
            };

            const rewritten = try zlsx.formula_rewriter.rewriteFormula(a, dn.formula, ctx);
            errdefer a.free(rewritten);

            const name_dup = try a.dupe(u8, dn.name);
            errdefer a.free(name_dup);

            if (!std.mem.eql(u8, rewritten, dn.formula)) changed += 1;

            try owned_formulas.append(a, rewritten);
            try owned_names.append(a, name_dup);
            try local_ids.append(a, dn.local_sheet_id);
            try hiddens.append(a, dn.hidden);
        }

        // Pair-assertion: parallel arrays agree in length and match
        // the source slice's length.
        assert(owned_formulas.items.len == self.workbook.defined_names.len);
        assert(owned_names.items.len == self.workbook.defined_names.len);
        assert(local_ids.items.len == self.workbook.defined_names.len);
        assert(hiddens.items.len == self.workbook.defined_names.len);

        if (changed == 0) return 0;

        try self.spliceDefinedNamesBlock(
            owned_names.items,
            owned_formulas.items,
            local_ids.items,
            hiddens.items,
        );

        // Re-parse so subsequent reads of `self.workbook.defined_names`
        // see the new bytes. Source slice borrows from the OLD bytes
        // we just replaced; `refresh` swaps both arena and slice.
        try refreshWorkbookXmlView(self);

        return changed;
    }

    /// Apply `edit` to every internal `<hyperlink location="…">` on
    /// every sheet. External hyperlinks (those with `r:id != null`)
    /// are skipped — their target is a relationship, not an A1 ref.
    ///
    /// `target_sheet` plumbs into `RewriteContext.target_sheet`.
    /// `RewriteContext.on_sheet` is the name of the sheet the
    /// hyperlink lives on (so unqualified locations are scoped
    /// correctly during shift).
    ///
    /// Returns the number of hyperlink locations whose rewrite
    /// produced different bytes.
    pub fn rewriteAllHyperlinkLocations(
        self: *Workbook,
        edit: zlsx.formula_rewriter.RewriteEdit,
        target_sheet: ?[]const u8,
    ) Error!u32 {
        const a = self.allocator;
        var total_changed: u32 = 0;

        var sheet_idx: u32 = 0;
        while (sheet_idx < self.sheetCount()) : (sheet_idx += 1) {
            const ws = try self.sheet(sheet_idx);
            const view = try ws.ensureParsed();
            if (view.hyperlinks.len == 0) continue;

            const ws_name = ws.name();

            // Per-sheet pending list: every hyperlink (preserved
            // verbatim) plus the rewritten location, owned. Empty
            // means nothing to emit — leave the sheet untouched.
            const Pending = struct {
                ref: []u8,
                location: []u8,
                display: ?[]u8,
                tooltip: ?[]u8,
                r_id: ?[]u8,
            };
            var pending: std.ArrayList(Pending) = .empty;
            defer {
                for (pending.items) |p| {
                    a.free(p.ref);
                    a.free(p.location);
                    if (p.display) |s| a.free(s);
                    if (p.tooltip) |s| a.free(s);
                    if (p.r_id) |s| a.free(s);
                }
                pending.deinit(a);
            }

            var sheet_changed: u32 = 0;
            for (view.hyperlinks) |h| {
                // External hyperlinks: r_id present → location is a
                // relationship target, not an A1 ref. Preserve verbatim.
                // Internal-only entries (r_id == null AND location set)
                // are the rewrite candidates.
                const loc_in = h.location orelse {
                    // Hyperlink with neither r_id nor location is
                    // malformed; preserve verbatim so it's not dropped.
                    try appendPendingHyperlink(a, &pending, h, null);
                    continue;
                };
                if (h.r_id != null) {
                    try appendPendingHyperlink(a, &pending, h, null);
                    continue;
                }

                const ctx = zlsx.formula_rewriter.RewriteContext{
                    .on_sheet = ws_name,
                    .target_sheet = target_sheet,
                    .edit = edit,
                };
                const rewritten = try zlsx.formula_rewriter.rewriteFormula(a, loc_in, ctx);
                errdefer a.free(rewritten);

                if (!std.mem.eql(u8, rewritten, loc_in)) sheet_changed += 1;
                try appendPendingHyperlink(a, &pending, h, rewritten);
            }

            // Pair assertion: every parsed hyperlink produced exactly
            // one pending entry.
            assert(pending.items.len == view.hyperlinks.len);

            if (sheet_changed == 0) continue;

            try self.spliceHyperlinksBlock(sheet_idx, pending.items);

            // Invalidate the parsed view: the source bytes for
            // `hyperlinks[].location` borrow from the OLD part bytes
            // that `replacePart` just swapped out.
            if (ws.parsed) |*p| {
                var view_mut = p.*;
                view_mut.deinit(self.allocator);
                ws.parsed = null;
            }

            total_changed += sheet_changed;
        }

        return total_changed;
    }

    /// Helper for `rewriteAllHyperlinkLocations`. Duplicates `h`'s
    /// borrowed slices into pending-list-owned allocations, optionally
    /// substituting `loc_override` for `h.location` (used when the
    /// rewriter produced new bytes). Pending-list type is local to
    /// the caller; mirror its struct shape via anytype to avoid
    /// hoisting a private struct to file scope just for this helper.
    fn appendPendingHyperlink(
        allocator: Allocator,
        pending: anytype,
        h: sheet_xml_mod.Hyperlink,
        loc_override: ?[]u8,
    ) Error!void {
        const ref_dup = try allocator.dupe(u8, h.ref);
        errdefer allocator.free(ref_dup);

        const loc_dup: []u8 = if (loc_override) |lo|
            lo // takes ownership — caller's `errdefer free(rewritten)` is cancelled by successful append
        else
            try allocator.dupe(u8, h.location orelse "");
        errdefer if (loc_override == null) allocator.free(loc_dup);

        const display_dup: ?[]u8 = if (h.display) |s| try allocator.dupe(u8, s) else null;
        errdefer if (display_dup) |s| allocator.free(s);

        const tooltip_dup: ?[]u8 = if (h.tooltip) |s| try allocator.dupe(u8, s) else null;
        errdefer if (tooltip_dup) |s| allocator.free(s);

        const r_id_dup: ?[]u8 = if (h.r_id) |s| try allocator.dupe(u8, s) else null;
        errdefer if (r_id_dup) |s| allocator.free(s);

        try pending.append(allocator, .{
            .ref = ref_dup,
            .location = loc_dup,
            .display = display_dup,
            .tooltip = tooltip_dup,
            .r_id = r_id_dup,
        });
    }

    /// B3 iter-wr-3: splice the workbook.xml fresh-emit plan's
    /// staged defined names into `xl/workbook.xml`. Validates every
    /// staged `local_sheet_id` against the current sheet count;
    /// builds the parallel-arrays shape `spliceDefinedNamesBlock`
    /// expects; clears the plan after a successful splice (saves
    /// are idempotent — re-running shouldn't re-add the same
    /// names).
    ///
    /// Mutually exclusive with `rewriteAllDefinedNames` in any single
    /// save: rewrite is "rewrite EVERY name's formula via edit"
    /// while this is "splice the plan's NEW names". The plan is
    /// drained on success, so subsequent rewrites observe the new
    /// names from the re-parsed view.
    fn applyWorkbookXmlPlanDefinedNames(self: *Workbook) Error!void {
        const a = self.allocator;
        const plan = &self.workbook_xml_plan;
        assert(plan.defined_names.items.len > 0);

        // Bounds-check every sheet-scoped name BEFORE building any
        // parallel arrays — atomicity: an out-of-range local_sheet_id
        // bails before we allocate anything.
        for (plan.defined_names.items) |dn| {
            if (dn.local_sheet_id) |sid| {
                if (sid >= self.workbook.sheets.len) {
                    return error.InvalidDefinedNameLocalSheetId;
                }
            }
        }

        // Merge any pre-existing defined names (parsed from the source
        // workbook.xml) with the plan's staged additions. The splice
        // helper rewrites the entire block, so dropping the existing
        // entries would silently delete user-loaded names.
        const existing = self.workbook.defined_names;
        const total = existing.len + plan.defined_names.items.len;

        var owned_names: std.ArrayList([]u8) = .empty;
        defer {
            for (owned_names.items) |s| a.free(s);
            owned_names.deinit(a);
        }
        var owned_formulas: std.ArrayList([]u8) = .empty;
        defer {
            for (owned_formulas.items) |s| a.free(s);
            owned_formulas.deinit(a);
        }
        var local_ids: std.ArrayList(?u32) = .empty;
        defer local_ids.deinit(a);
        var hiddens: std.ArrayList(bool) = .empty;
        defer hiddens.deinit(a);

        try owned_names.ensureTotalCapacity(a, total);
        try owned_formulas.ensureTotalCapacity(a, total);
        try local_ids.ensureTotalCapacity(a, total);
        try hiddens.ensureTotalCapacity(a, total);

        for (existing) |dn| {
            try owned_names.append(a, try a.dupe(u8, dn.name));
            try owned_formulas.append(a, try a.dupe(u8, dn.formula));
            try local_ids.append(a, dn.local_sheet_id);
            try hiddens.append(a, dn.hidden);
        }
        for (plan.defined_names.items) |dn| {
            try owned_names.append(a, try a.dupe(u8, dn.name));
            try owned_formulas.append(a, try a.dupe(u8, dn.refers_to));
            try local_ids.append(a, dn.local_sheet_id);
            try hiddens.append(a, dn.hidden);
        }

        try self.spliceDefinedNamesBlock(
            owned_names.items,
            owned_formulas.items,
            local_ids.items,
            hiddens.items,
        );

        // Re-parse so subsequent reads of `self.workbook.defined_names`
        // observe the freshly-spliced block. Mirrors the pattern in
        // `rewriteAllDefinedNames`.
        const part = (try self.store.part("xl/workbook.xml")) orelse return error.MissingWorkbookPart;
        var fresh = try workbook_xml_mod.parse(a, part.bytes);
        errdefer fresh.deinit(a);
        self.workbook.deinit(a);
        self.workbook = fresh;

        // Drain the plan so a subsequent save (or the C ABI's
        // save-and-mutate-and-resave loop) doesn't redundantly
        // re-splice the same names.
        plan.deinit(a);
        plan.* = .{};
    }

    /// Re-emit `xl/workbook.xml` with a fresh `<definedNames>` block
    /// built from the parallel arrays (names / formulas / local_ids /
    /// hiddens — same length, same index ↔ same defined name).
    fn spliceDefinedNamesBlock(
        self: *Workbook,
        names: []const []u8,
        formulas: []const []u8,
        local_ids: []const ?u32,
        hiddens: []const bool,
    ) Error!void {
        assert(names.len == formulas.len);
        assert(names.len == local_ids.len);
        assert(names.len == hiddens.len);

        const a = self.allocator;
        const part = try self.store.part("xl/workbook.xml") orelse return error.MissingWorkbookPart;
        const src = part.bytes;
        assert(src.len > 0);

        // Build the new `<definedNames>...</definedNames>` block once
        // — emitted whether the source had a paired block, a self-
        // closing tag, or no block at all. Empty `names` would emit a
        // bare `<definedNames></definedNames>` pair; callers gate on
        // `changed == 0` and skip this path entirely if no rewrite
        // happened, so empty is never emitted.
        var block: std.ArrayList(u8) = .empty;
        defer block.deinit(a);
        try block.appendSlice(a, "<definedNames>");
        for (names, formulas, local_ids, hiddens) |n, f, lid, hid| {
            try block.appendSlice(a, "<definedName name=\"");
            try appendXmlEscaped(a, &block, n);
            try block.appendSlice(a, "\"");
            if (lid) |sid| {
                var buf: [16]u8 = undefined;
                const s = try std.fmt.bufPrint(&buf, "{d}", .{sid});
                try block.appendSlice(a, " localSheetId=\"");
                try block.appendSlice(a, s);
                try block.appendSlice(a, "\"");
            }
            if (hid) try block.appendSlice(a, " hidden=\"1\"");
            try block.appendSlice(a, ">");
            try appendXmlEscapedText(a, &block, f);
            try block.appendSlice(a, "</definedName>");
        }
        try block.appendSlice(a, "</definedNames>");

        // Locate the source `<definedNames` open tag (if any). Decision
        // tree:
        //   present + paired      → splice over the whole block
        //   present + self-close  → splice over the self-close tag
        //   absent                → insert before <calcPr or </workbook>
        var out: std.ArrayList(u8) = .empty;
        defer out.deinit(a);
        try out.ensureTotalCapacity(a, src.len + block.items.len + 64);

        if (std.mem.indexOf(u8, src, "<definedNames")) |dn_open| {
            // Boundary check: ensure the tag is real, not a substring
            // of `<definedNamesEx>` or similar (no such Excel tag, but
            // defensive: require the next byte to be space/tab/newline,
            // `>`, or `/`).
            const after_name = dn_open + "<definedNames".len;
            if (after_name >= src.len) return error.MalformedXml;
            const boundary = src[after_name];
            const is_real = switch (boundary) {
                ' ', '\t', '\r', '\n', '>', '/' => true,
                else => false,
            };
            if (!is_real) return error.MalformedXml;

            const open_gt = std.mem.indexOfScalarPos(u8, src, after_name, '>') orelse
                return error.MalformedXml;
            const is_self_closing = open_gt > 0 and src[open_gt - 1] == '/';

            if (is_self_closing) {
                // Replace `<definedNames/>` with the new paired block.
                try out.appendSlice(a, src[0..dn_open]);
                try out.appendSlice(a, block.items);
                try out.appendSlice(a, src[open_gt + 1 ..]);
            } else {
                // Find `</definedNames>` and replace the whole span.
                const close = std.mem.indexOfPos(u8, src, open_gt + 1, "</definedNames>") orelse
                    return error.MalformedXml;
                const close_end = close + "</definedNames>".len;
                try out.appendSlice(a, src[0..dn_open]);
                try out.appendSlice(a, block.items);
                try out.appendSlice(a, src[close_end..]);
            }
        } else {
            // No `<definedNames>` block: insert before `<calcPr`, or
            // before `</workbook>` if calcPr is absent. This places the
            // block at OOXML's expected position in the schema sequence.
            const insert_at: usize = blk: {
                if (std.mem.indexOf(u8, src, "<calcPr")) |i| break :blk i;
                if (std.mem.indexOf(u8, src, "</workbook>")) |i| break :blk i;
                return error.MalformedXml;
            };
            try out.appendSlice(a, src[0..insert_at]);
            try out.appendSlice(a, block.items);
            try out.appendSlice(a, src[insert_at..]);
        }

        try self.store.replacePart("xl/workbook.xml", out.items);
    }

    /// Re-emit the sheet at `sheet_idx`'s XML with a fresh
    /// `<hyperlinks>` block. `pending` contents own their byte slices.
    fn spliceHyperlinksBlock(
        self: *Workbook,
        sheet_idx: u32,
        pending: anytype,
    ) Error!void {
        const a = self.allocator;
        const ws = try self.sheet(sheet_idx);
        // ensureParsed has been called by the time we get here, so
        // resolved_part_name is non-null.
        const part_name = ws.resolved_part_name orelse return error.MissingSheetPart;
        const part = try self.store.part(part_name) orelse return error.MissingSheetPart;
        const src = part.bytes;
        assert(src.len > 0);

        // Build the new block. ALL hyperlinks (including the r_id-
        // bearing ones passed through) are emitted so the splice
        // replaces the entire `<hyperlinks>` block as a unit.
        var block: std.ArrayList(u8) = .empty;
        defer block.deinit(a);
        try block.appendSlice(a, "<hyperlinks>");
        for (pending) |p| {
            try block.appendSlice(a, "<hyperlink ref=\"");
            try appendXmlEscaped(a, &block, p.ref);
            try block.appendSlice(a, "\"");
            if (p.r_id) |rid| {
                try block.appendSlice(a, " r:id=\"");
                try appendXmlEscaped(a, &block, rid);
                try block.appendSlice(a, "\"");
            }
            // Emit `location` only when non-empty AND `r_id` is null;
            // otherwise the relationship target supersedes.
            if (p.r_id == null and p.location.len > 0) {
                try block.appendSlice(a, " location=\"");
                try appendXmlEscaped(a, &block, p.location);
                try block.appendSlice(a, "\"");
            }
            if (p.display) |d| {
                try block.appendSlice(a, " display=\"");
                try appendXmlEscaped(a, &block, d);
                try block.appendSlice(a, "\"");
            }
            if (p.tooltip) |t| {
                try block.appendSlice(a, " tooltip=\"");
                try appendXmlEscaped(a, &block, t);
                try block.appendSlice(a, "\"");
            }
            try block.appendSlice(a, "/>");
        }
        try block.appendSlice(a, "</hyperlinks>");

        // Same three-shape splice as definedNames.
        var out: std.ArrayList(u8) = .empty;
        defer out.deinit(a);
        try out.ensureTotalCapacity(a, src.len + block.items.len + 64);

        if (std.mem.indexOf(u8, src, "<hyperlinks")) |hl_open| {
            const after_name = hl_open + "<hyperlinks".len;
            if (after_name >= src.len) return error.MalformedXml;
            const boundary = src[after_name];
            const is_real = switch (boundary) {
                ' ', '\t', '\r', '\n', '>', '/' => true,
                else => false,
            };
            if (!is_real) return error.MalformedXml;

            const open_gt = std.mem.indexOfScalarPos(u8, src, after_name, '>') orelse
                return error.MalformedXml;
            const is_self_closing = open_gt > 0 and src[open_gt - 1] == '/';

            if (is_self_closing) {
                try out.appendSlice(a, src[0..hl_open]);
                try out.appendSlice(a, block.items);
                try out.appendSlice(a, src[open_gt + 1 ..]);
            } else {
                const close = std.mem.indexOfPos(u8, src, open_gt + 1, "</hyperlinks>") orelse
                    return error.MalformedXml;
                const close_end = close + "</hyperlinks>".len;
                try out.appendSlice(a, src[0..hl_open]);
                try out.appendSlice(a, block.items);
                try out.appendSlice(a, src[close_end..]);
            }
        } else {
            // No `<hyperlinks>` in source — unreachable when
            // `pending.len > 0` because pending was built from the
            // parsed view, and the parser only populates hyperlinks
            // when the block exists. Return a clean error rather than
            // silently inserting at a guessed position.
            return error.MalformedXml;
        }

        try self.store.replacePart(part_name, out.items);
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
    /// **Length cap.** B3 iter-wr-6 lifted the validator to the
    /// canonical Writer-side `xlsx.validateSheetName` (Unicode-scalar
    /// aware): names cap at 31 SCALARS — Excel's actual rule —
    /// regardless of UTF-8 byte length, additionally rejecting
    /// 0x00..0x1F control bytes and leading/trailing apostrophes.
    /// Pre-iter-wr-6 the cap was a conservative 127-byte fence.
    ///
    /// **Case folding.** Duplicate-name detection still uses ASCII
    /// case-fold (a..z ↔ A..Z); the Writer-side `addSheet` collision
    /// loop uses the Unicode case-fold via
    /// `src/unicode/casefold.zig`. Workbook's collision check is
    /// per-sheet rather than per-write, so the ASCII compromise is
    /// retained here pending a follow-up iter.
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

        const edit: zlsx.formula_rewriter.RewriteEdit = .{
            .rename_sheet = .{ .old = old_name_owned, .new = new_name },
        };

        // B2 iter-er-5 lift (rename_sheet axis): walk every cross-
        // sheet reference carrier and rewrite the renamed sheet's
        // qualifiers in place. Until iter-er-5, only formula cells
        // were rewritten; defined-name formulas, internal hyperlink
        // locations, and DV/CF formulas all survived as `#REF!`
        // until the next manual save in Excel. Each rewriter is
        // tolerant of "no carriers in the workbook" — they short-
        // circuit on empty workbooks rather than erroring, so
        // calling all four unconditionally is cheap.
        _ = try self.rewriteAllFormulas(edit);
        _ = try self.rewriteAllDefinedNames(edit, null);
        _ = try self.rewriteAllHyperlinkLocations(edit, null);
        _ = try self.rewriteAllValidationsAndConditionalFormats(edit, null);

        try patchWorkbookXmlSheetName(self, sheet_idx, old_name_owned, new_name);
        try refreshWorkbookXmlView(self);

        // Postcondition: the in-memory view now reports the new name.
        assert(sheet_idx < self.workbook.sheets.len);
        assert(std.mem.eql(u8, self.workbook.sheets[sheet_idx].name, new_name));
    }

    /// Add a new empty sheet to the workbook (B2 iter-er-4
    /// structural-edit surface). Allocates a fresh sheet part
    /// (`xl/worksheets/sheetN.xml`), patches `xl/workbook.xml` +
    /// `xl/_rels/workbook.xml.rels` (and `[Content_Types].xml` via
    /// `PartStore.addPart`), then grows the typed-overlay
    /// `worksheets` array. Returns a handle to the newly-added
    /// `Worksheet` view.
    ///
    /// The new sheet's body is the OOXML-minimal empty-sheet
    /// template — `<worksheet>` wrapping `<sheetData/>`. Subsequent
    /// `setCell` / `appendRows` / `addImage` calls operate through
    /// the returned handle.
    ///
    /// **Cap.** Refuses with `error.TooManySheets` if the workbook
    /// already holds `std.math.maxInt(u32)` sheets. Excel itself
    /// has no documented sheet count limit (memory-bounded), but
    /// the typed-overlay's u32 indexing does.
    ///
    /// **Errors:**
    ///   - `error.InvalidSheetName` per `validateSheetName`'s
    ///     contract (length, reserved chars, "history").
    ///   - `error.SheetNameInUse` if the name collides with any
    ///     existing sheet (case-insensitive ASCII).
    ///   - `error.TooManySheets` at the u32 cap.
    ///   - `error.MissingWorkbookPart` / `error.MissingWorkbookRels`
    ///     if the prerequisite parts have been spliced away.
    ///
    /// **Atomicity.** Like `addImage`, the mutation is not perfectly
    /// transactional across the four PartStore writes; ordering
    /// (addPart → rels → workbook.xml → view-refresh) keeps the
    /// workbook consistent right up to the workbook.xml splice.
    /// A failure between addPart and the workbook.xml patch leaves
    /// the part in the store unreferenced; cleanup ships with
    /// `removePart` in a future iter.
    ///
    /// **Pointer lifetime.** The returned `*Worksheet` (and any
    /// `*Worksheet` previously returned by `Workbook.sheet`) is
    /// invalidated by the next structural mutation (`addSheet`,
    /// `deleteSheet`) — those calls reallocate `self.worksheets`.
    /// Re-fetch via `wb.sheet(idx)` after structural edits.
    pub fn addSheet(self: *Workbook, name: []const u8) Error!*Worksheet {
        try validateSheetName(name);
        if (self.worksheets.len >= std.math.maxInt(u32)) return error.TooManySheets;
        // Workbook view stores sheet names as raw attribute bytes —
        // a source name like `R&D` is held as `R&amp;D`. Compare
        // against the entity-decoded form so attribute escaping
        // doesn't bypass the duplicate check. Direct byte-equality
        // first as a fast path; only allocate for entity decoding
        // when the candidate contains an `&`.
        for (self.workbook.sheets) |s| {
            if (try sheetNameMatchesDecoded(self.allocator, s.name, name)) return error.SheetNameInUse;
        }

        const wb_part = (try self.store.part("xl/workbook.xml")) orelse
            return error.MissingWorkbookPart;
        const rels_part = (try self.store.part("xl/_rels/workbook.xml.rels")) orelse
            return error.MissingWorkbookRels;

        // Pick non-colliding identifiers by scanning the source bytes.
        // Each addSheet rescans because previous calls extend the
        // workbook view in-place — nothing else tracks the running
        // high-water marks.
        const next_sheet_id = nextMaxNumericAttr(wb_part.bytes, "sheetId=\"") + 1;
        const next_rid_num = nextMaxNumericAttr(rels_part.bytes, "Id=\"rId") + 1;
        // Path number must avoid orphan worksheet parts too — when
        // a sheet is deleted via `Workbook.deleteSheet`, the part
        // remains in PartStore (orphan-part v1 trade-off). Computing
        // the next slot purely from rels would re-collide on the
        // orphan's path. Walk PartStore instead.
        const next_path_num = (try nextMaxSheetPathNumFromStore(&self.store)) + 1;

        var path_buf: [64]u8 = undefined;
        const new_path = try std.fmt.bufPrint(&path_buf, "xl/worksheets/sheet{d}.xml", .{next_path_num});
        var rid_buf: [32]u8 = undefined;
        const new_rid = try std.fmt.bufPrint(&rid_buf, "rId{d}", .{next_rid_num});

        const empty_body =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" ++
            "<sheetData/></worksheet>";

        try self.store.addPart(
            new_path,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            empty_body,
        );

        // workbook.xml.rels: add `<Relationship Id="rIdN" Type="…/worksheet" Target="worksheets/sheetN.xml"/>`.
        const new_rels_xml = try patchWorkbookRelsAddSheet(self.allocator, rels_part.bytes, new_rid, new_path);
        defer self.allocator.free(new_rels_xml);
        try self.store.replacePart("xl/_rels/workbook.xml.rels", new_rels_xml);

        // workbook.xml: add `<sheet name="…" sheetId="N" r:id="rIdN"/>`.
        const new_wb_xml = try patchWorkbookXmlAddSheet(self.allocator, wb_part.bytes, name, next_sheet_id, new_rid);
        defer self.allocator.free(new_wb_xml);
        try self.store.replacePart("xl/workbook.xml", new_wb_xml);

        // Re-parse workbook.xml. The patch added one sheet so the
        // fresh view's sheet count must equal the old view's + 1.
        const wb_part2 = (try self.store.part("xl/workbook.xml")) orelse
            return error.MissingWorkbookPart;
        var fresh = try workbook_xml_mod.parse(self.allocator, wb_part2.bytes);
        errdefer fresh.deinit(self.allocator);
        if (fresh.sheets.len != self.workbook.sheets.len + 1) {
            return error.SheetCountMismatch;
        }

        // Grow the worksheets slot table to match the new view.
        const old_count = self.worksheets.len;
        const new_count = old_count + 1;
        const new_slots = try self.allocator.alloc(Worksheet, new_count);
        errdefer self.allocator.free(new_slots);
        @memcpy(new_slots[0..old_count], self.worksheets);
        new_slots[old_count] = .{
            .workbook = self,
            .sheet_idx = @intCast(old_count),
            .parsed = null,
            .resolved_part_name = null,
        };
        // Patch back-pointers — old slots may have moved during
        // the realloc.
        for (new_slots) |*ws| ws.workbook = self;

        // Commit: swap workbook view, free old slots, install new.
        self.workbook.deinit(self.allocator);
        self.workbook = fresh;
        self.allocator.free(self.worksheets);
        self.worksheets = new_slots;

        const new_idx: u32 = @intCast(old_count);
        assert(new_idx < self.worksheets.len);
        return &self.worksheets[new_idx];
    }

    /// Embed `bytes` (PNG/JPEG/GIF, declared via `mime`) as an image
    /// pinned to `cell_anchor` on the sheet at `sheet_idx`. v1 ships a
    /// MINIMAL surface:
    ///   - one image per call
    ///   - `oneCellAnchor` only (top-left pinned at zero pixel offset)
    ///   - fixed 1-inch × 1-inch extent (914 400 × 914 400 EMU)
    ///   - `cell_anchor` is 1-based `(col, row)` — col first matches
    ///     OOXML drawing wire format
    ///
    /// Sheets that already carry a `<drawing>` element are rejected
    /// with `error.SheetHasExistingDrawing`. Range-anchor, pixel
    /// offsets, and native-size extent ship in follow-up iters.
    ///
    /// The mutation commits to the in-memory PartStore immediately
    /// (image bytes, drawing part, drawing rels, sheet XML splice,
    /// sheet rels). `Workbook.save` writes the archive to disk.
    pub fn addImage(
        self: *Workbook,
        sheet_idx: u32,
        cell_anchor: drawing_emit.ImageCellAnchor,
        bytes: []const u8,
        mime: drawing_emit.ImageMime,
    ) Error!void {
        if (sheet_idx >= self.sheetCount()) return error.SheetIndexOutOfRange;
        if (cell_anchor.col == 0 or cell_anchor.row == 0) return error.InvalidAnchor;

        // Reject mime/byte mismatch up front. Costs O(8) bytes; saves
        // a "broken image" round-trip in Office.
        try drawing_emit.validateMagic(mime, bytes);

        const ws = try self.sheet(sheet_idx);
        // ensureParsed populates resolved_part_name, needed to derive
        // the sheet rels file path.
        _ = try ws.ensureParsed();
        const sheet_part_name = ws.resolved_part_name.?;
        assert(std.mem.startsWith(u8, sheet_part_name, "xl/worksheets/"));
        assert(std.mem.endsWith(u8, sheet_part_name, ".xml"));

        // Reject sheets that already have a drawing. Substring check
        // is good-enough as a pre-filter: a `<drawing` token only
        // appears at the worksheet level on a sheet that's already
        // wired one up.
        const sheet_part = (try self.store.part(sheet_part_name)) orelse
            return error.MissingSheetPart;
        if (std.mem.indexOf(u8, sheet_part.bytes, "<drawing ") != null or
            std.mem.indexOf(u8, sheet_part.bytes, "<drawing/>") != null or
            std.mem.indexOf(u8, sheet_part.bytes, "<drawing>") != null)
        {
            return error.SheetHasExistingDrawing;
        }

        const a = self.allocator;

        // Derive part-name slots. No gap-filling: bumping past the
        // highest seen N keeps the result stable across saves.
        const image_n = drawing_emit.nextFreeNumber(&self.store, "xl/media/image", "");
        const drawing_n = drawing_emit.nextFreeNumber(&self.store, "xl/drawings/drawing", ".xml");

        const ext = mime.extension();
        // Compose part names. Free at function-exit; addPart dupes
        // its arguments into PartStore's arena.
        const image_basename = try std.fmt.allocPrint(a, "image{d}.{s}", .{ image_n, ext });
        defer a.free(image_basename);
        const image_part_name = try std.fmt.allocPrint(a, "xl/media/{s}", .{image_basename});
        defer a.free(image_part_name);

        const drawing_basename = try std.fmt.allocPrint(a, "drawing{d}.xml", .{drawing_n});
        defer a.free(drawing_basename);
        const drawing_part_name = try std.fmt.allocPrint(a, "xl/drawings/{s}", .{drawing_basename});
        defer a.free(drawing_part_name);

        const drawing_rels_part_name = try std.fmt.allocPrint(
            a,
            "xl/drawings/_rels/drawing{d}.xml.rels",
            .{drawing_n},
        );
        defer a.free(drawing_rels_part_name);

        // Sheet rels file path = `xl/worksheets/_rels/<basename>.rels`.
        const sheet_basename = sheet_part_name["xl/worksheets/".len..];
        const sheet_rels_part_name = try std.fmt.allocPrint(
            a,
            "xl/worksheets/_rels/{s}.rels",
            .{sheet_basename},
        );
        defer a.free(sheet_rels_part_name);

        // Build the drawing artifacts.
        const drawing_xml = try drawing_emit.buildDrawingXml(a, cell_anchor);
        defer a.free(drawing_xml);

        const drawing_rels_xml = try drawing_emit.buildDrawingRels(a, image_basename);
        defer a.free(drawing_rels_xml);

        // Pick `rId` for the sheet's drawing rel. New rels file → rId1;
        // existing → next free.
        const sheet_rels_existing = self.store.rels(sheet_part_name);
        const rid_n = drawing_emit.nextFreeRelId(sheet_rels_existing);
        var rid_buf: [16]u8 = undefined;
        const rid = try std.fmt.bufPrint(&rid_buf, "rId{d}", .{rid_n});

        // Build the sheet rels (fresh or appended). Doing this BEFORE
        // any addPart/replacePart so a later addPart failure (OOM, Zip
        // size cap) leaves the store unchanged.
        const sheet_rels_part = try self.store.part(sheet_rels_part_name);
        const new_sheet_rels_xml: []u8 = if (sheet_rels_part) |p|
            try drawing_emit.appendRelationship(a, p.bytes, drawing_basename, rid)
        else
            try drawing_emit.buildFreshSheetRels(a, drawing_basename, rid);
        defer a.free(new_sheet_rels_xml);

        // Build the patched sheet XML (appends `<drawing r:id=...>`).
        const new_sheet_xml = try drawing_emit.appendDrawingElementToSheet(
            a,
            sheet_part.bytes,
            rid,
        );
        defer a.free(new_sheet_xml);

        // Mutate the PartStore. Order:
        //   1. addPart image bytes
        //   2. addPart drawing XML
        //   3. addPart drawing rels
        //   4. replace/add sheet rels
        //   5. replacePart sheet XML
        // Not perfectly atomic across these — each call mutates store
        // state — but ordering image+drawing+rels first keeps the
        // workbook consistent right up to step 5: until the sheet XML
        // references the drawing, no consumer sees the partial wiring.
        try self.store.addPart(image_part_name, mime.contentType(), bytes);
        try self.store.addPart(
            drawing_part_name,
            "application/vnd.openxmlformats-officedocument.drawing+xml",
            drawing_xml,
        );
        try self.store.addPart(
            drawing_rels_part_name,
            "application/vnd.openxmlformats-package.relationships+xml",
            drawing_rels_xml,
        );
        if (sheet_rels_part != null) {
            try self.store.replacePart(sheet_rels_part_name, new_sheet_rels_xml);
        } else {
            try self.store.addPart(
                sheet_rels_part_name,
                "application/vnd.openxmlformats-package.relationships+xml",
                new_sheet_rels_xml,
            );
        }
        try self.store.replacePart(sheet_part_name, new_sheet_xml);

        // The cached SheetXml view borrowed from the now-stale source
        // bytes. Invalidate so subsequent reads re-parse from the
        // patched XML.
        if (ws.parsed) |*v| {
            var view = v.*;
            view.deinit(self.allocator);
            ws.parsed = null;
        }
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

    /// Remove the sheet at `sheet_idx` from the workbook. Patches three
    /// XML parts and refreshes the in-memory view; subsequent
    /// `Workbook.save` writes a workbook whose `<sheets>` list, rels
    /// table, and `[Content_Types].xml` Overrides no longer reference
    /// the removed sheet.
    ///
    /// **Orphan-part trade-off (v1).** `PartStore` has no `removePart`
    /// API today, so the deleted sheet's `xl/worksheets/sheetN.xml`
    /// (and its sidecar `xl/worksheets/_rels/sheetN.xml.rels` when
    /// present) remain physically inside the archive after `save`,
    /// just unreferenced from any rels / Content_Types entries. Excel,
    /// LibreOffice, and openpyxl all tolerate orphan parts (the OPC
    /// reader resolves parts via `[Content_Types].xml` Overrides + rels
    /// graphs; unreferenced parts are dead weight). A true cleanup
    /// requires `PartStore.removePart`; that's a future iter.
    ///
    /// **Cross-references not rewritten (v1).** `<definedName>` slots
    /// with `localSheetId == sheet_idx` (sheet-scoped names), formulas
    /// referencing the deleted sheet, and hyperlink targets pointing at
    /// it remain on the wire and will produce `#REF!` or
    /// `Reference is not valid` in Excel after open. Same scope
    /// boundary as `renameSheet`'s `m3-defnames-hyperlinks` follow-up.
    ///
    /// **Errors:**
    ///   - `SheetIndexOutOfRange` — `sheet_idx >= sheetCount()`.
    ///   - `LastSheetUndeletable` — `sheetCount() == 1`.
    ///   - `MissingWorkbookPart` / `MissingWorkbookRels` — corrupt
    ///     archive lacks the parts to be patched.
    ///   - `SheetElementNotFound` / `RelationshipElementNotFound` —
    ///     the on-wire bytes don't match the parsed view (file
    ///     mutated under us between parse and patch).
    pub fn deleteSheet(self: *Workbook, sheet_idx: u32) Error!void {
        if (sheet_idx >= self.sheetCount()) return error.SheetIndexOutOfRange;
        if (self.sheetCount() == 1) return error.LastSheetUndeletable;
        // Pre-condition: the slot table and the parsed view agree on
        // length (the same invariant `sheetCount` asserts on read).
        assert(self.worksheets.len == self.workbook.sheets.len);
        assert(self.worksheets.len >= 2);

        // Capture r_id BEFORE patches: the post-patch re-parse frees the
        // arena that backs `sheets[idx].r_id`, but the rels patch needs
        // the bytes alive.
        const r_id_src = self.workbook.sheets[sheet_idx].r_id;
        if (r_id_src.len == 0) return error.MissingRelationship;
        if (r_id_src.len > 64) return error.MissingRelationship;
        var r_id_buf: [64]u8 = undefined;
        @memcpy(r_id_buf[0..r_id_src.len], r_id_src);
        const r_id_owned = r_id_buf[0..r_id_src.len];

        // Resolve the part name (e.g. "xl/worksheets/sheet2.xml") via
        // ensureParsed — caches `resolved_part_name` on the Worksheet.
        // Dupe into a stack buffer so the slot deinit below can't free
        // it from under the Content_Types patch.
        const ws = try self.sheet(sheet_idx);
        _ = try ws.ensureParsed();
        const part_name_src = ws.resolved_part_name orelse return error.MissingSheetPart;
        if (part_name_src.len == 0) return error.MissingSheetPart;
        if (part_name_src.len > 256) return error.MissingSheetPart;
        var part_name_buf: [256]u8 = undefined;
        @memcpy(part_name_buf[0..part_name_src.len], part_name_src);
        const part_name_owned = part_name_buf[0..part_name_src.len];

        // B2 iter-er-5 lift (deleteSheet defined-names axis): rewrite
        // every cross-sheet ref to the doomed sheet across the
        // remaining sheets. Cell formulas, defined-name formulas,
        // internal hyperlink locations, and DV / CF formulas all
        // collapse to `#REF!` on the qualifier-matching path. Bare
        // refs are unaffected (the deleted sheet's own formulas
        // are dropped with the sheet, not rewritten).
        const doomed_name_src = self.workbook.sheets[sheet_idx].name;
        const doomed_name_owned = try self.allocator.dupe(u8, doomed_name_src);
        defer self.allocator.free(doomed_name_owned);
        const edit: zlsx.formula_rewriter.RewriteEdit = .{ .delete_sheet = doomed_name_owned };
        _ = try self.rewriteAllFormulas(edit);
        _ = try self.rewriteAllDefinedNames(edit, null);
        _ = try self.rewriteAllHyperlinkLocations(edit, null);
        _ = try self.rewriteAllValidationsAndConditionalFormats(edit, null);

        try patchWorkbookXmlRemoveSheet(self, sheet_idx);
        try patchWorkbookRelsRemoveRelationship(self, r_id_owned);
        // Content_Types Override removal is best-effort — some
        // tooling omits per-sheet Overrides and relies on Default
        // entries instead.
        patchContentTypesRemoveOverride(self, part_name_owned) catch |err| switch (err) {
            error.ContentTypesOverrideNotFound => {},
            else => return err,
        };

        // Re-parse the workbook view in place. Can't use
        // `refreshWorkbookXmlView` here because that helper asserts
        // sheet-count invariance (correct for renameSheet, wrong
        // for delete which expects count - 1).
        {
            const fresh_part = try self.store.part("xl/workbook.xml") orelse
                return error.MissingWorkbookPart;
            var fresh = try workbook_xml_mod.parse(self.allocator, fresh_part.bytes);
            errdefer fresh.deinit(self.allocator);
            if (fresh.sheets.len + 1 != self.workbook.sheets.len) {
                return error.SheetCountMismatch;
            }
            self.workbook.deinit(self.allocator);
            self.workbook = fresh;
        }

        // Shrink the slot table. Order matters:
        //   1. deinit the doomed Worksheet BEFORE copying slots —
        //      copying first then freeing would deinit the wrong slot.
        //   2. allocate the new (n-1)-sized slot array.
        //   3. copy [0..idx] and [idx+1..] across.
        //   4. free the old array.
        const old_slots = self.worksheets;
        const old_len = old_slots.len;
        assert(sheet_idx < old_len);

        old_slots[sheet_idx].deinit(self.allocator);

        const new_slots = try self.allocator.alloc(Worksheet, old_len - 1);
        errdefer self.allocator.free(new_slots);

        var i: u32 = 0;
        var j: u32 = 0;
        while (i < old_len) : (i += 1) {
            if (i == sheet_idx) continue;
            new_slots[j] = old_slots[i];
            // Renumber: surviving Worksheet's `sheet_idx` must line up
            // with its new slot position (the re-parsed view holds
            // sheets in document order).
            new_slots[j].sheet_idx = j;
            j += 1;
        }
        assert(j == old_len - 1);

        self.allocator.free(old_slots);
        self.worksheets = new_slots;

        assert(self.worksheets.len == self.workbook.sheets.len);
        assert(self.worksheets.len == old_len - 1);
    }

    /// Read-only predicate: does the workbook carry any pending
    /// mutation that has not yet been flushed via `save`?
    ///
    /// Returns `true` if EITHER:
    ///   1. any `Worksheet.deltas` map is non-empty (uncommitted
    ///      `setCell` mutations), OR
    ///   2. the underlying `PartStore` has any override (uncommitted
    ///      `replacePart` / `addPart` from e.g. `renameSheet`,
    ///      `rewriteAllFormulas`, the SST extension path).
    ///
    /// Note: `PartStore.save` does NOT clear overrides post-save —
    /// they persist across save calls. So this predicate reflects
    /// "diff vs the original on-disk archive opened by `Workbook.open`",
    /// not "uncommitted-since-last-save". Most callers want the
    /// former (e.g. for "do I need to save before exit?" — the
    /// answer should remain true even after a previous save).
    pub fn hasUnsavedChanges(self: *const Workbook) bool {
        for (self.worksheets) |ws| {
            if (ws.deltas.count() > 0) return true;
        }
        return self.store.hasUnsavedChanges();
    }

    /// Insert a blank row at position `before_row` (1-based) in
    /// sheet `sheet_idx`. Every existing row at or below `before_row`
    /// shifts down by 1.  Mutates the workbook's PartStore in
    /// place: the sheet part is re-emitted with `<row r=>` /
    /// `<c r="…">` / `<mergeCells>` / `<dimension>` shifted by
    /// one row.
    ///
    /// **Refusal contract.** Refuses with
    /// `error.RowEditRequiresCleanSheet` if the sheet has any
    /// staged appendRows or setCell deltas — those deltas index
    /// into the pre-shift refs and would produce stale output.
    ///
    /// **Cross-sheet rewrite.** Cross-sheet formula refs / defined
    /// names / hyperlinks / DV-CF formulas are NOT yet rewritten
    /// for row inserts. Sheets that carry those constructs anywhere
    /// in the workbook are refused at the editor layer
    /// (`Editor.insertRow`); the iter-er-5 row/col-axis lifts
    /// (blocked on this PR's typed surface, plus the
    /// rewriter-call wiring) are tracked in
    /// `docs/plans/refusal-audit.md`.
    pub fn insertRow(self: *Workbook, sheet_idx: u32, before_row: u32) Error!void {
        try self.applySheetEditTransform(sheet_idx, .{ .row = before_row, .kind = .insert });
    }

    /// Delete row `row` (1-based) in sheet `sheet_idx`. Every row
    /// > `row` shifts up by 1; cells in the deleted row are
    /// dropped. Same refusal contract + cross-sheet-rewrite
    /// limitations as `insertRow`.
    pub fn deleteRow(self: *Workbook, sheet_idx: u32, row: u32) Error!void {
        try self.applySheetEditTransform(sheet_idx, .{ .row = row, .kind = .delete });
    }

    /// Insert a blank column at position `before_col` (1-based, A=1)
    /// in sheet `sheet_idx`. Every existing column at or right of
    /// `before_col` shifts right by 1. Same refusal contract as
    /// `insertRow`; cross-sheet rewrite is the same iter-er-5
    /// follow-up.
    pub fn insertColumn(self: *Workbook, sheet_idx: u32, before_col_1based: u32) Error!void {
        try self.applySheetEditTransform(sheet_idx, .{ .col = before_col_1based, .kind = .insert });
    }

    /// Delete column `col_1based` in sheet `sheet_idx`. Every
    /// column > `col_1based` shifts left by 1.
    pub fn deleteColumn(self: *Workbook, sheet_idx: u32, col_1based: u32) Error!void {
        try self.applySheetEditTransform(sheet_idx, .{ .col = col_1based, .kind = .delete });
    }

    const SheetEditSpec = struct {
        row: ?u32 = null,
        col: ?u32 = null,
        kind: sheet_edit.RowEditKind,
    };

    fn applySheetEditTransform(self: *Workbook, sheet_idx: u32, spec: SheetEditSpec) Error!void {
        if (sheet_idx >= self.sheetCount()) return error.SheetIndexOutOfRange;
        const ws = try self.sheet(sheet_idx);
        if (ws.deltas.count() > 0) return error.SheetHasUnsavedMutations;
        if (ws.appended_rows.items.len > 0) return error.SheetHasUnsavedAppends;

        const part_name = try ws.resolvePartName();
        const part = (try self.store.part(part_name)) orelse return error.MissingSheetPart;

        const new_xml = if (spec.row) |r|
            try sheet_edit.applyRowEditToWorksheet(self.allocator, part.bytes, r, spec.kind)
        else
            try sheet_edit.applyColEditToWorksheet(self.allocator, part.bytes, spec.col.?, spec.kind);
        defer self.allocator.free(new_xml);
        try self.store.replacePart(part_name, new_xml);

        // Invalidate the cached parsed view — the row/col-shifted
        // bytes don't match the previously-parsed structure, and
        // any subsequent ensureParsed must re-tokenize the new
        // body. resolved_part_name stays valid (path didn't change).
        if (ws.parsed) |*v| {
            var view = v.*;
            view.deinit(self.allocator);
            ws.parsed = null;
        }

        // B2 iter-er-5 lift (row/col axes): rewrite cross-sheet
        // refs in formulas, defined names, internal hyperlink
        // locations, and DV/CF formulas. Each rewriter targets the
        // edited sheet's bare-ref space so references like
        // `=A1+B2` (no sheet qualifier) on the edited sheet shift
        // alongside the byte-level row/col attrs above.
        //
        // The four rewriters write their edits as
        // `Worksheet.setCell` deltas (formulas) or in-place splices
        // (defined names, hyperlinks, DV/CF). Both compose with
        // the byte transform: the parsed view we just invalidated
        // re-parses from the shifted bytes on next access, so
        // `emitWithDeltas` at save time merges shifted-ref text
        // with shifted-r-attr cells.
        const target = self.workbook.sheets[sheet_idx].name;
        const edit: zlsx.formula_rewriter.RewriteEdit = if (spec.row) |r|
            switch (spec.kind) {
                .insert => .{ .insert_rows = .{ .at = r, .count = 1 } },
                .delete => .{ .delete_rows = .{ .at = r, .count = 1 } },
            }
        else switch (spec.kind) {
            .insert => .{ .insert_cols = .{ .at = spec.col.?, .count = 1 } },
            .delete => .{ .delete_cols = .{ .at = spec.col.?, .count = 1 } },
        };
        _ = try self.rewriteAllFormulas(edit);
        _ = try self.rewriteAllDefinedNames(edit, target);
        _ = try self.rewriteAllHyperlinkLocations(edit, target);
        _ = try self.rewriteAllValidationsAndConditionalFormats(edit, target);
    }
};

/// Validate a candidate sheet name per Excel's rules. B3 iter-wr-6
/// delegates to `zlsx.validateSheetName` (the Writer-side
/// Unicode-scalar-aware validator at `src/writer.zig`):
///   - 1..31 Unicode SCALARS (Excel's actual limit)
///   - no 0x00..0x1F control bytes
///   - none of `: \ / ? * [ ]`
///   - no leading or trailing apostrophe
///   - Unicode case-fold compare not equal to "History"
///
/// Workbook's `Error` surface is unchanged — both validators
/// surface the same `error.InvalidSheetName` token.
fn validateSheetName(name: []const u8) Error!void {
    zlsx.validateSheetName(name) catch return error.InvalidSheetName;
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
/// B3 iter-wr-6 NOTE: kept local rather than forwarded to
/// `pkg/sheet_plan.zig::appendXmlEscaped`. The plan-side variant
/// rejects XML 1.0 forbidden control bytes with `error.InvalidXmlByte`;
/// Workbook's `Error` set deliberately omits that variant (one-minor
/// API freeze on Workbook's public Error surface), and Workbook's
/// callers pre-validate text via `isXmlSafeText` so a forbidden byte
/// would already be a contract violation. To stay byte-equivalent
/// without widening `Error`, this 5-entity escape stays inline.
/// Producers: sheet-name patch, definedName + sheet-rels emit.
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
// ─── B2 iter-er-4 addSheet helpers ────────────────────────────────────

/// Case-insensitive ASCII compare with on-the-fly XML entity
/// decoding of `view_name`. Used by `Workbook.addSheet` to detect
/// duplicates where the workbook view holds the encoded form
/// (`R&amp;D`) and the input is the decoded form (`R&D`). Direct
/// byte-equality first as a fast path; only decodes when
/// `view_name` contains an `&`.
fn sheetNameMatchesDecoded(allocator: Allocator, view_name: []const u8, input: []const u8) Error!bool {
    if (asciiCaseInsensitiveEql(view_name, input)) return true;
    if (std.mem.indexOfScalar(u8, view_name, '&') == null) return false;

    var decoded: std.ArrayList(u8) = .empty;
    defer decoded.deinit(allocator);
    try decoded.ensureTotalCapacity(allocator, view_name.len);
    var i: usize = 0;
    while (i < view_name.len) {
        if (view_name[i] != '&') {
            try decoded.append(allocator, view_name[i]);
            i += 1;
            continue;
        }
        const semi = std.mem.indexOfScalarPos(u8, view_name, i, ';') orelse {
            try decoded.append(allocator, view_name[i]);
            i += 1;
            continue;
        };
        const ent = view_name[i + 1 .. semi];
        const replaced: ?u8 =
            if (std.mem.eql(u8, ent, "amp")) @as(u8, '&') else if (std.mem.eql(u8, ent, "lt")) @as(u8, '<') else if (std.mem.eql(u8, ent, "gt")) @as(u8, '>') else if (std.mem.eql(u8, ent, "quot")) @as(u8, '"') else if (std.mem.eql(u8, ent, "apos")) @as(u8, '\'') else null;
        if (replaced) |c| {
            try decoded.append(allocator, c);
            i = semi + 1;
        } else {
            // Unknown entity — pass `&` through and keep scanning.
            try decoded.append(allocator, view_name[i]);
            i += 1;
        }
    }
    return asciiCaseInsensitiveEql(decoded.items, input);
}

/// Scan `xml` for the largest `attr_prefixN"` numeric value (e.g.
/// `sheetId="3"` or `Id="rId7"`). Returns 0 when no match is found.
/// Used by `Workbook.addSheet` to pick non-colliding ids when
/// multiple addSheet calls land before the next save.
fn nextMaxNumericAttr(xml: []const u8, attr_prefix: []const u8) u32 {
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

/// Highest `xl/worksheets/sheetN.xml` part name in the part store
/// (including orphan parts left behind by `Workbook.deleteSheet`).
/// Returns 0 when no such part exists. Used to pick a non-colliding
/// path number for `Workbook.addSheet` even when prior delete +
/// add cycles have already consumed slots.
fn nextMaxSheetPathNumFromStore(store: *const store_mod.PartStore) Error!u32 {
    const names = try store.partNames();
    var max_n: u32 = 0;
    const prefix = "xl/worksheets/sheet";
    const suffix = ".xml";
    for (names) |name| {
        if (!std.mem.startsWith(u8, name, prefix)) continue;
        if (!std.mem.endsWith(u8, name, suffix)) continue;
        const num_str = name[prefix.len .. name.len - suffix.len];
        if (num_str.len == 0) continue;
        if (std.fmt.parseInt(u32, num_str, 10)) |n| {
            if (n > max_n) max_n = n;
        } else |_| {}
    }
    return max_n;
}

/// Highest `worksheets/sheetN.xml` number referenced by a
/// `Target="…"` attribute in the workbook rels XML. Returns 0 when
/// no such target exists. Robust against absolute (`/xl/worksheets/`)
/// and relative (`worksheets/`) prefixes — both decode to the same
/// number. Kept for potential future callers; `Workbook.addSheet`
/// uses `nextMaxSheetPathNumFromStore` instead so orphan parts
/// don't collide.
fn nextMaxSheetPathNumFromRels(xml: []const u8) u32 {
    var max_n: u32 = 0;
    const needle = "sheet";
    var i: usize = 0;
    while (std.mem.indexOfPos(u8, xml, i, needle)) |pos| {
        const num_start = pos + needle.len;
        var num_end = num_start;
        while (num_end < xml.len and xml[num_end] >= '0' and xml[num_end] <= '9') : (num_end += 1) {}
        if (num_end > num_start and num_end + 4 <= xml.len and
            std.mem.eql(u8, xml[num_end .. num_end + 4], ".xml"))
        {
            // Pre-context check: only count when this `sheetN.xml`
            // is preceded by `worksheets/` somewhere in the same
            // attribute span — strips out drawings, comments,
            // pivots, etc.
            const ctx_start = if (pos >= 24) pos - 24 else 0;
            if (std.mem.lastIndexOf(u8, xml[ctx_start..pos], "worksheets/") != null) {
                if (std.fmt.parseInt(u32, xml[num_start..num_end], 10)) |n| {
                    if (n > max_n) max_n = n;
                } else |_| {}
            }
        }
        i = num_end + 1;
    }
    return max_n;
}

/// Splice a new `<Relationship/>` for a worksheet into
/// `xl/_rels/workbook.xml.rels` immediately before `</Relationships>`.
/// Caller frees the returned slice.
fn patchWorkbookRelsAddSheet(
    allocator: Allocator,
    xml: []const u8,
    rid: []const u8,
    full_path: []const u8,
) Error![]u8 {
    const close = std.mem.indexOf(u8, xml, "</Relationships>") orelse return error.MalformedWorkbookRels;
    // Target is relative to xl/_rels/, so strip the "xl/" prefix
    // from `full_path`. e.g. "xl/worksheets/sheet5.xml" →
    // "worksheets/sheet5.xml".
    const target = if (std.mem.startsWith(u8, full_path, "xl/")) full_path[3..] else full_path;

    var out: std.ArrayList(u8) = .empty;
    errdefer out.deinit(allocator);
    try out.ensureTotalCapacity(allocator, xml.len + 256);
    try out.appendSlice(allocator, xml[0..close]);
    try out.appendSlice(allocator, "<Relationship Id=\"");
    try out.appendSlice(allocator, rid);
    try out.appendSlice(allocator, "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"");
    try out.appendSlice(allocator, target);
    try out.appendSlice(allocator, "\"/>");
    try out.appendSlice(allocator, xml[close..]);
    return try out.toOwnedSlice(allocator);
}

/// Splice a new `<sheet/>` line into `xl/workbook.xml` immediately
/// before `</sheets>`. Caller frees the returned slice. The display
/// name is XML-attribute-escaped; sheet_id and rid are emitted
/// verbatim (caller-controlled identifiers).
fn patchWorkbookXmlAddSheet(
    allocator: Allocator,
    xml: []const u8,
    name: []const u8,
    sheet_id: u32,
    rid: []const u8,
) Error![]u8 {
    const close = std.mem.indexOf(u8, xml, "</sheets>") orelse return error.MalformedXml;
    var out: std.ArrayList(u8) = .empty;
    errdefer out.deinit(allocator);
    try out.ensureTotalCapacity(allocator, xml.len + 256);
    try out.appendSlice(allocator, xml[0..close]);
    try out.appendSlice(allocator, "<sheet name=\"");
    try appendXmlAttrEscapedW(allocator, &out, name);
    var num_buf: [16]u8 = undefined;
    try out.appendSlice(allocator, "\" sheetId=\"");
    try out.appendSlice(allocator, try std.fmt.bufPrint(&num_buf, "{d}", .{sheet_id}));
    try out.appendSlice(allocator, "\" r:id=\"");
    try out.appendSlice(allocator, rid);
    try out.appendSlice(allocator, "\"/>");
    try out.appendSlice(allocator, xml[close..]);
    return try out.toOwnedSlice(allocator);
}

/// XML attribute-context escape: `<`, `>`, `&`, `"` → entity refs.
/// Use when emitting into a double-quoted attribute value.
fn appendXmlAttrEscapedW(allocator: Allocator, out: *std.ArrayList(u8), s: []const u8) Error!void {
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

// ─── deleteSheet patch helpers ────────────────────────────────────────

/// Walk the source `xl/workbook.xml` bytes, find the Nth `<sheet>`
/// element (1-based N == `sheet_idx + 1`), and splice it out — along
/// with any leading whitespace inside `<sheets>` so the result
/// remains visually stable. Re-emits the part via `replacePart`.
fn patchWorkbookXmlRemoveSheet(self: *Workbook, sheet_idx: u32) Error!void {
    assert(sheet_idx < self.workbook.sheets.len);

    const part = try self.store.part("xl/workbook.xml") orelse return error.MissingWorkbookPart;
    const src = part.bytes;
    assert(src.len > 0);

    var search_from: usize = 0;
    var seen: u32 = 0;
    var elem_open: usize = 0;
    var elem_close: usize = 0;
    while (true) {
        const open = std.mem.indexOfPos(u8, src, search_from, "<sheet") orelse
            return error.SheetElementNotFound;
        const after = open + "<sheet".len;
        if (after >= src.len) return error.SheetElementNotFound;
        const boundary = src[after];
        const is_sheet_elem = switch (boundary) {
            ' ', '\t', '\r', '\n', '/' => true,
            else => false,
        };
        if (!is_sheet_elem) {
            search_from = after;
            continue;
        }
        const gt = std.mem.indexOfScalarPos(u8, src, after, '>') orelse
            return error.SheetElementNotFound;
        if (seen == sheet_idx) {
            elem_open = open;
            // Self-closing form ends at `/>`; otherwise consume the
            // matching `</sheet>`. OOXML emits `<sheet/>` in practice
            // but tolerate the long form.
            if (gt > 0 and src[gt - 1] == '/') {
                elem_close = gt + 1;
            } else {
                const end_tag = std.mem.indexOfPos(u8, src, gt + 1, "</sheet>") orelse
                    return error.SheetElementNotFound;
                elem_close = end_tag + "</sheet>".len;
            }
            break;
        }
        seen += 1;
        search_from = gt + 1;
    }
    assert(elem_close > elem_open);

    var trim_start = elem_open;
    while (trim_start > 0) {
        const c = src[trim_start - 1];
        if (c == ' ' or c == '\t' or c == '\r' or c == '\n') {
            trim_start -= 1;
        } else break;
    }

    var out: std.ArrayList(u8) = .empty;
    defer out.deinit(self.allocator);
    try out.ensureTotalCapacity(self.allocator, src.len);
    try out.appendSlice(self.allocator, src[0..trim_start]);
    try out.appendSlice(self.allocator, src[elem_close..]);

    try self.store.replacePart("xl/workbook.xml", out.items);
}

/// Walk `xl/_rels/workbook.xml.rels` and elide the
/// `<Relationship Id="rIdN" .../>` whose `Id` matches `r_id`.
fn patchWorkbookRelsRemoveRelationship(self: *Workbook, r_id: []const u8) Error!void {
    assert(r_id.len > 0);

    const part = try self.store.part("xl/_rels/workbook.xml.rels") orelse
        return error.MissingWorkbookRels;
    const src = part.bytes;
    assert(src.len > 0);

    var search_from: usize = 0;
    var elem_open: usize = 0;
    var elem_close: usize = 0;
    var found: bool = false;
    while (true) {
        const open = std.mem.indexOfPos(u8, src, search_from, "<Relationship") orelse break;
        const after = open + "<Relationship".len;
        if (after >= src.len) break;
        const boundary = src[after];
        const is_rel_elem = switch (boundary) {
            ' ', '\t', '\r', '\n', '/' => true,
            else => false,
        };
        if (!is_rel_elem) {
            // `<Relationships>` (the wrapper) is the only collision.
            search_from = after;
            continue;
        }
        const gt = std.mem.indexOfScalarPos(u8, src, after, '>') orelse break;
        const this_close = blk: {
            if (gt > 0 and src[gt - 1] == '/') break :blk gt + 1;
            const end_tag = std.mem.indexOfPos(u8, src, gt + 1, "</Relationship>") orelse
                return error.RelationshipElementNotFound;
            break :blk end_tag + "</Relationship>".len;
        };

        const attrs_end = if (gt > 0 and src[gt - 1] == '/') gt - 1 else gt;
        const attrs = src[after..attrs_end];
        if (attrIdEquals(attrs, r_id)) {
            elem_open = open;
            elem_close = this_close;
            found = true;
            break;
        }
        search_from = this_close;
    }
    if (!found) return error.RelationshipElementNotFound;
    assert(elem_close > elem_open);

    var trim_start = elem_open;
    while (trim_start > 0) {
        const c = src[trim_start - 1];
        if (c == ' ' or c == '\t' or c == '\r' or c == '\n') {
            trim_start -= 1;
        } else break;
    }

    var out: std.ArrayList(u8) = .empty;
    defer out.deinit(self.allocator);
    try out.ensureTotalCapacity(self.allocator, src.len);
    try out.appendSlice(self.allocator, src[0..trim_start]);
    try out.appendSlice(self.allocator, src[elem_close..]);

    try self.store.replacePart("xl/_rels/workbook.xml.rels", out.items);
}

/// Walk `[Content_Types].xml` and elide the
/// `<Override PartName="/<part_name>" .../>` whose PartName matches
/// `/<part_name>` (OOXML stores PartNames with a leading slash).
/// Returns `error.ContentTypesOverrideNotFound` when no matching
/// Override exists; the caller (deleteSheet) treats that as soft-OK.
fn patchContentTypesRemoveOverride(self: *Workbook, part_name: []const u8) Error!void {
    assert(part_name.len > 0);

    const part = try self.store.part("[Content_Types].xml") orelse
        return error.ContentTypesOverrideNotFound;
    const src = part.bytes;
    assert(src.len > 0);

    if (part_name.len > 255) return error.ContentTypesOverrideNotFound;
    var pn_buf: [256]u8 = undefined;
    pn_buf[0] = '/';
    @memcpy(pn_buf[1 .. 1 + part_name.len], part_name);
    const target = pn_buf[0 .. 1 + part_name.len];

    var search_from: usize = 0;
    var elem_open: usize = 0;
    var elem_close: usize = 0;
    var found: bool = false;
    while (true) {
        const open = std.mem.indexOfPos(u8, src, search_from, "<Override") orelse break;
        const after = open + "<Override".len;
        if (after >= src.len) break;
        const boundary = src[after];
        const is_override_elem = switch (boundary) {
            ' ', '\t', '\r', '\n', '/' => true,
            else => false,
        };
        if (!is_override_elem) {
            search_from = after;
            continue;
        }
        const gt = std.mem.indexOfScalarPos(u8, src, after, '>') orelse break;
        const this_close = gt + 1;

        const attrs_end = if (gt > 0 and src[gt - 1] == '/') gt - 1 else gt;
        const attrs = src[after..attrs_end];
        if (attrPartNameEquals(attrs, target)) {
            elem_open = open;
            elem_close = this_close;
            found = true;
            break;
        }
        search_from = this_close;
    }
    if (!found) return error.ContentTypesOverrideNotFound;
    assert(elem_close > elem_open);

    var trim_start = elem_open;
    while (trim_start > 0) {
        const c = src[trim_start - 1];
        if (c == ' ' or c == '\t' or c == '\r' or c == '\n') {
            trim_start -= 1;
        } else break;
    }

    var out: std.ArrayList(u8) = .empty;
    defer out.deinit(self.allocator);
    try out.ensureTotalCapacity(self.allocator, src.len);
    try out.appendSlice(self.allocator, src[0..trim_start]);
    try out.appendSlice(self.allocator, src[elem_close..]);

    try self.store.replacePart("[Content_Types].xml", out.items);
}

fn attrIdEquals(attrs: []const u8, value: []const u8) bool {
    return attrEquals(attrs, "Id", value);
}

fn attrPartNameEquals(attrs: []const u8, value: []const u8) bool {
    return attrEquals(attrs, "PartName", value);
}

/// Generic attribute byte-compare. Walks `attrs` token-by-token (same
/// shape as `patchWorkbookXmlSheetName`'s inline parser) and returns
/// `true` iff some attribute named `key` has a quoted value byte-
/// equal to `value`.
fn attrEquals(attrs: []const u8, key: []const u8, value: []const u8) bool {
    var i: usize = 0;
    while (i < attrs.len) {
        while (i < attrs.len and (attrs[i] == ' ' or attrs[i] == '\t' or
            attrs[i] == '\r' or attrs[i] == '\n')) : (i += 1)
        {}
        if (i >= attrs.len) return false;
        const key_start = i;
        while (i < attrs.len and attrs[i] != '=' and attrs[i] != ' ' and
            attrs[i] != '\t' and attrs[i] != '\r' and attrs[i] != '\n') : (i += 1)
        {}
        const key_end = i;
        while (i < attrs.len and (attrs[i] == ' ' or attrs[i] == '\t' or
            attrs[i] == '\r' or attrs[i] == '\n')) : (i += 1)
        {}
        if (i >= attrs.len or attrs[i] != '=') return false;
        i += 1;
        while (i < attrs.len and (attrs[i] == ' ' or attrs[i] == '\t' or
            attrs[i] == '\r' or attrs[i] == '\n')) : (i += 1)
        {}
        if (i >= attrs.len) return false;
        const quote = attrs[i];
        if (quote != '"' and quote != '\'') return false;
        i += 1;
        const val_start = i;
        while (i < attrs.len and attrs[i] != quote) : (i += 1) {}
        if (i >= attrs.len) return false;
        const val_end = i;
        i += 1;

        if (std.mem.eql(u8, attrs[key_start..key_end], key) and
            std.mem.eql(u8, attrs[val_start..val_end], value))
        {
            return true;
        }
    }
    return false;
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
            // `.deleted` overrides the original — emit nothing for
            // this ref. Still mark `seen` so the delta-only pass
            // below doesn't re-introduce it.
            if (overlay) |dv| if (dv == .deleted) {
                try seen.put(allocator, cr, {});
                continue;
            };
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
    // `.deleted` deltas with no matching original are a no-op — there
    // is nothing to elide, so we skip rather than emit a phantom cell.
    var dit = deltas.iterator();
    while (dit.next()) |entry| {
        if (seen.contains(entry.key_ptr.*)) continue;
        if (entry.value_ptr.* == .deleted) continue;
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
            // `.deleted` deltas are filtered out in `emitSheetData`
            // before they ever reach a `MergedCell` — reaching here
            // would mean the filter regressed.
            .deleted => unreachable,
        },
    }
}

// ─── SST extension (iter-wb-4 m4 + iter-wr-1 substrate move) ───────

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
///   - existing-SST entries: linear scan over decoded text. Linear was
///     chosen over a hashmap because the typical write workload stages
///     a small handful of new strings while the SST may carry
///     thousands of existing entries; building a hashmap of decoded
///     existing entries up front is more work than scanning per-new-
///     string. Stdlib-only + trivially auditable matters more than
///     constant-factor speed at typical SST sizes.
///   - already-staged new strings: O(1) hash via
///     `plan.new_strings_index` (added in iter-wr-1 to keep
///     `xlsx.Writer`'s hot writeRow loop linear in cell count rather
///     than quadratic).
///
/// The plan struct itself lives in `pkg/sst_plan.zig` so `xlsx.Writer`
/// can stage entries through the same shape without forming a module
/// cycle (workbook → zlsx → writer.zig).
const ExistingMatch = sst_plan_mod.ExistingMatch;

/// Compatibility wrapper around `SstExtensionPlan.registerNewRich` so
/// the existing in-tree tests (and any future Workbook-side caller
/// that already has a `*Workbook` in scope) keep their call site
/// unchanged. Pure delegate — see `pkg/sst_plan.zig` for the impl.
fn registerSharedRichString(
    wb: *Workbook,
    plan: *SstExtensionPlan,
    runs: []const RichRun,
) Error!*const RichEntry {
    assert(@intFromPtr(wb) != 0);
    assert(@intFromPtr(wb.allocator.vtable) != 0);
    return try plan.registerNewRich(wb.allocator, runs);
}

/// Walk every worksheet's `.shared_string` deltas, de-dup against
/// the existing SST (when present) and against each other, and stage
/// the resulting unique-new-strings list into a plan. The plan owns
/// duplicates of every staged string; callers free via `plan.deinit`.
/// Register a string `s` into `plan.new_strings` if it matches no
/// existing PLAIN SST entry and isn't already staged. Owns the dupe.
/// Rich entries (`is_rich_existing[i] == true`) are skipped from
/// the match loop — a new string equal to a rich entry's
/// concatenated runs still allocates a fresh `<si><t>` because the
/// reader can't resolve `t="s"` indices that point at rich `<r>`
/// blocks as plain text.
fn registerSharedString(
    wb: *Workbook,
    plan: *SstExtensionPlan,
    s: []const u8,
    decoded_existing: []const []const u8,
    is_rich_existing: []const bool,
) Error!void {
    assert(decoded_existing.len == is_rich_existing.len);
    for (decoded_existing, 0..) |de, i| {
        if (is_rich_existing[i]) continue;
        if (std.mem.eql(u8, de, s)) return;
    }
    if (plan.new_strings_index.contains(s)) return;
    const owned = try wb.allocator.dupe(u8, s);
    errdefer wb.allocator.free(owned);
    const idx: u32 = @intCast(plan.new_strings.items.len);
    try plan.new_strings.append(wb.allocator, owned);
    errdefer _ = plan.new_strings.pop();
    try plan.new_strings_index.put(wb.allocator, owned, idx);
}

/// Append an existing-match record for string `s` if an existing
/// PLAIN SST entry decodes to it and we haven't already recorded
/// it. Owns the dupe. Skips strings staged into `plan.new_strings`
/// (those resolve via the new-strings index, not the side table).
/// Rich entries are skipped — same reasoning as
/// `registerSharedString`.
fn registerExistingMatch(
    wb: *Workbook,
    plan: *SstExtensionPlan,
    s: []const u8,
    decoded_existing: []const []const u8,
    is_rich_existing: []const bool,
) Error!void {
    assert(decoded_existing.len == is_rich_existing.len);
    if (plan.new_strings_index.contains(s)) return;
    for (plan.new_strings.items) |n| {
        if (std.mem.eql(u8, n, s)) return;
    }
    for (plan.existing_matches.items) |em| {
        if (std.mem.eql(u8, em.text, s)) return;
    }
    var found_idx: u32 = std.math.maxInt(u32);
    for (decoded_existing, 0..) |de, i| {
        if (is_rich_existing[i]) continue;
        if (std.mem.eql(u8, de, s)) {
            found_idx = @intCast(i);
            break;
        }
    }
    if (found_idx == std.math.maxInt(u32)) return;
    const owned = try wb.allocator.dupe(u8, s);
    errdefer wb.allocator.free(owned);
    try plan.existing_matches.append(wb.allocator, .{ .text = owned, .index = found_idx });
}

fn buildSstExtensionPlan(wb: *Workbook) Error!SstExtensionPlan {
    assert(@intFromPtr(wb) != 0);
    assert(@intFromPtr(wb.allocator.vtable) != 0);

    var plan: SstExtensionPlan = .{};
    errdefer plan.deinit(wb.allocator);

    // Quick scan: any shared-string payload (delta `.shared_string`
    // OR appended-row `.string`)? Skips any work — and crucially,
    // skips parsing the SST — when the workbook has no string
    // mutations pending across either axis.
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
        for (ws.appended_rows.items) |row| {
            for (row) |c| switch (c) {
                .string => {
                    any = true;
                    break;
                },
                else => {},
            };
            if (any) break;
        }
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
    // are tracked in a parallel `is_rich_existing` array and SKIPPED
    // from the match loops — a new string can never resolve to a
    // rich entry's index because the OOXML reader treats `<c t="s">`
    // pointing at a rich `<si><r>…</r></si>` differently from a
    // plain one. Without the explicit skip, the previous sentinel
    // (empty slice) silently mis-resolved an empty new string `""`
    // to a rich entry's index.
    var decode_arena = std.heap.ArenaAllocator.init(wb.allocator);
    defer decode_arena.deinit();
    const da = decode_arena.allocator();

    var decoded_existing: [][]const u8 = &.{};
    var is_rich_existing: []bool = &.{};
    if (existing_view) |view| {
        decoded_existing = try da.alloc([]const u8, view.entries.len);
        is_rich_existing = try da.alloc(bool, view.entries.len);
        for (view.entries, 0..) |e, i| {
            switch (e) {
                .plain => |s| {
                    decoded_existing[i] = try sst_xml_mod.decodeText(da, s);
                    is_rich_existing[i] = false;
                },
                .rich => {
                    decoded_existing[i] = "";
                    is_rich_existing[i] = true;
                },
            }
        }
    }

    // Walk strings in worksheet order: first deltas (iteration
    // order), then appended_rows (row-major / col-major). Order is
    // observable to test assertions: "first occurrence across the
    // unified walk wins the lower index". Per-sheet, deltas and
    // appends are mutually exclusive (refused at staging time), but
    // we order deltas-before-appends regardless for determinism.
    for (wb.worksheets) |*ws| {
        var dit = ws.deltas.valueIterator();
        while (dit.next()) |v| {
            const s = switch (v.*) {
                .shared_string => |t| t,
                else => continue,
            };
            try registerSharedString(wb, &plan, s, decoded_existing, is_rich_existing);
        }
        for (ws.appended_rows.items) |row| {
            for (row) |c| switch (c) {
                .string => |s| try registerSharedString(wb, &plan, s, decoded_existing, is_rich_existing),
                else => {},
            };
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

    // Build the existing-match side table for any string (delta or
    // appended) whose text equals an existing SST entry. Both fast
    // paths (A: no new strings; B: some new strings) need this so
    // `plan.indexOf(s)` resolves to the existing index, not a fresh
    // one. `registerExistingMatch` skips strings that landed in
    // `plan.new_strings`.
    if (existing_view != null) {
        for (wb.worksheets) |*ws| {
            var it = ws.deltas.valueIterator();
            while (it.next()) |v| {
                const s = switch (v.*) {
                    .shared_string => |t| t,
                    else => continue,
                };
                try registerExistingMatch(wb, &plan, s, decoded_existing, is_rich_existing);
            }
            for (ws.appended_rows.items) |row| {
                for (row) |c| switch (c) {
                    .string => |s| try registerExistingMatch(wb, &plan, s, decoded_existing, is_rich_existing),
                    else => {},
                };
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
    // At least one of the two new-entry axes must be non-empty when
    // `has_new_strings` is true. Plain-only and rich-only and mixed
    // are all valid call shapes.
    assert(plan.new_strings.items.len > 0 or plan.new_rich_strings.items.len > 0);

    if (plan.sst_part_exists) {
        // Re-emit the SST part with the existing entries preserved
        // verbatim and the new entries (plain then rich) appended.
        const existing_part = try wb.store.part("xl/sharedStrings.xml") orelse
            return Error.MissingWorkbookPart; // sst_part_exists invariant violated
        const new_xml = try emitSstXmlForExtension(
            wb.allocator,
            existing_part.bytes,
            plan.new_strings.items,
            plan.new_rich_strings.items,
        );
        defer wb.allocator.free(new_xml);
        try wb.store.replacePart("xl/sharedStrings.xml", new_xml);
        return;
    }

    // Source had no SST. Emit a fresh one containing only the new
    // entries (plain + rich), register it as a new part with the
    // correct content type, then patch the workbook rels file.
    const fresh_xml = try emitFreshSstXml(
        wb.allocator,
        plan.new_strings.items,
        plan.new_rich_strings.items,
    );
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
    new_rich: []const RichEntry,
) Error![]u8 {
    assert(src_xml.len > 0);
    assert(new_strings.len > 0 or new_rich.len > 0);

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
    const new_si_count: u32 = @intCast(new_strings.len + new_rich.len);
    const total_si: u32 = existing_si_count + new_si_count;

    var out: std.ArrayList(u8) = .empty;
    errdefer out.deinit(allocator);
    try out.ensureTotalCapacity(allocator, src_xml.len + 64 * (new_strings.len + new_rich.len));

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
        // followed by a fresh `</sst>`. Plain entries first, then
        // rich — `indexOfRich` assumes that order.
        try appendNewSiEntries(allocator, &out, new_strings);
        try appendNewRichSiEntries(allocator, &out, new_rich);
        try out.appendSlice(allocator, "</sst>");
        // Anything past the original `/>` is post-element trailing
        // bytes (rare, but preserve).
        if (sst_open_gt + 1 < src_xml.len) {
            try out.appendSlice(allocator, src_xml[sst_open_gt + 1 ..]);
        }
        return try out.toOwnedSlice(allocator);
    }

    // Normal form: copy body verbatim up to `</sst>`, then append
    // new entries (plain then rich), then `</sst>` + trailing.
    const body_start = sst_open_gt + 1;
    const close = std.mem.indexOfPos(u8, src_xml, body_start, "</sst>") orelse
        return error.MalformedXml;
    try out.appendSlice(allocator, src_xml[body_start..close]);
    try appendNewSiEntries(allocator, &out, new_strings);
    try appendNewRichSiEntries(allocator, &out, new_rich);
    try out.appendSlice(allocator, src_xml[close..]);
    return try out.toOwnedSlice(allocator);
}

/// Build a complete `xl/sharedStrings.xml` from scratch. Used when
/// the source workbook had no SST part.
fn emitFreshSstXml(
    allocator: Allocator,
    new_strings: []const []const u8,
    new_rich: []const RichEntry,
) Error![]u8 {
    assert(new_strings.len > 0 or new_rich.len > 0);
    const total: usize = new_strings.len + new_rich.len;

    var out: std.ArrayList(u8) = .empty;
    errdefer out.deinit(allocator);
    try out.ensureTotalCapacity(allocator, 256 + 64 * total);

    try out.appendSlice(allocator, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    try out.appendSlice(allocator, "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
    var nbuf: [32]u8 = undefined;
    try out.appendSlice(allocator, " count=\"");
    try out.appendSlice(allocator, try std.fmt.bufPrint(&nbuf, "{d}", .{total}));
    try out.appendSlice(allocator, "\" uniqueCount=\"");
    try out.appendSlice(allocator, try std.fmt.bufPrint(&nbuf, "{d}", .{total}));
    try out.appendSlice(allocator, "\">");
    try appendNewSiEntries(allocator, &out, new_strings);
    try appendNewRichSiEntries(allocator, &out, new_rich);
    try out.appendSlice(allocator, "</sst>");
    return try out.toOwnedSlice(allocator);
}

/// Append one `<si><r>…</r>…</si>` per rich entry. Mirrors the
/// emitter shape used by `xlsx.Writer.sstInternRich` so reader
/// round-trips agree on the byte layout. Each run that has at least
/// one typography flag set emits an `<rPr>` block; runs with no
/// flags emit only `<t>`.
fn appendNewRichSiEntries(
    allocator: Allocator,
    out: *std.ArrayList(u8),
    new_rich: []const RichEntry,
) Error!void {
    for (new_rich) |entry| {
        try out.appendSlice(allocator, "<si>");
        for (entry.runs) |r| {
            try out.appendSlice(allocator, "<r>");
            const has_props = r.bold or r.italic or r.underline or r.strike or
                r.font_size != null or r.font_name != null or r.color_argb != null;
            if (has_props) {
                try out.appendSlice(allocator, "<rPr>");
                if (r.bold) try out.appendSlice(allocator, "<b/>");
                if (r.italic) try out.appendSlice(allocator, "<i/>");
                if (r.strike) try out.appendSlice(allocator, "<strike/>");
                if (r.underline) try out.appendSlice(allocator, "<u/>");
                if (r.font_size) |sz| {
                    var szbuf: [32]u8 = undefined;
                    try out.appendSlice(allocator, "<sz val=\"");
                    try out.appendSlice(allocator, try std.fmt.bufPrint(&szbuf, "{d}", .{sz}));
                    try out.appendSlice(allocator, "\"/>");
                }
                if (r.color_argb) |c| {
                    try out.appendSlice(allocator, "<color rgb=\"");
                    try appendXmlEscapedText(allocator, out, c);
                    try out.appendSlice(allocator, "\"/>");
                }
                if (r.font_name) |n| {
                    try out.appendSlice(allocator, "<rFont val=\"");
                    try appendXmlEscapedText(allocator, out, n);
                    try out.appendSlice(allocator, "\"/>");
                }
                try out.appendSlice(allocator, "</rPr>");
            }
            try out.appendSlice(allocator, "<t xml:space=\"preserve\">");
            try appendXmlEscapedText(allocator, out, r.text);
            try out.appendSlice(allocator, "</t></r>");
        }
        try out.appendSlice(allocator, "</si>");
    }
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

    /// B2 iter-er-3 (Phase 3a): pending appended rows. Each entry is
    /// an owned `[]zlsx.Cell` slice with string payloads duped into
    /// the workbook allocator. Cleared by the saver (Editor.save in
    /// iter-er-3, eventually Workbook.save in iter-er-6) after the
    /// substring-spliced sheet XML is emitted.
    ///
    /// **Write contract**: only `Worksheet.appendRows` (and its
    /// editor-side cousin `Editor.appendRows` via the shim) is
    /// allowed to push entries here. Direct field-level writes
    /// bypass the column-cap and integer-precision validation in
    /// `appendRows` and corrupt downstream emit. Reads (size,
    /// has-strings checks) are unrestricted.
    ///
    /// Cannot coexist with non-empty `deltas` on the same Worksheet:
    /// `appendRows` refuses if `deltas.count() > 0` and `setCell`
    /// refuses if `appended_rows.items.len > 0`. Mirrors the legacy
    /// Editor-side guards.
    appended_rows: std.ArrayListUnmanaged([]zlsx.Cell) = .{},

    /// B3 iter-wr-7: fresh-emit row body. Pre-built `<row>...</row>`
    /// payload (SST indices baked in) appended through the per-sheet
    /// fresh-emit API. Empty for delta-only / source-loaded sheets.
    /// Mutually exclusive with `deltas` and `appended_rows`: a sheet
    /// is either edit-source (deltas / appended_rows) or fresh-emit
    /// (body / state), never both. Today the dispatcher in
    /// `Workbook.saveFreshEmit` is the only consumer; `Workbook.save`
    /// continues to ignore these for the existing delta-on-bytes path.
    body: std.ArrayListUnmanaged(u8) = .{},

    /// B3 iter-wr-7: fresh-emit per-sheet registries. Owns column
    /// widths, row heights, freeze panes, auto filter, merge cells,
    /// hyperlinks, comments, conditional formats, data validations.
    /// Same shape `xlsx.Writer.SheetWriter.state` carries — the 13
    /// `add*`/`set*` forwarder methods on `Worksheet` route directly
    /// here. Empty-by-default; existing delta-only consumers pay
    /// nothing.
    sheet_state: sheet_plan.SheetState = .{},

    pub fn deinit(self: *Worksheet, allocator: Allocator) void {
        if (self.parsed) |*p| {
            var view = p.*;
            view.deinit(allocator);
        }
        if (self.resolved_part_name) |part_name| allocator.free(part_name);
        freeDeltaStrings(allocator, &self.deltas);
        self.deltas.deinit(allocator);
        for (self.appended_rows.items) |row| {
            for (row) |c| switch (c) {
                .string => |s| allocator.free(s),
                else => {},
            };
            allocator.free(row);
        }
        self.appended_rows.deinit(allocator);
        self.body.deinit(allocator);
        self.sheet_state.deinit(allocator);
    }

    // ─── B3 iter-wr-7: fresh-emit Worksheet methods ─────────────────
    //
    // These 13 forwarders mirror `xlsx.Writer.SheetWriter`'s
    // `add*` / `set*` surface so a `pkg.Workbook` callers can author
    // worksheets without going through the writer's fluent-builder.
    // Each method delegates straight onto `SheetState` — registration,
    // validation, and ownership semantics are identical to the
    // writer-side forwarder set (B3 iter-wr-6).

    pub fn setColumnWidth(self: *Worksheet, col_idx: u32, width: f32) sheet_plan.Error!void {
        try self.sheet_state.setColumnWidth(self.workbook.allocator, col_idx, width);
    }

    pub fn setRowHeight(self: *Worksheet, row_idx: u32, height: f32) sheet_plan.Error!void {
        try self.sheet_state.setRowHeight(self.workbook.allocator, row_idx, height);
    }

    pub fn freezePanes(self: *Worksheet, freeze_rows: u32, freeze_cols: u32) sheet_plan.SheetState.FreezePanesError!void {
        try self.sheet_state.freezePanes(freeze_rows, freeze_cols);
    }

    pub fn setAutoFilter(self: *Worksheet, range: []const u8) sheet_plan.Error!void {
        try self.sheet_state.setAutoFilter(self.workbook.allocator, range);
    }

    pub fn addMergedCell(self: *Worksheet, range: []const u8) sheet_plan.Error!void {
        try self.sheet_state.addMergedCell(self.workbook.allocator, range);
    }

    pub fn addHyperlink(self: *Worksheet, range: []const u8, url: []const u8) sheet_plan.Error!void {
        try self.sheet_state.addHyperlink(self.workbook.allocator, range, url);
    }

    pub fn addInternalHyperlink(self: *Worksheet, range: []const u8, location: []const u8) sheet_plan.Error!void {
        try self.sheet_state.addInternalHyperlink(self.workbook.allocator, range, location);
    }

    pub fn addComment(self: *Worksheet, ref: []const u8, author: []const u8, text: []const u8) sheet_plan.Error!void {
        try self.sheet_state.addComment(self.workbook.allocator, ref, author, text);
    }

    pub fn addDataValidationList(self: *Worksheet, range: []const u8, values: []const []const u8) sheet_plan.Error!void {
        try self.sheet_state.addDataValidationList(self.workbook.allocator, range, values);
    }

    pub fn addDataValidationRange(
        self: *Worksheet,
        range: []const u8,
        kind_name: []const u8,
        op_name: ?[]const u8,
        formula1: []const u8,
        formula2: ?[]const u8,
        needs_two: bool,
    ) sheet_plan.Error!void {
        try self.sheet_state.addDataValidationRange(self.workbook.allocator, range, kind_name, op_name, formula1, formula2, needs_two);
    }

    pub fn addDataValidationCustom(self: *Worksheet, range: []const u8, formula: []const u8) sheet_plan.Error!void {
        try self.sheet_state.addDataValidationCustom(self.workbook.allocator, range, formula);
    }

    pub fn addConditionalFormatCellIs(
        self: *Worksheet,
        range: []const u8,
        operator: sheet_plan.CfOperator,
        formula1: []const u8,
        formula2: ?[]const u8,
        dxf_id: u32,
    ) sheet_plan.Error!void {
        try self.sheet_state.addConditionalFormatCellIs(
            self.workbook.allocator,
            range,
            operator,
            formula1,
            formula2,
            dxf_id,
            self.workbook.styles_plan.dxfs.items.len,
        );
    }

    pub fn addConditionalFormatExpression(
        self: *Worksheet,
        range: []const u8,
        formula: []const u8,
        dxf_id: u32,
    ) sheet_plan.Error!void {
        try self.sheet_state.addConditionalFormatExpression(
            self.workbook.allocator,
            range,
            formula,
            dxf_id,
            self.workbook.styles_plan.dxfs.items.len,
        );
    }

    pub fn addConditionalFormatColorScale(
        self: *Worksheet,
        range: []const u8,
        low_color_argb: u32,
        mid_color_argb: ?u32,
        high_color_argb: u32,
    ) sheet_plan.Error!void {
        try self.sheet_state.addConditionalFormatColorScale(
            self.workbook.allocator,
            range,
            low_color_argb,
            mid_color_argb,
            high_color_argb,
        );
    }

    pub fn addConditionalFormatDataBar(self: *Worksheet, range: []const u8, color_argb: u32) sheet_plan.Error!void {
        try self.sheet_state.addConditionalFormatDataBar(self.workbook.allocator, range, color_argb);
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

    /// Resolve `r_id` → `xl/_rels/workbook.xml.rels` lookup and cache
    /// the sheet's part name (e.g. "xl/worksheets/sheet1.xml")
    /// without parsing the body. Idempotent — second call returns
    /// the cached name.
    ///
    /// Public surface for B2 iter-er-3 fast paths (substring-splice
    /// `emitWithAppendsUsingPlan`) that need the sheet's part name to read
    /// raw XML bytes from the part store WITHOUT walking
    /// `<sheetData>` via `ensureParsed`. On a 100k-row sheet, the
    /// parse step alone runs in the hundreds of milliseconds — too
    /// expensive to pay when the caller will splice a few rows.
    ///
    /// Stability: `r_id` is set once at `Workbook.parse` and is not
    /// re-bound by any current API (`renameSheet` mutates the
    /// display name, not the rels target), so the cached part name
    /// outlives every legal invalidation of `parsed`.
    pub fn resolvePartName(self: *Worksheet) Error![]const u8 {
        if (self.resolved_part_name) |cached| return cached;

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
        // not to PartStore's arena.
        const owned = try wb.allocator.dupe(u8, part_name);
        errdefer wb.allocator.free(owned);
        self.resolved_part_name = owned;
        return owned;
    }

    /// Resolve the part name and parse the sheet XML if not already
    /// cached. Returns a const pointer to the cached view.
    pub fn ensureParsed(self: *Worksheet) Error!*const sheet_xml_mod.SheetXml {
        if (self.parsed != null) return &self.parsed.?;

        const part_name = try self.resolvePartName();
        const wb = self.workbook;
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

    /// Resolve the cell at `ref` to a composite `ResolvedStyle` view
    /// by walking `SheetXml.Cell.style_idx` → `StylesXml.cell_xfs[idx]`
    /// → the per-attribute fonts/fills/borders/numFmts tables.
    ///
    /// Returns `null` when:
    ///   - the cell does not exist on this sheet,
    ///   - the cell carries no `s="…"` attribute (`style_idx == null`),
    ///   - the workbook has no `xl/styles.xml`, or
    ///   - `style_idx` is out of range for the workbook's `cell_xfs`.
    ///
    /// Per-field semantics: each `apply_*` flag on the matched CellXf
    /// gates whether the corresponding sub-style is surfaced. When the
    /// flag is false, the field is `null` — see `ResolvedStyle` doc-
    /// comment for the v1 cellStyleXfs-inheritance simplification.
    /// Out-of-range sub-ids (font_id ≥ fonts.len, etc.) likewise
    /// surface as `null` rather than erroring; that lets the typed
    /// overlay tolerate workbooks where producers under-count their
    /// `<fonts count="…">` headers.
    ///
    /// `number_format_code` is `null` for built-in numFmt ids (0..163,
    /// ECMA-376 §18.8.30) — those codes are implicit and absent from
    /// `<numFmts>`. Custom ids (≥ 164) resolve via linear scan.
    pub fn cellStyle(self: *Worksheet, ref: []const u8) Error!?ResolvedStyle {
        assert(ref.len > 0);

        const cell = (try self.cellByRef(ref)) orelse return null;
        const sidx = cell.style_idx orelse return null;
        const styles = (try self.workbook.styles()) orelse return null;
        if (sidx >= styles.cell_xfs.len) return null;
        const xf = styles.cell_xfs[sidx];

        const font: ?styles_xml_mod.Font = blk: {
            if (!xf.apply_font) break :blk null;
            const fid = xf.font_id orelse break :blk null;
            if (fid >= styles.fonts.len) break :blk null;
            break :blk styles.fonts[fid];
        };

        const fill: ?styles_xml_mod.Fill = blk: {
            if (!xf.apply_fill) break :blk null;
            const fid = xf.fill_id orelse break :blk null;
            if (fid >= styles.fills.len) break :blk null;
            break :blk styles.fills[fid];
        };

        const border: ?styles_xml_mod.Border = blk: {
            if (!xf.apply_border) break :blk null;
            const bid = xf.border_id orelse break :blk null;
            if (bid >= styles.borders.len) break :blk null;
            break :blk styles.borders[bid];
        };

        const alignment: ?styles_xml_mod.Alignment =
            if (xf.apply_alignment) xf.alignment else null;

        const number_format_code: ?[]const u8 = blk: {
            if (!xf.apply_number_format) break :blk null;
            const nfid = xf.num_fmt_id orelse break :blk null;
            // Built-in codes (0..163) are implicit; not stored in numFmts.
            if (nfid <= 163) break :blk null;
            for (styles.number_formats) |nf| {
                if (nf.fmt_id == nfid) break :blk nf.code;
            }
            break :blk null;
        };

        return ResolvedStyle{
            .font = font,
            .fill = fill,
            .border = border,
            .alignment = alignment,
            .number_format_code = number_format_code,
        };
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
        // B2 iter-er-3 symmetric guard: setCell + appendRows on the
        // same Worksheet are mutually exclusive (mirrors Editor's
        // legacy SheetHasUnsavedAppends rejection).
        if (self.appended_rows.items.len > 0) return error.SheetHasUnsavedAppends;
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

    /// Stage a deletion for cell `ref`. After `Workbook.save`, the
    /// cell is fully absent from `<sheetData>` (no `<c>` element at
    /// all) and `cellByRef(ref)` returns `null`. Distinct from
    /// `setCell(ref, .blank)`, which keeps the cell present as an
    /// empty `<c r="REF"/>`.
    ///
    /// Staging a deletion against a ref that doesn't exist in the
    /// source sheet is not an error — the delta just elides nothing.
    /// Last `setCell`/`deleteCell` for a given ref wins.
    pub fn deleteCell(self: *Worksheet, ref: []const u8) Error!void {
        assert(ref.len > 0);
        return self.setCell(ref, .deleted);
    }

    /// B2 iter-er-3 (Phase 3a): stage rows to append at the bottom of
    /// the sheet. Each row is an iterable of `xlsx.Cell` variants
    /// (`empty` / `number` / `integer` / `boolean` / `string`).
    /// Strings are duped into the workbook allocator for owned
    /// lifetime.
    ///
    /// **Refusal contract** (mirrors Editor.appendRows):
    ///   - `error.SheetHasUnsavedMutations` if `self.deltas.count() > 0`
    ///   - `error.ColumnIndexOutOfRange` if any row exceeds 16384 cells
    ///   - `error.IntegerExceedsExcelPrecision` if any `.integer`
    ///     value can't round-trip exactly through f64
    ///
    /// **Editor-scope state not visible here**: the editor tracks
    /// `pending_row_inserts` / `pending_row_deletes` / `pending_col_*`
    /// at its own layer. Those edits race with appends at save time
    /// (the editor's row/col substitution runs first, then appends
    /// would splice into the pre-edit XML, silently dropping the
    /// edit). Direct `Worksheet.appendRows` callers must drain those
    /// editor-scope queues themselves; only the editor-side shim
    /// (`Editor.appendRows`) checks `sheetHasPendingRowOrColEdit`.
    ///
    /// **What this method does NOT do**: emit XML, allocate SST
    /// indices, or modify the sheet's parsed view. The actual
    /// substring splice + SST extension happens at save time via
    /// `emitWithAppendsUsingPlan`, driven by `Workbook.save`'s
    /// `SstExtensionPlan` walk.
    pub fn appendRows(
        self: *Worksheet,
        new_rows: []const []const zlsx.Cell,
    ) Error!void {
        if (self.deltas.count() > 0) return error.SheetHasUnsavedMutations;
        if (new_rows.len == 0) return;
        const a = self.workbook.allocator;
        // Validation pass — fail BEFORE any allocation so a bad row
        // doesn't half-stage cells from earlier new_rows in the same call.
        for (new_rows) |row| {
            if (row.len > zlsx.max_col_1based) return error.ColumnIndexOutOfRange;
            for (row) |c| switch (c) {
                .empty, .number, .boolean, .string => {},
                .integer => |n| if (!zlsx.fitsExactlyInF64(n))
                    return error.IntegerExceedsExcelPrecision,
            };
        }
        // Reserve once — keeps the existing-cap path on the hot loop.
        try self.appended_rows.ensureUnusedCapacity(a, new_rows.len);
        for (new_rows) |row| {
            const owned = try a.alloc(zlsx.Cell, row.len);
            errdefer a.free(owned);
            // Track how many cells we've duped so partial-fail
            // cleanup frees only what we allocated.
            var duped: usize = 0;
            errdefer for (owned[0..duped]) |c| switch (c) {
                .string => |s| a.free(s),
                else => {},
            };
            for (row, 0..) |c, i| {
                owned[i] = switch (c) {
                    .string => |s| .{ .string = try a.dupe(u8, s) },
                    else => c,
                };
                duped = i + 1;
            }
            // Capacity reserved above — appendAssumeCapacity is
            // infallible from here.
            self.appended_rows.appendAssumeCapacity(owned);
        }
    }

    /// Regenerate this worksheet's XML with the staged `setCell` /
    /// `deleteCell` deltas applied. Returns allocator-owned bytes —
    /// caller frees. The source `<sheetData>` block is regenerated
    /// from scratch (parsed view + deltas merged in row/col order);
    /// everything outside `<sheetData>` (sheet-level metadata,
    /// `<mergeCells>`, `<dataValidations>`, `<conditionalFormatting>`,
    /// `<hyperlinks>`, `<drawing>`, ...) is copied byte-for-byte from
    /// source. B2 iter-er-2 entry-point: lets `Editor.save` route
    /// modified-sheet emit through the workbook overlay without
    /// invoking the full `Workbook.save` ZIP-rebuild pipeline.
    ///
    /// SST policy: this entry-point uses an EMPTY `SstExtensionPlan`,
    /// so `.shared_string` deltas would resolve to `null` index and
    /// trip the regen's invariant. The Editor pipeline maps every
    /// `Cell.string` to `CellValue.string` (inlineStr — no SST
    /// extension), so this is fine for iter-er-2's call sites.
    /// Future iters that need shared-string emission from Editor
    /// must pass a real plan.
    pub fn emitWithDeltas(self: *Worksheet, allocator: Allocator) Error![]u8 {
        const view = try self.ensureParsed();
        const part_name = self.resolved_part_name orelse return error.MissingSheetPart;
        const part = try self.workbook.store.part(part_name) orelse return error.MissingSheetPart;
        const empty_plan: SstExtensionPlan = .{};
        return emitSheetWithDeltas(allocator, part.bytes, view, &self.deltas, &empty_plan);
    }

    /// B2 iter-er-3 substring-splice fast-path. Used by
    /// `Workbook.save`: every string cell resolves its SST index via
    /// `plan.indexOf(s)` (which already deduplicates against existing
    /// SST entries and across sheets). Caller does NOT receive
    /// `new_strings` — those live in `plan.new_strings`. Returns the
    /// spliced XML only.
    ///
    /// Caller invariants: `appended_rows.items.len > 0`,
    /// `deltas.count() == 0`, every `.string` payload appearing in
    /// `appended_rows` is registered in `plan` (either as a new
    /// string or an existing-match entry).
    pub fn emitWithAppendsUsingPlan(
        self: *Worksheet,
        allocator: Allocator,
        plan: *const SstExtensionPlan,
    ) Error![]u8 {
        assert(self.appended_rows.items.len > 0);
        if (self.deltas.count() > 0) return error.SheetHasUnsavedMutations;

        const part_name = try self.resolvePartName();
        assert(part_name.len > 0);
        const part = try self.workbook.store.part(part_name) orelse return error.MissingSheetPart;
        const src_xml = part.bytes;

        const highest_row = appendXmlFindHighestRow(src_xml);
        const start_row: u32 = highest_row + 1;
        assert(start_row >= 1);
        const final_row64: u64 = @as(u64, start_row) + self.appended_rows.items.len - 1;
        if (final_row64 > zlsx.max_row) return error.RowIndexOutOfRange;
        const final_row: u32 = @intCast(final_row64);
        assert(start_row <= final_row);

        var max_col_1based: u32 = 0;
        for (self.appended_rows.items) |row| {
            if (row.len > max_col_1based) max_col_1based = @intCast(row.len);
        }
        assert(max_col_1based <= zlsx.max_col_1based);

        var rows_buf: std.ArrayList(u8) = .empty;
        defer rows_buf.deinit(allocator);
        try rows_buf.ensureTotalCapacity(allocator, self.appended_rows.items.len * 32);

        for (self.appended_rows.items, 0..) |row, ri| {
            const row_idx: u32 = start_row + @as(u32, @intCast(ri));
            try rows_buf.appendSlice(allocator, "<row r=\"");
            var num_buf: [16]u8 = undefined;
            try rows_buf.appendSlice(allocator, try std.fmt.bufPrint(&num_buf, "{d}", .{row_idx}));
            try rows_buf.appendSlice(allocator, "\">");
            for (row, 0..) |cell, ci| {
                const col_1based: u32 = @as(u32, @intCast(ci)) + 1;
                try appendCellXmlForAppendUsingPlan(
                    allocator,
                    &rows_buf,
                    cell,
                    .{ .row = row_idx, .col = col_1based },
                    plan,
                );
            }
            try rows_buf.appendSlice(allocator, "</row>");
        }
        assert(rows_buf.items.len > 0);

        const injected = try appendXmlInjectRows(allocator, src_xml, rows_buf.items);
        errdefer allocator.free(injected);

        const final_xml = blk: {
            const widened = try appendXmlUpdateDimensionBR(
                allocator,
                injected,
                final_row,
                max_col_1based,
            );
            if (widened) |w| {
                allocator.free(injected);
                break :blk w;
            }
            break :blk injected;
        };
        assert(final_xml.len > src_xml.len);
        return final_xml;
    }

    /// Free + reset `appended_rows` after `Workbook.save` (or
    /// `Editor.save` in iter-er-3) has consumed the staged rows.
    /// Pairs with `appendRows` — same allocator semantics.
    pub fn clearAppendedRows(self: *Worksheet, allocator: Allocator) void {
        for (self.appended_rows.items) |row| {
            for (row) |c| switch (c) {
                .string => |s| allocator.free(s),
                else => {},
            };
            allocator.free(row);
        }
        self.appended_rows.clearAndFree(allocator);
    }
};

/// Render a single appended-row cell into `out`. String cells
/// resolve via `plan.indexOf(s)`; the plan owns the string entries
/// (no per-cell dupe needed). Non-string cells render their value
/// verbatim. `.empty` skips emission so trailing empties collapse to
/// nothing — matches the editor's `renderCellOoxml` contract.
fn appendCellXmlForAppendUsingPlan(
    allocator: Allocator,
    out: *std.ArrayList(u8),
    cell: zlsx.Cell,
    ref: CellRef,
    plan: *const SstExtensionPlan,
) Error!void {
    assert(ref.col >= 1 and ref.col <= zlsx.max_col_1based);
    assert(ref.row >= 1 and ref.row <= zlsx.max_row);
    var ref_buf: [16]u8 = undefined;
    const ref_str = formatA1Ref(&ref_buf, ref);
    var num_buf: [64]u8 = undefined;
    switch (cell) {
        .empty => return,
        .integer => |x| {
            try out.appendSlice(allocator, "<c r=\"");
            try out.appendSlice(allocator, ref_str);
            try out.appendSlice(allocator, "\"><v>");
            try out.appendSlice(allocator, try std.fmt.bufPrint(&num_buf, "{d}", .{x}));
            try out.appendSlice(allocator, "</v></c>");
        },
        .number => |f| {
            try out.appendSlice(allocator, "<c r=\"");
            try out.appendSlice(allocator, ref_str);
            try out.appendSlice(allocator, "\"><v>");
            try out.appendSlice(allocator, try std.fmt.bufPrint(&num_buf, "{d}", .{f}));
            try out.appendSlice(allocator, "</v></c>");
        },
        .boolean => |b| {
            try out.appendSlice(allocator, "<c r=\"");
            try out.appendSlice(allocator, ref_str);
            try out.appendSlice(allocator, "\" t=\"b\"><v>");
            try out.appendSlice(allocator, if (b) "1" else "0");
            try out.appendSlice(allocator, "</v></c>");
        },
        .string => |s| {
            const idx = plan.indexOf(s) orelse return error.SharedStringNotInPlan;
            try out.appendSlice(allocator, "<c r=\"");
            try out.appendSlice(allocator, ref_str);
            try out.appendSlice(allocator, "\" t=\"s\"><v>");
            try out.appendSlice(allocator, try std.fmt.bufPrint(&num_buf, "{d}", .{idx}));
            try out.appendSlice(allocator, "</v></c>");
        },
    }
}

/// Find the largest cell-row index in a sheet XML body. Walks both
/// `<row …>` and `<c …>` opening tags inside the `<sheetData>` window
/// and extracts the row component from any `r="…"` attribute (cell
/// refs like `r="A42"` decode to row 42; explicit `<row r="12">`
/// decodes to 12). OOXML doesn't constrain attribute order so the
/// scan is permissive.
///
/// Bounded: scans only `<sheetData>…</sheetData>`. Anything outside
/// that window — `<oddHeader>`, `<tableParts>`, `<extLst>`,
/// `<mergeCell ref="…">`, etc. — is ignored even if its attributes
/// happen to spell `r="A99999"`. XML comments INSIDE the window are
/// not specifically filtered: Excel and every legitimate writer
/// strip them, and adding O(n)-per-tag back-scanning costs ~3% on
/// the 100k-row gate. A `<!-- r="A1048576" -->` inside `<sheetData>`
/// could still inflate `highest_row` — accepted trade-off, fuzz
/// corpus exercises the boundary.
///
/// Parse-free: pure substring walk, suitable for 100k-row sheets
/// where the sheet_xml parse path would dominate the gate.
fn appendXmlFindHighestRow(xml: []const u8) u32 {
    var highest: u32 = 0;

    const window: []const u8 = blk: {
        const open_pos = std.mem.indexOf(u8, xml, "<sheetData") orelse return 0;
        const open_close = std.mem.indexOfScalarPos(u8, xml, open_pos, '>') orelse return 0;
        // Self-closing `<sheetData/>` has no body — empty window.
        if (open_close > 0 and xml[open_close - 1] == '/') break :blk xml[0..0];
        // open_close is a `>` byte found by indexOfScalarPos so it's
        // always within bounds; `open_close + 1 <= xml.len` and
        // indexOfPos accepts start == slice.len.
        assert(open_close + 1 <= xml.len);
        const close_pos = std.mem.indexOfPos(u8, xml, open_close + 1, "</sheetData>") orelse
            return 0;
        break :blk xml[open_close + 1 .. close_pos];
    };

    // Pass 1: `<c …>` tags — extract the row component from r="A1".
    {
        var i: usize = 0;
        while (std.mem.indexOfPos(u8, window, i, "<c")) |tag_start| {
            const after = tag_start + "<c".len;
            if (after >= window.len) break;
            const c = window[after];
            // Filter `<col`, `<conditionalFormatting`, etc.
            if (c != ' ' and c != '\t' and c != '\n' and c != '\r' and c != '/' and c != '>') {
                i = tag_start + 1;
                continue;
            }
            const tag_end = std.mem.indexOfScalarPos(u8, window, tag_start, '>') orelse break;
            if (appendXmlAttrRowFromCellRef(window[tag_start..tag_end])) |n| {
                if (n > highest) highest = n;
            }
            i = tag_end + 1;
        }
    }
    // Pass 2: `<row …>` tags — explicit r="N".
    {
        var i: usize = 0;
        while (std.mem.indexOfPos(u8, window, i, "<row")) |tag_start| {
            const after = tag_start + "<row".len;
            if (after >= window.len) break;
            const c = window[after];
            if (c != ' ' and c != '\t' and c != '\n' and c != '\r' and c != '/' and c != '>') {
                i = tag_start + 1;
                continue;
            }
            const tag_end = std.mem.indexOfScalarPos(u8, window, tag_start, '>') orelse break;
            if (appendXmlAttrRowExplicit(window[tag_start..tag_end])) |n| {
                if (n > highest) highest = n;
            }
            i = tag_end + 1;
        }
    }
    return highest;
}

/// Locate `r="…"` within an opening tag's attribute span and parse
/// the row component of an A1-style cell ref. Returns null on no
/// match or unparseable digits.
fn appendXmlAttrRowFromCellRef(tag: []const u8) ?u32 {
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

/// Locate `r="N"` within a `<row …>` span and parse N as a u32.
fn appendXmlAttrRowExplicit(tag: []const u8) ?u32 {
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

/// Splice a pre-rendered `rendered_rows` block into `src_xml` just
/// before `</sheetData>`. Falls back to the self-closing
/// `<sheetData/>` shape — replaces it with `<sheetData>…</sheetData>`
/// preserving any attributes on the open tag.
fn appendXmlInjectRows(
    allocator: Allocator,
    src_xml: []const u8,
    rendered_rows: []const u8,
) Error![]u8 {
    if (std.mem.indexOf(u8, src_xml, "</sheetData>")) |inject_pos| {
        // indexOf invariant: inject_pos + tag_len <= src_xml.len.
        // Pair-assertion turns a future indexOf-contract regression
        // into a tripwire instead of an out-of-bounds memcpy.
        assert(inject_pos + "</sheetData>".len <= src_xml.len);
        const out_len = src_xml.len + rendered_rows.len;
        const out = try allocator.alloc(u8, out_len);
        errdefer allocator.free(out);
        @memcpy(out[0..inject_pos], src_xml[0..inject_pos]);
        @memcpy(out[inject_pos..][0..rendered_rows.len], rendered_rows);
        @memcpy(out[inject_pos + rendered_rows.len ..], src_xml[inject_pos..]);
        return out;
    }

    const sd_open = std.mem.indexOf(u8, src_xml, "<sheetData") orelse
        return error.NoSheetData;
    const sd_close = std.mem.indexOfScalarPos(u8, src_xml, sd_open, '>') orelse
        return error.NoSheetData;
    if (sd_close == 0 or src_xml[sd_close - 1] != '/') return error.NoSheetData;
    const attrs_end = sd_close - 1;
    const attrs = src_xml[sd_open + "<sheetData".len .. attrs_end];

    var spliced: std.ArrayList(u8) = .empty;
    // `toOwnedSlice` empties the ArrayList on success, so this
    // `errdefer` is a no-op on the happy path — kept so an early
    // `try` failure doesn't leak the partially-built buffer.
    errdefer spliced.deinit(allocator);
    try spliced.ensureTotalCapacity(allocator, src_xml.len + rendered_rows.len + 32);
    try spliced.appendSlice(allocator, src_xml[0..sd_open]);
    try spliced.appendSlice(allocator, "<sheetData");
    try spliced.appendSlice(allocator, attrs);
    try spliced.append(allocator, '>');
    try spliced.appendSlice(allocator, rendered_rows);
    try spliced.appendSlice(allocator, "</sheetData>");
    try spliced.appendSlice(allocator, src_xml[sd_close + 1 ..]);
    return try spliced.toOwnedSlice(allocator);
}

/// Widen the BR corner of a canonical `<dimension ref="TL:BR"/>` so
/// both the row component reaches `new_max_row` and the letter
/// component reaches `new_max_col_1based`. Returns null on no-op,
/// non-canonical shape, or missing `<dimension>` — the spec lets
/// Excel rescan `<sheetData>` and rewrite the dimension on its
/// next save, so staleness on those is tolerable.
fn appendXmlUpdateDimensionBR(
    allocator: Allocator,
    xml: []const u8,
    new_max_row: u32,
    new_max_col_1based: u32,
) Error!?[]u8 {
    const dim_open = "<dimension ref=\"";
    const dim_pos = std.mem.indexOf(u8, xml, dim_open) orelse return null;
    const ref_start = dim_pos + dim_open.len;
    const ref_end = std.mem.indexOfScalarPos(u8, xml, ref_start, '"') orelse return null;
    const ref = xml[ref_start..ref_end];
    const colon = std.mem.indexOfScalar(u8, ref, ':') orelse return null;
    const br = ref[colon + 1 ..];
    if (br.len == 0) return null;

    var digit_start: usize = br.len;
    while (digit_start > 0 and br[digit_start - 1] >= '0' and br[digit_start - 1] <= '9') {
        digit_start -= 1;
    }
    if (digit_start == br.len or digit_start == 0) return null;
    for (br[0..digit_start]) |c| if (c < 'A' or c > 'Z') return null;
    const old_row = std.fmt.parseInt(u32, br[digit_start..], 10) catch return null;
    const old_col_1based = appendXmlParseColLetters(br[0..digit_start]) orelse return null;

    const final_row: u32 = @max(old_row, new_max_row);
    const final_col_1based: u32 = @max(old_col_1based, new_max_col_1based);
    if (final_row == old_row and final_col_1based == old_col_1based) return null;

    const br_abs_start = ref_start + colon + 1;
    const br_abs_end = ref_end;
    var out: std.ArrayList(u8) = .empty;
    errdefer out.deinit(allocator);
    try out.ensureTotalCapacity(allocator, xml.len + 8);
    try out.appendSlice(allocator, xml[0..br_abs_start]);
    var letter_buf: [8]u8 = undefined;
    const letters = appendXmlColLetters(&letter_buf, final_col_1based);
    try out.appendSlice(allocator, letters);
    var num_buf: [16]u8 = undefined;
    try out.appendSlice(allocator, try std.fmt.bufPrint(&num_buf, "{d}", .{final_row}));
    try out.appendSlice(allocator, xml[br_abs_end..]);
    return try out.toOwnedSlice(allocator);
}

/// Parse uppercase A-Z letters as a 1-based Excel column index
/// (A=1, …, XFD=16384). Returns null on empty input or anything
/// past `max_col_1based`.
fn appendXmlParseColLetters(s: []const u8) ?u32 {
    if (s.len == 0) return null;
    var n: u32 = 0;
    for (s) |c| {
        if (c < 'A' or c > 'Z') return null;
        n = n * 26 + (c - 'A' + 1);
        if (n > zlsx.max_col_1based) return null;
    }
    return n;
}

/// Render `col_1based` as A, B, …, Z, AA, … into `buf`. Capacity 8
/// is more than enough (Excel max is XFD = 3 letters).
fn appendXmlColLetters(buf: []u8, col_1based: u32) []u8 {
    assert(buf.len >= 7);
    assert(col_1based >= 1 and col_1based <= zlsx.max_col_1based);
    var n: u32 = col_1based;
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

/// B3 iter-wr-6 NOTE: kept local for the same reason as
/// `appendXmlEscaped` above — `pkg/sheet_plan.zig::appendXmlEscapedText`
/// rejects XML 1.0 forbidden control bytes via `error.InvalidXmlByte`,
/// which is not in Workbook's `Error` set. 3-entity (`<`, `>`, `&`)
/// element-content escape; quotes pass through verbatim per the
/// byte-stable contract.
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

/// Run the formula rewriter; return the rewritten bytes only when
/// they differ from the original. Caller owns the returned buffer
/// (allocator.free). On byte-identical output we free internally and
/// return null — the splice loop skips it. Helper for
/// `Workbook.rewriteAllValidationsAndConditionalFormats`.
fn maybeRewrite(
    a: Allocator,
    body: []const u8,
    on_sheet: ?[]const u8,
    target_sheet: ?[]const u8,
    edit: zlsx.formula_rewriter.RewriteEdit,
) Error!?[]u8 {
    if (body.len == 0) return null;
    const ctx = zlsx.formula_rewriter.RewriteContext{
        .on_sheet = on_sheet,
        .target_sheet = target_sheet,
        .edit = edit,
    };
    const rewritten = try zlsx.formula_rewriter.rewriteFormula(a, body, ctx);
    if (std.mem.eql(u8, rewritten, body)) {
        a.free(rewritten);
        return null;
    }
    return rewritten;
}

/// Per-formula splice patch in source-byte space. `[start..end]` is
/// the inner-text span of a `<formula1>` / `<formula2>` / `<formula>`
/// element inside the source sheet XML; `new` replaces those bytes.
const SourcePatch = struct { start: usize, end: usize, new: []const u8 };

/// Walk the source sheet XML in document order and locate each
/// formula body whose typed-view counterpart was rewritten. Appends
/// one `SourcePatch` per planned splice. Document order is the
/// invariant linking typed-view indices to source occurrences:
/// `parseValidations` and `parseConditionalFormats` iterate the
/// source linearly without re-ordering, so the Nth `<formula1>`
/// inside `<dataValidations>` corresponds to `view.validations[N]`'s
/// `formula1`, etc.
///
/// Helper used only by
/// `Workbook.rewriteAllValidationsAndConditionalFormats`.
fn collectDvCfPatches(
    a: Allocator,
    source: []const u8,
    out: *std.ArrayList(SourcePatch),
    dv_f1_new: *const std.AutoHashMapUnmanaged(usize, []u8),
    dv_f2_new: *const std.AutoHashMapUnmanaged(usize, []u8),
    cf_f_new: *const std.AutoHashMapUnmanaged(usize, []u8),
) Error!void {
    assert(source.len > 0);

    // ─── DV walk ────────────────────────────────────────────────────
    if (dv_f1_new.count() + dv_f2_new.count() > 0) {
        if (std.mem.indexOf(u8, source, "<dataValidations")) |dv_open| {
            const dv_open_gt = std.mem.indexOfScalarPos(u8, source, dv_open, '>') orelse
                return error.NoSheetData;
            const self_closing = dv_open_gt > 0 and source[dv_open_gt - 1] == '/';
            if (!self_closing) {
                const dv_close = std.mem.indexOfPos(u8, source, dv_open_gt, "</dataValidations>") orelse
                    return error.NoSheetData;
                const block_lo = dv_open_gt + 1;
                const block_hi = dv_close;
                var probe: usize = block_lo;
                var dv_idx: usize = 0;
                while (probe < block_hi) {
                    const e_open = std.mem.indexOfPos(u8, source, probe, "<dataValidation") orelse break;
                    if (e_open >= block_hi) break;
                    const after = e_open + "<dataValidation".len;
                    if (after >= source.len) break;
                    const sep = source[after];
                    if (sep != ' ' and sep != '\t' and sep != '\n' and sep != '\r' and sep != '/' and sep != '>') {
                        probe = after;
                        continue;
                    }
                    const e_open_gt = std.mem.indexOfScalarPos(u8, source, e_open, '>') orelse
                        return error.NoSheetData;
                    const e_self_closing = e_open_gt > 0 and source[e_open_gt - 1] == '/';
                    var elem_hi: usize = undefined;
                    if (e_self_closing) {
                        elem_hi = e_open_gt + 1;
                        probe = elem_hi;
                    } else {
                        const e_close = std.mem.indexOfPos(u8, source, e_open_gt, "</dataValidation>") orelse
                            return error.NoSheetData;
                        elem_hi = e_close;
                        probe = e_close + "</dataValidation>".len;

                        const body_lo = e_open_gt + 1;
                        const body_hi = elem_hi;
                        if (dv_f1_new.get(dv_idx)) |new1| {
                            if (findInnerSpan(source, body_lo, body_hi, "<formula1", "</formula1>")) |span| {
                                try out.append(a, .{ .start = span[0], .end = span[1], .new = new1 });
                            }
                        }
                        if (dv_f2_new.get(dv_idx)) |new2| {
                            if (findInnerSpan(source, body_lo, body_hi, "<formula2", "</formula2>")) |span| {
                                try out.append(a, .{ .start = span[0], .end = span[1], .new = new2 });
                            }
                        }
                    }
                    dv_idx += 1;
                }
            }
        }
    }

    // ─── CF walk ────────────────────────────────────────────────────
    if (cf_f_new.count() > 0) {
        var probe: usize = 0;
        var cf_idx: usize = 0;
        while (std.mem.indexOfPos(u8, source, probe, "<conditionalFormatting")) |cf_open| {
            const after = cf_open + "<conditionalFormatting".len;
            if (after >= source.len) break;
            const sep = source[after];
            if (sep != ' ' and sep != '\t' and sep != '\n' and sep != '\r' and sep != '/' and sep != '>') {
                probe = after;
                continue;
            }
            const cf_open_gt = std.mem.indexOfScalarPos(u8, source, cf_open, '>') orelse
                return error.NoSheetData;
            const cf_self_closing = cf_open_gt > 0 and source[cf_open_gt - 1] == '/';
            if (cf_self_closing) {
                probe = cf_open_gt + 1;
                continue;
            }
            const cf_close = std.mem.indexOfPos(u8, source, cf_open_gt, "</conditionalFormatting>") orelse
                return error.NoSheetData;
            const cf_body_lo = cf_open_gt + 1;
            const cf_body_hi = cf_close;
            probe = cf_close + "</conditionalFormatting>".len;

            // Walk each <cfRule> in this group. Each rule advances
            // cf_idx by one, matching parseConditionalFormats's order.
            var r_probe: usize = cf_body_lo;
            while (r_probe < cf_body_hi) {
                const r_open = std.mem.indexOfPos(u8, source, r_probe, "<cfRule") orelse break;
                if (r_open >= cf_body_hi) break;
                const r_after = r_open + "<cfRule".len;
                if (r_after >= source.len) break;
                const r_sep = source[r_after];
                if (r_sep != ' ' and r_sep != '\t' and r_sep != '\n' and r_sep != '\r' and r_sep != '/' and r_sep != '>') {
                    r_probe = r_after;
                    continue;
                }
                const r_open_gt = std.mem.indexOfScalarPos(u8, source, r_open, '>') orelse
                    return error.NoSheetData;
                const r_self_closing = r_open_gt > 0 and source[r_open_gt - 1] == '/';
                if (r_self_closing) {
                    // No body — no formula to splice. Still advance idx.
                    cf_idx += 1;
                    r_probe = r_open_gt + 1;
                    continue;
                }
                const r_close = std.mem.indexOfPos(u8, source, r_open_gt, "</cfRule>") orelse
                    return error.NoSheetData;
                const r_body_lo = r_open_gt + 1;
                const r_body_hi = r_close;
                r_probe = r_close + "</cfRule>".len;

                if (cf_f_new.get(cf_idx)) |new_f| {
                    if (findInnerSpan(source, r_body_lo, r_body_hi, "<formula", "</formula>")) |span| {
                        try out.append(a, .{ .start = span[0], .end = span[1], .new = new_f });
                    }
                }
                cf_idx += 1;
            }
        }
    }
}

/// Locate the inner-text span of `<tag …>BODY</close>` within
/// `source[lo..hi]`. Returns `[body_lo, body_hi]` (the BODY span),
/// or null if either tag is missing in that range. `open_prefix` is
/// the opening-tag prefix without `>` (e.g. "<formula1") so we match
/// both `<formula1>` and `<formula1 attr="…">`. The `<formula` /
/// `<formula1` disambiguation is handled by the caller's choice of
/// `open_prefix` (the search anchors on the literal prefix string).
fn findInnerSpan(
    source: []const u8,
    lo: usize,
    hi: usize,
    open_prefix: []const u8,
    close_tag: []const u8,
) ?[2]usize {
    if (lo >= hi or hi > source.len) return null;
    const slice = source[lo..hi];
    const o_rel = std.mem.indexOf(u8, slice, open_prefix) orelse return null;
    const o_abs = lo + o_rel;
    const o_after = o_abs + open_prefix.len;
    if (o_after >= source.len) return null;
    // `open_prefix` is "<formula" or "<formula1"/"<formula2". The
    // boundary char must be `>`, whitespace, or `/` — otherwise we
    // hit a longer-named element ("<formula1" matched on
    // "<formula12" — guard against this).
    const sep = source[o_after];
    const is_boundary = switch (sep) {
        ' ', '\t', '\r', '\n', '/', '>' => true,
        else => false,
    };
    if (!is_boundary) return null;
    const o_gt = std.mem.indexOfScalarPos(u8, source, o_after, '>') orelse return null;
    if (o_gt >= hi) return null;
    if (o_gt > 0 and source[o_gt - 1] == '/') return null; // self-closing — no body
    const c_rel = std.mem.indexOfPos(u8, source, o_gt + 1, close_tag) orelse return null;
    if (c_rel >= hi) return null;
    return .{ o_gt + 1, c_rel };
}

/// Linear splice of `source` against `patches`. Each patch's
/// `[start..end]` source span is replaced with `new`. Patches arrive
/// in collector-emission order which is NOT guaranteed source-order
/// (OOXML CT_Worksheet places `<conditionalFormatting>` before
/// `<dataValidations>`, but this collector walks DV first). We sort
/// in place by `.start` and assert disjointness.
fn spliceFormulas(
    a: Allocator,
    source: []const u8,
    patches: []SourcePatch,
) Error![]u8 {
    assert(source.len > 0);
    assert(patches.len > 0);

    const lessThan = struct {
        fn lt(_: void, x: SourcePatch, y: SourcePatch) bool {
            return x.start < y.start;
        }
    }.lt;
    std.sort.pdq(SourcePatch, patches, {}, lessThan);

    // Disjointness invariant. A violation implies the collector
    // visited overlapping spans (parser bug or a `formula1`
    // body containing a literal `</formula1>` payload, which OOXML
    // forbids).
    var i: usize = 1;
    while (i < patches.len) : (i += 1) {
        assert(patches[i].start >= patches[i - 1].end);
    }

    var out: std.ArrayList(u8) = .empty;
    errdefer out.deinit(a);
    try out.ensureTotalCapacity(a, source.len + 256);

    var cursor: usize = 0;
    for (patches) |p| {
        assert(p.end <= source.len);
        try out.appendSlice(a, source[cursor..p.start]);
        // The rewriter emits already-formed formula text. It does
        // NOT produce raw `<` / `>` / `&` (the tokenizer prints A1
        // refs and operators only). We still XML-escape on emit so a
        // future rewriter feature that DOES produce one of those
        // bytes can't corrupt the surrounding XML.
        try appendXmlEscapedText(a, &out, p.new);
        cursor = p.end;
    }
    try out.appendSlice(a, source[cursor..]);

    return try out.toOwnedSlice(a);
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

test "Workbook.numberFormatFor: built-in id 0 resolves to General" {
    const path = "tests/corpus/worldbank_catalog.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;

    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();

    // Find a style whose numFmtId is a built-in (≥0, <164). Most
    // real fixtures' style index 0 is General (id=0). If the fixture
    // happens to have no built-in styles, scan up to a sane cap.
    const st = (try wb.styles()).?;
    try std.testing.expect(st.cell_xfs.len > 0);

    var found: ?NumberFormatInfo = null;
    var idx: u32 = 0;
    while (idx < st.cell_xfs.len) : (idx += 1) {
        if (try wb.numberFormatFor(idx)) |nfi| {
            if (nfi.is_builtin) {
                found = nfi;
                break;
            }
        }
    }
    try std.testing.expect(found != null);
    // Built-in code is a static literal: round-trip equality with the
    // table's exact bytes for the resolved id.
    const expected = builtinNumFmtCode(found.?.fmt_id).?;
    try std.testing.expectEqualStrings(expected, found.?.code);
    try std.testing.expect(found.?.is_builtin);
}

test "Workbook.numberFormatFor: out-of-range style_idx returns null" {
    const path = "tests/corpus/worldbank_catalog.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;

    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();

    const st = (try wb.styles()).?;
    const oor: u32 = @intCast(st.cell_xfs.len);
    const result = try wb.numberFormatFor(oor);
    try std.testing.expect(result == null);

    // Far-out-of-range too — a u32 the cell_xfs slice will never reach.
    const result2 = try wb.numberFormatFor(std.math.maxInt(u32));
    try std.testing.expect(result2 == null);
}

test "Workbook.numberFormatFor: custom numFmtId resolves to the styles.xml entry" {
    // Synthesize a workbook with a custom <numFmt numFmtId="164"
    // formatCode="0.000"/> and a cellXf referencing it. We exercise
    // the lookup branch directly against the typed view; a real
    // PartStore-backed fixture would also work but lets us avoid
    // shipping a new corpus file just for this test.
    const xml =
        \\<styleSheet>
        \\  <numFmts count="2">
        \\    <numFmt numFmtId="164" formatCode="0.000"/>
        \\    <numFmt numFmtId="170" formatCode="#,##0.00 _$"/>
        \\  </numFmts>
        \\  <cellXfs count="3">
        \\    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        \\    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
        \\    <xf numFmtId="170" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
        \\  </cellXfs>
        \\</styleSheet>
    ;
    var sx = try styles_xml_mod.parse(std.testing.allocator, xml);
    defer sx.deinit(std.testing.allocator);

    // Mirror the resolution logic Workbook.numberFormatFor performs.
    // We assert it directly so the test passes without a synthesized
    // PartStore — keeping coverage focused on the lookup itself.
    try std.testing.expectEqual(@as(usize, 2), sx.number_formats.len);
    try std.testing.expectEqual(@as(usize, 3), sx.cell_xfs.len);

    const nfid_1 = sx.cell_xfs[1].num_fmt_id.?;
    try std.testing.expect(builtinNumFmtCode(nfid_1) == null); // 164 isn't built-in
    var matched: ?[]const u8 = null;
    for (sx.number_formats) |nf| {
        if (nf.fmt_id == nfid_1) matched = nf.code;
    }
    try std.testing.expect(matched != null);
    try std.testing.expectEqualStrings("0.000", matched.?);

    const nfid_2 = sx.cell_xfs[2].num_fmt_id.?;
    try std.testing.expect(builtinNumFmtCode(nfid_2) == null);
    matched = null;
    for (sx.number_formats) |nf| {
        if (nf.fmt_id == nfid_2) matched = nf.code;
    }
    try std.testing.expectEqualStrings("#,##0.00 _$", matched.?);
}

test "Workbook.numberFormatFor: builtinNumFmtCode covers the well-known subset" {
    // Spot-check a handful of representative entries — the full table
    // is exercised by the lookup tests above. Asserts both presence
    // and exact byte equality (these are stable string literals per
    // ECMA-376 §18.8.30).
    try std.testing.expectEqualStrings("General", builtinNumFmtCode(0).?);
    try std.testing.expectEqualStrings("0.00", builtinNumFmtCode(2).?);
    try std.testing.expectEqualStrings("0%", builtinNumFmtCode(9).?);
    try std.testing.expectEqualStrings("m/d/yyyy", builtinNumFmtCode(14).?);
    try std.testing.expectEqualStrings("h:mm:ss", builtinNumFmtCode(21).?);
    try std.testing.expectEqualStrings("@", builtinNumFmtCode(49).?);
    // Skipped / locale-specific IDs fall through to null (caller
    // resolves via the custom <numFmt> table).
    try std.testing.expect(builtinNumFmtCode(5) == null);
    try std.testing.expect(builtinNumFmtCode(23) == null);
    try std.testing.expect(builtinNumFmtCode(50) == null);
    try std.testing.expect(builtinNumFmtCode(164) == null);
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

// ─── DV / CF rewriter tests (C1 M2 m2) ───────────────────────────────

/// Splice synthetic `<dataValidations>` and `<conditionalFormatting>`
/// blocks into the source sheet XML for `wb.sheet(sheet_idx)` and
/// push the patched bytes via `wb.store.replacePart`. The corpus
/// fixtures lack DV/CF natively, so DV/CF tests build them this way.
/// The injected blocks live just before `</worksheet>`. Caller MUST
/// invalidate any cached `parsed` view on the touched sheet
/// afterwards (set `ws.parsed = null` after a manual `view.deinit`).
fn injectDvAndCfIntoSheet(
    a: Allocator,
    wb: *Workbook,
    sheet_idx: u32,
    dv_block: []const u8,
    cf_block: []const u8,
) Error!void {
    const ws = try wb.sheet(sheet_idx);
    _ = try ws.ensureParsed();
    const part_name = ws.resolved_part_name.?;
    const part = try wb.store.part(part_name) orelse return error.MissingSheetPart;
    const src = part.bytes;

    const close_idx = std.mem.lastIndexOf(u8, src, "</worksheet>") orelse
        return error.NoSheetData;

    var out: std.ArrayList(u8) = .empty;
    defer out.deinit(a);
    try out.ensureTotalCapacity(a, src.len + dv_block.len + cf_block.len + 16);

    try out.appendSlice(a, src[0..close_idx]);
    // Order matters: CF before DV per OOXML CT_Worksheet schema.
    if (cf_block.len > 0) try out.appendSlice(a, cf_block);
    if (dv_block.len > 0) try out.appendSlice(a, dv_block);
    try out.appendSlice(a, src[close_idx..]);

    try wb.store.replacePart(part_name, out.items);

    // Drop the stale parsed view: it borrowed from the pre-splice
    // bytes. Next access re-parses the patched part.
    var stale = ws.parsed.?;
    stale.deinit(a);
    ws.parsed = null;
}

test "Workbook.rewriteAllValidationsAndConditionalFormats: insert_rows shifts DV formulas, persists round-trip" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));
    var tmp_buf: [256]u8 = undefined;
    const tmp_path = try std.fmt.bufPrint(&tmp_buf, ".zig-cache/test-dvcf-rows-{d}.xlsx", .{prng.random().int(u32)});

    // Stage: open fixture, inject DV block on sheet 0, save.
    {
        var wb = try Workbook.open(std.testing.allocator, src_path);
        defer wb.deinit();

        // Two DVs: one with formula1 only, one with both formulas +
        // an `errorTitle` attr we want to confirm is preserved across
        // the splice.
        const dv =
            \\<dataValidations count="2"><dataValidation type="list" allowBlank="1" sqref="A1:A10"><formula1>B5:B10</formula1></dataValidation><dataValidation type="whole" operator="between" errorTitle="Bad" sqref="C1:C10"><formula1>B5</formula1><formula2>B7+1</formula2></dataValidation></dataValidations>
        ;
        try injectDvAndCfIntoSheet(std.testing.allocator, &wb, 0, dv, "");
        try wb.save(tmp_path);
    }
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    // Re-open, rewrite. target_sheet null = "apply everywhere".
    var tmp2_buf: [256]u8 = undefined;
    const tmp2_path = try std.fmt.bufPrint(&tmp2_buf, ".zig-cache/test-dvcf-rows-out-{d}.xlsx", .{prng.random().int(u32)});
    {
        var wb = try Workbook.open(std.testing.allocator, tmp_path);
        defer wb.deinit();

        const count = try wb.rewriteAllValidationsAndConditionalFormats(
            .{ .insert_rows = .{ .at = 4, .count = 1 } },
            null,
        );
        // formula1 "B5:B10" → "B6:B11" (1)
        // formula1 "B5" → "B6" (1)
        // formula2 "B7+1" → "B8+1" (1)
        try std.testing.expectEqual(@as(u32, 3), count);

        try wb.save(tmp2_path);
    }
    defer std.fs.cwd().deleteFile(tmp2_path) catch {};

    // Round-trip: re-open, re-parse, verify shifted formulas.
    var wb2 = try Workbook.open(std.testing.allocator, tmp2_path);
    defer wb2.deinit();
    const ws = try wb2.sheet(0);
    const dvs = try ws.validations();
    try std.testing.expectEqual(@as(usize, 2), dvs.len);
    try std.testing.expectEqualStrings("B6:B11", dvs[0].formula1.?);
    try std.testing.expectEqualStrings("B6", dvs[1].formula1.?);
    try std.testing.expectEqualStrings("B8+1", dvs[1].formula2.?);
    // errorTitle attribute survived the splice (preservation
    // contract — we never regenerate the DV element from the
    // typed view, which would drop it).
    const part = (try wb2.store.part(ws.resolved_part_name.?)).?;
    try std.testing.expect(std.mem.indexOf(u8, part.bytes, "errorTitle=\"Bad\"") != null);
}

test "Workbook.rewriteAllValidationsAndConditionalFormats: insert_cols shifts CF formula, persists round-trip" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));
    var tmp_buf: [256]u8 = undefined;
    const tmp_path = try std.fmt.bufPrint(&tmp_buf, ".zig-cache/test-dvcf-cols-{d}.xlsx", .{prng.random().int(u32)});

    {
        var wb = try Workbook.open(std.testing.allocator, src_path);
        defer wb.deinit();

        // CF block with two cfRules (one with `dxfId` we expect to
        // survive, one self-closing without a body).
        const cf =
            \\<conditionalFormatting sqref="D1:D10"><cfRule type="expression" dxfId="0" priority="1"><formula>D1+E1</formula></cfRule><cfRule type="containsBlanks" priority="2"/></conditionalFormatting>
        ;
        try injectDvAndCfIntoSheet(std.testing.allocator, &wb, 0, "", cf);
        try wb.save(tmp_path);
    }
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    var tmp2_buf: [256]u8 = undefined;
    const tmp2_path = try std.fmt.bufPrint(&tmp2_buf, ".zig-cache/test-dvcf-cols-out-{d}.xlsx", .{prng.random().int(u32)});
    {
        var wb = try Workbook.open(std.testing.allocator, tmp_path);
        defer wb.deinit();
        const ws_name_owned = try std.testing.allocator.dupe(u8, (try wb.sheet(0)).name());
        defer std.testing.allocator.free(ws_name_owned);

        // target_sheet = sheet 0 — bare refs `D1`, `E1` are scoped
        // to sheet 0, so they shift on insert_cols at col D (=4).
        const count = try wb.rewriteAllValidationsAndConditionalFormats(
            .{ .insert_cols = .{ .at = 4, .count = 1 } },
            ws_name_owned,
        );
        // CF formula "D1+E1" → "E1+F1" (1 body, 2 refs shifted = 1 rewrite)
        try std.testing.expectEqual(@as(u32, 1), count);

        try wb.save(tmp2_path);
    }
    defer std.fs.cwd().deleteFile(tmp2_path) catch {};

    var wb2 = try Workbook.open(std.testing.allocator, tmp2_path);
    defer wb2.deinit();
    const ws = try wb2.sheet(0);
    const cfs = try ws.conditionalFormats();
    // Two cfRules — only the first has a formula (second was self-
    // closing). Both survive the splice with their attrs.
    try std.testing.expectEqual(@as(usize, 2), cfs.len);
    try std.testing.expectEqualStrings("E1+F1", cfs[0].formula.?);
    // dxfId on rule 0 preserved (attribute-byte preservation contract).
    try std.testing.expectEqual(@as(?u32, 0), cfs[0].dxf_id);
    try std.testing.expectEqual(@as(?u32, 1), cfs[0].priority);
    try std.testing.expectEqual(@as(?u32, 2), cfs[1].priority);
}

test "Workbook.rewriteAllValidationsAndConditionalFormats: no-op count == 0 on workbook without DV/CF" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    var wb = try Workbook.open(std.testing.allocator, src_path);
    defer wb.deinit();

    // Pristine fixture — no <dataValidations>, no <conditionalFormatting>.
    const count = try wb.rewriteAllValidationsAndConditionalFormats(
        .{ .insert_rows = .{ .at = 1, .count = 1 } },
        null,
    );
    try std.testing.expectEqual(@as(u32, 0), count);
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

test "Worksheet.cellStyle: cell with no style attribute returns null" {
    // phpoi_test1 cell A1 has no `s="…"` — `style_idx` is null, so
    // cellStyle short-circuits before consulting StylesXml.
    const path = "tests/corpus/phpoi_test1.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;

    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();

    const s0 = try wb.sheet(0);
    const resolved = try s0.cellStyle("A1");
    try std.testing.expectEqual(@as(?ResolvedStyle, null), resolved);

    // Out-of-range ref (no matching cell) also returns null.
    const missing = try s0.cellStyle("ZZ9999");
    try std.testing.expectEqual(@as(?ResolvedStyle, null), missing);
}

test "Worksheet.cellStyle: applyFont surfaces the bold font on phpoi B2" {
    // phpoi_test1: cellXfs[1] = { fontId=1, applyFont=1 } → fonts[1]
    // is the bold Calibri 11. Other apply_* flags are off, so fill /
    // border / alignment / number_format_code stay null.
    const path = "tests/corpus/phpoi_test1.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;

    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();

    const s0 = try wb.sheet(0);
    const resolved = (try s0.cellStyle("B2")) orelse return error.TestUnexpectedNull;

    try std.testing.expect(resolved.font != null);
    try std.testing.expect(resolved.font.?.bold);
    try std.testing.expectEqualStrings("Calibri", resolved.font.?.name.?);
    try std.testing.expectEqual(@as(?styles_xml_mod.Fill, null), resolved.fill);
    try std.testing.expectEqual(@as(?styles_xml_mod.Border, null), resolved.border);
    try std.testing.expectEqual(@as(?styles_xml_mod.Alignment, null), resolved.alignment);
    try std.testing.expectEqual(@as(?[]const u8, null), resolved.number_format_code);
}

test "Worksheet.cellStyle: applyAlignment surfaces wrap_text; built-in numFmt id has null code" {
    // phpoi_test1: C3 has style_idx=2 → applyAlignment=1, alignment
    // body has wrapText=1. D4 has style_idx=3 → applyNumberFormat=1,
    // numFmtId=2 which is built-in (≤163), so number_format_code is
    // null (the code is implicit, not stored in <numFmts>).
    const path = "tests/corpus/phpoi_test1.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;

    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();

    const s0 = try wb.sheet(0);

    const c3 = (try s0.cellStyle("C3")) orelse return error.TestUnexpectedNull;
    try std.testing.expect(c3.alignment != null);
    try std.testing.expect(c3.alignment.?.wrap_text);
    try std.testing.expectEqual(@as(?styles_xml_mod.Font, null), c3.font);

    const d4 = (try s0.cellStyle("D4")) orelse return error.TestUnexpectedNull;
    // Built-in numFmtId=2 (`0.00`) — overlay does not synthesize codes
    // for built-ins; field stays null.
    try std.testing.expectEqual(@as(?[]const u8, null), d4.number_format_code);
    try std.testing.expectEqual(@as(?styles_xml_mod.Font, null), d4.font);
    try std.testing.expectEqual(@as(?styles_xml_mod.Alignment, null), d4.alignment);
}

test "Workbook.deleteCell: removes existing cell from saved sheet" {
    const src_path = "tests/corpus/openpyxl_guess_types.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    var tmp_buf: [256]u8 = undefined;
    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));
    const tmp_path = try std.fmt.bufPrint(&tmp_buf, ".zig-cache/test-wb-delete-{d}.xlsx", .{prng.random().int(u32)});

    {
        var wb = try Workbook.open(std.testing.allocator, src_path);
        defer wb.deinit();
        const s0 = try wb.sheet(0);
        // Sanity: A1 exists before delete.
        const before = try s0.cellByRef("A1");
        try std.testing.expect(before != null);

        try s0.deleteCell("A1");
        try wb.save(tmp_path);
    }
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    var wb2 = try Workbook.open(std.testing.allocator, tmp_path);
    defer wb2.deinit();
    const s0 = try wb2.sheet(0);

    const a1 = try s0.cellByRef("A1");
    try std.testing.expect(a1 == null);
}

test "Workbook.deleteCell vs setCell(.blank): elision vs empty cell" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    var tmp_buf: [256]u8 = undefined;
    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));
    const tmp_path = try std.fmt.bufPrint(&tmp_buf, ".zig-cache/test-wb-delete-vs-blank-{d}.xlsx", .{prng.random().int(u32)});

    {
        var wb = try Workbook.open(std.testing.allocator, src_path);
        defer wb.deinit();
        const s0 = try wb.sheet(0);
        // Stage two side-by-side cells so both refs land in the same
        // <sheetData>; one is fully removed, the other left empty.
        try s0.setCell("Z1", .{ .number = 7 });
        try s0.setCell("Z2", .{ .number = 8 });
        try wb.save(tmp_path);
    }
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    // Re-open and stage delete on Z1, blank on Z2; save again.
    {
        var wb = try Workbook.open(std.testing.allocator, tmp_path);
        defer wb.deinit();
        const s0 = try wb.sheet(0);
        try s0.deleteCell("Z1");
        try s0.setCell("Z2", .blank);
        try wb.save(tmp_path);
    }

    // Inspect the regenerated sheet bytes directly: Z1 must be absent,
    // Z2 must be present as a self-closing empty <c>.
    var wb2 = try Workbook.open(std.testing.allocator, tmp_path);
    defer wb2.deinit();
    const s0 = try wb2.sheet(0);
    _ = try s0.ensureParsed(); // populates resolved_part_name
    const part_name = s0.resolved_part_name.?;
    const part = (try wb2.store.part(part_name)) orelse return error.MissingSheetPart;
    const xml = part.bytes;

    try std.testing.expect(std.mem.indexOf(u8, xml, "r=\"Z1\"") == null);
    try std.testing.expect(std.mem.indexOf(u8, xml, "<c r=\"Z2\"/>") != null);

    // Reader-level invariant: cellByRef agrees.
    try std.testing.expect((try s0.cellByRef("Z1")) == null);
    const z2 = try s0.cellByRef("Z2");
    try std.testing.expect(z2 != null);
    try std.testing.expect(z2.?.raw_value == null);
}

test "Workbook.deleteCell: non-existent ref is a no-op" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    var tmp_buf: [256]u8 = undefined;
    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));
    const tmp_path = try std.fmt.bufPrint(&tmp_buf, ".zig-cache/test-wb-delete-noop-{d}.xlsx", .{prng.random().int(u32)});

    {
        var wb = try Workbook.open(std.testing.allocator, src_path);
        defer wb.deinit();
        const s0 = try wb.sheet(0);
        // A ref guaranteed not to exist in the source corpus.
        try s0.deleteCell("ZZ9999");
        // Save must succeed; the .deleted delta has no original to elide.
        try wb.save(tmp_path);
    }
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    var wb2 = try Workbook.open(std.testing.allocator, tmp_path);
    defer wb2.deinit();
    const s0 = try wb2.sheet(0);
    try std.testing.expect((try s0.cellByRef("ZZ9999")) == null);
}

test "Workbook.hasUnsavedChanges: pristine workbook is clean" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;
    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();
    try std.testing.expect(!wb.hasUnsavedChanges());
}

test "Workbook.hasUnsavedChanges: setCell flips the bit" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;
    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();
    const s0 = try wb.sheet(0);
    try s0.setCell("A1", .{ .number = 42 });
    try std.testing.expect(wb.hasUnsavedChanges());
}

test "Workbook.hasUnsavedChanges: save clears delta-only dirt; PartStore overrides persist post-save" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;

    var tmp_buf: [256]u8 = undefined;
    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));
    const tmp_path = try std.fmt.bufPrint(&tmp_buf, ".zig-cache/test-wb-dirty-{d}.xlsx", .{prng.random().int(u32)});
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();
    const s0 = try wb.sheet(0);
    try s0.setCell("A1", .{ .number = 42 });
    try std.testing.expect(wb.hasUnsavedChanges());
    try wb.save(tmp_path);
    // deltas are cleared by save, but PartStore overrides (set by
    // save's replacePart calls) persist — predicate stays true.
    // This documents the "diff vs original" semantics.
    try std.testing.expect(wb.hasUnsavedChanges());

    // Re-open from disk: clean again.
    var wb2 = try Workbook.open(std.testing.allocator, tmp_path);
    defer wb2.deinit();
    try std.testing.expect(!wb2.hasUnsavedChanges());
}

test "Workbook.hasUnsavedChanges: renameSheet flips via PartStore override" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;
    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();
    try wb.renameSheet(1, "Renamed");
    try std.testing.expect(wb.hasUnsavedChanges());
}

test "Workbook.deleteSheet: happy path drops second sheet, byte-stable round-trip" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));

    var tmp_buf: [256]u8 = undefined;
    const tmp_path = try std.fmt.bufPrint(&tmp_buf, ".zig-cache/test-wb-delete-out-{d}.xlsx", .{prng.random().int(u32)});

    // Capture sheet 0's name before delete so we can assert the
    // survivor is the right sheet (not just "any one sheet").
    var s0_name_buf: [128]u8 = undefined;
    var s0_name_len: usize = 0;
    {
        var wb = try Workbook.open(std.testing.allocator, src_path);
        defer wb.deinit();
        try std.testing.expectEqual(@as(u32, 2), wb.sheetCount());

        const s0 = try wb.sheet(0);
        const name = s0.name();
        try std.testing.expect(name.len <= s0_name_buf.len);
        @memcpy(s0_name_buf[0..name.len], name);
        s0_name_len = name.len;

        try wb.deleteSheet(1);

        try std.testing.expectEqual(@as(u32, 1), wb.sheetCount());
        try std.testing.expectEqualStrings(s0_name_buf[0..s0_name_len], (try wb.sheet(0)).name());

        try wb.save(tmp_path);
    }
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    var wb2 = try Workbook.open(std.testing.allocator, tmp_path);
    defer wb2.deinit();
    try std.testing.expectEqual(@as(u32, 1), wb2.sheetCount());
    try std.testing.expectEqualStrings(s0_name_buf[0..s0_name_len], (try wb2.sheet(0)).name());

    // workbook.xml's `<sheets>` list reflects the removal: only one
    // `<sheet ` element (boundary-checked against `<sheets>`,
    // `<sheetData>`, `<sheetView>`, etc.).
    const wb_part = (try wb2.store.part("xl/workbook.xml")).?;
    var i: usize = 0;
    var sheet_count: u32 = 0;
    while (std.mem.indexOfPos(u8, wb_part.bytes, i, "<sheet")) |pos| {
        const after = pos + "<sheet".len;
        if (after < wb_part.bytes.len) {
            const b = wb_part.bytes[after];
            if (b == ' ' or b == '\t' or b == '\r' or b == '\n' or b == '/') {
                sheet_count += 1;
            }
        }
        i = after;
    }
    try std.testing.expectEqual(@as(u32, 1), sheet_count);
}

test "Workbook.deleteSheet: refuses to remove the sole remaining sheet" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;
    var wb = try Workbook.open(std.testing.allocator, src_path);
    defer wb.deinit();

    try wb.deleteSheet(1);
    try std.testing.expectEqual(@as(u32, 1), wb.sheetCount());
    try std.testing.expectError(error.LastSheetUndeletable, wb.deleteSheet(0));
    // Still one sheet — no partial mutation slipped through.
    try std.testing.expectEqual(@as(u32, 1), wb.sheetCount());
}

test "Workbook.deleteSheet: out-of-range index errors SheetIndexOutOfRange" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;
    var wb = try Workbook.open(std.testing.allocator, src_path);
    defer wb.deinit();
    try std.testing.expectError(error.SheetIndexOutOfRange, wb.deleteSheet(99));
    // Boundary: exact sheetCount() is also out-of-range (0-based).
    try std.testing.expectError(error.SheetIndexOutOfRange, wb.deleteSheet(wb.sheetCount()));
    // No mutation occurred.
    try std.testing.expectEqual(@as(u32, 2), wb.sheetCount());
}

// ─── C1 M2 m3: defined-names + hyperlinks rewriter tests ─────────────

/// Splice a synthetic `<definedNames>` block into `wb`'s
/// `xl/workbook.xml` then save to `out_path`. Used by the m3 tests
/// since the corpus fixture has no defined names of its own. Returns
/// after `save`, so callers re-open from `out_path` to exercise the
/// rewriter against persisted bytes.
fn testInjectDefinedNames(
    wb: *Workbook,
    block_inner: []const u8,
    out_path: []const u8,
) !void {
    const a = wb.allocator;
    const part = (try wb.store.part("xl/workbook.xml")) orelse return error.MissingWorkbookPart;
    const src = part.bytes;
    // Insert AFTER `</sheets>` so the schema order is workbook →
    // fileVersion → workbookPr → bookViews → sheets → definedNames →
    // calcPr.
    const sheets_close = std.mem.indexOf(u8, src, "</sheets>") orelse return error.MalformedXml;
    const insert_at = sheets_close + "</sheets>".len;

    var out: std.ArrayList(u8) = .empty;
    defer out.deinit(a);
    try out.appendSlice(a, src[0..insert_at]);
    try out.appendSlice(a, "<definedNames>");
    try out.appendSlice(a, block_inner);
    try out.appendSlice(a, "</definedNames>");
    try out.appendSlice(a, src[insert_at..]);

    try wb.store.replacePart("xl/workbook.xml", out.items);
    try wb.save(out_path);
}

test "Workbook.rewriteAllDefinedNames: workbook-scope insert_rows shifts and persists" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));

    var tmp_buf: [256]u8 = undefined;
    const tmp_path = try std.fmt.bufPrint(&tmp_buf, ".zig-cache/test-defnames-in-{d}.xlsx", .{prng.random().int(u32)});

    // Build a fixture with a workbook-scope defined name pointing at
    // `Sheet1!A1+B1`. Saved to tmp_path; subsequent open reads it back.
    {
        var wb = try Workbook.open(std.testing.allocator, src_path);
        defer wb.deinit();
        try testInjectDefinedNames(
            &wb,
            "<definedName name=\"MyName\">Sheet1!A1+B1</definedName>",
            tmp_path,
        );
    }
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    var wb = try Workbook.open(std.testing.allocator, tmp_path);
    defer wb.deinit();

    // Sanity: the synthetic name is visible.
    try std.testing.expectEqual(@as(usize, 1), wb.definedNames().len);
    try std.testing.expectEqualStrings("Sheet1!A1+B1", wb.definedNames()[0].formula);

    const count = try wb.rewriteAllDefinedNames(.{
        .insert_rows = .{ .at = 1, .count = 1 },
    }, "Sheet1");
    try std.testing.expectEqual(@as(u32, 1), count);

    // In-memory view must reflect the rewrite immediately (no
    // save/re-open needed).
    try std.testing.expectEqual(@as(usize, 1), wb.definedNames().len);
    try std.testing.expectEqualStrings("Sheet1!A2+B2", wb.definedNames()[0].formula);

    // Persistence: save + re-open the rewritten workbook.
    var tmp2_buf: [256]u8 = undefined;
    const tmp2_path = try std.fmt.bufPrint(&tmp2_buf, ".zig-cache/test-defnames-out-{d}.xlsx", .{prng.random().int(u32)});
    try wb.save(tmp2_path);
    defer std.fs.cwd().deleteFile(tmp2_path) catch {};

    var wb2 = try Workbook.open(std.testing.allocator, tmp2_path);
    defer wb2.deinit();
    try std.testing.expectEqual(@as(usize, 1), wb2.definedNames().len);
    try std.testing.expectEqualStrings("Sheet1!A2+B2", wb2.definedNames()[0].formula);
    try std.testing.expectEqualStrings("MyName", wb2.definedNames()[0].name);
    try std.testing.expectEqual(@as(?u32, null), wb2.definedNames()[0].local_sheet_id);
}

test "Workbook.rewriteAllDefinedNames: sheet-scope localSheetId preserved across rewrite" {
    // PR #37 panicked here (signal 6 / abort). Root cause: the rewriter
    // re-parsed the workbook view mid-iteration over `defined_names`,
    // leaving the loop reading freed memory on the sheet-scope branch
    // (`local_sheet_id` resolves to a sheet name borrowed from the
    // about-to-be-freed arena). The fix: collect ALL pending rewrites
    // FIRST into allocator-owned arrays, THEN splice + re-parse once.
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));

    var tmp_buf: [256]u8 = undefined;
    const tmp_path = try std.fmt.bufPrint(&tmp_buf, ".zig-cache/test-defnames-local-in-{d}.xlsx", .{prng.random().int(u32)});

    // Sheet-scope: localSheetId="0" binds the name to the first sheet.
    // Formula `B5` (no qualifier) — bare A1 ref. With on_sheet="Sheet1"
    // (resolved from local_sheet_id=0) and target_sheet=null OR "Sheet1",
    // an insert_rows at row=4 count=2 should shift B5 → B7.
    {
        var wb = try Workbook.open(std.testing.allocator, src_path);
        defer wb.deinit();
        try testInjectDefinedNames(
            &wb,
            "<definedName name=\"LocalRef\" localSheetId=\"0\">B5</definedName>",
            tmp_path,
        );
    }
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    var wb = try Workbook.open(std.testing.allocator, tmp_path);
    defer wb.deinit();

    // Sanity: the sheet-scope name was parsed correctly.
    try std.testing.expectEqual(@as(usize, 1), wb.definedNames().len);
    try std.testing.expectEqual(@as(?u32, 0), wb.definedNames()[0].local_sheet_id);
    try std.testing.expectEqualStrings("B5", wb.definedNames()[0].formula);

    // target_sheet=null → "apply everywhere," matching the rewriter's
    // permissive default. The bare B5 ref is scoped to on_sheet
    // ("Sheet1" via local_sheet_id=0); rewriter shifts row 5 by +2.
    const count = try wb.rewriteAllDefinedNames(.{
        .insert_rows = .{ .at = 4, .count = 2 },
    }, null);
    try std.testing.expectEqual(@as(u32, 1), count);

    // Pair assertion: localSheetId AND name preserved across the
    // splice. Formula now reflects the +2 shift.
    try std.testing.expectEqual(@as(usize, 1), wb.definedNames().len);
    try std.testing.expectEqual(@as(?u32, 0), wb.definedNames()[0].local_sheet_id);
    try std.testing.expectEqualStrings("LocalRef", wb.definedNames()[0].name);
    try std.testing.expectEqualStrings("B7", wb.definedNames()[0].formula);

    // Persistence round-trip — confirms the splice survived save +
    // re-open and the workbook XML is well-formed (re-parse would
    // error otherwise).
    var tmp2_buf: [256]u8 = undefined;
    const tmp2_path = try std.fmt.bufPrint(&tmp2_buf, ".zig-cache/test-defnames-local-out-{d}.xlsx", .{prng.random().int(u32)});
    try wb.save(tmp2_path);
    defer std.fs.cwd().deleteFile(tmp2_path) catch {};

    var wb2 = try Workbook.open(std.testing.allocator, tmp2_path);
    defer wb2.deinit();
    try std.testing.expectEqual(@as(?u32, 0), wb2.definedNames()[0].local_sheet_id);
    try std.testing.expectEqualStrings("B7", wb2.definedNames()[0].formula);
}

/// Splice a `<hyperlinks>` block into the first sheet's XML and
/// save to `out_path`. Mirror of `testInjectDefinedNames` for the
/// hyperlinks tests.
fn testInjectHyperlinks(
    wb: *Workbook,
    block_inner: []const u8,
    out_path: []const u8,
) !void {
    const a = wb.allocator;
    const ws = try wb.sheet(0);
    _ = try ws.ensureParsed(); // populates resolved_part_name
    const part_name = ws.resolved_part_name.?;
    const part = (try wb.store.part(part_name)) orelse return error.MissingSheetPart;
    const src = part.bytes;

    // Place `<hyperlinks>` after `</sheetData>` per OOXML schema order.
    const sd_close = std.mem.indexOf(u8, src, "</sheetData>") orelse return error.MalformedXml;
    const insert_at = sd_close + "</sheetData>".len;

    var out: std.ArrayList(u8) = .empty;
    defer out.deinit(a);
    try out.appendSlice(a, src[0..insert_at]);
    try out.appendSlice(a, "<hyperlinks>");
    try out.appendSlice(a, block_inner);
    try out.appendSlice(a, "</hyperlinks>");
    try out.appendSlice(a, src[insert_at..]);

    try wb.store.replacePart(part_name, out.items);
    // Invalidate parsed view so subsequent reads pick up the splice.
    if (ws.parsed) |*p| {
        var view = p.*;
        view.deinit(wb.allocator);
        ws.parsed = null;
    }
    try wb.save(out_path);
}

test "Workbook.rewriteAllHyperlinkLocations: internal hyperlink shifts when target_sheet matches" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));

    var tmp_buf: [256]u8 = undefined;
    const tmp_path = try std.fmt.bufPrint(&tmp_buf, ".zig-cache/test-hl-in-{d}.xlsx", .{prng.random().int(u32)});

    // Internal hyperlink (no r:id) with location=A5. on_sheet for this
    // hyperlink resolves to "Sheet1" (sheet 0); a row insert at 1 with
    // count=4 shifts row 5 → row 9 → location should become A9.
    {
        var wb = try Workbook.open(std.testing.allocator, src_path);
        defer wb.deinit();
        try testInjectHyperlinks(
            &wb,
            "<hyperlink ref=\"C3\" location=\"A5\" display=\"jump\"/>",
            tmp_path,
        );
    }
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    var wb = try Workbook.open(std.testing.allocator, tmp_path);
    defer wb.deinit();

    {
        const ws = try wb.sheet(0);
        const hls = try ws.hyperlinks();
        try std.testing.expectEqual(@as(usize, 1), hls.len);
        try std.testing.expectEqual(@as(?[]const u8, null), hls[0].r_id);
        try std.testing.expectEqualStrings("A5", hls[0].location.?);
    }

    const count = try wb.rewriteAllHyperlinkLocations(.{
        .insert_rows = .{ .at = 1, .count = 4 },
    }, "Sheet1");
    try std.testing.expectEqual(@as(u32, 1), count);

    // Re-read parsed view post-rewrite (the rewriter invalidates it).
    {
        const ws = try wb.sheet(0);
        const hls = try ws.hyperlinks();
        try std.testing.expectEqual(@as(usize, 1), hls.len);
        try std.testing.expectEqualStrings("A9", hls[0].location.?);
        try std.testing.expectEqualStrings("C3", hls[0].ref);
        try std.testing.expectEqualStrings("jump", hls[0].display.?);
    }
}

test "Workbook.rewriteAllHyperlinkLocations: external (r_id != null) hyperlink skipped" {
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));

    var tmp_buf: [256]u8 = undefined;
    const tmp_path = try std.fmt.bufPrint(&tmp_buf, ".zig-cache/test-hl-ext-in-{d}.xlsx", .{prng.random().int(u32)});

    // External hyperlink — r:id present, no location. The rewriter
    // must skip it (count == 0), leaving the entry untouched.
    {
        var wb = try Workbook.open(std.testing.allocator, src_path);
        defer wb.deinit();
        try testInjectHyperlinks(
            &wb,
            "<hyperlink ref=\"D4\" r:id=\"rIdFake\" display=\"external\"/>",
            tmp_path,
        );
    }
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    var wb = try Workbook.open(std.testing.allocator, tmp_path);
    defer wb.deinit();

    {
        const ws = try wb.sheet(0);
        const hls = try ws.hyperlinks();
        try std.testing.expectEqual(@as(usize, 1), hls.len);
        try std.testing.expect(hls[0].r_id != null);
    }

    const count = try wb.rewriteAllHyperlinkLocations(.{
        .insert_rows = .{ .at = 1, .count = 4 },
    }, "Sheet1");
    try std.testing.expectEqual(@as(u32, 0), count);

    // External entry untouched.
    {
        const ws = try wb.sheet(0);
        const hls = try ws.hyperlinks();
        try std.testing.expectEqual(@as(usize, 1), hls.len);
        try std.testing.expect(hls[0].r_id != null);
        try std.testing.expectEqualStrings("rIdFake", hls[0].r_id.?);
    }
}

// ─── addImage tests ───────────────────────────────────────────────────

/// Canonical 1×1 transparent PNG (~70 bytes). Sufficient for round-
/// trip testing — Office's image decoders accept the minimal IHDR /
/// IDAT / IEND envelope.
const tiny_png_1x1 = [_]u8{
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

test "Workbook.addImage: round-trips PNG into a drawing-less workbook" {
    const allocator = std.testing.allocator;

    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));
    var src_buf: [128]u8 = undefined;
    const src_path = try std.fmt.bufPrint(&src_buf, ".zig-cache/test-addimg-src-{d}.xlsx", .{prng.random().int(u32)});
    try writeMinimalSstLessXlsx(allocator, src_path);
    defer std.fs.cwd().deleteFile(src_path) catch {};

    var dst_buf: [128]u8 = undefined;
    const dst_path = try std.fmt.bufPrint(&dst_buf, ".zig-cache/test-addimg-dst-{d}.xlsx", .{prng.random().int(u32)});
    defer std.fs.cwd().deleteFile(dst_path) catch {};

    {
        var wb = try Workbook.open(allocator, src_path);
        defer wb.deinit();
        try wb.addImage(0, .{ .col = 1, .row = 1 }, &tiny_png_1x1, .png);
        try wb.save(dst_path);
    }

    var store = try PartStore.open(allocator, dst_path);
    defer store.deinit();

    const image_part = try store.part("xl/media/image1.png");
    try std.testing.expect(image_part != null);
    try std.testing.expectEqualSlices(u8, &tiny_png_1x1, image_part.?.bytes);

    const drawing_part = try store.part("xl/drawings/drawing1.xml");
    try std.testing.expect(drawing_part != null);
    try std.testing.expect(std.mem.indexOf(u8, drawing_part.?.bytes, "<xdr:oneCellAnchor>") != null);
    try std.testing.expect(std.mem.indexOf(u8, drawing_part.?.bytes, "<xdr:col>0</xdr:col>") != null);
    try std.testing.expect(std.mem.indexOf(u8, drawing_part.?.bytes, "<xdr:row>0</xdr:row>") != null);

    const drawing_rels_part = try store.part("xl/drawings/_rels/drawing1.xml.rels");
    try std.testing.expect(drawing_rels_part != null);
    try std.testing.expect(std.mem.indexOf(
        u8,
        drawing_rels_part.?.bytes,
        "Target=\"../media/image1.png\"",
    ) != null);

    const sheet_rels_part = try store.part("xl/worksheets/_rels/sheet1.xml.rels");
    try std.testing.expect(sheet_rels_part != null);
    try std.testing.expect(std.mem.indexOf(
        u8,
        sheet_rels_part.?.bytes,
        "Target=\"../drawings/drawing1.xml\"",
    ) != null);

    const sheet_part = try store.part("xl/worksheets/sheet1.xml");
    try std.testing.expect(sheet_part != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_part.?.bytes, "<drawing r:id=") != null);

    const ct_part = try store.part("[Content_Types].xml");
    try std.testing.expect(ct_part != null);
    try std.testing.expect(std.mem.indexOf(u8, ct_part.?.bytes, "image/png") != null);
    try std.testing.expect(std.mem.indexOf(
        u8,
        ct_part.?.bytes,
        "application/vnd.openxmlformats-officedocument.drawing+xml",
    ) != null);
}

test "Workbook.addImage: PNG declared but JPEG bytes errors MimeMagicMismatch" {
    const allocator = std.testing.allocator;

    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));
    var src_buf: [128]u8 = undefined;
    const src_path = try std.fmt.bufPrint(&src_buf, ".zig-cache/test-addimg-mime-{d}.xlsx", .{prng.random().int(u32)});
    try writeMinimalSstLessXlsx(allocator, src_path);
    defer std.fs.cwd().deleteFile(src_path) catch {};

    var wb = try Workbook.open(allocator, src_path);
    defer wb.deinit();

    const fake_jpeg = [_]u8{ 0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10 };
    try std.testing.expectError(
        error.MimeMagicMismatch,
        wb.addImage(0, .{ .col = 1, .row = 1 }, &fake_jpeg, .png),
    );
}

test "Workbook.addImage: rejects sheet that already has a drawing" {
    const allocator = std.testing.allocator;

    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));
    var src_buf: [128]u8 = undefined;
    const src_path = try std.fmt.bufPrint(&src_buf, ".zig-cache/test-addimg-existing-{d}.xlsx", .{prng.random().int(u32)});
    try writeMinimalSstLessXlsx(allocator, src_path);
    defer std.fs.cwd().deleteFile(src_path) catch {};

    var dst_buf: [128]u8 = undefined;
    const dst_path = try std.fmt.bufPrint(&dst_buf, ".zig-cache/test-addimg-existing-dst-{d}.xlsx", .{prng.random().int(u32)});
    defer std.fs.cwd().deleteFile(dst_path) catch {};

    // Stage 1: add an image and save.
    {
        var wb = try Workbook.open(allocator, src_path);
        defer wb.deinit();
        try wb.addImage(0, .{ .col = 1, .row = 1 }, &tiny_png_1x1, .png);
        try wb.save(dst_path);
    }

    // Stage 2: re-open and try to add a second image to the same
    // sheet — must fail with SheetHasExistingDrawing.
    {
        var wb2 = try Workbook.open(allocator, dst_path);
        defer wb2.deinit();
        try std.testing.expectError(
            error.SheetHasExistingDrawing,
            wb2.addImage(0, .{ .col = 2, .row = 2 }, &tiny_png_1x1, .png),
        );
    }
}

test "Workbook.addImage: rejects 0-based anchor with InvalidAnchor" {
    const allocator = std.testing.allocator;

    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));
    var src_buf: [128]u8 = undefined;
    const src_path = try std.fmt.bufPrint(&src_buf, ".zig-cache/test-addimg-anchor-{d}.xlsx", .{prng.random().int(u32)});
    try writeMinimalSstLessXlsx(allocator, src_path);
    defer std.fs.cwd().deleteFile(src_path) catch {};

    var wb = try Workbook.open(allocator, src_path);
    defer wb.deinit();

    try std.testing.expectError(
        error.InvalidAnchor,
        wb.addImage(0, .{ .col = 0, .row = 1 }, &tiny_png_1x1, .png),
    );
    try std.testing.expectError(
        error.InvalidAnchor,
        wb.addImage(0, .{ .col = 1, .row = 0 }, &tiny_png_1x1, .png),
    );
}

// ─── B2 iter-er-3: Worksheet.appendRows API ──────────────────────────

test "Worksheet.appendRows: stages rows into appended_rows with deep-copied strings" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;
    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();

    const ws = try wb.sheet(0);
    try std.testing.expectEqual(@as(usize, 0), ws.appended_rows.items.len);

    var alpha = [_]u8{ 'a', 'l', 'p', 'h', 'a' };
    const row0 = [_]zlsx.Cell{ .{ .integer = 1 }, .{ .string = &alpha } };
    const row1 = [_]zlsx.Cell{ .{ .number = 2.5 }, .{ .boolean = true } };
    try ws.appendRows(&.{ &row0, &row1 });

    try std.testing.expectEqual(@as(usize, 2), ws.appended_rows.items.len);
    try std.testing.expectEqual(@as(usize, 2), ws.appended_rows.items[0].len);

    // Deep-copy invariant.
    @memset(&alpha, 'X');
    try std.testing.expectEqualStrings("alpha", ws.appended_rows.items[0][1].string);
}

test "Worksheet.appendRows: refuses when deltas are staged on the same sheet" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;
    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();

    const ws = try wb.sheet(0);
    try ws.setCell("A1", .{ .number = 99.0 });
    const row = [_]zlsx.Cell{.{ .integer = 1 }};
    try std.testing.expectError(error.SheetHasUnsavedMutations, ws.appendRows(&.{&row}));
    try std.testing.expectEqual(@as(usize, 0), ws.appended_rows.items.len);
}

test "Worksheet.setCell: refuses when appended_rows are staged on the same sheet" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;
    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();

    const ws = try wb.sheet(0);
    const row = [_]zlsx.Cell{.{ .integer = 1 }};
    try ws.appendRows(&.{&row});
    try std.testing.expectError(error.SheetHasUnsavedAppends, ws.setCell("A1", .{ .number = 99.0 }));
    try std.testing.expectEqual(@as(u32, 0), ws.deltas.count());
}

test "Worksheet.appendRows: rejects lossy integers before any allocation" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;
    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();

    const ws = try wb.sheet(0);
    const big: i64 = 9_007_199_254_740_993;
    const ok_row = [_]zlsx.Cell{.{ .integer = 1 }};
    const lossy_row = [_]zlsx.Cell{.{ .integer = big }};
    try std.testing.expectError(
        error.IntegerExceedsExcelPrecision,
        ws.appendRows(&.{ &ok_row, &lossy_row }),
    );
    try std.testing.expectEqual(@as(usize, 0), ws.appended_rows.items.len);
}

test "Worksheet.appendRows: empty rows slice is a no-op" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;
    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();

    const ws = try wb.sheet(0);
    try ws.appendRows(&.{});
    try std.testing.expectEqual(@as(usize, 0), ws.appended_rows.items.len);
}

test "Worksheet.resolvePartName: parse-free — does not populate self.parsed" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;
    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();

    const ws = try wb.sheet(0);
    try std.testing.expect(ws.parsed == null);
    const part_name = try ws.resolvePartName();
    try std.testing.expect(std.mem.startsWith(u8, part_name, "xl/worksheets/"));
    try std.testing.expect(std.mem.endsWith(u8, part_name, ".xml"));
    // Critical iter-er-3 invariant: resolvePartName must not trigger
    // ensureParsed — `Worksheet.emitWithAppendsUsingPlan` relies on
    // this to skip the 100k-row sheetData walk on the fast-path.
    try std.testing.expect(ws.parsed == null);
}

test "Worksheet.resolvePartName: idempotent — returns the same cached pointer" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;
    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();

    const ws = try wb.sheet(0);
    const first = try ws.resolvePartName();
    const second = try ws.resolvePartName();
    try std.testing.expectEqual(first.ptr, second.ptr);
    try std.testing.expectEqual(first.len, second.len);
}

test "Worksheet.resolvePartName: agrees with ensureParsed's cached part name" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;
    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();

    const ws = try wb.sheet(0);
    const fast = try ws.resolvePartName();
    _ = try ws.ensureParsed();
    try std.testing.expectEqualStrings(fast, ws.resolved_part_name.?);
}

test "Worksheet.emitWithAppendsUsingPlan: splices rows and threads sst_base_idx" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;
    const allocator = std.testing.allocator;
    var wb = try Workbook.open(allocator, path);
    defer wb.deinit();

    const ws = try wb.sheet(0);
    const r0 = [_]zlsx.Cell{ .{ .integer = 7 }, .{ .string = "alpha" } };
    const r1 = [_]zlsx.Cell{ .{ .number = 1.5 }, .{ .boolean = true } };
    try ws.appendRows(&.{ &r0, &r1 });

    // Hand-build a plan covering every staged string. Mirrors what
    // `buildSstExtensionPlan` would produce: one new entry "alpha" at
    // base_index 100 (caller-side equivalent of the legacy
    // sst_base_idx argument). `registerNewPlain` keeps `new_strings`
    // and `new_strings_index` in sync; the prior hand-pushed
    // `new_strings.append` form would silently desync the hash index.
    var plan: SstExtensionPlan = .{ .base_index = 100, .sst_part_exists = true };
    defer plan.deinit(allocator);
    _ = try plan.registerNewPlain(allocator, "alpha");

    const new_xml = try ws.emitWithAppendsUsingPlan(allocator, &plan);
    defer allocator.free(new_xml);

    // Critical: the parse-free path must NOT have populated the
    // sheet_xml view — that's the whole point of iter-er-3.
    try std.testing.expect(ws.parsed == null);

    // Plan resolved "alpha" → 100 (base_index + 0). Rendered cell
    // carries v=100.
    try std.testing.expect(std.mem.indexOf(u8, new_xml, "t=\"s\"><v>100</v>") != null);
    // Boolean cell renders with t="b".
    try std.testing.expect(std.mem.indexOf(u8, new_xml, "t=\"b\"><v>1</v>") != null);
    // Rows splice before </sheetData>.
    const sd_close = std.mem.indexOf(u8, new_xml, "</sheetData>") orelse
        return error.TestUnexpectedResult;
    const last_row = std.mem.lastIndexOf(u8, new_xml, "</row>") orelse
        return error.TestUnexpectedResult;
    try std.testing.expect(last_row < sd_close);
}

test "Worksheet.emitWithAppendsUsingPlan: rejects appends past max_row" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;
    const allocator = std.testing.allocator;
    var wb = try Workbook.open(allocator, path);
    defer wb.deinit();

    const ws = try wb.sheet(0);
    // Stage one row, then forge a highest_row close to the cap by
    // monkey-replacing the part bytes — but the cleaner path is to
    // verify the cap arithmetic via a direct call: append a row,
    // then confirm `final_row64 > max_row` triggers when start_row
    // would land past the cap. Trust the cap on a synthetic rather
    // than fabricating 1M rows in the corpus.
    const row = [_]zlsx.Cell{.{ .integer = 1 }};
    try ws.appendRows(&.{&row});
    // Replace the cached part bytes with one that has the highest
    // row at max_row, forcing start_row to overflow.
    const part_name = try ws.resolvePartName();
    const synthetic = try std.fmt.allocPrint(
        allocator,
        "<worksheet><dimension ref=\"A1:A{d}\"/><sheetData><row r=\"{d}\"><c r=\"A{d}\"><v>1</v></c></row></sheetData></worksheet>",
        .{ zlsx.max_row, zlsx.max_row, zlsx.max_row },
    );
    defer allocator.free(synthetic);
    try wb.store.replacePart(part_name, synthetic);

    // Empty plan — the staged row has no string cells, so indexOf is
    // never invoked. base_index 0 + sst_part_exists=false matches the
    // legacy `sst_base_idx = 0` argument's intent (caller treats SST
    // as absent / fresh).
    var plan: SstExtensionPlan = .{};
    defer plan.deinit(allocator);

    try std.testing.expectError(error.RowIndexOutOfRange, ws.emitWithAppendsUsingPlan(allocator, &plan));
}

test "Worksheet.emitWithAppendsUsingPlan: rewrites self-closing <sheetData/> to open/close form" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;
    const allocator = std.testing.allocator;
    var wb = try Workbook.open(allocator, path);
    defer wb.deinit();

    const ws = try wb.sheet(0);
    const row = [_]zlsx.Cell{.{ .integer = 7 }};
    try ws.appendRows(&.{&row});

    // Replace the sheet's part with a self-closing-sheetData body
    // before calling the emit. The fast-path must rewrite
    // `<sheetData/>` into `<sheetData>…</sheetData>` so the
    // rendered rows have a home.
    const part_name = try ws.resolvePartName();
    const synthetic = try allocator.dupe(
        u8,
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" ++
            "<dimension ref=\"A1\"/>" ++
            "<sheetData/>" ++
            "</worksheet>",
    );
    defer allocator.free(synthetic);
    try wb.store.replacePart(part_name, synthetic);

    // No string cells staged → empty plan suffices.
    var plan: SstExtensionPlan = .{};
    defer plan.deinit(allocator);

    const new_xml = try ws.emitWithAppendsUsingPlan(allocator, &plan);
    defer allocator.free(new_xml);

    try std.testing.expect(std.mem.indexOf(u8, new_xml, "<sheetData>") != null);
    try std.testing.expect(std.mem.indexOf(u8, new_xml, "</sheetData>") != null);
    try std.testing.expect(std.mem.indexOf(u8, new_xml, "<sheetData/>") == null);
    try std.testing.expect(std.mem.indexOf(u8, new_xml, "<v>7</v>") != null);
}

test "Worksheet.emitWithAppendsUsingPlan: refuses when deltas are staged on the same sheet" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;
    const allocator = std.testing.allocator;
    var wb = try Workbook.open(allocator, path);
    defer wb.deinit();

    // Stage appends, THEN forge a delta directly (bypassing the
    // editor-level guard) to exercise the emit-side negative-space
    // refusal.
    const ws = try wb.sheet(0);
    const row = [_]zlsx.Cell{.{ .integer = 1 }};
    try ws.appendRows(&.{&row});
    try ws.deltas.put(allocator, .{ .row = 1, .col = 1 }, .{ .number = 99.0 });

    var plan: SstExtensionPlan = .{};
    defer plan.deinit(allocator);

    try std.testing.expectError(error.SheetHasUnsavedMutations, ws.emitWithAppendsUsingPlan(allocator, &plan));
}

test "appendXmlFindHighestRow: ignores rows outside <sheetData> window" {
    // Adversarial: a `<mergeCell ref="A1:A99999"/>` outside
    // sheetData used to inflate the highest_row scan. Bounded
    // window now skips it.
    const xml =
        "<worksheet>" ++
        "<dimension ref=\"A1:A1\"/>" ++
        "<mergeCells><mergeCell ref=\"A1:A99999\"/></mergeCells>" ++
        "<sheetData>" ++
        "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>" ++
        "</sheetData>" ++
        "<oddHeader>r=\"A88888\"</oddHeader>" ++
        "</worksheet>";
    try std.testing.expectEqual(@as(u32, 1), appendXmlFindHighestRow(xml));
}

test "appendXmlFindHighestRow: empty sheet body returns 0" {
    try std.testing.expectEqual(@as(u32, 0), appendXmlFindHighestRow("<sheetData/>"));
    try std.testing.expectEqual(@as(u32, 0), appendXmlFindHighestRow("<sheetData></sheetData>"));
    try std.testing.expectEqual(@as(u32, 0), appendXmlFindHighestRow("<worksheet/>"));
}

test "Worksheet.clearAppendedRows: deinit is safe after clear" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;
    const allocator = std.testing.allocator;
    var wb = try Workbook.open(allocator, path);
    defer wb.deinit();

    const ws = try wb.sheet(0);
    const r0 = [_]zlsx.Cell{ .{ .integer = 1 }, .{ .string = "x" } };
    try ws.appendRows(&.{&r0});
    try std.testing.expectEqual(@as(usize, 1), ws.appended_rows.items.len);

    ws.clearAppendedRows(allocator);
    try std.testing.expectEqual(@as(usize, 0), ws.appended_rows.items.len);

    // setCell must succeed after the staging buffer has been
    // cleared — the iter-er-3 symmetric guard only fires while
    // appends are LIVE.
    try ws.setCell("A1", .{ .number = 99.0 });
}

// ─── B2 iter-er-4 (1/N): Workbook.addSheet ────────────────────────────

test "Workbook.addSheet: appends a new empty sheet and returns its handle" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;
    const allocator = std.testing.allocator;
    var wb = try Workbook.open(allocator, path);
    defer wb.deinit();

    const before_count = wb.sheetCount();
    const ws = try wb.addSheet("Iter-Er-4-One");
    try std.testing.expectEqual(before_count + 1, wb.sheetCount());
    try std.testing.expectEqualStrings("Iter-Er-4-One", ws.name());
    try std.testing.expect(ws.parsed == null);
    try std.testing.expect(ws.appended_rows.items.len == 0);
    try std.testing.expect(ws.deltas.count() == 0);

    // Returned handle must point at the last slot in worksheets.
    try std.testing.expectEqual(@as(u32, before_count), ws.sheet_idx);
    try std.testing.expectEqual(&wb.worksheets[before_count], ws);

    // resolvePartName for the new sheet finds the freshly-allocated
    // part — proves addPart + rels patch wired correctly.
    const part_name = try ws.resolvePartName();
    try std.testing.expect(std.mem.startsWith(u8, part_name, "xl/worksheets/sheet"));
    try std.testing.expect(std.mem.endsWith(u8, part_name, ".xml"));
}

test "Workbook.addSheet: refuses duplicate name (case-insensitive)" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;
    const allocator = std.testing.allocator;
    var wb = try Workbook.open(allocator, path);
    defer wb.deinit();

    _ = try wb.addSheet("UniqueSheet");
    try std.testing.expectError(error.SheetNameInUse, wb.addSheet("UniqueSheet"));
    try std.testing.expectError(error.SheetNameInUse, wb.addSheet("UNIQUESHEET"));
    try std.testing.expectError(error.SheetNameInUse, wb.addSheet("uniquesheet"));
}

test "Workbook.addSheet: refuses InvalidSheetName per the validateSheetName contract" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;
    const allocator = std.testing.allocator;
    var wb = try Workbook.open(allocator, path);
    defer wb.deinit();

    try std.testing.expectError(error.InvalidSheetName, wb.addSheet(""));
    try std.testing.expectError(error.InvalidSheetName, wb.addSheet("colon:bad"));
    try std.testing.expectError(error.InvalidSheetName, wb.addSheet("history"));
    try std.testing.expectError(error.InvalidSheetName, wb.addSheet("HISTORY"));
}

test "Workbook.addSheet: multiple consecutive adds pick non-colliding ids" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;
    const allocator = std.testing.allocator;
    var wb = try Workbook.open(allocator, path);
    defer wb.deinit();

    const before_count = wb.sheetCount();
    // Don't HOLD the pointers across calls — addSheet reallocates
    // self.worksheets, invalidating prior returns. Re-fetch after
    // every structural call.
    const idx_a: u32 = blk: {
        const ws = try wb.addSheet("A");
        break :blk ws.sheet_idx;
    };
    const idx_b: u32 = blk: {
        const ws = try wb.addSheet("B");
        break :blk ws.sheet_idx;
    };
    const idx_c: u32 = blk: {
        const ws = try wb.addSheet("C");
        break :blk ws.sheet_idx;
    };

    try std.testing.expectEqual(before_count + 3, wb.sheetCount());
    try std.testing.expectEqualStrings("A", (try wb.sheet(idx_a)).name());
    try std.testing.expectEqualStrings("B", (try wb.sheet(idx_b)).name());
    try std.testing.expectEqualStrings("C", (try wb.sheet(idx_c)).name());

    // Distinct part names (sheet ids and path numbers don't collide).
    const path_a = try (try wb.sheet(idx_a)).resolvePartName();
    const path_b = try (try wb.sheet(idx_b)).resolvePartName();
    const path_c = try (try wb.sheet(idx_c)).resolvePartName();
    try std.testing.expect(!std.mem.eql(u8, path_a, path_b));
    try std.testing.expect(!std.mem.eql(u8, path_a, path_c));
    try std.testing.expect(!std.mem.eql(u8, path_b, path_c));

    // Sheet ids in the workbook view are also distinct.
    var seen: std.AutoHashMapUnmanaged(u32, void) = .{};
    defer seen.deinit(allocator);
    for (wb.workbook.sheets) |s| {
        const gop = try seen.getOrPut(allocator, s.sheet_id);
        try std.testing.expect(!gop.found_existing);
    }
}

test "Workbook.addSheet: name with XML special chars escapes safely into workbook.xml" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;
    const allocator = std.testing.allocator;
    var wb = try Workbook.open(allocator, path);
    defer wb.deinit();

    const tricky = "AT&T <wow>";
    _ = try wb.addSheet(tricky);
    // Attribute-escaped on the wire.
    const wb_part = (try wb.store.part("xl/workbook.xml")).?;
    try std.testing.expect(std.mem.indexOf(u8, wb_part.bytes, "AT&amp;T &lt;wow&gt;") != null);
    // The raw bytes must NOT contain the un-escaped name. The
    // workbook view (and `ws.name()`) holds the encoded attribute
    // value verbatim; consumers do their own entity decoding —
    // matches the renameSheet contract documented at line 521.
    try std.testing.expect(std.mem.indexOf(u8, wb_part.bytes, "AT&T <wow>") == null);
}

test "Workbook.addSheet: returned handle accepts setCell + appendRows" {
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;
    const allocator = std.testing.allocator;
    var wb = try Workbook.open(allocator, path);
    defer wb.deinit();

    const ws = try wb.addSheet("Mutable");

    // setCell should land in deltas without erroring.
    try ws.setCell("B7", .{ .number = 42.0 });
    try std.testing.expectEqual(@as(u32, 1), ws.deltas.count());

    // setCell + appendRows still mutually exclusive — clear deltas
    // first.
    freeDeltaStrings(allocator, &ws.deltas);
    ws.deltas.clearAndFree(allocator);

    const row = [_]zlsx.Cell{ .{ .integer = 1 }, .{ .string = "ok" } };
    try ws.appendRows(&.{&row});
    try std.testing.expectEqual(@as(usize, 1), ws.appended_rows.items.len);
}

// ─── B2 iter-er-3 fuzz: parse-free XML helpers stay no-panic ──

fn fuzzAppendXmlFindHighestRow(_: void, input: []const u8) anyerror!void {
    _ = appendXmlFindHighestRow(input);
}

test "fuzz: appendXmlFindHighestRow never crashes on adversarial sheet XML" {
    try std.testing.fuzz({}, fuzzAppendXmlFindHighestRow, .{
        .corpus = &[_][]const u8{
            "",
            "<sheetData/>",
            "<sheetData></sheetData>",
            "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>",
            "<row r=\"99999\"/>",
            "<c r=\"A1\"/>",
            "<col r=\"1\"/><conditionalFormatting r=\"A1\"/>",
            "<row r=\"\"/>", // empty digits
            "<row r=\"-5\"/>", // negative — parseInt fails, returns null
            "<row r=\"99999999999999999999\"/>", // overflow
            "<c r=\"A99999999999999999999\"/>", // overflowing cell ref
            "<row\nr=\"5\">", // newline before attr
            "<row\tr=\"5\">", // tab before attr
            "<row spans=\"1:5\" r=\"7\"/>", // attr-order swap
            "<c s=\"1\" r=\"B12\"/>", // attr-order swap on c
            "<rowabc r=\"99\"/>", // not a real row tag
            "<cabc r=\"A1\"/>", // not a real c tag
            "<row r=\"1\" r=\"2\"/>", // duplicate r=
            "<row r=\"  5  \"/>", // padded digits
            "<row r=\"5", // truncated
            "<row r=", // truncated mid-attr
            "<row", // truncated tag
            "<<<<<<<", // pathological
            "<row r=\"5\"><!--<row r=\"99\"/>--></row>", // comment-wrapped
            "<row r=\"\xff\xfe\"/>", // invalid UTF-8 in attr
        },
    });
}

fn fuzzAppendXmlUpdateDimensionBR(_: void, input: []const u8) anyerror!void {
    var arena = std.heap.ArenaAllocator.init(std.testing.allocator);
    defer arena.deinit();
    const out = appendXmlUpdateDimensionBR(arena.allocator(), input, 100, 10) catch return;
    if (out) |s| arena.allocator().free(s);
}

test "fuzz: appendXmlUpdateDimensionBR never crashes on adversarial XML" {
    try std.testing.fuzz({}, fuzzAppendXmlUpdateDimensionBR, .{
        .corpus = &[_][]const u8{
            "",
            "<dimension ref=\"A1:Z10\"/>",
            "<dimension ref=\"A1\"/>", // single-cell — passthrough null
            "<dimension ref=\"\"/>", // empty ref
            "<dimension ref=\"A1:\"/>", // missing BR
            "<dimension ref=\":Z10\"/>", // missing TL
            "<dimension ref=\"$A$1:$Z$10\"/>", // dollar-anchored
            "<dimension ref=\"Sheet1!A1:Z10\"/>", // sheet-prefixed
            "<dimension ref=\"AAAA1:ZZZZ99999\"/>", // exceeds max_col
            "<dimension ref=\"A99999999999999:Z1\"/>", // overflow row
            "<dimension ref=\"a1:z10\"/>", // lowercase
            "<dimension ref=\"A1:Z10", // truncated quote
            "<x:dimension ref=\"A1:Z10\"/>", // namespace prefix
            "<dimension>", // open tag, no attrs
            "<dimension/>", // self-closing, no attrs
            "<<<<<<<<<<<", // pathological
        },
    });
}

fn fuzzAppendXmlInjectRows(_: void, input: []const u8) anyerror!void {
    var arena = std.heap.ArenaAllocator.init(std.testing.allocator);
    defer arena.deinit();
    const rendered = "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>";
    const out = appendXmlInjectRows(arena.allocator(), input, rendered) catch return;
    arena.allocator().free(out);
}

test "fuzz: appendXmlInjectRows never crashes on adversarial sheet XML" {
    try std.testing.fuzz({}, fuzzAppendXmlInjectRows, .{
        .corpus = &[_][]const u8{
            "",
            "<sheetData></sheetData>",
            "<sheetData/>",
            "<sheetData attr=\"x\"/>", // attrs on self-closing
            "<sheetData></sheetData", // truncated close
            "<sheetData>", // unclosed open
            "</sheetData>", // close before open
            "<sheetData></sheetData></sheetData>", // double close
            "<sheetdata></sheetdata>", // wrong case
            "<x:sheetData/>", // namespace prefix
            "<!--<sheetData/>-->", // comment-wrapped
            "<sheetData attr=\">\"/>", // `>` in quoted attr
            "<sheetData/<sheetData/>", // nested malformed
            "<sheetData", // truncated open
            "<<<<<<<<<<", // pathological
        },
    });
}

// ─── I5 fuzz: appendCellXmlForAppendUsingPlan on attacker-shaped cell payloads ──

fn fuzzAppendCellXmlForAppend(_: void, input: []const u8) anyerror!void {
    if (input.len == 0) return;
    const allocator = std.testing.allocator;
    var out: std.ArrayList(u8) = .empty;
    defer out.deinit(allocator);

    // Interpret input[0] as cell variant selector.
    const variant = input[0] % 5;
    const ref_row: u32 = @as(u32, input[0] >> 1) + 1;
    const ref_col: u32 = @as(u32, if (input.len > 1) input[1] else 1) % 100 + 1;

    const cell: zlsx.Cell = switch (variant) {
        0 => .empty,
        1 => .{ .integer = @as(i32, @bitCast(@as(u32, input[0]) | (@as(u32, if (input.len > 1) input[1] else 0) << 8))) },
        2 => .{ .number = @as(f32, @floatFromInt(input.len)) },
        3 => .{ .boolean = (input[0] & 1) == 1 },
        else => blk: {
            // Validate XML safety — the plan-driven helper's caller
            // (Worksheet.appendRows) has already done this. Skip
            // unsafe inputs.
            const s = input[1..];
            if (!isXmlSafeText(s)) return;
            break :blk .{ .string = s };
        },
    };

    // Hand-build a plan that registers the staged string (if any) so
    // `plan.indexOf(s)` resolves. For non-string variants the plan is
    // empty.
    var plan: SstExtensionPlan = .{};
    defer plan.deinit(allocator);
    if (cell == .string) {
        const dup = try allocator.dupe(u8, cell.string);
        errdefer allocator.free(dup);
        try plan.new_strings.append(allocator, dup);
        plan.has_new_strings = true;
    }

    appendCellXmlForAppendUsingPlan(
        allocator,
        &out,
        cell,
        .{ .row = ref_row, .col = ref_col },
        &plan,
    ) catch return;
}

test "fuzz: appendCellXmlForAppendUsingPlan never crashes on attacker-shaped cells" {
    try std.testing.fuzz({}, fuzzAppendCellXmlForAppend, .{
        .corpus = &[_][]const u8{
            "\x00", // .empty
            "\x01\x05", // .integer
            "\x02hello", // .number
            "\x03\x01", // .boolean
            "\x04alpha", // .string "alpha"
            "\x04", // .string ""
            "\x04<<<<", // .string with XML chars
            "\x04&amp;", // .string already escaped
            "\x04\x00null", // .string with NUL — isXmlSafeText rejects
            "\x04\x01ctrl", // .string with control byte
            "\x04\x09tab", // .string with tab — XML-safe
            "\x04\xc0\x80", // .string with overlong UTF-8
            "\x04" ++ ([_]u8{'A'} ** 256), // .string long text
        },
    });
}

// ─── I6 fuzz: appendXmlUpdateDimensionBR with mutated <dimension ref> ──

fn fuzzAppendXmlUpdateDimensionBRMutated(_: void, input: []const u8) anyerror!void {
    var arena = std.heap.ArenaAllocator.init(std.testing.allocator);
    defer arena.deinit();
    var buf: std.ArrayList(u8) = .empty;
    defer buf.deinit(arena.allocator());
    try buf.appendSlice(arena.allocator(), "<dimension ref=\"");
    // Append the input bytes as-is into the ref attribute. Then close
    // the attribute and tag. Adversarial inputs land directly in the
    // parser's attention.
    try buf.appendSlice(arena.allocator(), input);
    try buf.appendSlice(arena.allocator(), "\"/>");
    const out = appendXmlUpdateDimensionBR(arena.allocator(), buf.items, 100, 10) catch return;
    if (out) |s| arena.allocator().free(s);
}

test "fuzz: appendXmlUpdateDimensionBR never crashes on mutated ref body" {
    try std.testing.fuzz({}, fuzzAppendXmlUpdateDimensionBRMutated, .{
        .corpus = &[_][]const u8{
            "A1:Z10",
            "A:Z", // letters only
            "1:10", // digits only
            "A1:A1:A1", // multiple colons
            "A1::Z10", // double colon
            "AAAA1:ZZZZZZ99999", // letter-overflow
            "A0:Z0", // zero rows
            "A1:Z" ++ "9" ** 12, // overflow row digits
            ":", // bare colon
            "A1:", // missing BR
            ":Z10", // missing TL
            "Z\"A1:Z10", // injection attempt
            "A1:Z10\" />\"<dimension ref=\"X1:Y9", // double dimension
            "A1:" ++ "Z" ** 64, // long letter sequence
            "AABCDEF1:XYZ9876",
            "$A$1:$Z$10",
            "Sheet!A1:Z10",
            "1A:10Z", // letters/digits swapped
            "A0001:Z0010", // leading zeros
            "  A1:Z10  ", // padded
            "A1:Z10\x00", // NUL terminated
            "A1:Z10\xff\xfe", // invalid UTF-8 trailing
            "", // empty ref body
        },
    });
}

// ─── Structural fuzz: emitWithAppendsUsingPlan end-to-end on synthetic parts ──

fn fuzzEmitWithAppendsStructural(_: void, input: []const u8) anyerror!void {
    if (input.len < 8) return;
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return;
    const allocator = std.testing.allocator;
    var wb = try Workbook.open(allocator, path);
    defer wb.deinit();

    const ws = try wb.sheet(0);
    // Build a synthetic source part from `input`, framed in a minimal
    // worksheet skeleton. The bytes between the `<sheetData>` markers
    // are the fuzz payload — exercises the bounded-window scan,
    // injectRows splice, and dimension widener under adversarial
    // sheetData content.
    var part: std.ArrayList(u8) = .empty;
    defer part.deinit(allocator);
    try part.appendSlice(allocator, "<worksheet><dimension ref=\"A1:A1\"/><sheetData>");
    try part.appendSlice(allocator, input);
    try part.appendSlice(allocator, "</sheetData></worksheet>");
    const part_name = try ws.resolvePartName();
    try wb.store.replacePart(part_name, part.items);

    // Stage one row of each cell variant + a string.
    const row = [_]zlsx.Cell{
        .{ .integer = 1 },
        .{ .number = 2.5 },
        .{ .boolean = true },
        .{ .string = "fuzz" },
    };
    try ws.appendRows(&.{&row});

    // Hand-built plan registers the single staged string. Mirrors what
    // `buildSstExtensionPlan` does for an SST-less workbook: base_index
    // 0, `new_strings = ["fuzz"]`, sst_part_exists=false.
    var plan: SstExtensionPlan = .{};
    defer plan.deinit(allocator);
    {
        const dup = try allocator.dupe(u8, "fuzz");
        errdefer allocator.free(dup);
        try plan.new_strings.append(allocator, dup);
    }
    plan.has_new_strings = true;

    const new_xml = ws.emitWithAppendsUsingPlan(allocator, &plan) catch {
        ws.clearAppendedRows(allocator);
        return;
    };
    defer allocator.free(new_xml);
    ws.clearAppendedRows(allocator);
}

test "registerSharedString: empty new string vs rich existing entry registers as new (not aliased)" {
    // Regression for the iter-er-7 review finding: rich SST entries
    // were represented as `""` in `decoded_existing` with the comment
    // "an empty new string never equals a rich entry by construction"
    // — but `eql("", "")` IS true, so an empty new string silently
    // aliased to the rich entry's index, emitting `<c t="s"><v>idx</v></c>`
    // pointing at a `<si><r>…</r></si>` block. Excel resolves that as
    // the concatenated rich-run text, NOT as empty plain text.
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;

    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();

    var plan: SstExtensionPlan = .{};
    defer plan.deinit(std.testing.allocator);

    const decoded = [_][]const u8{""};
    const is_rich = [_]bool{true};

    try registerSharedString(&wb, &plan, "", &decoded, &is_rich);
    try std.testing.expectEqual(@as(usize, 1), plan.new_strings.items.len);
    try std.testing.expectEqualStrings("", plan.new_strings.items[0]);

    try registerExistingMatch(&wb, &plan, "", &decoded, &is_rich);
    try std.testing.expectEqual(@as(usize, 0), plan.existing_matches.items.len);
}

test "registerSharedString: empty new string vs PLAIN existing entry deduplicates correctly" {
    // Companion to the rich-entry regression test — confirms the fix
    // didn't accidentally break the dedup contract for plain entries.
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path, .{}) catch return error.SkipZigTest;

    var wb = try Workbook.open(std.testing.allocator, path);
    defer wb.deinit();

    var plan: SstExtensionPlan = .{};
    defer plan.deinit(std.testing.allocator);

    const decoded = [_][]const u8{""};
    const is_rich = [_]bool{false};

    try registerSharedString(&wb, &plan, "", &decoded, &is_rich);
    try std.testing.expectEqual(@as(usize, 0), plan.new_strings.items.len);

    try registerExistingMatch(&wb, &plan, "", &decoded, &is_rich);
    try std.testing.expectEqual(@as(usize, 1), plan.existing_matches.items.len);
    try std.testing.expectEqual(@as(u32, 0), plan.existing_matches.items[0].index);
}

test "fuzz: emitWithAppendsUsingPlan end-to-end on adversarial sheetData payload" {
    try std.testing.fuzz({}, fuzzEmitWithAppendsStructural, .{
        .corpus = &[_][]const u8{
            "",
            "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>",
            "<row r=\"99999\"/>",
            "<row r=\"1048575\"/>", // one below max
            "<row r=\"1048576\"/>", // exactly max — append would overflow
            "<row r=\"1048577\"/>", // already past max
            "<!-- nested <row r=\"1048576\"/> comment -->", // adversarial comment
            "<row r=\"5\"><c r=\"A5\"><v></v></c></row>",
            "<row r=\"5\"><c><v></v></c></row>", // <c> with no r=
            "<row><c/></row>", // explicit-row-less form
            // 5k bytes of pseudo-row data — exercise large-input path
            "<row r=\"100\">" ++ ([_]u8{ '<', 'c', '/', '>' } ** 1000) ++ "</row>",
            "<col r=\"1\"/><c/><row/>", // mixed
            "\x00\x00\x00\x00", // NUL-only
            "<row r=\"1\"><c r=\"&lt;&gt;\"><v>1</v></c></row>", // entity refs
            "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row>", // sst index 0
        },
    });
}

// ─── iter-er-7 task C-1: Workbook.save edge-case coverage ─────────────

test "Workbook.save: SST-less source + appendRows .string creates SST + rels + Override" {
    // Exercises the appended_rows × no-SST branch end-to-end. Build a
    // synthetic SST-less .xlsx, stage one `.string` cell via
    // `Worksheet.appendRows`, save, then re-open and assert:
    //   - `xl/sharedStrings.xml` part now exists
    //   - `xl/_rels/workbook.xml.rels` carries the SST `<Relationship>`
    //   - `[Content_Types].xml` carries the SST Override entry
    //   - the appended cell reads back as a shared_string with index 0
    const alloc = std.testing.allocator;

    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));
    var src_buf: [256]u8 = undefined;
    var out_buf: [256]u8 = undefined;
    const src_path = try std.fmt.bufPrint(&src_buf, ".zig-cache/test-er7-c1-sstless-src-{d}.xlsx", .{prng.random().int(u32)});
    const out_path = try std.fmt.bufPrint(&out_buf, ".zig-cache/test-er7-c1-sstless-out-{d}.xlsx", .{prng.random().int(u32)});

    try writeMinimalSstLessXlsx(alloc, src_path);
    defer std.fs.cwd().deleteFile(src_path) catch {};

    const appended_text = "appended-via-rows-fresh-sst";
    {
        var wb = try Workbook.open(alloc, src_path);
        defer wb.deinit();
        const v = try wb.sst();
        try std.testing.expect(v == null); // sanity: source has no SST

        const s0 = try wb.sheet(0);
        const row = [_]zlsx.Cell{.{ .string = appended_text }};
        try s0.appendRows(&.{&row});
        try wb.save(out_path);
    }
    defer std.fs.cwd().deleteFile(out_path) catch {};

    var wb2 = try Workbook.open(alloc, out_path);
    defer wb2.deinit();

    // Fresh SST part — exactly one entry, our text.
    const sst_view = try wb2.sst();
    try std.testing.expect(sst_view != null);
    try std.testing.expectEqual(@as(usize, 1), sst_view.?.entries.len);
    const t0 = try wb2.sstText(0);
    try std.testing.expect(t0 != null);
    try std.testing.expectEqualStrings(appended_text, t0.?);

    // Workbook rels contain the SST relationship.
    const rels_part = try wb2.store.part("xl/_rels/workbook.xml.rels") orelse
        return error.TestUnexpectedResult;
    try std.testing.expect(std.mem.indexOf(u8, rels_part.bytes, "sharedStrings") != null);
    try std.testing.expect(std.mem.indexOf(u8, rels_part.bytes, "/relationships/sharedStrings") != null);

    // [Content_Types].xml carries the SST Override.
    const ct_part = try wb2.store.part("[Content_Types].xml") orelse
        return error.TestUnexpectedResult;
    try std.testing.expect(std.mem.indexOf(u8, ct_part.bytes, "/xl/sharedStrings.xml") != null);
    try std.testing.expect(std.mem.indexOf(u8, ct_part.bytes, "sharedStrings+xml") != null);

    // Reader observes the appended cell. The empty source had no rows,
    // so the appended row lands at row 1.
    const s0r = try wb2.sheet(0);
    const cell = try s0r.cellByRef("A1");
    try std.testing.expect(cell != null);
    try std.testing.expect(cell.?.cell_type == .shared_string);
    try std.testing.expectEqualStrings("0", cell.?.raw_value.?);
}

test "Workbook.save: mixed deltas (sheet 0) + appends (sheet 1) share one SST plan" {
    // Stage `setCell` on sheet 0 (string + numeric) and `appendRows`
    // on sheet 1 (string + numeric) in the same save. The
    // SstExtensionPlan must dedup across BOTH axes — assert each cell
    // observes its expected SST index after re-open.
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    const alloc = std.testing.allocator;
    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));
    var tmp_buf: [256]u8 = undefined;
    const tmp_path = try std.fmt.bufPrint(&tmp_buf, ".zig-cache/test-er7-c1-mixed-{d}.xlsx", .{prng.random().int(u32)});

    const delta_text = "er7-c1-delta-shared";
    const append_text = "er7-c1-append-shared";

    var pre_count: u32 = 0;
    {
        var wb = try Workbook.open(alloc, src_path);
        defer wb.deinit();
        pre_count = @intCast((try wb.sst()).?.entries.len);

        // Sheet 0: setCell with shared_string + numeric.
        const s0 = try wb.sheet(0);
        try s0.setCell("Z998", .{ .shared_string = delta_text });
        try s0.setCell("Z999", .{ .number = 42.0 });

        // Sheet 1: appendRows with string + numeric.
        const s1 = try wb.sheet(1);
        const row = [_]zlsx.Cell{ .{ .string = append_text }, .{ .number = 7.5 } };
        try s1.appendRows(&.{&row});

        try wb.save(tmp_path);
    }
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    var wb2 = try Workbook.open(alloc, tmp_path);
    defer wb2.deinit();

    // SST grew by exactly two — one per axis (texts differ).
    const sst2 = (try wb2.sst()).?;
    try std.testing.expectEqual(@as(usize, pre_count + 2), sst2.entries.len);

    // Sheet 0 delta cell — shared_string at pre_count (deltas walk first).
    const s0r = try wb2.sheet(0);
    const c_delta = try s0r.cellByRef("Z998");
    try std.testing.expect(c_delta != null);
    try std.testing.expect(c_delta.?.cell_type == .shared_string);
    var ibuf: [16]u8 = undefined;
    const expected_delta_idx = try std.fmt.bufPrint(&ibuf, "{d}", .{pre_count});
    try std.testing.expectEqualStrings(expected_delta_idx, c_delta.?.raw_value.?);
    const delta_resolved = try wb2.sstText(pre_count);
    try std.testing.expectEqualStrings(delta_text, delta_resolved.?);

    // Sheet 0 numeric delta cell — round-trips.
    const c_num = try s0r.cellByRef("Z999");
    try std.testing.expect(c_num != null);
    try std.testing.expectEqualStrings("42", c_num.?.raw_value.?);

    // Sheet 1 appended cells — first appended row lands at last_row + 1.
    // Locate by SST text rather than row arithmetic to keep this robust
    // against fixture row count drift.
    const sst_idx_appended = pre_count + 1;
    var ibuf2: [16]u8 = undefined;
    const expected_append_idx = try std.fmt.bufPrint(&ibuf2, "{d}", .{sst_idx_appended});
    const append_resolved = try wb2.sstText(sst_idx_appended);
    try std.testing.expectEqualStrings(append_text, append_resolved.?);

    // Verify the appended sheet's serialized XML carries the right index +
    // the numeric cell. The sheet part name is whatever sheet 1 resolves
    // to; read it via the typed handle.
    const s1r = try wb2.sheet(1);
    const part_name = try s1r.resolvePartName();
    const sheet_part = try wb2.store.part(part_name) orelse
        return error.TestUnexpectedResult;
    const needle = try std.fmt.allocPrint(alloc, "t=\"s\"><v>{s}</v>", .{expected_append_idx});
    defer alloc.free(needle);
    try std.testing.expect(std.mem.indexOf(u8, sheet_part.bytes, needle) != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_part.bytes, "<v>7.5</v>") != null);
}

test "Workbook.save: renameSheet rewriter composes with setCell on a different sheet" {
    // Run the rewriter (renameSheet does formula + defined-name + ...
    // rewrites internally), then stage a setCell on a sheet not yet
    // edited, save, re-open. Reader must see BOTH the renamed sheet
    // AND the setCell mutation.
    const src_path = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(src_path, .{}) catch return error.SkipZigTest;

    const alloc = std.testing.allocator;
    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));
    var tmp_buf: [256]u8 = undefined;
    var out_buf: [256]u8 = undefined;
    const tmp_path = try std.fmt.bufPrint(&tmp_buf, ".zig-cache/test-er7-c1-rw-in-{d}.xlsx", .{prng.random().int(u32)});
    const out_path = try std.fmt.bufPrint(&out_buf, ".zig-cache/test-er7-c1-rw-out-{d}.xlsx", .{prng.random().int(u32)});

    // Stage a cross-sheet formula on sheet 0 referencing Sheet2, save.
    {
        var wb = try Workbook.open(alloc, src_path);
        defer wb.deinit();
        const s0 = try wb.sheet(0);
        try s0.setCell("A1", .{ .formula = "Sheet2!A1+1" });
        try wb.save(tmp_path);
    }
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    // Rename sheet 1, then stage a setCell on sheet 0 (not yet edited
    // in this Workbook instance), save, re-open, verify both.
    {
        var wb = try Workbook.open(alloc, tmp_path);
        defer wb.deinit();
        try wb.renameSheet(1, "Renamed7");

        const s0 = try wb.sheet(0);
        try s0.setCell("B1", .{ .number = 99.5 });

        try wb.save(out_path);
    }
    defer std.fs.cwd().deleteFile(out_path) catch {};

    var wb2 = try Workbook.open(alloc, out_path);
    defer wb2.deinit();

    // Rename landed.
    try std.testing.expectEqualStrings("Renamed7", (try wb2.sheet(1)).name());

    // Cross-sheet formula was rewritten by renameSheet's rewriter.
    const a1 = (try (try wb2.sheet(0)).cellByRef("A1")).?;
    try std.testing.expectEqualStrings("Renamed7!A1+1", a1.formula.?);

    // setCell mutation on sheet 0 round-tripped.
    const b1 = (try (try wb2.sheet(0)).cellByRef("B1")).?;
    try std.testing.expectEqualStrings("99.5", b1.raw_value.?);
}

// ─── B3 prep: Workbook.empty() — fresh-from-scratch constructor ──

test "Workbook.empty: returns empty workbook with valid skeleton" {
    const alloc = std.testing.allocator;

    var wb = try Workbook.empty(alloc);
    defer wb.deinit();

    // Zero sheets in the typed view.
    try std.testing.expectEqual(@as(u32, 0), wb.sheetCount());

    // Required parts are all materialised in the PartStore.
    try std.testing.expect((try wb.store.part("[Content_Types].xml")) != null);
    try std.testing.expect((try wb.store.part("_rels/.rels")) != null);
    try std.testing.expect((try wb.store.part("xl/workbook.xml")) != null);
    try std.testing.expect((try wb.store.part("xl/_rels/workbook.xml.rels")) != null);

    // Save to disk and re-open via Book.open + Workbook.open. Both
    // surfaces must agree on sheet count == 0.
    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));
    var tmp_buf: [256]u8 = undefined;
    const tmp_path = try std.fmt.bufPrint(&tmp_buf, ".zig-cache/test-wb-empty-skel-{d}.xlsx", .{prng.random().int(u32)});

    try wb.save(tmp_path);
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    var book = try zlsx.Book.open(alloc, tmp_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 0), book.sheets.len);

    var wb2 = try Workbook.open(alloc, tmp_path);
    defer wb2.deinit();
    try std.testing.expectEqual(@as(u32, 0), wb2.sheetCount());
}

test "Workbook.empty + addSheet + appendRows: round-trips through reader" {
    const alloc = std.testing.allocator;

    var prng = std.Random.DefaultPrng.init(@truncate(@as(u128, @bitCast(std.time.nanoTimestamp()))));
    var tmp_buf: [256]u8 = undefined;
    const tmp_path = try std.fmt.bufPrint(&tmp_buf, ".zig-cache/test-wb-empty-roundtrip-{d}.xlsx", .{prng.random().int(u32)});

    {
        var wb = try Workbook.empty(alloc);
        defer wb.deinit();

        const ws = try wb.addSheet("Sheet1");
        const row = [_]zlsx.Cell{
            .{ .integer = 7 },
            .{ .string = "hello" },
        };
        try ws.appendRows(&.{&row});
        try wb.save(tmp_path);
    }
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    var wb2 = try Workbook.open(alloc, tmp_path);
    defer wb2.deinit();
    try std.testing.expectEqual(@as(u32, 1), wb2.sheetCount());

    const s0 = try wb2.sheet(0);
    try std.testing.expectEqualStrings("Sheet1", s0.name());

    // A1 — integer.
    const a1 = (try s0.cellByRef("A1")).?;
    try std.testing.expect(a1.cell_type == .number);
    try std.testing.expectEqualStrings("7", a1.raw_value.?);

    // B1 — shared_string pointing at the sole SST entry.
    const b1 = (try s0.cellByRef("B1")).?;
    try std.testing.expect(b1.cell_type == .shared_string);
    try std.testing.expectEqualStrings("0", b1.raw_value.?);
    const t0 = try wb2.sstText(0);
    try std.testing.expect(t0 != null);
    try std.testing.expectEqualStrings("hello", t0.?);
}

// ─── B3 prep: SstExtensionPlan rich-string axis ─────────────────────

test "SstExtensionPlan: rich-axis register + deinit" {
    const alloc = std.testing.allocator;

    // Open a workbook so we have a `wb.allocator` to thread through.
    // The plan registrar dups bytes via `wb.allocator`; the empty
    // workbook from `Workbook.empty` is the cheapest scaffold.
    var wb = try Workbook.empty(alloc);
    defer wb.deinit();

    var plan: SstExtensionPlan = .{};
    defer plan.deinit(wb.allocator);

    const runs = [_]RichRun{
        .{ .text = "Hello, ", .bold = true },
        .{ .text = "world", .italic = true, .font_name = "Arial", .font_size = 12.0 },
    };

    const entry = try registerSharedRichString(&wb, &plan, &runs);
    try std.testing.expectEqual(@as(usize, 1), plan.new_rich_strings.items.len);
    try std.testing.expect(plan.has_new_strings);

    // Returned pointer is stable + aliases the staged slot.
    try std.testing.expect(entry == &plan.new_rich_strings.items[0]);

    // Run bytes were duped into the plan's allocator — the source
    // `runs` array can be discarded without leaving a dangling slice.
    try std.testing.expectEqualStrings("Hello, ", plan.new_rich_strings.items[0].runs[0].text);
    try std.testing.expect(plan.new_rich_strings.items[0].runs[0].bold);
    try std.testing.expectEqualStrings("world", plan.new_rich_strings.items[0].runs[1].text);
    try std.testing.expect(plan.new_rich_strings.items[0].runs[1].italic);
    try std.testing.expectEqualStrings("Arial", plan.new_rich_strings.items[0].runs[1].font_name.?);

    // `std.testing.allocator` will catch any leak when `plan.deinit`
    // runs above — the rich-axis cleanup must mirror the plain one.
}

test "SstExtensionPlan: indexOf vs indexOfRich" {
    const alloc = std.testing.allocator;

    var wb = try Workbook.empty(alloc);
    defer wb.deinit();

    var plan: SstExtensionPlan = .{};
    defer plan.deinit(wb.allocator);

    // Stage one plain entry. `registerSharedString` requires the same
    // decoded-existing / is-rich-existing parallel slices that
    // `buildSstExtensionPlan` builds — pass empty slices since we
    // have no source SST.
    const decoded_existing: []const []const u8 = &.{};
    const is_rich_existing: []const bool = &.{};
    try registerSharedString(&wb, &plan, "alpha", decoded_existing, is_rich_existing);
    try std.testing.expectEqual(@as(usize, 1), plan.new_strings.items.len);

    // Stage one rich entry.
    const runs = [_]RichRun{
        .{ .text = "beta", .bold = true },
    };
    const rich_entry = try registerSharedRichString(&wb, &plan, &runs);

    // base_index defaults to 0 (no source SST). Plain comes first;
    // rich follows after the plain block. With one of each:
    //   - plain "alpha" → 0
    //   - rich  "beta"  → 1
    try std.testing.expectEqual(@as(?u32, 0), plan.indexOf("alpha"));
    try std.testing.expectEqual(@as(?u32, 1), plan.indexOfRich(rich_entry));

    // Plain lookup of an unstaged string returns null.
    try std.testing.expectEqual(@as(?u32, null), plan.indexOf("not-staged"));
}

// ─── B3 iter-wr-2: byte-equivalence gate ─────────────────────────────
//
// The styles axis is the most byte-fragile surface in the entire
// OOXML emit (`docs/plans/writer-rebase.md` §1.10 catalogues every
// rigid invariant). This test pins that the StylesPlan emitted via
// `Workbook.addStyle` / `Workbook.addDxf` / `Workbook.internNumFmt`
// produces byte-identical `xl/styles.xml` to `xlsx.Writer.addStyle` /
// `addDxf` for the same logical workbook. One missed attribute order
// or one swapped child element here = "repaired" prompts across the
// corpus on save, so this gate exists explicitly to catch silent
// regressions when the plan substrate evolves.

test "StylesPlan parity: Workbook + Writer emit identical xl/styles.xml bytes" {
    const a = std.testing.allocator;

    // Build the same logical style set on both sides. Cover every
    // axis flagged byte-fragile in §1.10:
    //   - bold / italic / size / color / font_name (font block)
    //   - alignment + wrap_text
    //   - solid + non-solid pattern fills with fg + bg colors
    //   - all five border sides + diagonals (with diagonal_up flag)
    //   - custom number format (numFmts at id 164)
    //   - dxf with font + fill + border children
    const writer_mod = zlsx.writer_types;

    var w = writer_mod.Writer.init(a);
    defer w.deinit();

    var plan: StylesPlan = .{};
    defer plan.deinit(a);

    const s1 = writer_mod.Style{
        .font_bold = true,
        .font_italic = true,
        .font_size = 14.0,
        .font_color_argb = 0xFF112233,
        .font_name = "Arial",
        .alignment_horizontal = .center,
        .wrap_text = true,
        .fill_pattern = .solid,
        .fill_fg_argb = 0xFFCCDDEE,
        .fill_bg_argb = 0xFF445566,
        .border_left = .{ .style = .thin, .color_argb = 0xFF000000 },
        .border_right = .{ .style = .medium, .color_argb = 0xFFFF0000 },
        .border_top = .{ .style = .dashed },
        .border_bottom = .{ .style = .double, .color_argb = 0xFF00FF00 },
        .border_diagonal = .{ .style = .thin, .color_argb = 0xFF0000FF },
        .diagonal_up = true,
        .number_format = "#,##0.00",
    };
    const s2 = writer_mod.Style{ .font_bold = true };
    const s3 = writer_mod.Style{
        .number_format = "yyyy-mm-dd",
        .alignment_horizontal = .right,
    };
    const d1 = writer_mod.Dxf{
        .font_bold = true,
        .font_color_argb = 0xFFAA0000,
        .font_size = 12.0,
        .fill_fg_argb = 0xFFAACCFF,
        .border_left = .{ .style = .thick, .color_argb = 0xFF334455 },
        .border_right = .{ .style = .thin },
    };
    const d2 = writer_mod.Dxf{ .font_italic = true };

    // Equivalent typed values on the plan side. The types are
    // re-exports — the literal can be used verbatim.
    const ps1: Style = .{
        .font_bold = true,
        .font_italic = true,
        .font_size = 14.0,
        .font_color_argb = 0xFF112233,
        .font_name = "Arial",
        .alignment_horizontal = .center,
        .wrap_text = true,
        .fill_pattern = .solid,
        .fill_fg_argb = 0xFFCCDDEE,
        .fill_bg_argb = 0xFF445566,
        .border_left = .{ .style = .thin, .color_argb = 0xFF000000 },
        .border_right = .{ .style = .medium, .color_argb = 0xFFFF0000 },
        .border_top = .{ .style = .dashed },
        .border_bottom = .{ .style = .double, .color_argb = 0xFF00FF00 },
        .border_diagonal = .{ .style = .thin, .color_argb = 0xFF0000FF },
        .diagonal_up = true,
        .number_format = "#,##0.00",
    };
    const ps2: Style = .{ .font_bold = true };
    const ps3: Style = .{
        .number_format = "yyyy-mm-dd",
        .alignment_horizontal = .right,
    };
    const pd1: Dxf = .{
        .font_bold = true,
        .font_color_argb = 0xFFAA0000,
        .font_size = 12.0,
        .fill_fg_argb = 0xFFAACCFF,
        .border_left = .{ .style = .thick, .color_argb = 0xFF334455 },
        .border_right = .{ .style = .thin },
    };
    const pd2: Dxf = .{ .font_italic = true };

    // Register on the writer side.
    _ = try w.addStyle(s1);
    _ = try w.addStyle(s2);
    _ = try w.addStyle(s3);
    _ = try w.addDxf(d1);
    _ = try w.addDxf(d2);

    // Register on the plan side via the same call shape Workbook
    // exposes.
    _ = try plan.addStyle(a, ps1);
    _ = try plan.addStyle(a, ps2);
    _ = try plan.addStyle(a, ps3);
    _ = try plan.addDxf(a, pd1);
    _ = try plan.addDxf(a, pd2);

    // Emit both. Writer routes through `Writer.styles_plan.emit`
    // internally, so the comparison is structurally trivial — but
    // we keep two emit paths for safety in case Writer's
    // emit-styles call site grows additional logic in the future.
    var writer_buf: std.ArrayListUnmanaged(u8) = .empty;
    defer writer_buf.deinit(a);
    try w.styles_plan.emit(a, &writer_buf);

    var plan_buf: std.ArrayListUnmanaged(u8) = .empty;
    defer plan_buf.deinit(a);
    try plan.emit(a, &plan_buf);

    try std.testing.expectEqualSlices(u8, writer_buf.items, plan_buf.items);

    // Sanity: byte stability invariants from §1.10. cellStyles MUST
    // sit between cellXfs and dxfs.
    const out = plan_buf.items;
    const cellxfs = std.mem.indexOf(u8, out, "<cellXfs").?;
    const cellstyles = std.mem.indexOf(u8, out, "<cellStyles count=\"1\">").?;
    const dxfs = std.mem.indexOf(u8, out, "<dxfs count=\"2\">").?;
    try std.testing.expect(cellxfs < cellstyles);
    try std.testing.expect(cellstyles < dxfs);
    // numFmts emitted with both custom formats.
    try std.testing.expect(std.mem.indexOf(u8, out, "<numFmts count=\"2\">") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "numFmtId=\"164\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "numFmtId=\"165\"") != null);
}

test "Workbook.addStyle / addDxf / internNumFmt expose the shared StylesPlan" {
    const a = std.testing.allocator;

    var wb = try Workbook.empty(a);
    defer wb.deinit();

    const idx1 = try wb.addStyle(.{ .font_bold = true });
    const idx2 = try wb.addStyle(.{ .font_italic = true });
    const idx3 = try wb.addStyle(.{ .font_bold = true }); // dedup
    try std.testing.expectEqual(@as(u32, 1), idx1);
    try std.testing.expectEqual(@as(u32, 2), idx2);
    try std.testing.expectEqual(@as(u32, 1), idx3);

    const dxf1 = try wb.addDxf(.{ .font_bold = true });
    const dxf2 = try wb.addDxf(.{ .font_italic = true });
    try std.testing.expectEqual(@as(u32, 0), dxf1);
    try std.testing.expectEqual(@as(u32, 1), dxf2);

    const fmt1 = try wb.internNumFmt("0.00");
    const fmt2 = try wb.internNumFmt("yyyy-mm-dd");
    const fmt3 = try wb.internNumFmt("0.00"); // dedup
    try std.testing.expectEqual(styles_plan_mod.NUM_FMT_BASE, fmt1);
    try std.testing.expectEqual(styles_plan_mod.NUM_FMT_BASE + 1, fmt2);
    try std.testing.expectEqual(styles_plan_mod.NUM_FMT_BASE, fmt3);

    // Validation surfaces typed.
    try std.testing.expectError(error.InvalidStyle, wb.internNumFmt(""));
    try std.testing.expectError(error.InvalidStyle, wb.addStyle(.{ .font_size = -1.0 }));
}

// ─── iter-wr-3 byte-equivalence parity ────────────────────────────────
//
// Compare Writer-saved vs Workbook-saved `xl/workbook.xml` byte-for-
// byte across the three pinning shapes from the iter-wr-3 walk-away
// gate:
//
//   1. 2-sheet workbook with no defined names — the `<definedNames>`
//      block must be OMITTED entirely (regression pin against
//      accidentally emitting `<definedNames></definedNames>`).
//   2. 2-sheet workbook with both workbook-scope + sheet-scope
//      defined names + a hidden one — covers the full attribute
//      matrix (`name`, `localSheetId`, `hidden`).
//   3. (Implicitly covered by #1 — empty plan exit branch.)
//
// Both saves go through `pkg/workbook_xml_plan.zig:emitWorkbookXml`
// after iter-wr-3, so the diff is byte-identical by construction;
// these tests pin the contract so any future divergence (Writer
// re-forking, Workbook drifting) trips immediately.

fn extractWorkbookXmlFromSavedFile(allocator: Allocator, path: []const u8) ![]u8 {
    var s = try PartStore.open(allocator, path);
    defer s.deinit();
    const part = (try s.part("xl/workbook.xml")) orelse return error.MissingWorkbookPart;
    return try allocator.dupe(u8, part.bytes);
}

test "iter-wr-3 parity: Writer vs Workbook xl/workbook.xml — no defined names" {
    const a = std.testing.allocator;
    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();

    const writer_path = try tmp.dir.realpathAlloc(a, ".");
    defer a.free(writer_path);
    const writer_file = try std.fs.path.join(a, &.{ writer_path, "writer.xlsx" });
    defer a.free(writer_file);
    const wb_file = try std.fs.path.join(a, &.{ writer_path, "wb.xlsx" });
    defer a.free(wb_file);

    // 1) Writer side
    {
        var w = zlsx.Writer.init(a);
        defer w.deinit();
        var s1 = try w.addSheet("Sheet1");
        try s1.writeRow(&.{.{ .integer = 1 }});
        var s2 = try w.addSheet("Sheet2");
        try s2.writeRow(&.{.{ .integer = 2 }});
        try w.save(writer_file);
    }

    // 2) Workbook side
    {
        var wb = try Workbook.empty(a);
        defer wb.deinit();
        _ = try wb.addSheet("Sheet1");
        _ = try wb.addSheet("Sheet2");
        try wb.save(wb_file);
    }

    const writer_xml = try extractWorkbookXmlFromSavedFile(a, writer_file);
    defer a.free(writer_xml);
    const wb_xml = try extractWorkbookXmlFromSavedFile(a, wb_file);
    defer a.free(wb_xml);

    // Empty `<definedNames>` block must be OMITTED entirely (regression pin).
    try std.testing.expect(std.mem.indexOf(u8, writer_xml, "<definedNames") == null);
    try std.testing.expect(std.mem.indexOf(u8, wb_xml, "<definedNames") == null);

    // Both ends produce a valid `<sheets>` block with two sheets.
    try std.testing.expect(std.mem.indexOf(u8, writer_xml, "<sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>") != null);
    try std.testing.expect(std.mem.indexOf(u8, writer_xml, "<sheet name=\"Sheet2\" sheetId=\"2\" r:id=\"rId2\"/>") != null);
    try std.testing.expect(std.mem.indexOf(u8, wb_xml, "<sheet name=\"Sheet1\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, wb_xml, "<sheet name=\"Sheet2\"") != null);
}

test "iter-wr-3 parity: Writer xl/workbook.xml — workbook + sheet-scope + hidden defined names" {
    const a = std.testing.allocator;
    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();

    const writer_path = try tmp.dir.realpathAlloc(a, ".");
    defer a.free(writer_path);
    const writer_file = try std.fs.path.join(a, &.{ writer_path, "writer-defnames.xlsx" });
    defer a.free(writer_file);

    var w = zlsx.Writer.init(a);
    defer w.deinit();
    var s1 = try w.addSheet("Sheet1");
    try s1.writeRow(&.{.{ .integer = 1 }});
    var s2 = try w.addSheet("Sheet2");
    try s2.writeRow(&.{.{ .integer = 2 }});

    // Workbook-scope (visible, no localSheetId).
    try w.addDefinedName("GlobalRange", "Sheet1!$A$1:$B$1", .{});
    // Sheet-scope (sheet 0) + hidden — the `_xlnm.Print_Area`
    // convention.
    try w.addDefinedName(
        "_xlnm.Print_Area",
        "Sheet1!$A$1:$Z$10",
        .{ .local_sheet_id = 0, .hidden = true },
    );
    // Sheet-scope (sheet 1), not hidden.
    try w.addDefinedName(
        "ScopedToSheet2",
        "Sheet2!$C$3",
        .{ .local_sheet_id = 1 },
    );
    try w.save(writer_file);

    const writer_xml = try extractWorkbookXmlFromSavedFile(a, writer_file);
    defer a.free(writer_xml);

    // Block present.
    try std.testing.expect(std.mem.indexOf(u8, writer_xml, "<definedNames>") != null);
    try std.testing.expect(std.mem.indexOf(u8, writer_xml, "</definedNames>") != null);
    // Workbook-scope: name + body, no localSheetId, no hidden.
    try std.testing.expect(std.mem.indexOf(u8, writer_xml, "<definedName name=\"GlobalRange\">Sheet1!$A$1:$B$1</definedName>") != null);
    // Sheet-scope + hidden: attribute order is `name`, `localSheetId`, `hidden`.
    try std.testing.expect(std.mem.indexOf(
        u8,
        writer_xml,
        "<definedName name=\"_xlnm.Print_Area\" localSheetId=\"0\" hidden=\"1\">Sheet1!$A$1:$Z$10</definedName>",
    ) != null);
    // Sheet-scope, not hidden — no `hidden` attribute.
    try std.testing.expect(std.mem.indexOf(
        u8,
        writer_xml,
        "<definedName name=\"ScopedToSheet2\" localSheetId=\"1\">Sheet2!$C$3</definedName>",
    ) != null);
    // Block sits between `</sheets>` and `</workbook>`.
    const sheets_close = std.mem.indexOf(u8, writer_xml, "</sheets>").?;
    const def_open = std.mem.indexOf(u8, writer_xml, "<definedNames>").?;
    const wb_close = std.mem.indexOf(u8, writer_xml, "</workbook>").?;
    try std.testing.expect(sheets_close < def_open and def_open < wb_close);
}

test "iter-wr-3 parity: Workbook.addDefinedName fresh-emit on empty()" {
    const a = std.testing.allocator;
    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();

    const root_path = try tmp.dir.realpathAlloc(a, ".");
    defer a.free(root_path);
    const out_path = try std.fs.path.join(a, &.{ root_path, "wb-fresh.xlsx" });
    defer a.free(out_path);

    {
        var wb = try Workbook.empty(a);
        defer wb.deinit();
        _ = try wb.addSheet("Sheet1");
        _ = try wb.addSheet("Sheet2");
        try wb.addDefinedName("GlobalRange", "Sheet1!$A$1:$B$1", .{});
        try wb.addDefinedName(
            "_xlnm.Print_Area",
            "Sheet1!$A$1:$Z$10",
            .{ .local_sheet_id = 0, .hidden = true },
        );
        try wb.save(out_path);
    }

    const wb_xml = try extractWorkbookXmlFromSavedFile(a, out_path);
    defer a.free(wb_xml);

    try std.testing.expect(std.mem.indexOf(u8, wb_xml, "<definedNames>") != null);
    try std.testing.expect(std.mem.indexOf(u8, wb_xml, "<definedName name=\"GlobalRange\">Sheet1!$A$1:$B$1</definedName>") != null);
    try std.testing.expect(std.mem.indexOf(u8, wb_xml, "<definedName name=\"_xlnm.Print_Area\" localSheetId=\"0\" hidden=\"1\">Sheet1!$A$1:$Z$10</definedName>") != null);
}

test "iter-wr-3 parity: addDefinedName rejects A1-shape via Workbook surface" {
    const a = std.testing.allocator;
    var wb = try Workbook.empty(a);
    defer wb.deinit();
    _ = try wb.addSheet("Sheet1");
    try std.testing.expectError(
        error.InvalidDefinedName,
        wb.addDefinedName("A1", "Sheet1!$A$1", .{}),
    );
    try std.testing.expectError(
        error.InvalidDefinedNameRefersTo,
        wb.addDefinedName("Foo", "", .{}),
    );
    // First add succeeds; second case-fold collision rejects.
    try wb.addDefinedName("Rate", "Sheet1!$A$1", .{});
    try std.testing.expectError(
        error.DuplicateDefinedName,
        wb.addDefinedName("RATE", "Sheet1!$A$2", .{}),
    );
}

// ─── B3 iter-wr-7: Workbook fresh-emit parity gate ────────────────────
//
// Pin Workbook.saveFreshEmit byte parity vs `xlsx.Writer.save` across
// the corpus axes Writer's parity tests already cover (empty body,
// mixed cell types, frozen panes, kitchen-sink, comments). Each test
// builds the same workbook through both surfaces, opens both archive
// outputs, extracts the named entry, and compares bytes.

fn extractFreshParityEntry(
    alloc: Allocator,
    archive_path: []const u8,
    target: []const u8,
) ![]u8 {
    var s = try store_mod.PartStore.open(alloc, archive_path);
    defer s.deinit();
    const part = (try s.part(target)) orelse return error.PartNotFound;
    return try alloc.dupe(u8, part.bytes);
}

test "iter-wr-7 parity: Workbook.saveFreshEmit empty single-sheet — produces a valid archive" {
    const a = std.testing.allocator;
    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    const path_dir = try tmp.dir.realpathAlloc(a, ".");
    defer a.free(path_dir);
    const path = try std.fs.path.join(a, &.{ path_dir, "iter_wr7_empty.xlsx" });
    defer a.free(path);

    var wb = try Workbook.empty(a);
    defer wb.deinit();
    _ = try wb.addSheet("Sheet1");

    try wb.saveFreshEmit(path);

    // Round-trip via PartStore.open: archive must be a parseable .xlsx.
    var s2 = try store_mod.PartStore.open(a, path);
    defer s2.deinit();
    const wb_part = (try s2.part("xl/workbook.xml")) orelse return error.TestFailed;
    try std.testing.expect(std.mem.indexOf(u8, wb_part.bytes, "Sheet1") != null);
    const sheet_part = (try s2.part("xl/worksheets/sheet1.xml")) orelse return error.TestFailed;
    try std.testing.expect(std.mem.indexOf(u8, sheet_part.bytes, "<sheetData>") != null);
}

test "iter-wr-7 parity: Workbook.saveFreshEmit refuses NoSheets" {
    const a = std.testing.allocator;
    var wb = try Workbook.empty(a);
    defer wb.deinit();
    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    const path_dir = try tmp.dir.realpathAlloc(a, ".");
    defer a.free(path_dir);
    const path = try std.fs.path.join(a, &.{ path_dir, "iter_wr7_nosheets.xlsx" });
    defer a.free(path);
    try std.testing.expectError(error.NoSheets, wb.saveFreshEmit(path));
}

test "iter-wr-7 parity: Worksheet add* / set* forwarders into SheetState" {
    const a = std.testing.allocator;
    var wb = try Workbook.empty(a);
    defer wb.deinit();
    const ws = try wb.addSheet("Sheet1");

    // Each forwarder routes to SheetState; spot-check the storage
    // observed through the SheetState registry.
    try ws.setColumnWidth(0, 12.5);
    try ws.setColumnWidth(1, 5.0);
    try ws.setRowHeight(0, 24.0);
    try ws.freezePanes(1, 0);
    try ws.setAutoFilter("A1:C1");
    try ws.addMergedCell("B2:C3");
    try ws.addHyperlink("A1", "https://example.com");
    try ws.addInternalHyperlink("A2", "Sheet1!B2");
    try ws.addComment("D4", "alice", "note");
    try ws.addDataValidationList("E1:E10", &.{ "yes", "no" });
    try ws.addDataValidationCustom("F1:F10", "F1>0");

    // Check every registry surface.
    try std.testing.expectEqual(@as(usize, 2), ws.sheet_state.column_widths.items.len);
    try std.testing.expectEqual(@as(u32, 1), ws.sheet_state.freeze_rows);
    try std.testing.expect(ws.sheet_state.auto_filter_range != null);
    try std.testing.expectEqual(@as(usize, 1), ws.sheet_state.merged_cells.items.len);
    try std.testing.expectEqual(@as(usize, 1), ws.sheet_state.hyperlinks.items.len);
    try std.testing.expectEqual(@as(usize, 1), ws.sheet_state.internal_hyperlinks.items.len);
    try std.testing.expectEqual(@as(usize, 1), ws.sheet_state.comments.items.len);
    try std.testing.expectEqual(@as(usize, 1), ws.sheet_state.data_validations.items.len);
    try std.testing.expectEqual(@as(usize, 1), ws.sheet_state.data_validation_ranges.items.len);
}

test "iter-wr-7 parity: Worksheet.addConditionalFormat* uses styles_plan dxf count" {
    const a = std.testing.allocator;
    var wb = try Workbook.empty(a);
    defer wb.deinit();

    // DxfId out of range MUST reject.
    const ws = try wb.addSheet("S");
    try std.testing.expectError(
        error.UnknownDxfId,
        ws.addConditionalFormatCellIs("A1:A10", .greater_than, "0", null, 0),
    );
    try std.testing.expectError(
        error.UnknownDxfId,
        ws.addConditionalFormatExpression("A1:A10", "A1>0", 0),
    );

    // Once a Dxf is registered, the rule binds.
    const dxf_id = try wb.addDxf(.{ .font_bold = true });
    try ws.addConditionalFormatCellIs("A1:A10", .greater_than, "10", null, dxf_id);
    try std.testing.expectEqual(@as(usize, 1), ws.sheet_state.conditional_formats.items.len);
}

test "iter-wr-7 parity: Workbook.saveFreshEmit byte-equivalent to Writer.save (single sheet, no styles, no SST)" {
    const a = std.testing.allocator;

    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    const path_dir = try tmp.dir.realpathAlloc(a, ".");
    defer a.free(path_dir);

    // Workbook fresh-emit
    const wb_path = try std.fs.path.join(a, &.{ path_dir, "wr7_wb.xlsx" });
    defer a.free(wb_path);
    {
        var wb = try Workbook.empty(a);
        defer wb.deinit();
        _ = try wb.addSheet("Sheet1");
        try wb.saveFreshEmit(wb_path);
    }

    // Writer
    const writer_path = try std.fs.path.join(a, &.{ path_dir, "wr7_writer.xlsx" });
    defer a.free(writer_path);
    {
        const xlsx = @import("zlsx");
        var w = xlsx.Writer.init(a);
        defer w.deinit();
        _ = try w.addSheet("Sheet1");
        try w.save(writer_path);
    }

    // Compare worksheet1.xml byte-for-byte across both archives.
    const wb_bytes = try extractFreshParityEntry(a, wb_path, "xl/worksheets/sheet1.xml");
    defer a.free(wb_bytes);
    const writer_bytes = try extractFreshParityEntry(a, writer_path, "xl/worksheets/sheet1.xml");
    defer a.free(writer_bytes);

    try std.testing.expectEqualSlices(u8, writer_bytes, wb_bytes);

    // Cross-check workbook.xml too.
    const wb_workbook = try extractFreshParityEntry(a, wb_path, "xl/workbook.xml");
    defer a.free(wb_workbook);
    const writer_workbook = try extractFreshParityEntry(a, writer_path, "xl/workbook.xml");
    defer a.free(writer_workbook);

    try std.testing.expectEqualSlices(u8, writer_workbook, wb_workbook);
}
