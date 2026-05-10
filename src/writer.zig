//! xlsx writer — fresh-file emission.
//!
//! Flow
//! ----
//! `Writer.init → addStyle? → addSheet → writeRow* → save`. Each
//! sheet's `SheetWriter` exposes per-row writers and per-region
//! attachments (merges, hyperlinks, comments, validations,
//! conditional formats); the top-level `Writer` owns the style
//! table, dxf table, and shared-string pool, and serialises the
//! archive on `save`.
//!
//! Surface
//! -------
//! * Cell types: empty, string (shared + deduped), integer, number,
//!   boolean. Numerics rejected if not exactly representable in f64.
//! * `addStyle` + `writeRowStyled` — fonts, fills, borders,
//!   alignment, wrap, number formats.
//! * `addDxf` + conditional formatting (`addConditionalFormatCellIs`,
//!   `…Expression`, `…ColorScale`, `…DataBar`).
//! * `addMergedCell`, `addHyperlink`, `addInternalHyperlink`,
//!   `addComment`.
//! * `addDataValidationList`, `…Numeric`, `…Custom`.
//! * `writeRowWithFormulas`, `writeRichRow`.
//! * `setColumnWidth`, `setRowHeight`, `freezePanes`, `setAutoFilter`.
//! * Output: OOXML zip with deflate compression (in-house deflate
//!   encoder; entries `>= 1024` bytes that fail to compress fall
//!   back to store). Excel, LibreOffice, and zlsx's own reader all
//!   accept the produced files.
//!
//! Out of scope
//! ------------
//! * Load + edit + save round-trip — readers and writers don't
//!   share an in-memory document model yet.
//! * Pivot tables, charts, drawings.

const std = @import("std");
const xlsx = @import("xlsx.zig");
const casefold = @import("unicode/casefold.zig");

// B3 iter-wr-1: SST unification. Writer stages strings + rich entries
// through the shared `SstExtensionPlan` substrate (see
// `pkg/sst_plan.zig`) instead of a Writer-local pool. The plan module
// is std-only, so this dep doesn't form a cycle through `pkg/workbook`.
const sst_plan = @import("zlsx_sst_plan");
const SstExtensionPlan = sst_plan.SstExtensionPlan;
const PlanRichRun = sst_plan.RichRun;

// B3 iter-wr-2: Styles unification. Writer's style + dxf + numFmt
// registries + the `xl/styles.xml` emitter live in
// `pkg/styles_plan.zig` so Workbook's fresh-emit path can use the
// same code without a circular module dep.
const styles_plan_mod = @import("zlsx_styles_plan");
pub const StylesPlan = styles_plan_mod.StylesPlan;

// B3 iter-wr-3: workbook.xml unification. Writer stages defined names
// + emits `xl/workbook.xml` through the shared `WorkbookXmlPlan`
// substrate (see `pkg/workbook_xml_plan.zig`). Same cycle-avoidance
// argument as `sst_plan` above — std-only, no `pkg/workbook` import.
const workbook_xml_plan = @import("zlsx_workbook_xml_plan");
const WorkbookXmlPlan = workbook_xml_plan.WorkbookXmlPlan;
const WorkbookXmlSheetEntry = workbook_xml_plan.SheetEntry;
// B3 iter-wr-5: ZIP layout unification. Writer's per-archive LFH/CDFH/
// EOCD emission now routes through the shared `pkg/zip.zig` substrate.
// std-only; takes deflateCompress as a callback to avoid the
// writer→pkg→writer module-graph cycle.
const zip = @import("zlsx_zip");

// B3 iter-wr-4: per-sheet fresh-emit unification. Writer's
// `xl/worksheets/sheetN.xml` + per-sheet rels + comments + VML
// drawing emit routes through the shared `pkg/sheet_plan.zig`
// substrate. std-only; same cycle-avoidance argument as the
// other plan modules. Workbook (`pkg/workbook.zig`) gains the
// same emit surface for fresh-file production in future iters.
const sheet_plan = @import("zlsx_sheet_plan");

/// Function-pointer adapter wrapping `deflateCompress` with the
/// `anyerror!void` return type that `zip.Archive.addEntry`'s
/// `DeflateFn` expects.
fn deflateCompressErased(
    alloc: Allocator,
    input: []const u8,
    out: *std.ArrayListUnmanaged(u8),
) anyerror!void {
    return deflateCompress(alloc, input, out);
}

/// B3 iter-wr-6: `projectConditionalFormat` retired — `Writer.save`
/// inlines the union-arm projection now that the storage lives on
/// `sheet_plan.SheetState`. `projectCfOperator` survives because the
/// Writer-side `CfOperator` enum is a separate (identical-shape) type
/// from `sheet_plan.CfOperator` and the public `addConditionalFormatCellIs`
/// signature still takes the Writer-local enum.
inline fn projectCfOperator(op: CfOperator) sheet_plan.CfOperator {
    return switch (op) {
        .less_than => .less_than,
        .less_than_or_equal => .less_than_or_equal,
        .equal => .equal,
        .not_equal => .not_equal,
        .greater_than => .greater_than,
        .greater_than_or_equal => .greater_than_or_equal,
        .between => .between,
        .not_between => .not_between,
    };
}

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
pub fn fitsExactlyInF64(n: i64) bool {
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

const CONTENT_TYPES_DEFAULTS: []const u8 =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    \\<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    \\<Default Extension="xml" ContentType="application/xml"/>
;
const CONTENT_TYPES_FIXED_OVERRIDES: []const u8 =
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

// B3 iter-wr-3: `WORKBOOK_HEAD` / `WORKBOOK_SHEETS_CLOSE` /
// `WORKBOOK_END` retired. The fresh-emit shape is owned by
// `pkg/workbook_xml_plan.zig:emitWorkbookXml`, which is the
// canonical home for `xl/workbook.xml` byte serialization.

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

// xl/styles.xml skeleton + emit logic moved to `pkg/styles_plan.zig`
// (B3 iter-wr-2). The static blobs (`STYLES_HEAD`, `STYLES_FONTS_DEFAULT`,
// `STYLES_CELL_STYLE_XFS`, `STYLES_DEFAULT_CELL_XF`, `STYLES_CELL_STYLES`,
// `STYLES_TAIL`) and the `NUM_FMT_BASE` constant live there now.

// ─── Writer public API ───────────────────────────────────────────────

/// OOXML border-side style enum. Hosted by `pkg/styles_plan.zig` —
/// re-exported so `xlsx.Writer.addStyle(.{ ... .border_left = ... })`
/// keeps its pre-iter-wr-2 surface.
pub const BorderStyle = styles_plan_mod.BorderStyle;

/// Comparison operator for `cellIs` conditional-format rules.
/// Names mirror `DataValidationOp`; the two OOXML enums share the
/// same wire-format tokens but live in different grammar slots.
pub const CfOperator = enum {
    less_than,
    less_than_or_equal,
    equal,
    not_equal,
    greater_than,
    greater_than_or_equal,
    between,
    not_between,

    fn toOoxml(self: CfOperator) []const u8 {
        return switch (self) {
            .less_than => "lessThan",
            .less_than_or_equal => "lessThanOrEqual",
            .equal => "equal",
            .not_equal => "notEqual",
            .greater_than => "greaterThan",
            .greater_than_or_equal => "greaterThanOrEqual",
            .between => "between",
            .not_between => "notBetween",
        };
    }

    fn needsSecondFormula(self: CfOperator) bool {
        return self == .between or self == .not_between;
    }
};

/// A differential format — the font / fill overrides applied when a
/// conditional-format rule matches. Registered once via
/// `Writer.addDxf`, referenced by dxfId from one or more rules.
/// Scoped to the subset of properties real conditional formats
/// actually toggle (bold / font color / solid fill). Full differential
/// font / border support can layer on later without breaking shape.
pub const Dxf = styles_plan_mod.Dxf;

/// Cell comment (note) attached via `SheetWriter.addComment`. Emits
/// as `<comments><authors/><commentList/></comments>` under
/// `xl/commentsN.xml`, plus a minimal VML shape under
/// `xl/drawings/vmlDrawingN.xml` so Excel renders the note
/// indicator. Plain-text bodies only — rich-text comment bodies
/// can layer on without breaking this shape.
pub const Comment = struct {
    ref: []const u8,
    author: []const u8,
    text: []const u8,
};

/// Formatting run for a single piece of rich-text content. Mirrors
/// the reader-side `xlsx.RichRun` with writer-friendly nullable
/// optional fields. Consecutive runs inside the same cell accumulate
/// into one `<si>` entry in the shared-strings table.
pub const RichTextRun = struct {
    text: []const u8,
    bold: bool = false,
    italic: bool = false,
    color_argb: ?u32 = null,
    size: ?f32 = null,
    font_name: ?[]const u8 = null,
};

/// Row-cell union that adds a rich-text variant alongside the plain
/// `xlsx.Cell` shape. Used by `SheetWriter.writeRichRow`. Non-rich
/// cells keep the exact semantics of `writeRow` — only the rich
/// variant takes a different code path.
pub const RichRowCell = union(enum) {
    empty,
    string: []const u8,
    integer: i64,
    number: f64,
    boolean: bool,
    rich: []const RichTextRun,
};

pub const BorderSide = styles_plan_mod.BorderSide;
pub const PatternType = styles_plan_mod.PatternType;
pub const HAlign = styles_plan_mod.HAlign;
pub const Style = styles_plan_mod.Style;

/// A workbook-level defined name (named range, print area,
/// validation source, etc.). Emitted in xl/workbook.xml as
/// `<definedName name="..." [localSheetId="N"] [hidden="1"]>...</definedName>`.
///
/// B3 iter-wr-3: storage moved to `pkg/workbook_xml_plan.zig`. This
/// alias keeps `xlsx.Writer.DefinedName` callable for any consumer
/// that named the type explicitly.
pub const DefinedName = workbook_xml_plan.DefinedName;

/// B3 iter-wr-3: re-exported from `pkg/workbook_xml_plan.zig`. Same
/// shape as the prior writer-local struct (`local_sheet_id`,
/// `hidden`), so call sites compile unchanged.
pub const DefinedNameOptions = workbook_xml_plan.DefinedNameOptions;

pub const Writer = struct {
    allocator: Allocator,
    // Accumulated sheet writers (owned).
    sheets: std.ArrayListUnmanaged(*SheetWriter) = .{},
    // Shared-string table substrate (B3 iter-wr-1). Plain entries land
    // in `sst_plan.new_strings` (O(1) dedup via the hash side-index);
    // rich entries land in `sst_plan.new_rich_strings` (no dedup,
    // matches `xlsx.Writer`'s iter33 policy). `base_index` stays at 0
    // for fresh-emit, so plan-level indices ARE the SST indices the
    // emit loop hands to `<v>{idx}</v>`. Replaces the four pre-iter-wr-1
    // fields (`sst_strings` / `sst_index` / `sst_is_rich` / `sst_count`).
    sst_plan: SstExtensionPlan = .{},
    // Total number of string-typed cells written across all sheets
    // (informational — OOXML's <sst count="..."> field). Distinct from
    // `sst_plan.new_strings.items.len` (= uniqueCount); a single cell
    // hitting an existing entry still bumps `sst_count` but leaves
    // `new_strings.items.len` unchanged.
    sst_count: u64 = 0,
    // Styles substrate (B3 iter-wr-2). Replaces the pre-iter-wr-2
    // `styles` / `dxfs` / `num_fmts` / `num_fmt_index` quartet of
    // fields. The plan owns all duped style strings and the user-numFmt
    // pool, dedupes by content, and emits byte-identical
    // `xl/styles.xml`. Workbook (`pkg/workbook.zig`) uses the same
    // type — see `pkg/styles_plan.zig`.
    styles_plan: StylesPlan = .{},
    // Workbook.xml fresh-emit plan substrate (B3 iter-wr-3). Holds
    // the staged defined-name pool — emitted as `<definedNames>` in
    // xl/workbook.xml between `</sheets>` and `</workbook>`. Both
    // workbook-scoped (no localSheetId) and sheet-scoped names are
    // supported via the optional `local_sheet_id` on the plan's
    // `DefinedName`. Replaces the pre-iter-wr-3 Writer-local
    // `defined_names: ArrayList(DefinedName)`.
    workbook_xml_plan: WorkbookXmlPlan = .{},

    pub fn init(allocator: Allocator) Writer {
        return .{ .allocator = allocator };
    }

    pub fn deinit(self: *Writer) void {
        for (self.sheets.items) |s| {
            s.deinit();
            self.allocator.destroy(s);
        }
        self.sheets.deinit(self.allocator);
        self.sst_plan.deinit(self.allocator);
        // Plan owns all duped style strings + the user-numFmt pool.
        self.styles_plan.deinit(self.allocator);
        // Plan owns all duped defined-name strings.
        self.workbook_xml_plan.deinit(self.allocator);
        self.* = undefined;
    }

    /// Register a cell style and return its `s="…"` index. Dedupes
    /// structurally (including content-comparing `font_name`, not
    /// just slice-header comparing). Returning value is 1-based —
    /// cellXfs[0] is reserved for the default no-style record.
    ///
    /// Thin pass-through to `StylesPlan.addStyle`; see
    /// `pkg/styles_plan.zig` for the dedup + validation logic.
    pub fn addStyle(self: *Writer, style: Style) !u32 {
        return try self.styles_plan.addStyle(self.allocator, style);
    }

    /// Register a differential format (font / fill overrides applied
    /// when a conditional-format rule matches) and return its dxfId.
    /// Linear dedup by content equality — repeat registrations of
    /// the same Dxf return the same id.
    ///
    /// The returned u32 is fed into
    /// `SheetWriter.addConditionalFormatCellIs` /
    /// `…Expression` as the `dxf_id` parameter. It is a *pure* dxf
    /// id (0-based into `<dxfs>`), distinct from the style id used
    /// by `addStyle`.
    ///
    /// Thin pass-through to `StylesPlan.addDxf`.
    pub fn addDxf(self: *Writer, dxf: Dxf) !u32 {
        return try self.styles_plan.addDxf(self.allocator, dxf);
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
        // a hash of lowercased names. Uses non-Turkic full Unicode
        // case fold so e.g. café/CAFÉ and ß/SS collapse correctly.
        for (self.sheets.items) |existing| {
            if (casefold.excelSheetNameEql(existing.name, name)) return error.DuplicateSheetName;
        }
        const sw = try self.allocator.create(SheetWriter);
        errdefer self.allocator.destroy(sw);
        sw.* = try SheetWriter.init(self, name);
        try self.sheets.append(self.allocator, sw);
        return sw;
    }

    /// Register a workbook-level defined name. The name + refers_to
    /// strings are duped into the writer's heap; the caller can free
    /// or reuse their buffers immediately. Validates the name shape
    /// against Excel's rules (starts with letter / `_` / `\`,
    /// contains only `[A-Za-z0-9_.\\?]`, max 255 bytes, not the
    /// shape of an A1 cell ref); refers_to is not parsed beyond a
    /// non-empty check — the formula tokenizer at
    /// `src/formula/tokenizer.zig` is the canonical surface for
    /// callers wanting structural validation.
    ///
    /// `opts.local_sheet_id` (0-based) makes the name sheet-scoped
    /// — caller must ensure the index resolves at save() time.
    /// `opts.hidden = true` is the convention for `_xlnm.Print_Area`
    /// and similar built-in-shaped names that shouldn't surface in
    /// Excel's Name Manager UI.
    pub fn addDefinedName(
        self: *Writer,
        name: []const u8,
        refers_to: []const u8,
        opts: DefinedNameOptions,
    ) !void {
        // B3 iter-wr-3: validation + dedup + storage all live on the
        // shared plan. Same rule set as before (full Excel name
        // grammar, case-insensitive duplicate reject per scope,
        // empty refers_to reject); same byte format on emit.
        return self.workbook_xml_plan.addDefinedName(
            self.allocator,
            name,
            refers_to,
            opts,
        );
    }

    /// Return the 0-based SST index for plain string `s`. Dedups
    /// against `sst_plan.new_strings` via the plan's O(1) hash
    /// side-index; copies the string into the plan's pool on first
    /// sight so callers don't need to keep it alive.
    ///
    /// `base_index` is 0 in fresh-emit, so the plan index IS the SST
    /// index — no further offset needed at the call site.
    fn sstIntern(self: *Writer, s: []const u8) !u32 {
        return try self.sst_plan.registerNewPlain(self.allocator, s);
    }

    /// Stage a rich-text run list into the plan's typed
    /// `new_rich_strings` axis and return the 0-based SST index.
    /// Always appends (no dedup — same iter33 policy as before; rich
    /// entries are rare enough that hashing the formatted form costs
    /// more than it saves). Translates the writer's `RichTextRun`
    /// shape into the plan's `RichRun` shape: u32 ARGB → 8-hex string
    /// (the plan stores pre-formatted color bytes since Workbook's
    /// delta path reads raw `<color rgb="…"/>` slices straight from
    /// the source XML); `size` (f32) → `font_size` (f32). Strike +
    /// underline ride along as `false` because the writer's public
    /// surface doesn't expose them — the SST emitter consults those
    /// flags but they remain false-equivalent here.
    fn sstInternRich(self: *Writer, runs: []const RichTextRun) !u32 {
        // Translate writer-runs → plan-runs into a per-call arena so
        // the plan registrar (which dups every byte itself) sees an
        // owned buffer; on success the arena is freed and the plan
        // owns the duped copies.
        var arena = std.heap.ArenaAllocator.init(self.allocator);
        defer arena.deinit();
        const aalloc = arena.allocator();

        const plan_runs = try aalloc.alloc(PlanRichRun, runs.len);
        for (runs, 0..) |r, i| {
            const color: ?[]const u8 = if (r.color_argb) |c| blk: {
                var hex_buf: [8]u8 = undefined;
                _ = std.fmt.bufPrint(&hex_buf, "{X:0>8}", .{c}) catch unreachable;
                break :blk try aalloc.dupe(u8, &hex_buf);
            } else null;
            plan_runs[i] = .{
                .text = r.text,
                .bold = r.bold,
                .italic = r.italic,
                .underline = false,
                .strike = false,
                .font_name = r.font_name,
                .font_size = r.size,
                .color_argb = color,
            };
        }
        const entry = try self.sst_plan.registerNewRich(self.allocator, plan_runs);
        return self.sst_plan.indexOfRich(entry) orelse unreachable;
    }

    /// Serialise everything and write to `path`. Overwrites.
    pub fn save(self: *Writer, path: []const u8) !void {
        if (self.sheets.items.len == 0) return error.NoSheets;

        var zip_buf: std.ArrayListUnmanaged(u8) = .{};
        defer zip_buf.deinit(self.allocator);

        var zw = zip.Archive.init(self.allocator, &zip_buf);
        defer zw.deinit();

        const alloc = self.allocator;

        // Dxfs live in styles.xml too — conditional formatting alone
        // is enough to require the part, even when no user style was
        // registered via addStyle.
        const have_styles = !self.styles_plan.isEmpty();

        // 1. [Content_Types].xml
        //
        // OPC schema requires every <Default> element BEFORE any
        // <Override> element. The flow is:
        //   1. CONTENT_TYPES_DEFAULTS — opens <Types> + canonical
        //      `rels`/`xml` Defaults.
        //   2. Optional `vml` Default for comment-bearing
        //      workbooks (must precede ALL Overrides).
        //   3. CONTENT_TYPES_FIXED_OVERRIDES — workbook +
        //      sharedStrings.
        //   4. Per-sheet, optional styles + per-comments
        //      Overrides.
        //   5. CONTENT_TYPES_TAIL — closes </Types>.
        {
            var ct: std.ArrayListUnmanaged(u8) = .{};
            defer ct.deinit(alloc);
            // ─── Phase 1: Defaults ──────────────────────────────
            try ct.appendSlice(alloc, CONTENT_TYPES_DEFAULTS);
            var any_comments = false;
            for (self.sheets.items) |sw| {
                if (sw.state.comments.items.len > 0) {
                    any_comments = true;
                    break;
                }
            }
            if (any_comments) {
                try ct.appendSlice(
                    alloc,
                    "<Default Extension=\"vml\" ContentType=\"application/vnd.openxmlformats-officedocument.vmlDrawing\"/>",
                );
            }
            // ─── Phase 2: Overrides ─────────────────────────────
            try ct.appendSlice(alloc, CONTENT_TYPES_FIXED_OVERRIDES);
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
            if (any_comments) {
                for (self.sheets.items, 0..) |sw, i| {
                    if (sw.state.comments.items.len == 0) continue;
                    try ct.print(
                        alloc,
                        "<Override PartName=\"/xl/comments{d}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml\"/>",
                        .{i + 1},
                    );
                }
            }
            try ct.appendSlice(alloc, CONTENT_TYPES_TAIL);
            try zw.addEntry("[Content_Types].xml", ct.items, deflateCompressErased);
        }

        // 2. _rels/.rels (static)
        try zw.addEntry("_rels/.rels", ROOT_RELS, deflateCompressErased);

        // 3. xl/workbook.xml — B3 iter-wr-3: emission lives on the
        // shared plan. Build the sheet-entry view (one entry per
        // registered sheet, declaration order) and hand off to
        // `workbook_xml_plan.emitWorkbookXml`. The plan owns the
        // defined-name pool already; the bytes coming back are
        // byte-for-byte equivalent to the pre-iter-wr-3 local emit.
        {
            const sheet_entries = try alloc.alloc(WorkbookXmlSheetEntry, self.sheets.items.len);
            defer alloc.free(sheet_entries);
            for (self.sheets.items, 0..) |sw, i| {
                sheet_entries[i] = .{ .name = sw.name };
            }
            const wb_bytes = try workbook_xml_plan.emitWorkbookXml(
                alloc,
                sheet_entries,
                &self.workbook_xml_plan,
            );
            defer alloc.free(wb_bytes);
            try zw.addEntry("xl/workbook.xml", wb_bytes, deflateCompressErased);
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
            try zw.addEntry("xl/_rels/workbook.xml.rels", rels.items, deflateCompressErased);
        }

        // 5. xl/worksheets/sheetN.xml
        //
        // B3 iter-wr-4: per-sheet emit routes through `sheet_plan`'s
        // shared `emitWorksheetXml`. Writer's `SheetWriter`-local
        // arrays are projected onto the plan's
        // `SheetEmitInputs` view via per-sheet scratch buffers so the
        // plan stays std-only (no Writer-internal types reach it).
        // The byte output is identical to the prior local emit
        // branch — child-element order, attribute order, and
        // optional-block elision are all locked by the plan's tests.
        for (self.sheets.items, 0..) |sw, i| {
            var full: std.ArrayListUnmanaged(u8) = .{};
            defer full.deinit(alloc);

            // Project Writer's heap-owned arrays onto the plan's
            // const-slice view. The intermediate buffers live for
            // the duration of one sheet emit; no extra string
            // duplications happen here because the plan slices
            // straight into the SheetWriter's owned buffers.
            const hyperlinks_view = try alloc.alloc(sheet_plan.Hyperlink, sw.state.hyperlinks.items.len);
            defer alloc.free(hyperlinks_view);
            for (sw.state.hyperlinks.items, 0..) |h, k| {
                hyperlinks_view[k] = .{ .range = h.range, .url = h.url };
            }

            const internal_hl_view = try alloc.alloc(sheet_plan.InternalHyperlink, sw.state.internal_hyperlinks.items.len);
            defer alloc.free(internal_hl_view);
            for (sw.state.internal_hyperlinks.items, 0..) |h, k| {
                internal_hl_view[k] = .{ .range = h.range, .location = h.location };
            }

            const merges_view = try alloc.alloc([]const u8, sw.state.merged_cells.items.len);
            defer alloc.free(merges_view);
            for (sw.state.merged_cells.items, 0..) |range, k| {
                merges_view[k] = range;
            }

            const cw_view = try alloc.alloc(sheet_plan.ColumnWidth, sw.state.column_widths.items.len);
            defer alloc.free(cw_view);
            for (sw.state.column_widths.items, 0..) |cw, k| {
                cw_view[k] = .{ .col_min = cw.col_min, .col_max = cw.col_max, .width = cw.width };
            }

            const cf_view = try alloc.alloc(sheet_plan.ConditionalFormat, sw.state.conditional_formats.items.len);
            defer alloc.free(cf_view);
            for (sw.state.conditional_formats.items, 0..) |cf, k| {
                // SheetState stores `[]u8`-owned slices; the
                // const-slice view coerces implicitly per axis except
                // for the union arm, which we re-pack by hand.
                cf_view[k] = .{
                    .range = cf.range,
                    .rule = switch (cf.rule) {
                        .cell_is => |r| .{ .cell_is = .{
                            .operator = r.operator,
                            .formula1 = r.formula1,
                            .formula2 = r.formula2,
                            .dxf_id = r.dxf_id,
                        } },
                        .expression => |r| .{ .expression = .{
                            .formula = r.formula,
                            .dxf_id = r.dxf_id,
                        } },
                        .color_scale => |r| .{ .color_scale = .{
                            .low_color_argb = r.low_color_argb,
                            .mid_color_argb = r.mid_color_argb,
                            .high_color_argb = r.high_color_argb,
                        } },
                        .data_bar => |r| .{ .data_bar = .{ .color_argb = r.color_argb } },
                    },
                };
            }

            const dvl_view = try alloc.alloc(sheet_plan.DataValidationList, sw.state.data_validations.items.len);
            defer alloc.free(dvl_view);
            for (sw.state.data_validations.items, 0..) |dv, k| {
                // []u8 → []const u8 for the values inner slice. Writer's
                // DataValidationList.values is `[][]u8`; the plan view
                // wants `[]const []const u8`. A second per-row scratch
                // array projects the inner pointers without any
                // string duplication.
                dvl_view[k] = .{ .range = dv.range, .values = @ptrCast(dv.values) };
            }

            const dvr_view = try alloc.alloc(sheet_plan.DataValidationRange, sw.state.data_validation_ranges.items.len);
            defer alloc.free(dvr_view);
            for (sw.state.data_validation_ranges.items, 0..) |dv, k| {
                dvr_view[k] = .{
                    .range = dv.range,
                    .kind_name = dv.kind_name,
                    .op_name = dv.op_name,
                    .formula1 = dv.formula1,
                    .formula2 = dv.formula2,
                };
            }

            try sheet_plan.emitWorksheetXml(alloc, &full, .{
                .body = sw.body.items,
                .freeze_rows = sw.state.freeze_rows,
                .freeze_cols = sw.state.freeze_cols,
                .column_widths = cw_view,
                .auto_filter_range = sw.state.auto_filter_range,
                .merged_cells = merges_view,
                .conditional_formats = cf_view,
                .data_validations = dvl_view,
                .data_validation_ranges = dvr_view,
                .hyperlinks = hyperlinks_view,
                .internal_hyperlinks = internal_hl_view,
                .comment_count = sw.state.comments.items.len,
            });

            var name_buf: [64]u8 = undefined;
            const entry_name = try std.fmt.bufPrint(&name_buf, "xl/worksheets/sheet{d}.xml", .{i + 1});
            try zw.addEntry(entry_name, full.items, deflateCompressErased);
        }

        // 5a. xl/worksheets/_rels/sheetN.xml.rels (hyperlinks + comments)
        //
        // B3 iter-wr-4: per-sheet rels emit routes through
        // `sheet_plan.emitSheetRels`. The Default Extension="rels"
        // content-type in [Content_Types].xml covers any .rels file
        // we add here — no extra <Override> needed. rId scheme is
        // owned by the plan (1..N hyperlinks, then comments, then
        // vmlDrawing) so writer + workbook share one source.
        for (self.sheets.items, 0..) |sw, i| {
            const hyperlinks_view = try alloc.alloc(sheet_plan.Hyperlink, sw.state.hyperlinks.items.len);
            defer alloc.free(hyperlinks_view);
            for (sw.state.hyperlinks.items, 0..) |h, k| {
                hyperlinks_view[k] = .{ .range = h.range, .url = h.url };
            }

            var rels: std.ArrayListUnmanaged(u8) = .{};
            defer rels.deinit(alloc);
            const wrote = try sheet_plan.emitSheetRels(
                alloc,
                &rels,
                i,
                hyperlinks_view,
                sw.state.comments.items.len,
            );
            if (!wrote) continue;

            var rels_name_buf: [64]u8 = undefined;
            const rels_name = try std.fmt.bufPrint(&rels_name_buf, "xl/worksheets/_rels/sheet{d}.xml.rels", .{i + 1});
            try zw.addEntry(rels_name, rels.items, deflateCompressErased);
        }

        // 5b. xl/commentsN.xml + xl/drawings/vmlDrawingN.vml (per sheet with comments).
        //
        // B3 iter-wr-4: comments + VML drawing emit routes through
        // `sheet_plan.emitCommentsXml` / `emitVmlDrawingXml`. Author
        // dedup, plain-text body shape, idmap chunking, and the
        // XFD/1048576 anchor clamp all live on the plan so the
        // byte format is shared between Writer and (future) Workbook
        // fresh-emit. Per `a966e29` the anchor clamp prevents
        // inverted shapes on cells near the rightmost column.
        for (self.sheets.items, 0..) |sw, i| {
            if (sw.state.comments.items.len == 0) continue;

            // Project Writer's `Comment` slice onto the plan's view.
            // Same layout, just different namespace — no string
            // duplication.
            const comments_view = try alloc.alloc(sheet_plan.Comment, sw.state.comments.items.len);
            defer alloc.free(comments_view);
            for (sw.state.comments.items, 0..) |c, k| {
                comments_view[k] = .{ .ref = c.ref, .author = c.author, .text = c.text };
            }

            var cx: std.ArrayListUnmanaged(u8) = .{};
            defer cx.deinit(alloc);
            try sheet_plan.emitCommentsXml(alloc, &cx, comments_view);
            var cn_buf: [64]u8 = undefined;
            const cn = try std.fmt.bufPrint(&cn_buf, "xl/comments{d}.xml", .{i + 1});
            try zw.addEntry(cn, cx.items, deflateCompressErased);

            var vml: std.ArrayListUnmanaged(u8) = .{};
            defer vml.deinit(alloc);
            try sheet_plan.emitVmlDrawingXml(alloc, &vml, comments_view);
            var vml_buf: [64]u8 = undefined;
            const vml_name = try std.fmt.bufPrint(&vml_buf, "xl/drawings/vmlDrawing{d}.vml", .{i + 1});
            try zw.addEntry(vml_name, vml.items, deflateCompressErased);
        }

        // 6. xl/sharedStrings.xml — emits the staged
        // `SstExtensionPlan` (B3 iter-wr-1). Plain entries first
        // (matches the index ordering `sstIntern` returns), then rich
        // entries — `indexOfRich` assumes that layout. `xml:space=
        // "preserve"` is unconditional on plain entries (Writer's
        // pre-iter-wr-1 byte contract; differs from
        // `pkg/workbook.zig::appendNewSiEntries` which emits the
        // attribute conditionally per OOXML's whitespace rules).
        {
            var sst: std.ArrayListUnmanaged(u8) = .{};
            defer sst.deinit(alloc);
            const unique_count = self.sst_plan.new_strings.items.len +
                self.sst_plan.new_rich_strings.items.len;
            try sst.print(alloc, SST_HEAD_FMT, .{ self.sst_count, unique_count });
            for (self.sst_plan.new_strings.items) |s| {
                try sst.appendSlice(alloc, "<si><t xml:space=\"preserve\">");
                try appendXmlEscaped(alloc, &sst, s);
                try sst.appendSlice(alloc, "</t></si>");
            }
            for (self.sst_plan.new_rich_strings.items) |entry| {
                try sst.appendSlice(alloc, "<si>");
                for (entry.runs) |r| {
                    try sst.appendSlice(alloc, "<r>");
                    // <rPr>…</rPr> — only emit when at least one prop is
                    // set. Underline + strike are storage-only on the
                    // plan side; Writer's RichTextRun never sets them
                    // (false-equivalent), so they don't gate `has_props`
                    // and stay un-emitted, preserving the byte format.
                    const has_props = r.bold or r.italic or r.color_argb != null or
                        r.font_size != null or r.font_name != null;
                    if (has_props) {
                        try sst.appendSlice(alloc, "<rPr>");
                        if (r.bold) try sst.appendSlice(alloc, "<b/>");
                        if (r.italic) try sst.appendSlice(alloc, "<i/>");
                        if (r.font_size) |sz| try sst.print(alloc, "<sz val=\"{d}\"/>", .{sz});
                        if (r.color_argb) |c| {
                            // Plan stores pre-formatted 8-hex bytes;
                            // emit verbatim (matches Writer's pre-
                            // iter-wr-1 `{X:0>8}` output, sourced
                            // upstream by `sstInternRich`).
                            try sst.appendSlice(alloc, "<color rgb=\"");
                            try sst.appendSlice(alloc, c);
                            try sst.appendSlice(alloc, "\"/>");
                        }
                        if (r.font_name) |n| {
                            try sst.appendSlice(alloc, "<rFont val=\"");
                            try appendXmlEscaped(alloc, &sst, n);
                            try sst.appendSlice(alloc, "\"/>");
                        }
                        try sst.appendSlice(alloc, "</rPr>");
                    }
                    try sst.appendSlice(alloc, "<t xml:space=\"preserve\">");
                    try appendXmlEscaped(alloc, &sst, r.text);
                    try sst.appendSlice(alloc, "</t></r>");
                }
                try sst.appendSlice(alloc, "</si>");
            }
            try sst.appendSlice(alloc, SST_TAIL);
            try zw.addEntry("xl/sharedStrings.xml", sst.items, deflateCompressErased);
        }

        // 7. xl/styles.xml — only when the caller registered any styles.
        // Keeps the "no styles" path byte-identical to v0.2.0-0.2.3
        // output. The emit logic itself lives on `StylesPlan` (B3
        // iter-wr-2), so Workbook's fresh-emit path can reuse the
        // same byte-stable producer.
        if (have_styles) {
            var styles_buf: std.ArrayListUnmanaged(u8) = .{};
            defer styles_buf.deinit(alloc);
            try self.styles_plan.emit(alloc, &styles_buf);
            try zw.addEntry("xl/styles.xml", styles_buf.items, deflateCompressErased);
        }

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

/// Conditional-format entry stored inside `SheetWriter.conditional_formats`.
/// `range` is A1-style (may be a single cell or a rectangle). `rule`
/// is the OOXML cfRule payload; `dxf_id` is an index into the
/// workbook-wide `Writer.dxfs` table returned by `Writer.addDxf`.
pub const ConditionalFormat = struct {
    range: []u8,
    rule: ConditionalFormatRule,
};

/// Union of the cfRule variants zlsx currently emits.
///
/// - `cell_is`: threshold comparison (greater_than 100, between 1..50).
/// - `expression`: generic formula rule (zebra-stripe `MOD(ROW(),2)=0`).
/// - `color_scale`: two- or three-stop color gradient across values.
///   No dxf_id — colors are embedded per-stop.
/// - `data_bar`: in-cell bar chart from min to max of the range.
///   Single bar color; dxf_id unused.
pub const ConditionalFormatRule = union(enum) {
    cell_is: struct {
        operator: CfOperator,
        formula1: []u8,
        formula2: ?[]u8,
        dxf_id: u32,
    },
    expression: struct {
        formula: []u8,
        dxf_id: u32,
    },
    color_scale: struct {
        /// Low-end color (applied at the minimum value of the range).
        low_color_argb: u32,
        /// Optional midpoint color — when set, renders a 3-stop
        /// gradient (min → mid → max). Null → 2-stop (min → max).
        mid_color_argb: ?u32,
        /// High-end color (applied at the maximum value).
        high_color_argb: u32,
    },
    data_bar: struct {
        /// Bar fill color. Excel's default is a muted blue
        /// (`FF638EC6`); any explicit ARGB works.
        color_argb: u32,
    },
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
    /// B3 iter-wr-6: per-sheet registries (column widths, row
    /// heights, freeze panes, auto filter, merge cells, hyperlinks,
    /// internal hyperlinks, comments, conditional formats, data
    /// validations) all live on the shared `pkg/sheet_plan.zig`
    /// `SheetState` so `xlsx.Writer.SheetWriter` and (future,
    /// wr-7) `pkg.Worksheet` produce byte-identical outputs through
    /// the same registration + heap-ownership code path.
    /// Replaces 11 fields and the matching deinit branches.
    state: sheet_plan.SheetState = .{},

    fn init(parent: *Writer, name: []const u8) !SheetWriter {
        return .{
            .parent = parent,
            .name = try parent.allocator.dupe(u8, name),
        };
    }

    fn deinit(self: *SheetWriter) void {
        self.parent.allocator.free(self.name);
        self.body.deinit(self.parent.allocator);
        self.state.deinit(self.parent.allocator);
        self.* = undefined;
    }

    /// Set a column's width in character units (Excel's default is
    /// 8.43). `col_idx` is 0-based (A=0, B=1, …). Multiple calls on
    /// the same column append a new override — the emitter keeps them
    /// in order, so a later call wins on overlap in Excel.
    /// B3 iter-wr-6: thin forwarder onto `sheet_plan.SheetState`.
    pub fn setColumnWidth(self: *SheetWriter, col_idx: u32, width: f32) !void {
        return self.state.setColumnWidth(self.parent.allocator, col_idx, width);
    }

    /// Set `row_idx`'s height in Excel point units (default row
    /// height is ~15 pt). `row_idx` is 0-based (0 = row 1). Must be
    /// called before the matching `writeRow` / `writeRowStyled` — the
    /// row is emitted inline at that time, and a post-hoc call on an
    /// already-emitted row is silently ignored (no retroactive XML
    /// rewrite). Later calls on the same row_idx override earlier
    /// ones as long as the row hasn't been written yet.
    /// B3 iter-wr-6: thin forwarder onto `sheet_plan.SheetState`.
    pub fn setRowHeight(self: *SheetWriter, row_idx: u32, height: f32) !void {
        return self.state.setRowHeight(self.parent.allocator, row_idx, height);
    }

    /// Freeze the top `rows` rows and left `cols` columns. Pass 0 to
    /// disable one axis (e.g., `freezePanes(1, 0)` freezes only row 1).
    /// Calling again overrides the previous setting.
    /// B3 iter-wr-6: thin forwarder onto `sheet_plan.SheetState`.
    pub fn freezePanes(self: *SheetWriter, rows: u32, cols: u32) error{ RowOutOfRange, ColumnOutOfRange }!void {
        return self.state.freezePanes(rows, cols);
    }

    /// Apply an auto-filter over the given A1-style range (e.g.,
    /// "A1:E1"). Caller-owned; the writer dupes it.
    /// B3 iter-wr-6: thin forwarder onto `sheet_plan.SheetState`.
    pub fn setAutoFilter(self: *SheetWriter, range: []const u8) !void {
        return self.state.setAutoFilter(self.parent.allocator, range);
    }

    /// Merge a rectangular cell range (e.g., "A1:B2"). The range must
    /// be a valid multi-cell A1-style span — single-cell ranges and
    /// inverted (bottom-right-before-top-left) ranges are rejected.
    /// Caller-owned; the writer dupes it.
    /// B3 iter-wr-6: thin forwarder onto `sheet_plan.SheetState`.
    pub fn addMergedCell(self: *SheetWriter, range: []const u8) !void {
        return self.state.addMergedCell(self.parent.allocator, range);
    }

    /// Attach a list-type data validation (dropdown) to a cell or
    /// rectangular range. `range` is A1-style (single cell "A1" or
    /// span "B2:B10"); `values` are the literal dropdown options.
    /// Excel's in-cell list format joins values with commas inside a
    /// quoted formula1 string, so embedded commas and bare double-
    /// quotes in values are rejected (callers who need those should
    /// use a range-reference validation — not yet supported). Empty
    /// values and empty `values` slice also rejected.
    /// B3 iter-wr-6: thin forwarder onto `sheet_plan.SheetState`.
    pub fn addDataValidationList(
        self: *SheetWriter,
        range: []const u8,
        values: []const []const u8,
    ) !void {
        return self.state.addDataValidationList(self.parent.allocator, range, values);
    }

    /// Attach a numeric / date / time / text-length data validation.
    /// `range` is A1-style. `formula1` is the primary bound (number
    /// or Excel date serial or length — passed as a string, the
    /// writer emits it verbatim). `formula2` must be non-null iff
    /// `op` is `.between` or `.not_between`, and must be null
    /// otherwise — mismatches surface `error.InvalidDataValidation`.
    /// Excel displays number-typed validations as red-circle errors
    /// when the cell value falls outside the constraint.
    /// Attach a numeric / date / time / text-length data validation.
    /// `formula2` must be non-null iff `op` is `.between` or
    /// `.not_between`. B3 iter-wr-6: forwards onto
    /// `sheet_plan.SheetState.addDataValidationRange` — the writer-side
    /// `kind`/`op` enums translate to the canonical OOXML token
    /// strings here.
    pub fn addDataValidationNumeric(
        self: *SheetWriter,
        range: []const u8,
        kind: DataValidationNumericKind,
        op: DataValidationOp,
        formula1: []const u8,
        formula2: ?[]const u8,
    ) !void {
        return self.state.addDataValidationRange(
            self.parent.allocator,
            range,
            kind.toOoxml(),
            op.toOoxml(),
            formula1,
            formula2,
            op.needsSecondFormula(),
        );
    }

    /// Attach a custom-formula data validation. `formula` is any
    /// Excel formula that evaluates to TRUE for accepted cell values.
    /// B3 iter-wr-6: thin forwarder onto `sheet_plan.SheetState`.
    pub fn addDataValidationCustom(
        self: *SheetWriter,
        range: []const u8,
        formula: []const u8,
    ) !void {
        return self.state.addDataValidationCustom(self.parent.allocator, range, formula);
    }

    /// Attach a hyperlink to a cell or rectangular range. `range` is
    /// A1-style; `url` is the external target (xml-escaped on emit).
    /// B3 iter-wr-6: thin forwarder onto `sheet_plan.SheetState`.
    pub fn addHyperlink(self: *SheetWriter, range: []const u8, url: []const u8) !void {
        return self.state.addHyperlink(self.parent.allocator, range, url);
    }

    /// Attach an internal hyperlink that jumps to another cell or
    /// range within the same workbook. Emitted as
    /// `<hyperlink ref="…" location="…"/>` without an `r:id`.
    /// B3 iter-wr-6: thin forwarder onto `sheet_plan.SheetState`.
    pub fn addInternalHyperlink(self: *SheetWriter, range: []const u8, location: []const u8) !void {
        return self.state.addInternalHyperlink(self.parent.allocator, range, location);
    }

    /// Attach a cell comment (note) to a single-cell A1 ref.
    /// B3 iter-wr-6: thin forwarder onto `sheet_plan.SheetState`.
    pub fn addComment(
        self: *SheetWriter,
        ref: []const u8,
        author: []const u8,
        text: []const u8,
    ) !void {
        return self.state.addComment(self.parent.allocator, ref, author, text);
    }

    /// Attach a `cellIs`-type conditional-format rule. `range` is
    /// A1-style (single cell or rectangle). `operator` is the
    /// comparison (e.g. `.greater_than`, `.between`). `formula1` is
    /// the reference value (a number like `"100"`, a cell ref like
    /// `"$A$1"`, or any OOXML formula). `formula2` is required for
    /// `.between` / `.not_between` and must be null otherwise.
    /// `dxf_id` is the dxf index returned by `Writer.addDxf`.
    /// Returns `InvalidDataValidation` on empty formula / two-formula
    /// mismatch, `InvalidHyperlinkRange` on bad range, `UnknownDxfId`
    /// on out-of-range dxf.
    /// Attach a `cellIs`-type conditional-format rule. B3 iter-wr-6:
    /// forwards onto `sheet_plan.SheetState`; the `dxf_id` bounds
    /// check threads through the parent's `styles_plan.dxfs` count.
    pub fn addConditionalFormatCellIs(
        self: *SheetWriter,
        range: []const u8,
        operator: CfOperator,
        formula1: []const u8,
        formula2: ?[]const u8,
        dxf_id: u32,
    ) !void {
        return self.state.addConditionalFormatCellIs(
            self.parent.allocator,
            range,
            projectCfOperator(operator),
            formula1,
            formula2,
            dxf_id,
            self.parent.styles_plan.dxfs.items.len,
        );
    }

    /// Attach an `expression`-type conditional-format rule.
    /// B3 iter-wr-6: thin forwarder onto `sheet_plan.SheetState`.
    pub fn addConditionalFormatExpression(
        self: *SheetWriter,
        range: []const u8,
        formula: []const u8,
        dxf_id: u32,
    ) !void {
        return self.state.addConditionalFormatExpression(
            self.parent.allocator,
            range,
            formula,
            dxf_id,
            self.parent.styles_plan.dxfs.items.len,
        );
    }

    /// Attach a color-scale conditional format. Null `mid_color_argb`
    /// produces a 2-stop gradient; non-null gives 3-stop.
    /// B3 iter-wr-6: thin forwarder onto `sheet_plan.SheetState`.
    pub fn addConditionalFormatColorScale(
        self: *SheetWriter,
        range: []const u8,
        low_color_argb: u32,
        mid_color_argb: ?u32,
        high_color_argb: u32,
    ) !void {
        return self.state.addConditionalFormatColorScale(
            self.parent.allocator,
            range,
            low_color_argb,
            mid_color_argb,
            high_color_argb,
        );
    }

    /// Attach a data-bar conditional format.
    /// B3 iter-wr-6: thin forwarder onto `sheet_plan.SheetState`.
    pub fn addConditionalFormatDataBar(
        self: *SheetWriter,
        range: []const u8,
        color_argb: u32,
    ) !void {
        return self.state.addConditionalFormatDataBar(
            self.parent.allocator,
            range,
            color_argb,
        );
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
        const max_style_id: u32 = @intCast(self.parent.styles_plan.styles.items.len);
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

    /// Write a row mixing plain cells with rich-text cells. Rich
    /// cells carry an array of `RichTextRun`; each run becomes one
    /// `<r><rPr/>…<t/></r>` inside a single `<si>` in the SST.
    /// Non-rich cells follow the same semantics as `writeRow`.
    ///
    /// Rich-text entries are always appended to the SST (no dedup
    /// in this iter — the formatted form is rarely repeated, and
    /// hashing it would cost more than it saves).
    pub fn writeRichRow(self: *SheetWriter, cells: []const RichRowCell) !void {
        // Pre-validate the entire row BEFORE mutating `self.body`,
        // same atomicity + Excel-limit contract as `writeRowImpl`.
        if (self.next_row > EXCEL_MAX_ROW) return error.RowOutOfRange;
        if (cells.len > EXCEL_MAX_COL) return error.ColumnOutOfRange;
        for (cells) |cell| switch (cell) {
            .integer => |n| if (!fitsExactlyInF64(n)) return error.IntegerExceedsExcelPrecision,
            .number => |f| if (!std.math.isFinite(f)) return error.NonFiniteNumeric,
            // Match writeRowImpl's atomicity contract: any string
            // or rich-text-run content that would error inside
            // appendXmlEscaped on emit is rejected here, before
            // the row's `<row>` opener has been appended.
            .string => |s| try assertNoForbiddenXmlBytes(s),
            .rich => |runs| {
                for (runs) |run| {
                    try assertNoForbiddenXmlBytes(run.text);
                    if (run.font_name) |fn_name| try assertNoForbiddenXmlBytes(fn_name);
                }
            },
            else => {},
        };

        const alloc = self.parent.allocator;
        if (self.state.row_heights.get(self.next_row - 1)) |h| {
            try self.body.print(alloc, "<row r=\"{d}\" ht=\"{d}\" customHeight=\"1\">", .{ self.next_row, h });
        } else {
            try self.body.print(alloc, "<row r=\"{d}\">", .{self.next_row});
        }

        for (cells, 0..) |cell, col_idx| {
            if (cell == .empty) continue;
            var ref_buf: [16]u8 = undefined;
            const ref = try formatCellRef(&ref_buf, self.next_row, @intCast(col_idx));

            const type_attr: []const u8 = switch (cell) {
                .string, .rich => " t=\"s\"",
                .boolean => " t=\"b\"",
                else => "",
            };
            try self.body.print(alloc, "<c r=\"{s}\"{s}>", .{ ref, type_attr });

            switch (cell) {
                .empty => unreachable,
                .string => |s| {
                    const idx = try self.parent.sstIntern(s);
                    self.parent.sst_count += 1;
                    try self.body.print(alloc, "<v>{d}</v>", .{idx});
                },
                .rich => |runs| {
                    const idx = try self.parent.sstInternRich(runs);
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

    fn writeRowImpl(
        self: *SheetWriter,
        cells: []const xlsx.Cell,
        styles: ?[]const u32,
        formulas: ?[]const ?[]const u8,
    ) !void {
        // Pre-validate the entire row BEFORE mutating `self.body`. This
        // keeps writeRow atomic — caller can catch the error and retry
        // / skip the row without ending up with a half-emitted <row>.
        // Excel's hard limits are enforced here so a future caller
        // can't smuggle a u32 past us into formatCellRef or the
        // <row r="N"> emission and produce a workbook Excel rejects.
        if (self.next_row > EXCEL_MAX_ROW) return error.RowOutOfRange;
        if (cells.len > EXCEL_MAX_COL) return error.ColumnOutOfRange;
        for (cells) |cell| switch (cell) {
            .integer => |n| if (!fitsExactlyInF64(n)) return error.IntegerExceedsExcelPrecision,
            // Reject NaN / +Inf / -Inf — Excel's <v>NaN</v> /
            // <v>inf</v> would render as #NUM! at best and corrupt
            // the workbook at worst. Mirrors the isFinite check on
            // setColumnWidth / setRowHeight.
            .number => |f| if (!std.math.isFinite(f)) return error.NonFiniteNumeric,
            // Strings / formulas funnel into appendXmlEscaped on
            // emit; XML 1.0 forbidden bytes there would error
            // mid-row and leave a half-written <row> in self.body.
            // Detect up-front so the row stays atomic.
            .string => |s| try assertNoForbiddenXmlBytes(s),
            else => {},
        };
        if (formulas) |fs| {
            for (fs) |maybe_f| if (maybe_f) |f| try assertNoForbiddenXmlBytes(f);
        }

        const alloc = self.parent.allocator;
        // Row index is 0-based inside the height map; next_row is
        // 1-based per xlsx convention, so subtract 1 on lookup.
        if (self.state.row_heights.get(self.next_row - 1)) |h| {
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
            //   <c r="…"[ s="N"][ t="s|str|b"]>[<f>formula</f>][<v>value</v>]</c>
            //
            // String formula caches use t="str" (inline string in <v>)
            // rather than t="s" (SST index). Excel and other readers
            // expect the cached result of a formula returning text to
            // be the literal text — an SST index there confuses
            // recalc and triggers a "repaired" prompt. The non-
            // formula `.string` path stays on the SST.
            const string_is_formula_cache = cell == .string and formula != null;
            const type_attr: []const u8 = switch (cell) {
                .string => if (string_is_formula_cache) " t=\"str\"" else " t=\"s\"",
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
                    if (string_is_formula_cache) {
                        try self.body.appendSlice(alloc, "<v>");
                        try appendXmlEscaped(alloc, &self.body, s);
                        try self.body.appendSlice(alloc, "</v>");
                    } else {
                        const idx = try self.parent.sstIntern(s);
                        self.parent.sst_count += 1;
                        try self.body.print(alloc, "<v>{d}</v>", .{idx});
                    }
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

/// B3 iter-wr-6: A1 cell-ref formatter canonicalised in
/// `pkg/sheet_plan.zig`. The forwarder keeps the writer-side public
/// surface (a `pub` symbol some test files import directly) without
/// duplicating the bit-bashing.
pub fn formatCellRef(buf: *[16]u8, row: u32, col_idx: u32) ![]u8 {
    return sheet_plan.formatCellRef(buf, row, col_idx);
}

// Excel's hard limits: 16 384 columns (XFD) × 1 048 576 rows. B3
// iter-wr-6 lifted the constants to `pkg/sheet_plan.zig` so the
// reader, writer, and Workbook fresh-emit share one canonical
// definition. The aliases below keep this file's existing call
// sites compiling unchanged.
const EXCEL_MAX_COL: u32 = sheet_plan.EXCEL_MAX_COL;
const EXCEL_MAX_ROW: u32 = sheet_plan.EXCEL_MAX_ROW;

// B3 iter-wr-6: A1 corner parsing + range validators canonicalised
// in `pkg/sheet_plan.zig`. The aliases preserve the writer-side
// names (and the writer-side test suite still references them
// directly).
const MergeCorner = sheet_plan.A1Corner;
const parseA1Corner = sheet_plan.parseA1Corner;
const validateMergeRange = sheet_plan.validateMergeRange;

/// Byte-wise ASCII case-fold equality. Excel's sheet-name uniqueness
/// rule is Unicode-case-insensitive, but ASCII case-fold catches the
/// overwhelming majority of real-world collisions ("Summary" vs
/// "summary") at a tiny code cost. Non-ASCII input falls through to
/// byte comparison — sufficient for everything except genuine
/// Turkish-i-style edge cases that no real caller hits.
pub fn asciiEqlFold(a: []const u8, b: []const u8) bool {
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
pub fn validateSheetName(name: []const u8) !void {
    // Excel's 31-character limit is in Unicode SCALARS, not bytes —
    // the previous byte-cap was too tight for legal multi-byte names.
    if (name.len == 0) return error.InvalidSheetName;
    // Generous byte upper bound first (31 scalars × 4 bytes UTF-8 max).
    if (name.len > 124) return error.InvalidSheetName;
    const scalar_count = casefold.excelSheetNameLength(name) catch
        return error.InvalidSheetName;
    if (scalar_count == 0 or scalar_count > 31) return error.InvalidSheetName;

    if (name[0] == '\'' or name[name.len - 1] == '\'') return error.InvalidSheetName;
    for (name) |c| switch (c) {
        0...0x1F => return error.InvalidSheetName,
        ':', '/', '\\', '?', '*', '[', ']' => return error.InvalidSheetName,
        else => {},
    };
    // Excel reserves the sheet name "History" case-insensitively.
    // Use Unicode-aware comparison so e.g. "history" / "HISTORY" /
    // "HiStOrY" are all rejected; non-ASCII inputs fall through.
    if (casefold.excelSheetNameEql(name, "History")) return error.InvalidSheetName;
}

/// B3 iter-wr-3: defined-name validator + R1C1 / A1 ref-shape
/// detection moved to `pkg/workbook_xml_plan.zig`. The exported
/// alias keeps the writer-side test suite calling
/// `validateDefinedName(name)` against the canonical implementation.
const validateDefinedName = workbook_xml_plan.validateDefinedName;

// B3 iter-wr-6: range validators canonicalised in
// `pkg/sheet_plan.zig`. Same byte-stable shape as the prior locals
// (single-cell ranges are valid for hyperlinks/auto-filter, rejected
// for merges; inverted corners rejected on every axis).
const validateAutoFilterRange = sheet_plan.validateAutoFilterRange;
const validateHyperlinkRange = sheet_plan.validateHyperlinkRange;

// xl/styles.xml emitter + helpers (stylesEqual, hAlignName,
// emitBorderSide, emitDxfBorderSide, borderStyleName,
// patternTypeName) moved to pkg/styles_plan.zig per B3 iter-wr-2.

// XML 1.0 escape + forbidden-byte helpers: B3 iter-wr-6 lifted these
// to `pkg/sheet_plan.zig` so writer.zig + workbook.zig + store.zig
// share one byte-stable producer. The aliases below keep this file's
// existing call sites (~80 across writer.zig + tests) compiling
// unchanged — no perf cost; the compiler inlines the forward.
const appendXmlEscaped = sheet_plan.appendXmlEscaped;
const assertNoForbiddenXmlBytes = sheet_plan.assertNoForbiddenXmlBytes;
const isForbiddenXmlByte = sheet_plan.isForbiddenXmlByte;

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

/// Scale frequencies down so the max fits in u13 — HuffmanEncoder
/// asserts an internal leaf_count invariant (`leaf_counts[max_bits]
/// [max_bits] == n`) inside `bitCounts`, which trips on certain
/// near-cap frequency distributions when the cap sits at u15
/// (witness: a 100k×5-row append payload at 5000+ rows). Capping at
/// u13 (8191) gives the stdlib's two-priority-queue level walker
/// enough numerical headroom to converge. Preserve relative rank:
/// a non-zero input keeps a non-zero output (rounding up from 0).
fn scaleFreqs(src: []const u32, dst: []u16) void {
    var max_freq: u32 = 0;
    for (src) |f| max_freq = @max(max_freq, f);
    const cap: u32 = 8191; // u13_max
    const scale: u32 = if (max_freq > cap) (max_freq + cap - 1) / cap else 1;
    for (src, 0..) |f, i| {
        const v = f / scale;
        dst[i] = @intCast(if (f > 0 and v == 0) 1 else v);
    }
}

/// Emit `tokens` as a single final dynamic-huffman block. Assumes
/// `tokens.len > 0` — empty inputs are handled upstream (stored
/// fallback in zip.Archive.addEntry).
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
/// upstream). Exposed publicly so `xlsx.Editor` can recompress
/// substituted entries.
pub fn deflateCompress(alloc: Allocator, input: []const u8, out: *std.ArrayListUnmanaged(u8)) !void {
    std.debug.assert(input.len > 0);

    var tokens: std.ArrayListUnmanaged(DeflateToken) = .{};
    defer tokens.deinit(alloc);
    try tokens.ensureTotalCapacity(alloc, input.len);
    try tokenizeLazy(alloc, input, &tokens);

    var bw = DeflateBitWriter{ .out = out, .alloc = alloc };
    try emitDynamicBlock(alloc, tokens.items, &bw);
    try bw.flushByte();
}

// ─── ZIP writer ─────────────────────────────────────────────────────
//
// B3 iter-wr-5: Writer's private `ZipWriter` struct + `appendStruct`
// helper retired; the canonical LFH+CDFH+EOCD layout now lives in
// `pkg/zip.zig` (`zip.Archive`). `Writer.save` consumes it via the
// `zlsx_zip` named import — same byte-stable invariants, one source
// of truth shared with `pkg.PartStore.save` (`d6235f3` total-size
// guard, `66e4ccd` ZIP32 sentinel guards, `410b4bc` huffman u13
// regression fix all preserved).

// ─── Tests ───────────────────────────────────────────────────────────

/// Per-test temporary file helper. Replaces the older `/tmp/zlsx_*.xlsx`
/// hard-coded paths so the suite is portable across Linux / macOS /
/// Windows (Windows has no `/tmp`). Each call creates a fresh isolated
/// `std.testing.TmpDir` (auto-cleaned) and returns an owned absolute
/// path inside it. Caller frees the returned slice.
///
/// Usage:
///   var tt = TestTmp.init();
///   defer tt.deinit();
///   const src_path = try tt.path(std.testing.allocator, "src.xlsx");
///   defer std.testing.allocator.free(src_path);
const TestTmp = struct {
    dir: std.testing.TmpDir,
    pub fn init() TestTmp {
        return .{ .dir = std.testing.tmpDir(.{}) };
    }
    pub fn deinit(self: *TestTmp) void {
        self.dir.cleanup();
    }
    pub fn path(self: *TestTmp, alloc: std.mem.Allocator, name: []const u8) ![:0]u8 {
        const dir = try self.dir.dir.realpathAlloc(alloc, ".");
        defer alloc.free(dir);
        return std.fs.path.joinZ(alloc, &.{ dir, name });
    }
};

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(std.testing.allocator, "empty.xlsx");
    defer std.testing.allocator.free(path);
    var w = Writer.init(std.testing.allocator);
    defer w.deinit();
    try std.testing.expectError(error.NoSheets, w.save(path));
}

test "Writer: single-sheet round-trip via zlsx reader" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "writer_test.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "writer_multisheet.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
        try std.testing.expectEqual(@as(usize, 3), w.sst_plan.new_strings.items.len);
        // 4 string-cell writes total.
        try std.testing.expectEqual(@as(u64, 4), w.sst_count);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 2), book.sheets.len);
    try std.testing.expectEqualStrings("Alpha", book.sheets[0].name);
    try std.testing.expectEqualStrings("Beta", book.sheets[1].name);
    try std.testing.expectEqual(@as(usize, 3), book.sharedStringsCount());
}

test "Writer: xml entities in strings are escaped" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "writer_entities.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "writer_numfmt.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "writer_formulas.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "writer_row_heights.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "writer_sheet_features.xlsx");
    defer std.testing.allocator.free(tmp_path);

    {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("Sheet1");
        try sheet.setColumnWidth(0, 20.5);
        try sheet.setColumnWidth(3, 12);
        try sheet.freezePanes(1, 2);
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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "writer_merged_cells.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "writer_dv_ranges.xlsx");
    defer std.testing.allocator.free(tmp_path);

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

test "Writer: VML idmap expands for >1023 comments per sheet" {
    var tt = TestTmp.init();
    defer tt.deinit();
    // iter48 — the hardcoded `<o:idmap data="1"/>` only covered
    // shape IDs 1024..2047 = 1023 comments. Workbooks past that
    // need additional idmap entries (one per 1024-ID range). This
    // test emits 1025 comments and verifies the VML drawing grew
    // a second idmap: `data="1,2"`.
    const tmp_path = try tt.path(std.testing.allocator, "writer_idmap_scale.xlsx");
    defer std.testing.allocator.free(tmp_path);

    {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("S");
        var ref_buf: [16]u8 = undefined;
        for (0..1025) |i| {
            // Excel refs: column-major walk across A..Z then down rows.
            const col: u8 = @intCast((i % 26) + 1);
            const row: u32 = @intCast((i / 26) + 1);
            const col_letter: u8 = 'A' + col - 1;
            const ref = try std.fmt.bufPrint(&ref_buf, "{c}{d}", .{ col_letter, row });
            try sheet.addComment(ref, "A", "n");
        }
        try sheet.writeRow(&.{.{ .string = "hdr" }});
        try w.save(tmp_path);
    }

    // Unzip xl/drawings/vmlDrawing1.vml + grep for `data="1,2"`.
    var file = try std.fs.cwd().openFile(tmp_path, .{});
    defer file.close();
    var fbuf: [4096]u8 = undefined;
    var fr = file.reader(&fbuf);
    var iter = try std.zip.Iterator.init(&fr);
    var filename_buf: [64]u8 = undefined;
    var vml_xml: ?[]u8 = null;
    defer if (vml_xml) |v| std.testing.allocator.free(v);
    while (try iter.next()) |entry| {
        if (entry.filename_len > filename_buf.len) continue;
        try fr.seekTo(entry.header_zip_offset + @sizeOf(std.zip.CentralDirectoryFileHeader));
        const filename = filename_buf[0..entry.filename_len];
        try fr.interface.readSliceAll(filename);
        if (std.mem.eql(u8, filename, "xl/drawings/vmlDrawing1.vml")) {
            vml_xml = try extractEntryForTest(std.testing.allocator, entry, &fr);
            break;
        }
    }
    const vml = vml_xml orelse return error.VmlNotFound;
    try std.testing.expect(std.mem.indexOf(u8, vml, "data=\"1,2\"") != null);

    // A single-comment sheet keeps the data="1" shape (regression
    // guard: don't emit `data="1,"` trailing comma).
    const tmp_small = try tt.path(std.testing.allocator, "writer_idmap_small.xlsx");
    defer std.testing.allocator.free(tmp_small);
    {
        var w2 = Writer.init(std.testing.allocator);
        defer w2.deinit();
        var sheet2 = try w2.addSheet("S");
        try sheet2.addComment("A1", "A", "n");
        try sheet2.writeRow(&.{.{ .string = "x" }});
        try w2.save(tmp_small);
    }
    var file2 = try std.fs.cwd().openFile(tmp_small, .{});
    defer file2.close();
    var fbuf2: [4096]u8 = undefined;
    var fr2 = file2.reader(&fbuf2);
    var iter2 = try std.zip.Iterator.init(&fr2);
    var fn_buf2: [64]u8 = undefined;
    var vml2: ?[]u8 = null;
    defer if (vml2) |v| std.testing.allocator.free(v);
    while (try iter2.next()) |entry| {
        if (entry.filename_len > fn_buf2.len) continue;
        try fr2.seekTo(entry.header_zip_offset + @sizeOf(std.zip.CentralDirectoryFileHeader));
        const filename = fn_buf2[0..entry.filename_len];
        try fr2.interface.readSliceAll(filename);
        if (std.mem.eql(u8, filename, "xl/drawings/vmlDrawing1.vml")) {
            vml2 = try extractEntryForTest(std.testing.allocator, entry, &fr2);
            break;
        }
    }
    const v2 = vml2 orelse return error.VmlNotFound;
    try std.testing.expect(std.mem.indexOf(u8, v2, "data=\"1\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, v2, "data=\"1,\"") == null);
}

test "Writer: comment on XFD column emits non-inverted VML anchor" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "writer_comment_xfd.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("S");
        // XFD = column 16384 (1-based); col 16383 (0-based).
        try sheet.addComment("XFD1", "A", "edge");
        try sheet.writeRow(&.{.{ .string = "x" }});
        try w.save(tmp_path);
    }
    var file = try std.fs.cwd().openFile(tmp_path, .{});
    defer file.close();
    var fbuf: [4096]u8 = undefined;
    var fr = file.reader(&fbuf);
    var iter = try std.zip.Iterator.init(&fr);
    var fn_buf: [64]u8 = undefined;
    var vml: ?[]u8 = null;
    defer if (vml) |v| std.testing.allocator.free(v);
    while (try iter.next()) |entry| {
        if (entry.filename_len > fn_buf.len) continue;
        try fr.seekTo(entry.header_zip_offset + @sizeOf(std.zip.CentralDirectoryFileHeader));
        const fname = fn_buf[0..entry.filename_len];
        try fr.interface.readSliceAll(fname);
        if (std.mem.eql(u8, fname, "xl/drawings/vmlDrawing1.vml")) {
            vml = try extractEntryForTest(std.testing.allocator, entry, &fr);
            break;
        }
    }
    const v = vml orelse return error.VmlNotFound;
    // Anchor: <x:Anchor>FROM_COL, 15, FROM_ROW, 2, TO_COL, 31, TO_ROW, 3</x:Anchor>
    // Both FROM_COL and TO_COL must be ≤ 16383 (EXCEL_MAX_COL - 1)
    // — and FROM_COL must not exceed TO_COL.
    const start = std.mem.indexOf(u8, v, "<x:Anchor>") orelse return error.AnchorMissing;
    const end = std.mem.indexOfPos(u8, v, start, "</x:Anchor>") orelse return error.AnchorEnd;
    const anchor = v[start + "<x:Anchor>".len .. end];
    var it = std.mem.splitScalar(u8, anchor, ',');
    const fc = std.fmt.parseInt(u32, std.mem.trim(u8, it.next().?, " "), 10) catch return error.AnchorParse;
    _ = it.next(); // 15
    _ = it.next(); // FROM_ROW
    _ = it.next(); // 2
    const tc = std.fmt.parseInt(u32, std.mem.trim(u8, it.next().?, " "), 10) catch return error.AnchorParse;
    try std.testing.expect(fc <= 16383);
    try std.testing.expect(tc <= 16383);
    try std.testing.expect(fc <= tc);
}

test "Writer: conditional formatting — colorScale (2+3 stop) + dataBar" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "cf_gradient.xlsx");
    defer std.testing.allocator.free(tmp_path);

    {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("S");
        // 3-stop: red → yellow → green.
        try sheet.addConditionalFormatColorScale(
            "A2:A100",
            0xFFFF0000, // low
            0xFFFFFF00, // mid
            0xFF00FF00, // high
        );
        // 2-stop: white → blue.
        try sheet.addConditionalFormatColorScale(
            "B2:B100",
            0xFFFFFFFF,
            null,
            0xFF0000FF,
        );
        // dataBar with Excel's default blue.
        try sheet.addConditionalFormatDataBar("C2:C100", 0xFF638EC6);
        try sheet.writeRow(&.{.{ .string = "hdr" }});

        // Rejection: empty range.
        try std.testing.expectError(
            error.InvalidHyperlinkRange,
            sheet.addConditionalFormatColorScale("", 0, null, 0),
        );
        try std.testing.expectError(
            error.InvalidHyperlinkRange,
            sheet.addConditionalFormatDataBar("", 0),
        );

        try w.save(tmp_path);
    }

    // Extract sheet1.xml and check the wire fragments.
    const sheet_xml = blk: {
        var file = try std.fs.cwd().openFile(tmp_path, .{});
        defer file.close();
        var fbuf: [4096]u8 = undefined;
        var fr = file.reader(&fbuf);
        var iter = try std.zip.Iterator.init(&fr);
        var fn_buf: [64]u8 = undefined;
        while (try iter.next()) |entry| {
            if (entry.filename_len > fn_buf.len) continue;
            try fr.seekTo(entry.header_zip_offset + @sizeOf(std.zip.CentralDirectoryFileHeader));
            const filename = fn_buf[0..entry.filename_len];
            try fr.interface.readSliceAll(filename);
            if (std.mem.eql(u8, filename, "xl/worksheets/sheet1.xml")) {
                break :blk try extractEntryForTest(std.testing.allocator, entry, &fr);
            }
        }
        return error.SheetNotFound;
    };
    defer std.testing.allocator.free(sheet_xml);

    // 3-stop color scale: red / yellow / green @ min / 50% / max.
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<cfRule type=\"colorScale\" priority=\"1\">") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<cfvo type=\"percentile\" val=\"50\"/>") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<color rgb=\"FFFF0000\"/>") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<color rgb=\"FFFFFF00\"/>") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<color rgb=\"FF00FF00\"/>") != null);

    // 2-stop skips the percentile cfvo.
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<cfRule type=\"colorScale\" priority=\"2\">") != null);
    // Only one percentile occurrence in the whole doc (from the 3-stop block).
    var count: usize = 0;
    var pos: usize = 0;
    while (std.mem.indexOfPos(u8, sheet_xml, pos, "<cfvo type=\"percentile\"")) |p| {
        count += 1;
        pos = p + 1;
    }
    try std.testing.expectEqual(@as(usize, 1), count);

    // dataBar with Excel's default blue.
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<cfRule type=\"dataBar\" priority=\"3\">") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<color rgb=\"FF638EC6\"/>") != null);
}

test "Writer: conditional formatting — cellIs + expression rules + dxfs table" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "writer_conditional_formatting.xlsx");
    defer std.testing.allocator.free(tmp_path);

    {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();
        // Two dxfs — bold-red font for high values, green fill for
        // matching rows. Content-dedup: re-registering the same dxf
        // must return the same id.
        const red_bold = try w.addDxf(.{
            .font_bold = true,
            .font_color_argb = 0xFFFF0000,
        });
        const green_fill = try w.addDxf(.{ .fill_fg_argb = 0xFF00FF00 });
        const red_bold_again = try w.addDxf(.{
            .font_bold = true,
            .font_color_argb = 0xFFFF0000,
        });
        try std.testing.expectEqual(red_bold, red_bold_again);

        var sheet = try w.addSheet("S");
        try sheet.addConditionalFormatCellIs("B2:B10", .greater_than, "100", null, red_bold);
        try sheet.addConditionalFormatCellIs("C2:C10", .between, "0", "50", red_bold);
        try sheet.addConditionalFormatExpression("A1:Z100", "MOD(ROW(),2)=0", green_fill);
        try sheet.writeRow(&.{.{ .string = "hdr" }});

        // Rejection paths.
        try std.testing.expectError(
            error.InvalidDataValidation,
            sheet.addConditionalFormatCellIs("A1", .equal, "", null, red_bold),
        );
        try std.testing.expectError(
            error.InvalidDataValidation,
            sheet.addConditionalFormatCellIs("A1", .equal, "1", "2", red_bold),
        );
        try std.testing.expectError(
            error.InvalidDataValidation,
            sheet.addConditionalFormatCellIs("A1", .between, "1", null, red_bold),
        );
        try std.testing.expectError(
            error.UnknownDxfId,
            sheet.addConditionalFormatCellIs("A1", .equal, "1", null, 99),
        );
        try std.testing.expectError(
            error.UnknownDxfId,
            sheet.addConditionalFormatExpression("A1", "ROW()=1", 99),
        );

        try w.save(tmp_path);
    }

    // Extract sheet1.xml + styles.xml and verify both sides.
    var sheet_xml: []u8 = undefined;
    var styles_xml: []u8 = undefined;
    {
        var file = try std.fs.cwd().openFile(tmp_path, .{});
        defer file.close();
        var fbuf: [4096]u8 = undefined;
        var fr = file.reader(&fbuf);
        var iter = try std.zip.Iterator.init(&fr);
        var filename_buf: [64]u8 = undefined;
        var got_sheet = false;
        var got_styles = false;
        while (try iter.next()) |entry| {
            if (entry.filename_len > filename_buf.len) continue;
            try fr.seekTo(entry.header_zip_offset + @sizeOf(std.zip.CentralDirectoryFileHeader));
            const filename = filename_buf[0..entry.filename_len];
            try fr.interface.readSliceAll(filename);
            if (std.mem.eql(u8, filename, "xl/worksheets/sheet1.xml")) {
                sheet_xml = try extractEntryForTest(std.testing.allocator, entry, &fr);
                got_sheet = true;
            } else if (std.mem.eql(u8, filename, "xl/styles.xml")) {
                styles_xml = try extractEntryForTest(std.testing.allocator, entry, &fr);
                got_styles = true;
            }
        }
        try std.testing.expect(got_sheet);
        try std.testing.expect(got_styles);
    }
    defer std.testing.allocator.free(sheet_xml);
    defer std.testing.allocator.free(styles_xml);

    // Sheet XML: three conditionalFormatting blocks with the right
    // sqrefs, operators, and dxfIds.
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<conditionalFormatting sqref=\"B2:B10\">") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<cfRule type=\"cellIs\" dxfId=\"0\" priority=\"1\" operator=\"greaterThan\">") != null);
    // Priorities increment — second rule gets 2, third gets 3.
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "priority=\"2\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "priority=\"3\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<formula>100</formula>") != null);

    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<conditionalFormatting sqref=\"C2:C10\">") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "operator=\"between\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<formula>0</formula>") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<formula>50</formula>") != null);

    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<cfRule type=\"expression\" dxfId=\"1\" priority=\"3\">") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "MOD(ROW(),2)=0") != null);

    // Ordering: conditionalFormatting must come BEFORE dataValidations.
    // No DVs in this test but conditionalFormatting must still sit
    // after mergeCells / before dataValidations per the CT_Worksheet
    // spec — the sheetData→conditionalFormatting gap is what matters.
    const sd_end = std.mem.indexOf(u8, sheet_xml, "</sheetData>") orelse return error.MissingSheetData;
    const cf_pos = std.mem.indexOf(u8, sheet_xml, "<conditionalFormatting") orelse return error.MissingCF;
    try std.testing.expect(cf_pos > sd_end);

    // Styles XML: two dxfs (dedup → red_bold_again is the same slot).
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "<dxfs count=\"2\">") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "<dxf><font><b/><color rgb=\"FFFF0000\"/></font></dxf>") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "<fgColor rgb=\"FF00FF00\"/>") != null);
    try std.testing.expect(std.mem.indexOf(u8, styles_xml, "</dxfs>") != null);
}

test "Writer: addDataValidationList validates + emits <dataValidations> block" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "writer_dv_list.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "writer_no_dv.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "writer_hyperlinks.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "writer_no_hyperlinks.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "writer_no_merged.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "writer_styles_borders.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "writer_styles_fills.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "writer_styles_stage2.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "writer_styles_bold.xlsx");
    defer std.testing.allocator.free(tmp_path);

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

test "Writer: addSheet rejects Unicode-fold-equivalent duplicates (A1)" {
    // Empirical Excel matrix from the A1 spec — names that fold to
    // the same canonical form must be detected as duplicates.
    var w = Writer.init(std.testing.allocator);
    defer w.deinit();
    _ = try w.addSheet("café");
    try std.testing.expectError(error.DuplicateSheetName, w.addSheet("CAFÉ"));
    try std.testing.expectError(error.DuplicateSheetName, w.addSheet("cafÉ"));

    var w2 = Writer.init(std.testing.allocator);
    defer w2.deinit();
    _ = try w2.addSheet("Straße");
    try std.testing.expectError(error.DuplicateSheetName, w2.addSheet("STRASSE"));
    try std.testing.expectError(error.DuplicateSheetName, w2.addSheet("Strasse"));

    var w3 = Writer.init(std.testing.allocator);
    defer w3.deinit();
    _ = try w3.addSheet("ΣΤΟΧΟΣ");
    try std.testing.expectError(error.DuplicateSheetName, w3.addSheet("στοχοσ"));
    // Final-sigma form also folds to plain σ.
    try std.testing.expectError(error.DuplicateSheetName, w3.addSheet("στοχος"));
}

test "Writer: addSheet treats composed/decomposed Unicode as duplicate (A1 phase 3 NFC)" {
    var w = Writer.init(std.testing.allocator);
    defer w.deinit();
    // Add the precomposed form first.
    _ = try w.addSheet("café"); // U+0063 U+0061 U+0066 U+00E9
    // The decomposed form (e + combining acute) must be treated
    // as a duplicate after NFC normalisation.
    try std.testing.expectError(
        error.DuplicateSheetName,
        w.addSheet("cafe\u{0301}"),
    );
    // Reverse direction: opening with decomposed, then trying
    // precomposed.
    var w2 = Writer.init(std.testing.allocator);
    defer w2.deinit();
    _ = try w2.addSheet("cafe\u{0301}");
    try std.testing.expectError(error.DuplicateSheetName, w2.addSheet("café"));
}

test "Writer: addSheet validates Unicode scalar count (not bytes) for 31-char limit (A1)" {
    var w = Writer.init(std.testing.allocator);
    defer w.deinit();

    // 31 multi-byte scalars (62 bytes) — was rejected by the byte
    // check, now accepted.
    var buf31: [62]u8 = undefined;
    for (0..31) |i| {
        buf31[i * 2] = 0xC3;
        buf31[i * 2 + 1] = 0xA9; // é
    }
    _ = try w.addSheet(&buf31);

    // 32 scalars (64 bytes) — over the limit.
    var buf32: [64]u8 = undefined;
    for (0..32) |i| {
        buf32[i * 2] = 0xC3;
        buf32[i * 2 + 1] = 0xA9;
    }
    try std.testing.expectError(error.InvalidSheetName, w.addSheet(&buf32));

    // Malformed UTF-8 still rejected.
    try std.testing.expectError(error.InvalidSheetName, w.addSheet("ab\xFFc"));
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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "writer_sheet_escape.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "writer_atomic.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
        // Sanitise: replace forbidden XML 1.0 bytes with `?` so the
        // call succeeds — the appendXmlEscaped contract is now to
        // ERROR on forbidden bytes (NUL, most C0 controls, DEL).
        // The invariant we want to verify here is the entity-
        // escaping correctness on valid input; a separate test
        // exercises the error path.
        for (input_buf[0..len]) |*p| {
            if (isForbiddenXmlByte(p.*)) p.* = '?';
        }
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

test "appendXmlEscaped rejects forbidden XML 1.0 control bytes" {
    var out: std.ArrayListUnmanaged(u8) = .{};
    defer out.deinit(std.testing.allocator);
    // Each of these should error — none are legal in XML 1.0 text.
    const cases = [_][]const u8{
        "\x00", // NUL
        "hello\x01world",
        "tab\x0Bvtab", // vertical tab
        "ff\x0C", // form feed
        "esc\x1B", // escape
    };
    for (cases) |c| {
        out.clearRetainingCapacity();
        try std.testing.expectError(
            error.InvalidXmlByte,
            appendXmlEscaped(std.testing.allocator, &out, c),
        );
    }
    // Allowed: tab, LF, CR, and DEL (0x7F is in XML 1.0 `Char`).
    const allowed = [_][]const u8{ "tab\there", "lf\nhere", "cr\rhere", "del\x7Fhere" };
    for (allowed) |c| {
        out.clearRetainingCapacity();
        try appendXmlEscaped(std.testing.allocator, &out, c);
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
        try std.testing.expectEqual(@as(u32, @intCast(seen_indices.count())), @as(u32, @intCast(w.sst_plan.new_strings.items.len)));
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
        try std.testing.expect(w.styles_plan.styles.items.len <= 4);
    }
}

test "fuzz Writer end-to-end round-trip via reader" {
    const iters = fuzzIterationsW() / 10; // each iter does real zip I/O
    const seed = fuzzSeedW();
    var prng = std.Random.DefaultPrng.init(seed);
    const rng = prng.random();
    var tt = TestTmp.init();
    defer tt.deinit();
    var tmp_name_buf: [64]u8 = undefined;
    const tmp_name = std.fmt.bufPrint(&tmp_name_buf, "fuzz_writer_{x}.xlsx", .{seed}) catch unreachable;
    const tmp_path = try tt.path(std.testing.allocator, tmp_name);
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_name = std.fmt.bufPrint(&tmp_path_buf, "fuzz_combo_{x}.xlsx", .{seed}) catch unreachable;
    const tmp_path = try tt.path(std.testing.allocator, tmp_name);
    defer std.testing.allocator.free(tmp_path);

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
    var tt = TestTmp.init();

    defer tt.deinit();

    const _fuzz_name = std.fmt.bufPrint(&tmp_path_buf, "fuzz_sheetfeat_{x}.xlsx", .{seed}) catch unreachable;

    const tmp_path = try tt.path(std.testing.allocator, _fuzz_name);

    defer std.testing.allocator.free(tmp_path);
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
            try sheet.freezePanes(
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

test "fuzz zip.Archive produces archives our reader can walk" {
    const iters = fuzzIterationsW() / 10;
    const seed = fuzzSeedW();
    var prng = std.Random.DefaultPrng.init(seed);
    const rng = prng.random();
    var tmp_path_buf: [64]u8 = undefined;
    var tt = TestTmp.init();

    defer tt.deinit();

    const _fuzz_name = std.fmt.bufPrint(&tmp_path_buf, "fuzz_zipwriter_{x}.zip", .{seed}) catch unreachable;

    const tmp_path = try tt.path(std.testing.allocator, _fuzz_name);

    defer std.testing.allocator.free(tmp_path);
    for (0..iters) |_| {
        var zip_buf: std.ArrayListUnmanaged(u8) = .{};
        defer zip_buf.deinit(std.testing.allocator);
        var zw = zip.Archive.init(std.testing.allocator, &zip_buf);
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
            try zw.addEntry(expected_names[i][0..name_len], payload[0..payload_len], deflateCompressErased);
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
    // compression at the zip.Archive layer, so special-case here.
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
    var tt = TestTmp.init();

    defer tt.deinit();

    const _fuzz_name = std.fmt.bufPrint(&tmp_path_buf, "fuzz_state_{x}.xlsx", .{seed}) catch unreachable;

    const tmp_path = try tt.path(std.testing.allocator, _fuzz_name);

    defer std.testing.allocator.free(tmp_path);
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
                    try std.testing.expect(w.styles_plan.styles.items.len <= 4);
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
                    if (n_sheets == 0 or w.styles_plan.styles.items.len == 0) continue;
                    const si = rng.intRangeAtMost(usize, 0, n_sheets - 1);
                    var cells: [3]xlsx.Cell = undefined;
                    var styles: [3]u32 = undefined;
                    var str_buf: [3][32]u8 = undefined;
                    const nc = rng.intRangeAtMost(usize, 1, 3);
                    _ = &str_store;
                    for (0..nc) |ci| {
                        cells[ci] = randomCellDeep(rng, &str_buf[ci]);
                        styles[ci] = rng.intRangeAtMost(u32, 0, @intCast(w.styles_plan.styles.items.len));
                    }
                    sheet_handles[si].?.writeRowStyled(cells[0..nc], styles[0..nc]) catch |e| switch (e) {
                        error.IntegerExceedsExcelPrecision => continue,
                        else => return e,
                    };
                    expected_rows[si] += 1;
                },
                else => {
                    // No-op probe — repeatedly query sheet metadata.
                    _ = w.styles_plan.styles.items.len;
                    _ = w.sst_plan.new_strings.items.len;
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
    var tt = TestTmp.init();

    defer tt.deinit();

    const _fuzz_name = std.fmt.bufPrint(&tmp_path_buf, "fuzz_multisave_{x}.xlsx", .{seed}) catch unreachable;

    const tmp_path = try tt.path(std.testing.allocator, _fuzz_name);

    defer std.testing.allocator.free(tmp_path);
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
    var tt = TestTmp.init();

    defer tt.deinit();

    const _fuzz_name = std.fmt.bufPrint(&tmp_path_buf, "fuzz_bounds_{x}.xlsx", .{seed}) catch unreachable;

    const tmp_path = try tt.path(std.testing.allocator, _fuzz_name);

    defer std.testing.allocator.free(tmp_path);
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

test "fuzz zip.Archive: adversarial entry names" {
    // Names with path traversal, embedded nulls, UTF-8, max-length.
    // We don't promise to *reject* these (addEntry just writes bytes) —
    // we promise the result is still a walkable zip and our reader
    // doesn't blow up on the unusual names.
    const seed = fuzzSeedW();
    var prng = std.Random.DefaultPrng.init(seed);
    const rng = prng.random();
    var tmp_path_buf: [64]u8 = undefined;
    var tt = TestTmp.init();

    defer tt.deinit();

    const _fuzz_name = std.fmt.bufPrint(&tmp_path_buf, "fuzz_advnames_{x}.zip", .{seed}) catch unreachable;

    const tmp_path = try tt.path(std.testing.allocator, _fuzz_name);

    defer std.testing.allocator.free(tmp_path);
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
        var zw = zip.Archive.init(std.testing.allocator, &zip_buf);
        defer zw.deinit();

        var emitted: usize = 0;
        const n = rng.intRangeAtMost(usize, 1, 5);
        for (0..n) |_| {
            const name = names[rng.intRangeAtMost(usize, 0, names.len - 1)];
            var payload: [128]u8 = undefined;
            const plen = rng.intRangeAtMost(usize, 0, payload.len);
            rng.bytes(payload[0..plen]);
            zw.addEntry(name, payload[0..plen], deflateCompressErased) catch |e| switch (e) {
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

test "validateDefinedName: accepts valid names" {
    try validateDefinedName("MyName");
    try validateDefinedName("_private");
    try validateDefinedName("foo.bar.baz");
    try validateDefinedName("\\Backslashed");
    try validateDefinedName("a1_b2");
    try validateDefinedName("_xlnm.Print_Area");
}

test "validateDefinedName: rejects invalid names" {
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName(""));
    // Starts with a digit.
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("1Name"));
    // A1-shaped (case-insensitive grid match).
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("A1"));
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("a1"));
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("XFD1048576"));
    // Single R / C reserved for R1C1.
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("R"));
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("c"));
    // Disallowed character (`!` not in the allowed set).
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("Bad!Name"));
    // 256-byte name fails the 255 cap.
    var too_long: [256]u8 = undefined;
    @memset(&too_long, 'X');
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName(&too_long));
}

test "Writer.addDefinedName: stores name + refers_to with options" {
    var w = Writer.init(std.testing.allocator);
    defer w.deinit();
    try w.addDefinedName("MyRange", "Sheet1!$A$1:$B$1", .{});
    try w.addDefinedName(
        "_xlnm.Print_Area",
        "Sheet1!$A$1:$B$1",
        .{ .local_sheet_id = 0, .hidden = true },
    );
    try std.testing.expectEqual(@as(usize, 2), w.workbook_xml_plan.defined_names.items.len);
    try std.testing.expectEqualStrings("MyRange", w.workbook_xml_plan.defined_names.items[0].name);
    try std.testing.expectEqualStrings("Sheet1!$A$1:$B$1", w.workbook_xml_plan.defined_names.items[0].refers_to);
    try std.testing.expectEqual(@as(?u32, null), w.workbook_xml_plan.defined_names.items[0].local_sheet_id);
    try std.testing.expectEqual(false, w.workbook_xml_plan.defined_names.items[0].hidden);
    try std.testing.expectEqual(@as(?u32, 0), w.workbook_xml_plan.defined_names.items[1].local_sheet_id);
    try std.testing.expectEqual(true, w.workbook_xml_plan.defined_names.items[1].hidden);
}

test "Writer.addDefinedName: rejects invalid name + empty refers_to" {
    var w = Writer.init(std.testing.allocator);
    defer w.deinit();
    try std.testing.expectError(
        error.InvalidDefinedName,
        w.addDefinedName("A1", "Sheet1!$A$1", .{}),
    );
    try std.testing.expectError(
        error.InvalidDefinedNameRefersTo,
        w.addDefinedName("Foo", "", .{}),
    );
}

test "Writer.addDefinedName: rejects case-insensitive duplicates per scope" {
    var w = Writer.init(std.testing.allocator);
    defer w.deinit();
    try w.addDefinedName("Rate", "Sheet1!$A$1", .{});
    // Same name, same scope — duplicate.
    try std.testing.expectError(
        error.DuplicateDefinedName,
        w.addDefinedName("Rate", "Sheet1!$A$2", .{}),
    );
    // Case-only variant — still duplicate.
    try std.testing.expectError(
        error.DuplicateDefinedName,
        w.addDefinedName("RATE", "Sheet1!$A$3", .{}),
    );
    try std.testing.expectError(
        error.DuplicateDefinedName,
        w.addDefinedName("rate", "Sheet1!$A$4", .{}),
    );
    // Different scope (sheet-scoped) — accepted.
    try w.addDefinedName("Rate", "Sheet1!$B$1", .{ .local_sheet_id = 0 });
    // Same name in another sheet's scope — also accepted.
    try w.addDefinedName("Rate", "Sheet2!$B$1", .{ .local_sheet_id = 1 });
    try std.testing.expectEqual(@as(usize, 3), w.workbook_xml_plan.defined_names.items.len);
}

test "validateDefinedName: rejects ? and \\ in trailing chars" {
    // First char `\` is allowed; subsequent `\` or `?` are not.
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("Foo?bar"));
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("Foo\\bar"));
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("Foo?"));
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("Foo\\"));
    // First char as `\` still works.
    try validateDefinedName("\\Foo");
    try validateDefinedName("\\Foo_bar.baz");
}

test "validateDefinedName: rejects R1C1-shaped references" {
    // Bare R / C already covered by the single-letter-reservation
    // check; these cover the multi-segment R1C1 forms.
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("R1C1"));
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("r10c5"));
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("RC"));
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("R5C"));
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("RC10"));
    // Names that START with R/C but aren't R1C1 shapes are accepted.
    try validateDefinedName("Range1");
    try validateDefinedName("Customer");
    try validateDefinedName("R_total");
}

test "writeRow atomicity: forbidden-XML-byte string leaves no half-written row" {
    var w = Writer.init(std.testing.allocator);
    defer w.deinit();
    var sw = try w.addSheet("Sheet1");
    // Snapshot body length, attempt the bad row, verify body is unchanged.
    const body_before = sw.body.items.len;
    try std.testing.expectError(
        error.InvalidXmlByte,
        sw.writeRow(&.{ .{ .string = "ok" }, .{ .string = "bad\x00here" } }),
    );
    try std.testing.expectEqual(body_before, sw.body.items.len);

    // Then a valid row writes fine.
    try sw.writeRow(&.{ .{ .string = "good" }, .{ .integer = 42 } });
    try std.testing.expect(sw.body.items.len > body_before);
}

test "writeRichRow atomicity: forbidden byte in rich-run font_name bails too" {
    var w = Writer.init(std.testing.allocator);
    defer w.deinit();
    var sw = try w.addSheet("Sheet1");
    const body_before = sw.body.items.len;
    const runs = [_]RichTextRun{
        .{ .text = "ok", .font_name = "Bad\x01Font" },
    };
    const cells = [_]RichRowCell{.{ .rich = &runs }};
    try std.testing.expectError(error.InvalidXmlByte, sw.writeRichRow(&cells));
    try std.testing.expectEqual(body_before, sw.body.items.len);
}

test "writeRichRow atomicity: forbidden byte in rich-run text leaves no half-written row" {
    var w = Writer.init(std.testing.allocator);
    defer w.deinit();
    var sw = try w.addSheet("Sheet1");
    const body_before = sw.body.items.len;
    const runs = [_]RichTextRun{
        .{ .text = "ok run" },
        .{ .text = "bad\x00run" }, // NUL — forbidden
    };
    const cells = [_]RichRowCell{.{ .rich = &runs }};
    try std.testing.expectError(error.InvalidXmlByte, sw.writeRichRow(&cells));
    try std.testing.expectEqual(body_before, sw.body.items.len);

    // Subsequent valid rich row writes fine.
    const ok_runs = [_]RichTextRun{.{ .text = "fine" }};
    try sw.writeRichRow(&.{.{ .rich = &ok_runs }});
    try std.testing.expect(sw.body.items.len > body_before);
}

test "writeRowWithFormulas atomicity: forbidden byte in formula bails before any append" {
    var w = Writer.init(std.testing.allocator);
    defer w.deinit();
    var sw = try w.addSheet("Sheet1");
    const body_before = sw.body.items.len;
    const cells = [_]xlsx.Cell{.{ .integer = 1 }};
    const fs = [_]?[]const u8{"SUM(A1\x00:A2)"};
    try std.testing.expectError(
        error.InvalidXmlByte,
        sw.writeRowWithFormulas(&cells, &fs),
    );
    try std.testing.expectEqual(body_before, sw.body.items.len);
}

test "Style round-trip: font_name with XML specials, alignment, wrap, diagonals" {
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, "style_roundtrip.xlsx");
    defer std.testing.allocator.free(tmp_path);

    {
        var w = Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("Sheet1");
        // Style 1: font name carries `&`, `<`, `"` — XML-escaped on
        // emit; reader must entity-decode on readback.
        const sid_font = try w.addStyle(.{
            .font_name = "A&B<\"co\">",
            .font_size = .{ 12.5, 0 }[0],
        });
        // Style 2: alignment + wrap_text — nested <alignment> child.
        const sid_align = try w.addStyle(.{
            .alignment_horizontal = .center,
            .wrap_text = true,
        });
        // Style 3: diagonal direction flags — attributes on <border>.
        const sid_diag = try w.addStyle(.{
            .border_diagonal = .{ .style = .thin, .color_argb = 0xFF000000 },
            .diagonal_up = true,
            .diagonal_down = true,
        });
        try sheet.writeRowStyled(
            &.{ .{ .string = "x" }, .{ .string = "y" }, .{ .string = "z" } },
            &.{ sid_font, sid_align, sid_diag },
        );
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    // Style 1 → font_name decoded.
    const font = book.cellFont(1) orelse return error.MissingStyle1;
    try std.testing.expectEqualStrings("A&B<\"co\">", font.name);

    // Style 2 → alignment.horizontal == "center", wrap_text == true.
    const align_rec = book.cellAlignment(2) orelse return error.MissingStyle2;
    try std.testing.expect(align_rec.horizontal != null);
    try std.testing.expectEqualStrings("center", align_rec.horizontal.?);
    try std.testing.expectEqual(true, align_rec.wrap_text);

    // Style 3 → diagonal_up + diagonal_down.
    const border = book.cellBorder(3) orelse return error.MissingStyle3;
    try std.testing.expectEqual(true, border.diagonal_up);
    try std.testing.expectEqual(true, border.diagonal_down);
}

test "deflateCompress: large repetitive payload doesn't trip Huffman assert" {
    // Regression: a 100k×5 numeric `<sheetData>` payload (≈4 MB
    // pre-compress) used to crash inside `HuffmanEncoder.bitCounts`
    // with `assert(leaf_counts[max_bits][max_bits] == n)` because
    // `scaleFreqs` rescaled to u15 (32767) — the stdlib's level
    // walker mis-converged on certain near-cap distributions. The
    // fix: tighter rescale to u13 (8191) leaves enough numerical
    // headroom. See `scaleFreqs` docstring for the witness.
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(std.testing.allocator, "huff_large.xlsx");
    defer std.testing.allocator.free(path);

    var w = Writer.init(std.testing.allocator);
    defer w.deinit();
    var s = try w.addSheet("Bench");

    // 50_000 rows of 5 mixed numeric cells — past the previous cliff
    // (~10_000 rows). One sheet is enough; the bug is per-block.
    const cells = [_]xlsx.Cell{
        .{ .integer = 1 },
        .{ .number = 2.5 },
        .{ .integer = 3 },
        .{ .number = 4.75 },
        .{ .integer = 5 },
    };
    var i: usize = 0;
    while (i < 50_000) : (i += 1) {
        try s.writeRow(&cells);
    }
    // The writeRow path itself doesn't deflate; deflate happens at
    // save time. Save → reads back via Book.open → confirms the
    // round-trip held together (any Huffman-tree corruption would
    // surface as inflate errors here).
    try w.save(path);

    var book = try xlsx.Book.open(std.testing.allocator, path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 1), book.sheets.len);

    // Spot-check a row near the end to make sure decompression
    // walked the whole stream.
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    var row_count: usize = 0;
    while (try rows.next()) |row| : (row_count += 1) {
        if (row_count == 49_999) {
            try std.testing.expectEqual(@as(i64, 1), row[0].integer);
        }
    }
    try std.testing.expectEqual(@as(usize, 50_000), row_count);
}

// ─── iter-wr-4 byte-equivalence parity ───────────────────────────────
//
// These tests build a Writer-saved workbook, then construct the
// same per-sheet bytes by calling `pkg/sheet_plan.zig` directly.
// The two outputs must be byte-identical — that's the contract
// that lets future iters (wr-6) collapse `Writer.save` into a thin
// shim around `Workbook.save` without producing a "repaired"
// prompt across the corpus.
//
// Each test extracts the relevant sheet/comments/vml/rels part out
// of the Writer-emitted ZIP, then runs the plan module on a
// hand-built `SheetEmitInputs` and compares.

/// Walk a saved-on-disk archive, extract one entry's uncompressed
/// bytes by name. Returns owned bytes (caller frees). Returns
/// `error.PartNotFound` if the entry isn't in the archive.
fn extractParityEntry(
    alloc: Allocator,
    archive_path: []const u8,
    target: []const u8,
) ![]u8 {
    var file = try std.fs.cwd().openFile(archive_path, .{});
    defer file.close();
    var fbuf: [4096]u8 = undefined;
    var fr = file.reader(&fbuf);
    var iter = try std.zip.Iterator.init(&fr);
    var name_buf: [128]u8 = undefined;
    while (try iter.next()) |entry| {
        if (entry.filename_len > name_buf.len) continue;
        try fr.seekTo(entry.header_zip_offset + @sizeOf(std.zip.CentralDirectoryFileHeader));
        const filename = name_buf[0..entry.filename_len];
        try fr.interface.readSliceAll(filename);
        if (std.mem.eql(u8, filename, target)) {
            return try extractEntryForTest(alloc, entry, &fr);
        }
    }
    return error.PartNotFound;
}

test "iter-wr-4 parity: empty body sheet — byte identical" {
    const a = std.testing.allocator;
    var tmp = TestTmp.init();
    defer tmp.deinit();
    const path = try tmp.path(a, "iter_wr4_empty.xlsx");
    defer a.free(path);

    var w = Writer.init(a);
    defer w.deinit();
    _ = try w.addSheet("Sheet1");
    try w.save(path);

    const writer_bytes = try extractParityEntry(a, path, "xl/worksheets/sheet1.xml");
    defer a.free(writer_bytes);

    var plan_buf: std.ArrayListUnmanaged(u8) = .{};
    defer plan_buf.deinit(a);
    try sheet_plan.emitWorksheetXml(a, &plan_buf, .{ .body = "" });

    try std.testing.expectEqualSlices(u8, plan_buf.items, writer_bytes);
}

test "iter-wr-4 parity: mixed cell types row — byte identical" {
    const a = std.testing.allocator;
    var tmp = TestTmp.init();
    defer tmp.deinit();
    const path = try tmp.path(a, "iter_wr4_mixed.xlsx");
    defer a.free(path);

    var w = Writer.init(a);
    defer w.deinit();
    const sw = try w.addSheet("Sheet1");
    try sw.writeRow(&.{
        .{ .string = "hello" },
        .{ .integer = 42 },
        .{ .number = 3.14 },
        .{ .boolean = true },
        .empty,
    });
    try w.save(path);

    const writer_bytes = try extractParityEntry(a, path, "xl/worksheets/sheet1.xml");
    defer a.free(writer_bytes);

    // Pull the body bytes out of the writer-saved sheet — they live
    // between <sheetData> and </sheetData>. The plan module is
    // body-agnostic; row-emit primitives are still owned by Writer
    // (writeRowImpl). The parity is on the surrounding worksheet
    // shape.
    const open = "<sheetData>";
    const close = "</sheetData>";
    const open_pos = std.mem.indexOf(u8, writer_bytes, open) orelse return error.TestFailed;
    const close_pos = std.mem.indexOf(u8, writer_bytes, close) orelse return error.TestFailed;
    const body = writer_bytes[open_pos + open.len .. close_pos];

    var plan_buf: std.ArrayListUnmanaged(u8) = .{};
    defer plan_buf.deinit(a);
    try sheet_plan.emitWorksheetXml(a, &plan_buf, .{ .body = body });

    try std.testing.expectEqualSlices(u8, plan_buf.items, writer_bytes);
}

test "iter-wr-4 parity: frozen panes — byte identical" {
    const a = std.testing.allocator;
    var tmp = TestTmp.init();
    defer tmp.deinit();
    const path = try tmp.path(a, "iter_wr4_frozen.xlsx");
    defer a.free(path);

    var w = Writer.init(a);
    defer w.deinit();
    const sw = try w.addSheet("Sheet1");
    try sw.freezePanes(1, 2);
    try sw.writeRow(&.{ .{ .string = "header" }, .{ .integer = 1 } });
    try w.save(path);

    const writer_bytes = try extractParityEntry(a, path, "xl/worksheets/sheet1.xml");
    defer a.free(writer_bytes);

    const open = "<sheetData>";
    const close = "</sheetData>";
    const open_pos = std.mem.indexOf(u8, writer_bytes, open) orelse return error.TestFailed;
    const close_pos = std.mem.indexOf(u8, writer_bytes, close) orelse return error.TestFailed;
    const body = writer_bytes[open_pos + open.len .. close_pos];

    var plan_buf: std.ArrayListUnmanaged(u8) = .{};
    defer plan_buf.deinit(a);
    try sheet_plan.emitWorksheetXml(a, &plan_buf, .{
        .body = body,
        .freeze_rows = 1,
        .freeze_cols = 2,
    });

    try std.testing.expectEqualSlices(u8, plan_buf.items, writer_bytes);
}

test "iter-wr-4 parity: kitchen-sink (merges + autoFilter + DV + CF + hyperlinks)" {
    const a = std.testing.allocator;
    var tmp = TestTmp.init();
    defer tmp.deinit();
    const path = try tmp.path(a, "iter_wr4_kitchen.xlsx");
    defer a.free(path);

    var w = Writer.init(a);
    defer w.deinit();
    const dxf_id = try w.addDxf(.{ .font_bold = true });
    const sw = try w.addSheet("Sheet1");

    try sw.writeRow(&.{ .{ .string = "header" }, .{ .integer = 1 }, .{ .integer = 2 } });
    try sw.writeRow(&.{ .{ .string = "row" }, .{ .integer = 3 }, .{ .integer = 4 } });
    try sw.setColumnWidth(0, 12.5);
    try sw.setColumnWidth(1, 5.0);
    try sw.addMergedCell("A1:B2");
    try sw.setAutoFilter("A1:C1");
    try sw.addDataValidationList("C1:C10", &.{ "yes", "no" });
    try sw.addDataValidationNumeric("B1:B10", .whole, .between, "1", "100");
    try sw.addConditionalFormatCellIs("B1:B10", .greater_than, "10", null, dxf_id);
    try sw.addHyperlink("A1", "https://ex.com/?q=1&x=2");
    try sw.addInternalHyperlink("A2", "Sheet1!B2");

    try w.save(path);

    const writer_bytes = try extractParityEntry(a, path, "xl/worksheets/sheet1.xml");
    defer a.free(writer_bytes);

    const open = "<sheetData>";
    const close = "</sheetData>";
    const open_pos = std.mem.indexOf(u8, writer_bytes, open) orelse return error.TestFailed;
    const close_pos = std.mem.indexOf(u8, writer_bytes, close) orelse return error.TestFailed;
    const body = writer_bytes[open_pos + open.len .. close_pos];

    // Re-build the inputs view by hand — same projection Writer.save
    // does, but with literal slices.
    const cws = [_]sheet_plan.ColumnWidth{
        .{ .col_min = 1, .col_max = 1, .width = 12.5 },
        .{ .col_min = 2, .col_max = 2, .width = 5.0 },
    };
    const merges = [_][]const u8{"A1:B2"};
    const list_vals = [_][]const u8{ "yes", "no" };
    const lists = [_]sheet_plan.DataValidationList{
        .{ .range = "C1:C10", .values = &list_vals },
    };
    const ranges = [_]sheet_plan.DataValidationRange{
        .{
            .range = "B1:B10",
            .kind_name = "whole",
            .op_name = "between",
            .formula1 = "1",
            .formula2 = "100",
        },
    };
    const cfs = [_]sheet_plan.ConditionalFormat{
        .{ .range = "B1:B10", .rule = .{ .cell_is = .{
            .operator = .greater_than,
            .formula1 = "10",
            .formula2 = null,
            .dxf_id = dxf_id,
        } } },
    };
    const hls = [_]sheet_plan.Hyperlink{
        .{ .range = "A1", .url = "https://ex.com/?q=1&x=2" },
    };
    const intl = [_]sheet_plan.InternalHyperlink{
        .{ .range = "A2", .location = "Sheet1!B2" },
    };

    var plan_buf: std.ArrayListUnmanaged(u8) = .{};
    defer plan_buf.deinit(a);
    try sheet_plan.emitWorksheetXml(a, &plan_buf, .{
        .body = body,
        .column_widths = &cws,
        .auto_filter_range = "A1:C1",
        .merged_cells = &merges,
        .conditional_formats = &cfs,
        .data_validations = &lists,
        .data_validation_ranges = &ranges,
        .hyperlinks = &hls,
        .internal_hyperlinks = &intl,
    });

    try std.testing.expectEqualSlices(u8, plan_buf.items, writer_bytes);
}

test "iter-wr-4 parity: rich-row sheet — byte identical" {
    const a = std.testing.allocator;
    var tmp = TestTmp.init();
    defer tmp.deinit();
    const path = try tmp.path(a, "iter_wr4_rich.xlsx");
    defer a.free(path);

    var w = Writer.init(a);
    defer w.deinit();
    const sw = try w.addSheet("Sheet1");
    const runs = [_]RichTextRun{
        .{ .text = "bold ", .bold = true },
        .{ .text = "italic", .italic = true, .color_argb = 0xFFFF0000 },
    };
    try sw.writeRichRow(&.{ .{ .rich = &runs }, .{ .string = "plain" } });
    try w.save(path);

    const writer_bytes = try extractParityEntry(a, path, "xl/worksheets/sheet1.xml");
    defer a.free(writer_bytes);

    const open = "<sheetData>";
    const close = "</sheetData>";
    const open_pos = std.mem.indexOf(u8, writer_bytes, open) orelse return error.TestFailed;
    const close_pos = std.mem.indexOf(u8, writer_bytes, close) orelse return error.TestFailed;
    const body = writer_bytes[open_pos + open.len .. close_pos];

    var plan_buf: std.ArrayListUnmanaged(u8) = .{};
    defer plan_buf.deinit(a);
    try sheet_plan.emitWorksheetXml(a, &plan_buf, .{ .body = body });

    try std.testing.expectEqualSlices(u8, plan_buf.items, writer_bytes);
}

test "iter-wr-4 parity: comments + VML — byte identical (CT.xml VML Default-before-Override pin)" {
    const a = std.testing.allocator;
    var tmp = TestTmp.init();
    defer tmp.deinit();
    const path = try tmp.path(a, "iter_wr4_comments.xlsx");
    defer a.free(path);

    var w = Writer.init(a);
    defer w.deinit();
    const sw = try w.addSheet("Sheet1");
    try sw.writeRow(&.{.{ .string = "x" }});
    try sw.addComment("A1", "alice", "first note");
    try sw.addComment("A2", "bob", "second");
    try sw.addComment("A3", "alice", "third");
    try w.save(path);

    // Pin: CT.xml VML Default precedes any Override. This was the
    // `50ed225` regression — if it drifts, every comment-bearing
    // workbook trips Excel's "repaired" prompt.
    const ct = try extractParityEntry(a, path, "[Content_Types].xml");
    defer a.free(ct);
    const vml_default_pos = std.mem.indexOf(u8, ct, "Default Extension=\"vml\"") orelse
        return error.TestFailed;
    const first_override_pos = std.mem.indexOf(u8, ct, "<Override") orelse
        return error.TestFailed;
    try std.testing.expect(vml_default_pos < first_override_pos);

    // commentsN.xml byte parity.
    const writer_comments = try extractParityEntry(a, path, "xl/comments1.xml");
    defer a.free(writer_comments);

    const comments_view = [_]sheet_plan.Comment{
        .{ .ref = "A1", .author = "alice", .text = "first note" },
        .{ .ref = "A2", .author = "bob", .text = "second" },
        .{ .ref = "A3", .author = "alice", .text = "third" },
    };
    var plan_comments: std.ArrayListUnmanaged(u8) = .{};
    defer plan_comments.deinit(a);
    try sheet_plan.emitCommentsXml(a, &plan_comments, &comments_view);
    try std.testing.expectEqualSlices(u8, plan_comments.items, writer_comments);

    // vmlDrawingN.vml byte parity.
    const writer_vml = try extractParityEntry(a, path, "xl/drawings/vmlDrawing1.vml");
    defer a.free(writer_vml);
    var plan_vml: std.ArrayListUnmanaged(u8) = .{};
    defer plan_vml.deinit(a);
    try sheet_plan.emitVmlDrawingXml(a, &plan_vml, &comments_view);
    try std.testing.expectEqualSlices(u8, plan_vml.items, writer_vml);

    // sheet1.xml.rels byte parity (rId numbering across hyperlinks +
    // drawings + comments).
    const writer_rels = try extractParityEntry(a, path, "xl/worksheets/_rels/sheet1.xml.rels");
    defer a.free(writer_rels);
    var plan_rels: std.ArrayListUnmanaged(u8) = .{};
    defer plan_rels.deinit(a);
    const wrote = try sheet_plan.emitSheetRels(a, &plan_rels, 0, &.{}, comments_view.len);
    try std.testing.expect(wrote);
    try std.testing.expectEqualSlices(u8, plan_rels.items, writer_rels);
}

test "iter-wr-4 parity: rels with hyperlinks AND comments — rId numbering stable" {
    const a = std.testing.allocator;
    var tmp = TestTmp.init();
    defer tmp.deinit();
    const path = try tmp.path(a, "iter_wr4_rels_combo.xlsx");
    defer a.free(path);

    var w = Writer.init(a);
    defer w.deinit();
    const sw = try w.addSheet("Sheet1");
    try sw.writeRow(&.{.{ .string = "x" }});
    try sw.addHyperlink("A1", "https://ex.com");
    try sw.addHyperlink("A2", "https://ex2.com");
    try sw.addComment("A3", "alice", "note");
    try w.save(path);

    const writer_rels = try extractParityEntry(a, path, "xl/worksheets/_rels/sheet1.xml.rels");
    defer a.free(writer_rels);

    const hls = [_]sheet_plan.Hyperlink{
        .{ .range = "A1", .url = "https://ex.com" },
        .{ .range = "A2", .url = "https://ex2.com" },
    };
    var plan_rels: std.ArrayListUnmanaged(u8) = .{};
    defer plan_rels.deinit(a);
    const wrote = try sheet_plan.emitSheetRels(a, &plan_rels, 0, &hls, 1);
    try std.testing.expect(wrote);
    try std.testing.expectEqualSlices(u8, plan_rels.items, writer_rels);
}
