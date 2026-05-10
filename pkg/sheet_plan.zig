//! Per-sheet fresh-emit plan (B3 iter-wr-4).
//!
//! Storage + fresh-emit primitives for `xl/worksheets/sheetN.xml`,
//! `xl/worksheets/_rels/sheetN.xml.rels`, `xl/commentsN.xml`, and
//! `xl/drawings/vmlDrawingN.vml`. Lives outside `pkg/workbook.zig`
//! and `src/writer.zig` so both Workbook (delta-on-existing-bytes
//! editor, future-fresh producer) and `xlsx.Writer` (fresh-emit
//! producer) can stage per-sheet state through the same shape
//! without a circular module dependency (workbook.zig imports
//! `zlsx`, which contains writer.zig).
//!
//! Today this module owns:
//!   - opaque state types mirroring `SheetWriter`'s fields exactly:
//!     `Hyperlink`, `InternalHyperlink`, `Comment`, `ColumnWidth`,
//!     `ConditionalFormat`, `DataValidationList`, `DataValidationRange`.
//!   - the row-emit byte builder for `<sheetData>` rows
//!     (`emitRowOpener`, `emitCellRef`, `appendXmlEscaped`) — preserved
//!     byte-for-byte from `Writer.writeRowImpl` / `Writer.writeRichRow`.
//!   - the per-sheet `<worksheet>` body emitter
//!     (`emitWorksheetXml`) which renders the CT_Worksheet element
//!     order: `sheetViews → cols → sheetData → autoFilter →
//!     mergeCells → conditionalFormatting+ → dataValidations →
//!     hyperlinks → legacyDrawing` byte-for-byte equivalent to
//!     `xlsx.Writer.save`'s prior local emit branch.
//!   - the per-sheet rels emitter (`emitSheetRels`).
//!   - the comments + VML drawing emitters
//!     (`emitCommentsXml`, `emitVmlDrawingXml`).
//!   - the validators (`validateMergeRange`, `validateAutoFilterRange`,
//!     `validateHyperlinkRange`, `parseA1Corner`, `formatCellRef`,
//!     `assertNoForbiddenXmlBytes`, `appendXmlEscaped`).
//!
//! Stdlib only. Zig 0.15.2.
//!
//! ─── iter-wr-4 perf invariants ───
//!
//! Every emit primitive uses inline buffer appends (`appendSlice` +
//! `print` over a `std.ArrayListUnmanaged(u8)`). NO `std.fmt.format`
//! abstractions, NO intermediate copies. The byte order is locked
//! by the writer-tests at `src/writer.zig:3853`+ (CT_Worksheet child
//! order), `4427`+ (DV), `4535`+ (hyperlinks rId numbering), and
//! `5233` (row atomicity). The 1k×10 bench target is 6.7ms ± 0.3
//! with a strict 1.10× ceiling = 7.4ms.

const std = @import("std");

const Allocator = std.mem.Allocator;

// ─── Excel hard limits ────────────────────────────────────────────────

pub const EXCEL_MAX_COL: u32 = 16_384;
pub const EXCEL_MAX_ROW: u32 = 1_048_576;

pub const Error = error{
    OutOfMemory,
    RowOutOfRange,
    ColumnOutOfRange,
    InvalidMergeRange,
    InvalidAutoFilterRange,
    InvalidHyperlinkRange,
    InvalidXmlByte,
};

// ─── State types — identical shape to `xlsx.Writer.SheetWriter` ──────

/// External-URL hyperlink registered against a cell or range. Both
/// fields are caller-owned slices.
pub const Hyperlink = struct {
    range: []const u8,
    url: []const u8,
};

/// Internal-target hyperlink — jumps to another cell or range
/// within the same workbook. Emitted as
/// `<hyperlink ref="…" location="…"/>` (no r:id, no rels entry).
pub const InternalHyperlink = struct {
    range: []const u8,
    location: []const u8,
};

/// Cell comment (note). Plain-text body — rich-text comment bodies
/// are not yet exposed by the writer's public surface.
pub const Comment = struct {
    ref: []const u8,
    author: []const u8,
    text: []const u8,
};

/// Per-column width override. `col_min..=col_max` is the inclusive
/// 1-based range this width applies to.
pub const ColumnWidth = struct {
    col_min: u32,
    col_max: u32,
    width: f32,
};

/// Comparison operator for `cellIs` conditional-format rules.
pub const CfOperator = enum {
    less_than,
    less_than_or_equal,
    equal,
    not_equal,
    greater_than,
    greater_than_or_equal,
    between,
    not_between,

    pub fn toOoxml(self: CfOperator) []const u8 {
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

    pub fn needsSecondFormula(self: CfOperator) bool {
        return self == .between or self == .not_between;
    }
};

/// Conditional-format rule entry — the `range` is A1-style, the
/// `rule` is the OOXML cfRule payload.
pub const ConditionalFormat = struct {
    range: []const u8,
    rule: ConditionalFormatRule,
};

pub const ConditionalFormatRule = union(enum) {
    cell_is: struct {
        operator: CfOperator,
        formula1: []const u8,
        formula2: ?[]const u8,
        dxf_id: u32,
    },
    expression: struct {
        formula: []const u8,
        dxf_id: u32,
    },
    color_scale: struct {
        low_color_argb: u32,
        mid_color_argb: ?u32,
        high_color_argb: u32,
    },
    data_bar: struct {
        color_argb: u32,
    },
};

/// List-type data validation (dropdown).
pub const DataValidationList = struct {
    range: []const u8,
    values: []const []const u8,
};

/// Numeric / date / time / text-length / custom data validation.
pub const DataValidationRange = struct {
    range: []const u8,
    /// One of "whole", "decimal", "date", "time", "textLength", "custom".
    kind_name: []const u8,
    /// One of "between", "notBetween", "equal", "notEqual",
    /// "greaterThan", "lessThan", "greaterThanOrEqual",
    /// "lessThanOrEqual". `null` for `type="custom"`.
    op_name: ?[]const u8,
    formula1: []const u8,
    /// Required iff `op_name` is "between" or "notBetween".
    formula2: ?[]const u8,
};

/// All inputs for a per-sheet `xl/worksheets/sheetN.xml` emit.
/// Mirrors `SheetWriter`'s public field set; `body` is the
/// pre-built `<row>...</row>` bytes (the SST indices are baked
/// in by the row-emit primitives).
pub const SheetEmitInputs = struct {
    body: []const u8,
    freeze_rows: u32 = 0,
    freeze_cols: u32 = 0,
    column_widths: []const ColumnWidth = &.{},
    auto_filter_range: ?[]const u8 = null,
    merged_cells: []const []const u8 = &.{},
    conditional_formats: []const ConditionalFormat = &.{},
    data_validations: []const DataValidationList = &.{},
    data_validation_ranges: []const DataValidationRange = &.{},
    hyperlinks: []const Hyperlink = &.{},
    internal_hyperlinks: []const InternalHyperlink = &.{},
    /// Number of comments registered on this sheet — used to decide
    /// whether to emit a `<legacyDrawing>` link. The actual VML
    /// drawing payload is built separately via `emitVmlDrawingXml`.
    comment_count: usize = 0,
};

// ─── OOXML skeleton strings ──────────────────────────────────────────

const WORKSHEET_PROLOG: []const u8 =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
;

const RELS_HEAD: []const u8 =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
;
const RELS_TAIL: []const u8 = "</Relationships>";

// ─── Per-sheet `xl/worksheets/sheetN.xml` emitter ─────────────────────

/// Render a worksheet XML payload into `out`. CT_Worksheet child
/// order is fixed: `sheetViews → cols → sheetData → autoFilter →
/// mergeCells → conditionalFormatting+ → dataValidations →
/// hyperlinks → legacyDrawing`. Any drift = Excel "repaired"
/// prompt across the corpus, so the order is locked.
///
/// `out` is appended to (not reset). Caller owns lifetime.
pub fn emitWorksheetXml(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    inputs: SheetEmitInputs,
) Error!void {
    try out.appendSlice(allocator, WORKSHEET_PROLOG);

    // <sheetViews> — emitted when any pane is frozen.
    if (inputs.freeze_rows != 0 or inputs.freeze_cols != 0) {
        try out.appendSlice(allocator, "<sheetViews><sheetView workbookViewId=\"0\">");
        try out.appendSlice(allocator, "<pane");
        if (inputs.freeze_cols != 0) try out.print(allocator, " xSplit=\"{d}\"", .{inputs.freeze_cols});
        if (inputs.freeze_rows != 0) try out.print(allocator, " ySplit=\"{d}\"", .{inputs.freeze_rows});
        var tl_buf: [16]u8 = undefined;
        const top_left = try formatCellRef(&tl_buf, inputs.freeze_rows + 1, inputs.freeze_cols);
        const active_pane: []const u8 = if (inputs.freeze_rows != 0 and inputs.freeze_cols != 0)
            "bottomRight"
        else if (inputs.freeze_rows != 0)
            "bottomLeft"
        else
            "topRight";
        try out.print(allocator, " topLeftCell=\"{s}\" activePane=\"{s}\" state=\"frozen\"/>", .{ top_left, active_pane });
        try out.appendSlice(allocator, "</sheetView></sheetViews>");
    }

    // <cols> — one <col> per registered width override.
    if (inputs.column_widths.len > 0) {
        try out.appendSlice(allocator, "<cols>");
        for (inputs.column_widths) |cw| {
            try out.print(
                allocator,
                "<col min=\"{d}\" max=\"{d}\" width=\"{d}\" customWidth=\"1\"/>",
                .{ cw.col_min, cw.col_max, cw.width },
            );
        }
        try out.appendSlice(allocator, "</cols>");
    }

    try out.appendSlice(allocator, "<sheetData>");
    try out.appendSlice(allocator, inputs.body);
    try out.appendSlice(allocator, "</sheetData>");

    // <autoFilter> must come after </sheetData>.
    if (inputs.auto_filter_range) |range| {
        try out.appendSlice(allocator, "<autoFilter ref=\"");
        try appendXmlEscaped(allocator, out, range);
        try out.appendSlice(allocator, "\"/>");
    }

    // <mergeCells> follows <autoFilter> per ECMA-376 CT_Worksheet
    // child order. Ranges were validated on intake; defensively
    // xml-escape them on emit anyway.
    if (inputs.merged_cells.len > 0) {
        try out.print(allocator, "<mergeCells count=\"{d}\">", .{inputs.merged_cells.len});
        for (inputs.merged_cells) |range| {
            try out.appendSlice(allocator, "<mergeCell ref=\"");
            try appendXmlEscaped(allocator, out, range);
            try out.appendSlice(allocator, "\"/>");
        }
        try out.appendSlice(allocator, "</mergeCells>");
    }

    // <conditionalFormatting> — one block per rule. Priority increments
    // per rule so overlapping ranges produce a deterministic cascade.
    for (inputs.conditional_formats, 1..) |cf, cf_priority| {
        try out.appendSlice(allocator, "<conditionalFormatting sqref=\"");
        try appendXmlEscaped(allocator, out, cf.range);
        try out.appendSlice(allocator, "\">");
        switch (cf.rule) {
            .cell_is => |r| {
                try out.print(
                    allocator,
                    "<cfRule type=\"cellIs\" dxfId=\"{d}\" priority=\"{d}\" operator=\"{s}\">",
                    .{ r.dxf_id, cf_priority, r.operator.toOoxml() },
                );
                try out.appendSlice(allocator, "<formula>");
                try appendXmlEscaped(allocator, out, r.formula1);
                try out.appendSlice(allocator, "</formula>");
                if (r.formula2) |f2| {
                    try out.appendSlice(allocator, "<formula>");
                    try appendXmlEscaped(allocator, out, f2);
                    try out.appendSlice(allocator, "</formula>");
                }
                try out.appendSlice(allocator, "</cfRule>");
            },
            .expression => |r| {
                try out.print(
                    allocator,
                    "<cfRule type=\"expression\" dxfId=\"{d}\" priority=\"{d}\">",
                    .{ r.dxf_id, cf_priority },
                );
                try out.appendSlice(allocator, "<formula>");
                try appendXmlEscaped(allocator, out, r.formula);
                try out.appendSlice(allocator, "</formula>");
                try out.appendSlice(allocator, "</cfRule>");
            },
            .color_scale => |r| {
                try out.print(
                    allocator,
                    "<cfRule type=\"colorScale\" priority=\"{d}\"><colorScale>",
                    .{cf_priority},
                );
                if (r.mid_color_argb != null) {
                    // 3-stop: min / 50th percentile / max.
                    try out.appendSlice(allocator, "<cfvo type=\"min\"/><cfvo type=\"percentile\" val=\"50\"/><cfvo type=\"max\"/>");
                    try out.print(allocator, "<color rgb=\"{X:0>8}\"/>", .{r.low_color_argb});
                    try out.print(allocator, "<color rgb=\"{X:0>8}\"/>", .{r.mid_color_argb.?});
                    try out.print(allocator, "<color rgb=\"{X:0>8}\"/>", .{r.high_color_argb});
                } else {
                    // 2-stop: min / max only.
                    try out.appendSlice(allocator, "<cfvo type=\"min\"/><cfvo type=\"max\"/>");
                    try out.print(allocator, "<color rgb=\"{X:0>8}\"/>", .{r.low_color_argb});
                    try out.print(allocator, "<color rgb=\"{X:0>8}\"/>", .{r.high_color_argb});
                }
                try out.appendSlice(allocator, "</colorScale></cfRule>");
            },
            .data_bar => |r| {
                try out.print(
                    allocator,
                    "<cfRule type=\"dataBar\" priority=\"{d}\"><dataBar><cfvo type=\"min\"/><cfvo type=\"max\"/>",
                    .{cf_priority},
                );
                try out.print(allocator, "<color rgb=\"{X:0>8}\"/>", .{r.color_argb});
                try out.appendSlice(allocator, "</dataBar></cfRule>");
            },
        }
        try out.appendSlice(allocator, "</conditionalFormatting>");
    }

    // <dataValidations> — list entries first (iter13 ordering), then
    // numeric / custom range entries.
    const dv_list_count = inputs.data_validations.len;
    const dv_range_count = inputs.data_validation_ranges.len;
    if (dv_list_count + dv_range_count > 0) {
        try out.print(allocator, "<dataValidations count=\"{d}\">", .{dv_list_count + dv_range_count});
        for (inputs.data_validations) |dv| {
            try out.appendSlice(allocator, "<dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"");
            try appendXmlEscaped(allocator, out, dv.range);
            try out.appendSlice(allocator, "\"><formula1>&quot;");
            for (dv.values, 0..) |v, vi| {
                if (vi != 0) try out.append(allocator, ',');
                try appendXmlEscaped(allocator, out, v);
            }
            try out.appendSlice(allocator, "&quot;</formula1></dataValidation>");
        }
        for (inputs.data_validation_ranges) |dv| {
            try out.appendSlice(allocator, "<dataValidation type=\"");
            try out.appendSlice(allocator, dv.kind_name);
            try out.appendSlice(allocator, "\"");
            if (dv.op_name) |op| {
                try out.print(allocator, " operator=\"{s}\"", .{op});
            }
            try out.appendSlice(allocator, " allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"");
            try appendXmlEscaped(allocator, out, dv.range);
            try out.appendSlice(allocator, "\"><formula1>");
            try appendXmlEscaped(allocator, out, dv.formula1);
            try out.appendSlice(allocator, "</formula1>");
            if (dv.formula2) |f2| {
                try out.appendSlice(allocator, "<formula2>");
                try appendXmlEscaped(allocator, out, f2);
                try out.appendSlice(allocator, "</formula2>");
            }
            try out.appendSlice(allocator, "</dataValidation>");
        }
        try out.appendSlice(allocator, "</dataValidations>");
    }

    // <hyperlinks> — external entries (rIds) FIRST, then internal
    // (`location="…"`); r:id numbering matches the per-sheet rels.
    if (inputs.hyperlinks.len > 0 or inputs.internal_hyperlinks.len > 0) {
        try out.appendSlice(allocator, "<hyperlinks>");
        for (inputs.hyperlinks, 0..) |h, idx| {
            try out.appendSlice(allocator, "<hyperlink ref=\"");
            try appendXmlEscaped(allocator, out, h.range);
            try out.print(allocator, "\" r:id=\"rId{d}\"/>", .{idx + 1});
        }
        for (inputs.internal_hyperlinks) |h| {
            try out.appendSlice(allocator, "<hyperlink ref=\"");
            try appendXmlEscaped(allocator, out, h.range);
            try out.appendSlice(allocator, "\" location=\"");
            try appendXmlEscaped(allocator, out, h.location);
            try out.appendSlice(allocator, "\"/>");
        }
        try out.appendSlice(allocator, "</hyperlinks>");
    }

    // <legacyDrawing> when comments present. rId scheme: 1..N
    // external hyperlinks, N+1 = comments part, N+2 = vmlDrawing.
    if (inputs.comment_count > 0) {
        const vml_rid = inputs.hyperlinks.len + 2;
        try out.print(allocator, "<legacyDrawing r:id=\"rId{d}\"/>", .{vml_rid});
    }

    try out.appendSlice(allocator, "</worksheet>");
}

// ─── Per-sheet `xl/worksheets/_rels/sheetN.xml.rels` emitter ──────────

/// Render a per-sheet rels XML payload. rId scheme: 1..N external
/// hyperlinks, then comments, then vmlDrawing (when sheet has
/// comments). `sheet_idx` is 0-based; the comments / vml part
/// names use `sheet_idx + 1` (matching the workbook-level rId
/// numbering Writer.save uses).
///
/// Returns false if no rels payload was generated (caller skips
/// writing the part); true otherwise.
pub fn emitSheetRels(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    sheet_idx: usize,
    hyperlinks: []const Hyperlink,
    comment_count: usize,
) Error!bool {
    if (hyperlinks.len == 0 and comment_count == 0) return false;
    try out.appendSlice(allocator, RELS_HEAD);
    for (hyperlinks, 0..) |h, idx| {
        try out.print(allocator, "<Relationship Id=\"rId{d}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"", .{idx + 1});
        try appendXmlEscaped(allocator, out, h.url);
        try out.appendSlice(allocator, "\" TargetMode=\"External\"/>");
    }
    if (comment_count > 0) {
        const comments_rid = hyperlinks.len + 1;
        const vml_rid = hyperlinks.len + 2;
        try out.print(
            allocator,
            "<Relationship Id=\"rId{d}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" Target=\"../comments{d}.xml\"/>",
            .{ comments_rid, sheet_idx + 1 },
        );
        try out.print(
            allocator,
            "<Relationship Id=\"rId{d}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\" Target=\"../drawings/vmlDrawing{d}.vml\"/>",
            .{ vml_rid, sheet_idx + 1 },
        );
    }
    try out.appendSlice(allocator, RELS_TAIL);
    return true;
}

// ─── `xl/commentsN.xml` emitter ───────────────────────────────────────

/// Render `xl/commentsN.xml`. Authors are deduped O(N²) on emit;
/// first-occurrence wins on `authorId` numbering. Plain-text comment
/// bodies emit `<text><t xml:space="preserve">…</t></text>` — NO
/// synthetic `<r>` wrapper (a `<r>` would make the reader treat
/// every Writer-produced comment as rich, breaking the plain/rich
/// contract).
pub fn emitCommentsXml(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    comments: []const Comment,
) Error!void {
    var authors: std.ArrayListUnmanaged([]const u8) = .empty;
    defer authors.deinit(allocator);
    var author_ids: std.ArrayListUnmanaged(usize) = .empty;
    defer author_ids.deinit(allocator);
    for (comments) |c| {
        var found: ?usize = null;
        for (authors.items, 0..) |a, j| {
            if (std.mem.eql(u8, a, c.author)) {
                found = j;
                break;
            }
        }
        if (found) |j| {
            try author_ids.append(allocator, j);
        } else {
            try author_ids.append(allocator, authors.items.len);
            try authors.append(allocator, c.author);
        }
    }

    try out.appendSlice(allocator, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    try out.appendSlice(allocator, "<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
    try out.appendSlice(allocator, "<authors>");
    for (authors.items) |a| {
        try out.appendSlice(allocator, "<author>");
        try appendXmlEscaped(allocator, out, a);
        try out.appendSlice(allocator, "</author>");
    }
    try out.appendSlice(allocator, "</authors><commentList>");
    for (comments, author_ids.items) |c, aid| {
        try out.print(allocator, "<comment ref=\"{s}\" authorId=\"{d}\"><text><t xml:space=\"preserve\">", .{ c.ref, aid });
        try appendXmlEscaped(allocator, out, c.text);
        try out.appendSlice(allocator, "</t></text></comment>");
    }
    try out.appendSlice(allocator, "</commentList></comments>");
}

// ─── `xl/drawings/vmlDrawingN.vml` emitter ────────────────────────────

/// Render the legacy VML notes drawing for sheet `sheet_idx`. Shape
/// IDs start at 1025 and increment per comment. Every shape's
/// from/to anchor is clamped to `EXCEL_MAX_COL - 1` /
/// `EXCEL_MAX_ROW - 1` so notes on cells near XFD / row 1048576
/// don't reference off-sheet cells (un-clamped emit produces
/// inverted anchors per `a966e29`). `<o:idmap>` chunks cover 1024
/// shape IDs each — over-provision by one is harmless, under-
/// provision = unrendered notes.
pub fn emitVmlDrawingXml(
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    comments: []const Comment,
) Error!void {
    try out.appendSlice(
        allocator,
        \\<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
        ,
    );
    const num_idmaps: usize = comments.len / 1024 + 1;
    try out.appendSlice(allocator, "<o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"");
    for (0..num_idmaps) |k| {
        if (k != 0) try out.append(allocator, ',');
        try out.print(allocator, "{d}", .{k + 1});
    }
    try out.appendSlice(allocator, "\"/></o:shapelayout>");
    try out.appendSlice(
        allocator,
        \\<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>
        ,
    );
    for (comments, 0..) |c, shape_idx| {
        const rc = try parseA1Corner(c.ref);
        const row0 = rc.row - 1;
        const col0 = rc.col - 1;
        const from_col = @min(col0 + 1, EXCEL_MAX_COL - 1);
        const from_row = @min(row0, EXCEL_MAX_ROW - 1);
        const to_col = @min(col0 + 3, EXCEL_MAX_COL - 1);
        const to_row = @min(row0 + 4, EXCEL_MAX_ROW - 1);
        try out.print(
            allocator,
            "<v:shape id=\"_x0000_s{d}\" type=\"#_x0000_t202\" style=\"position:absolute;margin-left:60pt;margin-top:10pt;width:100pt;height:60pt;z-index:{d};visibility:hidden\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\"><v:fill color2=\"#ffffe1\"/><v:shadow on=\"t\" color=\"black\" obscured=\"t\"/><v:path o:connecttype=\"none\"/><v:textbox><div style=\"text-align:left\"/></v:textbox><x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>{d}, 15, {d}, 2, {d}, 31, {d}, 3</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>{d}</x:Row><x:Column>{d}</x:Column></x:ClientData></v:shape>",
            .{ 1025 + shape_idx, shape_idx + 1, from_col, from_row, to_col, to_row, row0, col0 },
        );
    }
    try out.appendSlice(allocator, "</xml>");
}

// ─── Helpers ─────────────────────────────────────────────────────────

pub const A1Corner = struct { col: u32, row: u32 };

pub fn parseA1Corner(s: []const u8) Error!A1Corner {
    if (s.len == 0) return error.InvalidMergeRange;
    var i: usize = 0;
    var col: u32 = 0;
    while (i < s.len and s[i] >= 'A' and s[i] <= 'Z') : (i += 1) {
        col = col * 26 + (s[i] - 'A' + 1);
        if (col > EXCEL_MAX_COL) return error.InvalidMergeRange;
    }
    // Need at least one letter and at least one digit after it.
    if (i == 0 or i == s.len) return error.InvalidMergeRange;
    if (s[i] == '0') return error.InvalidMergeRange;
    var row: u32 = 0;
    while (i < s.len and s[i] >= '0' and s[i] <= '9') : (i += 1) {
        row = row * 10 + (s[i] - '0');
        if (row > EXCEL_MAX_ROW) return error.InvalidMergeRange;
    }
    if (i != s.len) return error.InvalidMergeRange;
    return .{ .col = col, .row = row };
}

pub fn formatCellRef(buf: *[16]u8, row: u32, col_idx: u32) Error![]u8 {
    if (row == 0 or row > EXCEL_MAX_ROW) return error.RowOutOfRange;
    if (col_idx >= EXCEL_MAX_COL) return error.ColumnOutOfRange;
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
    return std.fmt.bufPrint(buf, "{s}{d}", .{ letters, row }) catch unreachable;
}

/// Validate an A1-style merge range: must be `TL:BR` form, both
/// corners valid A1, top-left must precede or equal bottom-right
/// on both axes, and the range must NOT collapse to a single cell
/// (Excel warns on 1×1 "merges").
pub fn validateMergeRange(range: []const u8) Error!void {
    const colon = std.mem.indexOfScalar(u8, range, ':') orelse return error.InvalidMergeRange;
    const tl = try parseA1Corner(range[0..colon]);
    const br = try parseA1Corner(range[colon + 1 ..]);
    if (tl.col > br.col or tl.row > br.row) return error.InvalidMergeRange;
    if (tl.col == br.col and tl.row == br.row) return error.InvalidMergeRange;
}

pub fn validateAutoFilterRange(range: []const u8) Error!void {
    if (range.len == 0) return error.InvalidAutoFilterRange;
    if (std.mem.indexOfScalar(u8, range, ':')) |colon| {
        const tl = parseA1Corner(range[0..colon]) catch return error.InvalidAutoFilterRange;
        const br = parseA1Corner(range[colon + 1 ..]) catch return error.InvalidAutoFilterRange;
        if (tl.col > br.col or tl.row > br.row) return error.InvalidAutoFilterRange;
    } else {
        _ = parseA1Corner(range) catch return error.InvalidAutoFilterRange;
    }
}

pub fn validateHyperlinkRange(range: []const u8) Error!void {
    if (range.len == 0) return error.InvalidHyperlinkRange;
    if (std.mem.indexOfScalar(u8, range, ':')) |colon| {
        const tl = parseA1Corner(range[0..colon]) catch return error.InvalidHyperlinkRange;
        const br = parseA1Corner(range[colon + 1 ..]) catch return error.InvalidHyperlinkRange;
        if (tl.col > br.col or tl.row > br.row) return error.InvalidHyperlinkRange;
    } else {
        _ = parseA1Corner(range) catch return error.InvalidHyperlinkRange;
    }
}

/// XML 1.0 forbids most C0 control bytes in document content.
/// Allowed: 0x09 (tab), 0x0A (LF), 0x0D (CR). Forbidden:
/// 0x00–0x08, 0x0B, 0x0C, 0x0E–0x1F.
///
/// 0x7F (DEL) is **valid** under XML 1.0 production [2] —
/// `Char ::= #x9 | #xA | #xD | [#x20-#xD7FF] | …` — DEL falls in
/// `[#x20-#xD7FF]` and many real-world workbooks carry it. B3
/// iter-wr-6 reconciled the lift: the Writer-side test pinned
/// 0x7F as legal (matching spec); the wr-4 plan-side variant
/// erroneously rejected it. Restored to spec here.
pub inline fn isForbiddenXmlByte(c: u8) bool {
    return switch (c) {
        0x00...0x08, 0x0B, 0x0C, 0x0E...0x1F => true,
        else => false,
    };
}

pub fn assertNoForbiddenXmlBytes(s: []const u8) Error!void {
    for (s) |c| if (isForbiddenXmlByte(c)) return error.InvalidXmlByte;
}

/// Append `s` to `out`, XML-escaping the five canonical entities
/// (`<`, `>`, `&`, `"`, `'`) and rejecting XML 1.0 forbidden control
/// bytes. Other bytes (including UTF-8 continuation bytes for non-
/// ASCII characters) pass through verbatim.
///
/// Use this for ATTRIBUTE values where `"` and `'` need escaping.
/// For ELEMENT text content, prefer `appendXmlEscapedText` so the
/// quote characters round-trip verbatim and the byte image matches
/// the prior `pkg/workbook.zig::appendXmlEscapedText` output.
pub fn appendXmlEscaped(allocator: Allocator, out: *std.ArrayListUnmanaged(u8), s: []const u8) Error!void {
    for (s) |ch| {
        if (isForbiddenXmlByte(ch)) return error.InvalidXmlByte;
        switch (ch) {
            '<' => try out.appendSlice(allocator, "&lt;"),
            '>' => try out.appendSlice(allocator, "&gt;"),
            '&' => try out.appendSlice(allocator, "&amp;"),
            '"' => try out.appendSlice(allocator, "&quot;"),
            '\'' => try out.appendSlice(allocator, "&apos;"),
            else => try out.append(allocator, ch),
        }
    }
}

/// Append `s` to `out`, XML-escaping the three element-content
/// entities (`<`, `>`, `&`) and rejecting XML 1.0 forbidden control
/// bytes. The two quote characters (`"`, `'`) pass through verbatim
/// — that's the byte-stable contract for element text bodies (see
/// `pkg/workbook.zig` pre-iter-wr-6 emission). Use `appendXmlEscaped`
/// for attribute values where the quote characters MUST be escaped.
pub fn appendXmlEscapedText(allocator: Allocator, out: *std.ArrayListUnmanaged(u8), s: []const u8) Error!void {
    for (s) |ch| {
        if (isForbiddenXmlByte(ch)) return error.InvalidXmlByte;
        switch (ch) {
            '<' => try out.appendSlice(allocator, "&lt;"),
            '>' => try out.appendSlice(allocator, "&gt;"),
            '&' => try out.appendSlice(allocator, "&amp;"),
            else => try out.append(allocator, ch),
        }
    }
}

// ─── Tests ────────────────────────────────────────────────────────────

test "formatCellRef A1, B2, Z1, AA1, AAA1" {
    var buf: [16]u8 = undefined;
    try std.testing.expectEqualStrings("A1", try formatCellRef(&buf, 1, 0));
    try std.testing.expectEqualStrings("B2", try formatCellRef(&buf, 2, 1));
    try std.testing.expectEqualStrings("Z1", try formatCellRef(&buf, 1, 25));
    try std.testing.expectEqualStrings("AA1", try formatCellRef(&buf, 1, 26));
    try std.testing.expectEqualStrings("AAA1", try formatCellRef(&buf, 1, 702));
    try std.testing.expectEqualStrings("XFD1048576", try formatCellRef(&buf, 1_048_576, 16_383));
}

test "formatCellRef rejects out-of-range" {
    var buf: [16]u8 = undefined;
    try std.testing.expectError(error.RowOutOfRange, formatCellRef(&buf, 0, 0));
    try std.testing.expectError(error.RowOutOfRange, formatCellRef(&buf, 1_048_577, 0));
    try std.testing.expectError(error.ColumnOutOfRange, formatCellRef(&buf, 1, 16_384));
}

test "appendXmlEscaped covers all 5 entities" {
    const a = std.testing.allocator;
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(a);
    try appendXmlEscaped(a, &buf, "<>&\"'");
    try std.testing.expectEqualStrings("&lt;&gt;&amp;&quot;&apos;", buf.items);
}

test "appendXmlEscaped rejects forbidden control bytes" {
    const a = std.testing.allocator;
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(a);
    try std.testing.expectError(error.InvalidXmlByte, appendXmlEscaped(a, &buf, "ok\x00bad"));
}

test "appendXmlEscaped permits tab/LF/CR" {
    const a = std.testing.allocator;
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(a);
    try appendXmlEscaped(a, &buf, "tab\there\nline\rrun");
    try std.testing.expectEqualStrings("tab\there\nline\rrun", buf.items);
}

test "appendXmlEscapedText escapes only <, >, & — quotes verbatim" {
    const a = std.testing.allocator;
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(a);
    try appendXmlEscapedText(a, &buf, "<>&\"'");
    try std.testing.expectEqualStrings("&lt;&gt;&amp;\"'", buf.items);
}

test "appendXmlEscapedText rejects forbidden control bytes" {
    const a = std.testing.allocator;
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(a);
    try std.testing.expectError(error.InvalidXmlByte, appendXmlEscapedText(a, &buf, "ok\x00bad"));
}

test "appendXmlEscapedText permits tab/LF/CR" {
    const a = std.testing.allocator;
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(a);
    try appendXmlEscapedText(a, &buf, "tab\there\nline\rrun");
    try std.testing.expectEqualStrings("tab\there\nline\rrun", buf.items);
}

test "parseA1Corner basic + bounds" {
    try std.testing.expectEqual(@as(u32, 1), (try parseA1Corner("A1")).col);
    try std.testing.expectEqual(@as(u32, 1), (try parseA1Corner("A1")).row);
    try std.testing.expectEqual(@as(u32, 26), (try parseA1Corner("Z99")).col);
    try std.testing.expectEqual(@as(u32, 99), (try parseA1Corner("Z99")).row);
    try std.testing.expectError(error.InvalidMergeRange, parseA1Corner(""));
    try std.testing.expectError(error.InvalidMergeRange, parseA1Corner("1A"));
    try std.testing.expectError(error.InvalidMergeRange, parseA1Corner("A0"));
}

test "validateMergeRange rejects single-cell + inverted" {
    try std.testing.expectError(error.InvalidMergeRange, validateMergeRange("A1:A1"));
    try std.testing.expectError(error.InvalidMergeRange, validateMergeRange("B2:A1"));
    try validateMergeRange("A1:B2");
}

test "validateHyperlinkRange accepts single-cell + range" {
    try validateHyperlinkRange("A1");
    try validateHyperlinkRange("A1:B2");
    try std.testing.expectError(error.InvalidHyperlinkRange, validateHyperlinkRange(""));
    try std.testing.expectError(error.InvalidHyperlinkRange, validateHyperlinkRange("B2:A1"));
}

test "emitWorksheetXml: empty body, no optional blocks" {
    const a = std.testing.allocator;
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(a);
    try emitWorksheetXml(a, &buf, .{ .body = "" });
    try std.testing.expect(std.mem.startsWith(u8, buf.items, "<?xml"));
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "<sheetData></sheetData>") != null);
    try std.testing.expect(std.mem.endsWith(u8, buf.items, "</worksheet>"));
    // No optional blocks at all.
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "<sheetViews") == null);
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "<cols") == null);
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "<autoFilter") == null);
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "<mergeCells") == null);
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "<conditionalFormatting") == null);
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "<dataValidations") == null);
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "<hyperlinks") == null);
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "<legacyDrawing") == null);
}

test "emitWorksheetXml: freeze panes — bottomRight (rows + cols)" {
    const a = std.testing.allocator;
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(a);
    try emitWorksheetXml(a, &buf, .{
        .body = "",
        .freeze_rows = 1,
        .freeze_cols = 2,
    });
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "xSplit=\"2\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "ySplit=\"1\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "topLeftCell=\"C2\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "activePane=\"bottomRight\"") != null);
}

test "emitWorksheetXml: freeze panes — bottomLeft (rows only)" {
    const a = std.testing.allocator;
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(a);
    try emitWorksheetXml(a, &buf, .{
        .body = "",
        .freeze_rows = 1,
    });
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "activePane=\"bottomLeft\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "xSplit=") == null);
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "ySplit=\"1\"") != null);
}

test "emitWorksheetXml: cols block + width" {
    const a = std.testing.allocator;
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(a);
    const cws = [_]ColumnWidth{
        .{ .col_min = 1, .col_max = 1, .width = 12.5 },
        .{ .col_min = 2, .col_max = 3, .width = 5.0 },
    };
    try emitWorksheetXml(a, &buf, .{ .body = "", .column_widths = &cws });
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "<col min=\"1\" max=\"1\" width=\"12.5\" customWidth=\"1\"/>") != null);
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "<col min=\"2\" max=\"3\" width=\"5\" customWidth=\"1\"/>") != null);
}

test "emitWorksheetXml: mergeCells with count attr even for N=1" {
    const a = std.testing.allocator;
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(a);
    const merges = [_][]const u8{"A1:B2"};
    try emitWorksheetXml(a, &buf, .{ .body = "", .merged_cells = &merges });
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "<mergeCells count=\"1\"><mergeCell ref=\"A1:B2\"/></mergeCells>") != null);
}

test "emitWorksheetXml: hyperlinks rId numbering matches rels" {
    const a = std.testing.allocator;
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(a);
    const hls = [_]Hyperlink{
        .{ .range = "A1", .url = "https://ex.com" },
        .{ .range = "A2", .url = "https://ex.com/?a=1&b=2" },
    };
    try emitWorksheetXml(a, &buf, .{ .body = "", .hyperlinks = &hls });
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "<hyperlink ref=\"A1\" r:id=\"rId1\"/>") != null);
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "<hyperlink ref=\"A2\" r:id=\"rId2\"/>") != null);
}

test "emitWorksheetXml: dataValidations list FIRST then range entries" {
    const a = std.testing.allocator;
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(a);
    const list_vals = [_][]const u8{ "yes", "no" };
    const lists = [_]DataValidationList{
        .{ .range = "A1:A10", .values = &list_vals },
    };
    const ranges = [_]DataValidationRange{
        .{
            .range = "B1:B10",
            .kind_name = "whole",
            .op_name = "between",
            .formula1 = "1",
            .formula2 = "10",
        },
    };
    try emitWorksheetXml(a, &buf, .{
        .body = "",
        .data_validations = &lists,
        .data_validation_ranges = &ranges,
    });
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "<dataValidations count=\"2\">") != null);
    const list_pos = std.mem.indexOf(u8, buf.items, "type=\"list\"") orelse return error.TestFailed;
    const whole_pos = std.mem.indexOf(u8, buf.items, "type=\"whole\"") orelse return error.TestFailed;
    try std.testing.expect(list_pos < whole_pos);
}

test "emitWorksheetXml: legacyDrawing rId scheme = hyperlinks + 2" {
    const a = std.testing.allocator;
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(a);
    const hls = [_]Hyperlink{.{ .range = "A1", .url = "https://ex.com" }};
    try emitWorksheetXml(a, &buf, .{
        .body = "",
        .hyperlinks = &hls,
        .comment_count = 3,
    });
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "<legacyDrawing r:id=\"rId3\"/>") != null);
}

test "emitSheetRels: returns false when both lists empty" {
    const a = std.testing.allocator;
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(a);
    const wrote = try emitSheetRels(a, &buf, 0, &.{}, 0);
    try std.testing.expect(!wrote);
    try std.testing.expectEqual(@as(usize, 0), buf.items.len);
}

test "emitSheetRels: hyperlinks only" {
    const a = std.testing.allocator;
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(a);
    const hls = [_]Hyperlink{
        .{ .range = "A1", .url = "https://ex.com/?q=1&x=2" },
    };
    const wrote = try emitSheetRels(a, &buf, 0, &hls, 0);
    try std.testing.expect(wrote);
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "Id=\"rId1\"") != null);
    // URL is XML-attribute-escaped — & becomes &amp;
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "https://ex.com/?q=1&amp;x=2") != null);
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "TargetMode=\"External\"") != null);
}

test "emitSheetRels: comments + hyperlinks rId scheme" {
    const a = std.testing.allocator;
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(a);
    const hls = [_]Hyperlink{
        .{ .range = "A1", .url = "https://a.com" },
        .{ .range = "A2", .url = "https://b.com" },
    };
    const wrote = try emitSheetRels(a, &buf, 1, &hls, 5);
    try std.testing.expect(wrote);
    // Comments rId = 3 (after 2 hyperlinks); vmlDrawing rId = 4.
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "Id=\"rId3\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "Target=\"../comments2.xml\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "Id=\"rId4\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "Target=\"../drawings/vmlDrawing2.vml\"") != null);
}

test "emitCommentsXml: dedupes authors O(N²) first-occurrence-wins" {
    const a = std.testing.allocator;
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(a);
    const comments = [_]Comment{
        .{ .ref = "A1", .author = "alice", .text = "hi" },
        .{ .ref = "A2", .author = "bob", .text = "hello" },
        .{ .ref = "A3", .author = "alice", .text = "again" },
    };
    try emitCommentsXml(a, &buf, &comments);
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "<author>alice</author>") != null);
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "<author>bob</author>") != null);
    // 2 authors total — count first occurrences.
    var n_auth: usize = 0;
    var pos: usize = 0;
    while (std.mem.indexOfPos(u8, buf.items, pos, "<author>")) |p| {
        n_auth += 1;
        pos = p + 1;
    }
    try std.testing.expectEqual(@as(usize, 2), n_auth);
    // Comment 3 references author 0 (alice).
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "<comment ref=\"A3\" authorId=\"0\">") != null);
}

test "emitCommentsXml: plain-text body — no <r> wrapper" {
    const a = std.testing.allocator;
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(a);
    const comments = [_]Comment{.{ .ref = "A1", .author = "x", .text = "plain" }};
    try emitCommentsXml(a, &buf, &comments);
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "<text><t xml:space=\"preserve\">plain</t></text>") != null);
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "<r>") == null);
}

test "emitVmlDrawingXml: idmap chunks scale at 1024-comment boundary" {
    const a = std.testing.allocator;
    // 1023 comments → 1 idmap; 1024 → 2 idmaps; 2048 → 3.
    {
        var buf: std.ArrayListUnmanaged(u8) = .empty;
        defer buf.deinit(a);
        var c: [1023]Comment = undefined;
        for (&c) |*x| x.* = .{ .ref = "A1", .author = "u", .text = "x" };
        try emitVmlDrawingXml(a, &buf, &c);
        try std.testing.expect(std.mem.indexOf(u8, buf.items, "data=\"1\"") != null);
    }
    {
        var buf: std.ArrayListUnmanaged(u8) = .empty;
        defer buf.deinit(a);
        var c: [1024]Comment = undefined;
        for (&c) |*x| x.* = .{ .ref = "A1", .author = "u", .text = "x" };
        try emitVmlDrawingXml(a, &buf, &c);
        try std.testing.expect(std.mem.indexOf(u8, buf.items, "data=\"1,2\"") != null);
    }
}

test "emitVmlDrawingXml: anchor clamp on XFD column" {
    const a = std.testing.allocator;
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(a);
    const comments = [_]Comment{.{ .ref = "XFD1", .author = "x", .text = "x" }};
    try emitVmlDrawingXml(a, &buf, &comments);
    // from_col + to_col both clamp to EXCEL_MAX_COL - 1 = 16383.
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "<x:Anchor>16383, 15, 0, 2, 16383, 31, 4, 3</x:Anchor>") != null);
}

test "isForbiddenXmlByte: rejects C0 control bytes, accepts tab/LF/CR/DEL" {
    try std.testing.expect(isForbiddenXmlByte(0x00));
    try std.testing.expect(isForbiddenXmlByte(0x01));
    try std.testing.expect(isForbiddenXmlByte(0x0B));
    try std.testing.expect(isForbiddenXmlByte(0x1F));
    try std.testing.expect(!isForbiddenXmlByte(0x09));
    try std.testing.expect(!isForbiddenXmlByte(0x0A));
    try std.testing.expect(!isForbiddenXmlByte(0x0D));
    try std.testing.expect(!isForbiddenXmlByte(0x20));
    // 0x7F (DEL) is valid per XML 1.0 production [2] (lies in
    // `[#x20-#xD7FF]`). B3 iter-wr-6 reconciled with the spec.
    try std.testing.expect(!isForbiddenXmlByte(0x7F));
}
