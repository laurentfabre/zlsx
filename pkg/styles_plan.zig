//! Styles extension-plan substrate (B3 iter-wr-2).
//!
//! Storage + dedup + fresh-emit types for the `xl/styles.xml`
//! pipeline. Lives outside `pkg/workbook.zig` so both Workbook
//! (delta-on-existing-bytes editor) and `xlsx.Writer` (fresh-emit
//! producer) can register styles through the same shape without a
//! circular module dependency (workbook.zig imports `zlsx`, which
//! contains writer.zig).
//!
//! Mirrors the iter-wr-1 `pkg/sst_plan.zig` architecture:
//!
//! - **Type-of-record** for `Style` / `Dxf` / `BorderSide` / `BorderStyle`
//!   / `PatternType` / `HAlign` lives here. `src/writer.zig` re-exports
//!   them so the writer's public API surface (`xlsx.Writer.addStyle(...)`
//!   etc.) is unchanged.
//! - **`StylesPlan`** is the registry. `addStyle`, `addDxf`,
//!   `internNumFmt` are the registration entry points. `emit(out)`
//!   produces byte-identical `xl/styles.xml` payload bytes (the
//!   pre-iter-wr-2 emitter from `src/writer.zig`'s `emitStylesXml`
//!   was lifted verbatim).
//!
//! Byte-stability invariants — see `docs/plans/archive/writer-rebase.md` §1.10.
//! The schema element order is rigid (`numFmts → fonts → fills →
//! borders → cellStyleXfs → cellXfs → cellStyles → dxfs`). Every
//! attribute order in `<xf>` and `<dxf>` blocks is fixed. Color
//! encoding is always 8-hex ARGB (`{X:0>8}`). User numFmts start at id
//! `NUM_FMT_BASE` (164). One missed attribute order = "repaired"
//! prompts across the corpus.
//!
//! Stdlib only. Zig 0.15.2.

const std = @import("std");

const Allocator = std.mem.Allocator;
const assert = std.debug.assert;

pub const Error = error{
    OutOfMemory,
    InvalidFontSize,
    InvalidFontName,
    InvalidNumberFormat,
};

/// OOXML reserves numFmtIds 0..=49 for built-ins; user numFmts must
/// start at 164 (Excel's convention — 50..=163 are "reserved").
pub const NUM_FMT_BASE: u32 = 164;

// ─── Public type-of-record (mirrors writer-side names) ──────────────

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

/// Cell style registered via `StylesPlan.addStyle`. Fields default to
/// "unset" so `addStyle(.{ .font_bold = true })` produces the
/// minimum-overhead styles.xml entry.
///
/// `font_name` is caller-owned for the duration of the `addStyle`
/// call; the registrar dupes it into its own pool so callers can free
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

/// A differential format — the font / fill / border overrides applied
/// when a conditional-format rule matches. Registered once via
/// `StylesPlan.addDxf`, referenced by dxfId from one or more rules.
/// Scoped to the subset of properties real conditional formats
/// actually toggle (bold / italic / font color + size / solid fill /
/// per-side borders).
pub const Dxf = struct {
    font_bold: bool = false,
    font_italic: bool = false,
    font_color_argb: ?u32 = null,
    /// Font size in points. Rare in CF rules but cheap to support —
    /// the `<sz val="…"/>` child renders the differential font at
    /// an explicit pt size instead of inheriting the cell style.
    font_size: ?f32 = null,
    fill_fg_argb: ?u32 = null,
    /// Per-side border overrides — emitted inside the dxf's
    /// `<border>` block. Use `.none` (the default) to inherit the
    /// cell's existing border on that side.
    border_left: BorderSide = .{},
    border_right: BorderSide = .{},
    border_top: BorderSide = .{},
    border_bottom: BorderSide = .{},
};

fn hasBorder(s: Style) bool {
    return s.border_left.style != .none or
        s.border_right.style != .none or
        s.border_top.style != .none or
        s.border_bottom.style != .none or
        s.border_diagonal.style != .none or
        s.diagonal_up or s.diagonal_down;
}

/// Content-compare two styles. Necessary because `std.meta.eql` on
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
    if ((a.font_name == null) != (b.font_name == null)) return false;
    if (a.font_name) |an| {
        if (!std.mem.eql(u8, an, b.font_name.?)) return false;
    }
    if ((a.number_format == null) != (b.number_format == null)) return false;
    if (a.number_format) |an| {
        if (!std.mem.eql(u8, an, b.number_format.?)) return false;
    }
    return true;
}

// ─── Static skeleton blobs ──────────────────────────────────────────
//
// Lifted verbatim from the pre-iter-wr-2 `src/writer.zig` constants;
// every byte these emit is the OOXML default backbone the corpus
// expects.

const STYLES_HEAD: []const u8 =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
;
const STYLES_FONTS_DEFAULT: []const u8 =
    \\<font><sz val="11"/><name val="Calibri"/></font>
;
const STYLES_FILLS_DEFAULT: []const u8 =
    \\<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
;
const STYLES_BORDER_DEFAULT: []const u8 =
    \\<border><left/><right/><top/><bottom/><diagonal/></border>
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

// ─── Helper functions ───────────────────────────────────────────────

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

/// Append `s` to `out`, escaping XML metacharacters (`<`, `>`, `&`,
/// `"`, `'`). Mirrors the writer-side `appendXmlEscaped` helper.
/// Caller is responsible for ensuring `s` carries no XML 1.0
/// forbidden control bytes — that's a writer-side intake concern,
/// not a styles emit one.
fn appendXmlEscaped(
    alloc: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    s: []const u8,
) Allocator.Error!void {
    for (s) |b| switch (b) {
        '<' => try out.appendSlice(alloc, "&lt;"),
        '>' => try out.appendSlice(alloc, "&gt;"),
        '&' => try out.appendSlice(alloc, "&amp;"),
        '"' => try out.appendSlice(alloc, "&quot;"),
        '\'' => try out.appendSlice(alloc, "&apos;"),
        else => try out.append(alloc, b),
    };
}

fn emitBorderSide(
    alloc: Allocator,
    buf: *std.ArrayListUnmanaged(u8),
    tag: []const u8,
    side: BorderSide,
) Allocator.Error!void {
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

/// Emit one `<left>` / `<right>` / `<top>` / `<bottom>` element for a
/// dxf `<border>` block. Self-closing when the side is `.none`; opens
/// with `style="…"` + nested `<color>` otherwise.
fn emitDxfBorderSide(
    alloc: Allocator,
    buf: *std.ArrayListUnmanaged(u8),
    tag: []const u8,
    side: BorderSide,
) Allocator.Error!void {
    if (side.style == .none) {
        try buf.print(alloc, "<{s}/>", .{tag});
        return;
    }
    try buf.print(alloc, "<{s} style=\"{s}\">", .{ tag, borderStyleName(side.style) });
    if (side.color_argb) |c| {
        try buf.print(alloc, "<color rgb=\"{X:0>8}\"/>", .{c});
    }
    try buf.print(alloc, "</{s}>", .{tag});
}

/// Where a plan's records land in a stylesheet: the id the first
/// record of each table takes. `fresh` is the fresh part's layout —
/// the OOXML default records in the first slots (`emit` writes them:
/// one font, the two conventional fills, one border, one `<xf>`, no
/// dxf) and custom number formats from `NUM_FMT_BASE`. An existing
/// part's extension starts each table at the count it already holds
/// (`Workbook.stylesBaseline`), so the index `addStyle` hands out is
/// the slot the record takes in the saved part, whichever part that
/// is. The plan itself stays fresh-relative; the mappers translate.
pub const Base = struct {
    /// The `numFmtId` the first interned format takes.
    num_fmt_next: u32 = NUM_FMT_BASE,
    fonts: u32 = 1,
    fills: u32 = 2,
    borders: u32 = 1,
    cell_xfs: u32 = 1,
    dxfs: u32 = 0,

    pub const fresh: Base = .{};

    /// The saved `s="…"` index of the style `addStyle` returned
    /// `fresh_idx` (1-based) for. Identity on `fresh`.
    pub fn xfIndex(self: Base, fresh_idx: u32) u32 {
        assert(fresh_idx >= 1);
        return self.cell_xfs + (fresh_idx - 1);
    }

    /// The saved dxfId of the dxf `addDxf` returned `fresh_id` for.
    pub fn dxfId(self: Base, fresh_id: u32) u32 {
        return self.dxfs + fresh_id;
    }

    /// The saved numFmtId of the format `internNumFmt` returned
    /// `fresh_id` for.
    pub fn numFmtId(self: Base, fresh_id: u32) u32 {
        assert(fresh_id >= NUM_FMT_BASE);
        return self.num_fmt_next + (fresh_id - NUM_FMT_BASE);
    }
};

/// `emitFragments`' output: the children each table gains, and how
/// many — the `count` attribute a table carries is the base plus the
/// added.
pub const Fragments = struct {
    num_fmts: std.ArrayListUnmanaged(u8) = .empty,
    fonts: std.ArrayListUnmanaged(u8) = .empty,
    fills: std.ArrayListUnmanaged(u8) = .empty,
    borders: std.ArrayListUnmanaged(u8) = .empty,
    cell_xfs: std.ArrayListUnmanaged(u8) = .empty,
    dxfs: std.ArrayListUnmanaged(u8) = .empty,
    num_fmts_added: u32 = 0,
    fonts_added: u32 = 0,
    fills_added: u32 = 0,
    borders_added: u32 = 0,
    cell_xfs_added: u32 = 0,
    dxfs_added: u32 = 0,

    pub fn deinit(self: *Fragments, allocator: Allocator) void {
        self.num_fmts.deinit(allocator);
        self.fonts.deinit(allocator);
        self.fills.deinit(allocator);
        self.borders.deinit(allocator);
        self.cell_xfs.deinit(allocator);
        self.dxfs.deinit(allocator);
        self.* = undefined;
    }
};

// ─── StylesPlan registry ────────────────────────────────────────────

/// Fresh-emit styles registry. Mirrors the pre-iter-wr-2
/// `xlsx.Writer.{styles, dxfs, num_fmts, num_fmt_index}` quartet of
/// fields plus the deduping `addStyle` / `addDxf` / `internNumFmt`
/// methods.
///
/// Indexing contract:
/// - `addStyle` returns a **1-based** index — slot 0 in the emitted
///   `<cellXfs>` is reserved for the default no-style record. The
///   returned value can be fed directly to a cell's `s="N"` attribute.
/// - `addDxf` returns a **0-based** dxf id — the index into the
///   emitted `<dxfs>` block.
/// - `internNumFmt` returns a numFmtId starting at `NUM_FMT_BASE`
///   (164), incrementing for each unique format string.
///
/// Allocator owns: every `font_name` / `number_format` slice on
/// staged styles (duped on insert), every entry of `num_fmts` (duped
/// on insert), the side-index hashmap. Dxfs carry no owned slices.
pub const StylesPlan = struct {
    /// Registered styles (unique). Index 0 in the emitted `<cellXfs>`
    /// is the default no-style entry; user styles start at 1 so the
    /// value returned from `addStyle()` can be used directly as the
    /// cell's `s="N"` attribute.
    styles: std.ArrayListUnmanaged(Style) = .empty,
    /// Differential formats used by conditional-formatting rules.
    /// One entry per unique Dxf; rules reference them by 0-based dxfId.
    dxfs: std.ArrayListUnmanaged(Dxf) = .empty,
    /// Number-format pool (stage 5). Owns the format strings.
    num_fmts: std.ArrayListUnmanaged([]u8) = .empty,
    /// Side-index over `num_fmts`: format string → numFmtId. All
    /// values are unique and >= `NUM_FMT_BASE`.
    num_fmt_index: std.StringHashMapUnmanaged(u32) = .empty,

    pub fn deinit(self: *StylesPlan, allocator: Allocator) void {
        for (self.styles.items) |s| {
            if (s.font_name) |n| allocator.free(n);
            if (s.number_format) |n| allocator.free(n);
        }
        self.styles.deinit(allocator);
        self.dxfs.deinit(allocator);
        for (self.num_fmts.items) |n| allocator.free(n);
        self.num_fmts.deinit(allocator);
        self.num_fmt_index.deinit(allocator);
        self.* = undefined;
    }

    pub fn isEmpty(self: *const StylesPlan) bool {
        return self.styles.items.len == 0 and self.dxfs.items.len == 0;
    }

    /// Whether a save has records to render — a style, a dxf or a
    /// number format. (`isEmpty` is the fresh emitter's question: a
    /// plan holding formats alone writes no part there.)
    pub fn hasWork(self: *const StylesPlan) bool {
        return self.styles.items.len > 0 or self.dxfs.items.len > 0 or self.num_fmts.items.len > 0;
    }

    /// Register a cell style and return its `s="…"` index. Dedupes
    /// structurally (including content-comparing `font_name` /
    /// `number_format`, not just slice-header comparing). Returning
    /// value is 1-based — cellXfs[0] is reserved for the default
    /// no-style record.
    ///
    /// Side effect: when `style.number_format` is set, the format
    /// string is registered into the numFmt pool *before* dedup of
    /// the parent Style runs, so a rejected style doesn't pollute
    /// the format pool.
    pub fn addStyle(self: *StylesPlan, allocator: Allocator, style: Style) Error!u32 {
        if (style.font_size) |s| {
            if (!std.math.isFinite(s) or s <= 0) return error.InvalidFontSize;
        }
        if (style.font_name) |n| {
            if (n.len == 0) return error.InvalidFontName;
        }
        if (style.number_format) |n| {
            if (n.len == 0) return error.InvalidNumberFormat;
        }

        if (style.number_format) |fmt| {
            _ = try self.internNumFmt(allocator, fmt);
        }

        // Linear-scan dedup (matches pre-iter-wr-2 writer behaviour).
        for (self.styles.items, 0..) |existing, i| {
            if (stylesEqual(existing, style)) return @intCast(i + 1);
        }

        var owned_style = style;
        if (style.font_name) |n| {
            owned_style.font_name = try allocator.dupe(u8, n);
        }
        errdefer if (owned_style.font_name) |n| allocator.free(n);
        if (style.number_format) |n| {
            owned_style.number_format = try allocator.dupe(u8, n);
        }
        errdefer if (owned_style.number_format) |n| allocator.free(n);
        try self.styles.append(allocator, owned_style);
        return @intCast(self.styles.items.len);
    }

    /// Register a differential format (font / fill / border overrides
    /// applied when a conditional-format rule matches) and return its
    /// 0-based dxfId. Linear dedup by content equality — Dxfs carry
    /// no owned slices, so a `std.meta.eql` is sufficient.
    pub fn addDxf(self: *StylesPlan, allocator: Allocator, dxf: Dxf) Error!u32 {
        for (self.dxfs.items, 0..) |existing, i| {
            if (std.meta.eql(existing, dxf)) return @intCast(i);
        }
        try self.dxfs.append(allocator, dxf);
        return @intCast(self.dxfs.items.len - 1);
    }

    /// Return the numFmtId for `fmt`, allocating a new entry at id >=
    /// `NUM_FMT_BASE` (164) on first sight. Subsequent calls with the
    /// same content return the same id.
    pub fn internNumFmt(self: *StylesPlan, allocator: Allocator, fmt: []const u8) Error!u32 {
        if (self.num_fmt_index.get(fmt)) |id| return id;
        const owned = try allocator.dupe(u8, fmt);
        errdefer allocator.free(owned);
        const id: u32 = @intCast(NUM_FMT_BASE + self.num_fmts.items.len);
        // The index's slot is reserved before the pool takes the string:
        // once appended, the pool owns it, and a failing `put` after the
        // append left the `errdefer` freeing what `deinit` frees again
        // (S3d slice 1's allocation sweep).
        try self.num_fmt_index.ensureUnusedCapacity(allocator, 1);
        try self.num_fmts.append(allocator, owned);
        self.num_fmt_index.putAssumeCapacity(owned, id);
        return id;
    }

    /// Emit `xl/styles.xml` to `out` — the fresh part: the OOXML
    /// default records in the first slots of every table, then this
    /// plan's records after them (`Base.fresh`). Byte-identical to the
    /// pre-iter-wr-2 `emitStylesXml` function from `src/writer.zig`.
    /// Caller owns `out`; this function only appends.
    ///
    /// Element order (rigid OOXML `CT_Stylesheet` schema):
    /// `numFmts → fonts → fills → borders → cellStyleXfs → cellXfs →
    /// cellStyles → dxfs`. `<cellStyles>` MUST sit between
    /// `<cellXfs>` and `<dxfs>` even when `<dxfs>` is absent — strict-
    /// mode validators reject otherwise.
    pub fn emit(
        self: *const StylesPlan,
        allocator: Allocator,
        out: *std.ArrayListUnmanaged(u8),
    ) Allocator.Error!void {
        var frags = try self.emitFragments(allocator, Base.fresh);
        defer frags.deinit(allocator);

        try out.appendSlice(allocator, STYLES_HEAD);

        // <numFmts> — emitted only when the user registered any custom
        // format. Built-ins (General / 0..=49) don't go here.
        if (frags.num_fmts_added > 0) {
            try out.print(allocator, "<numFmts count=\"{d}\">", .{frags.num_fmts_added});
            try out.appendSlice(allocator, frags.num_fmts.items);
            try out.appendSlice(allocator, "</numFmts>");
        }

        // <fonts>: default at index 0 + one per user style.
        try out.print(allocator, "<fonts count=\"{d}\">", .{Base.fresh.fonts + frags.fonts_added});
        try out.appendSlice(allocator, STYLES_FONTS_DEFAULT);
        try out.appendSlice(allocator, frags.fonts.items);
        try out.appendSlice(allocator, "</fonts>");

        // <fills>: 2 reserved slots (none, gray125 — conventional OOXML
        // defaults), then one user fill per style that sets any fill
        // field. Styles without a fill reference fillId=0.
        try out.print(allocator, "<fills count=\"{d}\">", .{Base.fresh.fills + frags.fills_added});
        try out.appendSlice(allocator, STYLES_FILLS_DEFAULT);
        try out.appendSlice(allocator, frags.fills.items);
        try out.appendSlice(allocator, "</fills>");

        // <borders>: default empty border at index 0, then one per
        // style that touches any border field. Styles without borders
        // keep borderId=0.
        try out.print(allocator, "<borders count=\"{d}\">", .{Base.fresh.borders + frags.borders_added});
        try out.appendSlice(allocator, STYLES_BORDER_DEFAULT);
        try out.appendSlice(allocator, frags.borders.items);
        try out.appendSlice(allocator, "</borders>");
        try out.appendSlice(allocator, STYLES_CELL_STYLE_XFS);

        // <cellXfs>: default at index 0 + one per user style.
        try out.print(allocator, "<cellXfs count=\"{d}\">", .{Base.fresh.cell_xfs + frags.cell_xfs_added});
        try out.appendSlice(allocator, STYLES_DEFAULT_CELL_XF);
        try out.appendSlice(allocator, frags.cell_xfs.items);
        try out.appendSlice(allocator, "</cellXfs>");

        // <cellStyles> sits between <cellXfs> and <dxfs> per the
        // OOXML stylesheet element-order schema. Strict-mode
        // validators reject styles.xml without it.
        try out.appendSlice(allocator, STYLES_CELL_STYLES);

        // <dxfs> — differential formats for conditional-formatting
        // rules. Emitted last per the schema.
        if (frags.dxfs_added > 0) {
            try out.print(allocator, "<dxfs count=\"{d}\">", .{frags.dxfs_added});
            try out.appendSlice(allocator, frags.dxfs.items);
            try out.appendSlice(allocator, "</dxfs>");
        }

        try out.appendSlice(allocator, STYLES_TAIL);
    }

    /// This plan's records as the children each stylesheet table
    /// gains, every cross-table id resolved against `base`: the font,
    /// fill and border a style's `<xf>` names are the ones this call
    /// appends, numbered from the slot each table's next record takes.
    /// The fresh part (`emit`) lays them after the OOXML defaults; an
    /// existing part's extension (`Workbook.save`) lays them after the
    /// records it already holds. One emitter for both, so a registered
    /// style is the same bytes whichever part it lands in — the fresh
    /// parity pins hold the extension to the writer's output.
    pub fn emitFragments(
        self: *const StylesPlan,
        allocator: Allocator,
        base: Base,
    ) Allocator.Error!Fragments {
        const styles = self.styles.items;
        const num_fmts: []const []u8 = self.num_fmts.items;
        const dxfs = self.dxfs.items;

        var frags: Fragments = .{};
        errdefer frags.deinit(allocator);

        for (num_fmts, 0..) |fmt, i| {
            const id: u32 = base.num_fmt_next + @as(u32, @intCast(i));
            try frags.num_fmts.print(allocator, "<numFmt numFmtId=\"{d}\" formatCode=\"", .{id});
            try appendXmlEscaped(allocator, &frags.num_fmts, fmt);
            try frags.num_fmts.appendSlice(allocator, "\"/>");
        }
        frags.num_fmts_added = @intCast(num_fmts.len);

        for (styles) |s| {
            const out = &frags.fonts;
            try out.appendSlice(allocator, "<font>");
            if (s.font_bold) try out.appendSlice(allocator, "<b/>");
            if (s.font_italic) try out.appendSlice(allocator, "<i/>");
            const size = s.font_size orelse 11.0;
            try out.print(allocator, "<sz val=\"{d}\"/>", .{size});
            if (s.font_color_argb) |c| try out.print(
                allocator,
                "<color rgb=\"{X:0>8}\"/>",
                .{c},
            );
            try out.appendSlice(allocator, "<name val=\"");
            if (s.font_name) |n| {
                try appendXmlEscaped(allocator, out, n);
            } else {
                try out.appendSlice(allocator, "Calibri");
            }
            try out.appendSlice(allocator, "\"/></font>");
        }
        frags.fonts_added = @intCast(styles.len);

        // One user fill per style that sets any fill field; styles
        // without a fill reference fillId=0.
        var fill_ids = try allocator.alloc(u32, styles.len);
        defer allocator.free(fill_ids);
        var next_user_fill_id: u32 = base.fills;
        for (styles, 0..) |s, i| {
            if (s.fill_pattern != .none or s.fill_fg_argb != null or s.fill_bg_argb != null) {
                fill_ids[i] = next_user_fill_id;
                next_user_fill_id += 1;
            } else {
                fill_ids[i] = 0;
            }
        }
        for (styles) |s| {
            const out = &frags.fills;
            if (s.fill_pattern == .none and s.fill_fg_argb == null and s.fill_bg_argb == null) continue;
            try out.print(
                allocator,
                "<fill><patternFill patternType=\"{s}\"",
                .{patternTypeName(s.fill_pattern)},
            );
            if (s.fill_fg_argb == null and s.fill_bg_argb == null) {
                try out.appendSlice(allocator, "/></fill>");
            } else {
                try out.appendSlice(allocator, ">");
                if (s.fill_fg_argb) |c| try out.print(allocator, "<fgColor rgb=\"{X:0>8}\"/>", .{c});
                if (s.fill_bg_argb) |c| try out.print(allocator, "<bgColor rgb=\"{X:0>8}\"/>", .{c});
                try out.appendSlice(allocator, "</patternFill></fill>");
            }
        }
        frags.fills_added = next_user_fill_id - base.fills;

        // One border per style that touches any border field; styles
        // without borders keep borderId=0.
        var border_ids = try allocator.alloc(u32, styles.len);
        defer allocator.free(border_ids);
        var next_user_border_id: u32 = base.borders;
        for (styles, 0..) |s, i| {
            if (hasBorder(s)) {
                border_ids[i] = next_user_border_id;
                next_user_border_id += 1;
            } else {
                border_ids[i] = 0;
            }
        }
        for (styles) |s| {
            const out = &frags.borders;
            if (!hasBorder(s)) continue;
            try out.appendSlice(allocator, "<border");
            if (s.diagonal_up) try out.appendSlice(allocator, " diagonalUp=\"1\"");
            if (s.diagonal_down) try out.appendSlice(allocator, " diagonalDown=\"1\"");
            try out.appendSlice(allocator, ">");
            try emitBorderSide(allocator, out, "left", s.border_left);
            try emitBorderSide(allocator, out, "right", s.border_right);
            try emitBorderSide(allocator, out, "top", s.border_top);
            try emitBorderSide(allocator, out, "bottom", s.border_bottom);
            try emitBorderSide(allocator, out, "diagonal", s.border_diagonal);
            try out.appendSlice(allocator, "</border>");
        }
        frags.borders_added = next_user_border_id - base.borders;

        // One <xf> per style, in registration order — the slot each
        // takes is `base.cell_xfs + i`, the index `addStyle` mapped.
        for (styles, 0..) |s, i| {
            const out = &frags.cell_xfs;
            const has_alignment = s.alignment_horizontal != .general or s.wrap_text;
            const fill_id = fill_ids[i];
            const border_id = border_ids[i];
            const num_fmt_id: u32 = if (s.number_format) |fmt|
                (if (self.num_fmt_index.get(fmt)) |fresh| base.numFmtId(fresh) else 0)
            else
                0;
            const font_id: u32 = base.fonts + @as(u32, @intCast(i));
            try out.print(
                allocator,
                "<xf numFmtId=\"{d}\" fontId=\"{d}\" fillId=\"{d}\" borderId=\"{d}\" xfId=\"0\" applyFont=\"1\"",
                .{ num_fmt_id, font_id, fill_id, border_id },
            );
            if (num_fmt_id != 0) try out.appendSlice(allocator, " applyNumberFormat=\"1\"");
            if (fill_id != 0) try out.appendSlice(allocator, " applyFill=\"1\"");
            if (border_id != 0) try out.appendSlice(allocator, " applyBorder=\"1\"");
            if (has_alignment) {
                try out.appendSlice(allocator, " applyAlignment=\"1\"><alignment");
                if (s.alignment_horizontal != .general) {
                    try out.print(allocator, " horizontal=\"{s}\"", .{hAlignName(s.alignment_horizontal)});
                }
                if (s.wrap_text) try out.appendSlice(allocator, " wrapText=\"1\"");
                try out.appendSlice(allocator, "/></xf>");
            } else {
                try out.appendSlice(allocator, "/>");
            }
        }
        frags.cell_xfs_added = @intCast(styles.len);

        for (dxfs) |dxf| {
            const out = &frags.dxfs;
            try out.appendSlice(allocator, "<dxf>");
            const has_font = dxf.font_bold or dxf.font_italic or
                dxf.font_color_argb != null or dxf.font_size != null;
            if (has_font) {
                try out.appendSlice(allocator, "<font>");
                if (dxf.font_bold) try out.appendSlice(allocator, "<b/>");
                if (dxf.font_italic) try out.appendSlice(allocator, "<i/>");
                if (dxf.font_color_argb) |c| {
                    try out.print(allocator, "<color rgb=\"{X:0>8}\"/>", .{c});
                }
                if (dxf.font_size) |sz| {
                    try out.print(allocator, "<sz val=\"{d}\"/>", .{sz});
                }
                try out.appendSlice(allocator, "</font>");
            }
            if (dxf.fill_fg_argb) |fg| {
                try out.print(
                    allocator,
                    "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"{X:0>8}\"/><bgColor rgb=\"{X:0>8}\"/></patternFill></fill>",
                    .{ fg, fg },
                );
            }
            const has_dxf_border = dxf.border_left.style != .none or
                dxf.border_right.style != .none or
                dxf.border_top.style != .none or
                dxf.border_bottom.style != .none;
            if (has_dxf_border) {
                try out.appendSlice(allocator, "<border>");
                try emitDxfBorderSide(allocator, out, "left", dxf.border_left);
                try emitDxfBorderSide(allocator, out, "right", dxf.border_right);
                try emitDxfBorderSide(allocator, out, "top", dxf.border_top);
                try emitDxfBorderSide(allocator, out, "bottom", dxf.border_bottom);
                try out.appendSlice(allocator, "</border>");
            }
            try out.appendSlice(allocator, "</dxf>");
        }
        frags.dxfs_added = @intCast(dxfs.len);

        return frags;
    }
};

// ─── Tests ────────────────────────────────────────────────────────────

test "StylesPlan: addStyle returns 1-based index, dedupes by content" {
    const a = std.testing.allocator;
    var plan: StylesPlan = .{};
    defer plan.deinit(a);

    const idx1 = try plan.addStyle(a, .{ .font_bold = true });
    const idx2 = try plan.addStyle(a, .{ .font_italic = true });
    const idx3 = try plan.addStyle(a, .{ .font_bold = true }); // dedup

    try std.testing.expectEqual(@as(u32, 1), idx1);
    try std.testing.expectEqual(@as(u32, 2), idx2);
    try std.testing.expectEqual(@as(u32, 1), idx3);
    try std.testing.expectEqual(@as(usize, 2), plan.styles.items.len);
}

test "StylesPlan: addStyle dedups font_name by content (not slice header)" {
    const a = std.testing.allocator;
    var plan: StylesPlan = .{};
    defer plan.deinit(a);

    var name_a = [_]u8{ 'A', 'r', 'i', 'a', 'l' };
    var name_b = [_]u8{ 'A', 'r', 'i', 'a', 'l' };
    const idx1 = try plan.addStyle(a, .{ .font_name = name_a[0..] });
    const idx2 = try plan.addStyle(a, .{ .font_name = name_b[0..] });
    try std.testing.expectEqual(idx1, idx2);
}

test "StylesPlan: addStyle rejects invalid inputs" {
    const a = std.testing.allocator;
    var plan: StylesPlan = .{};
    defer plan.deinit(a);

    try std.testing.expectError(
        error.InvalidFontSize,
        plan.addStyle(a, .{ .font_size = -1.0 }),
    );
    try std.testing.expectError(
        error.InvalidFontName,
        plan.addStyle(a, .{ .font_name = "" }),
    );
    try std.testing.expectError(
        error.InvalidNumberFormat,
        plan.addStyle(a, .{ .number_format = "" }),
    );
}

test "StylesPlan: addDxf returns 0-based id, dedupes" {
    const a = std.testing.allocator;
    var plan: StylesPlan = .{};
    defer plan.deinit(a);

    const id1 = try plan.addDxf(a, .{ .font_bold = true });
    const id2 = try plan.addDxf(a, .{ .font_italic = true });
    const id3 = try plan.addDxf(a, .{ .font_bold = true });

    try std.testing.expectEqual(@as(u32, 0), id1);
    try std.testing.expectEqual(@as(u32, 1), id2);
    try std.testing.expectEqual(@as(u32, 0), id3);
}

test "StylesPlan: internNumFmt assigns ids starting at NUM_FMT_BASE" {
    const a = std.testing.allocator;
    var plan: StylesPlan = .{};
    defer plan.deinit(a);

    const id1 = try plan.internNumFmt(a, "0.00");
    const id2 = try plan.internNumFmt(a, "yyyy-mm-dd");
    const id3 = try plan.internNumFmt(a, "0.00"); // dedup

    try std.testing.expectEqual(NUM_FMT_BASE, id1);
    try std.testing.expectEqual(NUM_FMT_BASE + 1, id2);
    try std.testing.expectEqual(NUM_FMT_BASE, id3);
}

test "StylesPlan: emit produces well-formed empty stylesheet (no styles, no dxfs)" {
    const a = std.testing.allocator;
    var plan: StylesPlan = .{};
    defer plan.deinit(a);

    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(a);
    try plan.emit(a, &buf);

    // Sanity-checks: schema element order, mandatory cellStyles block.
    const out = buf.items;
    try std.testing.expect(std.mem.indexOf(u8, out, "<fonts count=\"1\">") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "<fills count=\"2\">") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "<borders count=\"1\">") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "<cellXfs count=\"1\">") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "<cellStyles count=\"1\">") != null);
    // <numFmts> is omitted when empty.
    try std.testing.expect(std.mem.indexOf(u8, out, "<numFmts") == null);
    // <dxfs> is omitted when empty.
    try std.testing.expect(std.mem.indexOf(u8, out, "<dxfs") == null);
    // cellStyles MUST come before the closing </styleSheet>.
    const cs_pos = std.mem.indexOf(u8, out, "<cellStyles").?;
    const tail_pos = std.mem.indexOf(u8, out, "</styleSheet>").?;
    try std.testing.expect(cs_pos < tail_pos);
}

test "StylesPlan: emit element ordering with both styles and dxfs" {
    const a = std.testing.allocator;
    var plan: StylesPlan = .{};
    defer plan.deinit(a);

    _ = try plan.addStyle(a, .{ .font_bold = true });
    _ = try plan.addDxf(a, .{ .font_bold = true });

    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(a);
    try plan.emit(a, &buf);

    const out = buf.items;
    // numFmts → fonts → fills → borders → cellStyleXfs → cellXfs →
    // cellStyles → dxfs. cellStyles MUST sit between cellXfs and dxfs.
    const cellxfs_pos = std.mem.indexOf(u8, out, "<cellXfs").?;
    const cellstyles_pos = std.mem.indexOf(u8, out, "<cellStyles count=\"1\">").?;
    const dxfs_pos = std.mem.indexOf(u8, out, "<dxfs count=\"1\">").?;
    try std.testing.expect(cellxfs_pos < cellstyles_pos);
    try std.testing.expect(cellstyles_pos < dxfs_pos);
}

test "StylesPlan: emit styles with custom number_format registers numFmt at id 164" {
    const a = std.testing.allocator;
    var plan: StylesPlan = .{};
    defer plan.deinit(a);

    _ = try plan.addStyle(a, .{ .number_format = "0.00%" });

    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(a);
    try plan.emit(a, &buf);

    try std.testing.expect(std.mem.indexOf(u8, buf.items, "<numFmts count=\"1\">") != null);
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "numFmtId=\"164\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, buf.items, "formatCode=\"0.00%\"") != null);
}

test "StylesPlan: emitFragments against an existing part's layout numbers every cross-table id from the base" {
    const a = std.testing.allocator;
    var plan: StylesPlan = .{};
    defer plan.deinit(a);
    _ = try plan.addStyle(a, .{ .font_bold = true, .fill_pattern = .solid, .fill_fg_argb = 0xFF0000FF, .border_bottom = .{ .style = .dashed }, .number_format = "0.0" });
    _ = try plan.addStyle(a, .{ .font_italic = true });
    _ = try plan.addDxf(a, .{ .font_bold = true });
    const base: Base = .{ .num_fmt_next = 170, .fonts = 5, .fills = 3, .borders = 2, .cell_xfs = 7, .dxfs = 4 };
    var frags = try plan.emitFragments(a, base);
    defer frags.deinit(a);
    try std.testing.expectEqualStrings("<numFmt numFmtId=\"170\" formatCode=\"0.0\"/>", frags.num_fmts.items);
    try std.testing.expectEqualStrings(
        "<xf numFmtId=\"170\" fontId=\"5\" fillId=\"3\" borderId=\"2\" xfId=\"0\" applyFont=\"1\" applyNumberFormat=\"1\" applyFill=\"1\" applyBorder=\"1\"/>" ++
            "<xf numFmtId=\"0\" fontId=\"6\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>",
        frags.cell_xfs.items,
    );
    try std.testing.expectEqual(@as(u32, 1), frags.num_fmts_added);
    try std.testing.expectEqual(@as(u32, 2), frags.fonts_added);
    try std.testing.expectEqual(@as(u32, 1), frags.fills_added);
    try std.testing.expectEqual(@as(u32, 1), frags.borders_added);
    try std.testing.expectEqual(@as(u32, 2), frags.cell_xfs_added);
    try std.testing.expectEqual(@as(u32, 1), frags.dxfs_added);
    // The mappers: the index the registrations return for those slots.
    try std.testing.expectEqual(@as(u32, 7), base.xfIndex(1));
    try std.testing.expectEqual(@as(u32, 8), base.xfIndex(2));
    try std.testing.expectEqual(@as(u32, 4), base.dxfId(0));
    try std.testing.expectEqual(@as(u32, 170), base.numFmtId(NUM_FMT_BASE));
    // The fresh base is the identity, and the fresh emit is the
    // fragments behind the defaults.
    try std.testing.expectEqual(@as(u32, 1), Base.fresh.xfIndex(1));
    try std.testing.expectEqual(@as(u32, 0), Base.fresh.dxfId(0));
    try std.testing.expectEqual(NUM_FMT_BASE, Base.fresh.numFmtId(NUM_FMT_BASE));
    var fresh_frags = try plan.emitFragments(a, Base.fresh);
    defer fresh_frags.deinit(a);
    var whole: std.ArrayListUnmanaged(u8) = .empty;
    defer whole.deinit(a);
    try plan.emit(a, &whole);
    try std.testing.expect(std.mem.indexOf(u8, whole.items, fresh_frags.cell_xfs.items) != null);
    try std.testing.expect(std.mem.indexOf(u8, whole.items, fresh_frags.fonts.items) != null);
    try std.testing.expect(std.mem.indexOf(u8, whole.items, fresh_frags.dxfs.items) != null);
}
