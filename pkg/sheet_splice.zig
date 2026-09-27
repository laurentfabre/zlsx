//! S3d slice 2 — the per-sheet registrations on an OPENED sheet part.
//!
//! `Worksheet.setColumnWidth` / `setRowHeight` / `freezePanes` /
//! `setAutoFilter` / `addMergedCell` / `addHyperlink` /
//! `addInternalHyperlink` / `addComment` / `addDataValidation*` /
//! `addConditionalFormat*` register into `sheet_plan.SheetState` — the
//! fresh emitter's registry, which `emitWorksheetXml` renders whole.
//! An opened sheet part already has its elements; this module lands
//! the same registrations INTO it, every other byte preserved: each
//! element the fresh emitter would write is spelled by the same writer
//! (`sheet_plan.emit*Element*`), extended in place where the part
//! holds it, created at its CT_Worksheet schema slot where it does
//! not (after the last present predecessor, else right after the
//! root's open tag). The comments part, the VML drawing and the
//! sheet's relationships are extended the same way.
//!
//! The walk is lexical and reads the part as the reader reads it —
//! comments, CDATA, processing instructions and DOCTYPEs skipped,
//! quoted `>` respected — over the root's DIRECT children only, by
//! depth: an element's name inside an `<extLst>` extension is never
//! the sheet's (the S3d slice 1 rule). A part the walk cannot read in
//! place — no `<worksheet>` root, a self-closed one, an element that
//! never closes, an owned element under a prefix or inside a
//! markup-compatibility block — refuses `MalformedSheetXml`, and the
//! caller judges it at the first registration, nothing staged.
const std = @import("std");
const Allocator = std.mem.Allocator;
const assert = std.debug.assert;
const store_mod = @import("store.zig");
const wbxml = @import("typed_parts/workbook_xml.zig");
const sheet_plan = @import("zlsx_sheet_plan");

pub const Error = sheet_plan.Error || error{
    MalformedSheetXml,
    MalformedSheetRels,
    MalformedCommentsXml,
    MalformedVmlDrawing,
    IdSpaceExhausted,
};

/// The walk's own verdict; each caller names the part it was reading.
const ScanError = error{Malformed};

// ─── The element walk ────────────────────────────────────────────────

/// One element as the walk saw it. Every index is into the part.
pub const Element = struct {
    /// The tag's whole name, prefix included. Borrows the part.
    name: []const u8,
    /// The open tag's `<`.
    lt: usize,
    /// One past the open tag's `>`.
    open_end: usize,
    self_closing: bool,
    /// The close tag's `<`; `open_end` when self-closing.
    close_lt: usize,
    /// One past the element's last byte.
    end: usize,

    /// The open tag's attribute region — after the name, before the
    /// `>` (or the `/>`).
    pub fn attrs(self: Element, xml: []const u8) []const u8 {
        const from = self.lt + 1 + self.name.len;
        const to = if (self.self_closing) self.open_end - 2 else self.open_end - 1;
        return xml[from..@max(from, to)];
    }

    pub fn localName(self: Element) []const u8 {
        if (std.mem.indexOfScalar(u8, self.name, ':')) |c| return self.name[c + 1 ..];
        return self.name;
    }
};

fn isNameEnd(c: u8) bool {
    return std.ascii.isWhitespace(c) or c == '/' or c == '>';
}

/// The next element opening at `from` at the CURRENT level: decoys
/// skipped; null at a closing tag (the enclosing element's end) or at
/// `limit`. Malformed when the open tag never closes, the element
/// never closes, or the bytes are not markup.
fn nextElement(xml: []const u8, from: usize, limit: usize) ScanError!?Element {
    var i = from;
    while (i < limit) {
        const lt = std.mem.indexOfScalarPos(u8, xml, i, '<') orelse return null;
        if (lt >= limit) return null;
        const skip_to = wbxml.skipNonElement(xml, lt) catch return error.Malformed;
        if (skip_to != lt) {
            i = skip_to;
            continue;
        }
        if (lt + 1 >= xml.len) return error.Malformed;
        if (xml[lt + 1] == '/') return null;
        const gt = store_mod.xmlStartTagEnd(xml, lt) orelse return error.Malformed;
        var name_end = lt + 1;
        while (name_end < gt and !isNameEnd(xml[name_end])) name_end += 1;
        if (name_end == lt + 1) return error.Malformed;
        const self_closing = gt > lt + 1 and xml[gt - 1] == '/';
        var el: Element = .{
            .name = xml[lt + 1 .. name_end],
            .lt = lt,
            .open_end = gt + 1,
            .self_closing = self_closing,
            .close_lt = gt + 1,
            .end = gt + 1,
        };
        if (!self_closing) {
            const close = try elementClose(xml, gt + 1);
            el.close_lt = close.lt;
            el.end = close.end;
        }
        return el;
    }
    return null;
}

const CloseHit = struct { lt: usize, end: usize };

/// The close tag of the element whose body starts at `from`: a depth
/// walk over real markup. Malformed when the part ends first.
fn elementClose(xml: []const u8, from: usize) ScanError!CloseHit {
    var depth: usize = 0;
    var i = from;
    while (true) {
        const lt = std.mem.indexOfScalarPos(u8, xml, i, '<') orelse return error.Malformed;
        const skip_to = wbxml.skipNonElement(xml, lt) catch return error.Malformed;
        if (skip_to != lt) {
            i = skip_to;
            continue;
        }
        if (lt + 1 >= xml.len) return error.Malformed;
        if (xml[lt + 1] == '/') {
            const gt = std.mem.indexOfScalarPos(u8, xml, lt, '>') orelse return error.Malformed;
            if (depth == 0) return .{ .lt = lt, .end = gt + 1 };
            depth -= 1;
            i = gt + 1;
            continue;
        }
        const gt = store_mod.xmlStartTagEnd(xml, lt) orelse return error.Malformed;
        if (!(gt > lt + 1 and xml[gt - 1] == '/')) depth += 1;
        i = gt + 1;
    }
}

/// The elements of one level, `from` to `limit`, in document order.
fn children(a: Allocator, xml: []const u8, from: usize, limit: usize) (ScanError || Allocator.Error)![]Element {
    var list: std.ArrayListUnmanaged(Element) = .empty;
    errdefer list.deinit(a);
    var i = from;
    while (try nextElement(xml, i, limit)) |el| {
        try list.append(a, el);
        i = el.end;
    }
    return try list.toOwnedSlice(a);
}

/// The part's root: the first element, which must have a body.
fn rootElement(xml: []const u8) ScanError!Element {
    const root = (try nextElement(xml, 0, xml.len)) orelse return error.Malformed;
    if (root.self_closing) return error.Malformed;
    return root;
}

fn childNamed(list: []const Element, name: []const u8) ?Element {
    for (list) |el| if (std.mem.eql(u8, el.name, name)) return el;
    return null;
}

fn countNamed(list: []const Element, name: []const u8) usize {
    var n: usize = 0;
    for (list) |el| {
        if (std.mem.eql(u8, el.name, name)) n += 1;
    }
    return n;
}

// ─── The CT_Worksheet order ──────────────────────────────────────────

/// ECMA-376 CT_Worksheet's child sequence. An absent element is
/// created after the last present element that precedes it here.
const worksheet_order = [_][]const u8{
    "sheetPr",          "dimension",             "sheetViews",      "sheetFormatPr",    "cols",
    "sheetData",        "sheetCalcPr",           "sheetProtection", "protectedRanges",  "scenarios",
    "autoFilter",       "sortState",             "dataConsolidate", "customSheetViews", "mergeCells",
    "phoneticPr",       "conditionalFormatting", "dataValidations", "hyperlinks",       "printOptions",
    "pageMargins",      "pageSetup",             "headerFooter",    "rowBreaks",        "colBreaks",
    "customProperties", "cellWatches",           "ignoredErrors",   "smartTags",        "drawing",
    "legacyDrawing",    "legacyDrawingHF",       "drawingHF",       "picture",          "oleObjects",
    "controls",         "webPublishItems",       "tableParts",      "extLst",
};

fn rankOf(name: []const u8) ?usize {
    for (worksheet_order, 0..) |n, i| if (std.mem.eql(u8, n, name)) return i;
    return null;
}

/// The insertion point for an absent `target`: one past the last
/// direct child ranked before it, else right after the root's open tag.
fn slotFor(list: []const Element, target: []const u8, root_open_end: usize) usize {
    const target_rank = rankOf(target).?;
    var pos = root_open_end;
    for (list) |el| {
        const r = rankOf(el.name) orelse continue;
        if (r < target_rank) pos = el.end;
    }
    return pos;
}

/// The elements this module writes or rewrites: one of them under a
/// prefix, or inside a markup-compatibility block, is a part the
/// splice cannot extend in place (the S3d slice 1 r7 / r8 rule — the
/// walk would read it as absent and write a second one beside it).
const owned_elements = [_][]const u8{
    "sheetViews", "cols", "sheetData", "autoFilter", "mergeCells", "conditionalFormatting", "dataValidations", "hyperlinks", "legacyDrawing",
};

fn isOwnedLocal(name: []const u8) bool {
    for (owned_elements) |n| if (std.mem.eql(u8, n, name)) return true;
    return false;
}

/// True when `body` spells `<name` (as real markup) for an owned
/// element — the test a `mc:AlternateContent` child is put to.
fn bodyMentionsOwned(body: []const u8) bool {
    var i: usize = 0;
    while (std.mem.indexOfScalarPos(u8, body, i, '<')) |lt| {
        // A comment or CDATA spelling one of our names is not markup
        // (in-house r1 A-SPL-103).
        const skip_to = wbxml.skipNonElement(body, lt) catch return true;
        if (skip_to != lt) {
            i = skip_to;
            continue;
        }
        i = lt + 1;
        var name_end = lt + 1;
        while (name_end < body.len and !isNameEnd(body[name_end])) name_end += 1;
        var name = body[lt + 1 .. name_end];
        if (name.len > 0 and name[0] == '/') name = name[1..];
        if (std.mem.indexOfScalar(u8, name, ':')) |c| name = name[c + 1 ..];
        if (isOwnedLocal(name)) return true;
    }
    return false;
}

fn checkOwnedChildren(xml: []const u8, list: []const Element) ScanError!void {
    for (list) |el| {
        if (std.mem.indexOfScalar(u8, el.name, ':') != null) {
            if (isOwnedLocal(el.localName())) return error.Malformed;
            if (std.mem.eql(u8, el.localName(), "AlternateContent") and !el.self_closing and
                bodyMentionsOwned(xml[el.open_end..el.close_lt])) return error.Malformed;
        }
    }
}

// ─── Edits ───────────────────────────────────────────────────────────

/// `[start, end)` of the part replaced by `text`; an insertion when
/// `start == end`. `seq` orders insertions at one position.
const Edit = struct {
    start: usize,
    end: usize,
    text: []const u8,
    seq: usize,

    fn before(_: void, x: Edit, y: Edit) bool {
        if (x.start != y.start) return x.start < y.start;
        return x.seq < y.seq;
    }
};

const Edits = struct {
    arena: std.heap.ArenaAllocator,
    list: std.ArrayListUnmanaged(Edit) = .empty,

    fn init(a: Allocator) Edits {
        return .{ .arena = std.heap.ArenaAllocator.init(a) };
    }

    fn deinit(self: *Edits) void {
        self.arena.deinit();
    }

    fn scratch(self: *Edits) Allocator {
        return self.arena.allocator();
    }

    fn add(self: *Edits, start: usize, end: usize, text: []const u8) Allocator.Error!void {
        assert(start <= end);
        try self.list.append(self.scratch(), .{ .start = start, .end = end, .text = text, .seq = self.list.items.len });
    }

    /// The part with every edit applied, in position order. The edits
    /// never overlap: each touches its own element or slot.
    fn apply(self: *Edits, a: Allocator, xml: []const u8) Allocator.Error![]u8 {
        std.mem.sort(Edit, self.list.items, {}, Edit.before);
        var out: std.ArrayListUnmanaged(u8) = .empty;
        errdefer out.deinit(a);
        var extra: usize = 0;
        for (self.list.items) |e| extra += e.text.len;
        try out.ensureTotalCapacity(a, xml.len + extra);
        var cursor: usize = 0;
        for (self.list.items) |e| {
            assert(e.start >= cursor);
            try out.appendSlice(a, xml[cursor..e.start]);
            try out.appendSlice(a, e.text);
            cursor = e.end;
        }
        try out.appendSlice(a, xml[cursor..]);
        return try out.toOwnedSlice(a);
    }
};

// ─── Attribute rewriting ─────────────────────────────────────────────

const AttrSub = struct { name: []const u8, value: []const u8 };

/// `<name` + the attributes of `attrs` with every `subs` entry
/// substituted (added at the end when absent) + `tail` (`/>` or `>`).
/// An attribute that is not `name="…"` / `name='…'` ends the walk:
/// what follows is copied verbatim.
fn writeTagWithAttrs(a: Allocator, out: *std.ArrayListUnmanaged(u8), name: []const u8, attrs: []const u8, subs: []const AttrSub, tail: []const u8) Allocator.Error!void {
    try out.append(a, '<');
    try out.appendSlice(a, name);
    var written = [_]bool{false} ** 8;
    assert(subs.len <= written.len);
    var i: usize = 0;
    while (i < attrs.len) {
        if (std.ascii.isWhitespace(attrs[i])) {
            i += 1;
            continue;
        }
        const name_start = i;
        while (i < attrs.len and !std.ascii.isWhitespace(attrs[i]) and attrs[i] != '=') i += 1;
        const attr_name = attrs[name_start..i];
        var j = i;
        while (j < attrs.len and std.ascii.isWhitespace(attrs[j])) j += 1;
        if (j >= attrs.len or attrs[j] != '=') {
            try out.append(a, ' ');
            try out.appendSlice(a, attrs[name_start..]);
            break;
        }
        j += 1;
        while (j < attrs.len and std.ascii.isWhitespace(attrs[j])) j += 1;
        if (j >= attrs.len or (attrs[j] != '"' and attrs[j] != '\'')) {
            try out.append(a, ' ');
            try out.appendSlice(a, attrs[name_start..]);
            break;
        }
        const q = attrs[j];
        const value_start = j + 1;
        const value_end = std.mem.indexOfScalarPos(u8, attrs, value_start, q) orelse {
            try out.append(a, ' ');
            try out.appendSlice(a, attrs[name_start..]);
            break;
        };
        i = value_end + 1;
        var replaced = false;
        for (subs, 0..) |s, k| {
            if (std.mem.eql(u8, s.name, attr_name)) {
                try out.append(a, ' ');
                try out.appendSlice(a, s.name);
                try out.appendSlice(a, "=\"");
                try out.appendSlice(a, s.value);
                try out.append(a, '"');
                written[k] = true;
                replaced = true;
                break;
            }
        }
        if (!replaced) {
            try out.append(a, ' ');
            try out.appendSlice(a, attrs[name_start..i]);
        }
    }
    for (subs, 0..) |s, k| {
        if (written[k]) continue;
        try out.append(a, ' ');
        try out.appendSlice(a, s.name);
        try out.appendSlice(a, "=\"");
        try out.appendSlice(a, s.value);
        try out.append(a, '"');
    }
    try out.appendSlice(a, tail);
}

/// The open tag of a self-closed element, re-spelled as an opener.
fn reopen(a: Allocator, out: *std.ArrayListUnmanaged(u8), xml: []const u8, el: Element) Allocator.Error!void {
    try out.append(a, '<');
    try out.appendSlice(a, el.name);
    try out.appendSlice(a, el.attrs(xml));
    try out.append(a, '>');
}

fn parseU32Attr(attrs: []const u8, name: []const u8) ?u32 {
    const raw = store_mod.xmlAttrValue(attrs, name) orelse return null;
    return std.fmt.parseInt(u32, std.mem.trim(u8, raw, " \t\r\n"), 10) catch null;
}

fn fmtU64(buf: []u8, v: u64) []const u8 {
    return std.fmt.bufPrint(buf, "{d}", .{v}) catch unreachable;
}

// ─── The sheet part ──────────────────────────────────────────────────

pub const SheetOptions = struct {
    /// `r:id="rId{n}"` of the first staged external hyperlink; the
    /// rest follow in order.
    hyperlink_rid_base: u32 = 1,
    /// Insert `<legacyDrawing r:id="rId{n}"/>` — set when comments are
    /// staged and the sheet holds no `<legacyDrawing>`.
    legacy_drawing_rid: ?u32 = null,
    /// The relationships namespace of the package's dialect, declared
    /// on each element that names an `r:id` when the root does not
    /// declare the `r` prefix.
    ns_r: []const u8,
};

/// What the splice needs to know about the part before the save —
/// read at the first registration (the caller judges a part the walk
/// cannot read then) and again at the splice.
pub const SheetFacts = struct {
    /// The `r:id` of the sheet's `<legacyDrawing>`, decoded; null when
    /// the sheet holds none. Owned by the caller's allocator.
    legacy_drawing_rid: ?[]u8,
    /// The highest `priority` a `<cfRule>` / `<x14:cfRule>` spells.
    cf_priority_max: u32,
    /// The root declares `xmlns:r`.
    root_declares_r: bool,

    pub fn deinit(self: *SheetFacts, a: Allocator) void {
        if (self.legacy_drawing_rid) |r| a.free(r);
        self.* = undefined;
    }
};

/// Read the part once: the root, its direct children (the owned-name
/// rule applied), the legacy drawing's id, the CF priority ceiling.
pub fn readSheet(a: Allocator, xml: []const u8) Error!SheetFacts {
    const root = rootElement(xml) catch return error.MalformedSheetXml;
    if (!std.mem.eql(u8, root.name, "worksheet")) return error.MalformedSheetXml;
    const list = children(a, xml, root.open_end, root.close_lt) catch |e| switch (e) {
        error.Malformed => return error.MalformedSheetXml,
        error.OutOfMemory => return error.OutOfMemory,
    };
    defer a.free(list);
    checkOwnedChildren(xml, list) catch return error.MalformedSheetXml;

    var rid: ?[]u8 = null;
    errdefer if (rid) |r| a.free(r);
    if (childNamed(list, "legacyDrawing")) |ld| {
        const raw = store_mod.xmlAttrValue(ld.attrs(xml), "r:id") orelse return error.MalformedSheetXml;
        rid = try store_mod.decodeXmlEntities(a, raw);
    }

    // The two verdicts the splice would otherwise give at the SAVE —
    // a `<col>` record without a readable `min` / `max` (or `min` 0 or
    // past `max`), a rule priority at the ceiling — are judged here, at
    // the first registration, so a sheet admitted can always be saved
    // (in-house r3 B-ORC-301).
    if (childNamed(list, "cols")) |cols| {
        if (!cols.self_closing) {
            const inner = children(a, xml, cols.open_end, cols.close_lt) catch |e| switch (e) {
                error.Malformed => return error.MalformedSheetXml,
                error.OutOfMemory => return error.OutOfMemory,
            };
            defer a.free(inner);
            for (inner) |el| {
                if (!std.mem.eql(u8, el.name, "col")) continue;
                const attrs = el.attrs(xml);
                const min = parseU32Attr(attrs, "min") orelse return error.MalformedSheetXml;
                const max = parseU32Attr(attrs, "max") orelse return error.MalformedSheetXml;
                if (min > max or min == 0) return error.MalformedSheetXml;
            }
        }
    }

    var max_priority: u32 = 0;
    inline for (.{ "cfRule", "x14:cfRule" }) |tag| {
        var cursor: usize = 0;
        while (wbxml.findTagOpen(xml, cursor, tag) catch return error.MalformedSheetXml) |hit| {
            if (parseU32Attr(xml[hit.attrs_start..hit.attrs_end], "priority")) |p| {
                if (p > max_priority) max_priority = p;
            }
            cursor = hit.after_tag_close;
        }
    }

    if (max_priority == std.math.maxInt(u32)) return error.MalformedSheetXml;

    return .{
        .legacy_drawing_rid = rid,
        .cf_priority_max = max_priority,
        .root_declares_r = store_mod.xmlAttrValue(root.attrs(xml), "xmlns:r") != null,
    };
}

/// The sheet part with every staged registration landed; the caller
/// has already extended the relationships, the comments part and the
/// VML drawing the sheet's new elements name.
pub fn spliceSheet(a: Allocator, xml: []const u8, state: *const sheet_plan.SheetState, facts: *const SheetFacts, opts: SheetOptions) Error![]u8 {
    var edits = Edits.init(a);
    defer edits.deinit();
    const s = edits.scratch();

    const root = rootElement(xml) catch return error.MalformedSheetXml;
    if (!std.mem.eql(u8, root.name, "worksheet")) return error.MalformedSheetXml;
    const list = children(s, xml, root.open_end, root.close_lt) catch |e| switch (e) {
        error.Malformed => return error.MalformedSheetXml,
        error.OutOfMemory => return error.OutOfMemory,
    };
    checkOwnedChildren(xml, list) catch return error.MalformedSheetXml;
    const ns_r: ?[]const u8 = if (facts.root_declares_r) null else opts.ns_r;

    if (state.freeze_rows != 0 or state.freeze_cols != 0) {
        try splicePane(&edits, xml, list, root, state.freeze_rows, state.freeze_cols);
    }
    if (state.column_widths.items.len > 0) {
        try spliceCols(&edits, xml, list, root, state.column_widths.items);
    }
    if (state.row_heights.count() > 0) {
        try spliceRowHeights(&edits, xml, list, &state.row_heights);
    }
    if (state.auto_filter_range) |range| {
        var frag: std.ArrayListUnmanaged(u8) = .empty;
        try sheet_plan.emitAutoFilterElement(s, &frag, range);
        if (childNamed(list, "autoFilter")) |el| {
            try edits.add(el.lt, el.end, frag.items);
        } else {
            const at = slotFor(list, "autoFilter", root.open_end);
            try edits.add(at, at, frag.items);
        }
    }
    if (state.merged_cells.items.len > 0) {
        var frag: std.ArrayListUnmanaged(u8) = .empty;
        try sheet_plan.emitMergeCellElements(s, &frag, @ptrCast(state.merged_cells.items));
        try extendCountedTable(&edits, xml, list, root, "mergeCells", "mergeCell", state.merged_cells.items.len, frag.items);
    }
    if (state.conditional_formats.items.len > 0) {
        var frag: std.ArrayListUnmanaged(u8) = .empty;
        const view = try s.alloc(sheet_plan.ConditionalFormat, state.conditional_formats.items.len);
        for (state.conditional_formats.items, 0..) |cf, k| {
            view[k] = .{
                .range = cf.range,
                .rule = switch (cf.rule) {
                    .cell_is => |r| .{ .cell_is = .{ .operator = r.operator, .formula1 = r.formula1, .formula2 = r.formula2, .dxf_id = r.dxf_id } },
                    .expression => |r| .{ .expression = .{ .formula = r.formula, .dxf_id = r.dxf_id } },
                    .color_scale => |r| .{ .color_scale = .{ .low_color_argb = r.low_color_argb, .mid_color_argb = r.mid_color_argb, .high_color_argb = r.high_color_argb } },
                    .data_bar => |r| .{ .data_bar = .{ .color_argb = r.color_argb } },
                },
            };
        }
        // A part spelling a priority at the ceiling leaves no room.
        if (facts.cf_priority_max > std.math.maxInt(u32) - state.conditional_formats.items.len) return error.MalformedSheetXml;
        try sheet_plan.emitConditionalFormattingBlocks(s, &frag, view, facts.cf_priority_max);
        var at = slotFor(list, "conditionalFormatting", root.open_end);
        for (list) |el| if (std.mem.eql(u8, el.name, "conditionalFormatting")) {
            at = el.end;
        };
        try edits.add(at, at, frag.items);
    }
    if (state.data_validations.items.len + state.data_validation_ranges.items.len > 0) {
        var frag: std.ArrayListUnmanaged(u8) = .empty;
        const dvl = try s.alloc(sheet_plan.DataValidationList, state.data_validations.items.len);
        for (state.data_validations.items, 0..) |dv, k| dvl[k] = .{ .range = dv.range, .values = @ptrCast(dv.values) };
        const dvr = try s.alloc(sheet_plan.DataValidationRange, state.data_validation_ranges.items.len);
        for (state.data_validation_ranges.items, 0..) |dv, k| dvr[k] = .{ .range = dv.range, .kind_name = dv.kind_name, .op_name = dv.op_name, .formula1 = dv.formula1, .formula2 = dv.formula2 };
        try sheet_plan.emitDataValidationElements(s, &frag, dvl, dvr);
        try extendCountedTable(&edits, xml, list, root, "dataValidations", "dataValidation", dvl.len + dvr.len, frag.items);
    }
    if (state.hyperlinks.items.len + state.internal_hyperlinks.items.len > 0) {
        var frag: std.ArrayListUnmanaged(u8) = .empty;
        const ext = try s.alloc(sheet_plan.Hyperlink, state.hyperlinks.items.len);
        for (state.hyperlinks.items, 0..) |h, k| ext[k] = .{ .range = h.range, .url = h.url };
        const int = try s.alloc(sheet_plan.InternalHyperlink, state.internal_hyperlinks.items.len);
        for (state.internal_hyperlinks.items, 0..) |h, k| int[k] = .{ .range = h.range, .location = h.location };
        try sheet_plan.emitHyperlinkElements(s, &frag, ext, int, opts.hyperlink_rid_base, ns_r);
        try extendCountedTable(&edits, xml, list, root, "hyperlinks", null, 0, frag.items);
    }
    if (opts.legacy_drawing_rid) |rid| {
        assert(childNamed(list, "legacyDrawing") == null);
        var frag: std.ArrayListUnmanaged(u8) = .empty;
        try sheet_plan.emitLegacyDrawingElement(s, &frag, rid, ns_r);
        const at = slotFor(list, "legacyDrawing", root.open_end);
        try edits.add(at, at, frag.items);
    }
    return try edits.apply(a, xml);
}

/// `<pane>` into the first `<sheetView>`: replaced where one stands,
/// inserted first among its children where none, the view and the
/// views created where the sheet has none. A `<selection>` naming a
/// pane the new split lacks is left as the producer wrote it.
fn splicePane(edits: *Edits, xml: []const u8, list: []const Element, root: Element, rows: u32, cols: u32) Error!void {
    const s = edits.scratch();
    var pane: std.ArrayListUnmanaged(u8) = .empty;
    try sheet_plan.emitPaneElement(s, &pane, rows, cols);
    const views = childNamed(list, "sheetViews") orelse {
        var frag: std.ArrayListUnmanaged(u8) = .empty;
        try frag.appendSlice(s, "<sheetViews><sheetView workbookViewId=\"0\">");
        try frag.appendSlice(s, pane.items);
        try frag.appendSlice(s, "</sheetView></sheetViews>");
        const at = slotFor(list, "sheetViews", root.open_end);
        try edits.add(at, at, frag.items);
        return;
    };
    if (views.self_closing) {
        var frag: std.ArrayListUnmanaged(u8) = .empty;
        try reopen(s, &frag, xml, views);
        try frag.appendSlice(s, "<sheetView workbookViewId=\"0\">");
        try frag.appendSlice(s, pane.items);
        try frag.appendSlice(s, "</sheetView></sheetViews>");
        try edits.add(views.lt, views.end, frag.items);
        return;
    }
    const inner = children(s, xml, views.open_end, views.close_lt) catch |e| switch (e) {
        error.Malformed => return error.MalformedSheetXml,
        error.OutOfMemory => return error.OutOfMemory,
    };
    const view = childNamed(inner, "sheetView") orelse {
        var frag: std.ArrayListUnmanaged(u8) = .empty;
        try frag.appendSlice(s, "<sheetView workbookViewId=\"0\">");
        try frag.appendSlice(s, pane.items);
        try frag.appendSlice(s, "</sheetView>");
        try edits.add(views.open_end, views.open_end, frag.items);
        return;
    };
    if (view.self_closing) {
        var frag: std.ArrayListUnmanaged(u8) = .empty;
        try reopen(s, &frag, xml, view);
        try frag.appendSlice(s, pane.items);
        try frag.appendSlice(s, "</sheetView>");
        try edits.add(view.lt, view.end, frag.items);
        return;
    }
    const view_children = children(s, xml, view.open_end, view.close_lt) catch |e| switch (e) {
        error.Malformed => return error.MalformedSheetXml,
        error.OutOfMemory => return error.OutOfMemory,
    };
    if (childNamed(view_children, "pane")) |old| {
        try edits.add(old.lt, old.end, pane.items);
    } else {
        try edits.add(view.open_end, view.open_end, pane.items);
    }
}

/// A closed column interval with its width.
const ColSpan = struct { min: u32, max: u32, width: f32 };

/// The staged widths as disjoint intervals, the last registration
/// winning where two overlap.
fn stagedColSpans(s: Allocator, widths: []const sheet_plan.ColumnWidth) Allocator.Error![]ColSpan {
    var spans: std.ArrayListUnmanaged(ColSpan) = .empty;
    for (widths) |cw| {
        // The spans outside the new one survive, split where it cuts
        // them; rebuilt whole rather than edited in place (in-house r1
        // A-SPL-102: a removal followed by `i += 1` skipped the span
        // that slid into the slot).
        var kept: std.ArrayListUnmanaged(ColSpan) = .empty;
        for (spans.items) |old| {
            if (old.max < cw.col_min or old.min > cw.col_max) {
                try kept.append(s, old);
                continue;
            }
            if (old.min < cw.col_min) try kept.append(s, .{ .min = old.min, .max = cw.col_min - 1, .width = old.width });
            if (old.max > cw.col_max) try kept.append(s, .{ .min = cw.col_max + 1, .max = old.max, .width = old.width });
        }
        try kept.append(s, .{ .min = cw.col_min, .max = cw.col_max, .width = cw.width });
        spans.deinit(s);
        spans = kept;
    }
    return try spans.toOwnedSlice(s);
}

/// One `<col>` of the rewritten block: an existing record's bytes
/// (`raw`), or a piece of one with its bounds moved (`attrs` + the
/// substitutions), or a staged span (`attrs == null`).
const ColPiece = struct {
    min: u32,
    max: u32,
    width: ?f32,
    raw: ?[]const u8,
    attrs: ?[]const u8,

    fn before(_: void, x: ColPiece, y: ColPiece) bool {
        return x.min < y.min;
    }
};

/// `<cols>`: every existing `<col>` covering a staged column is split
/// around it, the covered piece keeping the record's other attributes
/// (`style`, `hidden`, …) with `width` and `customWidth` replaced; a
/// staged column no record covers is written as the fresh emitter
/// writes it; the block is re-emitted sorted by `min`.
fn spliceCols(edits: *Edits, xml: []const u8, list: []const Element, root: Element, widths: []const sheet_plan.ColumnWidth) Error!void {
    const s = edits.scratch();
    const spans = try stagedColSpans(s, widths);
    var pieces: std.ArrayListUnmanaged(ColPiece) = .empty;
    const cols = childNamed(list, "cols");
    var prefix: std.ArrayListUnmanaged(u8) = .empty;
    if (cols) |c| {
        if (!c.self_closing) {
            const inner = children(s, xml, c.open_end, c.close_lt) catch |e| switch (e) {
                error.Malformed => return error.MalformedSheetXml,
                error.OutOfMemory => return error.OutOfMemory,
            };
            for (inner) |el| {
                if (!std.mem.eql(u8, el.name, "col")) {
                    // Not a column record: kept ahead of the block, as found.
                    try prefix.appendSlice(s, xml[el.lt..el.end]);
                    continue;
                }
                const attrs = el.attrs(xml);
                const min = parseU32Attr(attrs, "min") orelse return error.MalformedSheetXml;
                const max = parseU32Attr(attrs, "max") orelse return error.MalformedSheetXml;
                if (min > max or min == 0) return error.MalformedSheetXml;
                // The record's columns not under a staged span keep its bytes.
                var cursor = min;
                var covered = false;
                for (spans) |sp| {
                    if (sp.max < min or sp.min > max) continue;
                    covered = true;
                }
                if (!covered) {
                    try pieces.append(s, .{ .min = min, .max = max, .width = null, .raw = xml[el.lt..el.end], .attrs = null });
                    continue;
                }
                // Walk the record column by column in span order.
                while (cursor <= max) {
                    var next_span: ?ColSpan = null;
                    for (spans) |sp| {
                        if (sp.max < cursor or sp.min > max) continue;
                        if (next_span == null or sp.min < next_span.?.min) next_span = sp;
                    }
                    const sp = next_span orelse {
                        try pieces.append(s, .{ .min = cursor, .max = max, .width = null, .raw = null, .attrs = attrs });
                        break;
                    };
                    const sp_min = @max(sp.min, cursor);
                    if (sp_min > cursor) {
                        try pieces.append(s, .{ .min = cursor, .max = sp_min - 1, .width = null, .raw = null, .attrs = attrs });
                    }
                    const sp_max = @min(sp.max, max);
                    try pieces.append(s, .{ .min = sp_min, .max = sp_max, .width = sp.width, .raw = null, .attrs = attrs });
                    if (sp_max == max) break;
                    cursor = sp_max + 1;
                }
            }
        }
    }
    // The staged columns no record covered.
    for (spans) |sp| {
        var cursor = sp.min;
        while (cursor <= sp.max) {
            // The first piece (from the records) at or after `cursor`
            // inside the span.
            var stop: u32 = sp.max + 1;
            var covered_here = false;
            for (pieces.items) |p| {
                if (p.attrs == null and p.raw == null) continue;
                if (p.max < cursor or p.min > sp.max) continue;
                if (p.min <= cursor) {
                    covered_here = true;
                    if (p.max + 1 > cursor) cursor = p.max + 1;
                    break;
                }
                if (p.min < stop) stop = p.min;
            }
            if (covered_here) continue;
            try pieces.append(s, .{ .min = cursor, .max = stop - 1, .width = sp.width, .raw = null, .attrs = null });
            cursor = stop;
        }
    }
    std.mem.sort(ColPiece, pieces.items, {}, ColPiece.before);

    var frag: std.ArrayListUnmanaged(u8) = .empty;
    if (cols) |c| {
        if (c.self_closing) try reopen(s, &frag, xml, c) else try frag.appendSlice(s, xml[c.lt..c.open_end]);
    } else {
        try frag.appendSlice(s, "<cols>");
    }
    try frag.appendSlice(s, prefix.items);
    for (pieces.items) |p| {
        if (p.raw) |raw| {
            try frag.appendSlice(s, raw);
            continue;
        }
        var min_buf: [16]u8 = undefined;
        var max_buf: [16]u8 = undefined;
        var w_buf: [64]u8 = undefined;
        if (p.attrs) |attrs| {
            if (p.width) |w| {
                const subs = [_]AttrSub{
                    .{ .name = "min", .value = fmtU64(&min_buf, p.min) },
                    .{ .name = "max", .value = fmtU64(&max_buf, p.max) },
                    .{ .name = "width", .value = std.fmt.bufPrint(&w_buf, "{d}", .{w}) catch unreachable },
                    .{ .name = "customWidth", .value = "1" },
                };
                try writeTagWithAttrs(s, &frag, "col", attrs, &subs, "/>");
            } else {
                const subs = [_]AttrSub{
                    .{ .name = "min", .value = fmtU64(&min_buf, p.min) },
                    .{ .name = "max", .value = fmtU64(&max_buf, p.max) },
                };
                try writeTagWithAttrs(s, &frag, "col", attrs, &subs, "/>");
            }
        } else {
            try sheet_plan.emitColElements(s, &frag, &.{.{ .col_min = p.min, .col_max = p.max, .width = p.width.? }});
        }
    }
    try frag.appendSlice(s, "</cols>");
    if (cols) |c| {
        try edits.add(c.lt, c.end, frag.items);
    } else {
        const at = slotFor(list, "cols", root.open_end);
        try edits.add(at, at, frag.items);
    }
}

const RowHeight = struct {
    row: u32,
    height: f32,
    fn before(_: void, x: RowHeight, y: RowHeight) bool {
        return x.row < y.row;
    }
};

/// Row heights: a `<row>` the sheet holds gets `ht` and `customHeight`
/// on its open tag; one it lacks is created empty in row order. A
/// `<row>` without a readable `r` is neither matched nor an anchor.
fn spliceRowHeights(edits: *Edits, xml: []const u8, list: []const Element, heights: *const std.AutoHashMapUnmanaged(u32, f32)) Error!void {
    const s = edits.scratch();
    const data = childNamed(list, "sheetData") orelse return error.MalformedSheetXml;
    var staged = try s.alloc(RowHeight, heights.count());
    var it = heights.iterator();
    var n: usize = 0;
    while (it.next()) |kv| : (n += 1) staged[n] = .{ .row = kv.key_ptr.* + 1, .height = kv.value_ptr.* };
    std.mem.sort(RowHeight, staged, {}, RowHeight.before);

    if (data.self_closing) {
        var frag: std.ArrayListUnmanaged(u8) = .empty;
        try reopen(s, &frag, xml, data);
        for (staged) |rh| try frag.print(s, "<row r=\"{d}\" ht=\"{d}\" customHeight=\"1\"/>", .{ rh.row, rh.height });
        try frag.appendSlice(s, "</sheetData>");
        try edits.add(data.lt, data.end, frag.items);
        return;
    }
    const rows = children(s, xml, data.open_end, data.close_lt) catch |e| switch (e) {
        error.Malformed => return error.MalformedSheetXml,
        error.OutOfMemory => return error.OutOfMemory,
    };
    var ri: usize = 0;
    for (staged) |rh| {
        // Advance to the first row at or past the target.
        var insert_at: usize = data.close_lt;
        var matched: ?Element = null;
        while (ri < rows.len) : (ri += 1) {
            const el = rows[ri];
            if (!std.mem.eql(u8, el.name, "row")) continue;
            const r = parseU32Attr(el.attrs(xml), "r") orelse continue;
            if (r < rh.row) continue;
            if (r == rh.row) matched = el;
            insert_at = el.lt;
            break;
        }
        var h_buf: [64]u8 = undefined;
        if (matched) |el| {
            var frag: std.ArrayListUnmanaged(u8) = .empty;
            const subs = [_]AttrSub{
                .{ .name = "ht", .value = std.fmt.bufPrint(&h_buf, "{d}", .{rh.height}) catch unreachable },
                .{ .name = "customHeight", .value = "1" },
            };
            try writeTagWithAttrs(s, &frag, "row", el.attrs(xml), &subs, if (el.self_closing) "/>" else ">");
            try edits.add(el.lt, el.open_end, frag.items);
            ri += 1;
        } else {
            var frag: std.ArrayListUnmanaged(u8) = .empty;
            try frag.print(s, "<row r=\"{d}\" ht=\"{d}\" customHeight=\"1\"/>", .{ rh.row, rh.height });
            try edits.add(insert_at, insert_at, frag.items);
        }
    }
}

/// A table of records (`<mergeCells>`, `<dataValidations>`,
/// `<hyperlinks>`): `fragment` appended after the records it holds, a
/// `count` rewritten where the open tag spells one (the records
/// counted, never the attribute trusted), a self-closed table opened,
/// an absent one created at its slot with the count when the schema
/// gives it one (`record` non-null).
fn extendCountedTable(edits: *Edits, xml: []const u8, list: []const Element, root: Element, table: []const u8, record: ?[]const u8, added: usize, fragment: []const u8) Error!void {
    const s = edits.scratch();
    var frag: std.ArrayListUnmanaged(u8) = .empty;
    var n_buf: [24]u8 = undefined;
    const el = childNamed(list, table) orelse {
        try frag.append(s, '<');
        try frag.appendSlice(s, table);
        if (record != null) {
            try frag.appendSlice(s, " count=\"");
            try frag.appendSlice(s, fmtU64(&n_buf, added));
            try frag.append(s, '"');
        }
        try frag.append(s, '>');
        try frag.appendSlice(s, fragment);
        try frag.appendSlice(s, "</");
        try frag.appendSlice(s, table);
        try frag.append(s, '>');
        const at = slotFor(list, table, root.open_end);
        try edits.add(at, at, frag.items);
        return;
    };
    const attrs = el.attrs(xml);
    const existing: usize = if (el.self_closing or record == null) 0 else blk: {
        const inner = children(s, xml, el.open_end, el.close_lt) catch |e| switch (e) {
            error.Malformed => return error.MalformedSheetXml,
            error.OutOfMemory => return error.OutOfMemory,
        };
        break :blk countNamed(inner, record.?);
    };
    const spells_count = record != null and store_mod.xmlAttrValue(attrs, "count") != null;
    if (spells_count) {
        const subs = [_]AttrSub{.{ .name = "count", .value = fmtU64(&n_buf, existing + added) }};
        try writeTagWithAttrs(s, &frag, table, attrs, &subs, ">");
    } else if (el.self_closing) {
        try reopen(s, &frag, xml, el);
    }
    if (el.self_closing) {
        try frag.appendSlice(s, fragment);
        try frag.appendSlice(s, "</");
        try frag.appendSlice(s, table);
        try frag.append(s, '>');
        try edits.add(el.lt, el.end, frag.items);
        return;
    }
    if (spells_count) try edits.add(el.lt, el.open_end, frag.items);
    try edits.add(el.close_lt, el.close_lt, fragment);
}

// ─── The sheet's relationships ───────────────────────────────────────

pub const RelEntry = struct {
    id: u32,
    type_uri: []const u8,
    target: []const u8,
    external: bool,
};

const rels_head =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";

fn writeRelationship(a: Allocator, out: *std.ArrayListUnmanaged(u8), e: RelEntry) Error!void {
    try out.print(a, "<Relationship Id=\"rId{d}\" Type=\"", .{e.id});
    try sheet_plan.appendXmlEscaped(a, out, e.type_uri);
    try out.appendSlice(a, "\" Target=\"");
    try sheet_plan.appendXmlEscaped(a, out, e.target);
    try out.append(a, '"');
    if (e.external) try out.appendSlice(a, " TargetMode=\"External\"");
    try out.appendSlice(a, "/>");
}

/// The next free `rId{n}` of a rels part, 1 when it holds none.
pub fn nextFreeRelId(rels_xml: []const u8) Error!u32 {
    var max_id: u32 = 0;
    var cursor: usize = 0;
    while (wbxml.findTagOpen(rels_xml, cursor, "Relationship") catch return error.MalformedSheetRels) |hit| {
        if (wbxml.getAttr(rels_xml[hit.attrs_start..hit.attrs_end], "Id")) |id| {
            if (std.mem.startsWith(u8, id, "rId")) {
                if (std.fmt.parseInt(u32, id["rId".len..], 10)) |n| {
                    if (n > max_id) max_id = n;
                } else |_| {}
            }
        }
        cursor = hit.after_tag_close;
    }
    if (max_id == std.math.maxInt(u32)) return error.IdSpaceExhausted;
    return max_id + 1;
}

/// `entries` appended before the root's close tag; a fresh part when
/// `rels_xml` is null.
pub fn appendRelationships(a: Allocator, rels_xml: ?[]const u8, entries: []const RelEntry) Error![]u8 {
    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(a);
    if (rels_xml) |xml| {
        // A self-closed root (`<Relationships/>`) is opened — the
        // counted tables' rule (in-house r1 B-ORC-103).
        const root = (nextElement(xml, 0, xml.len) catch return error.MalformedSheetRels) orelse return error.MalformedSheetRels;
        if (!std.mem.eql(u8, root.name, "Relationships")) return error.MalformedSheetRels;
        if (root.self_closing) {
            try out.appendSlice(a, xml[0..root.lt]);
            try reopen(a, &out, xml, root);
            for (entries) |e| try writeRelationship(a, &out, e);
            try out.appendSlice(a, "</Relationships>");
            try out.appendSlice(a, xml[root.end..]);
        } else {
            try out.appendSlice(a, xml[0..root.close_lt]);
            for (entries) |e| try writeRelationship(a, &out, e);
            try out.appendSlice(a, xml[root.close_lt..]);
        }
    } else {
        try out.appendSlice(a, rels_head);
        for (entries) |e| try writeRelationship(a, &out, e);
        try out.appendSlice(a, "</Relationships>");
    }
    return try out.toOwnedSlice(a);
}

// ─── The comments part ───────────────────────────────────────────────

/// The `ref` of every `<comment>` the part holds, decoded, in document
/// order; each string and the slice are the caller's.
pub fn commentRefs(a: Allocator, comments_xml: []const u8) Error![][]u8 {
    var refs: std.ArrayListUnmanaged([]u8) = .empty;
    errdefer {
        for (refs.items) |r| a.free(r);
        refs.deinit(a);
    }
    const root = rootElement(comments_xml) catch return error.MalformedCommentsXml;
    if (!std.mem.eql(u8, root.name, "comments")) return error.MalformedCommentsXml;
    const list = children(a, comments_xml, root.open_end, root.close_lt) catch |e| switch (e) {
        error.Malformed => return error.MalformedCommentsXml,
        error.OutOfMemory => return error.OutOfMemory,
    };
    defer a.free(list);
    const cl = childNamed(list, "commentList") orelse return try refs.toOwnedSlice(a);
    if (cl.self_closing) return try refs.toOwnedSlice(a);
    const records = children(a, comments_xml, cl.open_end, cl.close_lt) catch |e| switch (e) {
        error.Malformed => return error.MalformedCommentsXml,
        error.OutOfMemory => return error.OutOfMemory,
    };
    defer a.free(records);
    for (records) |el| {
        if (!std.mem.eql(u8, el.name, "comment")) continue;
        const raw = store_mod.xmlAttrValue(el.attrs(comments_xml), "ref") orelse return error.MalformedCommentsXml;
        const decoded = try store_mod.decodeXmlEntities(a, raw);
        errdefer a.free(decoded);
        try refs.append(a, decoded);
    }
    return try refs.toOwnedSlice(a);
}

/// The comments part with `comments` appended: an author the part
/// already names (its `<author>` text, decoded) keeps its id, a new
/// one is appended to `<authors>`; `<authors>` / `<commentList>` are
/// created where the part lacks them, a self-closed list opened.
pub fn extendComments(a: Allocator, comments_xml: []const u8, comments: []const sheet_plan.Comment) Error![]u8 {
    var edits = Edits.init(a);
    defer edits.deinit();
    const s = edits.scratch();
    const root = rootElement(comments_xml) catch return error.MalformedCommentsXml;
    if (!std.mem.eql(u8, root.name, "comments")) return error.MalformedCommentsXml;
    const list = children(s, comments_xml, root.open_end, root.close_lt) catch |e| switch (e) {
        error.Malformed => return error.MalformedCommentsXml,
        error.OutOfMemory => return error.OutOfMemory,
    };

    // The authors the part names.
    var authors: std.ArrayListUnmanaged([]const u8) = .empty;
    const authors_el = childNamed(list, "authors");
    if (authors_el) |ae| {
        if (!ae.self_closing) {
            const records = children(s, comments_xml, ae.open_end, ae.close_lt) catch |e| switch (e) {
                error.Malformed => return error.MalformedCommentsXml,
                error.OutOfMemory => return error.OutOfMemory,
            };
            for (records) |el| {
                if (!std.mem.eql(u8, el.name, "author")) continue;
                const body = if (el.self_closing) "" else comments_xml[el.open_end..el.close_lt];
                try authors.append(s, try store_mod.decodeXmlEntities(s, body));
            }
        }
    }
    const existing_authors = authors.items.len;
    var new_authors: std.ArrayListUnmanaged(u8) = .empty;
    var records: std.ArrayListUnmanaged(u8) = .empty;
    for (comments) |c| {
        var id: ?usize = null;
        for (authors.items, 0..) |name, k| {
            if (std.mem.eql(u8, name, c.author)) {
                id = k;
                break;
            }
        }
        if (id == null) {
            id = authors.items.len;
            try authors.append(s, c.author);
            try sheet_plan.emitAuthorElement(s, &new_authors, c.author);
        }
        try sheet_plan.emitCommentElement(s, &records, c.ref, id.?, c.text);
    }

    if (authors_el) |ae| {
        if (authors.items.len > existing_authors) {
            if (ae.self_closing) {
                var frag: std.ArrayListUnmanaged(u8) = .empty;
                try reopen(s, &frag, comments_xml, ae);
                try frag.appendSlice(s, new_authors.items);
                try frag.appendSlice(s, "</authors>");
                try edits.add(ae.lt, ae.end, frag.items);
            } else {
                try edits.add(ae.close_lt, ae.close_lt, new_authors.items);
            }
        }
    } else {
        var frag: std.ArrayListUnmanaged(u8) = .empty;
        try frag.appendSlice(s, "<authors>");
        try frag.appendSlice(s, new_authors.items);
        try frag.appendSlice(s, "</authors>");
        try edits.add(root.open_end, root.open_end, frag.items);
    }
    if (childNamed(list, "commentList")) |cl| {
        if (cl.self_closing) {
            var frag: std.ArrayListUnmanaged(u8) = .empty;
            try reopen(s, &frag, comments_xml, cl);
            try frag.appendSlice(s, records.items);
            try frag.appendSlice(s, "</commentList>");
            try edits.add(cl.lt, cl.end, frag.items);
        } else {
            try edits.add(cl.close_lt, cl.close_lt, records.items);
        }
    } else {
        var frag: std.ArrayListUnmanaged(u8) = .empty;
        try frag.appendSlice(s, "<commentList>");
        try frag.appendSlice(s, records.items);
        try frag.appendSlice(s, "</commentList>");
        const at = if (authors_el) |ae| ae.end else root.open_end;
        try edits.add(at, at, frag.items);
    }
    return try edits.apply(a, comments_xml);
}

// ─── The VML drawing ─────────────────────────────────────────────────

fn isVmlName(el: Element, local: []const u8) bool {
    return std.mem.eql(u8, el.localName(), local);
}

/// The VML drawing with one note shape per comment appended before
/// the root's close: shape ids past the highest `_x0000_s{n}` the
/// part spells (1025 upward when none), the note shape type added
/// when the part lacks it, the `<o:idmap>` extended to the blocks the
/// new ids fall in.
pub fn extendVml(a: Allocator, vml_xml: []const u8, comments: []const sheet_plan.Comment) Error![]u8 {
    var edits = Edits.init(a);
    defer edits.deinit();
    const s = edits.scratch();
    const root = rootElement(vml_xml) catch return error.MalformedVmlDrawing;
    const list = children(s, vml_xml, root.open_end, root.close_lt) catch |e| switch (e) {
        error.Malformed => return error.MalformedVmlDrawing,
        error.OutOfMemory => return error.OutOfMemory,
    };
    var max_id: u64 = 1024;
    var shapes: u64 = 0;
    var has_shapetype = false;
    var first_shape_lt: ?usize = null;
    var layout: ?Element = null;
    // Shape ids at EVERY depth — a shape inside a `v:group` holds an
    // id of the same space (in-house r1 B-VML-104).
    inline for (.{ "v:shape", "shape" }) |tag| {
        var cursor: usize = 0;
        while (wbxml.findTagOpen(vml_xml, cursor, tag) catch return error.MalformedVmlDrawing) |hit| {
            if (wbxml.getAttr(vml_xml[hit.attrs_start..hit.attrs_end], "id")) |id| {
                if (std.mem.startsWith(u8, id, "_x0000_s")) {
                    if (std.fmt.parseInt(u64, id["_x0000_s".len..], 10)) |n| {
                        if (n > max_id) max_id = n;
                    } else |_| {}
                }
            }
            cursor = hit.after_tag_close;
        }
    }
    for (list) |el| {
        if (isVmlName(el, "shape")) {
            shapes += 1;
            if (first_shape_lt == null) first_shape_lt = el.lt;
        } else if (isVmlName(el, "shapetype")) {
            if (store_mod.xmlAttrValue(el.attrs(vml_xml), "id")) |id| {
                if (std.mem.eql(u8, id, "_x0000_t202")) has_shapetype = true;
            }
        } else if (isVmlName(el, "shapelayout")) {
            layout = el;
        }
    }
    if (max_id > std.math.maxInt(u64) - comments.len) return error.MalformedVmlDrawing;

    // The id map: every 1024-block a new id falls in must be listed.
    const first_new = max_id + 1;
    const last_new = max_id + comments.len;
    const first_block = first_new / 1024;
    const last_block = last_new / 1024;
    var data: std.ArrayListUnmanaged(u8) = .empty;
    var idmap: ?Element = null;
    if (layout) |lo| {
        if (!lo.self_closing) {
            const inner = children(s, vml_xml, lo.open_end, lo.close_lt) catch |e| switch (e) {
                error.Malformed => return error.MalformedVmlDrawing,
                error.OutOfMemory => return error.OutOfMemory,
            };
            for (inner) |el| if (isVmlName(el, "idmap")) {
                idmap = el;
                break;
            };
        }
    }
    if (idmap) |im| {
        if (store_mod.xmlAttrValue(im.attrs(vml_xml), "data")) |d| try data.appendSlice(s, d);
    }
    var block = first_block;
    while (block <= last_block) : (block += 1) {
        var listed = false;
        var it = std.mem.splitScalar(u8, data.items, ',');
        while (it.next()) |tok| {
            const trimmed = std.mem.trim(u8, tok, " \t\r\n");
            if (std.fmt.parseInt(u64, trimmed, 10)) |n| {
                if (n == block) listed = true;
            } else |_| {}
        }
        if (listed) continue;
        if (data.items.len > 0) try data.append(s, ',');
        var b_buf: [24]u8 = undefined;
        try data.appendSlice(s, fmtU64(&b_buf, block));
    }
    if (idmap) |im| {
        var frag: std.ArrayListUnmanaged(u8) = .empty;
        const subs = [_]AttrSub{.{ .name = "data", .value = data.items }};
        try writeTagWithAttrs(s, &frag, im.name, im.attrs(vml_xml), &subs, if (im.self_closing) "/>" else ">");
        try edits.add(im.lt, im.open_end, frag.items);
    } else if (layout) |lo| {
        var frag: std.ArrayListUnmanaged(u8) = .empty;
        try frag.appendSlice(s, "<o:idmap v:ext=\"edit\" data=\"");
        try frag.appendSlice(s, data.items);
        try frag.appendSlice(s, "\"/>");
        if (lo.self_closing) {
            var whole: std.ArrayListUnmanaged(u8) = .empty;
            try reopen(s, &whole, vml_xml, lo);
            try whole.appendSlice(s, frag.items);
            try whole.appendSlice(s, "</");
            try whole.appendSlice(s, lo.name);
            try whole.append(s, '>');
            try edits.add(lo.lt, lo.end, whole.items);
        } else {
            try edits.add(lo.open_end, lo.open_end, frag.items);
        }
    } else {
        var frag: std.ArrayListUnmanaged(u8) = .empty;
        try frag.appendSlice(s, "<o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"");
        try frag.appendSlice(s, data.items);
        try frag.appendSlice(s, "\"/></o:shapelayout>");
        try edits.add(root.open_end, root.open_end, frag.items);
    }
    if (!has_shapetype) {
        const at = first_shape_lt orelse root.close_lt;
        try edits.add(at, at, sheet_plan.VML_NOTE_SHAPETYPE);
    }
    var frag: std.ArrayListUnmanaged(u8) = .empty;
    for (comments, 0..) |c, k| {
        try sheet_plan.emitVmlNoteShape(s, &frag, c.ref, first_new + k, shapes + k + 1);
    }
    try edits.add(root.close_lt, root.close_lt, frag.items);
    return try edits.apply(a, vml_xml);
}

/// The VML drawing readable by the walk — `addComment`'s check on a
/// drawing the save would extend, judged before anything is staged.
pub fn checkVml(a: Allocator, vml_xml: []const u8) Error!void {
    const root = rootElement(vml_xml) catch return error.MalformedVmlDrawing;
    const list = children(a, vml_xml, root.open_end, root.close_lt) catch |e| switch (e) {
        error.Malformed => return error.MalformedVmlDrawing,
        error.OutOfMemory => return error.OutOfMemory,
    };
    a.free(list);
}

// ─── Tests ───────────────────────────────────────────────────────────

const t = std.testing;

fn testState(a: Allocator) sheet_plan.SheetState {
    _ = a;
    return .{};
}

const ws_ns = "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"";
const ns_r_uri = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn splice(a: Allocator, xml: []const u8, state: *const sheet_plan.SheetState, opts: SheetOptions) Error![]u8 {
    var facts = try readSheet(a, xml);
    defer facts.deinit(a);
    return try spliceSheet(a, xml, state, &facts, opts);
}

test "S3d slice 2: an absent element lands at its schema slot, a present table is extended with its count rewritten, a self-closed one opened" {
    const a = t.allocator;
    var st: sheet_plan.SheetState = .{};
    defer st.deinit(a);
    try st.addMergedCell(a, "A1:B2");
    try st.addMergedCell(a, "C3:D4");
    try st.setAutoFilter(a, "A1:D1");
    try st.freezePanes(1, 0);
    try st.addDataValidationList(a, "E1", &.{ "x", "y" });
    try st.addInternalHyperlink(a, "F1", "Sheet1!A1");

    // Nothing but sheetData: every element is created, in schema order.
    {
        const src = "<worksheet " ++ ws_ns ++ "><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData><pageMargins left=\"0.7\"/></worksheet>";
        const out = try splice(a, src, &st, .{ .ns_r = ns_r_uri });
        defer a.free(out);
        try t.expectEqualStrings(
            "<worksheet " ++ ws_ns ++ "><sheetViews><sheetView workbookViewId=\"0\"><pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/></sheetView></sheetViews>" ++
                "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>" ++
                "<autoFilter ref=\"A1:D1\"/><mergeCells count=\"2\"><mergeCell ref=\"A1:B2\"/><mergeCell ref=\"C3:D4\"/></mergeCells>" ++
                "<dataValidations count=\"1\"><dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"E1\"><formula1>&quot;x,y&quot;</formula1></dataValidation></dataValidations>" ++
                "<hyperlinks><hyperlink ref=\"F1\" location=\"Sheet1!A1\"/></hyperlinks>" ++
                "<pageMargins left=\"0.7\"/></worksheet>",
            out,
        );
    }
    // Present tables: extended after their records, the count rewritten
    // where spelled (the records counted, not the attribute), left
    // absent where not; the autoFilter replaced; the pane replaced; a
    // self-closed dataValidations opened.
    {
        const src = "<worksheet " ++ ws_ns ++ "><sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"><pane xSplit=\"2\" topLeftCell=\"C1\" activePane=\"topRight\" state=\"frozen\"/><selection pane=\"topRight\"/></sheetView></sheetViews>" ++
            "<sheetData/><autoFilter ref=\"A1:A9\"><filterColumn colId=\"0\"/></autoFilter><mergeCells count=\"9\"><!-- <mergeCell ref=\"Z1:Z2\"/> --><mergeCell ref=\"X1:Y1\"/></mergeCells>" ++
            "<dataValidations count='0'/><hyperlinks><hyperlink ref=\"A1\" location=\"S!B1\"/></hyperlinks></worksheet>";
        const out = try splice(a, src, &st, .{ .ns_r = ns_r_uri });
        defer a.free(out);
        try t.expectEqualStrings(
            "<worksheet " ++ ws_ns ++ "><sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"><pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/><selection pane=\"topRight\"/></sheetView></sheetViews>" ++
                "<sheetData/><autoFilter ref=\"A1:D1\"/><mergeCells count=\"3\"><!-- <mergeCell ref=\"Z1:Z2\"/> --><mergeCell ref=\"X1:Y1\"/><mergeCell ref=\"A1:B2\"/><mergeCell ref=\"C3:D4\"/></mergeCells>" ++
                "<dataValidations count=\"1\"><dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"E1\"><formula1>&quot;x,y&quot;</formula1></dataValidation></dataValidations>" ++
                "<hyperlinks><hyperlink ref=\"A1\" location=\"S!B1\"/><hyperlink ref=\"F1\" location=\"Sheet1!A1\"/></hyperlinks></worksheet>",
            out,
        );
    }
}

test "S3d slice 2: columns — a covering record is split around the staged column, its other attributes kept; uncovered columns are written fresh; the block is sorted" {
    const a = t.allocator;
    var st: sheet_plan.SheetState = .{};
    defer st.deinit(a);
    try st.setColumnWidth(a, 2, 20); // C
    try st.setColumnWidth(a, 6, 7.5); // G, no record
    try st.setColumnWidth(a, 2, 25); // C again: the last wins
    const src = "<worksheet " ++ ws_ns ++ "><dimension ref=\"A1:H9\"/><cols><col min=\"1\" max=\"4\" width=\"12\" style=\"3\" hidden=\"1\" customWidth=\"1\"/><col min=\"5\" max=\"5\" width=\"9\"/></cols><sheetData/></worksheet>";
    const out = try splice(a, src, &st, .{ .ns_r = ns_r_uri });
    defer a.free(out);
    try t.expectEqualStrings(
        "<worksheet " ++ ws_ns ++ "><dimension ref=\"A1:H9\"/><cols>" ++
            "<col min=\"1\" max=\"2\" width=\"12\" style=\"3\" hidden=\"1\" customWidth=\"1\"/>" ++
            "<col min=\"3\" max=\"3\" width=\"25\" style=\"3\" hidden=\"1\" customWidth=\"1\"/>" ++
            "<col min=\"4\" max=\"4\" width=\"12\" style=\"3\" hidden=\"1\" customWidth=\"1\"/>" ++
            "<col min=\"5\" max=\"5\" width=\"9\"/>" ++
            "<col min=\"7\" max=\"7\" width=\"7.5\" customWidth=\"1\"/>" ++
            "</cols><sheetData/></worksheet>",
        out,
    );
    // No `<cols>`: created before sheetData, after sheetFormatPr.
    var st2: sheet_plan.SheetState = .{};
    defer st2.deinit(a);
    try st2.setColumnWidth(a, 0, 30);
    const src2 = "<worksheet " ++ ws_ns ++ "><sheetFormatPr defaultRowHeight=\"15\"/><sheetData/><pageMargins/></worksheet>";
    const out2 = try splice(a, src2, &st2, .{ .ns_r = ns_r_uri });
    defer a.free(out2);
    try t.expectEqualStrings("<worksheet " ++ ws_ns ++ "><sheetFormatPr defaultRowHeight=\"15\"/><cols><col min=\"1\" max=\"1\" width=\"30\" customWidth=\"1\"/></cols><sheetData/><pageMargins/></worksheet>", out2);
}

test "S3d slice 2: row heights — a held row's open tag gains ht + customHeight, a missing row is created in order, a self-closed sheetData opened" {
    const a = t.allocator;
    var st: sheet_plan.SheetState = .{};
    defer st.deinit(a);
    try st.setRowHeight(a, 1, 30); // row 2: held
    try st.setRowHeight(a, 3, 12.5); // row 4: between 2 and 5
    try st.setRowHeight(a, 8, 9); // row 9: past the last
    const src = "<worksheet " ++ ws_ns ++ "><sheetData><row r=\"2\" spans=\"1:1\" ht=\"15\" customHeight=\"1\"><c r=\"A2\"/></row><row r=\"5\"/></sheetData></worksheet>";
    const out = try splice(a, src, &st, .{ .ns_r = ns_r_uri });
    defer a.free(out);
    try t.expectEqualStrings(
        "<worksheet " ++ ws_ns ++ "><sheetData><row r=\"2\" spans=\"1:1\" ht=\"30\" customHeight=\"1\"><c r=\"A2\"/></row><row r=\"4\" ht=\"12.5\" customHeight=\"1\"/><row r=\"5\"/><row r=\"9\" ht=\"9\" customHeight=\"1\"/></sheetData></worksheet>",
        out,
    );
    const out2 = try splice(a, "<worksheet " ++ ws_ns ++ "><sheetData/></worksheet>", &st, .{ .ns_r = ns_r_uri });
    defer a.free(out2);
    try t.expectEqualStrings("<worksheet " ++ ws_ns ++ "><sheetData><row r=\"2\" ht=\"30\" customHeight=\"1\"/><row r=\"4\" ht=\"12.5\" customHeight=\"1\"/><row r=\"9\" ht=\"9\" customHeight=\"1\"/></sheetData></worksheet>", out2);
}

test "S3d slice 2: conditional formats land after the last block with priorities past the sheet's highest (x14 rules counted); hyperlinks and the legacy drawing declare r when the root does not" {
    const a = t.allocator;
    var st: sheet_plan.SheetState = .{};
    defer st.deinit(a);
    try st.addConditionalFormatDataBar(a, "A1:A9", 0xFF0000FF);
    try st.addHyperlink(a, "B1", "https://example.com/?a=1&b=2");
    try st.addComment(a, "C1", "me", "note");
    const src = "<worksheet " ++ ws_ns ++ "><sheetData/><conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"3\"><formula>1</formula></cfRule></conditionalFormatting><pageMargins/>" ++
        "<extLst><ext><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type=\"dataBar\" priority=\"7\"/></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst></worksheet>";
    var facts = try readSheet(a, src);
    defer facts.deinit(a);
    try t.expectEqual(@as(u32, 7), facts.cf_priority_max);
    try t.expect(facts.legacy_drawing_rid == null);
    try t.expect(!facts.root_declares_r);
    const out = try spliceSheet(a, src, &st, &facts, .{ .ns_r = ns_r_uri, .hyperlink_rid_base = 4, .legacy_drawing_rid = 6 });
    defer a.free(out);
    try t.expectEqualStrings(
        "<worksheet " ++ ws_ns ++ "><sheetData/><conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"3\"><formula>1</formula></cfRule></conditionalFormatting>" ++
            "<conditionalFormatting sqref=\"A1:A9\"><cfRule type=\"dataBar\" priority=\"8\"><dataBar><cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"FF0000FF\"/></dataBar></cfRule></conditionalFormatting>" ++
            "<hyperlinks><hyperlink ref=\"B1\" r:id=\"rId4\" xmlns:r=\"" ++ ns_r_uri ++ "\"/></hyperlinks><pageMargins/>" ++
            "<legacyDrawing r:id=\"rId6\" xmlns:r=\"" ++ ns_r_uri ++ "\"/>" ++
            "<extLst><ext><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type=\"dataBar\" priority=\"7\"/></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst></worksheet>",
        out,
    );
    // The root declaring `r`: no declaration on the elements; a held
    // legacy drawing is reported.
    const src2 = "<worksheet " ++ ws_ns ++ " xmlns:r=\"" ++ ns_r_uri ++ "\"><sheetData/><legacyDrawing r:id=\"rId&#50;\"/></worksheet>";
    var facts2 = try readSheet(a, src2);
    defer facts2.deinit(a);
    try t.expectEqualStrings("rId2", facts2.legacy_drawing_rid.?);
    try t.expect(facts2.root_declares_r);
}

test "S3d slice 2: the walk refuses what it cannot extend in place — no worksheet root, a self-closed one, an element that never closes, an owned element under a prefix or inside a markup-compatibility block; decoys are not elements" {
    const a = t.allocator;
    const refused = [_][]const u8{
        "<x:worksheet xmlns:x=\"u\"><x:sheetData/></x:worksheet>",
        "<worksheet " ++ ws_ns ++ "/>",
        "<worksheet " ++ ws_ns ++ "><sheetData><row r=\"1\"></sheetData></worksheet>",
        "<worksheet " ++ ws_ns ++ "><sheetData/><x:mergeCells xmlns:x=\"u\"/></worksheet>",
        "<worksheet " ++ ws_ns ++ "><sheetData/><mc:AlternateContent><mc:Choice><hyperlinks/></mc:Choice></mc:AlternateContent></worksheet>",
        "<worksheet " ++ ws_ns ++ "><sheetData",
    };
    for (refused) |src| {
        try t.expectError(error.MalformedSheetXml, readSheet(a, src));
    }
    // The save-time shapes, judged at the read (r3 B-ORC-301).
    try t.expectError(error.MalformedSheetXml, readSheet(a, "<worksheet " ++ ws_ns ++ "><cols><col max=\"2\" width=\"9\"/></cols><sheetData/></worksheet>"));
    try t.expectError(error.MalformedSheetXml, readSheet(a, "<worksheet " ++ ws_ns ++ "><cols><col min=\"3\" max=\"2\"/></cols><sheetData/></worksheet>"));
    try t.expectError(error.MalformedSheetXml, readSheet(a, "<worksheet " ++ ws_ns ++ "><sheetData/><conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"4294967295\"/></conditionalFormatting></worksheet>"));
    // An AlternateContent holding nothing of ours, a comment spelling a
    // close tag and a CDATA spelling an open one: readable, and the
    // merge lands after the real mergeCells.
    var st: sheet_plan.SheetState = .{};
    defer st.deinit(a);
    try st.addMergedCell(a, "A1:A2");
    const src = "<worksheet " ++ ws_ns ++ "><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t><![CDATA[<mergeCells>]]></t></is></c></row></sheetData><!-- </mergeCells> --><mergeCells><mergeCell ref=\"B1:B2\"/></mergeCells><mc:AlternateContent><mc:Choice Requires=\"x14\"><controls/></mc:Choice></mc:AlternateContent></worksheet>";
    const out = try splice(a, src, &st, .{ .ns_r = ns_r_uri });
    defer a.free(out);
    try t.expectEqualStrings(
        "<worksheet " ++ ws_ns ++ "><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t><![CDATA[<mergeCells>]]></t></is></c></row></sheetData><!-- </mergeCells> --><mergeCells><mergeCell ref=\"B1:B2\"/><mergeCell ref=\"A1:A2\"/></mergeCells><mc:AlternateContent><mc:Choice Requires=\"x14\"><controls/></mc:Choice></mc:AlternateContent></worksheet>",
        out,
    );
}

test "S3d slice 2: relationships — the next free id past the highest spelled, entries before the root's close, a fresh part when none" {
    const a = t.allocator;
    const rels = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId7\" Type=\"t\" Target=\"x\"/><!-- <Relationship Id=\"rId99\"/> --></Relationships>";
    try t.expectEqual(@as(u32, 8), try nextFreeRelId(rels));
    const out = try appendRelationships(a, rels, &.{.{ .id = 8, .type_uri = "T", .target = "https://e.com/?a=1&b=2", .external = true }});
    defer a.free(out);
    try t.expectEqualStrings(
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId7\" Type=\"t\" Target=\"x\"/><!-- <Relationship Id=\"rId99\"/> --><Relationship Id=\"rId8\" Type=\"T\" Target=\"https://e.com/?a=1&amp;b=2\" TargetMode=\"External\"/></Relationships>",
        out,
    );
    const fresh = try appendRelationships(a, null, &.{.{ .id = 1, .type_uri = "T", .target = "../comments1.xml", .external = false }});
    defer a.free(fresh);
    try t.expectEqualStrings(rels_head ++ "<Relationship Id=\"rId1\" Type=\"T\" Target=\"../comments1.xml\"/></Relationships>", fresh);
    try t.expectError(error.MalformedSheetRels, appendRelationships(a, "<Other/>", &.{}));
}

test "S3d slice 2: comments — a known author keeps its id, a new one is appended, the records land after the list's; refs read back decoded; a prefixed part refuses" {
    const a = t.allocator;
    const src = "<comments " ++ ws_ns ++ "><authors><author>bob</author><author>a&amp;b</author></authors><commentList><comment ref=\"A1\" authorId=\"0\"><text><t>x</t></text></comment></commentList></comments>";
    const refs = try commentRefs(a, src);
    defer {
        for (refs) |r| a.free(r);
        a.free(refs);
    }
    try t.expectEqual(@as(usize, 1), refs.len);
    try t.expectEqualStrings("A1", refs[0]);
    const out = try extendComments(a, src, &.{
        .{ .ref = "B2", .author = "a&b", .text = "two" },
        .{ .ref = "C3", .author = "carol", .text = "<3" },
    });
    defer a.free(out);
    try t.expectEqualStrings(
        "<comments " ++ ws_ns ++ "><authors><author>bob</author><author>a&amp;b</author><author>carol</author></authors><commentList><comment ref=\"A1\" authorId=\"0\"><text><t>x</t></text></comment>" ++
            "<comment ref=\"B2\" authorId=\"1\"><text><t xml:space=\"preserve\">two</t></text></comment><comment ref=\"C3\" authorId=\"2\"><text><t xml:space=\"preserve\">&lt;3</t></text></comment></commentList></comments>",
        out,
    );
    // A bare root: both tables created.
    const bare = try extendComments(a, "<comments " ++ ws_ns ++ "></comments>", &.{.{ .ref = "A1", .author = "me", .text = "n" }});
    defer a.free(bare);
    try t.expectEqualStrings("<comments " ++ ws_ns ++ "><authors><author>me</author></authors><commentList><comment ref=\"A1\" authorId=\"0\"><text><t xml:space=\"preserve\">n</t></text></comment></commentList></comments>", bare);
    try t.expectError(error.MalformedCommentsXml, commentRefs(a, "<x:comments/>"));
    try t.expectError(error.MalformedCommentsXml, extendComments(a, "<comments><commentList>", &.{}));
}

test "S3d slice 2: VML — shape ids past the highest held, the id map extended to the new block, the note shape type added when missing" {
    const a = t.allocator;
    const src = "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"><o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/></o:shapelayout>" ++
        "<v:shape id=\"_x0000_s2047\" type=\"#_x0000_t202\"><x:ClientData ObjectType=\"Note\"><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape></xml>";
    const out = try extendVml(a, src, &.{ .{ .ref = "B2", .author = "", .text = "" }, .{ .ref = "C3", .author = "", .text = "" } });
    defer a.free(out);
    var expect: std.ArrayListUnmanaged(u8) = .empty;
    defer expect.deinit(a);
    try expect.appendSlice(a, "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"><o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1,2\"/></o:shapelayout>");
    try expect.appendSlice(a, sheet_plan.VML_NOTE_SHAPETYPE);
    try expect.appendSlice(a, "<v:shape id=\"_x0000_s2047\" type=\"#_x0000_t202\"><x:ClientData ObjectType=\"Note\"><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape>");
    try sheet_plan.emitVmlNoteShape(a, &expect, "B2", 2048, 2);
    try sheet_plan.emitVmlNoteShape(a, &expect, "C3", 2049, 3);
    try expect.appendSlice(a, "</xml>");
    try t.expectEqualStrings(expect.items, out);
    // A drawing without a layout or a shape: the fresh part's shape.
    const bare = try extendVml(a, "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"></xml>", &.{.{ .ref = "A1", .author = "", .text = "" }});
    defer a.free(bare);
    var fresh: std.ArrayListUnmanaged(u8) = .empty;
    defer fresh.deinit(a);
    try fresh.appendSlice(a, "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"><o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/></o:shapelayout>");
    try fresh.appendSlice(a, sheet_plan.VML_NOTE_SHAPETYPE);
    try sheet_plan.emitVmlNoteShape(a, &fresh, "A1", 1025, 1);
    try fresh.appendSlice(a, "</xml>");
    try t.expectEqualStrings(fresh.items, bare);
    try t.expectError(error.MalformedVmlDrawing, extendVml(a, "<xml><v:shape", &.{}));
}

test "S3d slice 2: the attribute rewriter keeps unknown attributes and their spelling, replaces in place, appends what is missing" {
    const a = t.allocator;
    var out: std.ArrayListUnmanaged(u8) = .empty;
    defer out.deinit(a);
    const subs = [_]AttrSub{ .{ .name = "count", .value = "5" }, .{ .name = "new", .value = "y" } };
    try writeTagWithAttrs(a, &out, "mergeCells", " a='1' count = \"2\"  b=\"x>y\"", &subs, ">");
    try t.expectEqualStrings("<mergeCells a='1' count=\"5\" b=\"x>y\" new=\"y\">", out.items);
}

test "S3d slice 2 r1: a self-closed rels root is opened (B-ORC-103); a shape id nested in a group counts (B-VML-104); overlapping staged spans rebuild without skipping (A-SPL-102); a decoy inside an AlternateContent block is not a mention (A-SPL-103)" {
    const a = t.allocator;
    const out = try appendRelationships(a, "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>", &.{.{ .id = 1, .type_uri = "T", .target = "x", .external = false }});
    defer a.free(out);
    try t.expectEqualStrings("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"T\" Target=\"x\"/></Relationships>", out);

    const vml = try extendVml(a, "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"><v:group><v:shape id=\"_x0000_s2000\"/></v:group></xml>", &.{.{ .ref = "A1", .author = "", .text = "" }});
    defer a.free(vml);
    try t.expect(std.mem.indexOf(u8, vml, "id=\"_x0000_s2001\"") != null);

    const spans = try stagedColSpans(a, &.{ .{ .col_min = 1, .col_max = 1, .width = 1 }, .{ .col_min = 2, .col_max = 2, .width = 2 }, .{ .col_min = 1, .col_max = 2, .width = 3 } });
    defer a.free(spans);
    try t.expectEqual(@as(usize, 1), spans.len);
    try t.expectEqual(@as(u32, 1), spans[0].min);
    try t.expectEqual(@as(u32, 2), spans[0].max);
    const spans2 = try stagedColSpans(a, &.{ .{ .col_min = 1, .col_max = 5, .width = 1 }, .{ .col_min = 3, .col_max = 3, .width = 2 } });
    defer a.free(spans2);
    try t.expectEqual(@as(usize, 3), spans2.len);

    var st: sheet_plan.SheetState = .{};
    defer st.deinit(a);
    try st.addMergedCell(a, "A1:A2");
    const src = "<worksheet " ++ ws_ns ++ "><sheetData/><mc:AlternateContent><!-- <mergeCells/> --><mc:Choice><![CDATA[<hyperlinks>]]></mc:Choice></mc:AlternateContent></worksheet>";
    const spliced = try splice(a, src, &st, .{ .ns_r = ns_r_uri });
    defer a.free(spliced);
    try t.expect(std.mem.indexOf(u8, spliced, "<mergeCells count=\"1\"><mergeCell ref=\"A1:A2\"/></mergeCells><mc:AlternateContent>") != null);
}
