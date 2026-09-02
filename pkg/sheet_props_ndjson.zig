//! The `sheet-props` and `calc-props` NDJSON records — S3b's typed
//! read of freeze / split panes, the `<dimension>` extent and the
//! workbook's calculation properties (`docs/cli.md`, "sheet-props" /
//! "calc-props"), written once for every surface.
//!
//! `zlsx sheet-props` / `zlsx calc-props` emit these through the CLI's
//! selection and pagination; the C and Python legs of row S3b hand
//! over the same bytes when they land — the `pivot_ndjson.zig` /
//! `conditional_format_ndjson.zig` precedent.
//!
//! `sheet_props`: one record per workbook sheet, workbook order — the
//! sheet's `<dimension ref>` as authored (null when the element or the
//! attribute is absent) and the `<pane>` of its FIRST `<sheetView>` as
//! authored (null when there is none): `xSplit` / `ySplit` /
//! `topLeftCell` / `activePane` / `state`, each null when the source
//! omits it, no schema default applied, split panes reported as they
//! are — the lenient `Worksheet.freezePane` narrows to frozen panes
//! and keeps that contract; this read does not. `calc_props`: exactly
//! one record per workbook from `xl/workbook.xml`'s `<calcPr>` —
//! `calcId`, `fullCalcOnLoad`, `iterate`, `iterateCount`,
//! `iterateDelta` as authored, every field null when the element or
//! the attribute is absent (a workbook without `<calcPr>` is a record
//! of nulls, the `doc-props` convention).
//!
//! The read is STRICT where the Zig view is lenient — the
//! conditional-formats precedent: each sheet part is walked by the
//! namespace- and depth-aware scanner so the wire can prove what it
//! reports. A `dimension` is the extent only as a main-namespace
//! direct child of a worksheet / macrosheet root (the roots the schema
//! gives one), a `sheetViews` only as a direct child of a root whose
//! views carry panes (worksheet, macrosheet, dialogsheet — a
//! chartsheet's views cannot), a `sheetView` only as its direct child,
//! a `pane` only as a direct child of the FIRST `sheetView` — so a
//! `<pane>` under `<customSheetViews>`, a `<dimension>` under
//! `<extLst>` or inside a rebound subtree can never masquerade as the
//! slot, and a second `<dimension>`, `<sheetViews>` or first-view
//! `<pane>` (each maxOccurs=1) is a refusal, not a silent pick. Later
//! `<sheetView>` elements (one per extra Excel window) keep their own
//! panes in the part; the record reports the primary view's. The
//! `<calcPr>` slot rides the strict workbook read the sheet inventory
//! comes from (`conditional_format_ndjson.scanCalcPr`): a
//! main-namespace direct child of the root, one at most, refused when
//! an MCE branch the walk has no processor for could project one into
//! the slot.
//!
//! Decode discipline: `ref`, `topLeftCell`, `activePane`, `state` are
//! entities-only attribute carriers (no ST_Xstring layer, the C1
//! ruling) — a carrier that does not decode, decodes to non-UTF-8, or
//! carries markup refuses the whole read. The typed attributes decode
//! their entities FIRST (an undecodable carrier refuses), collapse
//! boundary XML whitespace (XSD's `whiteSpace="collapse"` for the
//! atomic types), then type lexically: `calcId` / `iterateCount`
//! digit-only u32 or null;
//! `xSplit` / `ySplit` / `iterateDelta` an `xsd:double` lexical or
//! null (a non-finite value — `INF`, `NaN`, an overflow — has no JSON
//! spelling and reads null with the rest); `fullCalcOnLoad` /
//! `iterate` an `xsd:boolean` lexical (`true` / `false` / `1` / `0`)
//! or null — the written-but-invalid-reads-absent convention of the
//! family's numeric attributes (Codex #215 r1 REL-104, r18 REL-1802).

const std = @import("std");
const formula_mod = @import("zlsx_formula");
const workbook_mod = @import("workbook.zig");
const store_mod = @import("store.zig");
const workbook_xml = @import("typed_parts/root.zig").workbook_xml;
const sheet_xml = @import("typed_parts/root.zig").sheet_xml;
const cf = @import("conditional_format_ndjson.zig");
const json = @import("json_text.zig");

const Allocator = std.mem.Allocator;
const PartStore = store_mod.PartStore;

/// Whether a record carries the `sheet` / `sheet_idx` envelope.
/// `compact` is the CLI's `--output compact-ndjson`, where a sheet
/// prologue record names the sheet once for every record after it.
pub const Envelope = enum { full, compact };

/// The first `<sheetView>`'s `<pane>` as authored. Every field is
/// null when the source omits the attribute — the schema's defaults
/// (`xSplit` / `ySplit` 0, `activePane` bottomRight, `state` split)
/// are the reader's to apply, not the wire's.
pub const Pane = struct {
    x_split: ?f64,
    y_split: ?f64,
    top_left_cell: ?[]const u8,
    active_pane: ?[]const u8,
    state: ?[]const u8,
};

/// One sheet's properties, attributed and decoded.
pub const SheetRecord = struct {
    sheet: []const u8,
    sheet_idx: u32,
    /// The `<dimension ref>` as authored, or null when the sheet
    /// spells neither the element nor its attribute.
    dimension: ?[]const u8,
    pane: ?Pane,
};

/// The workbook's `<calcPr>` as authored. Plain data — no strings, no
/// arena — so a caller holds it by value.
pub const CalcRecord = struct {
    calc_id: ?u32,
    full_calc_on_load: ?bool,
    iterate: ?bool,
    iterate_count: ?u32,
    iterate_delta: ?f64,

    /// The record of a workbook without `<calcPr>`.
    pub const absent: CalcRecord = .{
        .calc_id = null,
        .full_calc_on_load = null,
        .iterate = null,
        .iterate_count = null,
        .iterate_delta = null,
    };
};

/// Every sheet's properties, workbook order. Owns its decoded strings;
/// `deinit` frees them.
pub const SheetProps = struct {
    arena: std.heap.ArenaAllocator,
    /// Decoded sheet names in workbook order — the strict inventory
    /// (`conditional_format_ndjson.resolveSheets`) the CLI's selectors
    /// and the `sheet` field read from.
    sheet_names: []const []const u8,
    /// Parallel to `sheet_names`: one record per sheet.
    records: []const SheetRecord,

    pub fn deinit(self: *SheetProps) void {
        self.arena.deinit();
        self.* = undefined;
    }
};

pub const Error = error{
    /// A sheet list the strict workbook read cannot prove
    /// (`conditional_format_ndjson.resolveSheets`), or — on the
    /// calc-props read — a `<calcPr>` slot it cannot report faithfully:
    /// two elements at the slot, one an MCE branch could project
    /// there, a duplicate attribute, a carrier that does not decode.
    MalformedWorkbookXml,
    /// A sheet part the strict walk cannot prove a pane / extent for
    /// (mismatched nesting, a root that does not bind the main
    /// namespace as its default, the main namespace aliased to a
    /// prefix, a second `<dimension>` / `<sheetViews>` / first-view
    /// `<pane>`, a duplicate attribute on that machinery, an MCE
    /// construct at a recognized slot) — or a carrier the NDJSON
    /// cannot carry faithfully (one that does not decode, or is not
    /// UTF-8).
    MalformedSheetXml,
    OutOfMemory,
};

/// What `collect` itself can raise — `Error` plus the two verdicts
/// that keep their own names across every boundary: the archive-wide
/// decompression caps, and a sheet part the store lost between the
/// graph proof and the walk. Closed and explicit so the C boundary's
/// status mapping is compiler-checked, not assumed (the
/// conditional-formats read's rule, Codex #216 r1 S3B-ERR-602).
pub const CollectError = Error || error{ ZipBombSuspected, MissingSheetPart };

/// What `collectCalc` can raise: the workbook verdicts plus the
/// archive-wide decompression caps. No sheet part is read.
pub const CalcError = error{ MalformedWorkbookXml, OutOfMemory, ZipBombSuspected };

/// Collect every workbook sheet's properties. A read this module
/// cannot serve faithfully refuses whole — even for a sheet the CLI's
/// selection would exclude: the inventory is proven whole before
/// selection and pagination apply, the `anchors` rule.
pub fn collect(gpa: Allocator, wb: *workbook_mod.Workbook) CollectError!SheetProps {
    var arena = std.heap.ArenaAllocator.init(gpa);
    errdefer arena.deinit();
    const a = arena.allocator();

    // The AUTHORITATIVE sheet inventory — the strict workbook read the
    // conditional-formats and anchors reads share, names and parts
    // resolved into this view's arena.
    const inventory = try cf.resolveSheets(gpa, a, wb);
    const sheet_names = inventory.names;
    const sheet_parts = inventory.parts;

    const records = try a.alloc(SheetRecord, sheet_names.len);
    for (sheet_parts, 0..) |part_name, idx| {
        // Store failures fold like `Worksheet.ensureParsed`: memory
        // and the archive-wide budget keep their names, everything
        // else is "this sheet is not readable".
        const part = (wb.store.part(part_name) catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            error.ZipBombSuspected => return error.ZipBombSuspected,
            else => return error.MalformedSheetXml,
        }) orelse return error.MissingSheetPart;

        // Scratch (the walker's frame stack, the prefix scope) lives
        // on the caller's gpa and is reclaimed per sheet; only decoded
        // strings land on the view arena.
        const raw = try scanSheetProps(gpa, part.bytes);

        // The typed sheet view's parse verdict stays part of the
        // family contract for worksheet parts (the conditional-formats
        // rule, Codex #215 r4 REL-404): a part the view refuses
        // refuses here too. The other roots have no typed view; the
        // strict walk's verdict stands alone.
        if (raw.root == .worksheet) {
            var gate = sheet_xml.parse(gpa, part.bytes) catch |e| switch (e) {
                error.OutOfMemory => return error.OutOfMemory,
                else => return error.MalformedSheetXml,
            };
            gate.deinit(gpa);
        }

        records[idx] = .{
            .sheet = sheet_names[idx],
            .sheet_idx = @intCast(idx),
            .dimension = try decodeOpt(a, raw.dimension),
            .pane = if (raw.pane) |p| .{
                .x_split = try doubleAttr(gpa, p.x_split),
                .y_split = try doubleAttr(gpa, p.y_split),
                .top_left_cell = try decodeOpt(a, p.top_left_cell),
                .active_pane = try decodeOpt(a, p.active_pane),
                .state = try decodeOpt(a, p.state),
            } else null,
        };
    }

    return .{
        .arena = arena,
        .sheet_names = sheet_names,
        .records = records,
    };
}

/// The workbook's `<calcPr>`, read strictly. A workbook without one is
/// `CalcRecord.absent`; a slot the read cannot report faithfully
/// refuses whole.
pub fn collectCalc(gpa: Allocator, wb: *workbook_mod.Workbook) CalcError!CalcRecord {
    const slot = try cf.scanCalcPr(gpa, wb);
    // The schema's maxOccurs=1: a second element at the slot is not
    // "the first one wins" — which Excel honours is not the reader's
    // to guess. A branch an MCE processor could project into the slot
    // is a slot the walk cannot rule in or out (SEC-2201).
    if (slot.count > 1) return error.MalformedWorkbookXml;
    if (slot.mce_shadowed) return error.MalformedWorkbookXml;
    const attrs = slot.attrs orelse return CalcRecord.absent;
    return calcFromAttrs(gpa, attrs) catch |e| switch (e) {
        error.OutOfMemory => return error.OutOfMemory,
        // The lexical helpers speak the sheet part's error name; here
        // the carrier is the workbook part's.
        error.MalformedSheetXml => return error.MalformedWorkbookXml,
        error.MalformedWorkbookXml => return error.MalformedWorkbookXml,
    };
}

fn calcFromAttrs(scratch: Allocator, attrs: []const u8) Error!CalcRecord {
    // `uniqueAttr` refuses a duplicate of ANY attribute name on the
    // tag (the S7b-4 rule) on each lookup, so the five reads share one
    // verdict about the region.
    return .{
        .calc_id = try uintAttr(scratch, try cf.uniqueAttr(scratch, attrs, "calcId")),
        .full_calc_on_load = try boolAttr(scratch, try cf.uniqueAttr(scratch, attrs, "fullCalcOnLoad")),
        .iterate = try boolAttr(scratch, try cf.uniqueAttr(scratch, attrs, "iterate")),
        .iterate_count = try uintAttr(scratch, try cf.uniqueAttr(scratch, attrs, "iterateCount")),
        .iterate_delta = try doubleAttr(scratch, try cf.uniqueAttr(scratch, attrs, "iterateDelta")),
    };
}

/// One `{"kind":"sheet_props",…}` line. The field order is the
/// docs/cli.md contract; a change here is a wire-format change on
/// every surface at once.
pub fn writeSheetRecord(out: *std.Io.Writer, r: SheetRecord, envelope: Envelope) !void {
    try out.writeAll("{\"kind\":\"sheet_props\"");
    if (envelope == .full) {
        try out.writeAll(",\"sheet\":");
        try json.writeString(out, r.sheet);
        try out.print(",\"sheet_idx\":{d}", .{r.sheet_idx});
    }
    try out.writeAll(",\"dimension\":");
    try json.writeOptString(out, r.dimension);
    try out.writeAll(",\"pane\":");
    if (r.pane) |p| {
        try out.writeAll("{\"x_split\":");
        try json.writeOptF64(out, p.x_split);
        try out.writeAll(",\"y_split\":");
        try json.writeOptF64(out, p.y_split);
        try out.writeAll(",\"top_left_cell\":");
        try json.writeOptString(out, p.top_left_cell);
        try out.writeAll(",\"active_pane\":");
        try json.writeOptString(out, p.active_pane);
        try out.writeAll(",\"state\":");
        try json.writeOptString(out, p.state);
        try out.writeByte('}');
    } else {
        try out.writeAll("null");
    }
    try out.writeAll("}\n");
}

/// The unselected stream — every sheet record, workbook order. The C
/// leg's entry point.
pub fn writeAll(out: *std.Io.Writer, view: *const SheetProps) !void {
    for (view.records) |r| try writeSheetRecord(out, r, .full);
}

/// The one `{"kind":"calc_props",…}` line. There is no sheet
/// envelope to drop, so no `Envelope` — compact and full are one
/// shape.
pub fn writeCalcRecord(out: *std.Io.Writer, r: CalcRecord) !void {
    try out.writeAll("{\"kind\":\"calc_props\",\"calc_id\":");
    try json.writeOptU32(out, r.calc_id);
    try out.writeAll(",\"full_calc_on_load\":");
    try json.writeOptBool(out, r.full_calc_on_load);
    try out.writeAll(",\"iterate\":");
    try json.writeOptBool(out, r.iterate);
    try out.writeAll(",\"iterate_count\":");
    try json.writeOptU32(out, r.iterate_count);
    try out.writeAll(",\"iterate_delta\":");
    try json.writeOptF64(out, r.iterate_delta);
    try out.writeAll("}\n");
}

// ─── The strict per-sheet walk ───────────────────────────────────────

/// The first view's pane as sliced out of the part, undecoded; every
/// slice borrows the part's bytes.
const RawPane = struct {
    x_split: ?[]const u8 = null,
    y_split: ?[]const u8 = null,
    top_left_cell: ?[]const u8 = null,
    active_pane: ?[]const u8 = null,
    state: ?[]const u8 = null,
};

const RawSheetProps = struct {
    root: cf.SheetRoot,
    /// The `ref` attribute as sliced; null when the element or the
    /// attribute is absent.
    dimension: ?[]const u8,
    pane: ?RawPane,
};

/// Which roots give the schema slots this read classifies:
/// `CT_Worksheet` and `CT_Macrosheet` carry `<dimension>`;
/// `CT_Dialogsheet` shares their `CT_SheetViews` (whose views carry a
/// pane); `CT_Chartsheet`'s `CT_ChartsheetViews` carry none. An
/// element spelling a slot's name under a root without that slot is
/// opaque content, like a decoy under `<extLst>`.
fn hasDimensionSlot(root: cf.SheetRoot) bool {
    return switch (root) {
        .worksheet, .macrosheet => true,
        .chartsheet, .dialogsheet => false,
    };
}

fn hasPaneSlot(root: cf.SheetRoot) bool {
    return switch (root) {
        .worksheet, .macrosheet, .dialogsheet => true,
        .chartsheet => false,
    };
}

/// Walk one sheet part and slice out its extent and first-view pane —
/// the strict read the module doc-comment describes, on the
/// conditional-formats scanner's lexical layer: the tree is tokenized
/// whole (strict prolog / comment / PI skipping, DOCTYPE refusal,
/// quote-aware tag ends), element namespaces are tracked through
/// default declarations and the in-scope prefix stack, and
/// classification is by depth. Refusals: a root that is not a
/// sheet-family root under a main default, the main namespace bound
/// to a prefix anywhere, mismatched or unterminated nesting, a second
/// element at a maxOccurs=1 slot, a duplicate attribute on the slot
/// machinery, and — the walk has no MCE processor — an MCE-bound
/// element that is a direct child of the root, the views wrapper or
/// the first view, the `mc:ProcessContent` attribute anywhere, a
/// default binding of the MCE namespace (SEC-2201). Everything else
/// is an opaque subtree the walk steps over.
fn scanSheetProps(a: Allocator, xml: []const u8) Error!RawSheetProps {
    // The whole part must be UTF-8 before any byte-level scan (Codex
    // #215 r10 SEC-1001).
    if (!std.unicode.utf8ValidateSlice(xml)) return error.MalformedSheetXml;
    const Kind = enum { root, sheet_views, first_view, other };
    const Frame = struct { name: []const u8, kind: Kind, main_default: bool, prefix_mark: usize };
    var frames: std.ArrayListUnmanaged(Frame) = .empty;
    defer frames.deinit(a);
    var scope: cf.PrefixScope = .{};
    defer scope.deinit(a);
    var decl_seen: std.StringHashMapUnmanaged(void) = .empty;
    defer decl_seen.deinit(a);

    var root_seen = false;
    var root_closed = false;
    var root_kind: cf.SheetRoot = .worksheet;
    var dimension_count: u32 = 0;
    var dimension_ref: ?[]const u8 = null;
    var wrappers_seen: u32 = 0;
    var views_seen: u32 = 0;
    var panes_seen: u32 = 0;
    var pane: ?RawPane = null;
    var i: usize = 0;
    while (std.mem.indexOfScalarPos(u8, xml, i, '<')) |lt| {
        // OUTSIDE the root only comments, PIs and whitespace may
        // appear (a BOM at byte zero included) — REL-1102.
        if (frames.items.len == 0) {
            const gap_start: usize = if (i == 0 and std.mem.startsWith(u8, xml, "\xEF\xBB\xBF")) 3 else i;
            for (xml[gap_start..lt]) |c| {
                if (!cf.isXmlWs(c)) return error.MalformedSheetXml;
            }
            if (lt + 9 <= xml.len and std.mem.eql(u8, xml[lt .. lt + 9], "<![CDATA[")) {
                return error.MalformedSheetXml;
            }
        }
        // A DOCTYPE can define entities this byte walk cannot see —
        // refuse the declaration rather than skip it (SEC-302).
        if (lt + 1 < xml.len and xml[lt + 1] == '!') {
            const is_comment = lt + 4 <= xml.len and std.mem.eql(u8, xml[lt .. lt + 4], "<!--");
            const is_cdata = lt + 9 <= xml.len and std.mem.eql(u8, xml[lt .. lt + 9], "<![CDATA[");
            if (!is_comment and !is_cdata) return error.MalformedSheetXml;
            if (is_comment) {
                i = cf.skipStrictComment(xml, lt) orelse return error.MalformedSheetXml;
                continue;
            }
        }
        if (lt + 1 < xml.len and xml[lt + 1] == '?') {
            if (cf.isPrologXmlDecl(xml, lt)) {
                i = cf.skipStrictXmlDecl(xml, lt) orelse return error.MalformedSheetXml;
                continue;
            }
            i = cf.skipStrictPi(xml, lt) orelse return error.MalformedSheetXml;
            continue;
        }
        // What remains of the non-element constructs here is a CDATA
        // section inside the root — character data the walk steps over
        // whole (the conditional-formats walk's call, kept identical).
        const skip_to = workbook_xml.skipNonElement(xml, lt) catch return error.MalformedSheetXml;
        if (skip_to != lt) {
            i = skip_to;
            continue;
        }
        if (lt + 1 >= xml.len) return error.MalformedSheetXml;

        if (xml[lt + 1] == '/') {
            // Close tag: the name must match the open frame exactly.
            var j = lt + 2;
            const name_start = j;
            while (j < xml.len and xml[j] != '>' and !cf.isXmlWs(xml[j])) j += 1;
            const name = xml[name_start..j];
            while (j < xml.len and cf.isXmlWs(xml[j])) j += 1;
            if (j >= xml.len or xml[j] != '>') return error.MalformedSheetXml;
            if (frames.items.len == 0) return error.MalformedSheetXml;
            const top = frames.items[frames.items.len - 1];
            frames.items.len -= 1;
            if (!std.mem.eql(u8, top.name, name)) return error.MalformedSheetXml;
            scope.truncate(top.prefix_mark);
            if (top.kind == .root) root_closed = true;
            i = j + 1;
            continue;
        }

        // Open tag.
        const name_start = lt + 1;
        var j = name_start;
        while (j < xml.len and !cf.isTagBoundary(xml[j])) j += 1;
        if (j == name_start or j >= xml.len) return error.MalformedSheetXml;
        const qname = xml[name_start..j];
        if (!cf.validQName(qname)) return error.MalformedSheetXml;
        const te = cf.tagEnd(xml, j) orelse return error.MalformedSheetXml;
        const attrs = xml[j..te.attrs_end];
        if (root_closed) return error.MalformedSheetXml; // a second root

        // Namespace bookkeeping: the default declared here binds this
        // element too; the main namespace bound to a prefix is outside
        // the closed form and refuses (the host/SST predicate's rule).
        const prefix_mark = scope.mark();
        try cf.enterElementScope(a, &scope, qname, attrs);
        var elem_main = if (frames.items.len == 0)
            false
        else
            frames.items[frames.items.len - 1].main_default;
        {
            var it: cf.AttrScan = .{ .rest = attrs };
            // A namespace declaration repeated on one start tag makes
            // the binding ambiguous; refuse before reading either
            // value (SEC-703).
            defer cf.resetDeclSeen(a, &decl_seen);
            while (try it.next()) |attr| {
                const is_decl = std.mem.eql(u8, attr.name, "xmlns") or
                    std.mem.startsWith(u8, attr.name, "xmlns:");
                if (is_decl) {
                    const g = try decl_seen.getOrPut(a, attr.name);
                    if (g.found_existing) return error.MalformedSheetXml;
                }
                if (std.mem.eql(u8, attr.name, "xmlns")) {
                    elem_main = try cf.bindsMainNs(a, attr.value);
                } else if (std.mem.startsWith(u8, attr.name, "xmlns:")) {
                    if (try cf.bindsMainNs(a, attr.value)) return error.MalformedSheetXml;
                }
            }
        }
        const prefixed = std.mem.indexOfScalar(u8, qname, ':') != null;
        const is_main = !prefixed and elem_main;

        var kind: Kind = .other;
        if (frames.items.len == 0) {
            root_seen = true;
            // Every sheet-family root is WALKED whole, the slot-less
            // ones included: an early return would bypass the
            // second-root and unterminated-nesting checks (REL-201).
            root_kind = (try cf.classifySheetRoot(a, qname, attrs, is_main, elem_main)) orelse
                return error.MalformedSheetXml;
            kind = .root;
        } else {
            const parent = frames.items[frames.items.len - 1];
            // An MCE-bound element AT a recognized slot is where
            // branch substitution could add or hide the extent, the
            // views wrapper, the first view or its pane — deeper MCE
            // projects in place inside an opaque subtree and stays
            // walkable (SEC-2201).
            if (prefixed) {
                const c = std.mem.indexOfScalar(u8, qname, ':').?;
                const slot = switch (parent.kind) {
                    .root, .sheet_views, .first_view => true,
                    .other => false,
                };
                if (slot and scope.isMce(qname[0..c])) return error.MalformedSheetXml;
            }
            if (parent.kind == .root and is_main) {
                if (std.mem.eql(u8, qname, "dimension") and hasDimensionSlot(root_kind)) {
                    dimension_count += 1;
                    if (dimension_count > 1) return error.MalformedSheetXml;
                    dimension_ref = try cf.uniqueAttr(a, attrs, "ref");
                } else if (std.mem.eql(u8, qname, "sheetViews") and hasPaneSlot(root_kind)) {
                    wrappers_seen += 1;
                    if (wrappers_seen > 1) return error.MalformedSheetXml;
                    // The wrapper carries no attribute this read
                    // consumes, but it is slot machinery: a name twice
                    // on its start tag refuses like on `<pane>` (Codex
                    // #218 r1 S3B-REL-101).
                    _ = try cf.uniqueAttr(a, attrs, "");
                    kind = .sheet_views;
                }
            } else if (parent.kind == .sheet_views and is_main and std.mem.eql(u8, qname, "sheetView")) {
                views_seen += 1;
                // Only the FIRST view is a slot; later views are the
                // extra windows' — opaque here, reported by no field,
                // their attributes unpoliced like any opaque tag's.
                if (views_seen == 1) {
                    _ = try cf.uniqueAttr(a, attrs, "");
                    kind = .first_view;
                }
            } else if (parent.kind == .first_view and is_main and std.mem.eql(u8, qname, "pane")) {
                panes_seen += 1;
                if (panes_seen > 1) return error.MalformedSheetXml;
                pane = try paneAttrs(a, attrs);
            }
        }

        if (te.self_closing) {
            scope.truncate(prefix_mark);
            if (kind == .root) root_closed = true;
        } else {
            if (frames.items.len >= cf.max_depth) return error.MalformedSheetXml;
            try frames.append(a, .{ .name = qname, .kind = kind, .main_default = elem_main, .prefix_mark = prefix_mark });
        }
        i = te.after_gt;
    }
    // Trailing character data after the last construct must be
    // whitespace too (REL-1102).
    for (xml[i..]) |c| {
        if (!cf.isXmlWs(c)) return error.MalformedSheetXml;
    }
    if (!root_seen) return error.MalformedSheetXml;
    if (frames.items.len != 0) return error.MalformedSheetXml;
    return .{ .root = root_kind, .dimension = dimension_ref, .pane = pane };
}

/// The five pane attributes in one pass, refusing a duplicate of ANY
/// attribute name on the tag (the `uniqueAttr` rule, applied once).
fn paneAttrs(a: Allocator, attrs: []const u8) Error!RawPane {
    var seen: std.StringHashMapUnmanaged(void) = .empty;
    defer seen.deinit(a);
    var p: RawPane = .{};
    var it: cf.AttrScan = .{ .rest = attrs };
    while (try it.next()) |attr| {
        const g = try seen.getOrPut(a, attr.name);
        if (g.found_existing) return error.MalformedSheetXml;
        if (std.mem.eql(u8, attr.name, "xSplit")) {
            p.x_split = attr.value;
        } else if (std.mem.eql(u8, attr.name, "ySplit")) {
            p.y_split = attr.value;
        } else if (std.mem.eql(u8, attr.name, "topLeftCell")) {
            p.top_left_cell = attr.value;
        } else if (std.mem.eql(u8, attr.name, "activePane")) {
            p.active_pane = attr.value;
        } else if (std.mem.eql(u8, attr.name, "state")) {
            p.state = attr.value;
        }
    }
    return p;
}

// ─── Carrier decoding ────────────────────────────────────────────────

/// An optional entities-only attribute carrier, decoded into the view
/// arena — the conditional-formats `sqref` / `type` decode: raw markup
/// refuses, a bad entity refuses, non-UTF-8 refuses.
fn decodeOpt(a: Allocator, raw: ?[]const u8) Error!?[]const u8 {
    const r = raw orelse return null;
    return try cf.decodeRuleText(a, r);
}

/// A typed attribute: XML character references resolve FIRST (a
/// reference that does not decode refuses the read — `numericAttr`'s
/// rule, Codex #215 r18 REL-1802), THEN boundary XML whitespace
/// collapses — XSD fixes `whiteSpace` to `collapse` for every atomic
/// type these attributes carry, so `" 2 "` and `"&#x20;true "` are
/// the typed values `2` and `true` (Codex #218 r1 S3B-REL-103); an
/// interior run would leave the token invalid under every lexical
/// rule below, so trimming the ends is the whole of collapse here —
/// THEN the lexical rule applies.
fn typedAttr(comptime T: type, comptime lex: fn ([]const u8) ?T, scratch: Allocator, raw_opt: ?[]const u8) Error!?T {
    const raw = raw_opt orelse return null;
    if (std.mem.indexOfScalar(u8, raw, '&') == null) return lex(collapseWs(raw));
    const decoded = formula_mod.decode.decodeEntities(scratch, raw) catch |e| switch (e) {
        error.OutOfMemory => return error.OutOfMemory,
        else => return error.MalformedSheetXml,
    };
    defer scratch.free(decoded);
    return lex(collapseWs(decoded));
}

fn collapseWs(s: []const u8) []const u8 {
    return std.mem.trim(u8, s, " \t\n\r");
}

fn boolAttr(scratch: Allocator, raw_opt: ?[]const u8) Error!?bool {
    return typedAttr(bool, xsdBoolean, scratch, raw_opt);
}

fn doubleAttr(scratch: Allocator, raw_opt: ?[]const u8) Error!?f64 {
    return typedAttr(f64, xsdDouble, scratch, raw_opt);
}

/// `xsd:unsignedInt` — the conditional-formats `numericAttr` rule
/// (digit-only, `u32` range, else null) behind the same decode-then-
/// collapse discipline as the other typed attributes here.
fn uintAttr(scratch: Allocator, raw_opt: ?[]const u8) Error!?u32 {
    return typedAttr(u32, cf.digitOnlyU32, scratch, raw_opt);
}

/// `xsd:boolean`'s lexical space, exactly: `true` / `false` / `1` /
/// `0`. Anything else is written-but-invalid and reads null — the
/// lenient view's "anything but true is false" is its own contract.
pub fn xsdBoolean(s: []const u8) ?bool {
    if (std.mem.eql(u8, s, "true") or std.mem.eql(u8, s, "1")) return true;
    if (std.mem.eql(u8, s, "false") or std.mem.eql(u8, s, "0")) return false;
    return null;
}

/// `xsd:double`'s decimal lexical space — `(+|-)? (digits (. digits*)?
/// | . digits) ([eE] (+|-)? digits)?` — then the value.
/// `std.fmt.parseFloat`'s wider grammar (hex floats, digit
/// separators, `inf` / `nan` in any case) is not the schema's (the
/// `parseInt` shape of Codex #215 r1 REL-104), and the schema's own
/// `INF` / `-INF` / `NaN` have no JSON spelling, so they read null
/// with an overflow and every other non-finite value.
pub fn xsdDouble(s: []const u8) ?f64 {
    var i: usize = 0;
    if (i < s.len and (s[i] == '+' or s[i] == '-')) i += 1;
    var int_digits: usize = 0;
    while (i < s.len and s[i] >= '0' and s[i] <= '9') : (i += 1) int_digits += 1;
    var frac_digits: usize = 0;
    if (i < s.len and s[i] == '.') {
        i += 1;
        while (i < s.len and s[i] >= '0' and s[i] <= '9') : (i += 1) frac_digits += 1;
    }
    if (int_digits == 0 and frac_digits == 0) return null;
    if (i < s.len and (s[i] == 'e' or s[i] == 'E')) {
        i += 1;
        if (i < s.len and (s[i] == '+' or s[i] == '-')) i += 1;
        var exp_digits: usize = 0;
        while (i < s.len and s[i] >= '0' and s[i] <= '9') : (i += 1) exp_digits += 1;
        if (exp_digits == 0) return null;
    }
    if (i != s.len) return null;
    const v = std.fmt.parseFloat(f64, s) catch return null;
    if (!std.math.isFinite(v)) return null;
    return v;
}

// ─── Test fixture ────────────────────────────────────────────────────

/// Writes a real three-sheet workbook through the public Writer, then
/// splices what the fresh writer cannot author: `Data` (index 0)
/// carries the writer's own frozen pane (`freezePanes(1, 2)`) plus a
/// spliced `<dimension>`; `Report` (index 1) a spliced `<dimension>`
/// and a SPLIT pane with a fractional `ySplit`; `Bare` (index 2)
/// neither — its record is the two-null shape. `xl/workbook.xml`
/// gains a full `<calcPr>`. `src/cli.zig` and the tests below share
/// it.
pub const fixture = struct {
    pub fn write(allocator: Allocator, io: std.Io, path: []const u8) !void {
        {
            const zlsx = @import("zlsx");
            var w = zlsx.Writer.init(allocator);
            defer w.deinit();
            var data = try w.addSheet("Data");
            try data.writeRow(&.{ .{ .string = "Region" }, .{ .string = "Qty" } });
            try data.writeRow(&.{ .{ .string = "East" }, .{ .integer = 3 } });
            try data.writeRow(&.{ .{ .string = "West" }, .{ .integer = 4 } });
            try data.freezePanes(1, 2);
            var report = try w.addSheet("Report");
            try report.writeRow(&.{ .{ .string = "R&D" }, .{ .integer = 1 }, .{ .integer = 2 } });
            try report.writeRow(&.{ .{ .string = "Ops" }, .{ .integer = 3 }, .{ .integer = 4 } });
            var bare = try w.addSheet("Bare");
            try bare.writeRow(&.{.{ .integer = 1 }});
            try w.save(io, path);
        }
        // The fresh writer emits no `<dimension>` and no `<calcPr>`,
        // and has no split-pane surface; the schema order (dimension,
        // sheetViews, cols, sheetData) is kept at each splice.
        try patchPart(allocator, io, path, "xl/worksheets/sheet1.xml", "<sheetViews>", "<dimension ref=\"A1:B3\"/><sheetViews>");
        try patchPart(allocator, io, path, "xl/worksheets/sheet2.xml", "<sheetData>", "<dimension ref=\"A1:C2\"/><sheetViews><sheetView workbookViewId=\"0\"><pane xSplit=\"2865\" ySplit=\"1215.5\" topLeftCell=\"C4\" activePane=\"bottomRight\" state=\"split\"/></sheetView></sheetViews><sheetData>");
        try patchPart(allocator, io, path, "xl/workbook.xml", "</workbook>", "<calcPr calcId=\"191029\" fullCalcOnLoad=\"1\" iterate=\"true\" iterateCount=\"100\" iterateDelta=\"0.001\"/></workbook>");
    }

    /// Byte-replace the first `old` in one part of a saved workbook
    /// and save it back — the conditional-formats fixture's helper,
    /// shared rather than duplicated (this module already fronts that
    /// module's lexical layer).
    pub const patchPart = cf.fixture.patchPart;
};

// ─── Tests ───────────────────────────────────────────────────────────

const testing = std.testing;

const TestTmp = struct {
    dir: std.testing.TmpDir,
    fn init() TestTmp {
        return .{ .dir = std.testing.tmpDir(.{}) };
    }
    fn deinit(self: *TestTmp) void {
        self.dir.cleanup();
    }
    fn path(self: *TestTmp, alloc: Allocator, io: std.Io, name: []const u8) ![:0]u8 {
        const d = try self.dir.dir.realPathFileAlloc(io, ".", alloc);
        defer alloc.free(d);
        return std.fs.path.joinZ(alloc, &.{ d, name });
    }
};

const ws_open = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">";

test "writeSheetRecord: full and compact envelopes, exact bytes" {
    const rec: SheetRecord = .{
        .sheet = "R\"D",
        .sheet_idx = 1,
        .dimension = "A1:C2",
        .pane = .{
            .x_split = 2865,
            .y_split = 1215.5,
            .top_left_cell = "C4",
            .active_pane = "bottomRight",
            .state = "split",
        },
    };
    var buf: [512]u8 = undefined;
    var w = std.Io.Writer.fixed(&buf);
    try writeSheetRecord(&w, rec, .full);
    try testing.expectEqualStrings(
        "{\"kind\":\"sheet_props\",\"sheet\":\"R\\\"D\",\"sheet_idx\":1,\"dimension\":\"A1:C2\"," ++
            "\"pane\":{\"x_split\":2865,\"y_split\":1215.5,\"top_left_cell\":\"C4\",\"active_pane\":\"bottomRight\",\"state\":\"split\"}}\n",
        w.buffered(),
    );
    var w2 = std.Io.Writer.fixed(&buf);
    try writeSheetRecord(&w2, rec, .compact);
    try testing.expectEqualStrings(
        "{\"kind\":\"sheet_props\",\"dimension\":\"A1:C2\"," ++
            "\"pane\":{\"x_split\":2865,\"y_split\":1215.5,\"top_left_cell\":\"C4\",\"active_pane\":\"bottomRight\",\"state\":\"split\"}}\n",
        w2.buffered(),
    );
}

test "writeSheetRecord: absent extent and pane, and a pane of nulls" {
    var buf: [512]u8 = undefined;
    var w = std.Io.Writer.fixed(&buf);
    try writeSheetRecord(&w, .{ .sheet = "Bare", .sheet_idx = 2, .dimension = null, .pane = null }, .full);
    try testing.expectEqualStrings(
        "{\"kind\":\"sheet_props\",\"sheet\":\"Bare\",\"sheet_idx\":2,\"dimension\":null,\"pane\":null}\n",
        w.buffered(),
    );
    // `<pane/>` with no attributes: present, every field null — the
    // wire keeps "a pane with nothing said" apart from "no pane".
    var w2 = std.Io.Writer.fixed(&buf);
    try writeSheetRecord(&w2, .{
        .sheet = "S",
        .sheet_idx = 0,
        .dimension = "",
        .pane = .{ .x_split = null, .y_split = null, .top_left_cell = null, .active_pane = null, .state = null },
    }, .compact);
    try testing.expectEqualStrings(
        "{\"kind\":\"sheet_props\",\"dimension\":\"\"," ++
            "\"pane\":{\"x_split\":null,\"y_split\":null,\"top_left_cell\":null,\"active_pane\":null,\"state\":null}}\n",
        w2.buffered(),
    );
}

test "writeCalcRecord: exact bytes, and the absent record" {
    var buf: [512]u8 = undefined;
    var w = std.Io.Writer.fixed(&buf);
    try writeCalcRecord(&w, .{
        .calc_id = 191029,
        .full_calc_on_load = true,
        .iterate = false,
        .iterate_count = 100,
        .iterate_delta = 0.001,
    });
    try testing.expectEqualStrings(
        "{\"kind\":\"calc_props\",\"calc_id\":191029,\"full_calc_on_load\":true,\"iterate\":false,\"iterate_count\":100,\"iterate_delta\":0.001}\n",
        w.buffered(),
    );
    var w2 = std.Io.Writer.fixed(&buf);
    try writeCalcRecord(&w2, CalcRecord.absent);
    try testing.expectEqualStrings(
        "{\"kind\":\"calc_props\",\"calc_id\":null,\"full_calc_on_load\":null,\"iterate\":null,\"iterate_count\":null,\"iterate_delta\":null}\n",
        w2.buffered(),
    );
}

test "xsdDouble: the schema's lexical space, not parseFloat's" {
    const ok = [_]struct { s: []const u8, v: f64 }{
        .{ .s = "0", .v = 0 },
        .{ .s = "2865", .v = 2865 },
        .{ .s = "1215.5", .v = 1215.5 },
        .{ .s = "-1.5", .v = -1.5 },
        .{ .s = "+3", .v = 3 },
        .{ .s = ".5", .v = 0.5 },
        .{ .s = "5.", .v = 5 },
        .{ .s = "1e3", .v = 1000 },
        .{ .s = "1E-3", .v = 0.001 },
        .{ .s = "2.5e+2", .v = 250 },
        .{ .s = "0.001", .v = 0.001 },
    };
    for (ok) |c| try testing.expectEqual(@as(?f64, c.v), xsdDouble(c.s));
    const bad = [_][]const u8{
        "",    "+",   "-",     ".",   "abc", "1_0", "0x10", "1e",
        "1e+", "e5",  "1.5.2", " 1",  "1 ",  "INF", "-INF", "NaN",
        "inf", "nan", "1e400", "1,5",
        "１",
        "0b1", "1f",  "1.5d",
    };
    for (bad) |s| try testing.expectEqual(@as(?f64, null), xsdDouble(s));
}

test "typedAttr: boundary whitespace collapses after the decode; interior whitespace stays invalid (REL-103)" {
    const a = testing.allocator;
    try testing.expectEqual(@as(?f64, 2), try doubleAttr(a, " 2 "));
    try testing.expectEqual(@as(?f64, 0.5), try doubleAttr(a, "&#x20;.5&#x20;"));
    try testing.expectEqual(@as(?f64, 2865), try doubleAttr(a, "\t2865\r\n"));
    try testing.expectEqual(@as(?f64, null), try doubleAttr(a, "1 0"));
    try testing.expectEqual(@as(?f64, null), try doubleAttr(a, " "));
    try testing.expectEqual(@as(?bool, true), try boolAttr(a, " true "));
    try testing.expectEqual(@as(?bool, false), try boolAttr(a, "&#48; "));
    try testing.expectEqual(@as(?bool, null), try boolAttr(a, "tr ue"));
    try testing.expectEqual(@as(?u32, 191029), try uintAttr(a, " 191029 "));
    try testing.expectEqual(@as(?u32, null), try uintAttr(a, "19 1029"));
    try testing.expectEqual(@as(?u32, null), try uintAttr(a, ""));
    try testing.expectError(error.MalformedSheetXml, doubleAttr(a, " &bogus; "));
}

test "xsdBoolean: exactly the four lexicals" {
    try testing.expectEqual(@as(?bool, true), xsdBoolean("true"));
    try testing.expectEqual(@as(?bool, true), xsdBoolean("1"));
    try testing.expectEqual(@as(?bool, false), xsdBoolean("false"));
    try testing.expectEqual(@as(?bool, false), xsdBoolean("0"));
    for ([_][]const u8{ "", "TRUE", "True", "yes", "no", "01", "2", " 1", "1 " }) |s| {
        try testing.expectEqual(@as(?bool, null), xsdBoolean(s));
    }
}

fn scan(a: Allocator, xml: []const u8) Error!RawSheetProps {
    return scanSheetProps(a, xml);
}

test "scanSheetProps: the extent and the first view's pane at their schema slots" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    const r = try scan(a, ws_open ++
        "<sheetPr codeName=\"Sheet1\"/>" ++
        "<dimension ref=\"A1:D9\"/>" ++
        "<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\">" ++
        "<pane xSplit=\"2\" ySplit=\"1\" topLeftCell=\"C2\" activePane=\"bottomRight\" state=\"frozen\"/>" ++
        "<selection pane=\"bottomRight\" activeCell=\"C2\" sqref=\"C2\"/>" ++
        "</sheetView></sheetViews>" ++
        "<sheetData/></worksheet>");
    try testing.expectEqual(cf.SheetRoot.worksheet, r.root);
    try testing.expectEqualStrings("A1:D9", r.dimension.?);
    const p = r.pane.?;
    try testing.expectEqualStrings("2", p.x_split.?);
    try testing.expectEqualStrings("1", p.y_split.?);
    try testing.expectEqualStrings("C2", p.top_left_cell.?);
    try testing.expectEqualStrings("bottomRight", p.active_pane.?);
    try testing.expectEqualStrings("frozen", p.state.?);

    // A sheet spelling neither.
    const none = try scan(a, ws_open ++ "<sheetData/></worksheet>");
    try testing.expect(none.dimension == null);
    try testing.expect(none.pane == null);

    // A `<dimension>` without `ref`, a `<pane>` without attributes:
    // present, nothing said.
    const empty = try scan(a, ws_open ++ "<dimension/><sheetViews><sheetView><pane/></sheetView></sheetViews><sheetData/></worksheet>");
    try testing.expect(empty.dimension == null);
    try testing.expect(empty.pane != null);
    try testing.expect(empty.pane.?.state == null);

    // Whitespace around `=`, either quote, a pane that is not
    // self-closing — XML's grammar, not a byte pattern (REL-102).
    const ws = try scan(a, ws_open ++ "<dimension ref = 'B2:C3' /><sheetViews><sheetView><pane state = 'split' ></pane></sheetView></sheetViews><sheetData/></worksheet>");
    try testing.expectEqualStrings("B2:C3", ws.dimension.?);
    try testing.expectEqualStrings("split", ws.pane.?.state.?);
}

test "scanSheetProps: decoys at other depths are opaque, not the slot" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    // The corpus shape: `<customSheetViews>` carries its own `<pane>`.
    const custom = try scan(a, ws_open ++
        "<sheetData/><customSheetViews><customSheetView guid=\"{1}\"><pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/></customSheetView></customSheetViews></worksheet>");
    try testing.expect(custom.pane == null);

    // A `<dimension>` under `<extLst>`, and one inside a rebound subtree.
    const ext = try scan(a, ws_open ++
        "<sheetData/><extLst><ext uri=\"{x}\"><dimension ref=\"Z9\"/></ext></extLst></worksheet>");
    try testing.expect(ext.dimension == null);
    const rebound = try scan(a, ws_open ++
        "<foo xmlns=\"urn:x\"><dimension ref=\"Z9\"/><sheetViews><sheetView><pane state=\"frozen\"/></sheetView></sheetViews></foo><sheetData/></worksheet>");
    try testing.expect(rebound.dimension == null);
    try testing.expect(rebound.pane == null);

    // A pane directly under the wrapper (no view) is not the slot.
    const wrapper_pane = try scan(a, ws_open ++
        "<sheetViews><pane state=\"frozen\"/></sheetViews><sheetData/></worksheet>");
    try testing.expect(wrapper_pane.pane == null);

    // The SECOND view's pane is not the slot: the record reports the
    // primary window's, which here has none.
    const second_view = try scan(a, ws_open ++
        "<sheetViews><sheetView workbookViewId=\"0\"/><sheetView workbookViewId=\"1\"><pane ySplit=\"1\" state=\"frozen\"/></sheetView></sheetViews><sheetData/></worksheet>");
    try testing.expect(second_view.pane == null);
    // … and a second view's second pane, or its duplicate attribute,
    // is opaque, not a refusal — it is not slot machinery.
    const second_view_attrs = try scan(a, ws_open ++
        "<sheetViews><sheetView workbookViewId=\"0\"/><sheetView workbookViewId=\"1\" workbookViewId=\"2\"/></sheetViews><sheetData/></worksheet>");
    try testing.expect(second_view_attrs.pane == null);
    const second_view_dup = try scan(a, ws_open ++
        "<sheetViews><sheetView workbookViewId=\"0\"><pane state=\"split\"/></sheetView><sheetView workbookViewId=\"1\"><pane/><pane/></sheetView></sheetViews><sheetData/></worksheet>");
    try testing.expectEqualStrings("split", second_view_dup.pane.?.state.?);

    // A commented-out slot is text.
    const commented = try scan(a, ws_open ++
        "<!-- <dimension ref=\"Q1\"/> --><sheetData/><!-- <sheetViews><sheetView><pane state=\"frozen\"/></sheetView></sheetViews> --></worksheet>");
    try testing.expect(commented.dimension == null);
    try testing.expect(commented.pane == null);
}

test "scanSheetProps: sheet-family roots — chartsheet slot-less, dialogsheet pane only, macrosheet both" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();
    const main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    // A chartsheet has no dimension and its views carry no pane: both
    // spellings are opaque there, no refusal.
    const chart = try scan(a, "<chartsheet xmlns=\"" ++ main ++ "\"><dimension ref=\"A1\"/><sheetViews><sheetView zoomScale=\"100\"><pane state=\"frozen\"/></sheetView></sheetViews></chartsheet>");
    try testing.expectEqual(cf.SheetRoot.chartsheet, chart.root);
    try testing.expect(chart.dimension == null);
    try testing.expect(chart.pane == null);

    const dialog = try scan(a, "<dialogsheet xmlns=\"" ++ main ++ "\"><dimension ref=\"A1\"/><sheetViews><sheetView><pane ySplit=\"3\" state=\"frozen\"/></sheetView></sheetViews></dialogsheet>");
    try testing.expectEqual(cf.SheetRoot.dialogsheet, dialog.root);
    try testing.expect(dialog.dimension == null);
    try testing.expectEqualStrings("3", dialog.pane.?.y_split.?);

    const macro = try scan(a, "<macrosheet xmlns=\"" ++ main ++ "\"><dimension ref=\"A1:A2\"/><sheetViews><sheetView><pane xSplit=\"1\"/></sheetView></sheetViews></macrosheet>");
    try testing.expectEqual(cf.SheetRoot.macrosheet, macro.root);
    try testing.expectEqualStrings("A1:A2", macro.dimension.?);
    try testing.expectEqualStrings("1", macro.pane.?.x_split.?);

    // The canonical `<xm:macrosheet>` spelling (REL-2001).
    const xm = try scan(a, "<xm:macrosheet xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\" xmlns=\"" ++ main ++ "\"><dimension ref=\"B2\"/></xm:macrosheet>");
    try testing.expectEqual(cf.SheetRoot.macrosheet, xm.root);
    try testing.expectEqualStrings("B2", xm.dimension.?);

    // Foreign roots refuse.
    try testing.expectError(error.MalformedSheetXml, scan(a, "<worksheet xmlns=\"urn:x\"><dimension ref=\"A1\"/></worksheet>"));
    try testing.expectError(error.MalformedSheetXml, scan(a, "<x:worksheet xmlns:x=\"" ++ main ++ "\"/>"));
    try testing.expectError(error.MalformedSheetXml, scan(a, "<sheet xmlns=\"" ++ main ++ "\"/>"));
}

test "scanSheetProps: malformed structures refuse rather than pick" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();
    const mce = "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"";

    const cases = [_][]const u8{
        // maxOccurs=1 slots spelled twice.
        ws_open ++ "<dimension ref=\"A1\"/><dimension ref=\"B2\"/><sheetData/></worksheet>",
        ws_open ++ "<sheetViews/><sheetViews/><sheetData/></worksheet>",
        ws_open ++ "<sheetViews><sheetView><pane/><pane/></sheetView></sheetViews><sheetData/></worksheet>",
        // Duplicate attributes on the slot machinery.
        ws_open ++ "<dimension ref=\"A1\" ref=\"B2\"/><sheetData/></worksheet>",
        ws_open ++ "<sheetViews><sheetView><pane state=\"frozen\" state=\"split\"/></sheetView></sheetViews><sheetData/></worksheet>",
        ws_open ++ "<sheetViews><sheetView><pane xSplit=\"1\" foo=\"1\" foo=\"2\"/></sheetView></sheetViews><sheetData/></worksheet>",
        ws_open ++ "<sheetViews foo=\"1\" foo=\"2\"><sheetView/></sheetViews><sheetData/></worksheet>",
        ws_open ++ "<sheetViews><sheetView workbookViewId=\"0\" workbookViewId=\"1\"><pane/></sheetView></sheetViews><sheetData/></worksheet>",
        // MCE at the recognized slots (SEC-2201).
        ws_open ++ "<mc:AlternateContent " ++ mce ++ "><mc:Choice><dimension ref=\"A1\"/></mc:Choice></mc:AlternateContent><sheetData/></worksheet>",
        ws_open ++ "<sheetViews><mc:AlternateContent " ++ mce ++ "><mc:Choice><sheetView/></mc:Choice></mc:AlternateContent></sheetViews><sheetData/></worksheet>",
        ws_open ++ "<sheetViews><sheetView><mc:AlternateContent " ++ mce ++ "><mc:Choice><pane/></mc:Choice></mc:AlternateContent></sheetView></sheetViews><sheetData/></worksheet>",
        ws_open ++ "<sheetData mc:ProcessContent=\"x\" " ++ mce ++ "/></worksheet>",
        // The main namespace aliased to a prefix.
        ws_open ++ "<x:dimension xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ref=\"A1\"/><sheetData/></worksheet>",
        // Nesting: mismatched, unterminated, a second root, trailing data.
        ws_open ++ "<sheetViews><sheetView></sheetViews></sheetView><sheetData/></worksheet>",
        ws_open ++ "<sheetData/>",
        ws_open ++ "<sheetData/></worksheet><worksheet/>",
        ws_open ++ "<sheetData/></worksheet>x",
        // A DOCTYPE, a malformed comment, a raw `<` in a value.
        "<!DOCTYPE worksheet [<!ENTITY d \"<dimension ref='A1'/>\">]>" ++ ws_open ++ "&d;<sheetData/></worksheet>",
        ws_open ++ "<!-- bad -- <dimension ref=\"A1\"/> --><sheetData/></worksheet>",
        ws_open ++ "<dimension ref=\"A1<B\"/><sheetData/></worksheet>",
        // Non-UTF-8 anywhere in the part.
        ws_open ++ "<dimension ref=\"A\xff1\"/><sheetData/></worksheet>",
        // An undeclared prefix on the slot machinery (SEC-1101).
        ws_open ++ "<sheetViews><sheetView><pane p:state=\"frozen\"/></sheetView></sheetViews><sheetData/></worksheet>",
    };
    for (cases) |xml| {
        try testing.expectError(error.MalformedSheetXml, scan(a, xml));
    }

    // Inert MCE reads: the declaration, `mc:Ignorable`, and a deep
    // `mc:AlternateContent` inside an opaque subtree.
    const inert = try scan(a, "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " ++ mce ++ " mc:Ignorable=\"x14ac\">" ++
        "<dimension ref=\"A1\"/><sheetData/><oleObjects><mc:AlternateContent><mc:Choice Requires=\"x14\"><oleObject/></mc:Choice></mc:AlternateContent></oleObjects></worksheet>");
    try testing.expectEqualStrings("A1", inert.dimension.?);
}

test "scanSheetProps: nesting past the ceiling refuses; deep-but-bounded walks (PERF-201)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();
    var deep: std.ArrayListUnmanaged(u8) = .empty;
    try deep.appendSlice(a, ws_open);
    for (0..cf.max_depth + 2) |_| try deep.appendSlice(a, "<x>");
    for (0..cf.max_depth + 2) |_| try deep.appendSlice(a, "</x>");
    try deep.appendSlice(a, "</worksheet>");
    try testing.expectError(error.MalformedSheetXml, scan(a, deep.items));

    var bounded: std.ArrayListUnmanaged(u8) = .empty;
    try bounded.appendSlice(a, ws_open);
    for (0..cf.max_depth - 2) |_| try bounded.appendSlice(a, "<x>");
    for (0..cf.max_depth - 2) |_| try bounded.appendSlice(a, "</x>");
    try bounded.appendSlice(a, "<dimension ref=\"A1\"/></worksheet>");
    const r = try scan(a, bounded.items);
    try testing.expectEqualStrings("A1", r.dimension.?);
}

test "collect: fixture — frozen, split and bare sheets, workbook order, decoded values" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "sheet_props_fixture.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);

    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    var view = try collect(testing.allocator, &wb);
    defer view.deinit();

    try testing.expectEqual(@as(usize, 3), view.sheet_names.len);
    try testing.expectEqual(@as(usize, 3), view.records.len);

    const data = view.records[0];
    try testing.expectEqualStrings("Data", data.sheet);
    try testing.expectEqual(@as(u32, 0), data.sheet_idx);
    try testing.expectEqualStrings("A1:B3", data.dimension.?);
    const frozen = data.pane.?;
    try testing.expectEqual(@as(?f64, 2), frozen.x_split);
    try testing.expectEqual(@as(?f64, 1), frozen.y_split);
    try testing.expectEqualStrings("C2", frozen.top_left_cell.?);
    try testing.expectEqualStrings("bottomRight", frozen.active_pane.?);
    try testing.expectEqualStrings("frozen", frozen.state.?);

    const report = view.records[1];
    try testing.expectEqualStrings("Report", report.sheet);
    try testing.expectEqualStrings("A1:C2", report.dimension.?);
    const split = report.pane.?;
    try testing.expectEqual(@as(?f64, 2865), split.x_split);
    try testing.expectEqual(@as(?f64, 1215.5), split.y_split);
    try testing.expectEqualStrings("C4", split.top_left_cell.?);
    try testing.expectEqualStrings("split", split.state.?);

    const bare = view.records[2];
    try testing.expectEqualStrings("Bare", bare.sheet);
    try testing.expectEqual(@as(u32, 2), bare.sheet_idx);
    try testing.expect(bare.dimension == null);
    try testing.expect(bare.pane == null);

    const calc = try collectCalc(testing.allocator, &wb);
    try testing.expectEqual(@as(?u32, 191029), calc.calc_id);
    try testing.expectEqual(@as(?bool, true), calc.full_calc_on_load);
    try testing.expectEqual(@as(?bool, true), calc.iterate);
    try testing.expectEqual(@as(?u32, 100), calc.iterate_count);
    try testing.expectEqual(@as(?f64, 0.001), calc.iterate_delta);
}

test "collect: typed attributes decode entities first, then type lexically (REL-104 / REL-1802)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "sheet_props_typed.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);
    const sheet2 = "xl/worksheets/sheet2.xml";
    // Entity-spelled digits are the number; a schema-invalid lexical
    // is null; a fractional split is a double; the carriers decode.
    try fixture.patchPart(testing.allocator, io, path, sheet2, "xSplit=\"2865\"", "xSplit=\"&#50;865\"");
    try fixture.patchPart(testing.allocator, io, path, sheet2, "ySplit=\"1215.5\"", "ySplit=\"1_0\"");
    try fixture.patchPart(testing.allocator, io, path, sheet2, "topLeftCell=\"C4\"", "topLeftCell=\"C&#52;\"");
    try fixture.patchPart(testing.allocator, io, path, sheet2, "ref=\"A1:C2\"", "ref=\"A1:C&amp;2\"");
    // The workbook side: an entity-spelled or non-numeric `calcId` /
    // `iterateCount` / `iterateDelta` never reaches this read through
    // any opener — `Workbook.open`'s lenient parse refuses it first —
    // so the decode-first rule is exercised on a boolean carrier, and
    // the written-but-invalid rule on a boolean lexical.
    try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "iterate=\"true\"", "iterate=\"yes\"");
    try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "fullCalcOnLoad=\"1\"", "fullCalcOnLoad=\"&#48;\"");

    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    var view = try collect(testing.allocator, &wb);
    defer view.deinit();
    const report = view.records[1];
    try testing.expectEqualStrings("A1:C&2", report.dimension.?);
    try testing.expectEqual(@as(?f64, 2865), report.pane.?.x_split);
    try testing.expectEqual(@as(?f64, null), report.pane.?.y_split);
    try testing.expectEqualStrings("C4", report.pane.?.top_left_cell.?);

    const calc = try collectCalc(testing.allocator, &wb);
    try testing.expectEqual(@as(?u32, 191029), calc.calc_id);
    try testing.expectEqual(@as(?bool, false), calc.full_calc_on_load);
    try testing.expectEqual(@as(?bool, null), calc.iterate);
    try testing.expectEqual(@as(?u32, 100), calc.iterate_count);
    try testing.expectEqual(@as(?f64, 0.001), calc.iterate_delta);
}

test "calcFromAttrs: the numeric carriers decode entities first and type lexically (REL-1802)" {
    // Reachable only below the opener today (see above); the record
    // writer's contract is the family's regardless.
    const a = testing.allocator;
    const r = try calcFromAttrs(a, " calcId=\"&#49;91029\" iterateCount=\"1_0\" iterateDelta=\"1e-5\" iterate=\"&#49;\"");
    try testing.expectEqual(@as(?u32, 191029), r.calc_id);
    try testing.expectEqual(@as(?u32, null), r.iterate_count);
    try testing.expectEqual(@as(?f64, 1e-5), r.iterate_delta);
    try testing.expectEqual(@as(?bool, true), r.iterate);
    try testing.expectEqual(@as(?bool, null), r.full_calc_on_load);
    try testing.expectError(error.MalformedSheetXml, calcFromAttrs(a, " calcId=\"&bogus;\""));
    try testing.expectError(error.MalformedSheetXml, calcFromAttrs(a, " iterateDelta=\"&bogus;\""));
    try testing.expectError(error.MalformedSheetXml, calcFromAttrs(a, " calcId=\"1\" calcId=\"2\""));
    try testing.expectError(error.MalformedSheetXml, calcFromAttrs(a, " calcId=\"1\"x=\"2\""));
}

test "collect: whitespace-padded typed attributes read as their values on the wire (REL-103)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "sheet_props_ws.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);
    try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet2.xml", "xSplit=\"2865\"", "xSplit=\" 2865 \"");
    try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet2.xml", "ySplit=\"1215.5\"", "ySplit=\"&#x20;1215.5&#x9;\"");
    try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "iterate=\"true\"", "iterate=\"&#x20;true \"");
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    var view = try collect(testing.allocator, &wb);
    defer view.deinit();
    var buf: [512]u8 = undefined;
    var w = std.Io.Writer.fixed(&buf);
    try writeSheetRecord(&w, view.records[1], .compact);
    try testing.expectEqualStrings(
        "{\"kind\":\"sheet_props\",\"dimension\":\"A1:C2\"," ++
            "\"pane\":{\"x_split\":2865,\"y_split\":1215.5,\"top_left_cell\":\"C4\",\"active_pane\":\"bottomRight\",\"state\":\"split\"}}\n",
        w.buffered(),
    );
    const calc = try collectCalc(testing.allocator, &wb);
    try testing.expectEqual(@as(?bool, true), calc.iterate);
}

test "collect: a carrier the stream cannot carry faithfully refuses whole" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();

    const sheet2 = "xl/worksheets/sheet2.xml";
    const cases = [_]struct { name: []const u8, old: []const u8, new: []const u8 }{
        // An extent carrier that does not decode.
        .{ .name = "r1.xlsx", .old = "ref=\"A1:C2\"", .new = "ref=\"A1&bogus;C2\"" },
        // A pane carrier that does not decode.
        .{ .name = "r2.xlsx", .old = "topLeftCell=\"C4\"", .new = "topLeftCell=\"C&bogus;\"" },
        // A typed carrier that does not decode refuses — only a
        // DECODED non-lexical reads null (MNT-2301).
        .{ .name = "r3.xlsx", .old = "xSplit=\"2865\"", .new = "xSplit=\"&bogus;\"" },
        // Two extents, two wrappers, two first-view panes.
        .{ .name = "r4.xlsx", .old = "<dimension ref=\"A1:C2\"/>", .new = "<dimension ref=\"A1:C2\"/><dimension ref=\"A1\"/>" },
        .{ .name = "r5.xlsx", .old = "</sheetViews>", .new = "</sheetViews><sheetViews/>" },
        .{ .name = "r6.xlsx", .old = "state=\"split\"/>", .new = "state=\"split\"/><pane/>" },
        // A duplicate attribute on the pane.
        .{ .name = "r7.xlsx", .old = "state=\"split\"", .new = "state=\"split\" state=\"frozen\"" },
        // An MCE branch at the views slot.
        .{ .name = "r8.xlsx", .old = "<sheetViews>", .new = "<sheetViews><mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice Requires=\"x14\"/></mc:AlternateContent>" },
        // A truncated part.
        .{ .name = "r9.xlsx", .old = "</worksheet>", .new = "" },
    };
    for (cases) |case| {
        const path = try tt.path(testing.allocator, io, case.name);
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path);
        try fixture.patchPart(testing.allocator, io, path, sheet2, case.old, case.new);
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        try testing.expectError(error.MalformedSheetXml, collect(testing.allocator, &wb));
    }
}

test "collect: a worksheet part shares the typed view's parse verdict (REL-404)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "sheet_props_gate.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);
    // A balanced `<sheetDataQ>` wrapper is an opaque, accepted frame
    // to the strict walk; the typed view's lexical scan matches it as
    // `<sheetData` and refuses the part — the family refuses.
    try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet3.xml", "<sheetData>", "<sheetDataQ>");
    try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet3.xml", "</sheetData>", "</sheetDataQ>");
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    try testing.expectError(error.MalformedSheetXml, collect(testing.allocator, &wb));
}

test "collect: a sheet list the read cannot attribute against refuses whole" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "sheet_props_bad_sheet.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);
    try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "name=\"Report\"", "name=\"Rep&bogus;\"");
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    try testing.expectError(error.MalformedWorkbookXml, collect(testing.allocator, &wb));
    // The calc read runs the same strict workbook walk — a `<sheets>`
    // list it cannot prove is the same verdict there.
    try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "name=\"Rep&bogus;\"", "name=\"Report\"");
    try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "</sheets>", "</sheets><sheets/>");
    var wb2 = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb2.deinit();
    try testing.expectError(error.MalformedWorkbookXml, collectCalc(testing.allocator, &wb2));
}

test "collect: the corpus shapes — a customSheetViews pane and a missing extent — read as null" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "sheet_props_corpus_shapes.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);
    try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet3.xml", "</sheetData>", "</sheetData><customSheetViews><customSheetView guid=\"{A}\"><pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/><selection pane=\"bottomLeft\"/></customSheetView></customSheetViews>");
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    var view = try collect(testing.allocator, &wb);
    defer view.deinit();
    try testing.expect(view.records[2].dimension == null);
    try testing.expect(view.records[2].pane == null);
}

test "collectCalc: absent, empty, whitespace-closed, decoyed and doubled slots" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const full = "<calcPr calcId=\"191029\" fullCalcOnLoad=\"1\" iterate=\"true\" iterateCount=\"100\" iterateDelta=\"0.001\"/>";

    const Case = struct { name: []const u8, new: []const u8, want: ?CalcRecord };
    const cases = [_]Case{
        // No element at all: the record of nulls.
        .{ .name = "c1.xlsx", .new = "", .want = CalcRecord.absent },
        // An empty element: the same record — absent and empty are
        // one shape on the wire.
        .{ .name = "c2.xlsx", .new = "<calcPr/>", .want = CalcRecord.absent },
        // The corpus's `<calcPr calcId="40001" />` spelling.
        .{ .name = "c3.xlsx", .new = "<calcPr calcId=\"40001\" />", .want = .{ .calc_id = 40001, .full_calc_on_load = null, .iterate = null, .iterate_count = null, .iterate_delta = null } },
        // A decoy under `<extLst>` is not the slot.
        .{ .name = "c4.xlsx", .new = "<extLst><ext uri=\"{x}\"><calcPr calcId=\"7\"/></ext></extLst>", .want = CalcRecord.absent },
        // Two at the slot refuse.
        .{ .name = "c5.xlsx", .new = full ++ "<calcPr calcId=\"1\"/>", .want = null },
        // One an MCE branch could project into the slot refuses; the
        // real root-level `mc:AlternateContent` (absPath) stays inert.
        .{ .name = "c6.xlsx", .new = "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice Requires=\"x15\"><calcPr calcId=\"1\"/></mc:Choice></mc:AlternateContent>", .want = null },
        .{ .name = "c7.xlsx", .new = "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice Requires=\"x15\"><x15ac:absPath xmlns:x15ac=\"urn:x\" url=\"C:\\\"/></mc:Choice></mc:AlternateContent>" ++ full, .want = .{ .calc_id = 191029, .full_calc_on_load = true, .iterate = true, .iterate_count = 100, .iterate_delta = 0.001 } },
        // A duplicate attribute refuses; an undecodable carrier refuses.
        .{ .name = "c8.xlsx", .new = "<calcPr calcId=\"1\" calcId=\"2\"/>", .want = null },
        .{ .name = "c9.xlsx", .new = "<calcPr iterate=\"&bogus;\"/>", .want = null },
        // A `<calcPr>` under a rebound default namespace is foreign.
        .{ .name = "c10.xlsx", .new = "<foo xmlns=\"urn:x\"><calcPr calcId=\"1\"/></foo>", .want = CalcRecord.absent },
        // An MCE branch BELOW an ordinary wrapper cannot project into
        // the root slot: neither counted nor shadowing (Codex #218 r1
        // S3B-REL-102) — and the real root record beside it survives.
        .{ .name = "c11.xlsx", .new = "<extLst><ext uri=\"{x}\"><mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice Requires=\"x15\"><calcPr calcId=\"7\"/></mc:Choice></mc:AlternateContent></ext></extLst>", .want = CalcRecord.absent },
        .{ .name = "c12.xlsx", .new = full ++ "<extLst><ext uri=\"{x}\"><mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice Requires=\"x15\"><calcPr calcId=\"7\"/></mc:Choice></mc:AlternateContent></ext></extLst>", .want = .{ .calc_id = 191029, .full_calc_on_load = true, .iterate = true, .iterate_count = 100, .iterate_delta = 0.001 } },
        // A root MCE branch whose choice is an ordinary wrapper: after
        // projection the wrapper sits at the root, its `calcPr` below
        // it — not the slot.
        .{ .name = "c13.xlsx", .new = "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice Requires=\"x15\"><foo><calcPr calcId=\"7\"/></foo></mc:Choice></mc:AlternateContent>", .want = CalcRecord.absent },
        // … while a Choice-then-Fallback chain straight from the root
        // still shadows.
        .{ .name = "c14.xlsx", .new = "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice Requires=\"x15\"/><mc:Fallback><calcPr calcId=\"7\"/></mc:Fallback></mc:AlternateContent>", .want = null },
    };
    for (cases) |case| {
        const path = try tt.path(testing.allocator, io, case.name);
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path);
        try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", full, case.new);
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        if (case.want) |want| {
            const got = try collectCalc(testing.allocator, &wb);
            try testing.expectEqual(want, got);
        } else {
            try testing.expectError(error.MalformedWorkbookXml, collectCalc(testing.allocator, &wb));
        }
    }
}

test "collect: repeated reads leave the store's resident bytes flat (S3B-MEM-603)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "sheet_props_retention.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);

    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    {
        var primed = try collect(testing.allocator, &wb);
        primed.deinit();
        _ = try collectCalc(testing.allocator, &wb);
    }
    const before = wb.store.residentBytes();
    for (0..1024) |_| {
        var view = try collect(testing.allocator, &wb);
        view.deinit();
        _ = try collectCalc(testing.allocator, &wb);
    }
    try testing.expectEqual(before, wb.store.residentBytes());
}

test "collect: allocation failures surface as OutOfMemory, never a partial view" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "sheet_props_alloc_sweep.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);
    // An entity-spelled split so the typed decode allocates too.
    try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet2.xml", "xSplit=\"2865\"", "xSplit=\"&#50;865\"");
    try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "fullCalcOnLoad=\"1\"", "fullCalcOnLoad=\"&#49;\"");

    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    // Prime the workbook-owned caches (part bytes, rels) so the sweep
    // exercises the collectors' own allocations.
    {
        var primed = try collect(testing.allocator, &wb);
        primed.deinit();
        _ = try collectCalc(testing.allocator, &wb);
    }
    try std.testing.checkAllAllocationFailures(testing.allocator, struct {
        fn run(failing: Allocator, wb_: *workbook_mod.Workbook) !void {
            var view = try collect(failing, wb_);
            view.deinit();
        }
    }.run, .{&wb});
    try std.testing.checkAllAllocationFailures(testing.allocator, struct {
        fn run(failing: Allocator, wb_: *workbook_mod.Workbook) !void {
            _ = try collectCalc(failing, wb_);
        }
    }.run, .{&wb});
}

test "writeAll: the C handoff writes the full stream byte-exactly" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "sheet_props_write_all.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    var view = try collect(testing.allocator, &wb);
    defer view.deinit();
    var buf: [1024]u8 = undefined;
    var w = std.Io.Writer.fixed(&buf);
    try writeAll(&w, &view);
    try testing.expectEqualStrings(
        "{\"kind\":\"sheet_props\",\"sheet\":\"Data\",\"sheet_idx\":0,\"dimension\":\"A1:B3\"," ++
            "\"pane\":{\"x_split\":2,\"y_split\":1,\"top_left_cell\":\"C2\",\"active_pane\":\"bottomRight\",\"state\":\"frozen\"}}\n" ++
            "{\"kind\":\"sheet_props\",\"sheet\":\"Report\",\"sheet_idx\":1,\"dimension\":\"A1:C2\"," ++
            "\"pane\":{\"x_split\":2865,\"y_split\":1215.5,\"top_left_cell\":\"C4\",\"active_pane\":\"bottomRight\",\"state\":\"split\"}}\n" ++
            "{\"kind\":\"sheet_props\",\"sheet\":\"Bare\",\"sheet_idx\":2,\"dimension\":null,\"pane\":null}\n",
        w.buffered(),
    );
}

fn fuzzScanTarget(_: void, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();
    var smith_buf: [4096]u8 = undefined;
    const input = smith_buf[0..smith.slice(&smith_buf)];
    var arena = std.heap.ArenaAllocator.init(std.testing.allocator);
    defer arena.deinit();
    _ = scanSheetProps(arena.allocator(), input) catch {};
}

/// The shapes the slot rules and the review rounds turn on, plus the
/// truncations a hand-written test would not think to write.
const scan_fuzz_corpus = [_][]const u8{
    "",
    "<",
    ws_open,
    ws_open ++ "<dimension ref=\"A1:D9\"/><sheetViews><sheetView><pane xSplit=\"2\" ySplit=\"1\" topLeftCell=\"C2\" activePane=\"bottomRight\" state=\"frozen\"/></sheetView></sheetViews><sheetData/></worksheet>",
    ws_open ++ "<dimension ref=\"A1\"/><dimension ref=\"B2\"/></worksheet>",
    ws_open ++ "<sheetViews><sheetView><pane/><pane/></sheetView></sheetViews></worksheet>",
    ws_open ++ "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice><dimension ref=\"A1\"/></mc:Choice></mc:AlternateContent></worksheet>",
    ws_open ++ "<sheetViews><sheetView><pane state=\"frozen\" state=\"split\"/></sheetView></sheetViews></worksheet>",
    ws_open ++ "<!-- <dimension ref=\"Q1\"/> --><sheetData/></worksheet>",
    "<chartsheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetViews><sheetView><pane/></sheetView></sheetViews></chartsheet>",
    "<xm:macrosheet xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><dimension ref=\"B2\"/></xm:macrosheet>",
    "<?xml version=\"1.0\"?>" ++ ws_open ++ "<dimension ref='A1' /></worksheet>",
    "\xEF\xBB\xBF" ++ ws_open ++ "<sheetData/></worksheet>",
    "<!DOCTYPE x>" ++ ws_open ++ "</worksheet>",
    ws_open ++ "<sheetData><![CDATA[<dimension ref=\"A1\"/>]]></sheetData></worksheet>",
};

test "fuzz: the strict sheet-props scanner never crashes on adversarial XML" {
    try std.testing.fuzz({}, fuzzScanTarget, .{ .corpus = &scan_fuzz_corpus });
}
