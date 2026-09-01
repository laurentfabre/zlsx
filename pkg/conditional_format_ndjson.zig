//! The `conditional-formats` NDJSON records — S3b's typed
//! conditional-format read (`docs/cli.md`, "conditional-formats"),
//! written once for every surface.
//!
//! `zlsx conditional-formats` emits these through its own selection
//! and pagination; the C and Python legs of row S3b hand over the same
//! bytes when they land — the `pivot_ndjson.zig` /
//! `defined_name_ndjson.zig` / `anchor_ndjson.zig` precedent.
//!
//! One record per `<cfRule>`, sheets in workbook order, rules in
//! sheet-document order, each carrying its parent
//! `<conditionalFormatting>` block's `sqref`. The record reports the
//! rule envelope — where it applies, what kind it is, its formula
//! bodies, its differential style id and priority — not the visual
//! payload (`<colorScale>` / `<dataBar>` / `<iconSet>` children stay
//! in the part, byte-preserved, for callers that need them raw).
//!
//! The read is STRICT where the Zig view is lenient — the `anchors`
//! precedent (`Worksheet.conditionalFormats` keeps its historical
//! lexical contract for the rewriter; this module walks each sheet
//! part with the namespace- and depth-aware scanner so the wire can
//! prove its inventory whole). A `conditionalFormatting` element is a
//! rule block only as a main-namespace direct child of the sheet
//! root, a `cfRule` only as a direct child of a block, a `formula`
//! only as a text-only direct child of a rule — so an extension tree
//! that spells the same names (an `x14` subtree rebinding the default
//! namespace, a decoy under `<extLst>`) can never ghost a record, and
//! a `<formula>` that never closes is a refusal, not an absence
//! (Codex #215 r1 REL-101).
//!
//! Decode discipline: sheet names by their string carrier (entities +
//! ST_Xstring); `sqref`, the rule type and the formulas as
//! entities-only carriers (no ST_Xstring layer on any, the C1
//! ruling). A carrier that does not decode, decodes to non-UTF-8, or
//! carries embedded markup refuses the whole read — a partial or
//! wrong rule inventory is the shape of a guard hole, as the pivot,
//! defined-name and anchor reads established.

const std = @import("std");
const formula_mod = @import("zlsx_formula");
const workbook_mod = @import("workbook.zig");
const store_mod = @import("store.zig");
const drawings = @import("drawings.zig");
const workbook_xml = @import("typed_parts/root.zig").workbook_xml;
const json = @import("json_text.zig");

const Allocator = std.mem.Allocator;
const PartStore = store_mod.PartStore;

/// Whether a record carries the `sheet` / `sheet_idx` envelope.
/// `compact` is the CLI's `--output compact-ndjson`, where a sheet
/// prologue record names the sheet once for every record after it.
pub const Envelope = enum { full, compact };

/// One `<cfRule>`, attributed and decoded. `sqref` is the parent
/// `<conditionalFormatting>` block's target list as authored
/// (space-separated A1 areas); `formulas` holds the rule's up-to-three
/// `<formula>` bodies in document order (a `cellIs` `between` carries
/// two). `sqref` and `rule_type` are the empty string when the source
/// omits the attribute — absent and empty are one shape on the wire,
/// the documented boundary convention.
pub const Record = struct {
    sheet: []const u8,
    sheet_idx: u32,
    sqref: []const u8,
    rule_type: []const u8,
    formulas: []const []const u8,
    dxf_id: ?u32,
    priority: ?u32,
};

/// Every conditional-format rule of a workbook, attributed and in
/// emission order: sheets in workbook order, rules in sheet-document
/// order. Owns its decoded strings; `deinit` frees them.
pub const ConditionalFormats = struct {
    arena: std.heap.ArenaAllocator,
    /// Decoded sheet names, parallel to `WorkbookXml.sheets` — the
    /// inventory the CLI's selectors and the `sheet` field read from.
    sheet_names: []const []const u8,
    records: []const Record,

    pub fn deinit(self: *ConditionalFormats) void {
        self.arena.deinit();
        self.* = undefined;
    }
};

pub const Error = error{
    /// A sheet-name carrier that does not decode or decodes to
    /// non-UTF-8 — the NDJSON must stay parseable.
    MalformedWorkbookXml,
    /// A sheet part the strict walk cannot prove a rule inventory
    /// for (mismatched nesting, a root that does not bind the main
    /// namespace as its default, the main namespace aliased to a
    /// prefix, a `<formula>` that never closes or carries markup, a
    /// rule with more than the schema's three formulas, a duplicate
    /// attribute on the rule machinery) — or a rule field the NDJSON
    /// cannot carry faithfully (a carrier that does not decode, or
    /// one that is not UTF-8).
    MalformedSheetXml,
    OutOfMemory,
};

/// Collect every rule of every workbook sheet. A read this module
/// cannot serve faithfully refuses whole — even for a sheet the CLI's
/// selection would exclude: the inventory is proven whole before
/// selection and pagination apply, the `anchors` rule. A sheet the
/// workbook itself cannot place (a dangling relationship, a part the
/// archive does not hold) passes through as `Workbook`'s own error.
pub fn collect(gpa: Allocator, wb: *workbook_mod.Workbook) !ConditionalFormats {
    var arena = std.heap.ArenaAllocator.init(gpa);
    errdefer arena.deinit();
    const a = arena.allocator();

    const sheet_names = try a.alloc([]const u8, wb.workbook.sheets.len);
    for (wb.workbook.sheets, 0..) |s, i| sheet_names[i] = try decodeSheetName(a, s.name);

    // Resolve each workbook sheet to its part STRICTLY — the anchors
    // collector's rule, not `Worksheet.resolvePartName`'s first-match
    // fast path: the relationship must exist under the sheet's `r:id`,
    // be a sheet-family type, be internal, and reach a part the
    // archive holds, and no two `<sheet>` entries may reach one part
    // (the same rule would ride out twice under two identities) —
    // Codex #215 r4 REL-402.
    const wb_rels = wb.store.rels("xl/workbook.xml");
    const sheet_parts = try a.alloc([]const u8, wb.workbook.sheets.len);
    for (wb.workbook.sheets, 0..) |s, i| {
        const rel = (try relById(a, wb_rels, s.r_id)) orelse return error.MalformedWorkbookXml;
        var typed = false;
        for (sheet_rel_leaves) |leaf| typed = typed or drawings.relTypeIs(rel.type, leaf);
        if (!typed) return error.MalformedWorkbookXml;
        if (rel.target_mode == .external) return error.MalformedWorkbookXml;
        const name = (try wb.store.resolve("xl/workbook.xml", rel.target)) orelse
            return error.MalformedWorkbookXml;
        if ((try wb.store.part(name)) == null) return error.MalformedWorkbookXml;
        for (sheet_parts[0..i]) |prev| {
            if (std.mem.eql(u8, prev, name)) return error.MalformedWorkbookXml;
        }
        sheet_parts[i] = name;
    }

    var records: std.ArrayListUnmanaged(Record) = .empty;
    for (0..wb.workbook.sheets.len) |idx| {
        // Store failures fold like `Worksheet.ensureParsed`: memory
        // and the archive-wide budget keep their names, everything
        // else is "this sheet is not readable".
        const part = (wb.store.part(sheet_parts[idx]) catch |e| switch (e) {
            error.OutOfMemory, error.ZipBombSuspected => return e,
            else => return error.MalformedSheetXml,
        }) orelse return error.MissingSheetPart;

        // Scratch (the walker's frame stack, the raw-rule list) lives
        // on the caller's gpa and is reclaimed per sheet; only decoded
        // records land on the view arena, whose allocations are never
        // individually freed (Codex #215 r2 PERF-201).
        var raw_rules: std.ArrayListUnmanaged(RawRule) = .empty;
        defer raw_rules.deinit(gpa);
        const root = try scanSheetRules(gpa, part.bytes, &raw_rules);

        // The typed sheet view's parse verdict stays part of the
        // family contract for worksheet parts ("the same open-time
        // verdicts every `Worksheet` reader shares"): a part the view
        // refuses — a `<sheetDataQ>` decoy its lexical scan trips on —
        // refuses here too (Codex #215 r4 REL-404). The other roots
        // have no typed view; the strict walk's verdict stands alone.
        if (root == .worksheet) {
            const ws = try wb.sheet(@intCast(idx));
            _ = try ws.ensureParsed();
        }

        // One decode per block: every rule of a `<conditionalFormatting>`
        // shares the parent's `sqref` SLICE, so decoding it per rule
        // amplified a large shared sqref by the rule count (Codex #215
        // r4 PERF-401).
        var last_sqref_raw: ?[]const u8 = null;
        var last_sqref_dec: []const u8 = "";
        for (raw_rules.items) |raw| {
            if (last_sqref_raw == null or
                raw.sqref.ptr != last_sqref_raw.?.ptr or
                raw.sqref.len != last_sqref_raw.?.len)
            {
                last_sqref_dec = try decodeRuleText(a, raw.sqref);
                last_sqref_raw = raw.sqref;
            }
            const formulas = try a.alloc([]const u8, raw.formula_count);
            for (formulas, 0..) |*slot, fi| slot.* = try decodeRuleText(a, raw.formulas[fi]);
            try records.append(a, .{
                .sheet = sheet_names[idx],
                .sheet_idx = @intCast(idx),
                .sqref = last_sqref_dec,
                .rule_type = try decodeRuleText(a, raw.rule_type),
                .formulas = formulas,
                .dxf_id = raw.dxf_id,
                .priority = raw.priority,
            });
        }
    }

    return .{
        .arena = arena,
        .sheet_names = sheet_names,
        .records = try records.toOwnedSlice(a),
    };
}

/// One `{"kind":"conditional_format",…}` line. The field order is the
/// docs/cli.md contract; a change here is a wire-format change on
/// every surface at once.
pub fn writeRecord(out: *std.Io.Writer, r: Record, envelope: Envelope) !void {
    try out.writeAll("{\"kind\":\"conditional_format\"");
    if (envelope == .full) {
        try out.writeAll(",\"sheet\":");
        try json.writeString(out, r.sheet);
        try out.print(",\"sheet_idx\":{d}", .{r.sheet_idx});
    }
    try out.writeAll(",\"sqref\":");
    try json.writeString(out, r.sqref);
    try out.writeAll(",\"rule_type\":");
    try json.writeString(out, r.rule_type);
    try out.writeAll(",\"formulas\":[");
    for (r.formulas, 0..) |f, i| {
        if (i > 0) try out.writeByte(',');
        try json.writeString(out, f);
    }
    try out.writeAll("],\"dxf_id\":");
    try json.writeOptU32(out, r.dxf_id);
    try out.writeAll(",\"priority\":");
    try json.writeOptU32(out, r.priority);
    try out.writeAll("}\n");
}

/// The unselected stream — every record, emission order. The future C
/// leg's entry point.
pub fn writeAll(out: *std.Io.Writer, view: *const ConditionalFormats) !void {
    for (view.records) |r| try writeRecord(out, r, .full);
}

// ─── The strict per-sheet walk ───────────────────────────────────────

/// One rule as sliced out of the part, undecoded. `formulas[0..
/// formula_count]` borrow the part's bytes.
const RawRule = struct {
    sqref: []const u8,
    rule_type: []const u8,
    formulas: [3][]const u8,
    formula_count: u8,
    dxf_id: ?u32,
    priority: ?u32,
};

/// What the part's root turned out to be — the collect layer runs the
/// typed sheet view's parse gate for `worksheet` roots only (the other
/// kinds have no typed view; Codex #215 r4 REL-404).
const RootKind = enum { worksheet, macrosheet, barren };

const main_ns_transitional = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const main_ns_strict = "http://purl.oclc.org/ooxml/spreadsheetml/main";

fn isMainNs(uri: []const u8) bool {
    return std.mem.eql(u8, uri, main_ns_transitional) or
        std.mem.eql(u8, uri, main_ns_strict);
}

/// A namespace declaration binds by its DECODED value — XML resolves
/// character references in attribute values, so
/// `xmlns:x="…spreadsheetml&#47;2006/main"` IS a main-namespace alias
/// and a raw-byte compare read it as foreign, letting a prefixed rule
/// subtree traverse as opaque (Codex #215 r3 SEC-301; the S7b-5 host
/// predicate learned the same in Codex #206 r19 SEC-1901). The decode
/// is the STRICT one: the scalar decoder passes an unknown named
/// entity's `&` through verbatim, which read
/// `<conditionalFormatting xmlns="&bogus;">` as bound-to-foreign and
/// hid its rules — a binding that does not decode is not one the walk
/// can rule out (r4 SEC-401).
fn bindsMainNs(scratch: Allocator, raw: []const u8) Error!bool {
    if (std.mem.indexOfScalar(u8, raw, '&') == null) return isMainNs(raw);
    const decoded = formula_mod.decode.decodeEntities(scratch, raw) catch |e| switch (e) {
        error.OutOfMemory => return error.OutOfMemory,
        else => return error.MalformedSheetXml,
    };
    defer scratch.free(decoded);
    return isMainNs(decoded);
}

/// Walk one sheet part and collect its rules — the strict read the
/// module doc-comment describes. The tree is tokenized whole (the
/// workbook scanner's comment / CDATA / PI / DOCTYPE skipping, quote-
/// aware tag ends), element namespaces are tracked through default
/// declarations, and classification is by depth: a rule block is a
/// main-namespace `conditionalFormatting` that is a DIRECT child of
/// the sheet root, a rule a direct `cfRule` child of a block, a
/// formula a text-only direct `formula` child of a rule — everything
/// else is an opaque subtree the walk steps over. Refusals: a root
/// that is not a main-default `worksheet` / `macrosheet` (the two
/// sheet-family roots that carry conditional formatting; `chartsheet`
/// and `dialogsheet` cannot and return an empty inventory), the main
/// namespace bound to a prefix anywhere, mismatched or unterminated
/// nesting, a duplicate attribute on the rule machinery, markup where
/// the schema puts formula text (a CDATA section included — markup is
/// not the formula, the defined-names ruling), and a rule spelling
/// more than the schema's three formulas.
/// The walker's nesting ceiling. Excel writes sheet parts a handful
/// of levels deep; 1024 is two orders of magnitude of headroom, and a
/// bound keeps a crafted part from growing the frame stack without
/// limit (Codex #215 r2 PERF-201).
const max_depth = 1024;

fn scanSheetRules(a: Allocator, xml: []const u8, out: *std.ArrayListUnmanaged(RawRule)) Error!RootKind {
    const Kind = enum { root, barren_root, cf_block, cf_rule, other };
    const Frame = struct {
        name: []const u8,
        kind: Kind,
        main_default: bool,
        sqref: []const u8,
        rule: RawRule,
    };
    var frames: std.ArrayListUnmanaged(Frame) = .empty;
    defer frames.deinit(a);

    var root_seen = false;
    var root_closed = false;
    var root_kind: RootKind = .worksheet;
    var i: usize = 0;
    while (std.mem.indexOfScalarPos(u8, xml, i, '<')) |lt| {
        // A DOCTYPE (or any other `<!…>` declaration that is not a
        // comment or CDATA section) can define general entities whose
        // expansion this byte walk cannot see — `&cf;` standing for a
        // whole rule block would read as an empty success. Refuse the
        // declaration rather than skip it (Codex #215 r3 SEC-302);
        // Excel writes none.
        if (lt + 1 < xml.len and xml[lt + 1] == '!') {
            const is_comment = lt + 4 <= xml.len and std.mem.eql(u8, xml[lt .. lt + 4], "<!--");
            const is_cdata = lt + 9 <= xml.len and std.mem.eql(u8, xml[lt .. lt + 9], "<![CDATA[");
            if (!is_comment and !is_cdata) return error.MalformedSheetXml;
        }
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
            while (j < xml.len and xml[j] != '>' and !isXmlWs(xml[j])) j += 1;
            const name = xml[name_start..j];
            while (j < xml.len and isXmlWs(xml[j])) j += 1;
            if (j >= xml.len or xml[j] != '>') return error.MalformedSheetXml;
            if (frames.items.len == 0) return error.MalformedSheetXml;
            const top = frames.items[frames.items.len - 1];
            frames.items.len -= 1;
            if (!std.mem.eql(u8, top.name, name)) return error.MalformedSheetXml;
            switch (top.kind) {
                .cf_rule => try out.append(a, top.rule),
                .root, .barren_root => root_closed = true,
                else => {},
            }
            i = j + 1;
            continue;
        }

        // Open tag.
        const name_start = lt + 1;
        var j = name_start;
        while (j < xml.len and !isTagBoundary(xml[j])) j += 1;
        if (j == name_start or j >= xml.len) return error.MalformedSheetXml;
        const qname = xml[name_start..j];
        const te = tagEnd(xml, j) orelse return error.MalformedSheetXml;
        const attrs = xml[j..te.attrs_end];
        if (root_closed) return error.MalformedSheetXml; // a second root

        // Namespace bookkeeping: the default declared here binds this
        // element too; the main namespace bound to a prefix is outside
        // the closed form and refuses (the host/SST predicate's rule).
        var elem_main = if (frames.items.len == 0)
            false
        else
            frames.items[frames.items.len - 1].main_default;
        {
            var it: AttrScan = .{ .rest = attrs };
            var saw_default = false;
            while (try it.next()) |attr| {
                if (std.mem.eql(u8, attr.name, "xmlns")) {
                    if (saw_default) return error.MalformedSheetXml;
                    saw_default = true;
                    elem_main = try bindsMainNs(a, attr.value);
                } else if (std.mem.startsWith(u8, attr.name, "xmlns:")) {
                    if (try bindsMainNs(a, attr.value)) return error.MalformedSheetXml;
                }
            }
        }
        const prefixed = std.mem.indexOfScalar(u8, qname, ':') != null;
        const is_main = !prefixed and elem_main;

        var frame: Frame = .{
            .name = qname,
            .kind = .other,
            .main_default = elem_main,
            .sqref = "",
            .rule = undefined,
        };
        if (frames.items.len == 0) {
            if (!is_main) return error.MalformedSheetXml;
            root_seen = true;
            // A chartsheet or dialogsheet cannot carry conditional
            // formatting, so it contributes an empty inventory — but
            // it is still WALKED whole: an early return here would
            // bypass the second-root and unterminated-nesting checks,
            // so a truncated chartsheet or a `<worksheet/>` followed
            // by a rule-bearing second root would read as an empty
            // success (Codex #215 r2 REL-201).
            const barren = std.mem.eql(u8, qname, "chartsheet") or
                std.mem.eql(u8, qname, "dialogsheet");
            if (!barren and
                !std.mem.eql(u8, qname, "worksheet") and
                !std.mem.eql(u8, qname, "macrosheet"))
            {
                return error.MalformedSheetXml;
            }
            root_kind = if (barren)
                .barren
            else if (std.mem.eql(u8, qname, "macrosheet"))
                .macrosheet
            else
                .worksheet;
            frame.kind = if (barren) .barren_root else .root;
        } else {
            const parent = &frames.items[frames.items.len - 1];
            if (parent.kind == .root and is_main and std.mem.eql(u8, qname, "conditionalFormatting")) {
                frame.kind = .cf_block;
                frame.sqref = (try uniqueAttr(attrs, "sqref")) orelse "";
                // Validate the carrier at block open: a ruleless block
                // — self-closing or paired-empty — emits no record, so
                // a bad sqref there escaped the per-record decode and
                // the read succeeded against its own contract (Codex
                // #215 r4 REL-403).
                a.free(try decodeRuleText(a, frame.sqref));
            } else if (parent.kind == .cf_block and is_main and std.mem.eql(u8, qname, "cfRule")) {
                frame.kind = .cf_rule;
                frame.rule = .{
                    .sqref = parent.sqref,
                    .rule_type = (try uniqueAttr(attrs, "type")) orelse "",
                    .formulas = .{ "", "", "" },
                    .formula_count = 0,
                    .dxf_id = digitU32(try uniqueAttr(attrs, "dxfId")),
                    .priority = digitU32(try uniqueAttr(attrs, "priority")),
                };
            } else if (parent.kind == .cf_rule and is_main and std.mem.eql(u8, qname, "formula")) {
                // Text-only, consumed atomically: the FIRST markup
                // construct after the open tag must be the closing
                // `</formula>` — XML whitespace before its `>` allowed
                // (`</formula >` is valid; Codex #215 r2 REL-202).
                // Anything else — an element, a comment, a CDATA
                // section, a foreign close tag — is markup where the
                // schema puts formula text, or a formula that never
                // closes; both refuse, as does a fourth formula (the
                // schema's maxOccurs is 3).
                _ = try uniqueAttr(attrs, "");
                if (parent.rule.formula_count >= 3) return error.MalformedSheetXml;
                var body: []const u8 = "";
                if (te.self_closing) {
                    i = te.after_gt;
                } else {
                    const lt2 = std.mem.indexOfScalarPos(u8, xml, te.after_gt, '<') orelse
                        return error.MalformedSheetXml;
                    if (lt2 + 1 >= xml.len or xml[lt2 + 1] != '/') return error.MalformedSheetXml;
                    var j2 = lt2 + 2;
                    const ns2 = j2;
                    while (j2 < xml.len and xml[j2] != '>' and !isXmlWs(xml[j2])) j2 += 1;
                    if (!std.mem.eql(u8, xml[ns2..j2], "formula")) return error.MalformedSheetXml;
                    while (j2 < xml.len and isXmlWs(xml[j2])) j2 += 1;
                    if (j2 >= xml.len or xml[j2] != '>') return error.MalformedSheetXml;
                    body = xml[te.after_gt..lt2];
                    i = j2 + 1;
                }
                parent.rule.formulas[parent.rule.formula_count] = body;
                parent.rule.formula_count += 1;
                continue;
            }
        }

        if (te.self_closing) {
            switch (frame.kind) {
                .cf_rule => try out.append(a, frame.rule),
                .root, .barren_root => root_closed = true,
                // A self-closing block or opaque element contributes
                // nothing and opens no frame.
                else => {},
            }
        } else {
            if (frames.items.len >= max_depth) return error.MalformedSheetXml;
            try frames.append(a, frame);
        }
        i = te.after_gt;
    }
    if (!root_seen) return error.MalformedSheetXml;
    if (frames.items.len != 0) return error.MalformedSheetXml;
    return root_kind;
}

fn isXmlWs(c: u8) bool {
    return c == ' ' or c == '\t' or c == '\n' or c == '\r';
}

fn isTagBoundary(c: u8) bool {
    return isXmlWs(c) or c == '/' or c == '>';
}

const TagEnd = struct { attrs_end: usize, after_gt: usize, self_closing: bool };

/// Quote-aware scan from the end of the tag name to the tag's `>` —
/// an attribute value may contain `>` (Codex #215 r1 REL-103 caught
/// the rewriter's bare scan on the same shape).
fn tagEnd(xml: []const u8, attrs_start: usize) ?TagEnd {
    var i = attrs_start;
    while (i < xml.len) : (i += 1) {
        const c = xml[i];
        if (c == '"' or c == '\'') {
            i = std.mem.indexOfScalarPos(u8, xml, i + 1, c) orelse return null;
            continue;
        }
        if (c == '>') {
            const sc = i > attrs_start and xml[i - 1] == '/';
            return .{ .attrs_end = if (sc) i - 1 else i, .after_gt = i + 1, .self_closing = sc };
        }
    }
    return null;
}

const Attr = struct { name: []const u8, value: []const u8 };

/// Strict attribute tokenizer: `name = "value"` with XML whitespace
/// allowed around `=` (valid XML the exact-needle scan reported as
/// absent — Codex #215 r1 REL-102), either quote, nothing else. A
/// region that does not tokenize refuses the part.
const AttrScan = struct {
    rest: []const u8,

    fn next(self: *AttrScan) error{MalformedSheetXml}!?Attr {
        const s = self.rest;
        var i: usize = 0;
        while (i < s.len and isXmlWs(s[i])) i += 1;
        if (i >= s.len) {
            self.rest = s[s.len..];
            return null;
        }
        const name_start = i;
        while (i < s.len and s[i] != '=' and !isXmlWs(s[i])) i += 1;
        if (i == name_start) return error.MalformedSheetXml;
        const name = s[name_start..i];
        while (i < s.len and isXmlWs(s[i])) i += 1;
        if (i >= s.len or s[i] != '=') return error.MalformedSheetXml;
        i += 1;
        while (i < s.len and isXmlWs(s[i])) i += 1;
        if (i >= s.len or (s[i] != '"' and s[i] != '\'')) return error.MalformedSheetXml;
        const q = s[i];
        i += 1;
        const v_start = i;
        while (i < s.len and s[i] != q) i += 1;
        if (i >= s.len) return error.MalformedSheetXml;
        const value = s[v_start..i];
        self.rest = s[i + 1 ..];
        return .{ .name = name, .value = value };
    }
};

/// Look up `key` on a rule-machinery tag while refusing a duplicate
/// of ANY attribute name there (the S7b-4 "a name twice on one start
/// tag" rule, scoped to the elements this read depends on). Pass an
/// empty `key` to run the duplicate check alone.
fn uniqueAttr(attrs: []const u8, key: []const u8) Error!?[]const u8 {
    var names: [64][]const u8 = undefined;
    var n: usize = 0;
    var found: ?[]const u8 = null;
    var it: AttrScan = .{ .rest = attrs };
    while (try it.next()) |attr| {
        for (names[0..n]) |seen| {
            if (std.mem.eql(u8, seen, attr.name)) return error.MalformedSheetXml;
        }
        if (n >= names.len) return error.MalformedSheetXml;
        names[n] = attr.name;
        n += 1;
        if (key.len != 0 and std.mem.eql(u8, attr.name, key)) found = attr.value;
    }
    return found;
}

/// A plain base-10 unsigned integer, or null — `+1`, `1_0` and every
/// other spelling `std.fmt.parseInt` would admit read as absent, the
/// written-but-invalid convention `tableHeaderRowCount` set (Codex
/// #215 r1 REL-104).
fn digitU32(raw: ?[]const u8) ?u32 {
    const s = raw orelse return null;
    if (s.len == 0) return null;
    for (s) |c| {
        if (c < '0' or c > '9') return null;
    }
    return std.fmt.parseInt(u32, s, 10) catch null;
}

/// The sheet-family relationship types a `<sheet r:id>` may carry —
/// the pivots / anchors walk's list, matched exactly under a known
/// relationship-type namespace root by `drawings.relTypeIs`.
const sheet_rel_leaves = [_][]const u8{ "worksheet", "chartsheet", "dialogsheet", "xlMacrosheet", "xlIntlMacrosheet" };

/// Find a relationship by its raw `r:id` attribute value — the
/// `anchor_ndjson` helper, duplicated so this module keeps its imports
/// narrow. The store's `Relationship.id` is entity-decoded, so a rid
/// spelled without references matches byte-for-byte at any length;
/// only one that carries `&` needs the allocating decode.
fn relById(a: Allocator, rels: []const store_mod.Relationship, raw_rid: []const u8) !?store_mod.Relationship {
    if (raw_rid.len == 0) return null;
    for (rels) |rel| {
        if (std.mem.eql(u8, rel.id, raw_rid)) return rel;
    }
    if (std.mem.indexOfScalar(u8, raw_rid, '&') == null) return null;
    const decoded = try store_mod.decodeXmlEntities(a, raw_rid);
    defer a.free(decoded);
    if (decoded.len == 0) return null;
    for (rels) |rel| {
        if (std.mem.eql(u8, rel.id, decoded)) return rel;
    }
    return null;
}

fn decodeSheetName(a: Allocator, raw: []const u8) Error![]u8 {
    const decoded = formula_mod.decode.decodeAt(a, .sheet_name, raw) catch |e| switch (e) {
        error.OutOfMemory => return error.OutOfMemory,
        else => return error.MalformedWorkbookXml,
    };
    if (!std.unicode.utf8ValidateSlice(decoded)) return error.MalformedWorkbookXml;
    return decoded;
}

/// `sqref` / `type` are attribute values, the formulas element text —
/// all three are entities-only carriers (the formula carrier's
/// decode; no ST_Xstring layer). The strict walk already refused raw
/// markup in each; here a carrier that does not decode (a bad entity)
/// and a decoded value that is not UTF-8 refuse — the JSON writer
/// passes bytes through verbatim, so admitting one would emit invalid
/// NDJSON under an exit 0.
fn decodeRuleText(a: Allocator, raw: []const u8) Error![]u8 {
    if (std.mem.indexOfScalar(u8, raw, '<') != null) return error.MalformedSheetXml;
    const decoded = formula_mod.decode.decodeAt(a, .cell_formula_body, raw) catch |e| switch (e) {
        error.OutOfMemory => return error.OutOfMemory,
        else => return error.MalformedSheetXml,
    };
    if (!std.unicode.utf8ValidateSlice(decoded)) return error.MalformedSheetXml;
    return decoded;
}

// ─── Test fixture ────────────────────────────────────────────────────

/// Writes a real two-sheet workbook with conditional formats on BOTH
/// sheets through the public Writer, so the read exercises exactly
/// the bytes the write path produces: `Data` carries all four rule
/// families the writer can author (a `between` cellIs — two formula
/// bodies — an expression whose formula needs entity decoding, a
/// colorScale and a dataBar, priorities 1..4 in document order),
/// `Report` one expression rule whose formula carries a literal `&`.
/// `src/cli.zig` and the tests below share it.
pub const fixture = struct {
    pub fn write(allocator: Allocator, io: std.Io, path: []const u8) !void {
        const zlsx = @import("zlsx");
        var w = zlsx.Writer.init(allocator);
        defer w.deinit();
        const dxf = try w.addDxf(.{ .font_bold = true, .fill_fg_argb = 0xFFFFC7CE });
        var data = try w.addSheet("Data");
        try data.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 5 }, .{ .integer = 9 }, .{ .integer = 3 } });
        try data.writeRow(&.{ .{ .integer = 2 }, .{ .integer = 6 }, .{ .integer = 10 }, .{ .integer = 4 } });
        try data.addConditionalFormatCellIs("A1:A4", .between, "2", "4", dxf);
        try data.addConditionalFormatExpression("B1:B4", "B1>3", dxf);
        try data.addConditionalFormatColorScale("C1:C4", 0xFFF8696B, null, 0xFF63BE7B);
        try data.addConditionalFormatDataBar("D1:D4", 0xFF638EC6);
        var report = try w.addSheet("Report");
        try report.writeRow(&.{.{ .string = "R&D" }});
        try report.addConditionalFormatExpression("A1:A2", "$A1=\"R&D\"", dxf);
        try w.save(io, path);
    }

    /// Byte-replace the first `old` in one part of a saved workbook
    /// and save it back — how the refusal tests make a fixture wrong
    /// in exactly one place (the `pivots.fixture` / `anchor_ndjson`
    /// helper, duplicated so this module keeps its imports narrow).
    pub fn patchPart(allocator: Allocator, io: std.Io, path: []const u8, part: []const u8, old: []const u8, new: []const u8) !void {
        var store = try PartStore.open(allocator, io, path);
        defer store.deinit();
        const p = (try store.part(part)) orelse return error.PartNotFound;
        const at = std.mem.indexOf(u8, p.bytes, old) orelse return error.PatchAnchorNotFound;
        const patched = try std.mem.concat(allocator, u8, &.{ p.bytes[0..at], new, p.bytes[at + old.len ..] });
        defer allocator.free(patched);
        try store.replacePart(part, patched);
        try store.save(io, path);
    }
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

fn scanRules(a: Allocator, xml: []const u8) Error![]RawRule {
    var out: std.ArrayListUnmanaged(RawRule) = .empty;
    errdefer out.deinit(a);
    _ = try scanSheetRules(a, xml, &out);
    return try out.toOwnedSlice(a);
}

const ws_open = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">";

test "scanSheetRules: blocks, rules and formulas at their schema slots" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    const rules = try scanRules(a, ws_open ++
        "<sheetData/>" ++
        "<conditionalFormatting sqref=\"A1:A4 C2\">" ++
        "<cfRule type=\"cellIs\" dxfId=\"0\" priority=\"1\" operator=\"between\"><formula>2</formula><formula>4</formula></cfRule>" ++
        "<cfRule type=\"colorScale\" priority=\"2\"><colorScale><cfvo type=\"min\"/><cfvo type=\"max\"/></colorScale></cfRule>" ++
        "<cfRule type=\"containsBlanks\" priority=\"3\"/>" ++
        "</conditionalFormatting>" ++
        "<conditionalFormatting sqref=\"Z9\"/>" ++
        "</worksheet>");
    try testing.expectEqual(@as(usize, 3), rules.len);
    try testing.expectEqualStrings("A1:A4 C2", rules[0].sqref);
    try testing.expectEqualStrings("cellIs", rules[0].rule_type);
    try testing.expectEqual(@as(u8, 2), rules[0].formula_count);
    try testing.expectEqualStrings("2", rules[0].formulas[0]);
    try testing.expectEqualStrings("4", rules[0].formulas[1]);
    try testing.expectEqual(@as(?u32, 0), rules[0].dxf_id);
    try testing.expectEqual(@as(?u32, 1), rules[0].priority);
    // The colorScale's payload subtree is opaque; no formula leaks in.
    try testing.expectEqual(@as(u8, 0), rules[1].formula_count);
    try testing.expectEqual(@as(?u32, null), rules[1].dxf_id);
    // A self-closing rule is a record; a self-closing block is not.
    try testing.expectEqualStrings("containsBlanks", rules[2].rule_type);
}

test "scanSheetRules: XML whitespace around = and either quote read as authored (REL-102)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    const rules = try scanRules(a, ws_open ++
        "<conditionalFormatting sqref = 'A1:A4'>" ++
        "<cfRule type =\t\"cellIs\" dxfId\n= \"7\" priority = '1' operator=\"equal\"><formula>1</formula></cfRule>" ++
        "</conditionalFormatting>" ++
        "</worksheet>");
    try testing.expectEqual(@as(usize, 1), rules.len);
    try testing.expectEqualStrings("A1:A4", rules[0].sqref);
    try testing.expectEqualStrings("cellIs", rules[0].rule_type);
    try testing.expectEqual(@as(?u32, 7), rules[0].dxf_id);
    try testing.expectEqual(@as(?u32, 1), rules[0].priority);
}

test "scanSheetRules: numeric attrs are digit-only — parseInt's wider grammar reads null (REL-104)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    const rules = try scanRules(a, ws_open ++
        "<conditionalFormatting sqref=\"A1\">" ++
        "<cfRule type=\"cellIs\" dxfId=\"+1\" priority=\"1_0\" operator=\"equal\"><formula>1</formula></cfRule>" ++
        "<cfRule type=\"cellIs\" dxfId=\"4294967296\" priority=\"0\" operator=\"equal\"><formula>1</formula></cfRule>" ++
        "</conditionalFormatting>" ++
        "</worksheet>");
    try testing.expectEqual(@as(usize, 2), rules.len);
    try testing.expectEqual(@as(?u32, null), rules[0].dxf_id);
    try testing.expectEqual(@as(?u32, null), rules[0].priority);
    // Overflow is written-but-invalid too; a real 0 stays a number.
    try testing.expectEqual(@as(?u32, null), rules[1].dxf_id);
    try testing.expectEqual(@as(?u32, 0), rules[1].priority);
}

test "scanSheetRules: an extension tree can never ghost a record (REL-101)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    // An x14 subtree REBINDING THE DEFAULT namespace spells the same
    // unprefixed names — valid XML the lexical view would read as a
    // standard rule. Depth + namespace both exclude it here.
    const rules = try scanRules(a, ws_open ++
        "<conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting>" ++
        "<extLst><ext uri=\"{x14}\">" ++
        "<wrap xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">" ++
        "<conditionalFormatting sqref=\"Z1\"><cfRule type=\"expression\" priority=\"9\"><formula>GHOST</formula></cfRule></conditionalFormatting>" ++
        "</wrap></ext></extLst>" ++
        "</worksheet>");
    try testing.expectEqual(@as(usize, 1), rules.len);
    try testing.expectEqualStrings("1", rules[0].formulas[0]);

    // Same names nested under an extLst INSIDE a block — direct-child
    // classification steps over them.
    const rules2 = try scanRules(a, ws_open ++
        "<conditionalFormatting sqref=\"A1\">" ++
        "<extLst><cfRule type=\"expression\" priority=\"9\"><formula>GHOST</formula></cfRule></extLst>" ++
        "<cfRule type=\"expression\" priority=\"1\"><formula>real</formula></cfRule>" ++
        "</conditionalFormatting>" ++
        "</worksheet>");
    try testing.expectEqual(@as(usize, 1), rules2.len);
    try testing.expectEqualStrings("real", rules2[0].formulas[0]);
}

test "scanSheetRules: malformed structures refuse rather than thin (REL-101)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    const cases = [_][]const u8{
        // A recognised <formula> that never closes is not an absence.
        ws_open ++ "<conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</cfRule></conditionalFormatting></worksheet>",
        // Markup where the schema puts formula text — an element, a
        // comment, a CDATA section (markup is not the formula).
        ws_open ++ "<conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1<x/></formula></cfRule></conditionalFormatting></worksheet>",
        ws_open ++ "<conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1<!-- x --></formula></cfRule></conditionalFormatting></worksheet>",
        ws_open ++ "<conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula><![CDATA[1]]></formula></cfRule></conditionalFormatting></worksheet>",
        // A fourth formula is outside the schema.
        ws_open ++ "<conditionalFormatting sqref=\"A1\"><cfRule type=\"cellIs\" priority=\"1\"><formula>1</formula><formula>2</formula><formula>3</formula><formula>4</formula></cfRule></conditionalFormatting></worksheet>",
        // Mismatched nesting anywhere refuses the walk.
        ws_open ++ "<conditionalFormatting sqref=\"A1\"></worksheet>",
        ws_open ++ "<sheetData></worksheet>",
        // The main namespace aliased to a prefix is outside the
        // closed form (the host/SST predicate's rule).
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:zz=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData/></worksheet>",
        // A duplicate attribute on the rule machinery.
        ws_open ++ "<conditionalFormatting sqref=\"A1\" sqref=\"B1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting></worksheet>",
        ws_open ++ "<conditionalFormatting sqref=\"A1\"><cfRule type=\"a\" type=\"b\" priority=\"1\"/></conditionalFormatting></worksheet>",
        // A root that binds no main default cannot prove main-ns rules.
        "<worksheet><conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting></worksheet>",
    };
    for (cases) |xml| {
        try testing.expectError(error.MalformedSheetXml, scanRules(a, xml));
    }
}

test "scanSheetRules: sheet-family roots — chartsheet empty, macrosheet walked, foreign refused" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    const chart = try scanRules(a, "<chartsheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetViews/></chartsheet>");
    try testing.expectEqual(@as(usize, 0), chart.len);

    const macro = try scanRules(a, "<macrosheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" ++
        "<conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting>" ++
        "</macrosheet>");
    try testing.expectEqual(@as(usize, 1), macro.len);

    try testing.expectError(error.MalformedSheetXml, scanRules(a, "<chartml xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>"));

    // The Strict-conformance namespace is the main namespace too.
    const strict = try scanRules(a, "<worksheet xmlns=\"http://purl.oclc.org/ooxml/spreadsheetml/main\">" ++
        "<conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting>" ++
        "</worksheet>");
    try testing.expectEqual(@as(usize, 1), strict.len);
}

test "scanSheetRules: namespace bindings compare by their decoded value (SEC-301)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    // An entity-spelled prefix binding IS a main-namespace alias —
    // XML resolves character references in attribute values, so a
    // raw-byte compare let a prefixed rule subtree traverse as opaque.
    try testing.expectError(error.MalformedSheetXml, scanRules(a, "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " ++
        "xmlns:zz=\"http://schemas.openxmlformats.org/spreadsheetml&#47;2006/main\"><sheetData/></worksheet>"));

    // An entity-spelled DEFAULT main binding is the main namespace —
    // accepted, not a blanket rejection of every `&`.
    const rules = try scanRules(a, "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml&#47;2006/main\">" ++
        "<conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting>" ++
        "</worksheet>");
    try testing.expectEqual(@as(usize, 1), rules.len);

    // A binding value that does not decode is not one the walk can
    // rule out.
    try testing.expectError(error.MalformedSheetXml, scanRules(a, "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " ++
        "xmlns:zz=\"http://x&bogus;y\"><sheetData/></worksheet>"));
}

test "scanSheetRules: a DOCTYPE cannot smuggle rules past the walk (SEC-302)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    // An internal DTD can define an entity whose expansion IS a rule
    // block; the byte walk cannot expand it, so the declaration
    // refuses rather than the part reading as an empty success.
    try testing.expectError(error.MalformedSheetXml, scanRules(a, "<!DOCTYPE worksheet [" ++
        "<!ENTITY cf '<conditionalFormatting><cfRule type=\"expression\"><formula>1</formula></cfRule></conditionalFormatting>'" ++
        "]>" ++ ws_open ++ "&cf;</worksheet>"));

    // The external form refuses too.
    try testing.expectError(error.MalformedSheetXml, scanRules(a, "<!DOCTYPE worksheet SYSTEM \"sheet.dtd\">" ++
        ws_open ++ "<sheetData/></worksheet>"));

    // Comments and CDATA sections keep their handling: skipped as
    // non-elements, never declarations.
    const rules = try scanRules(a, ws_open ++
        "<!-- note --><conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting>" ++
        "</worksheet>");
    try testing.expectEqual(@as(usize, 1), rules.len);
}

test "scanSheetRules: an undecodable default binding cannot hide rules (SEC-401)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    // `xmlns="&bogus;"` does not decode — the lenient scalar decoder
    // passed the `&` through verbatim, read the block as foreign, and
    // its rules vanished under an exit 0.
    try testing.expectError(error.MalformedSheetXml, scanRules(a, ws_open ++
        "<conditionalFormatting xmlns=\"&bogus;\" sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting>" ++
        "</worksheet>"));
}

test "scanSheetRules: a ruleless block still polices its sqref carrier (REL-403)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    const cases = [_][]const u8{
        // Self-closing with a bad entity, non-UTF-8, and raw markup.
        ws_open ++ "<conditionalFormatting sqref=\"A1&bogus;\"/></worksheet>",
        ws_open ++ "<conditionalFormatting sqref=\"A1\xff\"/></worksheet>",
        ws_open ++ "<conditionalFormatting sqref=\"A1<B\"/></worksheet>",
        // Paired-empty is the same shape.
        ws_open ++ "<conditionalFormatting sqref=\"A1&bogus;\"></conditionalFormatting></worksheet>",
    };
    for (cases) |xml| {
        try testing.expectError(error.MalformedSheetXml, scanRules(a, xml));
    }
}

test "collect: the sheet graph resolves strictly — the anchors rule (REL-402)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();

    const rels = "xl/_rels/workbook.xml.rels";
    const cases = [_]struct { name: []const u8, old: []const u8, new: []const u8 }{
        // A wrong-type relationship with an extant target.
        .{ .name = "g1.xlsx", .old = "officeDocument/2006/relationships/worksheet\"", .new = "officeDocument/2006/relationships/hyperlink\"" },
        // A local-looking External target.
        .{ .name = "g2.xlsx", .old = "Target=\"worksheets/sheet1.xml\"", .new = "Target=\"worksheets/sheet1.xml\" TargetMode=\"External\"" },
        // Two <sheet> entries reaching one part.
        .{ .name = "g3.xlsx", .old = "worksheets/sheet2.xml", .new = "worksheets/sheet1.xml" },
        // A dangling target.
        .{ .name = "g4.xlsx", .old = "worksheets/sheet2.xml", .new = "worksheets/nope.xml" },
    };
    for (cases) |case| {
        const path = try tt.path(testing.allocator, io, case.name);
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path);
        try fixture.patchPart(testing.allocator, io, path, rels, case.old, case.new);
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        try testing.expectError(error.MalformedWorkbookXml, collect(testing.allocator, &wb));
    }
}

test "collect: a worksheet part shares the typed view's parse verdict (REL-404)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "cf_view_verdict.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);
    // A balanced `<sheetDataQ>` wrapper is an opaque, accepted frame
    // to the strict walk — but the typed view's lexical scan prefix-
    // matches it as `<sheetData` and refuses the part. The family
    // contract says every Worksheet reader shares that verdict.
    try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet1.xml", "<sheetData>", "<sheetDataQ>");
    try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet1.xml", "</sheetData>", "</sheetDataQ>");
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    try testing.expectError(error.MalformedSheetXml, collect(testing.allocator, &wb));
}

test "collect: rules of one block share one decoded sqref (PERF-401)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "cf_shared_sqref.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);
    // Give the between-rule's block a second rule so two records share
    // one parent sqref.
    try fixture.patchPart(
        testing.allocator,
        io,
        path,
        "xl/worksheets/sheet1.xml",
        "</cfRule></conditionalFormatting>",
        "</cfRule><cfRule type=\"expression\" priority=\"9\"><formula>1</formula></cfRule></conditionalFormatting>",
    );
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    var view = try collect(testing.allocator, &wb);
    defer view.deinit();
    try testing.expectEqual(@as(usize, 6), view.records.len);
    try testing.expectEqualStrings("A1:A4", view.records[0].sqref);
    try testing.expectEqualStrings("A1:A4", view.records[1].sqref);
    // The SAME decoded slice, not a copy per rule — a large shared
    // sqref must not amplify by the rule count.
    try testing.expect(view.records[0].sqref.ptr == view.records[1].sqref.ptr);
}

test "collect: a DOCTYPE or an entity-spelled main alias in a real workbook refuses (SEC-301/302)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();

    const cases = [_]struct { name: []const u8, old: []const u8, new: []const u8 }{
        .{ .name = "sec1.xlsx", .old = "<worksheet ", .new = "<worksheet xmlns:zz=\"http://schemas.openxmlformats.org/spreadsheetml&#47;2006/main\" " },
        .{ .name = "sec2.xlsx", .old = "<worksheet ", .new = "<!DOCTYPE worksheet [<!ENTITY cf '<cfRule/>'>]><worksheet " },
    };
    for (cases) |case| {
        const path = try tt.path(testing.allocator, io, case.name);
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path);
        try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet1.xml", case.old, case.new);
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        try testing.expectError(error.MalformedSheetXml, collect(testing.allocator, &wb));
    }
}

test "scanSheetRules: early-root shapes cannot bypass whole-part validation (REL-201)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    // A self-closing root followed by a second, rule-bearing root
    // must refuse — an early return at root classification read this
    // as an empty success and the second root's rules vanished.
    try testing.expectError(error.MalformedSheetXml, scanRules(a, "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>" ++
        ws_open ++ "<conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting></worksheet>"));

    // An unterminated chartsheet is mismatched nesting, not an empty
    // inventory.
    try testing.expectError(error.MalformedSheetXml, scanRules(a, "<chartsheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetViews/>"));

    // A lone self-closing worksheet root is well-formed: an empty
    // sheet holds no rules.
    const empty = try scanRules(a, "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>");
    try testing.expectEqual(@as(usize, 0), empty.len);

    // A chartsheet is walked whole; a rule-shaped subtree inside it
    // is opaque (the part cannot carry rules by schema), and its
    // nesting still has to prove.
    const chart = try scanRules(a, "<chartsheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" ++
        "<conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting>" ++
        "</chartsheet>");
    try testing.expectEqual(@as(usize, 0), chart.len);
    try testing.expectError(error.MalformedSheetXml, scanRules(a, "<chartsheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><a></chartsheet>"));
}

test "scanSheetRules: XML whitespace before a formula close's '>' reads (REL-202)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    const rules = try scanRules(a, ws_open ++
        "<conditionalFormatting sqref=\"A1\">" ++
        "<cfRule type=\"cellIs\" priority=\"1\" operator=\"between\"><formula>1</formula ><formula>2</formula\n></cfRule>" ++
        "</conditionalFormatting>" ++
        "</worksheet>");
    try testing.expectEqual(@as(usize, 1), rules.len);
    try testing.expectEqual(@as(u8, 2), rules[0].formula_count);
    try testing.expectEqualStrings("1", rules[0].formulas[0]);
    try testing.expectEqualStrings("2", rules[0].formulas[1]);
}

test "scanSheetRules: nesting past the ceiling refuses; deep-but-bounded walks (PERF-201)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(testing.allocator);
    try buf.appendSlice(testing.allocator, ws_open);
    for (0..max_depth) |_| try buf.appendSlice(testing.allocator, "<a>");
    // The ceiling trips at the push past max_depth — before EOF.
    try testing.expectError(error.MalformedSheetXml, scanRules(a, buf.items));

    buf.clearRetainingCapacity();
    try buf.appendSlice(testing.allocator, ws_open);
    for (0..1000) |_| try buf.appendSlice(testing.allocator, "<a>");
    for (0..1000) |_| try buf.appendSlice(testing.allocator, "</a>");
    try buf.appendSlice(testing.allocator, "<conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting></worksheet>");
    const rules = try scanRules(a, buf.items);
    try testing.expectEqual(@as(usize, 1), rules.len);
}

test "scanSheetRules: an attribute value may contain '>' — quote-aware tag ends (REL-103's shape)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    const rules = try scanRules(a, ws_open ++
        "<conditionalFormatting sqref=\"A1\">" ++
        "<cfRule type=\"expression\" priority=\"1\" x=\">\"><formula>1</formula></cfRule>" ++
        "</conditionalFormatting>" ++
        "</worksheet>");
    try testing.expectEqual(@as(usize, 1), rules.len);
    try testing.expectEqualStrings("1", rules[0].formulas[0]);
}

test "writeRecord: full and compact envelopes, exact bytes" {
    const rec: Record = .{
        .sheet = "R\"D",
        .sheet_idx = 1,
        .sqref = "A1:A4 C2",
        .rule_type = "cellIs",
        .formulas = &.{ "2", "SUM(A1,\"x\")" },
        .dxf_id = 0,
        .priority = 1,
    };
    var buf: [512]u8 = undefined;
    var w = std.Io.Writer.fixed(&buf);
    try writeRecord(&w, rec, .full);
    try testing.expectEqualStrings(
        "{\"kind\":\"conditional_format\",\"sheet\":\"R\\\"D\",\"sheet_idx\":1," ++
            "\"sqref\":\"A1:A4 C2\",\"rule_type\":\"cellIs\"," ++
            "\"formulas\":[\"2\",\"SUM(A1,\\\"x\\\")\"],\"dxf_id\":0,\"priority\":1}\n",
        w.buffered(),
    );
    var w2 = std.Io.Writer.fixed(&buf);
    try writeRecord(&w2, rec, .compact);
    try testing.expectEqualStrings(
        "{\"kind\":\"conditional_format\",\"sqref\":\"A1:A4 C2\",\"rule_type\":\"cellIs\"," ++
            "\"formulas\":[\"2\",\"SUM(A1,\\\"x\\\")\"],\"dxf_id\":0,\"priority\":1}\n",
        w2.buffered(),
    );
}

test "writeRecord: empty formulas, absent dxf and priority, empty sqref and type" {
    const rec: Record = .{
        .sheet = "S",
        .sheet_idx = 0,
        .sqref = "",
        .rule_type = "",
        .formulas = &.{},
        .dxf_id = null,
        .priority = null,
    };
    var buf: [256]u8 = undefined;
    var w = std.Io.Writer.fixed(&buf);
    try writeRecord(&w, rec, .full);
    try testing.expectEqualStrings(
        "{\"kind\":\"conditional_format\",\"sheet\":\"S\",\"sheet_idx\":0," ++
            "\"sqref\":\"\",\"rule_type\":\"\",\"formulas\":[],\"dxf_id\":null,\"priority\":null}\n",
        w.buffered(),
    );
}

test "collect: fixture rules attributed in sheet then document order, entities decoded" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "cf_fixture.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);

    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    var view = try collect(testing.allocator, &wb);
    defer view.deinit();

    try testing.expectEqual(@as(usize, 2), view.sheet_names.len);
    try testing.expectEqual(@as(usize, 5), view.records.len);

    const between = view.records[0];
    try testing.expectEqualStrings("Data", between.sheet);
    try testing.expectEqual(@as(u32, 0), between.sheet_idx);
    try testing.expectEqualStrings("A1:A4", between.sqref);
    try testing.expectEqualStrings("cellIs", between.rule_type);
    try testing.expectEqual(@as(usize, 2), between.formulas.len);
    try testing.expectEqualStrings("2", between.formulas[0]);
    try testing.expectEqualStrings("4", between.formulas[1]);
    try testing.expectEqual(@as(?u32, 0), between.dxf_id);
    try testing.expectEqual(@as(?u32, 1), between.priority);

    const expr = view.records[1];
    try testing.expectEqualStrings("expression", expr.rule_type);
    // The writer stored `B1&gt;3`; the read hands back what was
    // authored.
    try testing.expectEqual(@as(usize, 1), expr.formulas.len);
    try testing.expectEqualStrings("B1>3", expr.formulas[0]);
    try testing.expectEqual(@as(?u32, 2), expr.priority);

    const scale = view.records[2];
    try testing.expectEqualStrings("colorScale", scale.rule_type);
    try testing.expectEqual(@as(usize, 0), scale.formulas.len);
    try testing.expectEqual(@as(?u32, null), scale.dxf_id);
    try testing.expectEqual(@as(?u32, 3), scale.priority);

    const bar = view.records[3];
    try testing.expectEqualStrings("dataBar", bar.rule_type);
    try testing.expectEqualStrings("D1:D4", bar.sqref);
    try testing.expectEqual(@as(usize, 0), bar.formulas.len);

    const report = view.records[4];
    try testing.expectEqualStrings("Report", report.sheet);
    try testing.expectEqual(@as(u32, 1), report.sheet_idx);
    try testing.expectEqualStrings("$A1=\"R&D\"", report.formulas[0]);
    try testing.expectEqual(@as(?u32, 1), report.priority);
}

test "collect: allocation failures surface as OutOfMemory, never a partial view" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "cf_alloc_sweep.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);

    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    // Prime the workbook-owned caches (Worksheet handles, part names,
    // part bytes) so the sweep exercises collect's own allocations.
    {
        var primed = try collect(testing.allocator, &wb);
        primed.deinit();
    }
    try std.testing.checkAllAllocationFailures(testing.allocator, struct {
        fn run(failing: Allocator, wb_: *workbook_mod.Workbook) !void {
            var view = try collect(failing, wb_);
            view.deinit();
        }
    }.run, .{&wb});
}

test "collect: a rule field the stream cannot carry faithfully refuses whole" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();

    const sheet1 = "xl/worksheets/sheet1.xml";
    const cases = [_]struct { name: []const u8, part: []const u8, old: []const u8, new: []const u8 }{
        // A formula carrier that does not decode (a bad entity).
        .{ .name = "r1.xlsx", .part = sheet1, .old = "B1&gt;3", .new = "B1&bogus;3" },
        // Embedded markup where the schema puts formula text.
        .{ .name = "r2.xlsx", .part = sheet1, .old = "<formula>2</formula>", .new = "<formula>2<x/></formula>" },
        // A formula that decodes to non-UTF-8.
        .{ .name = "r3.xlsx", .part = sheet1, .old = "B1&gt;3", .new = "B1\xff3" },
        // A sqref carrier that does not decode.
        .{ .name = "r4.xlsx", .part = sheet1, .old = "sqref=\"A1:A4\"", .new = "sqref=\"A1&bogus;A4\"" },
        // A rule type that is not UTF-8.
        .{ .name = "r5.xlsx", .part = sheet1, .old = "type=\"dataBar\"", .new = "type=\"data\xffBar\"" },
        // A recognised <formula> that never closes is a refusal, not
        // an empty formulas list (REL-101): the only formula of
        // Report's rule, its closer gone, the rule's closer intact.
        .{ .name = "r6.xlsx", .part = "xl/worksheets/sheet2.xml", .old = "R&amp;D&quot;</formula>", .new = "R&amp;D&quot;" },
    };
    for (cases) |case| {
        const path = try tt.path(testing.allocator, io, case.name);
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path);
        try fixture.patchPart(testing.allocator, io, path, case.part, case.old, case.new);
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        try testing.expectError(error.MalformedSheetXml, collect(testing.allocator, &wb));
    }
}

test "collect: a sheet-name carrier that does not decode refuses whole" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "cf_bad_sheet_name.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);
    try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "name=\"Report\"", "name=\"Rep&bogus;\"");
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    try testing.expectError(error.MalformedWorkbookXml, collect(testing.allocator, &wb));
}

test "collect: absent sqref or type reads as the empty string — the wire's merge, pinned" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "cf_absent_attrs.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);
    try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet1.xml", " sqref=\"D1:D4\"", "");
    try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet1.xml", "type=\"dataBar\" ", "");
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    var view = try collect(testing.allocator, &wb);
    defer view.deinit();
    try testing.expectEqual(@as(usize, 5), view.records.len);
    try testing.expectEqualStrings("", view.records[3].sqref);
    try testing.expectEqualStrings("", view.records[3].rule_type);
}

test "collect: attribute spacing and a third formula through a real saved workbook (REL-102, MNT-101)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "cf_spacing_three.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);
    try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet1.xml", "sqref=\"A1:A4\"", "sqref = 'A1:A4'");
    try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet1.xml", "<formula>4</formula></cfRule>", "<formula>4</formula><formula>6</formula></cfRule>");
    try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet1.xml", "type=\"expression\" dxfId=\"0\"", "type=\"expression\" dxfId=\"+0\"");
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    var view = try collect(testing.allocator, &wb);
    defer view.deinit();
    try testing.expectEqual(@as(usize, 5), view.records.len);
    try testing.expectEqualStrings("A1:A4", view.records[0].sqref);
    try testing.expectEqual(@as(usize, 3), view.records[0].formulas.len);
    try testing.expectEqualStrings("6", view.records[0].formulas[2]);
    try testing.expectEqual(@as(?u32, null), view.records[1].dxf_id);
    // The lenient Zig view reads the same third body out of the same
    // real part — the `formula3` widening exercised through a saved
    // workbook, not only unit XML (Codex #215 r4 MNT-402).
    const ws = try wb.sheet(0);
    const cfs = try ws.conditionalFormats();
    try testing.expectEqualStrings("2", cfs[0].formula.?);
    try testing.expectEqualStrings("4", cfs[0].formula2.?);
    try testing.expectEqualStrings("6", cfs[0].formula3.?);
}

test "collect: a decoded control character rides out as a JSON escape, not raw bytes" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "cf_ctrl_char.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);
    try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet1.xml", "B1&gt;3", "B1&#7;3");
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    var view = try collect(testing.allocator, &wb);
    defer view.deinit();
    try testing.expectEqualStrings("B1\x073", view.records[1].formulas[0]);
    var buf: [512]u8 = undefined;
    var w = std.Io.Writer.fixed(&buf);
    try writeRecord(&w, view.records[1], .full);
    try testing.expect(std.mem.indexOf(u8, w.buffered(), "\\u0007") != null);
}

test "collect: an extension tree in a real workbook can never ghost a record (REL-101)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "cf_ghost_ext.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);
    try fixture.patchPart(
        testing.allocator,
        io,
        path,
        "xl/worksheets/sheet2.xml",
        "</worksheet>",
        "<extLst><ext uri=\"{x14}\"><wrap xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">" ++
            "<conditionalFormatting sqref=\"Z1\"><cfRule type=\"expression\" priority=\"9\"><formula>GHOST</formula></cfRule></conditionalFormatting>" ++
            "</wrap></ext></extLst></worksheet>",
    );
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    var view = try collect(testing.allocator, &wb);
    defer view.deinit();
    try testing.expectEqual(@as(usize, 5), view.records.len);
    for (view.records) |r| {
        try testing.expect(!std.mem.eql(u8, r.sqref, "Z1"));
    }
}

test "collect: a bodiless self-closing conditionalFormatting block emits nothing" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "cf_selfclosing_block.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);
    try fixture.patchPart(
        testing.allocator,
        io,
        path,
        "xl/worksheets/sheet2.xml",
        "<conditionalFormatting sqref=\"A1:A2\">",
        "<conditionalFormatting sqref=\"Z9\"/><conditionalFormatting sqref=\"A1:A2\">",
    );
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    var view = try collect(testing.allocator, &wb);
    defer view.deinit();
    // Still five records — the empty block contributes none.
    try testing.expectEqual(@as(usize, 5), view.records.len);
    try testing.expectEqualStrings("A1:A2", view.records[4].sqref);
}

test "collect: a workbook without conditional formats is an empty view" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "cf_none.xlsx");
    defer testing.allocator.free(path);
    {
        const zlsx = @import("zlsx");
        var w = zlsx.Writer.init(testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Only");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, path);
    }
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    var view = try collect(testing.allocator, &wb);
    defer view.deinit();
    try testing.expectEqual(@as(usize, 0), view.records.len);
    try testing.expectEqual(@as(usize, 1), view.sheet_names.len);
}
