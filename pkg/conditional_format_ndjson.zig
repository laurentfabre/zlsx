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
//!
//! The lexical layer below — the prefix scope, the attribute
//! tokenizer, the strict prolog / comment / PI skippers, the root
//! classifier, the carrier decoders — is `pub` for the sibling strict
//! reads that walk the same parts (the sheet-props read,
//! `sheet_props_ndjson.zig`), so one grammar rules every typed read;
//! the walk loops themselves stay per module, each classifying its
//! own schema slots.

const std = @import("std");
const formula_mod = @import("zlsx_formula");
const workbook_mod = @import("workbook.zig");
const store_mod = @import("store.zig");
const drawings = @import("drawings.zig");
const workbook_xml = @import("typed_parts/root.zig").workbook_xml;
const sheet_xml = @import("typed_parts/root.zig").sheet_xml;
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

/// What `collect` itself can raise — `Error` plus the two verdicts
/// that keep their own names across every boundary: the archive-wide
/// decompression caps, and a sheet part the store lost between the
/// graph proof and the walk. Closed and explicit so the C boundary's
/// status mapping is compiler-checked, not assumed (Codex #216 r1
/// S3B-ERR-602).
pub const CollectError = Error || error{ ZipBombSuspected, MissingSheetPart };

/// The workbook side of `Error` — what the strict `xl/workbook.xml`
/// and `workbook.xml.rels` readers raise. Narrower than `Error` so a
/// consumer that only needs the sheet inventory (the anchors
/// collector) inherits no sheet-part verdict it can never produce.
pub const WorkbookError = error{ MalformedWorkbookXml, OutOfMemory };

/// What `resolveSheets` can raise: the workbook verdicts plus the
/// archive-wide decompression caps, which keep their name across
/// every boundary.
pub const InventoryError = WorkbookError || error{ZipBombSuspected};

/// The strict sheet inventory of a workbook: every `<sheet>` entry
/// the strict read proves, decoded and placed. Both slices live in
/// the arena `resolveSheets` was handed and are parallel — `parts[i]`
/// is the worksheet part `names[i]` reaches.
pub const SheetInventory = struct {
    /// Decoded sheet names, workbook order.
    names: []const []const u8,
    /// The part each entry's relationship reaches, already probed
    /// (materialised) in the store.
    parts: []const []const u8,
};

/// The AUTHORITATIVE sheet inventory, shared by every read that
/// attributes per-sheet records (this module's rules, the anchors
/// collector): the strict namespace- and depth-aware read of
/// `xl/workbook.xml` (`scanWorkbookSheets`) — not the lenient lexical
/// projection, which counts `<sheet>` tags document-wide, drops
/// carrier-less entries silently, and only spells the literal `r:id`
/// (Codex #215 r5 SEC-502, r8 REL-801: a ghost under `<extLst>` is
/// excluded here, a real entry missing a carrier refuses, and a valid
/// alternate relationships prefix is an identity) — then each entry
/// resolved to its part strictly. Names and parts land in `a` (the
/// caller's view arena); scratch lives on `gpa` and is reclaimed
/// before return.
pub fn resolveSheets(gpa: Allocator, a: Allocator, wb: *workbook_mod.Workbook) InventoryError!SheetInventory {
    var strict_sheets: std.ArrayListUnmanaged(StrictSheet) = .empty;
    defer strict_sheets.deinit(gpa);
    {
        const wb_part = (wb.store.part("xl/workbook.xml") catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            error.ZipBombSuspected => return error.ZipBombSuspected,
            else => return error.MalformedWorkbookXml,
        }) orelse return error.MalformedWorkbookXml;
        // The calc-props slot rides the same walk; this read has no
        // use for it (`scanCalcPr` is the calc-props entry).
        var calc_slot: CalcPrSlot = .{};
        try scanWorkbookSheets(gpa, wb_part.bytes, &strict_sheets, &calc_slot);
    }
    const sheet_count = strict_sheets.items.len;

    const sheet_names = try a.alloc([]const u8, sheet_count);
    {
        // Two sheets whose names DECODE to one spelling would make
        // `--name` silently incomplete — the selector stops at the
        // first index and the later sheet's rules vanish under exit 0
        // (Codex #215 r19 REL-1901).
        var seen_names: std.StringHashMapUnmanaged(void) = .empty;
        defer seen_names.deinit(gpa);
        for (strict_sheets.items, 0..) |s, i| {
            sheet_names[i] = try decodeSheetName(a, s.name);
            const g = try seen_names.getOrPut(gpa, sheet_names[i]);
            if (g.found_existing) return error.MalformedWorkbookXml;
        }
    }

    // Resolve each sheet to its part STRICTLY — the anchors
    // collector's rule: the relationship under the entry's id must
    // exist, be a sheet-family type, be internal, and reach a part
    // the archive holds that no other sheet reaches (the same rule
    // would ride out twice under two identities) — Codex #215 r4
    // REL-402. Lookups are hashed: an id→relationship map and a
    // seen-targets set replace the per-sheet linear scans a crafted
    // graph amplified quadratically (r8 PERF-801).
    const wb_rels = wb.store.rels("xl/workbook.xml");
    // The store's relationship list is lexical too — verify it against
    // a strict read of the rels part before resolving through it, so a
    // nested decoy or a duplicate Id cannot select an arbitrary part
    // (Codex #215 r6 SEC-601).
    {
        const rels_part = wb.store.part("xl/_rels/workbook.xml.rels") catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            error.ZipBombSuspected => return error.ZipBombSuspected,
            else => return error.MalformedWorkbookXml,
        };
        if (rels_part) |rp| {
            try verifyWorkbookRels(gpa, rp.bytes, wb_rels);
        } else if (wb_rels.len != 0) {
            return error.MalformedWorkbookXml;
        }
    }
    var rel_by_id: std.StringHashMapUnmanaged(store_mod.Relationship) = .empty;
    defer rel_by_id.deinit(gpa);
    try rel_by_id.ensureTotalCapacity(gpa, @intCast(wb_rels.len));
    for (wb_rels) |rel| {
        const g = rel_by_id.getOrPutAssumeCapacity(rel.id);
        if (g.found_existing) return error.MalformedWorkbookXml;
        g.value_ptr.* = rel;
    }
    var seen_targets: std.StringHashMapUnmanaged(void) = .empty;
    defer seen_targets.deinit(gpa);
    const sheet_parts = try a.alloc([]const u8, sheet_count);
    for (strict_sheets.items, 0..) |s, i| {
        const rel = blk: {
            if (std.mem.indexOfScalar(u8, s.rid, '&') != null) {
                // The XML-semantic id is the DECODED spelling (r5
                // REL-503); one that does not decode refuses.
                const decoded = try decodeRelText(gpa, s.rid);
                defer gpa.free(decoded);
                break :blk rel_by_id.get(decoded) orelse return error.MalformedWorkbookXml;
            }
            break :blk rel_by_id.get(s.rid) orelse return error.MalformedWorkbookXml;
        };
        var typed = false;
        for (sheet_rel_leaves) |leaf| typed = typed or drawings.relTypeIs(rel.type, leaf);
        if (!typed) return error.MalformedWorkbookXml;
        if (rel.target_mode == .external) return error.MalformedWorkbookXml;
        // Resolved into the VIEW arena, not the store's lifetime arena
        // — a long-lived editor repeats this read per call, and the
        // store variant would retain every resolved path until close
        // (Codex #216 r1 S3B-MEM-603).
        const name = (try wb.store.resolveOwned(a, "xl/workbook.xml", rel.target)) orelse
            return error.MalformedWorkbookXml;
        // This lookup MATERIALISES on first touch — its store failures
        // fold like every other part read here, or a bad CRC would
        // escape with the zip layer's own name (Codex #216 r1
        // S3B-ERR-602). The sheet loop below then only cache-hits.
        const probed = wb.store.part(name) catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            error.ZipBombSuspected => return error.ZipBombSuspected,
            else => return error.MalformedWorkbookXml,
        };
        if (probed == null) return error.MalformedWorkbookXml;
        const g = try seen_targets.getOrPut(gpa, name);
        if (g.found_existing) return error.MalformedWorkbookXml;
        sheet_parts[i] = name;
    }
    return .{ .names = sheet_names, .parts = sheet_parts };
}

/// Collect every rule of every workbook sheet. A read this module
/// cannot serve faithfully refuses whole — even for a sheet the CLI's
/// selection would exclude: the inventory is proven whole before
/// selection and pagination apply, the `anchors` rule. A sheet the
/// workbook itself cannot place (a dangling relationship, a part the
/// archive does not hold) passes through as `Workbook`'s own error.
pub fn collect(gpa: Allocator, wb: *workbook_mod.Workbook) CollectError!ConditionalFormats {
    var arena = std.heap.ArenaAllocator.init(gpa);
    errdefer arena.deinit();
    const a = arena.allocator();

    const inventory = try resolveSheets(gpa, a, wb);
    const sheet_names = inventory.names;
    const sheet_parts = inventory.parts;
    const sheet_count = sheet_names.len;

    var records: std.ArrayListUnmanaged(Record) = .empty;
    for (0..sheet_count) |idx| {
        // Store failures fold like `Worksheet.ensureParsed`: memory
        // and the archive-wide budget keep their names, everything
        // else is "this sheet is not readable".
        const part = (wb.store.part(sheet_parts[idx]) catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            error.ZipBombSuspected => return error.ZipBombSuspected,
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
        // refuses here too (Codex #215 r4 REL-404). The gate parses
        // the STRICTLY-resolved bytes directly rather than going back
        // through `Worksheet` — whose raw-compare rid resolution
        // would spuriously refuse an entity-spelled `r:id` the strict
        // graph already placed (r5 REL-503). The other roots have no
        // typed view; the strict walk's verdict stands alone.
        if (root == .worksheet) {
            var gate = sheet_xml.parse(gpa, part.bytes) catch |e| switch (e) {
                error.OutOfMemory => return error.OutOfMemory,
                else => return error.MalformedSheetXml,
            };
            gate.deinit(gpa);
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

/// The unselected stream — every record, emission order. The C leg's
/// entry point (`zlsx_editor_conditional_formats_ndjson`).
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

/// The sheet-family roots a sheet part may open with — the one root
/// rule every strict sheet walk (this module's rules, the sheet-props
/// read) applies at its first element.
pub const SheetRoot = enum { worksheet, macrosheet, chartsheet, dialogsheet };

/// Classify a part's root element, or null when it is not a
/// sheet-family root. `is_main` / `elem_main` are the walk's namespace
/// verdicts for the element (unprefixed under a main default; the
/// main namespace is the element's default). Macro sheets have TWO
/// legal spellings: the noncanonical unprefixed main-ns
/// `<macrosheet>`, and Microsoft's canonical `<xm:macrosheet>` whose
/// prefix binds the macro namespace ON THIS TAG while the default
/// binds main for the children (Codex #215 r20 REL-2001) — an
/// arbitrary bound prefix is not that. Structural compare — a fixed
/// name buffer put an artificial ceiling on a valid prefix's length
/// (r21 REL-2101).
pub fn classifySheetRoot(scratch: Allocator, qname: []const u8, attrs: []const u8, is_main: bool, elem_main: bool) Error!?SheetRoot {
    if (is_main) {
        if (std.mem.eql(u8, qname, "worksheet")) return .worksheet;
        if (std.mem.eql(u8, qname, "macrosheet")) return .macrosheet;
        if (std.mem.eql(u8, qname, "chartsheet")) return .chartsheet;
        if (std.mem.eql(u8, qname, "dialogsheet")) return .dialogsheet;
        return null;
    }
    const c = std.mem.indexOfScalar(u8, qname, ':') orelse return null;
    if (!elem_main) return null;
    if (!std.mem.eql(u8, qname[c + 1 ..], "macrosheet")) return null;
    const p = qname[0..c];
    var it: AttrScan = .{ .rest = attrs };
    while (try it.next()) |attr| {
        if (std.mem.startsWith(u8, attr.name, "xmlns:") and
            std.mem.eql(u8, attr.name["xmlns:".len..], p))
        {
            return if (try bindsNs(scratch, attr.value, isMsMacroNs)) .macrosheet else null;
        }
    }
    return null;
}

/// Microsoft's macro-sheet namespace — the canonical macro-sheet
/// root is `<xm:macrosheet>` with `xm` bound to it and the DEFAULT
/// bound to SpreadsheetML main for the children (MS-OFFMACRO; Codex
/// #215 r20 REL-2001).
const ms_macro_ns = "http://schemas.microsoft.com/office/excel/2006/main";

pub fn isMsMacroNs(uri: []const u8) bool {
    return std.mem.eql(u8, uri, ms_macro_ns);
}

const main_ns_transitional = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const main_ns_strict = "http://purl.oclc.org/ooxml/spreadsheetml/main";

pub fn isMainNs(uri: []const u8) bool {
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
pub fn bindsMainNs(scratch: Allocator, raw: []const u8) Error!bool {
    return bindsNs(scratch, raw, isMainNs);
}

const rel_ns_transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const rel_ns_strict = "http://purl.oclc.org/ooxml/officeDocument/relationships";

fn isRelNs(uri: []const u8) bool {
    return std.mem.eql(u8, uri, rel_ns_transitional) or
        std.mem.eql(u8, uri, rel_ns_strict);
}

fn bindsRelNs(scratch: Allocator, raw: []const u8) Error!bool {
    return bindsNs(scratch, raw, isRelNs);
}

/// The OPC package-relationships namespace (`.rels` parts).
const pkg_rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships";

fn isPkgRelsNs(uri: []const u8) bool {
    return std.mem.eql(u8, uri, pkg_rels_ns);
}

/// The two RESERVED namespaces (Namespaces in XML 1.0 §3): `xml` is
/// implicitly bound to the first and may only be re-declared to it;
/// nothing may bind to the second, and the `xmlns` prefix may not be
/// declared at all.
const xml_ns_uri = "http://www.w3.org/XML/1998/namespace";
const xmlns_ns_uri = "http://www.w3.org/2000/xmlns/";

fn isXmlNsUri(uri: []const u8) bool {
    return std.mem.eql(u8, uri, xml_ns_uri);
}

fn isXmlnsNsUri(uri: []const u8) bool {
    return std.mem.eql(u8, uri, xmlns_ns_uri);
}

/// The Markup Compatibility and Extensibility namespace (ECMA-376
/// Part 3). Its DECLARATION is inert — every modern Excel part binds
/// `mc` and lists `mc:Ignorable` prefixes — but its processing
/// constructs re-project the tree: `mc:AlternateContent` substitutes
/// a branch in place, `mc:ProcessContent` lifts a wrapper's children.
/// The walk implements no MCE processor, so it refuses exactly the
/// constructs whose projection could reach a recognized slot (Codex
/// #215 r22 SEC-2201); a deep `mc:AlternateContent` — real worksheets
/// wrap `<oleObject>` alternates — projects in place inside an
/// already-opaque subtree and stays walkable.
const mce_ns = "http://schemas.openxmlformats.org/markup-compatibility/2006";

pub fn isMceNs(uri: []const u8) bool {
    return std.mem.eql(u8, uri, mce_ns);
}

/// Reset the per-tag declaration set: a declaration-free tag skips
/// the clear entirely, and oversized capacity is dropped so one wide
/// tag cannot tax every later tag with a full-capacity metadata
/// clear — `clearRetainingCapacity` zeroes the whole retained
/// capacity (Codex #215 r20 PERF-2001).
pub fn resetDeclSeen(a: Allocator, m: *std.StringHashMapUnmanaged(void)) void {
    if (m.count() == 0) return;
    if (m.capacity() > 64) {
        m.deinit(a);
        m.* = .empty;
    } else {
        m.clearRetainingCapacity();
    }
}

/// In-scope prefix bindings: a declaration stack (per-frame
/// watermarks pop exactly what a frame declared) beside a REFCOUNTED
/// hash map, so the per-attribute lookup is O(1) — the linear scan
/// was quadratic against a tag with many declarations and many
/// prefixed attributes (Codex #215 r15 PERF-1501).
pub const PrefixScope = struct {
    const Entry = struct { prefix: []const u8, prev_mce: ?bool };
    const Bind = struct { count: u32, mce: bool };
    stack: std.ArrayListUnmanaged(Entry) = .empty,
    counts: std.StringHashMapUnmanaged(Bind) = .empty,

    pub fn deinit(self: *PrefixScope, a: Allocator) void {
        self.stack.deinit(a);
        self.counts.deinit(a);
    }

    pub fn mark(self: *const PrefixScope) usize {
        return self.stack.items.len;
    }

    pub fn declare(self: *PrefixScope, a: Allocator, p: []const u8, mce: bool) !void {
        const prev: ?bool = if (self.counts.get(p)) |b| b.mce else null;
        try self.stack.append(a, .{ .prefix = p, .prev_mce = prev });
        const g = try self.counts.getOrPut(a, p);
        if (g.found_existing) {
            g.value_ptr.count += 1;
            g.value_ptr.mce = mce;
        } else {
            g.value_ptr.* = .{ .count = 1, .mce = mce };
        }
    }

    pub fn truncate(self: *PrefixScope, to: usize) void {
        while (self.stack.items.len > to) {
            const e = self.stack.items[self.stack.items.len - 1];
            self.stack.items.len -= 1;
            const v = self.counts.getPtr(e.prefix).?;
            if (v.count == 1) {
                _ = self.counts.remove(e.prefix);
            } else {
                v.count -= 1;
                // The popped binding shadowed an outer one — restore
                // that outer binding's classification.
                v.mce = e.prev_mce.?;
            }
        }
    }

    pub fn contains(self: *const PrefixScope, p: []const u8) bool {
        if (std.mem.eql(u8, p, "xml")) return true; // implicitly declared
        return self.counts.contains(p);
    }

    /// Whether the prefix's INNERMOST in-scope binding is the MCE
    /// namespace (SEC-2201).
    pub fn isMce(self: *const PrefixScope, p: []const u8) bool {
        return if (self.counts.get(p)) |b| b.mce else false;
    }
};

/// Process one element's namespace declarations into the in-scope
/// prefix stack and validate the reserved rules, then require the
/// element's own prefix and every prefixed non-declaration attribute
/// to be IN SCOPE — an undeclared `p:sheet` (or `p:cfRule`,
/// `p:Relationship`) is namespace-malformed XML, not harmless opaque
/// content it could thin the inventory as; and `xmlns:xmlns="…"` must
/// not mint the `xmlns` prefix as a relationship authorizer (Codex
/// #215 r11 SEC-1101). Declarations THEMSELVES are never candidates
/// for `<p>:id` matching — callers skip `xmlns`-family names.
pub fn enterElementScope(
    scratch: Allocator,
    scope: *PrefixScope,
    qname: []const u8,
    attrs: []const u8,
) Error!void {
    var it: AttrScan = .{ .rest = attrs };
    while (try it.next()) |attr| {
        if (std.mem.eql(u8, attr.name, "xmlns")) {
            // A DEFAULT binding may not name either reserved
            // namespace — `xmlns="http://www.w3.org/2000/xmlns/"` on
            // a nested rule block otherwise classified it foreign and
            // thinned the inventory under exit 0 (Codex #215 r12
            // SEC-1201).
            if (try bindsNs(scratch, attr.value, isXmlNsUri)) return error.MalformedSheetXml;
            if (try bindsNs(scratch, attr.value, isXmlnsNsUri)) return error.MalformedSheetXml;
            // A default binding of the MCE namespace is the one way
            // an UNPREFIXED element could be MCE-bound past the
            // element rule at the recognized slots; no real producer
            // default-binds it (SEC-2201).
            if (try bindsNs(scratch, attr.value, isMceNs)) return error.MalformedSheetXml;
        } else if (std.mem.startsWith(u8, attr.name, "xmlns:")) {
            const p = attr.name["xmlns:".len..];
            if (std.mem.eql(u8, p, "xmlns")) return error.MalformedSheetXml;
            // Namespaces in XML 1.0 forbids undeclaring a prefix.
            if (attr.value.len == 0) return error.MalformedSheetXml;
            const binds_xml = try bindsNs(scratch, attr.value, isXmlNsUri);
            if (std.mem.eql(u8, p, "xml")) {
                if (!binds_xml) return error.MalformedSheetXml;
                continue; // implicit; not re-recorded
            } else if (binds_xml) {
                return error.MalformedSheetXml;
            }
            if (try bindsNs(scratch, attr.value, isXmlnsNsUri)) return error.MalformedSheetXml;
            try scope.declare(scratch, p, try bindsNs(scratch, attr.value, isMceNs));
        }
    }
    if (std.mem.indexOfScalar(u8, qname, ':')) |c| {
        if (!scope.contains(qname[0..c])) return error.MalformedSheetXml;
    }
    var it2: AttrScan = .{ .rest = attrs };
    while (try it2.next()) |attr| {
        if (std.mem.eql(u8, attr.name, "xmlns")) continue;
        if (std.mem.startsWith(u8, attr.name, "xmlns:")) continue;
        if (std.mem.indexOfScalar(u8, attr.name, ':')) |c| {
            if (!scope.contains(attr.name[0..c])) return error.MalformedSheetXml;
            // `mc:Ignorable` (and the preserve lists) are inert for a
            // reader, but `mc:ProcessContent` marks a wrapper whose
            // children an MCE processor LIFTS — a chain of such
            // wrappers from a recognized slot down would project a
            // deep subtree into it (SEC-2201).
            if (scope.isMce(attr.name[0..c]) and
                std.mem.eql(u8, attr.name[c + 1 ..], "ProcessContent"))
            {
                return error.MalformedSheetXml;
            }
        }
    }
}

pub fn bindsNs(scratch: Allocator, raw: []const u8, comptime pred: fn ([]const u8) bool) Error!bool {
    // A binding that is not UTF-8, or that holds characters XML 1.0
    // forbids outright (`&#0;`, a literal C0 control, U+FFFE), is not
    // one the walk can rule out — each read as bound-to-foreign and
    // the block went opaque (Codex #215 r5 SEC-501, r7 SEC-701). The
    // XML-Char floor applies HERE, not to formula/record carriers,
    // whose decoded controls the JSON writer escapes.
    if (std.mem.indexOfScalar(u8, raw, '&') == null) {
        if (!std.unicode.utf8ValidateSlice(raw)) return error.MalformedSheetXml;
        if (!xmlCharsValid(raw)) return error.MalformedSheetXml;
        return pred(raw);
    }
    const decoded = formula_mod.decode.decodeEntities(scratch, raw) catch |e| switch (e) {
        error.OutOfMemory => return error.OutOfMemory,
        else => return error.MalformedSheetXml,
    };
    defer scratch.free(decoded);
    if (!std.unicode.utf8ValidateSlice(decoded)) return error.MalformedSheetXml;
    if (!xmlCharsValid(decoded)) return error.MalformedSheetXml;
    return pred(decoded);
}

/// XML 1.0 `Char`: #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD]
/// | [#x10000-#x10FFFF]. The input must already be valid UTF-8.
fn xmlCharsValid(s: []const u8) bool {
    var it = std.unicode.Utf8View.initUnchecked(s).iterator();
    while (it.nextCodepoint()) |cp| {
        const ok = cp == 0x9 or cp == 0xA or cp == 0xD or
            (cp >= 0x20 and cp <= 0xD7FF) or
            (cp >= 0xE000 and cp <= 0xFFFD) or
            (cp >= 0x10000 and cp <= 0x10FFFF);
        if (!ok) return false;
    }
    return true;
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
/// not the formula, the defined-names ruling), a rule spelling more
/// than the schema's three formulas, and — the walk has no MCE
/// processor — the MCE constructs whose projection could reach a
/// recognized slot: an MCE-bound element that is a direct child of
/// the root, a block or a rule, the `mc:ProcessContent` attribute
/// anywhere, a default binding of the MCE namespace. The `mc`
/// declaration, `mc:Ignorable`, and a deeper `mc:AlternateContent`
/// stay inert (SEC-2201).
/// The walker's nesting ceiling. Excel writes sheet parts a handful
/// of levels deep; 1024 is two orders of magnitude of headroom, and a
/// bound keeps a crafted part from growing the frame stack without
/// limit (Codex #215 r2 PERF-201).
pub const max_depth = 1024;

fn scanSheetRules(a: Allocator, xml: []const u8, out: *std.ArrayListUnmanaged(RawRule)) Error!RootKind {
    // The whole part must be UTF-8 before any byte-level scan — a
    // stray invalid byte inside a NAME otherwise turned an element
    // opaque instead of refusing (Codex #215 r10 SEC-1001).
    if (!std.unicode.utf8ValidateSlice(xml)) return error.MalformedSheetXml;
    const Kind = enum { root, barren_root, cf_block, cf_rule, other };
    const Frame = struct {
        name: []const u8,
        kind: Kind,
        main_default: bool,
        prefix_mark: usize,
        sqref: []const u8,
        rule: RawRule,
    };
    var frames: std.ArrayListUnmanaged(Frame) = .empty;
    defer frames.deinit(a);
    // The in-scope prefix bindings (Codex #215 r11 SEC-1101): each
    // frame records its watermark and pops its own declarations.
    var scope: PrefixScope = .{};
    defer scope.deinit(a);
    var decl_seen: std.StringHashMapUnmanaged(void) = .empty;
    defer decl_seen.deinit(a);

    var root_seen = false;
    var root_closed = false;
    var root_kind: RootKind = .worksheet;
    var i: usize = 0;
    while (std.mem.indexOfScalarPos(u8, xml, i, '<')) |lt| {
        // OUTSIDE the root only comments, PIs and whitespace may
        // appear (a BOM at byte zero included): stray character data
        // or a CDATA section there is not a well-formed document
        // (Codex #215 r11 REL-1102).
        if (frames.items.len == 0) {
            const gap_start: usize = if (i == 0 and std.mem.startsWith(u8, xml, "\xEF\xBB\xBF")) 3 else i;
            for (xml[gap_start..lt]) |c| {
                if (!isXmlWs(c)) return error.MalformedSheetXml;
            }
            if (lt + 9 <= xml.len and std.mem.eql(u8, xml[lt .. lt + 9], "<![CDATA[")) {
                return error.MalformedSheetXml;
            }
        }
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
            if (is_comment) {
                i = skipStrictComment(xml, lt) orelse return error.MalformedSheetXml;
                continue;
            }
        }
        if (lt + 1 < xml.len and xml[lt + 1] == '?') {
            if (isPrologXmlDecl(xml, lt)) {
                i = skipStrictXmlDecl(xml, lt) orelse return error.MalformedSheetXml;
                continue;
            }
            i = skipStrictPi(xml, lt) orelse return error.MalformedSheetXml;
            continue;
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
            scope.truncate(top.prefix_mark);
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
        if (!validQName(qname)) return error.MalformedSheetXml;
        const te = tagEnd(xml, j) orelse return error.MalformedSheetXml;
        const attrs = xml[j..te.attrs_end];
        if (root_closed) return error.MalformedSheetXml; // a second root

        // Namespace bookkeeping: the default declared here binds this
        // element too; the main namespace bound to a prefix is outside
        // the closed form and refuses (the host/SST predicate's rule).
        const prefix_mark = scope.mark();
        try enterElementScope(a, &scope, qname, attrs);
        var elem_main = if (frames.items.len == 0)
            false
        else
            frames.items[frames.items.len - 1].main_default;
        {
            var it: AttrScan = .{ .rest = attrs };
            // A namespace declaration repeated on one start tag —
            // default or prefixed, whichever value order — makes the
            // binding ambiguous; refuse before reading either value
            // (Codex #215 r7 SEC-703). Hashed and uncapped — a fixed
            // table rejected valid XML past its size (r13 REL-1301).
            defer resetDeclSeen(a, &decl_seen);
            while (try it.next()) |attr| {
                const is_decl = std.mem.eql(u8, attr.name, "xmlns") or
                    std.mem.startsWith(u8, attr.name, "xmlns:");
                if (is_decl) {
                    const g = try decl_seen.getOrPut(a, attr.name);
                    if (g.found_existing) return error.MalformedSheetXml;
                }
                if (std.mem.eql(u8, attr.name, "xmlns")) {
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
            .prefix_mark = prefix_mark,
            .sqref = "",
            .rule = undefined,
        };
        if (frames.items.len == 0) {
            root_seen = true;
            // A chartsheet or dialogsheet cannot carry conditional
            // formatting, so it contributes an empty inventory — but
            // it is still WALKED whole: an early return here would
            // bypass the second-root and unterminated-nesting checks,
            // so a truncated chartsheet or a `<worksheet/>` followed
            // by a rule-bearing second root would read as an empty
            // success (Codex #215 r2 REL-201). The root rule itself
            // is the family's (`classifySheetRoot`).
            const root = (try classifySheetRoot(a, qname, attrs, is_main, elem_main)) orelse
                return error.MalformedSheetXml;
            root_kind = switch (root) {
                .worksheet => .worksheet,
                .macrosheet => .macrosheet,
                .chartsheet, .dialogsheet => .barren,
            };
            frame.kind = if (root_kind == .barren) .barren_root else .root;
        } else {
            const parent = &frames.items[frames.items.len - 1];
            // An MCE-bound element AT a recognized slot is where
            // branch substitution could add or hide a block, a rule
            // or a formula the walk classifies — deeper MCE projects
            // in place inside an opaque subtree and stays walkable
            // (SEC-2201).
            if (std.mem.indexOfScalar(u8, qname, ':')) |c| {
                const slot = switch (parent.kind) {
                    .root, .cf_block, .cf_rule => true,
                    else => false,
                };
                if (slot and scope.isMce(qname[0..c])) return error.MalformedSheetXml;
            }
            if (parent.kind == .root and is_main and std.mem.eql(u8, qname, "conditionalFormatting")) {
                frame.kind = .cf_block;
                frame.sqref = (try uniqueAttr(a, attrs, "sqref")) orelse "";
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
                    .rule_type = (try uniqueAttr(a, attrs, "type")) orelse "",
                    .formulas = .{ "", "", "" },
                    .formula_count = 0,
                    .dxf_id = try numericAttr(a, try uniqueAttr(a, attrs, "dxfId")),
                    .priority = try numericAttr(a, try uniqueAttr(a, attrs, "priority")),
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
                _ = try uniqueAttr(a, attrs, "");
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
                scope.truncate(prefix_mark); // atomic element — its scope ends here
                continue;
            }
        }

        if (te.self_closing) {
            scope.truncate(prefix_mark);
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
    // Trailing character data after the last construct must be
    // whitespace too (REL-1102).
    for (xml[i..]) |c| {
        if (!isXmlWs(c)) return error.MalformedSheetXml;
    }
    if (!root_seen) return error.MalformedSheetXml;
    if (frames.items.len != 0) return error.MalformedSheetXml;
    return root_kind;
}

/// XML 1.0 comments forbid `--` inside the content and content
/// ending in `-` — a delimiter-only skip let
/// `<!-- bad -- <conditionalFormatting>…</conditionalFormatting> -->`
/// swallow a rule block into a malformed comment under exit 0 (Codex
/// #215 r15 REL-1501). Null = malformed or unterminated.
/// A strict PI skipper for the walkers: the target must be a
/// non-empty XML Name that is not `xml` in any case (the reserved
/// declaration target — the prolog's declaration is admitted
/// separately by position). A delimiter-only skip accepted
/// `<? <conditionalFormatting>…' as a PI and thinned the inventory
/// (Codex #215 r17 REL-1702). Null = malformed or unterminated.
pub fn skipStrictPi(xml: []const u8, lt: usize) ?usize {
    const end = std.mem.indexOfPos(u8, xml, lt + 2, "?>") orelse return null;
    const body = xml[lt + 2 .. end];
    var j: usize = 0;
    while (j < body.len and !isXmlWs(body[j])) j += 1;
    const target = body[0..j];
    // PITarget is an XML Name — colons are LEGAL there, unlike in the
    // namespace-constrained element/attribute QNames (Codex #215 r19
    // REL-1902).
    if (!validXmlName(target)) return null;
    if (target.len == 3 and
        (target[0] == 'x' or target[0] == 'X') and
        (target[1] == 'm' or target[1] == 'M') and
        (target[2] == 'l' or target[2] == 'L')) return null;
    return end + 2;
}

/// Strictly parse the document's XML declaration: a required
/// `version="1.0|1.1"`, then optionally `encoding` (a UTF-8 spelling
/// — the walkers already require UTF-8 bytes) and `standalone`
/// (`yes`/`no`), in that order, pseudo-attribute grammar only. The
/// positional predicate alone admitted `<?xml?>` and even rule-shaped
/// markup inside the declaration (Codex #215 r18 REL-1801). Null =
/// malformed.
pub fn skipStrictXmlDecl(xml: []const u8, lt: usize) ?usize {
    const end = std.mem.indexOfPos(u8, xml, lt + 2, "?>") orelse return null;
    if (lt + 5 > end) return null;
    const body = xml[lt + 5 .. end];
    var it: AttrScan = .{ .rest = body };
    const v = (it.next() catch return null) orelse return null;
    if (!std.mem.eql(u8, v.name, "version")) return null;
    if (!std.mem.eql(u8, v.value, "1.0") and !std.mem.eql(u8, v.value, "1.1")) return null;
    var saw_encoding = false;
    var saw_standalone = false;
    while (it.next() catch return null) |pa| {
        if (std.mem.eql(u8, pa.name, "encoding")) {
            if (saw_encoding or saw_standalone) return null;
            saw_encoding = true;
            if (!std.ascii.eqlIgnoreCase(pa.value, "UTF-8") and
                !std.ascii.eqlIgnoreCase(pa.value, "UTF8")) return null;
        } else if (std.mem.eql(u8, pa.name, "standalone")) {
            if (saw_standalone) return null;
            saw_standalone = true;
            if (!std.mem.eql(u8, pa.value, "yes") and !std.mem.eql(u8, pa.value, "no")) return null;
        } else {
            return null;
        }
    }
    return end + 2;
}

/// The document's XML declaration — admitted only as the very first
/// construct (a BOM aside).
pub fn isPrologXmlDecl(xml: []const u8, lt: usize) bool {
    const at_start = lt == 0 or
        (lt == 3 and std.mem.startsWith(u8, xml, "\xEF\xBB\xBF"));
    if (!at_start) return false;
    if (lt + 5 > xml.len or !std.mem.eql(u8, xml[lt .. lt + 5], "<?xml")) return false;
    return lt + 5 >= xml.len or isXmlWs(xml[lt + 5]) or xml[lt + 5] == '?';
}

pub fn skipStrictComment(xml: []const u8, lt: usize) ?usize {
    const end = std.mem.indexOfPos(u8, xml, lt + 4, "-->") orelse return null;
    const body = xml[lt + 4 .. end];
    if (std.mem.indexOf(u8, body, "--") != null) return null;
    if (body.len > 0 and body[body.len - 1] == '-') return null;
    return end + 3;
}

pub fn isXmlWs(c: u8) bool {
    return c == ' ' or c == '\t' or c == '\n' or c == '\r';
}

pub fn isTagBoundary(c: u8) bool {
    return isXmlWs(c) or c == '/' or c == '>';
}

pub const TagEnd = struct { attrs_end: usize, after_gt: usize, self_closing: bool };

/// Quote-aware scan from the end of the tag name to the tag's `>` —
/// an attribute value may contain `>` (Codex #215 r1 REL-103 caught
/// the rewriter's bare scan on the same shape).
pub fn tagEnd(xml: []const u8, attrs_start: usize) ?TagEnd {
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

pub const Attr = struct { name: []const u8, value: []const u8 };

/// Strict attribute tokenizer: `name = "value"` with XML whitespace
/// allowed around `=` (valid XML the exact-needle scan reported as
/// absent — Codex #215 r1 REL-102), either quote, nothing else.
/// Names follow XML's Name grammar (so `<cfRule/x="y"` — a slash
/// taken as a QName boundary — cannot smuggle a pseudo-attribute
/// past the tokenizer and classify as rule machinery; Codex #215 r5
/// REL-502), and a raw `<` in any value is ill-formed XML. A region
/// that does not tokenize refuses the part.
pub const AttrScan = struct {
    rest: []const u8,

    pub fn next(self: *AttrScan) error{MalformedSheetXml}!?Attr {
        const s = self.rest;
        var i: usize = 0;
        // XML requires whitespace BETWEEN attributes (and between the
        // element name and the first — the region handed here always
        // begins at that boundary): a quoted value directly followed
        // by another name (`sqref="A1"xmlns="urn:foreign"`) is
        // malformed, and tokenizing it let an unseparated namespace
        // rebinding turn a rule block — or a workbook `<sheet>` entry
        // — opaque under exit 0 (Codex #215 r9 SEC-901).
        if (s.len != 0 and !isXmlWs(s[0])) return error.MalformedSheetXml;
        while (i < s.len and isXmlWs(s[i])) i += 1;
        if (i >= s.len) {
            self.rest = s[s.len..];
            return null;
        }
        const name_start = i;
        while (i < s.len and s[i] != '=' and !isXmlWs(s[i])) i += 1;
        if (i == name_start) return error.MalformedSheetXml;
        const name = s[name_start..i];
        if (!validQName(name)) return error.MalformedSheetXml;
        while (i < s.len and isXmlWs(s[i])) i += 1;
        if (i >= s.len or s[i] != '=') return error.MalformedSheetXml;
        i += 1;
        while (i < s.len and isXmlWs(s[i])) i += 1;
        if (i >= s.len or (s[i] != '"' and s[i] != '\'')) return error.MalformedSheetXml;
        const q = s[i];
        i += 1;
        const v_start = i;
        while (i < s.len and s[i] != q) : (i += 1) {
            if (s[i] == '<') return error.MalformedSheetXml;
        }
        if (i >= s.len) return error.MalformedSheetXml;
        const value = s[v_start..i];
        self.rest = s[i + 1 ..];
        return .{ .name = name, .value = value };
    }
};

/// A QName: an NCName, or NCName `:` NCName — at most one colon,
/// neither side empty. `xmlns:=` (an empty prefix) and `:id` (an
/// empty first half) are malformed namespace XML that slipped a
/// colon-agnostic NameChar check and reopened the relationship-
/// identity boundary (Codex #215 r8 SEC-801).
pub fn validQName(name: []const u8) bool {
    if (name.len == 0) return false;
    if (std.mem.indexOfScalar(u8, name, ':')) |c| {
        if (c == 0 or c == name.len - 1) return false;
        if (std.mem.indexOfScalarPos(u8, name, c + 1, ':') != null) return false;
        return validNcName(name[0..c]) and validNcName(name[c + 1 ..]);
    }
    return validNcName(name);
}

/// An NCName by XML 1.0's CODE-POINT ranges — a `>= 0x80` byte
/// shortcut admitted `\xff` and U+00A0 in names, and a matched pair
/// like `<conditionalFormatting\xff>` then classified as opaque and
/// hid the rule subtree inside it (Codex #215 r10 SEC-1001).
fn validNcName(name: []const u8) bool {
    if (name.len == 0) return false;
    if (!std.unicode.utf8ValidateSlice(name)) return false;
    var it = std.unicode.Utf8View.initUnchecked(name).iterator();
    var first = true;
    while (it.nextCodepoint()) |cp| {
        if (first) {
            if (!ncNameStartCp(cp)) return false;
            first = false;
        } else if (!ncNameCharCp(cp)) {
            return false;
        }
    }
    return true;
}

/// An XML 1.0 `Name` — the NCName code points PLUS `:` anywhere.
/// Only PI targets use this; element and attribute names stay under
/// the namespace-constrained QName rule.
fn validXmlName(name: []const u8) bool {
    if (name.len == 0) return false;
    if (!std.unicode.utf8ValidateSlice(name)) return false;
    var it = std.unicode.Utf8View.initUnchecked(name).iterator();
    var first = true;
    while (it.nextCodepoint()) |cp| {
        if (first) {
            if (cp != ':' and !ncNameStartCp(cp)) return false;
            first = false;
        } else if (cp != ':' and !ncNameCharCp(cp)) {
            return false;
        }
    }
    return true;
}

fn ncNameStartCp(cp: u21) bool {
    return (cp >= 'A' and cp <= 'Z') or cp == '_' or (cp >= 'a' and cp <= 'z') or
        (cp >= 0xC0 and cp <= 0xD6) or (cp >= 0xD8 and cp <= 0xF6) or
        (cp >= 0xF8 and cp <= 0x2FF) or (cp >= 0x370 and cp <= 0x37D) or
        (cp >= 0x37F and cp <= 0x1FFF) or (cp >= 0x200C and cp <= 0x200D) or
        (cp >= 0x2070 and cp <= 0x218F) or (cp >= 0x2C00 and cp <= 0x2FEF) or
        (cp >= 0x3001 and cp <= 0xD7FF) or (cp >= 0xF900 and cp <= 0xFDCF) or
        (cp >= 0xFDF0 and cp <= 0xFFFD) or (cp >= 0x10000 and cp <= 0xEFFFF);
}

fn ncNameCharCp(cp: u21) bool {
    return ncNameStartCp(cp) or cp == '-' or cp == '.' or (cp >= '0' and cp <= '9') or
        cp == 0xB7 or (cp >= 0x300 and cp <= 0x36F) or (cp >= 0x203F and cp <= 0x2040);
}

/// Look up `key` on a rule-machinery tag while refusing a duplicate
/// of ANY attribute name there (the S7b-4 "a name twice on one start
/// tag" rule, scoped to the elements this read depends on). Pass an
/// empty `key` to run the duplicate check alone. Hashed and uncapped
/// — a fixed table rejected valid XML past its size (Codex #215 r13
/// REL-1301).
pub fn uniqueAttr(a: Allocator, attrs: []const u8, key: []const u8) Error!?[]const u8 {
    var seen: std.StringHashMapUnmanaged(void) = .empty;
    defer seen.deinit(a);
    var found: ?[]const u8 = null;
    var it: AttrScan = .{ .rest = attrs };
    while (try it.next()) |attr| {
        const g = try seen.getOrPut(a, attr.name);
        if (g.found_existing) return error.MalformedSheetXml;
        if (key.len != 0 and std.mem.eql(u8, attr.name, key)) found = attr.value;
    }
    return found;
}

/// A numeric rule attribute: XML character references resolve FIRST
/// (`dxfId="&#49;"` is 1 — the `styles_xml` numeric precedent; a
/// reference that does not decode refuses the inventory), THEN the
/// digit-only/u32-or-null lexical rule applies to the decoded value
/// (`+1`, `1_0` and overflow read as absent, the written-but-invalid
/// convention `tableHeaderRowCount` set) — Codex #215 r1 REL-104,
/// r18 REL-1802.
pub fn numericAttr(scratch: Allocator, raw_opt: ?[]const u8) Error!?u32 {
    const raw = raw_opt orelse return null;
    if (std.mem.indexOfScalar(u8, raw, '&') == null) return digitOnlyU32(raw);
    const decoded = formula_mod.decode.decodeEntities(scratch, raw) catch |e| switch (e) {
        error.OutOfMemory => return error.OutOfMemory,
        else => return error.MalformedSheetXml,
    };
    defer scratch.free(decoded);
    return digitOnlyU32(decoded);
}

pub fn digitOnlyU32(s: []const u8) ?u32 {
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

/// One sheet identity as the STRICT workbook read established it.
/// Slices borrow the workbook part's bytes (stable — the store caches
/// parts).
const StrictSheet = struct { name: []const u8, rid: []const u8 };

/// The AUTHORITATIVE sheet inventory: walk `xl/workbook.xml` with the
/// sheet-part walker's discipline (DOCTYPE refusal, decoded namespace
/// bindings, depth + namespace classification, QName grammar). A
/// `<sheet>` is an entry only as a main-namespace direct child of the
/// ONE main-namespace `<sheets>` child of the root; every entry must
/// carry the schema's required `name` and `sheetId` plus exactly one
/// id reference under a ROOT-declared relationships prefix. A ghost
/// under `<extLst>` is not an entry (SEC-502); a carrier-less real
/// entry refuses; and a standards-valid alternate prefix (`q:id`
/// under `xmlns:q="…relationships"`) IS an entry — sheet identities
/// come from this read, not from the lenient projection that only
/// spells `r:id` (Codex #215 r8 REL-801). The list must hold at least
/// one entry (CT_Sheets minOccurs=1): a missing `<sheets>` and an
/// empty `<sheets/>` are the same sheetless workbook, refused rather
/// than served as an empty inventory (REL-602, #219 r1 REL-101). The
/// same walk records the root's `<calcPr>` slot into `calc` for the
/// calc-props read — a capture only, no verdict of its own here
/// (`CalcPrSlot`).
fn scanWorkbookSheets(gpa: Allocator, xml: []const u8, out: *std.ArrayListUnmanaged(StrictSheet), calc: *CalcPrSlot) WorkbookError!void {
    if (!std.unicode.utf8ValidateSlice(xml)) return error.MalformedWorkbookXml;
    // Entries are counted relative to what the caller already holds,
    // so a reused list cannot vouch for an empty walk.
    const entries_before = out.items.len;
    const Kind = enum { root, sheets_wrap, other };
    const Frame = struct { name: []const u8, kind: Kind, main_default: bool, prefix_mark: usize, mce_chain: bool };
    var frames: std.ArrayListUnmanaged(Frame) = .empty;
    defer frames.deinit(gpa);
    var scope: PrefixScope = .{};
    defer scope.deinit(gpa);
    var decl_seen: std.StringHashMapUnmanaged(void) = .empty;
    defer decl_seen.deinit(gpa);

    // Prefixes the ROOT binds to a relationships namespace — the only
    // spellings an entry's `r:id` may use (a literal `r:id` under a
    // foreign or absent binding is not a relationship reference —
    // Codex #215 r6 SEC-601). A relationships binding below the root,
    // or a redeclaration of a collected prefix, is outside the closed
    // form and refuses, so the root's list stays authoritative.
    var rel_prefixes: std.StringHashMapUnmanaged(void) = .empty;
    defer rel_prefixes.deinit(gpa);

    var root_seen = false;
    var root_closed = false;
    var wraps_seen: usize = 0;
    var i: usize = 0;
    while (std.mem.indexOfScalarPos(u8, xml, i, '<')) |lt| {
        if (frames.items.len == 0) {
            const gap_start: usize = if (i == 0 and std.mem.startsWith(u8, xml, "\xEF\xBB\xBF")) 3 else i;
            for (xml[gap_start..lt]) |c| {
                if (!isXmlWs(c)) return error.MalformedWorkbookXml;
            }
            if (lt + 9 <= xml.len and std.mem.eql(u8, xml[lt .. lt + 9], "<![CDATA[")) {
                return error.MalformedWorkbookXml;
            }
        }
        if (lt + 1 < xml.len and xml[lt + 1] == '!') {
            const is_comment = lt + 4 <= xml.len and std.mem.eql(u8, xml[lt .. lt + 4], "<!--");
            const is_cdata = lt + 9 <= xml.len and std.mem.eql(u8, xml[lt .. lt + 9], "<![CDATA[");
            if (!is_comment and !is_cdata) return error.MalformedWorkbookXml;
            if (is_comment) {
                i = skipStrictComment(xml, lt) orelse return error.MalformedWorkbookXml;
                continue;
            }
        }
        if (lt + 1 < xml.len and xml[lt + 1] == '?') {
            if (isPrologXmlDecl(xml, lt)) {
                i = skipStrictXmlDecl(xml, lt) orelse return error.MalformedWorkbookXml;
                continue;
            }
            i = skipStrictPi(xml, lt) orelse return error.MalformedWorkbookXml;
            continue;
        }
        const skip_to = workbook_xml.skipNonElement(xml, lt) catch return error.MalformedWorkbookXml;
        if (skip_to != lt) {
            i = skip_to;
            continue;
        }
        if (lt + 1 >= xml.len) return error.MalformedWorkbookXml;

        if (xml[lt + 1] == '/') {
            var j = lt + 2;
            const name_start = j;
            while (j < xml.len and xml[j] != '>' and !isXmlWs(xml[j])) j += 1;
            const name = xml[name_start..j];
            while (j < xml.len and isXmlWs(xml[j])) j += 1;
            if (j >= xml.len or xml[j] != '>') return error.MalformedWorkbookXml;
            if (frames.items.len == 0) return error.MalformedWorkbookXml;
            const top = frames.items[frames.items.len - 1];
            frames.items.len -= 1;
            if (!std.mem.eql(u8, top.name, name)) return error.MalformedWorkbookXml;
            scope.truncate(top.prefix_mark);
            if (top.kind == .root) root_closed = true;
            i = j + 1;
            continue;
        }

        const name_start = lt + 1;
        var j = name_start;
        while (j < xml.len and !isTagBoundary(xml[j])) j += 1;
        if (j == name_start or j >= xml.len) return error.MalformedWorkbookXml;
        const qname = xml[name_start..j];
        if (!validQName(qname)) return error.MalformedWorkbookXml;
        const te = tagEnd(xml, j) orelse return error.MalformedWorkbookXml;
        const attrs = xml[j..te.attrs_end];
        if (root_closed) return error.MalformedWorkbookXml;

        const prefix_mark = scope.mark();
        enterElementScope(gpa, &scope, qname, attrs) catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            else => return error.MalformedWorkbookXml,
        };
        var elem_main = if (frames.items.len == 0)
            false
        else
            frames.items[frames.items.len - 1].main_default;
        {
            var it: AttrScan = .{ .rest = attrs };
            // Duplicate declarations refuse before either value is
            // interpreted — a correct-then-foreign (or reverse)
            // duplicate `xmlns:r` left the prefix collected under an
            // ambiguous binding (Codex #215 r7 SEC-703). Hashed and
            // uncapped (r13 REL-1301).
            defer resetDeclSeen(gpa, &decl_seen);
            while (it.next() catch return error.MalformedWorkbookXml) |attr| {
                const is_decl = std.mem.eql(u8, attr.name, "xmlns") or
                    std.mem.startsWith(u8, attr.name, "xmlns:");
                if (is_decl) {
                    const g = try decl_seen.getOrPut(gpa, attr.name);
                    if (g.found_existing) return error.MalformedWorkbookXml;
                }
                if (std.mem.eql(u8, attr.name, "xmlns")) {
                    elem_main = bindsMainNs(gpa, attr.value) catch |e| switch (e) {
                        error.OutOfMemory => return error.OutOfMemory,
                        else => return error.MalformedWorkbookXml,
                    };
                } else if (std.mem.startsWith(u8, attr.name, "xmlns:")) {
                    const main = bindsMainNs(gpa, attr.value) catch |e| switch (e) {
                        error.OutOfMemory => return error.OutOfMemory,
                        else => return error.MalformedWorkbookXml,
                    };
                    if (main) return error.MalformedWorkbookXml;
                    const prefix = attr.name["xmlns:".len..];
                    const is_rel = bindsRelNs(gpa, attr.value) catch |e| switch (e) {
                        error.OutOfMemory => return error.OutOfMemory,
                        else => return error.MalformedWorkbookXml,
                    };
                    if (frames.items.len == 0) {
                        if (is_rel) {
                            const g = try rel_prefixes.getOrPut(gpa, prefix);
                            if (g.found_existing) return error.MalformedWorkbookXml;
                        }
                    } else {
                        if (is_rel) return error.MalformedWorkbookXml;
                        if (rel_prefixes.contains(prefix)) return error.MalformedWorkbookXml;
                    }
                }
            }
        }
        const prefixed = std.mem.indexOfScalar(u8, qname, ':') != null;
        const is_main = !prefixed and elem_main;

        var kind: Kind = .other;
        // Whether THIS element's children would land at the root slot
        // under an MCE processor: true for the root itself, and for
        // an MCE-bound element whose every ancestor below the root is
        // MCE-bound too (the `mc:AlternateContent` / `mc:Choice` /
        // `mc:Fallback` chain a processor substitutes in place). Any
        // ordinary wrapper on the path — `<extLst>`, a foreign
        // container, a plain element inside a Choice — breaks the
        // chain: after projection it, not its content, sits at the
        // root (Codex #218 r1 S3B-REL-102).
        var mce_chain = false;
        if (frames.items.len == 0) {
            if (!is_main or !std.mem.eql(u8, qname, "workbook")) return error.MalformedWorkbookXml;
            root_seen = true;
            kind = .root;
            mce_chain = true;
            if (te.self_closing) {
                root_closed = true;
            }
        } else {
            const parent = frames.items[frames.items.len - 1];
            // Inside `<sheets>` an MCE branch could hide a `<sheet>`
            // from the identity walk; at workbook level the real
            // `mc:AlternateContent` (absPath et al.) stays opaque —
            // hiding `<sheets>` itself trips the exactly-one-wrapper
            // rule below (SEC-2201).
            if (prefixed) {
                const c = std.mem.indexOfScalar(u8, qname, ':').?;
                if (scope.isMce(qname[0..c])) {
                    if (parent.kind == .sheets_wrap) return error.MalformedWorkbookXml;
                    mce_chain = parent.mce_chain;
                }
            }
            if (parent.kind == .root and is_main and std.mem.eql(u8, qname, "sheets")) {
                kind = .sheets_wrap;
                wraps_seen += 1;
                if (wraps_seen > 1) return error.MalformedWorkbookXml;
            } else if (is_main and std.mem.eql(u8, qname, "calcPr")) {
                // The calc-props slot: counted at the root, flagged
                // when only MCE elements stand between it and the root
                // (the one path the processor the walk lacks could
                // project it into the slot by), opaque elsewhere
                // (`<extLst>` decoys are not entries — SEC-502).
                if (parent.kind == .root) {
                    calc.count += 1;
                    if (calc.count == 1) calc.attrs = attrs;
                } else if (parent.mce_chain) {
                    calc.mce_shadowed = true;
                }
            } else if (parent.kind == .sheets_wrap and is_main and std.mem.eql(u8, qname, "sheet")) {
                const name_attr = ((try wbAttr(gpa, attrs, "name")) orelse return error.MalformedWorkbookXml);
                if ((try wbAttr(gpa, attrs, "sheetId")) == null) return error.MalformedWorkbookXml;
                // The relationship reference is an attr named
                // `<p>:id` under a ROOT-declared relationships prefix
                // — exactly one (SEC-601).
                var rid_attr: ?[]const u8 = null;
                {
                    var it2: AttrScan = .{ .rest = attrs };
                    while (it2.next() catch return error.MalformedWorkbookXml) |attr2| {
                        // A DECLARATION is never an id candidate —
                        // `xmlns:id="rId1"` must not read as `<p>:id`
                        // under any authorized prefix (SEC-1101).
                        if (std.mem.eql(u8, attr2.name, "xmlns") or
                            std.mem.startsWith(u8, attr2.name, "xmlns:")) continue;
                        if (attr2.name.len > 3 and std.mem.endsWith(u8, attr2.name, ":id")) {
                            const p = attr2.name[0 .. attr2.name.len - 3];
                            if (rel_prefixes.contains(p)) {
                                if (rid_attr != null) return error.MalformedWorkbookXml;
                                rid_attr = attr2.value;
                            }
                        }
                    }
                }
                const rid = rid_attr orelse return error.MalformedWorkbookXml;
                try out.append(gpa, .{ .name = name_attr, .rid = rid });
            }
        }

        if (te.self_closing) {
            scope.truncate(prefix_mark);
        } else {
            if (frames.items.len >= max_depth) return error.MalformedWorkbookXml;
            try frames.append(gpa, .{ .name = qname, .kind = kind, .main_default = elem_main, .prefix_mark = prefix_mark, .mce_chain = mce_chain });
        }
        i = te.after_gt;
    }
    for (xml[i..]) |c| {
        if (!isXmlWs(c)) return error.MalformedWorkbookXml;
    }
    if (!root_seen) return error.MalformedWorkbookXml;
    if (frames.items.len != 0) return error.MalformedWorkbookXml;
    // Exactly one <sheets> wrapper — zero would make a sheetless
    // workbook an empty success with no inventory ever established
    // (Codex #215 r6 REL-602).
    if (wraps_seen != 1) return error.MalformedWorkbookXml;
    // … and at least one entry in it: CT_Sheets is minOccurs=1, and an
    // empty `<sheets/>` is the same sheetless workbook REL-602 refuses,
    // spelled with the wrapper present (Codex #219 r1 S3B-REL-101 —
    // every read over this inventory, and the C exports built on them,
    // promised a non-empty stream).
    if (out.items.len == entries_before) return error.MalformedWorkbookXml;
}

/// The `<calcPr>` slot as the strict workbook read found it, for the
/// sheet-props read (`sheet_props_ndjson.collectCalc`): the attribute
/// region of the first main-namespace `<calcPr>` that is a direct
/// child of the root (borrowing the cached part's bytes), how many
/// such elements the root carries (the schema allows one; the consumer
/// refuses more), and whether a main-namespace `calcPr` sits directly
/// under an unbroken chain of MCE-bound elements from the root — the
/// one path an MCE processor could project it INTO the root slot by,
/// which the walk (no MCE processor) cannot rule in or out (SEC-2201's
/// closed form). A `calcPr` below any ordinary wrapper — `<extLst>`,
/// a foreign container, a plain element inside a Choice — is neither:
/// not an entry, not a shadow (Codex #218 r1 S3B-REL-102). This
/// module and the anchors collector ignore the slot.
pub const CalcPrSlot = struct {
    attrs: ?[]const u8 = null,
    count: u32 = 0,
    mce_shadowed: bool = false,
};

/// The strict `<calcPr>` read of `xl/workbook.xml` — the walk
/// `resolveSheets` runs (so the `<sheets>` structure it polices holds
/// here too), without resolving the sheet parts: the calc-props read
/// has no per-sheet dimension and pays for none. The slot's slices
/// borrow the store's cached part.
pub fn scanCalcPr(gpa: Allocator, wb: *workbook_mod.Workbook) InventoryError!CalcPrSlot {
    const wb_part = (wb.store.part("xl/workbook.xml") catch |e| switch (e) {
        error.OutOfMemory => return error.OutOfMemory,
        error.ZipBombSuspected => return error.ZipBombSuspected,
        else => return error.MalformedWorkbookXml,
    }) orelse return error.MalformedWorkbookXml;
    var strict_sheets: std.ArrayListUnmanaged(StrictSheet) = .empty;
    defer strict_sheets.deinit(gpa);
    var slot: CalcPrSlot = .{};
    try scanWorkbookSheets(gpa, wb_part.bytes, &strict_sheets, &slot);
    return slot;
}

/// `uniqueAttr` with the workbook part's error name: a duplicate or
/// malformed region reads as "absent" for the caller's orelse-refusal
/// — but OutOfMemory PROPAGATES (the allocation-failure sweep caught
/// a `catch null` swallowing it into a spurious refusal).
fn wbAttr(a: Allocator, attrs: []const u8, key: []const u8) WorkbookError!?[]const u8 {
    return uniqueAttr(a, attrs, key) catch |e| switch (e) {
        error.OutOfMemory => return error.OutOfMemory,
        else => return null,
    };
}

/// SEC-601 verifier, second half: the store's lexical relationship
/// list for `xl/workbook.xml` is only trusted after a strict read of
/// `xl/_rels/workbook.xml.rels` agrees with it — OPC-namespace root,
/// `Relationship` elements only as its direct children (a nested or
/// foreign-slot decoy refuses), required `Id`/`Type`/`Target` (dup
/// attrs refuse), an exact `TargetMode` spelling when present, unique
/// STRICTLY-DECODED ids, and the store's decoded id sequence equal to
/// the strict one — so `relById`'s first match is THE match.
fn verifyWorkbookRels(gpa: Allocator, xml: []const u8, rels: []const store_mod.Relationship) WorkbookError!void {
    if (!std.unicode.utf8ValidateSlice(xml)) return error.MalformedWorkbookXml;
    const Kind = enum { root, other };
    const Frame = struct { name: []const u8, kind: Kind, pkg_default: bool, prefix_mark: usize };
    var frames: std.ArrayListUnmanaged(Frame) = .empty;
    defer frames.deinit(gpa);
    var scope: PrefixScope = .{};
    defer scope.deinit(gpa);
    var decl_seen: std.StringHashMapUnmanaged(void) = .empty;
    defer decl_seen.deinit(gpa);
    // Hashed — the linear seen-scan was quadratic in the relationship
    // count (Codex #215 r8 PERF-801). Keys are owned decoded ids.
    var ids: std.StringHashMapUnmanaged(void) = .empty;
    defer {
        var key_it = ids.keyIterator();
        while (key_it.next()) |k| gpa.free(k.*);
        ids.deinit(gpa);
    }

    var root_seen = false;
    var root_closed = false;
    var entry_idx: usize = 0;
    var i: usize = 0;
    while (std.mem.indexOfScalarPos(u8, xml, i, '<')) |lt| {
        if (frames.items.len == 0) {
            const gap_start: usize = if (i == 0 and std.mem.startsWith(u8, xml, "\xEF\xBB\xBF")) 3 else i;
            for (xml[gap_start..lt]) |c| {
                if (!isXmlWs(c)) return error.MalformedWorkbookXml;
            }
            if (lt + 9 <= xml.len and std.mem.eql(u8, xml[lt .. lt + 9], "<![CDATA[")) {
                return error.MalformedWorkbookXml;
            }
        }
        if (lt + 1 < xml.len and xml[lt + 1] == '!') {
            const is_comment = lt + 4 <= xml.len and std.mem.eql(u8, xml[lt .. lt + 4], "<!--");
            const is_cdata = lt + 9 <= xml.len and std.mem.eql(u8, xml[lt .. lt + 9], "<![CDATA[");
            if (!is_comment and !is_cdata) return error.MalformedWorkbookXml;
            if (is_comment) {
                i = skipStrictComment(xml, lt) orelse return error.MalformedWorkbookXml;
                continue;
            }
        }
        if (lt + 1 < xml.len and xml[lt + 1] == '?') {
            if (isPrologXmlDecl(xml, lt)) {
                i = skipStrictXmlDecl(xml, lt) orelse return error.MalformedWorkbookXml;
                continue;
            }
            i = skipStrictPi(xml, lt) orelse return error.MalformedWorkbookXml;
            continue;
        }
        const skip_to = workbook_xml.skipNonElement(xml, lt) catch return error.MalformedWorkbookXml;
        if (skip_to != lt) {
            i = skip_to;
            continue;
        }
        if (lt + 1 >= xml.len) return error.MalformedWorkbookXml;

        if (xml[lt + 1] == '/') {
            var j = lt + 2;
            const name_start = j;
            while (j < xml.len and xml[j] != '>' and !isXmlWs(xml[j])) j += 1;
            const name = xml[name_start..j];
            while (j < xml.len and isXmlWs(xml[j])) j += 1;
            if (j >= xml.len or xml[j] != '>') return error.MalformedWorkbookXml;
            if (frames.items.len == 0) return error.MalformedWorkbookXml;
            const top = frames.items[frames.items.len - 1];
            frames.items.len -= 1;
            if (!std.mem.eql(u8, top.name, name)) return error.MalformedWorkbookXml;
            scope.truncate(top.prefix_mark);
            if (top.kind == .root) root_closed = true;
            i = j + 1;
            continue;
        }

        const name_start = lt + 1;
        var j = name_start;
        while (j < xml.len and !isTagBoundary(xml[j])) j += 1;
        if (j == name_start or j >= xml.len) return error.MalformedWorkbookXml;
        const qname = xml[name_start..j];
        if (!validQName(qname)) return error.MalformedWorkbookXml;
        const te = tagEnd(xml, j) orelse return error.MalformedWorkbookXml;
        const attrs = xml[j..te.attrs_end];
        if (root_closed) return error.MalformedWorkbookXml;

        const prefix_mark = scope.mark();
        enterElementScope(gpa, &scope, qname, attrs) catch |e| switch (e) {
            error.OutOfMemory => return error.OutOfMemory,
            else => return error.MalformedWorkbookXml,
        };
        var elem_pkg = if (frames.items.len == 0)
            false
        else
            frames.items[frames.items.len - 1].pkg_default;
        {
            var it: AttrScan = .{ .rest = attrs };
            defer resetDeclSeen(gpa, &decl_seen);
            while (it.next() catch return error.MalformedWorkbookXml) |attr| {
                const is_decl = std.mem.eql(u8, attr.name, "xmlns") or
                    std.mem.startsWith(u8, attr.name, "xmlns:");
                if (is_decl) {
                    const g = try decl_seen.getOrPut(gpa, attr.name);
                    if (g.found_existing) return error.MalformedWorkbookXml;
                }
                if (std.mem.eql(u8, attr.name, "xmlns")) {
                    elem_pkg = bindsNs(gpa, attr.value, isPkgRelsNs) catch |e| switch (e) {
                        error.OutOfMemory => return error.OutOfMemory,
                        else => return error.MalformedWorkbookXml,
                    };
                } else if (std.mem.startsWith(u8, attr.name, "xmlns:")) {
                    const pkg = bindsNs(gpa, attr.value, isPkgRelsNs) catch |e| switch (e) {
                        error.OutOfMemory => return error.OutOfMemory,
                        else => return error.MalformedWorkbookXml,
                    };
                    if (pkg) return error.MalformedWorkbookXml;
                }
            }
        }
        const prefixed = std.mem.indexOfScalar(u8, qname, ':') != null;
        const is_pkg = !prefixed and elem_pkg;

        var kind: Kind = .other;
        if (frames.items.len == 0) {
            if (!is_pkg or !std.mem.eql(u8, qname, "Relationships")) return error.MalformedWorkbookXml;
            root_seen = true;
            kind = .root;
            if (te.self_closing) root_closed = true;
        } else if (std.mem.eql(u8, qname, "Relationship")) {
            // A Relationship-shaped element ANYWHERE but the schema
            // slot is a decoy the store's lexical scan may have
            // counted — refuse rather than reconcile.
            const parent = frames.items[frames.items.len - 1];
            if (parent.kind != .root or !is_pkg) return error.MalformedWorkbookXml;
            const id_raw = ((try wbAttr(gpa, attrs, "Id")) orelse return error.MalformedWorkbookXml);
            const type_raw = ((try wbAttr(gpa, attrs, "Type")) orelse return error.MalformedWorkbookXml);
            const target_raw = ((try wbAttr(gpa, attrs, "Target")) orelse return error.MalformedWorkbookXml);
            if (entry_idx >= rels.len) return error.MalformedWorkbookXml;
            // `Type` and `Target` verify by the STRICT decode too and
            // must equal what the lenient store cached — its wider
            // decoder (`&#X65;`) must not be the resolution's only
            // reading of a carrier (Codex #215 r7 SEC-702) — and the
            // `TargetMode` spelling must agree with the cached enum.
            {
                const type_dec = decodeRelText(gpa, type_raw) catch |e| switch (e) {
                    error.OutOfMemory => return error.OutOfMemory,
                    else => return error.MalformedWorkbookXml,
                };
                defer gpa.free(type_dec);
                if (!std.mem.eql(u8, type_dec, rels[entry_idx].type)) return error.MalformedWorkbookXml;
            }
            {
                const target_dec = decodeRelText(gpa, target_raw) catch |e| switch (e) {
                    error.OutOfMemory => return error.OutOfMemory,
                    else => return error.MalformedWorkbookXml,
                };
                defer gpa.free(target_dec);
                if (!std.mem.eql(u8, target_dec, rels[entry_idx].target)) return error.MalformedWorkbookXml;
            }
            const cached_external = rels[entry_idx].target_mode == .external;
            if (try wbAttr(gpa, attrs, "TargetMode")) |mode| {
                // Decoded like Id/Type/Target — `Inter&#110;al` is a
                // valid spelling of the enum and a raw compare refused
                // it (Codex #215 r23 REL-2301).
                const mode_dec = decodeRelText(gpa, mode) catch |e| switch (e) {
                    error.OutOfMemory => return error.OutOfMemory,
                    else => return error.MalformedWorkbookXml,
                };
                defer gpa.free(mode_dec);
                if (std.mem.eql(u8, mode_dec, "External")) {
                    if (!cached_external) return error.MalformedWorkbookXml;
                } else if (std.mem.eql(u8, mode_dec, "Internal")) {
                    if (cached_external) return error.MalformedWorkbookXml;
                } else {
                    return error.MalformedWorkbookXml;
                }
            } else if (cached_external) {
                return error.MalformedWorkbookXml;
            }
            const id = decodeRelText(gpa, id_raw) catch |e| switch (e) {
                error.OutOfMemory => return error.OutOfMemory,
                else => return error.MalformedWorkbookXml,
            };
            errdefer gpa.free(id);
            if (!std.mem.eql(u8, id, rels[entry_idx].id)) return error.MalformedWorkbookXml;
            const g = try ids.getOrPut(gpa, id);
            if (g.found_existing) return error.MalformedWorkbookXml;
            entry_idx += 1;
        }

        if (te.self_closing) {
            scope.truncate(prefix_mark);
        } else {
            if (frames.items.len >= max_depth) return error.MalformedWorkbookXml;
            try frames.append(gpa, .{ .name = qname, .kind = kind, .pkg_default = elem_pkg, .prefix_mark = prefix_mark });
        }
        i = te.after_gt;
    }
    for (xml[i..]) |c| {
        if (!isXmlWs(c)) return error.MalformedWorkbookXml;
    }
    if (!root_seen) return error.MalformedWorkbookXml;
    if (frames.items.len != 0) return error.MalformedWorkbookXml;
    if (entry_idx != rels.len) return error.MalformedWorkbookXml;
}

/// Strict entities-only decode for relationship ids in the verifier —
/// the sheet-side carrier decode with the workbook error name folded
/// by the caller.
fn decodeRelText(a: Allocator, raw: []const u8) WorkbookError![]u8 {
    if (std.mem.indexOfScalar(u8, raw, '<') != null) return error.MalformedWorkbookXml;
    const decoded = formula_mod.decode.decodeEntities(a, raw) catch |e| switch (e) {
        error.OutOfMemory => return error.OutOfMemory,
        else => return error.MalformedWorkbookXml,
    };
    errdefer a.free(decoded);
    if (!std.unicode.utf8ValidateSlice(decoded)) return error.MalformedWorkbookXml;
    if (!xmlCharsValid(decoded)) return error.MalformedWorkbookXml;
    return decoded;
}

fn decodeSheetName(a: Allocator, raw: []const u8) WorkbookError![]u8 {
    const decoded = formula_mod.decode.decodeAt(a, .sheet_name, raw) catch |e| switch (e) {
        error.OutOfMemory => return error.OutOfMemory,
        else => return error.MalformedWorkbookXml,
    };
    errdefer a.free(decoded);
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
pub fn decodeRuleText(a: Allocator, raw: []const u8) Error![]u8 {
    if (std.mem.indexOfScalar(u8, raw, '<') != null) return error.MalformedSheetXml;
    const decoded = formula_mod.decode.decodeAt(a, .cell_formula_body, raw) catch |e| switch (e) {
        error.OutOfMemory => return error.OutOfMemory,
        else => return error.MalformedSheetXml,
    };
    // The block-open validation calls this on the caller's gpa, so
    // the refusal path must free what the decode allocated (Codex
    // #215 r6 REL-601; a no-op for the arena-backed record path).
    errdefer a.free(decoded);
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

    /// Replace the workbook's whole `<sheets>…</sheets>` block with an
    /// empty `<sheets/>`, every sheet part and relationship left in
    /// the archive — the sheetless shape the strict inventory refuses
    /// (REL-602 / #219 r1 REL-101), built without knowing the writer's
    /// exact `<sheet>` spelling.
    pub fn emptySheets(allocator: Allocator, io: std.Io, path: []const u8) !void {
        var store = try PartStore.open(allocator, io, path);
        defer store.deinit();
        const p = (try store.part("xl/workbook.xml")) orelse return error.PartNotFound;
        const open = std.mem.indexOf(u8, p.bytes, "<sheets>") orelse return error.PatchAnchorNotFound;
        const close_tag = "</sheets>";
        const close = std.mem.indexOfPos(u8, p.bytes, open, close_tag) orelse return error.PatchAnchorNotFound;
        const patched = try std.mem.concat(allocator, u8, &.{ p.bytes[0..open], "<sheets/>", p.bytes[close + close_tag.len ..] });
        defer allocator.free(patched);
        try store.replacePart("xl/workbook.xml", patched);
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

test "scanSheetRules: non-UTF-8 namespace bindings cannot hide rules (SEC-501)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    // `xmlns="\xff"` is not a binding the walk can rule out — a
    // raw compare read it as bound-to-foreign and the block's rule
    // vanished under an exit 0.
    try testing.expectError(error.MalformedSheetXml, scanRules(a, ws_open ++
        "<conditionalFormatting xmlns=\"\xff\" sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting>" ++
        "</worksheet>"));
    try testing.expectError(error.MalformedSheetXml, scanRules(a, "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " ++
        "xmlns:zz=\"\xff\"><sheetData/></worksheet>"));
    // The decoded value is held to the same floor.
    try testing.expectError(error.MalformedSheetXml, scanRules(a, "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " ++
        "xmlns:zz=\"a&amp;\xff\"><sheetData/></worksheet>"));
}

test "scanSheetRules: XML-forbidden characters in a binding cannot ghost rules (SEC-701)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    // `&#0;` decodes, `&#7;` decodes, a literal C0 byte and U+FFFE
    // are valid UTF-8 — but none is an XML 1.0 Char, so none is a
    // binding the walk can rule out.
    const bindings = [_][]const u8{ "&#0;", "&#7;", "\x01", "\xEF\xBF\xBE" };
    inline for (bindings) |b| {
        try testing.expectError(error.MalformedSheetXml, scanRules(a, ws_open ++
            "<conditionalFormatting xmlns=\"" ++ b ++ "\" sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting>" ++
            "</worksheet>"));
    }
}

test "collect: lenient Type/Target decoding and duplicate declarations verify strictly (SEC-702/703)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();

    const rels = "xl/_rels/workbook.xml.rels";
    const rel_ns = "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";
    const cases = [_]struct { name: []const u8, part: []const u8, old: []const u8, new: []const u8 }{
        // The store's wider decoder reads `&#X65;`; the strict one
        // refuses — the resolution must not rest on the lenient
        // reading of a Type or Target carrier (SEC-702).
        .{ .name = "x1.xlsx", .part = rels, .old = "relationships/worksheet\" Target=\"worksheets/sheet2.xml\"", .new = "relationships/worksh&#X65;et\" Target=\"worksheets/sheet2.xml\"" },
        .{ .name = "x2.xlsx", .part = rels, .old = "Target=\"worksheets/sheet2.xml\"", .new = "Target=\"worksheets/sheet&#X32;.xml\"" },
        // A duplicate xmlns:r — either value order — is an ambiguous
        // binding, not an authorization (SEC-703).
        .{ .name = "x3.xlsx", .part = "xl/workbook.xml", .old = rel_ns, .new = rel_ns ++ " xmlns:r=\"urn:foreign\"" },
        .{ .name = "x4.xlsx", .part = "xl/workbook.xml", .old = rel_ns, .new = "xmlns:r=\"urn:foreign\" " ++ rel_ns },
        // An XML-forbidden character reference in a sheet-part binding
        // refuses through a real workbook too (SEC-701).
        .{ .name = "x5.xlsx", .part = "xl/worksheets/sheet1.xml", .old = "<conditionalFormatting sqref=\"A1:A4\">", .new = "<conditionalFormatting xmlns:q=\"&#0;\" sqref=\"A1:A4\">" },
    };
    for (cases) |case| {
        const path = try tt.path(testing.allocator, io, case.name);
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path);
        try fixture.patchPart(testing.allocator, io, path, case.part, case.old, case.new);
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        const expected: anyerror = if (std.mem.eql(u8, case.part, "xl/worksheets/sheet1.xml"))
            error.MalformedSheetXml
        else
            error.MalformedWorkbookXml;
        try std.testing.expectError(expected, collect(testing.allocator, &wb));
    }
}

test "scanSheetRules: attribute tokens follow XML's grammar (REL-502)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    const cases = [_][]const u8{
        // A slash taken as a QName boundary cannot smuggle a
        // pseudo-attribute and classify as rule machinery.
        ws_open ++ "<conditionalFormatting/x=\"y\" sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting></worksheet>",
        ws_open ++ "<conditionalFormatting sqref=\"A1\"><cfRule/x=\"y\" type=\"expression\" priority=\"1\"/></conditionalFormatting></worksheet>",
        // A raw `<` in ANY attribute value is ill-formed XML.
        ws_open ++ "<conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\" foo=\"a<b\"><formula>1</formula></cfRule></conditionalFormatting></worksheet>",
    };
    for (cases) |xml| {
        try testing.expectError(error.MalformedSheetXml, scanRules(a, xml));
    }
}

test "scanSheetRules: a refused sqref does not leak the scratch decode (REL-601)" {
    // std.testing.allocator leak-checks at test end — the block-open
    // validation runs on the caller's gpa, so its refusal path must
    // free what the decode allocated. Entity-free and entity-bearing
    // invalid UTF-8 both exercise the allocating decode.
    var out: std.ArrayListUnmanaged(RawRule) = .empty;
    defer out.deinit(testing.allocator);
    try testing.expectError(error.MalformedSheetXml, scanSheetRules(testing.allocator, ws_open ++
        "<conditionalFormatting sqref=\"A1\xff\"/></worksheet>", &out));
    try testing.expectError(error.MalformedSheetXml, scanSheetRules(testing.allocator, ws_open ++
        "<conditionalFormatting sqref=\"A1&amp;\xff\"/></worksheet>", &out));
}

test "collect: the relationship graph verifies strictly (SEC-601, REL-602)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();

    const cases = [_]struct { name: []const u8, part: []const u8, old: []const u8, new: []const u8 }{
        // `r:id` under a foreign binding is not a relationship
        // reference; without any root-declared relationships prefix
        // the entry has none at all.
        .{
            .name = "n1.xlsx",
            .part = "xl/workbook.xml",
            .old = "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"",
            .new = "xmlns:r=\"urn:vendor\"",
        },
        // A Relationship-shaped element outside the schema slot is a
        // decoy the lexical store may have counted.
        .{
            .name = "n2.xlsx",
            .part = "xl/_rels/workbook.xml.rels",
            .old = "</Relationships>",
            .new = "<Wrap><Relationship Id=\"rNest\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/></Wrap></Relationships>",
        },
        // Duplicate relationship ids make "the" relationship
        // ambiguous.
        .{ .name = "n3.xlsx", .part = "xl/_rels/workbook.xml.rels", .old = "Id=\"rId2\"", .new = "Id=\"rId1\"" },
        // No <sheets> wrapper at all is not an empty success —
        // there is nothing to verify the projection against.
        .{ .name = "n4.xlsx", .part = "xl/workbook.xml", .old = "<sheets>", .new = "<sheetsX>" },
    };
    for (cases, 0..) |case, ci| {
        const path = try tt.path(testing.allocator, io, case.name);
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path);
        try fixture.patchPart(testing.allocator, io, path, case.part, case.old, case.new);
        if (ci == 3) {
            try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "</sheets>", "</sheetsX>");
        }
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        try testing.expectError(error.MalformedWorkbookXml, collect(testing.allocator, &wb));
    }
}

test "collect: sheet identities come from the strict workbook read (SEC-502, REL-801)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();

    // A foreign-slot <sheet> decoy under extLst is NOT an identity —
    // the strict read excludes it, and the stream stays exactly the
    // real sheets' records.
    {
        const path = try tt.path(testing.allocator, io, "w1.xlsx");
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path);
        try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "</workbook>", "<extLst><sheet name=\"Ghost\" sheetId=\"9\" r:id=\"rId1\"/></extLst></workbook>");
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        var view = try collect(testing.allocator, &wb);
        defer view.deinit();
        try testing.expectEqual(@as(usize, 2), view.sheet_names.len);
        try testing.expectEqual(@as(usize, 5), view.records.len);
        for (view.sheet_names) |n| try testing.expect(!std.mem.eql(u8, n, "Ghost"));
    }
    // A real entry missing a required carrier refuses instead of
    // silently vanishing.
    {
        const path = try tt.path(testing.allocator, io, "w2.xlsx");
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path);
        try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", " name=\"Report\"", "");
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        try testing.expectError(error.MalformedWorkbookXml, collect(testing.allocator, &wb));
    }
    // A standards-valid ALTERNATE relationships prefix is an identity
    // (REL-801) — the strict read is authoritative, not the lenient
    // projection's literal `r:id` spelling.
    {
        const path = try tt.path(testing.allocator, io, "w3.xlsx");
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path);
        try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"", "xmlns:q=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"");
        try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "r:id=\"rId1\"", "q:id=\"rId1\"");
        try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "r:id=\"rId2\"", "q:id=\"rId2\"");
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        var view = try collect(testing.allocator, &wb);
        defer view.deinit();
        try testing.expectEqual(@as(usize, 5), view.records.len);
        try testing.expectEqualStrings("Report", view.records[4].sheet);
    }
}

test "scanSheetRules: prefixes resolve in scope; reserved declarations refuse (SEC-1101)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    // An undeclared prefixed element is namespace-malformed XML, not
    // harmless opaque content.
    try testing.expectError(error.MalformedSheetXml, scanRules(a, ws_open ++
        "<foo:bar/><conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting>" ++
        "</worksheet>"));
    // The xmlns prefix may not be declared; nothing may bind the
    // reserved XMLNS URI; xml may only rebind to its fixed URI.
    try testing.expectError(error.MalformedSheetXml, scanRules(a, "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " ++
        "xmlns:xmlns=\"urn:x\"><sheetData/></worksheet>"));
    try testing.expectError(error.MalformedSheetXml, scanRules(a, "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " ++
        "xmlns:zz=\"http://www.w3.org/2000/xmlns/\"><sheetData/></worksheet>"));
    try testing.expectError(error.MalformedSheetXml, scanRules(a, "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " ++
        "xmlns:xml=\"urn:not-the-xml-ns\"><sheetData/></worksheet>"));
    // A DECLARED foreign prefix stays ordinary opaque content, and
    // xml:* attributes are implicitly in scope.
    const rules = try scanRules(a, "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:f=\"urn:f\">" ++
        "<f:thing xml:space=\"preserve\"/>" ++
        "<conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting>" ++
        "</worksheet>");
    try testing.expectEqual(@as(usize, 1), rules.len);
}

test "scanSheetRules: reserved URIs as a DEFAULT binding refuse; empty prefixed bindings refuse (SEC-1201)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    // A rule block rebinding its default to a reserved namespace is
    // namespace-invalid XML, not a foreign subtree to step over.
    inline for (.{ "http://www.w3.org/2000/xmlns/", "http://www.w3.org/XML/1998/namespace" }) |uri| {
        try testing.expectError(error.MalformedSheetXml, scanRules(a, ws_open ++
            "<conditionalFormatting xmlns=\"" ++ uri ++ "\" sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting>" ++
            "</worksheet>"));
    }
    try testing.expectError(error.MalformedSheetXml, scanRules(a, "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " ++
        "xmlns:zz=\"\"><sheetData/></worksheet>"));
}

test "collect: a reserved default binding in a real workbook refuses (SEC-1201)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "resv.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);
    try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet1.xml", "<conditionalFormatting sqref=\"A1:A4\">", "<conditionalFormatting xmlns=\"http://www.w3.org/2000/xmlns/\" sqref=\"A1:A4\">");
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    try testing.expectError(error.MalformedSheetXml, collect(testing.allocator, &wb));
}

test "collect: reserved-prefix and undeclared-prefix shapes refuse across parts (SEC-1101)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();

    const cases = [_]struct { name: []const u8, part: []const u8, old: []const u8, new: []const u8, err: anyerror }{
        // `xmlns:xmlns="<rel-ns>"` must not authorize `xmlns:id` as an
        // id reference — the declaration refuses outright.
        .{ .name = "p1.xlsx", .part = "xl/workbook.xml", .old = "<workbook ", .new = "<workbook xmlns:xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ", .err = error.MalformedWorkbookXml },
        // Undeclared prefixed elements refuse in every walked part.
        .{ .name = "p2.xlsx", .part = "xl/worksheets/sheet1.xml", .old = "</worksheet>", .new = "<foo:bar/></worksheet>", .err = error.MalformedSheetXml },
        .{ .name = "p3.xlsx", .part = "xl/_rels/workbook.xml.rels", .old = "</Relationships>", .new = "<p:Relationship/></Relationships>", .err = error.MalformedWorkbookXml },
    };
    for (cases) |case| {
        const path = try tt.path(testing.allocator, io, case.name);
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path);
        try fixture.patchPart(testing.allocator, io, path, case.part, case.old, case.new);
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        try std.testing.expectError(case.err, collect(testing.allocator, &wb));
    }
}

test "scanSheetRules: the canonical xm:macrosheet root reads; foreign bindings refuse (REL-2001)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    // Microsoft's canonical macro-sheet root: prefix bound to the
    // macro namespace, default bound to main for the children.
    const rules = try scanRules(a, "<xm:macrosheet xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" ++
        "<conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting>" ++
        "</xm:macrosheet>");
    try testing.expectEqual(@as(usize, 1), rules.len);

    // An arbitrary bound prefix is not the canonical shape.
    try testing.expectError(error.MalformedSheetXml, scanRules(a, "<xm:macrosheet xmlns:xm=\"urn:foreign\" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData/></xm:macrosheet>"));
}

test "scanSheetRules: a long macro-sheet prefix is not a refusal (REL-2101)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    const lp = "pppppppppppppppppppppppppppppppppppppppppppppppppppppppppppp";
    const rules = try scanRules(a, "<" ++ lp ++ ":macrosheet xmlns:" ++ lp ++ "=\"http://schemas.microsoft.com/office/excel/2006/main\" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" ++
        "<conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting>" ++
        "</" ++ lp ++ ":macrosheet>");
    try testing.expectEqual(@as(usize, 1), rules.len);
}

test "scanSheetRules: MCE at a recognized slot refuses; inert MCE reads (SEC-2201)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    const cf = "<conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting>";
    // A Choice branch, a Fallback branch, an alternate prefix, a
    // ProcessContent wrapper, a block-slot AlternateContent and a
    // default MCE binding would each re-project rule machinery a walk
    // with no MCE processor cannot see.
    const refusals = [_][]const u8{
        ws_open ++ "<mc:AlternateContent xmlns:mc=\"" ++ mce_ns ++ "\"><mc:Choice Requires=\"x14\">" ++ cf ++ "</mc:Choice></mc:AlternateContent></worksheet>",
        ws_open ++ "<mc:AlternateContent xmlns:mc=\"" ++ mce_ns ++ "\"><mc:Fallback>" ++ cf ++ "</mc:Fallback></mc:AlternateContent></worksheet>",
        ws_open ++ "<q:AlternateContent xmlns:q=\"" ++ mce_ns ++ "\"><q:Choice Requires=\"x14\">" ++ cf ++ "</q:Choice></q:AlternateContent></worksheet>",
        ws_open ++ "<w:keep xmlns:w=\"urn:x\" xmlns:mc=\"" ++ mce_ns ++ "\" mc:ProcessContent=\"w:keep\">" ++ cf ++ "</w:keep></worksheet>",
        ws_open ++ "<conditionalFormatting sqref=\"A1\"><mc:AlternateContent xmlns:mc=\"" ++ mce_ns ++ "\"><mc:Choice Requires=\"x14\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></mc:Choice></mc:AlternateContent></conditionalFormatting></worksheet>",
        ws_open ++ "<wrap xmlns=\"" ++ mce_ns ++ "\">" ++ cf ++ "</wrap></worksheet>",
    };
    for (refusals) |xml| {
        try testing.expectError(error.MalformedSheetXml, scanRules(a, xml));
    }
    // The declaration and `mc:Ignorable` are how every modern Excel
    // part spells MCE, and a DEEP AlternateContent (the real
    // `<oleObject>` alternate shape) sits inside an opaque subtree —
    // all inert, the rule still reads.
    const inert = try scanRules(a, "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"" ++ mce_ns ++ "\" mc:Ignorable=\"x14ac\">" ++
        "<oleObjects><mc:AlternateContent><mc:Choice Requires=\"x14\"><oleObject shapeId=\"1025\"/></mc:Choice><mc:Fallback><oleObject shapeId=\"1025\"/></mc:Fallback></mc:AlternateContent></oleObjects>" ++
        cf ++ "</worksheet>");
    try testing.expectEqual(@as(usize, 1), inert.len);
}

test "collect: an MCE branch around a real sheet's rules refuses (SEC-2201)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "mce_sheet.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);
    try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet2.xml", "<conditionalFormatting", "<mc:AlternateContent xmlns:mc=\"" ++ mce_ns ++ "\"><mc:Choice Requires=\"x14\"><conditionalFormatting");
    try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet2.xml", "</conditionalFormatting>", "</conditionalFormatting></mc:Choice></mc:AlternateContent>");
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    try testing.expectError(error.MalformedSheetXml, collect(testing.allocator, &wb));
}

test "collect: an MCE branch inside <sheets> refuses (SEC-2201)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "mce_wb.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);
    try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "<sheet name=\"Report\"", "<mc:AlternateContent xmlns:mc=\"" ++ mce_ns ++ "\"><mc:Choice Requires=\"x15\"><sheet name=\"Report\"");
    try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "</sheets>", "</mc:Choice></mc:AlternateContent></sheets>");
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    try testing.expectError(error.MalformedWorkbookXml, collect(testing.allocator, &wb));
}

test "scanSheetRules: MCE binding state — entity spelling and shadow restore (MNT-2303)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    // An entity-spelled MCE binding IS the MCE binding — the decoded
    // compare every other namespace already gets (SEC-301's rule).
    try testing.expectError(error.MalformedSheetXml, scanRules(a, ws_open ++
        "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility&#47;2006\"><mc:Choice/></mc:AlternateContent></worksheet>"));
    // A foreign shadow of `mc` ends with its element: the outer MCE
    // binding restores on pop and the slot rule still fires.
    try testing.expectError(error.MalformedSheetXml, scanRules(a, "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"" ++ mce_ns ++ "\">" ++
        "<wrap xmlns:mc=\"urn:x\"><mc:thing/></wrap>" ++
        "<mc:AlternateContent><mc:Choice/></mc:AlternateContent></worksheet>"));
}

/// The rels-part fuzz seed and its matching store entry: with an
/// EMPTY expected list the verifier exits before the carrier
/// decoders, so the fuzz walk also runs against one matching internal
/// relationship (Codex #215 r24 MNT-2401).
const rels_seed_xml = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"t\" Target=\"x\"/></Relationships>";

fn verifySeedRels(a: Allocator, xml: []const u8) Error!void {
    const seed = [_]store_mod.Relationship{
        .{ .id = "rId1", .type = "t", .target = "x", .target_mode = .internal },
    };
    return verifyWorkbookRels(a, xml, &seed);
}

fn fuzzScannersTarget(_: void, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();
    // 4 KiB matches the sheet_edit harness's scratch bound.
    var smith_buf: [4096]u8 = undefined;
    const input = smith_buf[0..smith.slice(&smith_buf)];

    var arena = std.heap.ArenaAllocator.init(std.testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    var rules: std.ArrayListUnmanaged(RawRule) = .empty;
    _ = scanSheetRules(a, input, &rules) catch {};
    var sheets: std.ArrayListUnmanaged(StrictSheet) = .empty;
    var calc_slot: CalcPrSlot = .{};
    scanWorkbookSheets(a, input, &sheets, &calc_slot) catch {};
    verifyWorkbookRels(a, input, &.{}) catch {};
    verifySeedRels(a, input) catch {};
}

/// The shapes the review rounds actually broke, plus the truncations
/// a hand-written test would not think to write.
const scanner_fuzz_corpus = [_][]const u8{
    "",
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>",
    ws_open ++ "<conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting></worksheet>",
    "<xm:macrosheet xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>",
    ws_open ++ "<mc:AlternateContent xmlns:mc=\"" ++ mce_ns ++ "\"><mc:Choice/></mc:AlternateContent></worksheet>",
    ws_open ++ "<conditionalFormatting sqref=\"A1\"><cfRule type=\"x\" priority=\"1\"><formula>",
    ws_open ++ "<!--",
    ws_open ++ "<?tgt",
    "<?xml version=\"1.0\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"a\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>",
    rels_seed_xml,
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=",
    ws_open ++ "<a xmlns:b=\"c\" b:d=\"e\"",
    "<<<<>>>>",
};

test "fuzz: the strict scanners never crash on adversarial XML (MNT-2303)" {
    try std.testing.fuzz({}, fuzzScannersTarget, .{ .corpus = &scanner_fuzz_corpus });
}

test "verifyWorkbookRels: the canonical fuzz seed verifies whole (MNT-2401)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    try verifySeedRels(arena.allocator(), rels_seed_xml);
}

test "collect: an entity-spelled TargetMode verifies by its decoded value (REL-2301)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "tmode.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);
    // `Inter&#110;al` DECODES to Internal — the store cached it that
    // way, and the strict verifier must read the same spelling.
    try fixture.patchPart(testing.allocator, io, path, "xl/_rels/workbook.xml.rels", "Target=\"worksheets/sheet2.xml\"", "Target=\"worksheets/sheet2.xml\" TargetMode=\"Inter&#110;al\"");
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    var view = try collect(testing.allocator, &wb);
    defer view.deinit();
    try testing.expectEqual(@as(usize, 5), view.records.len);
}

test "collect: a decoded External sheet target refuses by POLICY, not by spelling (REL-2301)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "tmode_ext.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);
    try fixture.patchPart(testing.allocator, io, path, "xl/_rels/workbook.xml.rels", "Target=\"worksheets/sheet2.xml\"", "Target=\"worksheets/sheet2.xml\" TargetMode=\"Exter&#110;al\"");
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    // The spelling decodes and agrees with the cache; the refusal is
    // the sheet-graph rule that a sheet relationship must be internal.
    try testing.expectError(error.MalformedWorkbookXml, collect(testing.allocator, &wb));
}

test "writeAll: the C handoff writes the full stream byte-exactly (MNT-2302)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "writeall.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    var view = try collect(testing.allocator, &wb);
    defer view.deinit();
    var out: std.Io.Writer.Allocating = .init(testing.allocator);
    defer out.deinit();
    try writeAll(&out.writer, &view);
    try testing.expectEqualStrings(
        "{\"kind\":\"conditional_format\",\"sheet\":\"Data\",\"sheet_idx\":0,\"sqref\":\"A1:A4\"," ++
            "\"rule_type\":\"cellIs\",\"formulas\":[\"2\",\"4\"],\"dxf_id\":0,\"priority\":1}\n" ++
            "{\"kind\":\"conditional_format\",\"sheet\":\"Data\",\"sheet_idx\":0,\"sqref\":\"B1:B4\"," ++
            "\"rule_type\":\"expression\",\"formulas\":[\"B1>3\"],\"dxf_id\":0,\"priority\":2}\n" ++
            "{\"kind\":\"conditional_format\",\"sheet\":\"Data\",\"sheet_idx\":0,\"sqref\":\"C1:C4\"," ++
            "\"rule_type\":\"colorScale\",\"formulas\":[],\"dxf_id\":null,\"priority\":3}\n" ++
            "{\"kind\":\"conditional_format\",\"sheet\":\"Data\",\"sheet_idx\":0,\"sqref\":\"D1:D4\"," ++
            "\"rule_type\":\"dataBar\",\"formulas\":[],\"dxf_id\":null,\"priority\":4}\n" ++
            "{\"kind\":\"conditional_format\",\"sheet\":\"Report\",\"sheet_idx\":1,\"sqref\":\"A1:A2\"," ++
            "\"rule_type\":\"expression\",\"formulas\":[\"$A1=\\\"R&D\\\"\"],\"dxf_id\":0,\"priority\":1}\n",
        out.written(),
    );
}

test "collect: a canonical macro sheet in a real package keeps its rules (REL-2001)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "macro.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);
    // Convert Report into an xlMacrosheet: the relationship type and
    // the part's root spelling, canonical xm-prefixed.
    try fixture.patchPart(testing.allocator, io, path, "xl/_rels/workbook.xml.rels", "officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"", "officeDocument/2006/relationships/xlMacrosheet\" Target=\"worksheets/sheet2.xml\"");
    try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet2.xml", "<worksheet xmlns=", "<xm:macrosheet xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\" xmlns=");
    try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet2.xml", "</worksheet>", "</xm:macrosheet>");
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    var view = try collect(testing.allocator, &wb);
    defer view.deinit();
    try testing.expectEqual(@as(usize, 5), view.records.len);
    try testing.expectEqualStrings("Report", view.records[4].sheet);
    try testing.expectEqualStrings("A1:A2", view.records[4].sqref);
}

test "collect: decoded-duplicate sheet names refuse (REL-1901)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "dupname.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);
    // `D_x0061_ta` DECODES to `Data` — `--name Data` would silently
    // omit the later sheet's rules.
    try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "name=\"Report\"", "name=\"D_x0061_ta\"");
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    try testing.expectError(error.MalformedWorkbookXml, collect(testing.allocator, &wb));
}

test "scanSheetRules: a colon-bearing PI target is legal (REL-1902)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    const rules = try scanRules(a, ws_open ++
        "<?vendor:tool mode=\"x\" <cfRule/> ?>" ++
        "<conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting>" ++
        "</worksheet>");
    try testing.expectEqual(@as(usize, 1), rules.len);
}

test "scanSheetRules: the XML declaration itself parses strictly (REL-1801)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    // Version is mandatory; markup inside the declaration is not a
    // declaration.
    try testing.expectError(error.MalformedSheetXml, scanRules(a, "<?xml?>" ++ ws_open ++ "<sheetData/></worksheet>"));
    try testing.expectError(error.MalformedSheetXml, scanRules(a, "<?xml <conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting> ?>" ++ ws_open ++ "<sheetData/></worksheet>"));
    try testing.expectError(error.MalformedSheetXml, scanRules(a, "<?xml version=\"2.7\"?>" ++ ws_open ++ "<sheetData/></worksheet>"));
    // The full standard prolog reads.
    const rules = try scanRules(a, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++ ws_open ++
        "<conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting>" ++
        "</worksheet>");
    try testing.expectEqual(@as(usize, 1), rules.len);
}

test "collect: numeric attrs resolve entities before lexical typing (REL-1802)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();

    // `priority="&#49;"` IS 1 — references resolve before the
    // digit-only rule (the styles_xml numeric precedent)…
    {
        const path = try tt.path(testing.allocator, io, "num1.xlsx");
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path);
        try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet1.xml", "priority=\"1\" operator=\"between\"", "priority=\"&#49;\" operator=\"between\"");
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        var view = try collect(testing.allocator, &wb);
        defer view.deinit();
        try testing.expectEqual(@as(?u32, 1), view.records[0].priority);
    }
    // …a reference that does not decode refuses the inventory…
    {
        const path = try tt.path(testing.allocator, io, "num2.xlsx");
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path);
        try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet1.xml", "priority=\"1\" operator=\"between\"", "priority=\"&bogus;\" operator=\"between\"");
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        try testing.expectError(error.MalformedSheetXml, collect(testing.allocator, &wb));
    }
    // …and a decoded non-digit spelling still reads as absent.
    {
        const path = try tt.path(testing.allocator, io, "num3.xlsx");
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path);
        try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet1.xml", "priority=\"1\" operator=\"between\"", "priority=\"&#43;1\" operator=\"between\"");
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        var view = try collect(testing.allocator, &wb);
        defer view.deinit();
        try testing.expectEqual(@as(?u32, null), view.records[0].priority);
    }
}

test "scanSheetRules: PIs validate strictly; the prolog declaration stays legal (REL-1702)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    const cf_doc = "<conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting>";
    // An empty PITarget is not a PI — a delimiter-only skip let it
    // swallow rule-shaped bytes.
    try testing.expectError(error.MalformedSheetXml, scanRules(a, ws_open ++
        "<? " ++ cf_doc ++ " ?></worksheet>"));
    // The reserved `xml` target may only be the document's prolog
    // declaration, never an in-root PI.
    try testing.expectError(error.MalformedSheetXml, scanRules(a, ws_open ++
        "<?xml version=\"1.0\"?><sheetData/></worksheet>"));
    // A valid vendor PI is non-element content wherever it appears —
    // rule-shaped data inside it stays uncounted, and the prolog
    // declaration still reads.
    const rules = try scanRules(a, "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" ++ ws_open ++
        "<?vendor " ++ cf_doc ++ " ?>" ++ cf_doc ++ "</worksheet>");
    try testing.expectEqual(@as(usize, 1), rules.len);
}

test "collect: carrier classes stay split — ST_Xstring on names, entities-only elsewhere (MNT-1703)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();

    // `_x006F_` decodes in a sheet NAME (string carrier)…
    {
        const path = try tt.path(testing.allocator, io, "car1.xlsx");
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path);
        try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "name=\"Report\"", "name=\"Rep_x006F_rt\"");
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        var view = try collect(testing.allocator, &wb);
        defer view.deinit();
        try testing.expectEqualStrings("Report", view.sheet_names[1]);
        try testing.expectEqualStrings("Report", view.records[4].sheet);
    }
    // …stays LITERAL in sqref and formulas (formula carrier, no
    // ST_Xstring layer)…
    {
        const path = try tt.path(testing.allocator, io, "car2.xlsx");
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path);
        try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet1.xml", "sqref=\"A1:A4\"", "sqref=\"_x0041_1:A4\"");
        try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet1.xml", "<formula>2</formula>", "<formula>_x0041_1</formula>");
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        var view = try collect(testing.allocator, &wb);
        defer view.deinit();
        try testing.expectEqualStrings("_x0041_1:A4", view.records[0].sqref);
        try testing.expectEqualStrings("_x0041_1", view.records[0].formulas[0]);
    }
    // …and an ill-formed escape in a sheet name refuses.
    {
        const path = try tt.path(testing.allocator, io, "car3.xlsx");
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path);
        try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "name=\"Report\"", "name=\"Rep_xD800_rt\"");
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        try testing.expectError(error.MalformedWorkbookXml, collect(testing.allocator, &wb));
    }
}

test "scanSheetRules: malformed comments cannot swallow a rule block (REL-1501)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    const cf_doc = "<conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting>";
    // `--` inside comment content is malformed XML — a delimiter-only
    // skip let this comment eat the rule block under exit 0.
    try testing.expectError(error.MalformedSheetXml, scanRules(a, ws_open ++
        "<!-- bad -- " ++ cf_doc ++ " -->" ++
        "</worksheet>"));
    // Content ending in `-` is malformed too.
    try testing.expectError(error.MalformedSheetXml, scanRules(a, ws_open ++
        "<!-- x --->" ++ cf_doc ++ "</worksheet>"));
}

test "collect: a malformed comment in a real workbook refuses (REL-1501); notBetween pin (MNT-1501)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();

    {
        const path = try tt.path(testing.allocator, io, "cmt.xlsx");
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path);
        try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet1.xml", "<conditionalFormatting sqref=\"A1:A4\">", "<!-- a -- b --><conditionalFormatting sqref=\"A1:A4\">");
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        try testing.expectError(error.MalformedSheetXml, collect(testing.allocator, &wb));
    }
    // A notBetween rule wears the same `cellIs` type and two formula
    // bodies — the wire does not carry the operator, which is why the
    // docs' two-formula pipeline is operator-agnostic (MNT-1501).
    {
        const path = try tt.path(testing.allocator, io, "notbtw.xlsx");
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path);
        try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet1.xml", "operator=\"between\"", "operator=\"notBetween\"");
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        var view = try collect(testing.allocator, &wb);
        defer view.deinit();
        try testing.expectEqualStrings("cellIs", view.records[0].rule_type);
        try testing.expectEqual(@as(usize, 2), view.records[0].formulas.len);
    }
}

test "scanSheetRules: character data outside the root refuses; a legal prolog reads (REL-1102)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    const cf_doc = "<conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting>";
    try testing.expectError(error.MalformedSheetXml, scanRules(a, "junk" ++ ws_open ++ cf_doc ++ "</worksheet>"));
    try testing.expectError(error.MalformedSheetXml, scanRules(a, ws_open ++ cf_doc ++ "</worksheet>trailing"));
    try testing.expectError(error.MalformedSheetXml, scanRules(a, "<![CDATA[x]]>" ++ ws_open ++ cf_doc ++ "</worksheet>"));
    // BOM, XML declaration, comments and whitespace around the root
    // are the legal prolog/epilog.
    const rules = try scanRules(a, "\xEF\xBB\xBF<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<!-- prolog -->\n" ++
        ws_open ++ cf_doc ++ "</worksheet>\n<!-- epilog -->\n");
    try testing.expectEqual(@as(usize, 1), rules.len);
}

test "scanSheetRules: names validate by XML's code points, parts by UTF-8 whole (SEC-1001)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    // A matched `\xff`-suffixed pair used to classify as opaque and
    // hide the rule inside it; the whole-part UTF-8 gate refuses it.
    try testing.expectError(error.MalformedSheetXml, scanRules(a, ws_open ++
        "<conditionalFormatting\xff sqref=\"Z1\"><cfRule type=\"expression\" priority=\"9\"><formula>1</formula></cfRule></conditionalFormatting\xff>" ++
        "</worksheet>"));
    // Valid UTF-8 but not an XML name character (U+00A0) refuses at
    // the QName check rather than reading as an opaque element.
    try testing.expectError(error.MalformedSheetXml, scanRules(a, ws_open ++
        "<conditionalFormatting\xc2\xa0 sqref=\"Z1\"><cfRule type=\"expression\" priority=\"9\"><formula>1</formula></cfRule></conditionalFormatting\xc2\xa0>" ++
        "</worksheet>"));
    // A valid non-ASCII NCName is an ordinary opaque element.
    const rules = try scanRules(a, ws_open ++
        "<donn\xc3\xa9es><x/></donn\xc3\xa9es>" ++
        "<conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting>" ++
        "</worksheet>");
    try testing.expectEqual(@as(usize, 1), rules.len);
}

test "scanSheetRules: attributes require their XML separator (SEC-901)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    // A namespace rebinding riding in UNSEPARATED after a quoted
    // value tokenized, turned the block foreign, and its rule
    // vanished under exit 0.
    try testing.expectError(error.MalformedSheetXml, scanRules(a, ws_open ++
        "<conditionalFormatting sqref=\"A1\"xmlns=\"urn:foreign\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting>" ++
        "</worksheet>"));
}

test "collect: separatorless attributes cannot hide a block or a sheet (SEC-901)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();

    // A CF block cannot disappear behind an unseparated rebinding.
    {
        const path = try tt.path(testing.allocator, io, "s1.xlsx");
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path);
        try fixture.patchPart(testing.allocator, io, path, "xl/worksheets/sheet1.xml", "sqref=\"A1:A4\">", "sqref=\"A1:A4\"xmlns=\"urn:foreign\">");
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        try testing.expectError(error.MalformedSheetXml, collect(testing.allocator, &wb));
    }
    // A workbook <sheet> identity cannot disappear the same way.
    {
        const path = try tt.path(testing.allocator, io, "s2.xlsx");
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path);
        try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "name=\"Report\" ", "name=\"Report\"xmlns=\"urn:foreign\" ");
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        try testing.expectError(error.MalformedWorkbookXml, collect(testing.allocator, &wb));
    }
}

test "collect: malformed namespace prefixes cannot slip the identity gate (SEC-801)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();

    // An EMPTY prefix (`xmlns:=` + `:id`) is malformed namespace XML,
    // not a relationships binding.
    {
        const path = try tt.path(testing.allocator, io, "q1.xlsx");
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path);
        try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"", "xmlns:=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"");
        try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "r:id=\"rId1\"", ":id=\"rId1\"");
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        try testing.expectError(error.MalformedWorkbookXml, collect(testing.allocator, &wb));
    }
    // A multi-colon declaration name is not a QName.
    {
        const path = try tt.path(testing.allocator, io, "q2.xlsx");
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path);
        try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "<workbook ", "<workbook xmlns:a:b=\"urn:x\" ");
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        try testing.expectError(error.MalformedWorkbookXml, collect(testing.allocator, &wb));
    }
}

test "scanSheetRules: many distinct legal declarations read — no fixed table (REL-1301)" {
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(testing.allocator);
    try buf.appendSlice(testing.allocator, "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
    for (0..17) |k| {
        try buf.appendSlice(testing.allocator, " xmlns:f");
        try buf.append(testing.allocator, @intCast('a' + k));
        try buf.appendSlice(testing.allocator, "=\"urn:f\"");
    }
    try buf.appendSlice(testing.allocator, ">" ++
        "<conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting>" ++
        "</worksheet>");
    const rules = try scanRules(a, buf.items);
    try testing.expectEqual(@as(usize, 1), rules.len);
}

test "collect: a ninth relationships alias still authorizes its id (REL-1301)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "aliases.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);
    try fixture.patchPart(
        testing.allocator,
        io,
        path,
        "xl/workbook.xml",
        "<workbook ",
        "<workbook xmlns:q1=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:q2=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:q3=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:q4=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:q5=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:q6=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:q7=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:q8=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ",
    );
    try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "r:id=\"rId2\"", "q8:id=\"rId2\"");
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    var view = try collect(testing.allocator, &wb);
    defer view.deinit();
    try testing.expectEqual(@as(usize, 5), view.records.len);
}

test "collect: relationship ids resolve by their decoded spelling (REL-503)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();

    // An entity-spelled r:id is the same XML id and resolves.
    {
        const path = try tt.path(testing.allocator, io, "rid_ok.xlsx");
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path);
        try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "r:id=\"rId2\"", "r:id=\"rId&#50;\"");
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        var view = try collect(testing.allocator, &wb);
        defer view.deinit();
        try testing.expectEqual(@as(usize, 5), view.records.len);
    }
    // A rid whose reference does not decode refuses.
    {
        const path = try tt.path(testing.allocator, io, "rid_bad.xlsx");
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path);
        try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "r:id=\"rId2\"", "r:id=\"rId&bogus;\"");
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        try testing.expectError(error.MalformedWorkbookXml, collect(testing.allocator, &wb));
    }
    // A double-escaped relationship id is a DIFFERENT XML id: the
    // raw-first compare matched it; the decoded compare must not.
    {
        const path = try tt.path(testing.allocator, io, "rid_dbl.xlsx");
        defer testing.allocator.free(path);
        try fixture.write(testing.allocator, io, path);
        try fixture.patchPart(testing.allocator, io, path, "xl/_rels/workbook.xml.rels", "Id=\"rId2\"", "Id=\"r&amp;amp;d\"");
        try fixture.patchPart(testing.allocator, io, path, "xl/workbook.xml", "r:id=\"rId2\"", "r:id=\"r&amp;d\"");
        var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
        defer wb.deinit();
        try testing.expectError(error.MalformedWorkbookXml, collect(testing.allocator, &wb));
    }
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

test "collect: repeated reads leave the store's resident bytes flat (S3B-MEM-603)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "cf_retention.xlsx");
    defer testing.allocator.free(path);
    try fixture.write(testing.allocator, io, path);

    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    // Prime every store-side cache (part bytes, rels) so the loop
    // below measures only what a repeated read RETAINS.
    {
        var primed = try collect(testing.allocator, &wb);
        primed.deinit();
    }
    const before = wb.store.residentBytes();
    // 1024 reads resolve ~2 sheet targets each; through the store's
    // lifetime arena that is ~50 KiB of growth — far past any chunk
    // slack — so this pins that resolution lands in the VIEW's arena
    // (`resolveOwned`), reclaimed at deinit, and a revert to
    // `store.resolve` fails here even though the OOM sweep cannot see
    // the store's own allocator.
    for (0..1024) |_| {
        var view = try collect(testing.allocator, &wb);
        view.deinit();
    }
    try testing.expectEqual(before, wb.store.residentBytes());
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

test "resolveSheets / scanCalcPr: an empty <sheets/> is the sheetless workbook REL-602 refuses (#219 r1 REL-101)" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(testing.allocator, io, "cf_sheetless.xlsx");
    defer testing.allocator.free(path);
    {
        const zlsx = @import("zlsx");
        var w = zlsx.Writer.init(testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Only");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, path);
    }
    // The wrapper is present and well-formed; only its entries are
    // gone. The lenient opener accepts the shape (a zero-length typed
    // inventory), so the strict read is the only guard — every read
    // built over it, and the C exports promising a non-empty stream,
    // must refuse rather than serve an empty success.
    try fixture.emptySheets(testing.allocator, io, path);
    var wb = try workbook_mod.Workbook.open(testing.allocator, io, path);
    defer wb.deinit();
    try testing.expectError(error.MalformedWorkbookXml, collect(testing.allocator, &wb));
    var arena = std.heap.ArenaAllocator.init(testing.allocator);
    defer arena.deinit();
    try testing.expectError(error.MalformedWorkbookXml, resolveSheets(testing.allocator, arena.allocator(), &wb));
    // The calc read runs the same walk: the same verdict.
    try testing.expectError(error.MalformedWorkbookXml, scanCalcPr(testing.allocator, &wb));
}
