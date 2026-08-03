//! `xl/metadata.xml` — the typed reader, `cm`/`vm` resolution, and the
//! dialect primitives that depend on them (`goal_formula.md` §5.3b,
//! §5.6a, §5.8b).
//!
//! M4a of the tier-D1 ladder.
//!
//! Why a whole file for one boolean
//! --------------------------------
//! The only thing recalc *needs* from this part is whether a formula cell
//! was authored as a dynamic array, because §5.3b indexes the entire
//! shape/coercion table on that answer. The rest of the part is the
//! reason the file is not three functions: a cell can also point at a
//! **rich value** — a linked entity, a data type, an image-in-cell —
//! whose displayed content lives outside the cell's `<v>`. Recalculating
//! such a cell and writing a plain cached value silently destroys it, and
//! the destruction is invisible: the file still opens, the number still
//! looks right, the entity is simply gone.
//!
//! So the contract is **classify, never ignore**:
//!
//! * every element the schema defines has a row in `element_inventory`,
//!   and `childOf` — the parser's containment law — is *derived from that
//!   table*, so an element with no row cannot be parsed at all;
//! * every metadata *type* has a row in `type_classification`, and only
//!   `XLDAPR` is interpreted;
//! * anything else reached from an input or a result cell is a
//!   **pre-mutation typed refusal**, not a shrug.
//!
//! An unreferenced record refuses nothing. A workbook full of rich values
//! recalculates fine as long as the run does not read or write a cell
//! that points at one — the refusal is a property of the *reference*,
//! which is why resolution takes a `CellRole`.
//!
//! Provenance: this contract is spec-pinned
//! ----------------------------------------
//! No committed oracle manifest (`tests/oracle/fixtures/`) contains a
//! metadata row, and no workbook in `tests/corpus/` carries an
//! `xl/metadata.xml` part, so nothing here is oracle-decided. The pins:
//!
//! * **element and attribute inventory** — ECMA-376 `sml-sheetMetadata.xsd`
//!   (`CT_Metadata`, `CT_MetadataTypes`, `CT_MetadataType`,
//!   `CT_MetadataBlocks`, `CT_MetadataBlock`, `CT_MetadataRecord`,
//!   `CT_MetadataStrings`, `CT_MdxMetadata`, `CT_Mdx`, `CT_MdxTuple`,
//!   `CT_MdxSet`, `CT_MdxMemeberProp`, `CT_MdxKPI`,
//!   `CT_MetadataStringIndex`, `CT_FutureMetadata`,
//!   `CT_FutureMetadataBlock`);
//! * **`c@cm` / `c@vm`** — `CT_Cell`, `xsd:unsignedInt`, default `0`;
//! * **one-based indexing** — `goal_formula.md:133`; the base schema's
//!   prose and what Office writes differ, and this reader follows Office
//!   (`cm="1"` is the first `<bk>`, `rc@t="1"` the first `<metadataType>`),
//!   which is also what the `0` default forces: a zero-based `cm` could
//!   not distinguish "first block" from "no metadata";
//! * **type names** — `XLDAPR`, `XLRICHVALUE`, `XLMDX` are Office-defined
//!   producer strings, not schema enumerations, which is why `unknown` is
//!   a first-class row rather than an error case.
//!
//! M7b re-pins the *transitions* (which collection, which record, which
//! index, what a missing record means when writing) from byte-diffed
//! Excel references. Nothing here writes, so nothing here front-runs it.
//!
//! Allocation
//! ----------
//! `parse` allocates four flat arrays — types, and a (blocks, records)
//! pair per collection — and every text slice it keeps is **borrowed from
//! the input bytes**, which must outlive the result. Resolution allocates
//! nothing at all, which is what lets it sit on the per-cell path.

const std = @import("std");
const assert = std.debug.assert;

const value = @import("value.zig");
const parser = @import("parser.zig");
const env = @import("env.zig");

/// §10's plane-2 taxonomy has exactly one home. Importing it costs this
/// module the parser's compile time and buys a refusal that cannot drift
/// from the taxonomy it claims to belong to; `eval.zig:129-137` pays the
/// same price for the same reason.
pub const PlaneTwo = parser.PlaneTwo;

// ─── element inventory (ECMA-376 sml-sheetMetadata.xsd) ──────────

/// Every element the metadata part's schema defines. Names that the
/// schema overloads by position — `bk`, `t`, `s`, `n`, `p` — get distinct
/// members, because "which `bk` is this" is exactly the question a reader
/// that resolved names globally would get wrong.
pub const Element = enum {
    /// `metadata` — the part root.
    metadata,
    /// `metadataTypes`.
    metadata_types,
    /// `metadataType`.
    metadata_type,
    /// `metadataStrings`.
    metadata_strings,
    /// `s` — one MDX metadata string.
    metadata_string,
    /// `mdxMetadata`.
    mdx_metadata,
    /// `mdx`.
    mdx,
    /// `t` under `mdx` — `CT_MdxTuple`.
    mdx_tuple,
    /// `ms` under `mdx` — `CT_MdxSet`.
    mdx_set,
    /// `p` under `mdx` — `CT_MdxMemeberProp` (the schema's spelling).
    mdx_member_prop,
    /// `k` under `mdx` — `CT_MdxKPI`.
    mdx_kpi,
    /// `n` under `t`/`ms` — `CT_MetadataStringIndex`.
    mdx_string_index,
    /// `futureMetadata`.
    future_metadata,
    /// `bk` under `futureMetadata` — `CT_FutureMetadataBlock`.
    future_metadata_block,
    /// `cellMetadata`.
    cell_metadata,
    /// `valueMetadata`.
    value_metadata,
    /// `bk` under `cellMetadata`/`valueMetadata` — `CT_MetadataBlock`.
    /// This is the element `cm`/`vm` index into, one-based.
    metadata_block,
    /// `rc` — `CT_MetadataRecord`, the `(type, value)` pair.
    metadata_record,
    /// `extLst`.
    ext_lst,
    /// `ext` — foreign content; its subtree is skipped wholesale.
    ext,
};

pub const element_count = @typeInfo(Element).@"enum".fields.len;

pub const ElementTreatment = enum {
    /// The reader builds typed state from it. Its attributes are
    /// inventoried and an unknown one refuses.
    interpreted,
    /// Recognized and structurally skipped: it carries payload for a
    /// metadata type that is never interpreted, so any cell that could
    /// reach it refuses first. Attributes are not inspected.
    inert,
};

pub const ElementRow = struct {
    element: Element,
    /// The local name, as the schema spells it. Namespace prefixes are
    /// stripped before matching (the namespace *preflight* is M4b1's).
    xml_name: []const u8,
    /// Every position the schema allows this element in. `null` is the
    /// document root.
    parents: []const ?Element,
    treatment: ElementTreatment,
    /// What `interpreted` produces, or why `inert` is safe.
    note: []const u8,
};

/// The classification table. It is not documentation: `childOf` reads it,
/// so an element absent from it has no legal position anywhere and the
/// parse refuses.
pub const element_inventory = [_]ElementRow{
    .{
        .element = .metadata,
        .xml_name = "metadata",
        .parents = &.{null},
        .treatment = .interpreted,
        .note = "part root; carries only namespace declarations",
    },
    .{
        .element = .metadata_types,
        .xml_name = "metadataTypes",
        .parents = &.{.metadata},
        .treatment = .interpreted,
        .note = "the type table rc@t indexes, one-based",
    },
    .{
        .element = .metadata_type,
        .xml_name = "metadataType",
        .parents = &.{.metadata_types},
        .treatment = .interpreted,
        .note = "name -> TypeClass; the whole dialect answer starts here",
    },
    .{
        .element = .metadata_strings,
        .xml_name = "metadataStrings",
        .parents = &.{.metadata},
        .treatment = .inert,
        .note = "MDX string store; reachable only through XLMDX, which refuses",
    },
    .{
        .element = .metadata_string,
        .xml_name = "s",
        .parents = &.{.metadata_strings},
        .treatment = .inert,
        .note = "one MDX string",
    },
    .{
        .element = .mdx_metadata,
        .xml_name = "mdxMetadata",
        .parents = &.{.metadata},
        .treatment = .inert,
        .note = "OLAP payload; reachable only through XLMDX, which refuses",
    },
    .{
        .element = .mdx,
        .xml_name = "mdx",
        .parents = &.{.mdx_metadata},
        .treatment = .inert,
        .note = "one MDX record",
    },
    .{
        .element = .mdx_tuple,
        .xml_name = "t",
        .parents = &.{.mdx},
        .treatment = .inert,
        .note = "CT_MdxTuple",
    },
    .{
        .element = .mdx_set,
        .xml_name = "ms",
        .parents = &.{.mdx},
        .treatment = .inert,
        .note = "CT_MdxSet",
    },
    .{
        .element = .mdx_member_prop,
        .xml_name = "p",
        .parents = &.{.mdx},
        .treatment = .inert,
        .note = "CT_MdxMemeberProp",
    },
    .{
        .element = .mdx_kpi,
        .xml_name = "k",
        .parents = &.{.mdx},
        .treatment = .inert,
        .note = "CT_MdxKPI",
    },
    .{
        .element = .mdx_string_index,
        .xml_name = "n",
        .parents = &.{ .mdx_tuple, .mdx_set },
        .treatment = .inert,
        .note = "CT_MetadataStringIndex, under either MDX container",
    },
    .{
        .element = .future_metadata,
        .xml_name = "futureMetadata",
        .parents = &.{.metadata},
        .treatment = .inert,
        .note = "per-type value store; XLDAPR's payload is M7b/M7c, not M4a",
    },
    .{
        .element = .future_metadata_block,
        .xml_name = "bk",
        .parents = &.{.future_metadata},
        .treatment = .inert,
        .note = "one future-metadata value, rc@v indexes these zero-based",
    },
    .{
        .element = .cell_metadata,
        .xml_name = "cellMetadata",
        .parents = &.{.metadata},
        .treatment = .interpreted,
        .note = "the collection c@cm indexes, one-based",
    },
    .{
        .element = .value_metadata,
        .xml_name = "valueMetadata",
        .parents = &.{.metadata},
        .treatment = .interpreted,
        .note = "the collection c@vm indexes, one-based",
    },
    .{
        .element = .metadata_block,
        .xml_name = "bk",
        .parents = &.{ .cell_metadata, .value_metadata },
        .treatment = .interpreted,
        .note = "one indexable block of records",
    },
    .{
        .element = .metadata_record,
        .xml_name = "rc",
        .parents = &.{.metadata_block},
        .treatment = .interpreted,
        .note = "t = one-based type index, v = zero-based value index",
    },
    .{
        .element = .ext_lst,
        .xml_name = "extLst",
        .parents = &.{ .metadata, .future_metadata, .future_metadata_block },
        .treatment = .inert,
        .note = "extension container",
    },
    .{
        .element = .ext,
        .xml_name = "ext",
        .parents = &.{.ext_lst},
        .treatment = .inert,
        .note = "foreign content; the subtree is skipped, not classified",
    },
};

/// The containment law, read straight off the table. Returns null when
/// `name` has no legal position under `parent` — which the parser turns
/// into a refusal rather than a skip, because "an element we did not
/// recognize" and "an element we chose not to interpret" must not be the
/// same outcome in a part where the difference is a lost rich value.
pub fn childOf(parent: ?Element, name: []const u8) ?Element {
    for (element_inventory) |row| {
        if (!std.mem.eql(u8, row.xml_name, name)) continue;
        for (row.parents) |p| {
            if (p == null and parent == null) return row.element;
            if (p != null and parent != null and p.? == parent.?) return row.element;
        }
    }
    return null;
}

pub fn treatmentOf(el: Element) ElementTreatment {
    for (element_inventory) |row| {
        if (row.element == el) return row.treatment;
    }
    unreachable; // the table is exhaustive; `test "element inventory"` proves it
}

/// `CT_MetadataType`'s complete attribute list. Ships as data because the
/// policy is "unknown attribute → refusal": a list that lived in a
/// hand-written `if` chain would grow a hole the first time someone added
/// a case and forgot the else.
pub const metadata_type_attrs = [_][]const u8{
    "name",                "minSupportedVersion", "ghostRow",       "ghostCol",
    "edit",                "delete",              "copy",           "pasteAll",
    "pasteFormulas",       "pasteValues",         "pasteFormats",   "pasteComments",
    "pasteDataValidation", "pasteBorders",        "pasteColWidths", "pasteNumberFormats",
    "merge",               "splitFirst",          "splitAll",       "rowColShift",
    "clearAll",            "clearFormats",        "clearContents",  "clearComments",
    "assign",              "coerce",              "adjust",         "cellMeta",
};

// ─── metadata type classification ────────────────────────────────

/// What a `<metadataType name="…">` is. Producer-defined strings, so
/// `unknown` is a normal outcome rather than a malformed one.
pub const TypeClass = enum {
    /// `XLDAPR` — dynamic array properties. The one interpreted type:
    /// its presence on a cell is what "authored as a dynamic array"
    /// means, and therefore what `EvalEnv.dialectOf` answers from.
    dynamic_array_properties,
    /// `XLRICHVALUE` — rich values: linked entities, data types,
    /// images-in-cells. Refuses when referenced.
    rich_value,
    /// `XLMDX` — OLAP MDX metadata. Refuses when referenced.
    mdx,
    /// Any other name a producer wrote.
    unknown,
};

pub const TypeTreatment = enum {
    interpreted,
    /// Present in the part costs nothing; *reached from a cell the run
    /// reads or writes* is a refusal.
    refused_when_referenced,
};

pub const TypeRow = struct {
    class: TypeClass,
    /// The Office-defined name, or null for `unknown` — which has no
    /// single spelling by construction.
    name: ?[]const u8,
    treatment: TypeTreatment,
    why: []const u8,
};

pub const type_classification = [_]TypeRow{
    .{
        .class = .dynamic_array_properties,
        .name = "XLDAPR",
        .treatment = .interpreted,
        .why = "marks a cell as dynamic-array authored (Office writes it with cellMeta=\"1\")",
    },
    .{
        .class = .rich_value,
        .name = "XLRICHVALUE",
        .treatment = .refused_when_referenced,
        .why = "the cell's real content lives outside <v>; a recalc that rewrote <v> would delete it",
    },
    .{
        .class = .mdx,
        .name = "XLMDX",
        .treatment = .refused_when_referenced,
        .why = "the cached value belongs to an OLAP query zlsx cannot re-run",
    },
    .{
        .class = .unknown,
        .name = null,
        .treatment = .refused_when_referenced,
        .why = "a type nobody classified cannot be proven safe to overwrite",
    },
};

/// Exact, case-sensitive match. Office writes these names in upper case;
/// a differently-cased or entity-escaped spelling classifies as `unknown`
/// and therefore refuses, which is the safe direction — the alternative
/// is inferring "this is probably XLDAPR" about a part we are about to
/// overwrite values under.
pub fn classifyTypeName(name: []const u8) TypeClass {
    for (type_classification) |row| {
        const known = row.name orelse continue;
        if (std.mem.eql(u8, known, name)) return row.class;
    }
    return .unknown;
}

pub fn treatmentOfClass(class: TypeClass) TypeTreatment {
    for (type_classification) |row| {
        if (row.class == class) return row.treatment;
    }
    unreachable; // exhaustive by `test "type classification"`
}

// ─── refusals (§10) ──────────────────────────────────────────────

/// Which collection a refusal or a resolution walked.
pub const Collection = enum { cell_metadata, value_metadata };

/// Why a cell is being resolved. Both roles refuse identically — the
/// field exists so the diagnostic can say which side of the run tripped,
/// and so a caller cannot forget to check the result side (§5.7.3's
/// staged results are exactly as destructible as the inputs).
pub const CellRole = enum { input, result };

pub const LimitKind = enum { part_bytes, types, blocks, records, depth, skip_depth };

pub const Refusal = struct {
    reason: Reason,
    /// Set when the refusal came from resolving a cell.
    site: ?Site = null,
    /// Byte offset into the part where the refusal was detected. Zero for
    /// resolution refusals, which are about a cell, not about bytes.
    offset: u32 = 0,
    /// Set exactly when `reason == .limit_exceeded`.
    limit: ?LimitKind = null,

    pub const Site = struct {
        role: CellRole,
        collection: Collection,
        /// The one-based `cm`/`vm` the cell carried.
        index: u32,
        /// The offending `rc@t`, when the refusal names a record.
        type_index: ?u32 = null,
    };

    pub const Reason = enum {
        // ── structural: the part itself (parse time) ──
        /// Not valid UTF-8.
        invalid_utf8,
        /// Tag soup: unterminated markup, mismatched close, stray text.
        malformed_xml,
        /// `<!DOCTYPE` / `<!ENTITY`. Refused outright rather than
        /// resolved — an entity-expanding metadata reader is a
        /// denial-of-service surface with no upside.
        doctype_declaration,
        /// An element with no row in `element_inventory` legal here.
        unexpected_element,
        /// An attribute the schema does not define on an interpreted
        /// element.
        unexpected_attribute,
        /// A required attribute is missing.
        missing_attribute,
        /// A value that is not the lexical form its type demands.
        bad_attribute_value,
        /// Two `<cellMetadata>` (or two `<valueMetadata>`) in one part.
        duplicate_collection,
        /// Text where the schema allows only elements.
        unexpected_text,
        /// §9-shaped bound; `limit` names it.
        limit_exceeded,

        // ── resolution: a cell pointing into the part ──
        /// `cm`/`vm` names a block that does not exist.
        index_out_of_range,
        /// `rc@t` names a type that does not exist.
        type_index_out_of_range,
        /// A referenced block holds no records, so it classifies nothing.
        empty_record_block,
        /// `XLRICHVALUE` on an input or a result cell.
        rich_value_metadata,
        /// `XLMDX` on an input or a result cell.
        mdx_metadata,
        /// A producer-defined type nobody classified.
        unknown_metadata_type,
        /// Value metadata of the one interpreted type. `XLDAPR` is a
        /// *cell* metadata type; reached through `vm` it means something
        /// this reader has never been shown, and guessing is how rich
        /// content dies.
        dynamic_array_in_value_metadata,
    };

    /// Exhaustive by construction — a new `Reason` fails to compile until
    /// it has a §10 plane.
    pub fn planeTwo(self: Refusal) PlaneTwo {
        return switch (self.reason) {
            .invalid_utf8,
            .malformed_xml,
            .unexpected_attribute,
            .missing_attribute,
            .bad_attribute_value,
            .duplicate_collection,
            .unexpected_text,
            .index_out_of_range,
            .type_index_out_of_range,
            .empty_record_block,
            => .FormulaMalformedInput,

            .doctype_declaration,
            .unexpected_element,
            .rich_value_metadata,
            .mdx_metadata,
            .unknown_metadata_type,
            .dynamic_array_in_value_metadata,
            => .FormulaUnsupportedConstruct,

            .limit_exceeded => .FormulaLimitExceeded,
        };
    }
};

// ─── the parsed part ─────────────────────────────────────────────

pub const Limits = struct {
    /// Every diagnostic offset is a `u32`, so a part that could not be
    /// addressed by one is refused before it is scanned rather than
    /// truncated into a lie about where the problem is.
    max_bytes: u32 = std.math.maxInt(u32),
    max_types: u32 = 1 << 12,
    max_blocks: u32 = 1 << 20,
    max_records: u32 = 1 << 21,
    /// The schema's deepest classified path is
    /// metadata > futureMetadata > bk > extLst > ext = 5.
    max_depth: u32 = 8,
    /// Foreign content inside `<ext>` nests as deep as it likes; this is
    /// where "as deep as it likes" stops.
    max_skip_depth: u32 = 64,
};

pub const Options = struct { limits: Limits = .{} };

pub const TypeEntry = struct {
    /// Borrowed from the input bytes.
    name: []const u8,
    class: TypeClass,
    min_supported_version: ?u32,
    /// `cellMeta="1"` — Office's own declaration that this type belongs
    /// on `cm` rather than `vm`. Recorded, not enforced: the enforcement
    /// that matters is `classifyTypeName`, and M7b re-pins the collection
    /// rules from byte-diffed references.
    cell_meta: bool,
};

/// One `<rc>`: a type index and a value index, both exactly as written.
pub const Record = struct {
    /// `rc@t`, one-based into `types`.
    type_index: u32,
    /// `rc@v`, zero-based into that type's `futureMetadata` blocks. M4a
    /// records it and interprets nothing.
    value_index: u32,
};

/// A `<bk>`, as a half-open run of `records`.
pub const Block = struct { first: u32, len: u32 };

pub const Blocks = struct {
    blocks: []const Block = &.{},
    records: []const Record = &.{},
    /// The `count` attribute as written. A *hint*: resolution indexes by
    /// position and never consults it. Enforcing agreement would refuse
    /// files Excel opens — `tests/corpus/poi_MalformedSSTCount.xlsx` is
    /// this repo's standing evidence that producers miscount.
    declared_count: ?u32 = null,

    /// One-based, per `goal_formula.md:133`. Index 0 is "no metadata"
    /// and never reaches here.
    pub fn at(self: Blocks, one_based: u32) ?Block {
        assert(one_based != 0);
        if (one_based > self.blocks.len) return null;
        return self.blocks[one_based - 1];
    }

    pub fn recordsOf(self: Blocks, blk: Block) []const Record {
        return self.records[blk.first .. blk.first + blk.len];
    }
};

pub const Metadata = struct {
    types: []const TypeEntry = &.{},
    cell: Blocks = .{},
    value: Blocks = .{},
    /// How many of each element the part held. The inventory's proof that
    /// nothing was skipped silently: `test "every element in the
    /// inventory is reached"` asserts each row's counter moved.
    counts: [element_count]u32 = @splat(0),
    /// `metadataTypes@count` as written; a hint, like `Blocks`'.
    declared_type_count: ?u32 = null,

    /// A workbook with no `xl/metadata.xml` part. Every cell resolves
    /// `.legacy`, which is the correct reading: without the part there
    /// are no dynamic-array marks to find.
    pub const none: Metadata = .{};

    pub fn deinit(self: *Metadata, allocator: std.mem.Allocator) void {
        allocator.free(self.types);
        allocator.free(self.cell.blocks);
        allocator.free(self.cell.records);
        allocator.free(self.value.blocks);
        allocator.free(self.value.records);
        self.* = undefined;
    }

    pub fn typeAt(self: *const Metadata, one_based: u32) ?TypeEntry {
        if (one_based == 0 or one_based > self.types.len) return null;
        return self.types[one_based - 1];
    }

    fn collection(self: *const Metadata, which: Collection) Blocks {
        return switch (which) {
            .cell_metadata => self.cell,
            .value_metadata => self.value,
        };
    }

    // ── resolution (§5.3b dialect, §5.8b cm/vm) ──
    // The types these speak in are declared in the resolution section
    // below; the functions live here because they are how a parsed part
    // is asked a question.

    /// The dialect this cell was authored in, or the refusal that says
    /// the cell must not be recalculated at all.
    ///
    /// Order is deliberate: **value metadata is checked first**. A cell
    /// that carries both a dynamic-array mark and a rich value must
    /// refuse, not return a dialect a caller might act on before
    /// noticing.
    pub fn resolveCell(self: *const Metadata, role: CellRole, meta: CellMeta) Resolved {
        if (meta.vm != 0) {
            if (self.classifyReference(role, .value_metadata, meta.vm)) |refusal| {
                return .{ .refused = refusal };
            }
            // Every class refuses in `value_metadata`, including the
            // interpreted one, so a null return is impossible here.
            unreachable;
        }
        if (meta.cm == 0) return .{ .ok = .legacy };
        if (self.classifyReference(role, .cell_metadata, meta.cm)) |refusal| {
            return .{ .refused = refusal };
        }
        return .{ .ok = .dynamic_array };
    }

    /// Walks one referenced block and returns the first refusal in
    /// document order, or null when every record in it is interpretable
    /// *in that collection*. Deterministic by position: two runs over the
    /// same bytes name the same record.
    fn classifyReference(
        self: *const Metadata,
        role: CellRole,
        which: Collection,
        one_based: u32,
    ) ?Refusal {
        const site: Refusal.Site = .{ .role = role, .collection = which, .index = one_based };
        const blocks = self.collection(which);
        const blk = blocks.at(one_based) orelse
            return .{ .reason = .index_out_of_range, .site = site };
        const records = blocks.recordsOf(blk);
        if (records.len == 0) return .{ .reason = .empty_record_block, .site = site };

        for (records) |rec| {
            var record_site = site;
            record_site.type_index = rec.type_index;
            const entry = self.typeAt(rec.type_index) orelse
                return .{ .reason = .type_index_out_of_range, .site = record_site };
            const reason: ?Refusal.Reason = switch (entry.class) {
                .dynamic_array_properties => switch (which) {
                    .cell_metadata => null,
                    .value_metadata => .dynamic_array_in_value_metadata,
                },
                .rich_value => .rich_value_metadata,
                .mdx => .mdx_metadata,
                .unknown => .unknown_metadata_type,
            };
            if (reason) |r| return .{ .reason = r, .site = record_site };
        }
        return null;
    }

    /// Resolve every cell a run touches, all or nothing.
    ///
    /// The two passes are the contract, not an implementation detail.
    /// Phase 1 classifies every cell and writes nothing; only once no
    /// refusal is possible does phase 2 fill `out`. A single-pass loop
    /// would leave the prefix before a rich-value cell already resolved —
    /// a partially applied answer to a question the run is about to
    /// refuse, and precisely the shape of "half-done" §3 forbids.
    pub fn resolveRun(
        self: *const Metadata,
        cells: []const RunCell,
        out: []value.Dialect,
    ) RunResolved {
        assert(out.len == cells.len);

        for (cells, 0..) |c, i| {
            switch (self.resolveCell(c.role, c.meta)) {
                .refused => |r| return .{ .refused = .{ .refusal = r, .cell = @intCast(i) } },
                .ok => {},
            }
        }

        for (cells, 0..) |c, i| {
            // `resolveCell` is pure and was just proven not to refuse for
            // this exact argument; the `unreachable` is that proof,
            // written down where a future change would trip over it.
            out[i] = switch (self.resolveCell(c.role, c.meta)) {
                .ok => |d| d,
                .refused => unreachable,
            };
        }
        return .ok;
    }
};

pub const Parsed = union(enum) {
    ok: Metadata,
    refused: Refusal,

    pub fn deinit(self: *Parsed, allocator: std.mem.Allocator) void {
        switch (self.*) {
            .ok => |*m| m.deinit(allocator),
            .refused => {},
        }
        self.* = undefined;
    }
};

// ─── resolution (§5.3b dialect, §5.8b cm/vm) ─────────────────────

/// What one cell declares. Both default to `CT_Cell`'s own default of 0,
/// which means "no metadata".
pub const CellMeta = struct { cm: u32 = 0, vm: u32 = 0 };

pub const Resolved = union(enum) { ok: value.Dialect, refused: Refusal };

pub const RunCell = struct { role: CellRole, meta: CellMeta };

pub const RunRefusal = struct {
    refusal: Refusal,
    /// Index into the caller's `cells`.
    cell: u32,
};

pub const RunResolved = union(enum) { ok, refused: RunRefusal };

// ─── the EvalEnv seam (§5.6a) ────────────────────────────────────

/// Binds a parsed part to `env.DialectResolver`, so `EvalEnv.dialectOf`
/// answers from `cm`/`vm` exactly as a real workbook will at M4b1.
///
/// The refusal is kept here rather than thrown away at the interface:
/// `env.Error.MetadataRefused` says *that* a cell could not be
/// interpreted, and `last_refusal` says which record of which collection
/// said so — the diagnostic §5.7's report has to carry.
pub const CellDialectResolver = struct {
    part: *const Metadata,
    /// Which side of the run is asking. Both refuse identically; the
    /// field is what makes the diagnostic name the right side.
    role: CellRole = .input,
    last_refusal: ?Refusal = null,

    pub fn resolver(self: *CellDialectResolver) env.DialectResolver {
        return .{ .ctx = self, .resolve = resolveThrough };
    }

    fn resolveThrough(ctx: *anyopaque, cm: u32, vm: u32) env.Error!value.Dialect {
        const self: *CellDialectResolver = @ptrCast(@alignCast(ctx));
        switch (self.part.resolveCell(self.role, .{ .cm = cm, .vm = vm })) {
            .ok => |d| return d,
            .refused => |r| {
                self.last_refusal = r;
                return error.MetadataRefused;
            },
        }
    }
};

// ─── XML scanning ────────────────────────────────────────────────
//
// A part-local scanner rather than a shared one: `src/formula/` never
// imports `pkg/` (§3's module law), and the alternative — passing decoded
// structures across the boundary — would mean the engine trusting a
// reader it cannot test. The scanner is strict on purpose. It recognizes
// elements, attributes, comments, processing instructions and CDATA, and
// refuses everything else, including the doctype.

const Tag = struct {
    name: []const u8,
    attrs: []const u8,
    self_closing: bool,
    offset: u32,
};

const Event = union(enum) {
    open: Tag,
    close: struct { name: []const u8, offset: u32 },
    eof,
};

const ScanError = error{Refused};

const Scanner = struct {
    xml: []const u8,
    pos: usize = 0,
    refusal: ?Refusal = null,
    /// Set while skipping a foreign subtree: text is data there, not an
    /// error.
    in_foreign: bool = false,

    fn fail(self: *Scanner, reason: Refusal.Reason, at: usize) ScanError {
        self.refusal = .{ .reason = reason, .offset = @intCast(@min(at, std.math.maxInt(u32))) };
        return error.Refused;
    }

    fn next(self: *Scanner) ScanError!Event {
        while (true) {
            const lt = std.mem.indexOfScalarPos(u8, self.xml, self.pos, '<') orelse {
                if (!self.in_foreign and hasNonSpace(self.xml[self.pos..])) {
                    return self.fail(.unexpected_text, self.pos);
                }
                self.pos = self.xml.len;
                return .eof;
            };
            if (!self.in_foreign and hasNonSpace(self.xml[self.pos..lt])) {
                return self.fail(.unexpected_text, self.pos);
            }
            self.pos = lt;

            if (self.rest().len < 2) return self.fail(.malformed_xml, lt);
            const c1 = self.xml[lt + 1];

            if (c1 == '?') {
                const end = std.mem.indexOfPos(u8, self.xml, lt + 2, "?>") orelse
                    return self.fail(.malformed_xml, lt);
                self.pos = end + 2;
                continue;
            }
            if (c1 == '!') {
                if (std.mem.startsWith(u8, self.xml[lt..], "<!--")) {
                    const end = std.mem.indexOfPos(u8, self.xml, lt + 4, "-->") orelse
                        return self.fail(.malformed_xml, lt);
                    self.pos = end + 3;
                    continue;
                }
                if (std.mem.startsWith(u8, self.xml[lt..], "<![CDATA[")) {
                    const end = std.mem.indexOfPos(u8, self.xml, lt + 9, "]]>") orelse
                        return self.fail(.malformed_xml, lt);
                    if (!self.in_foreign and hasNonSpace(self.xml[lt + 9 .. end])) {
                        return self.fail(.unexpected_text, lt);
                    }
                    self.pos = end + 3;
                    continue;
                }
                return self.fail(.doctype_declaration, lt);
            }

            const gt = std.mem.indexOfScalarPos(u8, self.xml, lt + 1, '>') orelse
                return self.fail(.malformed_xml, lt);
            // A `>` inside an attribute value is legal XML. Re-scan the
            // element body so a name like `<rc t="1" v="2>3"/>` cannot
            // end the tag early.
            const body_end = self.elementEnd(lt + 1, gt) catch |e| return e;
            const inner = self.xml[lt + 1 .. body_end];
            self.pos = body_end + 1;

            if (inner.len == 0) return self.fail(.malformed_xml, lt);
            if (inner[0] == '/') {
                const name = std.mem.trim(u8, inner[1..], white);
                if (name.len == 0) return self.fail(.malformed_xml, lt);
                return .{ .close = .{ .name = localName(name), .offset = @intCast(lt) } };
            }

            var self_closing = false;
            var body = inner;
            if (body[body.len - 1] == '/') {
                self_closing = true;
                body = body[0 .. body.len - 1];
            }
            const name_end = std.mem.indexOfAny(u8, body, white) orelse body.len;
            const raw_name = body[0..name_end];
            if (raw_name.len == 0) return self.fail(.malformed_xml, lt);
            return .{ .open = .{
                .name = localName(raw_name),
                .attrs = body[name_end..],
                .self_closing = self_closing,
                .offset = @intCast(lt),
            } };
        }
    }

    /// Index of the `>` that really closes the element that starts at
    /// `from`, honoring quoted attribute values.
    fn elementEnd(self: *Scanner, from: usize, first_gt: usize) ScanError!usize {
        var i = from;
        var quote: u8 = 0;
        while (i < self.xml.len) : (i += 1) {
            const c = self.xml[i];
            if (quote != 0) {
                if (c == quote) quote = 0;
                continue;
            }
            switch (c) {
                '"', '\'' => quote = c,
                '>' => return i,
                else => {},
            }
        }
        return self.fail(.malformed_xml, first_gt);
    }

    fn rest(self: *const Scanner) []const u8 {
        return self.xml[self.pos..];
    }
};

const white = " \t\r\n";

fn hasNonSpace(s: []const u8) bool {
    return std.mem.trim(u8, s, white).len != 0;
}

/// Namespace prefixes are stripped and the local name matched. The
/// namespace *URI* preflight lands with M4b1 — this reader must not be
/// the place that decides a namespace is acceptable.
fn localName(name: []const u8) []const u8 {
    if (std.mem.lastIndexOfScalar(u8, name, ':')) |c| return name[c + 1 ..];
    return name;
}

const Attr = struct { name: []const u8, value: []const u8 };

const AttrIterator = struct {
    src: []const u8,
    pos: usize = 0,
    refusal: ?Refusal = null,
    /// Byte offset of `src` within the part, for diagnostics.
    base: u32,

    fn fail(self: *AttrIterator, reason: Refusal.Reason) ScanError {
        self.refusal = .{ .reason = reason, .offset = self.base };
        return error.Refused;
    }

    fn next(self: *AttrIterator) ScanError!?Attr {
        while (self.pos < self.src.len and isSpace(self.src[self.pos])) self.pos += 1;
        if (self.pos >= self.src.len) return null;

        const start = self.pos;
        const eq = std.mem.indexOfScalarPos(u8, self.src, start, '=') orelse
            return self.fail(.malformed_xml);
        const name = std.mem.trim(u8, self.src[start..eq], white);
        if (name.len == 0) return self.fail(.malformed_xml);
        // A name with a space in it means the previous attribute had no
        // value — `<c a b="1">`. Refusing here keeps that from arriving
        // downstream disguised as an unknown attribute.
        if (std.mem.indexOfAny(u8, name, white) != null) return self.fail(.malformed_xml);

        var i = eq + 1;
        while (i < self.src.len and isSpace(self.src[i])) i += 1;
        if (i >= self.src.len) return self.fail(.malformed_xml);
        const quote = self.src[i];
        if (quote != '"' and quote != '\'') return self.fail(.malformed_xml);
        const close = std.mem.indexOfScalarPos(u8, self.src, i + 1, quote) orelse
            return self.fail(.malformed_xml);
        self.pos = close + 1;
        return .{ .name = name, .value = self.src[i + 1 .. close] };
    }
};

fn isSpace(c: u8) bool {
    return c == ' ' or c == '\t' or c == '\r' or c == '\n';
}

/// `xmlns` and `xmlns:*` are namespace machinery, not schema attributes,
/// and appear on elements the schema declares attribute-free.
fn isNamespaceDecl(name: []const u8) bool {
    return std.mem.eql(u8, name, "xmlns") or std.mem.startsWith(u8, name, "xmlns:");
}

/// `xsd:unsignedInt`, strictly: digits only, no sign, no leading `+`, and
/// it must fit. Excel writes canonical integers; anything else here is a
/// producer we should not be guessing about.
fn parseUnsignedInt(text: []const u8) ?u32 {
    if (text.len == 0) return null;
    for (text) |c| {
        if (c < '0' or c > '9') return null;
    }
    return std.fmt.parseInt(u32, text, 10) catch null;
}

/// `xsd:boolean` has four lexical forms and Office writes two of them.
/// Parsed and discarded for every flag but `cellMeta` — discarded after
/// *validation*, so a malformed one still refuses.
fn parseBool(text: []const u8) ?bool {
    if (std.mem.eql(u8, text, "1") or std.mem.eql(u8, text, "true")) return true;
    if (std.mem.eql(u8, text, "0") or std.mem.eql(u8, text, "false")) return false;
    return null;
}

// ─── parse ───────────────────────────────────────────────────────

const Builder = struct {
    allocator: std.mem.Allocator,
    limits: Limits,
    types: std.ArrayListUnmanaged(TypeEntry) = .empty,
    cell_blocks: std.ArrayListUnmanaged(Block) = .empty,
    cell_records: std.ArrayListUnmanaged(Record) = .empty,
    value_blocks: std.ArrayListUnmanaged(Block) = .empty,
    value_records: std.ArrayListUnmanaged(Record) = .empty,
    counts: [element_count]u32 = @splat(0),
    declared_type_count: ?u32 = null,
    declared_cell_count: ?u32 = null,
    declared_value_count: ?u32 = null,
    seen_cell_metadata: bool = false,
    seen_value_metadata: bool = false,
    /// Which collection the open `<bk>` belongs to.
    current: ?Collection = null,
    block_start: u32 = 0,
    /// Where the open `<bk>` began, so a refusal raised when it closes
    /// still points at the element that caused it.
    block_offset: u32 = 0,

    fn deinit(self: *Builder) void {
        self.types.deinit(self.allocator);
        self.cell_blocks.deinit(self.allocator);
        self.cell_records.deinit(self.allocator);
        self.value_blocks.deinit(self.allocator);
        self.value_records.deinit(self.allocator);
    }
};

const ParseError = error{ OutOfMemory, Refused };

/// Read a metadata part. The result borrows every name from `xml`, which
/// must outlive it; free with `parsed.deinit(allocator)`.
///
/// Refusals here are *structural* — the part itself is unreadable or
/// holds something unclassifiable. A part that merely contains rich
/// values parses cleanly; the refusal for those is a property of the
/// cells that reference them (`resolveCell`).
pub fn parse(
    allocator: std.mem.Allocator,
    xml: []const u8,
    opts: Options,
) error{OutOfMemory}!Parsed {
    if (xml.len > opts.limits.max_bytes) {
        return .{ .refused = .{ .reason = .limit_exceeded, .limit = .part_bytes } };
    }
    if (!std.unicode.utf8ValidateSlice(xml)) {
        return .{ .refused = .{ .reason = .invalid_utf8 } };
    }

    var b: Builder = .{ .allocator = allocator, .limits = opts.limits };
    errdefer b.deinit();

    var sc: Scanner = .{ .xml = xml };
    var refusal: ?Refusal = null;
    run(&b, &sc, &refusal) catch |err| switch (err) {
        error.OutOfMemory => return error.OutOfMemory,
        error.Refused => {
            b.deinit();
            return .{ .refused = refusal orelse sc.refusal orelse
                .{ .reason = .malformed_xml } };
        },
    };

    var md: Metadata = .{
        .counts = b.counts,
        .declared_type_count = b.declared_type_count,
    };
    md.types = try b.types.toOwnedSlice(allocator);
    errdefer allocator.free(md.types);
    md.cell = .{
        .blocks = try b.cell_blocks.toOwnedSlice(allocator),
        .declared_count = b.declared_cell_count,
    };
    errdefer allocator.free(md.cell.blocks);
    md.cell.records = try b.cell_records.toOwnedSlice(allocator);
    errdefer allocator.free(md.cell.records);
    md.value = .{
        .blocks = try b.value_blocks.toOwnedSlice(allocator),
        .declared_count = b.declared_value_count,
    };
    errdefer allocator.free(md.value.blocks);
    md.value.records = try b.value_records.toOwnedSlice(allocator);
    return .{ .ok = md };
}

fn run(b: *Builder, sc: *Scanner, refusal: *?Refusal) ParseError!void {
    var stack: [16]Element = undefined;
    var depth: u32 = 0;
    var skip_depth: u32 = 0;
    var saw_root = false;

    while (true) {
        const ev = try sc.next();
        switch (ev) {
            .eof => {
                if (depth != 0 or skip_depth != 0) {
                    refusal.* = .{ .reason = .malformed_xml, .offset = @intCast(sc.xml.len) };
                    return error.Refused;
                }
                if (!saw_root) {
                    refusal.* = .{ .reason = .malformed_xml };
                    return error.Refused;
                }
                return;
            },
            .close => |c| {
                if (skip_depth > 0) {
                    skip_depth -= 1;
                    if (skip_depth == 0) sc.in_foreign = false;
                    continue;
                }
                if (depth == 0) {
                    refusal.* = .{ .reason = .malformed_xml, .offset = c.offset };
                    return error.Refused;
                }
                const open_el = stack[depth - 1];
                if (childOf(parentOf(&stack, depth), c.name) != open_el) {
                    refusal.* = .{ .reason = .malformed_xml, .offset = c.offset };
                    return error.Refused;
                }
                try closeElement(b, open_el, refusal);
                depth -= 1;
            },
            .open => |t| {
                if (skip_depth > 0) {
                    if (!t.self_closing) {
                        skip_depth += 1;
                        if (skip_depth > b.limits.max_skip_depth) {
                            refusal.* = .{
                                .reason = .limit_exceeded,
                                .limit = .skip_depth,
                                .offset = t.offset,
                            };
                            return error.Refused;
                        }
                    }
                    continue;
                }

                const parent: ?Element = if (depth == 0) null else stack[depth - 1];
                if (depth == 0 and saw_root) {
                    refusal.* = .{ .reason = .malformed_xml, .offset = t.offset };
                    return error.Refused;
                }
                const el = childOf(parent, t.name) orelse {
                    refusal.* = .{ .reason = .unexpected_element, .offset = t.offset };
                    return error.Refused;
                };
                if (depth == 0) saw_root = true;
                b.counts[@intFromEnum(el)] +|= 1;

                openElement(b, el, t, refusal) catch |err| switch (err) {
                    error.OutOfMemory => return error.OutOfMemory,
                    error.Refused => return error.Refused,
                };

                if (el == .ext) {
                    // Foreign content: well-formedness only from here
                    // down. Nothing inside is classified, because nothing
                    // inside is reachable without a type that refuses.
                    if (!t.self_closing) {
                        skip_depth = 1;
                        sc.in_foreign = true;
                    }
                    continue;
                }

                if (t.self_closing) {
                    try closeElement(b, el, refusal);
                } else {
                    if (depth >= stack.len or depth >= b.limits.max_depth) {
                        refusal.* = .{
                            .reason = .limit_exceeded,
                            .limit = .depth,
                            .offset = t.offset,
                        };
                        return error.Refused;
                    }
                    stack[depth] = el;
                    depth += 1;
                }
            },
        }
    }
}

fn parentOf(stack: *const [16]Element, depth: u32) ?Element {
    if (depth < 2) return null;
    return stack[depth - 2];
}

fn openElement(b: *Builder, el: Element, t: Tag, refusal: *?Refusal) ParseError!void {
    // An exhaustive switch: a new `Element` cannot land without a
    // decision about what reading it means.
    switch (el) {
        .metadata => try noAttrs(b, t, refusal),
        .metadata_types => b.declared_type_count = try countAttr(b, t, refusal),
        .metadata_type => try addType(b, t, refusal),

        .cell_metadata => {
            if (b.seen_cell_metadata) return refuse(refusal, .duplicate_collection, t.offset);
            b.seen_cell_metadata = true;
            b.current = .cell_metadata;
            b.declared_cell_count = try countAttr(b, t, refusal);
        },
        .value_metadata => {
            if (b.seen_value_metadata) return refuse(refusal, .duplicate_collection, t.offset);
            b.seen_value_metadata = true;
            b.current = .value_metadata;
            b.declared_value_count = try countAttr(b, t, refusal);
        },
        .metadata_block => {
            try noAttrs(b, t, refusal);
            b.block_start = @intCast(recordsOf(b).items.len);
            b.block_offset = t.offset;
        },
        .metadata_record => try addRecord(b, t, refusal),

        // Inert: recognized, counted, and left alone. Their payloads
        // belong to types that refuse the moment a cell points at them,
        // so reading them would buy nothing and cost a parse surface.
        .metadata_strings,
        .metadata_string,
        .mdx_metadata,
        .mdx,
        .mdx_tuple,
        .mdx_set,
        .mdx_member_prop,
        .mdx_kpi,
        .mdx_string_index,
        .future_metadata,
        .future_metadata_block,
        .ext_lst,
        .ext,
        => {},
    }
}

fn closeElement(b: *Builder, el: Element, refusal: *?Refusal) ParseError!void {
    switch (el) {
        .metadata_block => {
            // `childOf` only admits `bk` under a collection, so the
            // collection is open by construction.
            const which = b.current.?;
            const recs = recordsOf(b);
            const first = b.block_start;
            const len: u32 = @intCast(recs.items.len - first);
            const blocks = switch (which) {
                .cell_metadata => &b.cell_blocks,
                .value_metadata => &b.value_blocks,
            };
            if (blocks.items.len >= b.limits.max_blocks) {
                return refuseLimit(refusal, .blocks, b.block_offset);
            }
            try blocks.append(b.allocator, .{ .first = first, .len = len });
        },
        .cell_metadata, .value_metadata => b.current = null,
        else => {},
    }
}

fn recordsOf(b: *Builder) *std.ArrayListUnmanaged(Record) {
    return switch (b.current.?) {
        .cell_metadata => &b.cell_records,
        .value_metadata => &b.value_records,
    };
}

fn refuse(refusal: *?Refusal, reason: Refusal.Reason, offset: u32) ParseError {
    refusal.* = .{ .reason = reason, .offset = offset };
    return error.Refused;
}

fn refuseLimit(refusal: *?Refusal, limit: LimitKind, offset: u32) ParseError {
    refusal.* = .{ .reason = .limit_exceeded, .limit = limit, .offset = offset };
    return error.Refused;
}

fn attrs(t: Tag) AttrIterator {
    return .{ .src = t.attrs, .base = t.offset };
}

fn noAttrs(b: *Builder, t: Tag, refusal: *?Refusal) ParseError!void {
    _ = b;
    var it = attrs(t);
    while (it.next() catch {
        refusal.* = it.refusal;
        return error.Refused;
    }) |a| {
        if (isNamespaceDecl(a.name)) continue;
        return refuse(refusal, .unexpected_attribute, t.offset);
    }
}

fn countAttr(b: *Builder, t: Tag, refusal: *?Refusal) ParseError!?u32 {
    _ = b;
    var out: ?u32 = null;
    var it = attrs(t);
    while (it.next() catch {
        refusal.* = it.refusal;
        return error.Refused;
    }) |a| {
        if (isNamespaceDecl(a.name)) continue;
        if (!std.mem.eql(u8, localName(a.name), "count")) {
            return refuse(refusal, .unexpected_attribute, t.offset);
        }
        out = parseUnsignedInt(a.value) orelse
            return refuse(refusal, .bad_attribute_value, t.offset);
    }
    return out;
}

fn addType(b: *Builder, t: Tag, refusal: *?Refusal) ParseError!void {
    if (b.types.items.len >= b.limits.max_types) {
        return refuseLimit(refusal, .types, t.offset);
    }

    var name: ?[]const u8 = null;
    var min_version: ?u32 = null;
    var cell_meta = false;

    var it = attrs(t);
    while (it.next() catch {
        refusal.* = it.refusal;
        return error.Refused;
    }) |a| {
        if (isNamespaceDecl(a.name)) continue;
        const local = localName(a.name);
        if (!knownTypeAttr(local)) return refuse(refusal, .unexpected_attribute, t.offset);

        if (std.mem.eql(u8, local, "name")) {
            name = a.value;
        } else if (std.mem.eql(u8, local, "minSupportedVersion")) {
            min_version = parseUnsignedInt(a.value) orelse
                return refuse(refusal, .bad_attribute_value, t.offset);
        } else {
            // Every remaining attribute is `xsd:boolean`. Validated, then
            // kept only where it means something to M4a.
            const flag = parseBool(a.value) orelse
                return refuse(refusal, .bad_attribute_value, t.offset);
            if (std.mem.eql(u8, local, "cellMeta")) cell_meta = flag;
        }
    }

    const n = name orelse return refuse(refusal, .missing_attribute, t.offset);
    try b.types.append(b.allocator, .{
        .name = n,
        .class = classifyTypeName(n),
        .min_supported_version = min_version,
        .cell_meta = cell_meta,
    });
}

fn knownTypeAttr(local: []const u8) bool {
    for (metadata_type_attrs) |known| {
        if (std.mem.eql(u8, known, local)) return true;
    }
    return false;
}

fn addRecord(b: *Builder, t: Tag, refusal: *?Refusal) ParseError!void {
    var type_index: ?u32 = null;
    var value_index: u32 = 0;

    var it = attrs(t);
    while (it.next() catch {
        refusal.* = it.refusal;
        return error.Refused;
    }) |a| {
        if (isNamespaceDecl(a.name)) continue;
        const local = localName(a.name);
        if (std.mem.eql(u8, local, "t")) {
            type_index = parseUnsignedInt(a.value) orelse
                return refuse(refusal, .bad_attribute_value, t.offset);
        } else if (std.mem.eql(u8, local, "v")) {
            value_index = parseUnsignedInt(a.value) orelse
                return refuse(refusal, .bad_attribute_value, t.offset);
        } else {
            return refuse(refusal, .unexpected_attribute, t.offset);
        }
    }

    // `t` is what classification turns on, so a record without one is
    // unclassifiable rather than defaultable. `v` defaults to 0: M4a
    // never dereferences it.
    const ti = type_index orelse return refuse(refusal, .missing_attribute, t.offset);
    const recs = recordsOf(b);
    if (recs.items.len >= b.limits.max_records) {
        return refuseLimit(refusal, .records, t.offset);
    }
    try recs.append(b.allocator, .{ .type_index = ti, .value_index = value_index });
}

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;
const coords = @import("zlsx_refs");

fn cellOf(a1: []const u8) coords.Cell {
    return coords.parseCell(a1, .{ .dollar = .accept }) catch unreachable;
}

/// The shape Office writes for a workbook with one dynamic-array formula
/// (attributes as `XlsxWriter`'s `metadata.py` and Excel emit them).
/// Spec-pinned; no committed oracle decides it.
const fixture_dynamic_array =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">
    \\  <metadataTypes count="1">
    \\    <metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>
    \\  </metadataTypes>
    \\  <futureMetadata name="XLDAPR" count="1">
    \\    <bk>
    \\      <extLst>
    \\        <ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">
    \\          <xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>
    \\        </ext>
    \\      </extLst>
    \\    </bk>
    \\  </futureMetadata>
    \\  <cellMetadata count="1">
    \\    <bk>
    \\      <rc t="1" v="0"/>
    \\    </bk>
    \\  </cellMetadata>
    \\</metadata>
;

/// A workbook carrying both a dynamic array and rich values: XLDAPR in
/// `cellMetadata`, XLRICHVALUE in `valueMetadata`. This is the file shape
/// that makes "ignore what you do not understand" lethal.
const fixture_rich_value =
    \\<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    \\  <metadataTypes count="2">
    \\    <metadataType name="XLDAPR" minSupportedVersion="120000" cellMeta="1"/>
    \\    <metadataType name="XLRICHVALUE" minSupportedVersion="120000"/>
    \\  </metadataTypes>
    \\  <futureMetadata name="XLRICHVALUE" count="1">
    \\    <bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="0" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"/></ext></extLst></bk>
    \\  </futureMetadata>
    \\  <cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>
    \\  <valueMetadata count="1"><bk><rc t="2" v="0"/></bk></valueMetadata>
    \\</metadata>
;

/// Every element the inventory lists, in one part.
const fixture_all_elements =
    \\<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    \\  <metadataTypes count="3">
    \\    <metadataType name="XLDAPR" cellMeta="1"/>
    \\    <metadataType name="XLRICHVALUE"/>
    \\    <metadataType name="XLMDX"/>
    \\  </metadataTypes>
    \\  <metadataStrings count="2"><s v="[Measures]"/><s v="[Time]"/></metadataStrings>
    \\  <mdxMetadata count="1">
    \\    <mdx n="0" f="m">
    \\      <t c="1" si="0"><n x="0"/></t>
    \\      <ms ns="1" c="1"><n x="1"/></ms>
    \\      <p n="0" np="1"/>
    \\      <k n="0" np="1" p="v"/>
    \\    </mdx>
    \\  </mdxMetadata>
    \\  <futureMetadata name="XLDAPR" count="1">
    \\    <bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><foreign><deeper/></foreign></ext></extLst></bk>
    \\  </futureMetadata>
    \\  <cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>
    \\  <valueMetadata count="2"><bk><rc t="2" v="0"/></bk><bk><rc t="3" v="1"/></bk></valueMetadata>
    \\  <extLst><ext uri="{00000000-0000-0000-0000-000000000000}"><anything/></ext></extLst>
    \\</metadata>
;

fn parseOk(allocator: std.mem.Allocator, xml: []const u8) !Metadata {
    var p = try parse(allocator, xml, .{});
    switch (p) {
        .ok => |m| return m,
        .refused => |r| {
            std.debug.print("unexpected refusal: {t} at {d}\n", .{ r.reason, r.offset });
            p.deinit(allocator);
            return error.TestUnexpectedRefusal;
        },
    }
}

fn parseRefused(allocator: std.mem.Allocator, xml: []const u8) !Refusal {
    var p = try parse(allocator, xml, .{});
    switch (p) {
        .ok => {
            p.deinit(allocator);
            return error.TestExpectedRefusal;
        },
        .refused => |r| return r,
    }
}

test "element inventory: every schema element is classified exactly once" {
    // The count is the table's own gate: a schema element that grew a row
    // without a decision, or a row that lost its element, fails here.
    try testing.expectEqual(@as(usize, 20), element_inventory.len);
    try testing.expectEqual(element_count, element_inventory.len);

    inline for (@typeInfo(Element).@"enum".fields) |f| {
        const el: Element = @enumFromInt(f.value);
        var seen: usize = 0;
        for (element_inventory) |row| {
            if (row.element == el) seen += 1;
        }
        try testing.expectEqual(@as(usize, 1), seen);
    }

    // No row is decoration: every row is reachable from `childOf` under
    // each parent it claims, and the table is what `childOf` reads.
    for (element_inventory) |row| {
        try testing.expect(row.parents.len > 0);
        try testing.expect(row.xml_name.len > 0);
        try testing.expect(row.note.len > 0);
        for (row.parents) |p| {
            try testing.expectEqual(row.element, childOf(p, row.xml_name).?);
        }
    }

    // The overloaded names really are position-resolved.
    try testing.expectEqual(Element.metadata_block, childOf(.cell_metadata, "bk").?);
    try testing.expectEqual(Element.metadata_block, childOf(.value_metadata, "bk").?);
    try testing.expectEqual(Element.future_metadata_block, childOf(.future_metadata, "bk").?);
    try testing.expectEqual(Element.mdx_tuple, childOf(.mdx, "t").?);
    try testing.expect(childOf(.metadata_block, "t") == null);
    try testing.expect(childOf(null, "worksheet") == null);
}

test "element inventory: exactly six elements are interpreted" {
    var interpreted: usize = 0;
    for (element_inventory) |row| {
        switch (row.treatment) {
            .interpreted => interpreted += 1,
            .inert => {},
        }
        try testing.expectEqual(row.treatment, treatmentOf(row.element));
    }
    // metadata, metadataTypes, metadataType, cellMetadata, valueMetadata,
    // bk(metadata_block), rc(metadata_record) — seven, and naming them
    // here is how a silent promotion of an inert element gets caught.
    try testing.expectEqual(@as(usize, 7), interpreted);
    for ([_]Element{
        .metadata,
        .metadata_types,
        .metadata_type,
        .cell_metadata,
        .value_metadata,
        .metadata_block,
        .metadata_record,
    }) |el| {
        try testing.expectEqual(ElementTreatment.interpreted, treatmentOf(el));
    }
}

test "type classification: only XLDAPR is interpreted, and unknown has a row" {
    try testing.expectEqual(@typeInfo(TypeClass).@"enum".fields.len, type_classification.len);

    var interpreted: usize = 0;
    inline for (@typeInfo(TypeClass).@"enum".fields) |f| {
        const class: TypeClass = @enumFromInt(f.value);
        var seen: usize = 0;
        for (type_classification) |row| {
            if (row.class != class) continue;
            seen += 1;
            if (row.treatment == .interpreted) interpreted += 1;
        }
        try testing.expectEqual(@as(usize, 1), seen);
    }
    try testing.expectEqual(@as(usize, 1), interpreted);
    try testing.expectEqual(TypeTreatment.interpreted, treatmentOfClass(.dynamic_array_properties));

    try testing.expectEqual(TypeClass.dynamic_array_properties, classifyTypeName("XLDAPR"));
    try testing.expectEqual(TypeClass.rich_value, classifyTypeName("XLRICHVALUE"));
    try testing.expectEqual(TypeClass.mdx, classifyTypeName("XLMDX"));
    // Case and spelling are not negotiated.
    try testing.expectEqual(TypeClass.unknown, classifyTypeName("xldapr"));
    try testing.expectEqual(TypeClass.unknown, classifyTypeName("XLDAPR "));
    try testing.expectEqual(TypeClass.unknown, classifyTypeName(""));
    try testing.expectEqual(TypeClass.unknown, classifyTypeName("XLFUTURETHING"));
}

test "refusal: every reason has a plane-2 error, and the split is the documented one" {
    inline for (@typeInfo(Refusal.Reason).@"enum".fields) |f| {
        const reason: Refusal.Reason = @enumFromInt(f.value);
        const r: Refusal = .{ .reason = reason };
        const plane = r.planeTwo();
        // Nothing in this file may raise a plane-2 error M2 already owns
        // for another purpose, and nothing may claim a function refusal.
        try testing.expect(plane != .FormulaUnsupportedFunction);
    }
    try testing.expectEqual(PlaneTwo.FormulaUnsupportedConstruct, (Refusal{ .reason = .rich_value_metadata }).planeTwo());
    try testing.expectEqual(PlaneTwo.FormulaMalformedInput, (Refusal{ .reason = .malformed_xml }).planeTwo());
    try testing.expectEqual(PlaneTwo.FormulaLimitExceeded, (Refusal{ .reason = .limit_exceeded }).planeTwo());
}

test "parse: the dynamic-array part Office writes" {
    var md = try parseOk(testing.allocator, fixture_dynamic_array);
    defer md.deinit(testing.allocator);

    try testing.expectEqual(@as(usize, 1), md.types.len);
    try testing.expectEqualStrings("XLDAPR", md.types[0].name);
    try testing.expectEqual(TypeClass.dynamic_array_properties, md.types[0].class);
    try testing.expectEqual(@as(?u32, 120000), md.types[0].min_supported_version);
    try testing.expect(md.types[0].cell_meta);

    try testing.expectEqual(@as(usize, 1), md.cell.blocks.len);
    try testing.expectEqual(@as(usize, 1), md.cell.records.len);
    try testing.expectEqual(@as(u32, 1), md.cell.records[0].type_index);
    try testing.expectEqual(@as(u32, 0), md.cell.records[0].value_index);
    try testing.expectEqual(@as(usize, 0), md.value.blocks.len);

    // The `count` hints are recorded and not enforced.
    try testing.expectEqual(@as(?u32, 1), md.declared_type_count);
    try testing.expectEqual(@as(?u32, 1), md.cell.declared_count);
}

test "parse: every element in the inventory is reached, and none is skipped silently" {
    var md = try parseOk(testing.allocator, fixture_all_elements);
    defer md.deinit(testing.allocator);

    inline for (@typeInfo(Element).@"enum".fields) |f| {
        const el: Element = @enumFromInt(f.value);
        if (md.counts[f.value] == 0) {
            std.debug.print("element never reached: {t}\n", .{el});
            return error.TestUnreachedElement;
        }
    }
    // Two `ext` elements, one of them holding foreign content that is
    // skipped rather than classified — the nested `<deeper/>` is not an
    // element of this schema and must not refuse.
    try testing.expectEqual(@as(u32, 2), md.counts[@intFromEnum(Element.ext)]);
    try testing.expectEqual(@as(u32, 3), md.types.len);
    try testing.expectEqual(@as(usize, 1), md.cell.blocks.len);
    try testing.expectEqual(@as(usize, 2), md.value.blocks.len);
}

test "resolve: an XLDAPR-marked cell is dynamic_array, an unmarked one is legacy" {
    var md = try parseOk(testing.allocator, fixture_dynamic_array);
    defer md.deinit(testing.allocator);

    for ([_]CellRole{ .input, .result }) |role| {
        try testing.expectEqual(
            value.Dialect.dynamic_array,
            md.resolveCell(role, .{ .cm = 1 }).ok,
        );
        try testing.expectEqual(value.Dialect.legacy, md.resolveCell(role, .{}).ok);
    }

    // A workbook with no metadata part at all is legacy everywhere.
    const none = Metadata.none;
    try testing.expectEqual(value.Dialect.legacy, none.resolveCell(.input, .{}).ok);
    try testing.expectEqual(
        Refusal.Reason.index_out_of_range,
        none.resolveCell(.input, .{ .cm = 1 }).refused.reason,
    );
}

test "resolve: XLRICHVALUE refuses on an input cell and on a result cell" {
    var md = try parseOk(testing.allocator, fixture_rich_value);
    defer md.deinit(testing.allocator);

    // The same workbook's dynamic-array mark still resolves: the refusal
    // belongs to the reference, not to the file.
    try testing.expectEqual(value.Dialect.dynamic_array, md.resolveCell(.input, .{ .cm = 1 }).ok);

    for ([_]CellRole{ .input, .result }) |role| {
        const r = md.resolveCell(role, .{ .vm = 1 }).refused;
        try testing.expectEqual(Refusal.Reason.rich_value_metadata, r.reason);
        try testing.expectEqual(PlaneTwo.FormulaUnsupportedConstruct, r.planeTwo());
        try testing.expectEqual(role, r.site.?.role);
        try testing.expectEqual(Collection.value_metadata, r.site.?.collection);
        try testing.expectEqual(@as(u32, 1), r.site.?.index);
        try testing.expectEqual(@as(?u32, 2), r.site.?.type_index);

        // A rich value on a cell that is ALSO dynamic-array marked
        // refuses; the dialect never wins over the data loss.
        const both = md.resolveCell(role, .{ .cm = 1, .vm = 1 }).refused;
        try testing.expectEqual(Refusal.Reason.rich_value_metadata, both.reason);
    }
}

test "resolve: unknown and MDX metadata refuse, in either collection" {
    const xml =
        \\<metadata>
        \\  <metadataTypes count="3">
        \\    <metadataType name="XLFUTURE"/>
        \\    <metadataType name="XLMDX"/>
        \\    <metadataType name="XLDAPR" cellMeta="1"/>
        \\  </metadataTypes>
        \\  <cellMetadata count="2"><bk><rc t="1" v="0"/></bk><bk><rc t="2" v="0"/></bk></cellMetadata>
        \\  <valueMetadata count="1"><bk><rc t="3" v="0"/></bk></valueMetadata>
        \\</metadata>
    ;
    var md = try parseOk(testing.allocator, xml);
    defer md.deinit(testing.allocator);

    try testing.expectEqual(
        Refusal.Reason.unknown_metadata_type,
        md.resolveCell(.input, .{ .cm = 1 }).refused.reason,
    );
    try testing.expectEqual(
        Refusal.Reason.mdx_metadata,
        md.resolveCell(.result, .{ .cm = 2 }).refused.reason,
    );
    // XLDAPR through `vm` is not a dialect answer — it is a shape this
    // reader has never been shown.
    try testing.expectEqual(
        Refusal.Reason.dynamic_array_in_value_metadata,
        md.resolveCell(.input, .{ .vm = 1 }).refused.reason,
    );
}

test "resolve: broken indexes refuse rather than defaulting to a dialect" {
    const xml =
        \\<metadata>
        \\  <metadataTypes count="1"><metadataType name="XLDAPR" cellMeta="1"/></metadataTypes>
        \\  <cellMetadata count="3"><bk><rc t="1" v="0"/></bk><bk/><bk><rc t="9" v="0"/></bk></cellMetadata>
        \\</metadata>
    ;
    var md = try parseOk(testing.allocator, xml);
    defer md.deinit(testing.allocator);

    try testing.expectEqual(value.Dialect.dynamic_array, md.resolveCell(.input, .{ .cm = 1 }).ok);
    try testing.expectEqual(
        Refusal.Reason.empty_record_block,
        md.resolveCell(.input, .{ .cm = 2 }).refused.reason,
    );
    try testing.expectEqual(
        Refusal.Reason.type_index_out_of_range,
        md.resolveCell(.input, .{ .cm = 3 }).refused.reason,
    );
    try testing.expectEqual(
        Refusal.Reason.index_out_of_range,
        md.resolveCell(.input, .{ .cm = 4 }).refused.reason,
    );
}

test "resolveRun: a refusal leaves not one dialect written" {
    var md = try parseOk(testing.allocator, fixture_rich_value);
    defer md.deinit(testing.allocator);

    // Three resolvable cells, then the rich-value one. A single-pass
    // implementation would have written `.legacy` into slots 0..2 before
    // discovering the refusal; the poison value proves it did not.
    const cells = [_]RunCell{
        .{ .role = .input, .meta = .{} },
        .{ .role = .input, .meta = .{ .cm = 1 } },
        .{ .role = .result, .meta = .{} },
        .{ .role = .result, .meta = .{ .vm = 1 } },
    };
    var out: [cells.len]value.Dialect = @splat(.dynamic_array);

    const r = md.resolveRun(&cells, &out).refused;
    try testing.expectEqual(@as(u32, 3), r.cell);
    try testing.expectEqual(Refusal.Reason.rich_value_metadata, r.refusal.reason);
    try testing.expectEqual(CellRole.result, r.refusal.site.?.role);
    for (out) |d| try testing.expectEqual(value.Dialect.dynamic_array, d);

    // Drop the offending cell and the same call fills every slot.
    const clean = cells[0..3];
    var ok_out: [3]value.Dialect = @splat(.dynamic_array);
    try testing.expect(md.resolveRun(clean, &ok_out) == .ok);
    try testing.expectEqual(value.Dialect.legacy, ok_out[0]);
    try testing.expectEqual(value.Dialect.dynamic_array, ok_out[1]);
    try testing.expectEqual(value.Dialect.legacy, ok_out[2]);
}

test "EvalEnv: dialectOf answers from cm/vm, and a rich value refuses there too" {
    var md = try parseOk(testing.allocator, fixture_rich_value);
    defer md.deinit(testing.allocator);

    var fake = env.Fake.init(testing.allocator);
    defer fake.deinit();
    const sh = try fake.addSheet("Data");

    const marked = cellOf("A1");
    const plain = cellOf("B1");
    const rich = cellOf("C1");

    // Every cell stores `.dynamic_array` in the legacy field, so a
    // `.legacy` answer can only have come through the metadata part.
    try fake.put(sh, .stored, .{
        .row = marked.row,
        .col = marked.col,
        .v = value.ScalarValue.fromNumber(1),
        .cm = 1,
    });
    try fake.put(sh, .stored, .{
        .row = plain.row,
        .col = plain.col,
        .v = value.ScalarValue.fromNumber(2),
    });
    try fake.put(sh, .stored, .{
        .row = rich.row,
        .col = rich.col,
        .v = value.ScalarValue.fromNumber(3),
        .vm = 1,
    });

    var bridge: CellDialectResolver = .{ .part = &md };
    fake.dialect_resolver = bridge.resolver();
    const e = fake.evalEnv();

    try testing.expectEqual(
        value.Dialect.dynamic_array,
        try e.dialectOf(.{ .sheet = sh, .row = marked.row, .col = marked.col }),
    );
    try testing.expectEqual(
        value.Dialect.legacy,
        try e.dialectOf(.{ .sheet = sh, .row = plain.row, .col = plain.col }),
    );
    try testing.expectError(
        error.MetadataRefused,
        e.dialectOf(.{ .sheet = sh, .row = rich.row, .col = rich.col }),
    );
    try testing.expectEqual(Refusal.Reason.rich_value_metadata, bridge.last_refusal.?.reason);
    try testing.expectEqual(CellRole.input, bridge.last_refusal.?.site.?.role);

    // The result side asks the same question and gets the same refusal,
    // labelled for the side that asked.
    bridge.role = .result;
    bridge.last_refusal = null;
    try testing.expectError(
        error.MetadataRefused,
        e.dialectOf(.{ .sheet = sh, .row = rich.row, .col = rich.col }),
    );
    try testing.expectEqual(CellRole.result, bridge.last_refusal.?.site.?.role);
}

test "parse: structural refusals name what they refused" {
    const cases = [_]struct { xml: []const u8, reason: Refusal.Reason }{
        .{ .xml = "<metadata><pivotCache/></metadata>", .reason = .unexpected_element },
        .{ .xml = "<worksheet/>", .reason = .unexpected_element },
        .{ .xml = "<metadata><metadataTypes><rc t=\"1\"/></metadataTypes></metadata>", .reason = .unexpected_element },
        .{ .xml = "<!DOCTYPE metadata><metadata/>", .reason = .doctype_declaration },
        .{ .xml = "<metadata><cellMetadata/><cellMetadata/></metadata>", .reason = .duplicate_collection },
        .{ .xml = "<metadata unknown=\"1\"/>", .reason = .unexpected_attribute },
        .{ .xml = "<metadata><metadataTypes><metadataType name=\"X\" nope=\"1\"/></metadataTypes></metadata>", .reason = .unexpected_attribute },
        .{ .xml = "<metadata><metadataTypes><metadataType minSupportedVersion=\"1\"/></metadataTypes></metadata>", .reason = .missing_attribute },
        .{ .xml = "<metadata><metadataTypes><metadataType name=\"X\" cellMeta=\"yes\"/></metadataTypes></metadata>", .reason = .bad_attribute_value },
        .{ .xml = "<metadata><metadataTypes count=\"-1\"/></metadata>", .reason = .bad_attribute_value },
        .{ .xml = "<metadata><cellMetadata><bk><rc t=\"x\"/></bk></cellMetadata></metadata>", .reason = .bad_attribute_value },
        .{ .xml = "<metadata><cellMetadata><bk><rc v=\"0\"/></bk></cellMetadata></metadata>", .reason = .missing_attribute },
        .{ .xml = "<metadata><cellMetadata></metadata>", .reason = .malformed_xml },
        .{ .xml = "<metadata>text</metadata>", .reason = .unexpected_text },
        .{ .xml = "<metadata", .reason = .malformed_xml },
        .{ .xml = "", .reason = .malformed_xml },
        .{ .xml = "<metadata/><metadata/>", .reason = .malformed_xml },
        .{ .xml = "<metadata/>\xFF", .reason = .invalid_utf8 },
    };
    for (cases) |c| {
        const r = try parseRefused(testing.allocator, c.xml);
        if (r.reason != c.reason) {
            std.debug.print("input {s}: expected {t}, got {t}\n", .{ c.xml, c.reason, r.reason });
            return error.TestWrongRefusal;
        }
    }
}

test "parse: an empty part and a part with no collections are both legal" {
    var md = try parseOk(testing.allocator, "<metadata/>");
    defer md.deinit(testing.allocator);
    try testing.expectEqual(@as(usize, 0), md.types.len);
    try testing.expectEqual(value.Dialect.legacy, md.resolveCell(.input, .{}).ok);
}

test "parse: namespace prefixes, quoting, and comments do not change the answer" {
    const prefixed =
        \\<?xml version="1.0"?>
        \\<!-- Office writes a comment here in some builds -->
        \\<x:metadata xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        \\  <x:metadataTypes count='1'><x:metadataType name='XLDAPR' cellMeta='1'/></x:metadataTypes>
        \\  <x:cellMetadata count='1'><x:bk><x:rc t='1' v='0'/></x:bk></x:cellMetadata>
        \\</x:metadata>
    ;
    var md = try parseOk(testing.allocator, prefixed);
    defer md.deinit(testing.allocator);
    try testing.expectEqual(value.Dialect.dynamic_array, md.resolveCell(.input, .{ .cm = 1 }).ok);
}

test "parse: limits refuse instead of allocating without bound" {
    const many_types =
        "<metadata><metadataTypes>" ++
        ("<metadataType name=\"X\"/>" ** 8) ++
        "</metadataTypes></metadata>";
    var p = try parse(testing.allocator, many_types, .{ .limits = .{ .max_types = 4 } });
    defer p.deinit(testing.allocator);
    try testing.expectEqual(Refusal.Reason.limit_exceeded, p.refused.reason);
    try testing.expectEqual(LimitKind.types, p.refused.limit.?);

    // Offsets are u32; a part too large to address by one is refused
    // before it is scanned, not truncated into a misleading offset.
    var oversize = try parse(testing.allocator, fixture_dynamic_array, .{ .limits = .{ .max_bytes = 4 } });
    defer oversize.deinit(testing.allocator);
    try testing.expectEqual(LimitKind.part_bytes, oversize.refused.limit.?);

    const many_blocks =
        "<metadata><cellMetadata>" ++
        ("<bk><rc t=\"1\"/></bk>" ** 8) ++
        "</cellMetadata></metadata>";
    var blocks = try parse(testing.allocator, many_blocks, .{ .limits = .{ .max_blocks = 3 } });
    defer blocks.deinit(testing.allocator);
    try testing.expectEqual(LimitKind.blocks, blocks.refused.limit.?);

    const many_records =
        "<metadata><cellMetadata><bk>" ++
        ("<rc t=\"1\"/>" ** 8) ++
        "</bk></cellMetadata></metadata>";
    var records = try parse(testing.allocator, many_records, .{ .limits = .{ .max_records = 3 } });
    defer records.deinit(testing.allocator);
    try testing.expectEqual(LimitKind.records, records.refused.limit.?);

    const deep_foreign =
        "<metadata><extLst><ext uri=\"x\">" ++
        ("<a>" ** 40) ++ ("</a>" ** 40) ++
        "</ext></extLst></metadata>";
    var q = try parse(testing.allocator, deep_foreign, .{ .limits = .{ .max_skip_depth = 8 } });
    defer q.deinit(testing.allocator);
    try testing.expectEqual(LimitKind.skip_depth, q.refused.limit.?);
}

test "checkAllAllocationFailures: parsing leaks nothing under OOM" {
    const H = struct {
        fn run(allocator: std.mem.Allocator) !void {
            for ([_][]const u8{
                fixture_dynamic_array,
                fixture_rich_value,
                fixture_all_elements,
                "<metadata/>",
            }) |xml| {
                var p = try parse(allocator, xml, .{});
                defer p.deinit(allocator);
                const md = switch (p) {
                    .ok => |m| m,
                    .refused => return error.TestUnexpectedRefusal,
                };
                _ = md.resolveCell(.input, .{ .cm = 1 });
                _ = md.resolveCell(.result, .{ .vm = 1 });
            }
        }
    };
    try testing.checkAllAllocationFailures(testing.allocator, H.run, .{});
}

// ─── fuzz (§8.1: metadata) ───────────────────────────────────────

fn fuzzMetadataTarget(_: void, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();
    var buf: [1024]u8 = undefined;
    const input = buf[0..smith.slice(&buf)];

    var p = try parse(std.testing.allocator, input, .{});
    defer p.deinit(std.testing.allocator);
    const md = switch (p) {
        .refused => |r| {
            // A refusal is always classified, and always inside the
            // taxonomy.
            _ = r.planeTwo();
            return;
        },
        .ok => |m| m,
    };

    // Resolution is total: every index a cell could carry either answers
    // or refuses, for both roles, without reading out of bounds.
    var i: u32 = 0;
    while (i <= md.cell.blocks.len + 2) : (i += 1) {
        for ([_]CellRole{ .input, .result }) |role| {
            const first = md.resolveCell(role, .{ .cm = i });
            const second = md.resolveCell(role, .{ .cm = i });
            try std.testing.expectEqual(
                std.meta.activeTag(first),
                std.meta.activeTag(second),
            );
        }
    }
    i = 0;
    while (i <= md.value.blocks.len + 2) : (i += 1) {
        for ([_]CellRole{ .input, .result }) |role| {
            _ = md.resolveCell(role, .{ .vm = i });
        }
    }

    // All-or-nothing: a refused run writes nothing at all.
    const cells = [_]RunCell{
        .{ .role = .input, .meta = .{} },
        .{ .role = .input, .meta = .{ .cm = 1 } },
        .{ .role = .result, .meta = .{ .cm = 2 } },
        .{ .role = .result, .meta = .{ .vm = 1 } },
    };
    var out: [cells.len]value.Dialect = @splat(.dynamic_array);
    switch (md.resolveRun(&cells, &out)) {
        .refused => for (out) |d| try std.testing.expectEqual(value.Dialect.dynamic_array, d),
        .ok => {},
    }
}

test "fuzz: no metadata part can panic, leak, or resolve halfway" {
    try std.testing.fuzz({}, fuzzMetadataTarget, .{
        .corpus = &[_][]const u8{
            fixture_dynamic_array,
            fixture_rich_value,
            fixture_all_elements,
            "<metadata/>",
            "<metadata><cellMetadata><bk><rc t=\"1\" v=\"0\"/></bk></cellMetadata></metadata>",
            "<metadata><valueMetadata><bk><rc t=\"1\" v=\"0\"/></bk></valueMetadata></metadata>",
            "<metadata><metadataTypes><metadataType name=\"XLDAPR\"/></metadataTypes></metadata>",
            "<!DOCTYPE m><metadata/>",
            "<metadata><extLst><ext uri=\"x\"><a><b/></a></ext></extLst></metadata>",
            "<metadata><bk/></metadata>",
            "<metadata",
            "",
            "\xFF\xFE",
        },
    });
}
