//! emb-4B carrier catalogue — shared between the fixture generator and
//! the verifier so the two can never disagree about what was planted
//! where.
//!
//! **What emb-4B measures, and why it is not emb-4.** emb-4 asked
//! "does the `xl/zlsxEmbeddings/*` custom OPC part survive a passive
//! save?" and got a bad answer: Excel preserves it, but Apple Numbers
//! and LibreOffice Calc both rebuild the archive from their own model
//! and drop every unknown part. emb-4B asks the follow-up question —
//! *is there anywhere else in the package that does survive those
//! rebuilds?*
//!
//! The two failure modes deliberately do not overlap:
//!
//!   - The custom OPC part is invisible to Excel's Document Inspector
//!     but is erased by archive-rebuilding consumers.
//!   - Cell data is preserved by every consumer (it *is* the workbook
//!     model) but is enumerated by the Inspector and by any user who
//!     clicks Sheet ▸ Unhide.
//!
//! That is the argument for carrying a small **recovery record** in a
//! second carrier rather than picking one hiding spot and hoping. A
//! recovery record does not hold vectors — it holds provenance (model
//! id, dim, dtype, coverage ranges, content hash) so that a consumer
//! which finds the vectors gone can tell *that* they were stripped,
//! and re-embed from source rather than silently returning nothing.
//! Roughly 100–200 bytes, which is what makes the capacity-limited
//! carriers (custom document properties, defined names) viable here
//! even though they could never hold the vectors themselves.
//!
//! Every carrier below is planted in ONE fixture, so a single
//! open → save round-trip through a tool measures all six at once.
//! That matters most for the GUI-only legs (Excel, Numbers), where the
//! alternative is six manual saves per tool.

const std = @import("std");

/// Fixed nonce, not random: the fixture must be byte-stable across
/// runs so that a diff after a matrix leg is attributable to the
/// consumer tool and never to the generator. Same property emb-4's
/// fixture holds.
pub const NONCE: []const u8 = "8f3a2c1d";

pub const Carrier = enum {
    /// `xl/zlsxEmbeddings/index.xml` — the emb-1a/3a custom OPC part.
    /// Present as the control: its verdict here must reproduce emb-4,
    /// otherwise the fixture or the tool changed and the whole run is
    /// suspect.
    opc_part,
    /// `customXml/item1.xml` (+ itemProps + rels + workbook rel) — the
    /// standard OOXML custom-XML extension point. Rejected as the
    /// primary carrier in the design doc because Document Inspector ▸
    /// Custom XML Data ▸ Remove All targets it specifically; still
    /// worth measuring as a *secondary* carrier.
    custom_xml,
    /// `docProps/custom.xml` — a custom document property. Scalar-only
    /// and length-capped (~255 chars in Excel), so it can hold a
    /// recovery record but never the vectors.
    doc_props,
    /// A worksheet cell on a hidden sheet. Expected to survive
    /// everything (it is core workbook content); measured to confirm,
    /// and to record whether the *hidden* flag survives alongside the
    /// data.
    cell_data,
    /// A workbook-scoped `<definedName>` whose formula is a string
    /// literal. Capacity-limited like doc_props, but not enumerated by
    /// any Document Inspector module.
    defined_name,
    /// `<extLst><ext uri="…">` on `xl/workbook.xml` — the extension
    /// point ECMA-376 actually sanctions for vendor data, and the one
    /// Excel is contractually expected to round-trip.
    ext_lst,

    pub fn slug(self: Carrier) []const u8 {
        return switch (self) {
            .opc_part => "OPCPART",
            .custom_xml => "CUSTOMXML",
            .doc_props => "DOCPROPS",
            .cell_data => "CELLDATA",
            .defined_name => "DEFNAME",
            .ext_lst => "EXTLST",
        };
    }

    /// Human-readable location, for the verifier's report and for the
    /// matrix table in docs/plans/emb-4b-carrier-matrix.md.
    pub fn location(self: Carrier) []const u8 {
        return switch (self) {
            .opc_part => "xl/zlsxEmbeddings/index.xml",
            .custom_xml => "customXml/item1.xml",
            .doc_props => "docProps/custom.xml",
            .cell_data => "hidden sheet 'zlsxE4B' cell A1",
            .defined_name => "xl/workbook.xml <definedName ZlsxE4BRecovery>",
            .ext_lst => "xl/workbook.xml <extLst><ext>",
        };
    }
};

pub const ALL = [_]Carrier{
    .opc_part,
    .custom_xml,
    .doc_props,
    .cell_data,
    .defined_name,
    .ext_lst,
};

/// `ZLSX-E4B-<SLUG>-<NONCE>`. Distinct per carrier so a tool that
/// relocates or merges content cannot make one carrier's survival look
/// like another's.
pub fn marker(buf: []u8, c: Carrier) []const u8 {
    return std.fmt.bufPrint(buf, "ZLSX-E4B-{s}-{s}", .{ c.slug(), NONCE }) catch unreachable;
}

/// Longest possible marker, for caller stack buffers.
pub const MARKER_MAX: usize = "ZLSX-E4B-".len + "CUSTOMXML".len + 1 + 8;

/// Sheet that carries the `cell_data` marker.
pub const CELL_SHEET: []const u8 = "zlsxE4B";
/// Defined name that carries the `defined_name` marker.
pub const DEFINED_NAME: []const u8 = "ZlsxE4BRecovery";
/// `<ext uri>` GUID for the `ext_lst` carrier. Must be well-formed hex
/// — Excel validates the shape and a malformed URI would make this a
/// measurement of our typo rather than of the carrier. Vendor-owned,
/// so no collision with an Office-defined extension URI.
pub const EXT_URI: []const u8 = "{7A1C4E90-6B2D-4F31-9E5A-2C1D8F3A0001}";

test "markers are distinct and stable" {
    var seen: [ALL.len][MARKER_MAX]u8 = undefined;
    var lens: [ALL.len]usize = undefined;
    for (ALL, 0..) |c, i| {
        const m = marker(&seen[i], c);
        lens[i] = m.len;
    }
    for (0..ALL.len) |i| {
        for (i + 1..ALL.len) |j| {
            try std.testing.expect(!std.mem.eql(u8, seen[i][0..lens[i]], seen[j][0..lens[j]]));
        }
    }
    var buf: [MARKER_MAX]u8 = undefined;
    try std.testing.expectEqualStrings("ZLSX-E4B-OPCPART-8f3a2c1d", marker(&buf, .opc_part));
}
