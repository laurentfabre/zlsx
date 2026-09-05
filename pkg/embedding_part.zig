//! Embedding-part wire-format primitives (emb-1).
//!
//! Pure functions over byte slices. Allocators appear only for caller-
//! provided output buffers or temporary NFC normalization.
//!
//! See `docs/plans/embeddings-in-xlsx.md` for the format spec. This
//! file implements:
//!
//! - `VecHeader` (24 bytes) + `HashHeader` (16 bytes) byte layout
//!   with comptime size asserts pinning the disk offset of record 0.
//! - `Dtype` enum mapping the wire byte + the XML attribute string.
//! - `parseIndex` — low-allocation parser for
//!   `xl/zlsxEmbeddings/index.xml`, writing coverage records into a
//!   caller-provided slice.
//! - `parseIndexRelationships` + `validateIndexRelationships` —
//!   low-allocation validation for `index.xml.rels` before emb-2
//!   materializes vec/hash parts through PartStore.
//! - `parseVecPart` / `parseHashPart` — exact-size binary views with
//!   per-record helpers.
//! - f32 → binary16 / bfloat16 conversion (the compiler's IEEE 754
//!   lowering plus a vetted bf16 round-to-nearest-even path).
//! - int8 symmetric + asymmetric per-vector quantization +
//!   dequantization (used by Codec round-trip tests in this file
//!   and by the recall benchmark in emb-4).
//! - `parseVecHeader` / `parseHashHeader` — magic + version +
//!   reserved validation, returns the parsed struct or a typed
//!   error.
//! - `canonicalizeNumber` — Ryu shortest-round-trip via
//!   `std.fmt.float.render`, with the v1 special-case rules from
//!   the spec (`0.0` and `-0.0` both → `0`; NaN/inf rejected).
//! - `canonicalizeText` and `xxh3Canonical` — trim Unicode whitespace,
//!   NFC-normalize text, then hash the canonical row payload shape.
//! - `xxh3` thin wrapper over `std.hash.XxHash3` so callers don't
//!   pin a specific stdlib symbol.
//!
//! Tombstone marker (load-bearing): `TOMBSTONE_HASH` = `u64::MAX`.
//! Per the spec, query-time consumers MUST skip slots whose stored
//! hash equals this value (zero-vector slots cause NaN under
//! cosine).

const std = @import("std");
const Allocator = std.mem.Allocator;
const nfc = @import("zlsx_nfc");
const coords = @import("zlsx_refs");
// The writer's XML byte rule, so embedding metadata refuses exactly the
// bytes every other text channel refuses (one definition).
const sheet_plan = @import("zlsx_sheet_plan");

pub const Error = error{
    BadMagic,
    UnsupportedVersion,
    InvalidDtype,
    InvalidReservedBytes,
    HeaderTooShort,
    BodyTooShort,
    BodySizeMismatch,
    CountMismatch,
    DimensionMismatch,
    DtypeMismatch,
    DimensionOutOfRange,
    MalformedNumber,
    MalformedIndexXml,
    MalformedRelationshipsXml,
    MissingAttribute,
    MissingRelationship,
    UnsupportedHashAlgorithm,
    InvalidBoolean,
    InvalidCoverageId,
    DuplicateCoverageId,
    DuplicateRelationshipId,
    DuplicateRelationshipTarget,
    InvalidIndexRelationship,
    InvalidRange,
    CoverageCountMismatch,
    CoverageOverlap,
    /// A dtype with a decoder but no writer — binary16 / bfloat16 /
    /// the asymmetric int8 layout. Surfaces from `encodeVectorRecord`.
    UnsupportedDtype,
    /// A C0 control byte XML 1.0 forbids in a metadata string (`model`,
    /// `worksheet_target`) — the rule every other text channel of the
    /// writer enforces (`sheet_plan.isForbiddenXmlByte`).
    InvalidXmlByte,
    InvalidUtf8,
    UnicodeNormalizationFailed,
    BufferTooSmall,
    OutOfMemory,
};

// ---------------------------------------------------------------
// Constants
// ---------------------------------------------------------------

/// vec.bin magic. On disk: bytes 'Z','V','E','C'. As a little-endian
/// u32 that's 0x4345565A.
pub const VEC_MAGIC: u32 = 0x4345565A;

/// hashes.bin magic. On disk: bytes 'Z','H','S','H'. As a
/// little-endian u32 that's 0x4853485A.
pub const HASH_MAGIC: u32 = 0x4853485A;

pub const WIRE_VERSION: u32 = 1;
pub const EMBEDDINGS_DIR: []const u8 = "xl/zlsxEmbeddings";
pub const INDEX_PART_NAME: []const u8 = "xl/zlsxEmbeddings/index.xml";
pub const INDEX_RELS_PART_NAME: []const u8 = "xl/zlsxEmbeddings/_rels/index.xml.rels";
pub const INDEX_NAMESPACE: []const u8 = "http://schemas.fabre.me/zlsx/2026/embeddings";
pub const HASH_ALGO_XXH3_64: []const u8 = "xxh3-64";
pub const INDEX_CONTENT_TYPE: []const u8 = "application/vnd.fabre.zlsx.embedding-index+xml";
pub const VEC_CONTENT_TYPE: []const u8 = "application/vnd.fabre.zlsx.embedding-vec";
pub const HASH_CONTENT_TYPE: []const u8 = "application/vnd.fabre.zlsx.embedding-hash";
pub const REL_TYPE_EMBEDDINGS: []const u8 = "http://schemas.fabre.me/zlsx/2026/relationships/embeddings";
pub const REL_TYPE_VEC: []const u8 = "http://schemas.fabre.me/zlsx/2026/relationships/embedding-vec";
pub const REL_TYPE_HASH: []const u8 = "http://schemas.fabre.me/zlsx/2026/relationships/embedding-hash";

/// The sole no-vector / tombstone marker. Slot whose stored hash
/// equals this value MUST be skipped by query consumers (the vec
/// record at the same slot index is zeroed but its float
/// interpretation is undefined behavior for similarity scoring).
pub const TOMBSTONE_HASH: u64 = std.math.maxInt(u64);

/// Hash seed for `std.hash.XxHash3`. Pinned in v1 so that
/// independent reimplementations of the format produce the same
/// 64-bit value over byte-identical canonical input.
pub const HASH_SEED: u64 = 0;

// ---------------------------------------------------------------
// Dtype
// ---------------------------------------------------------------

pub const Dtype = enum(u8) {
    f32 = 0,
    binary16 = 1,
    bfloat16 = 2,
    int8_sym_per_vec = 3,
    int8_asym_per_vec = 4,

    pub fn fromU8(b: u8) Error!Dtype {
        return switch (b) {
            0 => .f32,
            1 => .binary16,
            2 => .bfloat16,
            3 => .int8_sym_per_vec,
            4 => .int8_asym_per_vec,
            else => Error.InvalidDtype,
        };
    }

    pub fn fromString(s: []const u8) Error!Dtype {
        if (std.mem.eql(u8, s, "f32")) return .f32;
        if (std.mem.eql(u8, s, "binary16")) return .binary16;
        if (std.mem.eql(u8, s, "bfloat16")) return .bfloat16;
        if (std.mem.eql(u8, s, "int8-sym-per-vec")) return .int8_sym_per_vec;
        if (std.mem.eql(u8, s, "int8-asym-per-vec")) return .int8_asym_per_vec;
        return Error.InvalidDtype;
    }

    pub fn string(self: Dtype) []const u8 {
        return switch (self) {
            .f32 => "f32",
            .binary16 => "binary16",
            .bfloat16 => "bfloat16",
            .int8_sym_per_vec => "int8-sym-per-vec",
            .int8_asym_per_vec => "int8-asym-per-vec",
        };
    }

    /// Wire-bytes per vector record for a given dimension. The
    /// header bytes are NOT included.
    pub fn recordBytes(self: Dtype, dim: u32) usize {
        const d: usize = dim;
        return switch (self) {
            .f32 => d * 4,
            .binary16, .bfloat16 => d * 2,
            .int8_sym_per_vec => 4 + d, // f32 scale + i8[dim]
            .int8_asym_per_vec => 4 + 1 + d, // f32 scale + i8 zero + i8[dim]
        };
    }
};

// ---------------------------------------------------------------
// index.xml parser
// ---------------------------------------------------------------

pub const EXCEL_MAX_ROW: u32 = coords.max_row;
pub const EXCEL_MAX_COL: u32 = coords.max_col_1based;

pub const A1Corner = struct {
    row: u32, // 1-based
    col: u32, // 0-based
};

pub const A1Range = struct {
    first: A1Corner,
    last: A1Corner,

    pub fn rowCount(self: A1Range) u32 {
        std.debug.assert(self.last.row >= self.first.row);
        return self.last.row - self.first.row + 1;
    }

    pub fn overlaps(self: A1Range, other: A1Range) bool {
        return self.first.row <= other.last.row and other.first.row <= self.last.row and
            self.first.col <= other.last.col and other.first.col <= self.last.col;
    }
};

pub const Coverage = struct {
    id: []const u8,
    worksheet_target: []const u8,
    range: []const u8,
    parsed_range: A1Range,
    column: []const u8,
    column_idx: u32,
    count: u32,
    include_formulas: bool,
    vec_rid: []const u8,
    hash_rid: []const u8,
};

pub const Index = struct {
    version: u32,
    model: []const u8,
    dim: u32,
    dtype: Dtype,
    hash_algo: []const u8,
    coverages: []const Coverage,
};

pub const TargetMode = enum {
    internal,
    external,
};

pub const IndexRelationship = struct {
    id: []const u8,
    type: []const u8,
    target: []const u8,
    target_mode: TargetMode,
};

pub fn parseIndex(xml: []const u8, coverage_buf: []Coverage) Error!Index {
    var pos: usize = 0;
    const root = (try nextXmlTag(xml, &pos)) orelse return Error.MalformedIndexXml;
    if (root.closing) return Error.MalformedIndexXml;
    if (!std.mem.eql(u8, localName(root.name), "embeddings")) return Error.MalformedIndexXml;

    const xmlns = try attrValueRequired(root.body, "xmlns");
    if (!std.mem.eql(u8, xmlns, INDEX_NAMESPACE)) return Error.MalformedIndexXml;

    const version = try parseU32Decimal(try attrValueRequired(root.body, "version"));
    if (version != WIRE_VERSION) return Error.UnsupportedVersion;
    const model = try attrValueRequired(root.body, "model");
    const dim = try parseU32Decimal(try attrValueRequired(root.body, "dim"));
    if (dim == 0) return Error.DimensionOutOfRange;
    const dtype = try Dtype.fromString(try attrValueRequired(root.body, "dtype"));
    const hash_algo = try attrValueRequired(root.body, "hash_algo");
    if (!std.mem.eql(u8, hash_algo, HASH_ALGO_XXH3_64)) return Error.UnsupportedHashAlgorithm;

    var coverage_count: usize = 0;
    var saw_root_close = root.self_closing;
    while (!saw_root_close) {
        const tag = (try nextXmlTag(xml, &pos)) orelse return Error.MalformedIndexXml;
        const lname = localName(tag.name);
        if (tag.closing) {
            if (std.mem.eql(u8, lname, "embeddings")) {
                saw_root_close = true;
                break;
            }
            continue;
        }
        if (!std.mem.eql(u8, lname, "coverage")) continue;
        if (coverage_count == coverage_buf.len) return Error.BufferTooSmall;

        const coverage = try parseCoverage(tag.body);
        for (coverage_buf[0..coverage_count]) |prev| {
            if (std.mem.eql(u8, prev.id, coverage.id)) return Error.DuplicateCoverageId;
            if (std.mem.eql(u8, prev.vec_rid, coverage.vec_rid) or
                std.mem.eql(u8, prev.vec_rid, coverage.hash_rid) or
                std.mem.eql(u8, prev.hash_rid, coverage.vec_rid) or
                std.mem.eql(u8, prev.hash_rid, coverage.hash_rid))
            {
                return Error.DuplicateRelationshipId;
            }
            if (std.mem.eql(u8, prev.worksheet_target, coverage.worksheet_target) and
                prev.parsed_range.overlaps(coverage.parsed_range))
            {
                return Error.CoverageOverlap;
            }
        }
        coverage_buf[coverage_count] = coverage;
        coverage_count += 1;
    }
    if (coverage_count == 0) return Error.MalformedIndexXml;
    return .{
        .version = version,
        .model = model,
        .dim = dim,
        .dtype = dtype,
        .hash_algo = hash_algo,
        .coverages = coverage_buf[0..coverage_count],
    };
}

pub fn parseIndexRelationships(xml: []const u8, rel_buf: []IndexRelationship) Error![]IndexRelationship {
    var pos: usize = 0;
    const root = (nextXmlTag(xml, &pos) catch return Error.MalformedRelationshipsXml) orelse
        return Error.MalformedRelationshipsXml;
    if (root.closing) return Error.MalformedRelationshipsXml;
    if (!std.mem.eql(u8, localName(root.name), "Relationships")) return Error.MalformedRelationshipsXml;

    var rel_count: usize = 0;
    var saw_root_close = root.self_closing;
    while (!saw_root_close) {
        const tag = (nextXmlTag(xml, &pos) catch return Error.MalformedRelationshipsXml) orelse
            return Error.MalformedRelationshipsXml;
        const lname = localName(tag.name);
        if (tag.closing) {
            if (std.mem.eql(u8, lname, "Relationships")) {
                saw_root_close = true;
                break;
            }
            continue;
        }
        if (!std.mem.eql(u8, lname, "Relationship")) continue;
        if (rel_count == rel_buf.len) return Error.BufferTooSmall;

        const rel = try parseIndexRelationship(tag.body);
        for (rel_buf[0..rel_count]) |prev| {
            if (std.mem.eql(u8, prev.id, rel.id)) return Error.DuplicateRelationshipId;
        }
        rel_buf[rel_count] = rel;
        rel_count += 1;
    }
    return rel_buf[0..rel_count];
}

/// Validate that every coverage's `vec_rId` and `hash_rId` resolves
/// to a non-external relationship with the v1 relationship type and
/// deterministic per-coverage target (`<id>/vec.bin` or
/// `<id>/hashes.bin`). This is intentionally independent of
/// `PartStore`; emb-2 will use PartStore only to normalize and
/// materialize the already-validated target.
pub fn validateIndexRelationships(index: Index, rels: []const IndexRelationship) Error!void {
    for (index.coverages) |coverage| {
        const vec_rel = try coverageRelationship(coverage, rels, .vec);
        const hash_rel = try coverageRelationship(coverage, rels, .hash);
        if (std.mem.eql(u8, vec_rel.target, hash_rel.target)) return Error.DuplicateRelationshipTarget;

        var vec_expected_buf: [80]u8 = undefined;
        const vec_expected = try expectedVecTarget(coverage.id, &vec_expected_buf);
        if (!std.mem.eql(u8, vec_rel.target, vec_expected)) return Error.InvalidIndexRelationship;

        var hash_expected_buf: [80]u8 = undefined;
        const hash_expected = try expectedHashTarget(coverage.id, &hash_expected_buf);
        if (!std.mem.eql(u8, hash_rel.target, hash_expected)) return Error.InvalidIndexRelationship;
    }

    for (index.coverages, 0..) |lhs, lhs_i| {
        const lhs_vec = try coverageRelationship(lhs, rels, .vec);
        const lhs_hash = try coverageRelationship(lhs, rels, .hash);
        for (index.coverages[0..lhs_i]) |rhs| {
            const rhs_vec = try coverageRelationship(rhs, rels, .vec);
            const rhs_hash = try coverageRelationship(rhs, rels, .hash);
            if (std.mem.eql(u8, lhs_vec.target, rhs_vec.target) or
                std.mem.eql(u8, lhs_vec.target, rhs_hash.target) or
                std.mem.eql(u8, lhs_hash.target, rhs_vec.target) or
                std.mem.eql(u8, lhs_hash.target, rhs_hash.target))
            {
                return Error.DuplicateRelationshipTarget;
            }
        }
    }
}

pub fn expectedVecTarget(id: []const u8, out: []u8) Error![]const u8 {
    return expectedCoverageTarget(id, "vec.bin", out);
}

pub fn expectedHashTarget(id: []const u8, out: []u8) Error![]const u8 {
    return expectedCoverageTarget(id, "hashes.bin", out);
}

fn parseCoverage(tag_body: []const u8) Error!Coverage {
    const id = try attrValueRequired(tag_body, "id");
    try validateCoverageId(id);
    const worksheet_target = try attrValueRequired(tag_body, "worksheet_target");
    const range_raw = try attrValueRequired(tag_body, "range");
    const parsed_range = try parseA1Range(range_raw);
    const column = try attrValueRequired(tag_body, "column");
    const column_idx = try validateCoverageColumn(column, parsed_range);
    const count = try parseU32Decimal(try attrValueRequired(tag_body, "count"));
    if (count != parsed_range.rowCount()) return Error.CoverageCountMismatch;
    const include_formulas = if (try attrValue(tag_body, "include_formulas")) |raw|
        try parseBool(raw)
    else
        false;
    const vec_rid = try attrValueRequired(tag_body, "vec_rId");
    const hash_rid = try attrValueRequired(tag_body, "hash_rId");
    if (std.mem.eql(u8, vec_rid, hash_rid)) return Error.DuplicateRelationshipId;
    return .{
        .id = id,
        .worksheet_target = worksheet_target,
        .range = range_raw,
        .parsed_range = parsed_range,
        .column = column,
        .column_idx = column_idx,
        .count = count,
        .include_formulas = include_formulas,
        .vec_rid = vec_rid,
        .hash_rid = hash_rid,
    };
}

fn parseIndexRelationship(tag_body: []const u8) Error!IndexRelationship {
    const id = try attrValueRequired(tag_body, "Id");
    const rtype = try attrValueRequired(tag_body, "Type");
    const target = try attrValueRequired(tag_body, "Target");
    const target_mode: TargetMode = if (try attrValue(tag_body, "TargetMode")) |mode| blk: {
        if (std.mem.eql(u8, mode, "External")) break :blk .external;
        if (std.mem.eql(u8, mode, "Internal")) break :blk .internal;
        return Error.InvalidIndexRelationship;
    } else .internal;
    return .{
        .id = id,
        .type = rtype,
        .target = target,
        .target_mode = target_mode,
    };
}

pub const CoverageRelKind = enum {
    vec,
    hash,
};

pub fn coverageRelationship(
    coverage: Coverage,
    rels: []const IndexRelationship,
    kind: CoverageRelKind,
) Error!IndexRelationship {
    const rid = switch (kind) {
        .vec => coverage.vec_rid,
        .hash => coverage.hash_rid,
    };
    const want_type = switch (kind) {
        .vec => REL_TYPE_VEC,
        .hash => REL_TYPE_HASH,
    };
    for (rels) |rel| {
        if (!std.mem.eql(u8, rel.id, rid)) continue;
        if (rel.target_mode != .internal) return Error.InvalidIndexRelationship;
        if (!std.mem.eql(u8, rel.type, want_type)) return Error.InvalidIndexRelationship;
        return rel;
    }
    return Error.MissingRelationship;
}

fn expectedCoverageTarget(id: []const u8, leaf: []const u8, out: []u8) Error![]const u8 {
    try validateCoverageId(id);
    return std.fmt.bufPrint(out, "{s}/{s}", .{ id, leaf }) catch Error.BufferTooSmall;
}

fn validateCoverageId(id: []const u8) Error!void {
    if (id.len == 0 or id.len > 63) return Error.InvalidCoverageId;
    for (id) |b| {
        const ok = (b >= 'A' and b <= 'Z') or
            (b >= 'a' and b <= 'z') or
            (b >= '0' and b <= '9') or
            b == '_' or b == '-';
        if (!ok) return Error.InvalidCoverageId;
    }
}

pub fn parseA1Range(s: []const u8) Error!A1Range {
    const colon = std.mem.indexOfScalar(u8, s, ':');
    const first = try parseA1Corner(if (colon) |c| s[0..c] else s);
    const last = try parseA1Corner(if (colon) |c| s[c + 1 ..] else s);
    if (first.row > last.row) return Error.InvalidRange;
    if (first.col > last.col) return Error.InvalidRange;
    return .{ .first = first, .last = last };
}

pub fn parseA1Corner(s: []const u8) Error!A1Corner {
    // M0 adapter over `zlsx_refs`. Policy preserved exactly:
    // coverage ranges are author-supplied, so letters are
    // case-insensitive and a leading-zero row is rejected.
    const cell = coords.parseCell(s, .{
        .case = .insensitive,
        .leading_zero_row = .reject,
    }) catch return Error.InvalidRange;
    return .{ .row = cell.row.oneBased(), .col = cell.col.zeroBased() };
}

/// Parse a bare column name ("A", "aa") to a **0-based** index.
pub fn parseColumnName(s: []const u8) Error!u32 {
    const col = coords.parseCol(s, .{ .case = .insensitive }) catch return Error.InvalidRange;
    return col.zeroBased();
}

/// The coverage-column rule, ONE for the index read (`parseCoverage`)
/// and the write (`Workbook.setEmbeddings`): a bare column name whose
/// index lies inside the coverage's range. Returns the 0-based index.
/// The writer used to skip it, so a coverage with `column = "Z"` on
/// `A2:A4` saved cleanly and `Workbook.embeddings()` then refused the
/// index it had just written (`InvalidRange`).
pub fn validateCoverageColumn(column: []const u8, range: A1Range) Error!u32 {
    const column_idx = try parseColumnName(column);
    if (column_idx < range.first.col or column_idx > range.last.col) return Error.InvalidRange;
    return column_idx;
}

/// The metadata byte rule: a `model` / `worksheet_target` string is an
/// XML attribute value, and XML 1.0 forbids most C0 control bytes in
/// one (0x00–0x08, 0x0B, 0x0C, 0x0E–0x1F; tab / LF / CR and DEL are
/// legal). The cell-text, sheet-name, comment, defined-name and
/// hyperlink writers refuse the same bytes through
/// `sheet_plan.isForbiddenXmlByte`; this is that rule on the one text
/// channel that escaped it (S0's surface audit). Checked by the writer
/// BEFORE its first part write and again by `appendXmlEscaped` when
/// the index is encoded.
pub fn validateMetadataText(s: []const u8) Error!void {
    for (s) |c| if (sheet_plan.isForbiddenXmlByte(c)) return Error.InvalidXmlByte;
}

const XmlTag = struct {
    name: []const u8,
    body: []const u8,
    closing: bool,
    self_closing: bool,
};

fn nextXmlTag(xml: []const u8, pos: *usize) Error!?XmlTag {
    while (pos.* < xml.len) {
        const lt = std.mem.indexOfScalarPos(u8, xml, pos.*, '<') orelse {
            pos.* = xml.len;
            return null;
        };
        if (std.mem.startsWith(u8, xml[lt..], "<!--")) {
            const end = std.mem.indexOfPos(u8, xml, lt + 4, "-->") orelse return Error.MalformedIndexXml;
            pos.* = end + 3;
            continue;
        }
        const gt = try findXmlTagEnd(xml, lt + 1);
        pos.* = gt + 1;
        var body = std.mem.trim(u8, xml[lt + 1 .. gt], " \t\r\n");
        if (body.len == 0) continue;
        if (body[0] == '?' or body[0] == '!') continue;

        const closing = body[0] == '/';
        if (closing) body = std.mem.trim(u8, body[1..], " \t\r\n");

        var self_closing = false;
        if (!closing and body.len > 0 and body[body.len - 1] == '/') {
            self_closing = true;
            body = std.mem.trim(u8, body[0 .. body.len - 1], " \t\r\n");
        }
        if (body.len == 0) return Error.MalformedIndexXml;
        const name_end = scanNameEnd(body, 0);
        return .{
            .name = body[0..name_end],
            .body = body,
            .closing = closing,
            .self_closing = self_closing,
        };
    }
    return null;
}

fn findXmlTagEnd(xml: []const u8, start: usize) Error!usize {
    var quote: u8 = 0;
    var i = start;
    while (i < xml.len) : (i += 1) {
        const b = xml[i];
        if (quote != 0) {
            if (b == quote) quote = 0;
            continue;
        }
        if (b == '"' or b == '\'') {
            quote = b;
            continue;
        }
        if (b == '>') return i;
    }
    return Error.MalformedIndexXml;
}

fn attrValueRequired(tag_body: []const u8, name: []const u8) Error![]const u8 {
    return (try attrValue(tag_body, name)) orelse Error.MissingAttribute;
}

fn attrValue(tag_body: []const u8, name: []const u8) Error!?[]const u8 {
    var i = scanNameEnd(tag_body, 0);
    while (i < tag_body.len) {
        skipXmlSpace(tag_body, &i);
        if (i >= tag_body.len) return null;
        const key_start = i;
        i = scanNameEnd(tag_body, i);
        if (i == key_start) return Error.MalformedIndexXml;
        const key = tag_body[key_start..i];
        skipXmlSpace(tag_body, &i);
        if (i >= tag_body.len or tag_body[i] != '=') return Error.MalformedIndexXml;
        i += 1;
        skipXmlSpace(tag_body, &i);
        if (i >= tag_body.len or (tag_body[i] != '"' and tag_body[i] != '\'')) return Error.MalformedIndexXml;
        const quote = tag_body[i];
        i += 1;
        const value_start = i;
        while (i < tag_body.len and tag_body[i] != quote) : (i += 1) {}
        if (i >= tag_body.len) return Error.MalformedIndexXml;
        const value = tag_body[value_start..i];
        i += 1;
        if (std.mem.eql(u8, key, name)) return value;
    }
    return null;
}

fn scanNameEnd(s: []const u8, start: usize) usize {
    var i = start;
    while (i < s.len) : (i += 1) {
        const b = s[i];
        if (isXmlSpace(b) or b == '=' or b == '/') break;
    }
    return i;
}

fn localName(name: []const u8) []const u8 {
    if (std.mem.lastIndexOfScalar(u8, name, ':')) |i| return name[i + 1 ..];
    return name;
}

fn skipXmlSpace(s: []const u8, i: *usize) void {
    while (i.* < s.len and isXmlSpace(s[i.*])) i.* += 1;
}

fn isXmlSpace(b: u8) bool {
    return b == ' ' or b == '\t' or b == '\n' or b == '\r';
}

fn isAsciiAlpha(b: u8) bool {
    return (b >= 'A' and b <= 'Z') or (b >= 'a' and b <= 'z');
}

fn parseU32Decimal(s: []const u8) Error!u32 {
    if (s.len == 0) return Error.MalformedIndexXml;
    var v: u32 = 0;
    for (s) |b| {
        if (b < '0' or b > '9') return Error.MalformedIndexXml;
        v = std.math.mul(u32, v, 10) catch return Error.MalformedIndexXml;
        v = std.math.add(u32, v, @as(u32, b - '0')) catch return Error.MalformedIndexXml;
    }
    return v;
}

fn parseBool(s: []const u8) Error!bool {
    if (std.mem.eql(u8, s, "true") or std.mem.eql(u8, s, "1")) return true;
    if (std.mem.eql(u8, s, "false") or std.mem.eql(u8, s, "0")) return false;
    return Error.InvalidBoolean;
}

// ---------------------------------------------------------------
// Header layouts
// ---------------------------------------------------------------

/// On-disk layout of the first 24 bytes of `vec.bin`. Reads via
/// `parseVecHeader`; writers either lay out by hand or use this
/// struct + `std.mem.bytesAsValue`. All fields little-endian.
pub const VecHeader = extern struct {
    magic: u32, // offset 0
    version: u32, // offset 4
    dim: u32, // offset 8
    count: u32, // offset 12
    dtype_byte: u8, // offset 16
    reserved: [7]u8, // offset 17..24
};

comptime {
    if (@sizeOf(VecHeader) != 24) @compileError("VecHeader must be exactly 24 bytes");
    if (@offsetOf(VecHeader, "magic") != 0) @compileError("VecHeader magic at offset 0");
    if (@offsetOf(VecHeader, "version") != 4) @compileError("VecHeader version at offset 4");
    if (@offsetOf(VecHeader, "dim") != 8) @compileError("VecHeader dim at offset 8");
    if (@offsetOf(VecHeader, "count") != 12) @compileError("VecHeader count at offset 12");
    if (@offsetOf(VecHeader, "dtype_byte") != 16) @compileError("VecHeader dtype_byte at offset 16");
    if (@offsetOf(VecHeader, "reserved") != 17) @compileError("VecHeader reserved at offset 17");
}

pub const VEC_HEADER_BYTES: usize = @sizeOf(VecHeader);

/// On-disk layout of the first 16 bytes of `hashes.bin`. All
/// fields little-endian.
pub const HashHeader = extern struct {
    magic: u32, // offset 0
    version: u32, // offset 4
    count: u32, // offset 8
    reserved: u32, // offset 12
};

comptime {
    if (@sizeOf(HashHeader) != 16) @compileError("HashHeader must be exactly 16 bytes");
}

pub const HASH_HEADER_BYTES: usize = @sizeOf(HashHeader);

/// Parsed view of a vec.bin header. Returned by `parseVecHeader`.
pub const ParsedVecHeader = struct {
    version: u32,
    dim: u32,
    count: u32,
    dtype: Dtype,
};

pub const ParsedHashHeader = struct {
    version: u32,
    count: u32,
};

pub const ParsedVecPart = struct {
    header: ParsedVecHeader,
    body: []const u8,

    pub fn record(self: ParsedVecPart, idx: u32) Error![]const u8 {
        if (idx >= self.header.count) return Error.InvalidRange;
        const record_len = self.header.dtype.recordBytes(self.header.dim);
        const start = @as(usize, idx) * record_len;
        return self.body[start .. start + record_len];
    }
};

pub const ParsedHashPart = struct {
    header: ParsedHashHeader,
    values_bytes: []const u8,

    pub fn value(self: ParsedHashPart, idx: u32) Error!u64 {
        if (idx >= self.header.count) return Error.InvalidRange;
        const start = @as(usize, idx) * @sizeOf(u64);
        return std.mem.readInt(u64, self.values_bytes[start .. start + @sizeOf(u64)][0..8], .little);
    }
};

pub fn parseVecHeader(bytes: []const u8) Error!ParsedVecHeader {
    if (bytes.len < VEC_HEADER_BYTES) return Error.HeaderTooShort;
    const magic = std.mem.readInt(u32, bytes[0..4], .little);
    if (magic != VEC_MAGIC) return Error.BadMagic;
    const version = std.mem.readInt(u32, bytes[4..8], .little);
    if (version != WIRE_VERSION) return Error.UnsupportedVersion;
    const dim = std.mem.readInt(u32, bytes[8..12], .little);
    if (dim == 0) return Error.DimensionOutOfRange;
    const count = std.mem.readInt(u32, bytes[12..16], .little);
    const dtype = try Dtype.fromU8(bytes[16]);
    for (bytes[17..VEC_HEADER_BYTES]) |b| {
        if (b != 0) return Error.InvalidReservedBytes;
    }
    // Body-size check: header must be followed by exactly
    // `count * recordBytes(dim)` bytes.
    const want_len = try checkedPartLen(VEC_HEADER_BYTES, count, dtype.recordBytes(dim));
    if (bytes.len < want_len) return Error.BodyTooShort;
    if (bytes.len != want_len) return Error.BodySizeMismatch;
    return .{ .version = version, .dim = dim, .count = count, .dtype = dtype };
}

pub fn parseVecPart(bytes: []const u8) Error!ParsedVecPart {
    const header = try parseVecHeader(bytes);
    return .{
        .header = header,
        .body = bytes[VEC_HEADER_BYTES..],
    };
}

pub fn parseHashHeader(bytes: []const u8) Error!ParsedHashHeader {
    if (bytes.len < HASH_HEADER_BYTES) return Error.HeaderTooShort;
    const magic = std.mem.readInt(u32, bytes[0..4], .little);
    if (magic != HASH_MAGIC) return Error.BadMagic;
    const version = std.mem.readInt(u32, bytes[4..8], .little);
    if (version != WIRE_VERSION) return Error.UnsupportedVersion;
    const count = std.mem.readInt(u32, bytes[8..12], .little);
    const reserved = std.mem.readInt(u32, bytes[12..16], .little);
    if (reserved != 0) return Error.InvalidReservedBytes;
    const want_len = try checkedPartLen(HASH_HEADER_BYTES, count, @sizeOf(u64));
    if (bytes.len < want_len) return Error.BodyTooShort;
    if (bytes.len != want_len) return Error.BodySizeMismatch;
    return .{ .version = version, .count = count };
}

pub fn parseHashPart(bytes: []const u8) Error!ParsedHashPart {
    const header = try parseHashHeader(bytes);
    return .{
        .header = header,
        .values_bytes = bytes[HASH_HEADER_BYTES..],
    };
}

/// Cross-check between a parsed VecHeader and a parsed HashHeader
/// for the same coverage: counts MUST match.
pub fn checkPairConsistent(vec: ParsedVecHeader, hash: ParsedHashHeader) Error!void {
    if (vec.count != hash.count) return Error.CountMismatch;
}

pub fn checkCoverageBinary(index: Index, coverage: Coverage, vec: ParsedVecHeader, hash: ParsedHashHeader) Error!void {
    if (vec.dim != index.dim) return Error.DimensionMismatch;
    if (vec.dtype != index.dtype) return Error.DtypeMismatch;
    if (vec.count != coverage.count) return Error.CountMismatch;
    if (hash.count != coverage.count) return Error.CountMismatch;
}

fn checkedPartLen(header_len: usize, count: u32, record_len: usize) Error!usize {
    const body_len = std.math.mul(usize, @intCast(count), record_len) catch
        return Error.BodyTooShort;
    return std.math.add(usize, header_len, body_len) catch Error.BodyTooShort;
}

// ---------------------------------------------------------------
// Float conversion: f32 → binary16 / bfloat16
// ---------------------------------------------------------------

/// f32 → IEEE 754 binary16 via the compiler's float lowering. The
/// returned u16 is the raw bit pattern. Behavior on overflow: the
/// Zig stdlib saturates to +/- infinity per IEEE 754. NaN is
/// preserved as a quiet NaN.
pub fn f32ToBinary16(x: f32) u16 {
    const half: f16 = @floatCast(x);
    return @bitCast(half);
}

/// IEEE 754 binary16 → f32. Lossless (binary16's representable
/// range is a strict subset of f32).
pub fn binary16ToF32(bits: u16) f32 {
    const half: f16 = @bitCast(bits);
    return @floatCast(half);
}

/// f32 → bfloat16 (8-bit exponent, 7-bit mantissa). Round-to-
/// nearest-even on the discarded low 16 bits. NaN is canonicalized
/// to a quiet NaN (0x7FC0).
pub fn f32ToBfloat16(x: f32) u16 {
    const bits: u32 = @bitCast(x);
    // NaN: emit canonical qNaN. (Exponent all-ones AND mantissa nonzero
    // identifies NaN regardless of sign.)
    if ((bits & 0x7F800000) == 0x7F800000 and (bits & 0x007FFFFF) != 0) {
        return 0x7FC0;
    }
    // RNE: bias 0x7FFF, plus the bit that breaks ties to even.
    const lsb_of_kept: u32 = (bits >> 16) & 1;
    const rounded: u32 = bits +% (0x7FFF + lsb_of_kept);
    return @truncate(rounded >> 16);
}

/// bfloat16 → f32. Lossless (bfloat16 is the top 16 bits of f32
/// plus zeroed low half).
pub fn bfloat16ToF32(bits: u16) f32 {
    const expanded: u32 = @as(u32, bits) << 16;
    return @bitCast(expanded);
}

// ---------------------------------------------------------------
// int8 quantization (symmetric + asymmetric, per-vector)
// ---------------------------------------------------------------

pub const QuantizedSym = struct {
    scale: f32,
};

pub const QuantizedAsym = struct {
    scale: f32,
    zero: i8,
};

/// Symmetric int8 quantization (one f32 scale per vector). Output
/// values are in [-127, 127]; the value -128 is reserved (never
/// produced by the quantizer) so dequantization symmetry holds.
/// Returns the computed scale (`f32`). On an all-zero input vector
/// the scale is 0 and the output is all zeroes — dequantization of
/// the zero vector returns zero, which is the desired behavior.
pub fn quantizeF32ToI8Sym(src: []const f32, out: []i8) QuantizedSym {
    std.debug.assert(src.len == out.len);
    // Find max abs value. Treat NaN as 0 (the caller is expected
    // to have sanitized vectors before quantizing).
    var max_abs: f32 = 0;
    for (src) |x| {
        const ax = if (std.math.isNan(x)) 0 else @abs(x);
        if (ax > max_abs) max_abs = ax;
    }
    if (max_abs == 0) {
        @memset(out, 0);
        return .{ .scale = 0 };
    }
    const scale = max_abs; // float ≈ (i8 / 127) * scale
    const inv = 127.0 / scale;
    for (src, out) |x, *q| {
        const xv = if (std.math.isNan(x)) 0 else x;
        // Round half away from zero; clamp to [-127, 127].
        var qf = xv * inv;
        if (qf >= 0) qf += 0.5 else qf -= 0.5;
        var qi: i32 = @intFromFloat(qf);
        if (qi > 127) qi = 127;
        if (qi < -127) qi = -127;
        q.* = @intCast(qi);
    }
    return .{ .scale = scale };
}

/// Dequantize one symmetric-quantized vector. Pure inverse of
/// `quantizeF32ToI8Sym` modulo quantization noise.
pub fn dequantizeI8Sym(values: []const i8, scale: f32, out: []f32) void {
    std.debug.assert(values.len == out.len);
    if (scale == 0) {
        @memset(out, 0);
        return;
    }
    const inv: f32 = scale / 127.0;
    for (values, out) |q, *v| {
        v.* = @as(f32, @floatFromInt(q)) * inv;
    }
}

/// Asymmetric int8 quantization. Stores one f32 scale + one i8
/// zero-point per vector. `out` has length `src.len`. The
/// dequantization formula is `float = ((q - zero) / 127) * scale`
/// — matching the spec. Encoding inverse:
/// `q = round(x * 127/scale + zero)`, clamped to [-127, 127].
///
/// Ideal scale + zero land q=-127 → lo and q=127 → hi exactly:
///   scale = (hi - lo) / 2
///   zero  = 127 - 254 * hi / (hi - lo)
/// When that ideal `zero` lies outside [-127, 127] (rare for
/// L2-normalized embeddings, common for distributions far from
/// origin), we clamp `zero` and widen `scale` so the q-range
/// still spans [lo, hi]. Endpoint exactness is sacrificed for
/// representability.
pub fn quantizeF32ToI8Asym(src: []const f32, out: []i8) QuantizedAsym {
    std.debug.assert(src.len == out.len);
    if (src.len == 0) return .{ .scale = 0, .zero = 0 };
    var lo: f32 = std.math.inf(f32);
    var hi: f32 = -std.math.inf(f32);
    for (src) |x| {
        const xv = if (std.math.isNan(x)) 0 else x;
        if (xv < lo) lo = xv;
        if (xv > hi) hi = xv;
    }
    if (hi <= lo) {
        @memset(out, 0);
        return .{ .scale = 0, .zero = 0 };
    }
    var scale: f32 = (hi - lo) / 2.0;
    const ideal_zero: f32 = 127.0 - 254.0 * hi / (hi - lo);
    var zero_i: i32 = @intFromFloat(@round(ideal_zero));
    if (zero_i > 127) {
        zero_i = 127;
        // q=-127 must cover lo: (-127 - 127)/127 * scale ≤ lo
        //   →   -2*scale ≤ lo   →   scale ≥ -lo/2 (assumes lo < 0).
        const need = -lo / 2.0;
        if (need > scale) scale = need;
    } else if (zero_i < -127) {
        zero_i = -127;
        // q=127 must cover hi: (127 + 127)/127 * scale ≥ hi
        //   →   2*scale ≥ hi    →    scale ≥ hi/2 (assumes hi > 0).
        const need = hi / 2.0;
        if (need > scale) scale = need;
    }
    const zero: i8 = @intCast(zero_i);
    const inv = 127.0 / scale;
    const zf: f32 = @floatFromInt(zero);
    for (src, out) |x, *q| {
        const xv = if (std.math.isNan(x)) 0 else x;
        var qf = xv * inv + zf;
        if (qf >= 0) qf += 0.5 else qf -= 0.5;
        var qi: i32 = @intFromFloat(qf);
        if (qi > 127) qi = 127;
        if (qi < -127) qi = -127;
        q.* = @intCast(qi);
    }
    return .{ .scale = scale, .zero = zero };
}

pub fn dequantizeI8Asym(values: []const i8, scale: f32, zero: i8, out: []f32) void {
    std.debug.assert(values.len == out.len);
    if (scale == 0) {
        @memset(out, 0);
        return;
    }
    const inv: f32 = scale / 127.0;
    const z: f32 = @floatFromInt(zero);
    for (values, out) |q, *v| {
        const qf: f32 = @floatFromInt(q);
        v.* = (qf - z) * inv;
    }
}

/// Decode an entire coverage's `vec.bin` body into `out` as f32,
/// row-major `[count][dim]`.
///
/// One call per coverage rather than per row: FFI consumers pay
/// per-call dispatch, and a 500-row × 1536-dim coverage is 500 crossings
/// the caller does not need to make. It also keeps every dtype's layout
/// knowledge here — a binding that dequantized on its own side would be
/// a second implementation to keep in step with `recordBytes`.
///
/// `out.len` must be exactly `count * dim`.
pub fn decodeAllF32(vec: ParsedVecPart, out: []f32) Error!void {
    const dim: usize = vec.header.dim;
    const count: usize = vec.header.count;
    if (out.len != count * dim) return Error.InvalidRange;

    var i: usize = 0;
    while (i < count) : (i += 1) {
        const rec = try vec.record(@intCast(i));
        const dst = out[i * dim ..][0..dim];
        switch (vec.header.dtype) {
            .f32 => {
                for (dst, 0..) |*v, j| {
                    v.* = @bitCast(std.mem.readInt(u32, rec[j * 4 ..][0..4], .little));
                }
            },
            .binary16 => {
                for (dst, 0..) |*v, j| {
                    v.* = binary16ToF32(std.mem.readInt(u16, rec[j * 2 ..][0..2], .little));
                }
            },
            .bfloat16 => {
                for (dst, 0..) |*v, j| {
                    v.* = bfloat16ToF32(std.mem.readInt(u16, rec[j * 2 ..][0..2], .little));
                }
            },
            .int8_sym_per_vec => {
                const scale: f32 = @bitCast(std.mem.readInt(u32, rec[0..4], .little));
                const q: []const i8 = @ptrCast(rec[4..]);
                dequantizeI8Sym(q, scale, dst);
            },
            .int8_asym_per_vec => {
                const scale: f32 = @bitCast(std.mem.readInt(u32, rec[0..4], .little));
                const zero: i8 = @bitCast(rec[4]);
                const q: []const i8 = @ptrCast(rec[5..]);
                dequantizeI8Asym(q, scale, zero, dst);
            },
        }
    }
}

// ---------------------------------------------------------------
// Number canonicalization (Ryu shortest-round-trip)
// ---------------------------------------------------------------

/// Reject-on-non-finite. Canonical form per the spec:
///
/// - `0.0` and `-0.0` both emit as `"0"`.
/// - Finite non-zero values emit via `std.fmt.float.render` in
///   scientific mode with default (Ryu shortest-round-trip)
///   precision.
/// - NaN, +inf, -inf return `Error.MalformedNumber` — they MUST
///   NOT appear in a hashable cell payload per the spec.
///
/// `out` receives the canonical bytes (no leading whitespace, no
/// trailing newline). Returns a sub-slice of `out` (typed as
/// `[]const u8` to match `std.fmt.float.render`'s return type).
pub fn canonicalizeNumber(src: []const u8, out: []u8) Error![]const u8 {
    // Trim ASCII whitespace (the OOXML `<v>` element body can carry
    // surrounding whitespace under some writers).
    var trimmed = src;
    while (trimmed.len > 0 and isAsciiSpace(trimmed[0])) trimmed = trimmed[1..];
    while (trimmed.len > 0 and isAsciiSpace(trimmed[trimmed.len - 1])) trimmed.len -= 1;
    if (trimmed.len == 0) return Error.MalformedNumber;

    // Reject locale-specific comma decimal. OOXML `<v>` is
    // locale-independent (period decimal separator only).
    for (trimmed) |b| {
        if (b == ',') return Error.MalformedNumber;
    }

    const v = std.fmt.parseFloat(f64, trimmed) catch return Error.MalformedNumber;
    if (std.math.isNan(v) or std.math.isInf(v)) return Error.MalformedNumber;
    // Both +0 and -0 emit as "0".
    if (v == 0.0) {
        if (out.len < 1) return Error.BufferTooSmall;
        out[0] = '0';
        return out[0..1];
    }
    return std.fmt.float.render(out, v, .{ .mode = .scientific, .precision = null }) catch
        Error.BufferTooSmall;
}

fn isAsciiSpace(b: u8) bool {
    return switch (b) {
        ' ', '\t', '\n', '\r', 0x0B, 0x0C => true,
        else => false,
    };
}

// ---------------------------------------------------------------
// Hash canonicalization
// ---------------------------------------------------------------

pub const CanonicalCell = union(enum) {
    blank,
    string: []const u8,
    boolean: bool,
    number: []const u8,
    error_value: []const u8,
};

/// Canonicalize visible text for embedding hashes: trim Unicode
/// `White_Space=Y` codepoints from both ends, then NFC-normalize.
/// The normalized bytes are appended to `out`.
pub fn canonicalizeText(
    allocator: Allocator,
    input: []const u8,
    out: *std.ArrayListUnmanaged(u8),
) Error!void {
    const trimmed = try trimUnicodeWhitespace(input);
    const normalized = nfc.normalize(allocator, trimmed) catch |err| switch (err) {
        error.OutOfMemory => return Error.OutOfMemory,
        error.InvalidUtf8 => return Error.InvalidUtf8,
        else => return Error.UnicodeNormalizationFailed,
    };
    defer allocator.free(normalized);
    try out.appendSlice(allocator, normalized);
}

/// Compose and hash:
///
/// `worksheet_target \x1F row_decimal \x1F cell_kind \x1F payload`
///
/// `scratch` is caller-owned and is cleared before use.
pub fn xxh3Canonical(
    allocator: Allocator,
    worksheet_target: []const u8,
    row: u32,
    cell: CanonicalCell,
    scratch: *std.ArrayListUnmanaged(u8),
) Error!u64 {
    if (row == 0 or row > EXCEL_MAX_ROW) return Error.InvalidRange;
    scratch.clearRetainingCapacity();
    try scratch.appendSlice(allocator, worksheet_target);
    try scratch.append(allocator, 0x1F);

    var row_buf: [10]u8 = undefined;
    const row_s = std.fmt.bufPrint(&row_buf, "{d}", .{row}) catch unreachable;
    try scratch.appendSlice(allocator, row_s);
    try scratch.append(allocator, 0x1F);

    switch (cell) {
        .blank => {
            try scratch.append(allocator, 'b');
            try scratch.append(allocator, 0x1F);
        },
        .string => |s| {
            try scratch.append(allocator, 's');
            try scratch.append(allocator, 0x1F);
            try canonicalizeText(allocator, s, scratch);
        },
        .boolean => |b| {
            try scratch.append(allocator, 'B');
            try scratch.append(allocator, 0x1F);
            try scratch.append(allocator, if (b) '1' else '0');
        },
        .number => |raw| {
            try scratch.append(allocator, 'n');
            try scratch.append(allocator, 0x1F);
            var num_buf: [64]u8 = undefined;
            const canonical = try canonicalizeNumber(raw, &num_buf);
            try scratch.appendSlice(allocator, canonical);
        },
        .error_value => |raw| {
            try scratch.append(allocator, 'e');
            try scratch.append(allocator, 0x1F);
            try scratch.appendSlice(allocator, raw);
        },
    }
    return xxh3(scratch.items);
}

fn trimUnicodeWhitespace(input: []const u8) Error![]const u8 {
    if (!std.unicode.utf8ValidateSlice(input)) return Error.InvalidUtf8;
    var i: usize = 0;
    var first_non_ws: ?usize = null;
    var end_non_ws: usize = 0;
    while (i < input.len) {
        const seq_len = std.unicode.utf8ByteSequenceLength(input[i]) catch return Error.InvalidUtf8;
        const cp = std.unicode.utf8Decode(input[i .. i + seq_len]) catch return Error.InvalidUtf8;
        if (!isUnicodeWhitespace(cp)) {
            if (first_non_ws == null) first_non_ws = i;
            end_non_ws = i + seq_len;
        }
        i += seq_len;
    }
    const start = first_non_ws orelse return input[0..0];
    return input[start..end_non_ws];
}

fn isUnicodeWhitespace(cp: u21) bool {
    return switch (cp) {
        0x0009...0x000D,
        0x0020,
        0x0085,
        0x00A0,
        0x1680,
        0x2000...0x200A,
        0x2028,
        0x2029,
        0x202F,
        0x205F,
        0x3000,
        => true,
        else => false,
    };
}

// ---------------------------------------------------------------
// xxh3-64 thin wrapper
// ---------------------------------------------------------------

/// xxh3-64 of `bytes` with the pinned `HASH_SEED`. Wraps
/// `std.hash.XxHash3` so consumers don't pin the specific stdlib
/// symbol; if Zig renames it across versions we change it in one
/// place.
pub fn xxh3(bytes: []const u8) u64 {
    return std.hash.XxHash3.hash(HASH_SEED, bytes);
}

// ---------------------------------------------------------------
// Encoders — emb-3a write path
// ---------------------------------------------------------------

/// Caller-fillable spec for one `<coverage>` element. Matches the
/// schema fields the design pins; XML escaping is handled by
/// `encodeIndexXml`.
pub const CoverageSpec = struct {
    id: []const u8,
    worksheet_target: []const u8,
    range: []const u8,
    column: []const u8,
    count: u32,
    include_formulas: bool,
    vec_rid: []const u8,
    hash_rid: []const u8,
};

pub const IndexSpec = struct {
    model: []const u8,
    dim: u32,
    dtype: Dtype,
    hash_algo: []const u8,
    coverages: []const CoverageSpec,
};

pub const RelSpec = struct {
    id: []const u8,
    type: []const u8,
    target: []const u8,
};

/// Write the 24-byte vec.bin header at the start of `out`. Asserts
/// `out.len >= VEC_HEADER_BYTES`. Returns a slice into `out` covering
/// the header.
pub fn writeVecHeader(out: []u8, header: ParsedVecHeader) []u8 {
    std.debug.assert(out.len >= VEC_HEADER_BYTES);
    std.mem.writeInt(u32, out[0..4], VEC_MAGIC, .little);
    std.mem.writeInt(u32, out[4..8], WIRE_VERSION, .little);
    std.mem.writeInt(u32, out[8..12], header.dim, .little);
    std.mem.writeInt(u32, out[12..16], header.count, .little);
    out[16] = @intFromEnum(header.dtype);
    @memset(out[17..VEC_HEADER_BYTES], 0);
    return out[0..VEC_HEADER_BYTES];
}

/// Write the 16-byte hashes.bin header at the start of `out`.
pub fn writeHashHeader(out: []u8, header: ParsedHashHeader) []u8 {
    std.debug.assert(out.len >= HASH_HEADER_BYTES);
    std.mem.writeInt(u32, out[0..4], HASH_MAGIC, .little);
    std.mem.writeInt(u32, out[4..8], WIRE_VERSION, .little);
    std.mem.writeInt(u32, out[8..12], header.count, .little);
    std.mem.writeInt(u32, out[12..16], 0, .little);
    return out[0..HASH_HEADER_BYTES];
}

/// Build a complete vec.bin: 24-byte header followed by
/// `header.count * header.dtype.recordBytes(header.dim)` body bytes
/// supplied by the caller. The body MUST already be in the on-disk
/// layout for the chosen dtype — for int8 quantization that means
/// `f32 scale; i8 values[dim]` per row (sym) or `f32 scale; i8 zero;
/// i8 values[dim]` (asym). Use `quantizeF32ToI8Sym/Asym` upstream to
/// produce the body bytes; this function does NOT quantize.
pub fn encodeVecPart(allocator: Allocator, header: ParsedVecHeader, body: []const u8) Error![]u8 {
    const expected_body: usize = @as(usize, header.count) * header.dtype.recordBytes(header.dim);
    if (body.len != expected_body) return Error.BodySizeMismatch;
    var buf = try allocator.alloc(u8, VEC_HEADER_BYTES + body.len);
    errdefer allocator.free(buf);
    _ = writeVecHeader(buf, header);
    @memcpy(buf[VEC_HEADER_BYTES..], body);
    return buf;
}

/// Build a complete hashes.bin: 16-byte header + count×u64 LE.
pub fn encodeHashesPart(allocator: Allocator, hashes: []const u64) Error![]u8 {
    if (hashes.len > std.math.maxInt(u32)) return Error.CountMismatch;
    var buf = try allocator.alloc(u8, HASH_HEADER_BYTES + hashes.len * @sizeOf(u64));
    errdefer allocator.free(buf);
    _ = writeHashHeader(buf, .{ .version = WIRE_VERSION, .count = @intCast(hashes.len) });
    for (hashes, 0..) |h, i| {
        const off = HASH_HEADER_BYTES + i * @sizeOf(u64);
        std.mem.writeInt(u64, buf[off..][0..8], h, .little);
    }
    return buf;
}

/// Encode ONE vector into its on-disk record for `dtype` — the f32
/// wire bytes, or the symmetric int8 quantizer's per-row
/// `f32 scale; i8[dim]` — into `dst`, which must be exactly
/// `dtype.recordBytes(dim)` bytes. The one encoder behind
/// `zlsx embed --vectors` and `zlsx_editor_set_embeddings`, so the
/// two surfaces cannot drift on the layout `decodeAllF32` reads back.
/// binary16 / bfloat16 / the asymmetric int8 layout have decoders but
/// no writer yet: `UnsupportedDtype`, before any byte of `dst` moves.
pub fn encodeVectorRecord(dtype: Dtype, dim: u32, vec: []const f32, dst: []u8) Error!void {
    if (dim == 0) return Error.DimensionOutOfRange;
    if (vec.len != dim) return Error.DimensionMismatch;
    if (dst.len != dtype.recordBytes(dim)) return Error.BodySizeMismatch;
    switch (dtype) {
        .f32 => for (vec, 0..) |f, i| {
            std.mem.writeInt(u32, dst[i * 4 ..][0..4], @bitCast(f), .little);
        },
        .int8_sym_per_vec => {
            const codes: []i8 = @ptrCast(dst[4..][0..dim]);
            const q = quantizeF32ToI8Sym(vec, codes);
            std.mem.writeInt(u32, dst[0..4], @bitCast(q.scale), .little);
        },
        .binary16, .bfloat16, .int8_asym_per_vec => return Error.UnsupportedDtype,
    }
}

/// Encode a row-major `[count][dim]` f32 matrix into a complete
/// vec.bin body (`count * dtype.recordBytes(dim)` bytes) for
/// `Workbook.setEmbeddings`. `vectors.len` must be a whole number of
/// rows (`DimensionMismatch` otherwise); an empty matrix is an empty
/// body. Caller frees.
pub fn encodeVectorBody(allocator: Allocator, dtype: Dtype, dim: u32, vectors: []const f32) Error![]u8 {
    if (dim == 0) return Error.DimensionOutOfRange;
    if (vectors.len % dim != 0) return Error.DimensionMismatch;
    switch (dtype) {
        .f32, .int8_sym_per_vec => {},
        .binary16, .bfloat16, .int8_asym_per_vec => return Error.UnsupportedDtype,
    }
    const count = vectors.len / dim;
    const rec = dtype.recordBytes(dim);
    const body = try allocator.alloc(u8, count * rec);
    errdefer allocator.free(body);
    for (0..count) |r| {
        try encodeVectorRecord(dtype, dim, vectors[r * dim ..][0..dim], body[r * rec ..][0..rec]);
    }
    return body;
}

/// Build the index.xml manifest from `spec`. XML-escapes user-
/// provided text fields (`model` and the per-coverage `id`,
/// `worksheet_target`, `range`, `column`).
pub fn encodeIndexXml(allocator: Allocator, spec: IndexSpec) Error![]u8 {
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    errdefer buf.deinit(allocator);

    try buf.appendSlice(allocator, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
    try buf.appendSlice(allocator, "<embeddings xmlns=\"");
    try appendXmlEscaped(&buf, allocator, INDEX_NAMESPACE);
    try buf.appendSlice(allocator, "\" version=\"1\" model=\"");
    try appendXmlEscaped(&buf, allocator, spec.model);
    try buf.appendSlice(allocator, "\" dim=\"");
    try appendDecimal(&buf, allocator, spec.dim);
    try buf.appendSlice(allocator, "\" dtype=\"");
    try buf.appendSlice(allocator, dtypeXmlName(spec.dtype));
    try buf.appendSlice(allocator, "\" hash_algo=\"");
    try appendXmlEscaped(&buf, allocator, spec.hash_algo);
    try buf.appendSlice(allocator, "\">\n");

    for (spec.coverages) |c| {
        try buf.appendSlice(allocator, "  <coverage id=\"");
        try appendXmlEscaped(&buf, allocator, c.id);
        try buf.appendSlice(allocator, "\" worksheet_target=\"");
        try appendXmlEscaped(&buf, allocator, c.worksheet_target);
        try buf.appendSlice(allocator, "\" range=\"");
        try appendXmlEscaped(&buf, allocator, c.range);
        try buf.appendSlice(allocator, "\" column=\"");
        try appendXmlEscaped(&buf, allocator, c.column);
        try buf.appendSlice(allocator, "\" count=\"");
        try appendDecimal(&buf, allocator, c.count);
        try buf.appendSlice(allocator, "\" include_formulas=\"");
        try buf.appendSlice(allocator, if (c.include_formulas) "true" else "false");
        try buf.appendSlice(allocator, "\" vec_rId=\"");
        try appendXmlEscaped(&buf, allocator, c.vec_rid);
        try buf.appendSlice(allocator, "\" hash_rId=\"");
        try appendXmlEscaped(&buf, allocator, c.hash_rid);
        try buf.appendSlice(allocator, "\"/>\n");
    }

    try buf.appendSlice(allocator, "</embeddings>\n");
    return buf.toOwnedSlice(allocator);
}

/// Build the index.xml.rels file linking each coverage's vec.bin /
/// hashes.bin sub-part via its rId.
pub fn encodeIndexRelsXml(allocator: Allocator, rels: []const RelSpec) Error![]u8 {
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    errdefer buf.deinit(allocator);

    try buf.appendSlice(allocator, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
    try buf.appendSlice(allocator, "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
    for (rels) |r| {
        try buf.appendSlice(allocator, "  <Relationship Id=\"");
        try appendXmlEscaped(&buf, allocator, r.id);
        try buf.appendSlice(allocator, "\" Type=\"");
        try appendXmlEscaped(&buf, allocator, r.type);
        try buf.appendSlice(allocator, "\" Target=\"");
        try appendXmlEscaped(&buf, allocator, r.target);
        try buf.appendSlice(allocator, "\"/>\n");
    }
    try buf.appendSlice(allocator, "</Relationships>\n");
    return buf.toOwnedSlice(allocator);
}

fn dtypeXmlName(d: Dtype) []const u8 {
    return switch (d) {
        .f32 => "f32",
        .binary16 => "binary16",
        .bfloat16 => "bfloat16",
        .int8_sym_per_vec => "int8-sym-per-vec",
        .int8_asym_per_vec => "int8-asym-per-vec",
    };
}

fn appendDecimal(buf: *std.ArrayListUnmanaged(u8), allocator: Allocator, value: u32) Error!void {
    var num: [16]u8 = undefined;
    const s = std.fmt.bufPrint(&num, "{d}", .{value}) catch return Error.BufferTooSmall;
    try buf.appendSlice(allocator, s);
}

fn appendXmlEscaped(buf: *std.ArrayListUnmanaged(u8), allocator: Allocator, s: []const u8) Error!void {
    // OOXML attribute-value escaping: & < > " (single quote left
    // alone since we use double-quoted attributes throughout). A C0
    // control byte has no escape — XML 1.0 forbids it outright — so
    // it refuses here as it does on every other writer channel.
    for (s) |c| switch (c) {
        0x00...0x08, 0x0B, 0x0C, 0x0E...0x1F => return Error.InvalidXmlByte,
        '&' => try buf.appendSlice(allocator, "&amp;"),
        '<' => try buf.appendSlice(allocator, "&lt;"),
        '>' => try buf.appendSlice(allocator, "&gt;"),
        '"' => try buf.appendSlice(allocator, "&quot;"),
        else => try buf.append(allocator, c),
    };
}

// ---------------------------------------------------------------
// Tests
// ---------------------------------------------------------------

const testing = std.testing;

test "VecHeader/HashHeader: compile-time size and offset asserts hold" {
    try testing.expectEqual(@as(usize, 24), VEC_HEADER_BYTES);
    try testing.expectEqual(@as(usize, 16), HASH_HEADER_BYTES);
    try testing.expectEqual(@as(usize, 0), @offsetOf(VecHeader, "magic"));
    try testing.expectEqual(@as(usize, 16), @offsetOf(VecHeader, "dtype_byte"));
    try testing.expectEqual(@as(usize, 17), @offsetOf(VecHeader, "reserved"));
}

test "Dtype.recordBytes: per-record sizing matches the spec" {
    try testing.expectEqual(@as(usize, 4 * 1536), Dtype.f32.recordBytes(1536));
    try testing.expectEqual(@as(usize, 2 * 1536), Dtype.binary16.recordBytes(1536));
    try testing.expectEqual(@as(usize, 2 * 1536), Dtype.bfloat16.recordBytes(1536));
    try testing.expectEqual(@as(usize, 4 + 1536), Dtype.int8_sym_per_vec.recordBytes(1536));
    try testing.expectEqual(@as(usize, 4 + 1 + 1536), Dtype.int8_asym_per_vec.recordBytes(1536));
}

test "Dtype.fromU8: known values map; out-of-range rejects" {
    try testing.expectEqual(Dtype.f32, try Dtype.fromU8(0));
    try testing.expectEqual(Dtype.int8_asym_per_vec, try Dtype.fromU8(4));
    try testing.expectError(Error.InvalidDtype, Dtype.fromU8(5));
    try testing.expectError(Error.InvalidDtype, Dtype.fromU8(255));
}

test "Dtype string mapping rejects ambiguous f16 spelling" {
    try testing.expectEqual(Dtype.f32, try Dtype.fromString("f32"));
    try testing.expectEqual(Dtype.binary16, try Dtype.fromString("binary16"));
    try testing.expectEqual(Dtype.bfloat16, try Dtype.fromString("bfloat16"));
    try testing.expectEqual(Dtype.int8_sym_per_vec, try Dtype.fromString("int8-sym-per-vec"));
    try testing.expectEqual(Dtype.int8_asym_per_vec, try Dtype.fromString("int8-asym-per-vec"));
    try testing.expectEqualStrings("int8-sym-per-vec", Dtype.int8_sym_per_vec.string());
    try testing.expectError(Error.InvalidDtype, Dtype.fromString("f16"));
}

test "parseA1Range: validates Excel bounds and row ordering" {
    const r = try parseA1Range("A2:C10");
    try testing.expectEqual(@as(u32, 2), r.first.row);
    try testing.expectEqual(@as(u32, 0), r.first.col);
    try testing.expectEqual(@as(u32, 10), r.last.row);
    try testing.expectEqual(@as(u32, 2), r.last.col);
    try testing.expectEqual(@as(u32, 9), r.rowCount());
    try testing.expectEqual(@as(u32, 16_383), try parseColumnName("XFD"));
    try testing.expectError(Error.InvalidRange, parseA1Range("A0:A1"));
    try testing.expectError(Error.InvalidRange, parseA1Range("A2:A1"));
    try testing.expectError(Error.InvalidRange, parseA1Range("XFE1:XFE2"));
}

test "parseIndex: parses plural coverage blocks" {
    const xml =
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<embeddings xmlns="http://schemas.fabre.me/zlsx/2026/embeddings"
        \\            version="1"
        \\            model="text-embedding-3-small"
        \\            dim="1536"
        \\            dtype="int8-sym-per-vec"
        \\            hash_algo="xxh3-64">
        \\  <coverage id="title"
        \\            worksheet_target="worksheets/sheet1.xml"
        \\            range="A2:A10001" column="A"
        \\            count="10000"
        \\            include_formulas="false"
        \\            vec_rId="rId1" hash_rId="rId2"/>
        \\  <coverage id="body"
        \\            worksheet_target="worksheets/sheet1.xml"
        \\            range="B2:B10001" column="B"
        \\            count="10000"
        \\            include_formulas="true"
        \\            vec_rId="rId3" hash_rId="rId4"/>
        \\</embeddings>
    ;
    var coverages: [2]Coverage = undefined;
    const index = try parseIndex(xml, &coverages);
    try testing.expectEqual(@as(u32, 1), index.version);
    try testing.expectEqualStrings("text-embedding-3-small", index.model);
    try testing.expectEqual(@as(u32, 1536), index.dim);
    try testing.expectEqual(Dtype.int8_sym_per_vec, index.dtype);
    try testing.expectEqual(@as(usize, 2), index.coverages.len);
    try testing.expectEqualStrings("title", index.coverages[0].id);
    try testing.expectEqual(@as(u32, 10_000), index.coverages[0].count);
    try testing.expect(!index.coverages[0].include_formulas);
    try testing.expect(index.coverages[1].include_formulas);
    try testing.expectEqual(@as(u32, 1), index.coverages[1].column_idx);
}

test "parseIndexRelationships: validates relationship types and deterministic targets" {
    const index_xml =
        \\<embeddings xmlns="http://schemas.fabre.me/zlsx/2026/embeddings" version="1" model="m" dim="3" dtype="f32" hash_algo="xxh3-64">
        \\  <coverage id="title" worksheet_target="worksheets/sheet1.xml" range="A1:A2" column="A" count="2" vec_rId="rId1" hash_rId="rId2"/>
        \\  <coverage id="body" worksheet_target="worksheets/sheet1.xml" range="B1:B2" column="B" count="2" vec_rId="rId3" hash_rId="rId4"/>
        \\</embeddings>
    ;
    var coverages: [2]Coverage = undefined;
    const index = try parseIndex(index_xml, &coverages);

    const rels_xml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" ++
        "<Relationship Id=\"rId1\" Type=\"" ++ REL_TYPE_VEC ++ "\" Target=\"title/vec.bin\"/>" ++
        "<Relationship Id=\"rId2\" Type=\"" ++ REL_TYPE_HASH ++ "\" Target=\"title/hashes.bin\"/>" ++
        "<Relationship Id=\"rId3\" Type=\"" ++ REL_TYPE_VEC ++ "\" Target=\"body/vec.bin\"/>" ++
        "<Relationship Id=\"rId4\" Type=\"" ++ REL_TYPE_HASH ++ "\" Target=\"body/hashes.bin\"/>" ++
        "</Relationships>";
    var rels_buf: [4]IndexRelationship = undefined;
    const rels = try parseIndexRelationships(rels_xml, &rels_buf);
    try testing.expectEqual(@as(usize, 4), rels.len);
    try validateIndexRelationships(index, rels);

    var target_buf: [80]u8 = undefined;
    try testing.expectEqualStrings("title/vec.bin", try expectedVecTarget("title", &target_buf));
    try testing.expectEqualStrings("title/hashes.bin", try expectedHashTarget("title", &target_buf));
}

test "parseIndexRelationships: rejects bad rel ids, targets, and modes" {
    const index_xml =
        \\<embeddings xmlns="http://schemas.fabre.me/zlsx/2026/embeddings" version="1" model="m" dim="3" dtype="f32" hash_algo="xxh3-64">
        \\  <coverage id="title" worksheet_target="worksheets/sheet1.xml" range="A1:A1" column="A" count="1" vec_rId="rId1" hash_rId="rId2"/>
        \\</embeddings>
    ;
    var coverages: [1]Coverage = undefined;
    const index = try parseIndex(index_xml, &coverages);

    const duplicate_id =
        "<Relationships>" ++
        "<Relationship Id=\"rId1\" Type=\"" ++ REL_TYPE_VEC ++ "\" Target=\"title/vec.bin\"/>" ++
        "<Relationship Id=\"rId1\" Type=\"" ++ REL_TYPE_HASH ++ "\" Target=\"title/hashes.bin\"/>" ++
        "</Relationships>";
    var rels_buf: [2]IndexRelationship = undefined;
    try testing.expectError(Error.DuplicateRelationshipId, parseIndexRelationships(duplicate_id, &rels_buf));

    const external =
        "<Relationships>" ++
        "<Relationship Id=\"rId1\" Type=\"" ++ REL_TYPE_VEC ++ "\" Target=\"https://example.com/v\" TargetMode=\"External\"/>" ++
        "<Relationship Id=\"rId2\" Type=\"" ++ REL_TYPE_HASH ++ "\" Target=\"title/hashes.bin\"/>" ++
        "</Relationships>";
    const external_rels = try parseIndexRelationships(external, &rels_buf);
    try testing.expectError(Error.InvalidIndexRelationship, validateIndexRelationships(index, external_rels));

    const duplicate_target =
        "<Relationships>" ++
        "<Relationship Id=\"rId1\" Type=\"" ++ REL_TYPE_VEC ++ "\" Target=\"title/vec.bin\"/>" ++
        "<Relationship Id=\"rId2\" Type=\"" ++ REL_TYPE_HASH ++ "\" Target=\"title/vec.bin\"/>" ++
        "</Relationships>";
    const duplicate_target_rels = try parseIndexRelationships(duplicate_target, &rels_buf);
    try testing.expectError(Error.DuplicateRelationshipTarget, validateIndexRelationships(index, duplicate_target_rels));

    const wrong_target =
        "<Relationships>" ++
        "<Relationship Id=\"rId1\" Type=\"" ++ REL_TYPE_VEC ++ "\" Target=\"other/vec.bin\"/>" ++
        "<Relationship Id=\"rId2\" Type=\"" ++ REL_TYPE_HASH ++ "\" Target=\"title/hashes.bin\"/>" ++
        "</Relationships>";
    const wrong_target_rels = try parseIndexRelationships(wrong_target, &rels_buf);
    try testing.expectError(Error.InvalidIndexRelationship, validateIndexRelationships(index, wrong_target_rels));

    const missing =
        "<Relationships>" ++
        "<Relationship Id=\"rId1\" Type=\"" ++ REL_TYPE_VEC ++ "\" Target=\"title/vec.bin\"/>" ++
        "</Relationships>";
    const missing_rels = try parseIndexRelationships(missing, &rels_buf);
    try testing.expectError(Error.MissingRelationship, validateIndexRelationships(index, missing_rels));
}

test "parseIndex: rejects duplicate ids, overlapping coverage, and count mismatch" {
    const duplicate_id =
        \\<embeddings xmlns="http://schemas.fabre.me/zlsx/2026/embeddings" version="1" model="m" dim="3" dtype="f32" hash_algo="xxh3-64">
        \\  <coverage id="dup" worksheet_target="worksheets/sheet1.xml" range="A1:A2" column="A" count="2" vec_rId="rId1" hash_rId="rId2"/>
        \\  <coverage id="dup" worksheet_target="worksheets/sheet2.xml" range="A1:A2" column="A" count="2" vec_rId="rId3" hash_rId="rId4"/>
        \\</embeddings>
    ;
    var coverages: [2]Coverage = undefined;
    try testing.expectError(Error.DuplicateCoverageId, parseIndex(duplicate_id, &coverages));

    const overlap =
        \\<embeddings xmlns="http://schemas.fabre.me/zlsx/2026/embeddings" version="1" model="m" dim="3" dtype="f32" hash_algo="xxh3-64">
        \\  <coverage id="a" worksheet_target="worksheets/sheet1.xml" range="A1:B2" column="A" count="2" vec_rId="rId1" hash_rId="rId2"/>
        \\  <coverage id="b" worksheet_target="worksheets/sheet1.xml" range="B2:C4" column="B" count="3" vec_rId="rId3" hash_rId="rId4"/>
        \\</embeddings>
    ;
    try testing.expectError(Error.CoverageOverlap, parseIndex(overlap, &coverages));

    const count_mismatch =
        \\<embeddings xmlns="http://schemas.fabre.me/zlsx/2026/embeddings" version="1" model="m" dim="3" dtype="f32" hash_algo="xxh3-64">
        \\  <coverage id="a" worksheet_target="worksheets/sheet1.xml" range="A1:A2" column="A" count="3" vec_rId="rId1" hash_rId="rId2"/>
        \\</embeddings>
    ;
    try testing.expectError(Error.CoverageCountMismatch, parseIndex(count_mismatch, coverages[0..1]));
}

test "parseIndex: rejects unsupported manifest values" {
    const bad_dtype =
        \\<embeddings xmlns="http://schemas.fabre.me/zlsx/2026/embeddings" version="1" model="m" dim="3" dtype="f16" hash_algo="xxh3-64">
        \\  <coverage id="a" worksheet_target="worksheets/sheet1.xml" range="A1:A1" column="A" count="1" vec_rId="rId1" hash_rId="rId2"/>
        \\</embeddings>
    ;
    var coverages: [1]Coverage = undefined;
    try testing.expectError(Error.InvalidDtype, parseIndex(bad_dtype, &coverages));

    const future_version =
        \\<embeddings xmlns="http://schemas.fabre.me/zlsx/2026/embeddings" version="2" model="m" dim="3" dtype="f32" hash_algo="xxh3-64">
        \\  <coverage id="a" worksheet_target="worksheets/sheet1.xml" range="A1:A1" column="A" count="1" vec_rId="rId1" hash_rId="rId2"/>
        \\</embeddings>
    ;
    try testing.expectError(Error.UnsupportedVersion, parseIndex(future_version, &coverages));

    const bad_hash =
        \\<embeddings xmlns="http://schemas.fabre.me/zlsx/2026/embeddings" version="1" model="m" dim="3" dtype="f32" hash_algo="sha256">
        \\  <coverage id="a" worksheet_target="worksheets/sheet1.xml" range="A1:A1" column="A" count="1" vec_rId="rId1" hash_rId="rId2"/>
        \\</embeddings>
    ;
    try testing.expectError(Error.UnsupportedHashAlgorithm, parseIndex(bad_hash, &coverages));
}

test "parseVecHeader: well-formed header round-trips" {
    var buf: [VEC_HEADER_BYTES + 16]u8 = undefined;
    @memset(&buf, 0);
    std.mem.writeInt(u32, buf[0..4], VEC_MAGIC, .little);
    std.mem.writeInt(u32, buf[4..8], WIRE_VERSION, .little);
    std.mem.writeInt(u32, buf[8..12], 4, .little); // dim 4
    std.mem.writeInt(u32, buf[12..16], 1, .little); // count 1
    buf[16] = @intFromEnum(Dtype.f32);
    // body: 1 record × 4 dim × 4 bytes = 16 bytes (zeroed)

    const got = try parseVecHeader(&buf);
    try testing.expectEqual(@as(u32, 1), got.version);
    try testing.expectEqual(@as(u32, 4), got.dim);
    try testing.expectEqual(@as(u32, 1), got.count);
    try testing.expectEqual(Dtype.f32, got.dtype);
}

test "parseVecHeader: bad magic rejected" {
    var buf: [VEC_HEADER_BYTES]u8 = undefined;
    @memset(&buf, 0);
    std.mem.writeInt(u32, buf[0..4], 0xDEADBEEF, .little);
    try testing.expectError(Error.BadMagic, parseVecHeader(&buf));
}

test "parseVecHeader: unsupported version rejected" {
    var buf: [VEC_HEADER_BYTES]u8 = undefined;
    @memset(&buf, 0);
    std.mem.writeInt(u32, buf[0..4], VEC_MAGIC, .little);
    std.mem.writeInt(u32, buf[4..8], 2, .little); // future version
    std.mem.writeInt(u32, buf[8..12], 1, .little);
    std.mem.writeInt(u32, buf[12..16], 0, .little);
    buf[16] = 0;
    try testing.expectError(Error.UnsupportedVersion, parseVecHeader(&buf));
}

test "parseVecHeader: non-zero reserved bytes rejected" {
    var buf: [VEC_HEADER_BYTES]u8 = undefined;
    @memset(&buf, 0);
    std.mem.writeInt(u32, buf[0..4], VEC_MAGIC, .little);
    std.mem.writeInt(u32, buf[4..8], WIRE_VERSION, .little);
    std.mem.writeInt(u32, buf[8..12], 1, .little);
    std.mem.writeInt(u32, buf[12..16], 0, .little);
    buf[16] = 0;
    buf[20] = 0xFF; // dirty reserved
    try testing.expectError(Error.InvalidReservedBytes, parseVecHeader(&buf));
}

test "parseVecHeader: short header rejected" {
    const buf: [10]u8 = .{0} ** 10;
    try testing.expectError(Error.HeaderTooShort, parseVecHeader(&buf));
}

test "parseVecHeader: missing body rejected" {
    var buf: [VEC_HEADER_BYTES]u8 = undefined;
    @memset(&buf, 0);
    std.mem.writeInt(u32, buf[0..4], VEC_MAGIC, .little);
    std.mem.writeInt(u32, buf[4..8], WIRE_VERSION, .little);
    std.mem.writeInt(u32, buf[8..12], 4, .little);
    std.mem.writeInt(u32, buf[12..16], 1, .little); // count 1 needs 16 body bytes
    buf[16] = @intFromEnum(Dtype.f32);
    try testing.expectError(Error.BodyTooShort, parseVecHeader(&buf));
}

test "parseHashHeader: well-formed round-trips" {
    var buf: [HASH_HEADER_BYTES + 8]u8 = undefined;
    @memset(&buf, 0);
    std.mem.writeInt(u32, buf[0..4], HASH_MAGIC, .little);
    std.mem.writeInt(u32, buf[4..8], WIRE_VERSION, .little);
    std.mem.writeInt(u32, buf[8..12], 1, .little);
    const got = try parseHashHeader(&buf);
    try testing.expectEqual(@as(u32, 1), got.version);
    try testing.expectEqual(@as(u32, 1), got.count);
}

test "parseVecPart/parseHashPart: exact-size views and record access" {
    var vec_buf: [VEC_HEADER_BYTES + 16]u8 = undefined;
    @memset(&vec_buf, 0);
    std.mem.writeInt(u32, vec_buf[0..4], VEC_MAGIC, .little);
    std.mem.writeInt(u32, vec_buf[4..8], WIRE_VERSION, .little);
    std.mem.writeInt(u32, vec_buf[8..12], 2, .little);
    std.mem.writeInt(u32, vec_buf[12..16], 2, .little);
    vec_buf[16] = @intFromEnum(Dtype.f32);
    vec_buf[VEC_HEADER_BYTES + 8] = 0xAB;

    const vec = try parseVecPart(&vec_buf);
    try testing.expectEqual(@as(u32, 2), vec.header.count);
    try testing.expectEqual(@as(usize, 16), vec.body.len);
    try testing.expectEqual(@as(usize, 8), (try vec.record(1)).len);
    try testing.expectEqual(@as(u8, 0xAB), (try vec.record(1))[0]);
    try testing.expectError(Error.InvalidRange, vec.record(2));

    var hash_buf: [HASH_HEADER_BYTES + 16]u8 = undefined;
    @memset(&hash_buf, 0);
    std.mem.writeInt(u32, hash_buf[0..4], HASH_MAGIC, .little);
    std.mem.writeInt(u32, hash_buf[4..8], WIRE_VERSION, .little);
    std.mem.writeInt(u32, hash_buf[8..12], 2, .little);
    std.mem.writeInt(u64, hash_buf[HASH_HEADER_BYTES..][0..8], 0x0102030405060708, .little);
    std.mem.writeInt(u64, hash_buf[HASH_HEADER_BYTES + 8 ..][0..8], TOMBSTONE_HASH, .little);

    const hashes = try parseHashPart(&hash_buf);
    try testing.expectEqual(@as(u64, 0x0102030405060708), try hashes.value(0));
    try testing.expectEqual(TOMBSTONE_HASH, try hashes.value(1));
    try testing.expectError(Error.InvalidRange, hashes.value(2));
}

test "parseVecHeader/parseHashHeader: trailing bytes rejected" {
    var vec_buf: [VEC_HEADER_BYTES + 1]u8 = undefined;
    @memset(&vec_buf, 0);
    std.mem.writeInt(u32, vec_buf[0..4], VEC_MAGIC, .little);
    std.mem.writeInt(u32, vec_buf[4..8], WIRE_VERSION, .little);
    std.mem.writeInt(u32, vec_buf[8..12], 1, .little);
    std.mem.writeInt(u32, vec_buf[12..16], 0, .little);
    vec_buf[16] = @intFromEnum(Dtype.f32);
    try testing.expectError(Error.BodySizeMismatch, parseVecHeader(&vec_buf));

    var hash_buf: [HASH_HEADER_BYTES + 1]u8 = undefined;
    @memset(&hash_buf, 0);
    std.mem.writeInt(u32, hash_buf[0..4], HASH_MAGIC, .little);
    std.mem.writeInt(u32, hash_buf[4..8], WIRE_VERSION, .little);
    std.mem.writeInt(u32, hash_buf[8..12], 0, .little);
    try testing.expectError(Error.BodySizeMismatch, parseHashHeader(&hash_buf));
}

test "checkPairConsistent: count mismatch flagged" {
    const v: ParsedVecHeader = .{ .version = 1, .dim = 4, .count = 10, .dtype = .f32 };
    const h_ok: ParsedHashHeader = .{ .version = 1, .count = 10 };
    const h_bad: ParsedHashHeader = .{ .version = 1, .count = 11 };
    try checkPairConsistent(v, h_ok);
    try testing.expectError(Error.CountMismatch, checkPairConsistent(v, h_bad));
}

test "checkCoverageBinary: validates dim, dtype, and counts against index coverage" {
    const index_xml =
        \\<embeddings xmlns="http://schemas.fabre.me/zlsx/2026/embeddings" version="1" model="m" dim="3" dtype="f32" hash_algo="xxh3-64">
        \\  <coverage id="a" worksheet_target="worksheets/sheet1.xml" range="A1:A2" column="A" count="2" vec_rId="rId1" hash_rId="rId2"/>
        \\</embeddings>
    ;
    var coverages: [1]Coverage = undefined;
    const index = try parseIndex(index_xml, &coverages);
    const coverage = index.coverages[0];

    try checkCoverageBinary(
        index,
        coverage,
        .{ .version = 1, .dim = 3, .count = 2, .dtype = .f32 },
        .{ .version = 1, .count = 2 },
    );
    try testing.expectError(
        Error.DimensionMismatch,
        checkCoverageBinary(index, coverage, .{ .version = 1, .dim = 4, .count = 2, .dtype = .f32 }, .{ .version = 1, .count = 2 }),
    );
    try testing.expectError(
        Error.DtypeMismatch,
        checkCoverageBinary(index, coverage, .{ .version = 1, .dim = 3, .count = 2, .dtype = .bfloat16 }, .{ .version = 1, .count = 2 }),
    );
    try testing.expectError(
        Error.CountMismatch,
        checkCoverageBinary(index, coverage, .{ .version = 1, .dim = 3, .count = 3, .dtype = .f32 }, .{ .version = 1, .count = 2 }),
    );
}

test "f32ToBfloat16: representable values round-trip exactly" {
    // bf16 is the top 16 bits of f32; exact-representable floats
    // (those with low 16 bits zero in their f32 repr) survive
    // unchanged.
    const exact = [_]f32{ 0.0, 1.0, -1.0, 2.0, -2.0, 0.5, -0.5 };
    for (exact) |x| {
        const b = f32ToBfloat16(x);
        const back = bfloat16ToF32(b);
        try testing.expectEqual(x, back);
    }
}

test "f32ToBfloat16: NaN canonicalized to qNaN 0x7FC0" {
    const nan = std.math.nan(f32);
    try testing.expectEqual(@as(u16, 0x7FC0), f32ToBfloat16(nan));
}

test "f32ToBfloat16: round-to-nearest-even on tie" {
    // 0x3F804000 = 1.0 + epsilon halfway between two bf16 reps.
    // bf16 round-to-nearest-even should pick the even neighbour.
    // 0x3F80 = 1.0 (LSB 0, even); 0x3F81 = 1.0 + 1 ULP (LSB 1).
    // Halfway value (low 16 bits == 0x8000) rounds to even => 0x3F80.
    const tie_bits: u32 = 0x3F808000;
    const x: f32 = @bitCast(tie_bits);
    try testing.expectEqual(@as(u16, 0x3F80), f32ToBfloat16(x));

    // Tie point above an odd LSB: 0x3F818000. The neighbours are
    // 0x3F81 (odd) and 0x3F82 (even). Round to 0x3F82.
    const tie_bits_odd: u32 = 0x3F818000;
    const x2: f32 = @bitCast(tie_bits_odd);
    try testing.expectEqual(@as(u16, 0x3F82), f32ToBfloat16(x2));
}

test "f32ToBinary16: compiler lowering round-trip for representable values" {
    const exact = [_]f32{ 0.0, 1.0, -1.0, 0.5, -0.5 };
    for (exact) |x| {
        const b = f32ToBinary16(x);
        const back = binary16ToF32(b);
        try testing.expectEqual(x, back);
    }
}

test "quantizeF32ToI8Sym: zero input maps to zero scale + zero values" {
    var values: [4]i8 = undefined;
    const src = [_]f32{ 0, 0, 0, 0 };
    const q = quantizeF32ToI8Sym(&src, &values);
    try testing.expectEqual(@as(f32, 0), q.scale);
    for (values) |v| try testing.expectEqual(@as(i8, 0), v);
}

test "quantizeF32ToI8Sym: max-abs value lands at 127 (or -127)" {
    var values: [3]i8 = undefined;
    const src = [_]f32{ -1.0, 0.5, 1.0 };
    const q = quantizeF32ToI8Sym(&src, &values);
    try testing.expectEqual(@as(f32, 1.0), q.scale);
    try testing.expectEqual(@as(i8, -127), values[0]);
    try testing.expectEqual(@as(i8, 127), values[2]);
    // 0.5 / scale * 127 = 63.5, rounds half-away to 64.
    try testing.expectEqual(@as(i8, 64), values[1]);
}

test "quantize+dequantize sym: relative error bounded by ~1/127 ulp of scale" {
    var prng = std.Random.DefaultPrng.init(0x5A1F1ED);
    const rng = prng.random();
    const dim: usize = 128;
    var src: [dim]f32 = undefined;
    for (&src) |*v| v.* = rng.float(f32) * 2 - 1; // [-1, 1]
    var quant: [dim]i8 = undefined;
    const q = quantizeF32ToI8Sym(&src, &quant);
    var deq: [dim]f32 = undefined;
    dequantizeI8Sym(&quant, q.scale, &deq);
    for (src, deq) |s, d| {
        // Worst case error per value: scale/127 (half a quantization step).
        // Allow a small slack for floating-point rounding.
        try testing.expect(@abs(s - d) <= q.scale / 127.0 + 1e-6);
    }
}

test "quantize+dequantize asym: realistic embedding-like distribution recovers within tolerance" {
    // Slightly asymmetric distribution in the typical embedding
    // value range [-0.2, 0.8]. Ideal zero-point is in i8 range,
    // so no scale widening.
    const dim: usize = 128;
    var src: [dim]f32 = undefined;
    for (&src, 0..) |*v, i| {
        v.* = -0.2 + @as(f32, @floatFromInt(i)) / @as(f32, dim);
    }
    var quant: [dim]i8 = undefined;
    const q = quantizeF32ToI8Asym(&src, &quant);
    // Range ≈ 0.992 → ideal scale ≈ 0.496; zero ≈ -76.
    try testing.expect(@abs(q.scale - 0.5) < 0.01);
    try testing.expect(q.zero <= -70 and q.zero >= -85);
    var deq: [dim]f32 = undefined;
    dequantizeI8Asym(&quant, q.scale, q.zero, &deq);
    for (src, deq) |s, d| {
        try testing.expect(@abs(s - d) <= q.scale / 127.0 + 1e-5);
    }
}

test "quantize+dequantize asym: range-far-from-origin still encodes (zero clamped, scale widened)" {
    // Pathological case the design accepts with degraded recall:
    // values in [5.0, 6.0] force the zero-point ideal out of i8
    // range. The algorithm clamps zero and widens scale; the
    // round-trip still covers the input range, just with bigger
    // per-step error.
    const dim: usize = 64;
    var src: [dim]f32 = undefined;
    for (&src, 0..) |*v, i| {
        v.* = 5.0 + @as(f32, @floatFromInt(i)) / @as(f32, dim);
    }
    var quant: [dim]i8 = undefined;
    const q = quantizeF32ToI8Asym(&src, &quant);
    try testing.expectEqual(@as(i8, -127), q.zero);
    var deq: [dim]f32 = undefined;
    dequantizeI8Asym(&quant, q.scale, q.zero, &deq);
    for (src, deq) |s, d| {
        try testing.expect(@abs(s - d) <= q.scale / 127.0 + 1e-4);
    }
}

test "canonicalizeNumber: +0 and -0 both emit \"0\"" {
    var buf: [64]u8 = undefined;
    const a = try canonicalizeNumber("0", &buf);
    try testing.expectEqualStrings("0", a);
    const b = try canonicalizeNumber("-0", &buf);
    try testing.expectEqualStrings("0", b);
    const c = try canonicalizeNumber("0.0", &buf);
    try testing.expectEqualStrings("0", c);
    const d = try canonicalizeNumber("-0.000", &buf);
    try testing.expectEqualStrings("0", d);
}

test "canonicalizeNumber: representationally equivalent inputs canonicalize identically" {
    // The whole point: Excel writes 0.1, LibreOffice may write
    // 1.0000000000000001E-1, but both parse to the same f64 and
    // therefore must emit the same canonical bytes.
    var buf_a: [64]u8 = undefined;
    var buf_b: [64]u8 = undefined;
    const a = try canonicalizeNumber("0.1", &buf_a);
    const b = try canonicalizeNumber("1.0000000000000001E-1", &buf_b);
    try testing.expectEqualStrings(a, b);
}

test "canonicalizeNumber: NaN and inf rejected" {
    var buf: [64]u8 = undefined;
    try testing.expectError(Error.MalformedNumber, canonicalizeNumber("nan", &buf));
    try testing.expectError(Error.MalformedNumber, canonicalizeNumber("inf", &buf));
    try testing.expectError(Error.MalformedNumber, canonicalizeNumber("-inf", &buf));
}

test "canonicalizeNumber: comma-decimal rejected (locale-quirk guard)" {
    var buf: [64]u8 = undefined;
    try testing.expectError(Error.MalformedNumber, canonicalizeNumber("0,1", &buf));
    try testing.expectError(Error.MalformedNumber, canonicalizeNumber("1,5e3", &buf));
}

test "canonicalizeNumber: surrounding ASCII whitespace tolerated" {
    var buf: [64]u8 = undefined;
    const a = try canonicalizeNumber("  42  ", &buf);
    const b = try canonicalizeNumber("42", &buf);
    try testing.expectEqualStrings(a, b);
}

test "canonicalizeNumber: empty input rejected" {
    var buf: [64]u8 = undefined;
    try testing.expectError(Error.MalformedNumber, canonicalizeNumber("", &buf));
    try testing.expectError(Error.MalformedNumber, canonicalizeNumber("   ", &buf));
}

test "canonicalizeText: trims Unicode whitespace and NFC-normalizes" {
    const a = testing.allocator;
    var out: std.ArrayListUnmanaged(u8) = .empty;
    defer out.deinit(a);

    try canonicalizeText(a, "\u{00A0}Cafe\u{301}\u{2003}", &out);
    try testing.expectEqualStrings("Caf\u{00E9}", out.items);
}

test "canonicalizeText: invalid UTF-8 rejected" {
    const a = testing.allocator;
    var out: std.ArrayListUnmanaged(u8) = .empty;
    defer out.deinit(a);
    try testing.expectError(Error.InvalidUtf8, canonicalizeText(a, "\xFF", &out));
}

test "xxh3Canonical: equivalent text and number payloads hash identically" {
    const a = testing.allocator;
    var scratch_a: std.ArrayListUnmanaged(u8) = .empty;
    defer scratch_a.deinit(a);
    var scratch_b: std.ArrayListUnmanaged(u8) = .empty;
    defer scratch_b.deinit(a);

    const text_a = try xxh3Canonical(a, "worksheets/sheet1.xml", 2, .{ .string = " Cafe\u{301} " }, &scratch_a);
    const text_b = try xxh3Canonical(a, "worksheets/sheet1.xml", 2, .{ .string = "Caf\u{00E9}" }, &scratch_b);
    try testing.expectEqual(text_a, text_b);

    const num_a = try xxh3Canonical(a, "worksheets/sheet1.xml", 2, .{ .number = "0.1" }, &scratch_a);
    const num_b = try xxh3Canonical(a, "worksheets/sheet1.xml", 2, .{ .number = "1.0000000000000001E-1" }, &scratch_b);
    try testing.expectEqual(num_a, num_b);
}

test "xxh3Canonical: worksheet target and row are part of the hash" {
    const a = testing.allocator;
    var scratch_a: std.ArrayListUnmanaged(u8) = .empty;
    defer scratch_a.deinit(a);
    var scratch_b: std.ArrayListUnmanaged(u8) = .empty;
    defer scratch_b.deinit(a);

    const row_2 = try xxh3Canonical(a, "worksheets/sheet1.xml", 2, .{ .boolean = true }, &scratch_a);
    const row_3 = try xxh3Canonical(a, "worksheets/sheet1.xml", 3, .{ .boolean = true }, &scratch_b);
    try testing.expect(row_2 != row_3);

    const sheet_1 = try xxh3Canonical(a, "worksheets/sheet1.xml", 2, .blank, &scratch_a);
    const sheet_2 = try xxh3Canonical(a, "worksheets/sheet2.xml", 2, .blank, &scratch_b);
    try testing.expect(sheet_1 != sheet_2);
}

fn allocationFailureCanonicalHash(allocator: Allocator) !void {
    var scratch: std.ArrayListUnmanaged(u8) = .empty;
    defer scratch.deinit(allocator);
    _ = try xxh3Canonical(
        allocator,
        "worksheets/sheet1.xml",
        2,
        .{ .string = "\u{00A0}Cafe\u{301}\u{2003}" },
        &scratch,
    );
}

test "xxh3Canonical: allocation failures clean up and propagate" {
    try testing.checkAllAllocationFailures(
        testing.allocator,
        allocationFailureCanonicalHash,
        .{},
    );
}

test "xxh3: matches stdlib XxHash3 with seed 0 (regression pin)" {
    const a = xxh3("");
    const b = std.hash.XxHash3.hash(0, "");
    try testing.expectEqual(a, b);
    const c = xxh3("the quick brown fox");
    const d = std.hash.XxHash3.hash(0, "the quick brown fox");
    try testing.expectEqual(c, d);
}

test "TOMBSTONE_HASH constant equals u64::MAX" {
    try testing.expectEqual(std.math.maxInt(u64), TOMBSTONE_HASH);
}

test "encodeVecPart: header + body round-trip through parseVecPart" {
    const a = testing.allocator;
    const dim: u32 = 2;
    const count: u32 = 3;
    var body: [3 * 2 * 4]u8 = undefined;
    @memset(&body, 0);
    // f32 record bytes: 4 per scalar. Write deterministic values.
    for (0..count) |row| {
        for (0..dim) |col| {
            const off = (row * dim + col) * 4;
            const v: f32 = @floatFromInt(row * 10 + col);
            std.mem.writeInt(u32, body[off..][0..4], @bitCast(v), .little);
        }
    }
    const part = try encodeVecPart(a, .{ .version = WIRE_VERSION, .dim = dim, .count = count, .dtype = .f32 }, &body);
    defer a.free(part);

    const parsed = try parseVecPart(part);
    try testing.expectEqual(@as(u32, dim), parsed.header.dim);
    try testing.expectEqual(@as(u32, count), parsed.header.count);
    try testing.expectEqual(Dtype.f32, parsed.header.dtype);
    try testing.expectEqualSlices(u8, &body, parsed.body);
}

test "encodeVecPart: body size mismatch rejected" {
    const a = testing.allocator;
    var bad_body: [4]u8 = .{0} ** 4; // 1×f32 — but header says count=2, dim=1 → expects 8
    try testing.expectError(
        Error.BodySizeMismatch,
        encodeVecPart(a, .{ .version = WIRE_VERSION, .dim = 1, .count = 2, .dtype = .f32 }, &bad_body),
    );
}

test "encodeHashesPart: round-trips through parseHashPart" {
    const a = testing.allocator;
    const hashes = [_]u64{ 0xDEADBEEF, TOMBSTONE_HASH, 0xCAFEBABE };
    const part = try encodeHashesPart(a, &hashes);
    defer a.free(part);

    const parsed = try parseHashPart(part);
    try testing.expectEqual(@as(u32, 3), parsed.header.count);
    for (hashes, 0..) |want, i| {
        try testing.expectEqual(want, try parsed.value(@intCast(i)));
    }
}

test "encodeIndexXml: round-trips through parseIndex with one coverage" {
    const a = testing.allocator;
    const cov_spec = [_]CoverageSpec{.{
        .id = "title",
        .worksheet_target = "worksheets/sheet1.xml",
        .range = "A2:A100",
        .column = "A",
        .count = 99,
        .include_formulas = false,
        .vec_rid = "rId1",
        .hash_rid = "rId2",
    }};
    const xml = try encodeIndexXml(a, .{
        .model = "text-embedding-3-small",
        .dim = 1536,
        .dtype = .int8_sym_per_vec,
        .hash_algo = HASH_ALGO_XXH3_64,
        .coverages = &cov_spec,
    });
    defer a.free(xml);

    var cov_buf: [4]Coverage = undefined;
    const idx = try parseIndex(xml, &cov_buf);
    try testing.expectEqualStrings("text-embedding-3-small", idx.model);
    try testing.expectEqual(@as(u32, 1536), idx.dim);
    try testing.expectEqual(Dtype.int8_sym_per_vec, idx.dtype);
    try testing.expectEqual(@as(usize, 1), idx.coverages.len);

    const c = idx.coverages[0];
    try testing.expectEqualStrings("title", c.id);
    try testing.expectEqualStrings("worksheets/sheet1.xml", c.worksheet_target);
    try testing.expectEqualStrings("rId1", c.vec_rid);
    try testing.expectEqualStrings("rId2", c.hash_rid);
    try testing.expectEqual(@as(u32, 99), c.count);
    try testing.expect(!c.include_formulas);
}

test "encodeIndexXml: XML-escapes special characters in model name" {
    const a = testing.allocator;
    const xml = try encodeIndexXml(a, .{
        .model = "model<test>&\"name\"",
        .dim = 4,
        .dtype = .f32,
        .hash_algo = HASH_ALGO_XXH3_64,
        .coverages = &.{},
    });
    defer a.free(xml);
    // The four special characters must not appear unescaped inside
    // the attribute value (the XML declaration's `?>` is OK).
    try testing.expect(std.mem.indexOf(u8, xml, "&lt;test&gt;") != null);
    try testing.expect(std.mem.indexOf(u8, xml, "&amp;") != null);
    try testing.expect(std.mem.indexOf(u8, xml, "&quot;name&quot;") != null);
    try testing.expect(std.mem.indexOf(u8, xml, "model<test>") == null);
}

test "encodeIndexRelsXml: round-trips through parseIndexRelationships" {
    const a = testing.allocator;
    const rels = [_]RelSpec{
        .{ .id = "rId1", .type = REL_TYPE_VEC, .target = "title/vec.bin" },
        .{ .id = "rId2", .type = REL_TYPE_HASH, .target = "title/hashes.bin" },
    };
    const xml = try encodeIndexRelsXml(a, &rels);
    defer a.free(xml);

    var rel_buf: [4]IndexRelationship = undefined;
    const parsed = try parseIndexRelationships(xml, &rel_buf);
    try testing.expectEqual(@as(usize, 2), parsed.len);
    try testing.expectEqualStrings("rId1", parsed[0].id);
    try testing.expectEqualStrings("title/vec.bin", parsed[0].target);
    try testing.expectEqual(TargetMode.internal, parsed[0].target_mode);
    try testing.expectEqualStrings("rId2", parsed[1].id);
    try testing.expectEqualStrings("title/hashes.bin", parsed[1].target);
}

test "decodeAllF32 round-trips int8-sym across every row" {
    // Two rows, dim 3, distinct scales so a row-indexing bug shows up
    // as a magnitude error rather than a subtle sign flip.
    const dim: u32 = 3;
    var body: [2 * (4 + 3)]u8 = undefined;
    std.mem.writeInt(u32, body[0..4], @bitCast(@as(f32, 1.27)), .little);
    body[4] = @bitCast(@as(i8, 127));
    body[5] = @bitCast(@as(i8, -127));
    body[6] = @bitCast(@as(i8, 0));
    std.mem.writeInt(u32, body[7..11], @bitCast(@as(f32, 2.54)), .little);
    body[11] = @bitCast(@as(i8, 127));
    body[12] = @bitCast(@as(i8, 0));
    body[13] = @bitCast(@as(i8, -127));

    const vec: ParsedVecPart = .{
        .header = .{ .version = 1, .dtype = .int8_sym_per_vec, .dim = dim, .count = 2 },
        .body = &body,
    };
    var out: [6]f32 = undefined;
    try decodeAllF32(vec, &out);

    try testing.expectApproxEqAbs(@as(f32, 1.27), out[0], 0.001);
    try testing.expectApproxEqAbs(@as(f32, -1.27), out[1], 0.001);
    try testing.expectApproxEqAbs(@as(f32, 0.0), out[2], 0.001);
    try testing.expectApproxEqAbs(@as(f32, 2.54), out[3], 0.001);
    try testing.expectApproxEqAbs(@as(f32, 0.0), out[4], 0.001);
    try testing.expectApproxEqAbs(@as(f32, -2.54), out[5], 0.001);
}

test "decodeAllF32 handles plain f32 bodies" {
    const dim: u32 = 2;
    var body: [2 * 2 * 4]u8 = undefined;
    const vals = [_]f32{ 1.5, -2.25, 0.0, 3.75 };
    for (vals, 0..) |v, i| std.mem.writeInt(u32, body[i * 4 ..][0..4], @bitCast(v), .little);
    const vec: ParsedVecPart = .{
        .header = .{ .version = 1, .dtype = .f32, .dim = dim, .count = 2 },
        .body = &body,
    };
    var out: [4]f32 = undefined;
    try decodeAllF32(vec, &out);
    for (vals, out) |want, got| try testing.expectEqual(want, got);
}

test "decodeAllF32 rejects a mis-sized output buffer" {
    var body: [4 + 2]u8 = @splat(0);
    const vec: ParsedVecPart = .{
        .header = .{ .version = 1, .dtype = .int8_sym_per_vec, .dim = 2, .count = 1 },
        .body = &body,
    };
    var too_small: [1]f32 = undefined;
    try testing.expectError(Error.InvalidRange, decodeAllF32(vec, &too_small));
}

test "encodeVectorRecord / encodeVectorBody: f32 and int8-sym bodies read back through decodeAllF32" {
    const a = testing.allocator;
    // 3 rows × 4: a plain row, the zero row (scale 0), a wide row.
    const vectors = [_]f32{ 1.0, -2.5, 0.25, 4.0, 0, 0, 0, 0, 100, -100, 50, 0.5 };

    const f32_body = try encodeVectorBody(a, .f32, 4, &vectors);
    defer a.free(f32_body);
    try testing.expectEqual(@as(usize, 48), f32_body.len);
    const f32_part = try encodeVecPart(a, .{ .version = WIRE_VERSION, .dim = 4, .count = 3, .dtype = .f32 }, f32_body);
    defer a.free(f32_part);
    var out: [12]f32 = undefined;
    try decodeAllF32(try parseVecPart(f32_part), &out);
    try testing.expectEqualSlices(f32, &vectors, &out);

    const i8_body = try encodeVectorBody(a, .int8_sym_per_vec, 4, &vectors);
    defer a.free(i8_body);
    try testing.expectEqual(@as(usize, 3 * (4 + 4)), i8_body.len);
    const i8_part = try encodeVecPart(a, .{ .version = WIRE_VERSION, .dim = 4, .count = 3, .dtype = .int8_sym_per_vec }, i8_body);
    defer a.free(i8_part);
    try decodeAllF32(try parseVecPart(i8_part), &out);
    for (vectors, out, 0..) |want, got, i| {
        const row_max: f32 = switch (i / 4) {
            0 => 4.0,
            1 => 0.0,
            else => 100.0,
        };
        try testing.expect(@abs(want - got) <= row_max / 127.0 + 1e-6);
    }
    try testing.expectEqual(@as(f32, 0), out[4]);
    // The on-disk record: the per-row f32 scale, then the codes.
    try testing.expectEqual(@as(f32, 4.0), @as(f32, @bitCast(std.mem.readInt(u32, i8_body[0..4], .little))));
    try testing.expectEqual(@as(i8, 127), @as(i8, @bitCast(i8_body[4 + 3])));
    try testing.expectEqual(@as(f32, 0), @as(f32, @bitCast(std.mem.readInt(u32, i8_body[8..12], .little))));
}

test "encodeVectorRecord / encodeVectorBody: the call's shape is checked before any byte moves" {
    const a = testing.allocator;
    const v = [_]f32{ 1, 2, 3 };
    var dst: [12]u8 = [_]u8{0xAA} ** 12;
    try testing.expectError(Error.DimensionOutOfRange, encodeVectorRecord(.f32, 0, &v, &dst));
    try testing.expectError(Error.DimensionMismatch, encodeVectorRecord(.f32, 2, &v, dst[0..8]));
    try testing.expectError(Error.BodySizeMismatch, encodeVectorRecord(.f32, 3, &v, dst[0..8]));
    try testing.expectError(Error.UnsupportedDtype, encodeVectorRecord(.binary16, 3, &v, dst[0..6]));
    try testing.expectError(Error.UnsupportedDtype, encodeVectorRecord(.bfloat16, 3, &v, dst[0..6]));
    try testing.expectError(Error.UnsupportedDtype, encodeVectorRecord(.int8_asym_per_vec, 3, &v, dst[0..8]));
    try testing.expectEqualSlices(u8, &([_]u8{0xAA} ** 12), &dst);
    try testing.expectError(Error.DimensionOutOfRange, encodeVectorBody(a, .f32, 0, &v));
    try testing.expectError(Error.DimensionMismatch, encodeVectorBody(a, .f32, 2, &v));
    try testing.expectError(Error.UnsupportedDtype, encodeVectorBody(a, .binary16, 3, &v));
    const empty = try encodeVectorBody(a, .f32, 3, &.{});
    defer a.free(empty);
    try testing.expectEqual(@as(usize, 0), empty.len);
}

test "validateCoverageColumn: the index read's rule, one definition" {
    const range = try parseA1Range("B2:D9");
    try testing.expectEqual(@as(u32, 1), try validateCoverageColumn("B", range));
    try testing.expectEqual(@as(u32, 3), try validateCoverageColumn("d", range));
    try testing.expectError(Error.InvalidRange, validateCoverageColumn("A", range));
    try testing.expectError(Error.InvalidRange, validateCoverageColumn("E", range));
    try testing.expectError(Error.InvalidRange, validateCoverageColumn("B2", range));
    try testing.expectError(Error.InvalidRange, validateCoverageColumn("", range));
}

test "validateMetadataText + appendXmlEscaped: XML 1.0's forbidden bytes refuse, its characters pass" {
    const a = testing.allocator;
    try validateMetadataText("model-v1 \t\n\r\x7f <&>\"'");
    var c: u8 = 0;
    while (true) : (c += 1) {
        const s = [_]u8{c};
        if (sheet_plan.isForbiddenXmlByte(c)) {
            try testing.expectError(Error.InvalidXmlByte, validateMetadataText(&s));
        } else {
            try validateMetadataText(&s);
        }
        if (c == 0x7f) break;
    }
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(a);
    try testing.expectError(Error.InvalidXmlByte, appendXmlEscaped(&buf, a, "m\x00"));
    buf.clearRetainingCapacity();
    try appendXmlEscaped(&buf, a, "a<b>&\"\t\x7f");
    try testing.expectEqualStrings("a&lt;b&gt;&amp;&quot;\t\x7f", buf.items);
    // The index encoder is the one carrier: a control byte in the
    // model refuses the whole part, so no index ever holds one.
    try testing.expectError(Error.InvalidXmlByte, encodeIndexXml(a, .{
        .model = "m\x01",
        .dim = 1,
        .dtype = .f32,
        .hash_algo = HASH_ALGO_XXH3_64,
        .coverages = &.{},
    }));
}
