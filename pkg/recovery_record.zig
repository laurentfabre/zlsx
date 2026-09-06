//! ER — the embedding recovery record.
//!
//! **What this is for.** emb-4 measured that Apple Numbers and
//! LibreOffice Calc rebuild the .xlsx archive on save and drop every
//! `xl/zlsxEmbeddings/*` part. The vectors are *removed*, not merely
//! unreferenced, so no reader-side scan gets them back. Without
//! something else in the package, a workbook that went through one of
//! those tools is indistinguishable from one that never had embeddings
//! — the caller silently gets nothing.
//!
//! emb-4B measured which carriers *do* survive those rebuilds. This
//! module encodes ~200 bytes of provenance into one of them, so the
//! stripped state becomes detectable and attributable. Per
//! `docs/plans/embeddings-in-xlsx.md` §Durability contract:
//!
//! > zlsx guarantees that a workbook which loses its vectors **says
//! > so**. It does not guarantee that every tool keeps the vectors.
//!
//! **This is not a backup.** The record cannot reconstruct the
//! vectors — 200 bytes against megabytes. It records *what was there*
//! so a caller can re-embed deliberately, with the original model and
//! ranges, instead of silently getting an empty result.
//!
//! ## Wire format
//!
//! One line, pipe-separated, ASCII only:
//!
//! ```text
//! zlsxER1|<model>|<dim>|<dtype>|<hash_algo>|<ncov>|<id>|<ws>|<range>|<count>|…|<digest>
//! ```
//!
//! Coverage fields repeat in groups of four, `ncov` times. `digest` is
//! 16 lowercase hex chars.
//!
//! Every variable-length text field is **percent-encoded**: any byte
//! outside `[A-Za-z0-9._-]` becomes `%XX`. That is stricter than
//! necessary for the separator alone, and deliberately so — it also
//! guarantees the payload survives:
//!
//!   - XML escaping (no `<`, `>`, `&`, `"` can appear),
//!   - Excel formula-string quoting (no `"` to double),
//!   - any consumer that re-encodes text (pure ASCII, no UTF-8
//!     sequences to mangle).
//!
//! ## Reading it back is the hard part
//!
//! The carriers are rewritten by the tools this record exists to
//! survive. Round-tripping a hidden `<definedName>` through
//! LibreOffice showed three normalizations, each of which silently
//! loses the record if the parser matches its own writer's bytes:
//!
//!   1. `hidden="1"` comes back as `hidden="true"` — both are valid
//!      `xsd:boolean` lexical forms.
//!   2. The payload comes back XML-escaped: `"x"` → `&quot;x&quot;`.
//!   3. The element gains attributes (`function="false"`,
//!      `vbProcedure="false"`).
//!
//! So: match on `name=`, never the whole tag; accept either boolean
//! form; unescape before parsing. `findDefinedNameValue` does all
//! three.

const std = @import("std");
const assert = std.debug.assert;
// The typed parser's scanner, for the ONE acceptance of what is not an
// element (comment / CDATA / PI / DOCTYPE): the lexical reader below
// must skip exactly what the parser and the strip skip. The module
// imports nothing but std, so this leaf stays a leaf.
const workbook_xml_mod = @import("typed_parts/workbook_xml.zig");

pub const Error = error{
    MalformedRecord,
    UnsupportedRecordVersion,
    BufferTooSmall,
    RecordTooLarge,
} || std.mem.Allocator.Error;

/// Format tag. Bump on any incompatible change to the field order or
/// separator scheme; readers reject a tag they do not know rather than
/// guessing, because a misparsed provenance record is worse than an
/// absent one — it reads as authoritative.
pub const MAGIC: []const u8 = "zlsxER1";

/// Base name for the chunk carrier. Excel caps a defined-name formula
/// near 255 chars, so a record longer than one chunk splits across
/// `_zlsxRecovery0`, `_zlsxRecovery1`, … and the reader concatenates
/// in index order.
///
/// The leading underscore matches Excel's own convention for
/// machine-owned names (`_xlnm.Print_Area`).
pub const NAME_PREFIX: []const u8 = "_zlsxRecovery";

/// Payload bytes per chunk. Conservative against the ~255-char
/// practical limit, leaving room for the surrounding quotes and for a
/// consumer that expands the text on re-encode.
pub const MAX_CHUNK: usize = 200;

/// Hard ceiling on chunk count. A record needing more than this is a
/// workbook with hundreds of coverages; `encodeChunked` refuses rather
/// than silently dropping coverages, because a truncated record claims
/// completeness it does not have.
pub const MAX_CHUNKS: usize = 16;

/// Custom document property name for the secondary carrier.
pub const DOC_PROP_NAME: []const u8 = "ZlsxEmbeddingRecovery";

/// Whether `name` is one of the record's chunk names: the exact
/// `_zlsxRecovery<digits>` form, compared case-insensitively because a
/// defined name's identity is (Excel folds case). A user's
/// `_zlsxRecoveryMine` is NOT ours — the strip used to delete any name
/// carrying the prefix, and a foreign tool's `_ZLSXRECOVERY0` used to
/// survive beside ours (in-house r2 S3C-REL-203).
pub fn isChunkName(name: []const u8) bool {
    if (name.len <= NAME_PREFIX.len) return false;
    if (!std.ascii.eqlIgnoreCase(name[0..NAME_PREFIX.len], NAME_PREFIX)) return false;
    for (name[NAME_PREFIX.len..]) |c| if (!std.ascii.isDigit(c)) return false;
    return true;
}

test "isChunkName: the exact chunk form, case-insensitive; a user's prefixed name is not ours" {
    try testing.expect(isChunkName("_zlsxRecovery0"));
    try testing.expect(isChunkName("_zlsxRecovery15"));
    try testing.expect(isChunkName("_ZLSXRECOVERY0"));
    try testing.expect(isChunkName("_ZlsxRecovery7"));
    try testing.expect(!isChunkName("_zlsxRecovery"));
    try testing.expect(!isChunkName("_zlsxRecoveryMine"));
    try testing.expect(!isChunkName("_zlsxRecovery0x"));
    try testing.expect(!isChunkName("zlsxRecovery0"));
    try testing.expect(!isChunkName(""));
}

/// Sheet that carries the record when the cell carrier is enabled.
///
/// Opt-in, because it is the one carrier a user can see: Sheet ▸ Unhide
/// reveals it. It is also the ONLY carrier Apple Numbers preserves —
/// measured 2026-07-27, Numbers 15.3 strips the other five. That is not
/// an implementation gap: Numbers rebuilds the file from its own
/// document model, so exactly what that model represents survives, and
/// everything invisible to the user is outside it. Invisibility and
/// Numbers-durability are mutually exclusive by construction.
pub const CELL_SHEET_NAME: []const u8 = "zlsxRecovery";

pub const RecoveredCoverage = struct {
    id: []const u8,
    worksheet_target: []const u8,
    range: []const u8,
    count: u32,
};

/// Which carrier a record was recovered from. Reported so a caller
/// (and the compat matrix) can tell which survived a given consumer.
pub const Carrier = enum { defined_name, doc_props, cell_data };

pub const RecoveryRecord = struct {
    model: []const u8,
    dim: u32,
    dtype: []const u8,
    hash_algo: []const u8,
    coverages: []const RecoveredCoverage,
    /// Content fingerprint at embed time: a fold over every coverage's
    /// per-row content hashes, in coverage order.
    ///
    /// It is **not** recomputable from the stripped file — the hashes
    /// went with the vectors. It is recomputable from the *current
    /// cells* via the same canonicalization, which is what makes it
    /// useful: equal means the covered content has not drifted since
    /// embedding, so a re-embed reproduces the same vectors; unequal
    /// means the content changed too.
    digest: u64,
    carrier: Carrier,
};

/// Input to the encoder. Mirrors `embedding_part.Index` but takes
/// plain slices so the encoder does not depend on the index types.
pub const RecordInput = struct {
    model: []const u8,
    dim: u32,
    dtype: []const u8,
    hash_algo: []const u8,
    coverages: []const RecoveredCoverage,
    digest: u64,
};

// ─── percent codec ──────────────────────────────────────────────────

fn isUnreserved(c: u8) bool {
    return (c >= 'A' and c <= 'Z') or
        (c >= 'a' and c <= 'z') or
        (c >= '0' and c <= '9') or
        c == '.' or c == '_' or c == '-';
}

fn hexDigit(nibble: u4) u8 {
    return "0123456789ABCDEF"[nibble];
}

fn unhex(c: u8) ?u8 {
    return switch (c) {
        '0'...'9' => c - '0',
        'A'...'F' => c - 'A' + 10,
        'a'...'f' => c - 'a' + 10,
        else => null,
    };
}

fn percentEncode(
    allocator: std.mem.Allocator,
    out: *std.ArrayListUnmanaged(u8),
    s: []const u8,
) !void {
    for (s) |c| {
        if (isUnreserved(c)) {
            try out.append(allocator, c);
        } else {
            try out.append(allocator, '%');
            try out.append(allocator, hexDigit(@intCast(c >> 4)));
            try out.append(allocator, hexDigit(@intCast(c & 0x0F)));
        }
    }
}

/// Decode in place into `buf`, returning the used prefix. Caller sizes
/// `buf` at least `s.len` — percent-decoding never grows.
fn percentDecode(buf: []u8, s: []const u8) Error![]const u8 {
    assert(buf.len >= s.len);
    var n: usize = 0;
    var i: usize = 0;
    while (i < s.len) {
        if (s[i] == '%') {
            if (i + 2 >= s.len) return Error.MalformedRecord;
            const hi = unhex(s[i + 1]) orelse return Error.MalformedRecord;
            const lo = unhex(s[i + 2]) orelse return Error.MalformedRecord;
            buf[n] = (@as(u8, hi) << 4) | lo;
            i += 3;
        } else {
            buf[n] = s[i];
            i += 1;
        }
        n += 1;
    }
    return buf[0..n];
}

// ─── encode ─────────────────────────────────────────────────────────

/// Serialize a record. Caller owns the returned slice.
pub fn encode(allocator: std.mem.Allocator, in: RecordInput) Error![]u8 {
    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(allocator);

    try out.appendSlice(allocator, MAGIC);
    try out.append(allocator, '|');
    try percentEncode(allocator, &out, in.model);
    try out.print(allocator, "|{d}|", .{in.dim});
    try percentEncode(allocator, &out, in.dtype);
    try out.append(allocator, '|');
    try percentEncode(allocator, &out, in.hash_algo);
    try out.print(allocator, "|{d}", .{in.coverages.len});

    for (in.coverages) |c| {
        try out.append(allocator, '|');
        try percentEncode(allocator, &out, c.id);
        try out.append(allocator, '|');
        try percentEncode(allocator, &out, c.worksheet_target);
        try out.append(allocator, '|');
        try percentEncode(allocator, &out, c.range);
        try out.print(allocator, "|{d}", .{c.count});
    }

    try out.print(allocator, "|{x:0>16}", .{in.digest});
    return out.toOwnedSlice(allocator);
}

/// Number of chunks `record` splits into.
pub fn chunkCount(record_len: usize) usize {
    if (record_len == 0) return 0;
    return (record_len + MAX_CHUNK - 1) / MAX_CHUNK;
}

/// Byte range of chunk `i`. Chunks are split on a fixed byte boundary,
/// not on field boundaries — the reader concatenates before parsing,
/// so a percent-escape may straddle a chunk and that is fine.
pub fn chunk(record: []const u8, i: usize) []const u8 {
    const start = i * MAX_CHUNK;
    if (start >= record.len) return record[record.len..];
    const end = @min(start + MAX_CHUNK, record.len);
    return record[start..end];
}

/// Name of chunk `i`, written into `buf`.
pub fn chunkName(buf: []u8, i: usize) []const u8 {
    return std.fmt.bufPrint(buf, "{s}{d}", .{ NAME_PREFIX, i }) catch unreachable;
}

// ─── decode ─────────────────────────────────────────────────────────

const FieldIter = struct {
    s: []const u8,
    pos: usize = 0,

    fn next(self: *FieldIter) ?[]const u8 {
        if (self.pos > self.s.len) return null;
        const start = self.pos;
        const rel = std.mem.indexOfScalarPos(u8, self.s, start, '|');
        if (rel) |p| {
            self.pos = p + 1;
            return self.s[start..p];
        }
        self.pos = self.s.len + 1;
        return self.s[start..];
    }
};

fn nextDecoded(it: *FieldIter, scratch: []u8) Error![]const u8 {
    const raw = it.next() orelse return Error.MalformedRecord;
    if (raw.len > scratch.len) return Error.BufferTooSmall;
    return percentDecode(scratch, raw);
}

fn nextInt(comptime T: type, it: *FieldIter) Error!T {
    const raw = it.next() orelse return Error.MalformedRecord;
    return std.fmt.parseInt(T, raw, 10) catch Error.MalformedRecord;
}

/// Parse a record.
///
/// `coverage_storage` receives the coverage entries; `text_scratch`
/// backs every decoded string. Both are borrowed by the result, so
/// they must outlive it. `BufferTooSmall` means one of them was too
/// small — grow and retry, the same contract `embedding_part.parseIndex`
/// uses.
pub fn decode(
    bytes: []const u8,
    carrier: Carrier,
    coverage_storage: []RecoveredCoverage,
    text_scratch: []u8,
) Error!RecoveryRecord {
    var it: FieldIter = .{ .s = bytes };

    const magic = it.next() orelse return Error.MalformedRecord;
    if (!std.mem.eql(u8, magic, MAGIC)) return Error.UnsupportedRecordVersion;

    // Hand out disjoint windows of the scratch buffer so every decoded
    // string stays live for the caller.
    var used: usize = 0;
    const take = struct {
        fn f(scratch: []u8, used_ptr: *usize, want: usize) Error![]u8 {
            if (used_ptr.* + want > scratch.len) return Error.BufferTooSmall;
            const s = scratch[used_ptr.* .. used_ptr.* + want];
            used_ptr.* += want;
            return s;
        }
    }.f;

    const model_raw = it.next() orelse return Error.MalformedRecord;
    const model = try percentDecode(try take(text_scratch, &used, model_raw.len), model_raw);

    const dim = try nextInt(u32, &it);

    const dtype_raw = it.next() orelse return Error.MalformedRecord;
    const dtype = try percentDecode(try take(text_scratch, &used, dtype_raw.len), dtype_raw);

    const algo_raw = it.next() orelse return Error.MalformedRecord;
    const hash_algo = try percentDecode(try take(text_scratch, &used, algo_raw.len), algo_raw);

    const ncov = try nextInt(u32, &it);
    if (ncov > coverage_storage.len) return Error.BufferTooSmall;

    var i: usize = 0;
    while (i < ncov) : (i += 1) {
        const id_raw = it.next() orelse return Error.MalformedRecord;
        const ws_raw = it.next() orelse return Error.MalformedRecord;
        const rg_raw = it.next() orelse return Error.MalformedRecord;
        coverage_storage[i] = .{
            .id = try percentDecode(try take(text_scratch, &used, id_raw.len), id_raw),
            .worksheet_target = try percentDecode(try take(text_scratch, &used, ws_raw.len), ws_raw),
            .range = try percentDecode(try take(text_scratch, &used, rg_raw.len), rg_raw),
            .count = try nextInt(u32, &it),
        };
    }

    const digest_raw = it.next() orelse return Error.MalformedRecord;
    const digest = std.fmt.parseInt(u64, digest_raw, 16) catch return Error.MalformedRecord;

    return .{
        .model = model,
        .dim = dim,
        .dtype = dtype,
        .hash_algo = hash_algo,
        .coverages = coverage_storage[0..ncov],
        .digest = digest,
        .carrier = carrier,
    };
}

// ─── carrier extraction ─────────────────────────────────────────────

/// XML-unescape the five predefined entities into `buf`.
///
/// Needed because LibreOffice returns the payload escaped
/// (`"x"` → `&quot;x&quot;`). Unknown entities pass through verbatim
/// rather than erroring: the payload is percent-encoded ASCII, so a
/// stray `&` is a consumer quirk, not a reason to discard a record we
/// can still read.
pub fn xmlUnescape(buf: []u8, s: []const u8) []const u8 {
    assert(buf.len >= s.len);
    const entities = [_]struct { name: []const u8, ch: u8 }{
        .{ .name = "&amp;", .ch = '&' },
        .{ .name = "&lt;", .ch = '<' },
        .{ .name = "&gt;", .ch = '>' },
        .{ .name = "&quot;", .ch = '"' },
        .{ .name = "&apos;", .ch = '\'' },
    };
    var n: usize = 0;
    var i: usize = 0;
    outer: while (i < s.len) {
        if (s[i] == '&') {
            for (entities) |e| {
                if (std.mem.startsWith(u8, s[i..], e.name)) {
                    buf[n] = e.ch;
                    n += 1;
                    i += e.name.len;
                    continue :outer;
                }
            }
        }
        buf[n] = s[i];
        n += 1;
        i += 1;
    }
    return buf[0..n];
}

/// Extract the inner text of `<definedName name="<wanted>" …>TEXT</definedName>`
/// from `xl/workbook.xml`, with the surrounding quotes of the string
/// literal stripped and XML entities resolved.
///
/// Matches on the `name=` attribute rather than the whole tag, because
/// consumers reorder and inject attributes — LibreOffice adds
/// `function="false"` and `vbProcedure="false"`, and rewrites
/// `hidden="1"` to `hidden="true"`. Matching the tag shape would lose
/// the record on every one of those.
pub fn findDefinedNameValue(
    workbook_xml: []const u8,
    wanted: []const u8,
    buf: []u8,
) ?[]const u8 {
    var search: usize = 0;
    while (std.mem.indexOfScalarPos(u8, workbook_xml, search, '<')) |lt| {
        // A `<definedName` inside a comment, a CDATA section, a
        // processing instruction or a DOCTYPE is text, not an element:
        // the construct is skipped whole, with the typed parser's own
        // scanner, so this reader cannot see the part differently from
        // the parser and the strip (a decoy chunk in a comment used to
        // be read as the record — in-house r4 S3C-REL-401; the other
        // constructs, r5 S3C-REL-502). An unterminated construct is not
        // a part to read a record from.
        const past = workbook_xml_mod.skipNonElement(workbook_xml, lt) catch return null;
        if (past != lt) {
            search = past;
            continue;
        }
        if (!std.mem.startsWith(u8, workbook_xml[lt..], "<definedName")) {
            search = lt + 1;
            continue;
        }
        const open = lt;
        const gt = std.mem.indexOfScalarPos(u8, workbook_xml, open, '>') orelse return null;
        const attrs = workbook_xml[open..gt];
        search = gt + 1;

        if (!definedNameMatches(attrs, wanted)) continue;
        // Self-closing name has no body, so no payload.
        if (gt > open and workbook_xml[gt - 1] == '/') continue;

        const close = std.mem.indexOfPos(u8, workbook_xml, gt + 1, "</definedName>") orelse return null;
        const body = workbook_xml[gt + 1 .. close];
        if (body.len > buf.len) return null;

        const unescaped = xmlUnescape(buf, body);
        return stripQuotes(unescaped);
    }
    return null;
}

/// True when `attrs` carries `name="<wanted>"` exactly.
///
/// Scans for the attribute rather than a fixed offset so attribute
/// order does not matter. Guards the preceding byte so `codeName=` and
/// `localSheetName=` cannot match `name=`.
fn definedNameMatches(attrs: []const u8, wanted: []const u8) bool {
    var i: usize = 0;
    while (std.mem.indexOfPos(u8, attrs, i, "name=")) |at| {
        i = at + "name=".len;
        const prev = attrs[at - 1];
        if (at > 0 and prev != ' ' and prev != '\t' and prev != '\n' and prev != '\r') continue;
        if (i >= attrs.len) return false;
        const q = attrs[i];
        if (q != '"' and q != '\'') continue;
        const end = std.mem.indexOfScalarPos(u8, attrs, i + 1, q) orelse return false;
        if (std.mem.eql(u8, attrs[i + 1 .. end], wanted)) return true;
    }
    return false;
}

/// Drop the wrapping quotes of an Excel string-literal formula.
/// Tolerates their absence — a consumer that normalizes the formula
/// may drop them, and the payload is unambiguous either way.
fn stripQuotes(s: []const u8) []const u8 {
    if (s.len >= 2 and s[0] == '"' and s[s.len - 1] == '"') return s[1 .. s.len - 1];
    return s;
}

/// Where a record sits in a text blob: `[start, end)`, from its magic
/// prefix to the first byte a percent-encoded record cannot contain.
pub const TextSpan = struct { start: usize, end: usize };

/// Locate a record anywhere in a text blob by its magic prefix — the
/// one locator the cell-carrier reader (`findRecordInText`) and the
/// strip's scrub share, so what one finds the other blanks.
///
/// The cell carrier's payload may land inline in the sheet XML or be
/// lifted into `sharedStrings.xml` depending on the consumer's
/// shared-string policy — so both scan rather than assume a location.
/// The record is percent-encoded ASCII, so it terminates at the first
/// byte that cannot appear in one.
pub fn recordSpanInText(text: []const u8) ?TextSpan {
    const at = std.mem.indexOf(u8, text, MAGIC) orelse return null;
    var end = at;
    while (end < text.len) : (end += 1) {
        const c = text[end];
        const ok = isUnreserved(c) or c == '%' or c == '|';
        if (!ok) break;
    }
    return .{ .start = at, .end = end };
}

/// Find a record anywhere in a text blob by its magic prefix — see
/// `recordSpanInText`.
pub fn findRecordInText(text: []const u8, buf: []u8) ?[]const u8 {
    const span = recordSpanInText(text) orelse return null;
    const raw = text[span.start..span.end];
    if (raw.len > buf.len) return null;
    return xmlUnescape(buf, raw);
}

/// Extract the record from `docProps/custom.xml` — the secondary
/// carrier. Same normalization tolerance as the defined-name path.
pub fn findDocPropValue(
    custom_xml: []const u8,
    buf: []u8,
) ?[]const u8 {
    var search: usize = 0;
    while (std.mem.indexOfPos(u8, custom_xml, search, "<property")) |open| {
        const gt = std.mem.indexOfScalarPos(u8, custom_xml, open, '>') orelse return null;
        const attrs = custom_xml[open..gt];
        search = gt + 1;
        if (!definedNameMatches(attrs, DOC_PROP_NAME)) continue;

        const close = std.mem.indexOfPos(u8, custom_xml, gt + 1, "</property>") orelse return null;
        const body = custom_xml[gt + 1 .. close];
        // Body is `<vt:lpwstr>TEXT</vt:lpwstr>`; take the inner text.
        const inner_open = std.mem.indexOfScalarPos(u8, body, 0, '>') orelse return null;
        const inner_close = std.mem.lastIndexOfScalar(u8, body, '<') orelse return null;
        if (inner_close <= inner_open) return null;
        const text = body[inner_open + 1 .. inner_close];
        if (text.len > buf.len) return null;
        return xmlUnescape(buf, text);
    }
    return null;
}

// ─── tests ──────────────────────────────────────────────────────────

const testing = std.testing;

fn sampleInput() RecordInput {
    return .{
        .model = "text-embedding-3-small",
        .dim = 1536,
        .dtype = "int8-sym-per-vec",
        .hash_algo = "xxh3-64",
        .coverages = &.{
            .{ .id = "title", .worksheet_target = "worksheets/sheet1.xml", .range = "A2:A500", .count = 499 },
            .{ .id = "body", .worksheet_target = "worksheets/sheet1.xml", .range = "B2:B500", .count = 499 },
        },
        .digest = 0xDEAD_BEEF_CAFE_1234,
    };
}

test "encode then decode round-trips every field" {
    const a = testing.allocator;
    const enc = try encode(a, sampleInput());
    defer a.free(enc);

    var covs: [8]RecoveredCoverage = undefined;
    var scratch: [512]u8 = undefined;
    const rec = try decode(enc, .defined_name, &covs, &scratch);

    try testing.expectEqualStrings("text-embedding-3-small", rec.model);
    try testing.expectEqual(@as(u32, 1536), rec.dim);
    try testing.expectEqualStrings("int8-sym-per-vec", rec.dtype);
    try testing.expectEqualStrings("xxh3-64", rec.hash_algo);
    try testing.expectEqual(@as(usize, 2), rec.coverages.len);
    try testing.expectEqualStrings("title", rec.coverages[0].id);
    try testing.expectEqualStrings("worksheets/sheet1.xml", rec.coverages[0].worksheet_target);
    try testing.expectEqualStrings("A2:A500", rec.coverages[0].range);
    try testing.expectEqual(@as(u32, 499), rec.coverages[0].count);
    try testing.expectEqualStrings("body", rec.coverages[1].id);
    try testing.expectEqual(@as(u64, 0xDEAD_BEEF_CAFE_1234), rec.digest);
    try testing.expectEqual(Carrier.defined_name, rec.carrier);
}

test "encoded record contains no XML- or formula-significant bytes" {
    const a = testing.allocator;
    // A model name full of exactly the characters that would break a
    // carrier: quotes, angle brackets, ampersand, and the separator.
    var in = sampleInput();
    in.model = "ev<il>&\"nasty\"|model spaces";
    const enc = try encode(a, in);
    defer a.free(enc);

    for (enc) |c| {
        try testing.expect(c != '<' and c != '>' and c != '&' and c != '"' and c != '\'');
        try testing.expect(c >= 0x20 and c < 0x7F); // printable ASCII only
    }
    // The separator survives as a separator: exactly one field holds
    // the model, and it decodes back to the original.
    var covs: [8]RecoveredCoverage = undefined;
    var scratch: [512]u8 = undefined;
    const rec = try decode(enc, .doc_props, &covs, &scratch);
    try testing.expectEqualStrings("ev<il>&\"nasty\"|model spaces", rec.model);
}

test "unknown magic is rejected rather than guessed" {
    var covs: [4]RecoveredCoverage = undefined;
    var scratch: [256]u8 = undefined;
    try testing.expectError(
        Error.UnsupportedRecordVersion,
        decode("zlsxER9|m|4|f32|xxh3-64|0|0000000000000000", .defined_name, &covs, &scratch),
    );
}

test "truncated record is rejected, not half-parsed" {
    var covs: [4]RecoveredCoverage = undefined;
    var scratch: [256]u8 = undefined;
    // Declares one coverage, supplies none.
    try testing.expectError(
        Error.MalformedRecord,
        decode("zlsxER1|m|4|f32|xxh3-64|1", .defined_name, &covs, &scratch),
    );
}

test "coverage storage too small reports BufferTooSmall" {
    const a = testing.allocator;
    const enc = try encode(a, sampleInput());
    defer a.free(enc);
    var covs: [1]RecoveredCoverage = undefined;
    var scratch: [512]u8 = undefined;
    try testing.expectError(Error.BufferTooSmall, decode(enc, .defined_name, &covs, &scratch));
}

test "chunking splits and rejoins losslessly" {
    const a = testing.allocator;
    var in = sampleInput();
    // Force multiple chunks.
    in.model = "m" ** 300;
    const enc = try encode(a, in);
    defer a.free(enc);

    const n = chunkCount(enc.len);
    try testing.expect(n > 1);

    var rejoined: std.ArrayListUnmanaged(u8) = .empty;
    defer rejoined.deinit(a);
    for (0..n) |i| try rejoined.appendSlice(a, chunk(enc, i));
    try testing.expectEqualStrings(enc, rejoined.items);
}

test "chunk names are stable and ordered" {
    var buf: [32]u8 = undefined;
    try testing.expectEqualStrings("_zlsxRecovery0", chunkName(&buf, 0));
    var buf2: [32]u8 = undefined;
    try testing.expectEqualStrings("_zlsxRecovery11", chunkName(&buf2, 11));
}

test "findDefinedNameValue survives the LibreOffice normalizations" {
    // Exactly the shape LibreOffice produced in the emb-4B measurement:
    // hidden rewritten to "true", payload XML-escaped, two attributes
    // injected, and attribute order changed.
    const xml =
        \\<workbook><definedNames><definedName function="false" hidden="true" name="_zlsxRecovery0" vbProcedure="false">&quot;zlsxER1|m|4|f32&quot;</definedName></definedNames></workbook>
    ;
    var buf: [256]u8 = undefined;
    const got = findDefinedNameValue(xml, "_zlsxRecovery0", &buf) orelse return error.TestExpectedValue;
    try testing.expectEqualStrings("zlsxER1|m|4|f32", got);
}

test "findDefinedNameValue reads the shape zlsx itself writes" {
    const xml =
        \\<workbook><definedNames><definedName name="_zlsxRecovery0" hidden="1">"zlsxER1|m|4|f32"</definedName></definedNames></workbook>
    ;
    var buf: [256]u8 = undefined;
    const got = findDefinedNameValue(xml, "_zlsxRecovery0", &buf) orelse return error.TestExpectedValue;
    try testing.expectEqualStrings("zlsxER1|m|4|f32", got);
}

test "findDefinedNameValue does not confuse a similarly-named attribute" {
    // `codeName=` ends in `name=`; matching it would read the wrong
    // element. Also pins that a different name is not returned.
    const xml =
        \\<workbook><definedNames><definedName codeName="_zlsxRecovery0" name="SomethingElse">"nope"</definedName></definedNames></workbook>
    ;
    var buf: [256]u8 = undefined;
    try testing.expect(findDefinedNameValue(xml, "_zlsxRecovery0", &buf) == null);
}

test "findDefinedNameValue skips non-matching names before the match" {
    const xml =
        \\<workbook><definedNames><definedName name="Other">"x"</definedName><definedName name="_zlsxRecovery0">"zlsxER1|ok"</definedName></definedNames></workbook>
    ;
    var buf: [256]u8 = undefined;
    const got = findDefinedNameValue(xml, "_zlsxRecovery0", &buf) orelse return error.TestExpectedValue;
    try testing.expectEqualStrings("zlsxER1|ok", got);
}

test "findDocPropValue extracts the lpwstr body" {
    const xml =
        \\<?xml version="1.0"?><Properties><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="ZlsxEmbeddingRecovery"><vt:lpwstr>zlsxER1|m|4|f32</vt:lpwstr></property></Properties>
    ;
    var buf: [256]u8 = undefined;
    const got = findDocPropValue(xml, &buf) orelse return error.TestExpectedValue;
    try testing.expectEqualStrings("zlsxER1|m|4|f32", got);
}

test "xmlUnescape resolves the five predefined entities and passes others through" {
    var buf: [128]u8 = undefined;
    try testing.expectEqualStrings(
        "a&b<c>d\"e'f",
        xmlUnescape(&buf, "a&amp;b&lt;c&gt;d&quot;e&apos;f"),
    );
    var buf2: [128]u8 = undefined;
    // An unknown entity is not a reason to discard a readable record.
    try testing.expectEqualStrings("x&nbsp;y", xmlUnescape(&buf2, "x&nbsp;y"));
}

test "percent-decode rejects a truncated escape" {
    var buf: [16]u8 = undefined;
    try testing.expectError(Error.MalformedRecord, percentDecode(&buf, "ab%4"));
    try testing.expectError(Error.MalformedRecord, percentDecode(&buf, "ab%ZZ"));
}

test "zero coverages encodes and decodes" {
    const a = testing.allocator;
    var in = sampleInput();
    in.coverages = &.{};
    const enc = try encode(a, in);
    defer a.free(enc);
    var covs: [4]RecoveredCoverage = undefined;
    var scratch: [256]u8 = undefined;
    const rec = try decode(enc, .defined_name, &covs, &scratch);
    try testing.expectEqual(@as(usize, 0), rec.coverages.len);
    try testing.expectEqual(@as(u64, 0xDEAD_BEEF_CAFE_1234), rec.digest);
}

test "a realistic two-coverage record fits in one chunk" {
    const a = testing.allocator;
    const enc = try encode(a, sampleInput());
    defer a.free(enc);
    // The design budgets ~200 bytes; hold that line so a format change
    // that silently doubles the size shows up here.
    try testing.expect(enc.len <= MAX_CHUNK);
    try testing.expectEqual(@as(usize, 1), chunkCount(enc.len));
}

test "findRecordInText pulls a record out of surrounding sheet XML" {
    const xml =
        \\<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>zlsxER1|m|4|f32|xxh3-64|0|0000000000000000</t></is></c></row></sheetData></worksheet>
    ;
    var buf: [256]u8 = undefined;
    const got = findRecordInText(xml, &buf) orelse return error.TestExpectedValue;
    try testing.expectEqualStrings("zlsxER1|m|4|f32|xxh3-64|0|0000000000000000", got);
}

test "findRecordInText survives the shared-strings landing site" {
    const sst =
        \\<sst count="2"><si><t>Alpha</t></si><si><t>zlsxER1|m|4|f32|xxh3-64|0|00000000000000ff</t></si></sst>
    ;
    var buf: [256]u8 = undefined;
    const got = findRecordInText(sst, &buf) orelse return error.TestExpectedValue;
    try testing.expectEqualStrings("zlsxER1|m|4|f32|xxh3-64|0|00000000000000ff", got);
}

test "findRecordInText returns null when no record is present" {
    var buf: [64]u8 = undefined;
    try testing.expect(findRecordInText("<sst><si><t>Alpha</t></si></sst>", &buf) == null);
}

test "cell carrier is a distinct Carrier variant" {
    // Guards against a future refactor collapsing it into doc_props:
    // the compat matrices report which carrier answered, and cell_data
    // answering means something different from the other two.
    try testing.expect(Carrier.cell_data != Carrier.doc_props);
    try testing.expect(Carrier.cell_data != Carrier.defined_name);
}

test "findDefinedNameValue: a chunk inside a comment, a CDATA section, a PI or a DOCTYPE is text — the live element after it is the record" {
    var buf: [64]u8 = undefined;
    const real = "<definedNames><definedName name=\"_zlsxRecovery0\" hidden=\"1\">\"real\"</definedName></definedNames></workbook>";
    const decoy = "<definedName name=\"_zlsxRecovery0\">decoy</definedName>";
    // Every construct the typed parser (and the strip) skips: the
    // comment (r4), and the three r5 S3C-REL-502 named — a decoy in a
    // CDATA section or a PI survived every re-embed and was read as the
    // record, so a stripped read answered the PREVIOUS generation.
    const openers = [_][]const u8{
        "<workbook><!-- " ++ decoy ++ " -->",
        "<workbook><![CDATA[ " ++ decoy ++ " ]]>",
        "<workbook><?zlsx " ++ decoy ++ " ?>",
        "<!DOCTYPE workbook [ <!-- " ++ decoy ++ " --> ]><workbook>",
    };
    inline for (openers) |opener| {
        const xml = opener ++ real;
        try testing.expectEqualStrings("real", findDefinedNameValue(xml, "_zlsxRecovery0", &buf).?);
        // Only the decoy: no record.
        try testing.expect(findDefinedNameValue(opener ++ "<definedNames/></workbook>", "_zlsxRecovery0", &buf) == null);
    }
    // An unterminated construct is not a part to read a record from.
    const torn = [_][]const u8{
        "<workbook><!-- " ++ decoy,
        "<workbook><![CDATA[ " ++ decoy,
        "<workbook><?zlsx " ++ decoy,
    };
    inline for (torn) |xml| {
        try testing.expect(findDefinedNameValue(xml, "_zlsxRecovery0", &buf) == null);
    }
}
