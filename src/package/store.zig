//! B0 phase 1 (post-0.2.9 roadmap): PartStore — read-only OOXML
//! package layer.
//!
//! Opens a .xlsx ZIP archive, walks the central directory,
//! decompresses every part eagerly, parses `[Content_Types].xml` and
//! every `_rels/*.rels`, and exposes a typed read-only API:
//!
//!   - `partNames()` — central-directory order, source-preserving.
//!   - `part(name)` — `Part { name, content_type, bytes, compression_method }`.
//!   - `rels(owner)` — `[]Relationship`.
//!   - `resolve(owner, target)` — collapse `..`/`.` segments per OPC.
//!
//! Out of scope for this milestone: `save`, `replacePart`, `addPart`,
//! dirty flags, typed overlays for known parts. Those land in B0 M2/M3
//! per docs/plans/post-0.2.9-roadmap.md.
//!
//! Relationship to `Editor` (src/xlsx.zig): the Editor has a battle-
//! tested ZIP scanner that supports the byte-preserving save path; we
//! deliberately keep PartStore independent of Editor's internals here
//! to ship M1 without disturbing the Editor surface. M2 will extract
//! a shared scanner (planned src/package/zip.zig) so both modules
//! consume the same code.

const std = @import("std");

pub const Part = struct {
    name: []const u8,
    /// Resolved via `[Content_Types].xml` (Override > Default by
    /// extension). `null` when neither matches — callers can treat
    /// that as "unknown content type".
    content_type: ?[]const u8,
    /// Decompressed part bytes. Owned by the PartStore arena;
    /// borrowed by the caller for the PartStore's lifetime.
    bytes: []const u8,
    /// ZIP compression method as recorded in the central directory
    /// (0 = stored, 8 = deflate). Useful for callers that want to
    /// re-emit byte-for-byte later.
    compression_method: u16,
};

pub const TargetMode = enum { internal, external };

pub const Relationship = struct {
    id: []const u8,
    type: []const u8,
    target: []const u8,
    target_mode: TargetMode,
};

pub const Error = error{
    NotPkzip,
    BadZip,
    UnsupportedCompression,
    Zip64NotSupported,
    EncryptedNotSupported,
    SplitArchiveNotSupported,
    DataDescriptorNotSupported,
    DuplicateContentType,
} || std.fs.File.OpenError || std.mem.Allocator.Error || std.fs.File.ReadError || std.fs.File.SeekError;

pub const PartStore = struct {
    allocator: std.mem.Allocator,
    /// Arena-owned storage for every borrowed slice exposed via the
    /// public API (part names, content types, rel attrs, decompressed
    /// part bytes). Caller-visible slices stay valid until `deinit`.
    arena: std.heap.ArenaAllocator,
    parts: []Part,
    /// Map from part name → relationships parsed from
    /// `_rels/<owner>.rels`. Empty list when the part has no rels.
    rels_by_owner: std.StringHashMapUnmanaged([]Relationship),

    pub fn open(allocator: std.mem.Allocator, path: []const u8) !PartStore {
        var arena = std.heap.ArenaAllocator.init(allocator);
        errdefer arena.deinit();
        const ar_alloc = arena.allocator();

        // Slurp the file into owned memory. Mirrors Editor's lifetime
        // story: we want to stay alive even after the source file is
        // closed, and we want a contiguous span for the central-
        // directory scan.
        const file = try std.fs.cwd().openFile(path, .{});
        defer file.close();
        const stat = try file.stat();
        if (stat.size > std.math.maxInt(u32)) return Error.Zip64NotSupported;
        const size: usize = @intCast(stat.size);
        const buf = try ar_alloc.alloc(u8, size);
        const n = try file.readAll(buf);
        if (n != size) return Error.BadZip;

        const entries = try scanCentralDirectory(ar_alloc, buf);

        // Decompress each entry eagerly. M1 chose eager because it
        // makes content-types / rels parsing trivial and keeps the
        // public API allocator-free.
        const parts = try ar_alloc.alloc(Part, entries.len);
        for (entries, 0..) |e, i| {
            const compressed = buf[e.payload_offset .. e.payload_offset + e.compressed_size];
            const bytes = try decompressPayload(
                ar_alloc,
                compressed,
                e.compression_method,
                e.uncompressed_size,
            );
            parts[i] = .{
                .name = e.name,
                .content_type = null, // resolved next pass
                .bytes = bytes,
                .compression_method = e.compression_method,
            };
        }

        // Resolve content types from `[Content_Types].xml`. Default
        // by extension, Override by part name (Override wins).
        try resolveContentTypes(parts);

        // Parse each `_rels/*.rels`. The result is keyed by the
        // owner part name (the document the rels file describes).
        var rels_by_owner: std.StringHashMapUnmanaged([]Relationship) = .empty;
        for (parts) |p| {
            const owner = (try relsOwner(ar_alloc, p.name)) orelse continue;
            const relationships = try parseRelationships(ar_alloc, p.bytes);
            try rels_by_owner.put(ar_alloc, owner, relationships);
        }

        return .{
            .allocator = allocator,
            .arena = arena,
            .parts = parts,
            .rels_by_owner = rels_by_owner,
        };
    }

    pub fn deinit(self: *PartStore) void {
        self.arena.deinit();
    }

    pub fn partNames(self: *const PartStore) []const []const u8 {
        // Build a names-only view on demand. Cheap because parts.len
        // is small; alternative is to cache a separate slice up-front.
        // For M1 keep the data structures minimal.
        const ar_alloc = @constCast(&self.arena).allocator();
        const out = ar_alloc.alloc([]const u8, self.parts.len) catch return &.{};
        for (self.parts, 0..) |p, i| out[i] = p.name;
        return out;
    }

    pub fn part(self: *const PartStore, name: []const u8) ?Part {
        for (self.parts) |p| {
            if (std.mem.eql(u8, p.name, name)) return p;
        }
        return null;
    }

    pub fn rels(self: *const PartStore, owner_part_name: []const u8) []const Relationship {
        return self.rels_by_owner.get(owner_part_name) orelse &.{};
    }

    /// Resolve a relationship `target` (which is interpreted relative
    /// to `owner_part_name`'s parent directory) into a normalised
    /// absolute part name. Returns `null` for external targets and
    /// for paths that escape the package root.
    pub fn resolve(
        self: *const PartStore,
        owner_part_name: []const u8,
        target: []const u8,
    ) !?[]const u8 {
        // Absolute target (rare): "/xl/foo.xml" → "xl/foo.xml".
        if (target.len == 0) return null;
        if (target[0] == '/') {
            return try self.dupeArena(target[1..]);
        }
        // Relative target: collapse against owner's parent dir.
        const owner_dir = parentDir(owner_part_name);
        var stack: std.ArrayListUnmanaged([]const u8) = .empty;
        defer stack.deinit(self.allocator);
        // Seed stack with owner_dir segments.
        if (owner_dir.len > 0) {
            var it = std.mem.splitScalar(u8, owner_dir, '/');
            while (it.next()) |seg| {
                if (seg.len == 0) continue;
                try stack.append(self.allocator, seg);
            }
        }
        // Walk target.
        var it = std.mem.splitScalar(u8, target, '/');
        while (it.next()) |seg| {
            if (seg.len == 0 or std.mem.eql(u8, seg, ".")) continue;
            if (std.mem.eql(u8, seg, "..")) {
                if (stack.items.len == 0) return null; // escaped package
                _ = stack.pop();
                continue;
            }
            try stack.append(self.allocator, seg);
        }
        // Join with `/`. Total length = sum + (n-1) separators.
        var total: usize = 0;
        for (stack.items, 0..) |s, i| {
            total += s.len;
            if (i + 1 < stack.items.len) total += 1;
        }
        const ar_alloc = @constCast(&self.arena).allocator();
        const out = try ar_alloc.alloc(u8, total);
        var w: usize = 0;
        for (stack.items, 0..) |s, i| {
            @memcpy(out[w .. w + s.len], s);
            w += s.len;
            if (i + 1 < stack.items.len) {
                out[w] = '/';
                w += 1;
            }
        }
        return out;
    }

    fn dupeArena(self: *const PartStore, s: []const u8) ![]u8 {
        const ar_alloc = @constCast(&self.arena).allocator();
        return ar_alloc.dupe(u8, s);
    }
};

// ─── ZIP scanner ──────────────────────────────────────────────────────

const ZipEntry = struct {
    name: []const u8,
    payload_offset: usize,
    compressed_size: u32,
    uncompressed_size: u32,
    compression_method: u16,
    crc32: u32,
};

const eocd_signature: u32 = 0x06054b50;
const cdfh_signature: u32 = 0x02014b50;
const lfh_signature: u32 = 0x04034b50;
const eocd_min_size: usize = 22;
const eocd_scan_window: usize = 65535 + eocd_min_size;

fn scanCentralDirectory(arena: std.mem.Allocator, buf: []const u8) ![]ZipEntry {
    if (buf.len < eocd_min_size) return Error.NotPkzip;
    const eocd_off = try findEocd(buf);

    const cd_size = std.mem.readInt(u32, buf[eocd_off + 12 ..][0..4], .little);
    const cd_offset = std.mem.readInt(u32, buf[eocd_off + 16 ..][0..4], .little);
    const total_records = std.mem.readInt(u16, buf[eocd_off + 10 ..][0..2], .little);

    if (cd_size == 0xFFFFFFFF or cd_offset == 0xFFFFFFFF) return Error.Zip64NotSupported;
    if (cd_offset + cd_size > buf.len) return Error.BadZip;

    var out: std.ArrayListUnmanaged(ZipEntry) = .empty;
    try out.ensureTotalCapacity(arena, total_records);

    var cur: usize = cd_offset;
    const cd_end = cd_offset + cd_size;
    var idx: usize = 0;
    while (cur + 46 <= cd_end and idx < total_records) : (idx += 1) {
        const sig = std.mem.readInt(u32, buf[cur..][0..4], .little);
        if (sig != cdfh_signature) return Error.BadZip;

        const flags = std.mem.readInt(u16, buf[cur + 8 ..][0..2], .little);
        if (flags & 0x0001 != 0) return Error.EncryptedNotSupported;
        // Data-descriptor flag (0x0008) means the LFH's CRC + sizes
        // are zero and the real values trail the payload. The CDFH
        // copies we read have the real sizes either way, so this is
        // OK as long as we use CDFH-recorded sizes (we do).

        const compression_method = std.mem.readInt(u16, buf[cur + 10 ..][0..2], .little);
        const crc32 = std.mem.readInt(u32, buf[cur + 16 ..][0..4], .little);
        const compressed_size = std.mem.readInt(u32, buf[cur + 20 ..][0..4], .little);
        const uncompressed_size = std.mem.readInt(u32, buf[cur + 24 ..][0..4], .little);
        const filename_len = std.mem.readInt(u16, buf[cur + 28 ..][0..2], .little);
        const extra_len = std.mem.readInt(u16, buf[cur + 30 ..][0..2], .little);
        const comment_len = std.mem.readInt(u16, buf[cur + 32 ..][0..2], .little);
        const lfh_offset = std.mem.readInt(u32, buf[cur + 42 ..][0..4], .little);

        if (compressed_size == 0xFFFFFFFF or uncompressed_size == 0xFFFFFFFF or lfh_offset == 0xFFFFFFFF) {
            return Error.Zip64NotSupported;
        }
        if (compression_method != 0 and compression_method != 8) {
            return Error.UnsupportedCompression;
        }

        const name_start = cur + 46;
        if (name_start + filename_len > cd_end) return Error.BadZip;
        const name = buf[name_start .. name_start + filename_len];

        // Compute payload offset by reading the LFH (its filename +
        // extra fields can differ from the CDFH copy).
        if (lfh_offset + 30 > buf.len) return Error.BadZip;
        const lfh_sig = std.mem.readInt(u32, buf[lfh_offset..][0..4], .little);
        if (lfh_sig != lfh_signature) return Error.BadZip;
        const lfh_name_len = std.mem.readInt(u16, buf[lfh_offset + 26 ..][0..2], .little);
        const lfh_extra_len = std.mem.readInt(u16, buf[lfh_offset + 28 ..][0..2], .little);
        const payload_offset = lfh_offset + 30 + @as(usize, lfh_name_len) + @as(usize, lfh_extra_len);
        if (payload_offset + compressed_size > buf.len) return Error.BadZip;

        try out.append(arena, .{
            .name = try arena.dupe(u8, name),
            .payload_offset = payload_offset,
            .compressed_size = compressed_size,
            .uncompressed_size = uncompressed_size,
            .compression_method = compression_method,
            .crc32 = crc32,
        });

        cur = name_start + filename_len + extra_len + comment_len;
    }

    if (idx != total_records) return Error.BadZip;
    return out.toOwnedSlice(arena);
}

fn findEocd(buf: []const u8) !usize {
    // Scan back from the end of buf looking for the EOCD signature.
    // ZIP allows up to 65535 bytes of comment after EOCD, so the
    // scan window is bounded; longer would mean a malformed archive.
    const max_back = @min(eocd_scan_window, buf.len);
    var i: usize = buf.len - 4;
    const lo = if (buf.len >= max_back) buf.len - max_back else 0;
    while (true) {
        const sig = std.mem.readInt(u32, buf[i..][0..4], .little);
        if (sig == eocd_signature) {
            const eocd_end = i + eocd_min_size;
            if (eocd_end > buf.len) return Error.BadZip;
            const comment_len = std.mem.readInt(u16, buf[i + 20 ..][0..2], .little);
            if (eocd_end + comment_len != buf.len) {
                // Comment must consume the file tail exactly. If not,
                // this is a stray signature inside another structure.
            } else {
                return i;
            }
        }
        if (i == lo) break;
        i -= 1;
    }
    return Error.NotPkzip;
}

fn decompressPayload(
    arena: std.mem.Allocator,
    payload: []const u8,
    method: u16,
    declared_uncompressed: u32,
) ![]u8 {
    if (method == 0) {
        if (payload.len != declared_uncompressed) return Error.BadZip;
        return try arena.dupe(u8, payload);
    } else if (method == 8) {
        var src_reader = std.Io.Reader.fixed(payload);
        var flate_buffer: [std.compress.flate.max_window_len]u8 = undefined;
        var dec = std.compress.flate.Decompress.init(&src_reader, .raw, &flate_buffer);
        const out = try arena.alloc(u8, declared_uncompressed);
        var out_writer = std.Io.Writer.fixed(out);
        dec.reader.streamExact64(&out_writer, declared_uncompressed) catch return Error.BadZip;
        return out;
    }
    return Error.UnsupportedCompression;
}

// ─── Content types ─ [Content_Types].xml ─────────────────────────────

fn resolveContentTypes(parts: []Part) !void {
    const ct_part = blk: {
        for (parts) |p| {
            if (std.mem.eql(u8, p.name, "[Content_Types].xml")) break :blk p;
        }
        // No content-types part — leave every Part.content_type as null.
        return;
    };
    const xml = ct_part.bytes;

    // Phase 1: collect Default extension → content type.
    // Phase 2: collect Override partName → content type.
    // Phase 3: assign each part's content_type (Override wins).

    var i: usize = 0;
    while (std.mem.indexOfPos(u8, xml, i, "<Default")) |pos| {
        const end = std.mem.indexOfScalarPos(u8, xml, pos, '>') orelse break;
        const attrs = xml[pos..end];
        const ext = attrAtSlice(attrs, "Extension") orelse {
            i = end + 1;
            continue;
        };
        const ct = attrAtSlice(attrs, "ContentType") orelse {
            i = end + 1;
            continue;
        };
        // Apply default to every part with that extension.
        for (parts) |*p| {
            if (p.content_type != null) continue; // Override may set later
            if (extensionEql(p.name, ext)) p.content_type = ct;
        }
        i = end + 1;
    }
    i = 0;
    while (std.mem.indexOfPos(u8, xml, i, "<Override")) |pos| {
        const end = std.mem.indexOfScalarPos(u8, xml, pos, '>') orelse break;
        const attrs = xml[pos..end];
        const part_name = attrAtSlice(attrs, "PartName") orelse {
            i = end + 1;
            continue;
        };
        const ct = attrAtSlice(attrs, "ContentType") orelse {
            i = end + 1;
            continue;
        };
        // PartName starts with `/`; strip to match part.name.
        const stripped = if (part_name.len > 0 and part_name[0] == '/') part_name[1..] else part_name;
        for (parts) |*p| {
            if (std.mem.eql(u8, p.name, stripped)) {
                p.content_type = ct;
                break;
            }
        }
        i = end + 1;
    }
}

fn attrAtSlice(attrs: []const u8, key: []const u8) ?[]const u8 {
    // Match `<key>="<value>"`. Tolerate single-quote variant too for
    // adversarial XML, but OOXML always emits double quotes.
    var search_buf: [64]u8 = undefined;
    if (key.len + 2 > search_buf.len) return null;
    @memcpy(search_buf[0..key.len], key);
    search_buf[key.len] = '=';
    search_buf[key.len + 1] = '"';
    const needle = search_buf[0 .. key.len + 2];
    const found = std.mem.indexOf(u8, attrs, needle) orelse return null;
    const start = found + needle.len;
    const close = std.mem.indexOfScalarPos(u8, attrs, start, '"') orelse return null;
    return attrs[start..close];
}

fn extensionEql(name: []const u8, ext: []const u8) bool {
    if (std.mem.lastIndexOfScalar(u8, name, '.')) |dot| {
        return std.ascii.eqlIgnoreCase(name[dot + 1 ..], ext);
    }
    return false;
}

// ─── Relationships ─ _rels/X.rels ────────────────────────────────────

/// Return the part name a `_rels/Y.rels` file describes, allocating
/// the result inside `arena`. Examples:
///   "_rels/.rels"                    → ""             (package level)
///   "xl/_rels/workbook.xml.rels"     → "xl/workbook.xml"
///   "xl/worksheets/_rels/sheet1.xml.rels" → "xl/worksheets/sheet1.xml"
/// Returns null if `name` isn't a `_rels/<base>.rels` shape.
fn relsOwner(arena: std.mem.Allocator, name: []const u8) !?[]const u8 {
    if (!std.mem.endsWith(u8, name, ".rels")) return null;
    // Find the `_rels` directory segment. It's either prefix-less
    // ("_rels/...") or trailing-prefix ("X/_rels/...").
    const marker = "_rels/";
    const marker_pos = std.mem.lastIndexOf(u8, name, marker) orelse return null;
    // Marker must end with the last `/` before the basename.
    const last_slash = std.mem.lastIndexOfScalar(u8, name, '/') orelse return null;
    if (marker_pos + marker.len - 1 != last_slash) return null;

    const prefix = if (marker_pos == 0) "" else name[0..marker_pos];
    const filename_start = last_slash + 1;
    const filename_end = name.len - ".rels".len;
    if (filename_end < filename_start) return null;
    const filename = name[filename_start..filename_end];

    if (filename.len == 0) {
        // Package-level rels: "_rels/.rels" → owner = "".
        return try arena.dupe(u8, "");
    }
    if (prefix.len == 0) {
        return try arena.dupe(u8, filename);
    }
    return try std.fs.path.join(arena, &.{ prefix, filename });
}

fn parseRelationships(arena: std.mem.Allocator, xml: []const u8) ![]Relationship {
    var out: std.ArrayListUnmanaged(Relationship) = .empty;
    var i: usize = 0;
    while (std.mem.indexOfPos(u8, xml, i, "<Relationship")) |pos| {
        // Skip `<Relationships ...>` (the wrapper) — only match
        // `<Relationship ` or `<Relationship>` followed by
        // attributes; the root tag is `<Relationships>` (with `s`).
        const after = pos + "<Relationship".len;
        if (after >= xml.len) break;
        const c = xml[after];
        if (c != ' ' and c != '\t' and c != '\n' and c != '\r' and c != '/' and c != '>') {
            i = after;
            continue;
        }
        const end = std.mem.indexOfScalarPos(u8, xml, pos, '>') orelse break;
        const attrs = xml[pos..end];
        const id = attrAtSlice(attrs, "Id") orelse {
            i = end + 1;
            continue;
        };
        const rtype = attrAtSlice(attrs, "Type") orelse {
            i = end + 1;
            continue;
        };
        const target = attrAtSlice(attrs, "Target") orelse {
            i = end + 1;
            continue;
        };
        const mode_str = attrAtSlice(attrs, "TargetMode");
        const target_mode: TargetMode = if (mode_str) |m|
            (if (std.mem.eql(u8, m, "External")) .external else .internal)
        else
            .internal;

        try out.append(arena, .{
            .id = try arena.dupe(u8, id),
            .type = try arena.dupe(u8, rtype),
            .target = try arena.dupe(u8, target),
            .target_mode = target_mode,
        });
        i = end + 1;
    }
    return out.toOwnedSlice(arena);
}

fn parentDir(name: []const u8) []const u8 {
    if (std.mem.lastIndexOfScalar(u8, name, '/')) |slash| return name[0..slash];
    return "";
}

// ─── Tests ────────────────────────────────────────────────────────────

test "PartStore.open: enumerates parts of frictionless_2sheets.xlsx" {
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(fixture, .{}) catch return error.SkipZigTest;

    var store = try PartStore.open(std.testing.allocator, fixture);
    defer store.deinit();

    const names = store.partNames();
    try std.testing.expect(names.len > 0);

    // Required OOXML parts must always be present.
    var seen_ct = false;
    var seen_wb = false;
    for (names) |n| {
        if (std.mem.eql(u8, n, "[Content_Types].xml")) seen_ct = true;
        if (std.mem.eql(u8, n, "xl/workbook.xml")) seen_wb = true;
    }
    try std.testing.expect(seen_ct);
    try std.testing.expect(seen_wb);
}

test "PartStore.part: workbook.xml has a workbook content type" {
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(fixture, .{}) catch return error.SkipZigTest;

    var store = try PartStore.open(std.testing.allocator, fixture);
    defer store.deinit();

    const wb = store.part("xl/workbook.xml") orelse return error.TestUnexpectedResult;
    const ct = wb.content_type orelse return error.TestUnexpectedResult;
    // OOXML ContentType for the workbook part. The canonical
    // workbook content-type is `…spreadsheetml.sheet.main+xml`
    // (NOT "…workbook+xml" — common confusion).
    try std.testing.expect(std.mem.indexOf(u8, ct, "spreadsheetml") != null);
    try std.testing.expect(std.mem.indexOf(u8, ct, "sheet.main") != null);
    // Bytes must be non-empty.
    try std.testing.expect(wb.bytes.len > 0);
    // workbook.xml is XML; should start with '<'.
    try std.testing.expectEqual(@as(u8, '<'), wb.bytes[0]);
}

test "PartStore.rels: package-root + workbook rels parse" {
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(fixture, .{}) catch return error.SkipZigTest;

    var store = try PartStore.open(std.testing.allocator, fixture);
    defer store.deinit();

    // Package-level rels (`_rels/.rels`, owner = "") must point at
    // `xl/workbook.xml`.
    const root_rels = store.rels("");
    try std.testing.expect(root_rels.len > 0);
    var found_wb = false;
    for (root_rels) |r| {
        if (std.mem.endsWith(u8, r.target, "workbook.xml")) found_wb = true;
    }
    try std.testing.expect(found_wb);

    // Workbook rels (`xl/_rels/workbook.xml.rels`, owner =
    // `xl/workbook.xml`) must list at least one worksheet.
    const wb_rels = store.rels("xl/workbook.xml");
    try std.testing.expect(wb_rels.len > 0);
    var found_ws = false;
    for (wb_rels) |r| {
        if (std.mem.indexOf(u8, r.target, "worksheets/sheet") != null) found_ws = true;
    }
    try std.testing.expect(found_ws);
}

test "relsOwner: shape decoder" {
    var arena = std.heap.ArenaAllocator.init(std.testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    try std.testing.expectEqualStrings("", (try relsOwner(a, "_rels/.rels")).?);
    try std.testing.expectEqualStrings(
        "xl/workbook.xml",
        (try relsOwner(a, "xl/_rels/workbook.xml.rels")).?,
    );
    try std.testing.expectEqualStrings(
        "xl/worksheets/sheet1.xml",
        (try relsOwner(a, "xl/worksheets/_rels/sheet1.xml.rels")).?,
    );
    // Negative cases.
    try std.testing.expectEqual(@as(?[]const u8, null), try relsOwner(a, "xl/workbook.xml"));
    try std.testing.expectEqual(@as(?[]const u8, null), try relsOwner(a, "xl/foo.rels"));
}

test "PartStore.resolve: relative + absolute targets" {
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(fixture, .{}) catch return error.SkipZigTest;

    var store = try PartStore.open(std.testing.allocator, fixture);
    defer store.deinit();

    // Relative: workbook.xml's rels target "worksheets/sheet1.xml" →
    // "xl/worksheets/sheet1.xml".
    const r1 = (try store.resolve("xl/workbook.xml", "worksheets/sheet1.xml")).?;
    try std.testing.expectEqualStrings("xl/worksheets/sheet1.xml", r1);

    // Relative with `..`: "../sharedStrings.xml" from a worksheet
    // resolves to "xl/sharedStrings.xml".
    const r2 = (try store.resolve("xl/worksheets/sheet1.xml", "../sharedStrings.xml")).?;
    try std.testing.expectEqualStrings("xl/sharedStrings.xml", r2);

    // Absolute: "/xl/workbook.xml" → "xl/workbook.xml".
    const r3 = (try store.resolve("anywhere", "/xl/workbook.xml")).?;
    try std.testing.expectEqualStrings("xl/workbook.xml", r3);
}

test "PartStore.open: rejects non-PK file" {
    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    try tmp.dir.writeFile(.{ .sub_path = "garbage.xlsx", .data = "not a zip" });
    const dir = try tmp.dir.realpathAlloc(std.testing.allocator, ".");
    defer std.testing.allocator.free(dir);
    const path = try std.fs.path.join(std.testing.allocator, &.{ dir, "garbage.xlsx" });
    defer std.testing.allocator.free(path);

    try std.testing.expectError(Error.NotPkzip, PartStore.open(std.testing.allocator, path));
}
