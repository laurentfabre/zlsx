//! Minimal ZIP reader for the oracle extractor (M1b).
//!
//! **Independence is the point.** The oracle exists to tell us whether
//! zlsx computed the right answer. If it read workbooks through
//! `pkg/zip.zig` and `src/xlsx.zig`, a bug in either would be present
//! on both sides of the comparison and cancel out — the oracle would
//! confirm zlsx against itself. So this reads the central directory by
//! hand and inflates through `std.compress.flate`, sharing no code with
//! the implementation under test.
//!
//! Scope is deliberately narrow: whole-file-in-memory, no ZIP64, no
//! encryption, no data descriptors, store and deflate only. Oracle
//! inputs are workbooks we or Excel/LibreOffice just wrote, all a few
//! KiB — the exotic archives `pkg/zip.zig` handles are not in play, and
//! every unsupported case is a typed error rather than a guess.

const std = @import("std");

pub const Error = error{
    ZipNoEndOfCentralDirectory,
    ZipTruncated,
    ZipBadCentralDirectorySignature,
    ZipBadLocalHeaderSignature,
    ZipMultiDiskUnsupported,
    ZipEncryptionUnsupported,
    ZipUnsupportedCompressionMethod,
    ZipEntryNotFound,
    ZipCrcMismatch,
    ZipSizeMismatch,
    ZipInflateFailed,
    OutOfMemory,
};

const eocd_sig = [4]u8{ 'P', 'K', 5, 6 };
const cd_sig = [4]u8{ 'P', 'K', 1, 2 };
const lfh_sig = [4]u8{ 'P', 'K', 3, 4 };

const eocd_len = 22;
const cd_header_len = 46;
const lfh_len = 30;

pub const Entry = struct {
    name: []const u8,
    compression_method: u16,
    crc32: u32,
    compressed_size: u32,
    uncompressed_size: u32,
    local_header_offset: u32,
};

pub const Archive = struct {
    bytes: []const u8,
    entries: []Entry,

    pub fn deinit(self: *Archive, allocator: std.mem.Allocator) void {
        allocator.free(self.entries);
        self.* = undefined;
    }

    /// Entry with exactly this name, or null. Part names in an xlsx are
    /// byte-exact and case-sensitive.
    pub fn find(self: Archive, name: []const u8) ?Entry {
        for (self.entries) |e| {
            if (std.mem.eql(u8, e.name, name)) return e;
        }
        return null;
    }

    /// Decompress one entry. Caller frees. CRC-32 and uncompressed size
    /// are both verified — a silently truncated inflate would otherwise
    /// look like a workbook that legitimately lacks the missing cells,
    /// which is exactly the kind of false oracle result this harness
    /// must never produce.
    pub fn read(self: Archive, allocator: std.mem.Allocator, name: []const u8) Error![]u8 {
        const entry = self.find(name) orelse return error.ZipEntryNotFound;

        const lfh_start = entry.local_header_offset;
        if (lfh_start + lfh_len > self.bytes.len) return error.ZipTruncated;
        const lfh = self.bytes[lfh_start..][0..lfh_len];
        if (!std.mem.eql(u8, lfh[0..4], &lfh_sig)) return error.ZipBadLocalHeaderSignature;

        // The local header repeats the name and extra-field lengths, and
        // they may differ from the central directory's — the payload
        // starts after the LOCAL copies.
        const local_name_len = std.mem.readInt(u16, lfh[26..28], .little);
        const local_extra_len = std.mem.readInt(u16, lfh[28..30], .little);
        const data_start = lfh_start + lfh_len + local_name_len + local_extra_len;
        const data_end = data_start + entry.compressed_size;
        if (data_end > self.bytes.len) return error.ZipTruncated;
        const compressed = self.bytes[data_start..data_end];

        const out = switch (entry.compression_method) {
            0 => blk: {
                if (compressed.len != entry.uncompressed_size) return error.ZipSizeMismatch;
                break :blk try allocator.dupe(u8, compressed);
            },
            8 => try inflateRaw(allocator, compressed, entry.uncompressed_size),
            else => return error.ZipUnsupportedCompressionMethod,
        };
        errdefer allocator.free(out);

        if (out.len != entry.uncompressed_size) return error.ZipSizeMismatch;
        if (std.hash.Crc32.hash(out) != entry.crc32) return error.ZipCrcMismatch;
        return out;
    }
};

fn inflateRaw(allocator: std.mem.Allocator, compressed: []const u8, hint: u32) Error![]u8 {
    var input: std.Io.Reader = .fixed(compressed);
    var window: [std.compress.flate.max_window_len]u8 = undefined;
    var decompress: std.compress.flate.Decompress = .init(&input, .raw, &window);

    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(allocator);
    try out.ensureTotalCapacity(allocator, hint);

    var writer: std.Io.Writer.Allocating = .fromArrayList(allocator, &out);
    defer out = writer.toArrayList();
    _ = decompress.reader.streamRemaining(&writer.writer) catch return error.ZipInflateFailed;

    return writer.toOwnedSlice() catch error.OutOfMemory;
}

/// Parse `bytes` (a whole .xlsx held in memory) into its entry list.
/// `bytes` is borrowed — it must outlive the returned `Archive`.
pub fn open(allocator: std.mem.Allocator, bytes: []const u8) Error!Archive {
    const eocd_offset = findEocd(bytes) orelse return error.ZipNoEndOfCentralDirectory;
    const eocd = bytes[eocd_offset..][0..eocd_len];

    const disk_number = std.mem.readInt(u16, eocd[4..6], .little);
    const cd_disk = std.mem.readInt(u16, eocd[6..8], .little);
    if (disk_number != 0 or cd_disk != 0) return error.ZipMultiDiskUnsupported;

    const record_count = std.mem.readInt(u16, eocd[10..12], .little);
    const cd_size = std.mem.readInt(u32, eocd[12..16], .little);
    const cd_offset = std.mem.readInt(u32, eocd[16..20], .little);
    if (@as(u64, cd_offset) + cd_size > bytes.len) return error.ZipTruncated;

    var entries: std.ArrayListUnmanaged(Entry) = .empty;
    errdefer entries.deinit(allocator);
    try entries.ensureTotalCapacity(allocator, record_count);

    var pos: usize = cd_offset;
    var seen: usize = 0;
    while (seen < record_count) : (seen += 1) {
        if (pos + cd_header_len > bytes.len) return error.ZipTruncated;
        const h = bytes[pos..][0..cd_header_len];
        if (!std.mem.eql(u8, h[0..4], &cd_sig)) return error.ZipBadCentralDirectorySignature;

        const flags = std.mem.readInt(u16, h[8..10], .little);
        if (flags & 0x1 != 0) return error.ZipEncryptionUnsupported;

        const name_len = std.mem.readInt(u16, h[28..30], .little);
        const extra_len = std.mem.readInt(u16, h[30..32], .little);
        const comment_len = std.mem.readInt(u16, h[32..34], .little);
        const name_start = pos + cd_header_len;
        const name_end = name_start + name_len;
        if (name_end > bytes.len) return error.ZipTruncated;

        entries.appendAssumeCapacity(.{
            .name = bytes[name_start..name_end],
            .compression_method = std.mem.readInt(u16, h[10..12], .little),
            .crc32 = std.mem.readInt(u32, h[16..20], .little),
            .compressed_size = std.mem.readInt(u32, h[20..24], .little),
            .uncompressed_size = std.mem.readInt(u32, h[24..28], .little),
            .local_header_offset = std.mem.readInt(u32, h[42..46], .little),
        });

        pos = name_end + extra_len + comment_len;
    }

    return .{ .bytes = bytes, .entries = try entries.toOwnedSlice(allocator) };
}

/// Scan backwards for the end-of-central-directory record. The comment
/// field is variable-length, so the signature can sit anywhere in the
/// last 64 KiB + 22 bytes; searching backwards finds the real record
/// rather than a signature-shaped run inside a compressed payload.
fn findEocd(bytes: []const u8) ?usize {
    if (bytes.len < eocd_len) return null;
    const max_comment = 0xFFFF;
    const search_floor = if (bytes.len > eocd_len + max_comment)
        bytes.len - eocd_len - max_comment
    else
        0;

    var i = bytes.len - eocd_len;
    while (true) : (i -= 1) {
        if (std.mem.eql(u8, bytes[i..][0..4], &eocd_sig)) {
            const comment_len = std.mem.readInt(u16, bytes[i + 20 ..][0..2], .little);
            if (i + eocd_len + comment_len == bytes.len) return i;
        }
        if (i == search_floor) return null;
    }
}

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

/// Read a whole file for a test, or skip when it isn't there. Corpus
/// workbooks are fetched by `scripts/fetch_test_corpus.sh` and absent on
/// a clean checkout, so every test that wants one has to tolerate its
/// absence rather than fail the build.
fn readFileOrSkip(io: std.Io, path: []const u8) ![]u8 {
    return std.Io.Dir.cwd().readFileAlloc(io, path, testing.allocator, .limited(32 << 20)) catch
        return error.SkipZigTest;
}

test "reads a real workbook's parts and verifies their CRCs" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();

    const bytes = try readFileOrSkip(io, "tests/corpus/openxlsx_loadExample.xlsx");
    defer testing.allocator.free(bytes);

    var archive = try open(testing.allocator, bytes);
    defer archive.deinit(testing.allocator);

    try testing.expect(archive.entries.len > 0);
    // Every xlsx has these two; if the central directory walk were off
    // by a field, the names would come back as garbage.
    try testing.expect(archive.find("[Content_Types].xml") != null);
    try testing.expect(archive.find("xl/workbook.xml") != null);

    const wb = try archive.read(testing.allocator, "xl/workbook.xml");
    defer testing.allocator.free(wb);
    try testing.expect(std.mem.indexOf(u8, wb, "<workbook") != null);
    try testing.expect(std.mem.indexOf(u8, wb, "<sheet") != null);
}

/// Corpus workbooks the extractor must decode COMPLETELY — every entry,
/// every CRC. Named explicitly rather than "everything that happens to
/// work", so a regression that starts failing one of them is a test
/// failure instead of a silently smaller success count.
const fully_decodable = [_][]const u8{
    "openxlsx_loadExample.xlsx",       "ons_cpi_detailed.xlsx",
    "phpsheet_3654c.xlsx",             "ecdc_covid.xlsx",
    "worldbank_catalog.xlsx",          "poi_58325_db.xlsx",
    "frictionless_2sheets.xlsx",       "calamine_encoded_entities.xlsx",
    "calamine_empty_s_attribute.xlsx", "poi_57893_many_merges.xlsx",
    "openpyxl_guess_types.xlsx",       "phpoi_test1.xlsx",
};

test "every entry of every known-good corpus workbook inflates with a matching CRC" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var checked: usize = 0;
    var files: usize = 0;
    for (fully_decodable) |name| {
        var path_buf: [128]u8 = undefined;
        const path = try std.fmt.bufPrint(&path_buf, "tests/corpus/{s}", .{name});
        const bytes = std.Io.Dir.cwd().readFileAlloc(io, path, testing.allocator, .limited(32 << 20)) catch
            continue; // corpus not fetched
        defer testing.allocator.free(bytes);
        files += 1;

        var archive = try open(testing.allocator, bytes);
        defer archive.deinit(testing.allocator);

        for (archive.entries) |entry| {
            const data = try archive.read(testing.allocator, entry.name);
            testing.allocator.free(data);
            checked += 1;
        }
    }
    if (files == 0) return error.SkipZigTest;
    try testing.expect(checked > 50);
}

test "adversarial corpus fixtures fail loudly, never silently" {
    // The corpus carries deliberately-broken archives. Each must produce
    // a typed error — the failure mode this harness cannot have is
    // returning plausible-looking bytes for a corrupt part, because the
    // oracle would then record a wrong expected value as ground truth.
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();

    const Case = struct { name: []const u8, part: ?[]const u8 };
    const cases = [_]Case{
        // Payload corrupted after the CRC was computed.
        .{ .name = "derived_bad_crc32.xlsx", .part = null },
        // Truncated three different ways.
        .{ .name = "derived_truncated_mid_payload.xlsx", .part = null },
        .{ .name = "derived_truncated_pre_eocd.xlsx", .part = null },
        .{ .name = "derived_truncated_signature.xlsx", .part = null },
        // Encrypted entries: refused, not decoded as garbage.
        .{ .name = "ziprs_aes_archive.zip", .part = null },
    };

    var seen: usize = 0;
    for (cases) |c| {
        var path_buf: [128]u8 = undefined;
        const path = try std.fmt.bufPrint(&path_buf, "tests/corpus/{s}", .{c.name});
        const bytes = std.Io.Dir.cwd().readFileAlloc(io, path, testing.allocator, .limited(32 << 20)) catch
            continue;
        defer testing.allocator.free(bytes);
        seen += 1;

        var archive = open(testing.allocator, bytes) catch continue; // refused at open — fine
        defer archive.deinit(testing.allocator);

        // Opened, so at least one entry must refuse on read.
        var refused = false;
        for (archive.entries) |entry| {
            if (archive.read(testing.allocator, entry.name)) |data| {
                testing.allocator.free(data);
            } else |_| {
                refused = true;
            }
        }
        if (!refused) {
            std.debug.print("adversarial fixture {s} decoded cleanly — expected a refusal\n", .{c.name});
            return error.TestUnexpectedResult;
        }
    }
    if (seen == 0) return error.SkipZigTest;
}

test "rejects a truncated archive rather than reading past the end" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();

    const bytes = try readFileOrSkip(io, "tests/corpus/openxlsx_loadExample.xlsx");
    defer testing.allocator.free(bytes);

    // Chopping the tail removes the end record entirely.
    try testing.expectError(
        error.ZipNoEndOfCentralDirectory,
        open(testing.allocator, bytes[0 .. bytes.len / 2]),
    );
    // An empty or sub-record-sized buffer is the same answer, not a panic.
    try testing.expectError(error.ZipNoEndOfCentralDirectory, open(testing.allocator, ""));
    try testing.expectError(error.ZipNoEndOfCentralDirectory, open(testing.allocator, "PK"));
}

test "a corrupted payload fails the CRC check instead of returning wrong bytes" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();

    const original = try readFileOrSkip(io, "tests/corpus/openxlsx_loadExample.xlsx");
    defer testing.allocator.free(original);

    const bytes = try testing.allocator.dupe(u8, original);
    defer testing.allocator.free(bytes);

    var archive = try open(testing.allocator, bytes);
    defer archive.deinit(testing.allocator);

    // Any sufficiently large deflated entry will do; hard-coding a part
    // name couples the test to one workbook's storage choices.
    var target: ?Entry = null;
    for (archive.entries) |e| {
        if (e.compression_method == 8 and e.compressed_size > 32) {
            target = e;
            break;
        }
    }
    const entry = target orelse return error.SkipZigTest;

    // Flip a bit deep inside the deflate stream. Inflate may still
    // succeed and produce plausible-looking bytes; the CRC is what
    // catches it.
    const lfh_start = entry.local_header_offset;
    const name_len = std.mem.readInt(u16, bytes[lfh_start + 26 ..][0..2], .little);
    const extra_len = std.mem.readInt(u16, bytes[lfh_start + 28 ..][0..2], .little);
    const data_start = lfh_start + lfh_len + name_len + extra_len;
    bytes[data_start + entry.compressed_size / 2] ^= 0xFF;

    const result = archive.read(testing.allocator, entry.name);
    if (result) |data| {
        testing.allocator.free(data);
        std.debug.print("corrupting {s} produced clean bytes\n", .{entry.name});
        return error.TestUnexpectedResult;
    } else |err| switch (err) {
        error.ZipCrcMismatch, error.ZipSizeMismatch, error.ZipInflateFailed => {},
        else => return err,
    }
}

test "unknown part name is a typed error" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();

    const bytes = try readFileOrSkip(io, "tests/corpus/openxlsx_loadExample.xlsx");
    defer testing.allocator.free(bytes);

    var archive = try open(testing.allocator, bytes);
    defer archive.deinit(testing.allocator);
    try testing.expectError(
        error.ZipEntryNotFound,
        archive.read(testing.allocator, "xl/nope.xml"),
    );
}
