//! B3 iter-wr-5: shared ZIP archive emit substrate.
//!
//! This module is the canonical home for the LFH + payload + CDFH +
//! EOCD layout used by both `xlsx.Writer.save` (fresh-file producer)
//! and `pkg.PartStore.save` (load-modify-save). Pre-iter-wr-5,
//! Writer carried a private `ZipWriter` struct (~165 LOC) at
//! `src/writer.zig:3265` that duplicated the byte-stable invariants
//! enforced inline by `PartStore.save`'s u32-sentinel guards, the
//! 1 KiB STORED-vs-DEFLATE policy, and the EOCD record layout.
//! Both producers now route through the same code paths.
//!
//! ## Surface
//!
//! `Archive` is the thin appender:
//!
//!   var arc = zip.Archive.init(alloc, &out_buf);
//!   defer arc.deinit();
//!   try arc.addEntry("[Content_Types].xml", ct_bytes, deflateFn);
//!   try arc.addEntry("xl/workbook.xml",     wb_bytes, deflateFn);
//!   ...
//!   try arc.finalize();
//!
//! `deflateFn` is supplied by the caller to keep this module
//! stdlib-only — the canonical implementation lives at
//! `src/writer.zig`'s `deflateCompress` (re-exported from `zlsx`).
//! Decoupling the deflater from the layout lets `pkg/zip.zig` sit
//! beneath every consumer in the module graph (no cycle through
//! `zlsx → writer.zig → pkg → zlsx`).
//!
//! ## Byte-stability invariants (preserved verbatim from
//! `src/writer.zig`'s legacy `ZipWriter`)
//!
//!   - **ZIP32 only.** Every serialised u32 size / offset field is
//!     rejected at `0xFFFFFFFF` (Zip64 sentinel). cd_size + cd_offset
//!     are re-checked AFTER the CD is fully written. Total file size
//!     is bounded so zlsx's own reader (which rejects > maxInt(u32))
//!     can round-trip the output (`d6235f3` total-size guard).
//!   - **Per-entry compression policy.** Sub-1 KiB inputs go STORED
//!     (method 0). ≥ 1 KiB run through `deflateFn`; if the compressed
//!     output is at least as large as the input, the entry falls back
//!     to STORED.
//!   - **No data descriptors.** All sizes + CRC32 live in the LFH.
//!   - **CRC32** is computed over the uncompressed bytes (per ZIP32).
//!   - **LFH version-needed.** 20 always. General-purpose flag = 0
//!     (no encryption, no data descriptors, no UTF-8 name bit-11).
//!   - **EOCD comment empty** (`comment_length = 0`); `disk_number = 0`;
//!     `central_directory_disk_number = 0`.
//!   - **Big-endian host byte-swap.** ZIP signatures + size fields
//!     are little-endian on disk; on a big-endian host we byte-swap
//!     each header before appending.
//!
//! ## Module-graph note
//!
//! `pkg/zip.zig` lives in `pkg/` because it's the typed-overlay /
//! package layer's home for OOXML container concerns. Wired into
//! the build as a standalone module (`zlsx_zip`) imported into
//! `writer_mod`, `package_mod`, and any other consumer that needs
//! to emit a ZIP. Std-only — no cycle through `zlsx_pkg`.

const std = @import("std");

const Allocator = std.mem.Allocator;

/// Caller-provided deflate function. Must compress `input` (which
/// the Archive guarantees is non-empty when called) and append the
/// raw deflate stream into `out`. The canonical implementation is
/// `xlsx.deflateCompress` (`src/writer.zig`).
pub const DeflateFn = *const fn (Allocator, []const u8, *std.ArrayListUnmanaged(u8)) anyerror!void;

pub const Error = error{
    EntryTooLarge,
    NameTooLong,
    TooManyZipEntries,
    ZipArchiveTooLarge,
} || Allocator.Error;

/// Minimal zip archive builder. Appends file entries to a byte buffer;
/// `finalize()` emits the central directory + end-of-central-directory
/// trailer. Each entry is deflate-compressed unless compression grows
/// the payload (empty entries, near-random bytes), in which case the
/// entry falls back to STORED (method 0). Both Excel and LibreOffice
/// accept mixed-method archives.
pub const Archive = struct {
    allocator: Allocator,
    out: *std.ArrayListUnmanaged(u8),
    /// Per-entry info accumulated for the central directory.
    entries: std.ArrayListUnmanaged(EntryMeta) = .empty,

    const EntryMeta = struct {
        /// Owned copy (arena would also work but the Archive's lifetime
        /// is short-lived around a save() call, so a per-entry dupe
        /// freed in deinit keeps the API self-contained).
        name: []u8,
        crc32: u32,
        compressed_size: u32,
        uncompressed_size: u32,
        local_offset: u32,
        method: std.zip.CompressionMethod,
    };

    pub fn init(alloc: Allocator, out: *std.ArrayListUnmanaged(u8)) Archive {
        return .{ .allocator = alloc, .out = out };
    }

    pub fn deinit(self: *Archive) void {
        for (self.entries.items) |e| self.allocator.free(e.name);
        self.entries.deinit(self.allocator);
    }

    /// Append one entry. `deflate` is the caller's compression
    /// function; passed in so `pkg/zip.zig` stays stdlib-only and
    /// downstream of the deflate implementation in the module graph.
    pub fn addEntry(
        self: *Archive,
        name: []const u8,
        data: []const u8,
        deflate: DeflateFn,
    ) !void {
        const alloc = self.allocator;
        if (data.len > std.math.maxInt(u32)) return Error.EntryTooLarge;
        if (name.len > std.math.maxInt(u16)) return Error.NameTooLong;

        const crc = std.hash.Crc32.hash(data);
        const offset: u32 = @intCast(self.out.items.len);

        // Sub-1 KB entries skip compression. The dynamic-huffman block
        // header adds ~60-120 bytes of fixed overhead that rarely pays
        // back on tiny XML fragments (Content_Types.xml, workbook rels,
        // empty sheet templates) — and the hash-chain init is pure waste.
        // The big entries (sheet1.xml, sharedStrings.xml, styles.xml)
        // dominate archive size, so bypassing small ones loses negligible
        // savings and shaves real per-entry wall time.
        //
        // If deflate still inflates a ≥ 1 KB payload (already-compressed
        // or near-random content), fall back to STORED.
        const COMPRESS_MIN: usize = 1024;
        var compressed: std.ArrayListUnmanaged(u8) = .empty;
        defer compressed.deinit(alloc);

        var method: std.zip.CompressionMethod = .deflate;
        var payload: []const u8 = undefined;
        if (data.len >= COMPRESS_MIN) {
            try deflate(alloc, data, &compressed);
        }
        if (data.len < COMPRESS_MIN or compressed.items.len >= data.len) {
            method = .store;
            payload = data;
        } else {
            payload = compressed.items;
        }
        if (payload.len > std.math.maxInt(u32)) return Error.EntryTooLarge;

        const hdr: std.zip.LocalFileHeader = .{
            .signature = std.zip.local_file_header_sig,
            .version_needed_to_extract = 20,
            .flags = .{ .encrypted = false, ._ = 0 },
            .compression_method = method,
            .last_modification_time = 0,
            .last_modification_date = 0x21, // 1980-01-01, minimum valid
            .crc32 = crc,
            .compressed_size = @intCast(payload.len),
            .uncompressed_size = @intCast(data.len),
            .filename_len = @intCast(name.len),
            .extra_len = 0,
        };
        try appendStruct(alloc, self.out, std.zip.LocalFileHeader, hdr);
        try self.out.appendSlice(alloc, name);
        try self.out.appendSlice(alloc, payload);

        const owned_name = try alloc.dupe(u8, name);
        errdefer alloc.free(owned_name);
        try self.entries.append(alloc, .{
            .name = owned_name,
            .crc32 = crc,
            .compressed_size = @intCast(payload.len),
            .uncompressed_size = @intCast(data.len),
            .local_offset = offset,
            .method = method,
        });
    }

    pub fn finalize(self: *Archive) !void {
        const alloc = self.allocator;
        // ZIP32 EOCD records the per-disk + total entry counts in u16
        // fields. >65535 entries needs Zip64 (which we don't emit).
        // Without this guard the @intCast at the EndRecord build trapped
        // in safe builds and silently truncated in ReleaseFast — both
        // produce a workbook Excel rejects.
        if (self.entries.items.len > std.math.maxInt(u16)) {
            return Error.TooManyZipEntries;
        }
        // cd_start IS written to the EOCD as u32, so 0xFFFFFFFF is
        // the Zip64 sentinel — readers (including zlsx's own) treat
        // it as "look for Zip64 extra fields". We don't emit Zip64,
        // so reject `>= 0xFFFFFFFF` strictly. (Total file length is
        // not on-wire, so a final byte count of 0xFFFFFFFF with a
        // smaller cd_start is fine — only the serialized field
        // matters.)
        if (self.out.items.len >= std.math.maxInt(u32)) {
            return Error.ZipArchiveTooLarge;
        }
        const cd_start: u32 = @intCast(self.out.items.len);

        for (self.entries.items) |e| {
            const cd: std.zip.CentralDirectoryFileHeader = .{
                .signature = std.zip.central_file_header_sig,
                .version_made_by = 20,
                .version_needed_to_extract = 20,
                .flags = .{ .encrypted = false, ._ = 0 },
                .compression_method = e.method,
                .last_modification_time = 0,
                .last_modification_date = 0x21,
                .crc32 = e.crc32,
                .compressed_size = e.compressed_size,
                .uncompressed_size = e.uncompressed_size,
                .filename_len = @intCast(e.name.len),
                .extra_len = 0,
                .comment_len = 0,
                .disk_number = 0,
                .internal_file_attributes = 0,
                .external_file_attributes = 0,
                .local_file_header_offset = e.local_offset,
            };
            try appendStruct(alloc, self.out, std.zip.CentralDirectoryFileHeader, cd);
            try self.out.appendSlice(alloc, e.name);
        }

        // What matters for ZIP32 sentinel-safety is the SERIALIZED
        // cd_size field — NOT the cd_end position (which isn't on
        // wire). Compute cd_size in u64 to avoid casting cd_end.
        const cd_size_u64 = self.out.items.len - cd_start;
        if (cd_size_u64 >= std.math.maxInt(u32)) {
            return Error.ZipArchiveTooLarge;
        }
        const cd_size: u32 = @intCast(cd_size_u64);
        // Round-trip guard: zlsx's reader rejects files whose
        // stat.size > maxInt(u32) (we don't support Zip64). Also
        // reject when the total written-so-far + EOCD would push
        // past 4 GiB, so we don't generate archives we can't open.
        // EndRecord is fixed-size; we use a generous overhead.
        if (self.out.items.len + 22 > std.math.maxInt(u32)) {
            return Error.ZipArchiveTooLarge;
        }

        const end: std.zip.EndRecord = .{
            .signature = std.zip.end_record_sig,
            .disk_number = 0,
            .central_directory_disk_number = 0,
            .record_count_disk = @intCast(self.entries.items.len),
            .record_count_total = @intCast(self.entries.items.len),
            .central_directory_size = cd_size,
            .central_directory_offset = cd_start,
            .comment_len = 0,
        };
        try appendStruct(alloc, self.out, std.zip.EndRecord, end);
    }
};

fn appendStruct(alloc: Allocator, out: *std.ArrayListUnmanaged(u8), comptime T: type, value: T) !void {
    // ZIP headers are defined little-endian on disk. On a big-endian
    // host, dumping the native struct bytes would emit byte-swapped
    // signatures/sizes/offsets and produce archives that Excel and
    // std.zip can't open. Mirrors the editor save path's pattern.
    var v = value;
    if (@import("builtin").cpu.arch.endian() != .little)
        std.mem.byteSwapAllFields(T, &v);
    const bytes = std.mem.asBytes(&v);
    try out.appendSlice(alloc, bytes);
}

// ─── Tests ───────────────────────────────────────────────────────────

const testing = std.testing;

/// Stub deflate that just stores. Lets the layout tests exercise the
/// LFH/CDFH/EOCD plumbing without dragging the writer's deflate impl
/// into this stdlib-only module's test graph.
fn stubDeflate(
    alloc: Allocator,
    input: []const u8,
    out: *std.ArrayListUnmanaged(u8),
) anyerror!void {
    // Append at-or-larger than input so the Archive falls back to
    // STORED (the policy when compressed.len >= data.len).
    try out.appendSlice(alloc, input);
    try out.append(alloc, 0);
}

test "Archive: empty archive finalises with EOCD only" {
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(testing.allocator);
    var arc = Archive.init(testing.allocator, &buf);
    defer arc.deinit();
    try arc.finalize();
    // EOCD is 22 bytes minimum; signature first.
    try testing.expectEqual(@as(usize, 22), buf.items.len);
    try testing.expectEqualSlices(u8, &std.zip.end_record_sig, buf.items[0..4]);
}

test "Archive: single small entry round-trips via std.zip.Iterator" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(testing.allocator);
    var arc = Archive.init(testing.allocator, &buf);
    defer arc.deinit();
    try arc.addEntry("hello.txt", "world", stubDeflate);
    try arc.finalize();

    // Spill to a tmp file + walk with std.zip.Iterator.
    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    {
        try tmp.dir.writeFile(io, .{ .sub_path = "a.zip", .data = buf.items });
    }
    var f = try tmp.dir.openFile(io, "a.zip", .{});
    defer f.close(io);
    var read_buf: [4096]u8 = undefined;
    var fr = f.reader(io, &read_buf);
    var iter = try std.zip.Iterator.init(&fr);
    var seen: usize = 0;
    while (try iter.next()) |_| seen += 1;
    try testing.expectEqual(@as(usize, 1), seen);
}

test "Archive: multiple entries preserve order" {
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(testing.allocator);
    var arc = Archive.init(testing.allocator, &buf);
    defer arc.deinit();
    try arc.addEntry("a", "aaaa", stubDeflate);
    try arc.addEntry("b", "bbbb", stubDeflate);
    try arc.addEntry("c", "cccc", stubDeflate);
    try arc.finalize();

    try testing.expectEqual(@as(usize, 3), arc.entries.items.len);
    try testing.expectEqualStrings("a", arc.entries.items[0].name);
    try testing.expectEqualStrings("b", arc.entries.items[1].name);
    try testing.expectEqualStrings("c", arc.entries.items[2].name);
}

test "Archive: stub deflate falls back to STORED" {
    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(testing.allocator);
    var arc = Archive.init(testing.allocator, &buf);
    defer arc.deinit();
    // Force the ≥ 1 KiB branch so deflate is actually called; stub
    // returns input + 1 byte, so the Archive selects STORED.
    var big: [1500]u8 = undefined;
    @memset(&big, 'x');
    try arc.addEntry("big.bin", &big, stubDeflate);
    try testing.expectEqual(std.zip.CompressionMethod.store, arc.entries.items[0].method);
    try arc.finalize();
}
