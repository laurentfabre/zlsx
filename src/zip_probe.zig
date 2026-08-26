//! Test support: a raw ZIP32 emitter whose central-directory fields are
//! whatever the caller says they are.
//!
//! The decompression limits (S1, `zlsx_control.decompress_limits`) are
//! checked against *declared* sizes, so exercising them needs archives
//! that declare what no real producer would — a 513 MiB part behind a
//! 128 KiB payload, five 512 MiB parts in a 640 KiB file. `Writer` cannot
//! emit those and mutating a real workbook's bytes can only reach the
//! values a random flip lands on. This builder writes each entry's LFH,
//! payload, CDFH and the EOCD from the fields it is given, so a test
//! chooses the exact declaration and the archive stays structurally
//! sound everywhere else (matching LFH/CDFH names, in-bounds offsets,
//! no data descriptor, no Zip64, no encryption): a refusal it triggers
//! is then attributable to the limit, not to a shape check that fired
//! first.
//!
//! Exported from `src/xlsx.zig` as `zlsx.zip_probe` for the same reason
//! `zlsx_pkg.fill_probe` is: `pkg/` tests reach the `zlsx` module and
//! nothing else in `src/`. Production code never calls it.

const std = @import("std");

/// One archive entry as the caller wants it *declared*. `payload` is
/// copied verbatim after the local header — its length is the
/// compressed size the directory records, so structural checks that
/// compare the payload span against the file length pass. Nothing here
/// inflates `payload`; a limit fires before any decompressor runs.
pub const Entry = struct {
    name: []const u8,
    payload: []const u8,
    /// 0 = stored, 8 = deflate. Only the number goes on the wire.
    method: u16 = 8,
    /// What the central directory (and the local header) claim the
    /// entry inflates to. Never verified by this builder.
    declared_uncompressed: u32,
    crc32: u32 = 0,
};

/// Emit a single-disk, comment-less ZIP32 archive over `entries`.
/// Caller owns the returned slice.
pub fn build(allocator: std.mem.Allocator, entries: []const Entry) ![]u8 {
    var out: std.Io.Writer.Allocating = .init(allocator);
    errdefer out.deinit();
    const w = &out.writer;

    const lfh_offsets = try allocator.alloc(u32, entries.len);
    defer allocator.free(lfh_offsets);

    for (entries, 0..) |e, i| {
        lfh_offsets[i] = @intCast(out.written().len);
        try w.writeAll(&std.zip.local_file_header_sig);
        try w.writeInt(u16, 20, .little); // version needed
        try w.writeInt(u16, 0, .little); // flags: no data descriptor, no encryption
        try w.writeInt(u16, e.method, .little);
        try w.writeInt(u16, 0, .little); // mtime
        try w.writeInt(u16, 0, .little); // mdate
        try w.writeInt(u32, e.crc32, .little);
        try w.writeInt(u32, @intCast(e.payload.len), .little);
        try w.writeInt(u32, e.declared_uncompressed, .little);
        try w.writeInt(u16, @intCast(e.name.len), .little);
        try w.writeInt(u16, 0, .little); // extra
        try w.writeAll(e.name);
        try w.writeAll(e.payload);
    }

    const cd_offset: u32 = @intCast(out.written().len);
    for (entries, 0..) |e, i| {
        try w.writeAll(&std.zip.central_file_header_sig);
        try w.writeInt(u16, 20, .little); // version made by
        try w.writeInt(u16, 20, .little); // version needed
        try w.writeInt(u16, 0, .little); // flags
        try w.writeInt(u16, e.method, .little);
        try w.writeInt(u16, 0, .little); // mtime
        try w.writeInt(u16, 0, .little); // mdate
        try w.writeInt(u32, e.crc32, .little);
        try w.writeInt(u32, @intCast(e.payload.len), .little);
        try w.writeInt(u32, e.declared_uncompressed, .little);
        try w.writeInt(u16, @intCast(e.name.len), .little);
        try w.writeInt(u16, 0, .little); // extra
        try w.writeInt(u16, 0, .little); // comment
        try w.writeInt(u16, 0, .little); // disk
        try w.writeInt(u16, 0, .little); // internal attrs
        try w.writeInt(u32, 0, .little); // external attrs
        try w.writeInt(u32, lfh_offsets[i], .little);
        try w.writeAll(e.name);
    }
    const cd_size: u32 = @intCast(out.written().len - cd_offset);

    try w.writeAll(&std.zip.end_record_sig);
    try w.writeInt(u16, 0, .little); // this disk
    try w.writeInt(u16, 0, .little); // cd disk
    try w.writeInt(u16, @intCast(entries.len), .little);
    try w.writeInt(u16, @intCast(entries.len), .little);
    try w.writeInt(u32, cd_size, .little);
    try w.writeInt(u32, cd_offset, .little);
    try w.writeInt(u16, 0, .little); // comment

    return try out.toOwnedSlice();
}

/// The three hostile shapes every opener must refuse, and the one
/// in-bounds control it must not, expressed against
/// `zlsx_control.decompress_limits` so the fixtures track the numbers
/// the owner gate settles on. Each returns an archive whose entries
/// live under `xl/media/` — a prefix no opener extracts — so the only
/// thing an opener can object to before the limits is nothing, and the
/// only thing after them is the missing workbook part.
pub const hostile = struct {
    const control = @import("zlsx_control");
    const limits = control.decompress_limits;

    /// One part declared one byte past the per-part cap, with enough
    /// payload behind it that the ratio cap is satisfied: only the
    /// per-part check can refuse this.
    pub fn oversizedPart(allocator: std.mem.Allocator) ![]u8 {
        const declared: u32 = @intCast(limits.max_part_size + 1);
        const payload_len: usize = @intCast(limits.max_part_size / limits.max_deflate_ratio + 1);
        const payload = try allocator.alloc(u8, payload_len);
        defer allocator.free(payload);
        @memset(payload, 0);
        return build(allocator, &.{.{ .name = "xl/media/a.bin", .payload = payload, .declared_uncompressed = declared }});
    }

    /// One byte of payload declaring `max_deflate_ratio + 1` bytes —
    /// far under the per-part cap, so only the ratio check can refuse.
    pub fn absurdRatio(allocator: std.mem.Allocator) ![]u8 {
        const declared: u32 = @intCast(limits.max_deflate_ratio + 1);
        return build(allocator, &.{.{ .name = "xl/media/a.bin", .payload = "x", .declared_uncompressed = declared }});
    }

    /// `n` parts each declaring exactly the per-part cap behind exactly
    /// the payload the ratio cap requires — every part is individually
    /// admissible. The sum crosses the aggregate iff
    /// `n * max_part_size > max_total_decompressed`.
    pub fn fullSizedParts(allocator: std.mem.Allocator, n: usize) ![]u8 {
        const declared: u32 = @intCast(limits.max_part_size);
        const payload_len: usize = @intCast(limits.max_part_size / limits.max_deflate_ratio);
        const payload = try allocator.alloc(u8, payload_len);
        defer allocator.free(payload);
        @memset(payload, 0);
        const entries = try allocator.alloc(Entry, n);
        defer allocator.free(entries);
        const names = try allocator.alloc([32]u8, n);
        defer allocator.free(names);
        for (entries, 0..) |*e, i| {
            const name = try std.fmt.bufPrint(&names[i], "xl/media/p{d}.bin", .{i});
            e.* = .{ .name = name, .payload = payload, .declared_uncompressed = declared };
        }
        return build(allocator, entries);
    }

    /// Smallest part count whose full-sized sum exceeds the aggregate.
    pub fn partsOverBudget() usize {
        return @intCast(limits.max_total_decompressed / limits.max_part_size + 1);
    }

    /// Largest part count whose full-sized sum stays within it.
    pub fn partsWithinBudget() usize {
        return @intCast(limits.max_total_decompressed / limits.max_part_size);
    }
};

test "build: std.zip walks what this emits and reads back the declared fields" {
    const a = std.testing.allocator;
    const bytes = try build(a, &.{
        .{ .name = "a.txt", .payload = "hello", .method = 0, .declared_uncompressed = 5 },
        .{ .name = "b/c.bin", .payload = "\x03\x00", .method = 8, .declared_uncompressed = 1234 },
    });
    defer a.free(bytes);

    var threaded: std.Io.Threaded = .init(a, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    try tmp.dir.writeFile(io, .{ .sub_path = "probe.zip", .data = bytes });
    const f = try tmp.dir.openFile(io, "probe.zip", .{});
    defer f.close(io);
    var buf: [4096]u8 = undefined;
    var fr = f.reader(io, &buf);
    var it = try std.zip.Iterator.init(&fr);
    const e0 = (try it.next()).?;
    try std.testing.expectEqual(@as(u64, 5), e0.compressed_size);
    try std.testing.expectEqual(@as(u64, 5), e0.uncompressed_size);
    try std.testing.expectEqual(std.zip.CompressionMethod.store, e0.compression_method);
    const e1 = (try it.next()).?;
    try std.testing.expectEqual(@as(u64, 2), e1.compressed_size);
    try std.testing.expectEqual(@as(u64, 1234), e1.uncompressed_size);
    try std.testing.expectEqual(std.zip.CompressionMethod.deflate, e1.compression_method);
    try std.testing.expectEqual(@as(?std.zip.Iterator.Entry, null), try it.next());
}

test "hostile: the four shapes have the sizes their names promise" {
    const a = std.testing.allocator;
    const l = hostile.limits;
    try std.testing.expect(hostile.partsOverBudget() * l.max_part_size > l.max_total_decompressed);
    try std.testing.expect(hostile.partsWithinBudget() * l.max_part_size <= l.max_total_decompressed);
    const over = try hostile.fullSizedParts(a, hostile.partsOverBudget());
    defer a.free(over);
    // Five 128 KiB payloads, not five 512 MiB ones: the archive itself
    // stays small enough to build in every test run.
    try std.testing.expect(over.len < 2 * hostile.partsOverBudget() * (l.max_part_size / l.max_deflate_ratio));
}
