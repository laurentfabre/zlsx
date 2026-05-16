//! Embedding-part wire-format primitives (emb-1a).
//!
//! Pure functions over byte slices. No XML, no Unicode text
//! canonicalization, no allocator-owning surfaces (allocator only
//! appears where the caller explicitly asked for an owned buffer
//! return). The XML + text canonicalization pieces ship in emb-1b.
//!
//! See `docs/plans/embeddings-in-xlsx.md` for the format spec. This
//! file implements:
//!
//! - `VecHeader` (24 bytes) + `HashHeader` (16 bytes) byte layout
//!   with comptime size asserts pinning the disk offset of record 0.
//! - `Dtype` enum mapping the wire byte + the XML attribute string.
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
//! - `xxh3` thin wrapper over `std.hash.XxHash3` so callers don't
//!   pin a specific stdlib symbol.
//!
//! Tombstone marker (load-bearing): `TOMBSTONE_HASH` = `u64::MAX`.
//! Per the spec, query-time consumers MUST skip slots whose stored
//! hash equals this value (zero-vector slots cause NaN under
//! cosine).

const std = @import("std");
const Allocator = std.mem.Allocator;

pub const Error = error{
    BadMagic,
    UnsupportedVersion,
    InvalidDtype,
    InvalidReservedBytes,
    HeaderTooShort,
    BodyTooShort,
    CountMismatch,
    DimensionOutOfRange,
    MalformedNumber,
    BufferTooSmall,
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
    const want_body: usize = @as(usize, count) * dtype.recordBytes(dim);
    if (bytes.len < VEC_HEADER_BYTES + want_body) return Error.BodyTooShort;
    return .{ .version = version, .dim = dim, .count = count, .dtype = dtype };
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
    const want_body: usize = @as(usize, count) * @sizeOf(u64);
    if (bytes.len < HASH_HEADER_BYTES + want_body) return Error.BodyTooShort;
    return .{ .version = version, .count = count };
}

/// Cross-check between a parsed VecHeader and a parsed HashHeader
/// for the same coverage: counts MUST match.
pub fn checkPairConsistent(vec: ParsedVecHeader, hash: ParsedHashHeader) Error!void {
    if (vec.count != hash.count) return Error.CountMismatch;
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

test "checkPairConsistent: count mismatch flagged" {
    const v: ParsedVecHeader = .{ .version = 1, .dim = 4, .count = 10, .dtype = .f32 };
    const h_ok: ParsedHashHeader = .{ .version = 1, .count = 10 };
    const h_bad: ParsedHashHeader = .{ .version = 1, .count = 11 };
    try checkPairConsistent(v, h_ok);
    try testing.expectError(Error.CountMismatch, checkPairConsistent(v, h_bad));
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
