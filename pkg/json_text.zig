//! JSON text scalars for the NDJSON contracts every surface shares.
//!
//! The CLI's record writers and the C ABI's buffer exports emit the
//! same bytes for the same workbook (S3a: the `pivots` shape rides
//! along into the C ABI + Python), so the escaper lives once, here,
//! and both call it. The escaping is RFC 8259's minimum: the two
//! JSON metacharacters, the five short escapes, and `\u00XX` for the
//! remaining C0 controls; every other byte (UTF-8 included) passes
//! through verbatim.

const std = @import("std");

pub fn writeString(w: *std.Io.Writer, s: []const u8) !void {
    try w.writeByte('"');
    for (s) |c| switch (c) {
        '"' => try w.writeAll("\\\""),
        '\\' => try w.writeAll("\\\\"),
        '\n' => try w.writeAll("\\n"),
        '\r' => try w.writeAll("\\r"),
        '\t' => try w.writeAll("\\t"),
        0x08 => try w.writeAll("\\b"),
        0x0c => try w.writeAll("\\f"),
        0...0x07, 0x0b, 0x0e...0x1f => try w.print("\\u{x:0>4}", .{c}),
        else => try w.writeByte(c),
    };
    try w.writeByte('"');
}

pub fn writeOptString(w: *std.Io.Writer, s: ?[]const u8) !void {
    if (s) |v| try writeString(w, v) else try w.writeAll("null");
}

pub fn writeOptU32(w: *std.Io.Writer, v: ?u32) !void {
    if (v) |n| try w.print("{d}", .{n}) else try w.writeAll("null");
}

pub fn writeOptBool(w: *std.Io.Writer, v: ?bool) !void {
    if (v) |b| try w.writeAll(if (b) "true" else "false") else try w.writeAll("null");
}

/// A finite double as a JSON number: the shortest round-trip digits
/// in plain decimal for the magnitudes real workbooks carry (pane
/// splits in twips, iteration deltas), exponent form outside
/// [1e-6, 1e21) so a wide magnitude does not print hundreds of digits.
/// Both spellings are RFC 8259 numbers. Callers keep non-finite
/// values off the wire — they have no JSON spelling — so one here is
/// a programming error, not an input.
pub fn writeF64(w: *std.Io.Writer, v: f64) !void {
    std.debug.assert(std.math.isFinite(v));
    const mag = @abs(v);
    if (mag == 0 or (mag >= 1e-6 and mag < 1e21)) {
        try w.print("{d}", .{v});
    } else {
        try w.print("{e}", .{v});
    }
}

pub fn writeOptF64(w: *std.Io.Writer, v: ?f64) !void {
    if (v) |x| try writeF64(w, x) else try w.writeAll("null");
}

test "writeString: metacharacters, short escapes, C0 controls, UTF-8 verbatim" {
    var buf: [256]u8 = undefined;
    var w = std.Io.Writer.fixed(&buf);
    try writeString(&w, "a\"b\\c\nd\re\tf\x08g\x0ch\x01i\x1fj\x7fk café");
    try std.testing.expectEqualStrings(
        "\"a\\\"b\\\\c\\nd\\re\\tf\\bg\\fh\\u0001i\\u001fj\x7fk café\"",
        w.buffered(),
    );
}

test "writeF64: plain decimal inside [1e-6, 1e21), exponent form outside, zero and sign kept" {
    var buf: [256]u8 = undefined;
    var w = std.Io.Writer.fixed(&buf);
    const vals = [_]f64{ 0, -0.0, 1, 2865, 1215.5, 0.001, 0.000001, 1e-7, 2.5e-7, 1e21, 1e20, -1.5, 1e300 };
    for (vals) |v| {
        try writeF64(&w, v);
        try w.writeByte(' ');
    }
    try std.testing.expectEqualStrings(
        "0 -0 1 2865 1215.5 0.001 0.000001 1e-7 2.5e-7 1e21 100000000000000000000 -1.5 1e300 ",
        w.buffered(),
    );
}

test "writeOptBool / writeOptF64: null spelled as JSON null" {
    var buf: [64]u8 = undefined;
    var w = std.Io.Writer.fixed(&buf);
    try writeOptBool(&w, null);
    try w.writeByte(' ');
    try writeOptBool(&w, true);
    try w.writeByte(' ');
    try writeOptBool(&w, false);
    try w.writeByte(' ');
    try writeOptF64(&w, null);
    try w.writeByte(' ');
    try writeOptF64(&w, 0.5);
    try std.testing.expectEqualStrings("null true false null 0.5", w.buffered());
}

test "writeOptString / writeOptU32: null spelled as JSON null" {
    var buf: [64]u8 = undefined;
    var w = std.Io.Writer.fixed(&buf);
    try writeOptString(&w, null);
    try w.writeByte(' ');
    try writeOptString(&w, "x");
    try w.writeByte(' ');
    try writeOptU32(&w, null);
    try w.writeByte(' ');
    try writeOptU32(&w, 7);
    try std.testing.expectEqualStrings("null \"x\" null 7", w.buffered());
}
