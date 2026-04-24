//! `zlsx` — read-only command-line interface over the zlsx library.
//!
//! Streams rows of the selected sheet to stdout in one of four formats.
//! Designed as a drop-in openpyxl replacement for reads: shell-friendly,
//! pipeable into jq / awk, no Python interpreter floor.

const std = @import("std");
const builtin = @import("builtin");
const xlsx = @import("xlsx.zig");

const Format = enum {
    /// NEW default: row envelope `{kind,sheet,sheet_idx,row,cells:[…]}`.
    jsonl,
    /// Bare `[…]` arrays — what iter54's `jsonl` emitted.
    legacy_jsonl,
    /// Bare `{col:val,…}` objects — what iter54's `jsonl-dict` emitted.
    legacy_jsonl_dict,
    tsv,
    csv,
};

const Args = struct {
    file: []const u8,
    sheet_index: ?usize = null,
    sheet_name: ?[]const u8 = null,
    format: Format = .jsonl,
    list_sheets: bool = false,
};

const ArgError = error{
    NoFile,
    HelpRequested,
    UnknownFlag,
    MissingValue,
    BadFormat,
    BadSheetIndex,
    SheetArgConflict,
};

fn parseArgs(argv: []const []const u8) ArgError!Args {
    var out: Args = .{ .file = "" };
    var i: usize = 0;
    while (i < argv.len) : (i += 1) {
        const a = argv[i];
        if (std.mem.eql(u8, a, "-h") or std.mem.eql(u8, a, "--help")) {
            return ArgError.HelpRequested;
        } else if (std.mem.eql(u8, a, "--list-sheets")) {
            out.list_sheets = true;
        } else if (std.mem.eql(u8, a, "--sheet")) {
            if (out.sheet_name != null) return ArgError.SheetArgConflict;
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            out.sheet_index = std.fmt.parseInt(usize, argv[i], 10) catch return ArgError.BadSheetIndex;
        } else if (std.mem.eql(u8, a, "--name")) {
            if (out.sheet_index != null) return ArgError.SheetArgConflict;
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            out.sheet_name = argv[i];
        } else if (std.mem.eql(u8, a, "--format")) {
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            const v = argv[i];
            if (std.mem.eql(u8, v, "jsonl")) {
                out.format = .jsonl;
            } else if (std.mem.eql(u8, v, "legacy-jsonl")) {
                out.format = .legacy_jsonl;
            } else if (std.mem.eql(u8, v, "legacy-jsonl-dict")) {
                out.format = .legacy_jsonl_dict;
            } else if (std.mem.eql(u8, v, "jsonl-dict")) {
                // Deprecated alias for `legacy-jsonl-dict`; kept silent
                // for one release cycle (iter55b will warn). Pre-iter55a
                // the only dict shape we shipped was the bare object.
                out.format = .legacy_jsonl_dict;
            } else if (std.mem.eql(u8, v, "tsv")) {
                out.format = .tsv;
            } else if (std.mem.eql(u8, v, "csv")) {
                out.format = .csv;
            } else return ArgError.BadFormat;
        } else if (a.len > 0 and a[0] == '-') {
            return ArgError.UnknownFlag;
        } else {
            if (out.file.len == 0) out.file = a else return ArgError.UnknownFlag;
        }
    }
    if (out.file.len == 0) return ArgError.NoFile;
    return out;
}

fn writeUsage(w: *std.Io.Writer) !void {
    try w.writeAll(
        \\usage: zlsx <file.xlsx> [options]
        \\
        \\  --sheet N         0-indexed sheet to read (default: 0)
        \\  --name NAME       select sheet by name (conflicts with --sheet)
        \\  --format FMT      jsonl | legacy-jsonl | legacy-jsonl-dict | jsonl-dict | tsv | csv
        \\                    (default: jsonl — NDJSON row envelope; iter55a.
        \\                    `jsonl-dict` is a deprecated alias for
        \\                    `legacy-jsonl-dict` — accepted this release.)
        \\  --list-sheets     print sheet names, one per line, and exit
        \\  -h, --help        show this help
        \\
        \\Formats
        \\  jsonl              NDJSON row envelope (default, iter55a):
        \\                     {"kind":"row","sheet":"S","sheet_idx":0,"row":1,
        \\                      "cells":[{"ref":"A1","col":1,"t":"str","v":"x"},…]}
        \\                     t ∈ {"str","int","num","bool"}; empty cells skipped.
        \\  legacy-jsonl       pre-iter55a bare arrays:  [1, "foo", null, true]
        \\  legacy-jsonl-dict  pre-iter55a bare objects: {"A": 1, "B": "foo"}
        \\                     (alias `jsonl-dict` accepted this release for back-
        \\                     compat; will warn in a future release)
        \\  tsv                tab-separated, \N for empty cells, control chars escaped
        \\  csv                RFC 4180, empty string for empty cells
        \\
        \\Exit codes
        \\  0  success
        \\  1  bad arguments
        \\  2  could not open or parse workbook
        \\  3  sheet not found
        \\
    );
}

fn colLetter(buf: *[8]u8, idx: usize) []const u8 {
    var i: usize = idx + 1; // xlsx columns are 1-based
    var pos: usize = buf.len;
    while (i > 0) {
        i -= 1;
        pos -= 1;
        buf[pos] = 'A' + @as(u8, @intCast(i % 26));
        i /= 26;
    }
    return buf[pos..];
}

fn writeJsonString(w: *std.Io.Writer, s: []const u8) !void {
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

fn writeJsonCell(w: *std.Io.Writer, cell: xlsx.Cell) !void {
    switch (cell) {
        .empty => try w.writeAll("null"),
        .string => |s| try writeJsonString(w, s),
        .integer => |x| try w.print("{d}", .{x}),
        .number => |f| {
            if (std.math.isFinite(f)) {
                try w.print("{d}", .{f});
            } else {
                // JSON has no NaN/Inf — emit null so parsers don't choke.
                try w.writeAll("null");
            }
        },
        .boolean => |b| try w.writeAll(if (b) "true" else "false"),
    }
}

fn writeTsvField(w: *std.Io.Writer, s: []const u8) !void {
    for (s) |c| switch (c) {
        '\t' => try w.writeAll("\\t"),
        '\n' => try w.writeAll("\\n"),
        '\r' => try w.writeAll("\\r"),
        '\\' => try w.writeAll("\\\\"),
        else => try w.writeByte(c),
    };
}

fn writeCsvField(w: *std.Io.Writer, s: []const u8) !void {
    var needs_quote = false;
    for (s) |c| {
        if (c == ',' or c == '"' or c == '\n' or c == '\r') {
            needs_quote = true;
            break;
        }
    }
    if (!needs_quote) {
        try w.writeAll(s);
        return;
    }
    try w.writeByte('"');
    for (s) |c| {
        if (c == '"') try w.writeAll("\"\"") else try w.writeByte(c);
    }
    try w.writeByte('"');
}

/// Per-cell `t` type tag for the envelope schema. Mirrors the
/// design-doc "cells" record but limited to the four primitive
/// types this slice emits — date/formula/error are future work.
fn envelopeTypeTag(cell: xlsx.Cell) []const u8 {
    return switch (cell) {
        .empty => unreachable, // caller skips empties
        .string => "str",
        .integer => "int",
        .number => "num",
        .boolean => "bool",
    };
}

/// Emit just the `[{ref,col,t,v},…]` array. Sparse: `.empty` slots
/// are skipped. `row_number` is the 1-based OOXML row used to build
/// each cell's `ref`.
fn writeEnvelopeCells(
    w: *std.Io.Writer,
    cells: []const xlsx.Cell,
    row_number: u32,
) !void {
    try w.writeByte('[');
    var first = true;
    for (cells, 0..) |c, i| {
        if (c == .empty) continue;
        if (!first) try w.writeByte(',');
        first = false;

        var col_buf: [8]u8 = undefined;
        const letters = colLetter(&col_buf, i);
        var ref_buf: [16]u8 = undefined;
        const ref = std.fmt.bufPrint(&ref_buf, "{s}{d}", .{ letters, row_number }) catch unreachable;

        try w.writeAll("{\"ref\":");
        try writeJsonString(w, ref);
        try w.print(",\"col\":{d},\"t\":\"{s}\",\"v\":", .{ i + 1, envelopeTypeTag(c) });
        try writeJsonCell(w, c);
        try w.writeByte('}');
    }
    try w.writeByte(']');
}

/// Emit one NDJSON envelope line:
/// `{"kind":"row","sheet":…,"sheet_idx":…,"row":…,"cells":[…]}\n`.
/// All-empty rows still emit the envelope with `"cells":[]` so
/// consumers can count rows without a second pass.
fn writeRowEnvelope(
    w: *std.Io.Writer,
    sheet_name: []const u8,
    sheet_idx: usize,
    row_number: u32,
    cells: []const xlsx.Cell,
) !void {
    try w.writeAll("{\"kind\":\"row\",\"sheet\":");
    try writeJsonString(w, sheet_name);
    try w.print(",\"sheet_idx\":{d},\"row\":{d},\"cells\":", .{ sheet_idx, row_number });
    try writeEnvelopeCells(w, cells, row_number);
    try w.writeAll("}\n");
}

/// Legacy emitter — covers the four bare/flat formats. The new
/// envelope format (`.jsonl`) goes through `writeRowEnvelope`, not
/// this function. Calling this with `.jsonl` is a programmer error.
fn writeRow(w: *std.Io.Writer, cells: []const xlsx.Cell, fmt: Format) !void {
    switch (fmt) {
        .jsonl => unreachable, // envelope path; use writeRowEnvelope
        .legacy_jsonl => {
            try w.writeByte('[');
            for (cells, 0..) |c, i| {
                if (i > 0) try w.writeAll(", ");
                try writeJsonCell(w, c);
            }
            try w.writeAll("]\n");
        },
        .legacy_jsonl_dict => {
            try w.writeByte('{');
            var first = true;
            for (cells, 0..) |c, i| {
                if (c == .empty) continue;
                if (!first) try w.writeAll(", ");
                first = false;
                var col_buf: [8]u8 = undefined;
                const col = colLetter(&col_buf, i);
                try w.writeByte('"');
                try w.writeAll(col);
                try w.writeAll("\": ");
                try writeJsonCell(w, c);
            }
            try w.writeAll("}\n");
        },
        .tsv => {
            for (cells, 0..) |c, i| {
                if (i > 0) try w.writeByte('\t');
                switch (c) {
                    .empty => try w.writeAll("\\N"),
                    .string => |s| try writeTsvField(w, s),
                    .integer => |x| try w.print("{d}", .{x}),
                    .number => |f| {
                        if (std.math.isFinite(f)) try w.print("{d}", .{f}) else try w.writeAll("\\N");
                    },
                    .boolean => |b| try w.writeAll(if (b) "true" else "false"),
                }
            }
            try w.writeByte('\n');
        },
        .csv => {
            for (cells, 0..) |c, i| {
                if (i > 0) try w.writeByte(',');
                switch (c) {
                    .empty => {},
                    .string => |s| try writeCsvField(w, s),
                    .integer => |x| try w.print("{d}", .{x}),
                    .number => |f| {
                        if (std.math.isFinite(f)) try w.print("{d}", .{f});
                    },
                    .boolean => |b| try w.writeAll(if (b) "true" else "false"),
                }
            }
            try w.writeByte('\n');
        },
    }
}

pub fn main() !u8 {
    // Debug builds use the leak-detecting allocator; release builds use
    // smp_allocator — fast, pure-Zig (no libc dep). smp_allocator asserts
    // !builtin.single_threaded, so single-threaded builds fall back to
    // page_allocator (also pure-Zig, slightly higher per-alloc cost but
    // fine for short-lived CLIs).
    var gpa: std.heap.DebugAllocator(.{}) = .init;
    defer if (builtin.mode == .Debug) {
        _ = gpa.deinit();
    };
    const release_alloc: std.mem.Allocator = if (builtin.single_threaded)
        std.heap.page_allocator
    else
        std.heap.smp_allocator;
    const alloc = if (builtin.mode == .Debug) gpa.allocator() else release_alloc;

    const raw_args = try std.process.argsAlloc(alloc);
    defer std.process.argsFree(alloc, raw_args);

    var stdout_buf: [16 * 1024]u8 = undefined;
    var stdout_file = std.fs.File.stdout().writer(&stdout_buf);
    const out = &stdout_file.interface;
    defer out.flush() catch {};

    var stderr_buf: [4 * 1024]u8 = undefined;
    var stderr_file = std.fs.File.stderr().writer(&stderr_buf);
    const err = &stderr_file.interface;
    defer err.flush() catch {};

    const args = parseArgs(raw_args[1..]) catch |e| switch (e) {
        ArgError.HelpRequested => {
            try writeUsage(out);
            return 0;
        },
        ArgError.NoFile => {
            try err.writeAll("zlsx: no input file\n\n");
            try writeUsage(err);
            return 1;
        },
        ArgError.UnknownFlag,
        ArgError.MissingValue,
        ArgError.BadFormat,
        ArgError.BadSheetIndex,
        ArgError.SheetArgConflict,
        => {
            try err.print("zlsx: bad arguments ({s})\n\n", .{@errorName(e)});
            try writeUsage(err);
            return 1;
        },
    };

    var book = xlsx.Book.open(alloc, args.file) catch |e| {
        try err.print("zlsx: cannot open '{s}': {s}\n", .{ args.file, @errorName(e) });
        return 2;
    };
    defer book.deinit();

    if (args.list_sheets) {
        for (book.sheets) |s| {
            try out.writeAll(s.name);
            try out.writeByte('\n');
        }
        return 0;
    }

    const Selected = struct { sheet: xlsx.Sheet, idx: usize };
    const selected: Selected = blk: {
        if (args.sheet_name) |n| {
            // Linear scan so we recover the 0-based index too — the
            // envelope emitter needs it per-record.
            for (book.sheets, 0..) |s, i| {
                if (std.mem.eql(u8, s.name, n)) break :blk .{ .sheet = s, .idx = i };
            }
            try err.print("zlsx: no sheet named '{s}'\n", .{n});
            return 3;
        }
        const idx = args.sheet_index orelse 0;
        if (book.sheets.len == 0) {
            try err.writeAll("zlsx: workbook has no sheets\n");
            return 3;
        }
        if (idx >= book.sheets.len) {
            try err.print("zlsx: sheet index {d} out of range (workbook has {d})\n", .{ idx, book.sheets.len });
            return 3;
        }
        break :blk .{ .sheet = book.sheets[idx], .idx = idx };
    };

    var rows = try book.rows(selected.sheet, alloc);
    defer rows.deinit();

    while (try rows.next()) |cells| {
        switch (args.format) {
            .jsonl => try writeRowEnvelope(out, selected.sheet.name, selected.idx, rows.currentRowNumber(), cells),
            else => try writeRow(out, cells, args.format),
        }
    }
    return 0;
}

// ─── Tests ───────────────────────────────────────────────────────────

test "colLetter A,B,Z,AA,AZ,BA,ZZ,AAA" {
    var buf: [8]u8 = undefined;
    try std.testing.expectEqualStrings("A", colLetter(&buf, 0));
    try std.testing.expectEqualStrings("B", colLetter(&buf, 1));
    try std.testing.expectEqualStrings("Z", colLetter(&buf, 25));
    try std.testing.expectEqualStrings("AA", colLetter(&buf, 26));
    try std.testing.expectEqualStrings("AZ", colLetter(&buf, 51));
    try std.testing.expectEqualStrings("BA", colLetter(&buf, 52));
    try std.testing.expectEqualStrings("ZZ", colLetter(&buf, 701));
    try std.testing.expectEqualStrings("AAA", colLetter(&buf, 702));
}

test "writeJsonString escapes" {
    var scratch: [256]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try writeJsonString(&w, "hi\n\"\\\t\x01");
    try std.testing.expectEqualStrings("\"hi\\n\\\"\\\\\\t\\u0001\"", w.buffered());
}

test "writeCsvField quoting" {
    var scratch: [256]u8 = undefined;
    {
        var w = std.Io.Writer.fixed(&scratch);
        try writeCsvField(&w, "plain");
        try std.testing.expectEqualStrings("plain", w.buffered());
    }
    {
        var w = std.Io.Writer.fixed(&scratch);
        try writeCsvField(&w, "has,comma");
        try std.testing.expectEqualStrings("\"has,comma\"", w.buffered());
    }
    {
        var w = std.Io.Writer.fixed(&scratch);
        try writeCsvField(&w, "has\"quote");
        try std.testing.expectEqualStrings("\"has\"\"quote\"", w.buffered());
    }
}

test "parseArgs basics" {
    const argv = [_][]const u8{ "file.xlsx", "--sheet", "2", "--format", "csv" };
    const a = try parseArgs(&argv);
    try std.testing.expectEqualStrings("file.xlsx", a.file);
    try std.testing.expectEqual(@as(?usize, 2), a.sheet_index);
    try std.testing.expectEqual(Format.csv, a.format);
}

test "parseArgs rejects both --sheet and --name" {
    const argv = [_][]const u8{ "f.xlsx", "--sheet", "0", "--name", "Sheet1" };
    try std.testing.expectError(ArgError.SheetArgConflict, parseArgs(&argv));
}

test "parseArgs help" {
    const argv = [_][]const u8{"-h"};
    try std.testing.expectError(ArgError.HelpRequested, parseArgs(&argv));
}

test "parseArgs maps jsonl to envelope and legacy-jsonl to bare array" {
    {
        const argv = [_][]const u8{ "f.xlsx", "--format", "jsonl" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Format.jsonl, a.format);
    }
    {
        const argv = [_][]const u8{ "f.xlsx", "--format", "legacy-jsonl" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Format.legacy_jsonl, a.format);
    }
    {
        const argv = [_][]const u8{ "f.xlsx", "--format", "legacy-jsonl-dict" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Format.legacy_jsonl_dict, a.format);
    }
    {
        // Deprecated alias still lands on the bare-dict path.
        const argv = [_][]const u8{ "f.xlsx", "--format", "jsonl-dict" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Format.legacy_jsonl_dict, a.format);
    }
}

test "writeRowEnvelope emits kind + sheet + sheet_idx + row + sparse cells" {
    var scratch: [512]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    const cells = [_]xlsx.Cell{
        .{ .string = "name" },
        .{ .integer = 42 },
        .empty, // sparse — must be skipped in the cells array
        .{ .number = 3.5 },
        .{ .boolean = true },
    };
    try writeRowEnvelope(&w, "Data", 0, 1, &cells);
    const expected =
        "{\"kind\":\"row\",\"sheet\":\"Data\",\"sheet_idx\":0,\"row\":1,\"cells\":[" ++
        "{\"ref\":\"A1\",\"col\":1,\"t\":\"str\",\"v\":\"name\"}," ++
        "{\"ref\":\"B1\",\"col\":2,\"t\":\"int\",\"v\":42}," ++
        "{\"ref\":\"D1\",\"col\":4,\"t\":\"num\",\"v\":3.5}," ++
        "{\"ref\":\"E1\",\"col\":5,\"t\":\"bool\",\"v\":true}" ++
        "]}\n";
    try std.testing.expectEqualStrings(expected, w.buffered());
}

test "writeRowEnvelope all-empty row emits envelope with empty cells array" {
    var scratch: [256]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    const cells = [_]xlsx.Cell{ .empty, .empty, .empty };
    try writeRowEnvelope(&w, "S", 2, 7, &cells);
    try std.testing.expectEqualStrings(
        "{\"kind\":\"row\",\"sheet\":\"S\",\"sheet_idx\":2,\"row\":7,\"cells\":[]}\n",
        w.buffered(),
    );
}

test "writeRowEnvelope escapes sheet name" {
    var scratch: [256]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    const cells = [_]xlsx.Cell{.{ .integer = 1 }};
    try writeRowEnvelope(&w, "She\"et\n", 0, 1, &cells);
    try std.testing.expectEqualStrings(
        "{\"kind\":\"row\",\"sheet\":\"She\\\"et\\n\",\"sheet_idx\":0,\"row\":1,\"cells\":[" ++
            "{\"ref\":\"A1\",\"col\":1,\"t\":\"int\",\"v\":1}" ++
            "]}\n",
        w.buffered(),
    );
}

test "writeRowEnvelope non-finite number becomes null v" {
    var scratch: [256]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    const cells = [_]xlsx.Cell{.{ .number = std.math.nan(f64) }};
    try writeRowEnvelope(&w, "S", 0, 1, &cells);
    // `t` stays `"num"` for the non-finite case — the type of the
    // cell didn't change, only its JSON-serializable value did.
    // This matches the pre-iter55a behaviour of writeJsonCell.
    try std.testing.expectEqualStrings(
        "{\"kind\":\"row\",\"sheet\":\"S\",\"sheet_idx\":0,\"row\":1,\"cells\":[" ++
            "{\"ref\":\"A1\",\"col\":1,\"t\":\"num\",\"v\":null}" ++
            "]}\n",
        w.buffered(),
    );
}

test "writeRow legacy-jsonl produces bare arrays (regression guard)" {
    var scratch: [256]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    const cells = [_]xlsx.Cell{
        .{ .string = "x" },
        .empty,
        .{ .integer = 9 },
    };
    try writeRow(&w, &cells, .legacy_jsonl);
    try std.testing.expectEqualStrings("[\"x\", null, 9]\n", w.buffered());
}

test "writeRow legacy-jsonl-dict produces bare objects (regression guard)" {
    var scratch: [256]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    const cells = [_]xlsx.Cell{
        .{ .string = "x" },
        .empty,
        .{ .integer = 9 },
    };
    try writeRow(&w, &cells, .legacy_jsonl_dict);
    try std.testing.expectEqualStrings("{\"A\": \"x\", \"C\": 9}\n", w.buffered());
}

// ─── Fuzz tests ──────────────────────────────────────────────────────

fn fuzzItersCli() usize {
    const env = std.process.getEnvVarOwned(std.heap.page_allocator, "XLSX_FUZZ_ITERS") catch return 1_000;
    defer std.heap.page_allocator.free(env);
    var digits: [32]u8 = undefined;
    var di: usize = 0;
    for (env) |c| {
        if (c == '_') continue;
        if (di == digits.len) break;
        digits[di] = c;
        di += 1;
    }
    return std.fmt.parseInt(usize, digits[0..di], 10) catch 1_000;
}

fn fuzzSeedCli() u64 {
    if (std.process.getEnvVarOwned(std.heap.page_allocator, "XLSX_FUZZ_SEED")) |s| {
        defer std.heap.page_allocator.free(s);
        return std.fmt.parseInt(u64, s, 10) catch 0xA1F8ED;
    } else |_| {
        return @bitCast(std.time.milliTimestamp());
    }
}

test "fuzz colLetter: output is uppercase A-Z" {
    const iters = fuzzItersCli();
    var prng = std.Random.DefaultPrng.init(fuzzSeedCli());
    const rng = prng.random();
    var buf: [8]u8 = undefined;
    for (0..iters) |_| {
        // xlsx max is column 16383 (XFD); cap at 2^20 — beyond that the
        // 8-byte buffer can't fit all letters and the function would
        // wrap around via pos underflow. This is documented: caller is
        // expected to stay within OOXML's column range.
        const idx = rng.intRangeAtMost(usize, 0, 1_048_575);
        const letters = colLetter(&buf, idx);
        try std.testing.expect(letters.len >= 1);
        for (letters) |c| {
            try std.testing.expect(c >= 'A' and c <= 'Z');
        }
    }
}

test "fuzz parseArgs: arbitrary tokens never panic" {
    const iters = fuzzItersCli();
    var prng = std.Random.DefaultPrng.init(fuzzSeedCli());
    const rng = prng.random();

    var token_pool: [8][32]u8 = undefined;
    for (0..token_pool.len) |i| rng.bytes(&token_pool[i]);

    for (0..iters) |_| {
        const n_tokens = rng.intRangeAtMost(usize, 0, 12);
        var argv_buf: [12][]const u8 = undefined;
        for (0..n_tokens) |i| {
            const k = rng.intRangeAtMost(usize, 0, token_pool.len - 1);
            const l = rng.intRangeAtMost(usize, 0, token_pool[k].len);
            argv_buf[i] = token_pool[k][0..l];
        }
        // Mix in a few well-known tokens so we hit more branches.
        if (n_tokens >= 1 and rng.boolean()) argv_buf[0] = "--help";
        if (n_tokens >= 2 and rng.boolean()) argv_buf[1] = "--format";

        // Must never panic; errors are fine.
        _ = parseArgs(argv_buf[0..n_tokens]) catch {};
    }
}

test "fuzz writeJsonString: no raw control chars survive" {
    const iters = fuzzItersCli();
    var prng = std.Random.DefaultPrng.init(fuzzSeedCli());
    const rng = prng.random();

    var input: [256]u8 = undefined;
    var scratch: [4096]u8 = undefined;

    for (0..iters) |_| {
        const l = rng.intRangeAtMost(usize, 0, input.len);
        rng.bytes(input[0..l]);
        var w = std.Io.Writer.fixed(&scratch);
        writeJsonString(&w, input[0..l]) catch continue;

        const out = w.buffered();
        try std.testing.expect(out.len >= 2); // at least "\"\""
        try std.testing.expect(out[0] == '"');
        try std.testing.expect(out[out.len - 1] == '"');

        // Walk the interior (between the outer quotes). No bare control
        // char (0..0x1f) except when preceded by a backslash. Quote and
        // backslash always escaped too.
        var i: usize = 1;
        while (i < out.len - 1) : (i += 1) {
            const c = out[i];
            if (c == '\\') {
                // Skip the escape sequence (\", \\, \n, \r, \t, \b, \f, \uXXXX).
                i += 1;
                if (i < out.len - 1 and out[i] == 'u') i += 4;
                continue;
            }
            try std.testing.expect(c >= 0x20);
            try std.testing.expect(c != '"');
        }
    }
}

test "fuzz writeCsvField: balanced quotes + no bare quote outside them" {
    const iters = fuzzItersCli();
    var prng = std.Random.DefaultPrng.init(fuzzSeedCli());
    const rng = prng.random();

    var input: [256]u8 = undefined;
    var scratch: [4096]u8 = undefined;

    for (0..iters) |_| {
        const l = rng.intRangeAtMost(usize, 0, input.len);
        rng.bytes(input[0..l]);
        var w = std.Io.Writer.fixed(&scratch);
        writeCsvField(&w, input[0..l]) catch continue;

        const out = w.buffered();
        // If any RFC-4180 trigger byte was present, output must be
        // quoted. Otherwise unquoted is fine.
        var needs_quote = false;
        for (input[0..l]) |c| {
            if (c == ',' or c == '"' or c == '\n' or c == '\r') {
                needs_quote = true;
                break;
            }
        }
        if (needs_quote) {
            try std.testing.expect(out.len >= 2);
            try std.testing.expectEqual(@as(u8, '"'), out[0]);
            try std.testing.expectEqual(@as(u8, '"'), out[out.len - 1]);
            // Every `"` inside must be doubled.
            var i: usize = 1;
            while (i < out.len - 1) : (i += 1) {
                if (out[i] == '"') {
                    try std.testing.expect(i + 1 < out.len - 1 and out[i + 1] == '"');
                    i += 1;
                }
            }
        }
    }
}
