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

/// iter56/57: first positional decides sub-command. `rows` is the
/// legacy envelope-row emitter; `cells` is the per-cell NDJSON stream;
/// `meta` emits a workbook record followed by per-sheet records;
/// `list_sheets` is the lighter NDJSON variant. Bare
/// `zlsx file.xlsx` (no sub-command token) still means `rows` so
/// existing scripts keep working — the short-alias re-point to `cells`
/// is an iter60+ breaking change with its own rollout.
const Subcommand = enum { rows, cells, meta, list_sheets };

const Args = struct {
    subcommand: Subcommand = .rows,
    file: []const u8,
    sheet_index: ?usize = null,
    sheet_name: ?[]const u8 = null,
    format: Format = .jsonl,
    list_sheets: bool = false,
    /// Set when the user passed the deprecated `--format jsonl-dict`
    /// spelling. `main` emits a one-line stderr deprecation warning
    /// so existing scripts keep working while their authors learn
    /// about the rename.
    deprecated_jsonl_dict: bool = false,
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
    // First positional may be a sub-command token. We only consume it
    // as a sub-command when it matches exactly; anything else (including
    // a filename that happens to be our first positional) falls through
    // to the file-path slot and defaults the sub-command to `rows`.
    var first_positional_seen = false;
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
                // Deprecated alias for `legacy-jsonl-dict` — routed
                // through the deprecation flag so `main` emits one
                // stderr warning. Pre-iter55a the only dict shape we
                // shipped was the bare object, so the intent is clear.
                out.format = .legacy_jsonl_dict;
                out.deprecated_jsonl_dict = true;
            } else if (std.mem.eql(u8, v, "tsv")) {
                out.format = .tsv;
            } else if (std.mem.eql(u8, v, "csv")) {
                out.format = .csv;
            } else return ArgError.BadFormat;
        } else if (a.len > 0 and a[0] == '-') {
            return ArgError.UnknownFlag;
        } else {
            if (!first_positional_seen) {
                first_positional_seen = true;
                if (std.mem.eql(u8, a, "cells")) {
                    out.subcommand = .cells;
                    continue;
                } else if (std.mem.eql(u8, a, "rows")) {
                    out.subcommand = .rows;
                    continue;
                } else if (std.mem.eql(u8, a, "meta")) {
                    out.subcommand = .meta;
                    continue;
                } else if (std.mem.eql(u8, a, "list-sheets")) {
                    out.subcommand = .list_sheets;
                    continue;
                }
                // Fall through: first positional is the file path,
                // sub-command stays at its default (`rows`).
            }
            if (out.file.len == 0) out.file = a else return ArgError.UnknownFlag;
        }
    }
    if (out.file.len == 0) return ArgError.NoFile;
    return out;
}

fn writeUsage(w: *std.Io.Writer) !void {
    try w.writeAll(
        \\usage: zlsx [<subcommand>] <file.xlsx> [options]
        \\
        \\  --sheet N         0-indexed sheet to read (default: 0)
        \\  --name NAME       select sheet by name (conflicts with --sheet)
        \\  --format FMT      jsonl | legacy-jsonl | legacy-jsonl-dict | jsonl-dict | tsv | csv
        \\                    (default: jsonl — NDJSON row envelope; iter55a.
        \\                    Applies to the `rows` sub-command only; ignored
        \\                    by `cells`, which always emits per-cell NDJSON.
        \\                    `jsonl-dict` is a deprecated alias for
        \\                    `legacy-jsonl-dict` — accepted this release.)
        \\  --list-sheets     print sheet names, one per line, and exit
        \\                    (legacy plain-text flag — still works.
        \\                    The `list-sheets` sub-command emits NDJSON.)
        \\  -h, --help        show this help
        \\
        \\Sub-commands
        \\  rows               (default) one NDJSON envelope per row — see Formats.
        \\                     Bare `zlsx file.xlsx` is an alias for `zlsx rows file.xlsx`.
        \\  cells              one NDJSON record per non-empty cell (iter56):
        \\                     {"kind":"cell","sheet":"S","sheet_idx":0,"ref":"A1",
        \\                      "row":1,"col":1,"t":"str","v":"x"}
        \\                     t ∈ {"str","int","num","bool"}. Empty cells skipped.
        \\                     --format is ignored; output shape is fixed.
        \\  meta               workbook summary as NDJSON (iter57). One
        \\                     workbook record first, then one sheet record per sheet:
        \\                     {"kind":"workbook","path":"f.xlsx","sheets":N,
        \\                      "sst":{"count":C,"rich":R},
        \\                      "has_styles":bool,"has_theme":bool,"has_comments":bool}
        \\                     {"kind":"sheet","sheet":"S","sheet_idx":0,
        \\                      "has_comments":bool}
        \\                     --format / --sheet / --name are ignored.
        \\  list-sheets        lighter NDJSON variant of `meta` (iter57):
        \\                     one {"kind":"sheet","sheet":…,"sheet_idx":…}
        \\                     record per sheet. For the plain-text one-name-
        \\                     per-line shape, use the legacy `--list-sheets` flag.
        \\
        \\Formats (rows only)
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

/// iter56: emit one NDJSON record for a single cell, matching the
/// `cells` sub-command wire format:
/// `{"kind":"cell","sheet":…,"sheet_idx":…,"ref":…,"row":…,"col":…,"t":…,"v":…}\n`.
/// Empty cells are a caller-skip: this function asserts out if handed
/// one (matches envelope semantics — sparse cells are suppressed at
/// the record level, not materialised as `v:null`).
fn writeCell(
    w: *std.Io.Writer,
    sheet_name: []const u8,
    sheet_idx: usize,
    ref: []const u8,
    row: u32,
    col: u32,
    cell: xlsx.Cell,
) !void {
    std.debug.assert(cell != .empty); // caller must skip empties
    try w.writeAll("{\"kind\":\"cell\",\"sheet\":");
    try writeJsonString(w, sheet_name);
    try w.print(",\"sheet_idx\":{d},\"ref\":", .{sheet_idx});
    try writeJsonString(w, ref);
    try w.print(",\"row\":{d},\"col\":{d},\"t\":\"{s}\",\"v\":", .{
        row,
        col,
        envelopeTypeTag(cell),
    });
    try writeJsonCell(w, cell);
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

    if (args.deprecated_jsonl_dict) {
        try err.writeAll(
            "zlsx: --format jsonl-dict is deprecated, use --format legacy-jsonl-dict (this alias will be removed in a future release)\n",
        );
        try err.flush();
    }

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

    // iter57 sub-commands — workbook-scoped, no per-sheet selection.
    switch (args.subcommand) {
        .meta => {
            try runMetaCommand(out, &book, args.file);
            return 0;
        },
        .list_sheets => {
            try runListSheetsCommand(out, &book);
            return 0;
        },
        .rows, .cells => {},
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

    switch (args.subcommand) {
        .rows => try runRowsCommand(out, &book, selected.sheet, selected.idx, args.format, alloc),
        .cells => try runCellsCommand(out, &book, selected.sheet, selected.idx, alloc),
        // Handled by the workbook-scoped early return above.
        .meta, .list_sheets => unreachable,
    }
    return 0;
}

fn runRowsCommand(
    out: *std.Io.Writer,
    book: *xlsx.Book,
    sheet: xlsx.Sheet,
    sheet_idx: usize,
    format: Format,
    alloc: std.mem.Allocator,
) !void {
    var rows = try book.rows(sheet, alloc);
    defer rows.deinit();

    while (try rows.next()) |cells| {
        switch (format) {
            .jsonl => try writeRowEnvelope(out, sheet.name, sheet_idx, rows.currentRowNumber(), cells),
            else => try writeRow(out, cells, format),
        }
    }
}

/// iter56: stream one NDJSON record per non-empty cell of the selected
/// sheet. Empty cells are suppressed (matches envelope semantics on
/// the rows path). `--format` is intentionally ignored here — the
/// `cells` sub-command has a single fixed wire shape.
fn runCellsCommand(
    out: *std.Io.Writer,
    book: *xlsx.Book,
    sheet: xlsx.Sheet,
    sheet_idx: usize,
    alloc: std.mem.Allocator,
) !void {
    var rows = try book.rows(sheet, alloc);
    defer rows.deinit();

    while (try rows.next()) |cells| {
        const row_number = rows.currentRowNumber();
        for (cells, 0..) |c, i| {
            if (c == .empty) continue;

            var col_buf: [8]u8 = undefined;
            const letters = colLetter(&col_buf, i);
            var ref_buf: [16]u8 = undefined;
            const ref = std.fmt.bufPrint(&ref_buf, "{s}{d}", .{ letters, row_number }) catch unreachable;

            try writeCell(
                out,
                sheet.name,
                sheet_idx,
                ref,
                row_number,
                @intCast(i + 1),
                c,
            );
        }
    }
}

/// iter57: emit the workbook record followed by one sheet record per
/// sheet. Fields deliberately limited to ones that are O(1) over the
/// reader APIs Book already exposes — `rows` / `cols` / `first_cell` /
/// `last_cell` / `format_version` are follow-up work (they need
/// sheet-iteration or version plumbing) per the iter57 scope note.
fn runMetaCommand(
    out: *std.Io.Writer,
    book: *const xlsx.Book,
    path: []const u8,
) !void {
    // Workbook-level `has_comments` is the OR across every sheet —
    // saves callers a reduce step when they only want "does this file
    // have any comments at all?".
    var any_comments = false;
    for (book.sheets) |s| {
        if (book.comments(s).len != 0) {
            any_comments = true;
            break;
        }
    }

    try out.writeAll("{\"kind\":\"workbook\",\"path\":");
    try writeJsonString(out, path);
    try out.print(
        ",\"sheets\":{d},\"sst\":{{\"count\":{d},\"rich\":{d}}}",
        .{ book.sheets.len, book.shared_strings.len, book.rich_runs_by_sst_idx.count() },
    );
    try out.print(
        ",\"has_styles\":{s},\"has_theme\":{s},\"has_comments\":{s}}}\n",
        .{
            if (book.styles_xml != null) "true" else "false",
            if (book.theme_xml != null) "true" else "false",
            if (any_comments) "true" else "false",
        },
    );

    for (book.sheets, 0..) |s, i| {
        const sheet_has_comments = book.comments(s).len != 0;
        try out.writeAll("{\"kind\":\"sheet\",\"sheet\":");
        try writeJsonString(out, s.name);
        try out.print(
            ",\"sheet_idx\":{d},\"has_comments\":{s}}}\n",
            .{ i, if (sheet_has_comments) "true" else "false" },
        );
    }
}

/// iter57: lighter NDJSON variant of `meta` — one record per sheet,
/// name + index only. Same envelope shape as `meta`'s sheet record
/// minus the workbook-scoped `has_comments` field, so consumers can
/// trivially swap between the two commands.
fn runListSheetsCommand(out: *std.Io.Writer, book: *const xlsx.Book) !void {
    for (book.sheets, 0..) |s, i| {
        try out.writeAll("{\"kind\":\"sheet\",\"sheet\":");
        try writeJsonString(out, s.name);
        try out.print(",\"sheet_idx\":{d}}}\n", .{i});
    }
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
        // Deprecated alias still lands on the bare-dict path AND
        // flips the deprecation flag for main's stderr warning.
        const argv = [_][]const u8{ "f.xlsx", "--format", "jsonl-dict" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Format.legacy_jsonl_dict, a.format);
        try std.testing.expect(a.deprecated_jsonl_dict);
    }
    {
        // Canonical `legacy-jsonl-dict` must NOT trip the warning.
        const argv = [_][]const u8{ "f.xlsx", "--format", "legacy-jsonl-dict" };
        const a = try parseArgs(&argv);
        try std.testing.expect(!a.deprecated_jsonl_dict);
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

test "parseArgs routes 'cells' as the cells sub-command" {
    // Bare file-path defaults to rows (back-compat).
    {
        const argv = [_][]const u8{"file.xlsx"};
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.rows, a.subcommand);
        try std.testing.expectEqualStrings("file.xlsx", a.file);
    }
    // Explicit `rows` is parsed as rows, file-path is the next positional.
    {
        const argv = [_][]const u8{ "rows", "file.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.rows, a.subcommand);
        try std.testing.expectEqualStrings("file.xlsx", a.file);
    }
    // `cells` flips the sub-command.
    {
        const argv = [_][]const u8{ "cells", "file.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.cells, a.subcommand);
        try std.testing.expectEqualStrings("file.xlsx", a.file);
    }
    // `cells` with flags behind it.
    {
        const argv = [_][]const u8{ "cells", "file.xlsx", "--sheet", "2" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.cells, a.subcommand);
        try std.testing.expectEqualStrings("file.xlsx", a.file);
        try std.testing.expectEqual(@as(?usize, 2), a.sheet_index);
    }
    // Flags before the sub-command still work — first POSITIONAL is
    // what decides, not the first argv slot.
    {
        const argv = [_][]const u8{ "--sheet", "1", "cells", "file.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.cells, a.subcommand);
        try std.testing.expectEqualStrings("file.xlsx", a.file);
    }
}

test "writeCell emits kind + sheet + sheet_idx + ref + row + col + t + v" {
    var scratch: [512]u8 = undefined;

    // string
    {
        var w = std.Io.Writer.fixed(&scratch);
        try writeCell(&w, "Data", 0, "A1", 1, 1, .{ .string = "name" });
        try std.testing.expectEqualStrings(
            "{\"kind\":\"cell\",\"sheet\":\"Data\",\"sheet_idx\":0,\"ref\":\"A1\",\"row\":1,\"col\":1,\"t\":\"str\",\"v\":\"name\"}\n",
            w.buffered(),
        );
    }
    // integer
    {
        var w = std.Io.Writer.fixed(&scratch);
        try writeCell(&w, "Data", 0, "B2", 2, 2, .{ .integer = 3 });
        try std.testing.expectEqualStrings(
            "{\"kind\":\"cell\",\"sheet\":\"Data\",\"sheet_idx\":0,\"ref\":\"B2\",\"row\":2,\"col\":2,\"t\":\"int\",\"v\":3}\n",
            w.buffered(),
        );
    }
    // number
    {
        var w = std.Io.Writer.fixed(&scratch);
        try writeCell(&w, "Data", 0, "C3", 3, 3, .{ .number = 3.5 });
        try std.testing.expectEqualStrings(
            "{\"kind\":\"cell\",\"sheet\":\"Data\",\"sheet_idx\":0,\"ref\":\"C3\",\"row\":3,\"col\":3,\"t\":\"num\",\"v\":3.5}\n",
            w.buffered(),
        );
    }
    // boolean
    {
        var w = std.Io.Writer.fixed(&scratch);
        try writeCell(&w, "Data", 0, "D4", 4, 4, .{ .boolean = true });
        try std.testing.expectEqualStrings(
            "{\"kind\":\"cell\",\"sheet\":\"Data\",\"sheet_idx\":0,\"ref\":\"D4\",\"row\":4,\"col\":4,\"t\":\"bool\",\"v\":true}\n",
            w.buffered(),
        );
    }
}

test "writeCell escapes sheet name" {
    var scratch: [256]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try writeCell(&w, "She\"et\n", 2, "A1", 1, 1, .{ .integer = 7 });
    try std.testing.expectEqualStrings(
        "{\"kind\":\"cell\",\"sheet\":\"She\\\"et\\n\",\"sheet_idx\":2,\"ref\":\"A1\",\"row\":1,\"col\":1,\"t\":\"int\",\"v\":7}\n",
        w.buffered(),
    );
}

test "cells loop skips empty cells from the stream" {
    // Mirrors runCellsCommand's inner loop: feed a mixed row, confirm
    // only non-empty cells surface, refs are built from (col,row_number).
    var scratch: [512]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);

    const cells = [_]xlsx.Cell{
        .{ .string = "name" }, // A1
        .empty, // B1 — must produce no output
        .{ .integer = 42 }, // C1
        .empty, // D1 — must produce no output
        .{ .boolean = false }, // E1
    };
    const row_number: u32 = 1;
    for (cells, 0..) |c, i| {
        if (c == .empty) continue;
        var col_buf: [8]u8 = undefined;
        const letters = colLetter(&col_buf, i);
        var ref_buf: [16]u8 = undefined;
        const ref = std.fmt.bufPrint(&ref_buf, "{s}{d}", .{ letters, row_number }) catch unreachable;
        try writeCell(&w, "S", 0, ref, row_number, @intCast(i + 1), c);
    }

    try std.testing.expectEqualStrings(
        "{\"kind\":\"cell\",\"sheet\":\"S\",\"sheet_idx\":0,\"ref\":\"A1\",\"row\":1,\"col\":1,\"t\":\"str\",\"v\":\"name\"}\n" ++
            "{\"kind\":\"cell\",\"sheet\":\"S\",\"sheet_idx\":0,\"ref\":\"C1\",\"row\":1,\"col\":3,\"t\":\"int\",\"v\":42}\n" ++
            "{\"kind\":\"cell\",\"sheet\":\"S\",\"sheet_idx\":0,\"ref\":\"E1\",\"row\":1,\"col\":5,\"t\":\"bool\",\"v\":false}\n",
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

test "parseArgs routes 'meta' and 'list-sheets' correctly" {
    // `meta` as first positional.
    {
        const argv = [_][]const u8{ "meta", "file.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.meta, a.subcommand);
        try std.testing.expectEqualStrings("file.xlsx", a.file);
    }
    // `list-sheets` as first positional flips the sub-command.
    {
        const argv = [_][]const u8{ "list-sheets", "file.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.list_sheets, a.subcommand);
        try std.testing.expectEqualStrings("file.xlsx", a.file);
    }
    // Sub-command token AFTER flags still works (positional decides).
    {
        const argv = [_][]const u8{ "--sheet", "1", "meta", "file.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.meta, a.subcommand);
    }
    // Legacy `--list-sheets` flag is NOT the `list-sheets` sub-command.
    // The flag flips `list_sheets` (legacy plain text), not `subcommand`.
    {
        const argv = [_][]const u8{ "--list-sheets", "file.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expect(a.list_sheets);
        try std.testing.expectEqual(Subcommand.rows, a.subcommand);
    }
}

test "runListSheetsCommand emits one sheet record per sheet" {
    const tmp_path = "/tmp/zlsx_cli_list_sheets_iter57.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{.{ .string = "hdr" }});
        var s1 = try w.addSheet("Other");
        try s1.writeRow(&.{.{ .integer = 1 }});
        var s2 = try w.addSheet("She\"et"); // name with a quote — must be JSON-escaped
        try s2.writeRow(&.{.{ .boolean = true }});
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    var scratch: [1024]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runListSheetsCommand(&w, &book);
    try std.testing.expectEqualStrings(
        "{\"kind\":\"sheet\",\"sheet\":\"Data\",\"sheet_idx\":0}\n" ++
            "{\"kind\":\"sheet\",\"sheet\":\"Other\",\"sheet_idx\":1}\n" ++
            "{\"kind\":\"sheet\",\"sheet\":\"She\\\"et\",\"sheet_idx\":2}\n",
        w.buffered(),
    );
}

test "runMetaCommand emits workbook record with sst/has_* fields then sheet records" {
    const tmp_path = "/tmp/zlsx_cli_meta_iter57.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        // Two distinct strings + one repeat → SST count of 2,
        // exercises the deduped path.
        try s0.writeRow(&.{ .{ .string = "alpha" }, .{ .string = "beta" } });
        try s0.writeRow(&.{.{ .string = "alpha" }});
        try s0.addComment("A1", "me", "hi there"); // forces has_comments=true for this sheet
        var s1 = try w.addSheet("NoComments");
        try s1.writeRow(&.{.{ .integer = 42 }});
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    var scratch: [4096]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runMetaCommand(&w, &book, tmp_path);

    const out = w.buffered();
    // Parse NDJSON line by line and assert field presence + values.
    var line_it = std.mem.splitScalar(u8, out, '\n');

    const wb_line = line_it.next() orelse return error.MissingWorkbookLine;
    // Structural probes — avoid order-sensitive equality because the
    // exact field ordering is an implementation detail the wire format
    // only loosely pins down. We pin the presence + values.
    try std.testing.expect(std.mem.indexOf(u8, wb_line, "\"kind\":\"workbook\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, wb_line, "\"sheets\":2") != null);
    try std.testing.expect(std.mem.indexOf(u8, wb_line, "\"sst\":{\"count\":2,\"rich\":0}") != null);
    // has_styles / has_theme reflect whether the writer chose to emit
    // those parts — we only pin field *presence* here, not the writer's
    // part-emission policy. The workbook-scoped `has_comments` is
    // deterministic given the addComment call above.
    try std.testing.expect(
        std.mem.indexOf(u8, wb_line, "\"has_styles\":true") != null or
            std.mem.indexOf(u8, wb_line, "\"has_styles\":false") != null,
    );
    try std.testing.expect(
        std.mem.indexOf(u8, wb_line, "\"has_theme\":true") != null or
            std.mem.indexOf(u8, wb_line, "\"has_theme\":false") != null,
    );
    try std.testing.expect(std.mem.indexOf(u8, wb_line, "\"has_comments\":true") != null);
    try std.testing.expect(std.mem.indexOf(u8, wb_line, "\"path\":") != null);

    const sheet0 = line_it.next() orelse return error.MissingSheet0;
    try std.testing.expect(std.mem.indexOf(u8, sheet0, "\"kind\":\"sheet\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet0, "\"sheet\":\"Data\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet0, "\"sheet_idx\":0") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet0, "\"has_comments\":true") != null);

    const sheet1 = line_it.next() orelse return error.MissingSheet1;
    try std.testing.expect(std.mem.indexOf(u8, sheet1, "\"sheet\":\"NoComments\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet1, "\"sheet_idx\":1") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet1, "\"has_comments\":false") != null);

    // Trailing empty token after the final '\n' — but no more records.
    const trailing = line_it.next();
    if (trailing) |t| try std.testing.expectEqualStrings("", t);
    try std.testing.expectEqual(@as(?[]const u8, null), line_it.next());
}

test "legacy --list-sheets flag still emits plain text (regression guard)" {
    // Regression guard: the legacy plain-text shape is exactly
    // `<name>\n` per sheet, no JSON, no sub-command routing. This
    // mirrors the code path in main() line-for-line so the flag
    // keeps working across iter57.
    const tmp_path = "/tmp/zlsx_cli_legacy_list_sheets.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{.{ .string = "x" }});
        var s1 = try w.addSheet("More");
        try s1.writeRow(&.{.{ .string = "y" }});
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    var scratch: [256]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    for (book.sheets) |s| {
        try w.writeAll(s.name);
        try w.writeByte('\n');
    }
    try std.testing.expectEqualStrings("Data\nMore\n", w.buffered());
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
