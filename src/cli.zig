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

/// iter56/57/58: first positional decides sub-command. `rows` is the
/// legacy envelope-row emitter; `cells` is the per-cell NDJSON stream;
/// `meta` emits a workbook record followed by per-sheet records;
/// `list_sheets` is the lighter NDJSON variant. iter58 adds the
/// five-way reader-surface exposure: `comments` / `validations` /
/// `hyperlinks` iterate every sheet (sheet-scoped records); `styles`
/// / `sst` are workbook-wide. Bare `zlsx file.xlsx` (no sub-command
/// token) still means `rows` so existing scripts keep working — the
/// short-alias re-point to `cells` is an iter60+ breaking change with
/// its own rollout.
const Subcommand = enum {
    rows,
    cells,
    meta,
    list_sheets,
    comments,
    validations,
    hyperlinks,
    styles,
    sst,
};

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
    /// iter59a: stream-native pagination over the emitted-record
    /// stream (rows / cells / comments / validations / hyperlinks /
    /// styles / sst). Both are applied GLOBALLY after sheet selection.
    skip: ?usize = null,
    take: ?usize = null,
    /// iter59b-1: per-sheet row-bounded filtering on the three
    /// sub-commands that emit row-keyed records (rows / cells /
    /// comments). Both endpoints are 1-based OOXML row numbers and
    /// inclusive: `start_row=3, end_row=5` emits rows 3, 4, 5.
    /// Applied BEFORE --skip/--take, so --skip counts post-row-filter
    /// records per the jq-for-excel design doc.
    start_row: ?u32 = null,
    end_row: ?u32 = null,
};

const ArgError = error{
    NoFile,
    HelpRequested,
    UnknownFlag,
    MissingValue,
    BadFormat,
    BadSheetIndex,
    BadArgValue,
    SheetArgConflict,
};

/// First-pass scan: identify the sub-command without validating
/// flag values. Lets the main pass relax --sheet / --name / --format
/// validation for workbook-scoped sub-commands that wrappers may
/// append those flags to universally. Skips `--sheet` / `--name` /
/// `--format` pairs so their values aren't mistaken for positionals.
fn detectSubcommand(argv: []const []const u8) Subcommand {
    var i: usize = 0;
    while (i < argv.len) : (i += 1) {
        const a = argv[i];
        if (std.mem.eql(u8, a, "--sheet") or
            std.mem.eql(u8, a, "--name") or
            std.mem.eql(u8, a, "--format") or
            std.mem.eql(u8, a, "--skip") or
            std.mem.eql(u8, a, "--take") or
            std.mem.eql(u8, a, "--start-row") or
            std.mem.eql(u8, a, "--end-row"))
        {
            i += 1; // skip paired value (bounds-checked by caller)
            continue;
        }
        if (a.len > 0 and a[0] == '-') continue; // flag with no value
        if (std.mem.eql(u8, a, "cells")) return .cells;
        if (std.mem.eql(u8, a, "rows")) return .rows;
        if (std.mem.eql(u8, a, "meta")) return .meta;
        if (std.mem.eql(u8, a, "list-sheets")) return .list_sheets;
        if (std.mem.eql(u8, a, "comments")) return .comments;
        if (std.mem.eql(u8, a, "validations")) return .validations;
        if (std.mem.eql(u8, a, "hyperlinks")) return .hyperlinks;
        if (std.mem.eql(u8, a, "styles")) return .styles;
        if (std.mem.eql(u8, a, "sst")) return .sst;
        return .rows; // first positional is the file path
    }
    return .rows;
}

fn parseArgs(argv: []const []const u8) ArgError!Args {
    const detected_sub = detectSubcommand(argv);
    // Workbook-scoped commands don't consume --sheet / --name /
    // --format, so wrappers that always append those flags should
    // not hit a hard error. Parse them tolerantly: missing-value is
    // still an error (user typo), but a malformed value is silently
    // dropped. Non-workbook commands keep strict validation.
    //
    // iter58: the three sheet-scoped newcomers (`comments` /
    // `validations` / `hyperlinks`) iterate every sheet by default,
    // so they join this group for flag tolerance even though their
    // records do carry `sheet` / `sheet_idx`. Narrowing via `--sheet`
    // is deferred to iter58-follow-up.
    const workbook_scoped = switch (detected_sub) {
        .meta,
        .list_sheets,
        .styles,
        .sst,
        => true,
        .rows, .cells, .comments, .validations, .hyperlinks => false,
    };

    var out: Args = .{ .file = "", .subcommand = detected_sub };
    var first_positional_seen = false;
    var i: usize = 0;
    while (i < argv.len) : (i += 1) {
        const a = argv[i];
        if (std.mem.eql(u8, a, "-h") or std.mem.eql(u8, a, "--help")) {
            return ArgError.HelpRequested;
        } else if (std.mem.eql(u8, a, "--list-sheets")) {
            out.list_sheets = true;
        } else if (std.mem.eql(u8, a, "--sheet")) {
            if (out.sheet_name != null and !workbook_scoped) return ArgError.SheetArgConflict;
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            const parsed = std.fmt.parseInt(usize, argv[i], 10) catch {
                if (workbook_scoped) continue; // ignore bad value for meta/list-sheets
                return ArgError.BadSheetIndex;
            };
            out.sheet_index = parsed;
        } else if (std.mem.eql(u8, a, "--name")) {
            if (out.sheet_index != null and !workbook_scoped) return ArgError.SheetArgConflict;
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
            } else {
                if (workbook_scoped) continue; // ignore unknown format for meta/list-sheets
                return ArgError.BadFormat;
            }
        } else if (std.mem.eql(u8, a, "--skip")) {
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            // --skip / --take are strict for EVERY sub-command
            // (unlike --sheet / --format whose tolerance depends on
            // workbook_scoped). Pagination is too useful on styles /
            // sst — those commands dump huge streams and a typoed
            // --take that silently returned everything would be a
            // very expensive surprise. For meta / list-sheets which
            // don't paginate, rejecting a --skip typo is also the
            // clearer user-signal: the flag is not effective there.
            out.skip = std.fmt.parseInt(usize, argv[i], 10) catch return ArgError.BadArgValue;
        } else if (std.mem.eql(u8, a, "--take")) {
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            out.take = std.fmt.parseInt(usize, argv[i], 10) catch return ArgError.BadArgValue;
        } else if (std.mem.eql(u8, a, "--start-row")) {
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            // Strict on every sub-command (same rationale as --skip/--take):
            // silently dropping a typoed row bound is an expensive surprise.
            // OOXML rows are 1-based; 0 is a user error and we reject it.
            const v = std.fmt.parseInt(u32, argv[i], 10) catch return ArgError.BadArgValue;
            if (v == 0) return ArgError.BadArgValue;
            out.start_row = v;
        } else if (std.mem.eql(u8, a, "--end-row")) {
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            const v = std.fmt.parseInt(u32, argv[i], 10) catch return ArgError.BadArgValue;
            if (v == 0) return ArgError.BadArgValue;
            out.end_row = v;
        } else if (a.len > 0 and a[0] == '-') {
            return ArgError.UnknownFlag;
        } else {
            if (!first_positional_seen) {
                first_positional_seen = true;
                // Sub-command token already handled by detectSubcommand.
                // Skip it here so it isn't mistaken for the file path.
                if (std.mem.eql(u8, a, "cells") or
                    std.mem.eql(u8, a, "rows") or
                    std.mem.eql(u8, a, "meta") or
                    std.mem.eql(u8, a, "list-sheets") or
                    std.mem.eql(u8, a, "comments") or
                    std.mem.eql(u8, a, "validations") or
                    std.mem.eql(u8, a, "hyperlinks") or
                    std.mem.eql(u8, a, "styles") or
                    std.mem.eql(u8, a, "sst"))
                {
                    continue;
                }
            }
            if (out.file.len == 0) out.file = a else return ArgError.UnknownFlag;
        }
    }
    if (out.file.len == 0) return ArgError.NoFile;

    // iter59b-1: --start-row / --end-row only map to sub-commands
    // that emit row-keyed records (rows / cells / comments). The
    // range-keyed commands (validations / hyperlinks) and the
    // workbook-scoped commands (meta / list-sheets / styles / sst)
    // have no per-record row number — reject the flag rather than
    // silently ignoring it.
    if (out.start_row != null or out.end_row != null) {
        switch (detected_sub) {
            .rows, .cells, .comments => {},
            .validations, .hyperlinks, .meta, .list_sheets, .styles, .sst => {
                return ArgError.BadArgValue;
            },
        }
    }
    // Empty emission ranges are caught at parse time — `start > end`
    // can never produce a record, which is almost certainly a typo.
    if (out.start_row) |s| if (out.end_row) |e| {
        if (s > e) return ArgError.BadArgValue;
    };
    // The legacy --list-sheets flag takes an early return in main
    // and emits plain sheet names — no row concept. Row bounds
    // passed alongside it would silently no-op, hiding typos.
    if (out.list_sheets and (out.start_row != null or out.end_row != null)) {
        return ArgError.BadArgValue;
    }
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
        \\  --skip N          drop the first N emitted records (iter59a).
        \\                    Applies globally to the record stream of
        \\                    rows / cells / comments / validations /
        \\                    hyperlinks / styles / sst. Ignored by meta
        \\                    and list-sheets.
        \\  --take N          stop after N emitted records. Same scope
        \\                    as --skip; combine for middle-slice paging.
        \\  --start-row R     (iter59b) 1-based OOXML row; drop records
        \\                    whose row < R. Per-sheet scope (each
        \\                    sheet's own rows, unlike --skip which is
        \\                    global). Valid for rows / cells / comments
        \\                    only; rejected on validations / hyperlinks
        \\                    / meta / list-sheets / styles / sst.
        \\  --end-row R       (iter59b) 1-based OOXML row; stop emitting
        \\                    after row R (inclusive). Same scope and
        \\                    sub-command constraints as --start-row.
        \\                    Applied BEFORE --skip / --take.
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
        \\  comments           one NDJSON record per cell comment across every
        \\                     sheet (iter58):
        \\                     {"kind":"comment","sheet":"S","sheet_idx":0,
        \\                      "ref":"A1","row":1,"col":1,"author":"Alice",
        \\                      "text":"…","runs":null}
        \\  validations        one NDJSON record per data-validation range
        \\                     across every sheet (iter58):
        \\                     {"kind":"validation","sheet":"S","sheet_idx":0,
        \\                      "range":"B2:B100","rule_type":"list","op":null,
        \\                      "formula1":"a,b","formula2":null,
        \\                      "values":["a","b"]}
        \\  hyperlinks         one NDJSON record per hyperlink across every
        \\                     sheet (iter58):
        \\                     {"kind":"hyperlink","sheet":"S","sheet_idx":0,
        \\                      "range":"A1","url":"https://…","location":null}
        \\  styles             one NDJSON record per cell-XF style entry
        \\                     (workbook-wide, iter58):
        \\                     {"kind":"style","idx":0,"font":{…}|null,
        \\                      "fill":{…}|null,"border":{…}|null,
        \\                      "num_fmt":"General"|null}
        \\  sst                one NDJSON record per shared-string entry
        \\                     (workbook-wide, iter58):
        \\                     {"kind":"sst","idx":0,"text":"…","runs":null}
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
        ArgError.BadArgValue,
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

    // iter57/58 sub-commands — no per-sheet selection. meta /
    // list-sheets / styles / sst are workbook-wide; comments /
    // validations / hyperlinks iterate every sheet internally.
    switch (args.subcommand) {
        .meta => {
            // Unix argv is raw bytes; only emit `path` as JSON when
            // valid UTF-8 so the NDJSON line stays parseable. Invalid
            // bytes → JSON null + stderr warning.
            const path_opt: ?[]const u8 = if (std.unicode.utf8ValidateSlice(args.file))
                args.file
            else blk: {
                try err.print(
                    "zlsx: workbook path contains non-UTF-8 bytes; emitting \"path\":null in meta output\n",
                    .{},
                );
                try err.flush();
                break :blk null;
            };
            try runMetaCommand(out, &book, path_opt);
            return 0;
        },
        .list_sheets => {
            try runListSheetsCommand(out, &book);
            return 0;
        },
        .comments => {
            const filter = resolveSheetFilter(&book, args) catch {
                try err.writeAll("zlsx: sheet not found\n");
                return 3;
            };
            try runCommentsCommand(out, &book, filter, args.skip, args.take, args.start_row, args.end_row);
            return 0;
        },
        .validations => {
            const filter = resolveSheetFilter(&book, args) catch {
                try err.writeAll("zlsx: sheet not found\n");
                return 3;
            };
            try runValidationsCommand(out, &book, filter, args.skip, args.take);
            return 0;
        },
        .hyperlinks => {
            const filter = resolveSheetFilter(&book, args) catch {
                try err.writeAll("zlsx: sheet not found\n");
                return 3;
            };
            try runHyperlinksCommand(out, &book, filter, args.skip, args.take);
            return 0;
        },
        .styles => {
            try runStylesCommand(out, &book, args.skip, args.take);
            return 0;
        },
        .sst => {
            try runSstCommand(out, &book, args.skip, args.take);
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
        .rows => try runRowsCommand(out, &book, selected.sheet, selected.idx, args.format, alloc, args.skip, args.take, args.start_row, args.end_row),
        .cells => try runCellsCommand(out, &book, selected.sheet, selected.idx, alloc, args.skip, args.take, args.start_row, args.end_row),
        // Handled by the workbook-scoped early return above.
        .meta,
        .list_sheets,
        .comments,
        .validations,
        .hyperlinks,
        .styles,
        .sst,
        => unreachable,
    }
    return 0;
}

/// iter59a: stream-native pagination. `consume()` returns one of
/// three verdicts per candidate record. The counters apply GLOBALLY
/// over the emitted-record stream of a single sub-command run, per
/// the jq-for-excel CLI conventions in docs/jq-for-excel.md.
const Pagination = struct {
    skip: ?usize,
    take: ?usize,
    skipped: usize = 0,
    taken: usize = 0,

    const Verdict = enum { drop, emit, stop };

    fn init(skip: ?usize, take: ?usize) Pagination {
        return .{ .skip = skip, .take = take };
    }

    /// Call once per candidate record before emitting. `.drop` means
    /// advance past this record; `.emit` means emit then mark taken;
    /// `.stop` means --take already satisfied — return early without
    /// emitting anything further.
    fn consume(self: *Pagination) Verdict {
        if (self.take) |t| if (self.taken >= t) return .stop;
        if (self.skip) |s| if (self.skipped < s) {
            self.skipped += 1;
            return .drop;
        };
        self.taken += 1;
        return .emit;
    }
};

fn runRowsCommand(
    out: *std.Io.Writer,
    book: *xlsx.Book,
    sheet: xlsx.Sheet,
    sheet_idx: usize,
    format: Format,
    alloc: std.mem.Allocator,
    skip: ?usize,
    take: ?usize,
    start_row: ?u32,
    end_row: ?u32,
) !void {
    var rows = try book.rows(sheet, alloc);
    defer rows.deinit();

    var pg = Pagination.init(skip, take);
    while (try rows.next()) |cells| {
        const row_number = rows.currentRowNumber();
        // Row bounds run BEFORE pagination (design doc v4.1).
        if (start_row) |s| if (row_number < s) continue;
        // OOXML rows are monotonic — once past end_row, no more records
        // in this sheet's stream can satisfy the bound.
        if (end_row) |e| if (row_number > e) break;
        switch (pg.consume()) {
            .drop => continue,
            .stop => return,
            .emit => {},
        }
        switch (format) {
            .jsonl => try writeRowEnvelope(out, sheet.name, sheet_idx, row_number, cells),
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
    skip: ?usize,
    take: ?usize,
    start_row: ?u32,
    end_row: ?u32,
) !void {
    var rows = try book.rows(sheet, alloc);
    defer rows.deinit();

    var pg = Pagination.init(skip, take);
    while (try rows.next()) |cells| {
        const row_number = rows.currentRowNumber();
        if (start_row) |s| if (row_number < s) continue;
        if (end_row) |e| if (row_number > e) break;
        for (cells, 0..) |c, i| {
            if (c == .empty) continue;

            switch (pg.consume()) {
                .drop => continue,
                .stop => return,
                .emit => {},
            }

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
    path: ?[]const u8,
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

    // `path` is null when the caller detected non-UTF-8 bytes in the
    // original argv — emit JSON `null` so the NDJSON line stays
    // parseable. main() has already logged the reason to stderr.
    try out.writeAll("{\"kind\":\"workbook\",\"path\":");
    if (path) |p| try writeJsonString(out, p) else try out.writeAll("null");
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

// ─── iter58: reader-surface sub-commands ─────────────────────────────

/// Resolve the sheet-selector flags (--sheet index / --name) to an
/// optional sheet filter. Null means "iterate every sheet" (the
/// default for sheet-scoped-but-multi-sheet commands like comments /
/// validations / hyperlinks). Returns error.SheetNotFound when a
/// concrete selector was given but doesn't match the workbook.
fn resolveSheetFilter(book: *const xlsx.Book, args: Args) !?usize {
    if (args.sheet_index) |idx| {
        if (idx >= book.sheets.len) return error.SheetNotFound;
        return idx;
    }
    if (args.sheet_name) |name| {
        for (book.sheets, 0..) |s, i| {
            if (std.mem.eql(u8, s.name, name)) return i;
        }
        return error.SheetNotFound;
    }
    return null;
}

/// Emit an `"A1"`-style ref into `buf` from a reader-shape CellRef
/// (`col` is 0-based — A→0, B→1 — and `row` is 1-based, matching
/// `xlsx.parseA1Ref`). Panics if the generated ref exceeds 16 bytes —
/// OOXML's max column XFD (=16 383) plus max row 1 048 576 fits in
/// 10 bytes, so the budget has a lot of slack. Callers must not hold
/// the returned slice past the buffer's lifetime.
fn refFromCellRef(buf: *[16]u8, ref: xlsx.CellRef) []const u8 {
    std.debug.assert(ref.row >= 1);
    var letters_buf: [8]u8 = undefined;
    const letters = colLetter(&letters_buf, ref.col);
    return std.fmt.bufPrint(buf, "{s}{d}", .{ letters, ref.row }) catch unreachable;
}

/// Emit `{"text":…,"bold":…,…}` fields for a single RichRun. Caller
/// wraps in the surrounding `[` / `]`. `bold`, `italic`, `color`,
/// `size`, `font_name` are each emitted only when set (matches the
/// design-doc "emitted only when true/non-null" shorthand).
fn writeRichRun(w: *std.Io.Writer, run: xlsx.RichRun) !void {
    try w.writeAll("{\"text\":");
    try writeJsonString(w, run.text);
    if (run.bold) try w.writeAll(",\"bold\":true");
    if (run.italic) try w.writeAll(",\"italic\":true");
    if (run.color_argb) |c| try w.print(",\"color\":\"{X:0>8}\"", .{c});
    if (run.size) |s| {
        if (std.math.isFinite(s)) try w.print(",\"size\":{d}", .{s});
    }
    if (run.font_name.len != 0) {
        try w.writeAll(",\"font_name\":");
        try writeJsonString(w, run.font_name);
    }
    try w.writeByte('}');
}

/// Emit `null` for plain strings, otherwise `[<run>,…]`. Shared by
/// `comments` and `sst` which use the same runs wire-shape.
fn writeRichRunsOrNull(w: *std.Io.Writer, runs: ?[]const xlsx.RichRun) !void {
    const rs = runs orelse {
        try w.writeAll("null");
        return;
    };
    try w.writeByte('[');
    for (rs, 0..) |r, i| {
        if (i != 0) try w.writeByte(',');
        try writeRichRun(w, r);
    }
    try w.writeByte(']');
}

/// Map the reader's DataValidationKind to the OOXML wire-form string
/// the design doc pins in the `rule_type` field. `.unknown` surfaces
/// as the literal `"unknown"` so consumers can still filter.
fn validationKindName(kind: xlsx.DataValidationKind) []const u8 {
    return switch (kind) {
        .list => "list",
        .whole => "whole",
        .decimal => "decimal",
        .date => "date",
        .time => "time",
        .text_length => "textLength",
        .custom => "custom",
        .unknown => "unknown",
    };
}

/// Map DataValidationOperator to its OOXML camelCase token.
fn validationOpName(op: xlsx.DataValidationOperator) []const u8 {
    return switch (op) {
        .between => "between",
        .not_between => "notBetween",
        .equal => "equal",
        .not_equal => "notEqual",
        .less_than => "lessThan",
        .less_than_or_equal => "lessThanOrEqual",
        .greater_than => "greaterThan",
        .greater_than_or_equal => "greaterThanOrEqual",
    };
}

/// Emit `"A1"` for a single-cell range or `"A1:B2"` for a rectangle
/// into the caller-provided 32-byte buffer. Uses `refFromColRow`
/// under the hood so both endpoints get identical formatting.
fn rangeFromBounds(buf: *[32]u8, top_left: xlsx.CellRef, bottom_right: xlsx.CellRef) []const u8 {
    var tl_buf: [16]u8 = undefined;
    const tl = refFromCellRef(&tl_buf, top_left);
    if (top_left.col == bottom_right.col and top_left.row == bottom_right.row) {
        return std.fmt.bufPrint(buf, "{s}", .{tl}) catch unreachable;
    }
    var br_buf: [16]u8 = undefined;
    const br = refFromCellRef(&br_buf, bottom_right);
    return std.fmt.bufPrint(buf, "{s}:{s}", .{ tl, br }) catch unreachable;
}

/// Emit one NDJSON record per comment. When `filter` is set, only
/// the matching sheet contributes; otherwise every sheet iterates.
fn runCommentsCommand(
    out: *std.Io.Writer,
    book: *const xlsx.Book,
    filter: ?usize,
    skip: ?usize,
    take: ?usize,
    start_row: ?u32,
    end_row: ?u32,
) !void {
    var pg = Pagination.init(skip, take);
    for (book.sheets, 0..) |s, sheet_idx| {
        if (filter) |f| if (sheet_idx != f) continue;
        for (book.comments(s)) |c| {
            // Comments are not guaranteed monotonic by row across a
            // sheet's comment list (OOXML preserves author/insertion
            // order). `continue` on both bounds — don't `break`.
            if (start_row) |sr| if (c.top_left.row < sr) continue;
            if (end_row) |er| if (c.top_left.row > er) continue;
            switch (pg.consume()) {
                .drop => continue,
                .stop => return,
                .emit => {},
            }
            var ref_buf: [16]u8 = undefined;
            const ref = refFromCellRef(&ref_buf, c.top_left);

            try out.writeAll("{\"kind\":\"comment\",\"sheet\":");
            try writeJsonString(out, s.name);
            try out.print(",\"sheet_idx\":{d},\"ref\":", .{sheet_idx});
            try writeJsonString(out, ref);
            // Reader `col` is 0-based (A=0); wire format is 1-based
            // (A=1) for consistency with `cells` / `rows` envelopes.
            try out.print(
                ",\"row\":{d},\"col\":{d},\"author\":",
                .{ c.top_left.row, c.top_left.col + 1 },
            );
            try writeJsonString(out, c.author);
            try out.writeAll(",\"text\":");
            try writeJsonString(out, c.text);
            try out.writeAll(",\"runs\":");
            try writeRichRunsOrNull(out, c.runs);
            try out.writeAll("}\n");
        }
    }
}

/// Emit one NDJSON record per data-validation range. When `filter`
/// is set, only the matching sheet contributes.
fn runValidationsCommand(
    out: *std.Io.Writer,
    book: *const xlsx.Book,
    filter: ?usize,
    skip: ?usize,
    take: ?usize,
) !void {
    var pg = Pagination.init(skip, take);
    for (book.sheets, 0..) |s, sheet_idx| {
        if (filter) |f| if (sheet_idx != f) continue;
        for (book.dataValidations(s)) |dv| {
            switch (pg.consume()) {
                .drop => continue,
                .stop => return,
                .emit => {},
            }
            var range_buf: [32]u8 = undefined;
            const range = rangeFromBounds(&range_buf, dv.top_left, dv.bottom_right);

            try out.writeAll("{\"kind\":\"validation\",\"sheet\":");
            try writeJsonString(out, s.name);
            try out.print(",\"sheet_idx\":{d},\"range\":", .{sheet_idx});
            try writeJsonString(out, range);
            try out.print(",\"rule_type\":\"{s}\",\"op\":", .{validationKindName(dv.kind)});
            if (dv.op) |op| try out.print("\"{s}\"", .{validationOpName(op)}) else try out.writeAll("null");

            try out.writeAll(",\"formula1\":");
            try writeJsonString(out, dv.formula1);
            try out.writeAll(",\"formula2\":");
            if (dv.formula2.len != 0) try writeJsonString(out, dv.formula2) else try out.writeAll("null");

            try out.writeAll(",\"values\":");
            if (dv.kind == .list and dv.values.len != 0) {
                try out.writeByte('[');
                for (dv.values, 0..) |v, i| {
                    if (i != 0) try out.writeByte(',');
                    try writeJsonString(out, v);
                }
                try out.writeByte(']');
            } else {
                try out.writeAll("null");
            }
            try out.writeAll("}\n");
        }
    }
}

/// Emit one NDJSON record per hyperlink. When `filter` is set, only
/// the matching sheet contributes.
fn runHyperlinksCommand(
    out: *std.Io.Writer,
    book: *const xlsx.Book,
    filter: ?usize,
    skip: ?usize,
    take: ?usize,
) !void {
    var pg = Pagination.init(skip, take);
    for (book.sheets, 0..) |s, sheet_idx| {
        if (filter) |f| if (sheet_idx != f) continue;
        for (book.hyperlinks(s)) |h| {
            switch (pg.consume()) {
                .drop => continue,
                .stop => return,
                .emit => {},
            }
            var range_buf: [32]u8 = undefined;
            const range = rangeFromBounds(&range_buf, h.top_left, h.bottom_right);

            try out.writeAll("{\"kind\":\"hyperlink\",\"sheet\":");
            try writeJsonString(out, s.name);
            try out.print(",\"sheet_idx\":{d},\"range\":", .{sheet_idx});
            try writeJsonString(out, range);
            try out.writeAll(",\"url\":");
            if (h.url.len != 0) try writeJsonString(out, h.url) else try out.writeAll("null");
            try out.writeAll(",\"location\":");
            if (h.location.len != 0) try writeJsonString(out, h.location) else try out.writeAll("null");
            try out.writeAll("}\n");
        }
    }
}

/// Emit `{…}` for a BorderSide or `null` when the side has no style.
fn writeBorderSideOrNull(w: *std.Io.Writer, side: xlsx.BorderSide) !void {
    if (side.style.len == 0) {
        try w.writeAll("null");
        return;
    }
    try w.writeAll("{\"style\":");
    try writeJsonString(w, side.style);
    try w.writeAll(",\"color\":");
    if (side.color_argb) |c| try w.print("\"{X:0>8}\"", .{c}) else try w.writeAll("null");
    try w.writeByte('}');
}

/// Emit one NDJSON record per cell-XF style entry. Workbook-scoped.
/// Every nested block (`font` / `fill` / `border`) is either the
/// resolved struct or JSON `null` when the getter returns null.
fn runStylesCommand(
    out: *std.Io.Writer,
    book: *const xlsx.Book,
    skip: ?usize,
    take: ?usize,
) !void {
    var pg = Pagination.init(skip, take);
    for (book.cell_xf_numfmt_ids, 0..) |_, i| {
        switch (pg.consume()) {
            .drop => continue,
            .stop => return,
            .emit => {},
        }
        const idx: u32 = @intCast(i);

        try out.print("{{\"kind\":\"style\",\"idx\":{d},\"font\":", .{idx});
        if (book.cellFont(idx)) |f| {
            try out.writeAll("{\"bold\":");
            try out.writeAll(if (f.bold) "true" else "false");
            try out.writeAll(",\"italic\":");
            try out.writeAll(if (f.italic) "true" else "false");
            try out.writeAll(",\"color\":");
            if (f.color_argb) |c| try out.print("\"{X:0>8}\"", .{c}) else try out.writeAll("null");
            try out.writeAll(",\"size\":");
            if (f.size) |s| {
                if (std.math.isFinite(s)) try out.print("{d}", .{s}) else try out.writeAll("null");
            } else try out.writeAll("null");
            try out.writeAll(",\"name\":");
            if (f.name.len != 0) try writeJsonString(out, f.name) else try out.writeAll("null");
            try out.writeByte('}');
        } else try out.writeAll("null");

        try out.writeAll(",\"fill\":");
        if (book.cellFill(idx)) |fl| {
            // Treat the default zlsx Fill (pattern="none", both
            // colors null) as "no fill" on the wire, same as when
            // cellFill returned null. Consumers can then trust
            // `fill != null` to mean "the style actually defines
            // a fill."
            const no_fill = (std.mem.eql(u8, fl.pattern, "none") or fl.pattern.len == 0) and
                fl.fg_color_argb == null and fl.bg_color_argb == null;
            if (no_fill) {
                try out.writeAll("null");
            } else {
                try out.writeAll("{\"pattern\":");
                try writeJsonString(out, fl.pattern);
                try out.writeAll(",\"fg\":");
                if (fl.fg_color_argb) |c| try out.print("\"{X:0>8}\"", .{c}) else try out.writeAll("null");
                try out.writeAll(",\"bg\":");
                if (fl.bg_color_argb) |c| try out.print("\"{X:0>8}\"", .{c}) else try out.writeAll("null");
                try out.writeByte('}');
            }
        } else try out.writeAll("null");

        try out.writeAll(",\"border\":");
        if (book.cellBorder(idx)) |b| {
            // Same contract: if every side is the zero BorderSide
            // (empty style + null color), emit `null` so unstyled
            // XFs read as "no border" on the wire.
            const side_empty = struct {
                fn f(s: xlsx.BorderSide) bool {
                    return s.style.len == 0 and s.color_argb == null;
                }
            }.f;
            const no_border = side_empty(b.left) and side_empty(b.right) and
                side_empty(b.top) and side_empty(b.bottom) and side_empty(b.diagonal);
            if (no_border) {
                try out.writeAll("null");
            } else {
                try out.writeAll("{\"left\":");
                try writeBorderSideOrNull(out, b.left);
                try out.writeAll(",\"right\":");
                try writeBorderSideOrNull(out, b.right);
                try out.writeAll(",\"top\":");
                try writeBorderSideOrNull(out, b.top);
                try out.writeAll(",\"bottom\":");
                try writeBorderSideOrNull(out, b.bottom);
                try out.writeAll(",\"diagonal\":");
                try writeBorderSideOrNull(out, b.diagonal);
                try out.writeByte('}');
            }
        } else try out.writeAll("null");

        try out.writeAll(",\"num_fmt\":");
        if (book.numberFormat(idx)) |nf| try writeJsonString(out, nf) else try out.writeAll("null");
        try out.writeAll("}\n");
    }
}

/// Emit one NDJSON record per shared-string entry.
fn runSstCommand(
    out: *std.Io.Writer,
    book: *const xlsx.Book,
    skip: ?usize,
    take: ?usize,
) !void {
    var pg = Pagination.init(skip, take);
    for (book.shared_strings, 0..) |s, i| {
        switch (pg.consume()) {
            .drop => continue,
            .stop => return,
            .emit => {},
        }
        try out.print("{{\"kind\":\"sst\",\"idx\":{d},\"text\":", .{i});
        try writeJsonString(out, s);
        try out.writeAll(",\"runs\":");
        try writeRichRunsOrNull(out, book.richRuns(i));
        try out.writeAll("}\n");
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

test "parseArgs tolerates bogus --sheet / --format values on workbook-scoped sub-commands" {
    // Wrappers that append --sheet/--format universally must still
    // reach `meta` / `list-sheets` without an exit-1. Values are
    // silently dropped on those sub-commands, not validated.
    {
        const argv = [_][]const u8{ "meta", "f.xlsx", "--sheet", "nope" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.meta, a.subcommand);
        try std.testing.expect(a.sheet_index == null);
    }
    {
        const argv = [_][]const u8{ "list-sheets", "f.xlsx", "--format", "bogus" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.list_sheets, a.subcommand);
    }
    // Non-workbook-scoped commands stay strict — bogus --sheet still
    // errors on `rows` / `cells`.
    {
        const argv = [_][]const u8{ "cells", "f.xlsx", "--sheet", "nope" };
        try std.testing.expectError(ArgError.BadSheetIndex, parseArgs(&argv));
    }
    {
        const argv = [_][]const u8{ "rows", "f.xlsx", "--format", "bogus" };
        try std.testing.expectError(ArgError.BadFormat, parseArgs(&argv));
    }
}

test "parseArgs --skip / --take round-trip and tolerance" {
    // Both flags parse as usize and live on Args.
    {
        const argv = [_][]const u8{ "rows", "f.xlsx", "--skip", "5", "--take", "10" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(@as(?usize, 5), a.skip);
        try std.testing.expectEqual(@as(?usize, 10), a.take);
    }
    // Bogus --skip / --take are hard errors on record-scoped commands.
    {
        const argv = [_][]const u8{ "rows", "f.xlsx", "--skip", "bogus" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    {
        const argv = [_][]const u8{ "cells", "f.xlsx", "--take", "nope" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    // --skip / --take are strict on every sub-command (unlike
    // --sheet / --format whose tolerance follows the workbook_scoped
    // group). Pagination is too useful on styles / sst — a typoed
    // --take that silently returned the full stream would be an
    // expensive surprise. On meta / list-sheets which don't paginate,
    // the error is also the clearer signal than silent no-op.
    inline for (.{ "meta", "list-sheets", "styles", "sst" }) |cmd| {
        {
            const argv = [_][]const u8{ cmd, "f.xlsx", "--skip", "bogus" };
            try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
        }
        {
            const argv = [_][]const u8{ cmd, "f.xlsx", "--take", "nope" };
            try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
        }
    }
    // --skip and --take default to null when absent — legacy callers
    // must see identical behavior to pre-iter59a.
    {
        const argv = [_][]const u8{ "cells", "f.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expect(a.skip == null);
        try std.testing.expect(a.take == null);
    }
}

test "parseArgs --start-row / --end-row round-trip and rejections" {
    // Happy path: both parse as u32 and live on Args.
    {
        const argv = [_][]const u8{ "rows", "f.xlsx", "--start-row", "5", "--end-row", "10" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(@as(?u32, 5), a.start_row);
        try std.testing.expectEqual(@as(?u32, 10), a.end_row);
    }
    // Bogus values error (strict on every sub-command).
    {
        const argv = [_][]const u8{ "rows", "f.xlsx", "--start-row", "bogus" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    {
        const argv = [_][]const u8{ "cells", "f.xlsx", "--end-row", "nope" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    // 0 is a user error: OOXML rows are 1-based.
    {
        const argv = [_][]const u8{ "rows", "f.xlsx", "--start-row", "0" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    // start_row > end_row is an empty emission range — caught at parse.
    {
        const argv = [_][]const u8{ "cells", "f.xlsx", "--start-row", "10", "--end-row", "5" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    // start_row == end_row is a valid single-row slice.
    {
        const argv = [_][]const u8{ "cells", "f.xlsx", "--start-row", "7", "--end-row", "7" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(@as(?u32, 7), a.start_row);
        try std.testing.expectEqual(@as(?u32, 7), a.end_row);
    }
    // Sub-commands without a row key reject --start-row / --end-row.
    inline for (.{ "validations", "hyperlinks", "meta", "list-sheets", "styles", "sst" }) |cmd| {
        {
            const argv = [_][]const u8{ cmd, "f.xlsx", "--start-row", "2" };
            try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
        }
        {
            const argv = [_][]const u8{ cmd, "f.xlsx", "--end-row", "5" };
            try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
        }
    }
    // Explicitly allowed on the three row-keyed sub-commands.
    inline for (.{ "rows", "cells", "comments" }) |cmd| {
        const argv = [_][]const u8{ cmd, "f.xlsx", "--start-row", "2", "--end-row", "4" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(@as(?u32, 2), a.start_row);
        try std.testing.expectEqual(@as(?u32, 4), a.end_row);
    }
    // Defaults to null when absent.
    {
        const argv = [_][]const u8{ "cells", "f.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expect(a.start_row == null);
        try std.testing.expect(a.end_row == null);
    }
    // Legacy --list-sheets flag takes the early-return path in main
    // and emits plain sheet names; row bounds passed alongside it
    // would silently no-op. parseArgs must reject.
    {
        const argv = [_][]const u8{ "f.xlsx", "--list-sheets", "--start-row", "2" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    {
        const argv = [_][]const u8{ "f.xlsx", "--list-sheets", "--end-row", "10" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
}

test "runCellsCommand --start-row / --end-row bound the emitted cell stream" {
    const tmp_path = "/tmp/zlsx_cli_rowbounds_iter59b.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        // 5 rows × 1 cell each → rows 1..5 in the OOXML sense.
        try s0.writeRow(&.{.{ .string = "c1" }});
        try s0.writeRow(&.{.{ .string = "c2" }});
        try s0.writeRow(&.{.{ .string = "c3" }});
        try s0.writeRow(&.{.{ .string = "c4" }});
        try s0.writeRow(&.{.{ .string = "c5" }});
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    const countLines = struct {
        fn f(s: []const u8) usize {
            var n: usize = 0;
            for (s) |c| if (c == '\n') {
                n += 1;
            };
            return n;
        }
    }.f;

    // --start-row 2 --end-row 4 → rows 2, 3, 4.
    {
        var scratch: [4096]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        try runCellsCommand(&w, &book, book.sheets[0], 0, std.testing.allocator, null, null, 2, 4);
        const out = w.buffered();
        try std.testing.expectEqual(@as(usize, 3), countLines(out));
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c1\"") == null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c2\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c3\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c4\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c5\"") == null);
    }
    // Row bounds run BEFORE --skip/--take. Of rows 2/3/4, --skip 1
    // drops c2 and --take 1 keeps exactly c3.
    {
        var scratch: [4096]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        try runCellsCommand(&w, &book, book.sheets[0], 0, std.testing.allocator, 1, 1, 2, 4);
        const out = w.buffered();
        try std.testing.expectEqual(@as(usize, 1), countLines(out));
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c3\"") != null);
    }
}

test "runCellsCommand --skip / --take slice the emitted cell stream" {
    const tmp_path = "/tmp/zlsx_cli_pagination_iter59a.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        // 5 rows × 1 cell each → 5 candidate cells in emit order.
        try s0.writeRow(&.{.{ .string = "c1" }});
        try s0.writeRow(&.{.{ .string = "c2" }});
        try s0.writeRow(&.{.{ .string = "c3" }});
        try s0.writeRow(&.{.{ .string = "c4" }});
        try s0.writeRow(&.{.{ .string = "c5" }});
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    const countLines = struct {
        fn f(s: []const u8) usize {
            var n: usize = 0;
            for (s) |c| if (c == '\n') {
                n += 1;
            };
            return n;
        }
    }.f;

    // Baseline — no pagination.
    {
        var scratch: [4096]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        try runCellsCommand(&w, &book, book.sheets[0], 0, std.testing.allocator, null, null, null, null);
        try std.testing.expectEqual(@as(usize, 5), countLines(w.buffered()));
    }
    // --skip 2 drops the first two cells (c1, c2).
    {
        var scratch: [4096]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        try runCellsCommand(&w, &book, book.sheets[0], 0, std.testing.allocator, 2, null, null, null);
        const out = w.buffered();
        try std.testing.expectEqual(@as(usize, 3), countLines(out));
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c1\"") == null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c2\"") == null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c3\"") != null);
    }
    // --take 3 keeps exactly the first three.
    {
        var scratch: [4096]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        try runCellsCommand(&w, &book, book.sheets[0], 0, std.testing.allocator, null, 3, null, null);
        const out = w.buffered();
        try std.testing.expectEqual(@as(usize, 3), countLines(out));
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c3\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c4\"") == null);
    }
    // --skip 1 --take 2 yields the exact middle slice: c2, c3.
    {
        var scratch: [4096]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        try runCellsCommand(&w, &book, book.sheets[0], 0, std.testing.allocator, 1, 2, null, null);
        const out = w.buffered();
        try std.testing.expectEqual(@as(usize, 2), countLines(out));
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c1\"") == null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c2\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c3\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c4\"") == null);
    }
}

test "runMetaCommand emits path:null on non-UTF-8 workbook path" {
    var scratch: [512]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);

    // Build a minimal Book-shaped view without actually opening a
    // file — runMetaCommand only dereferences book.sheets / sst /
    // styles_xml / theme_xml / rich_runs_by_sst_idx / comments.
    var empty_book: xlsx.Book = .{
        .allocator = std.testing.allocator,
        .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer empty_book.deinit();

    try runMetaCommand(&w, &empty_book, null);

    const out = scratch[0..w.end];
    // The path field must serialize as literal `null`, not a string.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"path\":null") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"kind\":\"workbook\"") != null);
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

// ─── iter58 tests ────────────────────────────────────────────────────

test "parseArgs routes iter58 sub-commands correctly" {
    const names = [_][]const u8{ "comments", "validations", "hyperlinks", "styles", "sst" };
    const expected = [_]Subcommand{ .comments, .validations, .hyperlinks, .styles, .sst };
    for (names, expected) |n, want| {
        const argv = [_][]const u8{ n, "file.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(want, a.subcommand);
        try std.testing.expectEqualStrings("file.xlsx", a.file);
    }
    // Styles / sst are workbook-scoped — bogus --sheet / --format
    // must be tolerated (per iter57's P2 fix).
    {
        const argv = [_][]const u8{ "styles", "f.xlsx", "--format", "bogus" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.styles, a.subcommand);
    }
    {
        const argv = [_][]const u8{ "sst", "f.xlsx", "--sheet", "bogus" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.sst, a.subcommand);
    }
    // Comments / validations / hyperlinks ARE sheet-scoped as of
    // iter58-P2 follow-up — bogus --sheet / --format must error so
    // callers don't get silently-misrouted output.
    {
        const argv = [_][]const u8{ "comments", "f.xlsx", "--sheet", "bogus" };
        try std.testing.expectError(ArgError.BadSheetIndex, parseArgs(&argv));
    }
    {
        const argv = [_][]const u8{ "hyperlinks", "f.xlsx", "--format", "bogus" };
        try std.testing.expectError(ArgError.BadFormat, parseArgs(&argv));
    }
    // Valid --sheet narrows the filter on sheet-scoped sub-commands.
    {
        const argv = [_][]const u8{ "comments", "f.xlsx", "--sheet", "1" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(@as(?usize, 1), a.sheet_index);
    }
}

test "runCommentsCommand emits one record per comment across every sheet" {
    const tmp_path = "/tmp/zlsx_cli_comments_iter58.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{.{ .string = "hdr" }});
        try s0.addComment("A1", "Alice", "needs review");
        var s1 = try w.addSheet("Other");
        try s1.writeRow(&.{.{ .integer = 1 }});
        try s1.addComment("B2", "Bob", "hi");
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    var scratch: [2048]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runCommentsCommand(&w, &book, null, null, null, null, null);

    const out = w.buffered();
    try std.testing.expect(std.mem.startsWith(u8, out, "{\"kind\":\"comment\""));
    try std.testing.expect(std.mem.indexOf(u8, out, "\"sheet\":\"Data\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"ref\":\"A1\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"row\":1,\"col\":1") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"author\":\"Alice\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"text\":\"needs review\"") != null);
    // The writer's comment body wraps text in `<r><t>` (even for
    // plain bodies), so the reader populates `runs` as a one-entry
    // array of `{text:"…"}`. `runs:null` would require an `<r>`-less
    // body, which the writer doesn't emit today — exercise the
    // populated-runs path instead.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"runs\":[{\"text\":\"needs review\"}]") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"sheet\":\"Other\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"ref\":\"B2\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"row\":2,\"col\":2") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"author\":\"Bob\"") != null);
}

test "runValidationsCommand emits list validation with values array" {
    const tmp_path = "/tmp/zlsx_cli_validations_iter58.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{.{ .string = "fruit" }});
        try s0.addDataValidationList("B2:B100", &.{ "apple", "banana" });
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    var scratch: [2048]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runValidationsCommand(&w, &book, null, null, null);

    const out = w.buffered();
    try std.testing.expect(std.mem.startsWith(u8, out, "{\"kind\":\"validation\""));
    try std.testing.expect(std.mem.indexOf(u8, out, "\"sheet\":\"Data\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"range\":\"B2:B100\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"rule_type\":\"list\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"op\":null") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"formula2\":null") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"values\":[\"apple\",\"banana\"]") != null);
}

test "runHyperlinksCommand emits url set + location null for external links" {
    const tmp_path = "/tmp/zlsx_cli_hyperlinks_iter58.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{.{ .string = "site" }});
        try s0.addHyperlink("A2", "https://example.com/");
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    var scratch: [2048]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runHyperlinksCommand(&w, &book, null, null, null);

    const out = w.buffered();
    try std.testing.expect(std.mem.startsWith(u8, out, "{\"kind\":\"hyperlink\""));
    try std.testing.expect(std.mem.indexOf(u8, out, "\"sheet\":\"Data\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"range\":\"A2\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"url\":\"https://example.com/\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"location\":null") != null);
}

test "runStylesCommand emits one record per cell-XF entry" {
    const tmp_path = "/tmp/zlsx_cli_styles_iter58.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        _ = try w.addStyle(.{ .font_bold = true });
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{.{ .string = "hdr" }});
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    var scratch: [4096]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runStylesCommand(&w, &book, null, null);

    const out = w.buffered();
    try std.testing.expect(std.mem.startsWith(u8, out, "{\"kind\":\"style\""));
    try std.testing.expect(std.mem.indexOf(u8, out, "\"idx\":0") != null);
    // The bold style registered at addStyle idx=1 (idx 0 is the default
    // no-style xf slot); the record MUST surface with bold:true.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"idx\":1") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"bold\":true") != null);
    // Each record also pins font / fill / border / num_fmt fields.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"font\":") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"fill\":") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"border\":") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"num_fmt\":") != null);
}

test "runSstCommand emits one record per shared-string entry" {
    const tmp_path = "/tmp/zlsx_cli_sst_iter58.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{ .{ .string = "header" }, .{ .string = "qty" } });
        try s0.writeRow(&.{ .{ .string = "apple" }, .{ .integer = 3 } });
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    var scratch: [4096]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runSstCommand(&w, &book, null, null);

    const out = w.buffered();
    try std.testing.expect(std.mem.startsWith(u8, out, "{\"kind\":\"sst\""));
    try std.testing.expect(std.mem.indexOf(u8, out, "\"idx\":0") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"text\":\"header\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"text\":\"qty\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"text\":\"apple\"") != null);
    // Plain strings — runs must be null on every record.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"runs\":null") != null);
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
