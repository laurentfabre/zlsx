//! The frozen oracle extractor (M1b, `goal_formula.md` §8.2).
//!
//! Reads a recalculated workbook and reports what is actually in it —
//! cached values, formula text, calc state — with **no code shared with
//! the implementation under test**. `zip_reader.zig` and `xml_scan.zig`
//! carry that independence; this file is the semantic layer over them.
//!
//! Frozen means versioned: `version` below is recorded in every
//! manifest's provenance. Changing what this file extracts, or how it
//! decodes, is a version bump plus a reviewed regeneration — otherwise
//! a manifest recorded last month and one recorded today are not
//! comparable, and the oracle silently stops being one.
//!
//! Decoding follows the carrier split this project already committed to
//! (`goal_formula.md` §5, M4b1), because manifests recorded now must
//! stay comparable with what the evaluator produces later:
//!
//!   * FORMULA carriers (`<f>` bodies) — XML entity decoding ONLY.
//!     ST_Xstring does not apply, so a literal `_x0041_` inside a
//!     formula survives byte-exact.
//!   * STRING carriers (shared strings, inline strings, `t="str"`
//!     values) — XML entities THEN ST_Xstring. Rich strings concatenate
//!     every visible `<r>/<t>` run in document order; `<rPh>` phonetic
//!     runs are excluded.

const std = @import("std");
const zip = @import("zip_reader.zig");
const xml = @import("xml_scan.zig");

/// Extractor version, recorded in provenance. Bump on ANY change to
/// what is extracted or how it is decoded.
pub const version = "oracle-extractor-1";

pub const Error = error{
    MissingWorkbookPart,
    MissingWorkbookRels,
    MissingSheetPart,
    UnexpectedCdata,
    MalformedCellRef,
    TooManyCells,
} || zip.Error || xml.Error;

/// Cap on cells per workbook. Oracle inputs are hand-sized; a workbook
/// large enough to hit this is not an oracle case and would make the
/// manifests unreviewable, which defeats "reviewed regeneration".
pub const max_cells: usize = 100_000;

pub const CellKind = enum {
    /// `t` absent or `t="n"` — a number in `<v>`.
    number,
    /// `t="s"` — `<v>` is a shared-string index.
    shared_string,
    /// `t="inlineStr"` — text in `<is>`.
    inline_string,
    /// `t="str"` — a formula that produced text.
    formula_string,
    /// `t="b"` — `<v>` is `0` or `1`.
    boolean,
    /// `t="e"` — `<v>` is an error spelling.
    err,
    /// `t="d"` — ISO 8601 date in `<v>` (deferred to M4b2 in the
    /// evaluator; recorded here so the oracle notices if Excel emits it).
    iso_date,
    /// A `<c>` with neither `<v>` nor `<is>`.
    blank,
};

pub const Formula = struct {
    /// Decoded formula body (XML entities only — see the header).
    text: []const u8,
    /// `t` attribute: null (normal), "shared", "array", "dataTable".
    kind: ?[]const u8,
    /// `ref` attribute on shared/array masters.
    ref: ?[]const u8,
    /// `si` shared-formula group index.
    si: ?[]const u8,
    /// `ca="1"` — "always calculate". This is how a volatile formula
    /// announces itself, and §8.2 excludes volatiles from every
    /// external value oracle, so it has to survive extraction.
    always_calc: bool,
};

pub const Cell = struct {
    sheet: []const u8,
    ref: []const u8,
    kind: CellKind,
    /// Raw `<v>` text after XML entity decoding. Null for a blank cell.
    /// Numbers keep their source spelling here; `manifest.zig` is what
    /// turns that into a binary64 and its bit pattern.
    value: ?[]const u8,
    /// Fully decoded text for the string kinds (shared string resolved,
    /// ST_Xstring applied). Null otherwise.
    text: ?[]const u8,
    formula: ?Formula,
};

pub const CalcState = struct {
    /// `calcMode`: "auto" (default), "manual", "autoNoTable".
    calc_mode: []const u8 = "auto",
    full_calc_on_load: bool = false,
    /// `fullPrecision="0"` means precision-as-displayed, which
    /// `goal_formula.md` refuses through all of v1 — the screen uses it.
    full_precision: bool = true,
    date1904: bool = false,
    calc_id: ?[]const u8 = null,
    /// Workbook declares links to other workbooks.
    has_external_references: bool = false,
};

pub const Workbook = struct {
    arena: std.heap.ArenaAllocator,
    /// Sheet names in workbook order.
    sheets: [][]const u8,
    cells: []Cell,
    calc: CalcState,
    /// SHA-256 of the whole .xlsx, lowercase hex. Provenance pins the
    /// exact bytes a manifest was derived from.
    digest: [64]u8,

    pub fn deinit(self: *Workbook) void {
        self.arena.deinit();
        self.* = undefined;
    }

    pub fn cell(self: Workbook, sheet: []const u8, ref: []const u8) ?Cell {
        for (self.cells) |c| {
            if (std.mem.eql(u8, c.sheet, sheet) and std.mem.eql(u8, c.ref, ref)) return c;
        }
        return null;
    }
};

/// Extract `bytes` (a whole .xlsx). Everything the result points at
/// lives in its arena; `bytes` need not outlive the call.
pub fn extract(allocator: std.mem.Allocator, bytes: []const u8) Error!Workbook {
    var arena: std.heap.ArenaAllocator = .init(allocator);
    errdefer arena.deinit();
    const a = arena.allocator();

    var archive = try zip.open(allocator, bytes);
    defer archive.deinit(allocator);

    var digest: [32]u8 = undefined;
    std.crypto.hash.sha2.Sha256.hash(bytes, &digest, .{});
    var digest_hex: [64]u8 = undefined;
    _ = std.fmt.bufPrint(&digest_hex, "{x}", .{&digest}) catch unreachable;

    const workbook_xml = archive.read(a, "xl/workbook.xml") catch return error.MissingWorkbookPart;
    const rels_xml = archive.read(a, "xl/_rels/workbook.xml.rels") catch return error.MissingWorkbookRels;

    const wb = try parseWorkbook(a, workbook_xml);
    const rels = try parseRels(a, rels_xml);
    const shared = parseSharedStrings(a, archive.read(a, "xl/sharedStrings.xml") catch null) catch
        &[_][]const u8{};

    var cells: std.ArrayListUnmanaged(Cell) = .empty;
    var sheet_names: std.ArrayListUnmanaged([]const u8) = .empty;

    for (wb.sheets) |sheet| {
        try sheet_names.append(a, sheet.name);
        const target = rels.get(sheet.rel_id) orelse return error.MissingSheetPart;
        const part = try resolveTarget(a, target);
        const sheet_xml = archive.read(a, part) catch return error.MissingSheetPart;
        try parseSheet(a, sheet.name, sheet_xml, shared, &cells);
        if (cells.items.len > max_cells) return error.TooManyCells;
    }

    return .{
        .arena = arena,
        .sheets = try sheet_names.toOwnedSlice(a),
        .cells = try cells.toOwnedSlice(a),
        .calc = wb.calc,
        .digest = digest_hex,
    };
}

// ─── workbook.xml ────────────────────────────────────────────────

const SheetRef = struct { name: []const u8, rel_id: []const u8 };
const ParsedWorkbook = struct { sheets: []SheetRef, calc: CalcState };

fn parseWorkbook(a: std.mem.Allocator, doc: []const u8) Error!ParsedWorkbook {
    var sheets: std.ArrayListUnmanaged(SheetRef) = .empty;
    var calc: CalcState = .{};

    var scanner: xml.Scanner = .init(doc);
    while (try scanner.next()) |ev| {
        const el = switch (ev) {
            .open, .self_closing => |e| e,
            else => continue,
        };
        const name = el.local();
        if (std.mem.eql(u8, name, "sheet")) {
            const sheet_name = el.attr("name") orelse continue;
            const rel = el.attr("id") orelse continue;
            try sheets.append(a, .{
                .name = try xml.decodeEntities(a, sheet_name),
                .rel_id = try a.dupe(u8, rel),
            });
        } else if (std.mem.eql(u8, name, "calcPr")) {
            if (el.attr("calcMode")) |m| calc.calc_mode = try a.dupe(u8, m);
            if (el.attr("fullCalcOnLoad")) |v| calc.full_calc_on_load = isTrue(v);
            if (el.attr("fullPrecision")) |v| calc.full_precision = isTrue(v);
            if (el.attr("calcId")) |v| calc.calc_id = try a.dupe(u8, v);
        } else if (std.mem.eql(u8, name, "workbookPr")) {
            if (el.attr("date1904")) |v| calc.date1904 = isTrue(v);
        } else if (std.mem.eql(u8, name, "externalReferences")) {
            calc.has_external_references = true;
        }
    }
    return .{ .sheets = try sheets.toOwnedSlice(a), .calc = calc };
}

/// OOXML booleans are `1`/`0` or `true`/`false`; both spellings appear
/// in the wild, and reading only one silently flips a screen result.
fn isTrue(v: []const u8) bool {
    return std.mem.eql(u8, v, "1") or std.ascii.eqlIgnoreCase(v, "true");
}

// ─── relationships ───────────────────────────────────────────────

const RelMap = std.StringHashMapUnmanaged([]const u8);

fn parseRels(a: std.mem.Allocator, doc: []const u8) Error!RelMap {
    var map: RelMap = .empty;
    var scanner: xml.Scanner = .init(doc);
    while (try scanner.next()) |ev| {
        const el = switch (ev) {
            .open, .self_closing => |e| e,
            else => continue,
        };
        if (!std.mem.eql(u8, el.local(), "Relationship")) continue;
        const id = el.attr("Id") orelse continue;
        const target = el.attr("Target") orelse continue;
        try map.put(a, try a.dupe(u8, id), try xml.decodeEntities(a, target));
    }
    return map;
}

/// Resolve a workbook-relative relationship target to a part name.
/// Targets are relative to `xl/` (the rels file's owner directory) and
/// may be written absolute (`/xl/worksheets/sheet1.xml`).
fn resolveTarget(a: std.mem.Allocator, target: []const u8) Error![]const u8 {
    if (target.len > 0 and target[0] == '/') return a.dupe(u8, target[1..]);
    return std.fmt.allocPrint(a, "xl/{s}", .{target}) catch error.OutOfMemory;
}

// ─── sharedStrings.xml ───────────────────────────────────────────

/// Each `<si>` becomes one string: every `<t>` outside a `<rPh>` block,
/// concatenated in document order. Phonetic runs are Excel's furigana
/// annotations — visible in the UI as a separate line, never part of
/// the cell's value.
fn parseSharedStrings(a: std.mem.Allocator, doc: ?[]const u8) Error![][]const u8 {
    const body = doc orelse return &[_][]const u8{};
    var out: std.ArrayListUnmanaged([]const u8) = .empty;

    var current: std.ArrayListUnmanaged(u8) = .empty;
    var in_si = false;
    var in_t = false;
    var phonetic_depth: usize = 0;

    var scanner: xml.Scanner = .init(body);
    while (try scanner.next()) |ev| switch (ev) {
        .open => |el| {
            const n = el.local();
            if (std.mem.eql(u8, n, "si")) {
                in_si = true;
                current = .empty;
            } else if (std.mem.eql(u8, n, "rPh")) {
                phonetic_depth += 1;
            } else if (std.mem.eql(u8, n, "t")) {
                in_t = true;
            }
        },
        .self_closing => |el| {
            const n = el.local();
            // `<si/>` and `<t/>` are both legal and both mean "empty".
            if (std.mem.eql(u8, n, "si")) try out.append(a, try a.dupe(u8, ""));
        },
        .close => |n| {
            const local = if (std.mem.indexOfScalar(u8, n, ':')) |c| n[c + 1 ..] else n;
            if (std.mem.eql(u8, local, "si")) {
                if (in_si) {
                    const decoded = try xml.decodeXstring(a, current.items);
                    try out.append(a, decoded);
                }
                in_si = false;
                current = .empty;
            } else if (std.mem.eql(u8, local, "rPh")) {
                if (phonetic_depth > 0) phonetic_depth -= 1;
            } else if (std.mem.eql(u8, local, "t")) {
                in_t = false;
            }
        },
        .text => |t| {
            if (in_si and in_t and phonetic_depth == 0) {
                const decoded = try xml.decodeEntities(a, t);
                try current.appendSlice(a, decoded);
            }
        },
        .cdata => return error.UnexpectedCdata,
    };

    return out.toOwnedSlice(a);
}

// ─── worksheet ───────────────────────────────────────────────────

fn parseSheet(
    a: std.mem.Allocator,
    sheet_name: []const u8,
    doc: []const u8,
    shared: []const []const u8,
    out: *std.ArrayListUnmanaged(Cell),
) Error!void {
    var scanner: xml.Scanner = .init(doc);

    var cell_ref: ?[]const u8 = null;
    var cell_type: ?[]const u8 = null;
    var in_cell = false;
    var in_v = false;
    var in_is_t = false;
    var in_is = false;
    var in_f = false;
    var v_text: std.ArrayListUnmanaged(u8) = .empty;
    var is_text: std.ArrayListUnmanaged(u8) = .empty;
    var f_text: std.ArrayListUnmanaged(u8) = .empty;
    var formula: ?Formula = null;
    var has_v = false;
    var has_is = false;

    while (try scanner.next()) |ev| switch (ev) {
        .open => |el| {
            const n = el.local();
            if (std.mem.eql(u8, n, "c")) {
                in_cell = true;
                cell_ref = if (el.attr("r")) |r| try a.dupe(u8, r) else null;
                cell_type = if (el.attr("t")) |t| try a.dupe(u8, t) else null;
                v_text = .empty;
                is_text = .empty;
                f_text = .empty;
                formula = null;
                has_v = false;
                has_is = false;
            } else if (std.mem.eql(u8, n, "v")) {
                in_v = true;
                has_v = true;
            } else if (std.mem.eql(u8, n, "is")) {
                in_is = true;
                has_is = true;
            } else if (std.mem.eql(u8, n, "t")) {
                if (in_is) in_is_t = true;
            } else if (std.mem.eql(u8, n, "f")) {
                in_f = true;
                formula = formulaFrom(a, el, "");
            }
        },
        .self_closing => |el| {
            const n = el.local();
            if (std.mem.eql(u8, n, "c")) {
                // `<c r="A1"/>` — a styled but empty cell.
                if (!cellRefOk(el.attr("r"))) return error.MalformedCellRef;
                try out.append(a, .{
                    .sheet = sheet_name,
                    .ref = try a.dupe(u8, el.attr("r").?),
                    .kind = .blank,
                    .value = null,
                    .text = null,
                    .formula = null,
                });
                cell_ref = null;
            } else if (std.mem.eql(u8, n, "f")) {
                // A shared-formula slave carries no body: `<f t="shared" si="0"/>`.
                formula = formulaFrom(a, el, "");
            }
        },
        .close => |raw| {
            const n = if (std.mem.indexOfScalar(u8, raw, ':')) |c| raw[c + 1 ..] else raw;
            if (std.mem.eql(u8, n, "v")) {
                in_v = false;
            } else if (std.mem.eql(u8, n, "t")) {
                in_is_t = false;
            } else if (std.mem.eql(u8, n, "is")) {
                in_is = false;
            } else if (std.mem.eql(u8, n, "f")) {
                in_f = false;
                if (formula) |*f| f.text = try a.dupe(u8, f_text.items);
            } else if (std.mem.eql(u8, n, "c")) {
                if (!in_cell) continue;
                in_cell = false;
                // A `<c>` without `r` is legal OOXML (MS-OE376 §2.1.624:
                // the column after its predecessor) and the evaluator
                // reconstructs it at M4b1. An ORACLE must not — pinning
                // an expected value onto a guessed coordinate is worse
                // than refusing the workbook.
                const ref = cell_ref orelse return error.MalformedCellRef;
                if (!cellRefOk(ref)) return error.MalformedCellRef;
                const kind = classify(cell_type, has_v, has_is);
                const raw_value: ?[]const u8 = if (has_v) v_text.items else null;
                try out.append(a, .{
                    .sheet = sheet_name,
                    .ref = ref,
                    .kind = kind,
                    .value = if (raw_value) |rv| try a.dupe(u8, rv) else null,
                    .text = try resolveText(a, kind, raw_value, is_text.items, shared),
                    .formula = formula,
                });
                cell_ref = null;
            }
        },
        .text => |t| {
            if (in_v) {
                try v_text.appendSlice(a, try xml.decodeEntities(a, t));
            } else if (in_is_t) {
                try is_text.appendSlice(a, try xml.decodeEntities(a, t));
            } else if (in_f) {
                // FORMULA carrier: entity decoding ONLY, no ST_Xstring.
                try f_text.appendSlice(a, try xml.decodeEntities(a, t));
            }
        },
        .cdata => return error.UnexpectedCdata,
    };
}

fn formulaFrom(a: std.mem.Allocator, el: xml.Element, text: []const u8) Formula {
    return .{
        .text = text,
        .kind = if (el.attr("t")) |t| (a.dupe(u8, t) catch null) else null,
        .ref = if (el.attr("ref")) |r| (a.dupe(u8, r) catch null) else null,
        .si = if (el.attr("si")) |s| (a.dupe(u8, s) catch null) else null,
        .always_calc = if (el.attr("ca")) |c| isTrue(c) else false,
    };
}

fn classify(t: ?[]const u8, has_v: bool, has_is: bool) CellKind {
    const tag = t orelse return if (has_v) .number else .blank;
    if (std.mem.eql(u8, tag, "s")) return .shared_string;
    if (std.mem.eql(u8, tag, "b")) return .boolean;
    if (std.mem.eql(u8, tag, "e")) return .err;
    if (std.mem.eql(u8, tag, "str")) return .formula_string;
    if (std.mem.eql(u8, tag, "d")) return .iso_date;
    if (std.mem.eql(u8, tag, "inlineStr")) return if (has_is) .inline_string else .blank;
    return if (has_v) .number else .blank;
}

/// Produce the decoded text of a string-carrying cell. Shared strings
/// resolve through the SST (already ST_Xstring-decoded); inline and
/// `t="str"` values are decoded here.
fn resolveText(
    a: std.mem.Allocator,
    kind: CellKind,
    raw_value: ?[]const u8,
    inline_text: []const u8,
    shared: []const []const u8,
) Error!?[]const u8 {
    return switch (kind) {
        .shared_string => blk: {
            const v = raw_value orelse break :blk null;
            const idx = std.fmt.parseInt(usize, std.mem.trim(u8, v, " \t\r\n"), 10) catch
                break :blk null;
            if (idx >= shared.len) break :blk null;
            break :blk shared[idx];
        },
        .inline_string => try xml.decodeXstring(a, inline_text),
        .formula_string => blk: {
            const v = raw_value orelse break :blk null;
            break :blk try xml.decodeXstring(a, v);
        },
        else => null,
    };
}

/// A cell reference must be `[A-Za-z]+[0-9]+`, optionally `$`-anchored.
/// The extractor does NOT reconstruct implicit coordinates (a `<c>`
/// without `r`): the evaluator will at M4b1, but an oracle that guesses
/// a coordinate could pin an expectation onto the wrong cell.
fn cellRefOk(ref: ?[]const u8) bool {
    const s = ref orelse return false;
    var i: usize = 0;
    if (i < s.len and s[i] == '$') i += 1;
    const letters = i;
    while (i < s.len and std.ascii.isAlphabetic(s[i])) : (i += 1) {}
    if (i == letters) return false;
    if (i < s.len and s[i] == '$') i += 1;
    const digits = i;
    while (i < s.len and std.ascii.isDigit(s[i])) : (i += 1) {}
    return i > digits and i == s.len;
}

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

test "extracts sheets, cells, formulas and calc state from a real workbook" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();

    const bytes = std.Io.Dir.cwd().readFileAlloc(
        io,
        "tests/corpus/openxlsx_loadExample.xlsx",
        testing.allocator,
        .limited(32 << 20),
    ) catch return error.SkipZigTest;
    defer testing.allocator.free(bytes);

    var wb = try extract(testing.allocator, bytes);
    defer wb.deinit();

    try testing.expect(wb.sheets.len >= 2);
    try testing.expect(wb.cells.len > 100);
    try testing.expectEqual(@as(usize, 64), wb.digest.len);

    // This workbook carries volatile shared formulas — `ca="1"` must
    // survive, since §8.2 excludes volatiles from external value oracles.
    var volatile_seen = false;
    var shared_seen = false;
    for (wb.cells) |c| {
        const f = c.formula orelse continue;
        if (f.always_calc) volatile_seen = true;
        if (f.kind) |k| {
            if (std.mem.eql(u8, k, "shared")) shared_seen = true;
        }
    }
    try testing.expect(volatile_seen);
    try testing.expect(shared_seen);
}

test "every corpus workbook extracts or fails with a typed error" {
    var threaded: std.Io.Threaded = .init(testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var dir = std.Io.Dir.cwd().openDir(io, "tests/corpus", .{ .iterate = true }) catch
        return error.SkipZigTest;
    defer dir.close(io);

    var extracted: usize = 0;
    var it = dir.iterate();
    while (try it.next(io)) |dirent| {
        if (dirent.kind != .file) continue;
        if (!std.mem.endsWith(u8, dirent.name, ".xlsx")) continue;
        const bytes = dir.readFileAlloc(io, dirent.name, testing.allocator, .limited(32 << 20)) catch
            continue;
        defer testing.allocator.free(bytes);

        // Typed error or success — never a panic, and never a partial
        // result presented as complete.
        var wb = extract(testing.allocator, bytes) catch continue;
        defer wb.deinit();
        extracted += 1;
    }
    try testing.expect(extracted >= 10);
}

const synth_prefix =
    \\<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
;
const synth_suffix = "</sheetData></worksheet>";

fn extractSynthetic(a: std.mem.Allocator, sheet_body: []const u8, sst: ?[]const u8) !std.ArrayListUnmanaged(Cell) {
    const doc = try std.mem.concat(a, u8, &.{ synth_prefix, sheet_body, synth_suffix });
    const shared = try parseSharedStrings(a, sst);
    var cells: std.ArrayListUnmanaged(Cell) = .empty;
    try parseSheet(a, "S", doc, shared, &cells);
    return cells;
}

test "cell kinds classify from the t attribute" {
    var arena: std.heap.ArenaAllocator = .init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    const cells = try extractSynthetic(a,
        \\<row r="1">
        \\<c r="A1"><v>42</v></c>
        \\<c r="B1" t="n"><v>1.5</v></c>
        \\<c r="C1" t="b"><v>1</v></c>
        \\<c r="D1" t="e"><v>#DIV/0!</v></c>
        \\<c r="E1" t="str"><f>"x"</f><v>x</v></c>
        \\<c r="F1" t="inlineStr"><is><t>inline</t></is></c>
        \\<c r="G1" s="3"/>
        \\<c r="H1" t="d"><v>2026-08-03T00:00:00</v></c>
        \\</row>
    , null);

    try testing.expectEqual(@as(usize, 8), cells.items.len);
    try testing.expectEqual(CellKind.number, cells.items[0].kind);
    try testing.expectEqual(CellKind.number, cells.items[1].kind);
    try testing.expectEqual(CellKind.boolean, cells.items[2].kind);
    try testing.expectEqual(CellKind.err, cells.items[3].kind);
    try testing.expectEqualStrings("#DIV/0!", cells.items[3].value.?);
    try testing.expectEqual(CellKind.formula_string, cells.items[4].kind);
    try testing.expectEqualStrings("x", cells.items[4].text.?);
    try testing.expectEqual(CellKind.inline_string, cells.items[5].kind);
    try testing.expectEqualStrings("inline", cells.items[5].text.?);
    try testing.expectEqual(CellKind.blank, cells.items[6].kind);
    try testing.expectEqual(CellKind.iso_date, cells.items[7].kind);
}

test "formula carriers get entity decoding only — no ST_Xstring" {
    // The load-bearing asymmetry of §5 M4b1. A literal `_x0041_` in a
    // formula must come back literal; the same bytes in a string cell
    // must come back as "A". Getting this wrong makes every later
    // formula-text comparison wrong in the same direction, which is the
    // hardest kind of oracle bug to notice.
    var arena: std.heap.ArenaAllocator = .init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    const cells = try extractSynthetic(a,
        \\<row r="1">
        \\<c r="A1" t="str"><f>IF(A2&gt;1,"_x0041_","b")</f><v>_x0041_</v></c>
        \\<c r="B1" t="inlineStr"><is><t>_x0041_</t></is></c>
        \\</row>
    , null);

    // Formula: `&gt;` decoded, `_x0041_` untouched.
    try testing.expectEqualStrings("IF(A2>1,\"_x0041_\",\"b\")", cells.items[0].formula.?.text);
    // `t="str"` value is a STRING carrier: ST_Xstring applies.
    try testing.expectEqualStrings("A", cells.items[0].text.?);
    // Inline string: likewise.
    try testing.expectEqualStrings("A", cells.items[1].text.?);
}

test "shared strings: rich runs concatenate, phonetic runs are excluded" {
    var arena: std.heap.ArenaAllocator = .init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    const sst =
        \\<sst xmlns="x">
        \\<si><t>plain</t></si>
        \\<si><r><t>rich </t></r><r><t>runs</t></r></si>
        \\<si><t>kanji</t><rPh sb="0" eb="2"><t>furigana</t></rPh></si>
        \\<si><t>_x0041_</t></si>
        \\<si/>
        \\</sst>
    ;
    const cells = try extractSynthetic(a,
        \\<row r="1">
        \\<c r="A1" t="s"><v>0</v></c>
        \\<c r="B1" t="s"><v>1</v></c>
        \\<c r="C1" t="s"><v>2</v></c>
        \\<c r="D1" t="s"><v>3</v></c>
        \\<c r="E1" t="s"><v>4</v></c>
        \\</row>
    , sst);

    try testing.expectEqualStrings("plain", cells.items[0].text.?);
    try testing.expectEqualStrings("rich runs", cells.items[1].text.?);
    // The furigana run must NOT join the value.
    try testing.expectEqualStrings("kanji", cells.items[2].text.?);
    // SST is a string carrier, so ST_Xstring applies.
    try testing.expectEqualStrings("A", cells.items[3].text.?);
    try testing.expectEqualStrings("", cells.items[4].text.?);
}

test "formula attributes survive: shared, array, ca" {
    var arena: std.heap.ArenaAllocator = .init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    const cells = try extractSynthetic(a,
        \\<row r="1">
        \\<c r="A1"><f t="shared" ref="A1:A3" si="0" ca="1">RAND()</f><v>0.5</v></c>
        \\<c r="A2"><f t="shared" si="0"/><v>0.25</v></c>
        \\<c r="B1"><f t="array" ref="B1:B2">SUM(A1:A2)</f><v>0.75</v></c>
        \\<c r="C1"><f>1+1</f><v>2</v></c>
        \\</row>
    , null);

    const master = cells.items[0].formula.?;
    try testing.expectEqualStrings("shared", master.kind.?);
    try testing.expectEqualStrings("A1:A3", master.ref.?);
    try testing.expectEqualStrings("0", master.si.?);
    try testing.expect(master.always_calc);
    try testing.expectEqualStrings("RAND()", master.text);

    // Slave: self-closing `<f>`, no body, still carries si.
    const slave = cells.items[1].formula.?;
    try testing.expectEqualStrings("0", slave.si.?);
    try testing.expectEqualStrings("", slave.text);
    try testing.expect(!slave.always_calc);

    const array = cells.items[2].formula.?;
    try testing.expectEqualStrings("array", array.kind.?);
    try testing.expectEqualStrings("B1:B2", array.ref.?);

    const plain = cells.items[3].formula.?;
    try testing.expect(plain.kind == null);
    try testing.expectEqualStrings("1+1", plain.text);
}

test "a cell without an r attribute is refused, not guessed" {
    var arena: std.heap.ArenaAllocator = .init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    // Implicit coordinates are real OOXML (MS-OE376 §2.1.624) and the
    // evaluator reconstructs them at M4b1. An ORACLE must not: pinning
    // an expected value onto a guessed coordinate is worse than
    // refusing the workbook.
    const doc = try std.mem.concat(a, u8, &.{
        synth_prefix, "<row r=\"1\"><c><v>1</v></c></row>", synth_suffix,
    });
    var cells: std.ArrayListUnmanaged(Cell) = .empty;
    try testing.expectError(error.MalformedCellRef, parseSheet(a, "S", doc, &.{}, &cells));
}

test "calc state parses both boolean spellings" {
    var arena: std.heap.ArenaAllocator = .init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    const numeric = try parseWorkbook(a,
        \\<workbook><workbookPr date1904="1"/><sheets/>
        \\<calcPr calcId="191029" calcMode="manual" fullCalcOnLoad="1" fullPrecision="0"/>
        \\<externalReferences><externalReference r:id="rId9"/></externalReferences></workbook>
    );
    try testing.expectEqualStrings("manual", numeric.calc.calc_mode);
    try testing.expect(numeric.calc.full_calc_on_load);
    try testing.expect(!numeric.calc.full_precision);
    try testing.expect(numeric.calc.date1904);
    try testing.expect(numeric.calc.has_external_references);
    try testing.expectEqualStrings("191029", numeric.calc.calc_id.?);

    const worded = try parseWorkbook(a,
        \\<workbook><workbookPr date1904="false"/><calcPr fullCalcOnLoad="true"/></workbook>
    );
    try testing.expect(worded.calc.full_calc_on_load);
    try testing.expect(!worded.calc.date1904);

    // Defaults when calcPr is absent entirely.
    const bare = try parseWorkbook(a, "<workbook><sheets/></workbook>");
    try testing.expectEqualStrings("auto", bare.calc.calc_mode);
    try testing.expect(bare.calc.full_precision);
    try testing.expect(!bare.calc.has_external_references);
    try testing.expect(!bare.calc.has_external_references);
}

test "relationship targets resolve relative to xl/ and absolute" {
    var arena: std.heap.ArenaAllocator = .init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    try testing.expectEqualStrings("xl/worksheets/sheet1.xml", try resolveTarget(a, "worksheets/sheet1.xml"));
    try testing.expectEqualStrings("xl/worksheets/sheet1.xml", try resolveTarget(a, "/xl/worksheets/sheet1.xml"));
}

test "cell-ref shape check" {
    try testing.expect(cellRefOk("A1"));
    try testing.expect(cellRefOk("$A$1"));
    try testing.expect(cellRefOk("XFD1048576"));
    try testing.expect(!cellRefOk(null));
    try testing.expect(!cellRefOk(""));
    try testing.expect(!cellRefOk("A"));
    try testing.expect(!cellRefOk("1"));
    try testing.expect(!cellRefOk("A1:B2"));
    try testing.expect(!cellRefOk("A1x"));
}
