//! Semantic oracle manifests (M1b, `goal_formula.md` §8.2 / §8.3).
//!
//! A manifest is what an oracle run leaves behind: decoded typed values
//! with their binary64 bit patterns, normalized error spellings, and a
//! provenance record. It is deliberately NOT a copy of the XML — an
//! oracle that recorded `<v>0.1</v>` would be recording a decimal
//! rendering, and §8.3's comparison rule is *bit-exact parsed
//! binary64*. Two runs that print `0.1` differently but parse to the
//! same bits agree; two that print it identically from different bits
//! do not, and only the bits reveal that.
//!
//! Three policies are pinned here because getting any of them wrong
//! makes every downstream comparison quietly wrong:
//!
//!  * **Signed zero is significant.** `-0.0` and `+0.0` are different
//!    bit patterns and therefore different manifest entries. Excel
//!    normalizes most negative zeros away on write, so seeing one is
//!    information, not noise.
//!  * **NaN and infinity are hard errors.** Neither is representable in
//!    a `<v>`; encountering one means the extractor, the application,
//!    or our reading of the file is wrong. Recording it would pin a
//!    nonsense expectation, so the manifest refuses instead.
//!  * **Error spellings normalize to a closed set, but unknown
//!    spellings are preserved verbatim** — the same open-set treatment
//!    M1a gave error literals in the tokenizer. Excel keeps adding
//!    errors; a manifest that dropped `#BLOCKED!` on the floor would
//!    silently record a blank where an error belongs.

const std = @import("std");
const extractor = @import("extractor.zig");
const provenance = @import("provenance.zig");

pub const schema = "zlsx-oracle-manifest-1";

pub const Error = error{
    ManifestNonFiniteValue,
    ManifestUnparsableNumber,
    ManifestBadBits,
    ManifestSchemaMismatch,
    ManifestUnknownKind,
    ManifestUnknownFidelity,
    ManifestDigestMismatch,
} || provenance.Error || std.mem.Allocator.Error;

/// The ten error spellings `goal_formula.md` freezes. Anything else is
/// carried as `.unknown` with its bytes intact.
pub const ErrorKind = enum {
    div0,
    na,
    name,
    null_value,
    num,
    ref,
    value,
    spill,
    calc,
    getting_data,
    /// A spelling outside the frozen ten — rich or future errors.
    unknown,

    /// Map a raw `<v>` spelling from a `t="e"` cell to its canonical
    /// kind. Matching is exact: an oracle that accepted near-misses
    /// would paper over a locale-translated error, which is precisely
    /// the divergence we want to see.
    pub fn normalize(spelling: []const u8) ErrorKind {
        const table = .{
            .{ "#DIV/0!", ErrorKind.div0 },
            .{ "#N/A", ErrorKind.na },
            .{ "#NAME?", ErrorKind.name },
            .{ "#NULL!", ErrorKind.null_value },
            .{ "#NUM!", ErrorKind.num },
            .{ "#REF!", ErrorKind.ref },
            .{ "#VALUE!", ErrorKind.value },
            .{ "#SPILL!", ErrorKind.spill },
            .{ "#CALC!", ErrorKind.calc },
            .{ "#GETTING_DATA", ErrorKind.getting_data },
        };
        inline for (table) |row| {
            if (std.mem.eql(u8, spelling, row[0])) return row[1];
        }
        return .unknown;
    }
};

/// Which comparison regime a manifest belongs to (§8.2 precedence,
/// §8.4 fidelity levels). The same workbook can have goldens in both,
/// and they are allowed to disagree — that disagreement is the record.
pub const Fidelity = enum {
    /// Match Excel, including its documented departures from IEEE-754.
    excel,
    /// Match the normative `ieee_fp_rules_v1` table; Excel is retained
    /// only as a recorded divergence witness.
    ieee,

    pub fn parse(s: []const u8) Error!Fidelity {
        return std.meta.stringToEnum(Fidelity, s) orelse error.ManifestUnknownFidelity;
    }
};

/// Why a cell was excluded from a manifest's value assertions. Recorded
/// rather than dropped: "there is no golden for this cell" and "nobody
/// looked at this cell" must never look the same.
pub const Exclusion = enum {
    /// `ca="1"` or a known volatile function. §8.2 excludes volatiles
    /// from EVERY external value oracle — they are covered by KATs,
    /// draw-count invariants and type/range checks instead.
    volatile_formula,
    /// CHAR/CODE over 127: the Mac build's code page differs from the
    /// CP-1252 spec cases, so Mac goldens would pin the wrong answer.
    char_code_high_byte,
    /// The workbook screen disqualified the whole file.
    screened_out,
};

pub const CellEntry = struct {
    sheet: []const u8,
    ref: []const u8,
    /// "number" | "text" | "boolean" | "error" | "blank"
    kind: []const u8,
    /// Numbers: the IEEE-754 bit pattern as `0x`-prefixed hex. Hex
    /// because JSON numbers cannot carry 64 bits losslessly, and the
    /// bits ARE the value under §8.3.
    bits: ?[]const u8 = null,
    /// Numbers: the source spelling, kept for human review of diffs.
    /// Never used for comparison.
    source: ?[]const u8 = null,
    /// Text cells: the fully decoded string.
    text: ?[]const u8 = null,
    boolean: ?bool = null,
    /// Error cells: canonical kind name plus the raw spelling.
    error_kind: ?[]const u8 = null,
    error_spelling: ?[]const u8 = null,
    /// Formula text, if the cell carries one.
    formula: ?[]const u8 = null,
    /// Set when the cell is present but carries no value assertion.
    excluded: ?[]const u8 = null,

    pub fn numberBits(self: CellEntry) Error!u64 {
        const s = self.bits orelse return error.ManifestBadBits;
        if (!std.mem.startsWith(u8, s, "0x")) return error.ManifestBadBits;
        return std.fmt.parseInt(u64, s[2..], 16) catch error.ManifestBadBits;
    }

    pub fn value(self: CellEntry) Error!f64 {
        return @bitCast(try self.numberBits());
    }
};

pub const CalcState = struct {
    calc_mode: []const u8,
    full_calc_on_load: bool,
    full_precision: bool,
    date1904: bool,
};

pub const Manifest = struct {
    schema: []const u8,
    /// Stable identifier for the case this manifest covers, e.g.
    /// "sentinel_stale_value" or "spec_arith".
    case: []const u8,
    fidelity: []const u8,
    provenance: provenance.Record,
    calc: CalcState,
    cells: []const CellEntry,

    pub fn validate(self: Manifest) Error!void {
        if (!std.mem.eql(u8, self.schema, schema)) return error.ManifestSchemaMismatch;
        _ = try Fidelity.parse(self.fidelity);
        try self.provenance.validate();
        for (self.cells) |c| try validateCell(c);
    }

    pub fn find(self: Manifest, sheet: []const u8, ref: []const u8) ?CellEntry {
        for (self.cells) |c| {
            if (std.mem.eql(u8, c.sheet, sheet) and std.mem.eql(u8, c.ref, ref)) return c;
        }
        return null;
    }

    /// Cells carrying an actual value assertion — i.e. not excluded.
    pub fn assertedCount(self: Manifest) usize {
        var n: usize = 0;
        for (self.cells) |c| {
            if (c.excluded == null) n += 1;
        }
        return n;
    }
};

fn validateCell(c: CellEntry) Error!void {
    if (c.sheet.len == 0 or c.ref.len == 0) return error.ManifestUnknownKind;
    if (c.excluded) |reason| {
        _ = std.meta.stringToEnum(Exclusion, reason) orelse return error.ManifestUnknownKind;
        return; // an excluded cell asserts nothing, so nothing else to check
    }
    if (std.mem.eql(u8, c.kind, "number")) {
        const bits = try c.numberBits();
        const f: f64 = @bitCast(bits);
        // The hard-error policy, enforced at the manifest boundary so a
        // non-finite can never reach a comparison.
        if (std.math.isNan(f) or std.math.isInf(f)) return error.ManifestNonFiniteValue;
    } else if (std.mem.eql(u8, c.kind, "text")) {
        if (c.text == null) return error.ManifestUnknownKind;
    } else if (std.mem.eql(u8, c.kind, "boolean")) {
        if (c.boolean == null) return error.ManifestUnknownKind;
    } else if (std.mem.eql(u8, c.kind, "error")) {
        const spelling = c.error_spelling orelse return error.ManifestUnknownKind;
        const kind_name = c.error_kind orelse return error.ManifestUnknownKind;
        const declared = std.meta.stringToEnum(ErrorKind, kind_name) orelse
            return error.ManifestUnknownKind;
        // The normalization must agree with what was recorded, so a
        // hand-edited manifest cannot claim `#VALUE!` is `div0`.
        if (ErrorKind.normalize(spelling) != declared) return error.ManifestUnknownKind;
    } else if (std.mem.eql(u8, c.kind, "blank")) {
        // nothing further
    } else return error.ManifestUnknownKind;
}

/// Parse and validate a manifest. Caller owns `parsed`; call
/// `parsed.deinit()`.
///
/// `alloc_always` is not a performance footnote. The default,
/// `alloc_if_needed`, makes every escape-free string BORROW from the
/// input buffer — so a caller that frees the JSON bytes after parsing
/// is left holding a `Manifest` full of dangling slices that still
/// compare and print as if they were fine.
pub fn parse(allocator: std.mem.Allocator, json: []const u8) !std.json.Parsed(Manifest) {
    const parsed = try std.json.parseFromSlice(Manifest, allocator, json, .{
        .ignore_unknown_fields = false,
        .allocate = .alloc_always,
    });
    errdefer parsed.deinit();
    try parsed.value.validate();
    return parsed;
}

// ─── building a manifest from an extraction ──────────────────────

pub const BuildOptions = struct {
    case: []const u8,
    fidelity: Fidelity,
    prov: provenance.Record,
    /// Volatile formulas are excluded from external value oracles
    /// (§8.2). Defaulting this to the adapter's own answer means a
    /// caller cannot forget it.
    exclude_volatiles: ?bool = null,
    /// Mac Excel's CHAR/CODE above 127 follow the Mac code page, not
    /// CP-1252, so those goldens must not be recorded from it.
    exclude_char_code_high: bool = false,
};

/// Turn an extraction into a manifest. Allocates into `allocator`;
/// everything borrows from `wb`, which must outlive the result.
pub fn build(
    allocator: std.mem.Allocator,
    wb: extractor.Workbook,
    opts: BuildOptions,
) Error!Manifest {
    const adapter = try opts.prov.adapterEnum();
    const exclude_volatiles = opts.exclude_volatiles orelse adapter.isExternalApp();

    var cells: std.ArrayListUnmanaged(CellEntry) = .empty;
    errdefer cells.deinit(allocator);

    for (wb.cells) |c| {
        var entry: CellEntry = .{
            .sheet = c.sheet,
            .ref = c.ref,
            .kind = "blank",
            .formula = if (c.formula) |f| (if (f.text.len > 0) f.text else null) else null,
        };

        if (exclude_volatiles and isVolatile(c)) {
            entry.excluded = @tagName(Exclusion.volatile_formula);
            try cells.append(allocator, entry);
            continue;
        }
        if (opts.exclude_char_code_high and mentionsCharCode(c)) {
            entry.excluded = @tagName(Exclusion.char_code_high_byte);
            try cells.append(allocator, entry);
            continue;
        }

        switch (c.kind) {
            .number => {
                const raw = c.value orelse {
                    try cells.append(allocator, entry);
                    continue;
                };
                const f = parseNumber(raw) catch return error.ManifestUnparsableNumber;
                if (std.math.isNan(f) or std.math.isInf(f)) return error.ManifestNonFiniteValue;
                entry.kind = "number";
                entry.source = raw;
                entry.bits = try std.fmt.allocPrint(
                    allocator,
                    "0x{X:0>16}",
                    .{@as(u64, @bitCast(f))},
                );
            },
            .boolean => {
                const raw = c.value orelse "0";
                entry.kind = "boolean";
                entry.boolean = std.mem.eql(u8, std.mem.trim(u8, raw, " \t\r\n"), "1");
            },
            .err => {
                const raw = c.value orelse "";
                entry.kind = "error";
                entry.error_spelling = raw;
                entry.error_kind = @tagName(ErrorKind.normalize(raw));
            },
            .shared_string, .inline_string, .formula_string => {
                entry.kind = "text";
                entry.text = c.text orelse "";
            },
            .iso_date => {
                // `t="d"` is deferred to M4b2 in the evaluator. Record
                // it as text so the manifest is honest about what the
                // file held, rather than inventing a serial number.
                entry.kind = "text";
                entry.text = c.value orelse "";
            },
            .blank => {},
        }
        try cells.append(allocator, entry);
    }

    return .{
        .schema = schema,
        .case = opts.case,
        .fidelity = @tagName(opts.fidelity),
        .provenance = opts.prov,
        .calc = .{
            .calc_mode = wb.calc.calc_mode,
            .full_calc_on_load = wb.calc.full_calc_on_load,
            .full_precision = wb.calc.full_precision,
            .date1904 = wb.calc.date1904,
        },
        .cells = try cells.toOwnedSlice(allocator),
    };
}

/// Parse a `<v>` number. OOXML writes plain decimal or scientific
/// notation; `parseFloat` covers both. Rejecting `nan`/`inf` spellings
/// here rather than after conversion keeps the hard-error policy at the
/// boundary where the text is still visible for a diagnostic.
fn parseNumber(raw: []const u8) !f64 {
    const trimmed = std.mem.trim(u8, raw, " \t\r\n");
    if (trimmed.len == 0) return error.ManifestUnparsableNumber;
    for (trimmed) |ch| {
        const ok = std.ascii.isDigit(ch) or ch == '.' or ch == '-' or ch == '+' or
            ch == 'e' or ch == 'E';
        if (!ok) return error.ManifestUnparsableNumber;
    }
    return std.fmt.parseFloat(f64, trimmed);
}

/// Volatile by declaration (`ca="1"`) or by name. Excel sets `ca` on
/// most volatile calls but not reliably on every one, so the name check
/// backs it up — an unexcluded volatile is a golden that fails on the
/// next run for no reason anybody can reproduce.
fn isVolatile(c: extractor.Cell) bool {
    const f = c.formula orelse return false;
    if (f.always_calc) return true;
    const names = [_][]const u8{
        "RAND(",   "RANDBETWEEN(", "RANDARRAY(", "NOW(",  "TODAY(",
        "OFFSET(", "INDIRECT(",    "CELL(",      "INFO(", "AREAS(",
    };
    for (names) |n| {
        if (containsIgnoreCase(f.text, n)) return true;
    }
    return false;
}

fn mentionsCharCode(c: extractor.Cell) bool {
    const f = c.formula orelse return false;
    return containsIgnoreCase(f.text, "CHAR(") or containsIgnoreCase(f.text, "CODE(") or
        containsIgnoreCase(f.text, "UNICHAR(") or containsIgnoreCase(f.text, "UNICODE(");
}

fn containsIgnoreCase(haystack: []const u8, needle: []const u8) bool {
    if (needle.len > haystack.len) return false;
    var i: usize = 0;
    while (i + needle.len <= haystack.len) : (i += 1) {
        if (std.ascii.eqlIgnoreCase(haystack[i .. i + needle.len], needle)) return true;
    }
    return false;
}

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

test "error spellings normalize to the frozen ten" {
    const cases = .{
        .{ "#DIV/0!", ErrorKind.div0 },
        .{ "#N/A", ErrorKind.na },
        .{ "#NAME?", ErrorKind.name },
        .{ "#NULL!", ErrorKind.null_value },
        .{ "#NUM!", ErrorKind.num },
        .{ "#REF!", ErrorKind.ref },
        .{ "#VALUE!", ErrorKind.value },
        .{ "#SPILL!", ErrorKind.spill },
        .{ "#CALC!", ErrorKind.calc },
        .{ "#GETTING_DATA", ErrorKind.getting_data },
    };
    inline for (cases) |c| try testing.expectEqual(c[1], ErrorKind.normalize(c[0]));
}

test "unknown error spellings stay unknown rather than being coerced" {
    // Same open-set treatment M1a gave the tokenizer: Excel keeps
    // shipping new errors, and coercing `#BLOCKED!` into the nearest
    // known kind would record a wrong expectation.
    for ([_][]const u8{ "#BLOCKED!", "#PYTHON!", "#BUSY!", "#CONNECT!", "#EXTERNAL!" }) |s| {
        try testing.expectEqual(ErrorKind.unknown, ErrorKind.normalize(s));
    }
    // Near-misses are NOT accepted — a locale-translated spelling must
    // surface as a divergence, not be silently folded in.
    for ([_][]const u8{ "#div/0!", "#DIV/0", "DIV/0!", "#WERT!", "" }) |s| {
        try testing.expectEqual(ErrorKind.unknown, ErrorKind.normalize(s));
    }
}

test "numbers record bit patterns, and signed zero is significant" {
    const pos: CellEntry = .{ .sheet = "S", .ref = "A1", .kind = "number", .bits = "0x0000000000000000" };
    const neg: CellEntry = .{ .sheet = "S", .ref = "A2", .kind = "number", .bits = "0x8000000000000000" };
    try validateCell(pos);
    try validateCell(neg);

    try testing.expectEqual(@as(f64, 0.0), try pos.value());
    try testing.expectEqual(@as(f64, -0.0), try neg.value());
    // `0.0 == -0.0` numerically, which is exactly why the manifest
    // compares BITS: the two entries must not be interchangeable.
    try testing.expect(try pos.numberBits() != try neg.numberBits());
    try testing.expect(std.math.signbit(try neg.value()));
    try testing.expect(!std.math.signbit(try pos.value()));
}

test "NaN and infinity are refused at the manifest boundary" {
    const nan: CellEntry = .{ .sheet = "S", .ref = "A1", .kind = "number", .bits = "0x7FF8000000000000" };
    try testing.expectError(error.ManifestNonFiniteValue, validateCell(nan));

    const inf: CellEntry = .{ .sheet = "S", .ref = "A1", .kind = "number", .bits = "0x7FF0000000000000" };
    try testing.expectError(error.ManifestNonFiniteValue, validateCell(inf));

    const neg_inf: CellEntry = .{ .sheet = "S", .ref = "A1", .kind = "number", .bits = "0xFFF0000000000000" };
    try testing.expectError(error.ManifestNonFiniteValue, validateCell(neg_inf));

    // The largest finite double is fine.
    const max: CellEntry = .{ .sheet = "S", .ref = "A1", .kind = "number", .bits = "0x7FEFFFFFFFFFFFFF" };
    try validateCell(max);
}

test "bit strings must be 0x-prefixed hex" {
    for ([_][]const u8{ "4008000000000000", "0xZZ", "", "0x" }) |bad| {
        const c: CellEntry = .{ .sheet = "S", .ref = "A1", .kind = "number", .bits = bad };
        try testing.expectError(error.ManifestBadBits, c.numberBits());
    }
    const missing: CellEntry = .{ .sheet = "S", .ref = "A1", .kind = "number" };
    try testing.expectError(error.ManifestBadBits, missing.numberBits());
}

test "an error entry cannot lie about its normalization" {
    const honest: CellEntry = .{
        .sheet = "S",
        .ref = "A1",
        .kind = "error",
        .error_spelling = "#DIV/0!",
        .error_kind = "div0",
    };
    try validateCell(honest);

    const lying: CellEntry = .{
        .sheet = "S",
        .ref = "A1",
        .kind = "error",
        .error_spelling = "#VALUE!",
        .error_kind = "div0",
    };
    try testing.expectError(error.ManifestUnknownKind, validateCell(lying));
}

test "0.1 + 0.2 records the bits, not the rendering" {
    // The case §8.3 exists for. Two implementations can both print
    // "0.30000000000000004" while holding different bits, and can print
    // differently while holding the same bits. Only the bits decide.
    // Through mutable vars deliberately: Zig folds `0.1 + 0.2` at
    // comptime in arbitrary precision, which lands exactly on 0.3 and
    // would make this test assert the opposite of what it means to.
    var x: f64 = 0.1;
    var y: f64 = 0.2;
    _ = &x;
    _ = &y;
    const sum: f64 = x + y;
    var buf: [32]u8 = undefined;
    const hex = try std.fmt.bufPrint(&buf, "0x{X:0>16}", .{@as(u64, @bitCast(sum))});
    const c: CellEntry = .{ .sheet = "S", .ref = "A1", .kind = "number", .bits = hex };
    try validateCell(c);
    try testing.expectEqual(sum, try c.value());
    try testing.expect(try c.value() != 0.3);
}

test "parseNumber accepts OOXML spellings and refuses everything else" {
    try testing.expectEqual(@as(f64, 42), try parseNumber("42"));
    try testing.expectEqual(@as(f64, 1.5), try parseNumber("1.5"));
    try testing.expectEqual(@as(f64, -0.25), try parseNumber("-0.25"));
    try testing.expectEqual(@as(f64, 2.1211775296413049e-2), try parseNumber("2.1211775296413049E-2"));
    try testing.expectEqual(@as(f64, 1000), try parseNumber(" 1000 "));

    // `parseFloat` would happily accept these; the manifest must not.
    for ([_][]const u8{ "nan", "inf", "-inf", "NaN", "0x1p3", "abc", "", "  " }) |bad| {
        try testing.expectError(error.ManifestUnparsableNumber, parseNumber(bad));
    }
}

test "volatile detection covers ca=1 and the volatile function names" {
    const declared: extractor.Cell = .{
        .sheet = "S",
        .ref = "A1",
        .kind = .number,
        .value = "0.5",
        .text = null,
        .formula = .{ .text = "SOMETHING()", .kind = null, .ref = null, .si = null, .always_calc = true },
    };
    try testing.expect(isVolatile(declared));

    for ([_][]const u8{ "RAND()", "1+NOW()", "TODAY()", "offset(A1,1,1)", "SUM(INDIRECT(\"A1\"))" }) |f| {
        const c: extractor.Cell = .{
            .sheet = "S",
            .ref = "A1",
            .kind = .number,
            .value = "1",
            .text = null,
            .formula = .{ .text = f, .kind = null, .ref = null, .si = null, .always_calc = false },
        };
        try testing.expect(isVolatile(c));
    }

    const plain: extractor.Cell = .{
        .sheet = "S",
        .ref = "A1",
        .kind = .number,
        .value = "3",
        .text = null,
        .formula = .{ .text = "SUM(A1:A2)", .kind = null, .ref = null, .si = null, .always_calc = false },
    };
    try testing.expect(!isVolatile(plain));

    // A cell with no formula at all is not volatile.
    const literal: extractor.Cell = .{
        .sheet = "S",
        .ref = "A1",
        .kind = .number,
        .value = "3",
        .text = null,
        .formula = null,
    };
    try testing.expect(!isVolatile(literal));
}

test "round-trips through JSON" {
    const json =
        \\{
        \\  "schema": "zlsx-oracle-manifest-1",
        \\  "case": "spec_arith",
        \\  "fidelity": "excel",
        \\  "provenance": {
        \\    "adapter": "hand_spec",
        \\    "app_build": "MS-OE376 §18.17",
        \\    "os": "n/a",
        \\    "locale": "en_US.UTF-8",
        \\    "extractor_version": "oracle-extractor-1",
        \\    "workbook_digest": "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef",
        \\    "recorded": "2026-08-03"
        \\  },
        \\  "calc": {
        \\    "calc_mode": "auto",
        \\    "full_calc_on_load": true,
        \\    "full_precision": true,
        \\    "date1904": false
        \\  },
        \\  "cells": [
        \\    {"sheet":"Sheet1","ref":"A1","kind":"number","bits":"0x3FF0000000000000","source":"1"},
        \\    {"sheet":"Sheet1","ref":"A2","kind":"error","error_kind":"div0","error_spelling":"#DIV/0!","formula":"1/0"},
        \\    {"sheet":"Sheet1","ref":"A3","kind":"text","text":"hello"},
        \\    {"sheet":"Sheet1","ref":"A4","kind":"boolean","boolean":true},
        \\    {"sheet":"Sheet1","ref":"A5","kind":"blank"},
        \\    {"sheet":"Sheet1","ref":"A6","kind":"number","excluded":"volatile_formula","formula":"RAND()"}
        \\  ]
        \\}
    ;
    const parsed = try parse(testing.allocator, json);
    defer parsed.deinit();
    const m = parsed.value;

    try testing.expectEqual(@as(usize, 6), m.cells.len);
    try testing.expectEqual(@as(usize, 5), m.assertedCount());
    try testing.expectEqual(@as(f64, 1.0), try m.find("Sheet1", "A1").?.value());
    try testing.expectEqualStrings("div0", m.find("Sheet1", "A2").?.error_kind.?);
    try testing.expectEqualStrings("hello", m.find("Sheet1", "A3").?.text.?);
    try testing.expect(m.find("Sheet1", "A4").?.boolean.?);
    try testing.expectEqualStrings("volatile_formula", m.find("Sheet1", "A6").?.excluded.?);
    try testing.expectEqual(Fidelity.excel, try Fidelity.parse(m.fidelity));
    try testing.expect(m.find("Sheet1", "Z99") == null);
}

test "a manifest with the wrong schema tag is refused" {
    const json =
        \\{"schema":"something-else","case":"x","fidelity":"excel",
        \\ "provenance":{"adapter":"hand_spec","app_build":"a","os":"b","locale":"c",
        \\ "extractor_version":"d","workbook_digest":"0000000000000000000000000000000000000000000000000000000000000000","recorded":"e"},
        \\ "calc":{"calc_mode":"auto","full_calc_on_load":false,"full_precision":true,"date1904":false},
        \\ "cells":[]}
    ;
    try testing.expectError(error.ManifestSchemaMismatch, parse(testing.allocator, json));
}

test "builds a manifest from an extraction, excluding volatiles for app adapters" {
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

    var wb = try extractor.extract(testing.allocator, bytes);
    defer wb.deinit();

    const prov: provenance.Record = .{
        .adapter = "excel_mac",
        .app_build = "test",
        .os = "test",
        .locale = "en_US.UTF-8",
        .extractor_version = extractor.version,
        .workbook_digest = &wb.digest,
        .recorded = "2026-08-03",
    };
    var arena: std.heap.ArenaAllocator = .init(testing.allocator);
    defer arena.deinit();

    const m = try build(arena.allocator(), wb, .{
        .case = "corpus_openxlsx",
        .fidelity = .excel,
        .prov = prov,
    });
    try m.validate();

    // This workbook is full of `RAND()`; every one must be excluded,
    // and the exclusions must be RECORDED, not dropped.
    var excluded: usize = 0;
    for (m.cells) |c| {
        if (c.excluded) |reason| {
            try testing.expectEqualStrings("volatile_formula", reason);
            excluded += 1;
        }
    }
    try testing.expect(excluded > 0);
    try testing.expect(m.assertedCount() > 0);
    try testing.expectEqual(m.cells.len, m.assertedCount() + excluded);
}

test "hand_spec keeps volatiles: a documented contract is not a draw" {
    var arena: std.heap.ArenaAllocator = .init(testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    var wb: extractor.Workbook = .{
        .arena = .init(testing.allocator),
        .sheets = try a.dupe([]const u8, &.{"S"}),
        .cells = try a.dupe(extractor.Cell, &.{.{
            .sheet = "S",
            .ref = "A1",
            .kind = .number,
            .value = "0.5",
            .text = null,
            .formula = .{ .text = "RAND()", .kind = null, .ref = null, .si = null, .always_calc = true },
        }}),
        .calc = .{},
        .digest = "0".* ** 64,
    };
    defer wb.deinit();

    const app = try build(a, wb, .{
        .case = "c",
        .fidelity = .excel,
        .prov = .{
            .adapter = "excel_mac",
            .app_build = "x",
            .os = "x",
            .locale = "x",
            .extractor_version = "x",
            .workbook_digest = "0" ** 64,
            .recorded = "x",
        },
    });
    try testing.expectEqualStrings("volatile_formula", app.cells[0].excluded.?);

    const spec = try build(a, wb, .{
        .case = "c",
        .fidelity = .excel,
        .prov = .{
            .adapter = "hand_spec",
            .app_build = "x",
            .os = "x",
            .locale = "x",
            .extractor_version = "x",
            .workbook_digest = "0" ** 64,
            .recorded = "x",
        },
    });
    try testing.expect(spec.cells[0].excluded == null);
}
