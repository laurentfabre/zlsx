//! workbook.xml fresh-emit plan (B3 iter-wr-3).
//!
//! Storage + validation + fresh-emit for `xl/workbook.xml`. Lives
//! outside `pkg/workbook.zig` so both Workbook (delta-on-existing-bytes
//! editor) and `xlsx.Writer` (fresh-emit producer) can stage workbook-
//! level state through the same shape without a circular module
//! dependency (workbook.zig imports `zlsx`, which contains writer.zig).
//!
//! Today this module owns:
//!   - the `DefinedName` record + its full Excel name-rule validator
//!     (`validateDefinedName`) lifted from `src/writer.zig` —
//!     R1C1 reject, A1-shape reject, case-insensitive duplicate reject.
//!   - the `WorkbookXmlPlan` registry that holds defined names through
//!     `addDefinedName` (validates → dedups → dups → appends).
//!   - `emitWorkbookXml` which renders `<workbook>…<sheets>…</sheets>
//!     [<definedNames>…</definedNames>]</workbook>` byte-for-byte
//!     compatible with the previous Writer-local emit branch.
//!
//! Stdlib only. Zig 0.15.2.

const std = @import("std");
const coords = @import("zlsx_refs");

const Allocator = std.mem.Allocator;

pub const Error = error{
    OutOfMemory,
    /// Empty / >255 chars / illegal first char / illegal interior char
    /// / single `R` / `C` / A1-shaped / R1C1-shaped — caught up front
    /// at `addDefinedName` so workbooks never get to save with names
    /// Excel will refuse on open.
    InvalidDefinedName,
    /// `refers_to` was empty. The full formula tokenizer
    /// (`src/formula/tokenizer.zig`) is the canonical surface for
    /// callers wanting structural validation; this checker is
    /// non-empty-only.
    InvalidDefinedNameRefersTo,
    /// Same name (case-insensitive) within the same scope — Excel
    /// treats workbook-scope and per-sheet scopes independently.
    DuplicateDefinedName,
    /// `local_sheet_id >= sheet_count` at emit time. Surfaced at
    /// `emitWorkbookXml` (not at `addDefinedName`) because the sheet
    /// count is the producer's invariant, not the plan's.
    InvalidDefinedNameLocalSheetId,
};

pub const DefinedName = struct {
    /// Owned by the plan allocator. Caller frees their staging copy
    /// immediately after `addDefinedName` returns.
    name: []const u8,
    /// Owned by the plan allocator.
    refers_to: []const u8,
    /// 0-based sheet index. `null` ⇒ workbook-scoped.
    local_sheet_id: ?u32 = null,
    /// Hidden names don't appear in Excel's Name Manager UI but still
    /// resolve in formulas (e.g. `_xlnm.Print_Area`).
    hidden: bool = false,
};

pub const DefinedNameOptions = struct {
    local_sheet_id: ?u32 = null,
    hidden: bool = false,
};

/// Plan registry for `xl/workbook.xml` fresh-emit. Today owns the
/// defined-name pool. Future iters extend with workbook-level
/// `<workbookPr>` / `<calcPr>` / `<fileVersion>` shape staging.
pub const WorkbookXmlPlan = struct {
    /// Allocator-owned vector of defined names. Each entry's
    /// `name` and `refers_to` slices are duped into the plan's
    /// allocator at `addDefinedName` time.
    defined_names: std.ArrayListUnmanaged(DefinedName) = .empty,

    pub fn deinit(self: *WorkbookXmlPlan, allocator: Allocator) void {
        for (self.defined_names.items) |dn| {
            allocator.free(dn.name);
            allocator.free(dn.refers_to);
        }
        self.defined_names.deinit(allocator);
        self.* = undefined;
    }

    /// Register a workbook-level defined name. Validates the name
    /// shape against Excel's rules, rejects empty `refers_to`,
    /// rejects case-insensitive duplicates within the same scope.
    /// On success the plan owns duped copies of `name` and
    /// `refers_to`; the caller can free their staging buffers.
    ///
    /// Atomicity: validation + dup-scan run before any allocation;
    /// an OOM during the dup of `refers_to` walks the dup of `name`
    /// back via errdefer, leaving the plan untouched.
    pub fn addDefinedName(
        self: *WorkbookXmlPlan,
        allocator: Allocator,
        name: []const u8,
        refers_to: []const u8,
        opts: DefinedNameOptions,
    ) Error!void {
        try validateDefinedName(name);
        if (refers_to.len == 0) return error.InvalidDefinedNameRefersTo;
        for (self.defined_names.items) |existing| {
            if (existing.local_sheet_id != opts.local_sheet_id) continue;
            if (std.ascii.eqlIgnoreCase(existing.name, name)) {
                return error.DuplicateDefinedName;
            }
        }
        const name_copy = try allocator.dupe(u8, name);
        errdefer allocator.free(name_copy);
        const refers_copy = try allocator.dupe(u8, refers_to);
        errdefer allocator.free(refers_copy);
        try self.defined_names.append(allocator, .{
            .name = name_copy,
            .refers_to = refers_copy,
            .local_sheet_id = opts.local_sheet_id,
            .hidden = opts.hidden,
        });
    }
};

/// One sheet in declared (= file/tab) order. The fresh-emit walker
/// builds `<sheet name=… sheetId=N r:id="rIdN"/>` in this order; both
/// `sheetId` and the `r:id` numeric tail are 1 + the slice index.
pub const SheetEntry = struct {
    /// Raw sheet name. Will be XML-attribute-escaped on emit; the
    /// caller must NOT pre-escape.
    name: []const u8,
};

/// Emit a fresh `xl/workbook.xml` payload byte-for-byte equivalent to
/// `xlsx.Writer.save`'s prior local emit branch. Returns the bytes —
/// caller takes ownership.
///
/// Element order: `<?xml…?><workbook…><sheets>…</sheets>
/// [<definedNames>…</definedNames>]</workbook>`. The
/// `<definedNames>` block is OMITTED entirely when the plan has no
/// names — preserves the pre-iter-wr-3 byte output for workbooks
/// that don't register any.
pub fn emitWorkbookXml(
    allocator: Allocator,
    sheets: []const SheetEntry,
    plan: *const WorkbookXmlPlan,
) Error![]u8 {
    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(allocator);

    try out.appendSlice(allocator, WORKBOOK_HEAD);
    for (sheets, 0..) |sw, i| {
        try out.appendSlice(allocator, "<sheet name=\"");
        try appendXmlEscaped(allocator, &out, sw.name);
        try out.print(allocator, "\" sheetId=\"{d}\" r:id=\"rId{d}\"/>", .{ i + 1, i + 1 });
    }
    try out.appendSlice(allocator, WORKBOOK_SHEETS_CLOSE);

    if (plan.defined_names.items.len > 0) {
        try out.appendSlice(allocator, "<definedNames>");
        for (plan.defined_names.items) |dn| {
            if (dn.local_sheet_id) |sid| {
                if (sid >= sheets.len) return error.InvalidDefinedNameLocalSheetId;
            }
            try out.appendSlice(allocator, "<definedName name=\"");
            try appendXmlEscaped(allocator, &out, dn.name);
            try out.appendSlice(allocator, "\"");
            if (dn.local_sheet_id) |sid| {
                try out.print(allocator, " localSheetId=\"{d}\"", .{sid});
            }
            if (dn.hidden) try out.appendSlice(allocator, " hidden=\"1\"");
            try out.appendSlice(allocator, ">");
            try appendXmlEscaped(allocator, &out, dn.refers_to);
            try out.appendSlice(allocator, "</definedName>");
        }
        try out.appendSlice(allocator, "</definedNames>");
    }
    try out.appendSlice(allocator, WORKBOOK_END);
    return out.toOwnedSlice(allocator);
}

// ─── OOXML skeleton strings ─────────────────────────────────────────

const WORKBOOK_HEAD: []const u8 =
    \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    \\<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>
;
const WORKBOOK_SHEETS_CLOSE: []const u8 = "</sheets>";
const WORKBOOK_END: []const u8 = "</workbook>";

// ─── Helpers ────────────────────────────────────────────────────────

/// Excel defined-name rules (per MS docs):
/// - 1..255 bytes
/// - First char: letter, `_`, or `\`
/// - Rest: letters, digits, `_`, `.` (NO `\`, NO `?`)
/// - Must NOT exactly match an A1 cell reference shape (column
///   [A, XFD] + row [1, 1048576] in any case combination)
/// - Must NOT be the single letter `R` or `C` (case-insensitive)
///   — Excel reserves these for R1C1 row/column references
/// We don't enforce the "not equal to a built-in name" rule because
/// the `_xlnm.Print_Area` family is intentionally usable.
pub fn validateDefinedName(name: []const u8) Error!void {
    if (name.len == 0 or name.len > 255) return error.InvalidDefinedName;
    const first = name[0];
    if (!isAsciiLetter(first) and first != '_' and first != '\\') {
        return error.InvalidDefinedName;
    }
    for (name[1..]) |c| {
        if (!isAsciiLetter(c) and !isAsciiDigit(c) and c != '_' and c != '.') {
            return error.InvalidDefinedName;
        }
    }
    if (name.len == 1 and (first == 'R' or first == 'r' or first == 'C' or first == 'c')) {
        return error.InvalidDefinedName;
    }
    if (looksLikeCellRef(name)) return error.InvalidDefinedName;
    // Excel also reserves R1C1-shaped names (`R[<digits>]C[<digits>]`,
    // case-insensitive). Bare `R` / `C` were already rejected above.
    if (looksLikeR1C1Ref(name)) return error.InvalidDefinedName;
}

/// True if `s` matches the case-insensitive R1C1 reference shape:
/// `R<digits>C<digits>`, `R<digits>C`, `RC<digits>`, or `RC`. The
/// digits may be absent (relative reference) or present (absolute).
fn looksLikeR1C1Ref(s: []const u8) bool {
    if (s.len < 2) return false;
    const first = s[0];
    if (first != 'R' and first != 'r') return false;
    var i: usize = 1;
    // Optional digits after R.
    while (i < s.len and isAsciiDigit(s[i])) : (i += 1) {}
    if (i >= s.len) return false;
    const c = s[i];
    if (c != 'C' and c != 'c') return false;
    i += 1;
    // Optional digits after C.
    while (i < s.len and isAsciiDigit(s[i])) : (i += 1) {}
    return i == s.len;
}

/// True if `s` matches the A1 cell-ref shape (case-insensitive
/// column letters in [A, XFD], row in [1, 1048576]). Used by the
/// defined-name validator to reject names that would be parsed as
/// cell refs.
fn looksLikeCellRef(s: []const u8) bool {
    // M0 adapter over `zlsx_refs`. The old hand-rolled 3-letter cap is
    // implied by the grid ceiling: a 4-letter run necessarily exceeds
    // XFD and is rejected. Leading-zero rows ("A01") parse here, as
    // they always have on this path.
    _ = coords.parseCell(s, .{
        .case = .insensitive,
        .leading_zero_row = .accept,
    }) catch return false;
    return true;
}

inline fn isAsciiLetter(c: u8) bool {
    return (c >= 'A' and c <= 'Z') or (c >= 'a' and c <= 'z');
}

inline fn isAsciiDigit(c: u8) bool {
    return c >= '0' and c <= '9';
}

/// Append `s` to `out`, XML-escaping the five canonical entities
/// (`<`, `>`, `&`, `"`, `'`). Other bytes (including UTF-8
/// continuation bytes for non-ASCII characters) pass through verbatim.
fn appendXmlEscaped(allocator: Allocator, out: *std.ArrayListUnmanaged(u8), s: []const u8) Error!void {
    for (s) |c| switch (c) {
        '<' => try out.appendSlice(allocator, "&lt;"),
        '>' => try out.appendSlice(allocator, "&gt;"),
        '&' => try out.appendSlice(allocator, "&amp;"),
        '"' => try out.appendSlice(allocator, "&quot;"),
        '\'' => try out.appendSlice(allocator, "&apos;"),
        else => try out.append(allocator, c),
    };
}

// ─── Tests ────────────────────────────────────────────────────────────

test "validateDefinedName: accepts valid names" {
    try validateDefinedName("MyName");
    try validateDefinedName("_private");
    try validateDefinedName("foo.bar.baz");
    try validateDefinedName("\\Backslashed");
    try validateDefinedName("a1_b2");
    try validateDefinedName("_xlnm.Print_Area");
}

test "validateDefinedName: rejects empty / oversized / bad-charclass" {
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName(""));
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("1Name"));
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("Bad!Name"));
    var too_long: [256]u8 = undefined;
    @memset(&too_long, 'X');
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName(&too_long));
}

test "validateDefinedName: rejects A1-shape" {
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("A1"));
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("a1"));
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("XFD1048576"));
}

test "validateDefinedName: rejects single R / C and R1C1-shape" {
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("R"));
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("c"));
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("R1C1"));
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("r10c5"));
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("RC"));
    // Names that START with R/C but aren't R1C1 shapes are accepted.
    try validateDefinedName("Range1");
    try validateDefinedName("Customer");
    try validateDefinedName("R_total");
}

test "validateDefinedName: rejects ? and \\ in trailing chars" {
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("Foo?bar"));
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("Foo\\bar"));
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("Foo?"));
    try std.testing.expectError(error.InvalidDefinedName, validateDefinedName("Foo\\"));
    try validateDefinedName("\\Foo");
    try validateDefinedName("\\Foo_bar.baz");
}

test "WorkbookXmlPlan: addDefinedName stores name + options" {
    const a = std.testing.allocator;
    var plan: WorkbookXmlPlan = .{};
    defer plan.deinit(a);

    try plan.addDefinedName(a, "MyRange", "Sheet1!$A$1:$B$1", .{});
    try plan.addDefinedName(
        a,
        "_xlnm.Print_Area",
        "Sheet1!$A$1:$B$1",
        .{ .local_sheet_id = 0, .hidden = true },
    );
    try std.testing.expectEqual(@as(usize, 2), plan.defined_names.items.len);
    try std.testing.expectEqualStrings("MyRange", plan.defined_names.items[0].name);
    try std.testing.expectEqualStrings("Sheet1!$A$1:$B$1", plan.defined_names.items[0].refers_to);
    try std.testing.expectEqual(@as(?u32, null), plan.defined_names.items[0].local_sheet_id);
    try std.testing.expectEqual(false, plan.defined_names.items[0].hidden);
    try std.testing.expectEqual(@as(?u32, 0), plan.defined_names.items[1].local_sheet_id);
    try std.testing.expectEqual(true, plan.defined_names.items[1].hidden);
}

test "WorkbookXmlPlan: rejects invalid name + empty refers_to" {
    const a = std.testing.allocator;
    var plan: WorkbookXmlPlan = .{};
    defer plan.deinit(a);

    try std.testing.expectError(
        error.InvalidDefinedName,
        plan.addDefinedName(a, "A1", "Sheet1!$A$1", .{}),
    );
    try std.testing.expectError(
        error.InvalidDefinedNameRefersTo,
        plan.addDefinedName(a, "Foo", "", .{}),
    );
}

test "WorkbookXmlPlan: rejects case-insensitive duplicates per scope" {
    const a = std.testing.allocator;
    var plan: WorkbookXmlPlan = .{};
    defer plan.deinit(a);

    try plan.addDefinedName(a, "Rate", "Sheet1!$A$1", .{});
    try std.testing.expectError(
        error.DuplicateDefinedName,
        plan.addDefinedName(a, "Rate", "Sheet1!$A$2", .{}),
    );
    try std.testing.expectError(
        error.DuplicateDefinedName,
        plan.addDefinedName(a, "RATE", "Sheet1!$A$3", .{}),
    );
    try std.testing.expectError(
        error.DuplicateDefinedName,
        plan.addDefinedName(a, "rate", "Sheet1!$A$4", .{}),
    );
    // Different scope (sheet-scoped) — accepted.
    try plan.addDefinedName(a, "Rate", "Sheet1!$B$1", .{ .local_sheet_id = 0 });
    try plan.addDefinedName(a, "Rate", "Sheet2!$B$1", .{ .local_sheet_id = 1 });
    try std.testing.expectEqual(@as(usize, 3), plan.defined_names.items.len);
}

test "emitWorkbookXml: sheets only, no definedNames block" {
    const a = std.testing.allocator;
    var plan: WorkbookXmlPlan = .{};
    defer plan.deinit(a);

    const sheets = [_]SheetEntry{
        .{ .name = "Sheet1" },
        .{ .name = "Sheet2" },
    };
    const out = try emitWorkbookXml(a, &sheets, &plan);
    defer a.free(out);

    // No `<definedNames>` substring at all.
    try std.testing.expect(std.mem.indexOf(u8, out, "<definedNames") == null);
    try std.testing.expect(std.mem.indexOf(u8, out, "<sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "<sheet name=\"Sheet2\" sheetId=\"2\" r:id=\"rId2\"/>") != null);
    try std.testing.expect(std.mem.endsWith(u8, out, "</workbook>"));
}

test "emitWorkbookXml: sheet-name XML escape" {
    const a = std.testing.allocator;
    var plan: WorkbookXmlPlan = .{};
    defer plan.deinit(a);

    const sheets = [_]SheetEntry{
        .{ .name = "R&D" },
        .{ .name = "x<y" },
    };
    const out = try emitWorkbookXml(a, &sheets, &plan);
    defer a.free(out);

    try std.testing.expect(std.mem.indexOf(u8, out, "name=\"R&amp;D\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "name=\"x&lt;y\"") != null);
}

test "emitWorkbookXml: definedNames with workbook + sheet scope + hidden" {
    const a = std.testing.allocator;
    var plan: WorkbookXmlPlan = .{};
    defer plan.deinit(a);

    const sheets = [_]SheetEntry{
        .{ .name = "Sheet1" },
        .{ .name = "Sheet2" },
    };
    try plan.addDefinedName(a, "Global", "Sheet1!$A$1", .{});
    try plan.addDefinedName(
        a,
        "_xlnm.Print_Area",
        "Sheet1!$A$1:$B$2",
        .{ .local_sheet_id = 0, .hidden = true },
    );
    try plan.addDefinedName(a, "ScopedToSheet2", "Sheet2!$C$3", .{ .local_sheet_id = 1 });

    const out = try emitWorkbookXml(a, &sheets, &plan);
    defer a.free(out);

    try std.testing.expect(std.mem.indexOf(u8, out, "<definedNames>") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "</definedNames>") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "<definedName name=\"Global\">Sheet1!$A$1</definedName>") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "localSheetId=\"0\" hidden=\"1\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "<definedName name=\"ScopedToSheet2\" localSheetId=\"1\">Sheet2!$C$3</definedName>") != null);
}

test "emitWorkbookXml: rejects out-of-range localSheetId at emit" {
    const a = std.testing.allocator;
    var plan: WorkbookXmlPlan = .{};
    defer plan.deinit(a);

    const sheets = [_]SheetEntry{.{ .name = "OnlySheet" }};
    // local_sheet_id=1 references a sheet that doesn't exist (only sheet 0).
    try plan.addDefinedName(a, "Bad", "Sheet1!$A$1", .{ .local_sheet_id = 1 });

    try std.testing.expectError(
        error.InvalidDefinedNameLocalSheetId,
        emitWorkbookXml(a, &sheets, &plan),
    );
}
