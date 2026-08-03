//! The decoded symbol layer — what a name in a formula can refer to,
//! and how a spelling finds it (`goal_formula.md` §5.9).
//!
//! M4b1 of the tier-D1 ladder, and the module `pkg/` imports.
//!
//! Two jobs, and the first is why the second is possible
//! -----------------------------------------------------
//! **Decoded.** Every identifier a workbook carries arrives XML-encoded
//! and, for the ST_Xstring-typed ones, escape-encoded on top of that. A
//! sheet named `R&D` is `R&amp;D` in `xl/workbook.xml`, and a formula
//! referring to it says `'R&D'!A1`. Matching the raw attribute against
//! the parsed formula compares `R&amp;D` with `R&D` and finds nothing —
//! `#NAME?` on a workbook that opens perfectly in Excel. So the table is
//! built from *decoded* bytes, once, and every lookup after that is a
//! comparison between two things in the same alphabet.
//!
//! **Ordered.** §5.9's resolution order is sheet-scoped name (which
//! shadows) → workbook name → table → `_xlnm.` builtin → `#NAME?`, and
//! it is an order rather than a set because the same spelling can name
//! two things. A workbook-scoped `Data` and a `Data` scoped to Sheet2
//! are different names, and which one a formula on Sheet2 means is not
//! negotiable.
//!
//! Matching is case-folded (§5.4b's one comparator: full non-Turkic
//! folding), because Excel's names are case-insensitive — `SUM(data)`
//! and `SUM(Data)` are the same reference. The fold is *injected* the
//! way `value.zig` injects it, so this file stays independent of the
//! concrete table.
//!
//! What is deliberately NOT here
//! -----------------------------
//! Name *bodies* are decoded and classified, never parsed: a name whose
//! body is a formula becomes a graph node at M5a (with the interim
//! depth guard at M4b3), and relative-reference names refuse there. What
//! this layer owes those rows is the decoded body and an honest
//! classification of the identifier — including the two prefixes that
//! refuse on sight.
//!
//! Function names resolve through `registry.zig`, not here: §5.9's call
//! position strips layered prefixes and asks the registry, and a
//! defined name never participates in that lookup.

const std = @import("std");
const assert = std.debug.assert;

const coords = @import("zlsx_refs");

/// The engine surface the package layer reaches through this module.
///
/// `pkg/workbook.zig` imports `zlsx_formula` and nothing else from the
/// engine. That is not a convenience: a file compiled into two modules
/// is two distinct types, so an adapter that named `env` and `symbols`
/// as separate modules would build an `EvalEnv` the evaluator could not
/// accept — and the error would arrive as a type mismatch nobody can
/// read. One module, one set of types.
pub const decode = @import("decode.zig");
pub const env = @import("env.zig");
pub const value = @import("value.zig");
pub const metadata = @import("metadata.zig");
/// Re-exported for `src/xlsx.zig`, which used to reach it by relative
/// path. Same rule, one layer up: `rewriter.zig` and this module both
/// contain `tokenizer.zig`, so they have to be the same module.
pub const rewriter = @import("rewriter.zig");

pub const Refusal = decode.Refusal;
pub const SheetIndex = env.SheetIndex;

// ─── name classification ─────────────────────────────────────────

/// Excel reserves two prefixes on defined names, and they mean opposite
/// things: `_xlnm.` is a *builtin* the application owns, `_xlpm.` is a
/// LAMBDA parameter that only means anything inside its own body.
pub const NameClass = enum {
    /// An ordinary user-defined name.
    user,
    /// A recognized `_xlnm.` builtin.
    builtin,
    /// An `_xlnm.` spelling nothing has classified. Refused when
    /// referenced rather than treated as a user name that happens to
    /// start with a reserved prefix.
    unknown_builtin,
    /// `_xlpm.` — a LAMBDA parameter. §5.9 refuses LAMBDA/LET through
    /// v1, and a parameter binding reached from outside its body is not
    /// a value this engine can produce.
    lambda_parameter,

    /// Whether merely *referring* to a name of this class refuses.
    pub fn refusesWhenReferenced(self: NameClass) bool {
        return switch (self) {
            .user, .builtin => false,
            .unknown_builtin, .lambda_parameter => true,
        };
    }
};

pub const builtin_prefix = "_xlnm.";
pub const lambda_param_prefix = "_xlpm.";

/// The `_xlnm.` builtins ECMA-376 defines (`ST_BuiltinName`-adjacent;
/// the spellings Office writes). Recognized so an unfamiliar one
/// refuses as *unknown* rather than resolving as a user name.
pub const builtin_names = [_][]const u8{
    "_xlnm.Print_Area",
    "_xlnm.Print_Titles",
    "_xlnm.Criteria",
    "_xlnm._FilterDatabase",
    "_xlnm.Extract",
    "_xlnm.Consolidate_Area",
    "_xlnm.Database",
    "_xlnm.Sheet_Title",
};

/// Classify a **decoded** identifier. Prefix matching is ASCII
/// case-insensitive: the prefix is Excel's, not the user's, and Excel
/// accepts its own spelling in either case.
pub fn classify(identifier: []const u8) NameClass {
    if (startsWithIgnoreAsciiCase(identifier, lambda_param_prefix)) return .lambda_parameter;
    if (!startsWithIgnoreAsciiCase(identifier, builtin_prefix)) return .user;
    for (builtin_names) |b| {
        if (std.ascii.eqlIgnoreCase(b, identifier)) return .builtin;
    }
    return .unknown_builtin;
}

fn startsWithIgnoreAsciiCase(s: []const u8, prefix: []const u8) bool {
    if (s.len < prefix.len) return false;
    return std.ascii.eqlIgnoreCase(s[0..prefix.len], prefix);
}

// ─── the table ───────────────────────────────────────────────────

pub const Sheet = struct {
    index: SheetIndex,
    /// Decoded (STRING carrier — `CT_Sheet@name` is ST_Xstring-typed).
    name: []const u8,
    folded: []const u8,
};

pub const Name = struct {
    /// Decoded identifier.
    identifier: []const u8,
    folded: []const u8,
    /// Decoded body (FORMULA carrier: XML entities only, so a literal
    /// `_x0041_` in a name's formula survives to the tokenizer).
    body: []const u8,
    /// Null for a workbook-scoped name; set for a sheet-scoped one,
    /// which shadows the workbook name of the same spelling on that
    /// sheet.
    scope: ?SheetIndex,
    hidden: bool,
    class: NameClass,
};

pub const Column = struct {
    /// Decoded (STRING carrier).
    name: []const u8,
    folded: []const u8,
    /// Decoded (FORMULA carrier), when the column is calculated.
    calculated_formula: ?[]const u8 = null,
    totals_formula: ?[]const u8 = null,
};

pub const Table = struct {
    name: []const u8,
    folded: []const u8,
    sheet: SheetIndex,
    /// Raw A1 range as written (`ref`), resolved by the consumer.
    ref: []const u8,
    columns: []const Column,

    /// Column lookup is by folded name, like every other symbol match.
    pub fn column(self: Table, folded_query: []const u8) ?*const Column {
        for (self.columns) |*c| {
            if (std.mem.eql(u8, c.folded, folded_query)) return c;
        }
        return null;
    }
};

/// What a spelling in value position resolves to (§5.9). `not_found`
/// is `#NAME?`'s evidence, and it is a *value* outcome, not a refusal —
/// the refusals here are the two reserved prefixes.
pub const Resolution = union(enum) {
    name: *const Name,
    table: *const Table,
    not_found,
    refused: Refusal,
};

pub const SymbolTable = struct {
    arena: std.heap.ArenaAllocator,
    collation: value.Collation,
    sheets: []const Sheet,
    names: []const Name,
    tables: []const Table,

    pub fn deinit(self: *SymbolTable) void {
        self.arena.deinit();
        self.* = undefined;
    }

    /// Sheet lookup by name, case-folded. Excel's sheet names are
    /// case-insensitive, so `'data'!A1` and `Data!A1` are one sheet.
    pub fn resolveSheet(
        self: *const SymbolTable,
        gpa: std.mem.Allocator,
        name: []const u8,
    ) error{OutOfMemory}!?SheetIndex {
        // Running out of memory is not a statement about the query, so
        // it is re-raised rather than answered with "no such sheet".
        const folded = self.fold(gpa, name) catch |err| {
            if (err == error.OutOfMemory) return error.OutOfMemory;
            // Anything else a fold can fail with is malformed UTF-8,
            // which no sheet name is.
            return null;
        };
        defer gpa.free(folded);
        for (self.sheets) |s| {
            if (std.mem.eql(u8, s.folded, folded)) return s.index;
        }
        return null;
    }

    pub fn sheetName(self: *const SymbolTable, idx: SheetIndex) ?[]const u8 {
        for (self.sheets) |s| {
            if (s.index == idx) return s.name;
        }
        return null;
    }

    /// §5.9's order, in one pass per tier: sheet-scoped name → workbook
    /// name → table. A `from` of null is a workbook-level context (a
    /// name's own body, a standalone evaluation) and skips the first
    /// tier rather than guessing a sheet.
    pub fn resolveName(
        self: *const SymbolTable,
        gpa: std.mem.Allocator,
        from: ?SheetIndex,
        query: []const u8,
    ) error{OutOfMemory}!Resolution {
        const folded = self.fold(gpa, query) catch |err| {
            if (err == error.OutOfMemory) return error.OutOfMemory;
            return .{ .refused = .{ .reason = .invalid_utf8 } };
        };
        defer gpa.free(folded);

        if (from) |sheet| {
            for (self.names) |*n| {
                if (n.scope) |s| {
                    if (s == sheet and std.mem.eql(u8, n.folded, folded)) return self.answer(n);
                }
            }
        }
        for (self.names) |*n| {
            if (n.scope == null and std.mem.eql(u8, n.folded, folded)) return self.answer(n);
        }
        for (self.tables) |*t| {
            if (std.mem.eql(u8, t.folded, folded)) return .{ .table = t };
        }
        return .not_found;
    }

    fn answer(self: *const SymbolTable, n: *const Name) Resolution {
        _ = self;
        if (n.class.refusesWhenReferenced()) {
            return .{ .refused = .{ .reason = switch (n.class) {
                .lambda_parameter => .lambda_parameter_name,
                .unknown_builtin => .unknown_builtin_name,
                .user, .builtin => unreachable,
            } } };
        }
        return .{ .name = n };
    }

    fn fold(
        self: *const SymbolTable,
        gpa: std.mem.Allocator,
        s: []const u8,
    ) anyerror![]u8 {
        return self.collation.fold(gpa, s);
    }
};

// ─── building it ─────────────────────────────────────────────────

pub const Built = union(enum) {
    ok: SymbolTable,
    refused: Refusal,
};

/// Accumulates decoded symbols, then hands back a table or the first
/// refusal.
///
/// Two-phase for the same reason `metadata.resolveRun` is: a caller
/// that got a half-built table back would have to know which half, and
/// "the workbook's symbols are partly readable" is not a state any
/// caller should have to reason about.
pub const Builder = struct {
    arena: std.heap.ArenaAllocator,
    collation: value.Collation,
    sheets: std.ArrayListUnmanaged(Sheet) = .empty,
    names: std.ArrayListUnmanaged(Name) = .empty,
    tables: std.ArrayListUnmanaged(Table) = .empty,
    refusal: ?Refusal = null,
    /// Set once the arena has left — by `finish` handing it to a
    /// `SymbolTable`, or by `deinit` freeing it.
    consumed: bool = false,

    pub fn init(gpa: std.mem.Allocator, collation: value.Collation) Builder {
        return .{ .arena = std.heap.ArenaAllocator.init(gpa), .collation = collation };
    }

    /// Free everything the build still owns.
    ///
    /// Idempotent, and a no-op once `finish` has succeeded, so the
    /// obvious `defer b.deinit()` is correct on every path — including
    /// the one where an injected allocation failure lands between the
    /// last `add` and the `finish`. An ownership rule a caller has to
    /// remember is one a caller gets wrong under OOM.
    pub fn deinit(self: *Builder) void {
        if (self.consumed) return;
        self.consumed = true;
        self.arena.deinit();
    }

    /// Add a sheet, in workbook order. `raw_name` is the undecoded
    /// `CT_Sheet@name` attribute.
    pub fn addSheet(self: *Builder, raw_name: []const u8) error{OutOfMemory}!void {
        if (self.refusal != null) return;
        const a = self.arena.allocator();
        const name = self.decodeInto(a, .sheet_name, raw_name) catch |err| return err;
        const folded = self.foldInto(a, name) catch |err| return err;
        const idx = SheetIndex.fromInt(@intCast(self.sheets.items.len));
        // A duplicate sheet name is not a thing Excel can produce and
        // not a thing a reference could resolve unambiguously.
        for (self.sheets.items) |s| {
            if (std.mem.eql(u8, s.folded, folded)) {
                self.refusal = .{ .reason = .duplicate_symbol };
                return;
            }
        }
        try self.sheets.append(a, .{ .index = idx, .name = name, .folded = folded });
    }

    /// Add a defined name. `raw_identifier` is `CT_DefinedName@name`
    /// (a STRING carrier); `raw_body` is the element's text, which is a
    /// FORMULA carrier and therefore entity-decoded ONLY.
    pub fn addName(
        self: *Builder,
        raw_identifier: []const u8,
        raw_body: []const u8,
        scope: ?SheetIndex,
        hidden: bool,
    ) error{OutOfMemory}!void {
        if (self.refusal != null) return;
        const a = self.arena.allocator();
        const identifier = try self.decodeInto(a, .defined_name_identifier, raw_identifier);
        const body = try self.decodeInto(a, .defined_name_body, raw_body);
        const folded = try self.foldInto(a, identifier);
        for (self.names.items) |n| {
            if (std.mem.eql(u8, n.folded, folded) and scopeEql(n.scope, scope)) {
                self.refusal = .{ .reason = .duplicate_symbol };
                return;
            }
        }
        try self.names.append(a, .{
            .identifier = identifier,
            .folded = folded,
            .body = body,
            .scope = scope,
            .hidden = hidden,
            .class = classify(identifier),
        });
    }

    /// Add a table and its columns. The names arrive already decoded
    /// from `decode.scanTable`, which is where the table part's own
    /// carriers are resolved; this only folds and files them.
    pub fn addTable(self: *Builder, sheet: SheetIndex, table: decode.Table) error{OutOfMemory}!void {
        if (self.refusal != null) return;
        const a = self.arena.allocator();
        const name = try a.dupe(u8, table.display_name);
        const folded = try self.foldInto(a, name);
        for (self.tables.items) |t| {
            if (std.mem.eql(u8, t.folded, folded)) {
                self.refusal = .{ .reason = .duplicate_symbol };
                return;
            }
        }
        const columns = try a.alloc(Column, table.columns.len);
        for (table.columns, columns) |src, *dst| {
            const col_name = try a.dupe(u8, src.name);
            dst.* = .{
                .name = col_name,
                .folded = try self.foldInto(a, col_name),
                .calculated_formula = if (src.calculated_formula) |f| try a.dupe(u8, f) else null,
                .totals_formula = if (src.totals_formula) |f| try a.dupe(u8, f) else null,
            };
        }
        try self.tables.append(a, .{
            .name = name,
            .folded = folded,
            .sheet = sheet,
            .ref = try a.dupe(u8, table.ref),
            .columns = columns,
        });
    }

    pub fn finish(self: *Builder) error{OutOfMemory}!Built {
        assert(!self.consumed);
        if (self.refusal) |r| {
            self.deinit();
            return .{ .refused = r };
        }
        const a = self.arena.allocator();
        // Every allocation happens before ownership moves: a failure
        // here leaves the builder holding its arena, which is exactly
        // what the caller's `deinit` expects to find.
        const sheets = try self.sheets.toOwnedSlice(a);
        const names = try self.names.toOwnedSlice(a);
        const tables = try self.tables.toOwnedSlice(a);
        self.consumed = true;
        return .{ .ok = .{
            .arena = self.arena,
            .collation = self.collation,
            .sheets = sheets,
            .names = names,
            .tables = tables,
        } };
    }

    fn decodeInto(
        self: *Builder,
        a: std.mem.Allocator,
        site: decode.Site,
        raw: []const u8,
    ) error{OutOfMemory}![]const u8 {
        return decode.decodeAt(a, site, raw) catch |err| switch (err) {
            error.OutOfMemory => error.OutOfMemory,
            error.BadEntity => {
                self.refusal = .{ .reason = .bad_entity };
                return "";
            },
            error.BadXstring => {
                self.refusal = .{ .reason = .bad_xstring_escape };
                return "";
            },
        };
    }

    fn foldInto(
        self: *Builder,
        a: std.mem.Allocator,
        s: []const u8,
    ) error{OutOfMemory}![]const u8 {
        return self.collation.fold(a, s) catch |err| {
            if (err == error.OutOfMemory) return error.OutOfMemory;
            // The only other failure a fold has is input that is not
            // UTF-8, which is a statement about the part.
            self.refusal = .{ .reason = .invalid_utf8 };
            return "";
        };
    }
};

fn scopeEql(a: ?SheetIndex, b: ?SheetIndex) bool {
    if (a == null and b == null) return true;
    if (a == null or b == null) return false;
    return a.? == b.?;
}

// ─── tests ───────────────────────────────────────────────────────
//
// The concrete fold is imported HERE and nowhere else, exactly as
// `value.zig` does it: a file-scope const referenced only from a test
// block is not resolved in a non-test build, so this module stays
// buildable without declaring `zlsx_casefold` while its fixtures still
// run against the shipped algorithm.

const casefold = @import("zlsx_casefold");
const testing = std.testing;

fn shippedFold(allocator: std.mem.Allocator, s: []const u8) anyerror![]u8 {
    return casefold.foldString(allocator, s);
}

const collation_v1: value.Collation = .{ .fold = &shippedFold };

fn sheetIdx(i: u32) SheetIndex {
    return SheetIndex.fromInt(i);
}

fn buildTable(
    sheets: []const []const u8,
    names: []const struct {
        id: []const u8,
        body: []const u8 = "",
        scope: ?u32 = null,
        hidden: bool = false,
    },
) !SymbolTable {
    var b = Builder.init(testing.allocator, collation_v1);
    defer b.deinit();
    for (sheets) |s| try b.addSheet(s);
    for (names) |n| {
        try b.addName(n.id, n.body, if (n.scope) |s| sheetIdx(s) else null, n.hidden);
    }
    return switch (try b.finish()) {
        .ok => |t| t,
        .refused => error.TestUnexpectedRefusal,
    };
}

test "symbols: an entity-encoded sheet name matches the formula's spelling" {
    // The bug this exists to prevent: `R&amp;D` never equals `R&D`, so
    // `'R&D'!A1` resolves to nothing on a workbook Excel opens fine.
    var t = try buildTable(&.{ "Sheet1", "R&amp;D" }, &.{});
    defer t.deinit();

    try testing.expectEqualStrings("R&D", t.sheets[1].name);
    try testing.expectEqual(
        @as(?SheetIndex, sheetIdx(1)),
        try t.resolveSheet(testing.allocator, "R&D"),
    );
    // …and the raw spelling is NOT a sheet.
    try testing.expectEqual(
        @as(?SheetIndex, null),
        try t.resolveSheet(testing.allocator, "R&amp;D"),
    );
}

test "symbols: sheet names are ST_Xstring, so an escape decodes" {
    var t = try buildTable(&.{"Tab_x0041_"}, &.{});
    defer t.deinit();
    try testing.expectEqualStrings("TabA", t.sheets[0].name);
    try testing.expectEqual(
        @as(?SheetIndex, sheetIdx(0)),
        try t.resolveSheet(testing.allocator, "TabA"),
    );
}

test "symbols: sheet and name matching is case-folded" {
    var t = try buildTable(&.{"Données"}, &.{.{ .id = "Chiffre", .body = "Sheet1!$A$1" }});
    defer t.deinit();

    try testing.expectEqual(
        @as(?SheetIndex, sheetIdx(0)),
        try t.resolveSheet(testing.allocator, "DONNÉES"),
    );
    const r = try t.resolveName(testing.allocator, null, "chiffre");
    try testing.expectEqualStrings("Chiffre", r.name.identifier);
}

test "symbols: a sheet-scoped name shadows the workbook name of the same spelling" {
    var t = try buildTable(
        &.{ "S0", "S1" },
        &.{
            .{ .id = "Data", .body = "S0!$A$1" },
            .{ .id = "Data", .body = "S1!$B$2", .scope = 1 },
        },
    );
    defer t.deinit();

    // On the scoping sheet, the local name wins…
    const local = try t.resolveName(testing.allocator, sheetIdx(1), "Data");
    try testing.expectEqualStrings("S1!$B$2", local.name.body);
    // …everywhere else, the workbook name stands.
    const global = try t.resolveName(testing.allocator, sheetIdx(0), "Data");
    try testing.expectEqualStrings("S0!$A$1", global.name.body);
    const none = try t.resolveName(testing.allocator, null, "Data");
    try testing.expectEqualStrings("S0!$A$1", none.name.body);
}

test "symbols: a name body is a FORMULA carrier — its _x0041_ survives" {
    // Same seven bytes in the identifier and in the body, two answers.
    var t = try buildTable(&.{"S"}, &.{.{
        .id = "N_x0041_",
        .body = "IF(S!A1&gt;0,&quot;_x0041_&quot;,&quot;b&quot;)",
    }});
    defer t.deinit();

    try testing.expectEqualStrings("NA", t.names[0].identifier);
    try testing.expectEqualStrings("IF(S!A1>0,\"_x0041_\",\"b\")", t.names[0].body);
}

test "symbols: an unresolved spelling is not_found, which is #NAME?'s evidence" {
    var t = try buildTable(&.{"S"}, &.{.{ .id = "Known" }});
    defer t.deinit();
    try testing.expect((try t.resolveName(testing.allocator, null, "Unknown")) == .not_found);
}

test "symbols: the two reserved prefixes classify, and refuse when referenced" {
    var t = try buildTable(&.{"S"}, &.{
        .{ .id = "_xlnm.Print_Area", .body = "S!$A$1:$C$9" },
        .{ .id = "_xlnm.Nope", .body = "S!$A$1" },
        .{ .id = "_xlpm.x", .body = "1" },
        .{ .id = "Plain", .body = "1" },
    });
    defer t.deinit();

    try testing.expectEqual(NameClass.builtin, t.names[0].class);
    try testing.expectEqual(NameClass.unknown_builtin, t.names[1].class);
    try testing.expectEqual(NameClass.lambda_parameter, t.names[2].class);
    try testing.expectEqual(NameClass.user, t.names[3].class);

    // A builtin resolves; the other two refuse rather than resolving to
    // something the engine would then have to guess about.
    try testing.expect((try t.resolveName(testing.allocator, null, "_xlnm.Print_Area")) == .name);
    const unknown = try t.resolveName(testing.allocator, null, "_xlnm.Nope");
    try testing.expectEqual(Refusal.Reason.unknown_builtin_name, unknown.refused.reason);
    try testing.expectEqual(
        decode.PlaneTwo.FormulaUnsupportedConstruct,
        unknown.refused.planeTwo(),
    );
    const lambda = try t.resolveName(testing.allocator, null, "_xlpm.x");
    try testing.expectEqual(Refusal.Reason.lambda_parameter_name, lambda.refused.reason);
}

test "symbols: a table resolves after names, and its columns fold too" {
    const table_xml = "<table xmlns=\"" ++ decode.ns_main ++ "\" name=\"Sales\" displayName=\"Sales\" ref=\"A1:B4\">" ++
        "<tableColumns><tableColumn name=\"R&amp;D\"/>" ++
        "<tableColumn name=\"Total\"><calculatedColumnFormula>SUM(Sales[R&amp;D])</calculatedColumnFormula></tableColumn>" ++
        "</tableColumns></table>";
    var scanned = switch (try decode.scanTable(testing.allocator, table_xml, .{})) {
        .ok => |ok| ok,
        .refused => return error.TestUnexpectedRefusal,
    };
    defer scanned.deinit();

    var b = Builder.init(testing.allocator, collation_v1);
    defer b.deinit();
    try b.addSheet("S");
    try b.addTable(sheetIdx(0), scanned);
    var t = switch (try b.finish()) {
        .ok => |ok| ok,
        .refused => return error.TestUnexpectedRefusal,
    };
    defer t.deinit();

    const r = try t.resolveName(testing.allocator, null, "sales");
    try testing.expectEqualStrings("Sales", r.table.name);
    try testing.expect(r.table.column("r&d") != null);
    try testing.expectEqualStrings(
        "SUM(Sales[R&D])",
        r.table.columns[1].calculated_formula.?,
    );

    // A name of the same spelling would win — §5.9's order is an order.
    var b2 = Builder.init(testing.allocator, collation_v1);
    defer b2.deinit();
    try b2.addSheet("S");
    try b2.addName("Sales", "S!$Z$1", null, false);
    try b2.addTable(sheetIdx(0), scanned);
    var t2 = switch (try b2.finish()) {
        .ok => |ok| ok,
        .refused => return error.TestUnexpectedRefusal,
    };
    defer t2.deinit();
    try testing.expect((try t2.resolveName(testing.allocator, null, "Sales")) == .name);
}

test "symbols: duplicates refuse rather than picking one" {
    var b = Builder.init(testing.allocator, collation_v1);
    defer b.deinit();
    try b.addSheet("Data");
    try b.addSheet("DATA");
    switch (try b.finish()) {
        .ok => |t| {
            var table = t;
            table.deinit();
            return error.TestExpectedRefusal;
        },
        .refused => |r| try testing.expectEqual(Refusal.Reason.duplicate_symbol, r.reason),
    }

    // Same spelling, different scopes, is NOT a duplicate.
    var t2 = try buildTable(&.{ "A", "B" }, &.{
        .{ .id = "N", .body = "1" },
        .{ .id = "N", .body = "2", .scope = 0 },
        .{ .id = "N", .body = "3", .scope = 1 },
    });
    defer t2.deinit();
    try testing.expectEqual(@as(usize, 3), t2.names.len);
}

test "symbols: a broken encoding refuses, and the table is not half-built" {
    var b = Builder.init(testing.allocator, collation_v1);
    defer b.deinit();
    try b.addSheet("Fine");
    try b.addName("Bad&nbsp;Name", "1", null, false);
    try b.addSheet("AlsoFine");
    switch (try b.finish()) {
        .ok => |t| {
            var table = t;
            table.deinit();
            return error.TestExpectedRefusal;
        },
        .refused => |r| {
            try testing.expectEqual(Refusal.Reason.bad_entity, r.reason);
            try testing.expectEqual(decode.PlaneTwo.FormulaMalformedInput, r.planeTwo());
        },
    }
}

test "symbols: hidden and scope survive for the rows that need them" {
    var t = try buildTable(&.{"S"}, &.{
        .{ .id = "Visible", .body = "1" },
        .{ .id = "Hidden", .body = "2", .hidden = true, .scope = 0 },
    });
    defer t.deinit();
    try testing.expect(!t.names[0].hidden);
    try testing.expect(t.names[1].hidden);
    try testing.expectEqual(@as(?SheetIndex, null), t.names[0].scope);
    try testing.expectEqual(@as(?SheetIndex, sheetIdx(0)), t.names[1].scope);
    try testing.expectEqualStrings("S", t.sheetName(sheetIdx(0)).?);
}

test "checkAllAllocationFailures: building and resolving leak nothing under OOM" {
    const H = struct {
        fn run(allocator: std.mem.Allocator) !void {
            var b = Builder.init(allocator, collation_v1);
            defer b.deinit();
            try b.addSheet("R&amp;D");
            try b.addSheet("Sheet_x0041_");
            try b.addName("Data", "'R&amp;D'!$A$1", null, false);
            try b.addName("Data", "'R&amp;D'!$B$2", SheetIndex.fromInt(1), false);
            try b.addName("_xlpm.p", "1", null, false);

            const table_xml = "<table xmlns=\"" ++ decode.ns_main ++ "\" name=\"T\" ref=\"A1:B2\">" ++
                "<tableColumns><tableColumn name=\"c\"/></tableColumns></table>";
            var scanned = switch (try decode.scanTable(allocator, table_xml, .{})) {
                .ok => |ok| ok,
                .refused => return error.TestUnexpectedRefusal,
            };
            defer scanned.deinit();
            try b.addTable(SheetIndex.fromInt(0), scanned);

            var t = switch (try b.finish()) {
                .ok => |ok| ok,
                .refused => return error.TestUnexpectedRefusal,
            };
            defer t.deinit();

            if ((try t.resolveSheet(allocator, "r&d")) == null) return error.WrongSheet;
            const local = try t.resolveName(allocator, SheetIndex.fromInt(1), "data");
            if (local != .name) return error.WrongName;
            if ((try t.resolveName(allocator, null, "_xlpm.p")) != .refused) return error.WrongClass;
            if ((try t.resolveName(allocator, null, "t")) != .table) return error.WrongTable;
        }
    };
    try testing.checkAllAllocationFailures(testing.allocator, H.run, .{});
}

test "coords stays reachable for the consumers of a table ref" {
    // `Table.ref` is handed on raw; this is the one assertion that the
    // range spelling a table carries is the one `zlsx_refs` parses, so
    // a consumer does not have to discover it at M7b.
    const r = try coords.parseRange("A1:B4", .{ .dollar = .accept });
    try testing.expectEqual(@as(u32, 1), r.first.row.oneBased());
    try testing.expectEqual(@as(u32, 1), r.last.col.zeroBased());
}
