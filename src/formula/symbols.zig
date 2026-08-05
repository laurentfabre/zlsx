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
pub const calc = @import("calc.zig");
/// M5b2: §5.7.6's calc-state writes and §5.7.7's mark-only eligibility.
/// Reached from `pkg/` by the prepare/swap transaction, which is the only
/// caller that decides what a run wants `<calcPr>` to say.
pub const calc_patch = @import("calc_patch.zig");
/// M5b1: the `ResolvedSheet` projection and its byte-confined patcher.
pub const resolved = @import("resolved.zig");
/// M3b: what a run is given and what it may spend. The transaction echoes
/// the resolved projection into its report (§5.7.8).
pub const run_inputs = @import("run_inputs.zig");
/// §5.4a's civil-date math. M6: the CLI's `--now` parser and its
/// resolved-input echo need the same days↔civil conversion the engine
/// dates with — a second copy would be a second calendar.
pub const serial_date = @import("serial_date.zig");
pub const names = @import("names.zig");
pub const env = @import("env.zig");
pub const value = @import("value.zig");
pub const metadata = @import("metadata.zig");
/// M4b3: the adapter's `Workbook.evaluate` parses a formula and runs the
/// evaluator, so both have to be reachable through the one module the
/// package layer imports — for the same reason everything above is.
pub const parser = @import("parser.zig");
pub const eval = @import("eval.zig");
/// Re-exported for `src/xlsx.zig`, which used to reach it by relative
/// path. Same rule, one layer up: `rewriter.zig` and this module both
/// contain `tokenizer.zig`, so they have to be the same module.
pub const rewriter = @import("rewriter.zig");
/// M5a1: the dependency graph. Reached through this module for the same
/// reason the evaluator is — `pkg/workbook.zig` imports one module, and
/// `graph.zig` and `eval.zig` share `env.zig`'s types, so they cannot be
/// two.
pub const graph = @import("graph.zig");
/// M5a2: the iteration engine and §5.6d's draw schedule, for the same
/// one-module reason as everything above — the engine's `Host` carries
/// `env.CellRef` and `value.ScalarValue`, and a second copy of either
/// would be a different type.
pub const iterate = @import("iterate.zig");
pub const draws = @import("draws.zig");
/// M5d2: `rng_v1`, as the recalc's draw source. Standalone evaluation
/// hands the evaluator a constant — a cache-based read must answer the
/// same call the same way twice — but a recalc is given a `RunInputs`
/// with a seed in it, and "equal inputs ⇒ equal output" is a promise
/// about *that* seed. It is reached through this module for the reason
/// everything above is: `Rng.drawSource()` returns an `eval.DrawSource`,
/// and a second `eval.zig` would make it a different type.
pub const rng = @import("rng.zig");

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
    /// What the `CT_DefinedName` attribute inventory decided about
    /// *referencing* this name (M4b3): `function`/`vbProcedure`/`xlm`
    /// put `macro_defined_name` here. Null for an ordinary name.
    ///
    /// Carried rather than folded into `class` because the two answer
    /// different questions — `class` is about the spelling, this is
    /// about the attributes — and a name can be both.
    attr_refusal: ?Refusal.Reason = null,
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
    /// `CT_Table@headerRowCount` / `@totalsRowCount`. Carried from M5a1
    /// on, because the row bands a structured reference denotes are
    /// arithmetic over these two numbers and the graph has to draw the
    /// edge whether or not the evaluator can yet compute the value.
    header_rows: u32 = 1,
    totals_rows: u32 = 0,
    columns: []const Column,

    /// Column lookup is by folded name, like every other symbol match.
    pub fn column(self: Table, folded_query: []const u8) ?*const Column {
        for (self.columns) |*c| {
            if (std.mem.eql(u8, c.folded, folded_query)) return c;
        }
        return null;
    }

    /// A column's **offset into the table**, which is what a structured
    /// reference's geometry is expressed in (M5a1). Same comparator as
    /// `column`; a separate function because a pointer cannot be
    /// subtracted from a slice the caller does not hold.
    pub fn columnIndex(self: Table, folded_query: []const u8) ?u32 {
        for (self.columns, 0..) |c, i| {
            if (std.mem.eql(u8, c.folded, folded_query)) return @intCast(i);
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

    /// §5.9's order — sheet-scoped name → workbook name → table →
    /// `_xlnm.` builtin → `#NAME?`.
    ///
    /// The order is not written here. `names.resolveInOrder` walks
    /// M2's exported `parser.value_resolution_order`, one tier per
    /// entry, and this supplies the tiers; the array is the single
    /// statement of the sequence and this file reads it. A `from` of
    /// null is a workbook-level context (a name's own body, a
    /// standalone evaluation) and the sheet-scoped tier declines rather
    /// than guessing a sheet.
    pub fn resolveName(
        self: *const SymbolTable,
        gpa: std.mem.Allocator,
        from: ?SheetIndex,
        query: []const u8,
    ) error{OutOfMemory}!Resolution {
        return self.resolveNameTraced(gpa, from, query, null);
    }

    /// `resolveName`, recording the tiers it visited. The trace exists
    /// for the tests that prove the walk followed the exported order
    /// and stopped where it says it stops.
    pub fn resolveNameTraced(
        self: *const SymbolTable,
        gpa: std.mem.Allocator,
        from: ?SheetIndex,
        query: []const u8,
        trace: ?*names.ValueTrace,
    ) error{OutOfMemory}!Resolution {
        const folded = self.fold(gpa, query) catch |err| {
            if (err == error.OutOfMemory) return error.OutOfMemory;
            return .{ .refused = .{ .reason = .invalid_utf8 } };
        };
        defer gpa.free(folded);

        const lookup: Tiers = .{ .table = self, .raw = query };
        const hit = names.resolveInOrder(
            Tiers,
            lookup,
            &parser.value_resolution_order,
            if (from) |s| s.toInt() else null,
            folded,
            trace,
        ) orelse return .not_found;
        return switch (hit) {
            .name => |n| self.answer(gpa, n),
            .table => |t| .{ .table = t },
            .refused => |r| .{ .refused = .{ .reason = r } },
        };
    }

    /// One tier of §5.9's order, answered on demand.
    ///
    /// `names.resolveInOrder` calls `at` once per entry in the order it
    /// was handed, so every branch here is a *tier*, never a sequence.
    /// The sequencing lives in the array.
    const Tiers = struct {
        table: *const SymbolTable,
        /// The unfolded spelling, for the `_xlnm.` tier — a reserved
        /// prefix is Excel's own ASCII and is classified on the source
        /// spelling, not on a fold of it.
        raw: []const u8,

        pub const Hit = union(enum) {
            name: *const Name,
            table: *const Table,
            refused: Refusal.Reason,
        };

        pub fn at(
            self: Tiers,
            stage: parser.ValueScope,
            from: ?u32,
            folded: []const u8,
        ) ?Hit {
            switch (stage) {
                .sheet_scoped_name => {
                    const sheet = from orelse return null;
                    for (self.table.names) |*n| {
                        const s = n.scope orelse continue;
                        if (s.toInt() == sheet and std.mem.eql(u8, n.folded, folded)) {
                            return .{ .name = n };
                        }
                    }
                    return null;
                },
                .workbook_name => {
                    for (self.table.names) |*n| {
                        if (n.scope == null and std.mem.eql(u8, n.folded, folded)) {
                            return .{ .name = n };
                        }
                    }
                    return null;
                },
                .table => {
                    for (self.table.tables) |*t| {
                        if (std.mem.eql(u8, t.folded, folded)) return .{ .table = t };
                    }
                    return null;
                },
                .builtin_xlnm => {
                    // Reached only when no `<definedName>` declares the
                    // spelling. A recognized builtin that the workbook
                    // never defined names nothing — `#NAME?`, so the
                    // tier declines and the terminal stage answers. An
                    // *unrecognized* `_xlnm.` spelling is a different
                    // thing: the prefix is reserved, so treating it as
                    // a user name that happens to start with it would
                    // be a guess.
                    return switch (classify(self.raw)) {
                        .unknown_builtin => .{ .refused = .unknown_builtin_name },
                        .lambda_parameter => .{ .refused = .lambda_parameter_name },
                        .user, .builtin => null,
                    };
                },
                // The terminal stage is the driver's; no tier supplies
                // "provably nowhere".
                .name_error => return null,
            }
        }
    };

    fn answer(
        self: *const SymbolTable,
        gpa: std.mem.Allocator,
        n: *const Name,
    ) error{OutOfMemory}!Resolution {
        _ = self;
        // Three ways a name that *resolved* still refuses, in the order
        // their evidence was gathered: the spelling (M4b1), the
        // attributes (M4b3's inventory), then the body — which costs a
        // parse and is therefore asked last.
        if (n.class.refusesWhenReferenced()) {
            return .{ .refused = .{ .reason = switch (n.class) {
                .lambda_parameter => .lambda_parameter_name,
                .unknown_builtin => .unknown_builtin_name,
                .user, .builtin => unreachable,
            } } };
        }
        if (n.attr_refusal) |r| return .{ .refused = .{ .reason = r } };
        if (try names.bodyRefusal(gpa, n.body)) |r| return .{ .refused = .{ .reason = r } };
        return .{ .name = n };
    }

    /// Expand a 3D sheet span over the workbook's own order (§5.6g).
    ///
    /// Inclusive at both ends, and `#REF!` when an endpoint names no
    /// sheet or the two arrive in the wrong order. The sheet lookup is
    /// the same case-folded one a single-sheet qualifier gets, because
    /// `sheet1:SHEET3` is the same span as `Sheet1:Sheet3`.
    pub fn resolveSheetSpan(
        self: *const SymbolTable,
        gpa: std.mem.Allocator,
        first: []const u8,
        last: []const u8,
    ) error{OutOfMemory}!names.SpanExpansion {
        const a = try self.resolveSheet(gpa, first);
        const b = try self.resolveSheet(gpa, last);
        return names.expandSpan(
            if (a) |x| x.toInt() else null,
            if (b) |y| y.toInt() else null,
        );
    }

    /// The one comparator. Public from M5a1 on, because the graph's
    /// structured-reference resolver has to match a column spelling the
    /// same way the evaluator matches a name — and a second, quietly
    /// different fold is exactly what injecting the collation was meant
    /// to prevent.
    pub fn fold(
        self: *const SymbolTable,
        gpa: std.mem.Allocator,
        s: []const u8,
    ) anyerror![]u8 {
        return self.collation.fold(gpa, s);
    }
};

/// The symbol table, as the evaluator's `env.NameResolver` seam.
///
/// Holds a scratch allocator because folding a query needs one and the
/// seam has no allocator parameter — the evaluator asks a question, it
/// does not fund one. Holds `last_refusal` for the same reason
/// `metadata.CellDialectResolver` does: the interface can only say
/// `error.NameRefused`, and the typed reason has to survive the trip
/// (M4a decision 17).
pub const NameResolution = struct {
    table: *const SymbolTable,
    gpa: std.mem.Allocator,
    last_refusal: ?Refusal = null,

    pub fn resolver(self: *NameResolution) env.NameResolver {
        return .{ .ctx = self, .resolve = resolve };
    }

    fn resolve(
        ctx: *anyopaque,
        from: ?SheetIndex,
        spelling: []const u8,
    ) env.Error!env.NameBinding {
        const self: *NameResolution = @ptrCast(@alignCast(ctx));
        const r = try self.table.resolveName(self.gpa, from, spelling);
        return switch (r) {
            .name => |n| .{ .body = .{ .text = n.body, .scope = n.scope } },
            .table => .table,
            .not_found => .not_found,
            .refused => |ref| {
                self.last_refusal = ref;
                return error.NameRefused;
            },
        };
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

    /// Everything about a defined name that is not its two carriers.
    /// A struct rather than positional parameters because M4b3 added
    /// the third field and the next row will not be the last: three
    /// bools and an optional in a row is a call site nobody can read.
    pub const NameFacts = struct {
        scope: ?SheetIndex = null,
        hidden: bool = false,
        /// What `classifyDefinedName` decided about referencing this
        /// name, from the `CT_DefinedName` attribute inventory.
        attr_refusal: ?Refusal.Reason = null,
    };

    /// Add a defined name. `raw_identifier` is `CT_DefinedName@name`
    /// (a STRING carrier); `raw_body` is the element's text, which is a
    /// FORMULA carrier and therefore entity-decoded ONLY.
    pub fn addName(
        self: *Builder,
        raw_identifier: []const u8,
        raw_body: []const u8,
        facts: NameFacts,
    ) error{OutOfMemory}!void {
        if (self.refusal != null) return;
        const a = self.arena.allocator();
        const identifier = try self.decodeInto(a, .defined_name_identifier, raw_identifier);
        const body = try self.decodeInto(a, .defined_name_body, raw_body);
        const folded = try self.foldInto(a, identifier);
        for (self.names.items) |n| {
            if (std.mem.eql(u8, n.folded, folded) and scopeEql(n.scope, facts.scope)) {
                self.refusal = .{ .reason = .duplicate_symbol };
                return;
            }
        }
        try self.names.append(a, .{
            .identifier = identifier,
            .folded = folded,
            .body = body,
            .scope = facts.scope,
            .hidden = facts.hidden,
            .class = classify(identifier),
            .attr_refusal = facts.attr_refusal,
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
            .header_rows = table.header_rows,
            .totals_rows = table.totals_rows,
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
        const defined = try self.names.toOwnedSlice(a);
        const tables = try self.tables.toOwnedSlice(a);
        self.consumed = true;
        return .{ .ok = .{
            .arena = self.arena,
            .collation = self.collation,
            .sheets = sheets,
            .names = defined,
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
    defined: []const struct {
        id: []const u8,
        body: []const u8 = "",
        scope: ?u32 = null,
        hidden: bool = false,
    },
) !SymbolTable {
    var b = Builder.init(testing.allocator, collation_v1);
    defer b.deinit();
    for (sheets) |s| try b.addSheet(s);
    for (defined) |n| {
        try b.addName(n.id, n.body, .{
            .scope = if (n.scope) |s| sheetIdx(s) else null,
            .hidden = n.hidden,
        });
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
    try b2.addName("Sales", "S!$Z$1", .{});
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
    try b.addName("Bad&nbsp;Name", "1", .{});
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
            try b.addName("Data", "'R&amp;D'!$A$1", .{});
            try b.addName("Data", "'R&amp;D'!$B$2", .{ .scope = SheetIndex.fromInt(1) });
            try b.addName("_xlpm.p", "1", .{});

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

// ─── §5.9 resolution over a real table (M4b3) ────────────────────

test "symbols: the tiers are walked in M2's exported order, and stop where they answer" {
    var t = try buildTable(
        &.{ "S0", "S1" },
        &.{
            .{ .id = "Data", .body = "S0!$A$1" },
            .{ .id = "Local", .body = "S1!$B$2", .scope = 1 },
        },
    );
    defer t.deinit();

    // A sheet-scoped hit stops at the first tier.
    var trace: names.ValueTrace = .{};
    const local = try t.resolveNameTraced(testing.allocator, sheetIdx(1), "Local", &trace);
    try testing.expectEqualStrings("S1!$B$2", local.name.body);
    try testing.expectEqualSlices(
        parser.ValueScope,
        &.{.sheet_scoped_name},
        trace.slice(),
    );

    // A workbook name is reached only after the sheet tier declines.
    var t2: names.ValueTrace = .{};
    _ = try t.resolveNameTraced(testing.allocator, sheetIdx(1), "Data", &t2);
    try testing.expectEqualSlices(
        parser.ValueScope,
        &.{ .sheet_scoped_name, .workbook_name },
        t2.slice(),
    );

    // Nowhere: every tier tried, in the exported order, ending at the
    // terminal stage that *is* `#NAME?`.
    var t3: names.ValueTrace = .{};
    const nowhere = try t.resolveNameTraced(testing.allocator, sheetIdx(0), "Nope", &t3);
    try testing.expect(nowhere == .not_found);
    try testing.expectEqualSlices(
        parser.ValueScope,
        &parser.value_resolution_order,
        t3.slice(),
    );
}

test "symbols: shadowing runs both ways, one fixture per direction" {
    var t = try buildTable(
        &.{ "S0", "S1" },
        &.{
            .{ .id = "Rate", .body = "S0!$A$1" },
            .{ .id = "Rate", .body = "S1!$B$2", .scope = 1 },
        },
    );
    defer t.deinit();

    // From the scoping sheet the local name wins…
    const from_local = try t.resolveName(testing.allocator, sheetIdx(1), "Rate");
    try testing.expectEqualStrings("S1!$B$2", from_local.name.body);
    try testing.expectEqual(@as(?SheetIndex, sheetIdx(1)), from_local.name.scope);
    // …and from any other sheet the workbook name does. The same
    // spelling, two answers, decided by where it was asked from.
    const from_other = try t.resolveName(testing.allocator, sheetIdx(0), "Rate");
    try testing.expectEqualStrings("S0!$A$1", from_other.name.body);
    try testing.expectEqual(@as(?SheetIndex, null), from_other.name.scope);
}

test "symbols: a macro name refuses when referenced, and carrying one does not" {
    var b = Builder.init(testing.allocator, collation_v1);
    defer b.deinit();
    try b.addSheet("S");
    try b.addName("Plain", "S!$A$1", .{});
    try b.addName("Macro", "MACRO_BODY", .{ .attr_refusal = .macro_defined_name });
    // Building succeeded: a workbook that merely *carries* a macro name
    // is a workbook Excel opens, and so is this one.
    var t = switch (try b.finish()) {
        .ok => |ok| ok,
        .refused => return error.TestUnexpectedRefusal,
    };
    defer t.deinit();

    try testing.expect((try t.resolveName(testing.allocator, null, "Plain")) == .name);
    const refused = try t.resolveName(testing.allocator, null, "Macro");
    try testing.expectEqual(Refusal.Reason.macro_defined_name, refused.refused.reason);
    try testing.expectEqual(
        decode.PlaneTwo.FormulaUnsupportedConstruct,
        refused.refused.planeTwo(),
    );
}

test "symbols: a relative body refuses when referenced; an anchored one resolves" {
    var t = try buildTable(&.{"S"}, &.{
        .{ .id = "Anchored", .body = "S!$A$1" },
        .{ .id = "Floating", .body = "S!A1" },
    });
    defer t.deinit();

    try testing.expect((try t.resolveName(testing.allocator, null, "Anchored")) == .name);
    const floating = try t.resolveName(testing.allocator, null, "Floating");
    try testing.expectEqual(Refusal.Reason.relative_reference_name, floating.refused.reason);
}

test "symbols: the `_xlnm.` tier answers only for a spelling nothing declares" {
    var t = try buildTable(&.{"S"}, &.{.{ .id = "_xlnm.Print_Area", .body = "S!$A$1:$C$9" }});
    defer t.deinit();

    // Declared: found at the name tier, before the builtin tier runs.
    var trace: names.ValueTrace = .{};
    const declared = try t.resolveNameTraced(
        testing.allocator,
        null,
        "_xlnm.Print_Area",
        &trace,
    );
    try testing.expect(declared == .name);
    try testing.expect(trace.slice()[trace.len - 1] != .builtin_xlnm);

    // Recognized but never declared: nothing names it, so `#NAME?`.
    try testing.expect(
        (try t.resolveName(testing.allocator, null, "_xlnm.Print_Titles")) == .not_found,
    );
    // Unrecognized under a reserved prefix: a guess this refuses to make.
    const unknown = try t.resolveName(testing.allocator, null, "_xlnm.Invented");
    try testing.expectEqual(Refusal.Reason.unknown_builtin_name, unknown.refused.reason);
    const lambda = try t.resolveName(testing.allocator, null, "_xlpm.p");
    try testing.expectEqual(Refusal.Reason.lambda_parameter_name, lambda.refused.reason);
}

test "symbols: a sheet span expands case-folded, inclusively, in workbook order" {
    var t = try buildTable(&.{ "Jan", "Feb", "Mar", "Données" }, &.{});
    defer t.deinit();

    const all = try t.resolveSheetSpan(testing.allocator, "Jan", "Mar");
    try testing.expectEqual(@as(u32, 0), all.members.first);
    try testing.expectEqual(@as(u32, 2), all.members.last);
    // The same fold every other symbol match uses.
    const folded = try t.resolveSheetSpan(testing.allocator, "jan", "DONNÉES");
    try testing.expectEqual(@as(u32, 3), folded.members.last);
    // One sheet is a span.
    try testing.expectEqual(
        @as(u32, 1),
        (try t.resolveSheetSpan(testing.allocator, "Feb", "Feb")).members.first,
    );
    // A deleted endpoint, and a pair that has swapped: both `#REF!`.
    try testing.expect((try t.resolveSheetSpan(testing.allocator, "Jan", "Gone")) == .ref_error);
    try testing.expect((try t.resolveSheetSpan(testing.allocator, "Mar", "Jan")) == .ref_error);
}

test "symbols: the resolver seam answers the evaluator's three questions" {
    var t = try buildTable(&.{"S"}, &.{
        .{ .id = "Good", .body = "S!$A$1" },
        .{ .id = "Macro", .body = "X" },
    });
    defer t.deinit();
    // Reach in and mark the second one, the way the inventory would.
    const marked: *Name = @constCast(&t.names[1]);
    marked.attr_refusal = .macro_defined_name;

    var seam: NameResolution = .{ .table = &t, .gpa = testing.allocator };
    const r = seam.resolver();

    const found = try r.resolveName(sheetIdx(0), "good");
    try testing.expectEqualStrings("S!$A$1", found.body.text);
    try testing.expect((try r.resolveName(sheetIdx(0), "Missing")) == .not_found);

    // A refusal arrives as one error and leaves its reason behind.
    try testing.expectError(error.NameRefused, r.resolveName(sheetIdx(0), "Macro"));
    try testing.expectEqual(Refusal.Reason.macro_defined_name, seam.last_refusal.?.reason);
}

test "checkAllAllocationFailures: M4b3 resolution leaks nothing under OOM" {
    const H = struct {
        fn run(allocator: std.mem.Allocator) !void {
            var b = Builder.init(allocator, collation_v1);
            defer b.deinit();
            try b.addSheet("S0");
            try b.addSheet("S1");
            try b.addName("Anchored", "S0!$A$1", .{});
            try b.addName("Floating", "S0!A1", .{});
            try b.addName("Macro", "M", .{ .attr_refusal = .macro_defined_name });
            try b.addName("Local", "S1!$B$2", .{ .scope = SheetIndex.fromInt(1) });
            var t = switch (try b.finish()) {
                .ok => |ok| ok,
                .refused => return error.TestUnexpectedRefusal,
            };
            defer t.deinit();

            if ((try t.resolveName(allocator, SheetIndex.fromInt(1), "local")) != .name) {
                return error.WrongName;
            }
            if ((try t.resolveName(allocator, null, "floating")) != .refused) return error.WrongBody;
            if ((try t.resolveName(allocator, null, "macro")) != .refused) return error.WrongAttr;
            if ((try t.resolveSheetSpan(allocator, "s0", "S1")) != .members) return error.WrongSpan;

            var seam: NameResolution = .{ .table = &t, .gpa = allocator };
            const r = seam.resolver();
            _ = r.resolveName(SheetIndex.fromInt(0), "anchored") catch |e| {
                if (e == error.OutOfMemory) return e;
                return error.WrongSeam;
            };
        }
    };
    try testing.checkAllAllocationFailures(testing.allocator, H.run, .{});
}
