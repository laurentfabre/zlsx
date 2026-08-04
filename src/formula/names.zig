//! What a *name* can carry, what a *table* can produce, and what a
//! reference that spans sheets is allowed to reach
//! (`goal_formula.md` §5.9, §5.6g, M4b3).
//!
//! M4b1 built the decoded symbol layer and M2 exported §5.9's two
//! resolution orders as checked arrays. This row is what consumes them,
//! plus the two inventories and the one matrix those orders need in
//! order to mean anything against a real workbook.
//!
//! The order is data, and the driver reads it
//! -------------------------------------------
//! §5.9's value-position order — sheet-scoped name, workbook name,
//! table, `_xlnm.` builtin, `#NAME?` — lives in `parser.zig` as
//! `value_resolution_order`, and `resolveInOrder` *iterates that array*.
//! Restating the order as a chain of `if`s here would leave two
//! statements of one rule, and the moment they disagreed the workbook
//! would resolve one way and the spec document another. The proof is a
//! test: handed a permuted order, the driver resolves differently. Code
//! that had the order baked in could not do that.
//!
//! Why the attributes are a table and not a switch
//! -----------------------------------------------
//! `CT_DefinedName` has sixteen attributes and the typed reader keeps
//! four (`workbook_xml.zig:58-68`). Three of the twelve it drops —
//! `function`, `vbProcedure`, `xlm` — mean the name is not a value at
//! all but a macro entry point, and a name resolved to a macro body is
//! a wrong answer rather than a missing one. So every attribute the
//! schema defines gets a row, the scanner reads the table, and an
//! attribute with no row refuses *before* anything is mutated. Same
//! rule, same reason, as M4b2's `CT_CellFormula` inventory.
//!
//! The macro three refuse **when referenced**, not when read: a
//! workbook may carry a macro name it never uses in a formula, and
//! refusing to open it would refuse files Excel opens. That is the same
//! shape `_xlpm.` and unknown `_xlnm.` already have (`symbols.zig`),
//! which is why the reason joins them in `decode.Refusal.Reason`
//! instead of starting a second vocabulary for one case.
//!
//! Table producers
//! ---------------
//! `<calculatedColumnFormula>` and `<totalsRowFormula>` are FORMULA
//! carriers M4b1 decoded and nothing has yet been allowed to *mean*
//! anything. They are producers: Excel materializes a calculated column
//! into every data cell of that column, so a member cell without its
//! own `<f>` is a cell this engine would recalculate as a blank. The
//! producer says what the cell should hold and the cell says nothing —
//! that disagreement refuses rather than picking a side.
//!
//! 3D references
//! -------------
//! `Sheet1:Sheet3!A1` is one reference over a *span* of sheets, and
//! §5.6g freezes what may consume one: exactly SUM, COUNT, COUNTA,
//! AVERAGE, MIN and MAX. Everything else refuses typed, and so does a
//! 3D span inside a CSE/DA array formula or under an intersection —
//! **pre-eval**, by walking the tree, because "refuses pre-persist" is
//! only true if the answer does not depend on having evaluated
//! anything.
//!
//! Provenance
//! ----------
//! Corpus-decided: the three `CT_DefinedName` attributes that actually
//! occur (`name`, `localSheetId`, `hidden` — 19/17/10 occurrences
//! across `tests/corpus/`), and the fact that **no** corpus workbook
//! carries a table producer or a 3D span, which is why every row below
//! that is not one of those three is **spec-pinned to ECMA-376 and
//! labelled as such** rather than claimed as oracle-derived. The
//! committed oracle manifests contain no 3D case either (§8.2), so the
//! whole of §5.6g ships spec-pinned; the per-function oracle legs land
//! with the functions themselves (AVERAGE/MIN/MAX are M4e).

const std = @import("std");
const assert = std.debug.assert;
const Allocator = std.mem.Allocator;

const coords = @import("zlsx_refs");
const decode = @import("decode.zig");
const parser = @import("parser.zig");
const registry = @import("registry.zig");

pub const PlaneTwo = decode.PlaneTwo;

// ─── refusals (§10) ──────────────────────────────────────────────

pub const Refusal = struct {
    reason: Reason,
    /// Byte offset into the part, when the refusal was found by
    /// scanning one; byte offset into the formula, when it was found by
    /// walking a tree.
    offset: u32 = 0,

    pub const Reason = enum {
        // ── `CT_DefinedName` attributes (§5.9) ──
        /// An attribute with no row in `defined_name_attrs`. Twelve of
        /// the sixteen are dropped by the typed reader and three of
        /// those change what the name *is*, so an unrecognized
        /// seventeenth refuses rather than being dropped with them.
        unknown_defined_name_attribute,
        /// A recognized attribute whose value is not the lexical form
        /// its type demands — a `localSheetId` that is not an unsigned
        /// int, a `hidden` that is not an `xsd:boolean`.
        bad_defined_name_attribute_value,
        /// `<definedName>` with no `name`. The schema marks it
        /// required; without it the entry names nothing.
        defined_name_missing_name,
        /// The part is not well-formed enough to find its names.
        malformed_defined_names_part,

        // ── table producers (§5.9) ──
        /// An attribute with no row in `table_formula_attrs`.
        unknown_table_formula_attribute,
        /// `array` present with a value outside `xsd:boolean`.
        bad_table_formula_attribute_value,
        /// A calculated column or totals row declares a producer and a
        /// cell it covers carries no `<f>`. Excel materializes the
        /// producer into every member cell; a member without one would
        /// recalculate as a blank, silently deleting a column.
        table_member_missing_formula,

        // ── 3D references (§5.6g) ──
        /// A 3D span reached a function outside the frozen v1 eligible
        /// list, or reached no function at all.
        three_d_ineligible_function,
        /// A 3D span inside a CSE or dynamic-array array formula.
        three_d_in_array_context,
        /// A 3D span under `@` or a legacy implicit intersection.
        three_d_in_intersection_context,
    };

    /// Exhaustive by construction — a new `Reason` fails to compile
    /// until it has a §10 plane.
    pub fn planeTwo(self: Refusal) PlaneTwo {
        return switch (self.reason) {
            .unknown_defined_name_attribute,
            .unknown_table_formula_attribute,
            .three_d_ineligible_function,
            .three_d_in_array_context,
            .three_d_in_intersection_context,
            => .FormulaUnsupportedConstruct,

            .bad_defined_name_attribute_value,
            .defined_name_missing_name,
            .malformed_defined_names_part,
            .bad_table_formula_attribute_value,
            .table_member_missing_formula,
            => .FormulaMalformedInput,
        };
    }
};

// ─── the `CT_DefinedName` attribute inventory (§5.9) ─────────────

/// What this engine does with an attribute. `classifyDefinedName`
/// switches on it, so a row cannot be added without deciding.
pub const Treatment = enum {
    /// Read into the typed model and acted on.
    modeled,
    /// Read, kept, carried back out — and read by nothing in v1.
    preserved,
    /// Its presence means the name is not a value. Referencing it
    /// refuses; carrying it does not.
    refuses_when_referenced,
};

pub const DefinedNameAttrRow = struct {
    /// The qualified name as it appears in the file. Exactly one row is
    /// prefixed (`xml:space`), and its prefix is part of the match.
    name: []const u8,
    treatment: Treatment,
    /// The refusal a `.refuses_when_referenced` row raises, in
    /// `decode`'s vocabulary — the one `symbols.Resolution` already
    /// speaks. Null for the others.
    refusal: ?decode.Refusal.Reason = null,
    note: []const u8,
};

/// ECMA-376 `CT_DefinedName` (§18.2.5), complete, plus the one
/// attribute the XML namespace contributes. `classifyDefinedName`
/// iterates this array — an attribute absent from it is
/// `unknown_defined_name_attribute` by construction, not by a forgotten
/// `else` branch.
///
/// Corpus-decided rows: `name`, `localSheetId`, `hidden` (the only
/// three that occur). Every other row is spec-pinned to ECMA-376.
pub const defined_name_attrs = [_]DefinedNameAttrRow{
    .{
        .name = "name",
        .treatment = .modeled,
        .note = "ST_DefinedName, required. A STRING carrier (M4b1 " ++
            "decision 6): entity-decoded and ST_Xstring-decoded, unlike " ++
            "the element's own body.",
    },
    .{
        .name = "localSheetId",
        .treatment = .modeled,
        .note = "`xsd:unsignedInt`. Present makes the name sheet-scoped, " ++
            "which is what shadows the workbook name of the same " ++
            "spelling (§5.9). Absent is workbook scope, not sheet 0.",
    },
    .{
        .name = "hidden",
        .treatment = .modeled,
        .note = "`xsd:boolean`, default false. A hidden name still " ++
            "resolves — hiding is a UI property, and Excel's own " ++
            "`_xlnm._FilterDatabase` is written hidden.",
    },
    .{
        .name = "function",
        .treatment = .refuses_when_referenced,
        .refusal = .macro_defined_name,
        .note = "The name is a macro function rather than a value. " ++
            "Resolving it as a value would answer a different question " ++
            "than the formula asked.",
    },
    .{
        .name = "vbProcedure",
        .treatment = .refuses_when_referenced,
        .refusal = .macro_defined_name,
        .note = "The name is a VBA procedure. Same reason as `function`, " ++
            "and the two co-occur.",
    },
    .{
        .name = "xlm",
        .treatment = .refuses_when_referenced,
        .refusal = .macro_defined_name,
        .note = "The name is an Excel 4.0 macro. The body is XLM, not a " ++
            "formula in the grammar §5.2 pins.",
    },
    .{
        .name = "functionGroupId",
        .treatment = .preserved,
        .note = "Which add-in function group a macro name belongs to. " ++
            "Meaningful only alongside `function`, which already refuses " ++
            "when referenced.",
    },
    .{
        .name = "publishToServer",
        .treatment = .preserved,
        .note = "Excel Services publication flag. No effect on what the " ++
            "name denotes.",
    },
    .{
        .name = "workbookParameter",
        .treatment = .preserved,
        .note = "Marks the name as an Excel Services parameter. The name " ++
            "still denotes its own body locally.",
    },
    .{
        .name = "comment",
        .treatment = .preserved,
        .note = "UI text shown in the name manager.",
    },
    .{
        .name = "customMenu",
        .treatment = .preserved,
        .note = "Legacy custom-menu text.",
    },
    .{
        .name = "description",
        .treatment = .preserved,
        .note = "UI description.",
    },
    .{
        .name = "help",
        .treatment = .preserved,
        .note = "Help-topic text.",
    },
    .{
        .name = "shortcutKey",
        .treatment = .preserved,
        .note = "Keyboard shortcut for a macro name.",
    },
    .{
        .name = "statusBar",
        .treatment = .preserved,
        .note = "Status-bar text.",
    },
    .{
        .name = "xml:space",
        .treatment = .preserved,
        .note = "`default` | `preserve`. The body is a FORMULA carrier " ++
            "preserved byte-exact either way, so this is recorded rather " ++
            "than acted on — the same answer M4b2 gave it on `<f>`.",
    },
};

comptime {
    // The inventory is complete only if it is also unambiguous: two rows
    // for one name would make the treatment depend on iteration order.
    for (defined_name_attrs, 0..) |a, i| {
        for (defined_name_attrs[i + 1 ..]) |b| {
            assert(!std.mem.eql(u8, a.name, b.name));
        }
        assert((a.refusal != null) == (a.treatment == .refuses_when_referenced));
    }
}

/// One `<definedName>`, with every attribute the schema defines
/// accounted for. Every slice borrows the part bytes and is still
/// **encoded** — decoding is the symbol layer's, and it decodes the
/// identifier and the body through different carrier classes.
pub const DefinedName = struct {
    /// Raw `name` attribute (a STRING carrier).
    raw_identifier: []const u8,
    /// Raw element text (a FORMULA carrier).
    raw_body: []const u8,
    local_sheet_id: ?u32,
    hidden: bool,
    /// Set by any of `function`, `vbProcedure`, `xlm`. Carried rather
    /// than refused, because a macro name a formula never mentions is
    /// not a reason to refuse a workbook.
    macro: bool,
    /// The refusal a *reference* to this name raises, or null. Derived
    /// from the inventory, so a new `.refuses_when_referenced` row
    /// reaches resolution without a second edit.
    refusal_when_referenced: ?decode.Refusal.Reason,
    /// The whole raw attribute region, borrowed from the part.
    raw_attrs: []const u8,
};

pub const DefinedNameResult = union(enum) {
    ok: DefinedName,
    refused: Refusal,
};

/// Classify one `<definedName>` against the inventory.
///
/// Pure, allocation-free and total: every byte string either produces a
/// `DefinedName` or a typed refusal, which is what makes the fuzz
/// target's "never classifies two ways" invariant checkable.
pub fn classifyDefinedName(el: decode.Element, body: []const u8) DefinedNameResult {
    var out: DefinedName = .{
        .raw_identifier = "",
        .raw_body = body,
        .local_sheet_id = null,
        .hidden = false,
        .macro = false,
        .refusal_when_referenced = null,
        .raw_attrs = el.attrs,
    };
    var saw_name = false;

    var it = el.attrIterator();
    while (it.next()) |a| {
        // Foreign-namespace attributes are exempt from the inventory —
        // M4b1 decision 8. `xml:space` is the one prefixed attribute
        // with a row, matched on its full qualified name.
        const row = rowForDefinedName(a.qname) orelse {
            if (a.prefix().len != 0 or std.mem.eql(u8, a.qname, "xmlns")) continue;
            return .{ .refused = .{
                .reason = .unknown_defined_name_attribute,
                .offset = @intCast(el.offset),
            } };
        };
        if (row.treatment == .refuses_when_referenced) {
            // The three macro flags are `xsd:boolean`: `function="0"` is
            // an ordinary name that a producer chose to be explicit
            // about, and refusing it would refuse a file that means
            // nothing unusual.
            const on = xsdBool(a.raw_value) orelse
                return badValue(el.offset);
            if (on) {
                out.macro = true;
                out.refusal_when_referenced = row.refusal.?;
            }
            continue;
        }
        if (std.mem.eql(u8, row.name, "name")) {
            out.raw_identifier = a.raw_value;
            saw_name = true;
        } else if (std.mem.eql(u8, row.name, "localSheetId")) {
            out.local_sheet_id = std.fmt.parseInt(u32, a.raw_value, 10) catch
                return badValue(el.offset);
        } else if (std.mem.eql(u8, row.name, "hidden")) {
            out.hidden = xsdBool(a.raw_value) orelse return badValue(el.offset);
        } else {
            // Every `.preserved` row lands here, and `raw_attrs` is what
            // preserves it. Listing them individually would be a second
            // inventory that could disagree with the first.
            assert(row.treatment == .preserved);
        }
    }

    if (!saw_name) return .{ .refused = .{
        .reason = .defined_name_missing_name,
        .offset = @intCast(el.offset),
    } };
    return .{ .ok = out };
}

fn badValue(offset: usize) DefinedNameResult {
    return .{ .refused = .{
        .reason = .bad_defined_name_attribute_value,
        .offset = @intCast(offset),
    } };
}

fn rowForDefinedName(qname: []const u8) ?DefinedNameAttrRow {
    for (defined_name_attrs) |row| {
        if (std.mem.eql(u8, row.name, qname)) return row;
    }
    return null;
}

/// `xsd:boolean`'s complete lexical space, and nothing else. Shared
/// with M4b2's inventory by value rather than by import: `calc.zig`
/// exports the same three lines, and neither file may import the other
/// (both are reached from the symbol layer).
pub fn xsdBool(s: []const u8) ?bool {
    if (std.mem.eql(u8, s, "1") or std.mem.eql(u8, s, "true")) return true;
    if (std.mem.eql(u8, s, "0") or std.mem.eql(u8, s, "false")) return false;
    return null;
}

pub const DefinedNames = struct {
    /// Holds the row list. Every slice inside a row borrows the part.
    arena: std.heap.ArenaAllocator,
    rows: []const DefinedName,

    pub fn deinit(self: *DefinedNames) void {
        self.arena.deinit();
        self.* = undefined;
    }
};

pub const DefinedNamesResult = union(enum) {
    ok: DefinedNames,
    refused: Refusal,
};

/// Scan `<definedNames>` out of `xl/workbook.xml`.
///
/// The typed reader keeps four fields and drops the attribute region
/// (`workbook_xml.zig:58-68`), which is exactly the region the
/// inventory is about — so this walks the part itself rather than
/// reading a view that has already thrown the evidence away.
pub fn scanDefinedNames(gpa: Allocator, xml: []const u8) error{OutOfMemory}!DefinedNamesResult {
    var arena = std.heap.ArenaAllocator.init(gpa);
    var keep = false;
    defer if (!keep) arena.deinit();
    const a = arena.allocator();

    var rows: std.ArrayListUnmanaged(DefinedName) = .empty;

    var sc = decode.Scanner.init(xml);
    var depth: usize = 0;
    var in_section: ?usize = null;
    var open: ?struct { el: decode.Element, body_start: usize, depth: usize } = null;

    while (sc.next() catch {
        return .{ .refused = .{ .reason = .malformed_defined_names_part } };
    }) |ev| {
        switch (ev) {
            .close => {
                if (depth == 0) {
                    return .{ .refused = .{ .reason = .malformed_defined_names_part } };
                }
                if (open) |o| {
                    if (depth == o.depth) {
                        // The body ends where the close tag begins.
                        const body = xml[o.body_start..closeTagStart(xml, sc.i)];
                        switch (classifyDefinedName(o.el, body)) {
                            .ok => |dn| try rows.append(a, dn),
                            .refused => |r| return .{ .refused = r },
                        }
                        open = null;
                    }
                }
                if (in_section) |d| {
                    if (depth == d) in_section = null;
                }
                depth -= 1;
                continue;
            },
            .text, .cdata, .doctype => continue,
            .open, .self_closing => {},
        }

        const el, const self_closing = switch (ev) {
            .open => |e| .{ e, false },
            .self_closing => |e| .{ e, true },
            else => unreachable,
        };
        if (!self_closing) depth += 1;
        const local = el.local();
        const el_depth = if (self_closing) depth + 1 else depth;

        if (el_depth == 2 and std.mem.eql(u8, local, "definedNames")) {
            if (!self_closing) in_section = depth;
        } else if (in_section != null and el_depth == in_section.? + 1 and
            std.mem.eql(u8, local, "definedName"))
        {
            if (self_closing) {
                // `<definedName name="x"/>` — a name with an empty body.
                switch (classifyDefinedName(el, "")) {
                    .ok => |dn| try rows.append(a, dn),
                    .refused => |r| return .{ .refused = r },
                }
            } else {
                open = .{ .el = el, .body_start = sc.i, .depth = depth };
            }
        }
    }

    const owned = try rows.toOwnedSlice(a);
    keep = true;
    return .{ .ok = .{ .arena = arena, .rows = owned } };
}

/// Where the close tag that ended at `end` began. The scanner reports
/// the byte *after* `</definedName>`; the body stops at its `<`.
fn closeTagStart(xml: []const u8, end: usize) usize {
    var i = end;
    while (i > 0) {
        i -= 1;
        if (xml[i] == '<') return i;
    }
    return end;
}

// ─── §5.9 resolution, driven by M2's exported orders ─────────────

/// A record of the stages a resolution actually visited, in order.
///
/// Its whole purpose is evidence: a caller can assert that the driver
/// walked §5.9's order and stopped where it says it stopped, without
/// the assertion being a second copy of the order.
pub fn Trace(comptime Stage: type) type {
    return struct {
        const Self = @This();
        /// Bounded by the enum: a driver cannot visit a stage twice.
        buf: [@typeInfo(Stage).@"enum".fields.len]Stage = undefined,
        len: usize = 0,

        pub fn push(self: *Self, s: Stage) void {
            assert(self.len < self.buf.len);
            self.buf[self.len] = s;
            self.len += 1;
        }

        pub fn slice(self: *const Self) []const Stage {
            return self.buf[0..self.len];
        }
    };
}

pub const ValueTrace = Trace(parser.ValueScope);
pub const CallTrace = Trace(parser.CallStage);

comptime {
    // §5.9's orders are only orders if they are also *complete*: a
    // stage missing from the array is a stage the driver never runs,
    // and a stage listed twice is one it runs twice.
    assertPermutation(parser.ValueScope, &parser.value_resolution_order);
    assertPermutation(parser.CallStage, &parser.call_resolution_order);
}

fn assertPermutation(comptime E: type, order: []const E) void {
    const fields = @typeInfo(E).@"enum".fields;
    assert(order.len == fields.len);
    for (order, 0..) |s, i| {
        for (order[i + 1 ..]) |t| assert(s != t);
    }
}

/// Walk a value-position resolution order and return the first stage
/// that answers.
///
/// `order` is a parameter, not a constant, and that is the point: the
/// driver has no opinion about §5.9's sequence, it *reads* one. Callers
/// pass `&parser.value_resolution_order`; the test that passes a
/// permuted array and watches the winner change is what proves this
/// file did not restate the order it was given.
///
/// `Lookup` supplies `fn at(self, stage, from, folded) ?Hit` and the
/// `Hit` type it answers with, so this file needs no knowledge of the
/// symbol layer that implements it (which imports *this* file).
pub fn resolveInOrder(
    comptime Lookup: type,
    lookup: Lookup,
    order: []const parser.ValueScope,
    from: ?u32,
    folded: []const u8,
    trace: ?*ValueTrace,
) ?Lookup.Hit {
    for (order) |stage| {
        if (trace) |t| t.push(stage);
        // `.name_error` is the terminal stage: "provably nowhere" is an
        // answer, and it is the one stage no lookup can supply.
        if (stage == .name_error) return null;
        if (lookup.at(stage, from, folded)) |hit| return hit;
    }
    return null;
}

/// What a call-position spelling resolves to (§5.9, §7).
pub const CallResolution = union(enum) {
    function: *const registry.Function,
    /// Not in the registry: `FormulaUnsupportedFunction`, never
    /// `#NAME?` — zlsx refuses rather than inventing an error value for
    /// a function it simply does not implement.
    unsupported,
};

/// A spelling with its layered prefixes removed.
pub const Stripped = struct {
    bare: []const u8,
    prefix: parser.Prefix,
};

/// `_xlfn.` and, layered inside it, `_xlws.` (§5.9). Excel writes them
/// so an older reader does not silently mis-evaluate a function it does
/// not know; the registry is keyed on the bare name.
///
/// Only ever layered outermost-first, which is why this peels rather
/// than looping: `_xlws._xlfn.X` is not a spelling Office produces, and
/// accepting it would make two spellings of one function.
pub fn stripLayeredPrefixes(raw: []const u8) Stripped {
    var out: Stripped = .{ .bare = raw, .prefix = .{} };
    if (startsWithIgnoreAsciiCase(out.bare, "_xlfn.")) {
        out.prefix.xlfn = true;
        out.bare = out.bare["_xlfn.".len..];
        if (startsWithIgnoreAsciiCase(out.bare, "_xlws.")) {
            out.prefix.xlws = true;
            out.bare = out.bare["_xlws.".len..];
        }
    }
    return out;
}

fn startsWithIgnoreAsciiCase(s: []const u8, prefix: []const u8) bool {
    if (s.len < prefix.len) return false;
    return std.ascii.eqlIgnoreCase(s[0..prefix.len], prefix);
}

/// Walk §5.9's call-position order: strip layered prefixes, ask the
/// registry, and otherwise refuse.
pub fn resolveCall(
    raw: []const u8,
    order: []const parser.CallStage,
    trace: ?*CallTrace,
) CallResolution {
    var bare = raw;
    for (order) |stage| {
        if (trace) |t| t.push(stage);
        switch (stage) {
            .strip_layered_prefixes => bare = stripLayeredPrefixes(bare).bare,
            .registry => if (registry.lookup(bare)) |f| return .{ .function = f },
            .unsupported_function => return .unsupported,
        }
    }
    return .unsupported;
}

// ─── name bodies: the interim depth guard (§5.9) ─────────────────

/// How deep a name may expand into another name before the run refuses.
///
/// Interim, and named as such: M5a makes name bodies graph nodes, where
/// a cycle through two names is a cycle like any other and the SCC
/// machinery answers it. Until then the evaluator expands a body
/// inline, and inline expansion of `A = B`, `B = A` has to stop
/// somewhere. Eight is far past any authored nesting and far short of
/// the expression-walk bound that would otherwise catch it as a stack
/// limit — a §9 limit with a name beats a stack limit without one.
pub const max_name_expansion_depth: usize = 8;

/// Whether a name body is position-dependent (§5.9).
///
/// A relative half — `A1`, `$A1`, `A$1`, a full row or column bound
/// without its `$` — means a different cell depending on where the name
/// is used from. Excel resolves that against the referencing cell; v1
/// does not carry the use site into name expansion, so referencing such
/// a name refuses (`relative_reference_name`) rather than expanding to
/// whichever cell the definition happened to be authored against.
/// Refused **when referenced**, not when read: a workbook may carry a
/// relative name no formula mentions.
///
/// Structured references and names nested inside the body are not
/// position-dependent in this sense — a table column is a table column
/// from anywhere — so they do not trip it.
pub fn bodyIsRelative(ast: parser.Ast) bool {
    return nodeIsRelative(ast, ast.root);
}

fn nodeIsRelative(ast: parser.Ast, i: parser.Index) bool {
    return switch (ast.node(i)) {
        .number, .string, .boolean, .error_lit, .missing_arg, .name, .structured => false,
        .ref_cell => |n| !(n.cell.anchor.col and n.cell.anchor.row),
        .ref_full_col => |n| !(n.first.absolute and n.last.absolute),
        .ref_full_row => |n| !(n.first.absolute and n.last.absolute),
        .array => |n| {
            for (ast.children(n.elems)) |e| {
                if (nodeIsRelative(ast, e)) return true;
            }
            return false;
        },
        .qualified => |n| nodeIsRelative(ast, n.target),
        .paren => |n| nodeIsRelative(ast, n.child),
        .call => |n| {
            for (ast.children(n.args)) |arg| {
                if (nodeIsRelative(ast, arg)) return true;
            }
            return false;
        },
        .unary => |n| nodeIsRelative(ast, n.child),
        .postfix => |n| nodeIsRelative(ast, n.child),
        .binary => |n| nodeIsRelative(ast, n.lhs) or nodeIsRelative(ast, n.rhs),
    };
}

/// Classify a decoded name body: parse it and answer the one question
/// resolution needs before M5a makes bodies graph nodes.
///
/// A body that does not parse answers `null` here — the expansion that
/// follows parses it too, and that parse's own refusal names the
/// construct rather than flattening it into "this name is bad".
pub fn bodyRefusal(gpa: Allocator, body: []const u8) error{OutOfMemory}!?decode.Refusal.Reason {
    var parsed = try parser.parse(gpa, body, .{});
    defer parsed.deinit(gpa);
    return switch (parsed) {
        .refused => null,
        .ok => |ast| if (bodyIsRelative(ast)) .relative_reference_name else null,
    };
}

// ─── the table producer inventory (§5.9) ─────────────────────────

pub const TableFormulaAttrRow = struct {
    name: []const u8,
    treatment: Treatment,
    note: []const u8,
};

/// ECMA-376 `CT_TableFormula` (§18.5.1.6, §18.5.1.88) — the type both
/// `<calculatedColumnFormula>` and `<totalsRowFormula>` have — plus the
/// XML namespace's contribution. Spec-pinned: no corpus workbook
/// carries either element.
pub const table_formula_attrs = [_]TableFormulaAttrRow{
    .{
        .name = "array",
        .treatment = .modeled,
        .note = "`xsd:boolean`, default false. The producer is a CSE " ++
            "array formula, which §5.6h places by declared range — and " ++
            "which §5.6g forbids a 3D span inside.",
    },
    .{
        .name = "xml:space",
        .treatment = .preserved,
        .note = "`default` | `preserve`. The body is a FORMULA carrier " ++
            "preserved byte-exact either way.",
    },
};

comptime {
    for (table_formula_attrs, 0..) |a, i| {
        for (table_formula_attrs[i + 1 ..]) |b| {
            assert(!std.mem.eql(u8, a.name, b.name));
        }
    }
}

/// Which producer an element is. They differ in what they cover, not in
/// what they are: one fills a column's data rows, the other one totals
/// cell.
pub const ProducerKind = enum { calculated_column, totals_row };

pub const TableFormula = struct {
    kind: ProducerKind,
    /// Decoded body (FORMULA carrier).
    text: []const u8,
    array: bool,
    raw_attrs: []const u8,
};

pub const TableFormulaResult = union(enum) {
    ok: TableFormula,
    refused: Refusal,
};

pub fn classifyTableFormula(
    kind: ProducerKind,
    raw_attrs: []const u8,
    text: []const u8,
    offset: u32,
) TableFormulaResult {
    var out: TableFormula = .{
        .kind = kind,
        .text = text,
        .array = false,
        .raw_attrs = raw_attrs,
    };
    var it: decode.AttrIterator = .{ .attrs = raw_attrs };
    while (it.next()) |a| {
        const row = rowForTableFormula(a.qname) orelse {
            if (a.prefix().len != 0 or std.mem.eql(u8, a.qname, "xmlns")) continue;
            return .{ .refused = .{
                .reason = .unknown_table_formula_attribute,
                .offset = offset,
            } };
        };
        if (std.mem.eql(u8, row.name, "array")) {
            out.array = xsdBool(a.raw_value) orelse return .{ .refused = .{
                .reason = .bad_table_formula_attribute_value,
                .offset = offset,
            } };
        } else {
            assert(row.treatment == .preserved);
        }
    }
    return .{ .ok = out };
}

fn rowForTableFormula(qname: []const u8) ?TableFormulaAttrRow {
    for (table_formula_attrs) |row| {
        if (std.mem.eql(u8, row.name, qname)) return row;
    }
    return null;
}

/// The cells one producer covers, in sheet coordinates.
pub const ProducerSpan = struct {
    col: coords.Col,
    first_row: coords.Row,
    last_row: coords.Row,
};

/// Where a producer's members sit, given the table's geometry.
///
/// Null when the geometry leaves the producer nothing to cover — an
/// empty data region under a calculated column, or a table with no
/// totals row under a totals producer. Nothing to check is not a
/// refusal; it is a table with no data rows yet.
pub fn producerSpan(
    kind: ProducerKind,
    ref: coords.Range,
    header_rows: u32,
    totals_rows: u32,
    col: coords.Col,
) ?ProducerSpan {
    const top = ref.first.row.oneBased();
    const bottom = ref.last.row.oneBased();
    if (bottom < top) return null;
    const height = bottom - top + 1;
    if (header_rows + totals_rows >= height) {
        // Header and totals consume the whole range: no data rows.
        if (kind == .calculated_column) return null;
    }
    switch (kind) {
        .calculated_column => {
            const first = top + header_rows;
            const last = bottom -| totals_rows;
            if (first > last) return null;
            return .{
                .col = col,
                .first_row = coords.Row.fromOneBased(first) catch return null,
                .last_row = coords.Row.fromOneBased(last) catch return null,
            };
        },
        .totals_row => {
            if (totals_rows == 0) return null;
            const first = bottom + 1 -| totals_rows;
            if (first > bottom) return null;
            return .{
                .col = col,
                .first_row = coords.Row.fromOneBased(first) catch return null,
                .last_row = coords.Row.fromOneBased(bottom) catch return null,
            };
        },
    }
}

/// Every member cell a producer covers must carry its own `<f>`.
///
/// Excel materializes a calculated column into each data cell, so the
/// part and the sheet say the same thing twice. When they disagree —
/// the table declares a producer and a member holds no formula — an
/// engine that writes values back would recalculate that member as a
/// blank and delete a column's worth of data. Neither side is
/// authoritative enough to pick, so this refuses.
///
/// `Cells` supplies `fn hasFormula(self, row, col) bool`.
pub fn checkProducerMembers(
    comptime Cells: type,
    cells: Cells,
    span: ProducerSpan,
) ?Refusal {
    var r = span.first_row.oneBased();
    while (r <= span.last_row.oneBased()) : (r += 1) {
        const row = coords.Row.fromOneBased(r) catch return null;
        if (!cells.hasFormula(row, span.col)) {
            return .{ .reason = .table_member_missing_formula };
        }
    }
    return null;
}

// ─── the 3D reference matrix (§5.6g) ─────────────────────────────

/// The frozen v1 eligible list. Exactly six, and the freeze is the
/// point: a seventh would need an oracle leg and a decision, not an
/// edit here.
pub const three_d_eligible = [_][]const u8{
    "SUM",
    "COUNT",
    "COUNTA",
    "AVERAGE",
    "MIN",
    "MAX",
};

comptime {
    // Eligible for v1 means *in* v1. A name that never reaches the
    // registry could not consume a 3D span whatever this list said.
    for (three_d_eligible) |name| {
        assert(inFrozenInventory(name));
    }
}

/// Whether a name has a row in the committed frozen inventory (§7).
fn inFrozenInventory(comptime name: []const u8) bool {
    @setEvalBranchQuota(200_000);
    var it = std.mem.splitScalar(u8, registry.inventory_v1, '\n');
    while (it.next()) |line| {
        if (line.len == 0 or line[0] == '#') continue;
        const tab = std.mem.indexOfScalar(u8, line, '\t') orelse continue;
        if (std.mem.eql(u8, line[0..tab], name)) return true;
    }
    return false;
}

/// Whether `name` — as spelled at a call site, prefixes included — may
/// consume a 3D span.
pub fn threeDEligible(name: []const u8) bool {
    const bare = stripLayeredPrefixes(name).bare;
    for (three_d_eligible) |e| {
        if (std.ascii.eqlIgnoreCase(e, bare)) return true;
    }
    return false;
}

/// The contexts §5.6g forbids a 3D span in, as the caller knows them
/// before evaluating anything.
pub const ThreeDContext = struct {
    /// The formula is a legacy CSE array (`<f t="array">`) or a
    /// dynamic-array formula. Both place a *result array*, and a 3D
    /// span inside one has no defined placement.
    array_formula: bool = false,
};

/// Walk a parsed formula and refuse any 3D span §5.6g does not allow.
///
/// Pre-eval by construction: it reads the tree and the context, never
/// the workbook, so "refuses pre-persist" is a property of the check
/// rather than a promise about where it is called.
///
/// Eligibility is judged against the **enclosing call argument**, with
/// parentheses and the three reference operators allowed in between —
/// `:`, ` ` and `,` compose references rather than consume them, so
/// `SUM(S1:S3!A1:B1)` is still SUM's argument. Every other operator
/// consumes a value: `SUM(S1:S3!A1*2)` hands `*` a multi-area
/// reference, which is a different question with a different answer,
/// and refusing is the direction an engine that writes values back can
/// afford to be wrong in. Spec-pinned — no oracle case exists (§8.2).
pub fn checkThreeD(ast: parser.Ast, ctx: ThreeDContext) ?Refusal {
    return walkThreeD(ast, ast.root, ctx, .{});
}

const Enclosing = struct {
    /// The callee spelling whose argument list we are directly inside,
    /// or null once an operator has intervened.
    callee: ?[]const u8 = null,
    /// Under `@`, `_xlfn.SINGLE`, or the legacy intersection operator.
    intersection: bool = false,
};

fn walkThreeD(
    ast: parser.Ast,
    i: parser.Index,
    ctx: ThreeDContext,
    encl: Enclosing,
) ?Refusal {
    switch (ast.node(i)) {
        .number, .string, .boolean, .error_lit, .missing_arg, .name, .structured => return null,
        .ref_cell, .ref_full_col, .ref_full_row => return null,
        // An array constant holds only literals today, but walking it
        // keeps this total over trees a later row may build.
        .array => |n| {
            for (ast.children(n.elems)) |e| {
                if (walkThreeD(ast, e, ctx, .{})) |r| return r;
            }
            return null;
        },
        .qualified => |n| {
            if (isSpan(n.sheet)) {
                if (ctx.array_formula) return .{
                    .reason = .three_d_in_array_context,
                    .offset = n.span.start,
                };
                if (encl.intersection) return .{
                    .reason = .three_d_in_intersection_context,
                    .offset = n.span.start,
                };
                const callee = encl.callee orelse return .{
                    .reason = .three_d_ineligible_function,
                    .offset = n.span.start,
                };
                if (!threeDEligible(callee)) return .{
                    .reason = .three_d_ineligible_function,
                    .offset = n.span.start,
                };
            }
            // The target of a qualified reference cannot itself be
            // qualified, but walking it costs nothing and keeps this
            // total over trees the parser did not build.
            return walkThreeD(ast, n.target, ctx, .{ .intersection = encl.intersection });
        },
        .paren => |n| return walkThreeD(ast, n.child, ctx, encl),
        .call => |n| {
            const callee = switch (ast.node(n.callee)) {
                .name => |c| c.raw,
                else => null,
            };
            if (walkThreeD(ast, n.callee, ctx, .{})) |r| return r;
            for (ast.children(n.args)) |arg| {
                if (walkThreeD(ast, arg, ctx, .{
                    .callee = callee,
                    .intersection = encl.intersection,
                })) |r| return r;
            }
            return null;
        },
        .unary => |n| return walkThreeD(ast, n.child, ctx, .{
            // `+x` and `-x` consume a value, so whatever call they sit
            // in is no longer the span's consumer.
            .callee = if (n.op == .implicit_intersection) encl.callee else null,
            .intersection = encl.intersection or n.op == .implicit_intersection,
        }),
        .postfix => |n| return walkThreeD(ast, n.child, ctx, .{
            // `A1#` names a reference; `x%` divides one by 100.
            .callee = if (n.op == .spill) encl.callee else null,
            .intersection = encl.intersection,
        }),
        .binary => |n| {
            // The three reference operators *compose* references —
            // `Sheet1:Sheet3!A1:B1` is one 3D area, not a span handed
            // to an operator — so the enclosing call is still the
            // consumer through them. Every other operator consumes a
            // value, and a span that reaches one has left the function
            // §5.6g made eligible.
            const composes = switch (n.op) {
                .range, .intersect, .union_op => true,
                else => false,
            };
            const inner: Enclosing = .{
                .callee = if (composes) encl.callee else null,
                .intersection = encl.intersection or n.op == .intersect,
            };
            if (walkThreeD(ast, n.lhs, ctx, inner)) |r| return r;
            return walkThreeD(ast, n.rhs, ctx, inner);
        },
    }
}

/// Whether a sheet prefix names a span rather than one sheet.
///
/// Two spellings reach here. `Sheet1:Sheet3!A1` arrives with `last`
/// set. `'Q1:Q4'!A1` arrives as one quoted token, because the tokenizer
/// cannot know whether the quotes hold one name with a colon in it or
/// two names around one — and Excel forbids `:` in a sheet name, so a
/// colon inside the quotes is a span every time.
pub fn isSpan(spec: parser.SheetSpec) bool {
    if (spec.last != null) return true;
    if (!spec.quoted) return std.mem.indexOfScalar(u8, spec.first, ':') != null;
    const raw = spec.first;
    const body = if (raw.len >= 2 and raw[0] == '\'' and raw[raw.len - 1] == '\'')
        raw[1 .. raw.len - 1]
    else
        raw;
    return std.mem.indexOfScalar(u8, body, ':') != null;
}

/// The two endpoint spellings of a span, unquoted by the caller.
pub const SpanEndpoints = struct {
    first: []const u8,
    last: []const u8,
};

/// Split a span spelling into its endpoints, or null when it is not a
/// span. Operates on the *already unquoted* sheet text, so the quoted
/// and unquoted spellings converge before they get here.
pub fn splitSpan(spec: parser.SheetSpec, unquoted_first: []const u8) ?SpanEndpoints {
    if (spec.last) |l| return .{ .first = unquoted_first, .last = l };
    const colon = std.mem.indexOfScalar(u8, unquoted_first, ':') orelse return null;
    return .{
        .first = unquoted_first[0..colon],
        .last = unquoted_first[colon + 1 ..],
    };
}

/// What a span's endpoints expand to.
pub const SpanExpansion = union(enum) {
    /// Inclusive workbook-order indices. `first <= last` always.
    members: struct { first: u32, last: u32 },
    /// `#REF!`, pinned: an endpoint names no sheet, or the two are in
    /// the wrong workbook order. A deleted endpoint is the ordinary way
    /// the first happens, and Excel leaves the surviving spelling in
    /// place rather than repairing the span (§5.6g).
    ref_error,
};

/// Expand a span over the workbook's sheet order.
///
/// Inclusive in both directions: `Sheet1:Sheet3` is three sheets, and
/// `Sheet2:Sheet2` is one. Reordering is *not* silently normalized —
/// §5.6g pins `#REF!` for it, on the ground that a span whose endpoints
/// have swapped is a span someone's edit broke, and computing over the
/// sheets between them would answer a question no one asked.
pub fn expandSpan(first: ?u32, last: ?u32) SpanExpansion {
    const f = first orelse return .ref_error;
    const l = last orelse return .ref_error;
    if (f > l) return .ref_error;
    return .{ .members = .{ .first = f, .last = l } };
}

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

fn elementOf(xml: []const u8) decode.Element {
    var sc = decode.Scanner.init(xml);
    while (sc.next() catch unreachable) |ev| {
        switch (ev) {
            .open, .self_closing => |e| return e,
            else => continue,
        }
    }
    unreachable;
}

test "defined names: every inventory row has a fixture, and the table is what refuses" {
    // The gate M4b2 set: walk the table, and fail on a row nothing
    // exercises. A row added without a fixture fails here, not in review.
    for (defined_name_attrs) |row| {
        const spelling = if (std.mem.eql(u8, row.name, "name"))
            "<definedName name=\"N\">1</definedName>"
        else if (std.mem.eql(u8, row.name, "localSheetId"))
            "<definedName name=\"N\" localSheetId=\"2\">1</definedName>"
        else if (std.mem.eql(u8, row.name, "hidden"))
            "<definedName name=\"N\" hidden=\"1\">1</definedName>"
        else if (row.treatment == .refuses_when_referenced)
            "<definedName name=\"N\" REPLACE=\"1\">1</definedName>"
        else
            "<definedName name=\"N\" REPLACE=\"x\">1</definedName>";

        var buf: [128]u8 = undefined;
        const xml = blk: {
            if (std.mem.indexOf(u8, spelling, "REPLACE")) |at| {
                @memcpy(buf[0..at], spelling[0..at]);
                @memcpy(buf[at..][0..row.name.len], row.name);
                const rest = spelling[at + "REPLACE".len ..];
                @memcpy(buf[at + row.name.len ..][0..rest.len], rest);
                break :blk buf[0 .. at + row.name.len + rest.len];
            }
            break :blk spelling;
        };

        const el = elementOf(xml);
        const got = classifyDefinedName(el, "1");
        try testing.expect(got == .ok);
        switch (row.treatment) {
            .modeled => {},
            .preserved => {
                // Preserved means *carried*, and `raw_attrs` is the
                // carrier. Nothing else in the record may mention it.
                try testing.expect(std.mem.indexOf(u8, got.ok.raw_attrs, row.name) != null);
                try testing.expect(!got.ok.macro);
            },
            .refuses_when_referenced => {
                try testing.expect(got.ok.macro);
                try testing.expectEqual(row.refusal.?, got.ok.refusal_when_referenced.?);
                try testing.expectEqual(
                    decode.PlaneTwo.FormulaUnsupportedConstruct,
                    (decode.Refusal{ .reason = got.ok.refusal_when_referenced.? }).planeTwo(),
                );
            },
        }
    }
}

test "defined names: the inventory is complete — sixteen rows, three of them macro" {
    try testing.expectEqual(@as(usize, 16), defined_name_attrs.len);
    var macro: usize = 0;
    var modeled: usize = 0;
    for (defined_name_attrs) |row| {
        switch (row.treatment) {
            .refuses_when_referenced => macro += 1,
            .modeled => modeled += 1,
            .preserved => {},
        }
    }
    try testing.expectEqual(@as(usize, 3), macro);
    try testing.expectEqual(@as(usize, 3), modeled);
}

test "defined names: an attribute with no row refuses before anything is mutated" {
    const el = elementOf("<definedName name=\"N\" futureThing=\"1\">1</definedName>");
    const got = classifyDefinedName(el, "1");
    try testing.expectEqual(Refusal.Reason.unknown_defined_name_attribute, got.refused.reason);
    try testing.expectEqual(PlaneTwo.FormulaUnsupportedConstruct, got.refused.planeTwo());

    // …but a foreign-namespace attribute is exempt (M4b1 decision 8).
    const ok = classifyDefinedName(
        elementOf("<definedName name=\"N\" x14:someExt=\"1\">1</definedName>"),
        "1",
    );
    try testing.expect(ok == .ok);
}

test "defined names: the three macro flags are booleans, and `0` is an ordinary name" {
    for ([_][]const u8{ "function", "vbProcedure", "xlm" }) |flag| {
        var buf: [96]u8 = undefined;
        const on = try std.fmt.bufPrint(&buf, "<definedName name=\"N\" {s}=\"1\"/>", .{flag});
        const got = classifyDefinedName(elementOf(on), "");
        try testing.expect(got.ok.macro);
        try testing.expectEqual(
            decode.Refusal.Reason.macro_defined_name,
            got.ok.refusal_when_referenced.?,
        );

        var buf2: [96]u8 = undefined;
        const off = try std.fmt.bufPrint(&buf2, "<definedName name=\"N\" {s}=\"0\"/>", .{flag});
        const plain = classifyDefinedName(elementOf(off), "");
        try testing.expect(!plain.ok.macro);
        try testing.expect(plain.ok.refusal_when_referenced == null);

        var buf3: [96]u8 = undefined;
        const bad = try std.fmt.bufPrint(&buf3, "<definedName name=\"N\" {s}=\"TRUE\"/>", .{flag});
        const refused = classifyDefinedName(elementOf(bad), "");
        try testing.expectEqual(
            Refusal.Reason.bad_defined_name_attribute_value,
            refused.refused.reason,
        );
    }
}

test "defined names: bad lexical values and a missing name refuse" {
    const bad_scope = classifyDefinedName(
        elementOf("<definedName name=\"N\" localSheetId=\"-1\"/>"),
        "",
    );
    try testing.expectEqual(
        Refusal.Reason.bad_defined_name_attribute_value,
        bad_scope.refused.reason,
    );
    const bad_hidden = classifyDefinedName(
        elementOf("<definedName name=\"N\" hidden=\"yes\"/>"),
        "",
    );
    try testing.expectEqual(
        Refusal.Reason.bad_defined_name_attribute_value,
        bad_hidden.refused.reason,
    );
    const no_name = classifyDefinedName(elementOf("<definedName hidden=\"1\"/>"), "");
    try testing.expectEqual(Refusal.Reason.defined_name_missing_name, no_name.refused.reason);
    try testing.expectEqual(PlaneTwo.FormulaMalformedInput, no_name.refused.planeTwo());
}

const wb_head = "<workbook xmlns=\"" ++ decode.ns_main ++ "\">";

test "defined names: the scanner finds them in the part, bodies raw" {
    const xml = wb_head ++
        "<sheets><sheet name=\"S\"/></sheets>" ++
        "<definedNames>" ++
        "<definedName name=\"Global\">S!$A$1</definedName>" ++
        "<definedName name=\"Local\" localSheetId=\"0\" hidden=\"1\">S!$B$2</definedName>" ++
        "<definedName name=\"Empty\"/>" ++
        "<definedName name=\"Lit\">\"_x0041_\"&amp;S!A1</definedName>" ++
        "</definedNames></workbook>";

    var got = switch (try scanDefinedNames(testing.allocator, xml)) {
        .ok => |g| g,
        .refused => return error.TestUnexpectedRefusal,
    };
    defer got.deinit();

    try testing.expectEqual(@as(usize, 4), got.rows.len);
    try testing.expectEqualStrings("Global", got.rows[0].raw_identifier);
    try testing.expectEqualStrings("S!$A$1", got.rows[0].raw_body);
    try testing.expectEqual(@as(?u32, null), got.rows[0].local_sheet_id);
    try testing.expectEqual(@as(?u32, 0), got.rows[1].local_sheet_id);
    try testing.expect(got.rows[1].hidden);
    try testing.expectEqualStrings("", got.rows[2].raw_body);
    // Still encoded: the body is a FORMULA carrier and decoding it is
    // the symbol layer's, one pass only.
    try testing.expectEqualStrings("\"_x0041_\"&amp;S!A1", got.rows[3].raw_body);
}

test "defined names: a name outside <definedNames> is not a defined name" {
    // `<definedName>` inside a sheet's extension list, say. Depth is
    // what distinguishes them, and a scanner that matched on the tag
    // alone would invent names from someone else's vocabulary.
    const xml = wb_head ++ "<extLst><ext><definedName name=\"Ghost\"/></ext></extLst></workbook>";
    var got = switch (try scanDefinedNames(testing.allocator, xml)) {
        .ok => |g| g,
        .refused => return error.TestUnexpectedRefusal,
    };
    defer got.deinit();
    try testing.expectEqual(@as(usize, 0), got.rows.len);
}

test "defined names: a refusal inside the section is the whole scan's answer" {
    const xml = wb_head ++ "<definedNames>" ++
        "<definedName name=\"Fine\">1</definedName>" ++
        "<definedName name=\"Bad\" nope=\"1\">1</definedName>" ++
        "</definedNames></workbook>";
    switch (try scanDefinedNames(testing.allocator, xml)) {
        .ok => |g| {
            var g2 = g;
            g2.deinit();
            return error.TestExpectedRefusal;
        },
        .refused => |r| try testing.expectEqual(
            Refusal.Reason.unknown_defined_name_attribute,
            r.reason,
        ),
    }
}

// ─── §5.9 order consumption ──────────────────────────────────────

const StubLookup = struct {
    /// Which stages have something to answer with.
    sheet_scoped: bool = false,
    workbook: bool = false,
    table: bool = false,
    builtin: bool = false,

    pub const Hit = parser.ValueScope;

    pub fn at(self: StubLookup, stage: parser.ValueScope, from: ?u32, folded: []const u8) ?Hit {
        _ = folded;
        return switch (stage) {
            .sheet_scoped_name => if (self.sheet_scoped and from != null) stage else null,
            .workbook_name => if (self.workbook) stage else null,
            .table => if (self.table) stage else null,
            .builtin_xlnm => if (self.builtin) stage else null,
            .name_error => null,
        };
    }
};

test "resolution: the driver walks M2's exported value order, and stops where it answers" {
    const lookup: StubLookup = .{ .workbook = true, .table = true };
    var trace: ValueTrace = .{};
    const hit = resolveInOrder(
        StubLookup,
        lookup,
        &parser.value_resolution_order,
        0,
        "n",
        &trace,
    );
    try testing.expectEqual(parser.ValueScope.workbook_name, hit.?);
    // The stages actually visited, and no more: the sheet-scoped tier
    // was asked and declined, the table tier was never reached.
    try testing.expectEqualSlices(
        parser.ValueScope,
        &.{ .sheet_scoped_name, .workbook_name },
        trace.slice(),
    );
}

test "resolution: shadowing is the order's, not the lookup's" {
    const lookup: StubLookup = .{ .sheet_scoped = true, .workbook = true };
    // From a sheet, the sheet-scoped tier wins…
    try testing.expectEqual(
        parser.ValueScope.sheet_scoped_name,
        resolveInOrder(StubLookup, lookup, &parser.value_resolution_order, 1, "n", null).?,
    );
    // …and from workbook level there is no sheet tier to win.
    try testing.expectEqual(
        parser.ValueScope.workbook_name,
        resolveInOrder(StubLookup, lookup, &parser.value_resolution_order, null, "n", null).?,
    );
}

test "resolution: the order is READ, not restated — a permuted order resolves differently" {
    // The proof the goal asks for. If §5.9's sequence were baked into
    // this file as a chain of `if`s, handing the driver a different
    // array could not change the answer. It does.
    const lookup: StubLookup = .{ .workbook = true, .table = true };
    const permuted = [_]parser.ValueScope{
        .table,
        .sheet_scoped_name,
        .workbook_name,
        .builtin_xlnm,
        .name_error,
    };
    try testing.expectEqual(
        parser.ValueScope.table,
        resolveInOrder(StubLookup, lookup, &permuted, 0, "n", null).?,
    );
    // …while the shipped order still puts the workbook name first.
    try testing.expectEqual(
        parser.ValueScope.workbook_name,
        resolveInOrder(StubLookup, lookup, &parser.value_resolution_order, 0, "n", null).?,
    );
}

test "resolution: nowhere is #NAME?'s evidence, and every stage was tried" {
    var trace: ValueTrace = .{};
    const hit = resolveInOrder(
        StubLookup,
        .{},
        &parser.value_resolution_order,
        0,
        "n",
        &trace,
    );
    try testing.expect(hit == null);
    try testing.expectEqualSlices(
        parser.ValueScope,
        &parser.value_resolution_order,
        trace.slice(),
    );
}

test "resolution: call position strips layered prefixes, then asks the registry" {
    var trace: CallTrace = .{};
    const r = resolveCall("_xlfn.SUM", &parser.call_resolution_order, &trace);
    try testing.expectEqualStrings("SUM", r.function.name);
    try testing.expectEqualSlices(
        parser.CallStage,
        &.{ .strip_layered_prefixes, .registry },
        trace.slice(),
    );

    // Unregistered is a refusal, never `#NAME?` (§7, decision 11).
    var t2: CallTrace = .{};
    try testing.expect(resolveCall("NOSUCHFN", &parser.call_resolution_order, &t2) == .unsupported);
    try testing.expectEqualSlices(parser.CallStage, &parser.call_resolution_order, t2.slice());
}

test "resolution: the call order is READ too — strip it and the prefixed name misses" {
    // The same proof the value order gets. With the strip stage removed
    // from the order, `_xlfn.SUM` reaches the registry spelled as the
    // file spells it and is not found — which is only possible if the
    // driver ran the array it was handed.
    const without_strip = [_]parser.CallStage{ .registry, .unsupported_function };
    try testing.expect(resolveCall("_xlfn.SUM", &without_strip, null) == .unsupported);
    // The bare spelling still resolves, so nothing else changed.
    try testing.expectEqualStrings(
        "SUM",
        resolveCall("SUM", &without_strip, null).function.name,
    );
}

test "resolution: `_xlws.` layers inside `_xlfn.`, and only there" {
    const both = stripLayeredPrefixes("_xlfn._xlws.FILTER");
    try testing.expectEqualStrings("FILTER", both.bare);
    try testing.expect(both.prefix.xlfn and both.prefix.xlws);

    // The reverse layering is not a spelling Office writes, and
    // accepting it would give one function two names.
    const reversed = stripLayeredPrefixes("_xlws._xlfn.FILTER");
    try testing.expectEqualStrings("_xlws._xlfn.FILTER", reversed.bare);
    try testing.expect(!reversed.prefix.xlfn and !reversed.prefix.xlws);

    // Excel's own prefix, in Excel's own casing or otherwise.
    try testing.expectEqualStrings("X", stripLayeredPrefixes("_XLFN.X").bare);
}

// ─── name bodies ─────────────────────────────────────────────────

test "name bodies: a relative half makes the name position-dependent" {
    for ([_][]const u8{
        "S!A1",
        "S!$A1",
        "S!A$1",
        "A1:$B$2",
        "SUM(S!A1:A9)",
        "A:A",
        "$A:A",
        "1:1",
        "-S!A1",
    }) |body| {
        const r = (try bodyRefusal(testing.allocator, body)) orelse {
            std.debug.print("expected relative: {s}\n", .{body});
            return error.TestExpectedRefusal;
        };
        try testing.expectEqual(decode.Refusal.Reason.relative_reference_name, r);
        try testing.expectEqual(
            decode.PlaneTwo.FormulaUnsupportedConstruct,
            (decode.Refusal{ .reason = r }).planeTwo(),
        );
    }

    // Fully anchored, or nothing to anchor: a name a formula may expand.
    for ([_][]const u8{
        "S!$A$1",
        "'R&D'!$A$1:$C$9",
        "SUM($A$1:$B$2)",
        "$A:$A",
        "$1:$5",
        "42",
        "\"text\"",
        "Other",
        "Sales[Total]",
    }) |body| {
        if (try bodyRefusal(testing.allocator, body)) |r| {
            std.debug.print("expected absolute: {s} → {t}\n", .{ body, r });
            return error.TestUnexpectedRefusal;
        }
    }
}

test "name bodies: a body that will not parse is the expansion's refusal, not the name's" {
    // `((` names nothing wrong about the *name*; the parse that expands
    // it says what is wrong, with the construct's own offset.
    try testing.expect((try bodyRefusal(testing.allocator, "((")) == null);
}

test "name bodies: LAMBDA and LET refuse through the registry, not a second rule" {
    // §5.9 refuses both through v1. They are unregistered calls, so the
    // call-position order already answers — inventing a name-level rule
    // would be a second statement of one refusal.
    for ([_][]const u8{ "LAMBDA", "LET", "_xlfn.LAMBDA" }) |name| {
        try testing.expect(resolveCall(name, &parser.call_resolution_order, null) == .unsupported);
    }
}

// ─── table producers ─────────────────────────────────────────────

test "table producers: both elements classify, and an unknown attribute refuses" {
    for ([_]ProducerKind{ .calculated_column, .totals_row }) |kind| {
        const plain = classifyTableFormula(kind, "", "SUM(T[c])", 0);
        try testing.expect(!plain.ok.array);
        try testing.expectEqual(kind, plain.ok.kind);

        const arr = classifyTableFormula(kind, " array=\"1\"", "SUM(T[c])", 0);
        try testing.expect(arr.ok.array);

        const bad = classifyTableFormula(kind, " array=\"maybe\"", "x", 0);
        try testing.expectEqual(
            Refusal.Reason.bad_table_formula_attribute_value,
            bad.refused.reason,
        );

        const unknown = classifyTableFormula(kind, " futureThing=\"1\"", "x", 0);
        try testing.expectEqual(
            Refusal.Reason.unknown_table_formula_attribute,
            unknown.refused.reason,
        );

        // `xml:space` is a row; a foreign namespace is exempt.
        try testing.expect(classifyTableFormula(kind, " xml:space=\"preserve\"", "x", 0) == .ok);
        try testing.expect(classifyTableFormula(kind, " x14:z=\"1\"", "x", 0) == .ok);
    }
}

const FakeCells = struct {
    /// Rows (1-based) that carry an `<f>` in the producer's column.
    with_formula: []const u32,

    pub fn hasFormula(self: FakeCells, row: coords.Row, col: coords.Col) bool {
        _ = col;
        for (self.with_formula) |r| {
            if (r == row.oneBased()) return true;
        }
        return false;
    }
};

test "table producers: a member cell without its own <f> refuses" {
    const ref = try coords.parseRange("A1:B4", .{});
    const span = producerSpan(.calculated_column, ref, 1, 0, ref.last.col).?;
    try testing.expectEqual(@as(u32, 2), span.first_row.oneBased());
    try testing.expectEqual(@as(u32, 4), span.last_row.oneBased());

    // Every data row carries one: nothing to say.
    try testing.expect(checkProducerMembers(
        FakeCells,
        .{ .with_formula = &.{ 2, 3, 4 } },
        span,
    ) == null);

    // One does not, and that member would recalculate as a blank.
    const missing = checkProducerMembers(
        FakeCells,
        .{ .with_formula = &.{ 2, 4 } },
        span,
    ).?;
    try testing.expectEqual(Refusal.Reason.table_member_missing_formula, missing.reason);
    try testing.expectEqual(PlaneTwo.FormulaMalformedInput, missing.planeTwo());
}

test "table producers: the totals producer covers the totals row, and only it" {
    const ref = try coords.parseRange("A1:B5", .{});
    const totals = producerSpan(.totals_row, ref, 1, 1, ref.first.col).?;
    try testing.expectEqual(@as(u32, 5), totals.first_row.oneBased());
    try testing.expectEqual(@as(u32, 5), totals.last_row.oneBased());

    // The calculated column stops short of it.
    const data = producerSpan(.calculated_column, ref, 1, 1, ref.first.col).?;
    try testing.expectEqual(@as(u32, 2), data.first_row.oneBased());
    try testing.expectEqual(@as(u32, 4), data.last_row.oneBased());

    // …and a totals member with no `<f>` refuses on the same rule the
    // calculated column does — one check, two producers.
    try testing.expect(checkProducerMembers(
        FakeCells,
        .{ .with_formula = &.{5} },
        totals,
    ) == null);
    try testing.expectEqual(
        Refusal.Reason.table_member_missing_formula,
        checkProducerMembers(FakeCells, .{ .with_formula = &.{} }, totals).?.reason,
    );

    // A table with no totals row has no totals members to check…
    try testing.expect(producerSpan(.totals_row, ref, 1, 0, ref.first.col) == null);
    // …and a header-only table has no data members either.
    const header_only = try coords.parseRange("A1:B1", .{});
    try testing.expect(
        producerSpan(.calculated_column, header_only, 1, 0, header_only.first.col) == null,
    );
}

// ─── the 3D matrix ───────────────────────────────────────────────

fn parseFor3D(src: []const u8) !parser.Ast {
    return switch (try parser.parse(testing.allocator, src, .{})) {
        .ok => |ast| ast,
        .refused => error.TestUnexpectedRefusal,
    };
}

test "3D: a fixture per eligible function — the frozen six accept a span" {
    try testing.expectEqual(@as(usize, 6), three_d_eligible.len);
    for (three_d_eligible) |name| {
        var buf: [64]u8 = undefined;
        const src = try std.fmt.bufPrint(&buf, "{s}(Sheet1:Sheet3!A1)", .{name});
        var ast = try parseFor3D(src);
        defer ast.deinit(testing.allocator);
        try testing.expect(checkThreeD(ast, .{}) == null);
        try testing.expect(threeDEligible(name));
    }
    // Case and the layered prefix are both Excel's, not the user's.
    try testing.expect(threeDEligible("sum"));
    try testing.expect(threeDEligible("_xlfn.SUM"));
}

test "3D: every other reference-consuming function refuses typed" {
    // Three ineligible ones, plus a span with no consumer at all.
    for ([_][]const u8{
        "COUNTBLANK(Sheet1:Sheet3!A1)",
        "SUMIF(Sheet1:Sheet3!A1,1)",
        "COUNTIF(Sheet1:Sheet3!A1,1)",
        "IF(Sheet1:Sheet3!A1,1,2)",
        "Sheet1:Sheet3!A1",
        "Sheet1:Sheet3!A1+1",
        // Eligible callee, but an operator intervened: the function
        // never sees the span, an operator does.
        "SUM(Sheet1:Sheet3!A1*2)",
    }) |src| {
        var ast = try parseFor3D(src);
        defer ast.deinit(testing.allocator);
        const r = checkThreeD(ast, .{}) orelse {
            std.debug.print("expected a refusal for: {s}\n", .{src});
            return error.TestExpectedRefusal;
        };
        try testing.expectEqual(Refusal.Reason.three_d_ineligible_function, r.reason);
        try testing.expectEqual(PlaneTwo.FormulaUnsupportedConstruct, r.planeTwo());
    }
}

test "3D: an array context refuses before evaluation, whatever consumes the span" {
    var ast = try parseFor3D("SUM(Sheet1:Sheet3!A1)");
    defer ast.deinit(testing.allocator);
    // The same formula that is legal as an ordinary cell formula…
    try testing.expect(checkThreeD(ast, .{}) == null);
    // …refuses as a CSE/DA array formula.
    const r = checkThreeD(ast, .{ .array_formula = true }).?;
    try testing.expectEqual(Refusal.Reason.three_d_in_array_context, r.reason);
}

test "3D: an intersection context refuses, in both of its spellings" {
    for ([_][]const u8{
        "SUM(@Sheet1:Sheet3!A1)",
        "SUM(_xlfn.SINGLE(Sheet1:Sheet3!A1))",
        "SUM(A1:A9 Sheet1:Sheet3!A1)",
    }) |src| {
        var ast = try parseFor3D(src);
        defer ast.deinit(testing.allocator);
        const r = checkThreeD(ast, .{}) orelse {
            std.debug.print("expected a refusal for: {s}\n", .{src});
            return error.TestExpectedRefusal;
        };
        try testing.expectEqual(Refusal.Reason.three_d_in_intersection_context, r.reason);
    }
}

test "3D: a formula with no span is untouched by the check" {
    for ([_][]const u8{
        "SUM(A1:A9)",
        "Sheet1!A1+Sheet2!B2",
        "COUNTBLANK(Sheet1!A1)",
        "@A1:A9",
        "{1,2;3,4}",
    }) |src| {
        var ast = try parseFor3D(src);
        defer ast.deinit(testing.allocator);
        try testing.expect(checkThreeD(ast, .{ .array_formula = true }) == null);
    }
}

test "3D: the quoted single-token spelling is a span too" {
    // `'Q1:Q4'!A1` reaches the parser as one quoted name, because a
    // tokenizer cannot tell one name with a colon from two around one.
    // Excel forbids `:` in a sheet name, so it is a span every time.
    var ast = try parseFor3D("COUNTBLANK('Q1:Q4'!A1)");
    defer ast.deinit(testing.allocator);
    try testing.expectEqual(
        Refusal.Reason.three_d_ineligible_function,
        checkThreeD(ast, .{}).?.reason,
    );

    var ok = try parseFor3D("SUM('Q1:Q4'!A1)");
    defer ok.deinit(testing.allocator);
    try testing.expect(checkThreeD(ok, .{}) == null);
}

test "3D: spans expand inclusively in workbook order; broken endpoints pin #REF!" {
    // Inclusive both ends — `Sheet1:Sheet3` is three sheets.
    const three = expandSpan(0, 2).members;
    try testing.expectEqual(@as(u32, 0), three.first);
    try testing.expectEqual(@as(u32, 2), three.last);
    // A span of one sheet is a span.
    try testing.expectEqual(@as(u32, 1), expandSpan(1, 1).members.last);

    // A missing endpoint — the ordinary shape of a deleted sheet.
    try testing.expect(expandSpan(null, 2) == .ref_error);
    try testing.expect(expandSpan(0, null) == .ref_error);
    // Reordered endpoints are not silently normalized (§5.6g).
    try testing.expect(expandSpan(2, 0) == .ref_error);
}

test "3D: splitting a span yields its two endpoint spellings" {
    const two_token = splitSpan(.{ .first = "S1", .last = "S3" }, "S1").?;
    try testing.expectEqualStrings("S1", two_token.first);
    try testing.expectEqualStrings("S3", two_token.last);

    const quoted = splitSpan(.{ .first = "'Q1:Q4'", .quoted = true }, "Q1:Q4").?;
    try testing.expectEqualStrings("Q1", quoted.first);
    try testing.expectEqualStrings("Q4", quoted.last);

    try testing.expect(splitSpan(.{ .first = "S1" }, "S1") == null);
}

test "checkAllAllocationFailures: scanning defined names leaks nothing under OOM" {
    const H = struct {
        fn run(allocator: Allocator) !void {
            const xml = wb_head ++ "<definedNames>" ++
                "<definedName name=\"A\">S!$A$1</definedName>" ++
                "<definedName name=\"B\" localSheetId=\"0\" hidden=\"1\">S!$B$2</definedName>" ++
                "<definedName name=\"C\" function=\"1\">MACRO</definedName>" ++
                "<definedName name=\"D\"/>" ++
                "</definedNames></workbook>";
            var got = switch (try scanDefinedNames(allocator, xml)) {
                .ok => |g| g,
                .refused => return error.TestUnexpectedRefusal,
            };
            defer got.deinit();
            if (got.rows.len != 4) return error.WrongCount;
            if (!got.rows[2].macro) return error.WrongClass;
        }
    };
    try testing.checkAllAllocationFailures(testing.allocator, H.run, .{});
}

// ─── fuzz (§8.1) ─────────────────────────────────────────────────

fn fuzzNamesTarget(_: void, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();
    var raw: [384]u8 = undefined;
    const input = raw[0..smith.slice(&raw)];
    const a = std.testing.allocator;

    // 1. The inventory, over an attribute region the fuzzer supplies.
    //    The element around it stays well-formed so the target is the
    //    table rather than the scanner.
    var buf: [512]u8 = undefined;
    const xml = std.fmt.bufPrint(&buf, "<definedName {s}/>", .{input}) catch return;
    var sc = decode.Scanner.init(xml);
    if (sc.next() catch null) |ev| {
        const el: ?decode.Element = switch (ev) {
            .open, .self_closing => |e| e,
            else => null,
        };
        if (el) |element| {
            const first = classifyDefinedName(element, "");
            const second = classifyDefinedName(element, "");
            // Total, and deterministic: every input is an answer, and
            // the same answer twice. A name that classified two ways
            // would resolve two ways.
            switch (first) {
                .ok => |ok| {
                    try std.testing.expect(second == .ok);
                    try std.testing.expectEqual(ok.macro, second.ok.macro);
                    try std.testing.expectEqual(ok.hidden, second.ok.hidden);
                    try std.testing.expectEqual(ok.local_sheet_id, second.ok.local_sheet_id);
                    // A macro flag and a referencing refusal are one
                    // fact; either both or neither.
                    try std.testing.expectEqual(ok.macro, ok.refusal_when_referenced != null);
                },
                .refused => |r| {
                    try std.testing.expect(second == .refused);
                    try std.testing.expectEqual(r.reason, second.refused.reason);
                    _ = r.planeTwo();
                },
            }
        }
    }

    // 2. The whole-part scanner over the same bytes: an answer or a
    //    typed refusal, never a crash and never a leaked arena.
    var part: [640]u8 = undefined;
    if (std.fmt.bufPrint(
        &part,
        wb_head ++ "<definedNames><definedName {s}/></definedNames></workbook>",
        .{input},
    )) |wb| {
        switch (try scanDefinedNames(a, wb)) {
            .ok => |g| {
                var g2 = g;
                g2.deinit();
            },
            .refused => |r| _ = r.planeTwo(),
        }
    } else |_| {}

    // 3. The 3D matrix, over whatever the input parses as. Two runs,
    //    one answer — a span that resolved two ways would be a span
    //    whose eligibility depended on when it was asked.
    var parsed = try parser.parse(a, input, .{});
    defer parsed.deinit(a);
    switch (parsed) {
        .refused => {},
        .ok => |ast| {
            const one = checkThreeD(ast, .{});
            const two = checkThreeD(ast, .{});
            try std.testing.expectEqual(one == null, two == null);
            if (one) |r| {
                try std.testing.expectEqual(r.reason, two.?.reason);
                _ = r.planeTwo();
                // An array context can only refuse *more*: a formula
                // that refuses plainly cannot become legal by being a
                // CSE array.
                try std.testing.expect(checkThreeD(ast, .{ .array_formula = true }) != null);
            }
        },
    }
}

test "fuzz: no defined-name attribute string or 3D span can panic, leak, or resolve two ways" {
    try std.testing.fuzz({}, fuzzNamesTarget, .{
        .corpus = &[_][]const u8{
            "name=\"N\"",
            "name=\"N\" localSheetId=\"0\" hidden=\"1\"",
            "name=\"N\" function=\"1\" vbProcedure=\"1\" xlm=\"1\"",
            "name=\"N\" comment=\"c\" description=\"d\" xml:space=\"preserve\"",
            "name=\"_xlnm.Print_Area\"",
            "SUM(Sheet1:Sheet3!A1)",
            "COUNTBLANK('Q1:Q4'!A1:B2)",
            "SUM(@Sheet1:Sheet3!A1)",
            "SUM(A1:A9 Sheet1:Sheet3!A1)",
            "_xlfn._xlws.FILTER(A1:A9,B1:B9)",
        },
    });
}
